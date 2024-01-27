import * as msal from '@azure/msal-node';
import {AccountInfo} from '@azure/msal-node';
import {
  DataProtectionScope,
  Environment,
  PersistenceCreator,
  PersistenceCachePlugin,
} from '@azure/msal-node-extensions';
import path from 'path';
import {auth} from './config';

let pca: msal.PublicClientApplication;

export class msalClient {
  private static pca: msal.PublicClientApplication;
  private static tokenCache: msal.TokenCache;
  private static accounts: msal.AccountInfo[];

  constructor() {
    msalClient.initialize();
  }

  static async initialize() {
    msalClient.pca = await msalClient.getPCA();
    msalClient.tokenCache = msalClient.pca.getTokenCache();
    msalClient.accounts = await msalClient.tokenCache.getAllAccounts();
    console.log(msalClient.accounts.length + ' accounts found');
  }

  static async getPCA() {
    const userRootDirectory = Environment.getUserRootDirectory();
    const cachePath = userRootDirectory
      ? path.join(userRootDirectory, './cache.json')
      : '';

    const persistenceConfiguration = {
      cachePath: cachePath,
      serviceName: "Microsoft Graph",
      accountName: "Graph CLI user",
      dataProtectionScope: DataProtectionScope.CurrentUser,
      usePlaintextFileOnLinux: false,
    };

    const persistence = await PersistenceCreator.createPersistence(
      persistenceConfiguration
    );

    const publicClientConfig = {
      auth,
      cache: {
        cachePlugin: new PersistenceCachePlugin(persistence),
      },
    };

    return new msal.PublicClientApplication(publicClientConfig);
  }

  static async authenticate(
    scopes: string[]
  ): Promise<msal.AuthenticationResult> {
    const openBrowser = async (url: string) => {
      // tsc transpiles import() to require(), despite various tsconfig.json settings I've tried.
      // runtime eval to force use of import() solves the problem.
      const {default: open} = await eval("import('open')");

      if (open) {
        await open(url);
      } else {
        throw new Error('Failed to import open module');
      }
    };

    const loginRequest = {
      scopes: scopes,
      openBrowser,
      successTemplate: 'Successfully signed in! You can close this window now.',
    };

    if (msalClient.accounts.length === 1) {
      const silentRequest = {
        account: msalClient.accounts[0],
        scopes: scopes,
      };
      return msalClient.pca
        .acquireTokenSilent(silentRequest)
        .catch((e: Error) => {
          if (e instanceof msal.InteractionRequiredAuthError) {
            return msalClient.pca.acquireTokenInteractive(loginRequest);
          }
          throw e;
        });
    } else if (msalClient.accounts.length > 1) {
      msalClient.accounts.forEach((account: AccountInfo) => {
        console.log(account.username);
      });
      return Promise.reject(
        'Multiple accounts found. Please select an account to use.'
      );
    } else {
      return msalClient.pca.acquireTokenInteractive(loginRequest);
    }
  }

  static async logout() {
    msalClient.accounts = await msalClient.tokenCache.getAllAccounts();
    if (msalClient.accounts.length > 0) {
      msalClient.accounts.forEach(async (account: AccountInfo) => {
        await msalClient.pca.getTokenCache().removeAccount(account);
      });
      console.log('Successfully signed out');
    } else {
      console.log('No accounts found');
    }
  }
}
