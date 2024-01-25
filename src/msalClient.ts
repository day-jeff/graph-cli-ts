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

async function getPCA() {
  const userRootDirectory = Environment.getUserRootDirectory();
  const cachePath = userRootDirectory
    ? path.join(userRootDirectory, './cache.json')
    : '';

  const persistenceConfiguration = {
    cachePath,
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

export async function authenticate(
  scopes: string[]
): Promise<msal.AuthenticationResult> {
  if (!pca) {
    pca = await getPCA();
  }

  const msalTokenCache = pca.getTokenCache();
  const accounts = await msalTokenCache.getAllAccounts();

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

  if (accounts.length === 1) {
    const silentRequest = {
      account: accounts[0],
      scopes: scopes,
    };
    return pca.acquireTokenSilent(silentRequest).catch((e: Error) => {
      if (e instanceof msal.InteractionRequiredAuthError) {
        return pca.acquireTokenInteractive(loginRequest);
      }
      throw e;
    });
  } else if (accounts.length > 1) {
    accounts.forEach((account: AccountInfo) => {
      console.log(account.username);
    });
    return Promise.reject(
      'Multiple accounts found. Please select an account to use.'
    );
  } else {
    return pca.acquireTokenInteractive(loginRequest);
  }
}
