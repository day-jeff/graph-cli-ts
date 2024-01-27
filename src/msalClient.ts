import * as msal from '@azure/msal-node';
import * as msalextensions from '@azure/msal-node-extensions';
import path from 'path';

import {auth} from './config';
import {open} from 'fs';

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

let pca: msal.PublicClientApplication;
let tokenCache: msal.TokenCache;
let accounts: msal.AccountInfo[];

export async function Initialize() {
  pca = await getPCA();
  tokenCache = pca.getTokenCache();
  accounts = await tokenCache.getAllAccounts();
  console.log(accounts.length + ' accounts found');
}

async function getPCA() {
  const userRootDirectory = msalextensions.Environment.getUserRootDirectory();
  const cachePath = userRootDirectory
    ? path.join(userRootDirectory, './cache.json')
    : '';

  const persistenceConfiguration = {
    cachePath: cachePath,
    serviceName: 'Microsoft Graph',
    accountName: 'Graph CLI user',
    dataProtectionScope: msalextensions.DataProtectionScope.CurrentUser,
    usePlaintextFileOnLinux: false,
  };

  const persistence = await msalextensions.PersistenceCreator.createPersistence(
    persistenceConfiguration
  );

  const publicClientConfig = {
    auth,
    cache: {
      cachePlugin: new msalextensions.PersistenceCachePlugin(persistence),
    },
  };

  return new msal.PublicClientApplication(publicClientConfig);
}

export async function Authenticate(
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

  if (accounts.length === 1) {
    const silentRequest = {
      account: accounts[0],
      scopes: scopes,
    };
    return pca.acquireTokenSilent(silentRequest).catch(async (e: Error) => {
      if (e instanceof msal.InteractionRequiredAuthError) {
        return pca.acquireTokenInteractive(loginRequest);
      }
      throw e;
    });
  } else if (accounts.length > 1) {
    accounts.forEach((account: msal.AccountInfo) => {
      console.log(account.username);
    });
    return Promise.reject(
      'Multiple accounts found. Please select an account to use.'
    );
  } else {
    return pca.acquireTokenInteractive(loginRequest);
  }
}

export async function Logout() {
  // The following code deletes credentials cached on the machine, but it doesn't clear browser cookies.
  // This means that if you sign in again, you probably won't be prompted for credentials.
  // This is problematic if you want to sign in with a different account.
  accounts = await tokenCache.getAllAccounts();
  if (accounts.length > 0) {
    accounts.forEach(async (account: msal.AccountInfo) => {
      await pca.getTokenCache().removeAccount(account);
    });
    console.log('Successfully signed out');
  } else {
    console.log('No accounts found');
  }

  // This is the "v1" sign out URL. It's not officially endorsed to use this URL, but it works.
  // The v2 URL prompts the user to choose which account they want to sign out of, which is clunky.
  // @azure/msal-node doesn't have a built-in way to sign out, so this is the best option for now.
  const logoutUri = 'https://login.microsoftonline.com/common/oauth2/logout';
  openBrowser(logoutUri);
}
