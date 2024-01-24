import * as msal from '@azure/msal-node';
import {AccountInfo} from '@azure/msal-node';
import {callMicrosoftGraph} from './graph';
import {
  DataProtectionScope,
  Environment,
  PersistenceCreator,
  PersistenceCachePlugin,
} from '@azure/msal-node-extensions';
import path from 'path';

const userRootDirectory = Environment.getUserRootDirectory();
const cachePath = userRootDirectory
  ? path.join(userRootDirectory, './cache.json')
  : '';

const persistenceConfiguration = {
  cachePath,
  dataProtectionScope: DataProtectionScope.CurrentUser,
  usePlaintextFileOnLinux: false,
};

PersistenceCreator.createPersistence(persistenceConfiguration).then(
  async (persistence: any) => {
    const publicClientConfig = {
      auth: {
        clientId: '72bfd166-a740-4899-9424-a018be5bae57',
        authority:
          'https://login.microsoftonline.com/f446b955-3842-4a00-8668-645bad7daad3',
      },
      cache: {
        cachePlugin: new PersistenceCachePlugin(persistence),
      },
    };

    // tsc transpiles import() to require(), despite various tsconfig.json settings I've tried.
    // runtime eval to force use of import() solves the problem.
    const {default: open} = await eval("import('open')");

    const openBrowser = async (url: string) => {
      if (open) {
        await open(url);
      } else {
        throw new Error('Failed to import open module');
      }
    };

    const loginRequest = {
      scopes: ['User.Read'],
      openBrowser,
      successTemplate: 'Successfully signed in! You can close this window now.',
    };

    // Create msal application object
    const pca = new msal.PublicClientApplication(publicClientConfig);

    const acquireToken = async () => {
      const msalTokenCache = pca.getTokenCache();
      const accounts = await msalTokenCache.getAllAccounts();
      if (accounts.length === 1) {
        const silentRequest = {
          account: accounts[0],
          scopes: ['User.Read'],
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
    };

    acquireToken()
      .then(response => {
        if (response) {
          return callMicrosoftGraph(response.accessToken)
            .then((graphResponse: msal.AuthenticationResult) => {
              console.log(graphResponse);
            })
            .catch((e: Error) => {
              console.error(e);
              throw e;
            });
        } else {
          throw new Error('Response is undefined.');
        }
      })
      .catch(e => {
        console.error(e);
        throw e;
      });
  }
);
