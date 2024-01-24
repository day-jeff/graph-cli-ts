import * as Msal from '@azure/msal-node';

export function InitializePersistentCache() {
  const {
    DataProtectionScope,
    Environment,
    PersistenceCreator,
    PersistenceCachePlugin,
  } = require('@azure/msal-node-extensions');

  // You can use the helper functions provided through the Environment class to construct your cache path
  // The helper functions provide consistent implementations across Windows, Mac and Linux.
  const cachePath = path.join(
    Environment.getUserRootDirectory(),
    './cache.json'
  );

  const persistenceConfiguration = {
    cachePath,
    dataProtectionScope: DataProtectionScope.CurrentUser,
    serviceName: '<SERVICE-NAME>',
    accountName: '<ACCOUNT-NAME>',
    usePlaintextFileOnLinux: false,
  };

  // The PersistenceCreator obfuscates a lot of the complexity by doing the following actions for you :-
  // 1. Detects the environment the application is running on and initializes the right persistence instance for the environment.
  // 2. Performs persistence validation for you.
  // 3. Performs any fallbacks if necessary.
  PersistenceCreator.createPersistence(persistenceConfiguration).then(
    async persistence => {
      const publicClientConfig = {
        auth: {
          clientId: '<CLIENT-ID>',
          authority: '<AUTHORITY>',
        },

        // This hooks up the cross-platform cache into MSAL
        cache: {
          cachePlugin: new PersistenceCachePlugin(persistence),
        },
      };

      const pca = new msal.PublicClientApplication(publicClientConfig);

      // Use the public client application as required...
    }
  );
}

interface ICachePlugin {
  beforeCacheAccess: (tokenCacheContext: TokenCacheContext) => Promise<void>;
  afterCacheAccess: (tokenCacheContext: TokenCacheContext) => Promise<void>;
}

class MyCachePlugin implements ICachePlugin {
  private client: ICacheClient;

  constructor(client: ICacheClient) {
    this.client = client; // client object to access the persistent cache
  }

  public async beforeCacheAccess(
    cacheContext: TokenCacheContext
  ): Promise<void> {
    const cacheData = await this.client.get(); // get the cache from persistence
    cacheContext.tokenCache.deserialize(cacheData); // deserialize it to in-memory cache
  }

  public async afterCacheAccess(
    cacheContext: TokenCacheContext
  ): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      await this.client.set(cacheContext.tokenCache.serialize()); // deserialize in-memory cache to persistence
    }
  }
}
