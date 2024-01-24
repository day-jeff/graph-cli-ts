import * as Msal from '@azure/msal-node';

export const cacheLocation = './src/data/cache.json';
export const cachePlugin = require('./cachePlugin')(cacheLocation);

// Define the type
interface AuthConfig {
  auth: {
    clientId: string;
    authority: string;
  };
  cache: {
    cachePlugin: Msal.DistributedCachePlugin;
  };
  system: {
    loggerOptions: {
      loggerCallback: (loglevel: Msal.LogLevel, message: string) => void;
      piiLoggingEnabled: boolean;
      logLevel: Msal.LogLevel;
    };
  };
}

// Assign the value
export const authConfig: AuthConfig = {
  auth: {
    clientId: '72bfd166-a740-4899-9424-a018be5bae57',
    authority:
      'https://login.microsoftonline.com/f446b955-3842-4a00-8668-645bad7daad3',
  },
  cache: {
    cachePlugin: cachePlugin,
  },
  system: {
    loggerOptions: {
      loggerCallback: (loglevel: Msal.LogLevel, message: string) => {
        console.log(message);
      },
      piiLoggingEnabled: false,
      logLevel: Msal.LogLevel.Trace,
    },
  },
};

export const graphMeEndpoint = 'https://graph.microsoft.com/v1.0/me';
