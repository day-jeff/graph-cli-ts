/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
import * as Msal from '@azure/msal-node';
import {authConfig} from './authConfig';
import {callMicrosoftGraph} from './graph';

const open = require('open');

// Before running the sample, you will need to replace the values in src/authConfig.js

// Open browser to sign user in and consent to scopes needed for application
const openBrowser = async (url: string) => {
  // You can open a browser window with any library or method you wish to use - the 'open' npm package is used here for demonstration purposes.
  open(url);
};

const loginRequest = {
  scopes: ['User.Read'],
  openBrowser,
  successTemplate: 'Successfully signed in! You can close this window now.',
};

// Create msal application object
const pcApp = new Msal.PublicClientApplication(authConfig);

const acquireToken = async () => {
  const msalTokenCache = pcApp.getTokenCache();
  const accounts = await msalTokenCache.getAllAccounts();
  if (accounts.length === 1) {
    const silentRequest = {
      account: accounts[0],
      scopes: ['User.Read'],
    };

    return pcApp.acquireTokenSilent(silentRequest).catch((e: Error) => {
      if (e instanceof Msal.InteractionRequiredAuthError) {
        return pcApp.acquireTokenInteractive(loginRequest);
      }
      throw e;
    });
  } else if (accounts.length > 1) {
    accounts.forEach(account => {
      console.log(account.username);
    });
    return Promise.reject(
      'Multiple accounts found. Please select an account to use.'
    );
  } else {
    return pcApp.acquireTokenInteractive(loginRequest);
  }
};

acquireToken()
  .then(response => {
    if (response) {
      return callMicrosoftGraph(response.accessToken)
        .then((graphResponse: Msal.AuthenticationResult) => {
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
