import * as Msal from '@azure/msal-node';

const axios = require('axios').default;
import {graphMeEndpoint} from './authConfig.js';

export async function callMicrosoftGraph(accessToken: string) {
  console.log('Calling Microsoft Graph');
  const options = {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  };

  try {
    const response = await axios.get(graphMeEndpoint, options);
    return response.data;
  } catch (error) {
    console.log(error);
    return error;
  }
}
