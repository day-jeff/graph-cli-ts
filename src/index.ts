import {callMicrosoftGraph} from './graph';
import * as msalClient from './msalClient';

msalClient.authenticate(['user.read']).then(async result => {
  if (result) {
    const graphResponse = await callMicrosoftGraph(result.accessToken);
    console.log(graphResponse);
  }
});
