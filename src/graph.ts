import axios from 'axios';

export async function callMicrosoftGraph(
  accessToken: string,
  graphUri: string
) {
  console.log('Calling Microsoft Graph');
  const options = {
    headers: {
      Authorization: `Bearer ${accessToken}`,
    },
  };

  const response = await axios.get(graphUri, options);
  return response.data;
}
