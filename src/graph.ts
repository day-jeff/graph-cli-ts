const axios = require('axios').default;

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

  try {
    const response = await axios.get(graphUri, options);
    return response.data;
  } catch (error) {
    console.log(error);
    return error;
  }
}
