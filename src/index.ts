import {msalClient} from './msalClient';
import figlet from 'figlet';
import {Command, Option, OptionValues} from 'commander';

async function main() {
  console.log(figlet.textSync('Graph CLI'));

  const options = getOptions();

  await msalClient.initialize();
  const result = await msalClient.authenticate(['user.read']);

  if (result) {
    let graphUri = '';
    if (options.me) {
      graphUri = 'https://graph.microsoft.com/v1.0/me';
      callMicrosoftGraph(result.accessToken, graphUri);
    }
    if (options.users) {
      graphUri = 'https://graph.microsoft.com/v1.0/users';
      callMicrosoftGraph(result.accessToken, graphUri);
    }
    if (options.logout) {
      msalClient.logout();
    }
  }
}

async function callMicrosoftGraph(accessToken: string, graphUri: string) {
  const graphResponse = await callMicrosoftGraph(accessToken, graphUri);
  console.log(graphResponse);
}

function getOptions(): OptionValues {
  const program = new Command();

  program
    .version('0.0.1')
    .description('A CLI for querying the Microsoft Graph')
    .addOption(new Option('-m, --me', 'View my profile').conflicts('users'))
    .addOption(new Option('-u, --users', 'View all users').conflicts('me'))
    .addOption(new Option('-o, --logout', 'Logout'))
    .parse(process.argv);

  const NO_COMMAND_SPECIFIED = process.argv.length < 3;
  if (NO_COMMAND_SPECIFIED) {
    program.help();
  }
  return program.opts();
}

main();
