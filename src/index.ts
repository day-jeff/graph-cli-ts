import {callMicrosoftGraph} from './graph';
import {msalClient} from './msalClient';
import figlet from 'figlet';
import {Command, Option, OptionValues} from 'commander';
import {get} from 'http';

async function main() {
  console.log(figlet.textSync('Graph CLI'));

  const options = getOptions();

  await msalClient.initialize();
  const result = await msalClient.authenticate(['user.read']);

  let graphUri = '';
  if (result) {
    if (options.me) {
      graphUri = 'https://graph.microsoft.com/v1.0/me';
    }
    if (options.users) {
      graphUri = 'https://graph.microsoft.com/v1.0/users';
    }

    const graphResponse = await callMicrosoftGraph(
      result.accessToken,
      graphUri
    );

    console.log(graphResponse);
  }
}

function getOptions(): OptionValues {
  const program = new Command();

  program
    .version('0.0.1')
    .description('A CLI for querying the Microsoft Graph')
    .addOption(new Option('-m, --me', 'View my profile').conflicts('users'))
    .addOption(new Option('-u, --users', 'View all users').conflicts('me'))
    .parse(process.argv);

  const NO_COMMAND_SPECIFIED = process.argv.length < 3;
  if (NO_COMMAND_SPECIFIED) {
    program.help();
  }
  return program.opts();
}

main();
