import figlet from 'figlet';
import chalk from 'chalk';

import * as msal from './msalClient';
import {callMicrosoftGraph} from './graph';
import {Command, Option, OptionValues} from 'commander';
import {isEmail} from 'validator';

async function main() {
  console.log(figlet.textSync('Graph CLI'));

  const options = getOptions();

  await msal.Initialize();
  const result = await msal.Authenticate(['user.read']);

  if (result) {
    let graphUri = '';
    if (options.allUsers) {
      graphUri = 'https://graph.microsoft.com/v1.0/users';
      callGraph(result.accessToken, graphUri);
    }
    if (options.files) {
      graphUri = 'https://graph.microsoft.com/v1.0/me/drive/root/children';
      callGraph(result.accessToken, graphUri);
    }
    if (options.logout) {
      msal.Logout();
    }
    if (options.me) {
      graphUri = 'https://graph.microsoft.com/v1.0/me';
      callGraph(result.accessToken, graphUri);
    }
    if (options.user) {
      const email = options.user;
      if (isEmail(email)) {
        graphUri = `https://graph.microsoft.com/v1.0/users/${email}`;
        callGraph(result.accessToken, graphUri);
      } else {
        console.log(
          chalk.red('Invalid email:'),
          `${email}. Please provide a complete email address.\n`
        );
      }
    }
  }
}

async function callGraph(accessToken: string, graphUri: string) {
  try {
    const graphResponse = await callMicrosoftGraph(accessToken, graphUri);
    console.log(graphResponse);
  } catch (err: any) {
    displayError(err);
  }
}

function getOptions(): OptionValues {
  const program = new Command();

  program
    .version('0.0.1')
    .description('A CLI for querying the Microsoft Graph')
    .addOption(
      new Option('-a, --all-users', 'View all users').conflicts([
        'files',
        'me',
        'user',
      ])
    )
    .addOption(
      new Option('-f, --files', 'View my files').conflicts([
        'all',
        'me',
        'user',
      ])
    )
    .addOption(
      new Option('-m, --me', 'View my profile').conflicts([
        'all',
        'files',
        'user',
      ])
    )
    .addOption(new Option('-o, --logout', 'Logout'))
    .addOption(
      new Option('-u, --user <email>', 'Look up user by email').conflicts([
        'all',
        'files',
        'me',
      ])
    )

    .parse(process.argv);

  const NO_COMMAND_SPECIFIED = process.argv.length < 3;
  if (NO_COMMAND_SPECIFIED) {
    program.help();
  }
  return program.opts();
}

function displayError(error: any) {
  const uri = error.response?.config.url;
  const status = error.response?.status;

  switch (status) {
    case 403:
      console.log(
        chalk.red('Not authorized error.'),
        `You are not authorized to view ${uri}.`
      );
      break;
    case 404:
      console.log(chalk.red(`HTTP ${status} error. URI not found:`), uri);
      break;
    default:
      console.log(chalk.red(`HTTP ${status} error`), `while querying ${uri}.`);
  }
  console.log('\n');
}

main();
