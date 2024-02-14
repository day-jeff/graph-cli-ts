import keytar = require('keytar');
import chalk from 'chalk';

const CRED_SERVICE = 'some_online_service';
const CRED_ACCOUNT = 'someone@ra.rockwell.com';
const CRED_PASSWORD = 'password123';

keytar
  .setPassword(CRED_SERVICE, CRED_ACCOUNT, CRED_PASSWORD)
  .then(() => {
    return keytar.getPassword(CRED_SERVICE, CRED_ACCOUNT);
  })
  .then((password: any) => {
    console.log(
      `\nSuccessfully saved password '${password}' to credential vault and read it back.`
    );
    console.log(chalk.green('\nKeytar functioning correctly.\n'));
  })
  .catch((err: Error) => {
    console.log('Error: ' + err.toString());
  });
