

import { emailTransporter, emailUser } from '../bibleSender/nodemailer';
import Moment from 'moment-timezone';
import *as gs from '@gzhangx/googleapi';
export type LoggerType = (...arg: any[]) => void;
async function createOps() {
    const gsKeyInfo: gs.gsAccount.IServiceAccountCreds = {
        client_email: process.env.GS_CLIENT_EMAIL,
        private_key: process.env.GS_PRIVATE_KEY.replace(/\\n/g, '\n'),
        private_key_id: process.env.GS_PRIVATE_KEY_ID,
    }; // = JSON.parse(fs.readFileSync('./data/secrets/gospelCamp.json').toString());    
    const client = gs.gsAccount.getClient(gsKeyInfo);
    return client;
}
async function readValues() {
    const client = await createOps();
    const ops = await client.getSheetOps("1qgGpKDF5blPj8c-pFDhz7VkT1tvyllkT5rl4s4I0Dd0");
  //const sheet = gsSheet.getOpsBySheetId('1uYTYzwjUN8tFpeejHtiGA5u_RtOoSBO8P1b2Qg-6Elk', logger);
  //const ops = await sheet.getOps();
    const ret = await ops.readData('EmailTemplate');
    //return ret.data;
    return ret.values;
}


export type BtaDataOpts = {
    date: Date;
    logger: (str: string, err: object) => void;
}
export async function sendBTAData(opts: BtaDataOpts) {
    const templates = await readValues();
    const nowInput = Moment(opts.date);
    if (!process.env.BTA_EMAIL) {
        return {
            err: 'must set process.env.BTA_EMAIL',
            message: 'must set process.env.BTA_EMAIL',
        }
    }
    // 0 Sunday, 1 mon, 2 tu, 3 wed 4 thu 5-Fri   6-sat
    const saturdayMMDDYYYY = nowInput.weekday(6).format('MM/DD/YYYY');
    const thursdayMMDDYYYY = nowInput.weekday(4).format('MM/DD/YYYY');
    const date = nowInput.format('YYYY-MM-DD');
    const template = templates.map(text => {
        return text[0].replace(/\$Date/g, date).replace(/\$Saturday/g, saturdayMMDDYYYY)
            .replace(/\$Thursday/g, thursdayMMDDYYYY);
    })

    const message = {
        from: `"Open Arms Saturday event Auto reminder" <${emailUser}>`,
        to: template[0].split(','),
        subject: template[1], // 'weekly opr email , ' + nowInput.format('YYYY-MM-DD'),
        text: template[2] || `Open Arms Saturday event:
School has started and we will be back to our normal Saturday event at true love daycare center.
Please confirm this Saturday （${saturdayMMDDYYYY}) event by Thursday (${thursdayMMDDYYYY})

Time 2:00pm-4:00pm
Location: TL DAYCARE CENTER
Address: 5050 Research Ct #650, Suwanee, GA 30024

1. Alan Zhou -
2. Brendon
3. Cayden -
4. Calvin & Andrew
5. Chase
6. Connor
7. Ethan Chen
8. Jackie & Ellie
9. Jeremy
10. Jessica
11. kevin -
12. Kian -
13. Koki -
14. Lena
15. Mathis
16. Max
17. Raylee & DD
18. Samuel & Jeremy Lin
19. Tyler

No walk in please. Please let me know if you have specific learning requirements.

@All 

`,
    };

    //console.log('sending ' + message.subject);    
    
    const sendEmailRes = await emailTransporter.sendMail(message).catch(err => {
        opts.logger('Error send email', err);
        return {
            error: err.message,
        }
    });
    opts.logger('SendEmailRes', sendEmailRes);
    return sendEmailRes;
}
