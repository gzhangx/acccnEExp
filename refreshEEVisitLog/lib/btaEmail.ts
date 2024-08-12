
import { emailTransporter, emailUser } from './nodemailer';
import Moment from 'moment-timezone';

export type BtaDataOpts = {
    date: Date;
    logger: (str: string, err: object) => void;
}
export async function sendBTAData(opts: BtaDataOpts) {
    const nowInput = Moment(opts.date);
    // 0 Sunday, 1 mon, 2 tu, 3 wed 4 thu 5-Fri   6-sat
    const saturdayMMDDYYYY = nowInput.weekday(6).format('MM/DD/YYYY');
    const thursdayMMDDYYYY = nowInput.weekday(4).format('MM/DD/YYYY');
    const message = {
        from: `"Hebrews Daily Bible verse" <${emailUser}>`,
        to: [process.env.BTA_EMAIL],
        subject: 'weekly opr email , ' + nowInput.format('YYYY-MM-DD'),
        text: `Open Arms Saturday event:
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
