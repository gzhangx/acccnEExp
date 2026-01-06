import * as fs from 'fs';
const jscfg = JSON.parse(fs.readFileSync('./.vscode/launch.json', 'utf8'));
        const envObj = jscfg.configurations.find((c: { name: string; }) => c.name === 'Run root test.ts').env
        Object.keys(envObj).forEach(k => {
            process.env[k] = envObj[k];
        });
import * as dailySender from './localMissionExpenseRecorder/bibleSender/getdata'
//import * as sendWeek from './localMissionExpenseRecorder/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail'
import * as sendWeek from './hebrewsEmailNotificationSender/lib/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail'
import { sendBTAData } from './refreshEEVisitLog/lib/btaEmail';


async function test(retFirst: string) {

    if (retFirst === "sendBtaEmail") {        
        const logger = (msg: any) => console.log(msg);
        
        //newGuestRegTest();
        
        await sendBTAData({
            date: new Date(),
            logger: s=>console.log(s),
        });
        return;
    }
    if (retFirst === 'sendSheetNotice') {

        
        const test = await sendWeek.sendSheetNotice({
            logger: console.log,
            sendEmail: 'N',
        });
        return console.log(test);
    }
    const got = await dailySender.sendBibleData({
        logger: console.log,
        sendEmail: 'N',
    });
    console.log(got);
}

test('sendSheetNotice');