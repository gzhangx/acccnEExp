import * as dailySender from './localMissionExpenseRecorder/bibleSender/getdata'
import * as sendWeek from './localMissionExpenseRecorder/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail'
async function test(retFirst: string) {

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