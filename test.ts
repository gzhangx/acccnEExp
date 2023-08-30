import * as dailySender from './localMissionExpenseRecorder/bibleSender/getdata'

async function test() {
    const got = await dailySender.sendBibleData({
        logger: console.log,
        sendEmail: 'N',
    });
    console.log(got);
}

test();