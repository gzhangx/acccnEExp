import { submitFile, resubmitLine } from './localMissionFile'

import * as newGuestRegUtil from '../../acccnGuestRegistration/util'
const logger = msg => console.log(msg);

newGuestRegTest();

async function newGuestRegTest() {
    const util = await newGuestRegUtil.getUtil('2022-04-test', console.log);
    return util.saveGuest({
        name:'test',
        email:'testemail',
        picture: 'testpic'
    })
}

//resubmitLine(15, logger);

function testold() {
    submitFile({
        payeeName: 'testpayeeName',
        amount: '0.01',
        description: 'testdesc',
        logger,
        reimbursementCat: 'Chinese New Year Carnival',
        attachements: [],
        ccList: [],
    }).catch(err => {
        console.log(err);
        console.log(err.stack)
        console.log(err.response ? err.response.data : err.message)
    }).then(() => {
        console.log('sending 16');
        return resubmitLine(15, logger);
    })
}
