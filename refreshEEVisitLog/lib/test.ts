import { submitFile } from './localMissionFile'


submitFile({
    payeeName: 'testpayeeName',
    amount: '0.01',
    description: 'testdesc',
    logger: msg => console.log(msg),
    reimbursementCat: 'Chinese New Year Carnival',
    attachements: [],
    ccList:[],
}).catch(err => {
    console.log(err);
    console.log(err.stack)
    console.log(err.response?err.response.data:err.message)
})