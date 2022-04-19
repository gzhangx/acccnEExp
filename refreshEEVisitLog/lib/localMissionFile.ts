//const Jimp = require('jimp');
import * as JSZip from "jszip"
import * as moment from 'moment-timezone';
//const email = require('./nodemailer');
import * as fs from 'fs';
//const { fstat } = require('fs');
import {msGraph} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt,IMsGraphDirPrms,IMsGraphExcelItemOpt} from "@gzhangx/googleapi/lib/msGraph/types";

import { getMSClientTenantInfo } from './ms'

interface ILocalCats {
    subCode: string;
    expCode: string;
    name: string;
}

interface INameBufAttachement {
    name: string;
    buffer: string;
}
export function getCategories() : ILocalCats[] {
    return `1	1600	Chinese New Year Carnival
2	1602	Ministry (Music Events, Guest Speaker)
3	1603	EE Training
5	1604	Organization support
4	1604A	Local Community Outreach Activity
6	1604	Family Keepers
6	1604	Family ministry Seminars (2)
6	1604	Herald Monthly
6	1604	美國華福總幹事 General Secretary, CCCOW
6	1605	Annual Budget Contribute to SECCC Pool of F
7	1607	Local Medias (Xin Times)
10	1612	In Town 信望愛 students and scholars Ministr
11	1611	(福音營)Gospel Camp financial aid`.split('\n').map((l, ind) => {
        const parts = l.split('\t');
        return {
            subCode: parts[0],
            expCode: parts[1],
            name: parts[2],
        }
    });
}

async function generateXlsx(xlsxFileName: string, replaces: { row: number; amount: string; column: string; }[], today:string, payeeName:string, description: string) {
    const zip = new JSZip();

    const z = await zip.loadAsync(fs.readFileSync('./files/ExpenseTemplate.xlsx'));

    const doFileReplace = async (fname: string, repFunc: (str:string)=>string) => {
        console.log(`Replcing ${fname} `);
        const origStr = await z.file(fname).async('string');
        console.log(`Replcing ${fname} ${origStr}`);
        zip.file(fname, repFunc(origStr));
    }

    await doFileReplace('xl/worksheets/sheet1.xml', origStr => {                
        const replaceOne = (acc,rpl) => {
            const { row, amount, column } = rpl;            
            console.log('before ' + acc.substr(acc.indexOf(`<c r="${column}${row}`), 40));
            const r1 = new RegExp(`<c r="${column}${row}" s="44"/>`);
            const r2 = new RegExp(`<c r="${column}${row}"([ ]*[s|t]="[0-9s]+"[ ]*)*><v>([0-9]+)?</v></c>`);

            const replaceTo = `<c r="${column}${row}" s="44"><v>${amount}</v></c>`;
            return acc.replace(r1, replaceTo).replace(r2, replaceTo);
        }
        return replaces.reduce((acc, rpl) => {            
            return replaceOne(acc, rpl);
        }, replaceOne(origStr, {
            column: 'J',
            row: 34,
            amount: '',
        }));
    });

    await doFileReplace('xl/sharedStrings.xml', origStr => {
        origStr = origStr.replace(/_2021-03-17_/g, `_${today}_`);
        origStr = origStr.replace('<t>Xin Times</t>', `<t>${payeeName}</t>`);
        origStr = origStr.replace('Payee (in English):           Xin Times', `Payee (in English):           ${payeeName}`);
        origStr = origStr.replace('<t>Safehouse Service Project - Provide and serve dinner for the homeless on Friday 1/29/2021</t>',`<t>${description}</t>`)
        return origStr;
    });

    const genb64 = await zip.generateAsync({ type: 'base64' });
    console.log(`writting file ${xlsxFileName}`);
    fs.writeFileSync(xlsxFileName, Buffer.from(genb64, 'base64'))
}


export interface ISubmitFileInterface{
    payeeName: string;
    reimbursementCat: string;
    amount: string;
    description: string;
    attachements: INameBufAttachement[];
    ccList: string[];
    logger: (msg: string) => void;
}

export async function submitFile({
    payeeName,
    reimbursementCat,
    amount,
    description,
    attachements = [],
    ccList,
    logger,
}: ISubmitFileInterface) {    
    // const fnt = await Jimp.loadFont(Jimp.FONT_SANS_12_BLACK);
    // const AMTX = 1220;
    // const AMTYSTART = 670;
    // const AMTYEND = 977;
    const AMTCATS = getCategories();
    //console.log(AMTCATS);

    const { found, row } = AMTCATS.reduce((acc, cat, pos) => {
        if (!acc.found) {
            if (cat.expCode === reimbursementCat || cat.name === reimbursementCat) {
                acc.found = cat;
                acc.row = pos + 25;
            }
        }
        return acc;
    }, {} as {found:ILocalCats, row: number});
    if (!found) {
        const err = { message: `not found ${reimbursementCat} ` };
        console.log(err.message);
        return err;
    }
    
    const nowMoment = moment();
    const today = nowMoment.format('YYYY-MM-DD');
    // console.log(`amtPos=${amtPos} ${today}`);
    // const img = await jimpRead('./files/expenseVoucher.PNG');
    const useDesc = description || found.name;
    // img.print(fnt, 272, 161, payeeName)
    //     .print(fnt, AMTX, amtPos, amount)
    //     .print(fnt, 1227, 1351, amount)
    //     .print(fnt, 1227, 1351, amount)
    //     .print(fnt, 232, 1517, 'Gang Zhang')
    //     .print(fnt, 790, 1517, today)
    //     .print(fnt, 232, 1583, submittedBy || payeeName)
    //     .print(fnt, 790, 1583, today)
    //     .print(fnt, 222, 1455, useDesc)
    //     .quality(60) // set JPEG quality
    //     //.greyscale() // set greyscale
    //     .write('./temp/accchForm.jpg'); // save
    
    const YYYY = moment().format('YYYY');
    const sheetOps = await msGraph.msExcell.getMsExcel(getMSClientTenantInfo(), {
        logger,
        sharedUrl: 'https://acccnusa-my.sharepoint.com/:x:/r/personal/gangzhang_acccn_org/Documents/Documents/safehouse/expenses.xlsx?d=wa6013afc83f64e6c9096851414d2d6b3&csf=1&web=1&e=3ga7Fb',
    }, {
        fileName: 'Documents/safehouse/expenses.xlsx',
    });
    const sheetName = nowMoment.format('YYYY');
    logger(`reading sheet ${sheetName}`);
    const curData = await sheetOps.readAll(sheetName);
    const vals = curData.values.filter(v => v[0]);
    console.log(vals)
    console.log(`length ${vals.length}`)
    vals.push([today, amount, found.subCode, found.expCode, payeeName, useDesc])
    console.log(vals)
    console.log(`length ${vals.length}`)
    await sheetOps.updateRange(sheetName, 'A1', `F${vals.length}`, vals);
    //await ops.append(`'LM${YYYY}'!A1`,
        //[[today, amount, found.subCode, found.expCode, useDesc, payeeName, today]]);
    console.log(`googlesheet appended`);
    const xlsxFileName = './temp/accchForm.xlsx';


    await generateXlsx(xlsxFileName, [
        { row, amount, column: 'J' },
        { row: 49, amount, column: 'J' }
    ], today, payeeName, description );
    console.log(`file generated`);
    const convertAttachement = (orig:INameBufAttachement) => {
        const origB64 = orig.buffer;
        const indPos = origB64.indexOf(',');
        const b64 = indPos >= 0 ? origB64.slice(indPos + 1) : origB64;
        //data:image/jpeg;base64,
        const matched = origB64.match(/data:(.+);base64,/);
        let contentType = '';
        if (matched) {
            contentType = matched[1]
        }
        console.log(orig.name+ " " + contentType);
        console.log(b64.slice(0, 20));
        //require('fs').writeFileSync('temp/test.jpg', Buffer.from(b64,'base64'))
        return {
            fileName: orig.name,
            content: Buffer.from(b64, 'base64'),
            path: '',
            contentType,
        }
    }
    const message = {
        from: '"LocalMissionBot" <gzhangx@gmail.com>',
        //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
        to: ['gzhangx@hotmail.com'].concat(ccList||[]),
        subject: `From ${payeeName} for ${found.name} Amount ${amount}`,
        text: `
        Date: ${today}
        subCode: ${found.subCode}
        expCode: ${found.expCode}
        amount: ${amount}
        payee: ${payeeName}
        `,
        attachments: [{
            fileName: 'expense.xlsx',
            path: xlsxFileName,
            contentType: '',
        }].concat(attachements.map(convertAttachement))
    };
    console.log(`sending email`);
    //await email.sendGmail(message);
    console.log(message);
}


module.exports = {
    getCategories,
    submitFile,
}