//const Jimp = require('jimp');
import * as moment from 'moment-timezone';
//const email = require('./nodemailer');
import * as fs from 'fs';
//const { fstat } = require('fs');
import {msGraph} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt,IMsGraphDirPrms,IMsGraphExcelItemOpt} from "@gzhangx/googleapi/lib/msGraph/types";

import { getMSClientTenantInfo } from './ms'
import { emailTransporter, emailUser} from './nodemailer'

interface ILocalCats {
    subCode: string;
    expCode: string;
    name: string;
}

interface INameBufAttachement {
    name: string;
    buffer: string;
}

const treatFileName = (path:string)=>path.replace(/[\\"|*<>?]/g, '')
export type ILogger = (msg:any)=> void;
export async function getCategories(logger: ILogger): Promise<ILocalCats[]> {
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const sheetOps = await msGraph.msExcell.getMsExcel(getMSClientTenantInfo(), msGrapDirPrms, {        
        fileName: 'Documents/safehouse/empty2022expense.xlsx',
    });
    const allSheet = await sheetOps.readAll('Table B');
    const data = allSheet.values;
    const res: ILocalCats[] = [];
    for (let i = 24; i <= 35; i++) {
        const subCode = data[i][6];
        const expCode = data[i][8];
        const name = data[i][7];
        res.push({
            expCode,
            name,
            subCode,
        })
        //console.log(`sc=${sc} exp=${exp} subCode=${subCode} expCode=${expCode}`)        
    }
    return res;
    /*
    return `01	1601	Chinese New Year Carnival
    02	1602	Ministry (Music Events, Guest Speaker)
    03	1603	EE Training
    05	1604	Organization Support
    04	1604A	Local Community Outreach Activities
    06	1606	Family Keepers
    06	1606	Family ministry Seminars (2)
    06	1606	Herald Monthly
    06	1606	美國華福總幹事 General Secretary, CCCOW
    06	1605	Annual Budget Contribute to SECCC
    07	1607	Local Medias (Xin Times)
    10	1612	In Town 信望愛 students and scholars Ministry`.split('\n').map((l, ind) => {
        const parts = l.split('\t');
        return {
            subCode: parts[0].trim(),
            expCode: parts[1],
            name: parts[2],
        }
    });
    */
}

export interface ISubmitFileInterface{
    payeeName: string;
    reimbursementCat: string;
    amount: string;
    description: string;
    attachements: INameBufAttachement[];
    ccList: string[];
    logger: ILogger;
}

function replaceStrUnderlines(orig: string, content: string) {
    let firstInd = orig.indexOf('_');
    let lastInd = orig.lastIndexOf('_');
    if (firstInd < 0) return orig + ' ' + content;
    if (lastInd - firstInd < content.length) return orig.substring(0, firstInd) + content + orig.substring(lastInd + 1);

    let start = Math.round((lastInd - firstInd - content.length) / 2);
    return orig.substring(0, firstInd + start) + content + orig.substring(firstInd + start + content.length);
}

function prepareExpenseSheet(found:ILocalCats,payeeName: string, amount: string, date: string, desc: string, data: string[][]) {
    let row = 50;
    row = 3;
    data[row][0] = replaceStrUnderlines(data[row][0], payeeName);
    const AMTPOS = 9;
    for (let i = 24; i <= 35; i++) {
        //const sc = parseInt(data[i][6]);
        //const exp = data[i][8];
        const cat = data[i][7];
        //console.log(`sc=${sc} exp=${exp} subCode=${subCode} expCode=${expCode}`)
        if (cat === found.name ) {            
            data[i][AMTPOS] = amount;
            break;
        }
    }
    data[48][AMTPOS] = amount;
    row = 50;
    data[row][0] = replaceStrUnderlines(data[row][0], desc || '');
    row = 51;
    data[row][0] = replaceStrUnderlines(data[row][0], 'Gang');
    const submitDatePos = 7;
    data[row][submitDatePos] = replaceStrUnderlines(data[row][submitDatePos], date);
    row = 53;
    data[row][0] = replaceStrUnderlines(data[row][0], 'Gang');
    data[row][submitDatePos] = replaceStrUnderlines(data[row][submitDatePos], date);
}

async function processRequestTemplateXlsx(fileInfo: ISubmitFileInterface, today:string, found:ILocalCats, logger: ILogger) {
    logger('fixing file');
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const msdirOps = await msGraph.msdir.getMsDir(getMSClientTenantInfo(), msGrapDirPrms);
    msGrapDirPrms.driveId = msdirOps.driveId;
    const newFileName = treatFileName(`${today}-${found.name}`);
    const newFileFullPath = `Documents/safehouse/safehouseRecords/${newFileName}.xlsx`;
    const newId = await msdirOps.copyItemByName('Documents/safehouse/empty2022expense.xlsx', newFileFullPath)
    console.log('newFileId is ', newId);
    const sheetOps = await msGraph.msExcell.getMsExcel(getMSClientTenantInfo(), msGrapDirPrms, {
        itemId: newId,
        //fileName: newFileFullPath,
    });
    logger('Reading sheet:Table B');
    const sheetRes = await sheetOps.readAll('Table B')
    //logger(sheetRes.values)
    //console.log(sheetRes.values);
    //sheetRes.values[50][0] = 'testtestesfaasdfadfaf';
    logger('prepareExpenseSheet');
    logger(found)
    logger(fileInfo)
    logger(today)
    prepareExpenseSheet(found, fileInfo.payeeName, fileInfo.amount, today, fileInfo.description, sheetRes.values);
    logger('done prepareExpenseSheet, update range');
    await sheetOps.updateRange('Table B', 'A1', `J${sheetRes.values.length}`, sheetRes.values);
    logger('done update range, get file by path ' + newFileFullPath);
    const newFileBuf = await msdirOps.getFileByPath(newFileFullPath);
    logger('got file content');
    return {
        newFileName,
        newFileBuf,
    }
}

function getGraphDirPrms(logger: ILogger) {
    const msGrapDirPrms: IMsGraphDirPrms = {
        logger,
        sharedUrl: 'https://acccnusa-my.sharepoint.com/:x:/r/personal/gangzhang_acccn_org/Documents/Documents/safehouse/empty2022expense.xlsx?d=w1a9a3f0fe89a4f9f93314efc910315fd&csf=1&web=1&e=WSHzge',
    }
    return msGrapDirPrms;
}
export async function submitFile(submitFileInfo: ISubmitFileInterface) {    
    const {
        payeeName,
        reimbursementCat,
        amount,
        description,
        attachements = [],
        ccList,
        logger,
    } = submitFileInfo;
    // const fnt = await Jimp.loadFont(Jimp.FONT_SANS_12_BLACK);
    // const AMTX = 1220;
    // const AMTYSTART = 670;
    // const AMTYEND = 977;
    const AMTCATS = await getCategories(logger);
    //console.log(AMTCATS);

    const found = AMTCATS.find(c => c.name === reimbursementCat);    
    if (!found) {
        const err = { message: `not found ${reimbursementCat} ` };
        console.log(err.message);
        return err;
    }
    
    const nowMoment = moment();
    const today = nowMoment.format('YYYY-MM-DD');
    // console.log(`amtPos=${amtPos} ${today}`);
    // const img = await jimpRead('./files/expenseVoucher.PNG');
    const useDesc = description;
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
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const sheetOps = await msGraph.msExcell.getMsExcel(getMSClientTenantInfo(), msGrapDirPrms, {
        fileName: 'Documents/safehouse/localMissionRecords.xlsx',
    });


    const sheetName = nowMoment.format('YYYY');
    logger(`reading sheet ${sheetName}`);
    const curData = await sheetOps.readAll(sheetName);
    const vals = curData.values.filter(v => v[0]);    
    const files = (attachements || []).map(a => a.name).join(',');
    vals.push([today, amount, found.subCode, found.expCode, payeeName, useDesc,found.name, files]);
    vals.forEach((vs,ind) => {
        if (vs.length === 8) return;
        while (vs.length < 8) {
            vs.push('');
        }
        if (vs.length > 8) {
            vs = vs.slice(0, 8);
            vals[ind] = vs;
        }
    })
    await sheetOps.updateRange(sheetName, 'A1', `H${vals.length}`, vals);
    //await ops.append(`'LM${YYYY}'!A1`,
        //[[today, amount, found.subCode, found.expCode, useDesc, payeeName, today]]);
    logger(`googlesheet appended`);
    

    logger('processRequestTemplateXlsx');    
    const newFileInfo = await processRequestTemplateXlsx(submitFileInfo, today, found, logger);



    logger(`file generated`);
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
        logger(orig.name+ " " + contentType);
        //logger(b64.slice(0, 20));
        //require('fs').writeFileSync('temp/test.jpg', Buffer.from(b64,'base64'))
        return {
            fileName: orig.name,
            content: Buffer.from(b64, 'base64'),
            //path: '',
            encoding:'',
            contentType,
        }
    }
    const message = {
        from: `"LocalMissionBot" <${emailUser}>`,
        //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
        to: ['gzhangx@hotmail.com'].concat(ccList||[]),
        subject: `From ${payeeName} for ${found.name} Amount ${amount}`,
        text: `
        Date: ${today}
        subCode: ${found.subCode}
        expCode: ${found.expCode}
        category: ${found.name}
        amount: ${amount}
        payee: ${payeeName}
        description: ${description}
        `,
        attachments: [{
            fileName: newFileInfo.newFileName,
            //path: xlsxFileName,
            content: newFileInfo.newFileBuf,
            //encoding:'base64',
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }].concat(attachements.map(convertAttachement))
    };
    logger(`sending email`);
    //await email.sendGmail(message);
    await emailTransporter.sendMail(message).catch(err => {
        logger(err);
    })
    return {
        message:'done'
    }
    //logger(message);
}


module.exports = {
    getCategories,
    submitFile,
}