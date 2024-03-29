//const Jimp = require('jimp');
import * as moment from 'moment';
//const email = require('./nodemailer');
import * as fs from 'fs';
//const { fstat } = require('fs');
import {msGraph, util} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt,IMsGraphDirPrms,IMsGraphExcelItemOpt} from "@gzhangx/googleapi/lib/msGraph/types";
import { ILogger } from '@gzhangx/googleapi/lib/msGraph/msauth';
import { getMSClientTenantInfo, treatFileName } from './ms'
import { emailTransporter, emailUser} from './nodemailer'
import { IMsDirOps } from '@gzhangx/googleapi/lib/msGraph/msdir';
import {v1 as uuidv1} from 'uuid';

const sleep = util.sleep;
interface ILocalCats {
    subCode: string;
    expCode: string;
    name: string;
}

interface IUserToLocalCats  {    
    shortName: string;
    fullName: string;
    catName: string;
}

interface INameBufAttachement {
    name: string;
    buffer: string;
}


async function getSheetData(logger: ILogger, fileName: string, sheetName: string) : Promise<string[][]> {
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const sheetOps = await msGraph.msExcell.getMsExcel(msGrapDirPrms, {
        fileName,
    });
    const allSheet = await sheetOps.readAll(sheetName);
    const data = allSheet.values;
    return data;
}

async function getLocalMissionRecordData(logger: ILogger, sheetName: string) {
    return await getSheetData(logger, 'localMissionRecords.xlsx', sheetName);
}
export async function getCategories(logger: ILogger): Promise<ILocalCats[]> {
    //const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    //const sheetOps = await msGraph.msExcell.getMsExcel(msGrapDirPrms, {        
        //fileName: 'Documents/safehouse/empty2022expense.xlsx',
    //});
    //const allSheet = await sheetOps.readAll('Table B');
    //const data = allSheet.values;
    const dataOrig = await getLocalMissionRecordData(logger, 'Categaries');
    const data = dataOrig.slice(1);
    const res: ILocalCats[] = [];
    for (let i = 0; i < data.length; i++) {
        const subCode = data[i][0];
        const name = data[i][1];
        const expCode = data[i][2];        
        res.push({            
            subCode,
            name,
            expCode,
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

export async function getUserToCategories(logger: ILogger): Promise<IUserToLocalCats[]> {
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);    
    const data = await getLocalMissionRecordData(logger, 'UserCats')    
    const res: IUserToLocalCats[] = [];
    for (let i = 0; i < data.length; i++) {
        const shortName = data[i][0];        
        const fullName = data[i][1];        
        const catName = data[i][2];
        res.push({
            shortName,
            fullName,
            catName,
        })
        //console.log(`sc=${sc} exp=${exp} subCode=${subCode} expCode=${expCode}`)        
    }
    return res;
}

export interface ISubmitFileInterface{
    payeeName: string;
    reimbursementCat: ILocalCats;
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

function prepareExpenseSheet(found:ILocalCats,payeeName: string, amount: string, date: string, desc: string, data: string[][], logger:ILogger) {
    let row = 50;
    row = 3;
    data[row][0] = replaceStrUnderlines(data[row][0], payeeName);
    const AMTPOS = 9;

    function* getCatRowIds() {
        for (let i = 24; i <= 35; i++) {
            yield i;
        }
    }
    let foundReplacement = false;
    for (let i of getCatRowIds()) {
        //const sc = parseInt(data[i][6]);
        //const exp = data[i][8];
        const cat = data[i][7];
        //console.log(`sc=${sc} exp=${exp} subCode=${subCode} expCode=${expCode}`)
        if (cat === found.name ) {            
            data[i][AMTPOS] = amount;
            foundReplacement = true;
            break;
        }
    }
    if (!foundReplacement) {
        logger(`not found, finding by ${found.expCode} to ${found.name}`)
        for (let i of getCatRowIds()) {
            const sc = parseInt(data[i][6]);
            const exp = parseInt(data[i][8]);
            const cat = data[i][7];
            
            if (sc === parseInt(found.subCode) && exp === parseInt(found.expCode)) {
                data[i][AMTPOS] = amount;
                data[i][7] = found.name;
                foundReplacement = true;
                logger(`not found, ffound by cat sc=${sc} exp=${exp} subCode=${sc} expCode=${exp}`)
                break;
            }
        }   
    }
    if (!foundReplacement) {
        logger('Warning, not found!!!!!!!');
        
    }
    data[48][AMTPOS] = amount;
    row = 50;
    data[row][0] = replaceStrUnderlines(data[row][0], desc || '');
    row = 51;
    data[row][0] = replaceStrUnderlines(data[row][0], 'Guangsen');
    const submitDatePos = 7;
    data[row][submitDatePos] = replaceStrUnderlines(data[row][submitDatePos], date);
    row = 53;
    data[row][0] = replaceStrUnderlines(data[row][0], 'Guangsen');
    data[row][submitDatePos] = replaceStrUnderlines(data[row][submitDatePos], date);
    return foundReplacement;
}

const SAVE_DOC_ROOT = 'savedRecords';
async function processRequestTemplateXlsx(msdirOps: IMsDirOps, newFileFullPath: string, fileInfo: ISubmitFileInterface, today:string, found:ILocalCats, logger: ILogger) {
    logger('fixing file', newFileFullPath);
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    //const msdirOps = await msGraph.msdir.getMsDir(getMSClientTenantInfo(), msGrapDirPrms);
    msGrapDirPrms.driveInfo = msdirOps.driveInfo;
    //const newFileName = treatFileName(`${today}-${found.name}`);
    
    const newId = await msdirOps.copyItemByName('empty2022expense.xlsx', newFileFullPath)
    logger('newFileId is ', newId);
    const sheetOps = await msGraph.msExcell.getMsExcel(msGrapDirPrms, {
        itemId: newId,
        //fileName: newFileFullPath,
    });
    logger('Reading sheet:Table B');
    const sheetRes = await sheetOps.readAll('Table B')
    //logger(sheetRes.values)
    //console.log(sheetRes.values);
    //sheetRes.values[50][0] = 'testtestesfaasdfadfaf';
    logger('prepareExpenseSheet');
    logger('found',found)
    logger('fileInfo',fileInfo)
    logger(today)
    const properlyReplaced = prepareExpenseSheet(found, fileInfo.payeeName, fileInfo.amount, today, fileInfo.description, sheetRes.values, logger);
    logger('done prepareExpenseSheet, update range ' + properlyReplaced);
    const origFileBuf = await msdirOps.getFileByPath(newFileFullPath);
    await sheetOps.updateRange('Table B', 'A1', `J${sheetRes.values.length}`, sheetRes.values);
    logger('done update range, get file by path ' + newFileFullPath);
    logger('done update range, sleep 500 before get file by path');
    await sleep(500);
    let newFileBuf = await msdirOps.getFileByPath(newFileFullPath);
    logger(`got file content debRmOrigFileBuf ================> orig=${origFileBuf.length}, newF=${newFileBuf.length}`);
    let curNewFileSize = 0;
    for (let sizeCheck = 0; sizeCheck < 10; sizeCheck++) {
        curNewFileSize = newFileBuf.length;
        if (newFileBuf.length !== origFileBuf.length) break;
        newFileBuf = await msdirOps.getFileByPath(newFileFullPath);
        logger(`got file content debRmOrigFileBuf ================> orig=${origFileBuf.length}, newF=${newFileBuf.length} at ${sizeCheck}`);
    }

    await sleep(500);
    newFileBuf = await msdirOps.getFileByPath(newFileFullPath);
    for (let sizeCheck = 0; sizeCheck < 10; sizeCheck++) {
        curNewFileSize = newFileBuf.length;
        if (newFileBuf.length === curNewFileSize) {
            logger(`newFileSizeCheck and orig are the same, file no longer grows ================> orig=${curNewFileSize}, newF=${newFileBuf.length} at ${sizeCheck}`);
            curNewFileSize = newFileBuf.length;
            break;
        }
        logger(`newFileSizeCheck and orig not the same, file still growing ================> orig=${curNewFileSize}, newF=${newFileBuf.length} at ${sizeCheck}`);
        curNewFileSize = newFileBuf.length;
        await sleep(500);
        newFileBuf = await msdirOps.getFileByPath(newFileFullPath);        
    }
    return {
        newFileBuf,
        msdirOps,
        properlyReplaced,
    }
}

function getGraphDirPrms(logger: ILogger) {
    const msGrapDirPrms: IMsGraphDirPrms = {
        creds: getMSClientTenantInfo(logger),
        sharedUrl: 'https://acccnusa.sharepoint.com/:x:/r/sites/LocalMission/_layouts/15/Doc.aspx?sourcedoc=%7BBE9B5D2A-2618-41CA-A7F7-660D68F641CB%7D&file=empty2022expense.xlsx&action=default&mobileredirect=true',
    }
    return msGrapDirPrms;
}

async function getSheetOps(logger:ILogger) {
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const sheetOps = await msGraph.msExcell.getMsExcel(msGrapDirPrms, {
        fileName: 'localMissionRecords.xlsx',
    });
    return {
        sheetOps,
        msGrapDirPrms,
    }
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
    const myId = uuidv1();
    logger('SubmitFile, id =', myId);
    const AMTCATS = await getCategories(logger);
    //console.log(AMTCATS);

    const found = AMTCATS.find(c => c.name === reimbursementCat.name);    
    if (!found) {
        const err = { message: `not found ${reimbursementCat} ` };
        console.log(err.message);
        return err;
    }
    
    const nowMoment = moment.default();
    const today = nowMoment.format('YYYY-MM-DD');
    const toddayWithSec = nowMoment.format('YYYY-MM-DD-HHmmss');
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
    
    const YYYY = moment.default().format('YYYY');
    const SAVE_DOC_ROOT_YYYY = `${SAVE_DOC_ROOT}/${YYYY}`;

    const {
        msGrapDirPrms,
        sheetOps,
    } = await getSheetOps(logger);


    const sheetName = nowMoment.format('YYYY');
    logger(`reading sheet ${sheetName}`);
    const curData = await sheetOps.readAll(sheetName);
    const vals = curData.values.filter(v => v[0]);    
    const files = (attachements || []).map(a => a.name).join(',');

    const convertAttachement = (orig: INameBufAttachement) => {
        const origB64 = orig.buffer;
        const indPos = origB64.indexOf(',');
        const b64 = indPos >= 0 ? origB64.slice(indPos + 1) : origB64;
        //data:image/jpeg;base64,
        const matched = origB64.match(/data:(.+);base64,/);
        let contentType = '';
        if (matched) {
            contentType = matched[1]
        }
        logger(orig.name + " " + contentType);
        //logger(b64.slice(0, 20));
        //require('fs').writeFileSync('temp/test.jpg', Buffer.from(b64,'base64'))
        return {
            fileName: orig.name,
            content: Buffer.from(b64, 'base64'),
            //path: '',
            encoding: '',
            contentType,
        }
    }
    const msgAttachements = attachements.map(convertAttachement);
    const newFileName = treatFileName(`${toddayWithSec}-${found.name}`);
    const newFileFullPath = `${SAVE_DOC_ROOT_YYYY}/${newFileName}.xlsx`;
    const actualNames = [];
    const msdirOps = await msGraph.msdir.getMsDir(msGrapDirPrms);
    for (let i = 0; i < msgAttachements.length; i++) {
        const att = msgAttachements[i];
        const sepInd = att.fileName.replace(/\\/g, '/').lastIndexOf('/');
        let filename = att.fileName;
        if (sepInd > 0) {
            filename = filename.substring(sepInd);
        }
        const saveFn = `${SAVE_DOC_ROOT_YYYY}/${newFileName}-${treatFileName(filename)}`;
        logger(`Saving ${saveFn}`);
        actualNames.push(saveFn);
        await msdirOps.createFile(saveFn, att.content);
    }
    const newRow = [today, amount, found.subCode, found.expCode, payeeName, useDesc, found.name, newFileFullPath, actualNames.join(','), files, myId];
    vals.push(newRow);
    vals.forEach((vs, ind) => {
        if (vs.length === newRow.length) return;
        while (vs.length < newRow.length) {
            vs.push('');
        }
        if (vs.length > newRow.length) {
            vs = vs.slice(0, newRow.length);
            vals[ind] = vs;
        }
    });
    
    const toColName = (ind: number) => String.fromCharCode('A'.charCodeAt(0) + ind);
    await sheetOps.updateRange(sheetName, 'A1', `${toColName(newRow.length-1)}${vals.length}`, vals);
    //await ops.append(`'LM${YYYY}'!A1`,
        //[[today, amount, found.subCode, found.expCode, useDesc, payeeName, today]]);
    logger(`googlesheet appended`);
    

    logger('processRequestTemplateXlsx');    
    const newFileInfo = await processRequestTemplateXlsx(msdirOps, newFileFullPath, submitFileInfo, today, found, logger);



    logger(`file generated newFileInfo=${newFileInfo.properlyReplaced}`);
    
    const emailsAry = await getLocalMissionRecordData(logger, 'emails');
    logger(`email lists`, emailsAry);
    const emailTo = emailsAry.map(e => e[0]);
    logger(`email lists To `, emailTo);
    const message = {
        from: `"LocalMissionBot" <${emailUser}>`,
        //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
        to: emailTo.concat(ccList||[]),
        subject: `From ${payeeName} for ${found.name} Amount ${amount} ${newFileInfo.properlyReplaced?'':' WARNING failed to replace cat'}`,
        text: `
        Dear brother George,
        Please see the attached reimbursement request for ${payeeName}, thanks!
        
        Date: ${today}
        subCode: ${found.subCode}
        expCode: ${found.expCode}
        category: ${found.name}
        amount: ${amount}
        payee: ${payeeName}
        description: ${description}
        `,
        attachments: [{
            fileName: newFileName,
            //path: xlsxFileName,
            content: newFileInfo.newFileBuf,
            //encoding:'base64',
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }].concat(msgAttachements)
    };    
    logger(`sending email`);
    //await email.sendGmail(message);
    const sendEmailRes = await emailTransporter.sendMail(message).catch(err => {
        logger(err);
        return {
            ...found,
            error: err.message,
        }
    });
    return {
        ...found,
        sendEmailRes,
    };
}


async function pmap<T, U>(items: T[], action: (data: T) => Promise<U>) {
    if (!items) return null;
    const res: U[] = [];
    for (let i = 0; i < items.length; i++) {
        res.push(await action(items[i]));
    }
    return res;
}

export async function resubmitLine(lineNum: number | string, logger: ILogger) {
    const msGrapDirPrms: IMsGraphDirPrms = getGraphDirPrms(logger);
    const sheetOps = await msGraph.msExcell.getMsExcel(msGrapDirPrms, {
        fileName: 'localMissionRecords.xlsx',
    });
    const YYYY = moment.default().format('YYYY');
    const allSheet = await sheetOps.readAll(YYYY);
    const datas = allSheet.values;    
    const [today, amount, subCode, expCode, payeeName, description, category, sheetName, filesByComma] = datas[lineNum as number] || datas.find(d => d[10] === lineNum);
    const msdirOps = await msGraph.msdir.getMsDir(msGrapDirPrms);

    const imgAttachements = await pmap(filesByComma.split(',').filter(x => x), async fileName => {
        let extInd = fileName.lastIndexOf('.');
        let ext = 'png';
        if (extInd > 0) {
            ext = fileName.substring(extInd+1)
        }
        const contentType = `image/${ext}`;
        return {
            fileName,
            content: await msdirOps.getFileByPath(fileName),
            contentType,
        }
    })
    const message = {
        from: `"LocalMissionBot" <${emailUser}>`,
        //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
        to: ['gzhangx@hotmail.com'],
        subject: `From ${payeeName} for ${category} Amount ${amount}`,
        text: `
        Date: ${today}
        subCode: ${subCode}
        expCode: ${expCode}
        category: ${category}
        amount: ${amount}
        payee: ${payeeName}
        description: ${description}
        `,
        attachments: [{
            fileName: sheetName,
            //path: xlsxFileName,
            content: await msdirOps.getFileByPath(sheetName),
            //encoding:'base64',
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }].concat(imgAttachements)
    };
    const sendEmailRes = await emailTransporter.sendMail(message).catch(err => {
        logger(err);
        return {
            error: err.message,
        }
    });
    return sendEmailRes;
}

export async function updateSums(logger:ILogger) {
    const nowMoment = moment.default();
    const sheetName = nowMoment.format('YYYY');
    logger(`reading sheet ${sheetName}`);
    const {
        sheetOps,
    } = await getSheetOps(logger);
    const sheetVals = await sheetOps.readAll(sheetName);
    const colNames = ['date', 'amount', 'code', 'exp', 'to', 'comment', 'cat']
    const sums = sheetVals.text.slice(1).reduce((acc, valAry) => {
        const val = colNames.reduce((valAcc, name, pos) => {
            let vv = valAcc[name] = valAry[pos].trim();
            if (name === 'amount') {
                vv.replace(/$/g, '');
                let amt = 0;
                if (vv.startsWith('(')) {
                    amt = -parseFloat(vv.replace(/[\(\)]/g, ''));
                } else {
                    amt = parseFloat(vv);
                }
                valAcc[name] = amt;
            }
            return valAcc;
        }, {} as { [name: string]: string | number; });
        val.id = `${val.exp}-${val.code}`;
        let existing = acc[val.id];
        if (!existing) {
            existing = {
                amt: val.amount as number,
                name: val.cat as string,
            };
            acc[val.id] = existing;
        } else {
            existing.amt += val.amount as number;
            if (!existing.name) existing.name = val.cat as string;
        }
        return acc;
    }, {} as { [name: string]: { amt: number; name: string; } });

    const keys = Object.keys(sums);
    keys.sort();
    const empty = ['', '', ''];
    const newData = keys.map(id => {
        const sum = sums[id];
        return [id, sum.amt.toFixed(2), sum.name];
    }).concat([empty, empty, empty]);    
    logger('resutling table', newData);

    //await sheetOps.createSheet('sums');
    await sheetOps.updateRange(`sums`, 'A1', `C${newData.length}`,newData);
    return newData;
}