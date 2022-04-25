import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { msGraph } from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";

import { getMsDirClientPrms } from '../refreshEEVisitLog/lib/ms'
import { IMsDirOps } from '@gzhangx/googleapi/lib/msGraph/msdir';
import { delay } from "@gzhangx/googleapi/lib/msGraph/msauth";
import * as moment from 'moment-timezone'

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name])) as string;
    const action = getPrm('action');    

    const today = getPrm('today');
    function addPathToImg(fname: string) {
        const todayMoment = moment(today);
        const quarter = Math.floor(((todayMoment.month() + 1) % 12) / 3 + 1);
        const year = todayMoment.format('YYYY');
        if (!fname) return fname;
        return `新人资料/${year}-Q${quarter}-DBGRM/${today}/${fname}`;
    }
    if (!today.match(/^[0-9]{4}-[0-9]{2}-[0-9]{2}$/)) {
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: {
                error: 'Bad date ' + today,
            }
        };
        return;
    }
    context.log(`action=${action}`);
    function checkFileName() {
        const fname = getPrm('name');
        if (!fname || !fname.trim()) {
            context.res = {
                body: 'No filename',
            };
            return null;
        }
        return addPathToImg(fname);
    }
    //await store.getAllDataNoCache();
    let responseMessage = null;
    const msGraphPrms: IMsGraphDirPrms = getMsDirClientPrms('https://acccnusa.sharepoint.com/:x:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99%E8%A1%A8%E6%B1%87%E6%80%BBnew.xlsx?d=wbd57c301f851467787c3b5405709c2bf&csf=1&web=1&e=HDYYri',
        context.log);    
    async function getMsDirOpt() {        
        const ops = await msGraph.msdir.getMsDir(msGraphPrms);
        return ops;
    }
    const xlsOps = await msGraph.msExcell.getMsExcel(msGraphPrms, {
        fileName: '新人资料/新人资料表汇总new.xlsx'
    });    
    //const today = moment().format('YYYY-MM-DD');
    await xlsOps.createSheet(today);
    for (let i = 0; i < 100; i++) {
        const sheets = await xlsOps.getWorkSheets();
        const found = sheets.value.find(v => v.name === today);
        context.log(`Sheets (trying to find ${today}), waiting ${i*500}`,found);
        if (found) break;
        await delay(500);
    }

    type IDataWithError = {
        error?: string;
        length?: number;
    }

    function returnError(error) {
        context.log(error);
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: {
                error: error
            }
        };
    }
    function getErrorHndl(inf: string) {
        return (err): IDataWithError => {
            responseMessage = {
                error: `${inf} ${err.message}`
            }
            context.log(responseMessage.error);
            return responseMessage;
        }
    }

    const loadTodayData = async () => {
        const todayData = await xlsOps.readAll(today);
        return todayData.values.filter(v => v[0]);
    }
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email') || '';
        const picture = addPathToImg(getPrm('picture') || '');
        context.log(`saveGuest for ${name}:${email}`);
        if (!name) {
            return returnError('Must have name or email');
        } else {
            const existingRaw = await loadTodayData();
            const COLWIDTH = 3;
            const toUpdate = existingRaw.map(ex => {
                while (ex.length < COLWIDTH) ex.push('');
                if (ex.length === 3) return ex;                
                return ex.slice(0,3);
            });
            toUpdate.push([name, email, picture])
            responseMessage = `user ${name} Saved`;
            await xlsOps.updateRange(today, 'A1', `C${toUpdate.length}`, toUpdate).catch(getErrorHndl(`user save error for ${name}:${email}`));
        }
    } else if (action === 'loadData') {
        responseMessage = await loadTodayData();
        //responseMessage = await store.loadData(msDirPrm).catch(getErrorHndl('loadData Error'));
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        const fname = checkFileName();
        if (!fname) {
            return returnError('bad file name, return')
        }
        const ops = await getMsDirOpt();
        const ary = await ops.getFileByPath(fname).catch(getErrorHndl(`unable to load image ${fname}`)) as IDataWithError;
        if (ary.error) {
            responseMessage = ary.error;
        } else {
            context.log(`image size ${ary.length}`)
            context.res = {
                headers: {
                    "Content-Type": "image/png"
                },
                isRaw: true,
                // status: 200, /* Defaults to 200 */
                body: ary, //new Uint8Array(buffer)
            };
            return;
        }
    } else if (action === 'saveImage') {
        const fname = checkFileName();
        if (!fname) return returnError('No filename for saveImage');
        let dataStr = getPrm('data') as string;
        const sub = dataStr.indexOf('base64,');
        if (sub > 0) {
            dataStr = dataStr.substring(sub + 7).trim();
        }
        const buf = Buffer.from(dataStr, 'base64');
        const ops = await getMsDirOpt();
        try {
            const res = await ops.createFile(fname, buf);
            context.res = {
                body: {
                    id: res.id,
                    file: res.file,
                    size: res.size,
                }
            };
        } catch (err) {
            getErrorHndl(`saveImage createFile error for ${fname} ${buf.length}`)(err);
        }
        return;
    }


    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;