import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { msGraph } from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";

import { getMsDirClientPrms } from '../refreshEEVisitLog/lib/ms'
import { IMsDirOps } from '@gzhangx/googleapi/lib/msGraph/msdir';
import { delay } from "@gzhangx/googleapi/lib/msGraph/msauth";
import * as moment from 'moment-timezone'

import { getUtil } from './util'
const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const getPrm = name => (req.query[name] || (req.body && req.body[name])) as string;
    const action = getPrm('action');    

    const today = getPrm('today');
    function returnError(error) {
        context.log(error);
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: {
                error: error
            }
        };
    }

    if (!today) {
        return returnError('Must set today');
    }
    const util = await getUtil(today, context.log);
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
        return (fname);
    }
    //await store.getAllDataNoCache();
    let responseMessage = null;
    

    type IDataWithError = {
        error?: string;
        length?: number;
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

    
    if (action === "saveGuest") {
        const name = getPrm('name');
        const email = getPrm('email') || '';
        const picture = util.addPathToImg(getPrm('picture') || '');
        context.log(`saveGuest for ${name}:${email}`);
        if (!name) {
            return returnError('Must have name or email');
        } else {
            responseMessage = await util.saveGuest({
                name,
                email,
                picture,
            }).then(() => {
                return `user ${name} Saved`;
            }).catch(getErrorHndl(`user save error for ${name}:${email}`));
        }
    } else if (action === 'loadData') {
        responseMessage = await util.loadTodayData();
        //responseMessage = await store.loadData(msDirPrm).catch(getErrorHndl('loadData Error'));
    } else if (action === 'loadImage') {
        context.res.setHeader("Content-Type", "image/png")
        const fname = checkFileName();
        if (!fname) {
            return returnError('bad file name, return')
        }        
        const ary = await util.getFileByPath(fname).catch(getErrorHndl(`unable to load image ${fname}`)) as IDataWithError;
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
        try {
            const res = await await util.saveImage(util.addPathToImg(fname), dataStr)
            context.res = {
                body: {
                    id: res.id,
                    file: res.file,
                    size: res.size,
                }
            };
        } catch (err) {
            getErrorHndl(`saveImage createFile error for ${fname} ${dataStr.length}`)(err);
        }
        return;
    }


    context.res = {
        // status: 200, /* Defaults to 200 */
        body: responseMessage
    };

};

export default httpTrigger;