/*
import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { ILogger } from "@gzhangx/googleapi/lib/msGraph/msauth";

import * as bibleSender from './lib/bibleSender/getdata';
import * as hebrewsSender from './lib/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail';
import { sendBTAData } from './lib/btaEmail/btaEmail';

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const reqBody = req.body;
    const action = reqBody?.action || req.query.action;
    let res = null;
    const logger: ILogger = (...msg) => context.log(...msg);
    logger(`invoked==========> ${action}`);
    
    if (action === 'sendHebrewsDailyEmail') {
        try {
            res = await bibleSender.sendBibleData({
                now: null,
                logger: context.log,
                sendEmail: 'Y',
            });
        } catch (err) {
            logger('Error sendHebrewsDailyEmail', err);
            res = {
                message: err.message,
            }
        }
    } else if (action === 'sendHebrewsWeeklyMeetingScheduleEmail') {
        try {
            res = await hebrewsSender.sendSheetNotice({
                logger: context.log,
            });
        } catch (err) {
            logger('Error sendHebrewsWeeklyMeetingScheduleEmail', err);
            res = {
                message: err.message,
            }
            if (err.steps) {
                res.steps = err.steps;
            }
        }
    } else if (action === 'sendBettinaEmail') {
        res = await sendBTAData({
            date: new Date(),
            logger: context.log,
        })
    } else {
        res = {
            message: 'bad action'
        }
    }
    context.res = {
        // status: 200, 
        body: res,
    };
    
};

export default httpTrigger;
*/