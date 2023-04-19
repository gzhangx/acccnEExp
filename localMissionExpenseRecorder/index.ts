import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { ILogger } from "@gzhangx/googleapi/lib/msGraph/msauth";

import { submitFile, getCategories, updateSums, getUserToCategories } from '../refreshEEVisitLog/lib/localMissionFile'
const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const name = (req.query.name || (req.body && req.body.name));
    const responseMessage = name
        ? "Hello, " + name + ". This HTTP triggered function executed successfully."
        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";
    
    function errorRsp(error: string) {
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: {
                error,
            },
        };
    }
    const reqBody = req.body;    
    const action = reqBody?.action || req.query.action;
    let res = null;
    const logger: ILogger = (...msg) => context.log(...msg);
    logger(`invoked==========> ${action}`);
    if (action === 'getCats') {
        try {
            res = await getCategories(logger);
        } catch (err) {
            logger('Error getCats', err);
            res = {
                message: err.message,
            }
        }
    } else if (action === 'getUserCats') {
        try {
            res = await getUserToCategories(logger);        
        } catch (err) {
            logger('Error getUserCats', err);
            res = {
                message: err.message,
            }
        }        
    } else if (action === 'saveFile') {
        logger(`invoked ${action} ${reqBody.payeeName} ${reqBody.amount} ${reqBody.reimbursementCat}`);
        if (!reqBody.amount) {
            return errorRsp('no amount');
        }
        if (!reqBody.payeeName) {
            return errorRsp('no payeeName');
        }
        if (!reqBody.reimbursementCat) {
            return errorRsp('no reimbursementCat');
        }
    
        res = await submitFile({
            amount: reqBody.amount,
            description: reqBody.description,
            logger,
            payeeName: reqBody.payeeName,
            reimbursementCat: reqBody.reimbursementCat,
            attachements: reqBody.attachements || [],
            ccList: reqBody.ccList,
        }).catch(err => {
            context.log(`error happened in submitFile`, err);
            return {
                error: err,
            }
        })
    } else if (action === 'updateSums') {
        res = await updateSums(context.log)
    } else {
        res = {
            message:'bad action'
        }
    }
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: res,
    };

};

export default httpTrigger;