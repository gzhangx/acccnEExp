import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { brotliDecompressSync } from "zlib";

import { submitFile, getCategories} from '../refreshEEVisitLog/lib/localMissionFile'
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
    let res = null;
    if (reqBody.action === 'getCats') {
        res = await getCategories(msg => context.log(msg));
    } else if (reqBody.action === 'saveFile') {
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
            logger: msg => context.log(msg),
            payeeName: reqBody.payeeName,
            reimbursementCat: reqBody.reimbursementCat,
            attachements: reqBody.attachements,
            ccList: reqBody.ccList,
        }).catch(err => {
            res = {
                error: err,
            }
            context.log(err);
        })
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