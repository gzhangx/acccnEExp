import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import {msGraph} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt,IMsGraphDirPrms,IMsGraphExcelItemOpt} from "@gzhangx/googleapi/lib/msGraph/types";



async function test(logger:(msg:string)=>void) {
    let refresh_token = process.env.REFRESH_TOKEN;
    const tenantClientInfo: IMsGraphCreds = {
        client_id: '72f543e0-817c-4939-8925-898b1048762c',
        refresh_token,
        tenantId:'60387d22-1b13-42a0-8894-208eeafd9e57', //https://portal.azure.com/#home, https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
    }
        
    const prm: IMsGraphDirPrms = {        
        logger,
        sharedUrl: 'https://acccnusa-my.sharepoint.com/:x:/r/personal/gangzhang_acccn_org/Documents/%E4%B8%89%E7%A6%8F%E6%8E%A2%E8%AE%BF%E8%AE%B0%E5%BD%95.xlsx?d=wf3a17698953344988a206fbe0fecded5&csf=1&web=1&e=sMhg4O',
        driveId:'',
    };
    const opt: IMsGraphExcelItemOpt = {
        //itemId: '01XX2KYFMYO2Q7GM4VTBCIUIDPXYH6ZXWV',
        fileName:'三福探访记录.xlsx'
    };    
    logger('getting sheet')
    const sheet = await msGraph.msExcell.getMsExcel(tenantClientInfo, prm, opt);
    
    const dataAll = await sheet.readAll('Sheet1');
    logger('got sheet done')
    logger(JSON.stringify(dataAll.text));

    const summary =dataAll.text.slice(1).reduce((acc, d) => {
        const leader = d[4];
        const std = d[5].split(/[,，]+/);
        const doAdd = (name: string) => {
            name = name.trim();
            acc[name] = (acc[name] || 0) + 1;
        }
        doAdd(leader);
        std.forEach(doAdd);
        return acc;
    }, {
    } as { [name: string]: number });
    logger(JSON.stringify(summary))
    const updateData = Object.keys(summary).sort().map(name => {
        return [name, summary[name].toString()];
    })
    logger(JSON.stringify(updateData));
    const creatRes = await sheet.createSheet('Summary');
    logger(`create res`);
    logger(JSON.stringify(creatRes));
    await sheet.updateRange('Summary', 'A1', `B${updateData.length}`, updateData);
    return updateData;
}

const httpTrigger: AzureFunction = async function (context: Context, req: HttpRequest): Promise<void> {
    context.log('HTTP trigger function processed a request.');
    const name = (req.query.name || (req.body && req.body.name));
    const responseMessage = name
        ? "Hello, " + name + ". This HTTP triggered function executed successfully."
        : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";

    const result = await test(msg => context.log(msg));
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: JSON.stringify(result)
    };

};

export default httpTrigger;