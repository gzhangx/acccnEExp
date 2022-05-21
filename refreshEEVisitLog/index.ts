import { AzureFunction, Context, HttpRequest } from "@azure/functions"
import { msGraph } from "@gzhangx/googleapi"
import { ILogger } from "@gzhangx/googleapi/lib/msGraph/msauth";
import { IMsGraphCreds, IAuthOpt,IMsGraphDirPrms,IMsGraphExcelItemOpt} from "@gzhangx/googleapi/lib/msGraph/types";
import { getMsDirClientPrms, generateRefreshTokenCode, getRefreshToken } from './lib/ms'


async function calculateEEVisitTimes(logger:ILogger) {
    const prm: IMsGraphDirPrms = getMsDirClientPrms('https://acccnusa-my.sharepoint.com/:x:/r/personal/gangzhang_acccn_org/Documents/%E4%B8%89%E7%A6%8F%E6%8E%A2%E8%AE%BF%E8%AE%B0%E5%BD%95.xlsx?d=wf3a17698953344988a206fbe0fecded5&csf=1&web=1&e=sMhg4O',
        logger);
    const opt: IMsGraphExcelItemOpt = {
        //itemId: '01XX2KYFMYO2Q7GM4VTBCIUIDPXYH6ZXWV',
        fileName:'三福探访记录.xlsx'
    };    
    logger('getting sheet')
    const sheet = await msGraph.msExcell.getMsExcel(prm, opt);
    logger('got sheet done, reading sheet1')
    const dataAll = await sheet.readAll('Sheet1');
    logger('got sheet read sheet 1 done')
    logger(JSON.stringify(dataAll.text));

    const studentMap = (await sheet.readAll('Students')).text.reduce((acc, line) => {
        acc[line[0]] = true;
        return acc;
    }, {} as { [name: string]: boolean });
    logger('studentMap', studentMap);

    const isLeader: {[name:string]:boolean} = {};
    const summary =dataAll.text.slice(1).reduce((acc, d) => {
        const leader = d[4];
        const std = d[5].split(/[,，]+/);
        const doAdd = (name: string) => {
            name = name.trim();
            if (name)
                acc[name] = (acc[name] || 0) + 1;
        }
        isLeader[leader] = true;
        doAdd(leader);
        std.forEach(doAdd);
        return acc;
    }, {
    } as { [name: string]: number });
    logger(JSON.stringify(summary))
    const updateDataParts = Object.keys(summary).sort().reduce((acc, name) => {
        const data = [name, summary[name].toString()] as [string, string];
        let ary = acc.leaders;
        if (studentMap[name]) {
            ary = acc.students;
        }
        ary.push(data);
        return acc;
    }, {
        leaders: [] as [string, string][],
        students: [] as [string, string][],
    });

    let updateData = [] as [string, string][];
    updateData = updateData.concat([['学员', '']]).concat(updateDataParts.students)
        .concat([['', '']])
        .concat([['老师,other', '']])
        .concat(updateDataParts.leaders);
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

    function retErr(error:string) {
        context.res = {
            body: {
                error,
            }
        }
    }
    context.log(`name is ${name} (can be refreshGetCode|waitToken)`);
    let result: any;
    if (!req.query.name) {
        result = await calculateEEVisitTimes(msg => context.log(msg));
    } else if (req.query.name === 'refreshGetCode') {
        context.log(`refreshGetCode`);
        result = await generateRefreshTokenCode(context.log);
        context.log(`refreshGetCode result`,result);
    } else if (req.query.name === 'waitToken') {        
        const device_code = req.query.device_code;
        if (!device_code) {
            return retErr('no device code');
        }
        context.log(`waitToken ${device_code}`);
        result = await getRefreshToken(context.log, device_code).catch(err => {
            context.log('error happened in getRefreshToken', err);
            return err;
        })
        context.log(result);
    }
    context.res = {
        // status: 200, /* Defaults to 200 */
        body: result
    };

};

export default httpTrigger;