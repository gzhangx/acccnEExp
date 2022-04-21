import {msGraph} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";
import { ILogger, IRefreshTokenResult } from "@gzhangx/googleapi/lib/msGraph/msauth";
import * as fs from 'fs'

function getTokenFileLoc() {
    return `d:/home/data/Functions/sampledata/refreshEEVisitLog.json`;
}
export function getMSClientTenantInfo(logger: ILogger): IMsGraphCreds {
    let refresh_token = process.env.REFRESH_TOKEN;
    try {
        const dec = JSON.parse(fs.readFileSync(getTokenFileLoc()).toString()) as IRefreshTokenResult;
        refresh_token = dec.refresh_token;
    } catch (err) {
        logger(`Error get refresh token from file`, err);
    }
    return {
        client_id: '72f543e0-817c-4939-8925-898b1048762c',
        refresh_token,
        tenantId: '60387d22-1b13-42a0-8894-208eeafd9e57',
        logger,
    }
}

export function getMsDirClientPrms(sharedUrl: string, logger:ILogger) {
    const prm: IMsGraphDirPrms = {
        creds: getMSClientTenantInfo(logger),
        sharedUrl,
        driveId: '',
    };
    return prm;
}

export async function generateRefreshTokenCode(logger: ILogger) {
    return await msGraph.msauth.getAuth(getMSClientTenantInfo(logger)).refreshTokenSeperated.getRefreshTokenPart1GetCodeWaitInfo();
}


export async function getRefreshToken(logger: ILogger, deviceCode:string) {
    return await msGraph.msauth.getAuth(getMSClientTenantInfo(logger)).refreshTokenSeperated.getRefreshTokenPartFinish(deviceCode, async (tk) => {
        logger(`Saving token to ${getTokenFileLoc()}`, tk);
        fs.writeFileSync(getTokenFileLoc(), JSON.stringify(tk));
        logger('done save token')
    });
}