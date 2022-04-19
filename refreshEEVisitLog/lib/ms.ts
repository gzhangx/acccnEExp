import {msGraph} from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";

const tenantClientInfo: IMsGraphCreds = {
    client_id: '72f543e0-817c-4939-8925-898b1048762c',
    refresh_token: process.env.REFRESH_TOKEN,
    tenantId:'60387d22-1b13-42a0-8894-208eeafd9e57', //https://portal.azure.com/#home, https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps
}
export function getMSClientTenantInfo() {        
    return tenantClientInfo;
}