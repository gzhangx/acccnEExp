
import * as moment from 'moment-timezone'
import { msGraph } from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";

import { getMsDirClientPrms } from '../refreshEEVisitLog/lib/ms'
import { IMsDirOps } from '@gzhangx/googleapi/lib/msGraph/msdir';
import { delay, ILogger } from "@gzhangx/googleapi/lib/msGraph/msauth";

interface IGuestInfo {
    name: string;
    email: string;
    picture: string;
}
export async function getUtil(today: string, logger: ILogger) {
    function addPathToImg(fname: string) {
        const todayMoment = moment(today);
        const quarter = Math.floor(((todayMoment.month() + 1) % 12) / 3 + 1);
        const year = todayMoment.format('YYYY');
        if (!fname) return fname;
        return `新人资料/${year}-Q${quarter}-DBGRM/${today}/${fname}`;
    }

    const msGraphPrms: IMsGraphDirPrms = getMsDirClientPrms('https://acccnusa.sharepoint.com/:x:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99%E8%A1%A8%E6%B1%87%E6%80%BBnew.xlsx?d=wbd57c301f851467787c3b5405709c2bf&csf=1&web=1&e=HDYYri',
        logger);
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
        logger(`Sheets (trying to find ${today}), waiting ${i * 500}`, found);
        if (found) break;
        await delay(500);
    }
    
    const loadTodayData = async () => {
        const todayData = await xlsOps.readAll(today);
        return todayData.values.filter(v => v[0]);
    }

    async function saveGuest({ name, email, picture }: IGuestInfo) {
        const existingRaw = await loadTodayData();
        const COLWIDTH = 3;
        const toUpdate = existingRaw.map(ex => {
            while (ex.length < COLWIDTH) ex.push('');
            if (ex.length === 3) return ex;
            return ex.slice(0, 3);
        });
        toUpdate.push([name, email, picture])        
        await xlsOps.updateRange(today, 'A1', `C${toUpdate.length}`, toUpdate);
        return `user ${name} Saved`;
    }
    return {
        addPathToImg,
        msGraphPrms,
        getMsDirOpt,
        xlsOps,
        loadTodayData,

        saveGuest,
    }
}