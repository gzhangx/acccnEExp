
import * as moment from 'moment-timezone'
import { msGraph } from "@gzhangx/googleapi"
import { IMsGraphCreds, IAuthOpt, IMsGraphDirPrms, IMsGraphExcelItemOpt } from "@gzhangx/googleapi/lib/msGraph/types";

import { getMsDirClientPrms, getStoreFileLoc } from '../refreshEEVisitLog/lib/ms'
import { IMsDirOps } from '@gzhangx/googleapi/lib/msGraph/msdir';
import { delay, ILogger } from "@gzhangx/googleapi/lib/msGraph/msauth";

import * as fs from 'fs';

interface IGuestInfo {
    name: string;
    email: string;
    picture: string;
}

function getGuestRegCacheFile() {
    return getStoreFileLoc('guestRegistrationCache.json');
}

interface IGuestRegCacheInfo {
    driveId: string;
    newGuestXlsxItemId: string;
    today: string;
}
export async function getUtil(today: string, logger: ILogger) {

    if (!today) {
        const error = 'Must define today!!!';
        logger(error);
        throw {
            message: error,
            error,
        }
    }

    let cache: IGuestRegCacheInfo = {} as IGuestRegCacheInfo;
    if (fs.existsSync(getGuestRegCacheFile())) {
        try {
            cache = JSON.parse(fs.readFileSync(getGuestRegCacheFile()).toString());
        } catch (err) {
            logger(`Cant load cache ${getGuestRegCacheFile()}`,err);
        }
    }
    function saveCache(actFor:string) {
        logger(`save Cache ${actFor}`);
        try {
            fs.writeFileSync(getGuestRegCacheFile(), JSON.stringify(cache, null, 2));
        } catch (err) {
            logger('Error save cache', err);
        }
    }

    function addPathToImg(fname: string) {
        const todayMoment = moment(today);
        const quarter = Math.floor(((todayMoment.month() + 1) % 12) / 3 + 1);
        const year = todayMoment.format('YYYY');
        if (!fname) return fname;
        return `新人资料/${year}-Q${quarter}-DBGRM/${today}/${fname}`;
    }

    const msGraphPrms: IMsGraphDirPrms = getMsDirClientPrms('https://acccnusa.sharepoint.com/:x:/r/sites/newcomer/Shared%20Documents/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99/%E6%96%B0%E4%BA%BA%E8%B5%84%E6%96%99%E8%A1%A8%E6%B1%87%E6%80%BBnew.xlsx?d=wbd57c301f851467787c3b5405709c2bf&csf=1&web=1&e=HDYYri',
        logger);
    
    
    if (msGraphPrms.driveId !== cache.driveId || !cache.driveId) {
        const oldCachedDriveId = cache.driveId;
        msGraphPrms.driveId = cache.driveId;
        if (!cache.driveId) {
            const ops = await getMsDirOpt();
            cache.driveId = msGraphPrms.driveId = ops.driveId;
        }
        saveCache(`msGraphPrms oldId=${oldCachedDriveId}, new=${msGraphPrms.driveId}`);
    }
    async function getMsDirOpt() {
        const ops = await msGraph.msdir.getMsDir(msGraphPrms);
        return ops;
    }
    const xlsPrms = {
        fileName: '新人资料/新人资料表汇总new.xlsx',
        itemId: cache.newGuestXlsxItemId,
    };
    const xlsOps = await msGraph.msExcell.getMsExcel(msGraphPrms, xlsPrms);
    if (xlsPrms.itemId !== cache.newGuestXlsxItemId) {
        const dbgRmCachedId = cache.newGuestXlsxItemId;
        cache.newGuestXlsxItemId = xlsPrms.itemId;
        saveCache(`xlsOps, newId ${xlsPrms.itemId}, old one ${dbgRmCachedId}`);
    }
    //const today = moment().format('YYYY-MM-DD');
    if (cache.today !== today) {
        await xlsOps.createSheet(today);
        
        for (let i = 0; i < 100; i++) {
            const sheets = await xlsOps.getWorkSheets();
            const found = sheets.value.find(v => v.name === today);
            logger(`Sheets (trying to find ${today}), waiting ${i * 500}`, found);
            if (found) break;
            await delay(500);
        }
        const dbgCached = cache.today;
        cache.today = today;
        saveCache(`todayChanged ${today}, cache was ${dbgCached}`);
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

    async function saveImage(fname:string, imgStr:string) {
        const sub = imgStr.indexOf('base64,');
        if (sub > 0) {
            imgStr = imgStr.substring(sub + 7).trim();
        }
        const buf = Buffer.from(imgStr, 'base64');
        const ops = await getMsDirOpt();
        const res = await ops.createFile(fname, buf);
        return res;
    }

    async function getFileByPath(fname: string) {
        const ops = await getMsDirOpt();
        const ary = await ops.getFileByPath(fname);
        return ary;
    }
    return {
        addPathToImg,
        msGraphPrms,
        getMsDirOpt,
        xlsOps,
        loadTodayData,

        saveGuest,
        saveImage,
        getFileByPath,
    }
}