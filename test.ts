import { loadEnvFromLaunchConfig } from './hebrewsEmailNotificationSender/lib/hebrewsFellowshipScheduleSender/eng_util';
loadEnvFromLaunchConfig();
import * as fs from 'fs';

import * as dailySender from './localMissionExpenseRecorder/bibleSender/getdata'
//import * as sendWeek from './localMissionExpenseRecorder/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail'
import * as sendWeek from './hebrewsEmailNotificationSender/lib/hebrewsFellowshipScheduleSender/sendHebrewsWeeklyEmail'
import { sendBTAData } from './refreshEEVisitLog/lib/btaEmail';
import moment from 'moment';
import * as gs from '@gzhangx/googleapi'
import { sum } from 'lodash';
import { creatOptsFromEnv } from './hebrewsEmailNotificationSender/lib/hebrewsFellowshipScheduleSender/gsSheet';


async function test(retFirst: string) {

    if (retFirst === "sendBtaEmail") {        
        const logger = (msg: any) => console.log(msg);
        
        //newGuestRegTest();
        
        await sendBTAData({
            date: new Date(),
            logger: s=>console.log(s),
        });
        return;
    }
    if (retFirst === 'sendSheetNotice') {

        
        const test = await sendWeek.sendSheetNotice({
            logger: console.log,
            sendEmail: 'Y',
        });
        return console.log(test);
    }
    const got = await dailySender.sendBibleData({
        logger: console.log,
        sendEmail: 'N',
    });
    console.log(got);
}


async function populateChurchEvent() {
    const client = creatOptsFromEnv();
    const hops = await client.getSheetOps(process.env.HEBREWS_SHEET_ID || '');
    const allData = [];
    for (let mon = 1; mon <= 12; mon++) {
        const monStr = moment('2026-01-01').add(mon - 1, 'months').format('YYYY-MM-DD');        
        const nextMonStr = moment('2026-01-01').add(mon, 'months').format('YYYY-MM-DD');
        console.log(monStr, nextMonStr);            
        const url = `https://clients6.google.com/calendar/v3/calendars/9ee668q0j6mmaum65303u7r194%40group.calendar.google.com/events?calendarId=9ee668q0j6mmaum65303u7r194%40group.calendar.google.com&singleEvents=true&eventTypes=default&eventTypes=focusTime&eventTypes=outOfOffice&timeZone=America/New_York&maxAttendees=1&maxResults=250&sanitizeHtml=true&timeMin=${monStr}T00%3A00%3A00%2B18%3A00&timeMax=${nextMonStr}T00%3A00%3A00-18%3A00&key=AIzaSyDOtGM5jr8bNp1utVpG2_gSRH03RNGBkI8&$unique=gc456`;
        const calRes = await gs.util.doHttpRequest({
            //url: ' https://clients6.google.com/calendar/v3/calendars/9ee668q0j6mmaum65303u7r194%40group.calendar.google.com/events?calendarId=9ee668q0j6mmaum65303u7r194%40group.calendar.google.com&singleEvents=true&eventTypes=default&eventTypes=focusTime&eventTypes=outOfOffice&timeZone=America/New_York&maxAttendees=1&maxResults=250&sanitizeHtml=true&timeMin=2026-01-01T00:00:00+18:00&timeMax=2026-12-31T00:00:00-18:00&key=AIzaSyDOtGM5jr8bNp1utVpG2_gSRH03RNGBkI8&$unique=gc456',            
            url,
            method: 'GET',
        });
        //console.log('Number items', (calRes.data as any).items.length, url);
        const descs = (calRes.data as any).items.map(itm=>{
            return {
                summary: itm.summary,
                //desc: itm.description,
                start: itm.start.date,
                end: itm.end.date,
            }
        }).filter(d=>(d.start || '').includes('2026')).map(d=>{
            const weekdays = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
            const [year, month, day] = d.start.split('-').map(Number);
            const dayOfWeek = weekdays[new Date(year, month - 1, day).getDay()];
            
            let time = '';
            let chineseText = '';
            
            // Check if tab-separated format
            if (d.summary.includes('\t')) {
                const parts = d.summary.split('\t');
                time = parts[0] || '';
                chineseText = parts[1] || '';
            } else {
                // Parse format: "12:45PM - 1:30PM 團契主席月會Fellowship Group Leaders Meeting"
                const match = d.summary.match(/^(.*?)([一-龥]+)(.*)?$/);
                time = match ? match[1].trim() : d.summary;
                chineseText = match ? match[2].trim() : '';
            }
            
            return {
                start: d.start,
                weekday: dayOfWeek,
                time: time,
                summary: chineseText,
                oldSummary: d.summary,
            }
        }); //.filter(d=>d.summary.includes('預查') || d.summary.includes('預查'));        
        descs.forEach(d=>{
            const data = [d.start, d.weekday, d.time, d.summary, d.oldSummary];
            console.log('data', data)
            allData.push(data)
        });
    }
    console.log(allData);
    await hops.append('北堂2026年历',allData.filter(d=>d[4].includes('預查')));
}

test('sendSheetNotice');

//populateChurchEvent();