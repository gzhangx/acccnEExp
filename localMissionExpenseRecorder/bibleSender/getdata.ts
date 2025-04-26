// Warning, moved to hebrewsEmailNotificationSender.ts
'use strict';

import * as moment from 'moment-timezone';

import { emailTransporter, emailUser } from './nodemailer';

//const DEFAULT_SCHEDULE_FILE = `${__dirname}/data/schedule.txt`;
//const BIBLE_TEXT_FILE = `${__dirname}/data/bibleUTF8.txt`;
import * as bibleData from './data/bibleData';
import * as schedule from './data/schedule';

//const schedule = require('./schedule');
//const dailyparts = require('./dailyparts');

function ParseLineData(data: string[]) {
    const lineData: string[] = [];
    for (let i = 0; i < 8; i++) lineData[i] = '';
    let who = 0;
    let prevChar = ' ';
    for (let i = 0; i < data.length; i++) {
        const d = data[i];
        if (who < lineData.length) lineData[who] = d;
        if (who > 0) {
            const c = d[0];
            if (c >= '0' && c <= '9') {
                lineData[who - 1] += d;
                continue;
            }
            if (c === ',') {
                lineData[who - 1] += c;
                prevChar = c;
                continue;
            }
            if (prevChar === ',') {
                lineData[who - 1] += c;
            }
            prevChar = c;
        }
        who++;
    }
    return lineData;
}


type JsonSchedule = {
    startDate: string;
    schedule: string[][];
    verses: {
        [name: string]: {
            pos: number;
        }
    }
}

//scheduleTxt = DEFAULT_SCHEDULE_FILE
function ScheduleToJson(scheduleLines?: string[]) : JsonSchedule {
    //console.log(`using  schedule ${scheduleTxt}`);
    //const data = fs.readFileSync(scheduleTxt, 'utf8').toString();
    //const lines = data.split('\n');
    //fs.writeFileSync('schedule.json', JSON.stringify(lines));
    const lines = scheduleLines || schedule.schedule;
    const startDate = moment.tz(lines[0].trim(),'EST').format('YYYY-MM-DD');
    const res: JsonSchedule = {startDate, schedule:[], verses: {}};
    let pos = 0;
    for (let i = 1; i < lines.length; i++) {
        const curLine = lines[i];
        const lineData = ParseLineData(curLine.split(/[\s\t]+/));
        res.schedule.push(lineData);
        //console.log('reading line '+ i + ' of ' + lines.length+ ' ' + curLine);
        for (let li in lineData) {
            if (parseInt(li) === 0) continue;
            res.verses[lineData[li]] = { pos: pos++};
        }
    }

    //const fcnt = JSON.stringify(res);
    //fs.writeFileSync('schedule.json', fcnt);
    return res;
}

function getTodayData(today: moment.Moment, scheduleData: JsonSchedule) {
    //const bibleData = fs.readFileSync(BIBLE_TEXT_FILE, 'utf8').toString().split('\n');
    //fs.writeFileSync('bibleData.json', JSON.stringify(bibleData));
   const searches = GetTodaysSearch(today, scheduleData);
    const data = GetOutput(bibleData.bibleData, searches.Verses);
    const ret = {
        subject: searches.subject.trim().replace(/ /g, ''),
        data
    };
    return ret;
}

function getTodayEST() {
    return moment.tz('EST');
}
function getDaysOffset(today: moment.Moment, scheduleData: JsonSchedule) {
    if (!today) today = getTodayEST();
    today.add(2, 'hours'); //for dst
    return today.diff(moment.tz(scheduleData.startDate,'EST'), 'days')%728;
}

type WeekData = {
    days: number; //day to use (%728)
    week: string[]; //the schedule week line
    day: number;    //days/days_per_line
    subject: string;
}
function getWeek(today: moment.Moment, scheduleData: JsonSchedule): WeekData{
    const DAYS_PER_LINE = 7;
    const lines = scheduleData.schedule;
    const days = getDaysOffset(today, scheduleData);
    const curLineDay = Math.floor(days / DAYS_PER_LINE);
    const lineData = lines[curLineDay];
    const day = days % DAYS_PER_LINE;
    const curdata1 = lineData[day + 1];
    return {
        days,
        week: lineData,
        day,
        subject: curdata1
    }
}


type TodaySearchResult = {
    days: number;
    subject: string;
    Verses: {
        Verse: string;
        Part?: number;
        Total?: number;
    }[];
}
function GetTodaysSearch(today: moment.Moment, scheduleData: JsonSchedule) {
    const week = getWeek(today, scheduleData);

    const retResult: TodaySearchResult = {
        days: week.days,
        Verses: [],
        subject: '',
    };
    const results = retResult.Verses;

    retResult.subject = week.subject;
    const curdataparts = week.subject.split(/[,]+/);

    for (let curdataii in curdataparts) {
        const curdata = curdataparts[curdataii];
        let numStart = 0;
        for (; numStart < curdata.length; numStart++) {
            if (!isNaN(parseInt(curdata[numStart]))) {
                break;
            }
        }
        const bookName = curdata.substring(0, numStart);

        let numbers = curdata.substring(numStart);

        //formats: book#-#
        //         book#:#-#
        //         book #(#/#)
        if (numbers.indexOf(":") > 0) {
            const verse = numbers.substring(0, numbers.indexOf(":"));
            numbers = numbers.substring(numbers.indexOf(":") + 1);
            if (numbers.indexOf("-") > 0) {
                const numberary = numbers.split('-');
                const fromVer = parseInt(numberary[0]);
                let toVer = parseInt(numberary[1]);
                for (let num = fromVer; num <= toVer; num++) {
                    results.push({Verse: bookName + verse + ":" + num + " "});
                }
            }
            else {
                results.push({Verse: curdata});
            }

        } else if (numbers.indexOf("-") > 0) {
            const numberary = numbers.split('-');
            const fromVer = parseInt(numberary[0]);
            const toVer = parseInt(numberary[1]);
            for (let num = fromVer; num <= toVer; num++) {
                results.push({Verse: bookName + num});
            }
        }
        else if (numbers.indexOf("(") > 0) {
            const chapterN = numbers.indexOf("(");
            const partialStr = numbers.substring(chapterN);
            const matches = partialStr.match(/\((\d+)\/(\d+)\)/)
            const pt =
            {
                Verse: bookName + numbers.substring(0, chapterN),
                Part: parseInt(matches[1]),
                Total: parseInt(matches[2]),
            };
            
            //const startNTotal = partialStr.split(/[\(/\)]/);
            //pt.Part = parseInt(startNTotal[0]);
            //pt.Total = parseInt(startNTotal[1]);
            results.push(pt);
        }
        else {
            results.push({Verse: curdata});
        }
    }

    for (const rii in results) {
        const r = results[rii];
        if (r.Verse.indexOf(":") >= 0) continue;
        const lastChar = r.Verse[r.Verse.length - 1];
        if (lastChar >= '0' && lastChar <= '9')
            r.Verse += ":";
    }
    return retResult;
}


function GetOutput(all, shows) {
    let sb = '';
    for (let showi in shows) {
        const show = shows[showi];
        const result = [];
        for (let ti in all) {
            const t = all[ti];
            if (t.startsWith(show.Verse)) {
                result.push(t);
            }
        }
        let startLimit = 0;
        let endLimit = result.length;
        if (show.Part !== 0) {
            startLimit = (show.Part - 1) * result.length / show.Total;
            endLimit = (show.Part) * result.length / show.Total;
            if (show.Part === result.length) endLimit++;
        }
        for (let i = 0; i < result.length; i++) {
            if (i < startLimit) continue;
            if (i >= endLimit) continue;
            const t = result[i];
            sb += (t) + ("\r\n");
        }
    }
    return sb;
}

function loadData(today: moment.Moment, scheduleLines?:string[]) {
    const scheduleData = ScheduleToJson(scheduleLines);
    return getTodayData(today, scheduleData)

}

//var resss = loadData(new Date());
//console.log(resss.Data);


type SendBibleDataInput = {
    now?: moment.Moment;
    sendEmail?: 'Y' | 'N';
    logger: (...args) => void;
}
export async function sendBibleData(opts: SendBibleDataInput)
{    
    if (!opts.now) opts.now = getTodayEST();
    if (!opts.logger) opts.logger = console.log;
    const data = loadData(opts.now);
    const message = {
        from: `"Hebrews Daily Bible verse" <${emailUser}>`,
        to: ['hebrewsofacccn@googlegroups.com'],  
        subject: data.subject + ', ' + opts.now.format('YYYY-MM-DD'),
        text: data.data,
    };
    
    //console.log('sending ' + message.subject);    
    if (opts.sendEmail !== 'N') {        
        const sendEmailRes = await emailTransporter.sendMail(message).catch(err => {            
            opts.logger('Error send email', err);
            return {
                error: err.message,
            }
        });
        opts.logger('SendEmailRes', sendEmailRes);
        return sendEmailRes;
    }
    return message;
}

//ScheduleToJson();
//init_createAllData();
/*
for (var i = 0; i < 1000; i++) {
    var now = new Date();
    now.setDate(now.getDate()+i);
    console.log(now+"\r\n");
    var data = loadData(now);
    console.log(data.Subject+"\r\n");
    console.log('data='+data.Data+"\r\n");
}
/* */
//for (let i = 0; i < 100;i++) {
//    console.log(getWeek(moment().add(i, 'days')));
//    console.log(SendEmail(moment().add(i, 'days')));
//    console.log(loadData(moment().add(i, 'days')).Data);
//}

//getd.SendEmail();

