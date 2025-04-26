
import moment from 'moment-timezone';
import * as mailTool from '../bibleSender/nodemailer'
import * as lodash from 'lodash';
const { flow, get, mapValues, keyBy } = lodash;
import * as gsSheet from './gsSheet';
import { LoggerType } from './gsSheet';
async function readValues(range: string, logger: LoggerType) {
  const sheet = gsSheet.getOpsBySheetId('1uYTYzwjUN8tFpeejHtiGA5u_RtOoSBO8P1b2Qg-6Elk', logger);
  const ops = await sheet.getOps();
  const ret = await ops.readData(range);
  //return ret.data;
  return ret.values;
}

const addrToPos = (x:string) => x.charCodeAt(0) - 65;
//return test();

//given dayOfweek, get the next day of week.
//function getNextDayOfWeek(date: Date, dayOfWeek:number) {
//  const res = new Date(date.getTime());
//  res.setDate(date.getDate() + (7 + dayOfWeek - date.getDay()) % 7);
//  return res;
//}


type SubjectText = {
  subject: string;
  text: string;
}

type TemplateData = {
  names: {
    [name: string]: {
      name: string;
      address: string;
      additionalInformation: string;
    }
  };
  curDate: Date;
  toAddrs: string[];
  template: {
    have: SubjectText;
    havenot: SubjectText;
    haveOffline: SubjectText;
    havenotOffline: SubjectText;
  }
}

async function getTemplateData(curDate: Date, logger: LoggerType): Promise<TemplateData> {
  //await sheet.readSheet(spreadsheetId,range);
  const getRows = async range => (await readValues(range, logger)).splice(1);
  
  const rows = await getRows('Addresses');  
  const names = keyBy(rows.map(r=>({name: r[0], address:r[1], additionalInformation:r[2]})).filter(x=>x.name), 'name');
  const templateRows = await getRows('Template');  
  const toAddrs: string[] = [];
  for (let i = 0; i < templateRows.length; i++) {
    const em = templateRows[i][2];
    if (em) {
      toAddrs.push(em);
    }
  }
  console.log(toAddrs);
  return {
    names,
    curDate,
    toAddrs,
    template: {
      have: {
        subject: templateRows[0][0],
        text: templateRows[1][0],
      },
      havenot: {
        subject: templateRows[0][1],
        text: templateRows[1][1],
      },
      haveOffline: {
        subject: templateRows[0][4],
        text: templateRows[1][4],
      },
      havenotOffline: {
        subject: templateRows[0][5],
        text: templateRows[1][5],
      }
    }
  }
}

type StringDict = {
  [name: string]: string;
};

type ScheduleData = {
  date: string;
  diff: number;
  row: StringDict;
}


function getLookupOwner(owner:string) {
  const churchNumber = owner.match(/(教会){0,1}(?<number>\d+)(教室){0,1}/)?.groups?.number;
  const lookupOwnerName = churchNumber ? '教会' : owner;
  let displayOwnerHomeName = churchNumber ? `教会${churchNumber}教室` : `${owner}家`;
  if (owner.toUpperCase() === 'ACCCN') {
    displayOwnerHomeName = 'ACCCN';
  }
  return {
    lookupOwnerName,
    displayOwnerHomeName,
  }
}
function createMessage2021(templateAll: TemplateData, first: ScheduleData, have: string) {
  const getRowData = (who: string) => first.row[who] || 'NA';
  
  
  const dateVal = getRowData('日期');
  const offLine = getRowData('聚会地点');
  if (offLine.toUpperCase() === 'YES') {
    have += 'Offline'
  }
  const ownerLookup = (ownerName: string, name: string) => get(templateAll, `names[${ownerName}].${name}`, '');
  //const additionalInformation = get(templateAll, `names[${openHomeOwner}.additionalInformation`,'');  
  first.row['_dayOfWeek'] = ['日', '一', '二', '三', '四', '五', '六'][moment(dateVal).weekday()];    

  const template = templateAll.template[have] as SubjectText;
  const ownerColRegRes = template.text.match(/\$\{address\((.+)\)}/);

  const ownerColFromAddr = (ownerColRegRes && ownerColRegRes[1]) ? ownerColRegRes[1] : null;
  const ownerCol = '聚会地点'
  if (ownerColFromAddr !== ownerCol && ownerColFromAddr) {
    console.log(`Warning, owner col ${ownerColFromAddr} is not ${ownerCol}`);
  }
  const owner = getRowData(ownerCol);
  
  const {  lookupOwnerName, displayOwnerHomeName } = getLookupOwner(owner);  
  first.row['聚会地点'] = displayOwnerHomeName;
  const lookupInfo = ownerLookup(lookupOwnerName, 'address');
  first.row['address'] = lookupInfo ? lookupInfo : `Can't find address for ${lookupOwnerName}`;
  const additionalInformation = ownerLookup(lookupOwnerName, 'additionalInformation');
  first.row['additionalInformation'] = additionalInformation || '';
  first.row['year'] = moment(dateVal).format('YYYY');

  const rpls = [
    (data: string) => data.replace(new RegExp('[$]{([^{}]+)}', 'gi'), (...m) => {    
      return getRowData(m[1]);
    }),
  ];

  
  const ret = mapValues(template, flow(rpls));
  return ret;
}


type SendSeehtNoticeParms = {
  curDateD?: Date;
  sendEmail?: 'Y' | 'N';
  logger: LoggerType;  
}

export async function sendSheetNotice(opts: SendSeehtNoticeParms) {
  const steps: string[] = [];
  try {
    const res = await sendSheetNoticeInner(opts, steps);
    return res;
  } catch (err) {
    return {
      steps,
      error: err.message,
    }
  }
}

async function sendSheetNoticeInner(opts: SendSeehtNoticeParms, steps: string[]) {
  const logInfo = (...args) => {
    steps.push(args[0]);
    opts.logger(...args);
  }
  if (!opts.curDateD) opts.curDateD = new Date();
  logInfo(`${opts.curDateD} sendEmail=${opts.sendEmail}`);
  const curDate = moment(opts.curDateD).startOf('day');
  const sheetName = `${curDate.format('YYYY') }行事历`
  logInfo(`${sheetName} ${curDate.format('YYYY-MM-DD')} sendEmail=${opts.sendEmail}`);
  const valuesRange = ['A', 'L'];  
  
  //for (let cnt = valuesRange[0].charCodeAt(0); cnt <= valuesRange[1].charCodeAt(0); cnt++) {
  //  columnNames.push(String.fromCharCode(cnt));
  //}
  logInfo(`${sheetName} ${curDate.format('YYYY-MM-DD')} sendEmail=${opts.sendEmail} here`);
  const scheduleData = await readValues(`'${sheetName}'!${valuesRange[0]}:${valuesRange[1]}`, logInfo).then(r => {    
    const columnNames: string[] = r[1].map(r => r);
    return r.map(curRow => {
      const cur = columnNames.reduce((acc, name, i) => { 
        acc[name] = curRow[i];
        return acc;
      }, {} as StringDict);
      const dateStr = cur['日期'];
      if (!dateStr) return;
      if (!dateStr.match(/^\d\d-\d\d$/)
        && !dateStr.match(/^\d\d\d\d-\d\d-\d\d$/)) return;
      const date = dateStr.length === 5 ?
        moment(`${dateStr}`, 'MM-DD') :
        moment(`${dateStr}`, 'YYYY-MM-DD');
      if (!date.isValid()) return null;
      const ret = {
        date,
        row: cur       
      };
      //if (ret[addrToPos('B')] !== '7:30-9:30pm') return null;
      if ((ret.row['Send Notice'] || '').trim().toLowerCase() !== 'yes') return null;
      return ret;
    }).filter(x => x);
  });

  logInfo(`scheduleData len=${scheduleData.length}`);
  const found = scheduleData.reduce((acc, d) => {
    if (acc) return acc;
    if (d.date.isSameOrAfter(curDate)) return {
      diff: d.date.diff((curDate), 'days'),
      date: d.date.format('YYYY-MM-DD'),
      row: d.row,
    };
  }, null) as ScheduleData;  
  //console.log(found);  

  const message = {
    from: `"HebrewsBot" <${mailTool.emailUser}>`,
    //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
    to: '',
    subject: 'NA',
    text: 'NA',
  };
  const template = await getTemplateData(opts.curDateD, logInfo);
  message.to = template.toAddrs.join(',');
  const messageData = createMessage2021(template, found, found.diff < 7 ? 'have' : 'havenot');
  Object.assign(message, messageData);
  logInfo(message.to);
  logInfo(message.subject);
  logInfo(message.text);  
  
  if (opts.sendEmail === 'N') return {
    message,
    steps,
  };
  const mailResult = await mailTool.emailTransporter.sendMail(message);
  //email.sendGmail(message);
  return {
    message,
    steps,
    mailResult,
  };
}

