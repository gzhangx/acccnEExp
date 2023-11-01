
import moment from 'moment-timezone';
import * as mailTool from '../bibleSender/nodemailer'
import { flow, get, mapValues, keyBy } from 'lodash';
import * as gsSheet from './gsSheet';
async function readValues(range: string) {
  const sheet = gsSheet.getOpsBySheetId('1uYTYzwjUN8tFpeejHtiGA5u_RtOoSBO8P1b2Qg-6Elk');
  const ops = await sheet.getOps();
  const ret = await ops.readData(range);
  return ret.data;
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

async function getTemplateData(curDate: Date): Promise<TemplateData> {
  //await sheet.readSheet(spreadsheetId,range);
  const getRows = async range => (await readValues(range)).splice(1);
  
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

type ScheduleData = {
  date: string;
  diff: number;
  row: string[];
}


function getLookupOwner(owner:string) {
  const churchNumber = owner.match(/(教会){0,1}(?<number>\d+)(教室){0,1}/)?.groups?.number;
  const lookupOwnerName = churchNumber ? '教会' : owner;
  const displayOwnerHomeName = churchNumber ? `教会${churchNumber}教室` : `${owner}家`;
  return {
    lookupOwnerName,
    displayOwnerHomeName,
  }
}
function createMessage2021(templateAll: TemplateData, first: ScheduleData, have: string) {
  const getRowData = (who: number) => first.row[who] || 'NA';
  const getRowDataByLetter = (x: string | string[]) => getRowData(addrToPos(x[1] || x[0]));
  const openHomeOwner = getRowDataByLetter('F');
  const timeVal = getRowDataByLetter('B');
  const dateVal = getRowDataByLetter('A');
  const offLine = getRowDataByLetter('D');
  if (offLine.toUpperCase() === 'YES') {
    have += 'Offline'
  }
  const ownerLookup = (ownerName: string, name: string) => get(templateAll, `names[${ownerName}].${name}`, '');
  //const additionalInformation = get(templateAll, `names[${openHomeOwner}.additionalInformation`,'');  
  const map = [
    {
      name: '_time',
      val: timeVal,
    },
    {
      name: '_date',
      val: dateVal,
    },
    {
      name: '_dayOfWeek',
      val: moment(dateVal).weekday().toString(),
    },
  ];

  const template = templateAll.template[have] as SubjectText;
  const ownerColRegRes = template.text.match(/\$\{address\(([A-Z])\)}/);

  const ownerCol = (ownerColRegRes && ownerColRegRes[1]) ? ownerColRegRes[1] : null;
  const owner = ownerCol ? getRowDataByLetter(ownerCol) : '';
  
  const {  lookupOwnerName, displayOwnerHomeName } = getLookupOwner(owner);  
  

  const rpls = [
    (data: string) => map.reduce((acc, cur) => {
      return acc.replace(`\${${cur.name}}`, cur.val);
    }, data),
    (data: string) => data.replace(new RegExp('[$]{([A-Z])}', 'gi'), (...m) => {
      if (m[1] === ownerCol) {        
        return displayOwnerHomeName;
      }
      return getRowDataByLetter(m);
    }),
  ...['address', 'additionalInformation'].map(name =>
    (data: string) => data.replace(new RegExp(`[$]{${name}[(]([A-Z])[)]}`, 'gi'), (...m) => {
      //const owner = getRowDataByLetter(m);      
      const lookupInfo = ownerLookup(lookupOwnerName, name);
      if (name === 'address' && !lookupInfo) {
        return `Can't find address for ${lookupOwnerName}`;
      }
      return lookupInfo;
    })
  ),
  ];

  
  const ret = mapValues(template, flow(rpls));
  return ret;
}


type Logger = (...args) => void;
type SendSeehtNoticeParms = {
  curDateD?: Date;
  sendEmail?: 'Y' | 'N';
  logger: Logger;  
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
  logInfo(`${curDate.format('YYYY-MM-DD')} sendEmail=${opts.sendEmail}`);
  const valuesRange = ['A', 'L'];  
  const columnNames: string[] = [];
  for (let cnt = valuesRange[0].charCodeAt(0); cnt <= valuesRange[1].charCodeAt(0); cnt++) {
    columnNames.push(String.fromCharCode(cnt));
  }
  const scheduleData = await readValues(`'Schedule'!${valuesRange[0]}:${valuesRange[1]}`).then(r => {    
    return r.map(d => {
      const dateStr = d[0];
      if (!dateStr) return;
      if (!dateStr.match(/^\d\d-\d\d$/)
        && !dateStr.match(/^\d\d\d\d-\d\d-\d\d$/)) return;
      const date = dateStr.length === 5 ?
        moment(`${dateStr}`, 'MM-DD') :
        moment(`${dateStr}`, 'YYYY-MM-DD');
      if (!date.isValid()) return null;
      const ret = columnNames.reduce((acc, who) => { 
        const pos = addrToPos(who);
        acc.row[pos] = d[pos];
        return acc;
      }, {
        date,
        row: [] as string[]
        //time: d[1],
        //location: d[5],
        //who: d[6],
        //what: d[8],
      });
      //if (ret[addrToPos('B')] !== '7:30-9:30pm') return null;
      if ((ret.row[addrToPos('L')] || '').trim().toLowerCase() !== 'yes') return null;
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
  const template = await getTemplateData(opts.curDateD);
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
