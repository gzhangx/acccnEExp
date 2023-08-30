
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
async function test(days = 0) {
  const start = moment();
  for(let i = 0; i < 1;i++) {
    await checkSheetNotice(start.add(days, 'days').toDate(), true);
  }
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

export function createMessage2021(templateAll: TemplateData, first: ScheduleData, have: string) {
  const getRowData = (who: number) => first.row[who] || 'NA';
  const getRowDataByLetter = (x: string | string[]) => getRowData(addrToPos(x[1] || x[0]));
  const openHomeOwner = getRowDataByLetter('F');
  const timeVal = getRowDataByLetter('B');
  const dateVal = getRowDataByLetter('A');
  const offLine = getRowDataByLetter('D');
  if (offLine.toUpperCase() === 'YES') {
    have += 'Offline'
  }
  const ownerLookup = (openHomeowner: string, name: string) => get(templateAll, `names[${openHomeOwner}].${name}`, '');
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
      val: templateAll.curDate.getDay().toString(),
    },
  ];

    
  const rpls = [
    (data: string) => map.reduce((acc, cur) => {
      return acc.replace(`\${${cur.name}}`, cur.val);
    }, data),
    (data: string) => data.replace(new RegExp('[$]{([A-Z])}', 'gi'), (...m) => {
      return getRowDataByLetter(m);
    }),
  ...['address', 'additionalInformation'].map(name =>
    (data: string) => data.replace(new RegExp(`[$]{${name}[(]([A-Z])[)]}`, 'gi'), (...m) => {
      if (name === 'address') {
        if (m[1].indexOf('教会') >= 0) return m[1];
        if (m[1].match(/\d/)) {
          return `教会${m[1]}教室`
        }
      }
      const owner = getRowDataByLetter(m);
      const ret = ownerLookup(owner, name);
      if (name === 'address' && !ret) {
        return `Can't find address for ${openHomeOwner}`;
      }
      return ret;
    })
  ),
  ];

  const template = templateAll.template[have] as SubjectText;
  const ret = mapValues(template, flow(rpls));
  return ret;
}

export async function checkSheetNotice(curDateD = new Date(), sendEmail = true) {
  console.log(`${curDateD} sendEmail=${sendEmail}`);
  const curDate = moment(curDateD).startOf('day');
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
    from: '"HebrewsBot" <gzhangx@gmail.com>',
    //to: 'hebrewsofacccn@googlegroups.com',  //nodemailer settings, not used here
    to: '',
    subject: 'NA',
    text: 'NA',
  };
  const template = await getTemplateData(curDateD);
  message.to = template.toAddrs.join(',');
  const messageData = createMessage2021(template, found, found.diff < 7 ? 'have' : 'havenot');
  Object.assign(message, messageData);
  console.log(message.to);
  console.log(message.subject);
  console.log(message.text);  
  
  if (!sendEmail) return message;
  //email.sendGmail(message);
  return message;
}

