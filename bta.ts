import * as fs from 'fs';
const jscfg = JSON.parse(fs.readFileSync('./.vscode/launch.json', 'utf8'));
        const envObj = jscfg.configurations.find((c: { name: string; }) => c.name === 'Run root test.ts').env
        Object.keys(envObj).forEach(k => {
            process.env[k] = envObj[k];
        });
import { sendBTAData } from './refreshEEVisitLog/lib/btaEmail';


async function test() {        
        await sendBTAData({
            date: new Date(),
            logger: s=>console.log(s),
        });
        return;
    }

test();