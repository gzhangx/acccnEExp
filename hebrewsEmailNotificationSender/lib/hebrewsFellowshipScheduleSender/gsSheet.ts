import *as gs from '@gzhangx/googleapi';
//import gsKeyInfo from './data/secrets/gospelCamp.json'
import * as fs from 'fs';

const clientCache ={
    client: null,
    ops: {},
} as {
    client: gs.gsAccount.IGoogleClient | null;
    ops: { [id: string]: gs.gsAccount.IGetSheetOpsReturn };
};

export type LoggerType = (...arg: any[]) => void;
async function createOps(id: string, logger: LoggerType) {
    if (!clientCache.client) {
        const gsKeyInfo: gs.gsAccount.IServiceAccountCreds = {
            client_email: process.env.GS_CLIENT_EMAIL,
            private_key: process.env.GS_PRIVATE_KEY.replace(/\\n/g, '\n'),
            private_key_id: process.env.GS_PRIVATE_KEY_ID,
        }; // = JSON.parse(fs.readFileSync('./data/secrets/gospelCamp.json').toString());
        logger(`creating ops  for ${id}`)
        const client = gs.gsAccount.getClient(gsKeyInfo);        
        clientCache.client = client;        
    }    
    if (!clientCache.ops[id]) {
        const ops = await clientCache.client.getSheetOps(id);
        clientCache.ops[id] = ops;
    }
    return clientCache.ops[id];
}

async function readDataByColumnName(id: string, name: string, logger: LoggerType) {
    const ops = await createOps(id, logger);
    return ops.readDataByColumnName(name);
}

export function getOpsBySheetId(id: string, logger: LoggerType) {
    return {        
        readDataByColumnName: (name) => readDataByColumnName(id, name, logger),        
        getOps: () => createOps(id, logger),        
    }
}