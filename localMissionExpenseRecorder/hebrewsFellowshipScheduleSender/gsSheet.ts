import *as gs from '@gzhangx/googleapi';
import gsKeyInfo from './data/secrets/gospelCamp.json'

const clientCache ={
    client: null,
    ops: {},
} as {
    client: gs.gsAccount.IGoogleClient | null;
    ops: { [id: string]: gs.gsAccount.IGetSheetOpsReturn };
};

async function createOps(id: string) {
    if (!clientCache.client) {
        const client = gs.gsAccount.getClient(gsKeyInfo);        
        clientCache.client = client;        
    }
    if (!clientCache.ops[id]) {
        const ops = await clientCache.client.getSheetOps(id);
        clientCache.ops[id] = ops;
    }
    return clientCache.ops[id];
}

async function readDataByColumnName(id: string, name: string) {
    const ops = await createOps(id);
    return ops.readDataByColumnName(name);
}

export function getOpsBySheetId(id: string) {
    return {        
        readDataByColumnName: (name) => readDataByColumnName(id, name),        
        getOps: () => createOps(id),        
    }
}