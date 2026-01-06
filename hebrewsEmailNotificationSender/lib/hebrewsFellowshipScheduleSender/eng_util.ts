import * as fs from 'fs';

export function loadEnvFromLaunchConfig() {
const jscfg = JSON.parse(fs.readFileSync('./.vscode/launch.json', 'utf8'));
        const envObj = jscfg.configurations.find((c: { name: string; }) => c.name === 'Run root test.ts').env
        Object.keys(envObj).forEach(k => {
            process.env[k] = envObj[k];
        });
}