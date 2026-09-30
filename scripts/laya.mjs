/** Isolated optional Laya service: npm run setup:laya, then npm run laya. */
import {spawnSync, spawn} from 'node:child_process';
import {existsSync, readFileSync, writeFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import {join} from 'node:path';
const root = fileURLToPath(new URL('..', import.meta.url));
const environment = join(root, '.venv-laya');
const bin = join(environment, process.platform === 'win32' ? 'Scripts' : 'bin');
const python = join(bin, process.platform === 'win32' ? 'python.exe' : 'python');
function run(command, args) {
    const result = spawnSync(command, args, {cwd:root, stdio:'inherit'});
    if (result.error || result.status !== 0) { console.error(result.error?.message || 'Laya setup failed.'); process.exit(result.status || 1); }
}
if (process.argv.includes('--setup')) {
    if (!existsSync(python)) run(process.env.LAYA_PYTHON || (process.platform === 'win32' ? 'python' : 'python3'), ['-m','venv',environment]);
    run(python, ['-m','pip','install','-r','requirements-laya.txt']);
    console.log('Laya installed. Run npm run laya while online once to download the multilingual checkpoint.');
} else {
    const executable = join(bin, process.platform === 'win32' ? 'laya-serve.exe' : 'laya-serve');
    if (!existsSync(executable)) { console.error('Run npm run setup:laya first (Python 3.10+ required).'); process.exit(1); }
    const appPython = join(root, 'venv', process.platform === 'win32' ? 'Scripts/python.exe' : 'bin/python');
    if (!existsSync(appPython)) { console.error('Run npm run setup first to install the local connector dependencies.'); process.exit(1); }
    const originsFile = join(root, '.laya-origins.json');
    const index = process.argv.indexOf('--origin');
    let origins = process.env.LAYA_ALLOWED_ORIGINS || '';
    if (index !== -1) {
        const input = process.argv[index + 1];
        let origin;
        try {
            const url = new URL(input);
            if (url.username || url.password || url.pathname !== '/' || url.search || url.hash || url.hostname.includes('*')
                || !(url.protocol === 'https:' || (url.protocol === 'http:' && ['localhost','127.0.0.1','[::1]'].includes(url.hostname)))) throw new Error();
            origin = url.origin;
        } catch { console.error('--origin needs your exact website origin, such as https://your-app.vercel.app (no path).'); process.exit(1); }
        origins = origin;
        writeFileSync(originsFile, JSON.stringify({origins}, null, 2) + '\n', {mode:0o600});
    } else if (!origins && existsSync(originsFile)) {
        try { origins = JSON.parse(readFileSync(originsFile, 'utf8')).origins || ''; }
        catch { console.error('Invalid .laya-origins.json. Run again with --origin to replace it.'); process.exit(1); }
    }
    const modelPort = process.env.LAYA_PORT || '8000';
    const children = [
        spawn(executable, [], {cwd:root, stdio:'inherit', env:{...process.env,
            LAYA_HOST:'127.0.0.1', LAYA_PORT:modelPort,
            LAYA_MODELS:process.env.LAYA_MODELS || 'multilingual', LAYA_PRELOAD:'1',
            ...(process.argv.includes('--offline') ? {HF_HUB_OFFLINE:'1', TRANSFORMERS_OFFLINE:'1'} : {})}}),
        spawn(appPython, ['local_laya_bridge.py'], {cwd:root, stdio:'inherit', env:{...process.env,
            LAYA_BASE_URL:`http://127.0.0.1:${modelPort}`, LAYA_ALLOWED_ORIGINS:origins}})
    ];
    let stopping = false;
    function stop(code = 0) {
        if (stopping) return;
        stopping = true; process.exitCode = code;
        for (const child of children) if (child.exitCode === null) child.kill('SIGTERM');
    }
    for (const signal of ['SIGINT','SIGTERM']) process.on(signal,()=>stop());
    for (const child of children) {
        child.on('error', error=>{console.error(error.message); stop(1);});
        child.on('exit', code=>stop(code || 0));
    }
}
