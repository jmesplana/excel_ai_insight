#!/usr/bin/env node
/**
 * Dev launcher for the Flask app, so `npm run dev` works on a fresh clone.
 *
 * Steps:
 *   1. Find a usable Python 3 interpreter.
 *   2. Create ./venv if missing.
 *   3. Install requirements.txt when it is newer than the last install stamp.
 *   4. Exec app.py with the venv's Python.
 *
 * Flags:
 *   --setup-only   Do the venv/dependency work, then exit.
 *   --no-debug     Run without the Flask reloader (production-ish local run).
 */

import { spawn, spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import process from 'node:process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const isWindows = process.platform === 'win32';
const venvDir = join(root, 'venv');
const venvPython = join(venvDir, isWindows ? 'Scripts' : 'bin', isWindows ? 'python.exe' : 'python');
const requirements = join(root, 'requirements.txt');
const stampFile = join(venvDir, '.requirements.stamp');

const args = new Set(process.argv.slice(2));
const setupOnly = args.has('--setup-only');
const noDebug = args.has('--no-debug');

const color = (code, text) => (process.stdout.isTTY ? `\x1b[${code}m${text}\x1b[0m` : text);
const info = (msg) => console.log(`${color('36', '›')} ${msg}`);
const ok = (msg) => console.log(`${color('32', '✓')} ${msg}`);
const fail = (msg) => console.error(`${color('31', '✗')} ${msg}`);

// Some pinned dependencies (jiter, pydantic_core, numpy, pandas) ship wheels
// only for released Python versions. On a very new interpreter pip falls back
// to building from source, which needs a Rust/C toolchain and usually fails.
// Prefer an interpreter we know has wheels, newest first, before falling back
// to whatever `python3` happens to be.
const PREFERRED_PYTHONS = [
  'python3.12',
  'python3.11',
  'python3.10',
  'python3.9',
  'python3',
  'python',
];

// Highest Python minor version known to have wheels for our pinned deps.
const MAX_SUPPORTED_MINOR = 12;

/** Return [major, minor] for an interpreter, or null if it is not usable. */
function pythonVersion(candidate) {
  const probe = spawnSync(
    candidate,
    ['-c', 'import sys; print("%d.%d" % sys.version_info[:2])'],
    { encoding: 'utf8' },
  );
  if (probe.status !== 0) return null;
  const parts = probe.stdout.trim().split('.').map(Number);
  if (parts.length !== 2 || Number.isNaN(parts[0])) return null;
  return parts;
}

/**
 * Locate the best Python 3 interpreter on PATH.
 * Returns { cmd, version, risky } — risky flags a version that is newer than
 * what our pinned requirements have wheels for.
 */
function findSystemPython() {
  const seen = new Set();
  let fallback = null;

  for (const candidate of PREFERRED_PYTHONS) {
    const version = pythonVersion(candidate);
    if (!version || version[0] !== 3) continue;
    const key = version.join('.');
    if (seen.has(key)) continue;
    seen.add(key);

    if (version[1] <= MAX_SUPPORTED_MINOR) {
      return { cmd: candidate, version: key, risky: false };
    }
    if (!fallback) fallback = { cmd: candidate, version: key, risky: true };
  }
  return fallback;
}

/** requirements.txt is newer than the stamp (or was never installed). */
function requirementsAreStale() {
  if (!existsSync(stampFile)) return true;
  try {
    return statSync(requirements).mtimeMs > statSync(stampFile).mtimeMs;
  } catch {
    return true;
  }
}

function run(cmd, cmdArgs, label, hint) {
  const result = spawnSync(cmd, cmdArgs, { stdio: 'inherit', cwd: root });
  if (result.status !== 0) {
    fail(`${label} failed (exit code ${result.status}).`);
    if (hint) info(hint);
    process.exit(result.status ?? 1);
  }
}

// --- 1. venv ---------------------------------------------------------------
if (!existsSync(venvPython)) {
  const systemPython = findSystemPython();
  if (!systemPython) {
    fail('No Python 3 interpreter found on PATH. Install Python 3 and try again.');
    process.exit(1);
  }
  if (systemPython.risky) {
    info(
      `Only Python ${systemPython.version} was found. Some pinned dependencies ` +
        `have no prebuilt wheels for it and may fail to compile.`,
    );
    info(`If the install fails, install Python 3.12 or 3.11 and re-run.`);
  }
  info(`Creating virtual environment in ./venv (Python ${systemPython.version}) ...`);
  run(systemPython.cmd, ['-m', 'venv', 'venv'], 'venv creation');
  ok('Virtual environment created.');
}

// --- 2. dependencies -------------------------------------------------------
if (requirementsAreStale()) {
  info('Installing Python dependencies (this runs only when requirements.txt changes) ...');
  run(venvPython, ['-m', 'pip', 'install', '--upgrade', 'pip', '--quiet'], 'pip upgrade');
  run(
    venvPython,
    ['-m', 'pip', 'install', '-r', 'requirements.txt', '--quiet'],
    'dependency install',
    'If wheels failed to build, your Python version is likely too new for the ' +
      'pinned versions. Install Python 3.12 or 3.11, remove ./venv, and re-run.',
  );
  writeFileSync(stampFile, new Date().toISOString());
  ok('Dependencies installed.');
}

if (setupOnly) {
  ok('Setup complete. Run `npm run dev` to start the app.');
  process.exit(0);
}

// --- 3. .env hint ----------------------------------------------------------
const envFile = join(root, '.env');
if (!existsSync(envFile) && existsSync(join(root, '.env.example'))) {
  info('No .env found. Copy .env.example to .env to configure a server-side model.');
} else if (existsSync(envFile)) {
  const contents = readFileSync(envFile, 'utf8');
  if (/^\s*AZURE_OPENAI_ENDPOINT=\S/m.test(contents)) {
    ok('Azure AI Foundry configuration detected in .env');
  } else if (/^\s*OPENAI_API_KEY=\S/m.test(contents)) {
    ok('OpenAI configuration detected in .env');
  }
}

// --- 4. run ----------------------------------------------------------------
const port = process.env.PORT || '8080';
const host = process.env.HOST || '127.0.0.1';
ok(`Starting Aidstack Insights on http://${host}:${port}`);
info('Press Ctrl+C to stop.\n');

const child = spawn(venvPython, ['app.py'], {
  cwd: root,
  stdio: 'inherit',
  env: {
    ...process.env,
    PORT: port,
    HOST: host,
    FLASK_DEBUG: noDebug ? '0' : '1',
    PYTHONUNBUFFERED: '1',
  },
});

// Forward termination signals so Ctrl+C stops Flask cleanly.
for (const signal of ['SIGINT', 'SIGTERM']) {
  process.on(signal, () => child.kill(signal));
}
child.on('exit', (code) => process.exit(code ?? 0));
