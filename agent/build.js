'use strict';

/**
 * Builds LocaLinter-Agent.exe — the agent as one file a tester can double-click.
 *
 * A tester should not have to install Node, clone a repo, or open a terminal to
 * scan a device, so the whole agent (Node runtime included) is folded into a
 * single Windows executable using Node's Single Executable Application support:
 *
 *   1. esbuild flattens server.js and its dependencies into one CommonJS file.
 *   2. The route maps under routes/ are collected into an asset the executable
 *      carries, and seeds into the data directory on first run.
 *   3. `node --experimental-sea-config` turns that into a blob.
 *   4. The blob is injected into a copy of this machine's node.exe.
 *
 * Run it with `npm run build:exe` from agent/. Output: agent/dist/.
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const AGENT_DIR = __dirname;
const BUILD_DIR = path.join(AGENT_DIR, 'build');
const DIST_DIR = path.join(AGENT_DIR, 'dist');
const EXE_NAME = 'LocaLinter-Agent.exe';

// Node's own sentinel. postject looks for this string inside the binary to know
// where the blob may go; it is not something we get to choose.
const SENTINEL = 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2';

function run(cmd, args, opts = {}) {
  return execFileSync(cmd, args, { stdio: 'inherit', cwd: AGENT_DIR, shell: process.platform === 'win32', ...opts });
}

function step(n, msg) {
  console.log(`\n[${n}/5] ${msg}`);
}

function main() {
  if (process.platform !== 'win32') {
    // Cross-building would mean fetching another platform's node binary; this
    // script only ever injects the runtime it is standing on.
    console.error(`This build produces a Windows .exe and must run on Windows (this is ${process.platform}).`);
    process.exit(1);
  }

  fs.rmSync(BUILD_DIR, { recursive: true, force: true });
  fs.mkdirSync(BUILD_DIR, { recursive: true });
  fs.mkdirSync(DIST_DIR, { recursive: true });

  step(1, 'Bundling the agent into one file…');
  const bundle = path.join(BUILD_DIR, 'agent.cjs');
  run('npx', [
    '--yes', 'esbuild', 'server.js',
    '--bundle',
    '--platform=node',
    '--target=node20',
    '--format=cjs',
    `--outfile=${bundle}`,
    // Keeps stack traces from the shipped binary pointing at real code.
    '--sourcemap=inline',
  ]);

  step(2, 'Collecting route maps…');
  const seed = {};
  const routesDir = path.join(AGENT_DIR, 'routes');
  for (const file of fs.existsSync(routesDir) ? fs.readdirSync(routesDir) : []) {
    if (!file.endsWith('.json')) continue;
    try {
      seed[file] = JSON.parse(fs.readFileSync(path.join(routesDir, file), 'utf8'));
      console.log(`      + ${file}`);
    } catch (e) {
      console.warn(`      ! skipping ${file}: ${e.message}`);
    }
  }
  const seedPath = path.join(BUILD_DIR, 'routes-seed.json');
  fs.writeFileSync(seedPath, JSON.stringify(seed), 'utf8');

  step(3, 'Writing the SEA config…');
  const seaConfig = path.join(BUILD_DIR, 'sea-config.json');
  fs.writeFileSync(seaConfig, JSON.stringify({
    main: bundle,
    output: path.join(BUILD_DIR, 'agent.blob'),
    disableExperimentalSEAWarning: true,
    // Node resolves asset paths relative to the config file's directory.
    assets: { 'routes-seed.json': seedPath },
  }, null, 2), 'utf8');

  step(4, 'Generating the blob…');
  run(process.execPath, ['--experimental-sea-config', seaConfig], { shell: false });

  step(5, 'Injecting it into a copy of node.exe…');
  const exe = path.join(DIST_DIR, EXE_NAME);
  fs.copyFileSync(process.execPath, exe);
  run('npx', ['--yes', 'postject', exe, 'NODE_SEA_BLOB', path.join(BUILD_DIR, 'agent.blob'),
    '--sentinel-fuse', SENTINEL]);

  const mb = (fs.statSync(exe).size / 1024 / 1024).toFixed(1);
  console.log(`\nBuilt ${path.relative(AGENT_DIR, exe)} (${mb} MB)`);
  console.log('Double-click it, then open LocaLinter and go to the Device Scan tab.');
}

main();
