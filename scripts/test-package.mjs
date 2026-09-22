import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, mkdtemp, symlink, rm } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { resolve, join, dirname } from 'node:path';
import { createHash, webcrypto } from 'node:crypto';
import vm from 'node:vm';
import { build } from 'esbuild';
import { gzipSync } from 'node:zlib';
const pkg = JSON.parse(await readFile('package.json', 'utf8'));
const temporary = await mkdtemp(join(tmpdir(), 'vibe-package-'));
try {
  const [packed] = JSON.parse(execFileSync('npm', ['pack', '--ignore-scripts', '--json', '--pack-destination', temporary, '--cache', join(temporary, 'cache')], { encoding: 'utf8' }));
  assert(packed.files.every(({ path }) => path.startsWith('dist/') || ['package.json','README.md','LICENSE'].includes(path)), 'Unexpected files in package');
  const packagePath = join(temporary, 'node_modules', pkg.name);
  await mkdir(packagePath, { recursive: true });
  execFileSync('tar', ['-xzf', join(temporary, packed.filename), '--strip-components=1', '-C', packagePath]);
  // Use the installed, locked dependency graph without downloading anything. The
  // toolkit itself MUST resolve from the tarball, not from the source checkout.
  for (const name of Object.keys(pkg.dependencies)) {
    const target = join(temporary, 'node_modules', name);
    await mkdir(dirname(target), { recursive: true });
    await symlink(resolve('node_modules', name), target, 'dir');
  }
  for (const name of ['', '/auth', '/sharepoint', '/copilot', '/redirect-bridge']) {
    const entry = join(temporary, 'consumer.mjs');
    await writeFile(entry, `import * as api from ${JSON.stringify(pkg.name + name)}; if (!Object.keys(api).length) throw Error('Empty exports');`);
    execFileSync(process.execPath, [entry]);
  }
  const consumer = join(temporary, 'consumer.ts');
  await writeFile(consumer, `import { createVibeMicrosoftClient } from '${pkg.name}';\nimport { VibeAuthClient } from '${pkg.name}/auth';\nconst auth = {clientId:'app', tenantId:'tenant', redirectUri:'https://example.com/callback'};\nconst client = createVibeMicrosoftClient({auth, copilot:{environmentId:'env',schemaName:'agent'}});\nclient.getCopilot().sendActivity({type:'message',text:'Hello'},'conversation');\nnew VibeAuthClient(auth);`);
  execFileSync(process.execPath, [resolve('node_modules/typescript/bin/tsc'), '--noEmit', '--strict', '--skipLibCheck', '--target', 'ES2022', '--module', 'ESNext', '--moduleResolution', 'Bundler', consumer]);

  const config = { auth: { clientId: 'test', tenantId: 'tenant', redirectUri: 'https://example.com/callback' }, sharePoint: { siteUrl: 'https://contoso.sharepoint.com/sites/test', lists: { requests: { name: 'Requests' } } } };
  async function exercise(code, label) {
    const requests = [];
    const sandbox = { URL, URLSearchParams, Headers, Request, Response, AbortController, TextEncoder, TextDecoder, crypto: webcrypto, console, setTimeout, clearTimeout, setInterval, clearInterval, queueMicrotask,
      fetch: async (url, init) => { requests.push([url, init]); return new Response(JSON.stringify({ value: [{ Id: 1 }] }), { headers: { 'content-type': 'application/json' } }); },
    };
    sandbox.window = { document: {}, location: { origin: 'https://example.com' }, crypto: webcrypto };
    const context = vm.createContext(sandbox);
    try { vm.runInContext(code, context, { timeout: 10000 }); } catch (error) { throw new Error(`${label}: ${error.message}`); }
    const api = context.VibeToolkitMicrosoft;
    assert.equal(api.VIBE_TOOLKIT_VERSION, pkg.version);
    let initialized = false;
    const account = { homeAccountId: 'account' };
    const msal = { initialize: async () => { initialized = true; }, getActiveAccount: () => account, getAllAccounts: () => [account], acquireTokenSilent: async ({ scopes }) => { assert(initialized); assert.equal(scopes[0], 'https://contoso.sharepoint.com/.default'); return { accessToken: 'package-test' }; } };
    const client = api.createVibeMicrosoftClient(config, { msalInstance: msal });
    const items = await client.getListItems('requests');
    assert.equal(items[0].Id, 1);
    assert.equal(new Headers(requests[0][1].headers).get('authorization'), 'Bearer package-test');
    console.log(`${label}: exports, initialization and real PnPjs request passed`);
  }
  const entry = join(temporary, 'browser-consumer.mjs');
  await writeFile(entry, `export * from '${pkg.name}';`);
  const built = await build({ entryPoints: [entry], bundle: true, platform: 'browser', format: 'iife', globalName: 'VibeToolkitMicrosoft', write: false, minify: true });
  await exercise(built.outputFiles[0].text, 'npm browser consumer');
  const globalFile = join(packagePath, 'dist/index.global.js');
  await exercise(await readFile(globalFile, 'utf8'), 'CDN script');
  const esm = await build({ entryPoints: [join(packagePath, 'dist/index.browser.js')], bundle: true, platform: 'browser', format: 'iife', globalName: 'VibeToolkitMicrosoft', write: false });
  await exercise(esm.outputFiles[0].text, 'CDN ESM');
  const authOnly = await build({ stdin: { contents: `export { VibeAuthClient } from '${pkg.name}/auth';`, resolveDir: temporary }, bundle: true, platform: 'browser', write: false, metafile: true });
  assert(!Object.keys(authOnly.metafile.inputs).some(p => p.includes('@pnp') || p.includes('@microsoft/agents')), 'Auth-only entry includes unrelated SDKs');
  const integrity = JSON.parse(await readFile(join(packagePath, 'dist/integrity.json'), 'utf8'));
  for (const [file, expected] of Object.entries(integrity)) {
    assert.equal('sha384-' + createHash('sha384').update(await readFile(join(packagePath, 'dist', file))).digest('base64'), expected);
  }
  const bridgeContext = vm.createContext({ console });
  vm.runInContext(await readFile(join(packagePath, 'dist/redirect-bridge.global.js'), 'utf8'), bridgeContext);
  assert.equal(typeof bridgeContext.VibeToolkitRedirectBridge.broadcastResponseToMainFrame, 'function');
  const size = (await readFile(globalFile)).length;
  assert(size < 850000, 'CDN size exceeded 850 kB budget');
  console.log(`Package checks passed; CDN ${size} bytes / ${gzipSync(await readFile(globalFile)).length} bytes gzip. No live tenant requests made.`);
} finally { await rm(temporary, { recursive: true, force: true }); }
