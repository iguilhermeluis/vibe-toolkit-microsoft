import { build } from 'esbuild';
import { readFile, writeFile, readdir } from 'node:fs/promises';
import { createRequire } from 'node:module';
import { dirname, resolve, join } from 'node:path';
import { createHash } from 'node:crypto';
const pkg = JSON.parse(await readFile('package.json', 'utf8'));
const common = { bundle: true, platform: 'browser', target: 'es2022', minify: true, sourcemap: true, legalComments: 'eof', metafile: true, logLevel: 'info' };
const outputs = [
  ['src/index.ts', 'dist/index.global.js', 'iife', 'VibeToolkitMicrosoft'],
  ['src/index.ts', 'dist/index.browser.js', 'esm'],
  ['src/redirect-bridge.ts', 'dist/redirect-bridge.browser.js', 'esm'],
  ['src/redirect-bridge.ts', 'dist/redirect-bridge.global.js', 'iife', 'VibeToolkitRedirectBridge'],
];
const integrity = {};
for (const [entry, outfile, format, globalName] of outputs) {
  const result = await build({ ...common, entryPoints: [entry], outfile, format, globalName,
    banner: { js: `/*! ${pkg.name} v${pkg.version} | MIT | Third-party notices: THIRD-PARTY-NOTICES.txt */` },
  });
  for (const output of Object.values(result.metafile.outputs)) {
    if (output.imports.some(item => item.external)) throw new Error(`External dependency in browser bundle: ${outfile}`);
  }
  integrity[outfile.slice(5)] = 'sha384-' + createHash('sha384').update(await readFile(outfile)).digest('base64');
}
await writeFile('dist/integrity.json', JSON.stringify(integrity, null, 2) + '\n');

// Include licenses for the complete runtime dependency graph, including dependencies
// already bundled inside the Copilot SDK's browser distribution.
const visited = new Set();
const notices = [];
async function visit(name, from = resolve('package.json')) {
  const resolver = createRequire(from);
  let manifest;
  try { manifest = resolver.resolve(name + '/package.json'); }
  catch {
    let directory = dirname(resolver.resolve(name));
    while (true) {
      try { const candidate = join(directory, 'package.json'); if (JSON.parse(await readFile(candidate, 'utf8')).name === name) { manifest = candidate; break; } } catch {}
      const parent = dirname(directory); if (parent === directory) throw new Error(`Cannot find license for ${name}`); directory = parent;
    }
  }
  if (visited.has(manifest)) return;
  visited.add(manifest);
  const data = JSON.parse(await readFile(manifest, 'utf8'));
  const directory = dirname(manifest);
  const files = (await readdir(directory)).filter(file => /^(license|licence|copying|copyright|notice)(\.|$)/i.test(file));
  let license = (await Promise.all(files.map(file => readFile(join(directory, file), 'utf8')))).join('\n\n');
  if (!license && name.startsWith('@microsoft/agents-')) license = await readFile('licenses/agents-sdk-MIT.txt', 'utf8');
  if (!license) throw new Error(`Missing license text: ${name}`);
  notices.push(`${data.name}@${data.version} (${data.license})\n${license}`);
  for (const dependency of Object.keys(data.dependencies ?? {})) await visit(dependency, manifest);
}
for (const name of Object.keys(pkg.dependencies)) await visit(name);
await writeFile('dist/THIRD-PARTY-NOTICES.txt', notices.join('\n\n' + '='.repeat(72) + '\n\n'));
