import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { mkdtemp, readFile, readdir, rm, writeFile, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

const execute = promisify(execFile);
const root = fileURLToPath(new URL('../', import.meta.url));
const args = process.argv.slice(2);
if (args.length && (args.length !== 2 || args[0] !== '--node')) {
  throw new Error('Usage: bun scripts/verify-package.mjs [--node <executable>]');
}
const bun = process.versions.bun ? process.execPath : 'bun';
const node = args[1] ?? 'node';
const workspace = await mkdtemp(join(tmpdir(), 'sharepoint-package-'));

async function run(command, commandArgs, cwd) {
  try {
    return await execute(command, commandArgs, { cwd, maxBuffer: 4 * 1024 * 1024, timeout: 120_000 });
  } catch (error) {
    throw new Error(`${command} ${commandArgs.join(' ')} failed\n${error.stdout ?? ''}${error.stderr ?? ''}`, { cause: error });
  }
}

try {
  const filename = (await run(bun, ['pm', 'pack', '--quiet', '--destination', workspace], root)).stdout.trim();
  const tarball = resolve(workspace, filename);
  const paths = (await run('tar', ['-tzf', tarball], root)).stdout.trim().split('\n')
    .filter(path => !path.endsWith('/')).map(path => {
      assert.ok(path.startsWith('package/'), `Unexpected tarball entry: ${path}`);
      return path.slice('package/'.length);
    });
  for (const required of ['dist/esm/index.js', 'dist/esm/index.d.ts', 'dist/cjs/index.js',
    'dist/cjs/index.d.ts', 'dist/cjs/package.json', 'src/index.ts', 'README.md', 'LICENSE', 'package.json']) {
    assert.ok(paths.includes(required), `Tarball is missing ${required}`);
  }
  for (const path of paths) {
    assert.ok(/^(dist\/|src\/.*\.ts$|README\.md$|LICENSE$|package\.json$)/.test(path), `Unexpected packaged file: ${path}`);
    assert.ok(!/\.rs$|(^|\/)(Cargo\.(toml|lock)|target|test|tests|docs|\.superpowers)(\/|$)/.test(path), `Development/Rust artifact packaged: ${path}`);
  }
  for (const manager of ['bun', 'npm']) {
    const consumer = join(workspace, `consumer-${manager}`);
    await mkdir(consumer);
    await writeFile(join(consumer, 'package.json'), JSON.stringify({ private: true, type: 'module' }));
    const installArgs = manager === 'bun'
      ? ['add', '--production', '--ignore-scripts', tarball]
      : ['install', '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund', tarball];
    await run(manager === 'bun' ? bun : 'npm', installArgs, consumer);
    const installedRoot = join(consumer, 'node_modules/sharepoint-fetch');
    const installed = JSON.parse(await readFile(join(installedRoot, 'package.json'), 'utf8'));
    assert.deepEqual(Object.keys(installed.dependencies), ['tough-cookie']);
    for (const developmentTool of ['typescript', 'tsx', '@types/node']) {
      await assert.rejects(readFile(join(consumer, 'node_modules', developmentTool, 'package.json')), { code: 'ENOENT' });
    }
    assert.deepEqual(await readFile(join(installedRoot, 'LICENSE')), await readFile(join(root, 'LICENSE')));
    for (const format of ['esm', 'cjs']) {
      const map = JSON.parse(await readFile(join(installedRoot, 'dist', format, 'auth.js.map'), 'utf8'));
      assert.ok(map.sourcesContent?.[0]?.includes('export function buildClient'), `${format} source map lacks inline source`);
    }

    const runtimeBody = String.raw`
const originalFetch = globalThis.fetch;
globalThis.fetch = () => { throw new Error('Network activity during import'); };
const { SharePointTarget, buildClient, downloadFile } = LOAD_PACKAGE;
globalThis.fetch = originalFetch;
const itemId = '01234567-89AB-CDEF-0123-456789ABCDEF';
const binary = Buffer.from([0, 255, 128, 10]);
const requests = [];
const server = createServer((req, res) => {
  requests.push({ url: req.url, method: req.method, cookie: req.headers.cookie });
  if (req.url !== '/final') {
    res.writeHead(302, { location: '/final', 'set-cookie': ['FedAuth=updated; Path=/', 'rtFa=other; Path=/; Expires=Wed, 01 Jan 2031 00:00:00 GMT'] });
    res.end('discard');
  } else res.end(binary);
});
server.listen(0, '127.0.0.1');
await once(server, 'listening');
try {
  const origin = 'http://127.0.0.1:' + server.address().port;
  const target = SharePointTarget.fromDocumentUrl(origin + '/sites/demo/doc.aspx?sourcedoc=%7B01234567-89ab-cdef-0123-456789abcdef%7D&file=%E8%B3%87%E6%96%99.xlsx');
  assert.equal(target.itemId, itemId);
  assert.equal(target.label, '資料.xlsx');
  for (const auth of [{ kind: 'cookieHeader', cookieHeader: 'FedAuth=sample==; rtFa=second' }, { kind: 'fedAuthRtFa', fedAuth: 'sample==', rtFa: 'second' }]) {
    requests.length = 0;
    const response = await downloadFile(buildClient(auth, target.siteUrl), target);
    assert.equal(response.ok, true);
    assert.equal(response.bodyUsed, false);
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), binary);
    assert.equal(requests.length, 2);
    assert.equal(requests[0].url, "/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
    assert.equal(requests[0].method, 'GET');
    assert.deepEqual(new Set(requests[0].cookie.split(/;\s*/)), new Set(['FedAuth=sample==', 'rtFa=second']));
    assert.deepEqual(new Set(requests[1].cookie.split(/;\s*/)), new Set(['FedAuth=updated', 'rtFa=other']));
  }
} finally {
  const closed = new Promise((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
  server.closeAllConnections();
  await closed;
}
  `;
    const esm = `import assert from 'node:assert/strict';\nimport { createServer } from 'node:http';\nimport { once } from 'node:events';\n${runtimeBody.replace('LOAD_PACKAGE', "await import('sharepoint-fetch')")}`;
    const cjs = `const assert = require('node:assert/strict');\nconst { createServer } = require('node:http');\nconst { once } = require('node:events');\n(async () => {\n${runtimeBody.replace('LOAD_PACKAGE', "require('sharepoint-fetch')")}\n})().catch(error => { console.error(error); process.exitCode = 1; });`;
    await writeFile(join(consumer, 'consumer.mjs'), esm);
    await writeFile(join(consumer, 'consumer.cjs'), cjs);
    await writeFile(join(consumer, 'consumer.ts'), `import type { SharePointTargetInit } from 'sharepoint-fetch';\n${esm.replace("const itemId =", "const itemId: SharePointTargetInit['itemId'] =")}`);
    for (const [name, runtime] of [['Bun', bun], ['Node', node]]) {
      await run(runtime, ['consumer.mjs'], consumer);
      await run(runtime, ['consumer.cjs'], consumer);
      if (name === 'Bun') await run(runtime, ['consumer.ts'], consumer);
      console.log(`${manager} installed package: ${name} import/require${name === 'Bun' ? '/TypeScript' : ''} passed (${(await run(runtime, ['--version'], consumer)).stdout.trim()})`);
    }

    const types = `import { SharePointTarget, buildClient, downloadFile, type AuthMaterial, type SharePointTargetInit, type ClientOptions, type DownloadOptions, type SharePointClient } from 'sharepoint-fetch';
type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
type Expect<T extends true> = T;
type ClientReturn = Expect<Equal<ReturnType<typeof buildClient>, SharePointClient>>;
type DownloadReturn = Expect<Equal<ReturnType<typeof downloadFile>, Promise<Response>>>;
type GetReturn = Expect<Equal<ReturnType<SharePointClient['get']>, Promise<Response>>>;
type ClientArgs = Expect<Equal<Parameters<typeof buildClient>, [AuthMaterial, string, ClientOptions?]>>;
type DownloadArgs = Expect<Equal<Parameters<typeof downloadFile>, [SharePointClient, SharePointTarget, DownloadOptions?]>>;
type GetArgs = Expect<Equal<Parameters<SharePointClient['get']>, [string | URL, DownloadOptions?]>>;
type AuthShape = Expect<Equal<AuthMaterial, { kind: 'cookieHeader'; cookieHeader: string } | { kind: 'fedAuthRtFa'; fedAuth: string; rtFa: string }>>;
type InitShape = Expect<Equal<SharePointTargetInit, { siteUrl: string; itemId: string; label?: string; documentUrl?: string }>>;
type ClientOptionsShape = Expect<Equal<ClientOptions, { fetch?: typeof globalThis.fetch; userAgent?: string; maxRedirects?: number }>>;
type DownloadOptionsShape = Expect<Equal<DownloadOptions, { signal?: AbortSignal }>>;
type ConstructorArgs = Expect<Equal<ConstructorParameters<typeof SharePointTarget>, [SharePointTargetInit]>>;
type ParseArgs = Expect<Equal<Parameters<typeof SharePointTarget.fromDocumentUrl>, [string]>>;
type ParseReturn = Expect<Equal<ReturnType<typeof SharePointTarget.fromDocumentUrl>, SharePointTarget>>;
type LabelArgs = Expect<Equal<Parameters<SharePointTarget['withLabel']>, [string]>>;
type LabelReturn = Expect<Equal<ReturnType<SharePointTarget['withLabel']>, SharePointTarget>>;
type UrlArgs = Expect<Equal<Parameters<SharePointTarget['downloadUrl']>, []>>;
type UrlReturn = Expect<Equal<ReturnType<SharePointTarget['downloadUrl']>, string>>;
const init: SharePointTargetInit = { siteUrl: 'https://example.sharepoint.com/sites/demo', itemId: '01234567-89AB-CDEF-0123-456789ABCDEF', label: '資料.xlsx', documentUrl: 'https://example.sharepoint.com/doc' };
const target: SharePointTarget = new SharePointTarget(init).withLabel('download.bin');
const site: string = target.siteUrl;
const id: string = target.itemId;
const label: string = target.label;
const documentUrl: string | undefined = target.documentUrl;
const url: string = target.downloadUrl();
const parsed: SharePointTarget = SharePointTarget.fromDocumentUrl(init.documentUrl!);
const cookieAuth: AuthMaterial = { kind: 'cookieHeader', cookieHeader: 'FedAuth=test' };
const fedAuth: AuthMaterial = { kind: 'fedAuthRtFa', fedAuth: 'a', rtFa: 'b' };
const options: ClientOptions = { fetch: globalThis.fetch, userAgent: 'test', maxRedirects: 2 };
const downloadOptions: DownloadOptions = { signal: new AbortController().signal };
const client: SharePointClient = buildClient(cookieAuth, site, options);
const response: Promise<Response> = downloadFile(client, target, downloadOptions);
client.get(new URL(url), downloadOptions);
// @ts-expect-error both auth cookie values are required
const invalidAuth: AuthMaterial = { kind: 'fedAuthRtFa', fedAuth: 'a' };
// @ts-expect-error target properties are readonly
target.itemId = 'other';
// @ts-expect-error target properties are readonly
target.siteUrl = 'https://other.example';
// @ts-expect-error target properties are readonly
target.label = 'other';
// @ts-expect-error target properties are readonly
target.documentUrl = undefined;
// @ts-expect-error return value is a Response, not bytes
const invalidResponse: Promise<Uint8Array> = downloadFile(client, target);
  `;
    await writeFile(join(consumer, 'consumer.mts'), types);
    await writeFile(join(consumer, 'consumer.cts'), types);
    const readme = await readFile(join(installedRoot, 'README.md'), 'utf8');
    const examples = [...readme.matchAll(/```(?:ts|typescript)\n([\s\S]*?)```/g)];
    assert.ok(examples.length > 0, 'README must contain TypeScript examples');
    for (const [index, example] of examples.entries()) for (const suffix of ['mts', 'cts']) {
      await writeFile(join(consumer, `readme-${index}.${suffix}`), example[1]);
    }
    // The development compiler/types are tooling only; the actual consumer remains --omit=dev.
    const compiler = fileURLToPath(import.meta.resolve('typescript/bin/tsc'));
    const nodeTypes = dirname(dirname(fileURLToPath(import.meta.resolve('@types/node/package.json'))));
    const typeFiles = (await readdir(consumer)).filter(file => /\.(mts|cts)$/.test(file));
    await run(process.execPath, [compiler, '--noEmit', '--strict', '--target', 'ES2022', '--module', 'NodeNext',
      '--moduleResolution', 'NodeNext', '--types', 'node', '--typeRoots', nodeTypes, ...typeFiles], consumer);
    console.log(`${manager} installed package: .mts/.cts NodeNext types and ${examples.length} README examples passed`);
  }
  console.log(`Tarball contents and Bun/npm runtime-only installations passed (${paths.length} files)`);
} finally {
  await rm(workspace, { recursive: true, force: true });
}
