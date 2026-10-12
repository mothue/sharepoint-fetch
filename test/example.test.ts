import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { mkdir, mkdtemp, readFile, readdir, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { spawnInTerminal, waitForExit } from './terminal.js';

const itemId = '01234567-89AB-CDEF-0123-456789ABCDEF';
const binary = Buffer.from([0, 255, 128, 10, 13]);
const sample = fileURLToPath(new URL('../examples/download.ts', import.meta.url));
const tsx = fileURLToPath(import.meta.resolve('tsx'));

async function runSample(label: string, auth: '1' | '2', terminal = false, status = 200) {
  const parent = await mkdtemp(join(tmpdir(), 'sharepoint-example-'));
  const cwd = join(parent, 'downloads');
  await mkdir(cwd);
  const requests: Array<{ url?: string; cookie?: string }> = [];
  const server = createServer((req, res) => {
    requests.push({ url: req.url, cookie: req.headers.cookie });
    res.writeHead(status);
    res.end(binary);
  });
  server.listen(0, '127.0.0.1');
  try {
    await once(server, 'listening');
    const address = server.address();
    assert.ok(address && typeof address !== 'string');
    const url = new URL(`http://127.0.0.1:${address.port}/sites/demo/document.aspx`);
    url.searchParams.set('sourcedoc', itemId);
    url.searchParams.set('file', label);
    const lines = [url.href, auth, ...(auth === '1' ? ['FedAuth=sample==; rtFa=second'] : ['sample==', 'second'])];
    const prompts = ['Document URL: ', 'Authentication (1: Cookie header, 2: FedAuth/rtFa): ',
      ...(auth === '1' ? ['Cookie header: '] : ['FedAuth: ', 'rtFa: '])];
    const runtimeArgs = process.versions.bun
      ? [process.execPath, sample]
      : [process.execPath, '--import', tsx, sample];
    const args = terminal
      ? ['/bin/sh', '-c', 'before=$(stty -g); "$@"; status=$?; after=$(stty -g); if [ "$before" != "$after" ]; then echo "Terminal state was not restored"; exit 99; fi; echo "Terminal state restored"; exit "$status"', 'sh', ...runtimeArgs]
      : runtimeArgs;
    // script allocates a real pseudo-terminal; ordinary runs deliver all lines at once.
    const child = terminal
      ? spawnInTerminal(args, cwd)
      : spawn(args[0]!, args.slice(1), { cwd, detached: true });
    let stdout = '';
    let stderr = '';
    let sent = 0;
    child.stdout.on('data', (chunk: Buffer) => {
      stdout += chunk.toString();
      if (terminal) while (sent < prompts.length && stdout.includes(prompts[sent]!)) {
        child.stdin.write(`${lines[sent]}\n`);
        sent++;
        if (sent === prompts.length) child.stdin.end();
      }
    });
    child.stderr.on('data', (chunk: Buffer) => { stderr += chunk.toString(); });
    if (!terminal) child.stdin.end(`${lines.join('\n')}\n`);
    const exitCode = await waitForExit(child);
    const names = await readdir(cwd);
    const files = await Promise.all(names.map(async name => [name, await readFile(join(cwd, name))] as const));
    const siblings = await readdir(parent);
    return { exitCode, stdout, stderr, names, files, siblings, requests };
  } finally {
    const closed = new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
    server.closeAllConnections();
    await closed;
    await rm(parent, { recursive: true, force: true });
  }
}

// Catches selecting the wrong auth branch, lost piped lines, or corrupted/unsaved bytes.
for (const terminal of [false, true]) for (const auth of ['1', '2'] as const) {
  test(`saves a download using authentication ${auth} through ${terminal ? 'terminal' : 'pipe'}`, async () => {
    const result = await runSample('資料.xlsx', auth, terminal);
    assert.equal(result.exitCode, 0, result.stderr || result.stdout);
    assert.deepEqual(result.names, ['資料.xlsx']);
    assert.deepEqual(result.files[0]?.[1], binary);
    assert.equal(result.requests[0]?.url, "/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
    assert.deepEqual(new Set(result.requests[0]?.cookie?.split(/;\s*/)), new Set(['FedAuth=sample==', 'rtFa=second']));
    assert.doesNotMatch(result.stdout + result.stderr, /sample==|second/);
    if (terminal) assert.match(result.stdout, /Terminal state restored/);
  });
}

// Catches traversal, forbidden separators, missing GUID fallback, or removed Unicode.
test('sanitizes output names', async () => {
  const cases: Array<[string, string]> = [
    ['a/b\\c:d*e?f"g<h>i|j.xlsx', 'abcdefghij.xlsx'],
    ['../escape.bin', '..escape.bin'],
    ['', itemId], ['.', itemId], ['..', itemId], ['/\\:*?"<>|', itemId],
  ];
  for (const [label, expected] of cases) {
    const result = await runSample(label, '1');
    assert.equal(result.exitCode, 0, result.stderr);
    assert.deepEqual(result.names, [expected]);
    assert.deepEqual(result.files[0]?.[1], binary);
    assert.deepEqual(result.siblings, ['downloads']);
  }
});

// Catches saving HTTP error bodies or reporting failure as successful completion.
test('does not save an HTTP failure', async () => {
  for (const terminal of [false, true]) for (const auth of ['1', '2'] as const) {
    const result = await runSample('failure.bin', auth, terminal, 403);
    assert.notEqual(result.exitCode, 0);
    assert.match(result.stderr + result.stdout, /403/);
    assert.doesNotMatch(result.stdout + result.stderr, /sample==|second/);
    if (terminal) assert.match(result.stdout, /Terminal state restored/);
    assert.deepEqual(result.names, []);
  }
});
