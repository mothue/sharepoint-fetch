import assert from 'node:assert/strict';
import { createServer, type RequestListener } from 'node:http';
import { once } from 'node:events';
import test from 'node:test';
import { buildClient, downloadFile, SharePointTarget } from '../src/index.js';

async function withServer(handler: RequestListener, run: (origin: string) => Promise<void>): Promise<void> {
  const server = createServer(handler);
  server.listen(0, '127.0.0.1');
  try {
    await once(server, 'listening');
    const address = server.address();
    assert.ok(address && typeof address !== 'string');
    await run(`http://127.0.0.1:${address.port}`);
  } finally {
    const closed = new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
    server.closeAllConnections();
    await closed;
  }
}

const itemId = '01234567-89AB-CDEF-0123-456789ABCDEF';

// Catches accepted site page URLs sending GET to the page instead of the file.
for (const suffix of ['?web=1', '#overview', '/?web=1', '/#overview']) {
  test(`requests the file endpoint for explicit sites with query or fragment: ${suffix}`, async () => {
    let path: string | undefined;
    await withServer((req, res) => {
      path = req.url;
      res.end('file');
    }, async origin => {
      const client = buildClient({ kind: 'cookieHeader', cookieHeader: 'FedAuth=test' }, origin);
      const response = await downloadFile(client, new SharePointTarget({ siteUrl: `${origin}/sites/demo${suffix}`, itemId }));
      assert.equal(await response.text(), 'file');
      assert.equal(path, "/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
    });
  });
}

// Catches a wrong REST URL, request headers/method, or changed/consumed bytes.
test('downloads unchanged binary bytes', async () => {
  let observed: { url?: string; method?: string; cookie?: string; agent?: string } | undefined;
  await withServer((req, res) => {
    observed = { url: req.url, method: req.method, cookie: req.headers.cookie, agent: req.headers['user-agent'] };
    res.end(Buffer.from([0, 1, 255, 128]));
  }, async origin => {
    const client = buildClient({ kind: 'fedAuthRtFa', fedAuth: 'a=b==', rtFa: 'second' }, `${origin}/sites/demo`);
    const target = new SharePointTarget({ siteUrl: `${origin}/sites/demo`, itemId });
    const response = await downloadFile(client, target);
    assert.equal(response.bodyUsed, false);
    assert.deepEqual(new Uint8Array(await response.arrayBuffer()), new Uint8Array([0, 1, 255, 128]));
    assert.equal(observed?.url, "/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
    assert.equal(observed?.method, 'GET');
    assert.deepEqual(new Set(observed?.cookie?.split(/;\s*/)), new Set(['FedAuth=a=b==', 'rtFa=second']));
    assert.equal(observed?.agent, 'sharepoint-fetch-bot');
  });
});

// Catches treating HTTP status failures as transport exceptions.
test('returns HTTP errors for callers', async () => {
  for (const status of [401, 403, 404]) await withServer((_req, res) => {
    res.writeHead(status);
    res.end(`error-${status}`);
  }, async origin => {
    const client = buildClient({ kind: 'cookieHeader', cookieHeader: 'FedAuth=test' }, origin);
    const response = await downloadFile(client, new SharePointTarget({ siteUrl: origin, itemId }));
    assert.equal(response.status, status);
    assert.equal(response.bodyUsed, false);
    assert.equal(await response.text(), `error-${status}`);
  });
});

// Catches auto-follow bypassing jar updates and returning the redirect Response.
test('reuses cookies after a real redirect', async () => {
  const requests: Array<{ url?: string; cookie?: string }> = [];
  await withServer((req, res) => {
    requests.push({ url: req.url, cookie: req.headers.cookie });
    if (req.url !== '/final') {
      res.writeHead(302, { location: '/final', 'set-cookie': ['FedAuth=new; Path=/', 'rtFa=updated; Path=/; Expires=Wed, 01 Jan 2031 00:00:00 GMT'] });
      res.end('discard');
    } else res.end('final');
  }, async origin => {
    const client = buildClient({ kind: 'fedAuthRtFa', fedAuth: 'old', rtFa: 'old-second' }, origin);
    const response = await downloadFile(client, new SharePointTarget({ siteUrl: origin, itemId }));
    assert.equal(response.status, 200);
    assert.equal(response.bodyUsed, false);
    assert.equal(await response.text(), 'final');
    assert.equal(requests.length, 2);
    assert.equal(requests[1].url, '/final');
    assert.deepEqual(new Set(requests[1].cookie?.split(/;\s*/)), new Set(['FedAuth=new', 'rtFa=updated']));
  });
});

// Catches swallowed transport errors and dropped download cancellation signal.
test('rejects transport failures and cancellation', async () => {
  await withServer(req => req.socket.destroy(), async origin => {
    const client = buildClient({ kind: 'cookieHeader', cookieHeader: 'FedAuth=test' }, origin);
    await assert.rejects(downloadFile(client, new SharePointTarget({ siteUrl: origin, itemId })), TypeError);
  });
  let reached: () => void = () => {};
  const requestReceived = new Promise<void>(resolve => { reached = resolve; });
  await withServer(() => reached(), async origin => {
    const client = buildClient({ kind: 'cookieHeader', cookieHeader: 'FedAuth=test' }, origin);
    const controller = new AbortController();
    const downloading = downloadFile(client, new SharePointTarget({ siteUrl: origin, itemId }), { signal: controller.signal });
    const rejected = assert.rejects(downloading, error => error instanceof Error && error.name === 'AbortError');
    await requestReceived;
    controller.abort();
    await rejected;
  });
});
