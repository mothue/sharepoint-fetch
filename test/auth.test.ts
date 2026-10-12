import assert from 'node:assert/strict';
import test from 'node:test';
import { buildClient } from '../src/index.js';
import type { AuthMaterial } from '../src/index.js';

const site = 'https://tenant.sharepoint.com/sites/demo';
const auth: AuthMaterial = { kind: 'cookieHeader', cookieHeader: 'FedAuth=a=b==; rtFa=second' };
const cookies = (init?: RequestInit) => new Set((new Headers(init?.headers).get('cookie') ?? '').split(/;\s*/).filter(Boolean));

// Catches dropped auth variants, truncated '=' values, and consumed final bodies.
test('accepts both auth variants', async () => {
  for (const material of [auth, { kind: 'fedAuthRtFa', fedAuth: 'a=b==', rtFa: 'second' } as const]) {
    const client = buildClient(material, site, { fetch: async (_url, init) => {
      assert.deepEqual(cookies(init), new Set(['FedAuth=a=b==', 'rtFa=second']));
      return new Response(new Uint8Array([0, 1, 255, 128]));
    } });
    const response = await client.get(`${site}/file`);
    assert.equal(response.bodyUsed, false);
    assert.deepEqual(new Uint8Array(await response.arrayBuffer()), new Uint8Array([0, 1, 255, 128]));
  }
});

// Catches asynchronous validation and unsafe header/cookie interpolation.
test('validates auth synchronously without leaking values', () => {
  const invalid: AuthMaterial[] = [
    { kind: 'cookieHeader', cookieHeader: 'ignored; =; empty=' },
    { kind: 'cookieHeader', cookieHeader: 'bad name=private-secret' },
    { kind: 'cookieHeader', cookieHeader: 'FedAuth=private-secret\r\nInjected=yes' },
    { kind: 'fedAuthRtFa', fedAuth: '', rtFa: 'private-secret' },
    { kind: 'fedAuthRtFa', fedAuth: 'private-secret', rtFa: '' },
    { kind: 'fedAuthRtFa', fedAuth: 'private-secret;Injected=yes', rtFa: 'ok' },
    { kind: 'fedAuthRtFa', fedAuth: 'ok', rtFa: 'private-secret\nInjected=yes' },
  ];
  for (const material of invalid) {
    assert.throws(() => buildClient(material, site), error => error instanceof Error && !error.message.includes('private-secret'));
  }
});

test('ignores empty cookie fragments when valid cookies remain', async () => {
  const client = buildClient({ kind: 'cookieHeader', cookieHeader: 'ignored; =missing; empty=; FedAuth=valid' }, site, {
    fetch: async (_url, init) => {
      assert.deepEqual(cookies(init), new Set(['FedAuth=valid']));
      return new Response();
    },
  });
  await client.get(site);
});

// Catches wrong request method, auto redirect, omitted cancellation, or ignored UA.
test('uses configurable request options', async () => {
  const controller = new AbortController();
  for (const userAgent of [undefined, 'custom-agent']) {
    const client = buildClient(auth, site, { userAgent, fetch: async (url, init) => {
      assert.equal(String(url), `${site}/file`);
      assert.equal(init?.method, 'GET');
      assert.equal(init?.redirect, 'manual');
      assert.equal(init?.signal, controller.signal);
      assert.equal(new Headers(init?.headers).get('user-agent'), userAgent ?? 'sharepoint-fetch-bot');
      return new Response();
    } });
    await client.get(new URL(`${site}/file`), { signal: controller.signal });
  }
});

// Catches copying auth headers across hosts, subdomains, or insecure transitions.
test('keeps authentication scoped', async () => {
  for (const destination of ['https://other.example/file', 'https://child.tenant.sharepoint.com/file', 'http://tenant.sharepoint.com/file']) {
    const client = buildClient(auth, site, { fetch: async (url, init) => {
      if (String(url) === site) {
        assert.deepEqual(cookies(init), new Set(['FedAuth=a=b==', 'rtFa=second']));
        return new Response('redirect', { status: 302, headers: { location: destination } });
      }
      assert.equal(String(url), destination);
      assert.deepEqual(cookies(init), new Set());
      return new Response('final');
    } });
    assert.equal(await (await client.get(site)).text(), 'final');
  }
  const localClient = buildClient(auth, 'https://localhost/sites/demo', { fetch: async (_url, init) => {
    assert.deepEqual(cookies(init), new Set());
    return new Response();
  } });
  await localClient.get('http://localhost/file');
});

// Catches comma splitting of Set-Cookie and failure to persist jar changes.
test('updates multiple response cookies', async () => {
  let request = 0;
  const client = buildClient(auth, site, { fetch: async (_url, init) => {
    if (request++ === 0) {
      const headers = new Headers();
      headers.append('set-cookie', 'FedAuth=new=value; Path=/; Expires=Wed, 01 Jan 2031 00:00:00 GMT');
      headers.append('set-cookie', 'rtFa=new-second; Path=/; Expires=Wed, 01 Jan 2031 00:00:00 GMT');
      return new Response(null, { headers });
    }
    assert.deepEqual(cookies(init), new Set(['FedAuth=new=value', 'rtFa=new-second']));
    return new Response();
  } });
  await client.get(site);
  await client.get(site);
});

// Catches disregarded cookie scope/expiry, or stale cookies after deletion.
test('obeys response cookie attributes', async () => {
  let first = true;
  const client = buildClient(auth, site, { fetch: async (url, init) => {
    if (first) {
      first = false;
      const headers = new Headers();
      for (const cookie of [
        'FedAuth=deleted; Max-Age=0; Path=/', 'rtFa=deleted; Expires=Thu, 01 Jan 1970 00:00:00 GMT; Path=/',
        'scoped=path; Path=/private', 'secure=only; Secure; Path=/',
        'domain=shared; Domain=tenant.sharepoint.com; Path=/',
        'foreign=rejected; Domain=other.example; Path=/',
        'expired=no; Expires=Thu, 01 Jan 1970 00:00:00 GMT; Path=/',
      ]) headers.append('set-cookie', cookie);
      return new Response(null, { headers });
    }
    const expected = new Map([
      ['https://tenant.sharepoint.com/private/file', ['scoped=path', 'secure=only', 'domain=shared']],
      ['https://tenant.sharepoint.com/public', ['secure=only', 'domain=shared']],
      ['http://tenant.sharepoint.com/private/file', ['scoped=path', 'domain=shared']],
      ['https://child.tenant.sharepoint.com/private/file', ['domain=shared']],
    ]);
    assert.deepEqual(cookies(init), new Set(expected.get(String(url))));
    return new Response();
  } });
  await client.get(site);
  for (const url of ['https://tenant.sharepoint.com/private/file', 'https://tenant.sharepoint.com/public', 'http://tenant.sharepoint.com/private/file', 'https://child.tenant.sharepoint.com/private/file']) await client.get(url);
});

// Catches missing status handling or incorrect relative URL resolution.
test('follows each redirect status', async () => {
  for (const status of [301, 302, 303, 307, 308]) {
    let calls = 0;
    const final = new Response('unread');
    const client = buildClient(auth, site, { fetch: async (url, init) => {
      assert.equal(init?.method, 'GET');
      if (calls++ === 0) return new Response('redirect', { status, headers: { location: '../final' } });
      assert.equal(String(url), 'https://tenant.sharepoint.com/sites/final');
      return final;
    } });
    assert.equal(await client.get(`${site}/file`), final);
    assert.equal(final.bodyUsed, false);
    const noLocation = new Response('no-location', { status });
    const stationary = buildClient(auth, site, { fetch: async () => noLocation });
    assert.equal(await stationary.get(site), noLocation);
    assert.equal(await noLocation.text(), 'no-location');
  }
});

// Catches off-by-one limits, invalid settings, and leaked redirect bodies.
test('enforces redirect limits', async () => {
  for (const maxRedirects of [-1, 0.5, NaN, Infinity]) assert.throws(() => buildClient(auth, site, { maxRedirects }));
  for (const maxRedirects of [undefined, 0, 1]) {
    let requests = 0;
    let cancelled = 0;
    const client = buildClient(auth, site, { maxRedirects, fetch: async () => {
      requests++;
      return new Response(new ReadableStream({ cancel() { cancelled++; } }), { status: 302, headers: { location: '/again' } });
    } });
    await assert.rejects(client.get(site), /redirect/i);
    assert.equal(requests, (maxRedirects ?? 20) + 1);
    assert.equal(cancelled, requests);
  }
  let requests = 0;
  let cancelled = false;
  const client = buildClient(auth, site, { maxRedirects: 1, fetch: async () => ++requests === 1
    ? new Response(new ReadableStream({ cancel() { cancelled = true; } }), { status: 302, headers: { location: '/final' } })
    : new Response('success') });
  assert.equal(await (await client.get(site)).text(), 'success');
  assert.equal(cancelled, true);
});
