import { Cookie, CookieJar, MemoryCookieStore } from 'tough-cookie';
import type { AuthMaterial, ClientOptions, DownloadOptions, SharePointClient } from './models.js';

const cookieName = /^[!#$%&'*+\-.^_`|~\da-z]+$/i;
const cookieValue = /^[\x21\x23-\x2B\x2D-\x3A\x3C-\x5B\x5D-\x7E]+$/;
const redirectStatuses = new Set([301, 302, 303, 307, 308]);

export function buildClient(auth: AuthMaterial, siteUrl: string, options: ClientOptions = {}): SharePointClient {
  const site = httpUrl(siteUrl);
  const maxRedirects = options.maxRedirects ?? 20;
  if (!Number.isInteger(maxRedirects) || maxRedirects < 0) {
    throw new Error('maxRedirects must be a non-negative integer');
  }
  const jar = new CookieJar(new MemoryCookieStore(), { allowSecureOnLocal: false });
  try {
    for (const [key, value] of authCookies(auth)) {
      jar.setCookieSync(new Cookie({ key, value, path: '/', secure: site.protocol === 'https:' }), site.href);
    }
  } catch {
    // Never expose a parser's diagnostic containing authentication input.
    throw new Error('Invalid authentication material');
  }
  const request = options.fetch ?? globalThis.fetch;
  const userAgent = options.userAgent ?? 'sharepoint-fetch-bot';

  return {
    async get(url: string | URL, downloadOptions: DownloadOptions = {}): Promise<Response> {
      let current = httpUrl(url);
      let redirects = 0;
      while (true) {
        const headers = new Headers({ 'user-agent': userAgent });
        const cookie = jar.getCookieStringSync(current.href);
        if (cookie) headers.set('cookie', cookie);
        const response = await request(current.href, {
          method: 'GET', headers, redirect: 'manual', signal: downloadOptions.signal,
        });
        for (const cookie of response.headers.getSetCookie()) {
          // Invalid server cookies are ignored rather than exposing their values.
          jar.setCookieSync(cookie, current.href, { ignoreError: true });
        }
        const location = response.headers.get('location');
        if (!redirectStatuses.has(response.status) || location === null) return response;
        await response.body?.cancel();
        if (redirects >= maxRedirects) throw new Error('Maximum redirect count exceeded');
        current = httpUrl(new URL(location, current));
        redirects++;
      }
    },
  };
}

function authCookies(auth: AuthMaterial): Array<[string, string]> {
  let entries: Array<[string, string]>;
  if (auth.kind === 'cookieHeader') {
    if (/[\r\n]/.test(auth.cookieHeader)) throw new Error();
    entries = [];
    for (const fragment of auth.cookieHeader.split(';')) {
      const separator = fragment.indexOf('=');
      if (separator === -1) continue;
      const name = fragment.slice(0, separator).trim();
      const value = fragment.slice(separator + 1).trim();
      if (name && value) entries.push([name, value]);
    }
  } else if (auth.kind === 'fedAuthRtFa') {
    entries = [['FedAuth', auth.fedAuth], ['rtFa', auth.rtFa]];
  } else {
    throw new Error();
  }
  if (!entries.length || entries.some(([name, value]) => !cookieName.test(name) || !cookieValue.test(value))) {
    throw new Error();
  }
  return entries;
}

function httpUrl(raw: string | URL): URL {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new Error('URL must be an absolute HTTP or HTTPS URL');
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error('URL must be an absolute HTTP or HTTPS URL');
  }
  return url;
}
