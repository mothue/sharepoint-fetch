# sharepoint-fetch

A TypeScript library for downloading SharePoint Online files with cookie authentication. Bun is the default tool for development, testing, and running the example. The library supports Bun 1.4.2 or later and Node.js 22 or later, with JavaScript and type declarations for both ESM and CommonJS. Its only direct runtime dependency is `tough-cookie`.

## Build and install locally

Create an npm-compatible tarball from this repository, then install it in your project with Bun or npm.

```sh
bun install --frozen-lockfile
bun run verify
bun pm pack
# In your project, specify the path to the generated tarball
bun add /path/to/sharepoint-fetch-0.1.0.tgz
# You can also use npm in a Node.js project
npm install /path/to/sharepoint-fetch-0.1.0.tgz
```

`bun pm pack` runs the build through the `prepack` script. The package includes `dist`, TypeScript source files, this README, and the existing LICENSE. Development dependencies are locked in `bun.lock`. Bun is not required to run the installed library in Node.js.

## Download and save a file from a URL

Provide cookies from an authenticated browser session. Authentication accepts either a `FedAuth` and `rtFa` pair or a Cookie header string. Use valid cookies from a session with permission to access the target file.

```typescript
import { writeFile } from 'node:fs/promises';
import { buildClient, downloadFile, SharePointTarget, type AuthMaterial } from 'sharepoint-fetch';

async function main(): Promise<void> {
  const target = SharePointTarget.fromDocumentUrl(
    'https://example.sharepoint.com/sites/demo/_layouts/15/Doc.aspx?sourcedoc={01234567-89AB-CDEF-0123-456789ABCDEF}&file=report.xlsx',
  );
  const auth: AuthMaterial = { kind: 'fedAuthRtFa', fedAuth: 'your-fedauth', rtFa: 'your-rtfa' };
  // For Cookie header authentication, pass this value to buildClient instead.
  const cookieAuth: AuthMaterial = { kind: 'cookieHeader', cookieHeader: 'FedAuth=your-fedauth; rtFa=your-rtfa' };
  const client = buildClient(auth, target.siteUrl);
  const response = await downloadFile(client, target);
  if (!response.ok) {
    await response.body?.cancel();
    throw new Error(`Download failed: HTTP ${response.status}`);
  }
  await writeFile('downloaded.xlsx', new Uint8Array(await response.arrayBuffer()));
}

main().catch(console.error);
```

`fromDocumentUrl` parses `sites`, `teams`, `personal`, SharePoint root sites, and URLs nested in query parameters or fragments. GUIDs are normalized to uppercase, and Unicode characters in the `file` parameter are preserved in the label. Labels are not sanitized for use as filenames, so choose the output path in your application.

`downloadFile` returns a standard `Response` with its body unread. HTTP errors such as 401, 403, and 404 also return a response; check `response.ok` before saving it. Network failures, cancellation, and exceeding the redirect limit reject the promise.

## Create a target explicitly

```typescript
import { SharePointTarget } from 'sharepoint-fetch';

const target = new SharePointTarget({
  siteUrl: 'https://example.sharepoint.com/sites/demo/',
  itemId: '{01234567-89ab-cdef-0123-456789abcdef}',
  label: 'report.xlsx',
});
const renamed = target.withLabel('downloaded.xlsx');
console.log(renamed.downloadUrl());
```

`withLabel` returns a new instance without modifying the original target. `siteUrl` must be an absolute HTTP or HTTPS URL, and `itemId` must be a valid GUID. Query parameters, fragments, and trailing slashes are removed from `siteUrl` to normalize it to the site's base URL.

In CommonJS JavaScript, use `const { SharePointTarget, buildClient, downloadFile } = require('sharepoint-fetch')`. Here is a CommonJS TypeScript example:

```typescript
import sharepoint = require('sharepoint-fetch');

const target = new sharepoint.SharePointTarget({
  siteUrl: 'https://example.sharepoint.com/sites/demo',
  itemId: '01234567-89AB-CDEF-0123-456789ABCDEF',
});
const client = sharepoint.buildClient({ kind: 'cookieHeader', cookieHeader: 'FedAuth=your-fedauth' }, target.siteUrl);
async function download(): Promise<Response> {
  return sharepoint.downloadFile(client, target);
}
```

## Public API

| API | Description |
| --- | --- |
| `AuthMaterial` | `{ kind: 'cookieHeader', cookieHeader }` or `{ kind: 'fedAuthRtFa', fedAuth, rtFa }` |
| `SharePointTargetInit` | `siteUrl`, `itemId`, and optional `label` and `documentUrl` |
| `SharePointTarget` | Readonly `siteUrl`, `itemId`, `label`, and `documentUrl`; methods `fromDocumentUrl`, `withLabel`, and `downloadUrl` |
| `buildClient(auth, siteUrl, options?)` | Validates authentication synchronously and creates a `SharePointClient` |
| `ClientOptions` | Optional `fetch`, `userAgent`, and `maxRedirects` (a nonnegative integer, default: 20) |
| `SharePointClient.get(url, options?)` | Sends a GET request to a `string` or `URL` and returns `Promise<Response>` |
| `downloadFile(client, target, options?)` | Fetches the file's REST URL and returns `Promise<Response>` |
| `DownloadOptions` | Optional `signal: AbortSignal` |

The default User-Agent is `sharepoint-fetch-bot`. Initial cookies are restricted to the target host and marked Secure for HTTPS. Response `Set-Cookie` headers update the cookie jar, and cookies are selected for each redirect according to their domain, path, secure flag, and expiry. Importing the library does not start network requests, prompt for input, or save files.

## Development and interactive example

```sh
bun install --frozen-lockfile
bun run typecheck
bun test
bun run test:node
bun run build
bun run verify:package
bun run verify
bun run example
# Verify the tests and package with a different Node.js version
/path/to/node --import tsx --test test/*.test.ts
bun scripts/verify-package.mjs --node /path/to/node
```

`bun run example` executes TypeScript directly and prompts for the URL, authentication method (1: Cookie header, 2: FedAuth/rtFa), and credentials in that order. It saves the body to the current directory only when the HTTP response is successful. Credentials entered in a terminal are hidden, and piped input is also supported. The output filename removes `/`, `\`, `:`, `*`, `?`, `"`, `<`, `>`, `|`, and control characters; if the result is empty, `.`, or `..`, it uses the GUID instead. On failure, the example exits with a nonzero status and displays the HTTP status or error message.

`bun run verify` runs type checking, tests on Bun and Node.js, builds for both module formats, and package verification. Node compatibility checks require `node` and `npm`. The development dependency `tsx` runs TypeScript tests and the example on Node; it is not needed for Bun execution or the installed package. The TypeScript compiler handles type checking and declaration generation.

Terminal tests for the example use the macOS/Linux commands `script`, `stty`, and `ps`; package verification also uses `tar`. Package verification creates a tarball with Bun and installs it into separate temporary projects with Bun and npm, excluding development dependencies. Each installation is checked for ESM and CommonJS execution on both Bun and Node.js, direct TypeScript execution on Bun, NodeNext type resolution, and the TypeScript examples in this README.

## Migrating from the Rust version

| Rust | TypeScript |
| --- | --- |
| `build_client` | `buildClient` |
| `download_file` | `downloadFile` |
| `SharePointTarget::from_document_url` | `SharePointTarget.fromDocumentUrl` |
| `with_label` / `download_url` | `withLabel` / `downloadUrl` |
| `site_url` / `item_id` / `document_url` | `siteUrl` / `itemId` / `documentUrl` |
| `AuthMaterial::CookieHeader` | `{ kind: 'cookieHeader', cookieHeader }` |
| `AuthMaterial::FedAuthRtFa` | `{ kind: 'fedAuthRtFa', fedAuth, rtFa }` |
| Creating a target with public fields | `new SharePointTarget(init)` |

The Rust source and Cargo files have been replaced with the TypeScript implementation. See [LICENSE](LICENSE) for the license terms.
