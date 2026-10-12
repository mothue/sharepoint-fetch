# sharepoint-fetch

SharePoint Online のファイルを Cookie 認証で取得する TypeScript ライブラリです。開発・テスト・サンプル実行は Bun を標準とし、Bun 1.4.2 以上と Node.js 22 以上で利用できます。ESM と CommonJS の両方で JavaScript と型定義を配布し、実行時依存は `tough-cookie` のみです。

## ローカルでビルド・インストール

このリポジトリから npm 互換の tarball を作成し、Bun または npm で利用側のプロジェクトにインストールできます。

```sh
bun install --frozen-lockfile
bun run verify
bun pm pack
# 利用側プロジェクトで、作成した tarball のパスを指定
bun add /path/to/sharepoint-fetch-0.1.0.tgz
# Node.js のプロジェクトでは npm でもインストールできます
npm install /path/to/sharepoint-fetch-0.1.0.tgz
```

`bun pm pack` は `prepack` でビルドを実行します。配布物には `dist`、TypeScript ソース、README、既存の LICENSE が含まれます。依存関係は `bun.lock` で固定します。利用側が配布済みライブラリを Node.js で実行する際に Bun は不要です。

## URL から取得して保存

ブラウザのログイン済みセッションから取得した Cookie を渡します。認証方式は `FedAuth` と `rtFa` の組、または Cookie ヘッダー文字列です。値は対象ファイルへのアクセス権を持つ有効なものを指定してください。

```typescript
import { writeFile } from 'node:fs/promises';
import { buildClient, downloadFile, SharePointTarget, type AuthMaterial } from 'sharepoint-fetch';

async function main(): Promise<void> {
  const target = SharePointTarget.fromDocumentUrl(
    'https://example.sharepoint.com/sites/demo/_layouts/15/Doc.aspx?sourcedoc={01234567-89AB-CDEF-0123-456789ABCDEF}&file=資料.xlsx',
  );
  const auth: AuthMaterial = { kind: 'fedAuthRtFa', fedAuth: 'your-fedauth', rtFa: 'your-rtfa' };
  // Cookie ヘッダー方式では、こちらを buildClient に渡します。
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

`fromDocumentUrl` は `sites`、`teams`、`personal`、SharePoint のルートサイトと、クエリやフラグメント内の入れ子の URL を解析します。GUID は大文字に正規化され、日本語の `file` 値をラベルとして保持します。ラベルはファイル名として安全化されていないため、保存先は利用側で指定してください。

`downloadFile` は本文を未消費の標準 `Response` として返します。401・403・404 などでも応答を返すため、保存前に `response.ok` を確認します。通信失敗・キャンセル・リダイレクト上限超過は Promise を reject します。

## ターゲットを明示的に作成

```typescript
import { SharePointTarget } from 'sharepoint-fetch';

const target = new SharePointTarget({
  siteUrl: 'https://example.sharepoint.com/sites/demo/',
  itemId: '{01234567-89ab-cdef-0123-456789abcdef}',
  label: '資料.xlsx',
});
const renamed = target.withLabel('downloaded.xlsx');
console.log(renamed.downloadUrl());
```

`withLabel` は元のターゲットを変更せず新しいインスタンスを返します。`siteUrl` は絶対 HTTP/HTTPS URL、`itemId` は有効な GUID が必要です。`siteUrl` のクエリ・フラグメントと末尾のスラッシュは除去し、サイトのベース URL に正規化します。

CommonJS の JavaScript では `const { SharePointTarget, buildClient, downloadFile } = require('sharepoint-fetch')` を利用できます。TypeScript の CommonJS 利用例です。

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

## 公開 API

| API | 役割 |
| --- | --- |
| `AuthMaterial` | `{ kind: 'cookieHeader', cookieHeader }` または `{ kind: 'fedAuthRtFa', fedAuth, rtFa }` |
| `SharePointTargetInit` | `siteUrl`、`itemId`、任意の `label`・`documentUrl` |
| `SharePointTarget` | readonly の `siteUrl`、`itemId`、`label`、`documentUrl`。`fromDocumentUrl`、`withLabel`、`downloadUrl` |
| `buildClient(auth, siteUrl, options?)` | 同期的な認証検証と `SharePointClient` の生成 |
| `ClientOptions` | 任意の `fetch`、`userAgent`、`maxRedirects`（既定 20、0 以上の整数） |
| `SharePointClient.get(url, options?)` | `string` または `URL` を GET し、`Promise<Response>` を返す |
| `downloadFile(client, target, options?)` | REST URL から取得し、`Promise<Response>` を返す |
| `DownloadOptions` | 任意の `signal: AbortSignal` |

既定 User-Agent は `sharepoint-fetch-bot` です。初期 Cookie は対象ホストに限定され、HTTPS の場合は Secure を付けます。応答の `Set-Cookie` を保持し、リダイレクトごとに domain・path・secure・expiry に従って送信 Cookie を選び直します。ライブラリの読み込みだけで通信・対話入力・ファイル保存は起きません。

## 開発と対話式サンプル

```sh
bun install --frozen-lockfile
bun run typecheck
bun test
bun run test:node
bun run build
bun run verify:package
bun run verify
bun run example
# 別バージョンの Node.js で既存テストと配布物を検証
/path/to/node --import tsx --test test/*.test.ts
bun scripts/verify-package.mjs --node /path/to/node
```

`bun run example` は TypeScript を直接実行し、URL、認証方式（1: Cookie ヘッダー、2: FedAuth/rtFa）、認証値を順に読み、成功した本文だけを現在のディレクトリへ保存します。端末で入力する認証値は表示しません。パイプ入力にも対応します。保存名から `/`、`\`、`:`、`*`、`?`、`"`、`<`、`>`、`|` と制御文字を除去し、空・`.`・`..` になる場合は GUID を使います。失敗時は非ゼロで終了し、HTTP ステータスまたはエラーメッセージを表示します。

`bun run verify` は型チェック、Bun と Node.js のテスト、両モジュール形式のビルド、配布検証をまとめて実行します。Node 互換検証には `node` と `npm` が必要です。`tsx` は Node 上で TypeScript のテスト・サンプルを実行するための開発依存で、Bun の実行や配布物には不要です。型チェック・型定義生成には TypeScript コンパイラを使います。

サンプルの terminal テストには macOS/Linux の `script`・`stty`・`ps` コマンド、配布検証には `tar` を使用します。配布検証は Bun で tarball を作り、一時プロジェクトへ Bun と npm の両方で開発依存を除いてインストールします。それぞれの配布物について Bun と Node.js の ESM・CommonJS 実行、Bun の TypeScript 直接実行、NodeNext 型解決、README の TypeScript 例を確認します。

## Rust 版からの移行

| Rust | TypeScript |
| --- | --- |
| `build_client` | `buildClient` |
| `download_file` | `downloadFile` |
| `SharePointTarget::from_document_url` | `SharePointTarget.fromDocumentUrl` |
| `with_label` / `download_url` | `withLabel` / `downloadUrl` |
| `site_url` / `item_id` / `document_url` | `siteUrl` / `itemId` / `documentUrl` |
| `AuthMaterial::CookieHeader` | `{ kind: 'cookieHeader', cookieHeader }` |
| `AuthMaterial::FedAuthRtFa` | `{ kind: 'fedAuthRtFa', fedAuth, rtFa }` |
| 公開フィールドによるターゲット作成 | `new SharePointTarget(init)` |

Rust ソースと Cargo ファイルは TypeScript 実装へ置き換えました。ライセンスの条件は [LICENSE](LICENSE) を参照してください。
