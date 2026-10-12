# sharepoint-fetch TypeScript Library Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Rust環境を必要とせず、TypeScriptから型付きで利用できるsharepoint-fetch npmライブラリに置き換える。

**Architecture:** URL解析、Cookie付きHTTPクライアント、ダウンロードを小さなモジュールに分ける。標準fetchとtough-cookieを使い、未消費のResponseを利用側へ返す。TypeScriptからESM・CommonJSと対応する型定義を生成し、配布物を別プロジェクトで検証する。

**Tech Stack:** Node.js 22以上、TypeScript、Node.js標準fetch・node:test、tough-cookie、tsx、npm。

**Spec:** [承認済みの設計書](../specs/2026-10-12-typescript-library-design.md)。2026-10-12のユーザーの「ええよ」を設計書の承認として受領した。

## Global Constraints

- 想定する利用環境はNode.js 22以上。
- パッケージ名は`sharepoint-fetch`、初期バージョンは既存の`0.1.0`を維持する。
- TypeScriptのstrictチェックを使用する。
- Cookie管理の実行時依存は`tough-cookie`に限定する。
- `dist/esm`はESM、`dist/cjs`はCommonJSとして解釈される構成にする。
- npmパッケージにはビルド成果物、TypeScriptソース、README、LICENSEを含める。
- 既存の`LICENSE`は維持し、READMEはTypeScript版に置き換える。
- ライブラリを読み込むだけでは、ネットワーク通信、対話入力、ファイル保存を行わない。
- `downloadFile`の戻り値は未消費の標準`Response`である。
- 認証情報をログやエラーメッセージへ埋め込まない。
- npmへの公開やリモートリポジトリへのpushはこの移植作業の範囲に含めない。

## Review Focus

1. ラッパーURLと内部URLが別サイトを指すとき、GUIDとサイトは同じ候補から採用する（Task 1）。
2. 不正なパーセントエンコードと日本語のファイル名が混在しても、有効な内部URLを解析できる（Task 1）。
3. Expires属性にカンマがある複数のSet-Cookieを、別々のCookieとして反映する（Task 2）。
4. Cookieヘッダーの不正な名前・改行を拒否し、認証値をエラーに露出しない（Task 2）。
5. パッケージを開発依存なしでインストールした利用側でも、ESM・CommonJSの実行と型解決が成立する（Task 3）。

---

## File Structure

| ファイル | 責務 |
| --- | --- |
| `src/index.ts` | 公開エクスポート |
| `src/models.ts` | 認証素材・オプション・クライアントの型 |
| `src/target.ts` | URL候補探索、GUID検証、ターゲット生成 |
| `src/auth.ts` | Cookie jar、入力検証、認証付きGETとリダイレクト |
| `src/sharepoint.ts` | ターゲットを使ったファイル取得 |
| `examples/download.ts` | 対話入力、HTTPエラー処理、ファイル保存 |
| `test/target.test.ts` | URL解析とターゲット作成の検証 |
| `test/auth.test.ts` | 認証・Cookie・リダイレクトの検証 |
| `test/download.test.ts` | ローカルHTTPサーバーによる通信検証 |
| `test/example.test.ts` | サンプルの保存・失敗時動作の検証 |
| `scripts/build.mjs` | ESM・CommonJSビルドと一時ソースの片付け |
| `scripts/verify-package.mjs` | tarballの内容、利用側の実行・型解決の検証 |
| `package.json`, `package-lock.json` | 配布情報、実行コマンド、固定した依存関係 |
| `tsconfig.json`, `tsconfig.esm.json`, `tsconfig.cjs.json` | 開発用型チェックと配布物の生成 |

## Task 1: npm開発環境とターゲット解析

**Files:** Create `package.json`, `package-lock.json`, `tsconfig.json`, `src/target.ts`, `test/target.test.ts`; modify `.gitignore`.

**Interfaces:**

- Consumes: 設計書の`SharePointTargetInit`と`SharePointTarget`の正確な公開署名。
- Produces: `SharePointTargetInit`、`SharePointTarget`（`siteUrl`, `itemId`, `label`, `documentUrl`, `fromDocumentUrl`, `withLabel`, `downloadUrl`）。
- `SharePointTarget`と`SharePointTargetInit`は`src/target.ts`で定義し、最終的に`src/index.ts`から再エクスポートする。

- [x] **Step 1: テスト実行に必要なnpm設定を作成する。**

`type: module`、`engines.node: >=22`、TypeScript・@types/node・tsxを開発依存、tough-cookieを実行時依存として設定する。`npm run typecheck`は`tsc --noEmit`、`npm test`は`node --import tsx --test test/*.test.ts`とする。`npm install`でロックファイルを生成し、`.gitignore`に`node_modules/`, `dist/`, `.build/`, `*.tgz`を追加する。型チェック設定はstrictとNodeNextを使い、`src/`、`test/`、`examples/`を対象とする。

- [x] **Step 2: `test/target.test.ts`に失敗するテストを書く。**

固定GUIDは`01234567-89ab-cdef-0123-456789abcdef`、期待する正規化値は`01234567-89AB-CDEF-0123-456789ABCDEF`とする。各ケースの名前と検証内容：

| テスト名 | 入力と期待結果 |
| --- | --- |
| `parses document links` | `https://tenant.sharepoint.com/:x:/r/sites/demo/_layouts/15/Doc.aspx?sourcedoc=%7B01234567-89ab-cdef-0123-456789abcdef%7D&file=example.xlsx` → サイト`https://tenant.sharepoint.com/sites/demo`、正規化GUID、ラベル`example.xlsx` |
| `supports site scopes and root sites` | sites・teams・personalを各々生成し、そのサイトURLを得る。ルートサイトは`https://tenant.sharepoint.com` |
| `supports query casing and fallback labels` | `SOURCEDOC`と`FILE`も解析でき、fileが空または欠落ならGUIDをラベルにする |
| `follows nested query and fragment URLs` | 直接・二重・三重のエンコードを含むクエリ値とフラグメントから内部の有効なURLを解析する |
| `keeps nested site and item together` | outerの`/sites/wrapper`に内部URLを付けた入力 → 内部の`/sites/demo`を採用する |
| `preserves Unicode with malformed sibling input` | 不正な`%ZZ`の別クエリ値と、有効な内部URLの`file=報告書.xlsx` → ラベルを壊さず解析できる |
| `rejects invalid targets` | 欠落・不正なGUID、相対URL、非HTTP URL、空のサイト名、スコープのないURLの不正なSharePointホストを拒否する |
| `creates explicit targets and copies labels` | 明示作成は末尾`/`を除去し、`withLabel`後も元のラベルが変わらない。REST URLが設計書どおりで、元のdocumentUrlを維持する |

代表的なテスト本体：

```ts
const input = 'https://tenant.sharepoint.com/:x:/r/sites/demo/_layouts/15/Doc.aspx?sourcedoc=%7B01234567-89ab-cdef-0123-456789abcdef%7D&file=example.xlsx';
const target = SharePointTarget.fromDocumentUrl(input);
assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo');
assert.equal(target.itemId, '01234567-89AB-CDEF-0123-456789ABCDEF');
assert.equal(target.label, 'example.xlsx');
assert.equal(target.documentUrl, input);
assert.equal(target.downloadUrl(), "https://tenant.sharepoint.com/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
assert.equal(target.withLabel('renamed.xlsx').label, 'renamed.xlsx');
assert.equal(target.label, 'example.xlsx');
```

- [x] **Step 3: REDを確認する。**

Run: `node --import tsx --test test/target.test.ts`。
Expected: `src/target.ts`がまだ存在しないことによる失敗。依存関係やテスト構文による失敗なら先に修正する。

- [x] **Step 4: `src/target.ts`を実装する。**

公開署名は設計書と一致させる。URLとURLSearchParamsで候補を幅優先に探索し、候補の再処理を防ぐ。内部URLのデコードを最大3段扱い、不正なデコードを局所的に無視する。同じ候補からサイト・GUID・ラベルを採用する。`sharepoint.com`の一致はドメイン境界で検証する。

- [x] **Step 5: GREENと型チェックを確認する。**

Run: `node --import tsx --test test/target.test.ts`、`npm run typecheck`。
Expected: すべて成功。公開フィールドと署名が設計書と一致する。

- [x] **Step 6: 検証済みのTask 1をコミットする。**

`feat: port SharePoint target parsing to TypeScript`。

## Task 2: Cookie認証付きクライアントとファイル取得

**Files:** Create `src/models.ts`, `src/auth.ts`, `src/sharepoint.ts`, `src/index.ts`, `test/auth.test.ts`, `test/download.test.ts`.

**Interfaces:**

- Consumes: Task 1の`SharePointTarget.downloadUrl(): string`とターゲット型。
- Produces: 設計書どおりの`AuthMaterial`, `ClientOptions`, `DownloadOptions`, `SharePointClient`, `buildClient(auth, siteUrl, options?): SharePointClient`, `downloadFile(client, target, options?): Promise<Response>`。
- 認証・オプション・クライアントの型は`src/models.ts`に定義する。公開する値と型をすべて`src/index.ts`から再エクスポートする。

- [x] **Step 1: `test/auth.test.ts`に失敗する認証・リダイレクトテストを書く。**

注入したfetchでリクエストのURL、method、Cookie、User-Agent、redirect、signalを観測する。Cookie比較はセミコロンで分けた名前・値で行い、無関係な順序を制約しない。

| テスト名 | 期待する検証 |
| --- | --- |
| `accepts both auth variants` | 両方式でFedAuth/rtFaを送る。`FedAuth=a=b==`の値内の`=`を維持する |
| `validates auth synchronously without leaking values` | 有効なCookieなし、片方が空のFedAuth/rtFa、不正な名前・改行を同期的に拒否する。エラーに入力した認証値を含めない |
| `uses configurable request options` | GET、既定UA`sharepoint-fetch-bot`、UA変更、signal、manual redirectを確認する |
| `keeps authentication scoped` | 別ホストとHTTPSからHTTPへの移動で初期Cookieを送らない |
| `updates multiple response cookies` | Expires属性にカンマを含む2個のSet-Cookieを別々に反映する |
| `obeys response cookie attributes` | domain・path・secure・expiry属性とCookie削除が後続リクエストへ適用される |
| `follows each redirect status` | 301・302・303・307・308と相対Locationを追跡する。LocationなしはそのResponseを返す |
| `enforces redirect limits` | 既定20回、指定0回・1回、負数・小数・NaNの拒否、リダイレクト元ボディの解放を確認する |

値内の`=`を保持することを確認するテスト本体：

```ts
const client = buildClient(
  { kind: 'cookieHeader', cookieHeader: 'FedAuth=a=b==; rtFa=second' },
  'https://tenant.sharepoint.com/sites/demo',
  { fetch: async (_url, init) => {
    const headers = new Headers(init?.headers);
    assert.equal(init?.method, 'GET');
    assert.equal(init?.redirect, 'manual');
    assert.equal(headers.get('user-agent'), 'sharepoint-fetch-bot');
    assert.deepEqual(new Set((headers.get('cookie') ?? '').split(/;\s*/)), new Set(['FedAuth=a=b==', 'rtFa=second']));
    return new Response(new Uint8Array([0, 1, 255, 128]));
  } },
);
const response = await client.get('https://tenant.sharepoint.com/sites/demo/file');
assert.deepEqual(new Uint8Array(await response.arrayBuffer()), new Uint8Array([0, 1, 255, 128]));
```

- [x] **Step 2: `test/download.test.ts`に失敗するローカルHTTP通信テストを書く。**

`node:http`のサーバーを空きポートで起動し、各テストで必ず停止する。外部テナントや実際の認証情報に依存しない。

| テスト名 | 期待する検証 |
| --- | --- |
| `downloads unchanged binary bytes` | 明示作成したターゲットに対しREST URL・GET・Cookie・UAをサーバーで確認する。バイナリ`[0, 1, 255, 128]`をそのまま読み取れる |
| `returns HTTP errors for callers` | 401・403・404のResponseをthrowせず返し、status・本文を読める |
| `reuses cookies after a real redirect` | 応答Cookieを相対リダイレクト先へ反映し、最終Responseを返す |
| `rejects transport failures and cancellation` | サーバーの接続切断とAbortControllerによる通信キャンセルでrejectする |

- [x] **Step 3: REDを確認する。**

Run: `node --import tsx --test test/auth.test.ts test/download.test.ts`。
Expected: 未実装モジュールまたはAPIによる失敗。

- [x] **Step 4: `src/models.ts`・`src/auth.ts`を実装する。**

CookieJarの同期メモリstoreと`allowSecureOnLocal: false`を使用する。初期Cookieはホスト限定、`Path=/`、HTTPSならSecureに設定する。入力に含まれる最初の`=`だけを区切りとし、不正な認証情報のエラーを秘匿する。各GETでCookieを選択し、`Headers.getSetCookie()`で応答Cookieを反映する。リダイレクト先ごとにCookieを選び直し、上限20回を既定とする。最終Responseは本文を読まずに返し、通信エラーとキャンセルを識別できるよう維持する。

- [x] **Step 5: `src/sharepoint.ts`・`src/index.ts`を実装する。**

`downloadFile`は`target.downloadUrl()`への`client.get`にオプションを渡す。エントリーポイントは副作用を持たず、Task 1とTask 2の公開値・型を再エクスポートする。

- [x] **Step 6: GREENと型チェックを確認する。**

Run: `npm test`、`npm run typecheck`。
Expected: ターゲット・認証・実HTTP通信の全テストが成功する。

- [x] **Step 7: 検証済みのTask 2をコミットする。**

`feat: add cookie-authenticated SharePoint downloads`。

## Task 3: サンプル、npm配布、Rust版の置き換え

**Files:** Create `examples/download.ts`, `test/example.test.ts`, `scripts/build.mjs`, `scripts/verify-package.mjs`, `tsconfig.esm.json`, `tsconfig.cjs.json`; modify `package.json`, `README.md`, `.gitignore`; delete `src/*.rs`, `Cargo.toml`, `Cargo.lock`.

**Interfaces:**

- Consumes: Task 1とTask 2の全公開API。
- Produces: `npm run build`, `npm run verify:package`, `npm run example`と、install済みパッケージからの`import`・`require`。
- `npm run verify`は型チェック・全テスト・ビルド・配布検証を順に実行する。
- `scripts/verify-package.mjs --node <executable>`で、利用側の実行に使うNode.jsを指定できる。指定しない場合は`process.execPath`を使う。
- サンプルは`tsx examples/download.ts`で実行し、ソースの`src/index.ts`を拡張子`.js`の相対importで利用する。

- [x] **Step 1: `test/example.test.ts`に失敗するサンプルのテストを書く。**

サーバーはTask 2と同様にローカルで起動する。サンプルは一時的なcwdで子プロセスとして実行し、対話入力を渡す。

| テスト名 | 期待する検証 |
| --- | --- |
| `saves a download using either authentication mode` | 2方式それぞれで認証・URLを渡し、保存ファイルのバイナリ内容を確認する |
| `sanitizes output names` | `/`, `\\`, `:`, `*`, `?`, `"`, `<`, `>`, `|`を安全な名前にし、空・`.`・`..`はGUIDを使う。cwdの外にファイルを作らない |
| `does not save an HTTP failure` | 403応答では非ゼロ終了し、ファイルを保存しない。エラーにはステータスを含める |

子プロセスの結果と一時cwdを観測して、成功時は`assert.equal(exitCode, 0)`と保存されたバイト列の`assert.deepEqual`、403時は`assert.notEqual(exitCode, 0)`・`assert.match(stderr, /403/)`・`assert.deepEqual(await readdir(cwd), [])`を確認する。

- [x] **Step 2: REDを確認する。**

Run: `node --import tsx --test test/example.test.ts`。
Expected: 未実装サンプルによる失敗。

- [x] **Step 3: 対話式サンプルを実装する。**

`node:readline`でURLと認証方式・値を読み、入力インターフェースを必ず閉じる。ライブラリAPIで取得して`Response.ok`を判定し、成功時だけ`arrayBuffer()`をファイルへ保存する。サンプル内でファイル名を整え、設計書のGUIDフォールバックを適用する。

- [x] **Step 4: パッケージ利用側の検証スクリプトを作り、失敗を確認する。**

`scripts/verify-package.mjs`は`npm pack --json`でtarballを生成し、一時ディレクトリへ`npm install --omit=dev`する。ESM・CommonJS利用側でターゲット解析と認証付きダウンロードを実行する。`.mts`と`.cts`の利用側をNodeNextで型チェックし、`AuthMaterial`等の型と戻り値の署名を確認する。tarballにESM・CommonJS・型定義・src・README・LICENSEが含まれ、Rust成果物・テスト・開発文書を含まないことを確認する。一時ディレクトリとtarballはfinallyで片付ける。

Run: `node scripts/verify-package.mjs`。
Expected: まだビルド成果物またはexportsがないため失敗する。

- [x] **Step 5: 配布ビルドとエクスポートを実装する。**

ESMはsrcをNodeNextで`dist/esm`へコンパイルする。CommonJSは`.build/cjs`にsrcと`type: commonjs`のpackage.jsonを一時コピーし、NodeNextで`dist/cjs`へコンパイルする。`dist/cjs/package.json`にも`type: commonjs`を記載し、一時ソースを片付ける。両方で型定義とsource mapを生成し、一時ソースの削除後もデバッグできるよう`inlineSources`を有効にする。package.jsonのimport・require条件に各`index.js`・`index.d.ts`を指定し、`files`を`dist`・`src`に限定する。READMEとLICENSEの自動収録も配布検証で確認する。`prepack`はビルドを実行する。

- [x] **Step 6: READMEを更新し、Rust成果物を削除する。**

Node.js 22以上、npm開発コマンド、認証の2方式、URL解析と明示作成、バイナリ保存、HTTPエラー判定、公開API、Rust版からの名前の対応を記載する。既存LICENSEの本文は変更しない。TypeScript版に不要なRustソースとCargoファイルを削除する。

- [x] **Step 7: 全体の検証を行う。**

Run: `npm ci`、`npm run verify`。
Expected: クリーンな依存インストール、型チェック、全テスト、両形式のビルド、配布物の利用側検証が成功する。

追加の実行環境検証：Node.js 22.0.0で全テストと配布物の利用側実行を確認し、複数のSet-Cookieを分離して扱えることも確認する。既存のNode.js 22.0.0がなければ次を使用する。npm操作は開発環境のNode.jsで行い、指定したNode.jsは利用側コードの実行に使う。

```sh
npx --yes --package=node@22.0.0 -- node --import tsx --test test/*.test.ts
task_node22_path=$(npx --yes --package=node@22.0.0 -- node -p process.execPath)
node scripts/verify-package.mjs --node "$task_node22_path"
```

- [x] **Step 8: 全体レビューと完了監査を行う。**

設計書の8個の完了条件と各公開署名を、最新のファイル・コマンド結果で照合する。READMEの例は利用側の型検証に含める。副作用のないimport、Cookieのスコープ、Unicode、HTTPエラー、両モジュール形式の型解決を重点的にレビューする。

- [x] **Step 9: 検証済みのTask 3をコミットする。**

`feat: ship TypeScript package and replace Rust implementation`。

## Execution Handoff

計画の自己レビューで、設計書の完了条件1・6はTask 3のビルドと利用側検証、条件2はTask 3の最終検証、条件3はTask 1、条件4・5はTask 2、条件7はTask 3のサンプル、条件8はTask 3のREADME・Rust削除・配布内容検証に対応していることを確認した。公開署名は設計書を参照し、後続タスクで名前を変更しない。

本計画は公開型を共有する3個の順次タスクである。実装者がこのセッションで全タスクを順に実行するNative方式を推奨する。
サブエージェント方式を選ぶ場合は、各タスクの実装・レビューを分け、Task 1 → Task 2 → Task 3の依存順を維持する。
どちらの場合も最後に全体レビューと設計書の完了監査を行う。
