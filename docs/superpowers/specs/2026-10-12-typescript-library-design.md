# sharepoint-fetch TypeScriptライブラリ設計

## 目的と承認済みの方針

ユーザーの目的は「TypeScriptのライブラリとして使えるようにRewriteする」こと。
既存のRustクレートを、Rustの実行環境を必要としないTypeScript製npmパッケージに置き換える。
URL解析、Cookie認証、ファイル取得、対話式サンプルを移植し、ESM・CommonJSの両方に対応する型定義を提供する。
この方針は2026-10-12にユーザーが「いいよ」と承認した。

## 現在の実装と移植範囲

- `src/sharepoint.rs`：ドキュメントURLの解析、GUIDの正規化、REST API URL生成、ダウンロード。
- `src/auth.rs`・`src/models.rs`：CookieヘッダーまたはFedAuth/rtFaによるクライアント生成。
- `src/config.rs`：既定のUser-Agent `sharepoint-fetch-bot`。
- `src/main.rs`：URLと認証情報を対話入力し、ファイル名を整えて保存するサンプル。
- `src/lib.rs`：公開APIのエントリーポイント。

移植完了時にRustソース、`Cargo.toml`、`Cargo.lock`を削除する。
既存の`LICENSE`は維持し、READMEはTypeScript版に置き換える。
ライブラリを読み込むだけでは、ネットワーク通信、対話入力、ファイル保存を行わない。

## 実行環境と配布

想定する利用環境はNode.js 22以上。ブラウザからのCookieヘッダー設定を前提とした動作は対象外。
パッケージ名は`sharepoint-fetch`、初期バージョンは既存の`0.1.0`を維持する。

- TypeScriptのstrictチェックを使用する。
- 通信にはNode.js標準の`fetch`、`URL`、`Headers`、`Response`を使用する。
- Cookie管理の実行時依存は`tough-cookie`に限定する。
- 開発用にはTypeScript、Node.jsの型定義、TypeScriptサンプル・テストの実行用`tsx`を使用する。
- `package.json`の条件付き`exports`で、ESM・CommonJSそれぞれのJavaScriptと型定義を指定する。
- `dist/esm`はESM、`dist/cjs`はCommonJSとして解釈される構成にする。
- npmパッケージにはビルド成果物、TypeScriptソース、README、LICENSEを含める。
- ビルド・型チェック・テスト・サンプル実行をnpm scriptsで提供する。

Rustを残したラッパー方式は、利用者にRustまたはバイナリ配布が必要になるため採用しない。
RustとTypeScriptの併存方式も、同じ機能を二重に保守することになるため採用しない。

## 公開API

Rustの名前をTypeScriptのcamelCaseへ変換する。

```ts
export type AuthMaterial =
  | { kind: 'cookieHeader'; cookieHeader: string }
  | { kind: 'fedAuthRtFa'; fedAuth: string; rtFa: string };

export interface SharePointTargetInit {
  siteUrl: string;
  itemId: string;
  label?: string;
  documentUrl?: string;
}

export class SharePointTarget {
  readonly siteUrl: string;
  readonly itemId: string;
  readonly label: string;
  readonly documentUrl: string | undefined;
  constructor(init: SharePointTargetInit);
  static fromDocumentUrl(documentUrl: string): SharePointTarget;
  withLabel(label: string): SharePointTarget;
  downloadUrl(): string;
}

export interface ClientOptions {
  fetch?: typeof globalThis.fetch;
  userAgent?: string;
  maxRedirects?: number;
}

export interface DownloadOptions {
  signal?: AbortSignal;
}

export interface SharePointClient {
  get(url: string | URL, options?: DownloadOptions): Promise<Response>;
}

export function buildClient(
  auth: AuthMaterial,
  siteUrl: string,
  options?: ClientOptions,
): SharePointClient;

export function downloadFile(
  client: SharePointClient,
  target: SharePointTarget,
  options?: DownloadOptions,
): Promise<Response>;
```

`SharePointTargetInit`による明示的なターゲット作成も、従来のRustの公開フィールドによる作成に相当する機能として提供する。
`withLabel`は元のターゲットを変更せず、新しいターゲットを返す。
`buildClient`は同期的に認証情報を検証してCookie jarを準備する。
`downloadFile`の戻り値は未消費の標準`Response`である。

## URL解析とターゲット検証

入力の前後の空白と囲みのダブルクォートを除去し、絶対HTTP/HTTPS URLを候補として扱う。
クエリ値とフラグメントに含まれるURLを幅優先で探索し、同じ候補は再処理しない。
入れ子のURL値について、既存コードと同様に複数段のパーセントエンコードを扱う。
デコードで日本語などのUnicode文字列を壊さず、不正なエンコードでも探索処理全体を停止しない。

- `sites/<name>`、`teams/<name>`、`personal/<name>`からサイトURLを推定する。
- 上記のパスがないSharePointドメインでは、サイトURLをオリジンとする。
- SharePointドメインの判定は`sharepoint.com`自身またはそのサブドメインへの一致を使う。
- `sourcedoc`と`file`のクエリキーは大文字小文字を区別しない。
- GUIDは前後の空白と波括弧を除去し、`8-4-4-4-12`の16進数形式を検証して大文字へ正規化する。
- サイトURLとGUIDは同じURL候補から採用し、ラッパーURLのサイトと別URLのGUIDを組み合わせない。
- ラベルは採用した候補の空でない`file`値を使用し、なければGUIDを使用する。
- `documentUrl`には呼び出し元が渡した元の文字列を保存する。

明示的なターゲット作成でも、HTTP/HTTPSサイトURLとGUIDを検証する。
サイトURLの末尾のスラッシュを除去し、REST URLは次の形にする。

```text
<siteUrl>/_api/web/GetFileById(guid'<itemId>')/$value
```

## Cookie認証とHTTP通信

Cookieヘッダーはセミコロンで分割し、最初の`=`で名前と値を分ける。
空の名前・値、`=`のない要素は無視するが、有効なCookieが一つもなければエラーにする。
FedAuth/rtFa方式は両方の空でない値を必要とする。
HTTPヘッダーまたはCookieを別の値として注入できる改行や不正な名前を受け付けない。
Cookie値に含まれる`=`は維持する。

初期Cookieは渡されたサイトのホストに限定し、`Path=/`として登録する。
HTTPSサイトの初期CookieにはSecure属性を付ける。
応答の`Set-Cookie`はCookie jarに反映し、domain・path・secure・expiry属性に従って後続リクエストへ適用する。
手動リダイレクトにより、各遷移先についてCookie jarからCookieを選び直す。
初期認証情報を無関係なホストへ転送しない。

- GETでファイルを取得する。
- User-Agentの既定値は`sharepoint-fetch-bot`とする。
- 301・302・303・307・308のLocationをたどる。相対Locationも扱う。
- 既定のリダイレクト上限は20回で、変更値は0以上の整数に限定する。
- Locationのない応答は、そのまま呼び出し元へ返す。
- リダイレクト前の応答ボディは解放し、通信失敗または上限超過はPromiseのrejectとする。
- `AbortSignal`を通信に渡し、ダウンロードのキャンセルを可能にする。
- 401・403・404などのHTTP応答はthrowせず、未消費の`Response`として返す。
- 認証情報をログやエラーメッセージへ埋め込まない。

`fetch`の注入はテストと利用者の通信設定のために提供する。
標準の通信失敗・キャンセルを識別できるようにし、URL解析や入力検証には具体的な`Error`メッセージを付ける。

## ファイル構成

- `src/index.ts`：値と型の公開エクスポート。
- `src/models.ts`：認証情報、クライアント・ダウンロードオプションの型。
- `src/target.ts`：ターゲットの作成、URL解析、GUIDの検証。
- `src/auth.ts`：Cookie jarと認証付きHTTPクライアント。
- `src/sharepoint.ts`：ターゲットからのダウンロード。
- `examples/download.ts`：既存の対話入力・ファイル保存サンプルの移植。
- `test/`：URL解析、認証、HTTP、公開API、配布パッケージの検証。
- `scripts/`：必要なビルドと配布検証の補助処理。
- `package.json`・ロックファイル・TypeScript設定：再現可能なインストールとビルド。

## サンプルとREADME

READMEには動作環境、インストール・開発コマンド、公開APIとRust版からの名前の対応を記載する。
認証の2方式、URLからのターゲット生成、明示的なターゲット生成、`Response.ok`の判定、バイナリ保存の例を掲載する。
対話式サンプルはURLと認証方式・値を受け取り、ファイル名から区切り文字とOS上の禁止文字を除去して保存する。
空または`.`・`..`になる名前にはGUIDを使用する。
失敗時はファイルを保存せず、HTTPステータスまたは通信エラーを表示する。
サンプル以外でファイルシステムや対話入力を呼び出さない。

## 完了条件と検証

1. TypeScriptソースからESM・CommonJS両方のJavaScriptと型定義を生成できる。
2. `npm ci`で依存関係を再現でき、型チェックとテストが成功する。
3. 代表的なSharePoint URL、sites/teams/personal、ルートサイト、ネストしたクエリ・フラグメント、GUIDの不正入力、日本語ファイル名をテストする。
4. Cookieの2方式、値内の`=`、無効な認証情報、Set-Cookieの更新、Cookieのスコープ、リダイレクト上限をテストする。
5. ローカルHTTPサーバーでGETのURL・ヘッダー、バイナリ本文、HTTPエラー、リダイレクト、通信失敗・キャンセルを確認する。
6. `npm pack`した成果物を一時的な利用側プロジェクトへインストールし、ESM・CommonJSからの実行と両方のTypeScript型解決を確認する。
7. サンプルをローカルHTTPサーバーで実行し、保存内容・保存ファイル名・HTTPエラー時の処理を確認する。
8. READMEの例が公開APIと一致し、不要なRust成果物が配布パッケージに含まれない。

実際のSharePointテナントへの通信は、利用可能な認証情報がないため完了検証には使用しない。
npmへの公開やリモートリポジトリへのpushはこの移植作業の範囲に含めない。

## 実装前に確認した技術的前提

2026-10-12に以下を確認した。これは実装・動作検証の代わりではなく、設計の依存先と検証対象を明確にするための調査である。

- [Node.js 22.0.0の公式資料](https://nodejs.org/download/release/v22.0.0/docs/api/globals.html#fetch)：標準の`fetch`・`Headers`・`Response`は安定版APIとして提供されている。
- `npm view tough-cookie version engines exports --json`：公開版は`6.0.2`、Node.jsの要求は`>=16`、ESM・CommonJSそれぞれに実装と型定義のエクスポートがある。Node.js 22以上を前提にする設計と整合している。
- [CookieJarの公式API](https://github.com/salesforce/tough-cookie/blob/master/api/docs/tough-cookie.cookiejar.md)：同期Cookie storeで`setCookieSync`・`getCookieStringSync`を利用できるため、`buildClient`を同期APIとして提供できる。
- [tough-cookieの公式README](https://github.com/salesforce/tough-cookie)：既定ではlocalhostをSecure Cookieの適用先として扱う。設計のsecure属性に従う検証では`allowSecureOnLocal: false`を明示し、HTTPSからHTTPへの変更で初期Secure Cookieを送信しないことを確認する。
- [Undiciの公式Headers API](https://github.com/nodejs/undici/blob/main/docs/docs/api/Fetch.md#headersgetsetcookie)：`getSetCookie()`は個々のSet-Cookieを配列として返す。CookieのExpires属性にはカンマが含まれるため、結合されたヘッダーをカンマで分割する実装は使用しない。実装時にはNode.js 22でこのAPIの動作も確認する。
