export type AuthMaterial =
  | { kind: 'cookieHeader'; cookieHeader: string }
  | { kind: 'fedAuthRtFa'; fedAuth: string; rtFa: string };

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
