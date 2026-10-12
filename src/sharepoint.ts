import type { DownloadOptions, SharePointClient } from './models.js';
import type { SharePointTarget } from './target.js';

export function downloadFile(client: SharePointClient, target: SharePointTarget, options?: DownloadOptions): Promise<Response> {
  return client.get(target.downloadUrl(), options);
}
