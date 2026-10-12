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

  constructor(init: SharePointTargetInit) {
    const siteUrl = parseHttpUrl(init.siteUrl);
    if (!siteUrl) {
      throw new Error('siteUrl must be an absolute HTTP or HTTPS URL');
    }
    const itemId = normalizeGuid(init.itemId);
    if (!itemId) {
      throw new Error('itemId must be a GUID in 8-4-4-4-12 hexadecimal format');
    }
    siteUrl.search = '';
    siteUrl.hash = '';
    this.siteUrl = siteUrl.href.replace(/\/+$/, '');
    this.itemId = itemId;
    this.label = init.label ?? itemId;
    this.documentUrl = init.documentUrl;
  }

  static fromDocumentUrl(documentUrl: string): SharePointTarget {
    const queue = [documentUrl];
    const seen = new Set<string>();

    for (let index = 0; index < queue.length; index++) {
      const raw = trimCandidate(queue[index]);
      if (!raw || seen.has(raw)) continue;
      seen.add(raw);
      const url = parseHttpUrl(raw);
      if (!url) continue;

      const siteUrl = inferSiteUrl(url);
      let itemId: string | undefined;
      let label: string | undefined;
      for (const [key, value] of url.searchParams) {
        if (key.toLowerCase() === 'sourcedoc' && !itemId) {
          itemId = normalizeGuid(value);
        }
        if (key.toLowerCase() === 'file' && !label) {
          label = value.trim() || undefined;
        }
      }
      if (siteUrl && itemId) {
        return new SharePointTarget({ siteUrl, itemId, label, documentUrl });
      }

      for (const value of url.searchParams.values()) {
        enqueueNestedUrls(queue, value);
      }
      enqueueNestedUrls(queue, url.hash.slice(1));
    }

    throw new Error('documentUrl must contain a valid SharePoint HTTP or HTTPS URL with a sourcedoc GUID');
  }

  withLabel(label: string): SharePointTarget {
    return new SharePointTarget({
      siteUrl: this.siteUrl,
      itemId: this.itemId,
      label,
      documentUrl: this.documentUrl,
    });
  }

  downloadUrl(): string {
    return `${this.siteUrl}/_api/web/GetFileById(guid'${this.itemId}')/$value`;
  }
}

function trimCandidate(raw: string): string {
  return raw.trim().replace(/^"+|"+$/g, '').trim();
}

function parseHttpUrl(raw: string): URL | undefined {
  try {
    const url = new URL(raw);
    return url.protocol === 'http:' || url.protocol === 'https:' ? url : undefined;
  } catch {
    return undefined;
  }
}

function normalizeGuid(raw: string): string | undefined {
  const value = raw.trim().replace(/^[{}]+|[{}]+$/g, '');
  return /^[\da-f]{8}-[\da-f]{4}-[\da-f]{4}-[\da-f]{4}-[\da-f]{12}$/i.test(value)
    ? value.toUpperCase()
    : undefined;
}

function inferSiteUrl(url: URL): string | undefined {
  const segments = url.pathname.split('/');
  const scopeIndex = segments.findIndex(segment =>
    segment === 'sites' || segment === 'teams' || segment === 'personal');
  if (scopeIndex !== -1) {
    const name = segments[scopeIndex + 1];
    return name ? `${url.origin}/${segments[scopeIndex]}/${name}` : undefined;
  }
  return url.hostname === 'sharepoint.com' || url.hostname.endsWith('.sharepoint.com')
    ? url.origin
    : undefined;
}

function enqueueNestedUrls(queue: string[], raw: string): void {
  let current = trimCandidate(raw);
  for (let depth = 0; depth <= 3; depth++) {
    if (/^https?:\/\//i.test(current)) queue.push(current);
    if (depth === 3) break;
    try {
      const decoded = trimCandidate(decodeURIComponent(current));
      if (decoded === current) break;
      current = decoded;
    } catch {
      // A malformed sibling must not prevent traversal of other URL values.
      break;
    }
  }
}
