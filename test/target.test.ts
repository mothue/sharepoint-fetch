import assert from 'node:assert/strict';
import test from 'node:test';
import { SharePointTarget } from '../src/target.js';

const guid = '01234567-89ab-cdef-0123-456789abcdef';
const normalized = '01234567-89AB-CDEF-0123-456789ABCDEF';
const input = `https://tenant.sharepoint.com/:x:/r/sites/demo/_layouts/15/Doc.aspx?sourcedoc=%7B${guid}%7D&file=example.xlsx`;

// Catches incorrect extraction, normalization, REST composition, or original URL loss.
test('parses document links', () => {
  const target = SharePointTarget.fromDocumentUrl(input);
  assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo');
  assert.equal(target.itemId, normalized);
  assert.equal(target.label, 'example.xlsx');
  assert.equal(target.documentUrl, input);
  assert.equal(target.downloadUrl(), "https://tenant.sharepoint.com/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
  const quoted = `  "${input}"  `;
  assert.equal(SharePointTarget.fromDocumentUrl(quoted).siteUrl, target.siteUrl);
  assert.equal(SharePointTarget.fromDocumentUrl(quoted).documentUrl, quoted);
});

// Catches a missing scope branch or failure to accept a root site.
test('supports site scopes and root sites', () => {
  for (const scope of ['sites', 'teams', 'personal']) {
    const target = SharePointTarget.fromDocumentUrl(`https://tenant.sharepoint.com/${scope}/demo/Doc.aspx?sourcedoc=${guid}`);
    assert.equal(target.siteUrl, `https://tenant.sharepoint.com/${scope}/demo`);
  }
  assert.equal(SharePointTarget.fromDocumentUrl(`https://tenant.sharepoint.com/Doc.aspx?sourcedoc=${guid}`).siteUrl, 'https://tenant.sharepoint.com');
  assert.equal(SharePointTarget.fromDocumentUrl(`http://sharepoint.com:8080/Doc.aspx?sourcedoc=${guid}`).siteUrl, 'http://sharepoint.com:8080');
});

// Catches incorrectly imposing the root-site hostname rule on scoped HTTP sites.
test('supports scoped document links on local HTTP hosts', () => {
  for (const scope of ['sites', 'teams', 'personal']) {
    const target = SharePointTarget.fromDocumentUrl(`http://127.0.0.1:8080/${scope}/demo/Doc.aspx?sourcedoc=${guid}&file=local.xlsx`);
    assert.equal(target.siteUrl, `http://127.0.0.1:8080/${scope}/demo`);
    assert.equal(target.itemId, normalized);
    assert.equal(target.label, 'local.xlsx');
  }
});

// Catches case-sensitive lookup or retaining an empty label.
test('supports query casing and fallback labels', () => {
  assert.equal(SharePointTarget.fromDocumentUrl(`https://tenant.sharepoint.com/?SOURCEDOC=${guid}&FILE=Example.xlsx`).label, 'Example.xlsx');
  for (const suffix of ['', '&file=', '&file=%20%20']) {
    assert.equal(SharePointTarget.fromDocumentUrl(`https://tenant.sharepoint.com/?sourcedoc=${guid}${suffix}`).label, normalized);
  }
});

// Catches omitted query/fragment traversal or insufficient decoding depth.
test('follows nested query and fragment URLs', () => {
  let nested = input;
  for (let depth = 0; depth <= 3; depth++) {
    for (const wrapper of [`https://wrapper.example/?url=${nested}`, `https://wrapper.example/#${nested}`]) {
      const target = SharePointTarget.fromDocumentUrl(wrapper);
      assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo');
      assert.equal(target.itemId, normalized);
    }
    nested = encodeURIComponent(nested);
  }
});

// Catches combining the wrapper's site with the nested candidate's item or label.
test('keeps nested site and item together', () => {
  const wrapper = `https://tenant.sharepoint.com/sites/wrapper?file=wrong.xlsx&url=${encodeURIComponent(input)}`;
  const target = SharePointTarget.fromDocumentUrl(wrapper);
  assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo');
  assert.equal(target.itemId, normalized);
  assert.equal(target.label, 'example.xlsx');
});

// Catches aborting sibling traversal on decoding errors or corrupting UTF-8.
test('preserves Unicode with malformed sibling input', () => {
  const nested = `https://tenant.sharepoint.com/sites/demo/Doc.aspx?sourcedoc=${guid}&file=報告書.xlsx`;
  const wrapper = `https://wrapper.example/?broken=%ZZ&url=${encodeURIComponent(encodeURIComponent(nested))}`;
  assert.equal(SharePointTarget.fromDocumentUrl(wrapper).label, '報告書.xlsx');
});

// Catches missing input, protocol, scope, GUID, or domain-boundary validation.
test('rejects invalid targets', () => {
  for (const invalid of [
    'https://tenant.sharepoint.com/sites/demo/Doc.aspx',
    'https://tenant.sharepoint.com/?sourcedoc=invalid',
    `https://tenant.sharepoint.com/?sourcedoc=${guid}0`,
    `/sites/demo/Doc.aspx?sourcedoc=${guid}`,
    `ftp://tenant.sharepoint.com/?sourcedoc=${guid}`,
    `https://tenant.sharepoint.com/sites/?sourcedoc=${guid}`,
    `https://tenant.sharepoint.com/sites//demo?sourcedoc=${guid}`,
    `https://notsharepoint.com/Doc.aspx?sourcedoc=${guid}`,
    `https://tenant.sharepoint.com.evil.example/Doc.aspx?sourcedoc=${guid}`,
    `https://example.com/Doc.aspx?sourcedoc=${guid}`,
  ]) {
    assert.throws(() => SharePointTarget.fromDocumentUrl(invalid), Error, invalid);
  }
});

// Catches missing constructor validation/normalization or mutating label copies.
test('creates explicit targets and copies labels', () => {
  const target = new SharePointTarget({ siteUrl: 'https://tenant.sharepoint.com/sites/demo///', itemId: ` {${guid}} `, label: 'original.xlsx', documentUrl: input });
  assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo');
  assert.equal(target.itemId, normalized);
  assert.equal(target.downloadUrl(), "https://tenant.sharepoint.com/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value");
  const copy = target.withLabel('renamed.xlsx');
  assert.equal(copy.label, 'renamed.xlsx');
  assert.equal(target.label, 'original.xlsx');
  assert.equal(copy.siteUrl, target.siteUrl);
  assert.equal(copy.itemId, target.itemId);
  assert.equal(copy.documentUrl, input);
  const fallback = new SharePointTarget({ siteUrl: 'http://localhost:8080/', itemId: guid });
  assert.equal(fallback.label, normalized);
  assert.equal(fallback.documentUrl, undefined);
  for (const siteUrl of ['/sites/demo', 'ftp://tenant.sharepoint.com/', 'invalid']) {
    assert.throws(() => new SharePointTarget({ siteUrl, itemId: guid }), Error);
  }
  assert.throws(() => new SharePointTarget({ siteUrl: 'https://tenant.sharepoint.com', itemId: 'invalid' }), Error);
});

// Catches query/fragment text absorbing the REST suffix, including after a slash.
for (const suffix of ['?web=1', '#overview', '/?web=1', '/#overview']) {
  test(`normalizes explicit site page URLs to a site base: ${suffix}`, () => {
    const target = new SharePointTarget({ siteUrl: `https://tenant.sharepoint.com/sites/demo${suffix}`, itemId: guid });
    assert.equal(target.siteUrl, 'https://tenant.sharepoint.com/sites/demo', suffix);
    assert.equal(target.downloadUrl(), "https://tenant.sharepoint.com/sites/demo/_api/web/GetFileById(guid'01234567-89AB-CDEF-0123-456789ABCDEF')/$value", suffix);
  });
}
