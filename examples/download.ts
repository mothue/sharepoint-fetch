import { writeFile } from 'node:fs/promises';
import { createInterface } from 'node:readline';
import { stdin, stdout } from 'node:process';
import { Writable } from 'node:stream';
import { buildClient, downloadFile, SharePointTarget, type AuthMaterial } from '../src/index.js';

async function main(): Promise<void> {
  const terminal = stdin.isTTY === true;
  let secret = false;
  // Readline controls terminal echo in raw mode; suppress its writes for credentials.
  const output = new Writable({
    write(chunk, _encoding, callback) {
      if (!secret) stdout.write(chunk);
      callback();
    },
  });
  const input = createInterface({ input: stdin, output, terminal });
  // Register the iterator before asking, so piped lines are queued even at EOF.
  const lines = input[Symbol.asyncIterator]();
  async function ask(prompt: string, hidden = false): Promise<string> {
    secret = hidden;
    stdout.write(prompt);
    try {
      const line = await lines.next();
      if (line.done) throw new Error('Input ended before all answers were provided');
      return line.value;
    } finally {
      if (terminal && hidden) stdout.write('\n');
      secret = false;
    }
  }
  try {
    const target = SharePointTarget.fromDocumentUrl(await ask('Document URL: '));
    const choice = (await ask('Authentication (1: Cookie header, 2: FedAuth/rtFa): ')).trim();
    let auth: AuthMaterial;
    if (choice === '1') {
      auth = { kind: 'cookieHeader', cookieHeader: await ask('Cookie header: ', true) };
    } else if (choice === '2') {
      auth = { kind: 'fedAuthRtFa', fedAuth: await ask('FedAuth: ', true), rtFa: await ask('rtFa: ', true) };
    } else {
      throw new Error('Choose authentication mode 1 or 2');
    }
    const client = buildClient(auth, target.siteUrl);
    const response = await downloadFile(client, target);
    if (!response.ok) {
      await response.body?.cancel();
      throw new Error(`Download failed: HTTP ${response.status}`);
    }
    const cleaned = target.label.replace(/[\/\\:*?"<>|\x00-\x1f\x7f]/g, '').trim();
    const filename = !cleaned || cleaned === '.' || cleaned === '..' ? target.itemId : cleaned;
    await writeFile(filename, new Uint8Array(await response.arrayBuffer()));
    stdout.write(`Saved ${filename}\n`);
  } finally {
    input.close();
  }
}

main().catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : 'Download failed');
  process.exitCode = 1;
});
