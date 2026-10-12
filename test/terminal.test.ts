import assert from 'node:assert/strict';
import { tmpdir } from 'node:os';
import test from 'node:test';
import { setTimeout as delay } from 'node:timers/promises';
import { spawnInTerminal, waitForExit } from './terminal.js';

// Catches killing only the shell while descendants keep inherited pipes open.
test('terminal timeout settles and terminates the command and its descendant', async () => {
  const program = `
    const { spawn } = require('node:child_process');
    process.on('SIGHUP', () => {});
    const child = spawn(process.execPath, ['-e', "process.on('SIGHUP', () => {}); setInterval(() => {}, 1000)"], { stdio: 'inherit' });
    console.log('READY:' + process.pid + ':' + child.pid);
    setInterval(() => {}, 1000);
  `;
  const child = spawnInTerminal([process.execPath, '-e', program], tmpdir());
  const pids: number[] = [];
  let stdout = '';
  const ready = new Promise<void>(resolve => {
    child.stdout.on('data', (chunk: Buffer) => {
      stdout += chunk.toString();
      const match = /READY:(\d+):(\d+)/.exec(stdout);
      if (match) {
        pids.splice(0, pids.length, Number(match[1]), Number(match[2]));
        resolve();
      }
    });
  });
  child.stderr.resume();
  let watchdog: ReturnType<typeof setTimeout> | undefined;
  try {
    await Promise.race([ready, new Promise<never>((_, reject) => {
      watchdog = setTimeout(() => reject(new Error('Terminal fixture did not start')), 2_000);
    })]);
    clearTimeout(watchdog);
    const started = performance.now();
    await assert.rejects(Promise.race([waitForExit(child, 100), new Promise<never>((_, reject) => {
      watchdog = setTimeout(() => reject(new Error('Timeout did not settle')), 1_000);
    })]), /Process timed out/);
    assert.ok(performance.now() - started < 1_000);
    // Allow the OS to reap the killed descendants before checking liveness.
    for (let attempt = 0; attempt < 50; attempt++) {
      if (pids.every(pid => {
        try { process.kill(pid, 0); return false; } catch { return true; }
      })) break;
      await delay(10);
    }
    for (const pid of pids) assert.throws(() => process.kill(pid, 0), { code: 'ESRCH' });
  } finally {
    clearTimeout(watchdog);
    // Independently clean up even when the regression is present.
    for (const pid of pids.reverse()) {
      try { process.kill(pid, 'SIGKILL'); } catch { /* Already gone. */ }
    }
    try { process.kill(-child.pid!, 'SIGKILL'); } catch { /* Already gone. */ }
    child.stdin.destroy();
    child.stdout.destroy();
    child.stderr.destroy();
  }
});
