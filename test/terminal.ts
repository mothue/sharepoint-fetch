import { execFileSync, spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';

export function spawnInTerminal(args: string[], cwd: string): ChildProcessWithoutNullStreams {
  // Own a process group so a terminal command and its pipeline can be torn down together.
  return process.platform === 'darwin'
    // macOS script accepts a pipe but cannot read Node's socket-backed stdin.
    ? spawn('/bin/sh', ['-c', 'cat | script -q /dev/null "$@"', 'sh', ...args], { cwd, detached: true })
    : spawn('script', ['-q', '-e', '-c', args.map(arg => `'${arg.replaceAll("'", "'\\''")}'`).join(' '), '/dev/null'], { cwd, detached: true });
}

function terminate(child: ChildProcessWithoutNullStreams): void {
  const descendants: number[] = [];
  // script creates a new terminal process group, so killing only our group is insufficient.
  try {
    const rows = execFileSync('ps', ['-axo', 'pid=,ppid='], { encoding: 'utf8', timeout: 250 })
      .trim().split('\n').map(line => line.trim().split(/\s+/).map(Number));
    const parents = [child.pid!];
    for (let index = 0; index < parents.length; index++) {
      for (const [pid, parent] of rows) if (parent === parents[index]) {
        descendants.push(pid!);
        parents.push(pid!);
      }
    }
  } finally {
    for (const pid of descendants.reverse()) {
      try { process.kill(pid, 'SIGKILL'); } catch { /* Already gone. */ }
    }
    try { if (child.pid) process.kill(-child.pid, 'SIGKILL'); } catch { /* Already gone. */ }
    // Settle even if a descendant retained a pipe outside the captured tree.
    child.stdin.destroy();
    child.stdout.destroy();
    child.stderr.destroy();
  }
}

export function waitForExit(child: ChildProcessWithoutNullStreams, timeoutMs = 8_000): Promise<number | null> {
  return new Promise((resolve, reject) => {
    const cleanup = () => {
      clearTimeout(timeout);
      child.off('close', onClose);
      child.off('error', onError);
    };
    const onClose = (code: number | null) => { cleanup(); resolve(code); };
    const onError = (error: Error) => { cleanup(); reject(error); };
    const timeout = setTimeout(() => {
      try {
        terminate(child);
      } catch {
        // Cleanup and bounded settlement still apply if process inspection fails.
      } finally {
        cleanup();
        reject(new Error(`Process timed out after ${timeoutMs}ms`));
      }
    }, timeoutMs);
    child.once('close', onClose);
    child.once('error', onError);
  });
}
