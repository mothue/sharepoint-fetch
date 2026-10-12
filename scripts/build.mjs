import { execFile } from 'node:child_process';
import { cp, mkdir, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { promisify } from 'node:util';
import { fileURLToPath } from 'node:url';

const root = fileURLToPath(new URL('../', import.meta.url));
const compiler = fileURLToPath(import.meta.resolve('typescript/bin/tsc'));
const execute = promisify(execFile);
const temporary = join(root, '.build/cjs');
async function compile(config) {
  try {
    await execute(process.execPath, [compiler, '-p', config], { cwd: root });
  } catch (error) {
    throw new Error(`TypeScript build failed\n${error.stdout ?? ''}${error.stderr ?? ''}`, { cause: error });
  }
}

try {
  await rm(join(root, 'dist'), { recursive: true, force: true });
  await rm(temporary, { recursive: true, force: true });
  await compile('tsconfig.esm.json');
  await cp(join(root, 'src'), temporary, { recursive: true });
  await writeFile(join(temporary, 'package.json'), JSON.stringify({ type: 'commonjs' }));
  await compile('tsconfig.cjs.json');
  await mkdir(join(root, 'dist/cjs'), { recursive: true });
  await writeFile(join(root, 'dist/cjs/package.json'), JSON.stringify({ type: 'commonjs' }) + '\n');
  // Keep stdout available for npm pack --json, including its prepack lifecycle.
  console.error('Built ESM and CommonJS JavaScript, declarations, and source maps');
} finally {
  await rm(temporary, { recursive: true, force: true });
}
