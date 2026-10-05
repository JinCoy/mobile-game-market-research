import {spawnSync} from 'node:child_process';
import {homedir} from 'node:os';
import {join} from 'node:path';

const candidates = [process.env.PLAYFIELD_PYTHON, join(process.cwd(), '.venv/bin/python'), 'python3',
  join(homedir(), '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3')].filter(Boolean);
const python = candidates.find(candidate => spawnSync(candidate, ['-c', 'import openpyxl'], {stdio:'ignore'}).status === 0);
if (!python) {
  console.error('Python with openpyxl is required. Install requirements.txt in a virtual environment, or set PLAYFIELD_PYTHON.');
  process.exit(1);
}
const result = spawnSync(python, ['scripts/import_data.py', ...process.argv.slice(2)], {stdio:'inherit'});
process.exit(result.status ?? 1);
