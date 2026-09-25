// Optional Windows check: open a generated DOCX with locally installed Microsoft Word.
import assert from 'node:assert/strict';
import { mkdtemp, writeFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { buildDocx } from './docx.js';

const run = promisify(execFile);
const folder = await mkdtemp(path.join(tmpdir(), 'opencode', 'instruction-word-'));
const documentPath = path.join(folder, 'proverka.docx');
try {
  await writeFile(documentPath, buildDocx([{ text: 'ПРОВЕРКА ИНСТРУКЦИИ', style: 'Title' }, { text: 'Текст & проверка', style: 'Normal' }]));
  const script = `$word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0; try { $doc = $word.Documents.Open('${documentPath.replace(/'/g, "''")}', $false, $true); try { if (-not $doc.Content.Text.Contains('ПРОВЕРКА ИНСТРУКЦИИ')) { throw 'Нет текста документа' }; 'WORD_OK' } finally { $doc.Close($false) } } finally { $word.Quit() }`;
  const { stdout } = await run('powershell.exe', ['-NoProfile', '-Command', script], { timeout: 30000, windowsHide: true });
  assert.match(stdout, /WORD_OK/);
  console.log('Microsoft Word открыл DOCX без ошибок.');
} finally { await rm(folder, { recursive: true, force: true }); }
