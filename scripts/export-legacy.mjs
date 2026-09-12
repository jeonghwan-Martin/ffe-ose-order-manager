// Read-only extraction from the existing Apps Script API. Never invokes syncNow.
import fs from 'node:fs/promises';
import path from 'node:path';
const [outputPath] = process.argv.slice(2);
if (!outputPath) throw new Error('Usage: node scripts/export-legacy.mjs /private/path/legacy-backup.json');
const source = await fs.readFile(new URL('../src/api.js', import.meta.url), 'utf8');
const url = process.env.LEGACY_APPS_SCRIPT_URL || source.match(/APPS_SCRIPT_URL = "([^"]+)"/)?.[1];
if (!url) throw new Error('LEGACY_APPS_SCRIPT_URL is required');
async function get(action, params = {}) {
  const target = new URL(url);
  target.search = new URLSearchParams({action, ...params});
  const response = await fetch(target, {signal: AbortSignal.timeout(90000)});
  if (!response.ok) throw new Error(`${action}: HTTP ${response.status}`);
  const data = await response.json();
  if (data.error) throw new Error(`${action}: ${data.error}`);
  return data;
}
const startedAt = new Date().toISOString();
const first = (await get('listProjects')).list;
if (!Array.isArray(first) || first.some(x => !x.id || !x.name)) throw new Error('Invalid project index');
if (new Set(first.map(x => x.id)).size !== first.length) throw new Error('Duplicate project IDs in index');
const projects = new Array(first.length);
let next = 0;
await Promise.all(Array.from({length: Math.min(4, first.length)}, async () => {
  while (next < first.length) {
    const i = next++, entry = first[i];
    const {value} = await get('getProject', {id:entry.id});
    if (typeof value !== 'string') throw new Error(`Invalid payload for project ${i + 1}`);
    projects[i] = {...entry, rawValue:value, payload:value ? JSON.parse(value) : null};
    console.log(`Read ${i + 1}/${first.length}: ${value ? 'saved payload' : 'no saved payload'}`);
  }
}));
const catalog = await get('getCatalog');
const last = (await get('listProjects')).list;
if (JSON.stringify(first) !== JSON.stringify(last)) throw new Error('Project index changed during export. Retry.');
const output = {format:'ffe-ose-legacy-export-v1', startedAt, exportedAt:new Date().toISOString(),
  scope:'Indexed projects and item-catalog only; orphan script properties require owner export.',
  consistency:'Read-only API snapshot, not a transaction. Re-export during a save pause before cutover.',
  index:first, projects, catalogRawValue:catalog.value || '', catalog:catalog.value ? JSON.parse(catalog.value) : null};
await fs.mkdir(path.dirname(outputPath), {recursive:true});
await fs.writeFile(outputPath, JSON.stringify(output,null,2), {flag:'wx',mode:0o600});
console.log(JSON.stringify({projects:projects.length,saved:projects.filter(x=>x.payload).length,empty:projects.filter(x=>!x.payload).length}));
