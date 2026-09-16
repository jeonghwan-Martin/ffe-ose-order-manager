import fs from 'node:fs/promises';
import {planLegacyImport} from '../src/legacyMigration.js';
const [backupPath,outputPath]=process.argv.slice(2);
if(!backupPath||!outputPath)throw new Error('Usage: node scripts/audit-legacy.mjs backup.json audit.json');
const code=await fs.readFile(new URL('../src/projectIdApi.js',import.meta.url),'utf8');
const url=process.env.SUPABASE_URL||code.match(/SUPABASE_URL = "([^"]+)"/)[1];
const key=process.env.SUPABASE_ANON_KEY||code.match(/SUPABASE_ANON_KEY =\s*"([^"]+)"/)[1];
const res=await fetch(`${url}/rest/v1/projects?select=id,name,client_id&order=id`,{headers:{apikey:key,Authorization:`Bearer ${key}`},signal:AbortSignal.timeout(90000)});
if(!res.ok)throw new Error(`Project lookup failed: ${res.status}`);
// Require an explicit small result; avoid silently treating a paginated result as complete.
const projects=await res.json();if(projects.length>=1000)throw new Error('Pagination required');
const backup=JSON.parse(await fs.readFile(backupPath,'utf8'));
const plan=planLegacyImport(backup,projects);
const summary=plan.reduce((a,x)=>(a[x.status]=(a[x.status]||0)+1,a),{});
const audit={checkedAt:new Date().toISOString(),summary,sourceProjects:backup.projects.length,targetProjects:projects.length,
  roomTypes:plan.reduce((n,p)=>n+p.roomTypes,0),orderItems:plan.reduce((n,p)=>n+p.orderItems,0),
  projects:plan.map(({payload,...p})=>p)};
await fs.writeFile(outputPath,JSON.stringify(audit,null,2),{flag:'wx',mode:0o600});
console.log(JSON.stringify({...audit,projects:undefined}));
