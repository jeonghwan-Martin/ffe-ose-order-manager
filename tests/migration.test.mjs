import test from 'node:test';
import assert from 'node:assert/strict';
import {planLegacyImport} from '../src/legacyMigration.js';
import {createProjectDocumentStore} from '../src/projectDocumentStore.js';
const payload={roomTypes:[{id:1,byFloor:{'2층':8}}],ffeItems:{1:[{id:2,unitPrice:136.3636363636,vendorId:'vendor-1',cartonSize:20}]},oseItems:[]};
const snapshot=(entries)=>({format:'ffe-ose-legacy-export-v1',projects:entries});
test('matches explicit IDs and preserves every payload field and precision',()=>{
 const result=planLegacyImport(snapshot([{id:'ov-abc',name:'old name',payload}]),[{id:'abc',name:'new name'}]);
 assert.equal(result[0].status,'ready');assert.deepEqual(result[0].payload,payload);assert.equal(result[0].orderItems,1);
});
test('does not guess similar names or replace missing payload with empty data',()=>{
 const r=planLegacyImport(snapshot([{id:'a',name:'대전',payload},{id:'b',name:'수원',payload:null}]),[{id:'c',name:'대전 호텔'}]);
 assert.equal(r[0].status,'unmatched');assert.equal(r[1].status,'no-saved-payload');
});
test('duplicate destinations and orphan room items require review',()=>{
 const r=planLegacyImport(snapshot([{id:'a',name:'same',payload},{id:'b',name:'same',payload}]),[{id:'c',name:'same'}]);
 assert.ok(r.every(x=>x.status==='duplicate-target'));
 const invalid=planLegacyImport(snapshot([{id:'a',name:'a',payload:{...payload,roomTypes:[]}}]),[]);
 assert.equal(invalid[0].status,'invalid');
});
test('empty response is migration-required, never an empty editable project',async()=>{
 const store=createProjectDocumentStore({url:'https://test.invalid',key:'test',fetchImpl:async()=>new Response('[]')});
 await assert.rejects(store.load('a'),/이전이 아직/);
});
test('version conflict is surfaced without overwriting or retrying',async()=>{
 let calls=0;
 const store=createProjectDocumentStore({url:'https://test.invalid',key:'test',fetchImpl:async(_,init)=>{
   calls++;assert.equal(JSON.parse(init.body).p_expected_version,3);
   return new Response(JSON.stringify({code:'40001'}),{status:409});
 }});
 await assert.rejects(store.save('a',payload,3),/다른 팀원/);assert.equal(calls,1);
 await assert.rejects(store.save('a',payload,undefined),/최신 데이터/);assert.equal(calls,1);
});
