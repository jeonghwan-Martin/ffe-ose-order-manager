import test from 'node:test';
import assert from 'node:assert/strict';
import {toRow,fromRow,saveOrderItems} from '../src/orderItemsApi.js';
test('vendor, specification, unit, brand and remark survive a round trip',()=>{
 const input={name:'타월',vendorId:'vendor-1',spec:'40x80',unit:'EA',brand:'제공된 브랜드',remark:'확인 필요',qtyPerRoom:6,unitPrice:136.3636363636,actualUnitPrice:0};
 const output=fromRow(toRow(input,{projectId:'p',roomTypeId:'r'}));
 for(const key of Object.keys(input))assert.equal(output[key],input[key],key);
});
test('unmapped room type fails before any network mutation',async()=>{
 const original=globalThis.fetch;let called=false;
 globalThis.fetch=async()=>{called=true;throw new Error('must not request');};
 try{await assert.rejects(saveOrderItems('p',{missing:[{name:'침대'}]},[],{}),/객실 유형 연결/);assert.equal(called,false);}
 finally{globalThis.fetch=original;}
});
