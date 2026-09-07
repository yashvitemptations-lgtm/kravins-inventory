// ════════════════════════════════════════════════════════════════
//  KRAVINS INVENTORY — Google Apps Script  (v4)
//  Paste into Extensions > Apps Script in your Google Sheet
//  Deploy as Web App: Execute as Me | Who has access: Anyone
// ════════════════════════════════════════════════════════════════

const SHEET_ID = '1PU6Z0R7yHMx8UT61zYGqPUt4W9Toxg0PXU81EMFggSw';

const TABS = {
  STATE:        'State',
  TRANSACTIONS: 'Transactions',
  FIFO_RM:      'FIFO_RM',       // old RM (flavour-based) — kept for safety
  FIFO_RM_ITEM: 'FIFO_RMItem',   // NEW: individual RM ingredients
  FIFO_PP:      'FIFO_PP',       // NEW: powder & paste
  FIFO_SFG1:    'FIFO_SFG1',     // NEW: loose candy
  FIFO_BM:      'FIFO_BM',       // NEW: batch mix store
  FIFO_SFG:     'FIFO_SFG',      // SFG-2 pouches
  FIFO_FG:      'FIFO_FG',
  BATCH_SEQ:    'BatchSeq',
};

// ════════════════════════════════════════════════════════════════
//  ENTRY POINT — GET with base64 encoded payload
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.p)
      return respond({ ok:true, message:'Kravins API v4 running', ts:Date.now() });
    const bytes = Utilities.base64Decode(e.parameter.p);
    const decoded = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    return respond(route(JSON.parse(decoded)));
  } catch(err) { return respond({ ok:false, error:err.toString() }); }
}

function doPost(e) {
  try { return respond(route(JSON.parse(e.postData.contents))); }
  catch(err) { return respond({ ok:false, error:err.toString() }); }
}

function route(p) {
  if (p.action==='load')         return loadAll();
  if (p.action==='saveState')    return saveState(p.data);
  if (p.action==='addTxn')       return addTransaction(p.data);
  if (p.action==='addFIFO')      return addFIFOBatch(p.kind, p.data);
  if (p.action==='consumeFIFO')  return consumeFIFOBatch(p.kind, p.id, p.qty);
  if (p.action==='nextBatchSeq') return nextBatchSeq(p.key);
  if (p.action==='initSheets')   return initSheets();
  if (p.action==='setBatchSeq')  return setBatchSeq(p.key, p.value);
  return { ok:false, error:'Unknown action: '+p.action };
}

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ════════════════════════════════════════════════════════════════
function ss()          { return SpreadsheetApp.openById(SHEET_ID); }
function getSheet(name){ let sh=ss().getSheetByName(name); if(!sh)sh=ss().insertSheet(name); return sh; }

function sheetToObjects(sh) {
  const data=sh.getDataRange().getValues();
  if(data.length<2) return [];
  const H=data[0].map(String);
  return data.slice(1).map(row=>{ const o={}; H.forEach((h,i)=>o[h]=row[i]); return o; });
}

function ensureHeaders(sh,headers) {
  if(sh.getLastRow()===0){ sh.appendRow(headers); sh.getRange(1,1,1,headers.length).setFontWeight('bold'); sh.setFrozenRows(1); }
}

function clearAndWrite(sh,headers,rows) {
  sh.clearContents();
  const data=[headers,...rows.map(r=>headers.map(h=>r[h]!=null?r[h]:''))];
  sh.getRange(1,1,data.length,headers.length).setValues(data);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold'); sh.setFrozenRows(1);
}

// ════════════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════════════
function initSheets() {
  const defs = {
    [TABS.STATE]:        ['key','valueJson'],
    [TABS.TRANSACTIONS]: ['ts','type','cat','stage','itemId','itemName','itemCode','qty','date','batchCode','supplier','invoice','distId','distName','rmUsed','sfgUsed','fgUsed','remarks'],
    [TABS.FIFO_RM]:      ['flavourId','batchCode','qty','date','ts'],
    [TABS.FIFO_RM_ITEM]: ['rmItemId','batchCode','qty','date','ts'],
    [TABS.FIFO_PP]:      ['ppItemId','batchCode','qty','date','ts'],
    [TABS.FIFO_BM]:      ['bmiId','batchCode','qty','date','ts'],
    [TABS.FIFO_SFG1]:    ['sfg1Id','batchCode','qty','date','ts'],
    [TABS.FIFO_SFG]:     ['sfgId','batchCode','qty','date','ts'],
    [TABS.FIFO_FG]:      ['skuId','batchCode','qty','date','ts'],
    [TABS.BATCH_SEQ]:    ['key','seq'],
  };
  const spreadsheet=ss();
  Object.entries(defs).forEach(([name,headers])=>{
    let sh=spreadsheet.getSheetByName(name);
    if(!sh){ sh=spreadsheet.insertSheet(name); sh.appendRow(headers); sh.getRange(1,1,1,headers.length).setFontWeight('bold'); sh.setFrozenRows(1); }
  });
  return { ok:true, message:'All sheets ready' };
}

// ════════════════════════════════════════════════════════════════
//  LOAD ALL
// ════════════════════════════════════════════════════════════════
function loadAll() {
  const stateMap={};
  sheetToObjects(getSheet(TABS.STATE)).forEach(r=>{ try{stateMap[r.key]=JSON.parse(r.valueJson);}catch(e){stateMap[r.key]=r.valueJson;} });

  const transactions=sheetToObjects(getSheet(TABS.TRANSACTIONS)).map(r=>({
    ts:Number(r.ts),type:String(r.type||''),cat:String(r.cat||''),stage:String(r.stage||''),
    itemId:String(r.itemId||''),itemName:String(r.itemName||''),itemCode:String(r.itemCode||''),
    qty:Number(r.qty),date:String(r.date||''),batchCode:String(r.batchCode||''),
    supplier:String(r.supplier||''),invoice:String(r.invoice||''),
    distId:String(r.distId||''),distName:String(r.distName||''),
    rmUsed:String(r.rmUsed||''),sfgUsed:String(r.sfgUsed||''),fgUsed:String(r.fgUsed||''),remarks:String(r.remarks||''),
  }));

  function buildFIFO(tab,keyField) {
    return sheetToObjects(getSheet(tab)).filter(r=>Number(r.qty)>0).sort((a,b)=>Number(a.ts)-Number(b.ts))
      .reduce((m,r)=>{ const k=String(r[keyField]); if(!m[k])m[k]=[]; m[k].push({batchCode:String(r.batchCode),qty:Number(r.qty),date:String(r.date),ts:Number(r.ts)}); return m; },{});
  }

  function buildBM(tab,keyField) {
    return sheetToObjects(getSheet(tab)).filter(r=>Number(r.qty)>0).sort((a,b)=>Number(a.ts)-Number(b.ts))
      .reduce((m,r)=>{ const k=String(r[keyField]); if(!m[k])m[k]=[]; m[k].push({batchCode:String(r.batchCode),qty:Number(r.qty),date:String(r.date),ts:Number(r.ts)}); return m; },{});
  }
  return {
    ok:true, state:stateMap, transactions,
    rmBatches:    buildFIFO(TABS.FIFO_RM,'flavourId'),
    bmBatches:    buildBM(TABS.FIFO_BM,'bmiId'),
    rmItemBatches:buildFIFO(TABS.FIFO_RM_ITEM,'rmItemId'),
    ppBatches:    buildFIFO(TABS.FIFO_PP,'ppItemId'),
    sfg1Batches:  buildFIFO(TABS.FIFO_SFG1,'sfg1Id'),
    sfgBatches:   buildFIFO(TABS.FIFO_SFG,'sfgId'),
    fgBatches:    buildFIFO(TABS.FIFO_FG,'skuId'),
  };
}

// ════════════════════════════════════════════════════════════════
//  SAVE STATE
// ════════════════════════════════════════════════════════════════
function saveState(data) {
  const sh=getSheet(TABS.STATE);
  clearAndWrite(sh,['key','valueJson'],Object.entries(data).map(([key,val])=>({key,valueJson:JSON.stringify(val)})));
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
//  ADD TRANSACTION
// ════════════════════════════════════════════════════════════════
function addTransaction(txn) {
  const sh=getSheet(TABS.TRANSACTIONS);
  const H=['ts','type','cat','stage','itemId','itemName','itemCode','qty','date','batchCode','supplier','invoice','distId','distName','rmUsed','sfgUsed','fgUsed','remarks'];
  ensureHeaders(sh,H);
  sh.appendRow(H.map(h=>txn[h]!=null?txn[h]:''));
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
//  ADD FIFO BATCH
// ════════════════════════════════════════════════════════════════
function addFIFOBatch(kind, batch) {
  const tabMap={ rm:TABS.FIFO_RM, rmItem:TABS.FIFO_RM_ITEM, ppItem:TABS.FIFO_PP, bmBatch:TABS.FIFO_BM, sfg1:TABS.FIFO_SFG1, sfg:TABS.FIFO_SFG, fg:TABS.FIFO_FG };
  const keyMap={ rm:'flavourId', rmItem:'rmItemId', ppItem:'ppItemId', bmBatch:'bmiId', sfg1:'sfg1Id', sfg:'sfgId', fg:'skuId' };
  const sh=getSheet(tabMap[kind]);
  const keyField=keyMap[kind];
  const headers=[keyField,'batchCode','qty','date','ts'];
  ensureHeaders(sh,headers);
  sh.appendRow([batch[keyField],batch.batchCode,batch.qty,batch.date,batch.ts]);
  return { ok:true };
}

// ════════════════════════════════════════════════════════════════
//  CONSUME FIFO BATCH
// ════════════════════════════════════════════════════════════════
function consumeFIFOBatch(kind, id, needed) {
  needed=Number(needed);
  const tabMap={ rm:TABS.FIFO_RM, rmItem:TABS.FIFO_RM_ITEM, ppItem:TABS.FIFO_PP, bmBatch:TABS.FIFO_BM, sfg1:TABS.FIFO_SFG1, sfg:TABS.FIFO_SFG, fg:TABS.FIFO_FG };
  const keyMap={ rm:'flavourId', rmItem:'rmItemId', ppItem:'ppItemId', bmBatch:'bmiId', sfg1:'sfg1Id', sfg:'sfgId', fg:'skuId' };
  const sh=getSheet(tabMap[kind]);
  const data=sh.getDataRange().getValues();
  if(data.length<2) return { ok:false, error:'No stock in sheet for '+kind };
  const H=data[0].map(String);
  const keyIdx=H.indexOf(keyMap[kind]), qtyIdx=H.indexOf('qty'), codeIdx=H.indexOf('batchCode'), tsIdx=H.indexOf('ts');
  const rows=[];
  for(let i=1;i<data.length;i++){
    if(String(data[i][keyIdx])===String(id)&&Number(data[i][qtyIdx])>0)
      rows.push({r:i+1,code:String(data[i][codeIdx]),qty:Number(data[i][qtyIdx]),ts:Number(data[i][tsIdx])});
  }
  rows.sort((a,b)=>a.ts-b.ts);
  const avail=rows.reduce((s,r)=>s+r.qty,0);
  if(avail<needed) return { ok:false, error:'Need '+needed+' but only '+avail.toFixed(3)+' available' };
  let rem=needed; const consumed=[];
  for(const row of rows){
    if(rem<=0)break;
    const take=Math.min(row.qty,rem);
    sh.getRange(row.r,qtyIdx+1).setValue(+(row.qty-take).toFixed(6));
    consumed.push({batchCode:row.code,qty:take}); rem-=take;
  }
  return { ok:true, consumed };
}

// ════════════════════════════════════════════════════════════════
//  BATCH SEQUENCE
// ════════════════════════════════════════════════════════════════
function nextBatchSeq(key) {
  const sh=getSheet(TABS.BATCH_SEQ);
  ensureHeaders(sh,['key','seq']);
  const data=sh.getDataRange().getValues();
  const H=data[0].map(String), keyIdx=H.indexOf('key'), seqIdx=H.indexOf('seq');
  for(let i=1;i<data.length;i++){
    if(String(data[i][keyIdx])===String(key)){ const n=Number(data[i][seqIdx])+1; sh.getRange(i+1,seqIdx+1).setValue(n); return{ok:true,seq:n}; }
  }
  sh.appendRow([key,1]); return{ok:true,seq:1};
}

// ════════════════════════════════════════════════════════════════
//  SET BATCH SEQ — admin override
// ════════════════════════════════════════════════════════════════
function setBatchSeq(key, value) {
  const sh=getSheet(TABS.BATCH_SEQ);
  ensureHeaders(sh,['key','seq']);
  const data=sh.getDataRange().getValues();
  const H=data[0].map(String), keyIdx=H.indexOf('key'), seqIdx=H.indexOf('seq');
  for(let i=1;i<data.length;i++){
    if(String(data[i][keyIdx])===String(key)){ sh.getRange(i+1,seqIdx+1).setValue(Number(value)); return{ok:true}; }
  }
  sh.appendRow([key,Number(value)]); return{ok:true};
}
