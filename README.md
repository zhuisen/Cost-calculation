<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>成本计算器｜SheetA 成本计算 / SheetB 成本反推倍数</title>
<style>
  :root{
    --bg:#f5f7fb;--card:#ffffff;--muted:#5b6b82;--text:#0f172a;--accent:#2563eb;--line:#e5e7eb
  }
  *{box-sizing:border-box}
  body{margin:0;font:14px/1.45 system-ui,Segoe UI,Helvetica,Arial;background:var(--bg);color:var(--text)}
  header{padding:14px 18px;border-bottom:1px solid var(--line);background:#fff;z-index:5}
  h1{margin:0;font-size:18px;letter-spacing:.2px}
  .sub{color:#64748b;font-size:12px}
  .wrap{max-width:1200px;margin:0 auto;padding:14px}

  /* 全局参数吸顶栏 */
  #globalBar{
    position:sticky; top:0; z-index:50;
    display:flex; gap:6px; flex-wrap:wrap; /* 窄屏自动换行，避免横向滚动条 */
    margin:8px 0 10px; padding:4px;
    background:#F6F9FFCC; border:1px solid #dbe6ff; border-radius:10px;
    align-items:center; justify-content:flex-start;
    box-shadow:0 2px 8px rgba(2,6,23,.06); backdrop-filter:saturate(1.15) blur(2px);
  }
  .tool{background:#ffffff;border:1.5px solid #94a3b8;padding:4px 6px;border-radius:14px;display:flex;align-items:center;gap:4px}
  .tool label{color:#0f172a;font-size:14px;font-weight:600;white-space:nowrap}
  input, select{
    background:#fff;border:1.5px solid #94a3b8;color:var(--text);
    padding:3px 5px;border-radius:14px;outline:none;min-width:48px;
    font-size:14px;font-weight:500;text-align:center
  }
  input:focus, select:focus{border-color:var(--accent);box-shadow:0 0 0 3px #2563eb22}

  /* 页签 */
  .tabs{display:flex;gap:8px;margin-top:6px;margin-bottom:10px}
  .tabs button{background:#fff;border:1px solid #dbe6ff;border-bottom:2px solid transparent;padding:8px 12px;border-radius:8px;cursor:pointer;color:#1e3a8a}
  .tabs button.active{border-color:#2563eb;background:#eef4ff;color:#0f172a;font-weight:700}
  .sheet{display:none}
  .sheet.is-active{display:block}

  /* Sheet 内部工具栏（按钮行） */
  .toolbar.actions{
    display:flex; gap:10px; flex-wrap:wrap;
    background:#F6F9FF; border:1px solid #dbe6ff; border-radius:10px;
    padding:6px; margin:6px 0 10px; justify-content:flex-start
  }

  button{background:var(--accent);border:1px solid #1e40af;color:#fff;padding:8px 12px;border-radius:10px;cursor:pointer}
  button.secondary{background:#eef2ff;color:#1e40af;border-color:#c7d2fe}
  button.ghost{background:#fff;color:#1e3a8a;border-color:var(--line)}
  button:hover{filter:brightness(1.03)}

  table{width:100%;border-collapse:separate;border-spacing:0;background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
  thead th{background:#f8fafc;border-bottom:1px solid var(--line);color:#0f172a;font-weight:600;padding:8px 6px;font-size:12px;white-space:nowrap}
  tbody td{border-bottom:1px solid var(--line);padding:4px 6px}
  tbody tr:last-child td{border-bottom:none}
  .num{text-align:right}
  tfoot td{padding:10px;border-top:1px solid var(--line);background:#f8fafc}
  .row-actions{display:flex;gap:6px}
  .nowrap{white-space:nowrap}
  td input, td select{min-width:60px;width:90px;transition:border-color .15s ease, box-shadow .15s ease}
  td.wide input{width:110px}

  /* 视觉分组：输入 vs 计算 */
  .group-input{background:#FCFEFF}
  .group-calc{background:#F8FAFC}
  .vol, .weight, .cost, .ex{background:#F1F5F9}
  td input:hover, td select:hover{border-color:var(--accent);box-shadow:0 0 0 3px #2563eb22}

  /* 合计/总额条 */
  .totalbar{
    margin:12px 0 6px;padding:10px 12px;background:#fff;
    border:1px solid #dbe6ff;border-radius:10px;font-size:20px;font-weight:700;color:#0f172a
  }
  .totalbar .currency{font-weight:600;color:#2563eb}
  .hint{color:#475569;font-size:12px}
</style>
</head>
<body>
  <header>
    <h1>成本计算器（SheetA 成本计算 / SheetB 成本反推倍数）</h1>
    <div class="sub">体积单位 <b>cm³</b>（按 mm 尺寸计算后 ÷1000），密度默认 <b>7.85 g/cm³</b>（可编辑）。</div>
  </header>

  <div class="wrap">
    <!-- 全局顶部参数（两张 Sheet 共用） -->
    <div id="globalBar">
      <div class="tool"><label for="density">密度(g/cm³)</label><input id="density" type="number" step="0.01" value="7.85"></div>
      <div class="tool"><label for="priceLine">线 单价(元/kg)</label><input id="priceLine" type="number" step="0.01" value="5"></div>
      <div class="tool"><label for="pricePipe">管 单价(元/kg)</label><input id="pricePipe" type="number" step="0.01" value="7"></div>
      <div class="tool"><label for="defaultMult">默认倍数</label><input id="defaultMult" type="number" step="0.01" value="3"></div>
    </div>

    <!-- 页签 -->
    <div class="tabs">
      <button id="tab-cost" class="active">Sheet A · 成本计算</button>
      <button id="tab-alt">Sheet B · 成本反推倍数</button>
    </div>

    <!-- Sheet A：成本计算 -->
    <div id="sheet-cost" class="sheet is-active">
      <div class="toolbar actions">
        <button id="addRow">+ 新增一行</button>
        <button id="exportCsv" class="secondary">导出 CSV（Excel 可打开）</button>
        <button id="reset" class="ghost">清空表格</button>
      </div>

      <table id="tbl">
        <thead>
          <tr>
            <th>类型</th>
            <th class="nowrap">长 / mm</th>
            <th class="nowrap">直径 / 管宽1 / mm</th>
            <th class="nowrap">管宽2 / mm <span class="hint">（方管第二边，线/圆管可留空）</span></th>
            <th class="nowrap">壁厚 / mm</th>
            <th class="nowrap">数量 / pc</th>
            <th class="nowrap">体积 / cm³</th>
            <th class="nowrap">重量 / kg</th>
            <th class="nowrap">成本(元)</th>
            <th>倍数</th>
            <th class="nowrap">出厂价(元)</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody id="body"></tbody>
        <tfoot>
          <tr>
            <td colspan="7" class="hint">合计</td>
            <td class="num" id="sumWeight">0.00</td>
            <td class="num" id="sumCost">0.00</td>
            <td></td>
            <td class="num" id="sumEx">0.00</td>
            <td><button id="copyAll" class="secondary">复制全部到 Excel</button></td>
          </tr>
        </tfoot>
      </table>

      <div class="totalbar">总出厂价合计（含倍数）：<span class="currency">¥</span> <span id="grandEx">0.00</span></div>

      <p class="hint" style="margin-top:10px">说明：
        <br>• <b>线（实心圆）</b>：需填写 <b>长、直径、数量</b>
        <br>• <b>圆管</b>：需填写 <b>长、管宽1、壁厚、数量</b>
        <br>• <b>方管</b>：需填写 <b>长、管宽1、管宽2、壁厚、数量</b>
      </p>
    </div>

    <!-- Sheet B：成本反推倍数 -->
    <div id="sheet-alt" class="sheet">
      <div class="toolbar actions">
        <button id="addRowB">+ 新增一行</button>
        <button id="exportCsvB" class="secondary">导出 CSV（Excel 可打开）</button>
        <button id="resetB" class="ghost">清空表格</button>
      </div>

      <table id="tblB">
        <thead>
          <tr>
            <th>类型</th>
            <th class="nowrap">重量 / kg</th>
            <th class="nowrap">数量 / pc</th>
            <th class="nowrap">出厂价(元)</th>
            <th class="nowrap">倍数（反推）</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody id="bodyB"></tbody>
      </table>

      <div class="totalbar">总倍数：<span id="sumMultB">0.00</span></div>

      <p class="hint" style="margin-top:10px">说明（Sheet B）：倍数 = 出厂价 ÷ （重量 × 数量 × 单价）。其中：线使用“线 单价(元/kg)”，圆管/方管使用“管 单价(元/kg)”。</p>
    </div>
  </div>

<script>
/* ========= 共用小工具 ========= */
const $ = (s,p=document) => p.querySelector(s);
const $$ = (s,p=document) => [...p.querySelectorAll(s)];
const fmt = n => (isFinite(n) ? Number(n).toFixed(2) : '0.00');

/* ========= 页签切换 ========= */
function switchSheet(id){
  $$('.sheet').forEach(s=>s.classList.remove('is-active'));
  $$('.tabs button').forEach(b=>b.classList.remove('active'));
  const btnId = id==='sheet-cost' ? 'tab-cost' : 'tab-alt';
  document.getElementById(id)?.classList.add('is-active');
  document.getElementById(btnId)?.classList.add('active');
}
document.getElementById('tab-cost').addEventListener('click',()=>switchSheet('sheet-cost'));
document.getElementById('tab-alt').addEventListener('click',()=>switchSheet('sheet-alt'));

/* ========= 全局参数 ========= */
function currentDensity(){ return +$('#density').value || 7.85; }
function priceByType(type){
  const lineP = +$('#priceLine').value||0;
  const pipeP = +$('#pricePipe').value||0;
  return (type==='线') ? lineP : pipeP;
}

/* ========= A：成本计算 ========= */
function computeVolumeCm3(type, Lmm, W1mm, W2mm, Tmm){
  // 输入单位：mm，结果：cm^3（所以 ÷1000）
  const L = Number(Lmm||0);
  const W1 = Number(W1mm||0);
  const W2 = Number(W2mm||0);
  const T  = Number(Tmm||0);
  const PI = Math.PI;

  if(type==='线'){
    const r = W1/2; // 直径 = W1
    const vol_mm3 = PI * r*r * L;
    return vol_mm3/1000;
  }
  if(type==='圆管'){
    const R = W1/2; let r = R - T; if(r<0) r = 0;
    const vol_mm3 = PI * (R*R - r*r) * L;
    return vol_mm3/1000;
  }
  if(type==='方管'){
    const a = W1; const b = (W2>0?W2:W1);
    let ai = a - 2*T, bi = b - 2*T;
    if(ai<=0 || bi<=0){ // 实心方柱兜底
      return (a*b*L)/1000;
    }
    const area_mm2 = (a*b - ai*bi);
    return (area_mm2 * L)/1000;
  }
  return 0;
}

function rowTemplate(data={}){
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td class="group-input"><select class="type">
      <option value="线">线</option>
      <option value="圆管">圆管</option>
      <option value="方管">方管</option>
    </select></td>
    <td class="group-input"><input class="len num" type="number" step="0.01" value="${data.len??''}" placeholder="mm"></td>
    <td class="group-input"><input class="w1 num" type="number" step="0.01" value="${data.w1??''}" placeholder="直径/管宽1"></td>
    <td class="group-input"><input class="w2 num" type="number" step="0.01" value="${data.w2??''}" placeholder="管宽2(方管)"></td>
    <td class="group-input"><input class="t  num" type="number" step="0.01" value="${data.t??''}"  placeholder="壁厚"></td>
    <td class="group-input"><input class="qty num" type="number" step="1"    value="${data.qty??1}"></td>
    <td class="group-calc num vol">0.00</td>
    <td class="group-calc num weight">0.00</td>
    <td class="group-calc num cost">0.00</td>
    <td class="group-input"><input class="mult num" type="number" step="0.01" value="${data.mult??''}" placeholder="默认"></td>
    <td class="group-calc num ex">0.00</td>
    <td class="row-actions">
      <button class="dup"  title="复制一行">复制</button>
      <button class="copy" title="复制到Excel" type="button">复制到Excel</button>
      <button class="del"  title="删除" type="button">删除</button>
    </td>
  `;
  tr.$ = (s)=>tr.querySelector(s);
  tr.typeSel = tr.$('.type');
  tr.len = tr.$('.len');
  tr.w1  = tr.$('.w1');
  tr.w2  = tr.$('.w2');
  tr.t   = tr.$('.t');
  tr.qty = tr.$('.qty');
  tr.mult= tr.$('.mult');
  if(data.mult==null) tr.mult.value = $('#defaultMult').value;
  tr.vol = tr.$('.vol');
  tr.weight = tr.$('.weight');
  tr.cost = tr.$('.cost');
  tr.ex   = tr.$('.ex');

  if(data.type){ tr.typeSel.value = data.type }

  [tr.typeSel,tr.len,tr.w1,tr.w2,tr.t,tr.qty,tr.mult]
    .forEach(el=>el.addEventListener('input',()=>recalcRow(tr,true)));
  tr.$('.dup').addEventListener('click',()=>{ addRow(getRowData(tr)); });
  tr.$('.del').addEventListener('click',()=>{ tr.remove(); recalcTotals(); saveAll(); });
  tr.$('.copy').addEventListener('click',()=>copyRowToClipboard(tr));

  recalcRow(tr,false);
  return tr;
}

function getRowData(tr){
  return {
    type: tr.typeSel.value,
    len: +tr.len.value||0,
    w1 : +tr.w1.value||0,
    w2 : +tr.w2.value||0,
    t  : +tr.t.value||0,
    qty: +tr.qty.value||1,
    mult: tr.mult.value===''? null : (+tr.mult.value||0)
  }
}

function recalcRow(tr, doSave){
  const d = getRowData(tr);
  const mult = d.mult ?? (+$('#defaultMult').value||1);

  const volPer = computeVolumeCm3(d.type, d.len, d.w1, d.w2, d.t);
  const vol = volPer * (d.qty||1);
  const weightKg = vol * currentDensity() / 1000; // g -> kg
  const unitPrice = priceByType(d.type);
  const cost = weightKg * unitPrice;
  const ex = cost * mult;

  tr.vol.textContent    = fmt(vol);
  tr.weight.textContent = fmt(weightKg);
  tr.cost.textContent   = fmt(cost);
  tr.ex.textContent     = fmt(ex);

  if(doSave){ saveAll(); }
  recalcTotals();
}

function recalcTotals(){
  let sumW=0,sumC=0,sumE=0;
  $$('#body tr').forEach(tr=>{
    sumW += +tr.querySelector('.weight').textContent||0;
    sumC += +tr.querySelector('.cost').textContent||0;
    sumE += +tr.querySelector('.ex').textContent||0;
  });
  $('#sumWeight').textContent = fmt(sumW);
  $('#sumCost').textContent   = fmt(sumC);
  $('#sumEx').textContent     = fmt(sumE);
  const g = $('#grandEx'); if(g) g.textContent = fmt(sumE);
}

function addRow(init={}){ $('#body').appendChild(rowTemplate(init)); saveAll(); }

function copyRowToClipboard(tr){
  const d = getRowData(tr);
  const cells = [
    d.type, d.len, d.w1, d.w2, d.t, d.qty,
    tr.querySelector('.vol').textContent,
    tr.querySelector('.weight').textContent,
    tr.querySelector('.cost').textContent,
    (d.mult ?? (+$('#defaultMult').value||1)),
    tr.querySelector('.ex').textContent
  ];
  navigator.clipboard.writeText(cells.join('\t')).then(()=>alert('已复制当前行，直接粘贴到 Excel 即可'));
}

function exportCsv(){
  const header = ['类型','长(mm)','直径/管宽1(mm)','管宽2(mm)','壁厚(mm)','数量(pc)','体积(cm3)','重量(kg)','成本(元)','倍数','出厂价(元)'];
  const rows = $$('#body tr').map(tr=>{
    const d = getRowData(tr);
    return [
      d.type, d.len, d.w1, d.w2, d.t, d.qty,
      tr.querySelector('.vol').textContent,
      tr.querySelector('.weight').textContent,
      tr.querySelector('.cost').textContent,
      (d.mult??(+$('#defaultMult').value||1)),
      tr.querySelector('.ex').textContent
    ];
  });
  const all = [header, ...rows];
  const csv = all.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '成本计算导出.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

function copyAllToClipboard(){
  const header = ['类型','长(mm)','直径/管宽1(mm)','管宽2(mm)','壁厚(mm)','数量(pc)','体积(cm3)','重量(kg)','成本(元)','倍数','出厂价(元)'];
  const rows = $$('#body tr').map(tr=>{
    const d = getRowData(tr);
    return [
      d.type, d.len, d.w1, d.w2, d.t, d.qty,
      tr.querySelector('.vol').textContent,
      tr.querySelector('.weight').textContent,
      tr.querySelector('.cost').textContent,
      (d.mult??(+$('#defaultMult').value||1)),
      tr.querySelector('.ex').textContent
    ];
  });
  const text = [header, ...rows].map(r=>r.join('\t')).join('\n');
  navigator.clipboard.writeText(text).then(()=>alert('✅ 已复制整表内容，可直接粘贴到 Excel'));
}

/* ========= B：成本反推倍数 ========= */
function priceByTypeB(type){ return priceByType(type); }

function rowTemplateB(data={}){
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td class="group-input"><select class="typeB">
      <option value="线">线</option>
      <option value="圆管">圆管</option>
      <option value="方管">方管</option>
    </select></td>
    <td class="group-input"><input class="wkg num"  type="number" step="0.01" value="${data.wkg??''}" placeholder="重量kg"></td>
    <td class="group-input"><input class="qtyB num"  type="number" step="1"    value="${data.qty??1}"></td>
    <td class="group-input"><input class="exB  num"  type="number" step="0.01" value="${data.ex??''}" placeholder="出厂价"></td>
    <td class="group-calc num multB">0.00</td>
    <td class="row-actions">
      <button class="dupB" title="复制一行">复制</button>
      <button class="delB" title="删除" type="button">删除</button>
    </td>
  `;
  tr.$ = (s)=>tr.querySelector(s);
  tr.typeSel = tr.$('.typeB');
  tr.wkg = tr.$('.wkg');
  tr.qty = tr.$('.qtyB');
  tr.ex  = tr.$('.exB');
  tr.multCell = tr.$('.multB');

  if(data.type){ tr.typeSel.value = data.type }

  [tr.typeSel,tr.wkg,tr.qty,tr.ex].forEach(el=>el.addEventListener('input',()=>recalcRowB(tr,true)));
  tr.$('.dupB').addEventListener('click',()=>{ addRowB(getRowDataB(tr)); });
  tr.$('.delB').addEventListener('click',()=>{ tr.remove(); saveAll(); recalcTotalsB(); });

  recalcRowB(tr,false);
  return tr;
}

function getRowDataB(tr){
  return {
    type: tr.typeSel.value,
    wkg: +tr.wkg.value||0,
    qty: +tr.qty.value||1,
    ex : +tr.ex.value||0
  }
}

function recalcRowB(tr, doSave){
  const d = getRowDataB(tr);
  const unit = priceByTypeB(d.type);
  const denom = (d.wkg||0) * (d.qty||1) * (unit||0);
  const mult = denom>0 ? (d.ex||0)/denom : 0;
  tr.multCell.textContent = fmt(mult);
  if(doSave) saveAll();
  recalcTotalsB();
}

function addRowB(init={}){ $('#bodyB').appendChild(rowTemplateB(init)); saveAll(); recalcTotalsB(); }

function recalcTotalsB(){
  const cells = $$('#bodyB .multB');
  let sum = 0, n = 0;
  cells.forEach(c=>{ const v = +c.textContent; if(isFinite(v) && v>0){ sum += v; n++; } });
  const avg = n ? sum / n : 0;   // 这里用“平均倍数”，如需改为求和，把 avg 改成 sum 即可
  const el = $('#sumMultB'); if(el) el.textContent = fmt(avg);
}

function exportCsvB(){
  const header = ['类型','重量(kg)','数量(pc)','出厂价(元)','倍数(反推)'];
  const rows = $$('#bodyB tr').map(tr=>{
    const d = getRowDataB(tr);
    return [d.type, d.wkg, d.qty, d.ex, tr.querySelector('.multB').textContent];
  });
  const csv = [header, ...rows].map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '成本反推倍数-导出.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ========= 存储（A+B+全局参数） ========= */
const STORAGE_KEY = 'cost-calc-v3';

function saveAll(){
  const rowsA = $$('#body tr').map(tr=>getRowData(tr));
  const rowsB = $$('#bodyB tr').map(tr=>getRowDataB(tr));
  const state = {
    defaults:{
      density:+$('#density').value||7.85,
      priceLine:+$('#priceLine').value||5,
      pricePipe:+$('#pricePipe').value||7,
      mult:+$('#defaultMult').value||3
    },
    rowsA, rowsB
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function loadAll(){
  const s = localStorage.getItem(STORAGE_KEY);
  if(!s) return false;
  try{
    const st = JSON.parse(s);
    if(st?.defaults){
      $('#density').value    = st.defaults.density ?? 7.85;
      $('#priceLine').value  = st.defaults.priceLine ?? 5;
      $('#pricePipe').value  = st.defaults.pricePipe ?? 7;
      $('#defaultMult').value= st.defaults.mult ?? 3;
    }
    if(Array.isArray(st.rowsA)) st.rowsA.forEach(r=>addRow(r));
    if(Array.isArray(st.rowsB)) st.rowsB.forEach(r=>addRowB(r));
    recalcTotals();
    recalcTotalsB();
    return true;
  }catch(e){ console.warn(e); }
  return false;
}

/* ========= 事件绑定 ========= */
$('#addRow').addEventListener('click',()=>addRow());
$('#copyAll').addEventListener('click', copyAllToClipboard);
$('#exportCsv').addEventListener('click', exportCsv);
$('#reset').addEventListener('click',()=>{
  if(confirm('确定要清空 Sheet A 吗？')){
    $('#body').innerHTML='';
    recalcTotals();
    saveAll();
  }
});
['priceLine','pricePipe','defaultMult','density'].forEach(id=>{
  $('#'+id).addEventListener('input',()=>{
    // 全局参数变化 → 两张表都重算
    $$('#body tr').forEach(tr=>recalcRow(tr,false));
    $$('#bodyB tr').forEach(tr=>recalcRowB(tr,false));
    recalcTotals(); recalcTotalsB(); saveAll();
  });
});

/* Sheet B 事件 */
$('#addRowB').addEventListener('click',()=>addRowB());
$('#exportCsvB').addEventListener('click',exportCsvB);
$('#resetB').addEventListener('click',()=>{
  if(confirm('确定要清空 Sheet B 吗？')){ $('#bodyB').innerHTML=''; recalcTotalsB(); saveAll(); }
});

/* ========= 初始化 ========= */
if(!loadAll()){
  // 初次打开不预置行，保持空白
}
</script>
</body>
</html>
