
/* sync.js — Integra o frontend com Excel (Office.js) sem alterar layout/JS do app. */
(function(){
  const headersObras = ["obra_id","fluig","mf","obra","cliente","lider","local","peso_orcado","peso_liberar","valor_fechado","valor_liberar","status"];
  const headersEtapas = ["obra_id","obra","mf_etapa","nome_etapa","peso","valor","status","envio_est","lib_est","envio_atu","lib_atu","deadline","aprovacao","lib_real"];

  function safeNumber(v){ const n = Number(v); return Number.isFinite(n) ? n : 0; }
  function obrasToRows(list){
    return list.map(o => [
      o.id||"", o.fluig||"", o.mf||"", o.nome||"", o.cliente||"", o.lider||"", o.local||"",
      safeNumber(o.pesoOrcado), safeNumber(o.pesoLiberar), safeNumber(o.valorFechado), safeNumber(o.valorLiberar), o.status||""
    ]);
  }
  function etapasToRows(list){
    const rows = [];
    list.forEach(o => {
      const ets = (o.etapas && o.etapas.length) ? o.etapas : [{
        mfEtapa:"—", peso:0, valor:0, status:"Etapa única",
        envioEst:"—", libEst:"—", envioAtu:"—", libAtu:"—",
        deadline:"—", aprovacao:"—", libReal:"—", nome:"Etapa única"
      }];
      ets.forEach(e => {
        rows.push([
          o.id||"", o.nome||"", e.mfEtapa||"", e.nome||"",
          safeNumber(e.peso), safeNumber(e.valor), e.status||"",
          e.envioEst||"", e.libEst||"", e.envioAtu||"", e.libAtu||"",
          e.deadline||"", e.aprovacao||"", e.libReal||""
        ]);
      });
    });
    return rows;
  }
  function rowsToObras(obrasRows, etapasRows){
    const map = new Map();
    (obrasRows||[]).forEach(r => {
      const [id, fluig, mf, obra, cliente, lider, local, peso_orcado, peso_liberar, valor_fechado, valor_liberar, status] = r;
      map.set(id, {
        id: String(id||""), fluig: String(fluig||""), mf: String(mf||""), nome: String(obra||""),
        cliente: String(cliente||""), lider: String(lider||""), local: String(local||""),
        pesoOrcado: safeNumber(peso_orcado), pesoLiberar: safeNumber(peso_liberar),
        valorFechado: safeNumber(valor_fechado), valorLiberar: safeNumber(valor_liberar),
        status: String(status||""), etapas: []
      });
    });
    (etapasRows||[]).forEach(r => {
      const [obra_id, obra, mf_etapa, nome_etapa, peso, valor, status, envio_est, lib_est, envio_atu, lib_atu, deadline, aprovacao, lib_real] = r;
      const host = map.get(String(obra_id||""));
      if(!host){ return; }
      host.etapas = host.etapas || [];
      host.etapas.push({
        mfEtapa: String(mf_etapa||""),
        peso: safeNumber(peso), valor: safeNumber(valor), status: String(status||""),
        envioEst: String(envio_est||""), libEst: String(lib_est||""),
        envioAtu: String(envio_atu||""), libAtu: String(lib_atu||""),
        deadline: String(deadline||""), aprovacao: String(aprovacao||""), libReal: String(lib_real||""),
        nome: String(nome_etapa||"")
      });
    });
    return Array.from(map.values());
  }
  async function loadFromExcel(){
    if(!(window.Office && window.Excel && Office.context)){ return; }
    try{
      await Excel.run(async (ctx)=>{
        const tables = ctx.workbook.tables;
        const tObras = tables.getItemOrNullObject("tbl_obras");
        const tEtapas = tables.getItemOrNullObject("tbl_etapas");
        tObras.load("name"); tEtapas.load("name");
        await ctx.sync();
        if(tObras.isNullObject){ await writeAllToExcel(ctx); return; }
        const ro = tObras.getDataBodyRange();
        const re = tEtapas.isNullObject ? null : tEtapas.getDataBodyRange();
        if(re){ ro.load("values"); re.load("values"); } else { ro.load("values"); }
        await ctx.sync();
        const obrasRows = (ro && ro.values && ro.values.length) ? ro.values : [];
        const etapasRows = (re && re.values && re.values.length) ? re.values : [];
        if(obrasRows.length){
          const lista = rowsToObras(obrasRows, etapasRows);
          if(Array.isArray(window.obras)){
            window.obras.length = 0; Array.prototype.push.apply(window.obras, lista);
            if(typeof window.renderKPIs === "function") window.renderKPIs();
            if(typeof window.initFiltroLider === "function") window.initFiltroLider();
            if(typeof window.renderTable === "function") window.renderTable();
          }
        } else {
          await writeAllToExcel(ctx);
        }
      });
    }catch(err){ console.error("Erro loadFromExcel:", err); }
  }
  async function writeAllToExcel(ctx){
    const wsObras0 = ctx.workbook.worksheets.getItemOrNullObject("Obras");
    const wsEtapas0 = ctx.workbook.worksheets.getItemOrNullObject("Etapas");
    await ctx.sync();
    const wsObras = wsObras0.isNullObject ? ctx.workbook.worksheets.add("Obras") : wsObras0;
    const wsEtapas = wsEtapas0.isNullObject ? ctx.workbook.worksheets.add("Etapas") : wsEtapas0;
    const rowsO = obrasToRows(window.obras||[]);
    const rowsE = etapasToRows(window.obras||[]);
    try { ctx.workbook.tables.getItem("tbl_obras").delete(); } catch(_){}
    try { ctx.workbook.tables.getItem("tbl_etapas").delete(); } catch(_){}
    const allO = [headersObras].concat(rowsO);
    const allE = [headersEtapas].concat(rowsE);
    const rangeO = wsObras.getRangeByIndexes(0, 0, max1(allO.length), headersObras.length);
    rangeO.values = allO;
    const rangeE = wsEtapas.getRangeByIndexes(0, 0, max1(allE.length), headersEtapas.length);
    rangeE.values = allE;
    ctx.workbook.tables.add(rangeO, true).name = "tbl_obras";
    ctx.workbook.tables.add(rangeE, true).name = "tbl_etapas";
    await ctx.sync();
  }
  function max1(n){ return Math.max(1, n); }
  async function syncToExcel(){
    if(!(window.Office && window.Excel)){ return; }
    try{
      await Excel.run(async (ctx)=>{
        let wsObras = ctx.workbook.worksheets.getItemOrNullObject("Obras");
        let wsEtapas = ctx.workbook.worksheets.getItemOrNullObject("Etapas");
        await ctx.sync();
        wsObras = wsObras.isNullObject ? ctx.workbook.worksheets.add("Obras") : wsObras;
        wsEtapas = wsEtapas.isNullObject ? ctx.workbook.worksheets.add("Etapas") : wsEtapas;
        const rowsO = obrasToRows(window.obras||[]);
        const rowsE = etapasToRows(window.obras||[]);
        try { ctx.workbook.tables.getItem("tbl_obras").delete(); } catch(_){}
        try { ctx.workbook.tables.getItem("tbl_etapas").delete(); } catch(_){}
        const allO = [headersObras].concat(rowsO);
        const allE = [headersEtapas].concat(rowsE);
        const rangeO = wsObras.getRangeByIndexes(0, 0, max1(allO.length), headersObras.length);
        rangeO.values = allO;
        const rangeE = wsEtapas.getRangeByIndexes(0, 0, max1(allE.length), headersEtapas.length);
        rangeE.values = allE;
        ctx.workbook.tables.add(rangeO, true).name = "tbl_obras";
        ctx.workbook.tables.add(rangeE, true).name = "tbl_etapas";
        await ctx.sync();
      });
    }catch(err){ console.error("Erro syncToExcel:", err); }
  }
  async function registerChangeWatchers(){
    if(!(window.Office && window.Excel)){ return; }
    try{
      await Excel.run(async (ctx)=>{
        const wsO = ctx.workbook.worksheets.getItemOrNullObject("Obras");
        const wsE = ctx.workbook.worksheets.getItemOrNullObject("Etapas");
        await ctx.sync();
        if(!wsO.isNullObject){ wsO.onChanged.add(()=> setTimeout(loadFromExcel, 150)); }
        if(!wsE.isNullObject){ wsE.onChanged.add(()=> setTimeout(loadFromExcel, 150)); }
        await ctx.sync();
      });
    }catch(err){ console.warn("Watcher falhou:", err); }
  }
  function attachPersistenceHooks(){
    const sNova = document.getElementById("salvarNova");
    if(sNova){ sNova.addEventListener("click", ()=> setTimeout(syncToExcel, 120)); }
    const geAdd = document.getElementById("ge_add");
    if(geAdd){ geAdd.addEventListener("click", ()=> setTimeout(syncToExcel, 120)); }
    const geLista = document.getElementById("ge_lista");
    if(geLista){
      geLista.addEventListener("click", (ev)=>{
        const btn = ev.target.closest('button[data-del]');
        if(btn){ setTimeout(syncToExcel, 200); }
      }, true);
    }
  }
  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", ()=>{
      attachPersistenceHooks();
      loadFromExcel().then(registerChangeWatchers);
    });
  } else {
    attachPersistenceHooks();
    loadFromExcel().then(registerChangeWatchers);
  }
})();
