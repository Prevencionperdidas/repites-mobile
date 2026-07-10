
/* ════════════════════════════════════════════════════════════════
   MENÚ EN GOOGLE SHEETS — aparece automáticamente al abrir la planilla
   Úsalo para configurar hojas sin necesitar abrir la app
════════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 Control ISOs')
    .addItem('⚙ Configurar hojas y Dashboard', 'runSetup')
    .addItem('📊 Resumen rápido', 'showSummary')
    .addSeparator()
    .addItem('ℹ️  Instrucciones', 'showInstructions')
    .addToUi();
}

function runSetup() {
  var result = serverSetup();
  var ui = SpreadsheetApp.getUi();
  if (result.ok) {
    ui.alert('✅ Configuración completada',
      'Se crearon las pestañas por etapa y el Dashboard.\n\nAbre la pestaña 📊 Dashboard para ver el resumen.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Error', result.error || 'Error desconocido', ui.ButtonSet.OK);
  }
}

function showSummary() {
  var isos = readAll();
  var counts = {};
  isos.forEach(function(iso) {
    counts[iso.stage] = (counts[iso.stage] || 0) + 1;
  });
  var msg = 'ISOs por etapa:\n\n';
  var stageOrder = ['llegada','recepcion','bodega','retiro','anden','anulado'];
  var labels = {llegada:'Llegada',recepcion:'Recepción L.Inv.',
                bodega:'Bodega Rechazo',retiro:'Retiro',anden:'Andén',anulado:'Anulados'};
  stageOrder.forEach(function(key) {
    msg += '  ' + labels[key] + ': ' + (counts[key] || 0) + '\n';
  });
  msg += '\nTOTAL: ' + isos.length;
  SpreadsheetApp.getUi().alert('📊 Resumen ISOs/ASOs', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function showInstructions() {
  var url = ScriptApp.getService().getUrl();
  var ui = SpreadsheetApp.getUi();
  ui.alert('ℹ️ Instrucciones de uso',
    'URL de la app (PC):\n' + url + '\n\n' +
    'URL app móvil:\n' + url + '?m=1\n\n' +
    'IMPORTANTE: Abre siempre la app desde estas URLs,\n' +
    'NO abras el archivo HTML directamente.',
    ui.ButtonSet.OK);
}

/* ════════════════════════════════════════════════════════════════
   CONTROL DE ISOs/ASOs — Backend Google Apps Script
   IKEA Home Delivery & Logistics
   ════════════════════════════════════════════════════════════════
   INSTALACIÓN:
   1. Extensiones → Apps Script en tu planilla
   2. Pega este Code.gs completo
   3. Crea App.html    → "+" → HTML → nombre: App    → pega control-isos.html
   4. Crea Mobile.html → "+" → HTML → nombre: Mobile → pega mobile.html
   5. Implementar → Nueva implementación → Aplicación web
      · Ejecutar como: Yo  ·  Acceso: Cualquier usuario
   6. Para configurar las hojas: abre tu planilla de Google Sheets
      y usa el menú  Control ISOs → ⚙ Configurar hojas y Dashboard
════════════════════════════════════════════════════════════════ */

var STAGES = [
  { key:'llegada',   label:'Llegada',          tab:'01-Llegada',        hex:'#1A5EA8' },
  { key:'recepcion', label:'Recepción L.Inv.',  tab:'02-Recepción LI',   hex:'#1269C0' },
  { key:'bodega',    label:'Bodega Rechazo',    tab:'03-Bodega Rechazo', hex:'#2978D4' },
  { key:'retiro',    label:'Retiro',            tab:'04-Retiro',         hex:'#3B87E0' },
  { key:'anden',     label:'Andén',             tab:'05-Andén',          hex:'#1F8A4C' },
  { key:'anulado',   label:'Anulados',          tab:'06-Anulados',       hex:'#D7263D' },
];
var MAIN_TAB = 'ISOs';
var DASH_TAB = '📊 Dashboard';

/* ════════════════════════════════════════════════════════════════
   MENÚ EN GOOGLE SHEETS — aparece automáticamente al abrir
════════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛡 Control ISOs')
    .addItem('⚙ Configurar hojas y Dashboard', 'setupAll')
    .addSeparator()
    .addItem('📊 Ir al Dashboard', 'goToDashboard')
    .addItem('🔄 Actualizar pestañas de etapas', 'setupStageTabs')
    .addToUi();
}

function goToDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d  = ss.getSheetByName(DASH_TAB);
  if (d) ss.setActiveSheet(d);
  else SpreadsheetApp.getUi().alert('Primero ejecuta "Configurar hojas y Dashboard".');
}

/* ════════════════════════════════════════════════════════════════
   HOJA PRINCIPAL ISOs
   Columnas: A:ID  B:Código  C:Etapa  D:Creado  E:Actualizado
             F:JSON  G:Entrada a etapa  H:Responsable  I:Motivo  J:Transportista
════════════════════════════════════════════════════════════════ */
function getMainSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s  = ss.getSheetByName(MAIN_TAB);
  if (!s) {
    s = ss.insertSheet(MAIN_TAB);
    s.appendRow(['ID','Código','Etapa','Creado','Actualizado',
                 'JSON','Entrada a etapa','Responsable','Motivo','Transportista']);
    s.setFrozenRows(1);
    s.getRange('1:1').setFontWeight('bold')
      .setBackground('#0058A3').setFontColor('#FFFFFF');
    [1,2,3,4,5,7,8,9,10].forEach(function(c){ s.setColumnWidth(c,140); });
    s.setColumnWidth(6,500);
  }
  return s;
}

function readAll() {
  var s=getMainSheet(), last=s.getLastRow();
  if(last<2) return [];
  var rows=s.getRange(2,1,last-1,6).getValues(), out=[];
  rows.forEach(function(r){ try{ if(r[5]) out.push(JSON.parse(r[5])); }catch(e){} });
  return out;
}

function findRow(s, id) {
  var last=s.getLastRow(); if(last<2) return -1;
  var ids=s.getRange(2,1,last-1,1).getValues();
  for(var i=0;i<ids.length;i++){ if(ids[i][0]===id) return i+2; }
  return -1;
}

function saveOne(iso) {
  if(!iso||!iso.id) return;
  var lock=LockService.getScriptLock(); lock.waitLock(15000);
  try{
    var s=getMainSheet(), now=new Date().toISOString();
    var hist=iso.history||[], last=hist[hist.length-1]||{};
    var row=[
      iso.id, iso.code||'', iso.stage||'',
      iso.createdAt||now, iso.updatedAt||now, JSON.stringify(iso),
      iso.stageEnteredAt||now, last.responsable||'',
      (iso.data&&iso.data.llegada&&iso.data.llegada.motivoDevolucion)||'',
      (iso.data&&iso.data.llegada&&iso.data.llegada.empresaTransportista)||'',
    ];
    var idx=findRow(s,iso.id);
    if(idx===-1) s.appendRow(row);
    else s.getRange(idx,1,1,10).setValues([row]);
  }finally{ lock.releaseLock(); }
}

/* ════════════════════════════════════════════════════════════════
   CONFIGURACIÓN DE HOJAS (llamable desde menú Y desde la app)
════════════════════════════════════════════════════════════════ */
function setupAll() {
  getMainSheet();      // asegura que exista
  setupStageTabs();
  setupDashboard();
  try{
    SpreadsheetApp.getUi().alert(
      '✅ Listo',
      'Hojas configuradas:\n' +
      '• 01-Llegada\n• 02-Recepción LI\n• 03-Bodega Rechazo\n' +
      '• 04-Retiro\n• 05-Andén\n• 06-Anulados\n• 📊 Dashboard',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }catch(e){}
}

function setupStageTabs() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  STAGES.forEach(function(st){
    var s=ss.getSheetByName(st.tab)||ss.insertSheet(st.tab);
    s.clearContents().clearFormats();
    // Fila 1: título con color
    s.getRange(1,1,1,6).merge()
      .setValue('ISOs en etapa: '+st.label+' — CD El Sauce IKEA')
      .setBackground(st.hex).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
    s.setRowHeight(1,36);
    // Fila 2: contador dinámico
    s.getRange(2,1,1,6).merge()
      .setFormula('="Total: "&COUNTIF(ISOs!C:C,"'+st.key+'")&" ISO(s) en esta etapa"')
      .setBackground('#F4F8FD').setFontColor(st.hex)
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    s.setRowHeight(2,28);
    // Fila 3+: datos con QUERY
    var q='=IFERROR(QUERY(ISOs!A:J,"SELECT B,C,G,H,I,J WHERE C=\''+st.key+
          '\' AND A<>\'\' ORDER BY G ASC LABEL B \'Código ISO/ASO\',C \'Etapa\''+
          ',G \'Entrada a etapa\',H \'Responsable\',I \'Motivo devolución\',J \'Transportista\'",0)'+
          ',"Sin ISOs en esta etapa actualmente")';
    s.getRange(3,1).setFormula(q);
    s.setFrozenRows(3);
    [160,120,170,160,200,180].forEach(function(w,i){ s.setColumnWidth(i+1,w); });
    s.setTabColor(st.hex);
  });
}

function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var d  = ss.getSheetByName(DASH_TAB) || ss.insertSheet(DASH_TAB, 0);
  d.clear(); d.clearFormats(); d.setTabColor('#FFDB00');

  /* ── Columnas de la hoja ─────────────────────────────────────
     A:ID B:Código C:Etapa D:Creado E:Actualizado F:JSON
     G:Entrada a etapa H:Responsable I:Motivo J:Transportista
  ──────────────────────────────────────────────────────────── */

  /* ── Helpers ── */
  function rng(r1,c1,r2,c2){ return d.getRange(r1,c1,r2-r1+1,c2-c1+1); }
  function hdr(r,c1,c2,txt,bg,fg,sz){
    rng(r,c1,r,c2).merge()
      .setValue(txt).setBackground(bg||'#002C52').setFontColor(fg||'#FFFFFF')
      .setFontWeight('bold').setFontSize(sz||10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    d.setRowHeight(r, sz&&sz>12 ? 50 : 26);
  }

  /* ══════════════════ TÍTULO ══════════════════ */
  hdr(1,1,14,'📊  DASHBOARD — Control ISOs / ASOs   ·   CD El Sauce IKEA Home Delivery','#002C52','#FFDB00',15);
  d.setRowHeight(1,52);
  rng(2,1,2,14).merge()
    .setFormula('="Última actualización: "&TEXT(MAX(ISOs!E:E),"dd/mm/yyyy hh:mm")  &"     |     Total ISOs registrados: "&(COUNTA(ISOs!A:A)-1)')
    .setBackground('#0058A3').setFontColor('#BBD6EF')
    .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle');
  d.setRowHeight(2,28); d.setRowHeight(3,10);

  /* ══════════════════ KPIs GLOBALES (fila 4-6) ══════════════════ */
  hdr(4,1,14,'KPIs GLOBALES','#0058A3','#FFDB00',10);

  var kpis=[
    {c:1, label:'TOTAL ISOs',          formula:'=COUNTA(ISOs!A:A)-1',                              color:'#002C52'},
    {c:3, label:'ACTIVOS (en proceso)', formula:'=COUNTIF(ISOs!C:C,"llegada")+COUNTIF(ISOs!C:C,"recepcion")+COUNTIF(ISOs!C:C,"bodega")+COUNTIF(ISOs!C:C,"retiro")', color:'#0058A3'},
    {c:5, label:'EN ANDÉN (completos)', formula:'=COUNTIF(ISOs!C:C,"anden")',                        color:'#1F8A4C'},
    {c:7, label:'ANULADOS',            formula:'=COUNTIF(ISOs!C:C,"anulado")',                       color:'#D7263D'},
    {c:9, label:'CREADOS HOY',         formula:'=COUNTIFS(ISOs!D:D,">="&TODAY(),ISOs!D:D,"<"&TODAY()+1)', color:'#B9790A'},
    {c:11,label:'COMPLETADOS HOY',     formula:'=COUNTIFS(ISOs!C:C,"anden",ISOs!G:G,">="&TODAY())', color:'#1F8A4C'},
  ];

  kpis.forEach(function(k){
    rng(5,k.c,5,k.c+1).merge().setValue(k.label)
      .setBackground('#EAF1FB').setFontColor('#6B7C8F')
      .setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    rng(6,k.c,6,k.c+1).merge().setFormula(k.formula)
      .setBackground('#FFFFFF').setFontColor(k.color)
      .setFontSize(32).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setNumberFormat('0');
    d.setRowHeight(5,22); d.setRowHeight(6,52);
  });

  // Tasa completitud %
  rng(5,13,5,14).merge().setValue('TASA COMPLETITUD')
    .setBackground('#EAF1FB').setFontColor('#6B7C8F')
    .setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  rng(6,13,6,14).merge()
    .setFormula('=IFERROR(TEXT(COUNTIF(ISOs!C:C,"anden")/(COUNTA(ISOs!A:A)-1),"0.0%"),"—")')
    .setBackground('#FFFFFF').setFontColor('#1F8A4C')
    .setFontSize(26).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  d.setRowHeight(7,10);

  /* ══════════════════ TABLA POR ETAPA (fila 8-15) ══════════════════ */
  hdr(8,1,8,'DISTRIBUCIÓN POR ETAPA — ISOs activos, críticos y tiempos','#002C52','#FFFFFF',10);

  rng(9,1,9,8).setValues([['ETAPA','ISOs ACTIVOS','% TOTAL','🔴 CRÍTICOS +48h','⚠️ ADVERTENCIA +24h','ÚLT. INGRESO A ETAPA','TIEMPO PROM. EN ETAPA','BARRAS']])
    .setBackground('#0058A3').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle');
  d.setRowHeight(9,28);

  var stages=[
    {key:'llegada',   label:'1 · Llegada',          bg:'#EAF1FB',fg:'#1A5EA8'},
    {key:'recepcion', label:'2 · Recepción L.Inv.',  bg:'#F4F8FD',fg:'#1269C0'},
    {key:'bodega',    label:'3 · Bodega Rechazo',    bg:'#EAF1FB',fg:'#2978D4'},
    {key:'retiro',    label:'4 · Retiro',            bg:'#F4F8FD',fg:'#3B87E0'},
    {key:'anden',     label:'5 · Andén (completo)',  bg:'#E7F5EC',fg:'#1F8A4C'},
  ];

  stages.forEach(function(st,i){
    var r=10+i; d.setRowHeight(r,30);
    d.getRange(r,1).setValue(st.label).setFontColor(st.fg).setFontWeight('bold').setFontSize(10);
    d.getRange(r,2).setFormula('=COUNTIF(ISOs!C:C,"'+st.key+'")').setFontSize(18).setFontWeight('bold').setFontColor(st.fg).setHorizontalAlignment('center');
    d.getRange(r,3).setFormula('=IFERROR(TEXT(B'+r+'/$B$15,"0.0%"),"—")').setHorizontalAlignment('center');
    d.getRange(r,4).setFormula('=IFERROR(COUNTIFS(ISOs!C:C,"'+st.key+'",ISOs!G:G,"<"&(NOW()-2)),0)').setFontColor('#D7263D').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(14);
    d.getRange(r,5).setFormula('=IFERROR(COUNTIFS(ISOs!C:C,"'+st.key+'",ISOs!G:G,">="&(NOW()-2),ISOs!G:G,"<"&(NOW()-1)),0)').setFontColor('#B9790A').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(14);
    d.getRange(r,6).setFormula('=IFERROR(TEXT(MAXIFS(ISOs!G:G,ISOs!C:C,"'+st.key+'"),"dd/mm hh:mm"),"—")').setHorizontalAlignment('center').setFontSize(9);
    d.getRange(r,7).setFormula('=IFERROR(TEXT(AVERAGEIF(ISOs!C:C,"'+st.key+'",NOW()-ISOs!G:G)*24,"0.0")&"h","—")').setHorizontalAlignment('center').setFontSize(10);
    d.getRange(r,8).setFormula('=IFERROR(REPT("█",ROUND(B'+r+'*20/MAX($B$10:$B$14+0.01))),"")').setFontColor(st.fg).setFontSize(9);
    rng(r,1,r,8).setBackground(st.bg);
  });

  // Fila TOTAL
  var rT=15; d.setRowHeight(rT,32);
  rng(rT,1,rT,8).setBackground('#002C52').setFontColor('#FFDB00').setFontWeight('bold');
  d.getRange(rT,1).setValue('TOTAL ACTIVOS (excl. Andén)');
  d.getRange(rT,2).setFormula('=B10+B11+B12+B13').setFontSize(18).setHorizontalAlignment('center').setFontColor('#FFDB00');
  d.getRange(rT,3).setValue('100%').setHorizontalAlignment('center');
  d.getRange(rT,4).setFormula('=SUM(D10:D14)').setFontColor('#FF9FAE').setHorizontalAlignment('center').setFontSize(14);
  d.getRange(rT,5).setFormula('=SUM(E10:E14)').setFontColor('#FFDB00').setHorizontalAlignment('center').setFontSize(14);

  d.setRowHeight(16,10);

  /* ══════════════════ FILA DE FLUJO VISUAL (fila 17-18) ══════════════════ */
  hdr(17,1,12,'BARRA DE FLUJO — % de ISOs en cada etapa','#0058A3','#FFDB00',9);
  d.setRowHeight(17,24);
  var flowStages=[
    {key:'llegada',r:10,col:'#1A5EA8'},{key:'recepcion',r:11,col:'#1269C0'},
    {key:'bodega',r:12,col:'#2978D4'},{key:'retiro',r:13,col:'#3B87E0'},{key:'anden',r:14,col:'#1F8A4C'}
  ];
  flowStages.forEach(function(st,i){
    var c=1+i*2; var r=st.r;
    d.setRowHeight(18,34); d.setRowHeight(19,26);
    rng(18,c,18,c+1).merge()
      .setFormula('="Etapa "&'+(i+1))
      .setBackground(st.col).setFontColor('#FFFFFF').setFontWeight('bold')
      .setFontSize(11).setHorizontalAlignment('center').setVerticalAlignment('middle');
    rng(19,c,19,c+1).merge()
      .setFormula('=IFERROR(TEXT(B'+r+'/(B10+B11+B12+B13+B14),"0.0%"),"0%")')
      .setBackground('#FFFFFF').setFontColor(st.col)
      .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    // Label debajo
    rng(20,c,20,c+1).merge()
      .setValue(['Llegada','Recepción','Bodega','Retiro','Andén'][i])
      .setBackground('#EAF1FB').setFontColor(st.col)
      .setFontSize(8).setHorizontalAlignment('center');
    d.setRowHeight(20,18);
  });
  d.setRowHeight(21,10);

  /* ══════════════════ MOTIVOS DE DEVOLUCIÓN (fila 22+) ══════════════════ */
  hdr(22,1,6,'MOTIVOS DE DEVOLUCIÓN — frecuencia por tipo','#002C52');
  rng(23,1,23,3).setValues([['MOTIVO','CANTIDAD','% DEL TOTAL']])
    .setBackground('#0058A3').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(9);
  d.setRowHeight(23,24);
  d.getRange(24,1).setFormula('=IFERROR(QUERY(ISOs!I:I,"SELECT I,COUNT(I) WHERE I<>\'\' GROUP BY I ORDER BY COUNT(I) DESC LABEL I \'Motivo\',COUNT(I) \'Cantidad\'",0),"Sin datos")');
  // % para cada motivo (filas 24-33, fila dinámica)
  for(var r2=24;r2<=35;r2++){
    d.getRange(r2,3).setFormula('=IFERROR(TEXT(B'+r2+'/(COUNTA(ISOs!A:A)-1),"0.0%"),"")').setHorizontalAlignment('center');
  }

  /* ══════════════════ ISOs CRÍTICOS (col 8+) ══════════════════ */
  hdr(22,8,14,'🔴  ISOs CRÍTICOS — Más de 48h sin avanzar','#D7263D');
  rng(23,8,23,12).setValues([['CÓDIGO','ETAPA','ENTRADA A ETAPA','RESPONSABLE','TRANSPORTISTA']])
    .setBackground('#D7263D').setFontColor('#FFFFFF').setFontWeight('bold')
    .setHorizontalAlignment('center').setFontSize(9);
  d.setRowHeight(22,26); d.setRowHeight(23,24);
  var critDate='TEXT(NOW()-2,"yyyy-mm-dd")';
  d.getRange(24,8).setFormula('=IFERROR(QUERY(ISOs!B:J,"SELECT B,C,G,H,J WHERE C<>\'anden\' AND C<>\'anulado\' AND C<>\'\' AND G < date \'"&'+critDate+'&"\' ORDER BY G ASC",0),"✓ Sin ISOs críticos")');

  /* ══════════════════ ANCHOS DE COLUMNA ══════════════════ */
  var widths=[155,88,80,90,100,125,120,155,100,130,130,140,0,0];
  widths.forEach(function(w,i){if(w)d.setColumnWidth(i+1,w);});

  /* ══════════════════ GRÁFICOS ══════════════════ */
  try{
    d.getCharts().forEach(function(c){d.removeChart(c);});

    // Gráfico 1: Barras por etapa (col 9-11, filas 4-8)
    var g1=d.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(d.getRange('A10:B14'))
      .setOption('title','ISOs activos por etapa')
      .setOption('titleTextStyle',{color:'#002C52',fontSize:12,bold:true})
      .setOption('colors',['#0058A3'])
      .setOption('backgroundColor','#F4F8FD')
      .setOption('legend',{position:'none'})
      .setOption('hAxis',{title:'Cantidad',titleTextStyle:{color:'#6B7C8F'},minValue:0})
      .setOption('vAxis',{textStyle:{color:'#002C52',bold:true,fontSize:10}})
      .setOption('chartArea',{width:'65%',height:'80%',left:'30%'})
      .setPosition(4,10,0,10)
      .setOption('width',340).setOption('height',240)
      .build();
    d.insertChart(g1);

    // Gráfico 2: Dona de distribución %
    var g2=d.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(d.getRange('A10:B14'))
      .setOption('title','Distribución % por etapa')
      .setOption('titleTextStyle',{color:'#002C52',fontSize:11,bold:true})
      .setOption('colors',['#1A5EA8','#1269C0','#2978D4','#3B87E0','#1F8A4C'])
      .setOption('backgroundColor','#F4F8FD')
      .setOption('pieHole',0.4)
      .setOption('legend',{position:'bottom',textStyle:{color:'#002C52',fontSize:8}})
      .setOption('chartArea',{width:'90%',height:'75%'})
      .setPosition(4,12,0,10)
      .setOption('width',320).setOption('height',240)
      .build();
    d.insertChart(g2);

    // Gráfico 3: Motivos de devolución
    var g3=d.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(d.getRange('A24:B34'))
      .setOption('title','Motivos de devolución (top 10)')
      .setOption('titleTextStyle',{color:'#002C52',fontSize:11,bold:true})
      .setOption('colors',['#FFDB00'])
      .setOption('backgroundColor','#F4F8FD')
      .setOption('legend',{position:'none'})
      .setOption('hAxis',{minValue:0,gridlines:{color:'#EAF1FB'}})
      .setOption('vAxis',{textStyle:{color:'#002C52',fontSize:8}})
      .setOption('chartArea',{width:'60%',height:'85%',left:'35%'})
      .setPosition(22,8,410,10)
      .setOption('width',320).setOption('height',260)
      .build();
    d.insertChart(g3);
  } catch(e){ Logger.log('Chart error (non-fatal): '+e.message); }

  ss.setActiveSheet(d);
  return { ok:true, msg:'Dashboard creado correctamente con gráficos.' };
}


/* ════════════════════════════════════════════════════════════════
   FUNCIONES DE SERVIDOR — llamadas desde App vía google.script.run
════════════════════════════════════════════════════════════════ */
function serverGetAll() {
  try{ return { ok:true, isos:readAll() }; }
  catch(e){ return { ok:false, error:e.message }; }
}

function serverSave(isoJson) {
  try{ saveOne(JSON.parse(isoJson)); return { ok:true }; }
  catch(e){ return { ok:false, error:e.message }; }
}

function serverBulk(isosJson) {
  try{
    JSON.parse(isosJson).forEach(saveOne);
    return { ok:true, count:JSON.parse(isosJson).length };
  }catch(e){ return { ok:false, error:e.message }; }
}

// serverSetup llamada desde google.script.run (dentro de /exec)
function serverSetup() {
  try{
    getMainSheet();
    setupStageTabs();
    setupDashboard();
    return { ok:true, msg:'Hojas configuradas. Revisa tu planilla de Google Sheets.' };
  }catch(e){ return { ok:false, error:e.message }; }
}

/* ════════════════════════════════════════════════════════════════
   PUNTO DE ENTRADA HTTP
════════════════════════════════════════════════════════════════ */
function doGet(e) {
  var p=(e&&e.parameter)||{}, fn=p.fn||'';
  if(!fn){
    var mob=p.m==='1';
    return HtmlService.createHtmlOutputFromFile(mob?'Mobile':'App')
      .setTitle(mob?'Scanner ISO/ASO — Bodega · Retiro · Andén':'Control de ISOs/ASOs — IKEA')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport','width=device-width,initial-scale=1');
  }
  var data=p.data;
  if(fn==='serverGetAll')  return jsonOut(serverGetAll());
  if(fn==='serverSave')    return jsonOut(serverSave(data||'{}'));
  if(fn==='serverBulk')    return jsonOut(serverBulk(data||'[]'));
  if(fn==='serverSetup')   return jsonOut(serverSetup());
  return jsonOut({ok:false,error:'Función desconocida: '+fn});
}
function doPost(e){ return doGet(e); }

function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
