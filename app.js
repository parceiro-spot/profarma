/* Dashboard Profarma: implementação nova baseada somente na configuração declarativa report-spec.json. */
const $ = (s, r=document) => r.querySelector(s);
const state = { pages: [], files: [], rows: [], page: 0, from: null, to: null, dateTouched: false, filters: {}, cross: null, openSlicer: null, slicerSearch: '', sort: {}, tablePages: {}, min: null, max: null };
const MONTHS = ['JANEIRO','FEVEREIRO','MARÇO','ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO'];
const DAYS = ['domingo','segunda-feira','terça-feira','quarta-feira','quinta-feira','sexta-feira','sábado'];
let toastTimer = 0;

function esc(value) { return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch])); }
function fixEncoding(value) {
  if (typeof value !== 'string' || !/[ÃÂ]/.test(value)) return value;
  try { const bytes=[]; for (const ch of value) { const n=ch.codePointAt(0); if (n<=255) bytes.push(n); else return value; } return new TextDecoder('utf-8',{fatal:true}).decode(new Uint8Array(bytes)); } catch { return value; }
}
function norm(value) { return String(fixEncoding(value ?? '')).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim(); }
function num(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  let s=String(value ?? '').trim().replace(/[^\d,.-]/g,'');
  if (!s) return 0;
  if (s.includes(',') && s.includes('.')) s=s.replace(/\./g,'').replace(',','.');
  else if (s.includes(',')) s=s.replace(',','.');
  const n=Number(s); return Number.isFinite(n)?n:0;
}
function dateKey(value) {
  if (value instanceof Date && !Number.isNaN(+value)) return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;
  if (typeof value === 'number' && window.XLSX?.SSF) { const d=XLSX.SSF.parse_date_code(value); if(d) return `${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`; }
  const s=String(fixEncoding(value ?? '')).trim();
  let m=/^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s); if(m) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  m=/^(\d{1,2})[\/.](\d{1,2})[\/.](\d{2,4})/.exec(s); if(m){let y=+m[3];if(y<100)y+=2000;return `${y}-${String(+m[2]).padStart(2,'0')}-${String(+m[1]).padStart(2,'0')}`;}
  return '';
}
function dateObj(key) { if(!key)return null; const [y,m,d]=key.split('-').map(Number); return new Date(y,m-1,d); }
function formatDate(key) { const d=dateObj(key); return d ? `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}` : '—'; }
function formatNum(n) { return n==null || !Number.isFinite(+n) ? '—' : Math.round(+n).toLocaleString('pt-BR'); }
function formatPct(n) { return n==null || !Number.isFinite(+n) ? '—' : (+n*100).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+'%'; }
function formatMoney(n) { return n==null || !Number.isFinite(+n) ? '—' : (+n).toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function isPct(metric) { return ['On time','Aderência','AderênciaComZero','Apontamento','Impacto Aderência'].includes(metric); }
function monthOfDate(key) { const d=dateObj(key); return d ? d.getMonth() : -1; }
function operation(priority) {
  const p=Number(priority); const exact={75:'Agendamento',40:'Specialty',31:'TP Barra Mansa',20:'TP Barra Mansa',24:'Viagem CD',32:'Viagem CD',29:'TP São Pedro',16:'TP São Pedro',15:'TP Friburgo',26:'TP Friburgo',11:'TP Itaocara',27:'TP Itaocara',9:'TP Campos',46:'TP Campos',81:'Baixada',71:'Zona Norte',61:'Zona Oeste',55:'Zona Sul',43:'Niterói - SG',82:'Grandes Contas',80:'GS Farma'};
  return exact[p] || (p>=47&&p<=54?'TMO 700':p>=64&&p<=69?'DML Rua':p>=83&&p<=89?'DML Shopping':'Cadastrar');
}
function zone(priority) {
  const p=Number(priority); const exact={43:'Dist NIT/SG',55:'Dist Zona Sul',61:'Dist Zona Oeste',71:'Dist Zona Norte',81:'Dist Baixada',9:'TP Campos',46:'TP Campos',11:'TP Itaocara',27:'TP Itaocara',15:'TP Friburgo',26:'TP Friburgo',16:'TP Sao Pedro',29:'TP Sao Pedro',24:'Viagem CD',32:'Viagem CD',20:'Barra Mansa',31:'Barra Mansa',40:'Specialty',75:'Agendamento'};
  return p>=83&&p<90?'DMIL Shopping':p>=64&&p<70?'DMIL Rua':p>=47&&p<53?'Tamoio NIT/SG':p>=53&&p<55?'Maricá/Alcântara':(exact[p]||'');
}
function transformRow(input) {
  const r={}; for(const [k,v] of Object.entries(input)) r[String(fixEncoding(k)).trim()]=typeof v==='string'?fixEncoding(v.trim()):v;
  const dt=dateKey(r.DATA_ITINERARIO); if(!dt)return null;
  const priority=num(r.PRIORIDADE), transport=String(r.DESCRICAO_TRANSPSEC||'').trim(), transfer=String(r.TRANSF||'').trim();
  const occurrence=String(r['TIPO_OCORRÊNCIA']||r['TIPO_OCORRÃŠNCIA']||'').trim();
  const network=String(r.REDE_CLIENTE??''), brand=String(r.DESCRICAO_REDE??'');
  const ap=String(r.DATA_OCORRENCIA??'').trim();
  const sourceSegment=norm(r.SEGMENTO);
  const explicitSegment=sourceSegment.includes('varejo')?'VAREJO':sourceSegment.includes('distribu')?'DISTRIBUIÇÃO':null;
  const day=dateObj(dt), first=new Date(day.getFullYear(),day.getMonth(),1);
  const week=Math.floor((day.getDate()-1+first.getDay())/7)+1;
  const result={...r,_dt:dt,_ot:norm(r.ON_TIME),_ce:norm(r.ENTREGA_DENTRO_CERCA),_sl:norm(r.ENTREGA_DENTRO_SLA),_ft:norm(r.LINK_FOTO),_sa:!ap,_vo:num(r.VOLUME),_en:dt+String(r.CODIGO_INTERNO_CLIENTE??''),
    'Semana do Mês':week,'Semana do mes txt':'Sem '+week,'Dia Itinerário':DAYS[day.getDay()],'Ordem Dia':day.getDay()||7,
    SEGMENTO:explicitSegment||([1155,263].includes(num(r.REDE_CLIENTE))?'VAREJO':'DISTRIBUIÇÃO'),'Zona Prioridade':zone(priority),'Operação':operation(priority),
    'Transp BI':/^PARCEIRO\b/i.test(transport)?'PARCEIRO':transport,'TP BI':transfer,
    NOME_MOTORISTA:String(r.NOME_MOTORISTA||'').slice(0,22).toUpperCase(),
    'REDE CONCAT':network+' - '+brand,'CLIENTE CONCAT':String(r.CODIGO_INTERNO_CLIENTE??'')+' - '+String(r.CLIENTE??''),
    'COLUNA APONTAMENTO':ap?'OK':'SEM APONTAMENTO','EMISSAO ABREVIADA':String(r.EMISSAO_DOCUMENTO??'').slice(0,5),VALOR:num(r.VALOR),_occurrence:occurrence};
  return result;
}
function prepareRows(rawRows) {
  const sorted=[...rawRows].sort((a,b)=>dateKey(a.DATA_ITINERARIO).localeCompare(dateKey(b.DATA_ITINERARIO)));
  const seen=new Set(), out=[];
  for(const raw of sorted){
    const r=transformRow(raw); if(!r)continue;
    if(norm(r.STATUS_ITINERARIO)!=='criado'||norm(r.ROTA_STATUS)==='cancelado')continue;
    const op=operation(num(r.PRIORIDADE)); if(op==='Grandes Contas'||op==='Specialty')continue;
    if(!String(r.DESCRICAO_TRANSPSEC||'').trim())continue;
    if(r._occurrence==='CARGA SINISTRADA')continue;
    const key=String(r.REGIAO_DOCUMENTO??'')+String(r.DOCUMENTO??'')+String(r.EMISSAO_DOCUMENTO??'');
    if(seen.has(key))continue; seen.add(key); out.push(r);
  }
  return out;
}
function measure(rows) {
  let noTime=0,late=0,ok=0,noAp=0,ray=0,notInFence=0,nok=0,volume=0,openValue=0;
  const clients=new Set(), deliveries=new Set();
  for(const r of rows){
    if(r._ot==='sim')noTime++; else if(r._ot===''||r._ot==='nao')late++;
    if(r._ce==='sim')ok++;
    if(r._sa){noAp++;openValue+=r.VALOR;}
    if(r._ce==='nao'){notInFence++;if(r._sl==='sim')ray++;if(r._ft==='sim')nok++;}
    volume+=r._vo; clients.add(r.CODIGO_INTERNO_CLIENTE); deliveries.add(r._en);
  }
  const n=rows.length, adherent=n?ok/n:null;
  return {'Total NFs':n,'No Prazo':noTime,'Fora do prazo':late,'On time':noTime+late?noTime/(noTime+late):null,OK:ok,'Aderência':adherent,'AderênciaComZero':adherent,'OkComZERO':ok,'SemApontComZero':noAp,'Apontamento':n?n/(n+noAp):null,'Raio NOK':ray,'Total fora':notInFence,NOK:nok,'Impacto Aderência':n?ray/n:null,ClientesDia:clients.size,Volumetria:volume,Entregas:deliveries.size,'NFs em aberto':noAp,'Valor aberto':openValue};
}
function displayMetric(metric,value) { return isPct(metric)?formatPct(value):metric==='Valor aberto'||metric==='VALOR'?formatMoney(value):formatNum(value); }
function sameValue(a,b) { return a===b || (a!=null&&b!=null&&String(a)===String(b)); }
function slicerValueLabel(value) { return value==null || value==='' ? '(Em branco)' : String(value); }
function selectionMatches(row, values) { return Object.entries(values||{}).every(([k,v])=>sameValue(row[k],v)); }
function filteredRows(pageFilter=null,excludeSource=null) {
  let rows=state.rows.filter(r=>(!state.from||r._dt>=state.from)&&(!state.to||r._dt<=state.to));
  for(const [field,set] of Object.entries(state.filters)) if(set&&set.size) rows=rows.filter(r=>set.has(r[field]));
  if(state.cross&&state.cross.source!==excludeSource) rows=rows.filter(r=>selectionMatches(r,state.cross.values));
  if(pageFilter) rows=rows.filter(r=>r['COLUNA APONTAMENTO']===pageFilter);
  return rows;
}
function setCross(source, values) {
  if(state.cross&&state.cross.source===source&&selectionMatches(state.cross.values,values)&&selectionMatches(values,state.cross.values)) state.cross=null;
  else state.cross={source,values};
  state.openSlicer=null; render();
}
function clearFilters() { state.filters={};state.cross=null;state.dateTouched=false;state.from=state.min;state.to=state.max;state.tablePages={};render(); }
function groupRows(rows,keyFn,sorter) {
  const m=new Map(); for(const r of rows){const k=keyFn(r);if(!m.has(k))m.set(k,[]);m.get(k).push(r);}
  const a=[...m.entries()].map(([key,items])=>({key,rows:items})); return a.sort(sorter||((x,y)=>String(x.key).localeCompare(String(y.key),'pt',{numeric:true})));
}
function visualElement(v,index) {
  const d=document.createElement('div');d.className='visual'+(v.sh?' shadow':'');d.dataset.visual=`${state.page}_${index}`;
  d.style.cssText=`left:${v.x}px;top:${v.y}px;width:${v.w}px;height:${v.h}px;z-index:${v.z||index+1}`;
  return d;
}
function ccColor(rules,value){
  if(value==null||!Number.isFinite(+value))return null;
  const ops=[(a,b)=>a===b,(a,b)=>a>b,(a,b)=>a>=b,(a,b)=>a<b,(a,b)=>a<=b];
  for(const [conds,color] of rules){ if(conds.every(([op,t])=>ops[op](+value,t))) return color; }
  return null;
}
function renderCard(v,rows,d) {
  const m=measure(rows);const value=m[v.m];d.classList.add('card-visual');const condMeasure=(v.m==='SemApontComZero'||v.m==='Valor aberto')?'Aderência':v.m;d.style.color=(v.cc&&ccColor(v.cc,m[condMeasure]))||v.c||'#006a4e';
  d.innerHTML=`<div class="card-value">${esc(displayMetric(v.m,value))}</div><div class="card-label">${esc(v.l||'')}</div>`;
  d.title=v.m==='AderênciaComZero'?'Aderência = OK / Total de NFs após filtros; medida DAX do PBIX ainda não foi confirmada.':v.l||'';
}
function colorFor(metric,value) { if(metric==='On time')return value>=.88?'#006440':'#D64550'; if(/Volum|Entreg/.test(metric))return '#006440'; return '#D64550'; }
function makeCategories(v,rows) {
  let groups;
  if(v.cat==='H') groups=groupRows(rows,r=>r._dt,(a,b)=>String(a.key).localeCompare(String(b.key)));
  else if(v.cat==='Dia Itinerário') groups=groupRows(rows,r=>r['Ordem Dia'],(a,b)=>a.key-b.key);
  else if(v.k==='col') groups=groupRows(rows,r=>String(r.NOME_MOTORISTA||'')+'|'+String(r['Transp BI']||''),()=>0);
  else groups=groupRows(rows,r=>r[v.cat],(a,b)=>String(a.key??'').localeCompare(String(b.key??''),'pt',{numeric:true}));
  let cats=groups.map(g=>{
    const first=g.rows[0]||{};let label=slicerValueLabel(g.key),group='';let values={};
    if(v.cat==='H'){label=String(Number(g.key.slice(8,10))).padStart(2,'0');const dt=dateObj(g.key);group=dt?MONTHS[dt.getMonth()].toLowerCase():'';values={_dt:g.key};}
    else if(v.cat==='Dia Itinerário'){label=first['Dia Itinerário']||label;values={'Ordem Dia':g.key};}
    else if(v.k==='col'){label=first.NOME_MOTORISTA||'(Em branco)';group=first['Transp BI']||'';values={NOME_MOTORISTA:first.NOME_MOTORISTA,'Transp BI':first['Transp BI']};}
    else values={[v.cat]:g.key};
    const mm=measure(g.rows);return {key:g.key,label,group,rows:g.rows,base:mm[v.ym],line:v.ym2?mm[v.ym2]:null,values};
  });
  if(v.k==='col')cats=cats.sort((a,b)=>(b.base||0)-(a.base||0)).slice(0,15);
  return cats;
}
function selectedCategory(c,source) { return !!(state.cross&&state.cross.source===source&&selectionMatches(c.values,state.cross.values)&&selectionMatches(state.cross.values,c.values)); }
function renderColumnChart(v,cats,d,source) {
  const W=v.w,H=v.h-50,pt=8,pb=49,ph=Math.max(20,H-pt-pb),n=Math.max(cats.length,1),step=(W-18)/n,bw=step*.78,mx=Math.max(1,...cats.map(c=>c.base||0));let svg='';
  cats.forEach((c,i)=>{const x=9+step*i+step/2,h=(c.base||0)/mx*ph,y=pt+ph-h,fill=selectedCategory(c,source)?'#94313b':'#D64550';let label=c.label.trim();const mc=Math.max(6,Math.floor(step/7.4)-1);if(label.length>mc)label=label.slice(0,mc-1)+'…';
    const labelInside=h>=28,labelY=labelInside?Math.min(pt+ph-5,y+15):Math.max(pt+12,y-6);svg+=`<g class="chart-point${selectedCategory(c,source)?' is-selected':''}" data-i="${i}" role="button" tabindex="0" aria-label="${esc(c.label)}: ${formatNum(c.base)}"><rect x="${x-bw/2}" y="${y}" width="${bw}" height="${Math.max(h,1)}" fill="${fill}"/>${c.base>0?`<text x="${x}" y="${labelY}" text-anchor="middle" font-size="13" font-weight="700" fill="${labelInside?'#fff':'#252423'}">${formatNum(c.base)}</text>`:''}<text x="${x}" y="${pt+ph+17}" text-anchor="middle" font-size="11" fill="#777">${esc(c.group)}</text><text x="${x}" y="${pt+ph+34}" text-anchor="middle" font-size="11" fill="#777">${esc(label)}</text></g>`;
  });
  d.innerHTML=`<div class="chart-title" style="font-size:21px">${esc((v.ttl||'').replace(/\s+/g,' '))}</div><svg width="${W}" height="${H}" role="img" aria-label="${esc(v.ttl||'Gráfico')}">${svg}</svg>`;
  wireChartPoints(d,cats,source);
}
function renderChart(v,rows,d,source) {
  const cats=makeCategories(v,rows); if(!cats.length){d.innerHTML=`<div class="chart-title">${esc(v.ttl||'')}</div><div class="chart-empty">Sem dados para os filtros selecionados</div>`;return;}
  if(v.k==='col'){renderColumnChart(v,cats,d,source);return;}
  const W=v.w,H=v.h-(v.ttl?56:6),pt=28,sy=['NFs em aberto','Valor aberto'].includes(v.ym),pb=cats.some(c=>c.group)?(sy?66:48):34,ph=Math.max(24,H-pt-pb),n=cats.length,step=(W-24)/n,bw=n===1?step*.21:step*(['On Time x Aderência','On Time','Volumetria','Clientes Atendidos'].includes((v.ttl||'').trim())?.83:.76),ax=v.ax||[];
  const lo=Number.isFinite(ax[0])?ax[0]:0,rawMax=Math.max(0,...cats.map(c=>c.base||0)),hi=(Number.isFinite(ax[1])&&ax[1]>lo)?ax[1]:Math.max(1,rawMax*1.15),lo2=Number.isFinite(ax[2])?ax[2]:0,hi2=(Number.isFinite(ax[3])&&ax[3]>lo2)?ax[3]:1;
  const Y=x=>pt+ph-(x-lo)/(hi-lo)*ph,Y2=x=>pt+ph-(x-lo2)/(hi2-lo2)*ph;let bars='',labels='',points=[],groups=[];
  cats.forEach((c,i)=>{const x=12+step*i+step/2,base=c.base||0,y=Y(base),bh=Math.max(0,pt+ph-y),rot=n>14||bw<46,fmt=isPct(v.ym)?formatPct:v.ym==='Valor aberto'?(x=>'R$ '+formatMoney(x)):formatNum,fill=colorFor(v.ym,base);
    bars+=`<g class="chart-point${selectedCategory(c,source)?' is-selected':''}" data-i="${i}" role="button" tabindex="0" aria-label="${esc(c.label)}: ${esc(fmt(base))}${c.line!=null?' · '+esc(formatPct(c.line)):''}"><rect x="${x-bw/2}" y="${y}" width="${bw}" height="${Math.max(bh,1)}" fill="${fill}"/>`;
    const t=fmt(base),inside=!v.out&&bh>=28, safeY=inside?Math.min(pt+ph-5,y+15):Math.max(pt+12,y-6);if(inside){labels+=`<text transform="translate(${x+(rot?4:0)},${safeY})${rot?' rotate(-90)':''}" ${rot?'text-anchor="end"':'text-anchor="middle"'} font-size="11" font-weight="700" fill="#fff">${esc(t)}</text>`;}else{labels+=`<text x="${x}" y="${safeY}" text-anchor="middle" font-size="12" font-weight="700" fill="#252423">${esc(t)}</text>`;}
    if(c.line!=null){points.push([x,Y2(c.line)]);labels+=`<text x="${x}" y="${Y2(c.line)-8}" text-anchor="middle" font-size="11" font-weight="700">${esc(formatPct(c.line))}</text>`;}
    let label=c.label;if(v.k==='col'&&label.length>17)label=label.slice(0,16)+'…';bars+=`<text x="${x}" y="${pt+ph+16}" text-anchor="middle" font-size="${v.k==='col'?11:12}" font-weight="700" fill="#6f7471">${esc(label)}</text>`;
    if(c.group){const gs=groups[groups.length-1];if(gs&&gs.name===c.group)gs.x2=x;else groups.push({name:c.group,x1:x,x2:x,year:String(c.key).slice(0,4)});} bars+='</g>';
  });
  groups.forEach(g=>{const cx=(g.x1+g.x2)/2;labels+=`<text x="${cx}" y="${pt+ph+34}" text-anchor="middle" font-size="11" font-weight="700" fill="#666">${esc(g.name)}</text>`;if(sy)labels+=`<text x="${(12+W-12)/2}" y="${pt+ph+54}" text-anchor="middle" font-size="11" font-weight="700" fill="#666">${esc(g.year)}</text>`;});
  const line=points.length?`<polyline fill="none" stroke="#222" stroke-width="2" stroke-dasharray="6 4" points="${points.map(p=>p.join(',')).join(' ')}"/>`:'';
  d.innerHTML=`${v.ttl?`<div class="chart-title">${esc(v.ttl)}</div>`:''}<svg width="${W}" height="${H}" role="img" aria-label="${esc(v.ttl||'Gráfico')}">${line}${bars}${labels}</svg>`;
  wireChartPoints(d,cats,source);
}
function wireChartPoints(root,cats,source) {
  root.querySelectorAll('.chart-point').forEach(el=>{const act=()=>{const c=cats[Number(el.dataset.i)];if(c)setCross(source,c.values);};el.addEventListener('click',e=>{e.stopPropagation();act();});el.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();act();}});});
}
const CELL_STYLES={'No Prazo':['#777','#d9ffc3'],'Fora do prazo':['#d64550','#faf0f0'],'On time':['#094780','#e8ecf3'],'Total fora':['#777'],'OK':['#777'],'NOK':['#d64550']};
function renderTable(v,rows,d,source) {
  const dims=v.cols.filter(c=>!c[2]),m=new Map();
  for(const r of rows){const key=JSON.stringify(dims.map(c=>r[c[0]]??''));if(!m.has(key))m.set(key,[]);m.get(key).push(r);}
  let grouped=[...m.values()].map(items=>{const row={};dims.forEach(c=>row[c[0]]=items[0][c[0]]);return Object.assign(row,measure(items));});
  const sort=state.sort[source]||[v.so,v.sd||'Ascending'],key=sort[0],dir=sort[1]==='Ascending'?1:-1;
  if(key)grouped.sort((a,b)=>{const x=a[key],y=b[key];return dir*(typeof x==='number'&&typeof y==='number'?x-y:String(x??'').localeCompare(String(y??''),'pt',{numeric:true}));});
  const totals=measure(rows),max={};v.cols.forEach(c=>{if(c[2])max[c[0]]=Math.max(1e-9,...grouped.map(r=>r[c[0]]||0));});
  const pageSize=60,totalPages=Math.max(1,Math.ceil(grouped.length/pageSize)),page=Math.min(state.tablePages[source]||0,totalPages-1),shown=grouped.slice(page*pageSize,(page+1)*pageSize);
  const cell=(c,r,isTotal=false)=>{const prop=c[0],value=r[prop],sty=CELL_STYLES[prop]||[],metric=c[2];let text=metric?displayMetric(prop,value):prop==='VALOR'?formatMoney(num(value)):value instanceof Date?formatDate(dateKey(value)):esc(value??'');let style='';if(sty[0]&&!isTotal)style+=`color:${sty[0]};`;if(metric&&(prop==='On time'||prop==='Total NFs'||sty[1]))style+='text-align:right;';if(sty[1]&&!isTotal&&value)style+=`background:linear-gradient(90deg,${sty[1]} ${isPct(prop)?value*100:(value/max[prop])*100}%,transparent 0);`;return `<td class="${metric?'num':''}" style="${style}">${text}</td>`;};
  const trs=shown.map(row=>{const values=Object.fromEntries(dims.map(c=>[c[0],row[c[0]]]));const selected=state.cross&&state.cross.source===source&&selectionMatches(state.cross.values,values)&&selectionMatches(values,state.cross.values);return `<tr data-row="${grouped.indexOf(row)}" class="${selected?'selected':''}" aria-selected="${!!selected}">${v.cols.map(c=>cell(c,row)).join('')}</tr>`;}).join('');
  d.innerHTML=`<div class="table-wrap"><table class="report-table"><thead><tr>${v.cols.map(c=>`<th data-key="${esc(c[0])}">${esc(c[1])}${key===c[0]?(dir>0?' ▲':' ▼'):''}</th>`).join('')}</tr></thead><tbody>${trs||`<tr><td colspan="${v.cols.length}">Sem dados para os filtros selecionados</td></tr>`}</tbody><tfoot><tr>${v.cols.map((c,i)=>i===0?'<td>Total</td>':c[2]?cell(c,totals,true):'<td></td>').join('')}</tr></tfoot></table></div><div class="table-pager"><span>${grouped.length.toLocaleString('pt-BR')} linhas · página ${page+1}/${totalPages}</span><button data-page="prev" ${page===0?'disabled':''}>Anterior</button><button data-page="next" ${page>=totalPages-1?'disabled':''}>Próxima</button></div>`;
  d.querySelectorAll('th[data-key]').forEach(th=>th.addEventListener('click',()=>{const p=th.dataset.key;state.sort[source]=[p,key===p&&dir>0?'Descending':'Ascending'];render();}));
  d.querySelectorAll('tr[data-row]').forEach(tr=>tr.addEventListener('click',()=>{const row=grouped[Number(tr.dataset.row)];const values=Object.fromEntries(dims.map(c=>[c[0],row[c[0]]]));setCross(source,values);}));
  d.querySelectorAll('button[data-page]').forEach(b=>b.addEventListener('click',()=>{state.tablePages[source]=Math.max(0,Math.min(totalPages-1,page+(b.dataset.page==='next'?1:-1)));render();}));
}
function renderSlicer(v,d) {
  const slicerId=d.dataset.visual;
  d.classList.add('slicer');d.classList.toggle('slicer-open',state.openSlicer===slicerId);const field=v.f;d.innerHTML=`<div class="slicer-head">${esc(v.l||field)}</div>`;
  if(field==='DATA_ITINERARIO'){
    d.innerHTML+=`<div class="date-range"><input type="date" aria-label="Data inicial" value="${esc(state.from||'')}" min="${esc(state.min||'')}" max="${esc(state.max||'')}" data-date="from"><input type="date" aria-label="Data final" value="${esc(state.to||'')}" min="${esc(state.min||'')}" max="${esc(state.max||'')}" data-date="to"></div>`;
    d.querySelectorAll('input[type=date]').forEach(inp=>inp.addEventListener('change',()=>{state[inp.dataset.date]=inp.value||null;state.dateTouched=true;if(state.from&&state.to&&state.from>state.to){const old=state.from;state.from=state.to;state.to=old;}state.tablePages={};render();}));return;
  }
  const options=[...new Set(state.rows.map(r=>r[field]).filter(x=>x!==null&&x!==undefined))].sort((a,b)=>String(a).localeCompare(String(b),'pt',{numeric:true}));const selected=state.filters[field];
  const label=!selected?'Tudo':selected.size===1?slicerValueLabel([...selected][0]):`${selected.size} selecionados`;
  d.innerHTML+=`<button class="slicer-button" type="button"><span>${esc(label)}</span><span aria-hidden="true">⌄</span></button>`;
  d.querySelector('.slicer-button').addEventListener('click',e=>{e.stopPropagation();state.openSlicer=state.openSlicer===slicerId?null:slicerId;state.slicerSearch='';render();});
  if(state.openSlicer!==slicerId)return;
  const pop=document.createElement('div');pop.className='slicer-popup';pop.addEventListener('click',e=>e.stopPropagation());
  const q=norm(state.slicerSearch),visible=options.filter(x=>norm(x).includes(q));
  pop.innerHTML=`<input type="search" placeholder="Pesquisar" value="${esc(state.slicerSearch)}"><label><input type="checkbox" data-all ${!selected?'checked':''}>Selecionar tudo</label>`+visible.map((value,i)=>`<label><input type="checkbox" data-option="${i}" ${!selected||selected.has(value)?'checked':''}>${esc(slicerValueLabel(value))}</label>`).join('');
  const search=pop.querySelector('input[type=search]');search.addEventListener('input',()=>{state.slicerSearch=search.value;render();const next=d.querySelector('.slicer-popup input[type=search]');next?.focus();next?.setSelectionRange(next.value.length,next.value.length);});
  pop.querySelector('[data-all]').addEventListener('change',()=>{delete state.filters[field];state.openSlicer=slicerId;render();});
  pop.querySelectorAll('[data-option]').forEach(inp=>inp.addEventListener('change',()=>{let set=selected?new Set(selected):new Set(options);const value=visible[Number(inp.dataset.option)];if(inp.checked)set.add(value);else set.delete(value);if(!set.size||set.size===options.length)delete state.filters[field];else state.filters[field]=set;state.openSlicer=slicerId;render();}));
  d.appendChild(pop);
}
function renderVisual(v,index,page) {
  const d=visualElement(v,index);const source=`${state.page}_${index}`;const rows=filteredRows(page.pf,source);
  if(v.t==='text'){d.className='visual text-visual';d.style.color=v.c||'#006a4e';d.style.fontSize=`${(v.sz||14)*1.333}px`;d.style.justifyContent=v.al==='center'?'center':'flex-start';d.textContent=v.tx||'';}
  else if(v.t==='img'){d.classList.add('img-visual');d.innerHTML='<img src="profarma-logo.png" alt="Profarma">';}
  else if(v.t==='line'){d.style.background='transparent';d.innerHTML='<div class="rule-line"></div>';}
  else if(v.t==='box'){d.style.background='transparent';d.innerHTML='<div class="shape-box"></div>';}
  else if(v.t==='card')renderCard(v,rows,d);
  else if(v.t==='ch')renderChart(v,rows,d,source);
  else if(v.t==='tb')renderTable(v,rows,d,source);
  else if(v.t==='sl')renderSlicer(v,d);
  return d;
}
function updateInfo() {
  const loaded=state.files.length, active=state.files.filter(f=>f.on).length, info=$('#info');
  if(!state.rows.length){info.textContent=loaded?`${loaded} arquivo(s) carregado(s); sem registros válidos após os filtros de preparação.`:'Carregue uma planilha para começar.';return;}
  info.textContent=`${state.rows.length.toLocaleString('pt-BR')} linhas válidas · ${active}/${loaded} planilha(s) ativa(s) · ${formatDate(state.from||state.min)} a ${formatDate(state.to||state.max)} · dados neste navegador`;
}
function renderTabs() {
  const tabs=$('#tabs');let html='';
  if(state.cross)html+=`<span class="active-filter" title="${esc(Object.entries(state.cross.values).map(([k,v])=>`${k}: ${v}`).join(' · '))}">Filtro: ${esc(Object.entries(state.cross.values).map(([k,v])=>`${k}=${v}`).join(' · '))}</span><button class="nav-action" id="clearSelection" type="button">Limpar seleção</button>`;
  html+=`<button class="nav-action primary" id="loadButton" type="button">Carregar planilhas</button>`;
  state.pages.forEach((p,i)=>{html+=`<button class="tab ${i===state.page?'active':''}" data-tab="${i}" type="button">${esc(p.n)}</button>`;});
  tabs.innerHTML=html;
  $('#loadButton').addEventListener('click',()=>$('#fileInput').click());
  $('#clearSelection')?.addEventListener('click',()=>{state.cross=null;render();});
  tabs.querySelectorAll('[data-tab]').forEach(b=>b.addEventListener('click',()=>{state.page=Number(b.dataset.tab);state.cross=null;state.openSlicer=null;state.tablePages={};render();window.scrollTo(0,0);}));
}
function render() {
  const page=state.pages[state.page];if(!page)return;
  const canvas=$('#canvas');const H=Math.max(...page.v.map(v=>v.y+v.h))+30;canvas.replaceChildren();canvas.style.width=page.w+'px';canvas.style.height=H+'px';canvas.style.background=page.bg||'#fff';
  if(state.rows.length){page.v.forEach((v,i)=>canvas.appendChild(renderVisual(v,i,page)));addShotButton(canvas,page);}
  else {const empty=document.createElement('section');empty.id='empty';empty.innerHTML='<h2>Dashboard Profarma</h2><p>Adicione a planilha de operação para carregar os indicadores, filtros e tabelas. A planilha será processada e armazenada somente neste navegador.</p><button type="button" id="emptyUpload">Selecionar planilha</button>';canvas.appendChild(empty);$('#emptyUpload').addEventListener('click',()=>$('#fileInput').click());}
  const mobile=matchMedia('(max-width:700px)').matches;const scale=mobile?1:Math.max(.35,Math.min(2.5,document.documentElement.clientWidth/page.w));canvas.style.transform=`scale(${scale})`;$('#viewport').style.height=`${H*scale}px`;$('#viewport').style.overflowX=mobile?'auto':'hidden';
  renderTabs();updateInfo();
}
function csvRows(text) {
  if(text.charCodeAt(0)===0xfeff)text=text.slice(1);const first=text.split(/\r?\n/,1)[0]||'';const delim=(first.split(';').length>first.split(',').length)?';':',';
  const rows=[];let row=[],field='',quoted=false;
  for(let i=0;i<text.length;i++){const c=text[i];if(quoted){if(c==='"'&&text[i+1]==='"'){field+='"';i++;}else if(c==='"')quoted=false;else field+=c;}else if(c==='"')quoted=true;else if(c===delim){row.push(field);field='';}else if(c==='\n'){row.push(field.replace(/\r$/,''));rows.push(row);row=[];field='';}else field+=c;}
  if(field||row.length){row.push(field.replace(/\r$/,''));rows.push(row);}const headers=(rows.shift()||[]).map(x=>x.trim());return rows.filter(r=>r.length>1).map(r=>Object.fromEntries(headers.map((h,i)=>[h,fixEncoding(r[i]??'')])));
}
async function readRows(name,buffer) {
  if(/\.xlsx?$/i.test(name)){
    if(!window.XLSX)throw new Error('Não foi possível carregar a biblioteca de leitura Excel. Verifique sua conexão e tente novamente.');
    const wb=XLSX.read(buffer,{type:'array',cellDates:true});const ws=wb.Sheets['1']||wb.Sheets[wb.SheetNames[0]];if(!ws)throw new Error('A planilha não contém uma aba legível.');
    return XLSX.utils.sheet_to_json(ws,{defval:'',raw:true,blankrows:false}).map(row=>Object.fromEntries(Object.entries(row).map(([k,v])=>[String(fixEncoding(k)).trim(),typeof v==='string'?fixEncoding(v):v])));
  }
  const bytes=new Uint8Array(buffer);let text=new TextDecoder('utf-8').decode(bytes);if(text.includes('\uFFFD'))text=new TextDecoder('windows-1252').decode(bytes);return csvRows(text);
}
function monthFolder(name,rows) {
  const n=norm(name),idx=MONTHS.findIndex(m=>n.includes(norm(m)));let year=(String(name).match(/20\d{2}/)||[])[0]||'';
  const dates=[...new Set(rows.map(r=>dateKey(r.DATA_ITINERARIO)).filter(Boolean))];const months=new Set(dates.map(d=>d.slice(0,7)));
  let month=idx;
  if(month<0&&months.size===1){const [y,m]=[...months][0].split('-').map(Number);year=year||String(y);month=m-1;}
  if(month<0){const stamp=String(name).match(/(20\d{2})(0[1-9]|1[0-2])([0-3]\d)/);if(stamp){year=year||stamp[1];month=Number(stamp[2])-1;}}
  if(month<0)return months.size?'VÁRIOS MESES':'SEM MÊS';
  if(!year){const years=dates.filter(d=>Number(d.slice(5,7))===month+1).map(d=>Number(d.slice(0,4)));if(years.length)year=String(Math.max(...years));}
  return MONTHS[month]+(year?' '+year:'');
}
function idbOpen() { return new Promise((resolve,reject)=>{if(!window.indexedDB){reject(new Error('Armazenamento local não disponível neste navegador.'));return;}const req=indexedDB.open('profarma-dashboard-test',1);req.onupgradeneeded=()=>req.result.createObjectStore('files',{keyPath:'name'});req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);}); }
async function idbAll() { const db=await idbOpen();return new Promise((resolve,reject)=>{const tx=db.transaction('files','readonly'),req=tx.objectStore('files').getAll();req.onsuccess=()=>resolve(req.result||[]);req.onerror=()=>reject(req.error);}); }
async function idbSave(file) { const db=await idbOpen();return new Promise((resolve,reject)=>{const tx=db.transaction('files','readwrite');tx.objectStore('files').put({name:file.name,buffer:file.buffer,lm:file.lm,on:file.on,folder:file.folder});tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);}); }
async function idbRemove(name) { const db=await idbOpen();return new Promise((resolve,reject)=>{const tx=db.transaction('files','readwrite');tx.objectStore('files').delete(name);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error);}); }
async function addFiles(fileList) {
  const errors=[];
  for(const file of fileList){try{const buffer=await file.arrayBuffer(),rows=await readRows(file.name,buffer),f={name:file.name,buffer,lm:file.lastModified||Date.now(),on:true,folder:monthFolder(file.name,rows),rawRows:rows};state.files=state.files.filter(x=>x.name!==f.name);state.files.push(f);await idbSave(f);}catch(e){errors.push(`${file.name}: ${e.message||'erro de leitura'}`);}}
  state.openSlicer=null;rebuild();render();if(errors.length)showToast('Alguns arquivos não foram carregados: '+errors.join(' | '));
}
function rebuild() {
  const active=[...state.files].filter(f=>f.on).sort((a,b)=>b.lm-a.lm),owner=new Map();
  for(const f of active){for(const r of f.rawRows){const d=dateKey(r.DATA_ITINERARIO);if(d&&!owner.has(d))owner.set(d,f);}}
  const raw=[];for(const f of active)for(const r of f.rawRows){const d=dateKey(r.DATA_ITINERARIO);if(d&&owner.get(d)===f)raw.push(r);}
  state.rows=prepareRows(raw);state.min=state.rows.reduce((a,r)=>!a||r._dt<a?r._dt:a,null);state.max=state.rows.reduce((a,r)=>!a||r._dt>a?r._dt:a,null);
  if(!state.dateTouched){state.from=state.min;state.to=state.max;}
  state.cross=null;state.tablePages={};
}
function showToast(message) { const old=$('.toast');old?.remove();const t=document.createElement('div');t.className='toast';t.textContent=message;document.body.appendChild(t);clearTimeout(toastTimer);toastTimer=setTimeout(()=>t.remove(),5000); }

function filePeriod(file) {
  const dates=(file.rawRows||[]).map(r=>dateKey(r.DATA_ITINERARIO)).filter(Boolean).sort();
  return dates.length ? {from:dates[0],to:dates[dates.length-1]} : null;
}
function individualTitle(file) {
  const p=filePeriod(file), folder=file.folder||'SEM MÊS';
  return p ? `ANALISE ${folder} DO DIA ${formatDate(p.from)} ATE ${formatDate(p.to)}` : String(file.name||'itineraries').replace(/\.[^.]+$/,'');
}
function workbookRows(files) {
  const rows=[];
  for(const f of files) for(const r of (f.rawRows||[])) rows.push({...r});
  return rows;
}
function xlsxBlob(rows, sheetName='itineraries') {
  if(!window.XLSX) throw new Error('A biblioteca Excel ainda não foi carregada.');
  const clean=rows.map(row=>Object.fromEntries(Object.entries(row).filter(([k])=>!String(k).startsWith('_'))));
  const headers=[...new Set(clean.flatMap(r=>Object.keys(r)))];
  const normalized=clean.map(r=>Object.fromEntries(headers.map(h=>[h,r[h]??''])));
  const ws=XLSX.utils.json_to_sheet(normalized,{header:headers,skipHeader:false});
  ws['!cols']=headers.map(h=>({wch:Math.min(38,Math.max(12,String(h).length+2))}));
  ws['!autofilter']={ref:`A1:${XLSX.utils.encode_col(Math.max(headers.length-1,0))}${Math.max(normalized.length+1,1)}`};
  ws['!freeze']={xSplit:0,ySplit:1};
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,(sheetName||'itineraries').slice(0,31));
  return new Blob([XLSX.write(wb,{bookType:'xlsx',type:'array',compression:true})],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function downloadWorkbook(rows,filename) {
  if(!rows.length){showToast('Não há linhas para exportar.');return;}
  const url=URL.createObjectURL(xlsxBlob(rows,'itineraries')),a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),15000);
}
function monthFiles(folder) { return state.files.filter(f=>f.folder===folder).sort((a,b)=>(b.lm||0)-(a.lm||0)); }
function latestRowsForFolder(folder) {
  const files=monthFiles(folder), owner=new Map();
  for(const f of files) for(const r of (f.rawRows||[])){const d=dateKey(r.DATA_ITINERARIO);if(d&&!owner.has(d))owner.set(d,f);}
  const out=[]; for(const f of files) for(const r of (f.rawRows||[])){const d=dateKey(r.DATA_ITINERARIO);if(!d||owner.get(d)===f)out.push(r);}
  return out;
}
function showPreview(file) {
  const root=$('#modalRoot'), rows=file.rawRows||[], headers=[...new Set(rows.flatMap(r=>Object.keys(r)))].filter(k=>!String(k).startsWith('_')), shown=rows.slice(0,120);
  root.innerHTML=`<div id="previewModal" role="dialog" aria-modal="true"><div class="preview-box"><div class="settings-head"><span>${esc(individualTitle(file))}</span><button data-action="close-preview" aria-label="Fechar">×</button></div><div class="preview-body"><p>${rows.length.toLocaleString('pt-BR')} linhas · mostrando ${shown.length.toLocaleString('pt-BR')} para visualização.</p><div class="table-wrap"><table class="file-preview"><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${shown.map(r=>`<tr>${headers.map(h=>`<td>${esc(r[h]??'')}</td>`).join('')}</tr>`).join('')}</tbody></table></div></div><div class="settings-foot"><button data-action="download-file" data-name="${esc(file.name)}">Baixar XLSX</button><button data-action="close-preview">Fechar</button></div></div></div>`;
}

const CRC_TABLE=(()=>{const t=new Uint32Array(256);for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=c&1?0xedb88320^(c>>>1):c>>>1;t[n]=c>>>0;}return t;})();
function crc32(bytes){let c=0xffffffff;for(const b of bytes)c=CRC_TABLE[(c^b)&255]^(c>>>8);return(c^0xffffffff)>>>0;}
function zipBlob(files,folderMode) {
  const enc=new TextEncoder(),parts=[],central=[];let offset=0;
  for(const f of files){const safe=String(f.name).replace(/[\\/\0]/g,'_'),folder=String(f.folder||'SEM MÊS').replace(/[\\/\0]/g,'_'),path=folderMode?folder+'/'+safe:safe,name=enc.encode(path),data=new Uint8Array(f.buffer),crc=crc32(data),h=new Uint8Array(30),v=new DataView(h.buffer);
    v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x0800,true);v.setUint16(8,0,true);v.setUint16(10,0,true);v.setUint16(12,0x21,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,data.length,true);v.setUint16(26,name.length,true);v.setUint16(28,0,true);parts.push(h,name,data);
    const c=new Uint8Array(46),d=new DataView(c.buffer);d.setUint32(0,0x02014b50,true);d.setUint16(4,20,true);d.setUint16(6,20,true);d.setUint16(8,0x0800,true);d.setUint16(10,0,true);d.setUint16(12,0,true);d.setUint16(14,0x21,true);d.setUint32(16,crc,true);d.setUint32(20,data.length,true);d.setUint32(24,data.length,true);d.setUint16(28,name.length,true);d.setUint16(30,0,true);d.setUint16(32,0,true);d.setUint16(34,0,true);d.setUint16(36,0,true);d.setUint32(38,0,true);d.setUint32(42,offset,true);central.push(c,name);offset+=h.length+name.length+data.length;
  }
  const centralSize=central.reduce((n,x)=>n+x.length,0),end=new Uint8Array(22),e=new DataView(end.buffer);e.setUint32(0,0x06054b50,true);e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);e.setUint32(12,centralSize,true);e.setUint32(16,offset,true);return new Blob([...parts,...central,end],{type:'application/zip'});
}
function downloadZip(files,folderMode,filename) { if(!files.length)return;const url=URL.createObjectURL(zipBlob(files,folderMode)),a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),15000); }
function renderSettings() {
  const root=$('#modalRoot'),folders=[...new Set(state.files.map(f=>f.folder||'SEM MÊS'))].sort((a,b)=>a.localeCompare(b,'pt'));
  const sections=folders.map(folder=>{const files=monthFiles(folder);return `<section class="month-folder"><div class="folder-head"><strong>📁 ${esc(folder)}</strong><span class="small-muted">${files.length} planilha(s) · dashboard usa a mais recente por data</span><button class="primary" data-action="download-month" data-folder="${esc(folder)}">Baixar ${esc(folder)}.xlsx</button></div><table class="file-list"><thead><tr><th>Usar</th><th>Planilha individual</th><th>Linhas</th><th>Período</th><th>Ações</th></tr></thead><tbody>${files.map(f=>{const ds=filePeriod(f);return `<tr><td><input type="checkbox" data-action="toggle" data-name="${esc(f.name)}" ${f.on?'checked':''}></td><td><strong>${esc(individualTitle(f))}</strong><br><span class="small-muted">origem: ${esc(f.name)}</span></td><td>${(f.rawRows||[]).length.toLocaleString('pt-BR')}</td><td>${ds?formatDate(ds.from)+' a '+formatDate(ds.to):'—'}</td><td><button data-action="view" data-name="${esc(f.name)}">Visualizar</button> <button data-action="download-file" data-name="${esc(f.name)}">Baixar</button> <button data-action="delete" data-name="${esc(f.name)}">Excluir</button></td></tr>`;}).join('')}</tbody></table></section>`;}).join('');
  root.innerHTML=`<div id="settingsModal" role="dialog" aria-modal="true" aria-labelledby="settingsTitle"><div class="settings-box"><div class="settings-head"><span id="settingsTitle">⚙ Planilhas e pastas mensais</span><button data-action="close" aria-label="Fechar">×</button></div><div class="settings-body"><p>Os dados ficam neste navegador. Cada mês pode ser baixado como um único arquivo Excel formatado, enquanto as planilhas individuais continuam disponíveis para visualizar, baixar ou excluir.</p><div class="settings-actions"><button class="primary" data-action="add">+ Adicionar planilha</button></div>${sections||'<div class="empty-folder">Nenhuma planilha carregada. Adicione um arquivo Excel ou CSV.</div>'}</div><div class="settings-foot"><button data-action="close">Fechar</button></div></div></div>`;
}

function closeSettings(){ $('#modalRoot').replaceChildren(); }
$('#settingsButton').addEventListener('click',renderSettings);
$('#modalRoot').addEventListener('click',async e=>{if(e.target.id==='settingsModal'||e.target.closest('[data-action="close"]')){closeSettings();return;}const b=e.target.closest('[data-action]');if(!b)return;const action=b.dataset.action;
  if(action==='add'){$('#fileInput').click();return;}
  if(action==='close-preview'){closeSettings();return;}
  if(action==='download-month'){const folder=b.dataset.folder;downloadWorkbook(latestRowsForFolder(folder),`${folder}.xlsx`);return;}
  if(action==='view'){const f=state.files.find(x=>x.name===b.dataset.name);if(f)showPreview(f);return;}
  if(action==='download-file'){const f=state.files.find(x=>x.name===b.dataset.name);if(f)downloadWorkbook(f.rawRows||[],`${individualTitle(f)}.xlsx`);return;}
  if(action==='delete'){if(!confirm(`Excluir ${b.dataset.name} deste navegador?`))return;state.files=state.files.filter(f=>f.name!==b.dataset.name);await idbRemove(b.dataset.name);rebuild();render();renderSettings();return;}
});
$('#modalRoot').addEventListener('change',async e=>{const t=e.target;if(t.dataset.action==='move'){const f=state.files.find(x=>x.name===t.dataset.name);if(f){f.folder=t.value;await idbSave(f);renderSettings();}}else if(t.dataset.action==='toggle'){const f=state.files.find(x=>x.name===t.dataset.name);if(f){f.on=t.checked;await idbSave(f);rebuild();render();renderSettings();}}});
$('#fileInput').addEventListener('change',async e=>{if(e.target.files.length)await addFiles([...e.target.files]);e.target.value='';if($('#settingsModal'))renderSettings();});
let dragCount=0;window.addEventListener('dragenter',e=>{e.preventDefault();dragCount++;$('#dropOverlay').style.display='flex';});window.addEventListener('dragleave',()=>{dragCount=Math.max(0,dragCount-1);if(!dragCount)$('#dropOverlay').style.display='none';});window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',async e=>{e.preventDefault();dragCount=0;$('#dropOverlay').style.display='none';if(e.dataTransfer.files.length)await addFiles([...e.dataTransfer.files]);});
document.addEventListener('click',e=>{if(state.openSlicer&&!e.target.closest('.slicer-popup')&&!e.target.closest('.slicer-button')){state.openSlicer=null;render();}});
window.addEventListener('resize',render);

/* ===== Copiar print da tela atual (vai direto para a área de transferência) ===== */
const SHOT_HTML_PROPS=['display','position','left','top','right','bottom','width','height','min-width','min-height','max-width','max-height','margin-top','margin-right','margin-bottom','margin-left','padding-top','padding-right','padding-bottom','padding-left','box-sizing','overflow-x','overflow-y','z-index','color','background-color','background-image','border-top-width','border-top-style','border-top-color','border-right-width','border-right-style','border-right-color','border-bottom-width','border-bottom-style','border-bottom-color','border-left-width','border-left-style','border-left-color','border-top-left-radius','border-top-right-radius','border-bottom-left-radius','border-bottom-right-radius','border-collapse','box-shadow','font-family','font-size','font-weight','font-style','line-height','letter-spacing','text-align','text-overflow','text-transform','white-space','vertical-align','flex-direction','flex-wrap','flex-grow','flex-shrink','align-items','justify-content','gap','opacity','object-fit','object-position'];
const SHOT_SVG_PROPS=['fill','stroke','stroke-width','stroke-dasharray','font-family','font-size','font-weight','font-style','text-anchor','opacity'];
function addShotButton(canvas,page) {
  const t=page.v.find(v=>v.t==='text'&&v.al==='left'&&v.y>=40&&v.y<=70)||{x:50,y:59,w:157,h:57};
  const b=document.createElement('button');b.type='button';b.className='shot-btn';b.dataset.nocapture='1';b.title='Copiar um print desta tela para colar (Ctrl+V) em outro lugar';
  b.textContent='📋 Copiar print';b.style.left=`${t.x+t.w+10}px`;b.style.top=`${t.y+Math.max(0,(t.h-30)/2)}px`;
  b.addEventListener('click',e=>{e.stopPropagation();copyShot(b);});canvas.appendChild(b);
}
function blobToDataURL(blob){return new Promise((ok,no)=>{const r=new FileReader();r.onload=()=>ok(r.result);r.onerror=()=>no(r.error);r.readAsDataURL(blob);});}
async function pageToBlob() {
  const canvas=$('#canvas'),W=canvas.offsetWidth,H=canvas.offsetHeight;
  const clone=canvas.cloneNode(true);
  const keep=el=>!el.closest('[data-nocapture],.slicer-popup');
  const src=[canvas,...canvas.querySelectorAll('*')].filter(keep),dst=[clone,...clone.querySelectorAll('*')].filter(keep);
  src.forEach((s,i)=>{const d=dst[i];if(!d)return;const cs=getComputedStyle(s),svg=s instanceof SVGElement;let css='';
    for(const p of (svg?SHOT_SVG_PROPS:SHOT_HTML_PROPS)){const v=cs.getPropertyValue(p);if(v)css+=`${p}:${v};`;}
    d.setAttribute('style',css);});
  clone.querySelectorAll('[data-nocapture],.slicer-popup').forEach(e=>e.remove());
  clone.querySelectorAll('input[type=date]').forEach(inp=>{const sp=document.createElement('div');sp.setAttribute('style',inp.getAttribute('style')||'');sp.style.display='flex';sp.style.alignItems='center';sp.textContent=formatDate(inp.value);inp.replaceWith(sp);});
  for(const img of clone.querySelectorAll('img')){try{const r=await fetch(img.getAttribute('src'));img.setAttribute('src',await blobToDataURL(await r.blob()));}catch{}}
  clone.setAttribute('style',`position:relative;left:0;top:0;margin:0;width:${W}px;height:${H}px;background:#fff;overflow:hidden;font-family:"Segoe UI",Arial,sans-serif;color:#252423;`);
  const xhtml=new XMLSerializer().serializeToString(clone);
  const svg=`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}"><foreignObject x="0" y="0" width="${W}" height="${H}">${xhtml}</foreignObject></svg>`;
  const img=new Image();await new Promise((ok,no)=>{img.onload=ok;img.onerror=()=>no(new Error('não consegui desenhar a tela'));img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);});
  const k=Math.min(2,Math.max(1,window.devicePixelRatio||1)),cv=document.createElement('canvas');cv.width=Math.round(W*k);cv.height=Math.round(H*k);
  const ctx=cv.getContext('2d');ctx.fillStyle='#fff';ctx.fillRect(0,0,cv.width,cv.height);ctx.scale(k,k);ctx.drawImage(img,0,0);
  return await new Promise((ok,no)=>cv.toBlob(b=>b?ok(b):no(new Error('falha ao gerar a imagem')),'image/png'));
}
function downloadBlob(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),15000);}
async function copyShot(btn) {
  const label='📋 Copiar print',pageName=(state.pages[state.page]?.n||'tela').toLowerCase();
  btn.disabled=true;btn.textContent='Copiando…';
  try {
    const promise=pageToBlob();
    if(navigator.clipboard&&window.ClipboardItem){await navigator.clipboard.write([new ClipboardItem({'image/png':promise})]);showToast('✔ Print copiado! É só colar (Ctrl+V).');}
    else{downloadBlob(await promise,`print-${pageName}.png`);showToast('Este navegador não permite copiar imagem direto; o print foi baixado.');}
  } catch(e) {
    try{downloadBlob(await pageToBlob(),`print-${pageName}.png`);showToast('O navegador bloqueou a cópia direta; baixei o print em PNG.');}
    catch(e2){showToast('Não foi possível gerar o print: '+(e2.message||'erro'));}
  } finally { btn.disabled=false;btn.textContent=label; }
}
async function init() {
  try {
    const response=await fetch('report-spec.json',{cache:'no-store'});if(!response.ok)throw new Error('Não foi possível carregar a configuração visual.');
    const spec=await response.json();state.pages=spec.pages||[];
    try{for(const saved of await idbAll()){const rawRows=await readRows(saved.name,saved.buffer);state.files.push({...saved,rawRows});}}catch(e){showToast('O armazenamento local não foi restaurado: '+(e.message||'erro'));}
    rebuild();render();
  } catch(e) { $('#info').textContent=e.message||'Falha ao abrir o dashboard.';showToast(e.message||'Falha ao abrir o dashboard.'); }
}
init();
