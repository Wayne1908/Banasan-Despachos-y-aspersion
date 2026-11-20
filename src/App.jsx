import React,{useState,useRef,useEffect} from 'react';
import htm from 'htm';
import ExcelJS from 'exceljs/dist/exceljs.min.js';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {createClient} from '@supabase/supabase-js';

pdfjsLib.GlobalWorkerOptions.workerSrc=pdfWorker;
const html=htm.bind(React.createElement);

/* ---- Supabase (puedes sobreescribir con VITE_SUPABASE_URL y VITE_SUPABASE_ANON_KEY) ---- */
const SB_URL=import.meta.env.VITE_SUPABASE_URL||'https://ywykyfatumkfjwyembqr.supabase.co';
const SB_KEY=import.meta.env.VITE_SUPABASE_ANON_KEY||'sb_publishable_TYpKSrMMdqwh05g7j67VPw_SLRLqeqc';
const sb=createClient(SB_URL,SB_KEY);

const CSS=`*,::before,::after{border-color:#e2e8f0}:root{color-scheme:light}body{background:#f1f5f9;color:#0f172a;font-family:Calibri,'Segoe UI',system-ui,sans-serif;margin:0}@keyframes sp{to{transform:rotate(360deg)}}.spin{animation:sp 1s linear infinite}td,th{font-size:13px}`;

const BLUE='#0054A6',YEL='#FFD100';
const CONV=['CEIBA','PORVENIR','VILLA BEATRIZ','BRETAÑITA','LAS MERCEDES','JOSEFINA MT','FABLISKA','ARENAL','VIJAGUAL','ANGELES','MARTE','SAN PEDRO','RUBI'];
const ORG=['NARANJITOS','DESPENSA','MANANTIAL','GISELLE BEATRIZ','PLANTACION'];
const ALL=[...CONV,...ORG];
const D=['LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO','DOMINGO'];
const DL=['LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES','SÁBADO','DOMINGO'];
const MO=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
const norm=s=>String(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase();
const P2=n=>String(n).padStart(2,'0');
const addD=(d,n)=>new Date(d.getTime()+n*864e5);
const fR=d=>`${P2(d.getUTCDate())} ${MO[d.getUTCMonth()]}`;
const fD=d=>`${P2(d.getUTCDate())}/${P2(d.getUTCMonth()+1)}/${d.getUTCFullYear()}`;
const fX=d=>`${d.getUTCDate()}/${d.getUTCMonth()+1}/${d.getUTCFullYear()}`;
const isoW1=y=>{const j=new Date(Date.UTC(y,0,4));return addD(j,-((j.getUTCDay()+6)%7))};
const isoWk=(d,m,y)=>{const dt=new Date(Date.UTC(y,m,d));const day=(dt.getUTCDay()+6)%7;dt.setUTCDate(dt.getUTCDate()-day+3);const w1=new Date(Date.UTC(dt.getUTCFullYear(),0,4));return 1+Math.round(((dt-w1)/864e5-3+(w1.getUTCDay()+6)%7)/7)};
const mkWeeks=s=>Array.from({length:52},(_,i)=>({n:i+1,start:addD(s,i*7)}));
const wl=w=>`SEM ${w.n}: ${fR(w.start)} - ${fR(addD(w.start,6))}`;
const bdg=f=>ORG.includes(f)?'bg-emerald-100 text-emerald-800':'bg-sky-100 text-sky-800';

const IC={upload:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>',
download:'<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" x2="12" y1="15" y2="3"/>',
plus:'<path d="M5 12h14"/><path d="M12 5v14"/>',pencil:'<path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/>',x:'<path d="M18 6 6 18"/><path d="m6 6 12 12"/>',
file:'<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/><path d="M16 13H8"/><path d="M16 17H8"/>',
ok:'<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="m9 11 3 3L22 4"/>',
load:'<path d="M21 12a9 9 0 1 1-6.219-8.56"/>',
warn:'<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>',
trash:'<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>'};
const Ic=(n,s=18,c='')=>html`<svg width=${s} height=${s} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" class=${c} dangerouslySetInnerHTML=${{__html:IC[n]}}/>`;

/* ---- Filtro estricto de las 18 fincas oficiales de Banasan ---- */
const CN={};ALL.forEach(f=>CN[norm(f)]=f);
const PAT=[...Object.keys(CN).map(k=>[k,CN[k]]),['LA CEIBA','CEIBA'],['VIJAGUAL BUFFER','VIJAGUAL'],['MERCEDES','LAS MERCEDES'],['JOSEFINA MARIA TERESA','JOSEFINA MT'],['JOSEFINA M T','JOSEFINA MT'],['MARIA TERESA','JOSEFINA MT'],['JOSEFINA','JOSEFINA MT'],['GISELLE','GISELLE BEATRIZ'],['LOS ANGELES','ANGELES'],['NARANJITO','NARANJITOS'],['MARTES','MARTE']].sort((a,b)=>b[0].length-a[0].length);
const findF=(t,mt)=>{let s=' '+norm(t).replace(/[._\-–—/]+/g,' ').replace(/\s+/g,' ')+' ';
s=s.replace(/\b(COOMULBANANO|EMPREBANCOOP|BANAFRUCOOP|COOBAMAG|COOBAFRIO|UNIBAN|TECNIBAN)\s+[A-Z0-9]+(?:\s+[A-Z0-9]+)?\b/g,' ')
   .replace(/\bGLORIA\s+MERCEDES\b/g,' ')
   .replace(/\bMERCEDES\s+II\b/g,' ')
   .replace(/\b(?!LA\b)[A-Z0-9]+\s+CEIBA\b/g,' ')
   .replace(/\b[A-Z0-9]+\s+VIJAGUAL\b/g,' ')
   .replace(/\bVIJAGUAL\s+(?!BUFFER\b)[A-Z]{2,}\b/g,' ');
const r=[];for(const [k,f] of PAT){if(k==='MARTES'&&!mt)continue;const re=new RegExp('(?<![A-Z])'+k+'(?![A-Z])','g');if(re.test(s)){s=s.replace(re,' ');if(!r.includes(f))r.push(f)}}return r};

function parseAsp(t){const out=[];let wk=null,dy=null;for(const raw of t.split(/\r?\n/)){const line=norm(raw).replace(/(MERCEDES)\s*-?\s*[A-D](?:\s*[,/&-]\s*[A-D](?![A-Z]))*(?![A-Z])/g,'$1');if(!line.trim())continue;const m=line.match(/SEM(?:ANA)?\s*\.?\s*(\d{1,2})(?!\d)/);if(m){wk=+m[1];dy=null}
const cells=line.split(/[,;\t|]/);
if(m&&cells.length>=4){cells.slice(1,8).forEach((c,i)=>findF(c,true).forEach(f=>out.push([wk,i,f])));continue}
let rest=m?line.replace(m[0],' '):line;
const dm=rest.trim().match(/^(LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO|DOMINGO)\b/);
if(dm){dy=D.indexOf(dm[1]);rest=rest.trim().slice(dm[0].length)}
if(wk!==null&&dy!==null)findF(rest,false).forEach(f=>out.push([wk,dy,f]))}
return out}

/* ---- Renderizado de PDF en alta nitidez e IA ---- */
async function pdfInfo(file,max,img){const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;const blobs=[];let text='';
for(let i=1;i<=Math.min(pdf.numPages,max);i++){const p=await pdf.getPage(i);const tc=await p.getTextContent();const rows={};
tc.items.forEach(it=>{const y=Math.round(it.transform[5]/5)*5;(rows[y]=rows[y]||[]).push(it)});
text+=Object.keys(rows).sort((a,b)=>b-a).map(y=>rows[y].sort((a,b)=>a.transform[4]-b.transform[4]).map(x=>x.str).join(' ')).join('\n')+'\n';
if(img){const v=p.getViewport({scale:2.2});const c=document.createElement('canvas');c.width=v.width;c.height=v.height;await p.render({canvasContext:c.getContext('2d'),viewport:v}).promise;blobs.push(await new Promise(r=>c.toBlob(r,'image/png')))}}
return {text,blobs}}

async function shrink(b){try{const bm=await createImageBitmap(b);const k=Math.min(1,2400/Math.max(bm.width,bm.height));const c=document.createElement('canvas');c.width=Math.round(bm.width*k);c.height=Math.round(bm.height*k);const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(bm,0,0,c.width,c.height);return await new Promise(r=>c.toBlob(r,'image/jpeg',0.9))}catch(e){return b}}
const b64=b=>new Promise(r=>{const fr=new FileReader();fr.onload=()=>r(String(fr.result).split(',')[1]);fr.readAsDataURL(b)});

async function ai(prompt,images){try{const imgs=[];for(const b of images||[]){const s=await shrink(b);imgs.push({type:s.type||'image/jpeg',data:await b64(s)})}
const r=await fetch('/api/leer',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prompt,images:imgs})});
if(!r.ok)return null;return (await r.json()).data}catch(e){return null}}

const isImg=f=>/\.(jpe?g|png|webp)$/i.test(f.name)||f.type.startsWith('image/');
const isPdf=f=>/\.pdf$/i.test(f.name);

const aspPrompt=t=>`Analiza con máxima precisión visual esta imagen de un Programa de Aspersión Semanal de fincas bananeras.
Puede venir en dos formatos:
1) Reporte por empresa fumigadora (columnas FECHA a la izquierda con celdas combinadas verticalmente, y FINCA a la derecha). Mira con cuidado las líneas horizontales de los bordes de cada celda combinada en la columna FECHA para saber exactamente qué fecha corresponde a cada fila de FINCA.
2) Matriz semanal (columnas SEMANAS, LUNES, MARTES, MIERCOLES, JUEVES, VIERNES, SABADO, DOMINGO).

Responde ÚNICAMENTE un arreglo JSON válido con este esquema exacto:
[{"date":"YYYY-MM-DD","week":45,"day":"VIERNES","farm":"NOMBRE EXACTO DE LA CELDA FINCA"}]

Reglas estrictas:
- Si la fila tiene una fecha visible (ej: "viernes, 6 de noviembre de 2026"), pon "date":"2026-11-06" (ignora la fecha de "Vigencia" del pie/encabezado del formato). Si no hay fecha visible, pon "date":null.
- En "day" escribe el día en mayúsculas sin tildes: LUNES, MARTES, MIERCOLES, JUEVES, VIERNES, SABADO o DOMINGO.
- En "farm" copia TEXTUALMENTE el nombre de la celda FINCA sin recortarlo ni abreviarlo (por ejemplo, si dice "EMPREBANCOOP CEIBA" o "COOMULBANANO VIJAGUAL", escríbelo completo; si dice "LAS MERCEDES A", escríbelo completo).
- No omitas ninguna fila de la tabla ni inventes datos.${t?'\n\nTexto auxiliar:\n'+t.slice(0,4000):''}`;

/* ---- componentes ---- */
function FL({files}){return html`<ul class="mt-3 space-y-1.5">${files.map(f=>html`<li key=${f.id} class="flex items-start gap-2 text-sm bg-slate-50 border rounded px-2 py-1.5">
<span class=${f.st==='ok'?'text-emerald-600':f.st==='err'?'text-red-600':'text-blue-600'}>${Ic(f.st==='ok'?'ok':f.st==='err'?'warn':'load',18,f.st==='proc'?'spin':'')}</span>
<span class="min-w-0"><b class="break-all">${f.name}</b><br/><span class="text-xs text-slate-500">${f.note}</span></span></li>`)}</ul>`}
function Drop({title,sub,onFiles,files,children}){const [o,setO]=useState(false);const r=useRef();
return html`<div class="bg-white rounded-xl shadow border border-slate-200 p-5">
<h3 class="font-bold text-lg" style=${{color:BLUE}}>${title}</h3><p class="text-sm text-slate-500 mb-3">${sub}</p>
<div onDragOver=${e=>{e.preventDefault();setO(true)}} onDragLeave=${()=>setO(false)} onDrop=${e=>{e.preventDefault();setO(false);onFiles([...e.dataTransfer.files])}} onClick=${()=>r.current.click()}
class=${'cursor-pointer rounded-lg border-2 border-dashed p-6 text-center flex flex-col items-center gap-1 '+(o?'bg-yellow-50 border-yellow-400':'border-slate-300 hover:border-blue-500')}>
<span style=${{color:BLUE}}>${Ic('upload',30)}</span><b>Arrastra y suelta o haz clic</b><span class="text-xs text-slate-500">JPG · PNG · PDF · TXT · CSV — uno o varios archivos</span></div>
<input ref=${r} type="file" multiple accept=".jpg,.jpeg,.png,.pdf,.txt,.csv,image/*,application/pdf,text/plain,text/csv" class="hidden" onChange=${e=>{onFiles([...e.target.files]);e.target.value=''}}/>
${children}<${FL} files=${files}/></div>`}

export default function App(){
const [tab,setTab]=useState('carga'),[weeks,setWeeks]=useState([]),[restr,setRestr]=useState({}),[calF,setCalF]=useState([]),[aspF,setAspF]=useState([]);
const [provs,setProvs]=useState([]),[ap,setAp]=useState(''),[newP,setNewP]=useState(null),[desp,setDesp]=useState({}),[conf,setConf]=useState(0),[edP,setEdP]=useState(null),[delP,setDelP]=useState(false);
const [flt,setFlt]=useState(''),[onlyB,setOnlyB]=useState(false),[form,setForm]=useState({w:'',d:'0',placa:'',oc:'',finca:''});
const [err,setErr]=useState(''),[ok,setOk]=useState(''),[toast,setToast]=useState(''),[yr,setYr]=useState(2026);
const upd=(set,id,p)=>set(a=>a.map(x=>x.id===id?{...x,...p}:x));
const afterW=ws=>{setWeeks(ws);const t=Date.now();const c=ws.find(w=>t>=w.start.getTime()&&t<w.start.getTime()+7*864e5);setForm(f=>({...f,w:String((c||ws[0]).n)}))};
const setF=(k,v)=>setForm(f=>({...f,[k]:v}));

/* ---- Sincronización con Supabase ---- */
const [cloud,setCloud]=useState('cargando'),[st,setSt]=useState('cargando'),[tick,setTick]=useState(0);
const just=useRef(false),ready=useRef(false),lastJ=useRef(''),dirty=useRef(false);
const snap=()=>({v:1,cal:weeks.length?weeks[0].start.toISOString():null,restr,provs,ap,desp,conf});
const load=async first=>{try{const {data,error}=await sb.from('estado_banasan').select('datos').eq('id','principal').maybeSingle();
if(error)throw error;
if(data&&data.datos){const d=data.datos;just.current=true;
if(d.cal){const ws=mkWeeks(new Date(d.cal));first?afterW(ws):setWeeks(ws)}else setWeeks([]);
setRestr({...(d.restr||{})});setProvs([...(d.provs||[])]);setAp(d.ap||'');setDesp({...(d.desp||{})});setConf(d.conf||0)}
else just.current=!ready.current;
ready.current=true;setCloud('listo');setSt('guardado')}
catch(e){if(!ready.current){setCloud('error');setSt('error');setToast('No se pudo leer la nube: '+(e.message||e))}}};
useEffect(()=>{load(true);const f=()=>{if(!dirty.current&&document.visibilityState==='visible')load(false)};
document.addEventListener('visibilitychange',f);window.addEventListener('focus',f);
return()=>{document.removeEventListener('visibilitychange',f);window.removeEventListener('focus',f)}},[]);
useEffect(()=>{if(cloud!=='listo')return;const j=JSON.stringify(snap());
if(just.current){just.current=false;lastJ.current=j;return}
if(j===lastJ.current)return;
dirty.current=true;setSt('guardando');
const t=setTimeout(async()=>{try{const {error}=await sb.from('estado_banasan').upsert({id:'principal',datos:JSON.parse(j)});
if(error)throw error;lastJ.current=j;dirty.current=false;setSt('guardado')}
catch(e){setSt('error');setToast('No se pudo guardar en la nube: '+(e.message||e))}},600);
return()=>clearTimeout(t)},[weeks,restr,provs,ap,desp,conf,cloud,tick]);

const doCal=async fs=>{for(const f of fs){const id=Math.random();setCalF(a=>[...a,{id,name:f.name,st:'proc',note:'Procesando…'}]);
try{let text='',imgs=[];if(isImg(f))imgs=[f];else if(isPdf(f)){const r=await pdfInfo(f,2,true);text=r.text;imgs=r.blobs}else text=await f.text();
const y=+((text+' '+f.name).match(/20\d\d/)||[2026])[0];let start=null,via=`regla ISO del año ${y} (SEM 1 = semana del 1 de enero)`;
const j=await ai(`Este archivo es el Calendario Bananero (SEM 1 a SEM 52, semanas de lunes a domingo). Responde SOLO un JSON {"week1Start":"YYYY-MM-DD"} con el lunes en que inicia la SEM 1; verifícalo contando semanas desde una fecha visible.${text?'\n\nTexto:\n'+text.slice(0,6000):''}`,imgs);
if(j&&j.week1Start){const d=new Date(j.week1Start+'T00:00:00Z');if(!isNaN(d)&&d.getUTCDay()===1){start=d;via='leído con IA'}}
if(!start)start=isoW1(y);afterW(mkWeeks(start));
upd(setCalF,id,{st:'ok',note:`SEM 1–52 habilitadas (${via}). SEM 1: ${fR(start)}`})}
catch(e){upd(setCalF,id,{st:'err',note:'No se pudo leer: '+(e.message||e)})}}};

const doAsp=async fs=>{for(const f of fs){const id=Math.random();setAspF(a=>[...a,{id,name:f.name,st:'proc',note:'Leyendo…'}]);
try{let text='',imgs=[],res=[],via='';
if(isImg(f))imgs=[f];else if(isPdf(f)){const r=await pdfInfo(f,40,true);text=r.text;imgs=r.blobs}else text=await f.text();
let fail=0,totalRows=0;
const fromAI=async()=>{const out=[],tx=isImg(f)||isPdf(f)?'':text;
for(let i=0;i<Math.max(imgs.length,1);i++){
  if(imgs.length>1)upd(setAspF,id,{note:`Analizando página ${i+1} de ${imgs.length} con IA…`});
  const j=await ai(aspPrompt(tx),imgs.slice(i,i+1));
  if(!Array.isArray(j)){fail++;continue}
  totalRows+=j.length;
  j.forEach(x=>{
    let w=+x.week,d=D.indexOf(norm(x.day||''));
    const dm=String(x.date||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if(dm){const yy=+dm[1],mm=+dm[2]-1,dd=+dm[3],dt=new Date(Date.UTC(yy,mm,dd));if(!isNaN(dt)){d=(dt.getUTCDay()+6)%7;w=isoWk(dd,mm,yy)}}
    const fs=findF(x.farm||'',!dm);
    if(w>0&&w<54&&d>=0)fs.forEach(fi=>out.push([w,d,fi]));
  });
}
return out};

if(isImg(f)||isPdf(f)){res=await fromAI();via='visión IA'}
else{res=parseAsp(text);via='texto';if(!res.length){res=await fromAI();via='IA'}}

if(!res.length){
  if((isImg(f)||isPdf(f))&&fail===0&&totalRows>0){
    upd(setAspF,id,{st:'ok',note:'0 fincas de Banasan en este archivo (solo cooperativas o terceros)'});continue;
  }
  upd(setAspF,id,{st:'err',note:'No se pudo conectar con la IA (/api/leer). Revisa que exista api/leer.js en la raíz del repo y la variable GEMINI_API_KEY en Vercel.'});continue;
}
setRestr(r=>{const n={...r};res.forEach(([w,d,fi])=>{const k=w+'|'+d,a=n[k]||[];if(!a.includes(fi))n[k]=[...a,fi]});return n});
upd(setAspF,id,{st:'ok',note:`${res.length} fincas-día de Banasan cargadas (${via}${imgs.length>1?', '+imgs.length+' págs':''})`+(fail?` · ${fail} pág(s) con error de red`:'')})}
catch(e){upd(setAspF,id,{st:'err',note:'Error: '+(e.message||e)})}}};

const rm=(w,d,f)=>setRestr(r=>({...r,[w+'|'+d]:(r[w+'|'+d]||[]).filter(x=>x!==f)}));
const add=(w,d,f)=>f&&setRestr(r=>({...r,[w+'|'+d]:[...(r[w+'|'+d]||[]),f]}));
const nB=Object.values(restr).reduce((a,b)=>a+b.length,0);
const wkOf=n=>weeks.find(w=>w.n===+n);
const dm=(ap&&desp[ap])||{};
const fw=wkOf(form.w);
const live=fw&&form.finca&&(restr[form.w+'|'+form.d]||[]).includes(form.finca);

const reg=()=>{setOk('');const {w,d,placa,oc,finca}=form;
if(!ap)return setErr('Primero crea o selecciona un proveedor.');
if(!fw||!oc.trim()||!finca)return setErr('Completa semana, día, OC y finca destino. La placa es opcional: puedes escribirla después en la tabla.');
if((restr[w+'|'+d]||[]).includes(finca)){setConf(c=>c+1);return setErr(`⛔ OPERACIÓN NO PERMITIDA: La finca ${finca} está en aspersión el ${DL[d]} ${fD(addD(fw.start,+d))}`)}
const k=w+'|'+d;setDesp(s=>({...s,[ap]:{...(s[ap]||{}),[k]:[...((s[ap]||{})[k]||[]),{placa:placa.trim().toUpperCase(),oc:oc.trim(),finca}]}}));
setErr('');setOk('Despacho registrado correctamente.');setForm(f=>({...f,placa:'',oc:''}))};
const delE=(k,i)=>setDesp(s=>({...s,[ap]:{...s[ap],[k]:s[ap][k].filter((_,j)=>j!==i)}}));
const addProv=()=>{const n=(newP||'').trim().toUpperCase();if(n&&!provs.includes(n)){setProvs(p=>[...p,n]);setAp(n)}setNewP(null)};
const renProv=()=>{const n=(edP||'').trim().toUpperCase();if(!n||n===ap){setEdP(null);return}if(provs.includes(n)){setToast('Ya existe un proveedor con ese nombre.');return}
setProvs(p=>p.map(x=>x===ap?n:x));setDesp(s=>{const {[ap]:cur,...rest}=s;return cur?{...rest,[n]:cur}:rest});setAp(n);setEdP(null)};
const delProv=()=>{const r=provs.filter(x=>x!==ap);setProvs(r);setDesp(s=>{const {[ap]:_x,...rest}=s;return rest});setAp(r[0]||'');setDelP(false)};
const setE=(k,i,p)=>setDesp(s=>({...s,[ap]:{...s[ap],[k]:s[ap][k].map((e,j)=>j===i?{...e,...p}:e)}}));

async function exportX(){try{
const wb=new ExcelJS.Workbook();const th={style:'thin',color:{argb:'FF000000'}},bd={top:th,left:th,bottom:th,right:th};
const put=(ws,r,c,v,o={})=>{const x=ws.getCell(r,c);x.value=v;x.border=bd;x.alignment={horizontal:o.l?'left':'center',vertical:'middle',wrapText:true};x.font={name:'Calibri',size:11,bold:!!o.b,color:{argb:o.col||'FF000000'}};if(o.bg)x.fill={type:'pattern',pattern:'solid',fgColor:{argb:o.bg}}};
const G='FFD9D9D9',Y='FFFFFF00';
const w1=wb.addWorksheet('Programacion_Despachos');w1.getColumn(1).width=8;
for(let d=0;d<5;d++){[10,10,23].forEach((w,i)=>w1.getColumn(2+d*3+i).width=w);for(let i=0;i<3;i++)put(w1,2,2+d*3+i,null,{b:
