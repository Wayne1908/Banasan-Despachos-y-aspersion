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

/* ---- reconocimiento de fincas ---- */
const CN={};ALL.forEach(f=>CN[norm(f)]=f);
const PAT=[...Object.keys(CN).map(k=>[k,CN[k]]),['LA CEIBA','CEIBA'],['VIJAGUAL BUFFER','VIJAGUAL'],['MERCEDES','LAS MERCEDES'],['JOSEFINA MARIA TERESA','JOSEFINA MT'],['JOSEFINA M T','JOSEFINA MT'],['MARIA TERESA','JOSEFINA MT'],['JOSEFINA','JOSEFINA MT'],['GISELLE','GISELLE BEATRIZ'],['LOS ANGELES','ANGELES'],['NARANJITO','NARANJITOS'],['MARTES','MARTE']].sort((a,b)=>b[0].length-a[0].length);
const findF=(t,mt)=>{let s=' '+norm(t).replace(/[._\-–—/]+/g,' ').replace(/\s+/g,' ')+' ';s=s.replace(/\b(COOMULBANANO|EMPREBANCOOP|BANAFRUCOOP|COOBAMAG|UNIBAN|TECNIBAN)\s+[A-Z0-9]+(?:\s+[A-Z0-9]+)?\b/g,' ').replace(/\bGLORIA\s+MERCEDES\b/g,' ').replace(/\b(?!LA\b)[A-Z0-9]+\s+CEIBA\b/g,' ').replace(/\b[A-Z0-9]+\s+VIJAGUAL\b/g,' ').replace(/\bVIJAGUAL\s+(?!BUFFER\b)[A-Z]{2,}\b/g,' ');const r=[];for(const [k,f] of PAT){if(k==='MARTES'&&!mt)continue;const re=new RegExp('(?<![A-Z])'+k+'(?![A-Z])','g');if(re.test(s)){s=s.replace(re,' ');if(!r.includes(f))r.push(f)}}return r};
function parseAsp(t){const out=[];let wk=null,dy=null;for(const raw of t.split(/\r?\n/)){const line=norm(raw).replace(/(MERCEDES)\s*-?\s*[A-D](?:\s*[,/&-]\s*[A-D](?![A-Z]))*(?![A-Z])/g,'$1');if(!line.trim())continue;const m=line.match(/SEM(?:ANA)?\s*\.?\s*(\d{1,2})(?!\d)/);if(m){wk=+m[1];dy=null}
const cells=line.split(/[,;\t|]/);
if(m&&cells.length>=4){cells.slice(1,8).forEach((c,i)=>findF(c,true).forEach(f=>out.push([wk,i,f])));continue}
let rest=m?line.replace(m[0],' '):line;
const dm=rest.trim().match(/^(LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO|DOMINGO)\b/);
if(dm){dy=D.indexOf(dm[1]);rest=rest.trim().slice(dm[0].length)}
if(wk!==null&&dy!==null)findF(rest,false).forEach(f=>out.push([wk,dy,f]))}
return out}

/* ---- PDF e IA ---- */
async function pdfInfo(file,max,img){const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;const blobs=[];let text='';const OPS=pdfjsLib.OPS;
const meses={ENERO:0,FEBRERO:1,MARZO:2,ABRIL:3,MAYO:4,JUNIO:5,JULIO:6,AGOSTO:7,SEPTIEMBRE:8,OCTUBRE:9,NOVIEMBRE:10,DICIEMBRE:11};
const isoWk=(d,m,y)=>{const dt=new Date(Date.UTC(y,m,d));const day=(dt.getUTCDay()+6)%7;dt.setUTCDate(dt.getUTCDate()-day+3);const w1=new Date(Date.UTC(dt.getUTCFullYear(),0,4));return 1+Math.round(((dt-w1)/864e5-3+(w1.getUTCDay()+6)%7)/7)};
const clusterY=(arr,tol)=>{const s=[...arr].sort((a,b)=>a.vy-b.vy);const out=[];for(const it of s){const g=out[out.length-1];if(g&&it.vy-g.yLast<=tol){g.items.push(it);g.yLast=it.vy}else out.push({y0:it.vy,yLast:it.vy,items:[it]})}return out};
for(let i=1;i<=Math.min(pdf.numPages,max);i++){const p=await pdf.getPage(i);const vp=p.getViewport({scale:1});const tc=await p.getTextContent();
const items=tc.items.filter(it=>it.str.trim()).map(it=>{const [vx,vy]=vp.convertToViewportPoint(it.transform[4],it.transform[5]);return {str:it.str,vx,vy}});
const rawStream=norm(items.map(it=>it.str).join(' '));
const sortedTxt=clusterY(items,5).map(g=>g.items.sort((a,b)=>a.vx-b.vx).map(x=>x.str).join(' ')).join('\n');
const cleanNoVig=(rawStream+' '+sortedTxt).replace(/VIGENCIA\s*:?\s*\d{1,2}\s*\/\s*\d{1,2}\s*\/\s*20\d\d/g,' ');
const yrM=cleanNoVig.match(/(?:ANO|DE)\s*(202\d)\b/)||cleanNoVig.match(/\b(202[4-9])\b/);const docYr=yrM?+yrM[1]:2026;
const opList=await p.getOperatorList();let ctm=[1,0,0,1,0,0],stk=[];const hSegs=[],vSegs=[];
const mul=(m1,m2)=>[m1[0]*m2[0]+m1[2]*m2[1],m1[1]*m2[0]+m1[3]*m2[1],m1[0]*m2[2]+m1[2]*m2[3],m1[1]*m2[2]+m1[3]*m2[3],m1[0]*m2[4]+m1[2]*m2[5]+m1[4],m1[1]*m2[4]+m1[3]*m2[5]+m1[5]];
const pt=(x,y)=>{const px=ctm[0]*x+ctm[2]*y+ctm[4],py=ctm[1]*x+ctm[3]*y+ctm[5];return vp.convertToViewportPoint(px,py)};
const addSeg=(x1,y1,x2,y2)=>{const [ax,ay]=pt(x1,y1),[bx,by]=pt(x2,y2);
if(Math.abs(ay-by)<3&&Math.abs(bx-ax)>20)hSegs.push({x1:Math.min(ax,bx),x2:Math.max(ax,bx),y:(ay+by)/2});
else if(Math.abs(ax-bx)<3&&Math.abs(by-ay)>35)vSegs.push({x:(ax+bx)/2,y1:Math.min(ay,by),y2:Math.max(ay,by)})};
const addRect=(x,y,w,h)=>{addSeg(x,y,x+w,y);addSeg(x,y+h,x+w,y+h);addSeg(x,y,x,y+h);addSeg(x+w,y,x+w,y+h)};
for(let k=0;k<opList.fnArray.length;k++){const fn=opList.fnArray[k],args=opList.argsArray[k];
if(fn===OPS.save)stk.push([...ctm]);else if(fn===OPS.restore){if(stk.length)ctm=stk.pop()}else if(fn===OPS.transform)ctm=mul(ctm,args);
else if(fn===OPS.rectangle)addRect(args[0],args[1],args[2],args[3]);
else if(fn===OPS.constructPath&&args){const subOps=(args[0]&&typeof args[0].length==='number')?args[0]:[args[0]];const rawC=Array.isArray(args[1])&&args[1].length&&typeof args[1][0]!=='number'?args[1][0]:args[1];
if(rawC&&typeof rawC.length==='number'){let ci=0,cx=0,cy=0,sx=0,sy=0;for(const so of subOps){
if(so===0||so===OPS.moveTo){cx=sx=rawC[ci++];cy=sy=rawC[ci++]}
else if(so===1||so===OPS.lineTo){const nx=rawC[ci++],ny=rawC[ci++];addSeg(cx,cy,nx,ny);cx=nx;cy=ny}
else if(so===2||so===OPS.curveTo){ci+=4;cx=rawC[ci++];cy=rawC[ci++]}
else if(so===3||so===OPS.curveTo2){ci+=2;cx=rawC[ci++];cy=rawC[ci++]}
else if(so===OPS.curveTo3){ci+=2;cx=rawC[ci++];cy=rawC[ci++]}
else if(so===4||so===OPS.closePath){addSeg(cx,cy,sx,sy);cx=sx;cy=sy}
else if(so===OPS.rectangle){addRect(rawC[ci],rawC[ci+1],rawC[ci+2],rawC[ci+3]);ci+=4}}}}}
const fechaHdr=items.find(it=>/^FECHA$/i.test(norm(it.str.trim())));
const fincaHdr=items.find(it=>/^FINCA$/i.test(norm(it.str.trim())));
const nextHdr=fincaHdr&&items.filter(it=>Math.abs(it.vy-fincaHdr.vy)<15&&it.vx>fincaHdr.vx+15).sort((a,b)=>a.vx-b.vx)[0];
const vLeft=fechaHdr&&fincaHdr&&vSegs.filter(v=>v.x>fechaHdr.vx&&v.x<fincaHdr.vx).sort((a,b)=>(b.y2-b.y1)-(a.y2-a.y1))[0];
const vRight=fincaHdr&&nextHdr&&vSegs.filter(v=>v.x>fincaHdr.vx&&v.x<nextHdr.vx).sort((a,b)=>(b.y2-b.y1)-(a.y2-a.y1))[0];
const fxMin=vLeft?vLeft.x-2:(fechaHdr&&fincaHdr?(fechaHdr.vx+fincaHdr.vx)/2:(fincaHdr?fincaHdr.vx-95:160));
const fxMax=vRight?vRight.x+2:(fincaHdr&&nextHdr?(fincaHdr.vx+nextHdr.vx)/2:(fincaHdr?fincaHdr.vx+145:360));
const allYs=hSegs.filter(s=>s.x1<=fxMin-25&&s.x2>=fxMin-5).map(s=>s.y).sort((a,b)=>a-b);
const cutYs=[];let last=-1e9;allYs.forEach(y=>{if(y-last>3){cutYs.push(y);last=y}});
const dayRe=/^(LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO|DOMINGO)\s*,\s*(\d{1,2})\s+DE\s+([A-Z]+)(?:\s+DE\s+(202\d))?/i;
const leftItems=items.filter(it=>it.vx<fxMin);
let boxes=[];
if(cutYs.length>=2){for(let j=0;j<cutYs.length-1;j++){const yTop=cutYs[j],yBottom=cutYs[j+1];
const cellItems=leftItems.filter(it=>it.vy>=yTop-2&&it.vy<=yBottom+2).sort((a,b)=>Math.abs(a.vy-b.vy)<5?a.vx-b.vx:a.vy-b.vy);
const txt=norm(cellItems.map(x=>x.str).join(' ')).replace(/\s*,\s*/g,', ').replace(/\s+/g,' ').trim();
const m=txt.match(dayRe);if(m)boxes.push({d:m[1],dd:+m[2],mm:meses[m[3]],yy:m[4]?+m[4]:docYr,yTop,yBottom,cy:(yTop+yBottom)/2})}}
if(!boxes.length){const days=clusterY(leftItems,16).map(g=>{const arr=g.items.sort((a,b)=>Math.abs(a.vy-b.vy)<5?a.vx-b.vx:a.vy-b.vy);
const txt=norm(arr.map(x=>x.str).join(' ')).replace(/\s*,\s*/g,', ').replace(/\s+/g,' ').trim();return {m:txt.match(dayRe),vy:(g.y0+g.yLast)/2}}).filter(x=>x.m);
boxes=days.map((x,idx)=>({d:x.m[1],dd:+x.m[2],mm:meses[x.m[3]],yy:x.m[4]?+x.m[4]:docYr,yTop:idx>0?(days[idx-1].vy+x.vy)/2:-1e9,yBottom:idx<days.length-1?(x.vy+days[idx+1].vy)/2:1e9,cy:x.vy}))}
let wkNum=(boxes.length&&boxes[0].mm!==undefined)?isoWk(boxes[0].dd,boxes[0].mm,boxes[0].yy):null;
if(!wkNum){const wm=cleanNoVig.match(/SEMANA\s*[:\-]?\s*(\d{1,2})\b/);if(wm)wkNum=+wm[1]}
if(wkNum)text+=`SEM ${wkNum}\n`;
if(boxes.length){const lines=clusterY(items.filter(it=>it.vx>=fxMin&&it.vx<=fxMax),4).map(g=>({vy:(g.y0+g.yLast)/2,t:g.items.sort((a,b)=>a.vx-b.vx).map(x=>x.str).join(' ')})).filter(r=>!/(TOTAL|SISTEMA DE GESTION|RECOMENDACIONES|FINCA|PERIODO)/i.test(norm(r.t)));
lines.forEach(r=>{const hit=boxes.find(b=>r.vy>=b.yTop-2&&r.vy<=b.yBottom+2)||boxes.reduce((a,b)=>Math.abs(b.cy-r.vy)<Math.abs(a.cy-r.vy)?b:a);text+=`${hit.d} ${r.t}\n`})}
else{text+=sortedTxt+'\n'}
if(img){const v=p.getViewport({scale:2.5});const c=document.createElement('canvas');c.width=v.width;c.height=v.height;await p.render({canvasContext:c.getContext('2d'),viewport:v}).promise;blobs.push(await new Promise(r=>c.toBlob(r,'image/png')))}}
return {text,blobs}}
async function shrink(b){try{const bm=await createImageBitmap(b);const k=Math.min(1,2800/Math.max(bm.width,bm.height));const c=document.createElement('canvas');c.width=Math.round(bm.width*k);c.height=Math.round(bm.height*k);const x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(bm,0,0,c.width,c.height);return await new Promise(r=>c.toBlob(r,'image/jpeg',0.92))}catch(e){return b}}
const b64=b=>new Promise(r=>{const fr=new FileReader();fr.onload=()=>r(String(fr.result).split(',')[1]);fr.readAsDataURL(b)});
/* Lectura de imágenes/PDF con Claude a través de la función serverless /api/leer (clave ANTHROPIC_API_KEY en Vercel) */
async function ai(prompt,images){try{const imgs=[];for(const b of images||[]){const s=await shrink(b);imgs.push({type:s.type||'image/jpeg',data:await b64(s)})}
const r=await fetch('/api/leer',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({prompt,images:imgs})});
if(!r.ok)return null;return (await r.json()).data}catch(e){return null}}
const isImg=f=>/\.(jpe?g|png|webp)$/i.test(f.name)||f.type.startsWith('image/');
const isPdf=f=>/\.pdf$/i.test(f.name);
const aspPrompt=t=>`Este archivo es un Programa de Aspersión Semanal de fincas bananeras (columnas SEMANAS, LUNES a DOMINGO; cada celda con nombre de finca; la semana aparece como SEM XX). El archivo puede tener varias páginas y algunas no repiten el encabezado: las columnas siempre van en este orden fijo: SEMANAS, LUNES, MARTES, MIERCOLES, JUEVES, VIERNES, SABADO, DOMINGO. Responde SOLO un JSON: [{"week":44,"day":"LUNES","farm":"RUBI"}] con una entrada por cada finca de cada celda. Catálogo de fincas: ${ALL.join(', ')} (equivalencias: LAS MERCEDES A, B, C o D = LAS MERCEDES, pero GLORIA MERCEDES NO es LAS MERCEDES y se ignora por completo; MERCEDES sola=LAS MERCEDES, LA CEIBA=CEIBA ÚNICAMENTE si la celda dice exactamente "CEIBA" o "LA CEIBA" sola — NUNCA abrevies "EMPREBANCOOP CEIBA", "COOMULBANANO CEIBA" ni ningún otro nombre compuesto con CEIBA como "CEIBA", escríbelos completos tal cual aparecen; JOSEFINA-MARIA TERESA o JOSEFINA MARIA TERESA=JOSEFINA MT, MARTES en una celda de finca=MARTE, PLANTACIÓN=PLANTACION). En el campo farm escribe el nombre exactamente como aparece en la celda y no omitas ninguna celda con nombre. No inventes nada.${t?'\n\nTexto:\n'+t.slice(0,8000):''}`;

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
const unk=[];let fail=0;const fromAI=async()=>{const out=[],tx=isImg(f)||isPdf(f)?'':text;
for(let i=0;i<Math.max(imgs.length,1);i+=2){if(imgs.length>2)upd(setAspF,id,{note:`Leyendo páginas ${i+1}–${Math.min(i+2,imgs.length)} de ${imgs.length}…`});
const j=await ai(aspPrompt(tx),imgs.slice(i,i+2));if(!Array.isArray(j)){fail++;continue}
j.forEach(x=>{const d=D.indexOf(norm(x.day||'')),w=+x.week,fs=findF(x.farm||'',true);if(!fs.length&&x.farm&&!/GLORIA/.test(norm(x.farm)))unk.push(x.farm);if(w>0&&w<54&&d>=0)fs.forEach(fi=>out.push([w,d,fi]))})}
return out};
if(isImg(f)){res=await fromAI();via='lectura con IA'}
else if(isPdf(f)){res=await fromAI();via='lectura con IA';if(!res.length){res=parseAsp(text);via='texto del PDF'}}
else{res=parseAsp(text);via='texto';if(!res.length){res=await fromAI();via='IA'}}
if(!res.length){upd(setAspF,id,{st:'err',note:'No se detectaron semana/día/finca. Si es imagen o PDF escaneado, revisa que ANTHROPIC_API_KEY esté configurada en Vercel.'});continue}
setRestr(r=>{const n={...r};res.forEach(([w,d,fi])=>{const k=w+'|'+d,a=n[k]||[];if(!a.includes(fi))n[k]=[...a,fi]});return n});
upd(setAspF,id,{st:'ok',note:`${res.length} fincas-día extraídas (${via}${imgs.length>1?', '+imgs.length+' págs':''})`+(fail?` · ${fail} lote(s) de páginas no se pudieron leer: vuelve a subir el archivo`:'')+(unk.length?' · Nombres no reconocidos (omitidos): '+[...new Set(unk)].join(', '):'')})}
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
for(let d=0;d<5;d++){[10,10,23].forEach((w,i)=>w1.getColumn(2+d*3+i).width=w);for(let i=0;i<3;i++)put(w1,2,2+d*3+i,null,{b:1,bg:Y});w1.mergeCells(2,2+d*3,2,4+d*3);w1.getCell(2,2+d*3).value=D[d]}
for(let c=1;c<=16;c++)put(w1,1,c,null,{b:1,bg:'FF0054A6',col:'FFFFFFFF'});w1.mergeCells(1,1,1,16);w1.getCell(1,1).value=`PROGRAMACIÓN DE DESPACHOS (${ap||'NOMBRE DEL PROVEEDOR'})`;w1.getRow(1).height=22;put(w1,2,1,null);w1.getRow(2).height=18;
let wks=weeks.filter(w=>[0,1,2,3,4].some(d=>(dm[w.n+'|'+d]||[]).length));if(!wks.length&&weeks.length)wks=[fw||weeks[0]];
let r=3;wks.forEach(w=>{const n=Math.max(6,...[0,1,2,3,4].map(d=>(dm[w.n+'|'+d]||[]).length));
put(w1,r,1,'SEM',{b:1,bg:G});for(let d=0;d<5;d++){put(w1,r,2+d*3,'PLACAS',{b:1,bg:G});put(w1,r,3+d*3,'OC',{b:1,bg:G});put(w1,r,4+d*3,`${D[d]} ${fX(addD(w.start,d))}`,{b:1,bg:G})}
w1.getRow(r).height=32;const top=r+1;
for(let i=0;i<n;i++){put(w1,top+i,1,null,{b:1});for(let d=0;d<5;d++){const e=(dm[w.n+'|'+d]||[])[i]||{};put(w1,top+i,2+d*3,e.placa||null,{l:1});put(w1,top+i,3+d*3,e.oc||null,{b:1});put(w1,top+i,4+d*3,e.finca||null,{b:1,l:1})}}
w1.mergeCells(top,1,top+n-1,1);w1.getCell(top,1).value=w.n;r=top+n});
const w2=wb.addWorksheet('Programa_Aspersion');[12,18,18,18,18,18,18,18].forEach((w,i)=>w2.getColumn(i+1).width=w);
for(let c=1;c<=8;c++)put(w2,1,c,null,{b:1,bg:'FF0054A6',col:'FFFFFFFF'});w2.mergeCells('A1:H1');w2.getCell('A1').value='PROGRAMA DE ASPERSIÓN SEMANAL';w2.getRow(1).height=24;
['SEMANAS',...D].forEach((h,i)=>put(w2,2,i+1,h,{b:1,bg:Y}));
let q=3;weeks.forEach(w=>{const n=Math.max(0,...D.map((_,d)=>(restr[w.n+'|'+d]||[]).length));for(let i=0;i<n;i++){put(w2,q,1,`SEM ${w.n}`,{b:1});D.forEach((_,d)=>put(w2,q,2+d,(restr[w.n+'|'+d]||[])[i]||null));q++}});
w2.autoFilter='A2:H2';
const buf=await wb.xlsx.writeBuffer();const fn=`BANASAN_Programacion_${ap||'despachos'}.xlsx`.replace(/\s+/g,'_');
const u=URL.createObjectURL(new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));const lk=document.createElement('a');lk.href=u;lk.download=fn;document.body.appendChild(lk);lk.click();lk.remove();setTimeout(()=>URL.revokeObjectURL(u),3000);
setToast('Excel generado con 2 pestañas.')}
catch(e){setToast(e&&e.code==='declined'?'Descarga cancelada.':'No se pudo exportar: '+((e&&e.message)||e))}}

const vis=weeks.filter(w=>(!flt||w.n===+flt)&&(!onlyB||D.some((_,d)=>(restr[w.n+'|'+d]||[]).length)));
const Tab=(k,t)=>html`<button onClick=${()=>setTab(k)} class=${'px-4 py-2 rounded-t-lg font-bold text-sm border-b-4 '+(tab===k?'text-white':'bg-white text-slate-600 border-transparent hover:bg-slate-50')} style=${tab===k?{background:BLUE,borderColor:YEL}:{}}>${t}</button>`;
const sel='border border-slate-300 rounded px-2 py-1.5 text-sm bg-white';
const nrows=Math.max(6,...[0,1,2,3,4].map(d=>(dm[(fw?fw.n:0)+'|'+d]||[]).length));

if(cloud==='cargando')return html`<div class="p-10 text-center text-slate-500"><style>${CSS}</style>Cargando datos desde la nube…</div>`;
const stTxt={cargando:'Cargando…',guardando:'Guardando...',guardado:'Guardado en la nube ✓',error:cloud==='error'?'Sin conexión con la nube: los cambios no se guardarán':'No se pudo guardar en la nube'}[st];
return html`<div><style>${CSS}</style>
<div class=${'text-xs text-right px-4 py-1 border-b '+(st==='error'?'bg-red-50 text-red-700':'bg-slate-100 text-slate-500')}>${stTxt}${st==='error'&&html` <button class="underline font-bold" onClick=${()=>cloud==='error'?load(true):setTick(t=>t+1)}>Reintentar</button>`}</div>
<header class="bg-white border-b-4 shadow-sm" style=${{borderColor:BLUE}}><div class="max-w-7xl mx-auto px-4 py-3 flex flex-wrap items-center gap-4">
<img src="/logo.jpg" alt="BANASAN" style=${{height:52}}/>
<div class="flex-1 min-w-[200px]"><div class="font-bold text-lg leading-tight" style=${{color:BLUE}}>Programación de Despachos de Materia Orgánica</div><div class="text-xs text-slate-500">Control de cruces con el cronograma de aspersión</div></div>
<div class="flex gap-2 text-center"><div title="Veces que se intentó registrar un despacho a una finca que estaba en aspersión ese día y el sistema lo rechazó" class="rounded-lg px-3 py-1 cursor-help" style=${{background:YEL}}><div class="text-xl font-extrabold" style=${{color:BLUE}}>${conf}</div><div class="text-[10px] font-bold uppercase">Cruces bloqueados</div></div>
<div class="rounded-lg px-3 py-1 bg-red-50 border border-red-200"><div class="text-xl font-extrabold text-red-600">${nB}</div><div class="text-[10px] font-bold uppercase text-red-700">Aspersiones cargadas</div></div></div>
<button onClick=${exportX} class="flex items-center gap-2 text-white font-bold px-4 py-2 rounded-lg shadow hover:opacity-90" style=${{background:BLUE}}>${Ic('download')} Exportar Excel (.xlsx)</button>
</div>${toast&&html`<div class="max-w-7xl mx-auto px-4 pb-2 text-sm text-blue-800 flex gap-2 items-center">${toast}<button onClick=${()=>setToast('')}>${Ic('x',14)}</button></div>`}</header>

<main class="max-w-7xl mx-auto px-4 py-5">
<div class="flex gap-1 mb-0 flex-wrap">${Tab('carga','1 · Carga de archivos')}${Tab('asp','Hoja 1 · Programa de Aspersión')}${Tab('desp','Hoja 2 · Programación de Despachos')}</div>
<div class="bg-slate-50 border border-slate-200 rounded-b-xl rounded-tr-xl p-4">

${tab==='carga'&&html`<div class="grid md:grid-cols-2 gap-4">
<${Drop} title="Caja 1 · Calendario Corporativo Bananero" sub="Sube el calendario (foto, PDF o texto). Se habilitan SEM 1 a SEM 52 con su rango lunes–domingo." onFiles=${doCal} files=${calF}>
<div class="flex items-center gap-2 mt-3 text-sm"><span class="text-slate-500">¿Sin archivo? Generar año:</span><input type="number" value=${yr} onChange=${e=>setYr(+e.target.value)} class=${sel+' w-24'}/>
<button onClick=${()=>afterW(mkWeeks(isoW1(yr)))} class="px-3 py-1.5 rounded font-bold text-white text-sm" style=${{background:BLUE}}>Generar</button></div>
${weeks.length>0&&html`<div class="mt-3 text-sm rounded bg-emerald-50 border border-emerald-200 text-emerald-800 px-3 py-2">${weeks.length} semanas habilitadas · ${wl(weeks[0])} → ${wl(weeks[weeks.length-1])}</div>`}
<//>
<${Drop} title="Caja 2 · Programa de Aspersión Semanal" sub="Sube uno o varios archivos; se suman a los anteriores. Solo se cargan las fincas y días que contienen." onFiles=${doAsp} files=${aspF}>
<div class="mt-3 text-sm flex items-center gap-3"><span class="text-slate-600">Bloqueos actuales: <b class="text-red-600">${nB}</b></span>
${nB>0&&html`<button onClick=${()=>setRestr({})} class="flex items-center gap-1 text-red-600 border border-red-300 rounded px-2 py-1 hover:bg-red-50">${Ic('trash',14)} Vaciar todo</button>`}</div>
<//></div>`}

${tab==='asp'&&html`<div>
<div class="flex flex-wrap items-center gap-3 mb-3 text-sm"><label class="font-bold">Semana:</label>
<select class=${sel} value=${flt} onChange=${e=>setFlt(e.target.value)}><option value="">Todas (SEM 1 – 52)</option>${weeks.map(w=>html`<option key=${w.n} value=${w.n}>${wl(w)}</option>`)}</select>
<label class="flex items-center gap-1"><input type="checkbox" checked=${onlyB} onChange=${e=>setOnlyB(e.target.checked)}/> Solo semanas con bloqueos</label>
<span class="text-slate-500">Haz clic en una celda (lista “+ agregar”) para añadir; usa la × para quitar.</span></div>
${!weeks.length?html`<div class="p-6 text-center text-slate-500">Primero carga el calendario (pestaña 1) para habilitar las semanas.</div>`:html`
<div class="overflow-x-auto bg-white border rounded"><table class="w-full border-collapse"><thead>
<tr><th colSpan="8" class="text-white py-2 text-base" style=${{background:BLUE}}>PROGRAMA DE ASPERSIÓN SEMANAL</th></tr>
<tr><th class="border border-slate-400 p-2" style=${{background:YEL,color:BLUE}}>SEMANAS</th>${DL.map(x=>html`<th key=${x} class="border border-slate-400 p-2" style=${{background:YEL,color:BLUE}}>${x}</th>`)}</tr></thead>
<tbody>${vis.map(w=>html`<tr key=${w.n}><td class="border border-slate-300 p-2 font-bold whitespace-nowrap align-top bg-yellow-50">SEM ${w.n}<div class="text-[11px] font-normal text-slate-500">${fR(w.start)} - ${fR(addD(w.start,6))}</div></td>
${D.map((_,d)=>{const k=w.n+'|'+d,a=restr[k]||[];return html`<td key=${d} class="border border-slate-300 align-top p-1 min-w-[125px]">
${a.map(f=>html`<span key=${f} class="inline-flex items-center gap-1 m-0.5 px-2 py-0.5 rounded text-xs font-bold text-white bg-red-600">${f}<button onClick=${()=>rm(w.n,d,f)}>${Ic('x',12)}</button></span>`)}
<select class="text-xs w-full mt-1 border rounded text-slate-400 bg-white" value="" onChange=${e=>add(w.n,d,e.target.value)}><option value="">+ agregar</option>${ALL.filter(f=>!a.includes(f)).map(f=>html`<option key=${f} value=${f}>${f}</option>`)}</select></td>`})}</tr>`)}</tbody></table></div>
${nB===0&&html`<div class="mt-3 text-sm text-slate-500">La tabla inicia con 0 bloqueos: se llena solo con tus archivos o con registros manuales.</div>`}`}
</div>`}

${tab==='desp'&&html`<div>
<div class="flex flex-wrap items-center gap-2 mb-3"><span class="font-bold text-sm">Proveedor:</span>
${provs.map(p=>html`<button key=${p} onClick=${()=>{setAp(p);setEdP(null);setDelP(false)}} class=${'px-3 py-1.5 rounded-full text-sm font-bold border '+(ap===p?'text-white':'bg-white text-slate-600')} style=${ap===p?{background:BLUE,borderColor:BLUE}:{}}>${p}</button>`)}
${ap&&edP===null&&!delP&&html`<span class="flex gap-1"><button title="Editar nombre" onClick=${()=>setEdP(ap)} class="p-1.5 rounded border bg-white text-slate-600 hover:bg-slate-100">${Ic('pencil',15)}</button><button title="Eliminar proveedor" onClick=${()=>setDelP(true)} class="p-1.5 rounded border bg-white text-red-600 hover:bg-red-50">${Ic('trash',15)}</button></span>`}
${edP!==null&&html`<span class="flex gap-1"><input autoFocus value=${edP} onChange=${e=>setEdP(e.target.value)} onKeyDown=${e.key==='Enter'&&renProv()} class=${sel}/><button onClick=${renProv} class="px-3 rounded text-white text-sm font-bold" style=${{background:BLUE}}>Guardar nombre</button><button onClick=${()=>setEdP(null)} class="px-2 text-slate-500">${Ic('x',16)}</button></span>`}
${delP&&html`<span class="flex items-center gap-2 text-sm bg-red-50 border border-red-200 rounded px-2 py-1">¿Eliminar <b>${ap}</b> y todos sus despachos?<button onClick=${delProv} class="px-2 py-0.5 rounded bg-red-600 text-white font-bold">Sí, eliminar</button><button onClick=${()=>setDelP(false)} class="px-2 text-slate-600">Cancelar</button></span>`}
${newP===null?html`<button onClick=${()=>setNewP('')} class="flex items-center gap-1 px-3 py-1.5 rounded-full text-sm font-bold border-2 border-dashed" style=${{color:BLUE,borderColor:BLUE}}>${Ic('plus',14)} Nuevo Proveedor</button>`
:html`<span class="flex gap-1"><input autofocus placeholder="Nombre del proveedor" value=${newP} onChange=${e=>setNewP(e.target.value)} onKeyDown=${e.key==='Enter'&&addProv()} class=${sel}/><button onClick=${addProv} class="px-3 rounded text-white text-sm font-bold" style=${{background:BLUE}}>Guardar</button><button onClick=${()=>setNewP(null)} class="px-2 text-slate-500">${Ic('x',16)}</button></span>`}</div>
${!weeks.length?html`<div class="p-6 text-center text-slate-500">Primero carga el calendario (pestaña 1).</div>`:html`
<div class="bg-white border rounded-lg p-3 mb-3"><div class="grid grid-cols-2 md:grid-cols-6 gap-2 items-end text-sm">
<label class="md:col-span-2 font-bold">Semana<select class=${sel+' w-full font-normal'} value=${form.w} onChange=${e=>setF('w',e.target.value)}>${weeks.map(w=>html`<option key=${w.n} value=${w.n}>${wl(w)}</option>`)}</select></label>
<label class="font-bold">Día<select class=${sel+' w-full font-normal'} value=${form.d} onChange=${e=>setF('d',e.target.value)}>${DL.slice(0,5).map((x,i)=>html`<option key=${i} value=${i}>${x}${fw?' '+fR(addD(fw.start,i)):''}</option>`)}</select></label>
<label class="font-bold">Placas <span class="font-normal text-slate-400">(opcional)</span><input class=${sel+' w-full font-normal'} placeholder="WNQ-072" value=${form.placa} onChange=${e=>setF('placa',e.target.value)}/></label>
<label class="font-bold">OC<input class=${sel+' w-full font-normal'} placeholder="40449" value=${form.oc} onChange=${e=>setF('oc',e.target.value)}/></label>
<label class="font-bold">Finca destino<select class=${sel+' w-full font-normal '+(live?'border-red-600 bg-red-50 text-red-700':'')} value=${form.finca} onChange=${e=>setF('finca',e.target.value)}><option value="">Seleccionar…</option>
<optgroup label="Orgánicas">${ORG.map(f=>html`<option key=${f}>${f}</option>`)}</optgroup><optgroup label="Convencionales">${CONV.map(f=>html`<option key=${f}>${f}</option>`)}</optgroup></select></label></div>
<div class="mt-3 flex flex-wrap items-center gap-3"><button onClick=${reg} class="flex items-center gap-1 px-4 py-2 rounded-lg text-white font-bold" style=${{background:BLUE}}>${Ic('plus',16)} Registrar despacho</button>
${live&&!err&&html`<span class="text-red-600 text-sm font-bold">${Ic('warn',16,'inline')} Esta finca está en aspersión ese día: el guardado será bloqueado.</span>`}
${ok&&html`<span class="text-emerald-700 text-sm font-bold">${ok}</span>`}</div>
${err&&html`<div class="mt-3 rounded-lg bg-red-600 text-white font-bold px-4 py-3">${err}</div>`}</div>
<div class="overflow-x-auto bg-white border rounded"><table class="border-collapse w-full" style=${{minWidth:1000}}><thead>
<tr><th colSpan="16" class="text-white py-2 text-base" style=${{background:BLUE}}>PROGRAMACIÓN DE DESPACHOS (${ap||'NOMBRE DEL PROVEEDOR'})</th></tr><tr><th class="border border-slate-400"></th>${DL.slice(0,5).map(x=>html`<th key=${x} colSpan="3" class="border border-slate-400 py-1" style=${{background:'#FFFF00'}}>${x.replace('É','E')}</th>`)}</tr>
<tr><th class="border border-slate-400 bg-slate-300 px-2">SEM</th>${[0,1,2,3,4].map(d=>html`<${React.Fragment} key=${d}><th class="border border-slate-400 bg-slate-300 px-1">PLACAS</th><th class="border border-slate-400 bg-slate-300 px-1">OC</th><th class="border border-slate-400 bg-slate-300 px-1">${fw?D[d]+' '+fX(addD(fw.start,d)):D[d]}</th><//>`)}</tr></thead>
<tbody>${Array.from({length:nrows},(_,i)=>html`<tr key=${i}>${i===0&&html`<td rowSpan=${nrows} class="border border-slate-400 text-center font-bold text-lg w-12">${fw?fw.n:''}</td>`}
${[0,1,2,3,4].map(d=>{const k=(fw?fw.n:0)+'|'+d,e=(dm[k]||[])[i];const bad=e&&(restr[k]||[]).includes(e.finca);
return html`<${React.Fragment} key=${d}><td class=${'border border-slate-300 h-7 '+(e&&!e.placa?'bg-amber-50':'')}>${e&&html`<input value=${e.placa} placeholder="agregar placa" onChange=${ev=>setE(k,i,{placa:ev.target.value.toUpperCase()})} class="w-full px-1 bg-transparent outline-none focus:bg-yellow-50 placeholder:text-amber-500 placeholder:text-xs" style=${{minWidth:80}}/>`}</td><td class="border border-slate-300 font-bold text-center">${e&&html`<input value=${e.oc} onChange=${ev=>setE(k,i,{oc:ev.target.value})} class="w-full px-1 bg-transparent outline-none text-center font-bold focus:bg-yellow-50" style=${{minWidth:60}}/>`}</td>
<td class=${'border border-slate-300 px-1 font-bold '+(e?(bad?'bg-red-600 text-white':bdg(e.finca)):'')}>${e&&html`<span class="flex justify-between gap-1">${e.finca}<button onClick=${()=>delE(k,i)} title="Eliminar">${Ic('x',12)}</button></span>`}</td><//>`})}</tr>`)}</tbody></table></div>
<div class="mt-2 text-xs text-slate-500 flex gap-3 flex-wrap"><span class="px-2 rounded bg-emerald-100 text-emerald-800">Orgánica</span><span class="px-2 rounded bg-sky-100 text-sky-800">Convencional</span><span class="px-2 rounded bg-amber-100 text-amber-800">Placa pendiente (clic en la celda para escribirla)</span><span class="px-2 rounded bg-red-600 text-white">Cruce con aspersión (agregado después)</span></div>`}
</div>`}
</div></main></div>`}
