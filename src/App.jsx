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
const fR=d=>`${P2(d.getUTCDate())} ${MO[d.getUTCMonth()]}`; const fD=d=>`${P2(d.getUTCDate())}/${P2(d.getUTCMonth()+1)}/${d.getUTCFullYear()}`;
const fX=d=>`${d.getUTCDate()}/${d.getUTCMonth()+1}/${d.getUTCFullYear()}`; const isoW1=y=>{const j=new Date(Date.UTC(y,0,4));return addD(j,-((j.getUTCDay()+6)\%7))}; const isoWk=(d,m,y)=>{const dt=new Date(Date.UTC(y,m,d));const day=(dt.getUTCDay()+6)\%7;dt.setUTCDate(dt.getUTCDate()-day+3);const w1=new Date(Date.UTC(dt.getUTCFullYear(),0,4));return 1+Math.round(((dt-w1)/864e5-3+(w1.getUTCDay()+6)\%7)/7)}; const mkWeeks=s=>Array.from({length:52},(_,i)=>({n:i+1,start:addD(s,i*7)})); const wl=w=>`SEM ${w.n}: ${fR(w.start)} -${fR(addD(w.start,6))}`;
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
2) Matriz semanal (columnas SEMANAS, LUNES, MARTES, MIERCOLES, JUEVES, VIERNES, SAB
