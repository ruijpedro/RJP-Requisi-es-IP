/* RJP V4: exportação local; XLSX usa o modelo XLSM original sem executar macros. */
(function(){
 const name=r=>'Requisicao_'+String(r.reqNo||'sem_numero').replace(/[^\w-]/g,'_');
 async function excel(){
  if(typeof XLSX==='undefined'){alert('Biblioteca Excel indisponível. Ligue à Internet na primeira utilização.');return}
  const r=collect();
  try{
   const res=await fetch('./modelo-original.xlsm');if(!res.ok)throw Error('Modelo não encontrado');
   const wb=XLSX.read(await res.arrayBuffer(),{type:'array',cellFormula:true});
   const ws=wb.Sheets[wb.SheetNames[0]];
   function put(cell,val){ws[cell]={t:typeof val==='number'?'n':'s',v:val}}
   // Preserva o desenho original. Campos colocados nas células adjacentes às legendas.
   put('D7',r.action||'');put('D9',r.reqNo||'');put('F9',r.costCenter||'');
   put('D10',r.sigma||'');put('F10',r.date||'');put('D11',r.arco||'');put('F11',r.location||'');put('C15',r.description||'');
   // Limpa os itens de exemplo e as fórmulas de consulta externa.
   const end=Math.max(100,22+r.lines.length);
   for(let i=22;i<=end;i++){for(const c of ['B','C','D','E','F','G'])delete ws[c+i]}
   r.lines.forEach((x,i)=>{const n=22+i;put('B'+n,x.code);put('C'+n,x.desc);put('D'+n,x.unit);put('E'+n,x.qty);put('F'+n,x.price);ws['G'+n]={t:'n',f:`E${n}*F${n}`,v:x.qty*x.price}});
   ws.G12={t:'n',f:`SUM(G22:G${Math.max(22,21+r.lines.length)})`,v:r.total};
   const range=XLSX.utils.decode_range(ws['!ref']||'A1:I100');range.e.r=Math.max(range.e.r,21+r.lines.length);ws['!ref']=XLSX.utils.encode_range(range);
   // Remove folhas de catálogo histórico para não perpetuar preços antigos no documento exportado.
   wb.SheetNames=[wb.SheetNames[0]];wb.Sheets={[wb.SheetNames[0]]:ws};
   XLSX.writeFile(wb,name(r)+'.xlsx',{bookType:'xlsx',compression:true});
  }catch(e){console.error(e);alert('Não foi possível exportar o Excel: '+e.message)}
 }
 function pdf(){const r=collect();if(!r.lines.length){alert('Adicione artigos antes de exportar.');return}const w=window.open('','_blank');if(!w){alert('Autorize janelas pop-up para exportar PDF.');return}
 const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 w.document.write(`<!doctype html><html lang="pt"><head><meta charset="utf-8"><title>${name(r)}</title><style>body{font:12px Arial;margin:22mm;color:#182e38}h1{font-size:20px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #aaa;padding:6px;text-align:left}th{background:#eee}.sum{text-align:right;font-size:17px;margin-top:18px}@page{size:A4;margin:15mm}</style></head><body><h1>REQUISIÇÃO DE TRABALHOS</h1><p><b>N.º:</b> ${escape(r.reqNo)} &nbsp; <b>Data:</b> ${escape(r.date)} &nbsp; <b>Tipo de ação:</b> ${escape(r.action)}</p><p><b>Centro de custo:</b> ${escape(r.costCenter)} &nbsp; <b>Ação SIGMA:</b> ${escape(r.sigma)} &nbsp; <b>Ref. ARCO:</b> ${escape(r.arco)}</p><p><b>Local:</b> ${escape(r.location)}</p><p><b>Descrição:</b> ${escape(r.description)}</p><table><thead><tr><th>Cód. LPU</th><th>Designação</th><th>Un.</th><th>Qtd.</th><th>Preço un.</th><th>Valor</th></tr></thead><tbody>${r.lines.map(x=>`<tr><td>${escape(x.code)}</td><td>${escape(x.desc)}</td><td>${escape(x.unit)}</td><td>${escape(x.qty)}</td><td>${euro(x.price)}</td><td>${euro(x.value)}</td></tr>`).join('')}</tbody></table><div class="sum"><b>Custo estimado: ${euro(r.total)}</b></div></body></html>`);w.document.close();setTimeout(()=>w.print(),350);
 }
 function add(){const bar=document.querySelector('#editor .bar');if(!bar||document.getElementById('v4excel'))return;const e=document.createElement('button');e.id='v4excel';e.className='btn primary no-print';e.textContent='⬇ Excel original';e.onclick=excel;const p=document.createElement('button');p.className='btn secondary no-print';p.textContent='⬇ PDF';p.onclick=pdf;bar.append(e,p)}
 add();
})();
