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
 function pdfDocument(){
  const r=collect();if(!r.lines.length){alert('Adicione artigos antes de gerar o PDF.');return null}
  const fmt=n=>Number(n||0).toLocaleString('pt-PT',{minimumFractionDigits:2,maximumFractionDigits:2});
  const pages=[];let page=[],y=0;
  function next(){if(page.length)pages.push(page);page=[];y=0}
  function row(left,right='',bold=false){if(y>=48)next();page.push({left:String(left),right:String(right),bold});y++}
  function wrap(t,max=83){let words=String(t??'').replace(/\s+/g,' ').split(' '),out=[],line='';for(const word of words){if((line+' '+word).trim().length>max){if(line)out.push(line);line=word;while(line.length>max){out.push(line.slice(0,max));line=line.slice(max)}}else line+=(line?' ':'')+word}if(line)out.push(line);return out.length?out:['']}
  row('INFRAESTRUTURAS DE PORTUGAL', '',true);row('REQUISIÇÃO DE TRABALHOS', '',true);row('');
  row('N.º: '+(r.reqNo||'RASCUNHO')+'     Data: '+datePT(r.date));
  row('Centro de Custo: '+(r.costCenter||''));
  row('Tipo de ação: '+(r.action||''));row('Ação SIGMA: '+(r.sigma||'')+'     Referência ARCO: '+(r.arco||''));
  for(const t of wrap('Local de intervenção: '+(r.location||'')))row(t);
  for(const t of wrap('Descrição dos trabalhos: '+(r.description||'')))row(t);
  row('');row('ARTIGOS / DESIGNAÇÃO', 'VALOR (€)',true);row('--------------------------------------------------------------');
  r.lines.forEach((x,i)=>{for(const t of wrap((i+1)+'. '+x.code+' — '+x.desc,78))row(t);row('    '+fmt(x.qty)+' '+x.unit+'  ×  '+fmt(x.price)+' €',fmt(x.value)+' €');row('')});
  row('TOTAL ESTIMADO',fmt(r.total)+' €',true);next();
  // PDF 1.4 com fonte Helvetica e codificação WinAnsi, compatível com acentos portugueses.
  const encode=s=>{let a=[];for(const ch of s){const n=ch.codePointAt(0);const cp={0x20ac:128,0x201a:130,0x0192:131,0x201e:132,0x2026:133,0x2020:134,0x2021:135,0x02c6:136,0x2030:137,0x0160:138,0x2039:139,0x0152:140,0x017d:142,0x2018:145,0x2019:146,0x201c:147,0x201d:148,0x2022:149,0x2013:150,0x2014:151,0x02dc:152,0x2122:153,0x0161:154,0x203a:155,0x0153:156,0x017e:158,0x0178:159};a.push(n<=255?n:cp[n]||63)}return a};
  const escape=s=>{let o='';for(const n of encode(s)){if(n===40||n===41||n===92)o+='\\'+String.fromCharCode(n);else if(n<32)o+=' ';else o+=String.fromCharCode(n)}return o};
  const objs=[null],add=v=>(objs.push(v),objs.length-1),font=add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>'),bold=add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>');
  const pageRefs=[],pagesId=add('');
  pages.forEach((rows,index)=>{let content='BT /F2 9 Tf 42 808 Td (RJP REQUISIÇÕES IP) Tj ET\n';rows.forEach((v,i)=>{let yy=779-i*15;content+=`BT /F${v.bold?2:1} 9 Tf 42 ${yy} Td (${escape(v.left)}) Tj ET\n`;if(v.right)content+=`BT /F${v.bold?2:1} 9 Tf 474 ${yy} Td (${escape(v.right)}) Tj ET\n`});content+=`BT /F1 8 Tf 42 38 Td (${escape('Centro de Custo: '+r.costCenter)}) Tj ET\nBT /F1 8 Tf 485 38 Td (${index+1}/${pages.length}) Tj ET\n`;const bytes=encode(content);const stream=add('<< /Length '+bytes.length+' >>\nstream\n'+content+'endstream');const id=add(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${font} 0 R /F2 ${bold} 0 R >> >> /Contents ${stream} 0 R >>`);pageRefs.push(id)});
  objs[pagesId]=`<< /Type /Pages /Kids [${pageRefs.map(i=>i+' 0 R').join(' ')}] /Count ${pageRefs.length} >>`;
  const catalog=add(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);let data='%PDF-1.4\n%âãÏÓ\n',offsets=[0];for(let i=1;i<objs.length;i++){offsets[i]=encode(data).length;data+=`${i} 0 obj\n${objs[i]}\nendobj\n`}const xref=encode(data).length;data+=`xref\n0 ${objs.length}\n0000000000 65535 f \n`;for(let i=1;i<objs.length;i++)data+=String(offsets[i]).padStart(10,'0')+' 00000 n \n';data+=`trailer\n<< /Size ${objs.length} /Root ${catalog} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Blob([new Uint8Array(encode(data))],{type:'application/pdf'});
 }
 function downloadPdf(blob,filename){const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=filename;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),60000)}
 async function pdf(){const blob=pdfDocument();if(!blob)return;const r=collect(),file=new File([blob],name(r)+'.pdf',{type:'application/pdf'});if(navigator.canShare?.({files:[file]})){try{await navigator.share({files:[file],title:'Requisição de trabalhos'});return}catch(e){if(e.name==='AbortError')return}}downloadPdf(blob,file.name)}
 function printPdf(){const blob=pdfDocument();if(!blob)return;const url=URL.createObjectURL(blob);window.open(url,'_blank');toast('PDF aberto: utilize Partilhar / Imprimir no visualizador.');setTimeout(()=>URL.revokeObjectURL(url),120000)}
 window.rjpExportPDF=pdf;window.rjpPrintPDF=printPdf;
 function add(){const bar=document.querySelector('#editor .bar');if(!bar||document.getElementById('v4excel'))return;const e=document.createElement('button');e.id='v4excel';e.className='btn primary no-print';e.textContent='⬇ Excel original';e.onclick=excel;const p=document.createElement('button');p.className='btn secondary no-print';p.textContent='⬇ PDF';p.onclick=pdf;bar.append(e,p)}
 add();
})();
