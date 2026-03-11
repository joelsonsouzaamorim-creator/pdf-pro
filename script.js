// ===================== ESTADO GLOBAL DE RESULTADOS =====================
// Guarda bytes prontos em memória para permitir baixar ou enviar para outra ferramenta

let resultadoJuntarBytes=null
let resultadoJuntarNome=""
let resultadosDividirList=[]   // [{bytes, nome}]
let resultadosCompactarList=[] // [{bytes, nome}]


// ===================== NAVEGAÇÃO =====================

function abrirUnificador(filesParaCarregar){
  ocultarTudo()
  document.getElementById("unificador").style.display="block"
  if(filesParaCarregar && filesParaCarregar.length>0){
    carregarFilesNoJuntar(filesParaCarregar)
  }
}

function abrirDividir(fileParaCarregar){
  ocultarTudo()
  document.getElementById("dividir-pdf").style.display="block"
  if(fileParaCarregar){
    carregarFileNoDividir(fileParaCarregar)
  }
}

function abrirOCR(fileParaCarregar){
  ocultarTudo()
  document.getElementById("ocr-section").style.display="block"
  if(fileParaCarregar){
    carregarFileNoOCR(fileParaCarregar)
  }
}

function abrirCompactar(filesParaCarregar){
  ocultarTudo()
  document.getElementById("compactar-section").style.display="block"
  if(filesParaCarregar && filesParaCarregar.length>0){
    carregarFilesNoCompactar(filesParaCarregar)
  }
}

function abrirConverter(fileParaCarregar){
  ocultarTudo()
  document.getElementById("converter-section").style.display="block"
  if(fileParaCarregar){
    carregarFileNoConverter(fileParaCarregar)
  }
}

function abrirJuntarPastas(){
  ocultarTudo()
  document.getElementById("juntarpastas-section").style.display="block"
}

function abrirWordParaPDF(fileParaCarregar){
  ocultarTudo()
  document.getElementById("wordpdf-section").style.display="block"
  if(fileParaCarregar) carregarFileNoWordPDF(fileParaCarregar)
}

function abrirExcelParaPDF(fileParaCarregar){
  ocultarTudo()
  document.getElementById("excelpdf-section").style.display="block"
  if(fileParaCarregar) carregarFileNoExcelPDF(fileParaCarregar)
}

function abrirExcelParaWord(fileParaCarregar){
  ocultarTudo()
  document.getElementById("excelword-section").style.display="block"
  if(fileParaCarregar) carregarFileNoExcelWord(fileParaCarregar)
}

function ocultarTudo(){
  document.getElementById("menu-principal").style.display="none"
  document.getElementById("unificador").style.display="none"
  document.getElementById("dividir-pdf").style.display="none"
  document.getElementById("ocr-section").style.display="none"
  document.getElementById("compactar-section").style.display="none"
  document.getElementById("converter-section").style.display="none"
  document.getElementById("wordpdf-section").style.display="none"
  document.getElementById("excelpdf-section").style.display="none"
  document.getElementById("excelword-section").style.display="none"
  document.getElementById("juntarpastas-section").style.display="none"
}

function voltarMenu(){
  document.getElementById("menu-principal").style.display="block"
  document.getElementById("unificador").style.display="none"
  document.getElementById("dividir-pdf").style.display="none"
  document.getElementById("ocr-section").style.display="none"
  document.getElementById("compactar-section").style.display="none"
  document.getElementById("converter-section").style.display="none"
  document.getElementById("wordpdf-section").style.display="none"
  document.getElementById("excelpdf-section").style.display="none"
  document.getElementById("excelword-section").style.display="none"
  document.getElementById("juntarpastas-section").style.display="none"
}


// ===================== UTILIDADES =====================

function bytesParaFile(bytes, nome){
  const blob=new Blob([bytes],{type:"application/pdf"})
  return new File([blob], nome, {type:"application/pdf"})
}

function baixarArquivo(bytes, nome){
  const blob=new Blob([bytes],{type:"application/pdf"})
  const link=document.createElement("a")
  link.href=URL.createObjectURL(blob)
  link.download=nome
  link.click()
}

async function renderizarCapa(file){
  const arrayBuffer=await file.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  const page=await pdf.getPage(1)
  const viewport=page.getViewport({scale:0.4})
  const canvas=document.createElement("canvas")
  canvas.width=viewport.width
  canvas.height=viewport.height
  await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise
  return {canvas, totalPaginas:pdf.numPages}
}

async function renderizarCapaNoCanvas(file, canvasEl){
  const arrayBuffer=await file.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  const page=await pdf.getPage(1)
  const viewport=page.getViewport({scale:0.5})
  canvasEl.width=viewport.width
  canvasEl.height=viewport.height
  await page.render({canvasContext:canvasEl.getContext("2d"),viewport}).promise
  return pdf.numPages
}

async function abrirPaginasModal(file){
  document.getElementById("modal-titulo").textContent=file.name
  const grid=document.getElementById("modal-paginas-grid")
  grid.innerHTML="<p style='color:#6b7280'>Carregando páginas...</p>"
  document.getElementById("modal-paginas").style.display="flex"

  const arrayBuffer=await file.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  grid.innerHTML=""

  for(let i=1;i<=pdf.numPages;i++){
    const page=await pdf.getPage(i)
    const viewport=page.getViewport({scale:0.5})
    const canvas=document.createElement("canvas")
    canvas.width=viewport.width
    canvas.height=viewport.height
    await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise

    const item=document.createElement("div")
    item.className="modal-page-item"
    const label=document.createElement("span")
    label.textContent="Página "+i
    item.appendChild(canvas)
    item.appendChild(label)
    grid.appendChild(item)
  }
}

function fecharModal(e){
  if(e.target===document.getElementById("modal-paginas")){
    document.getElementById("modal-paginas").style.display="none"
  }
}

function fecharModalBtn(){
  document.getElementById("modal-paginas").style.display="none"
}

async function criarThumbCard(file){
  const {canvas, totalPaginas}=await renderizarCapa(file)
  const card=document.createElement("div")
  card.className="thumb-card"
  card.onclick=()=>abrirPaginasModal(file)

  const nome=document.createElement("div")
  nome.className="thumb-nome"
  nome.textContent=file.name

  const pags=document.createElement("div")
  pags.className="thumb-paginas"
  pags.textContent=totalPaginas+" página"+(totalPaginas>1?"s":"")

  card.appendChild(canvas)
  card.appendChild(nome)
  card.appendChild(pags)
  return {card, totalPaginas}
}


// ===================== JUNTAR PDF =====================

let arquivos=[]
let pastasJuntar=[]
let resultadosPastasJuntar=[]

const fileInput=document.getElementById("fileInput")
const gallery=document.getElementById("file-gallery")

fileInput.addEventListener("change",async(e)=>{
  await carregarFilesNoJuntar(Array.from(e.target.files))
  fileInput.value=""
})

function selecionarPasta(){
  const input=document.createElement("input")
  input.type="file"
  input.webkitdirectory=true
  input.style.display="none"
  document.body.appendChild(input)
  input.addEventListener("change",async(e)=>{
    const files=Array.from(e.target.files).filter(f=>f.type==="application/pdf"||f.type.startsWith("image/"))
    document.body.removeChild(input)
    if(files.length===0){alert("Nenhum PDF ou imagem encontrado na pasta.");return}
    const nomePasta=files[0].webkitRelativePath.split("/")[0]||("Pasta "+(pastasJuntar.length+1))
    pastasJuntar.push({nomePasta,arquivos:files})
    renderizarPastaThumb(pastasJuntar.length-1)
    atualizarInfoJuntar()
  })
  input.click()
}

function renderizarPastaThumb(idx){
  const {nomePasta, arquivos:arqs}=pastasJuntar[idx]
  const row=document.getElementById("pastas-adicionadas")

  const card=document.createElement("div")
  card.className="pasta-thumb-card"
  card.id="pasta-thumb-"+idx

  const btnRemover=document.createElement("button")
  btnRemover.className="pasta-thumb-remover"
  btnRemover.textContent="✕"
  btnRemover.onclick=(e)=>{e.stopPropagation();removerPastaJuntar(idx)}

  const icon=document.createElement("div")
  icon.className="pasta-thumb-icon"
  icon.textContent="📁"

  const nome=document.createElement("div")
  nome.className="pasta-thumb-nome"
  nome.textContent=nomePasta

  const qtd=document.createElement("div")
  qtd.className="pasta-thumb-qtd"
  qtd.textContent=arqs.length+" doc"+(arqs.length>1?"s":"")

  const thumbsArea=document.createElement("div")
  thumbsArea.className="pasta-thumbs-expandida"
  thumbsArea.style.display="none"
  thumbsArea.id="pasta-thumbs-exp-"+idx

  card.appendChild(btnRemover)
  card.appendChild(icon)
  card.appendChild(nome)
  card.appendChild(qtd)
  card.appendChild(thumbsArea)

  card.onclick=async()=>{
    const area=document.getElementById("pasta-thumbs-exp-"+idx)
    if(area.style.display==="none"){
      if(area.children.length===0){
        area.style.display="flex"
        area.innerHTML="<div style='color:#6b7280;font-size:11px;padding:8px'>Carregando...</div>"
        area.innerHTML=""
        for(let file of arqs.slice(0,12)){
          try{
            const {card:thumb}=await criarThumbCard(file)
            thumb.style.width="80px"
            area.appendChild(thumb)
          }catch(e){}
        }
        if(arqs.length>12){
          const mais=document.createElement("div")
          mais.className="thumb-mais"
          mais.style.cssText="width:80px;height:100px;font-size:12px"
          mais.textContent="+"+(arqs.length-12)
          area.appendChild(mais)
        }
      }else{
        area.style.display="flex"
      }
      card.classList.add("expandida")
    }else{
      area.style.display="none"
      card.classList.remove("expandida")
    }
  }

  row.appendChild(card)
}

function removerPastaJuntar(idx){
  pastasJuntar.splice(idx,1)
  document.getElementById("pastas-adicionadas").innerHTML=""
  for(let i=0;i<pastasJuntar.length;i++) renderizarPastaThumb(i)
  atualizarInfoJuntar()
}

// Converte imagem para PDF de uma página
async function imagemParaPDF(file){
  const novoPdf=await PDFLib.PDFDocument.create()
  let img
  if(file.type==="image/png"){
    img=await novoPdf.embedPng(await file.arrayBuffer())
  }else{
    img=await novoPdf.embedJpg(await file.arrayBuffer())
  }
  const page=novoPdf.addPage([img.width,img.height])
  page.drawImage(img,{x:0,y:0,width:img.width,height:img.height})
  const pdfBytes=await novoPdf.save()
  return new File([pdfBytes],file.name.replace(/\.[^.]+$/,"")+".pdf",{type:"application/pdf"})
}

async function carregarFilesNoJuntar(files){
  for(let file of files){
    let fileFinal=file
    if(file.type!=="application/pdf"){
      document.getElementById("juntar-info").textContent="Convertendo imagem "+file.name+"..."
      try{ fileFinal=await imagemParaPDF(file) }
      catch(e){ alert("Não foi possível converter: "+file.name); continue }
    }
    arquivos.push(fileFinal)
    const {card}=await criarThumbCard(fileFinal)
    gallery.appendChild(card)
  }
  atualizarInfoJuntar()
}

function atualizarInfoJuntar(){
  const ta=arquivos.length
  const tp=pastasJuntar.length
  let info=""
  if(tp>0) info+=tp+" pasta"+(tp>1?"s":"")
  if(ta>0) info+=(info?" + ":"")+ta+" arquivo"+(ta>1?"s":"")
  document.getElementById("file-counter").textContent=info||"Nenhum arquivo carregado"
  document.getElementById("juntar-info").textContent=info
}

function limparJuntar(){
  arquivos=[]
  pastasJuntar=[]
  resultadoJuntarBytes=null
  resultadoJuntarNome=""
  resultadosPastasJuntar=[]
  gallery.innerHTML=""
  document.getElementById("pastas-adicionadas").innerHTML=""
  document.getElementById("file-counter").textContent="Nenhum arquivo carregado"
  document.getElementById("juntar-info").textContent=""
  document.getElementById("resultado-juntar").style.display="none"
  document.getElementById("resultado-pastas").style.display="none"
  document.getElementById("resultado-pastas-btns").innerHTML=""
  document.getElementById("finalFileName").value=""
}

document.getElementById("actionBtn").addEventListener("click",async()=>{
  if(arquivos.length===0 && pastasJuntar.length===0){
    alert("Selecione pelo menos um arquivo ou pasta")
    return
  }
  if(pastasJuntar.length>0){
    await gerarPDFsPorPasta()
  }else{
    await gerarPDFAvulsos()
  }
})

async function gerarPDFAvulsos(){
  document.getElementById("juntar-info").textContent="Consolidando PDFs..."
  const mergedPdf=await PDFLib.PDFDocument.create()
  for(let file of arquivos){
    const bytes=await file.arrayBuffer()
    const pdf=await PDFLib.PDFDocument.load(bytes)
    const pages=await mergedPdf.copyPages(pdf,pdf.getPageIndices())
    pages.forEach(p=>mergedPdf.addPage(p))
  }
  resultadoJuntarBytes=await mergedPdf.save()
  const nomeInput=document.getElementById("finalFileName").value.trim()
  resultadoJuntarNome=(nomeInput||"arquivo_unificado")+".pdf"
  const resultadoFile=bytesParaFile(resultadoJuntarBytes,resultadoJuntarNome)
  const previewCanvas=document.getElementById("preview-juntar")
  const totalPags=await renderizarCapaNoCanvas(resultadoFile,previewCanvas)
  document.getElementById("resultado-juntar-nome").textContent=resultadoJuntarNome
  document.getElementById("resultado-juntar-pags").textContent=totalPags+" página"+(totalPags>1?"s":"")
  document.getElementById("resultado-juntar").style.display="block"
  document.getElementById("resultado-juntar").scrollIntoView({behavior:"smooth"})
  document.getElementById("juntar-info").textContent=arquivos.length+" arquivo(s) consolidados em "+totalPags+" páginas"
}

async function gerarPDFsPorPasta(){
  resultadosPastasJuntar=[]
  const btns=document.getElementById("resultado-pastas-btns")
  btns.innerHTML=""
  document.getElementById("resultado-pastas").style.display="none"

  const todasPastas=[...pastasJuntar]
  if(arquivos.length>0){
    const nomeAvulso=document.getElementById("finalFileName").value.trim()||"Arquivos_Avulsos"
    todasPastas.push({nomePasta:nomeAvulso,arquivos})
  }

  for(let i=0;i<todasPastas.length;i++){
    const {nomePasta,arquivos:arqs}=todasPastas[i]
    document.getElementById("juntar-info").textContent="Gerando: "+nomePasta+" ("+(i+1)+"/"+todasPastas.length+")..."
    const mergedPdf=await PDFLib.PDFDocument.create()
    const ordenados=[...arqs].sort((a,b)=>(a.webkitRelativePath||a.name).localeCompare(b.webkitRelativePath||b.name))
    for(let file of ordenados){
      try{
        let f=file
        if(file.type!=="application/pdf") f=await imagemParaPDF(file)
        const bytes=await f.arrayBuffer()
        const pdf=await PDFLib.PDFDocument.load(bytes)
        const pages=await mergedPdf.copyPages(pdf,pdf.getPageIndices())
        pages.forEach(p=>mergedPdf.addPage(p))
      }catch(e){ console.warn("Erro:",file.name,e) }
    }
    const pdfBytes=await mergedPdf.save()
    const nomeFinal=nomePasta+".pdf"
    resultadosPastasJuntar.push({nome:nomeFinal,bytes:pdfBytes})
    const btn=document.createElement("button")
    btn.className="btn generate"
    btn.textContent="⬇ "+nomeFinal
    btn.onclick=(()=>{const b=pdfBytes,n=nomeFinal;return()=>baixarArquivo(b,n)})()
    btns.appendChild(btn)
  }

  if(resultadosPastasJuntar.length>1){
    const btnZip=document.createElement("button")
    btnZip.className="btn primary"
    btnZip.textContent="📦 Baixar Tudo (ZIP)"
    btnZip.onclick=baixarTudoZipado
    btns.insertBefore(btnZip,btns.firstChild)
  }

  document.getElementById("juntar-info").textContent=todasPastas.length+" PDF"+(todasPastas.length>1?"s gerados":"gerado")+"!"
  document.getElementById("resultado-pastas").style.display="block"
  document.getElementById("resultado-pastas").scrollIntoView({behavior:"smooth"})
}

async function baixarTudoZipado(){
  const JSZip=await carregarJSZip()
  const zip=new JSZip()
  for(let item of resultadosPastasJuntar) zip.file(item.nome,item.bytes)
  const blob=await zip.generateAsync({type:"blob"})
  const link=document.createElement("a")
  link.href=URL.createObjectURL(blob)
  link.download="PDFs_Consolidados.zip"
  link.click()
}

function baixarResultadoJuntar(){
  if(!resultadoJuntarBytes) return
  baixarArquivo(resultadoJuntarBytes,resultadoJuntarNome)
}

function enviarResultadoParaDividir(){
  if(!resultadoJuntarBytes) return
  const file=bytesParaFile(resultadoJuntarBytes,resultadoJuntarNome)
  limparJuntar(); abrirDividir(file)
}

function enviarResultadoParaCompactar(){
  if(!resultadoJuntarBytes) return
  const file=bytesParaFile(resultadoJuntarBytes,resultadoJuntarNome)
  limparJuntar(); abrirCompactar([file])
}

function enviarResultadoParaOCR(){
  if(!resultadoJuntarBytes) return
  const file=bytesParaFile(resultadoJuntarBytes,resultadoJuntarNome)
  limparJuntar(); abrirOCR(file)
}


// ===================== DIVIDIR PDF =====================

let splitFileRef=null

document.getElementById("splitFileInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  await carregarFileNoDividir(file)
})

async function carregarFileNoDividir(file){
  splitFileRef=file
  resultadosDividirList=[]
  document.getElementById("resultado-dividir").style.display="none"

  const splitGallery=document.getElementById("split-gallery")
  splitGallery.innerHTML=""

  const {card, totalPaginas}=await criarThumbCard(file)
  splitGallery.appendChild(card)

  document.getElementById("dividir-info").textContent="1 arquivo | "+totalPaginas+" páginas | Clique na miniatura para ver páginas"
}

function limparDividir(){
  splitFileRef=null
  resultadosDividirList=[]
  document.getElementById("split-gallery").innerHTML=""
  document.getElementById("dividir-info").textContent=""
  document.getElementById("splitFileInput").value=""
  document.getElementById("resultado-dividir").style.display="none"
}

async function dividirPDF(){
  if(!splitFileRef){
    alert("Selecione um PDF primeiro")
    return
  }

  resultadosDividirList=[]
  document.getElementById("resultado-dividir").style.display="none"

  const file=splitFileRef
  const bytes=await file.arrayBuffer()
  const pdf=await PDFLib.PDFDocument.load(bytes)
  const totalPaginas=pdf.getPageCount()
  const modo=document.querySelector('input[name="splitMode"]:checked').value

  // DIVIDIR POR QUANTIDADE DE PAGINAS
  if(modo==="pages"){
    const qtd=parseInt(document.getElementById("pagesPerFile").value)
    if(!qtd || qtd<=0){
      alert("Digite a quantidade de páginas")
      return
    }
    let contador=1
    for(let i=0;i<totalPaginas;i+=qtd){
      const novoPdf=await PDFLib.PDFDocument.create()
      const paginas=[]
      for(let j=i;j<i+qtd && j<totalPaginas;j++) paginas.push(j)
      const copied=await novoPdf.copyPages(pdf,paginas)
      copied.forEach(p=>novoPdf.addPage(p))
      const pdfBytes=await novoPdf.save()
      resultadosDividirList.push({bytes:pdfBytes, nome:"parte_"+contador+".pdf"})
      contador++
    }
  }

  // PAGINAS IMPARES
  if(modo==="odd"){
    const novoPdf=await PDFLib.PDFDocument.create()
    const paginas=[]
    for(let i=0;i<totalPaginas;i++) if((i+1)%2!==0) paginas.push(i)
    const copied=await novoPdf.copyPages(pdf,paginas)
    copied.forEach(p=>novoPdf.addPage(p))
    const pdfBytes=await novoPdf.save()
    resultadosDividirList.push({bytes:pdfBytes, nome:"paginas_impares.pdf"})
  }

  // PAGINAS PARES
  if(modo==="even"){
    const novoPdf=await PDFLib.PDFDocument.create()
    const paginas=[]
    for(let i=0;i<totalPaginas;i++) if((i+1)%2===0) paginas.push(i)
    const copied=await novoPdf.copyPages(pdf,paginas)
    copied.forEach(p=>novoPdf.addPage(p))
    const pdfBytes=await novoPdf.save()
    resultadosDividirList.push({bytes:pdfBytes, nome:"paginas_pares.pdf"})
  }

  document.getElementById("dividir-info").textContent="Divisão concluída: "+resultadosDividirList.length+" arquivo(s) gerado(s)"
  document.getElementById("resultado-dividir").style.display="block"
  document.getElementById("resultado-dividir").scrollIntoView({behavior:"smooth"})
}

function baixarResultadosDividir(){
  for(let item of resultadosDividirList){
    baixarArquivo(item.bytes, item.nome)
  }
}

function enviarPartsParaCompactar(){
  if(resultadosDividirList.length===0) return
  const files=resultadosDividirList.map(item=>bytesParaFile(item.bytes, item.nome))
  limparDividir()
  abrirCompactar(files)
}

function enviarPartsParaJuntar(){
  if(resultadosDividirList.length===0) return
  const files=resultadosDividirList.map(item=>bytesParaFile(item.bytes, item.nome))
  limparDividir()
  abrirUnificador(files)
}


// ===================== OCR =====================

let ocrTextoGlobal=""
let ocrFileRef=null

document.getElementById("ocrInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  await carregarFileNoOCR(file)
})

async function carregarFileNoOCR(file){
  ocrFileRef=file
  ocrTextoGlobal=""
  document.getElementById("ocr-resultado-box").style.display="none"
  document.getElementById("ocr-status").textContent=""

  const ocrGallery=document.getElementById("ocr-gallery")
  ocrGallery.innerHTML=""

  const {card, totalPaginas}=await criarThumbCard(file)
  ocrGallery.appendChild(card)

  document.getElementById("ocr-info").textContent="1 arquivo | "+totalPaginas+" páginas | Clique na miniatura para ver páginas"
}

function limparOCR(){
  ocrTextoGlobal=""
  ocrFileRef=null
  document.getElementById("ocr-gallery").innerHTML=""
  document.getElementById("ocr-status").textContent=""
  document.getElementById("ocr-resultado-box").style.display="none"
  document.getElementById("ocr-info").textContent=""
  document.getElementById("ocrInput").value=""
}

async function reconhecerTexto(){
  const file=ocrFileRef

  if(!file){
    alert("Selecione um PDF primeiro")
    return
  }

  const status=document.getElementById("ocr-status")
  const resultadoBox=document.getElementById("ocr-resultado-box")

  status.textContent="Processando, aguarde..."
  resultadoBox.style.display="none"
  ocrTextoGlobal=""

  const arrayBuffer=await file.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  let textoCompleto=""

  for(let i=1;i<=pdf.numPages;i++){
    status.textContent="Processando página "+i+" de "+pdf.numPages+"..."
    const page=await pdf.getPage(i)
    const content=await page.getTextContent()
    const textoPage=content.items.map(item=>item.str).join(" ")
    textoCompleto+=textoPage+"\n"
  }

  if(textoCompleto.trim().length>0){
    ocrTextoGlobal=textoCompleto
    status.textContent="Texto reconhecido com sucesso! Clique em Copiar Texto."
    resultadoBox.style.display="block"
    resultadoBox.scrollIntoView({behavior:"smooth"})
  }else{
    status.textContent="PDF sem texto detectável."
  }
}

function copiarTexto(){
  if(!ocrTextoGlobal){
    alert("Nenhum texto reconhecido ainda.")
    return
  }
  navigator.clipboard.writeText(ocrTextoGlobal).then(()=>{
    alert("Texto copiado!")
  })
}

function enviarOCRFileParaJuntar(){
  if(!ocrFileRef) return
  const file=ocrFileRef
  limparOCR()
  abrirUnificador([file])
}

function enviarOCRFileParaCompactar(){
  if(!ocrFileRef) return
  const file=ocrFileRef
  limparOCR()
  abrirCompactar([file])
}

function enviarOCRFileParaDividir(){
  if(!ocrFileRef) return
  const file=ocrFileRef
  limparOCR()
  abrirDividir(file)
}


// ===================== COMPACTAR PDF =====================

let arquivosCompactar=[]

document.getElementById("compactarInput").addEventListener("change",async(e)=>{
  await carregarFilesNoCompactar(Array.from(e.target.files).filter(f=>f.type==="application/pdf"))
})

document.getElementById("compactarPastaInput").addEventListener("change",async(e)=>{
  await carregarFilesNoCompactar(Array.from(e.target.files).filter(f=>f.type==="application/pdf"))
})

async function carregarFilesNoCompactar(files){
  for(let file of files){
    arquivosCompactar.push(file)
    const {card}=await criarThumbCard(file)
    document.getElementById("compactar-lista").appendChild(card)
  }
  atualizarInfoCompactar()
}

function atualizarInfoCompactar(){
  const total=arquivosCompactar.length
  document.getElementById("compactar-info").textContent=total+" arquivo"+(total!==1?"s":"")+" | Clique na miniatura para ver páginas"
}

function limparCompactar(){
  arquivosCompactar=[]
  resultadosCompactarList=[]
  document.getElementById("compactar-lista").innerHTML=""
  document.getElementById("compactar-info").textContent=""
  document.getElementById("compactar-status").textContent=""
  document.getElementById("compactarInput").value=""
  document.getElementById("compactarPastaInput").value=""
  document.getElementById("resultado-compactar").style.display="none"
}

function formatarTamanho(bytes){
  if(bytes>=1024*1024) return (bytes/(1024*1024)).toFixed(2)+" MB"
  return (bytes/1024).toFixed(1)+" KB"
}

// Renderiza cada página do PDF como imagem JPEG recomprimida e monta novo PDF
async function compactarComImagensRecomprimidas(arrayBuffer, qualidade){
  const pdfSrc=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  const novoPdf=await PDFLib.PDFDocument.create()
  const escala=qualidade>0.5?1.2:0.9

  for(let i=1;i<=pdfSrc.numPages;i++){
    const page=await pdfSrc.getPage(i)
    const viewport=page.getViewport({scale:escala})
    const canvas=document.createElement("canvas")
    canvas.width=viewport.width
    canvas.height=viewport.height
    await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise

    const jpegDataUrl=canvas.toDataURL("image/jpeg", qualidade)
    const jpegBase64=jpegDataUrl.split(",")[1]
    const jpegBytes=Uint8Array.from(atob(jpegBase64),c=>c.charCodeAt(0))

    const jpegImage=await novoPdf.embedJpg(jpegBytes)
    const pdfPage=novoPdf.addPage([viewport.width, viewport.height])
    pdfPage.drawImage(jpegImage,{x:0,y:0,width:viewport.width,height:viewport.height})
  }

  return await novoPdf.save({useObjectStreams:true})
}

async function compactarPDFs(){
  if(arquivosCompactar.length===0){
    alert("Selecione pelo menos um PDF")
    return
  }

  resultadosCompactarList=[]
  document.getElementById("resultado-compactar").style.display="none"

  const modo=document.querySelector('input[name="compactMode"]:checked').value
  const status=document.getElementById("compactar-status")

  // qualidade JPEG: normal=0.75, extrema=0.35
  const qualidade=modo==="extrema"?0.35:0.75

  let totalOriginal=0
  let totalFinal=0
  let resumoLinhas=[]

  for(let i=0;i<arquivosCompactar.length;i++){
    const file=arquivosCompactar[i]
    status.textContent="Compactando "+file.name+" ("+(i+1)+" de "+arquivosCompactar.length+")..."

    const arrayBuffer=await file.arrayBuffer()
    const tamanhoOriginal=arrayBuffer.byteLength

    const pdfBytes=await compactarComImagensRecomprimidas(arrayBuffer, qualidade)
    const tamanhoFinal=pdfBytes.byteLength

    totalOriginal+=tamanhoOriginal
    totalFinal+=tamanhoFinal

    const reducao=((1-(tamanhoFinal/tamanhoOriginal))*100).toFixed(1)
    const nomeOriginal=file.name.replace(".pdf","")
    const nomeFinal=nomeOriginal+"_compactado.pdf"

    resultadosCompactarList.push({bytes:pdfBytes, nome:nomeFinal})
    resumoLinhas.push(
      file.name+": "+formatarTamanho(tamanhoOriginal)+" → "+formatarTamanho(tamanhoFinal)+" (−"+reducao+"%)"
    )
  }

  const reducaoTotal=((1-(totalFinal/totalOriginal))*100).toFixed(1)

  // Atualizar painel resultado com resumo
  const acoes=document.querySelector("#resultado-compactar .resultado-acoes")
  let resumoHtml='<div class="compactar-resumo">'
  resumoHtml+='<div class="resumo-total">Total: '+formatarTamanho(totalOriginal)+' → '+formatarTamanho(totalFinal)+' <strong style="color:#10b981">−'+reducaoTotal+'%</strong></div>'
  resumoLinhas.forEach(linha=>{
    resumoHtml+='<div class="resumo-linha">'+linha+'</div>'
  })
  resumoHtml+='</div>'

  // Remove resumo anterior se existir
  const resumoAnterior=document.getElementById("compactar-resumo-box")
  if(resumoAnterior) resumoAnterior.remove()

  const resumoEl=document.createElement("div")
  resumoEl.id="compactar-resumo-box"
  resumoEl.innerHTML=resumoHtml
  acoes.insertBefore(resumoEl, acoes.querySelector(".resultado-label"))

  status.textContent=""
  document.getElementById("resultado-compactar").style.display="block"
  document.getElementById("resultado-compactar").scrollIntoView({behavior:"smooth"})
}

function baixarResultadosCompactar(){
  for(let item of resultadosCompactarList){
    baixarArquivo(item.bytes, item.nome)
  }
}

function enviarCompactadosParaJuntar(){
  if(resultadosCompactarList.length===0) return
  const files=resultadosCompactarList.map(item=>bytesParaFile(item.bytes, item.nome))
  limparCompactar()
  abrirUnificador(files)
}

function enviarCompactadosParaDividir(){
  if(resultadosCompactarList.length===0) return
  const file=bytesParaFile(resultadosCompactarList[0].bytes, resultadosCompactarList[0].nome)
  limparCompactar()
  abrirDividir(file)
}

function enviarCompactadosParaOCR(){
  if(resultadosCompactarList.length===0) return
  const file=bytesParaFile(resultadosCompactarList[0].bytes, resultadosCompactarList[0].nome)
  limparCompactar()
  abrirOCR(file)
}


// ===================== CONVERTER PDF PARA WORD =====================

let converterFileRef=null
let resultadoWordBlob=null
let resultadoWordNome=""

document.getElementById("converterInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  await carregarFileNoConverter(file)
})

async function carregarFileNoConverter(file){
  converterFileRef=file
  resultadoWordBlob=null
  document.getElementById("resultado-converter").style.display="none"
  document.getElementById("converter-status").textContent=""

  const convGallery=document.getElementById("converter-gallery")
  convGallery.innerHTML=""

  const {card, totalPaginas}=await criarThumbCard(file)
  convGallery.appendChild(card)

  document.getElementById("converter-info").textContent="1 arquivo | "+totalPaginas+" páginas | Clique na miniatura para ver páginas"
}

function limparConverter(){
  converterFileRef=null
  resultadoWordBlob=null
  resultadoWordNome=""
  document.getElementById("converter-gallery").innerHTML=""
  document.getElementById("converter-status").textContent=""
  document.getElementById("resultado-converter").style.display="none"
  document.getElementById("converter-info").textContent=""
  document.getElementById("converterInput").value=""
}

// Gera arquivo .docx com o texto extraído do PDF usando estrutura XML mínima
async function converterParaWord(){
  if(!converterFileRef){
    alert("Selecione um PDF primeiro")
    return
  }

  const status=document.getElementById("converter-status")
  status.textContent="Extraindo texto do PDF..."
  document.getElementById("resultado-converter").style.display="none"

  const arrayBuffer=await converterFileRef.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise

  let paragrafos=[]

  for(let i=1;i<=pdf.numPages;i++){
    status.textContent="Processando página "+i+" de "+pdf.numPages+"..."
    const page=await pdf.getPage(i)
    const content=await page.getTextContent()

    // Agrupa itens por linha usando posição Y
    const linhas={}
    for(let item of content.items){
      const y=Math.round(item.transform[5])
      if(!linhas[y]) linhas[y]=[]
      linhas[y].push(item.str)
    }

    const ysOrdenados=Object.keys(linhas).map(Number).sort((a,b)=>b-a)
    for(let y of ysOrdenados){
      const linha=linhas[y].join(" ").trim()
      if(linha) paragrafos.push(linha)
    }

    // Separador de página
    if(i<pdf.numPages) paragrafos.push("--- Página "+(i+1)+" ---")
  }

  status.textContent="Gerando arquivo Word..."

  // Monta XML do docx manualmente (formato mínimo válido)
  const xmlParas=paragrafos.map(p=>{
    const texto=p
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;")
    return `<w:p><w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`
  }).join("\n")

  const docXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
${xmlParas}
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body>
</w:document>`

  const relsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

  const wordRelsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`

  const contentTypesXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`

  // Monta ZIP (docx é um ZIP) usando biblioteca JSZip via CDN inline
  const JSZip=await carregarJSZip()
  const zip=new JSZip()
  zip.file("[Content_Types].xml", contentTypesXml)
  zip.file("_rels/.rels", relsXml)
  zip.file("word/document.xml", docXml)
  zip.file("word/_rels/document.xml.rels", wordRelsXml)

  const blob=await zip.generateAsync({type:"blob", mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})

  resultadoWordBlob=blob
  resultadoWordNome=converterFileRef.name.replace(".pdf","")+".docx"

  status.textContent="Conversão concluída!"
  document.getElementById("resultado-converter").style.display="block"
  document.getElementById("resultado-converter").scrollIntoView({behavior:"smooth"})
}

function carregarJSZip(){
  return new Promise((resolve,reject)=>{
    if(window.JSZip){resolve(window.JSZip);return}
    const script=document.createElement("script")
    script.src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"
    script.onload=()=>resolve(window.JSZip)
    script.onerror=()=>reject(new Error("Falha ao carregar JSZip"))
    document.head.appendChild(script)
  })
}

function baixarResultadoWord(){
  if(!resultadoWordBlob) return
  const link=document.createElement("a")
  link.href=URL.createObjectURL(resultadoWordBlob)
  link.download=resultadoWordNome
  link.click()
}

function enviarConverterParaJuntar(){
  if(!converterFileRef) return
  const file=converterFileRef
  limparConverter()
  abrirUnificador([file])
}

function enviarConverterParaCompactar(){
  if(!converterFileRef) return
  const file=converterFileRef
  limparConverter()
  abrirCompactar([file])
}


// ===================== UTILITÁRIOS COMPARTILHADOS DE CONVERSÃO =====================

function criarInfoBox(secaoId, nomeId, nome){
  const box=document.getElementById(secaoId)
  document.getElementById(nomeId).textContent="📄 "+nome
  box.style.display="block"
}

async function carregarJSZipSeNecessario(){
  return carregarJSZip()
}

async function carregarSheetJS(){
  return new Promise((resolve,reject)=>{
    if(window.XLSX){resolve(window.XLSX);return}
    const script=document.createElement("script")
    script.src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"
    script.onload=()=>resolve(window.XLSX)
    script.onerror=()=>reject(new Error("Falha ao carregar SheetJS"))
    document.head.appendChild(script)
  })
}

async function gerarDocxDeParagrafos(paragrafos, nome){
  const JSZip=await carregarJSZip()
  const xmlParas=paragrafos.map(p=>{
    const texto=p
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;")
    return `<w:p><w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`
  }).join("\n")
  const docXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>${xmlParas}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
  const relsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
  const wordRelsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
  const contentTypesXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
  const zip=new JSZip()
  zip.file("[Content_Types].xml",contentTypesXml)
  zip.file("_rels/.rels",relsXml)
  zip.file("word/document.xml",docXml)
  zip.file("word/_rels/document.xml.rels",wordRelsXml)
  return await zip.generateAsync({type:"blob",mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"})
}

async function gerarPDFDeTabela(linhas, titulo){
  const novoPdf=await PDFLib.PDFDocument.create()
  const fontSize=10
  const margin=40
  const lineHeight=16
  const pageWidth=842
  const pageHeight=595
  let page=novoPdf.addPage([pageWidth,pageHeight])
  let y=pageHeight-margin

  // Título
  page.drawText(titulo,{x:margin,y,size:14,color:PDFLib.rgb(0.2,0.2,0.8)})
  y-=lineHeight*2

  for(let linha of linhas){
    if(y<margin+lineHeight){
      page=novoPdf.addPage([pageWidth,pageHeight])
      y=pageHeight-margin
    }
    const texto=linha.join("   |   ").substring(0,120)
    const textoSeguro=texto.replace(/[^\x20-\x7E]/g," ")
    page.drawText(textoSeguro,{x:margin,y,size:fontSize,color:PDFLib.rgb(0.1,0.1,0.1)})
    y-=lineHeight
  }

  return await novoPdf.save()
}


// ===================== WORD PARA PDF =====================

let wordpdfFileRef=null
let resultadoWordPDFBytes=null
let resultadoWordPDFNome=""

document.getElementById("wordpdfInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  carregarFileNoWordPDF(file)
})

function carregarFileNoWordPDF(file){
  wordpdfFileRef=file
  resultadoWordPDFBytes=null
  document.getElementById("resultado-wordpdf").style.display="none"
  document.getElementById("wordpdf-status").textContent=""
  criarInfoBox("wordpdf-info-box","wordpdf-nome",file.name)
  document.getElementById("wordpdf-bar-info").textContent="1 arquivo selecionado: "+file.name
}

function limparWordPDF(){
  wordpdfFileRef=null
  resultadoWordPDFBytes=null
  resultadoWordPDFNome=""
  document.getElementById("wordpdf-info-box").style.display="none"
  document.getElementById("wordpdf-status").textContent=""
  document.getElementById("resultado-wordpdf").style.display="none"
  document.getElementById("wordpdf-bar-info").textContent=""
  document.getElementById("wordpdfInput").value=""
}

async function converterWordParaPDF(){
  if(!wordpdfFileRef){alert("Selecione um arquivo Word primeiro");return}
  const status=document.getElementById("wordpdf-status")
  status.textContent="Lendo documento Word..."

  const mammoth=await carregarMammoth()
  const arrayBuffer=await wordpdfFileRef.arrayBuffer()
  const resultado=await mammoth.extractRawText({arrayBuffer})
  const texto=resultado.value

  status.textContent="Gerando PDF..."

  const paragrafos=texto.split("\n").filter(p=>p.trim().length>0)
  const novoPdf=await PDFLib.PDFDocument.create()
  const fontSize=11
  const margin=50
  const lineHeight=18
  const pageWidth=595
  const pageHeight=842
  let page=novoPdf.addPage([pageWidth,pageHeight])
  let y=pageHeight-margin

  for(let paragrafo of paragrafos){
    // Quebra linhas longas
    const palavras=paragrafo.split(" ")
    let linhaAtual=""
    for(let palavra of palavras){
      const teste=linhaAtual?linhaAtual+" "+palavra:palavra
      if(teste.length>80){
        if(y<margin+lineHeight){page=novoPdf.addPage([pageWidth,pageHeight]);y=pageHeight-margin}
        const textoSeguro=linhaAtual.replace(/[^\x20-\x7E]/g," ")
        if(textoSeguro.trim()) page.drawText(textoSeguro,{x:margin,y,size:fontSize,color:PDFLib.rgb(0,0,0)})
        y-=lineHeight
        linhaAtual=palavra
      }else{
        linhaAtual=teste
      }
    }
    if(linhaAtual.trim()){
      if(y<margin+lineHeight){page=novoPdf.addPage([pageWidth,pageHeight]);y=pageHeight-margin}
      const textoSeguro=linhaAtual.replace(/[^\x20-\x7E]/g," ")
      page.drawText(textoSeguro,{x:margin,y,size:fontSize,color:PDFLib.rgb(0,0,0)})
      y-=lineHeight
    }
    y-=4 // espaço entre parágrafos
  }

  resultadoWordPDFBytes=await novoPdf.save()
  resultadoWordPDFNome=wordpdfFileRef.name.replace(/\.(doc|docx)$/i,"")+".pdf"
  status.textContent="Conversão concluída!"
  document.getElementById("resultado-wordpdf").style.display="block"
  document.getElementById("resultado-wordpdf").scrollIntoView({behavior:"smooth"})
}

function carregarMammoth(){
  return new Promise((resolve,reject)=>{
    if(window.mammoth){resolve(window.mammoth);return}
    const script=document.createElement("script")
    script.src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js"
    script.onload=()=>resolve(window.mammoth)
    script.onerror=()=>reject(new Error("Falha ao carregar Mammoth"))
    document.head.appendChild(script)
  })
}

function baixarResultadoWordPDF(){
  if(!resultadoWordPDFBytes) return
  baixarArquivo(resultadoWordPDFBytes, resultadoWordPDFNome)
}

function enviarWordPDFParaCompactar(){
  if(!resultadoWordPDFBytes) return
  const file=bytesParaFile(resultadoWordPDFBytes, resultadoWordPDFNome)
  limparWordPDF()
  abrirCompactar([file])
}

function enviarWordPDFParaDividir(){
  if(!resultadoWordPDFBytes) return
  const file=bytesParaFile(resultadoWordPDFBytes, resultadoWordPDFNome)
  limparWordPDF()
  abrirDividir(file)
}

function enviarWordPDFParaJuntar(){
  if(!resultadoWordPDFBytes) return
  const file=bytesParaFile(resultadoWordPDFBytes, resultadoWordPDFNome)
  limparWordPDF()
  abrirUnificador([file])
}


// ===================== EXCEL PARA PDF =====================

let excelpdfFileRef=null
let resultadoExcelPDFBytes=null
let resultadoExcelPDFNome=""

document.getElementById("excelpdfInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  carregarFileNoExcelPDF(file)
})

function carregarFileNoExcelPDF(file){
  excelpdfFileRef=file
  resultadoExcelPDFBytes=null
  document.getElementById("resultado-excelpdf").style.display="none"
  document.getElementById("excelpdf-status").textContent=""
  criarInfoBox("excelpdf-info-box","excelpdf-nome",file.name)
  document.getElementById("excelpdf-bar-info").textContent="1 arquivo selecionado: "+file.name
}

function limparExcelPDF(){
  excelpdfFileRef=null
  resultadoExcelPDFBytes=null
  resultadoExcelPDFNome=""
  document.getElementById("excelpdf-info-box").style.display="none"
  document.getElementById("excelpdf-status").textContent=""
  document.getElementById("resultado-excelpdf").style.display="none"
  document.getElementById("excelpdf-bar-info").textContent=""
  document.getElementById("excelpdfInput").value=""
}

async function converterExcelParaPDF(){
  if(!excelpdfFileRef){alert("Selecione um arquivo Excel primeiro");return}
  const status=document.getElementById("excelpdf-status")
  status.textContent="Lendo planilha Excel..."

  const XLSX=await carregarSheetJS()
  const arrayBuffer=await excelpdfFileRef.arrayBuffer()
  const workbook=XLSX.read(arrayBuffer,{type:"array"})

  status.textContent="Gerando PDF..."

  const novoPdf=await PDFLib.PDFDocument.create()
  const fontSize=9
  const margin=30
  const lineHeight=14
  const pageWidth=842
  const pageHeight=595

  for(let i=0;i<workbook.SheetNames.length;i++){
    const sheetName=workbook.SheetNames[i]
    status.textContent="Processando aba: "+sheetName+" ("+(i+1)+" de "+workbook.SheetNames.length+")..."
    const sheet=workbook.Sheets[sheetName]
    const dados=XLSX.utils.sheet_to_json(sheet,{header:1,defval:""})

    let page=novoPdf.addPage([pageWidth,pageHeight])
    let y=pageHeight-margin

    // Título da aba
    page.drawText("Aba: "+sheetName,{x:margin,y,size:12,color:PDFLib.rgb(0.2,0.2,0.8)})
    y-=lineHeight*2

    for(let linha of dados){
      if(y<margin+lineHeight){
        page=novoPdf.addPage([pageWidth,pageHeight])
        y=pageHeight-margin
      }
      const texto=linha.map(c=>String(c).substring(0,20)).join("  |  ").substring(0,130)
      const textoSeguro=texto.replace(/[^\x20-\x7E]/g," ")
      page.drawText(textoSeguro,{x:margin,y,size:fontSize,color:PDFLib.rgb(0.1,0.1,0.1)})
      y-=lineHeight
    }
  }

  resultadoExcelPDFBytes=await novoPdf.save()
  resultadoExcelPDFNome=excelpdfFileRef.name.replace(/\.(xls|xlsx)$/i,"")+".pdf"
  status.textContent="Conversão concluída!"
  document.getElementById("resultado-excelpdf").style.display="block"
  document.getElementById("resultado-excelpdf").scrollIntoView({behavior:"smooth"})
}

function baixarResultadoExcelPDF(){
  if(!resultadoExcelPDFBytes) return
  baixarArquivo(resultadoExcelPDFBytes, resultadoExcelPDFNome)
}

function enviarExcelPDFParaCompactar(){
  if(!resultadoExcelPDFBytes) return
  const file=bytesParaFile(resultadoExcelPDFBytes, resultadoExcelPDFNome)
  limparExcelPDF()
  abrirCompactar([file])
}

function enviarExcelPDFParaDividir(){
  if(!resultadoExcelPDFBytes) return
  const file=bytesParaFile(resultadoExcelPDFBytes, resultadoExcelPDFNome)
  limparExcelPDF()
  abrirDividir(file)
}

function enviarExcelPDFParaJuntar(){
  if(!resultadoExcelPDFBytes) return
  const file=bytesParaFile(resultadoExcelPDFBytes, resultadoExcelPDFNome)
  limparExcelPDF()
  abrirUnificador([file])
}


// ===================== EXCEL PARA WORD =====================

let excelwordFileRef=null
let resultadoExcelWordBlob=null
let resultadoExcelWordNome=""

document.getElementById("excelwordInput").addEventListener("change",async(e)=>{
  const file=e.target.files[0]
  if(!file) return
  carregarFileNoExcelWord(file)
})

function carregarFileNoExcelWord(file){
  excelwordFileRef=file
  resultadoExcelWordBlob=null
  document.getElementById("resultado-excelword").style.display="none"
  document.getElementById("excelword-status").textContent=""
  criarInfoBox("excelword-info-box","excelword-nome",file.name)
  document.getElementById("excelword-bar-info").textContent="1 arquivo selecionado: "+file.name
}

function limparExcelWord(){
  excelwordFileRef=null
  resultadoExcelWordBlob=null
  resultadoExcelWordNome=""
  document.getElementById("excelword-info-box").style.display="none"
  document.getElementById("excelword-status").textContent=""
  document.getElementById("resultado-excelword").style.display="none"
  document.getElementById("excelword-bar-info").textContent=""
  document.getElementById("excelwordInput").value=""
}

async function converterExcelParaWord(){
  if(!excelwordFileRef){alert("Selecione um arquivo Excel primeiro");return}
  const status=document.getElementById("excelword-status")
  status.textContent="Lendo planilha Excel..."

  const XLSX=await carregarSheetJS()
  const arrayBuffer=await excelwordFileRef.arrayBuffer()
  const workbook=XLSX.read(arrayBuffer,{type:"array"})

  status.textContent="Gerando Word..."

  let paragrafos=[]

  for(let i=0;i<workbook.SheetNames.length;i++){
    const sheetName=workbook.SheetNames[i]
    const sheet=workbook.Sheets[sheetName]
    const dados=XLSX.utils.sheet_to_json(sheet,{header:1,defval:""})

    paragrafos.push("=== Aba: "+sheetName+" ===")
    paragrafos.push("")

    for(let linha of dados){
      const texto=linha.map(c=>String(c)).join("\t | \t")
      if(texto.trim()) paragrafos.push(texto)
    }
    paragrafos.push("")
    paragrafos.push("--- Fim da aba "+sheetName+" ---")
    paragrafos.push("")
  }

  resultadoExcelWordBlob=await gerarDocxDeParagrafos(paragrafos, excelwordFileRef.name)
  resultadoExcelWordNome=excelwordFileRef.name.replace(/\.(xls|xlsx)$/i,"")+".docx"
  status.textContent="Conversão concluída!"
  document.getElementById("resultado-excelword").style.display="block"
  document.getElementById("resultado-excelword").scrollIntoView({behavior:"smooth"})
}

function baixarResultadoExcelWord(){
  if(!resultadoExcelWordBlob) return
  const link=document.createElement("a")
  link.href=URL.createObjectURL(resultadoExcelWordBlob)
  link.download=resultadoExcelWordNome
  link.click()
}


// ===================== JUNTAR PASTAS =====================

const MAX_PASTAS = 20
let pastasLista = [] // [{nomePasta, arquivos:[File]}]
let resultadosPastasBytes = [] // [{nome, bytes}]
let pastaInputAtivo = null

function adicionarPasta(){
  if(pastasLista.length >= MAX_PASTAS){
    alert("Limite de "+MAX_PASTAS+" pastas atingido.")
    return
  }
  const input = document.createElement("input")
  input.type = "file"
  input.webkitdirectory = true
  input.style.display = "none"
  document.body.appendChild(input)
  input.addEventListener("change", async(e)=>{
    const files = Array.from(e.target.files).filter(f=>f.type==="application/pdf")
    if(files.length===0){
      alert("Nenhum PDF encontrado nessa pasta.")
      document.body.removeChild(input)
      return
    }
    // Nome da pasta: pega do caminho do primeiro arquivo
    const nomePasta = files[0].webkitRelativePath.split("/")[0] || ("Pasta "+(pastasLista.length+1))
    pastasLista.push({nomePasta, arquivos: files})
    document.body.removeChild(input)
    await renderizarPastaCard(pastasLista.length-1)
    atualizarInfoJuntarPastas()
  })
  input.click()
}

async function renderizarPastaCard(idx){
  const {nomePasta, arquivos} = pastasLista[idx]
  const lista = document.getElementById("pastas-lista")

  const card = document.createElement("div")
  card.className = "pasta-card"
  card.id = "pasta-card-"+idx

  // Header da pasta
  const header = document.createElement("div")
  header.className = "pasta-card-header"

  const numero = document.createElement("div")
  numero.className = "pasta-numero"
  numero.textContent = (idx+1)+"."

  const info = document.createElement("div")
  info.className = "pasta-card-info"

  const nomeEl = document.createElement("div")
  nomeEl.className = "pasta-nome"
  nomeEl.textContent = "📁 "+nomePasta

  const qtdEl = document.createElement("div")
  qtdEl.className = "pasta-qtd"
  qtdEl.textContent = arquivos.length+" PDF"+(arquivos.length>1?"s":"")+" encontrado"+(arquivos.length>1?"s":"")

  const btnRemover = document.createElement("button")
  btnRemover.className = "btn danger btn-sm"
  btnRemover.textContent = "✕ Remover"
  btnRemover.onclick = ()=> removerPasta(idx)

  info.appendChild(nomeEl)
  info.appendChild(qtdEl)
  header.appendChild(numero)
  header.appendChild(info)
  header.appendChild(btnRemover)
  card.appendChild(header)

  // Miniaturas dos arquivos da pasta
  const thumbs = document.createElement("div")
  thumbs.className = "pasta-thumbs"

  for(let file of arquivos.slice(0,10)){
    try{
      const {card:thumb} = await criarThumbCard(file)
      thumbs.appendChild(thumb)
    }catch(e){}
  }
  if(arquivos.length>10){
    const mais = document.createElement("div")
    mais.className = "thumb-mais"
    mais.textContent = "+"+(arquivos.length-10)+" mais"
    thumbs.appendChild(mais)
  }

  card.appendChild(thumbs)
  lista.appendChild(card)
}

function removerPasta(idx){
  pastasLista.splice(idx,1)
  // Re-renderizar tudo
  document.getElementById("pastas-lista").innerHTML=""
  for(let i=0;i<pastasLista.length;i++){
    renderizarPastaCard(i)
  }
  atualizarInfoJuntarPastas()
}

function atualizarInfoJuntarPastas(){
  const total = pastasLista.length
  document.getElementById("pastas-contador").textContent = total+"/"+MAX_PASTAS+" pastas"
  document.getElementById("juntarpastas-info").textContent = total+" pasta"+(total!==1?"s":"")+" | "+(total<MAX_PASTAS?"Pode adicionar mais":"Limite atingido")
  document.getElementById("btn-add-pasta").disabled = total>=MAX_PASTAS
  document.getElementById("resultado-juntarpastas").style.display="none"
}

function limparJuntarPastas(){
  pastasLista=[]
  resultadosPastasBytes=[]
  document.getElementById("pastas-lista").innerHTML=""
  document.getElementById("juntarpastas-status").textContent=""
  document.getElementById("resultado-juntarpastas").style.display="none"
  document.getElementById("resultado-juntarpastas-btns").innerHTML=""
  document.getElementById("juntarpastas-info").textContent=""
  document.getElementById("pastas-contador").textContent=""
  document.getElementById("btn-add-pasta").disabled=false
}

async function gerarPDFsDasPastas(){
  if(pastasLista.length===0){
    alert("Adicione pelo menos uma pasta.")
    return
  }

  resultadosPastasBytes=[]
  const status=document.getElementById("juntarpastas-status")
  const btns=document.getElementById("resultado-juntarpastas-btns")
  btns.innerHTML=""
  document.getElementById("resultado-juntarpastas").style.display="none"

  for(let i=0;i<pastasLista.length;i++){
    const {nomePasta, arquivos} = pastasLista[i]
    status.textContent="Gerando PDF da "+nomePasta+" ("+(i+1)+" de "+pastasLista.length+")..."

    const mergedPdf = await PDFLib.PDFDocument.create()

    // Ordena arquivos por caminho para manter subpastas em ordem
    const arquivosOrdenados = [...arquivos].sort((a,b)=>a.webkitRelativePath.localeCompare(b.webkitRelativePath))

    for(let file of arquivosOrdenados){
      try{
        const bytes = await file.arrayBuffer()
        const pdf = await PDFLib.PDFDocument.load(bytes)
        const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices())
        pages.forEach(p=>mergedPdf.addPage(p))
      }catch(e){
        console.warn("Erro ao processar "+file.name, e)
      }
    }

    const pdfBytes = await mergedPdf.save()
    const nomeFinal = nomePasta+".pdf"
    resultadosPastasBytes.push({nome:nomeFinal, bytes:pdfBytes, qtd:arquivos.length})

    // Criar botão de download para esta pasta
    const btn = document.createElement("button")
    btn.className = "btn generate"
    btn.textContent = "⬇ "+nomeFinal+" ("+arquivos.length+" docs)"
    btn.onclick = (()=>{
      const b=pdfBytes, n=nomeFinal
      return ()=>baixarArquivo(b,n)
    })()
    btns.appendChild(btn)
  }

  // Botão baixar todos
  if(resultadosPastasBytes.length>1){
    const btnTodos = document.createElement("button")
    btnTodos.className = "btn primary"
    btnTodos.textContent = "⬇ Baixar Todos ("+resultadosPastasBytes.length+")"
    btnTodos.onclick = ()=>{
      for(let item of resultadosPastasBytes) baixarArquivo(item.bytes, item.nome)
    }
    btns.insertBefore(btnTodos, btns.firstChild)
  }

  status.textContent="Concluído! "+resultadosPastasBytes.length+" PDF"+(resultadosPastasBytes.length>1?"s gerados":"gerado")+"."
  document.getElementById("resultado-juntarpastas").style.display="block"
  document.getElementById("resultado-juntarpastas").scrollIntoView({behavior:"smooth"})
}

// ===================== DRAG AND DROP GLOBAL =====================
// Permite arrastar arquivos ou pastas diretamente para as áreas

document.addEventListener("dragover",(e)=>{
  e.preventDefault()
})

document.addEventListener("drop",(e)=>{
  e.preventDefault()
})

// ---------- JUNTAR PDF DROP ----------

const dropJuntar=document.getElementById("unificador")

if(dropJuntar){

dropJuntar.addEventListener("dragover",(e)=>{
  e.preventDefault()
  dropJuntar.classList.add("dragover")
})

dropJuntar.addEventListener("dragleave",(e)=>{
  dropJuntar.classList.remove("dragover")
})

dropJuntar.addEventListener("drop",async(e)=>{
  e.preventDefault()
  dropJuntar.classList.remove("dragover")

  const items=e.dataTransfer.items

  if(items){
    let files=[]

    for(let i=0;i<items.length;i++){
      const entry=items[i].webkitGetAsEntry()

      if(entry && entry.isDirectory){
        const reader=entry.createReader()

        await new Promise(resolve=>{
          reader.readEntries(async(entries)=>{
            for(let ent of entries){
              if(ent.isFile){
                await new Promise(res=>{
                  ent.file(file=>{
                    files.push(file)
                    res()
                  })
                })
              }
            }
            resolve()
          })
        })

      }else{
        const file=items[i].getAsFile()
        if(file) files.push(file)
      }
    }

    if(files.length>0){
      carregarFilesNoJuntar(files)
    }
  }

})

}

// ---------- DIVIDIR PDF DROP ----------

const dropDividir=document.getElementById("dividir-pdf")

if(dropDividir){

dropDividir.addEventListener("dragover",(e)=>{
  e.preventDefault()
  dropDividir.classList.add("dragover")
})

dropDividir.addEventListener("dragleave",()=>{
  dropDividir.classList.remove("dragover")
})

dropDividir.addEventListener("drop",async(e)=>{
  e.preventDefault()
  dropDividir.classList.remove("dragover")

  const file=e.dataTransfer.files[0]

  if(file && file.type==="application/pdf"){
    carregarFileNoDividir(file)
  }else{
    alert("Arraste um arquivo PDF.")
  }

})

}

// ---------- COMPACTAR PDF DROP ----------

const dropCompactar=document.getElementById("compactar-section")

if(dropCompactar){

dropCompactar.addEventListener("dragover",(e)=>{
  e.preventDefault()
  dropCompactar.classList.add("dragover")
})

dropCompactar.addEventListener("dragleave",()=>{
  dropCompactar.classList.remove("dragover")
})

dropCompactar.addEventListener("drop",async(e)=>{
  e.preventDefault()
  dropCompactar.classList.remove("dragover")

  const files=Array.from(e.dataTransfer.files).filter(f=>f.type==="application/pdf")

  if(files.length>0){
    carregarFilesNoCompactar(files)
  }else{
    alert("Arraste arquivos PDF.")
  }

})

}
dropJuntar.addEventListener("drop", async (e) => {
  e.preventDefault()
  dropJuntar.classList.remove("dragover")

  const items = e.dataTransfer.items
  let arquivosDrop = []

  async function lerEntrada(entry) {
    return new Promise((resolve) => {

      if (entry.isFile) {

        entry.file(file => {
          arquivosDrop.push(file)
          resolve()
        })

      }

      else if (entry.isDirectory) {

        const reader = entry.createReader()

        reader.readEntries(async (entries) => {

          for (const ent of entries) {
            await lerEntrada(ent)
          }

          resolve()

        })

      }

    })
  }

  for (let i = 0; i < items.length; i++) {

    const entry = items[i].webkitGetAsEntry()

    if (entry) {
      await lerEntrada(entry)
    }

  }

  if (arquivosDrop.length > 0) {
    carregarFilesNoJuntar(arquivosDrop)
  }

})
