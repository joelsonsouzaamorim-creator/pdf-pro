// ===================== ESTADO GLOBAL DE RESULTADOS =====================
// Guarda bytes prontos em memória para permitir baixar ou enviar para outra ferramenta

let resultadoJuntarBytes=null
let resultadoJuntarNome=""
let resultadosDividirList=[]   // [{bytes, nome}]
let resultadosCompactarList=[] // [{bytes, nome}]
let pastaArquivoArrastandoInfo=null
const ordemPaginasPorArquivo=new WeakMap()
let modalPaginasEstado={file:null, ordem:[], itens:new Map(), total:0}
let modalPaginaArrastandoPos=null

const historicoAtualizacoes=[
  {
    titulo:"Organizacao de folhas dentro do PDF",
    texto:"Ao abrir um PDF pela miniatura, agora voce pode arrastar as paginas para definir a ordem antes de juntar ou dividir."
  },
  {
    titulo:"Selecao de pasta com ZIP integrado",
    texto:"Compactados encontrados dentro da pasta principal agora entram na mesma pasta escolhida, em vez de aparecer separados."
  },
  {
    titulo:"Ordenacao por arrastar nas pastas",
    texto:"Os PDFs dentro das pastas do Juntar PDF agora podem ser reorganizados por arrastar as miniaturas."
  },
  {
    titulo:"Tela de progresso centralizada",
    texto:"Juntar PDF e Dividir PDF ganharam uma tela dedicada de progresso com resultado central e botao de download."
  },
  {
    titulo:"Botao voltar no processamento",
    texto:"A tela de progresso agora permite voltar para a ferramenta sem perder o andamento."
  },
  {
    titulo:"Compactar PDF em breve",
    texto:"A opcao de compactacao foi sinalizada como Em breve para evitar uso de um fluxo destrutivo."
  },
  {
    titulo:"Layout mais profissional",
    texto:"Topo, cards, areas de upload, barra fixa e estados visuais receberam refinamento mais corporativo."
  },
  {
    titulo:"Assinatura visual preservada",
    texto:"Juntar, dividir e agrupar pastas tentam manter a aparencia da assinatura visivel nos PDFs editados."
  },
  {
    titulo:"Suporte a ZIP, RAR e 7Z",
    texto:"Arquivos compactados passaram a ser aceitos nos fluxos de juntar e juntar pastas."
  },
  {
    titulo:"Miniaturas menores e mais limpas",
    texto:"As miniaturas foram reduzidas e as pastas ganharam um visual expandido mais organizado."
  },
  {
    titulo:"Progresso mais intuitivo",
    texto:"A barra de progresso ficou mais suave, com etapas mais claras e finalizacao menos brusca."
  }
]
const historicoAtualizacoesVersao="2026-04-09-2"
const chaveAtualizacoesLidas="pdflicita_pro_updates_read_version"


// ===================== NAVEGAÇÃO =====================

function atualizarAtalhosTopo(mostrar){
  const atalhos=document.getElementById("topbar-shortcuts")
  if(!atalhos) return
  atalhos.style.display=mostrar ? "flex" : "none"
}

inicializarPainelAtualizacoes()

function inicializarPainelAtualizacoes(){
  const btn=document.getElementById("updates-toggle")
  const painel=document.getElementById("updates-panel")
  const lista=document.getElementById("updates-list")
  const count=document.getElementById("updates-count")
  const badge=document.getElementById("updates-panel-badge")
  if(!btn || !painel || !lista || !count || !badge) return

  lista.innerHTML=""
  historicoAtualizacoes.forEach((item, idx)=>{
    const card=document.createElement("div")
    card.className="update-item"
    card.innerHTML=`
      <div class="update-item-title">${String(idx+1).padStart(2,"0")} • ${item.titulo}</div>
      <div class="update-item-text">${item.texto}</div>
    `
    lista.appendChild(card)
  })

  const atualizarEstadoBadge=()=>{
    const lido=localStorage.getItem(chaveAtualizacoesLidas)===historicoAtualizacoesVersao
    count.textContent=lido ? "0" : String(historicoAtualizacoes.length)
    badge.textContent=historicoAtualizacoes.length+" itens"
    btn.classList.toggle("is-read", lido)
    count.classList.toggle("is-zero", lido)
  }

  const marcarAtualizacoesComoLidas=()=>{
    try{
      localStorage.setItem(chaveAtualizacoesLidas, historicoAtualizacoesVersao)
    }catch(e){}
    atualizarEstadoBadge()
  }

  atualizarEstadoBadge()

  btn.onclick=(e)=>{
    e.stopPropagation()
    const aberto=painel.style.display==="block"
    painel.style.display=aberto ? "none" : "block"
    btn.setAttribute("aria-expanded", aberto ? "false" : "true")
    if(!aberto) marcarAtualizacoesComoLidas()
  }

  painel.addEventListener("click",(e)=>e.stopPropagation())
  document.addEventListener("click",()=>{
    painel.style.display="none"
    btn.setAttribute("aria-expanded","false")
  })
}

function abrirUnificador(filesParaCarregar){
  ocultarTudo()
  document.getElementById("unificador").style.display="block"
  atualizarAtalhosTopo(true)
  if(filesParaCarregar && filesParaCarregar.length>0){
    carregarFilesNoJuntar(filesParaCarregar)
  }
}

function abrirDividir(fileParaCarregar){
  ocultarTudo()
  document.getElementById("dividir-pdf").style.display="block"
  atualizarAtalhosTopo(true)
  if(fileParaCarregar){
    carregarFileNoDividir(fileParaCarregar)
  }
}

function abrirOCR(fileParaCarregar){
  ocultarTudo()
  document.getElementById("ocr-section").style.display="block"
  atualizarAtalhosTopo(true)
  if(fileParaCarregar){
    carregarFileNoOCR(fileParaCarregar)
  }
}

function abrirCompactar(filesParaCarregar){
  ocultarTudo()
  document.getElementById("compactar-section").style.display="block"
  atualizarAtalhosTopo(true)
  if(filesParaCarregar && filesParaCarregar.length>0){
    carregarFilesNoCompactar(filesParaCarregar)
  }
}

function abrirConverter(fileParaCarregar){
  ocultarTudo()
  document.getElementById("converter-section").style.display="block"
  atualizarAtalhosTopo(true)
  if(fileParaCarregar){
    carregarFileNoConverter(fileParaCarregar)
  }
}

function abrirJuntarPastas(){
  ocultarTudo()
  document.getElementById("juntarpastas-section").style.display="block"
  configurarInterfaceJuntarPastas()
  atualizarAtalhosTopo(true)
}

function configurarInterfaceJuntarPastas(){
  const secao=document.getElementById("juntarpastas-section")
  if(!secao) return

  const descricao=secao.querySelector(".unificador-header p")
  if(descricao){
    descricao.textContent="Cada pasta ou compactado gera um PDF separado - ate 20 entradas"
  }

  const btnCompactado=document.getElementById("btn-add-compactado")
  if(btnCompactado){
    btnCompactado.textContent="+ Adicionar Pasta Compactada"
    btnCompactado.classList.add("primary","btn-compactado")
    btnCompactado.classList.remove("atalho")
  }

  let hint=secao.querySelector(".pastas-adicionar-hint")
  if(!hint){
    hint=document.createElement("p")
    hint.className="pastas-adicionar-hint"
    const alvo=document.getElementById("pastas-adicionar-btn")
    if(alvo?.parentNode){
      alvo.parentNode.insertBefore(hint, alvo.nextSibling)
    }
  }
  if(hint){
    hint.textContent="Aceita ZIP, RAR, 7Z, TAR e outros compactados compativeis."
  }
}

function abrirWordParaPDF(fileParaCarregar){
  ocultarTudo()
  document.getElementById("wordpdf-section").style.display="block"
  atualizarAtalhosTopo(true)
  if(fileParaCarregar) carregarFileNoWordPDF(fileParaCarregar)
}

function abrirExcelParaPDF(fileParaCarregar){
  ocultarTudo()
  document.getElementById("excelpdf-section").style.display="block"
  atualizarAtalhosTopo(true)
  if(fileParaCarregar) carregarFileNoExcelPDF(fileParaCarregar)
}

function abrirExcelParaWord(fileParaCarregar){
  ocultarTudo()
  document.getElementById("excelword-section").style.display="block"
  atualizarAtalhosTopo(true)
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

function limparEstadosAoVoltarMenu(){
  try{ limparJuntar() }catch(e){}
  try{ limparDividir() }catch(e){}
  try{ limparOCR() }catch(e){}
  try{ limparCompactar() }catch(e){}
  try{ limparConverter() }catch(e){}
  try{ limparWordPDF() }catch(e){}
  try{ limparExcelPDF() }catch(e){}
  try{ limparExcelWord() }catch(e){}
  try{ limparJuntarPastas() }catch(e){}
}

function voltarMenu(){
  limparEstadosAoVoltarMenu()
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
  atualizarAtalhosTopo(false)
}

function irParaFerramenta(destino){
  limparEstadosAoVoltarMenu()
  if(destino==="juntar") abrirUnificador()
  else if(destino==="dividir") abrirDividir()
  else if(destino==="compactar") abrirCompactar()
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

async function carregarPdfPreservandoAssinaturaVisual(bytes){
  const pdf=await PDFLib.PDFDocument.load(bytes)
  try{
    const form=pdf.getForm()
    form.flatten()
  }catch(e){}
  return pdf
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

function ordemPaginasNatural(totalPaginas){
  return Array.from({length:totalPaginas}, (_,idx)=>idx)
}

function obterOrdemPaginasArquivo(file, totalPaginas){
  const ordemSalva=ordemPaginasPorArquivo.get(file)
  if(!Array.isArray(ordemSalva) || ordemSalva.length!==totalPaginas){
    return ordemPaginasNatural(totalPaginas)
  }
  const unicos=new Set(ordemSalva)
  if(unicos.size!==totalPaginas) return ordemPaginasNatural(totalPaginas)
  return [...ordemSalva]
}

function ordemPaginasEhPersonalizada(ordem){
  return ordem.some((paginaIdx, posicao)=>paginaIdx!==posicao)
}

function salvarOrdemPaginasArquivo(file, ordem){
  if(!file) return
  if(!ordemPaginasEhPersonalizada(ordem)){
    ordemPaginasPorArquivo.delete(file)
    return
  }
  ordemPaginasPorArquivo.set(file, [...ordem])
}

async function recriarPdfComOrdemAplicada(pdf, file){
  if(!file) return pdf
  const ordem=obterOrdemPaginasArquivo(file, pdf.getPageCount())
  if(!ordemPaginasEhPersonalizada(ordem)) return pdf
  const novoPdf=await PDFLib.PDFDocument.create()
  const paginas=await novoPdf.copyPages(pdf, ordem)
  paginas.forEach(pagina=>novoPdf.addPage(pagina))
  return novoPdf
}

function garantirFerramentasModalPaginas(){
  const box=document.querySelector("#modal-paginas .modal-box")
  const grid=document.getElementById("modal-paginas-grid")
  if(!box || !grid) return {}

  let toolbar=document.getElementById("modal-paginas-toolbar")
  if(!toolbar){
    toolbar=document.createElement("div")
    toolbar.id="modal-paginas-toolbar"
    toolbar.className="modal-pages-toolbar"
    toolbar.innerHTML=`
      <div id="modal-paginas-info" class="modal-pages-info"></div>
      <div class="modal-pages-actions">
        <button id="modal-reset-order" type="button" class="btn atalho btn-sm">Restaurar ordem</button>
      </div>
    `
    box.insertBefore(toolbar, grid)
  }

  const info=document.getElementById("modal-paginas-info")
  const resetBtn=document.getElementById("modal-reset-order")
  if(resetBtn){
    resetBtn.onclick=()=>{
      if(!modalPaginasEstado.file || modalPaginasEstado.total<=1) return
      modalPaginasEstado.ordem=ordemPaginasNatural(modalPaginasEstado.total)
      salvarOrdemPaginasArquivo(modalPaginasEstado.file, modalPaginasEstado.ordem)
      renderizarGridModalPaginas()
    }
  }

  return {toolbar, info, resetBtn}
}

function atualizarStatusModalPaginas(){
  const info=document.getElementById("modal-paginas-info")
  const resetBtn=document.getElementById("modal-reset-order")
  if(!info || !resetBtn) return
  const total=modalPaginasEstado.total||0
  const personalizada=ordemPaginasEhPersonalizada(modalPaginasEstado.ordem||[])
  if(total<=1){
    info.textContent="Documento com 1 pagina."
    resetBtn.disabled=true
    return
  }
  info.textContent=personalizada
    ? "Ordem personalizada ativa. Arraste as paginas para reorganizar o PDF."
    : "Arraste as paginas para reorganizar. A ordem sera salva automaticamente."
  resetBtn.disabled=!personalizada
}

function renderizarGridModalPaginas(){
  const grid=document.getElementById("modal-paginas-grid")
  if(!grid) return
  grid.innerHTML=""
  modalPaginasEstado.ordem.forEach((paginaIdx, posicao)=>{
    const item=modalPaginasEstado.itens.get(paginaIdx)
    if(!item) return
    item.dataset.modalPos=String(posicao)
    const ordemEl=item.querySelector(".modal-page-order")
    const labelEl=item.querySelector(".modal-page-label")
    if(ordemEl) ordemEl.textContent=String(posicao+1).padStart(2,"0")
    if(labelEl) labelEl.textContent="Pagina "+(paginaIdx+1)
    grid.appendChild(item)
  })
  atualizarStatusModalPaginas()
}

function resetarEstadoModalPaginas(){
  modalPaginaArrastandoPos=null
  modalPaginasEstado={file:null, ordem:[], itens:new Map(), total:0}
}

async function abrirPaginasModal(file){
  document.getElementById("modal-titulo").textContent=file.name
  const grid=document.getElementById("modal-paginas-grid")
  grid.innerHTML="<p style='color:#6b7280'>Carregando páginas...</p>"
  document.getElementById("modal-paginas").style.display="flex"
  garantirFerramentasModalPaginas()

  const arrayBuffer=await file.arrayBuffer()
  const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
  modalPaginasEstado={
    file,
    ordem:obterOrdemPaginasArquivo(file, pdf.numPages),
    itens:new Map(),
    total:pdf.numPages
  }

  for(const paginaIdx of modalPaginasEstado.ordem){
    const page=await pdf.getPage(paginaIdx+1)
    const viewport=page.getViewport({scale:1.45})
    const canvas=document.createElement("canvas")
    canvas.width=viewport.width
    canvas.height=viewport.height
    const ctx=canvas.getContext("2d")
    ctx.fillStyle="#ffffff"
    ctx.fillRect(0,0,canvas.width,canvas.height)
    await page.render({canvasContext:ctx,viewport}).promise

    const item=document.createElement("div")
    item.className="modal-page-item"
    item.dataset.pageIndex=String(paginaIdx)
    if(pdf.numPages>1){
      item.draggable=true
      item.classList.add("modal-page-item-sortable")
      item.addEventListener("dragstart",(e)=>{
        modalPaginaArrastandoPos=Number(item.dataset.modalPos||0)
        item.classList.add("modal-page-item-dragging")
        if(e.dataTransfer){
          e.dataTransfer.effectAllowed="move"
          e.dataTransfer.setData("text/plain", item.dataset.modalPos||"0")
        }
      })
      item.addEventListener("dragend",()=>{
        modalPaginaArrastandoPos=null
        item.classList.remove("modal-page-item-dragging")
        document.querySelectorAll(".modal-page-item-drop-target").forEach(el=>el.classList.remove("modal-page-item-drop-target"))
      })
      item.addEventListener("dragover",(e)=>{
        const destinoPos=Number(item.dataset.modalPos||0)
        if(modalPaginaArrastandoPos===null || modalPaginaArrastandoPos===destinoPos) return
        e.preventDefault()
        item.classList.add("modal-page-item-drop-target")
        if(e.dataTransfer) e.dataTransfer.dropEffect="move"
      })
      item.addEventListener("dragleave",()=>{
        item.classList.remove("modal-page-item-drop-target")
      })
      item.addEventListener("drop",(e)=>{
        const destinoPos=Number(item.dataset.modalPos||0)
        if(modalPaginaArrastandoPos===null || modalPaginaArrastandoPos===destinoPos) return
        e.preventDefault()
        item.classList.remove("modal-page-item-drop-target")
        const ordemAtual=[...modalPaginasEstado.ordem]
        const [movida]=ordemAtual.splice(modalPaginaArrastandoPos,1)
        ordemAtual.splice(destinoPos,0,movida)
        modalPaginasEstado.ordem=ordemAtual
        salvarOrdemPaginasArquivo(file, ordemAtual)
        renderizarGridModalPaginas()
      })
    }

    const head=document.createElement("div")
    head.className="modal-page-head"
    const ordemBadge=document.createElement("span")
    ordemBadge.className="modal-page-order"
    const label=document.createElement("span")
    label.className="modal-page-label"
    label.textContent="Pagina "+(paginaIdx+1)
    head.appendChild(ordemBadge)
    head.appendChild(label)
    if(pdf.numPages>1){
      const dragHint=document.createElement("span")
      dragHint.className="modal-page-drag-hint"
      dragHint.textContent="Arraste para mover"
      head.appendChild(dragHint)
    }
    item.appendChild(head)
    item.appendChild(canvas)
    modalPaginasEstado.itens.set(paginaIdx, item)
  }
  renderizarGridModalPaginas()
}

function fecharModal(e){
  if(e.target===document.getElementById("modal-paginas")){
    document.getElementById("modal-paginas").style.display="none"
    resetarEstadoModalPaginas()
  }
}

function fecharModalBtn(){
  document.getElementById("modal-paginas").style.display="none"
  resetarEstadoModalPaginas()
}

async function criarThumbCard(file, idx=null, options={}){
  const compacto=options.compacto===true
  const ordemTexto=options.ordemTexto ?? (idx!==null ? String(idx+1).padStart(2,"0") : "PDF")
  const classesExtras=Array.isArray(options.classesExtras) ? options.classesExtras : []
  const canvasWidth=compacto ? 126 : 148
  const canvasHeight=compacto ? 146 : 172
  const card=document.createElement("div")
  card.className="thumb-card"
  if(compacto) card.classList.add("thumb-card-compact")
  classesExtras.forEach(classe=>card.classList.add(classe))
  card.onclick=()=>{
    if(Date.now()<bloquearCliqueThumbAte) return
    abrirPaginasModal(file)
  }

  const btnRemover=document.createElement("button")
  btnRemover.className="thumb-remover"
  btnRemover.type="button"
  btnRemover.setAttribute("aria-label","Remover arquivo")
  btnRemover.textContent="✕"
  btnRemover.onclick=(e)=>{
    e.stopPropagation()
    if(idx!==null) removerArquivoJuntar(idx)
  }

  const ordemBadge=document.createElement("div")
  ordemBadge.className="thumb-order-badge"
  ordemBadge.textContent=ordemTexto

  // Placeholder imediato — sem bloquear
  const canvas=document.createElement("canvas")
  canvas.width=canvasWidth
  canvas.height=canvasHeight
  const ctx=canvas.getContext("2d")
  const grad=ctx.createLinearGradient(0,0,0,canvasHeight)
  grad.addColorStop(0,"#f8fafc")
  grad.addColorStop(1,"#e5edf8")
  ctx.fillStyle=grad
  ctx.fillRect(0,0,canvasWidth,canvasHeight)
  ctx.fillStyle="#4f46e5"
  ctx.font=compacto ? "700 36px sans-serif" : "700 44px sans-serif"
  ctx.textAlign="center"
  ctx.fillText("PDF",canvasWidth/2,compacto ? 82 : 96)
  ctx.fillStyle="#94a3b8"
  ctx.font=compacto ? "600 11px sans-serif" : "600 12px sans-serif"
  ctx.fillText("Carregando prévia",74,122)

  const nome=document.createElement("div")
  nome.className="thumb-nome"
  nome.textContent=file.name

  const pags=document.createElement("div")
  pags.className="thumb-paginas"
  pags.textContent="..."

  card.appendChild(ordemBadge)
  if(idx!==null) card.appendChild(btnRemover)
  card.appendChild(canvas)
  card.appendChild(nome)
  card.appendChild(pags)

  if(idx!==null){
    card.draggable=true
    card.classList.add("thumb-card-sortable")
    card.setAttribute("aria-label","Arraste para reorganizar o PDF")

    card.addEventListener("dragstart",(e)=>{
      arquivoJuntarArrastandoIdx=idx
      bloquearCliqueThumbAte=Date.now()+350
      card.classList.add("thumb-card-dragging")
      if(e.dataTransfer){
        e.dataTransfer.effectAllowed="move"
        e.dataTransfer.setData("text/plain", String(idx))
      }
    })

    card.addEventListener("dragend",()=>{
      arquivoJuntarArrastandoIdx=null
      bloquearCliqueThumbAte=Date.now()+350
      card.classList.remove("thumb-card-dragging")
      document.querySelectorAll(".thumb-card-drop-target").forEach(el=>el.classList.remove("thumb-card-drop-target"))
    })

    card.addEventListener("dragover",(e)=>{
      e.preventDefault()
      if(arquivoJuntarArrastandoIdx===null || arquivoJuntarArrastandoIdx===idx) return
      card.classList.add("thumb-card-drop-target")
      if(e.dataTransfer) e.dataTransfer.dropEffect="move"
    })

    card.addEventListener("dragleave",()=>{
      card.classList.remove("thumb-card-drop-target")
    })

    card.addEventListener("drop",(e)=>{
      e.preventDefault()
      e.stopPropagation()
      card.classList.remove("thumb-card-drop-target")
      if(arquivoJuntarArrastandoIdx===null || arquivoJuntarArrastandoIdx===idx) return
      reordenarArquivoJuntar(arquivoJuntarArrastandoIdx, idx)
    })
  }

  if(idx!==null){
    const mover=document.createElement("div")
    mover.className="thumb-move-controls"

    const btnSubir=document.createElement("button")
    btnSubir.type="button"
    btnSubir.className="thumb-move-btn"
    btnSubir.textContent="↑"
    btnSubir.disabled=idx===0
    btnSubir.setAttribute("aria-label","Mover arquivo para cima")
    btnSubir.onclick=(e)=>{
      e.stopPropagation()
      moverArquivoJuntar(idx,-1)
    }

    const btnDescer=document.createElement("button")
    btnDescer.type="button"
    btnDescer.className="thumb-move-btn"
    btnDescer.textContent="↓"
    btnDescer.disabled=idx===arquivos.length-1
    btnDescer.setAttribute("aria-label","Mover arquivo para baixo")
    btnDescer.onclick=(e)=>{
      e.stopPropagation()
      moverArquivoJuntar(idx,1)
    }

    mover.appendChild(btnSubir)
    mover.appendChild(btnDescer)
    card.appendChild(mover)
  }

  // Renderizar capa em background sem bloquear o loop
  let totalPaginas=1
  renderizarCapaAsync(file, canvas, pags).then(n=>{ totalPaginas=n })

  return {card, totalPaginas}
}

async function renderizarCapaAsync(file, canvas, pagsEl){
  try{
    const arrayBuffer=await file.arrayBuffer()
    const pdf=await pdfjsLib.getDocument({data:arrayBuffer}).promise
    const page=await pdf.getPage(1)
    const viewport=page.getViewport({scale:0.8})
    canvas.width=viewport.width
    canvas.height=viewport.height
    const ctx=canvas.getContext("2d")
    ctx.fillStyle="#ffffff"
    ctx.fillRect(0,0,canvas.width,canvas.height)
    await page.render({canvasContext:ctx,viewport}).promise
    pagsEl.textContent=pdf.numPages+" página"+(pdf.numPages>1?"s":"")
    return pdf.numPages
  }catch(e){
    pagsEl.textContent="—"
    return 1
  }
}


// ===================== JUNTAR PDF =====================

let arquivos=[]
let pastasJuntar=[]
let resultadosPastasJuntar=[]
let arquivoJuntarArrastandoIdx=null
let bloquearCliqueThumbAte=0
let pastaJuntarArquivoArrastandoInfo=null

const fileInput=document.getElementById("fileInput")
const gallery=document.getElementById("file-gallery")


function arquivoEhCompactado(file){
  const nome=(file?.name||"").toLowerCase()
  const tipo=(file?.type||"").toLowerCase()
  return (
    /\.(zip|rar|7z|tar|tgz|gz|bz2|xz)$/i.test(nome) ||
    nome.endsWith(".tar.gz") ||
    nome.endsWith(".tar.bz2") ||
    nome.endsWith(".tar.xz") ||
    tipo==="application/zip" ||
    tipo==="application/x-zip-compressed" ||
    tipo==="application/x-rar-compressed" ||
    tipo==="application/vnd.rar" ||
    tipo==="application/x-7z-compressed" ||
    tipo==="application/gzip"
  )
}

function arquivoEhZip(file){
  return /\.zip$/i.test(file?.name||"") || file?.type==="application/zip" || file?.type==="application/x-zip-compressed"
}

function obterTipoPorNome(nome){
  const n=(nome||"").toLowerCase()
  if(n.endsWith(".pdf")) return "application/pdf"
  if(n.endsWith(".png")) return "image/png"
  if(n.endsWith(".jpg")||n.endsWith(".jpeg")) return "image/jpeg"
  if(n.endsWith(".webp")) return "image/webp"
  return ""
}

function arquivoValidoJuntarPorNome(nome){
  return /\.(pdf|png|jpe?g|webp)$/i.test(nome||"")
}

function nomeBaseSemExtensao(nome){
  return (nome||"arquivo").replace(/\.[^.]+$/,"")
}

function rerenderPastasJuntar(){
  const row=document.getElementById("pastas-adicionadas")
  if(!row) return
  row.innerHTML=""
  for(let i=0;i<pastasJuntar.length;i++) renderizarPastaThumb(i)
}

function adicionarOuSubstituirPastaJuntar(nomePasta, arquivosPasta){
  const idxExistente=pastasJuntar.findIndex(p=>p.nomePasta===nomePasta && p.arquivos.length===arquivosPasta.length)
  if(idxExistente>=0){
    const desejaSubstituir=confirm('Já consta uma pasta com o mesmo nome ("'+nomePasta+'") e a mesma quantidade de documentos ('+arquivosPasta.length+').\\n\\nClique em OK para substituir ou em Cancelar para manter a atual.')
    if(!desejaSubstituir) return {adicionada:false, substituida:false}
    pastasJuntar[idxExistente]={
      nomePasta,
      arquivos:ordenarArquivosPastaInicialmente(arquivosPasta),
      expandida:pastasJuntar[idxExistente]?.expandida===true
    }
    rerenderPastasJuntar()
    atualizarInfoJuntar()
    return {adicionada:true, substituida:true}
  }
  pastasJuntar.push({
    nomePasta,
    arquivos:ordenarArquivosPastaInicialmente(arquivosPasta),
    expandida:false
  })
  renderizarPastaThumb(pastasJuntar.length-1)
  atualizarInfoJuntar()
  return {adicionada:true, substituida:false}
}

async function extrairPastasDeZipParaJuntar(zipFile){
  const JSZip=await carregarJSZip()
  const zip=await JSZip.loadAsync(await zipFile.arrayBuffer())
  const grupos=new Map()
  const zipBase=nomeBaseSemExtensao(zipFile.name)

  const entradas=Object.values(zip.files)
  for(const entry of entradas){
    if(entry.dir) continue
    if(!arquivoValidoJuntarPorNome(entry.name)) continue

    const partes=entry.name.split("/").filter(Boolean)
    if(partes.length===0) continue

    let nomePasta=zipBase
    let caminhoInterno=entry.name

    if(partes.length>1){
      nomePasta=partes[0]
      caminhoInterno=partes.slice(1).join("/")
    }

    const tipo=obterTipoPorNome(entry.name)
    const bytes=await entry.async("uint8array")
    const arquivo=new File([bytes], partes[partes.length-1], {type:tipo||"application/octet-stream"})
    try{
      Object.defineProperty(arquivo, "webkitRelativePath", {
        value: nomePasta + "/" + caminhoInterno,
        configurable: true
      })
    }catch(e){}
    if(!grupos.has(nomePasta)) grupos.set(nomePasta, [])
    grupos.get(nomePasta).push(arquivo)
  }

  return Array.from(grupos.entries()).map(([nomePasta, arquivos])=>({nomePasta, arquivos}))
}

let libarchiveRef=null
let libarchiveInitPromise=null

async function carregarLibarchive(){
  if(libarchiveRef) return libarchiveRef
  if(!libarchiveInitPromise){
    libarchiveInitPromise=(async()=>{
      const mod=await import("https://cdn.jsdelivr.net/npm/libarchive.js@2.0.2/dist/libarchive.js")
      const Archive=mod.Archive || mod.default?.Archive || mod.default || mod
      if(Archive?.init){
        Archive.init({
          workerUrl:"https://cdn.jsdelivr.net/npm/libarchive.js@2.0.2/dist/worker-bundle.js"
        })
      }
      libarchiveRef=Archive
      return Archive
    })()
  }
  return await libarchiveInitPromise
}

async function extrairPastasDeCompactadoParaJuntar(archiveFile){
  if(arquivoEhZip(archiveFile)){
    return await extrairPastasDeZipParaJuntar(archiveFile)
  }

  const Archive=await carregarLibarchive()
  const archive=await Archive.open(archiveFile)
  const grupos=new Map()
  const archiveBase=nomeBaseSemExtensao(archiveFile.name)

  await archive.extractFiles((entry)=>{
    const fileExtraido=entry?.file
    if(!fileExtraido) return

    const caminhoBase=String(entry?.path||"").replace(/\\/g,"/").replace(/^\/+/,"")
    const caminhoCompleto=(caminhoBase ? caminhoBase+fileExtraido.name : fileExtraido.name).replace(/\\/g,"/")
    if(!arquivoValidoJuntarPorNome(caminhoCompleto)) return

    const partes=caminhoCompleto.split("/").filter(Boolean)
    if(partes.length===0) return

    let nomePasta=archiveBase
    let caminhoInterno=caminhoCompleto
    if(partes.length>1){
      nomePasta=partes[0]
      caminhoInterno=partes.slice(1).join("/")
    }

    const tipo=obterTipoPorNome(fileExtraido.name)
    const arquivo=new File([fileExtraido], partes[partes.length-1], {type:tipo||fileExtraido.type||"application/octet-stream"})
    try{
      Object.defineProperty(arquivo, "webkitRelativePath", {
        value: nomePasta + "/" + caminhoInterno,
        configurable: true
      })
    }catch(e){}
    if(!grupos.has(nomePasta)) grupos.set(nomePasta, [])
    grupos.get(nomePasta).push(arquivo)
  })

  return Array.from(grupos.entries()).map(([nomePasta, arquivos])=>({nomePasta, arquivos}))
}

async function processarZipNoJuntar(zipFile){
  mostrarLoading("Lendo ZIP: "+zipFile.name+"...")
  const grupos=await extrairPastasDeZipParaJuntar(zipFile)
  if(grupos.length===0){
    esconderLoading()
    alert("Nenhum PDF ou imagem válido foi encontrado no ZIP.")
    return
  }

  let adicionadas=0
  let substituidas=0
  for(const grupo of grupos){
    const r=adicionarOuSubstituirPastaJuntar(grupo.nomePasta, grupo.arquivos)
    if(r.adicionada) adicionadas++
    if(r.substituida) substituidas++
  }

  esconderLoading()
  if(substituidas>0){
    mostrarSucesso("ZIP carregado com substituição de pasta(s).")
  }else{
    mostrarSucesso(adicionadas+" pasta"+(adicionadas>1?"s adicionadas":" adicionada")+" via ZIP!")
  }
}

async function processarCompactadoNoJuntar(archiveFile){
  if(arquivoEhZip(archiveFile)){
    await processarZipNoJuntar(archiveFile)
    return
  }

  mostrarLoading("Lendo compactado: "+archiveFile.name+"...")
  let grupos=[]
  try{
    grupos=await extrairPastasDeCompactadoParaJuntar(archiveFile)
  }catch(e){
    console.error(e)
    esconderLoading()
    alert("Não foi possível abrir esse compactado. Verifique se ele não está protegido por senha ou corrompido.")
    return
  }
  if(grupos.length===0){
    esconderLoading()
    alert("Nenhum PDF ou imagem válido foi encontrado no compactado.")
    return
  }

  let adicionadas=0
  let substituidas=0
  for(const grupo of grupos){
    const r=adicionarOuSubstituirPastaJuntar(grupo.nomePasta, grupo.arquivos)
    if(r.adicionada) adicionadas++
    if(r.substituida) substituidas++
  }

  esconderLoading()
  if(substituidas>0){
    mostrarSucesso("Compactado carregado com substituição de pasta(s).")
  }else{
    mostrarSucesso(adicionadas+" pasta"+(adicionadas>1?"s adicionadas":" adicionada")+" via compactado!")
  }
}

async function extrairCompactadosParaPastaPrincipal(compactados, nomePastaPrincipal){
  const arquivosExtraidos=[]
  let processados=0
  let ignorados=0

  for(const compactado of compactados){
    try{
      const grupos=await extrairPastasDeCompactadoParaJuntar(compactado)
      if(grupos.length===0){
        ignorados++
        continue
      }

      processados++
      const baseCompactado=nomeBaseSemExtensao(compactado.name)

      for(const grupo of grupos){
        for(const arquivo of grupo.arquivos){
          const caminhoOriginal=(arquivo.webkitRelativePath || `${grupo.nomePasta}/${arquivo.name}`)
            .replace(/\\/g,"/")
            .replace(/^\/+/,"")
          const partes=caminhoOriginal.split("/").filter(Boolean)
          const caminhoInterno=partes.length>1 ? partes.slice(1).join("/") : arquivo.name
          const caminhoFinal=`${nomePastaPrincipal}/${baseCompactado}/${caminhoInterno}`
          try{
            Object.defineProperty(arquivo, "webkitRelativePath", {
              value: caminhoFinal,
              configurable: true
            })
          }catch(e){}
          arquivosExtraidos.push(arquivo)
        }
      }
    }catch(err){
      console.warn("Nao foi possivel ler compactado da pasta:", compactado?.name, err)
      ignorados++
    }
  }

  return {arquivosExtraidos, processados, ignorados}
}

async function processarEntradaJuntar(files){
  const compactados=files.filter(arquivoEhCompactado)
  const comuns=files.filter(f=>!arquivoEhCompactado(f))
  if(comuns.length>0) await carregarFilesNoJuntar(comuns)
  for(const compactado of compactados){
    await processarCompactadoNoJuntar(compactado)
  }
}

fileInput.addEventListener("change",async(e)=>{
  await processarEntradaJuntar(Array.from(e.target.files))
  fileInput.value=""
})

async function renderizarArquivosJuntar(){
  gallery.innerHTML=""
  for(let i=0;i<arquivos.length;i++){
    const {card}=await criarThumbCard(arquivos[i], i)
    gallery.appendChild(card)
  }
}

function removerArquivoJuntar(idx){
  arquivos.splice(idx,1)
  renderizarArquivosJuntar()
  atualizarInfoJuntar()
}

function moverArquivoJuntar(idx, deslocamento){
  const novoIdx=idx+deslocamento
  if(novoIdx<0 || novoIdx>=arquivos.length) return
  const [arquivo]=arquivos.splice(idx,1)
  arquivos.splice(novoIdx,0,arquivo)
  renderizarArquivosJuntar()
  atualizarInfoJuntar()
}

function reordenarArquivoJuntar(origemIdx, destinoIdx){
  if(origemIdx===destinoIdx) return
  if(origemIdx<0 || destinoIdx<0) return
  if(origemIdx>=arquivos.length || destinoIdx>=arquivos.length) return
  const [arquivo]=arquivos.splice(origemIdx,1)
  arquivos.splice(destinoIdx,0,arquivo)
  renderizarArquivosJuntar()
  atualizarInfoJuntar()
}

function obterChaveOrdenacaoPasta(file){
  return (file?.webkitRelativePath || file?.name || "").toLowerCase()
}

function ordenarArquivosPastaInicialmente(listaArquivos){
  return [...listaArquivos].sort((a,b)=>obterChaveOrdenacaoPasta(a).localeCompare(obterChaveOrdenacaoPasta(b), "pt-BR", {numeric:true, sensitivity:"base"}))
}

// ===================== DRAG AND DROP UNIFICADO =====================

const DROP_ZONES={
  "#drop-zone":           "juntar",
  "#compactar-drop-zone": "compactar",
  "#split-drop-zone":     "dividir",
  "#ocr-drop-zone":       "ocr"
}

function getZona(el){
  for(let [sel,tipo] of Object.entries(DROP_ZONES)){
    if(el.closest(sel)) return {zone:el.closest(sel), tipo}
  }
  return null
}

document.addEventListener("dragenter",(e)=>{
  const r=getZona(e.target)
  if(r){e.preventDefault();r.zone.classList.add("drag-over")}
})

document.addEventListener("dragover",(e)=>{
  const r=getZona(e.target)
  if(r){e.preventDefault();r.zone.classList.add("drag-over")}
})

document.addEventListener("dragleave",(e)=>{
  const r=getZona(e.target)
  if(r&&!r.zone.contains(e.relatedTarget)) r.zone.classList.remove("drag-over")
})

document.addEventListener("drop",async(e)=>{
  const r=getZona(e.target)
  if(!r) return
  e.preventDefault()
  e.stopPropagation()
  r.zone.classList.remove("drag-over")

  const {tipo}=r
  const items=Array.from(e.dataTransfer.items||[])

  // Processar via FileSystemEntry (detecta pastas)
  if(items.length>0 && items[0].webkitGetAsEntry){
    const entries=items.map(i=>i.webkitGetAsEntry()).filter(Boolean)
    for(let entry of entries){
      if(entry.isDirectory){
        mostrarLoading("Lendo pasta: "+entry.name+"...")
        const todos=await lerDiretorio(entry)
        if(tipo==="juntar"){
          const validos=todos.filter(f=>f.type==="application/pdf"||f.type.startsWith("image/"))
          const compactados=todos.filter(arquivoEhCompactado)
          const {arquivosExtraidos}=await extrairCompactadosParaPastaPrincipal(compactados, entry.name)
          const arquivosDaPasta=[...validos, ...arquivosExtraidos]
          if(arquivosDaPasta.length===0){esconderLoading();alert("Nenhum PDF encontrado: "+entry.name);continue}
          adicionarOuSubstituirPastaJuntar(entry.name, arquivosDaPasta)
        }else if(tipo==="compactar"){
          const pdfs=todos.filter(f=>f.type==="application/pdf")
          if(pdfs.length>0) await carregarFilesNoCompactar(pdfs)
        }else if(tipo==="dividir"){
          const pdf=todos.find(f=>f.type==="application/pdf")
          if(pdf) await carregarFileNoDividir(pdf)
        }else if(tipo==="ocr"){
          const pdfs=todos.filter(f=>f.type==="application/pdf"||/\.pdf$/i.test(f.name))
          if(pdfs.length>0) await carregarFilesNoOCR(pdfs)
        }
      }else{
        const file=await entryParaFile(entry)
        if(!file) continue
        if(tipo==="juntar"){
          if(arquivoEhCompactado(file)) await processarCompactadoNoJuntar(file)
          else if(file.type==="application/pdf"||file.type.startsWith("image/")) await carregarFilesNoJuntar([file])
        }
        else if(tipo==="compactar"&&file.type==="application/pdf")
          await carregarFilesNoCompactar([file])
        else if(tipo==="dividir"&&file.type==="application/pdf")
          await carregarFileNoDividir(file)
        else if(tipo==="ocr"&&(file.type==="application/pdf"||/\.pdf$/i.test(file.name)))
          await carregarFilesNoOCR([file])
      }
    }
    if(tipo==="juntar") atualizarInfoJuntar()
    esconderLoading()
    mostrarSucesso("Carregado!")
    return
  }

  // Fallback sem FileSystemEntry
  const files=Array.from(e.dataTransfer.files||[])
  if(tipo==="juntar"){
    await processarEntradaJuntar(files)
  }else if(tipo==="compactar"){
    const pdfs=files.filter(f=>f.type==="application/pdf")
    if(pdfs.length>0) await carregarFilesNoCompactar(pdfs)
  }else if(tipo==="dividir"){
    const pdf=files.find(f=>f.type==="application/pdf")
    if(pdf) await carregarFileNoDividir(pdf)
  }
})

// Lê recursivamente todos os arquivos de um diretório
async function lerDiretorio(dirEntry){
  const result=[]
  const reader=dirEntry.createReader()

  // readEntries retorna no máximo 100 por vez — precisa chamar até retornar array vazio
  async function lerTodosOsLotes(){
    return new Promise((resolve,reject)=>{
      reader.readEntries(resolve, reject)
    })
  }

  let lote
  do{
    lote=await lerTodosOsLotes()
    for(let entry of lote){
      if(entry.isFile){
        const file=await entryParaFile(entry)
        if(file) result.push(file)
      }else if(entry.isDirectory){
        const sub=await lerDiretorio(entry)
        result.push(...sub)
      }
    }
  }while(lote.length>0)

  return result
}

function entryParaFile(entry){
  return new Promise((resolve)=>{
    entry.file(f=>resolve(f),()=>resolve(null))
  })
}

function selecionarPasta(){
  const input=document.createElement("input")
  input.type="file"
  input.webkitdirectory=true
  input.style.display="none"
  document.body.appendChild(input)
  input.addEventListener("change",async(e)=>{
    document.body.removeChild(input)
    const selecionados=Array.from(e.target.files||[])
    const files=selecionados.filter(f=>f.type==="application/pdf"||f.type.startsWith("image/"))
    const compactados=selecionados.filter(arquivoEhCompactado)
    if(files.length===0 && compactados.length===0){alert("Nenhum PDF, imagem ou compactado encontrado na pasta.");return}
    mostrarLoading("Carregando pasta selecionada...")
    const nomePasta=(selecionados[0]?.webkitRelativePath||"").split("/")[0]||("Pasta "+(pastasJuntar.length+1))
    const {arquivosExtraidos, processados:compactadosProcessados, ignorados:compactadosIgnorados}=await extrairCompactadosParaPastaPrincipal(compactados, nomePasta)
    const arquivosDaPasta=[...files, ...arquivosExtraidos]
    if(arquivosDaPasta.length===0){
      esconderLoading()
      alert("Nenhum PDF valido foi encontrado dentro da pasta ou dos compactados.")
      return
    }
    let resultado={adicionada:false, substituida:false}
    if(arquivosDaPasta.length>0){
      resultado=adicionarOuSubstituirPastaJuntar(nomePasta, arquivosDaPasta)
    }
    esconderLoading()
    if(resultado.adicionada || compactadosProcessados>0){
      if(compactadosProcessados>0 || compactadosIgnorados>0){
        const resumoIgnorados=compactadosIgnorados>0 ? " "+compactadosIgnorados+" compactado(s) ignorado(s)." : ""
        mostrarSucesso('Pasta "'+nomePasta+'" carregada com '+compactadosProcessados+' compactado(s) integrado(s).'+resumoIgnorados)
      }else{
        mostrarSucesso(resultado.substituida ? 'Pasta "'+nomePasta+'" substituída!' : 'Pasta "'+nomePasta+'" carregada!')
      }
    }
  })
  input.click()
}

function renderizarPastaThumb(idx){
  const {nomePasta, arquivos:arqs, expandida}=pastasJuntar[idx]
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
  thumbsArea.style.display=expandida ? "flex" : "none"
  thumbsArea.id="pasta-thumbs-exp-"+idx

  card.appendChild(btnRemover)
  card.appendChild(icon)
  card.appendChild(nome)
  card.appendChild(qtd)
  card.appendChild(thumbsArea)

  if(expandida){
    card.classList.add("expandida")
    renderizarArquivosDaPastaJuntarExpandida(idx, thumbsArea)
  }

  card.onclick=()=>{
    alternarPastaJuntarExpandida(idx)
  }

  row.appendChild(card)
}

async function renderizarArquivosDaPastaJuntarExpandida(idx, area){
  const pasta=pastasJuntar[idx]
  if(!pasta || !area) return
  area.innerHTML=""
  for(let i=0;i<pasta.arquivos.length;i++){
    try{
      const {card:thumb}=await criarThumbCard(pasta.arquivos[i])
      thumb.classList.add("thumb-card-pasta-expandida")
      thumb.addEventListener("click",(e)=>e.stopPropagation())
      const canvas=thumb.querySelector("canvas")
      if(canvas) canvas.classList.add("thumb-canvas-pasta-expandida")
      const badge=thumb.querySelector(".thumb-order-badge")
      if(badge) badge.textContent=String(i+1).padStart(2,"0")
      habilitarOrdenacaoArquivoNaPastaJuntar(thumb, idx, i)
      area.appendChild(thumb)
    }catch(e){}
  }
}

function alternarPastaJuntarExpandida(idx){
  const pasta=pastasJuntar[idx]
  if(!pasta) return
  pasta.expandida=!pasta.expandida
  rerenderPastasJuntar()
}

function reordenarArquivoNaPastaJuntar(pastaIdx, origemIdx, destinoIdx){
  const pasta=pastasJuntar[pastaIdx]
  if(!pasta) return
  if(origemIdx===destinoIdx) return
  if(origemIdx<0 || destinoIdx<0) return
  if(origemIdx>=pasta.arquivos.length || destinoIdx>=pasta.arquivos.length) return
  const [arquivo]=pasta.arquivos.splice(origemIdx,1)
  pasta.arquivos.splice(destinoIdx,0,arquivo)
  pasta.expandida=true
  rerenderPastasJuntar()
}

function habilitarOrdenacaoArquivoNaPastaJuntar(card, pastaIdx, arquivoIdx){
  card.draggable=true
  card.classList.add("thumb-card-sortable")
  card.setAttribute("aria-label","Arraste para reorganizar o PDF dentro da pasta")

  card.addEventListener("dragstart",(e)=>{
    pastaJuntarArquivoArrastandoInfo={pastaIdx, arquivoIdx}
    bloquearCliqueThumbAte=Date.now()+350
    card.classList.add("thumb-card-dragging")
    if(e.dataTransfer){
      e.dataTransfer.effectAllowed="move"
      e.dataTransfer.setData("text/plain", `${pastaIdx}:${arquivoIdx}`)
    }
  })

  card.addEventListener("dragend",()=>{
    pastaJuntarArquivoArrastandoInfo=null
    bloquearCliqueThumbAte=Date.now()+350
    card.classList.remove("thumb-card-dragging")
    document.querySelectorAll(".thumb-card-drop-target").forEach(el=>el.classList.remove("thumb-card-drop-target"))
  })

  card.addEventListener("dragover",(e)=>{
    if(!pastaJuntarArquivoArrastandoInfo) return
    if(pastaJuntarArquivoArrastandoInfo.pastaIdx!==pastaIdx) return
    if(pastaJuntarArquivoArrastandoInfo.arquivoIdx===arquivoIdx) return
    e.preventDefault()
    card.classList.add("thumb-card-drop-target")
    if(e.dataTransfer) e.dataTransfer.dropEffect="move"
  })

  card.addEventListener("dragleave",()=>{
    card.classList.remove("thumb-card-drop-target")
  })

  card.addEventListener("drop",(e)=>{
    if(!pastaJuntarArquivoArrastandoInfo) return
    if(pastaJuntarArquivoArrastandoInfo.pastaIdx!==pastaIdx) return
    if(pastaJuntarArquivoArrastandoInfo.arquivoIdx===arquivoIdx) return
    e.preventDefault()
    e.stopPropagation()
    card.classList.remove("thumb-card-drop-target")
    reordenarArquivoNaPastaJuntar(pastaIdx, pastaJuntarArquivoArrastandoInfo.arquivoIdx, arquivoIdx)
  })
}

function removerPastaJuntar(idx){
  pastasJuntar.splice(idx,1)
  document.getElementById("pastas-adicionadas").innerHTML=""
  for(let i=0;i<pastasJuntar.length;i++) renderizarPastaThumb(i)
  atualizarInfoJuntar()
}

// ===================== TOAST DE CARREGAMENTO =====================

let toastTimer=null

function mostrarLoading(msg){
  let toast=document.getElementById("toast-loading")
  if(!toast){
    toast=document.createElement("div")
    toast.id="toast-loading"
    document.body.appendChild(toast)
  }
  toast.innerHTML=`<div class="toast-spinner"></div><span>${msg}</span>`
  toast.classList.add("visivel")
  if(toastTimer) clearTimeout(toastTimer)
}

function esconderLoading(){
  const toast=document.getElementById("toast-loading")
  if(toast) toast.classList.remove("visivel")
}

function mostrarSucesso(msg){
  let toast=document.getElementById("toast-loading")
  if(!toast){
    toast=document.createElement("div")
    toast.id="toast-loading"
    document.body.appendChild(toast)
  }
  toast.innerHTML=`<span>✅ ${msg}</span>`
  toast.classList.add("visivel","sucesso")
  if(toastTimer) clearTimeout(toastTimer)
  toastTimer=setTimeout(()=>{
    toast.classList.remove("visivel","sucesso")
  },2500)
}

function atualizarProgressoJuntarAntigo(percentual, titulo="", subtitulo=""){
  const card=document.getElementById("juntar-progress-card")
  const fill=document.getElementById("juntar-progress-fill")
  const label=document.getElementById("juntar-progress-label")
  const percent=document.getElementById("juntar-progress-percent")
  const sub=document.getElementById("juntar-progress-subtext")
  if(!card || !fill || !label || !percent || !sub) return
  card.style.display="block"
  const seguro=Math.max(0,Math.min(100,Math.round(percentual||0)))
  fill.style.width=seguro+"%"
  percent.textContent=seguro+"%"
  if(titulo) label.textContent=titulo
  if(subtitulo!==undefined) sub.textContent=subtitulo
}

function ocultarProgressoJuntarAntigo(){
  const card=document.getElementById("juntar-progress-card")
  const fill=document.getElementById("juntar-progress-fill")
  const label=document.getElementById("juntar-progress-label")
  const percent=document.getElementById("juntar-progress-percent")
  const sub=document.getElementById("juntar-progress-subtext")
  if(card) card.style.display="block"
  if(fill) fill.style.width="0%"
  if(label) label.textContent="Aguardando arquivos"
  if(percent) percent.textContent="0%"
  if(sub) sub.textContent="A barra de progresso aparece aqui durante a geração do PDF."
}
function garantirTelaProcesso(secaoId, telaId, config){
  if(document.getElementById(telaId)) return
  const secao=document.getElementById(secaoId)
  const barraFixa=secao?.querySelector(".fixed-bar")
  if(!secao || !barraFixa) return

  const tela=document.createElement("div")
  tela.id=telaId
  tela.className="process-screen"
  tela.style.display="none"
  tela.innerHTML=`
    <div class="process-screen-card">
      <button id="${telaId}-voltar" type="button" class="process-back-btn">Voltar</button>
      <div id="${telaId}-progress" class="process-progress-block">
        <span class="process-kicker">${config.kicker}</span>
        <h3 id="${telaId}-titulo" class="process-title">${config.tituloInicial}</h3>
        <p id="${telaId}-subtitulo" class="process-subtitle">${config.subtituloInicial}</p>
        <div class="progress-head progress-head-centered">
          <strong>Progresso</strong>
          <span id="${telaId}-percent">0%</span>
        </div>
        <div class="progress-track">
          <div id="${telaId}-fill" class="progress-fill"></div>
        </div>
      </div>
      <div id="${telaId}-resultado" class="process-result-block" style="display:none">
        <span class="process-kicker">${config.kickerResultado}</span>
        <h3 id="${telaId}-resultado-titulo" class="process-title">${config.tituloResultado}</h3>
        <p id="${telaId}-resultado-texto" class="process-subtitle">${config.textoResultado}</p>
        <button id="${telaId}-download" class="btn generate">${config.botaoResultado}</button>
      </div>
    </div>
  `

  secao.insertBefore(tela, barraFixa)
  const btnVoltar=tela.querySelector(`#${telaId}-voltar`)
  if(btnVoltar){
    btnVoltar.onclick=()=>fecharTelaProcesso(secaoId, telaId)
  }
}

function garantirTelasDeProcesso(){
  garantirTelaProcesso("unificador","juntar-processo",{
    kicker:"Processando arquivos",
    tituloInicial:"Preparando arquivos...",
    subtituloInicial:"Aguarde enquanto o PDF esta sendo gerado.",
    kickerResultado:"Arquivo pronto",
    tituloResultado:"Processamento concluido",
    textoResultado:"Seu arquivo ja pode ser baixado.",
    botaoResultado:"Baixar PDF"
  })

  garantirTelaProcesso("dividir-pdf","dividir-processo",{
    kicker:"Dividindo arquivo",
    tituloInicial:"Preparando divisao...",
    subtituloInicial:"Aguarde enquanto os arquivos sao separados.",
    kickerResultado:"Arquivos prontos",
    tituloResultado:"Divisao concluida",
    textoResultado:"Os arquivos gerados ja podem ser baixados.",
    botaoResultado:"Baixar arquivos"
  })
}

function mostrarTelaProcesso(secaoId, telaId){
  const secao=document.getElementById(secaoId)
  const tela=document.getElementById(telaId)
  if(!secao || !tela) return
  secao.classList.add("process-mode")
  tela.style.display="flex"
}

function fecharTelaProcesso(secaoId, telaId){
  const secao=document.getElementById(secaoId)
  const tela=document.getElementById(telaId)
  if(secao) secao.classList.remove("process-mode")
  if(tela) tela.style.display="none"
}

function ocultarTelaProcesso(secaoId, telaId, tituloPadrao, subtituloPadrao){
  const secao=document.getElementById(secaoId)
  const tela=document.getElementById(telaId)
  const fill=document.getElementById(`${telaId}-fill`)
  const percent=document.getElementById(`${telaId}-percent`)
  const titulo=document.getElementById(`${telaId}-titulo`)
  const subtitulo=document.getElementById(`${telaId}-subtitulo`)
  const progress=document.getElementById(`${telaId}-progress`)
  const resultado=document.getElementById(`${telaId}-resultado`)
  resetarAnimacaoTelaProcesso(telaId)
  if(secao) secao.classList.remove("process-mode")
  if(tela) tela.style.display="none"
  if(fill) fill.style.width="0%"
  if(percent) percent.textContent="0%"
  if(titulo) titulo.textContent=tituloPadrao
  if(subtitulo) subtitulo.textContent=subtituloPadrao
  if(progress) progress.style.display="block"
  if(resultado) resultado.style.display="none"
}

const estadosTelaProcesso={}

function esperar(ms){
  return new Promise(resolve=>setTimeout(resolve, ms))
}

function obterEstadoTelaProcesso(telaId){
  if(!estadosTelaProcesso[telaId]){
    estadosTelaProcesso[telaId]={
      atual: 0,
      alvo: 0,
      raf: null
    }
  }
  return estadosTelaProcesso[telaId]
}

function resetarAnimacaoTelaProcesso(telaId){
  const estado=estadosTelaProcesso[telaId]
  if(estado?.raf) cancelAnimationFrame(estado.raf)
  estadosTelaProcesso[telaId]={
    atual: 0,
    alvo: 0,
    raf: null
  }
}

function definirPercentualTelaProcesso(telaId, percentual, instantaneo=false){
  const fill=document.getElementById(`${telaId}-fill`)
  const percent=document.getElementById(`${telaId}-percent`)
  if(!fill || !percent) return Promise.resolve()

  const destino=Math.max(0,Math.min(100,Number(percentual)||0))
  const estado=obterEstadoTelaProcesso(telaId)

  if(estado.raf){
    cancelAnimationFrame(estado.raf)
    estado.raf=null
  }

  if(instantaneo){
    estado.atual=destino
    estado.alvo=destino
    fill.style.width=destino.toFixed(1)+"%"
    percent.textContent=Math.round(destino)+"%"
    return Promise.resolve()
  }

  const inicio=estado.atual
  const delta=destino-inicio
  if(Math.abs(delta)<0.2){
    estado.atual=destino
    estado.alvo=destino
    fill.style.width=destino.toFixed(1)+"%"
    percent.textContent=Math.round(destino)+"%"
    return Promise.resolve()
  }

  const duracao=Math.max(420, Math.min(950, Math.abs(delta)*18))
  estado.alvo=destino

  return new Promise(resolve=>{
    const inicioTempo=performance.now()
    const animar=(agora)=>{
      const progresso=Math.min(1,(agora-inicioTempo)/duracao)
      const eased=1-Math.pow(1-progresso,3)
      const valor=inicio+(delta*eased)
      estado.atual=valor
      fill.style.width=valor.toFixed(1)+"%"
      percent.textContent=Math.round(valor)+"%"
      if(progresso<1){
        estado.raf=requestAnimationFrame(animar)
      }else{
        estado.atual=destino
        estado.alvo=destino
        estado.raf=null
        fill.style.width=destino.toFixed(1)+"%"
        percent.textContent=Math.round(destino)+"%"
        resolve()
      }
    }
    estado.raf=requestAnimationFrame(animar)
  })
}

function atualizarTelaProcesso(secaoId, telaId, percentual, titulo="", subtitulo=""){
  mostrarTelaProcesso(secaoId, telaId)
  const progress=document.getElementById(`${telaId}-progress`)
  const resultado=document.getElementById(`${telaId}-resultado`)
  const label=document.getElementById(`${telaId}-titulo`)
  const sub=document.getElementById(`${telaId}-subtitulo`)
  if(!progress || !resultado || !label || !sub) return
  progress.style.display="block"
  resultado.style.display="none"
  const seguro=Math.max(0,Math.min(100,Math.round(percentual||0)))
  definirPercentualTelaProcesso(telaId, seguro)
  if(titulo) label.textContent=titulo
  if(subtitulo!==undefined) sub.textContent=subtitulo
}

async function finalizarTelaProcesso(secaoId, telaId, titulo, subtitulo, botaoTexto, onDownload){
  mostrarTelaProcesso(secaoId, telaId)
  const progress=document.getElementById(`${telaId}-progress`)
  const resultado=document.getElementById(`${telaId}-resultado`)
  const tituloEl=document.getElementById(`${telaId}-resultado-titulo`)
  const textoEl=document.getElementById(`${telaId}-resultado-texto`)
  const btn=document.getElementById(`${telaId}-download`)
  await definirPercentualTelaProcesso(telaId, 100)
  await esperar(320)
  if(progress) progress.style.display="none"
  if(resultado) resultado.style.display="flex"
  if(tituloEl) tituloEl.textContent=titulo
  if(textoEl) textoEl.textContent=subtitulo
  if(btn){
    btn.textContent=botaoTexto
    btn.onclick=onDownload || null
  }
}

function atualizarProgressoJuntar(percentual, titulo="", subtitulo=""){
  atualizarTelaProcesso("unificador","juntar-processo",percentual,titulo,subtitulo)
}

function ocultarProgressoJuntar(){
  ocultarTelaProcesso(
    "unificador",
    "juntar-processo",
    "Preparando arquivos...",
    "Aguarde enquanto o PDF esta sendo gerado."
  )
}

function atualizarProgressoDividir(percentual, titulo="", subtitulo=""){
  atualizarTelaProcesso("dividir-pdf","dividir-processo",percentual,titulo,subtitulo)
}

function ocultarProgressoDividir(){
  ocultarTelaProcesso(
    "dividir-pdf",
    "dividir-processo",
    "Preparando divisao...",
    "Aguarde enquanto os arquivos sao separados."
  )
}

garantirTelasDeProcesso()

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
  mostrarLoading("Carregando "+files.length+" arquivo"+(files.length>1?"s":"")+"...")
  for(let file of files){
    let fileFinal=file
    if(file.type!=="application/pdf"){
      mostrarLoading("Convertendo imagem: "+file.name+"...")
      try{ fileFinal=await imagemParaPDF(file) }
      catch(e){ esconderLoading(); alert("Não foi possível converter: "+file.name); continue }
    }
    arquivos.push(fileFinal)
  }
  await renderizarArquivosJuntar()
  esconderLoading()
  mostrarSucesso(files.length+" arquivo"+(files.length>1?"s carregados":"carregado")+"!")
  atualizarInfoJuntar()
}

function atualizarInfoJuntar(){
  const ta=arquivos.length
  const tp=pastasJuntar.length
  let info=""
  if(tp>0) info+=tp+" pasta"+(tp>1?"s":"")
  if(ta>0) info+=(info?" + ":"")+ta+" arquivo"+(ta>1?"s":"")
  document.getElementById("file-counter").textContent=info||"Nenhum arquivo carregado"
  document.getElementById("juntar-info").textContent=ta>1 ? info+" | Arraste as miniaturas para reorganizar" : info
  // Atualiza nome do botão
  const btn=document.getElementById("actionBtn")
  if(tp>0){
    btn.textContent=tp>1?"GERAR "+tp+" PDFs (1 por pasta)":"GERAR PDF DA PASTA"
  }else{
    btn.textContent="CONSOLIDAR PDF"
  }
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
  document.getElementById("resultado-pastas").style.display="none"
  document.getElementById("resultado-pastas-btns").innerHTML=""
  const nomeInput=document.getElementById("finalFileName")
  if(nomeInput) nomeInput.value=""
  ocultarProgressoJuntar()
}

document.getElementById("actionBtn").addEventListener("click",async()=>{
  if(arquivos.length===0 && pastasJuntar.length===0){
    alert("Selecione pelo menos um arquivo ou pasta")
    return
  }
  try{
    if(pastasJuntar.length>0){
      await gerarPDFsPorPasta()
    }else{
      await gerarPDFAvulsos()
    }
  }catch(e){
    console.error(e)
    ocultarProgressoJuntar()
    alert("Nao foi possivel concluir o processamento dos arquivos.")
  }
})

async function gerarPDFAvulsos(){
  document.getElementById("resultado-juntar").style.display="none"
  atualizarProgressoJuntar(6,"Consolidando PDFs","Preparando arquivos e conferindo a ordem...")
  document.getElementById("juntar-info").textContent="Consolidando PDFs..."

  const mergedPdf=await PDFLib.PDFDocument.create()
  const totalArquivos=arquivos.length

  for(let i=0;i<totalArquivos;i++){
    const file=arquivos[i]
    atualizarProgressoJuntar(12+((i/Math.max(totalArquivos,1))*66),"Consolidando PDFs",`Lendo ${i+1} de ${totalArquivos}: ${file.name}`)
    const bytes=await file.arrayBuffer()
    const pdf=await carregarPdfPreservandoAssinaturaVisual(bytes)
    const ordemPaginas=obterOrdemPaginasArquivo(file, pdf.getPageCount())
    const pages=await mergedPdf.copyPages(pdf,ordemPaginas)
    pages.forEach(p=>mergedPdf.addPage(p))
    atualizarProgressoJuntar(18+(((i+1)/Math.max(totalArquivos,1))*66),"Consolidando PDFs",`Arquivo ${i+1} de ${totalArquivos} incorporado ao PDF final.`)
  }

  atualizarProgressoJuntar(92,"Finalizando PDF","Salvando e preparando o download...")
  resultadoJuntarBytes=await mergedPdf.save()
  const nomeInput=(document.getElementById("finalFileName")?.value || "").trim()
  resultadoJuntarNome=(nomeInput||"arquivo_unificado")+".pdf"
  const pdfFinal=await PDFLib.PDFDocument.load(resultadoJuntarBytes)
  const totalPags=pdfFinal.getPageCount()
  document.getElementById("juntar-info").textContent=arquivos.length+" arquivo(s) consolidados em "+totalPags+" páginas"
  await finalizarTelaProcesso(
    "unificador",
    "juntar-processo",
    "PDF concluído",
    resultadoJuntarNome+" pronto para download.",
    "Baixar PDF",
    baixarResultadoJuntar
  )
}

async function gerarPDFsPorPasta(){
  resultadosPastasJuntar=[]

  const todasPastas=[...pastasJuntar]
  if(arquivos.length>0){
    const nomeAvulso=(document.getElementById("finalFileName")?.value || "").trim()||"Arquivos_Avulsos"
    todasPastas.push({nomePasta:nomeAvulso,arquivos})
  }

  for(let i=0;i<todasPastas.length;i++){
    const {nomePasta,arquivos:arqs}=todasPastas[i]
    atualizarProgressoJuntar(8+((i/Math.max(todasPastas.length,1))*72),"Gerando PDFs por pasta",`Pasta ${i+1} de ${todasPastas.length}: ${nomePasta}`)
    document.getElementById("juntar-info").textContent="Gerando: "+nomePasta+" ("+(i+1)+"/"+todasPastas.length+")..."
    const mergedPdf=await PDFLib.PDFDocument.create()
    const ordenados=[...arqs]
    for(let j=0;j<ordenados.length;j++){
      const file=ordenados[j]
      try{
        atualizarProgressoJuntar((((i)+(j+1)/Math.max(ordenados.length,1))/Math.max(todasPastas.length,1))*100,"Gerando PDFs por pasta",`Pasta ${i+1}/${todasPastas.length} · arquivo ${j+1}/${ordenados.length}: ${file.name}`)
        let f=file
        if(file.type!=="application/pdf") f=await imagemParaPDF(file)
        const bytes=await f.arrayBuffer()
        const pdf=await carregarPdfPreservandoAssinaturaVisual(bytes)
        const ordemPaginas=obterOrdemPaginasArquivo(file, pdf.getPageCount())
        const pages=await mergedPdf.copyPages(pdf,ordemPaginas)
        pages.forEach(p=>mergedPdf.addPage(p))
      }catch(e){ console.warn("Erro:",file.name,e) }
    }
    const pdfBytes=await mergedPdf.save()
    const nomeFinal=nomePasta+".pdf"
    resultadosPastasJuntar.push({nome:nomeFinal,bytes:pdfBytes})
  }


  atualizarProgressoJuntar(92,"Finalizando arquivos","Empacotando os PDFs gerados para download...")
  document.getElementById("juntar-info").textContent=todasPastas.length+" PDF"+(todasPastas.length>1?"s gerados":"gerado")+"!"
  if(resultadosPastasJuntar.length===1){
    await finalizarTelaProcesso(
      "unificador",
      "juntar-processo",
      "PDF concluído",
      resultadosPastasJuntar[0].nome+" pronto para download.",
      "Baixar PDF",
      ()=>baixarArquivo(resultadosPastasJuntar[0].bytes,resultadosPastasJuntar[0].nome)
    )
    return
  }

  await finalizarTelaProcesso(
    "unificador",
    "juntar-processo",
    "Arquivos concluídos",
    resultadosPastasJuntar.length+" PDFs prontos para download em ZIP.",
    "Baixar arquivos",
    baixarTudoZipado
  )
}

async function baixarResultadosPastasAutomaticamente(){
  if(resultadosPastasJuntar.length===1){
    baixarArquivo(resultadosPastasJuntar[0].bytes,resultadosPastasJuntar[0].nome)
    mostrarSucesso("Download iniciado automaticamente.")
    return
  }

  if(resultadosPastasJuntar.length>1){
    await baixarTudoZipado()
    mostrarSucesso("ZIP iniciado automaticamente.")
  }
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

// ===================== DIVIDIR PDF =====================

function mostrarInputDivisao(modo){
  document.getElementById("input-pages").style.display = modo==="pages" ? "block" : "none"
  document.getElementById("input-size").style.display  = modo==="size"  ? "block" : "none"
}

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
  ocultarProgressoDividir()

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
  // Resetar modo de divisão
  document.querySelectorAll('input[name="splitMode"]').forEach(r=>r.checked=false)
  document.getElementById("input-pages").style.display="none"
  document.getElementById("input-size").style.display="none"
  document.getElementById("pagesPerFile").value=""
  document.getElementById("sizePerFile").value=""
  ocultarProgressoDividir()
}

async function dividirPDF(){
  if(!splitFileRef){
    alert("Selecione um PDF primeiro")
    return
  }

  const modoSelecionado=document.querySelector('input[name="splitMode"]:checked')
  if(!modoSelecionado){
    alert("Selecione um modo de divisão antes de continuar.")
    return
  }

  resultadosDividirList=[]
  document.getElementById("resultado-dividir").style.display="none"
  try{
    const file=splitFileRef
    const bytes=await file.arrayBuffer()
    let pdf=await carregarPdfPreservandoAssinaturaVisual(bytes)
    pdf=await recriarPdfComOrdemAplicada(pdf, file)
    const totalPaginas=pdf.getPageCount()
    const modo=modoSelecionado.value

    if(modo==="size"){
      alert("A divisao por tamanho ainda nao esta disponivel.")
      return
    }

    atualizarProgressoDividir(8,"Preparando divisao","Lendo o arquivo selecionado e validando as paginas...")
    document.getElementById("dividir-info").textContent="Dividindo PDF..."

  // DIVIDIR POR QUANTIDADE DE PAGINAS
  if(modo==="pages"){
    const qtd=parseInt(document.getElementById("pagesPerFile").value)
    if(!qtd || qtd<=0){
      ocultarProgressoDividir()
      alert("Digite a quantidade de páginas")
      return
    }
    let contador=1
    for(let i=0;i<totalPaginas;i+=qtd){
      atualizarProgressoDividir(
        (Math.min(i+qtd,totalPaginas)/Math.max(totalPaginas,1))*84,
        "Separando paginas",
        `Bloco ${contador} em preparacao...`
      )
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
    for(let i=0;i<totalPaginas;i++){
      if((i+1)%2!==0) paginas.push(i)
      atualizarProgressoDividir(
        ((i+1)/Math.max(totalPaginas,1))*84,
        "Separando paginas impares",
        `Analisando pagina ${i+1} de ${totalPaginas}...`
      )
    }
    const copied=await novoPdf.copyPages(pdf,paginas)
    copied.forEach(p=>novoPdf.addPage(p))
    const pdfBytes=await novoPdf.save()
    resultadosDividirList.push({bytes:pdfBytes, nome:"paginas_impares.pdf"})
  }

  // PAGINAS PARES
  if(modo==="even"){
    const novoPdf=await PDFLib.PDFDocument.create()
    const paginas=[]
    for(let i=0;i<totalPaginas;i++){
      if((i+1)%2===0) paginas.push(i)
      atualizarProgressoDividir(
        ((i+1)/Math.max(totalPaginas,1))*84,
        "Separando paginas pares",
        `Analisando pagina ${i+1} de ${totalPaginas}...`
      )
    }
    const copied=await novoPdf.copyPages(pdf,paginas)
    copied.forEach(p=>novoPdf.addPage(p))
    const pdfBytes=await novoPdf.save()
    resultadosDividirList.push({bytes:pdfBytes, nome:"paginas_pares.pdf"})
  }

    atualizarProgressoDividir(92,"Finalizando divisao","Organizando os arquivos gerados para download...")
    document.getElementById("dividir-info").textContent="Divisão concluída: "+resultadosDividirList.length+" arquivo(s) gerado(s)"
    await finalizarTelaProcesso(
      "dividir-pdf",
      "dividir-processo",
      "Divisao concluida",
      resultadosDividirList.length+" arquivo(s) pronto(s) para download.",
      resultadosDividirList.length>1 ? "Baixar arquivos" : "Baixar arquivo",
      baixarResultadosDividir
    )
  }catch(e){
    console.error(e)
    ocultarProgressoDividir()
    alert("Nao foi possivel dividir o arquivo.")
  }
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
let ocrFiles=[]
let ocrResultados=[]
let ocrWorkerRef=null

document.getElementById("ocrInput").addEventListener("change",async(e)=>{
  const files=Array.from(e.target.files||[])
  if(!files.length) return
  await carregarFilesNoOCR(files)
  e.target.value=""
})

async function carregarFilesNoOCR(files){
  const validos=files.filter(f=>f && (f.type==="application/pdf" || /\.pdf$/i.test(f.name)))
  if(validos.length===0){
    alert("Selecione PDF(s) válido(s).")
    return
  }

  const existentes=new Set(ocrFiles.map(f=>`${f.name}_${f.size}_${f.lastModified}`))
  const novos=validos.filter(f=>!existentes.has(`${f.name}_${f.size}_${f.lastModified}`))
  if(novos.length===0) return

  ocrFiles.push(...novos)
  ocrTextoGlobal=""
  ocrResultados=[]
  document.getElementById("ocr-resultado-box").style.display="none"
  document.getElementById("ocr-status").textContent=""

  const gallery=document.getElementById("ocr-gallery")
  for(const file of novos){
    try{
      const {card}=await criarThumbCard(file)
      gallery.appendChild(card)
    }catch{
      const div=document.createElement("div")
      div.className="thumb-card"
      div.innerHTML=`<div class="thumb-header"><strong>${file.name}</strong></div><div class="thumb-meta">PDF</div>`
      gallery.appendChild(div)
    }
  }

  const totalArquivos=ocrFiles.length
  document.getElementById("ocr-info").textContent=`${totalArquivos} arquivo(s) pronto(s) para OCR`
}

function limparOCR(){
  ocrTextoGlobal=""
  ocrFiles=[]
  ocrResultados=[]
  document.getElementById("ocr-gallery").innerHTML=""
  document.getElementById("ocr-status").textContent=""
  document.getElementById("ocr-resultado-box").style.display="none"
  document.getElementById("ocr-info").textContent=""
  document.getElementById("ocrInput").value=""
}

async function obterOCRWorker(statusEl){
  if(ocrWorkerRef) return ocrWorkerRef
  if(!window.Tesseract) throw new Error("Tesseract não carregado.")
  statusEl.textContent="Carregando motor OCR..."
  ocrWorkerRef=await Tesseract.createWorker('por', 1, {
    logger: (m)=>{
      if(m.status==='recognizing text' && Number.isFinite(m.progress)){
        statusEl.textContent=`OCR em andamento... ${Math.round(m.progress*100)}%`
      }
    }
  })
  try{
    await ocrWorkerRef.setParameters({
      tessedit_pageseg_mode: Tesseract.PSM.AUTO,
      preserve_interword_spaces: '1'
    })
  }catch{}
  return ocrWorkerRef
}

function nomePdfOCR(nome){
  return nome.replace(/\.pdf$/i,'') + '_OCR.pdf'
}

async function reconhecerTexto(){
  if(ocrFiles.length===0){
    alert("Selecione ao menos um PDF primeiro")
    return
  }

  const status=document.getElementById("ocr-status")
  const resultadoBox=document.getElementById("ocr-resultado-box")
  status.textContent="Preparando OCR..."
  resultadoBox.style.display="none"
  ocrTextoGlobal=""
  ocrResultados=[]

  const worker=await obterOCRWorker(status)
  const canvas=document.createElement('canvas')
  const ctx=canvas.getContext('2d', { willReadFrequently: true })

  for(let fileIndex=0; fileIndex<ocrFiles.length; fileIndex++){
    const file=ocrFiles[fileIndex]
    status.textContent=`Abrindo ${file.name} (${fileIndex+1}/${ocrFiles.length})...`

    const sourceBytes=new Uint8Array(await file.arrayBuffer())
    const pdfjsDoc=await pdfjsLib.getDocument({data:sourceBytes}).promise
    const sourcePdfLib=await PDFLib.PDFDocument.load(sourceBytes)
    const outputPdf=await PDFLib.PDFDocument.create()
    const outputFont=await outputPdf.embedFont(PDFLib.StandardFonts.Helvetica)
    let textoArquivo=''

    for(let i=1;i<=pdfjsDoc.numPages;i++){
      const page=await pdfjsDoc.getPage(i)
      status.textContent=`${file.name}: página ${i} de ${pdfjsDoc.numPages}...`

      const content=await page.getTextContent()
      const textoExistente=(content.items||[]).map(item=>item.str).join(' ').trim()
      if(textoExistente.length>20){
        const [copiedPage]=await outputPdf.copyPages(sourcePdfLib,[i-1])
        outputPdf.addPage(copiedPage)
        textoArquivo+=textoExistente + "\n"
        continue
      }

      const viewport=page.getViewport({scale:1.25})
      canvas.width=Math.ceil(viewport.width)
      canvas.height=Math.ceil(viewport.height)
      ctx.clearRect(0,0,canvas.width,canvas.height)
      await page.render({canvasContext:ctx, viewport}).promise

      const {data}=await worker.recognize(canvas)
      textoArquivo += (data.text||'') + "\n"

      const jpgUrl=canvas.toDataURL('image/jpeg',0.72)
      const jpgImg=await outputPdf.embedJpg(jpgUrl)
      const pdfPage=outputPdf.addPage([viewport.width, viewport.height])
      pdfPage.drawImage(jpgImg,{x:0,y:0,width:viewport.width,height:viewport.height})

      const words=(data.words||[]).filter(w=>w && w.text && w.bbox)
      for(const word of words){
        const x0=word.bbox.x0 ?? word.bbox.left ?? 0
        const y0=word.bbox.y0 ?? word.bbox.top ?? 0
        const x1=word.bbox.x1 ?? word.bbox.right ?? x0
        const y1=word.bbox.y1 ?? word.bbox.bottom ?? y0
        const width=Math.max(1, x1-x0)
        const height=Math.max(7, y1-y0)
        const y=viewport.height - y1
        const safeText=String(word.text).replace(/\s+/g,' ').trim()
        if(!safeText) continue
        pdfPage.drawText(safeText,{
          x:x0,
          y:y,
          size:Math.max(6, Math.min(22, height*0.85)),
          font:outputFont,
          color:PDFLib.rgb(1,1,1),
          opacity:0.01,
          maxWidth:width+2,
          lineHeight:height
        })
      }
    }

    const resultBytes=await outputPdf.save()
    ocrResultados.push({
      nome:nomePdfOCR(file.name),
      bytes:resultBytes,
      texto:textoArquivo.trim()
    })
    ocrTextoGlobal += (textoArquivo.trim() + "\n\n")
  }

  status.textContent=`OCR concluído para ${ocrResultados.length} arquivo(s).`
  resultadoBox.style.display="block"
  resultadoBox.scrollIntoView({behavior:"smooth"})
}

async function baixarResultadosOCR(){
  if(ocrResultados.length===0){
    alert("Nenhum resultado OCR disponível.")
    return
  }
  if(ocrResultados.length===1){
    const item=ocrResultados[0]
    baixarArquivo(item.bytes,item.nome)
    return
  }
  const JSZipCtor=await carregarJSZip()
  const zip=new JSZipCtor()
  for(const item of ocrResultados){
    zip.file(item.nome,item.bytes)
  }
  const blob=await zip.generateAsync({type:'blob'})
  const link=document.createElement('a')
  link.href=URL.createObjectURL(blob)
  link.download='OCR_Resultados.zip'
  link.click()
}

function copiarTexto(){
  if(!ocrTextoGlobal.trim()){
    alert("Nenhum texto reconhecido ainda.")
    return
  }
  navigator.clipboard.writeText(ocrTextoGlobal).then(()=>{
    alert("Texto copiado!")
  })
}

async function encerrarOCRWorker(){
  if(ocrWorkerRef){
    try{ await ocrWorkerRef.terminate() }catch{}
    ocrWorkerRef=null
  }
}
window.addEventListener('beforeunload', ()=>{ encerrarOCRWorker() })

// ===================== COMPACTAR PDF =====================

let arquivosCompactar=[]

document.getElementById("compactarInput").addEventListener("change",async(e)=>{
  await carregarFilesNoCompactar(Array.from(e.target.files).filter(f=>f.type==="application/pdf"))
  document.getElementById("compactarInput").value=""
})

function selecionarPastaCompactar(){
  const input=document.createElement("input")
  input.type="file"
  input.webkitdirectory=true
  input.style.display="none"
  document.body.appendChild(input)
  input.addEventListener("change",async(e)=>{
    const files=Array.from(e.target.files).filter(f=>f.type==="application/pdf")
    document.body.removeChild(input)
    if(files.length===0){alert("Nenhum PDF encontrado na pasta.");return}
    mostrarLoading("Carregando pasta: "+files.length+" PDF"+(files.length>1?"s":"")+"...")
    await carregarFilesNoCompactar(files)
  })
  input.click()
}

async function carregarFilesNoCompactar(files){
  mostrarLoading("Carregando "+files.length+" arquivo"+(files.length>1?"s":"")+"...")
  for(let file of files){
    arquivosCompactar.push(file)
    const {card}=await criarThumbCard(file)
    document.getElementById("compactar-lista").appendChild(card)
  }
  esconderLoading()
  mostrarSucesso(files.length+" arquivo"+(files.length>1?"s carregados":"carregado")+"!")
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
  document.getElementById("resultado-compactar").style.display="none"
}

function formatarTamanho(bytes){
  if(bytes>=1024*1024) return (bytes/(1024*1024)).toFixed(2)+" MB"
  return (bytes/1024).toFixed(1)+" KB"
}

// Pausa para não travar o browser
function yield_(){
  return new Promise(r=>setTimeout(r,0))
}

// Renderiza cada página do PDF como imagem e monta novo PDF
async function compactarComImagensRecomprimidas(arrayBuffer, qualidade, onProgresso){
  const pdfSrc=await pdfjsLib.getDocument({data:new Uint8Array(arrayBuffer)}).promise
  const novoPdf=await PDFLib.PDFDocument.create()

  // Normal: JPEG 80% escala 1.2 — boa legibilidade com redução real (~40-60%)
  // Extrema: JPEG 60% escala 1.0 — redução agressiva (~60-80%)
  const escala=qualidade>0.5?1.2:1.0
  const qualidadeJpeg=qualidade>0.5?0.80:0.60

  const canvas=document.createElement("canvas")
  const ctx=canvas.getContext("2d")

  for(let i=1;i<=pdfSrc.numPages;i++){
    if(onProgresso) onProgresso(i, pdfSrc.numPages)
    await yield_()

    const page=await pdfSrc.getPage(i)
    const viewport=page.getViewport({scale:escala})
    canvas.width=viewport.width
    canvas.height=viewport.height
    ctx.fillStyle="#ffffff"
    ctx.fillRect(0,0,canvas.width,canvas.height)
    await page.render({canvasContext:ctx, viewport}).promise
    await yield_()

    const blob=await new Promise(res=>canvas.toBlob(res,"image/jpeg",qualidadeJpeg))
    const jpegBytes=new Uint8Array(await blob.arrayBuffer())
    const img=await novoPdf.embedJpg(jpegBytes)
    const pdfPage=novoPdf.addPage([viewport.width, viewport.height])
    pdfPage.drawImage(img,{x:0,y:0,width:viewport.width,height:viewport.height})
  }

  canvas.remove()
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

    const pdfBytes=await compactarComImagensRecomprimidas(arrayBuffer, qualidade, (pag,total)=>{
      status.textContent="Compactando "+file.name+" ("+(i+1)+"/"+arquivosCompactar.length+") — página "+pag+" de "+total+"..."
    })
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

async function adicionarPastaListaJuntarPastas(nomePasta, arquivos){
  if(pastasLista.length >= MAX_PASTAS) return false
  pastasLista.push({
    nomePasta,
    arquivos: ordenarArquivosPastaInicialmente(arquivos),
    expandida: false
  })
  await renderizarPastaCard(pastasLista.length-1)
  return true
}

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
    document.body.removeChild(input)
    const arquivosSelecionados = Array.from(e.target.files||[])
    const pdfs = arquivosSelecionados.filter(f=>f.type==="application/pdf" || /\.pdf$/i.test(f.name))
    const compactados = arquivosSelecionados.filter(arquivoEhCompactado)
    if(pdfs.length===0 && compactados.length===0){
      alert("Nenhum PDF ou compactado encontrado nessa pasta.")
      return
    }

    const nomePasta = arquivosSelecionados[0]?.webkitRelativePath?.split("/")[0] || ("Pasta "+(pastasLista.length+1))
    let adicionadas = 0
    let compactadosLidos = 0
    let compactadosIgnorados = 0

    mostrarLoading("Lendo pasta selecionada...")

    if(pdfs.length>0 && pastasLista.length < MAX_PASTAS){
      const ok = await adicionarPastaListaJuntarPastas(nomePasta, pdfs)
      if(ok) adicionadas++
    }

    for(const compactado of compactados){
      if(pastasLista.length >= MAX_PASTAS) break
      try{
        const grupos = await extrairPastasDeCompactadoParaJuntar(compactado)
        if(grupos.length===0){
          compactadosIgnorados++
          continue
        }
        compactadosLidos++
        const baseCompactado = nomeBaseSemExtensao(compactado.name)
        for(const grupo of grupos){
          if(pastasLista.length >= MAX_PASTAS) break
          const pdfsGrupo = grupo.arquivos.filter(f=>f.type==="application/pdf" || /\.pdf$/i.test(f.name))
          if(pdfsGrupo.length===0) continue
          const nomeGrupo = grupos.length===1
            ? baseCompactado
            : `${baseCompactado} - ${grupo.nomePasta}`
          const ok = await adicionarPastaListaJuntarPastas(nomeGrupo, pdfsGrupo)
          if(ok) adicionadas++
        }
      }catch(err){
        console.warn("Nao foi possivel ler o compactado dentro da pasta:", compactado.name, err)
        compactadosIgnorados++
      }
    }

    esconderLoading()
    atualizarInfoJuntarPastas()

    if(adicionadas===0){
      alert("Nenhum PDF valido foi encontrado na pasta ou nos compactados.")
      return
    }

    if(compactados.length>0){
      const resumoCompactados = compactadosLidos>0
        ? ` ${compactadosLidos} compactado(s) lido(s).`
        : ""
      const resumoIgnorados = compactadosIgnorados>0
        ? ` ${compactadosIgnorados} compactado(s) nao puderam ser usados.`
        : ""
      mostrarSucesso(adicionadas+" entrada"+(adicionadas>1?"s adicionadas":" adicionada")+"."+resumoCompactados+resumoIgnorados)
      return
    }
  })
  input.click()
}

async function adicionarCompactadoJuntarPastas(){
  if(pastasLista.length >= MAX_PASTAS){
    alert("Limite de "+MAX_PASTAS+" pastas atingido.")
    return
  }

  const input=document.getElementById("pastaCompactadoInput")
  if(!input) return

  input.onchange=async(e)=>{
    const arquivo=e.target.files?.[0]
    input.value=""
    if(!arquivo) return

    mostrarLoading("Lendo compactado: "+arquivo.name+"...")
    let grupos=[]
    try{
      grupos=await extrairPastasDeCompactadoParaJuntar(arquivo)
    }catch(err){
      console.error(err)
      esconderLoading()
      alert("Não foi possível abrir esse compactado. Verifique se ele não está protegido por senha ou corrompido.")
      return
    }

    if(grupos.length===0){
      esconderLoading()
      alert("Nenhum PDF encontrado nesse compactado.")
      return
    }

    let adicionadas=0
    for(const grupo of grupos){
      const ok=await adicionarPastaListaJuntarPastas(grupo.nomePasta, grupo.arquivos.filter(f=>f.type==="application/pdf"))
      if(ok) adicionadas++
      if(pastasLista.length >= MAX_PASTAS) break
    }

    esconderLoading()
    atualizarInfoJuntarPastas()

    if(adicionadas===0){
      alert("Não foi possível adicionar novas pastas. Verifique o limite atual.")
      return
    }

    if(grupos.length>adicionadas){
      mostrarSucesso(adicionadas+" pasta(s) adicionadas. O restante foi ignorado por causa do limite.")
    }else{
      mostrarSucesso(adicionadas+" pasta"+(adicionadas>1?"s adicionadas":" adicionada")+" via compactado!")
    }
  }

  input.click()
}

async function renderizarPastaCard(idx){
  const {nomePasta, arquivos, expandida} = pastasLista[idx]
  const lista = document.getElementById("pastas-lista")

  const card = document.createElement("div")
  card.className = "pasta-card"
  if(expandida) card.classList.add("expandida")
  card.id = "pasta-card-"+idx

  // Header da pasta
  const header = document.createElement("div")
  header.className = "pasta-card-header"

  const numero = document.createElement("div")
  numero.className = "pasta-numero"
  numero.textContent = String(idx+1).padStart(2,"0")

  const info = document.createElement("div")
  info.className = "pasta-card-info"

  const nomeEl = document.createElement("div")
  nomeEl.className = "pasta-nome"
  nomeEl.textContent = nomePasta

  const qtdEl = document.createElement("div")
  qtdEl.className = "pasta-qtd"
  qtdEl.textContent = arquivos.length+" PDF"+(arquivos.length>1?"s":"")

  const hintEl = document.createElement("div")
  hintEl.className = "pasta-hint"
  hintEl.textContent = expandida
    ? "Arraste os PDFs para organizar a ordem final."
    : "Clique para abrir e organizar os PDFs desta pasta."

  const acoes = document.createElement("div")
  acoes.className = "pasta-card-actions"

  const btnExpandir = document.createElement("button")
  btnExpandir.className = "btn atalho btn-sm"
  btnExpandir.type = "button"
  btnExpandir.textContent = expandida ? "Ocultar PDFs" : "Organizar PDFs"
  btnExpandir.onclick = (e)=>{
    e.stopPropagation()
    alternarExpandirPasta(idx)
  }

  const btnRemover = document.createElement("button")
  btnRemover.className = "btn danger btn-sm"
  btnRemover.type = "button"
  btnRemover.textContent = "Remover"
  btnRemover.onclick = (e)=>{
    e.stopPropagation()
    removerPasta(idx)
  }

  info.appendChild(nomeEl)
  info.appendChild(qtdEl)
  info.appendChild(hintEl)
  acoes.appendChild(btnExpandir)
  acoes.appendChild(btnRemover)
  header.appendChild(numero)
  header.appendChild(info)
  header.appendChild(acoes)
  card.appendChild(header)

  const thumbs = document.createElement("div")
  thumbs.className = "pasta-thumbs"
  thumbs.onclick = ()=>alternarExpandirPasta(idx)

  for(let i=0;i<Math.min(arquivos.length,5);i++){
    try{
      const {card:thumb} = await criarThumbCard(arquivos[i])
      thumb.classList.add("thumb-card-folder-preview")
      const badge=thumb.querySelector(".thumb-order-badge")
      if(badge) badge.textContent = String(i+1).padStart(2,"0")
      thumbs.appendChild(thumb)
    }catch(e){}
  }
  if(arquivos.length>5){
    const mais = document.createElement("div")
    mais.className = "thumb-mais"
    mais.textContent = "+"+(arquivos.length-5)
    thumbs.appendChild(mais)
  }

  card.appendChild(thumbs)

  if(expandida){
    const painel = document.createElement("div")
    painel.className = "pasta-expandida-painel"

    const topo = document.createElement("div")
    topo.className = "pasta-expandida-topo"

    const titulo = document.createElement("div")
    titulo.className = "pasta-expandida-titulo"
    titulo.textContent = "Ordem dos PDFs desta pasta"

    const texto = document.createElement("div")
    texto.className = "pasta-expandida-texto"
    texto.textContent = "Arraste uma miniatura sobre a outra para reorganizar o PDF final."

    topo.appendChild(titulo)
    topo.appendChild(texto)

    const grid = document.createElement("div")
    grid.className = "pasta-thumbs-expandida-grid"

    for(let i=0;i<arquivos.length;i++){
      try{
        const {card:thumb} = await criarThumbCard(arquivos[i])
        thumb.classList.add("thumb-card-folder")
        const badge=thumb.querySelector(".thumb-order-badge")
        if(badge) badge.textContent = String(i+1).padStart(2,"0")
        habilitarOrdenacaoArquivoDentroDaPasta(thumb, idx, i)
        grid.appendChild(thumb)
      }catch(e){}
    }

    painel.appendChild(topo)
    painel.appendChild(grid)
    card.appendChild(painel)
  }

  lista.appendChild(card)
}

function removerPasta(idx){
  pastasLista.splice(idx,1)
  rerenderPastasLista()
  atualizarInfoJuntarPastas()
}

async function rerenderPastasLista(){
  const lista=document.getElementById("pastas-lista")
  if(!lista) return
  lista.innerHTML=""
  for(let i=0;i<pastasLista.length;i++){
    await renderizarPastaCard(i)
  }
}

function alternarExpandirPasta(idx){
  const pasta=pastasLista[idx]
  if(!pasta) return
  pasta.expandida=!pasta.expandida
  rerenderPastasLista()
}

function reordenarArquivosDaPasta(pastaIdx, origemIdx, destinoIdx){
  const pasta=pastasLista[pastaIdx]
  if(!pasta) return
  if(origemIdx===destinoIdx) return
  if(origemIdx<0 || destinoIdx<0) return
  if(origemIdx>=pasta.arquivos.length || destinoIdx>=pasta.arquivos.length) return
  const [arquivo]=pasta.arquivos.splice(origemIdx,1)
  pasta.arquivos.splice(destinoIdx,0,arquivo)
  pasta.expandida=true
  rerenderPastasLista()
}

function habilitarOrdenacaoArquivoDentroDaPasta(card, pastaIdx, arquivoIdx){
  card.draggable=true
  card.classList.add("thumb-card-sortable")
  card.setAttribute("aria-label","Arraste para reorganizar o PDF dentro da pasta")

  card.addEventListener("dragstart",(e)=>{
    pastaArquivoArrastandoInfo={pastaIdx, arquivoIdx}
    bloquearCliqueThumbAte=Date.now()+350
    card.classList.add("thumb-card-dragging")
    if(e.dataTransfer){
      e.dataTransfer.effectAllowed="move"
      e.dataTransfer.setData("text/plain", `${pastaIdx}:${arquivoIdx}`)
    }
  })

  card.addEventListener("dragend",()=>{
    pastaArquivoArrastandoInfo=null
    bloquearCliqueThumbAte=Date.now()+350
    card.classList.remove("thumb-card-dragging")
    document.querySelectorAll(".thumb-card-drop-target").forEach(el=>el.classList.remove("thumb-card-drop-target"))
  })

  card.addEventListener("dragover",(e)=>{
    if(!pastaArquivoArrastandoInfo) return
    if(pastaArquivoArrastandoInfo.pastaIdx!==pastaIdx) return
    if(pastaArquivoArrastandoInfo.arquivoIdx===arquivoIdx) return
    e.preventDefault()
    card.classList.add("thumb-card-drop-target")
    if(e.dataTransfer) e.dataTransfer.dropEffect="move"
  })

  card.addEventListener("dragleave",()=>{
    card.classList.remove("thumb-card-drop-target")
  })

  card.addEventListener("drop",(e)=>{
    if(!pastaArquivoArrastandoInfo) return
    if(pastaArquivoArrastandoInfo.pastaIdx!==pastaIdx) return
    if(pastaArquivoArrastandoInfo.arquivoIdx===arquivoIdx) return
    e.preventDefault()
    e.stopPropagation()
    card.classList.remove("thumb-card-drop-target")
    reordenarArquivosDaPasta(pastaIdx, pastaArquivoArrastandoInfo.arquivoIdx, arquivoIdx)
  })
}

function atualizarInfoJuntarPastas(){
  const total = pastasLista.length
  document.getElementById("pastas-contador").textContent = total+"/"+MAX_PASTAS+" pastas"
  document.getElementById("juntarpastas-info").textContent = total+" pasta"+(total!==1?"s":"")+" | "+(total<MAX_PASTAS?"Pode adicionar pasta ou compactado":"Limite atingido")
  document.getElementById("btn-add-pasta").disabled = total>=MAX_PASTAS
  const btnCompactado=document.getElementById("btn-add-compactado")
  if(btnCompactado) btnCompactado.disabled = total>=MAX_PASTAS
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
  const btnCompactado=document.getElementById("btn-add-compactado")
  if(btnCompactado) btnCompactado.disabled=false
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

    const arquivosOrdenados = [...arquivos]

    for(let file of arquivosOrdenados){
      try{
        const bytes = await file.arrayBuffer()
        const pdf = await carregarPdfPreservandoAssinaturaVisual(bytes)
        const ordemPaginas = obterOrdemPaginasArquivo(file, pdf.getPageCount())
        const pages = await mergedPdf.copyPages(pdf, ordemPaginas)
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
