const configuracoes = [
  { id:"computadores", nome:"Computadores", aba:"Notebook's", icone:"▣", cor:"azul" },
  { id:"monitores", nome:"Monitores", aba:"Monitores", icone:"▤", cor:"verde" },
  { id:"switches", nome:"Switches", aba:"Switches", icone:"⌘", cor:"roxo" },
  { id:"aps", nome:"Access Points", aba:"AP's", icone:"⌁", cor:"amarelo" },
  { id:"servidores", nome:"Servidores", aba:"Servidores", icone:"▥", cor:"azul" },
  { id:"antispam", nome:"E-mails tratados", aba:"Antispam", icone:"✉", cor:"verde", colunaTotal:"Total de e-mails tratados" },
  { id:"virus", nome:"Vírus tratados", aba:"Virus", icone:"✓", cor:"amarelo" },
  { id:"chamados", nome:"Chamados", aba:"Chamados", icone:"◷", cor:"roxo" }
];

// Dados iniciais retirados da planilha enviada. Ao importar um novo Excel,
// estes valores são substituídos automaticamente.
let banco = {
  "Notebook's": dadosMensais(["Data","Desktop","Note Próprio","Mac Próprio","Estoque","Aluguel"], [
    ["2026-03-01",3,222,4,8,0],["2026-04-01",3,221,8,5,0],["2026-05-01",3,217,8,9,1],
    ["2026-06-01",3,222,8,3,1],["2026-07-01",3,222,8,2,1],["2026-08-01",3,222,8,2,1]
  ]),
  "Monitores": dadosMensais(["Data","CNseg","Home Office","Estoque"], [
    ["2026-02-01",175,23,5],["2026-03-01",175,11,5],["2026-04-01",180,2,2],
    ["2026-05-01",180,2,2],["2026-06-01",180,0,3],["2026-07-01",180,0,3]
  ]),
  "Switches": dadosMensais(["Data","Estoque","9º Andar","8º Andar","7º Andar","17º Andar","16º Andar","15º Andar","12º Andar","Brasília"], [
    ["2026-06-01",12,4,6,13,2,3,5,6,7],["2026-07-01",12,4,6,13,2,3,5,6,7],["2026-08-01",12,4,6,13,2,3,5,6,7]
  ]),
  "AP's": dadosMensais(["Data","12º Andar","15º Andar","16º Andar","17º Andar","7º Andar","8º Andar","9º Andar","Brasília"], [
    ["2026-03-01",5,5,5,2,5,7,3,10],["2026-04-01",5,5,5,2,5,7,3,10],["2026-05-01",5,5,5,2,5,7,3,10],
    ["2026-06-01",5,5,5,2,5,7,3,10],["2026-07-01",5,5,5,2,5,7,3,10],["2026-08-01",5,5,5,2,5,7,3,10]
  ]),
  "Servidores": dadosMensais(["Data","Servidores Virtuais Equinix","Servidores Virtuais CNseg","Servidores Físicos CNseg RJ","Servidores Físicos CNseg BSB"], [
    ["2026-02-01",26,54,4,0],["2026-03-01",26,54,4,0],["2026-04-01",26,54,4,0],
    ["2026-05-01",26,54,4,0],["2026-06-01",26,54,4,0],["2026-08-01",26,54,4,0]
  ]),
  "Antispam": dadosMensais(["Data","Total de e-mails suspeitos - Spam","Total de e-mails legítimos","Total de e-mails tratados"], [
    ["2026-02-01",109769,189548,435004],["2026-03-01",250430,160361,453525],["2026-04-01",184500,186122,233602],
    ["2026-05-01",357933,162627,374332],["2026-06-01",305444,178831,333207],["2026-07-01",305444,178831,333207]
  ]),
  "Virus": dadosMensais(["Data","Ocorrências de vírus tratadas"], [
    ["2026-02-01",13],["2026-03-01",6],["2026-04-01",3],["2026-05-01",6],["2026-06-01",10],["2026-07-01",10]
  ]),
  "Chamados": dadosMensais(["Data","Infraestrutura","Service Desk","Sistemas","Suporte Dataseg","Suporte DBA","Suporte MXM"], [
    ["2026-01-01",71,146,21,28,35,38],["2026-02-01",22,83,17,23,18,38],["2026-03-01",67,109,3,16,5,30],
    ["2026-04-01",25,142,11,17,24,28],["2026-05-01",25,142,11,17,24,28],["2026-06-01",25,142,11,17,24,28]
  ])
};

const elementos = {
  competencia: document.querySelector("#competencia"), arquivo: document.querySelector("#arquivoExcel"),
  cards: document.querySelector("#cardsKpi"), corpoTabela: document.querySelector("#corpoTabela"),
  indicadorGrafico: document.querySelector("#indicadorGrafico"), canvas: document.querySelector("#graficoEvolucao"),
  listaAtualizacao: document.querySelector("#listaAtualizacao"), nomeArquivo: document.querySelector("#nomeArquivo"),
  statusArquivo: document.querySelector("#statusArquivo"), rodape: document.querySelector("#rodapeStatus"),
  toast: document.querySelector("#toast"), exportar: document.querySelector("#exportarCsv")
};

iniciar();

function iniciar(){
  preencherIndicadores(); preencherCompetencias(); renderizarTudo();
  elementos.competencia.addEventListener("change", renderizarTudo);
  elementos.indicadorGrafico.addEventListener("change", renderizarGrafico);
  elementos.arquivo.addEventListener("change", importarPlanilha);
  elementos.exportar.addEventListener("click", exportarResumo);
  document.querySelectorAll(".menu-item").forEach(item => item.addEventListener("click", () => {
    document.querySelectorAll(".menu-item").forEach(i => i.classList.remove("ativo")); item.classList.add("ativo");
  }));
  window.addEventListener("resize", renderizarGrafico);
}

function dadosMensais(colunas, linhas){ return { colunas, linhas: linhas.map(l => ({ data:new Date(l[0]+"T12:00:00"), valores:l.slice(1) })) }; }

function preencherIndicadores(){
  elementos.indicadorGrafico.innerHTML = configuracoes.map(c => `<option value="${c.id}">${c.nome}</option>`).join("");
}

function preencherCompetencias(){
  const datas = [...new Set(Object.values(banco).flatMap(a => a.linhas.map(l => chaveData(l.data))))].sort().reverse();
  elementos.competencia.innerHTML = datas.map(d => `<option value="${d}">${formatarMes(lerChaveData(d))}</option>`).join("");
}

function renderizarTudo(){ renderizarCards(); renderizarTabela(); renderizarAtualizacoes(); renderizarGrafico(); }

function obterRegistro(config, dataEscolhida){
  const aba = banco[config.aba]; if(!aba) return null;
  const elegiveis = aba.linhas.filter(l => chaveData(l.data) <= chaveData(dataEscolhida)).sort((a,b) => b.data-a.data);
  return elegiveis[0] || [...aba.linhas].sort((a,b) => b.data-a.data)[0] || null;
}

function totalRegistro(config, registro){
  if(!registro) return 0; const aba = banco[config.aba];
  if(config.colunaTotal){ const indice = aba.colunas.indexOf(config.colunaTotal)-1; return numero(registro.valores[indice]); }
  return registro.valores.reduce((s,v) => s + numero(v), 0);
}

function variacao(config, registro){
  const aba = banco[config.aba]; const ordenadas = [...aba.linhas].sort((a,b)=>a.data-b.data);
  const indice = ordenadas.findIndex(l => chaveData(l.data)===chaveData(registro.data));
  if(indice <= 0) return {texto:"Sem comparação", classe:"neutra"};
  const atual=totalRegistro(config,registro), anterior=totalRegistro(config,ordenadas[indice-1]);
  if(!anterior) return {texto:"Sem comparação",classe:"neutra"};
  const percentual=((atual-anterior)/anterior)*100, sinal=percentual>0?"+":"";
  return {texto:`${sinal}${percentual.toFixed(1).replace(".",",")}% vs. anterior`,classe:percentual>0?"positiva":percentual<0?"negativa":"neutra"};
}

function renderizarCards(){
  const escolhida=lerChaveData(elementos.competencia.value);
  elementos.cards.innerHTML=configuracoes.map(config=>{
    const registro=obterRegistro(config,escolhida), total=totalRegistro(config,registro), varInfo=variacao(config,registro);
    const atrasado=registro && chaveData(registro.data)!==chaveData(escolhida);
    return `<article class="card" data-cor="${config.cor}">${atrasado?'<span class="atrasado">DADO ANTERIOR</span>':''}
      <div class="card-topo"><div class="card-icone">${config.icone}</div></div><p class="card-titulo">${config.nome}</p>
      <h2 class="card-valor">${formatarNumero(total)}</h2><div class="card-rodape"><span class="variacao ${varInfo.classe}">${varInfo.texto}</span>
      <span class="competencia-card">${registro?formatarMes(registro.data):"Sem dados"}</span></div></article>`;
  }).join("");
}

function renderizarTabela(){
  const escolhida=lerChaveData(elementos.competencia.value);
  elementos.corpoTabela.innerHTML=configuracoes.map(config=>{
    const aba=banco[config.aba], registro=obterRegistro(config,escolhida), atrasado=registro&&chaveData(registro.data)!==chaveData(escolhida);
    const composicao=registro?aba.colunas.slice(1).map((nome,i)=>`${nome}: ${formatarNumero(registro.valores[i])}`).join(" • "):"Sem dados";
    return `<tr><td class="indicador-nome">${config.nome}</td><td>${registro?formatarMes(registro.data):"—"}</td>
      <td class="total">${formatarNumero(totalRegistro(config,registro))}</td><td class="composicao">${composicao}</td>
      <td><span class="situacao ${atrasado?'pendente':'ok'}">● ${atrasado?'Último disponível':'Atualizado'}</span></td></tr>`;
  }).join("");
}

function renderizarAtualizacoes(){
  const escolhida=lerChaveData(elementos.competencia.value);
  elementos.listaAtualizacao.innerHTML=configuracoes.map(config=>{
    const ultima=[...(banco[config.aba]?.linhas||[])].sort((a,b)=>b.data-a.data)[0];
    const ok=ultima&&chaveData(ultima.data)>=chaveData(escolhida);
    return `<div class="linha-atualizacao"><div><strong>${config.nome}</strong><small>${ultima?formatarMes(ultima.data):"Sem dados"}</small></div>
      <span class="selo ${ok?'ok':'pendente'}">${ok?'ATUAL':'PENDENTE'}</span></div>`;
  }).join("");
}

function renderizarGrafico(){
  const config=configuracoes.find(c=>c.id===elementos.indicadorGrafico.value)||configuracoes[0], aba=banco[config.aba];
  const dados=[...(aba?.linhas||[])].sort((a,b)=>a.data-b.data).slice(-12).map(l=>({label:formatarMesCurto(l.data),valor:totalRegistro(config,l)}));
  document.querySelector("#descricaoGrafico").textContent=`Evolução de ${config.nome.toLowerCase()} nos últimos meses`;
  desenharGrafico(elementos.canvas,dados);
}

function desenharGrafico(canvas,dados){
  const ctx=canvas.getContext("2d"), dpr=window.devicePixelRatio||1, box=canvas.getBoundingClientRect();
  canvas.width=box.width*dpr; canvas.height=box.height*dpr; ctx.scale(dpr,dpr); ctx.clearRect(0,0,box.width,box.height);
  if(!dados.length) return; const margem={e:48,d:15,t:18,b:34}, w=box.width-margem.e-margem.d,h=box.height-margem.t-margem.b;
  const max=Math.max(...dados.map(d=>d.valor),1)*1.12, min=0; ctx.font="9px Inter";ctx.fillStyle="#98a2b2";ctx.textAlign="right";
  for(let i=0;i<=4;i++){const y=margem.t+h-(h*i/4),valor=min+(max-min)*i/4;ctx.beginPath();ctx.strokeStyle="#e9edf2";ctx.setLineDash([4,4]);ctx.moveTo(margem.e,y);ctx.lineTo(margem.e+w,y);ctx.stroke();ctx.fillText(formatarCompacto(valor),margem.e-8,y+3)}
  ctx.setLineDash([]); const pontos=dados.map((d,i)=>({x:margem.e+(dados.length===1?w/2:i*w/(dados.length-1)),y:margem.t+h-(d.valor/max)*h}));
  const grad=ctx.createLinearGradient(0,margem.t,0,margem.t+h);grad.addColorStop(0,"rgba(21,94,239,.22)");grad.addColorStop(1,"rgba(21,94,239,0)");
  ctx.beginPath();ctx.moveTo(pontos[0].x,margem.t+h);pontos.forEach(p=>ctx.lineTo(p.x,p.y));ctx.lineTo(pontos.at(-1).x,margem.t+h);ctx.closePath();ctx.fillStyle=grad;ctx.fill();
  ctx.beginPath();pontos.forEach((p,i)=>i?ctx.lineTo(p.x,p.y):ctx.moveTo(p.x,p.y));ctx.strokeStyle="#155eef";ctx.lineWidth=2.5;ctx.stroke();
  pontos.forEach((p,i)=>{ctx.beginPath();ctx.arc(p.x,p.y,4,0,Math.PI*2);ctx.fillStyle="#fff";ctx.fill();ctx.strokeStyle="#155eef";ctx.lineWidth=2;ctx.stroke();ctx.fillStyle="#7d8899";ctx.textAlign="center";ctx.font="8px Inter";ctx.fillText(dados[i].label,p.x,margem.t+h+20)});
}

async function importarPlanilha(evento){
  const arquivo=evento.target.files[0]; if(!arquivo) return;
  if(typeof XLSX==="undefined"){mostrarToast("Não foi possível carregar o leitor de Excel. Verifique a internet.");return;}
  try{
    elementos.statusArquivo.textContent="Lendo planilha..."; const buffer=await arquivo.arrayBuffer();
    const workbook=XLSX.read(buffer,{type:"array",cellDates:true}); const novoBanco={};
    configuracoes.forEach(config=>{const nome=workbook.SheetNames.find(n=>normalizar(n)===normalizar(config.aba));if(!nome)return;
      const matriz=XLSX.utils.sheet_to_json(workbook.Sheets[nome],{header:1,raw:true,defval:null}); if(matriz.length<2)return;
      const colunas=matriz[0].map(v=>String(v??"").trim()); const linhas=matriz.slice(1).filter(l=>l[0]!=null).map(l=>({data:converterData(l[0]),valores:l.slice(1,colunas.length)})).filter(l=>l.data&&!Number.isNaN(l.data));
      novoBanco[config.aba]={colunas,linhas};
    });
    if(Object.keys(novoBanco).length<1)throw new Error("Abas não reconhecidas");
    banco={...banco,...novoBanco}; elementos.nomeArquivo.textContent=arquivo.name;elementos.statusArquivo.textContent=`${Object.keys(novoBanco).length} abas carregadas`;
    elementos.rodape.textContent="● Planilha importada"; preencherCompetencias();renderizarTudo();mostrarToast("Planilha importada com sucesso!");
  }catch(erro){console.error(erro);elementos.statusArquivo.textContent="Erro na importação";mostrarToast("Não foi possível ler essa planilha.");}
}

function exportarResumo(){
  const data=lerChaveData(elementos.competencia.value), linhas=[["Indicador","Competência","Total","Composição"]];
  configuracoes.forEach(config=>{const aba=banco[config.aba],registro=obterRegistro(config,data);linhas.push([config.nome,registro?formatarMes(registro.data):"",totalRegistro(config,registro),registro?aba.colunas.slice(1).map((n,i)=>`${n}: ${registro.valores[i]??0}`).join(" | "):""])});
  const csv="\uFEFF"+linhas.map(l=>l.map(v=>`"${String(v).replaceAll('"','""')}"`).join(";")).join("\n"),blob=new Blob([csv],{type:"text/csv;charset=utf-8"}),link=document.createElement("a");
  link.href=URL.createObjectURL(blob);link.download=`resumo-kpi-${elementos.competencia.value}.csv`;link.click();URL.revokeObjectURL(link.href);mostrarToast("Resumo exportado!");
}

function converterData(valor){if(valor instanceof Date)return new Date(valor.getFullYear(),valor.getMonth(),1,12);if(typeof valor==="number"){const d=new Date(Date.UTC(1899,11,30)+valor*86400000);return new Date(d.getUTCFullYear(),d.getUTCMonth(),1,12)}const d=new Date(valor);return Number.isNaN(d)?null:new Date(d.getFullYear(),d.getMonth(),1,12)}
function numero(v){const n=Number(v);return Number.isFinite(n)?n:0}function chaveData(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`}function lerChaveData(c){const [a,m]=c.split("-").map(Number);return new Date(a,m-1,1,12)}
function formatarMes(d){return new Intl.DateTimeFormat("pt-BR",{month:"long",year:"numeric"}).format(d).replace(/^./,c=>c.toUpperCase())}function formatarMesCurto(d){return new Intl.DateTimeFormat("pt-BR",{month:"short",year:"2-digit"}).format(d).replace(" de ","/")}
function formatarNumero(v){return new Intl.NumberFormat("pt-BR").format(v)}function formatarCompacto(v){return new Intl.NumberFormat("pt-BR",{notation:"compact",maximumFractionDigits:1}).format(v)}function normalizar(t){return String(t).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().replace(/[^a-z0-9]/g,"")}
let timerToast;function mostrarToast(msg){clearTimeout(timerToast);elementos.toast.textContent=msg;elementos.toast.classList.add("visivel");timerToast=setTimeout(()=>elementos.toast.classList.remove("visivel"),3000)}
