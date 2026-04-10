console.log("App iniciada");
async function obterListaPedidosSegura(){

    const pedidos = await obterPedidosFaturas();

    if(!pedidos || !pedidos.value){
        console.warn("Sem pedidos ou erro no Graph");
        return [];
    }

    return pedidos.value;

}
async function uploadPdfSharePoint(ficheiro){

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const nomeFicheiro = ficheiro.name;

    /* caminho onde vai guardar */
    const caminho = `DocumentosAprovacao/${nomeFicheiro}`;

    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${caminho}:/content`;

    const resp = await fetch(url, {
        method: "PUT",
        headers: {
            Authorization: "Bearer " + token,
            "Content-Type": ficheiro.type
        },
        body: ficheiro
    });

    if(!resp.ok){
    const erroTexto = await resp.text();
    console.error("ERRO UPLOAD PDF:", erroTexto);
    alert("Erro no upload do PDF");
    throw new Error("Erro upload");
}

    const data = await resp.json();

    console.log("UPLOAD GRAPH:", data);

    return {
    webUrl: data.webUrl,
    name: data.name,
    id: data.id,
    downloadUrl: `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${data.id}/content`
};

}
const appDiv = document.getElementById("app");

if(appDiv){
    appDiv.innerHTML = `
    <button onclick="login()">Iniciar sessão Microsoft</button>
    `;
}

function mostrarDashboard(utilizador){

    document.getElementById("app").innerHTML = `

        <h2>Dashboard</h2>

        <p>Bem-vindo ${utilizador.name} | Monte do Pasto</p>

        <button onclick="testarLigacao()">Testar ligação</button>

    `;

}



async function testarLigacao(){

    console.log("A testar ligação ao Microsoft Graph");

    const utilizador = await testarGraph();
 
    console.log("Utilizador Graph:", utilizador);
document.getElementById("utilizador").innerText =
"Bem-vindo " + utilizador.displayName;
    const site = await obterSiteApp();
    console.log("Site da aplicação:", site);

    const listas = await obterListas();
    console.log("Listas do site:", listas);

    const pedidos = await obterPedidosFaturas();
    console.log("Pedidos da lista:", pedidos);

}
async function carregarDashboard(){
const perfil = await obterPerfilUtilizador();

console.log("Perfil do utilizador:", perfil);

const btnFatura = document.getElementById("btnNovaFatura");
if(btnFatura){
    btnFatura.onclick = () => {
        window.location.href = "nova-fatura.html";
    };
}

const btnDespesa = document.getElementById("btnNovaDespesa");
const btnAdmin = document.getElementById("btnAdmin");
const menuPedidos = document.getElementById("menuPedidos");

const utilizador = await testarGraph();
const email = utilizador.mail || utilizador.userPrincipalName;

/* METE AQUI O TEU EMAIL */
const adminEmail = "TEU_EMAIL_AQUI";

/* CONTROLO POR PERFIL */

/* NOVA FATURA */
if(btnFatura){
    if(perfil === "Admin" || perfil === "GestorFaturas"){
        btnFatura.style.display = "inline-block";
    } else {
        btnFatura.style.display = "none";
    }
}

/* NOVA DESPESA */
if(btnDespesa){
    if(perfil === "Admin" || perfil === "GestorFaturas" || perfil === "Utilizador"){
        btnDespesa.style.display = "inline-block";
    } else {
        btnDespesa.style.display = "none";
    }
}

/* ADMIN */
if(btnAdmin){
    if(perfil === "Admin"){
        btnAdmin.style.display = "inline-block";
    } else {
        btnAdmin.style.display = "none";
    }
}

/* PEDIDOS */
if(menuPedidos){
    menuPedidos.style.display = "none";
}
const lista = await obterListaPedidosSegura();
const paraMim = [];
const outros = [];

lista.forEach(p => {

const f = p.fields;

if(
f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
)
){
paraMim.push(p);
}else{
outros.push(p);
}

});

const meusPendentes = lista.filter(p => {

const f = p.fields;

return f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
);

});

lista.sort((a,b)=>{

const estadoA = a.fields.EstadoPedido;
const estadoB = b.fields.EstadoPedido;

if(estadoA === "Pendente" && estadoB !== "Pendente") return -1;
if(estadoA !== "Pendente" && estadoB === "Pendente") return 1;

return 0;

});
    const total = lista.length;

    let pendentes = 0;
    let aprovados = 0;
    let rejeitados = 0;

    lista.forEach(p => {

        const estado = p.fields.EstadoPedido;

        if(estado === "Pendente") pendentes++;
        if(estado === "Aprovado") aprovados++;
        if(estado === "Rejeitado") rejeitados++;

    });

    document.getElementById("totalPedidos").innerText = total;
    document.getElementById("pendentes").innerText = pendentes;
    document.getElementById("aprovados").innerText = aprovados;
    document.getElementById("rejeitados").innerText = rejeitados;
    document.getElementById("meusPendentes").innerText = meusPendentes.length;
const tabela = document.getElementById("listaPedidos");

tabela.innerHTML = "";

lista.forEach(p => {

    const f = p.fields;

    const linha = document.createElement("tr");

if(
f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
)
){
linha.classList.add("linha-para-aprovar");
}
linha.onclick = () => {

if(f.PdfUrl){
    window.open(f.PdfUrl, "_blank");
}else{
    alert("Este pedido não tem PDF associado.");
}

};
const aprovadores = [
    f.Aprovador1Email,
    f.Aprovador2Email
]
.filter(a => a && a !== "")
.map(a => a.split("@")[0]) // só nome antes do email
.join(" / ");
linha.innerHTML = `
<td>
<input type="checkbox" class="checkPedido" value="${p.id}" onclick="event.stopPropagation()">
</td>
<td>${f.NumeroInterno || ""}</td>
<td>${f.Fornecedor || ""}</td>
<td>${f.NumeroFaturaOriginal || ""}</td>
<td>${aprovadores || "-"}</td>
<td>${badgeEstado(f.EstadoPedido)}</td>
`;

    tabela.appendChild(linha);

});
}
window.addEventListener("load", async () => {

    console.log("A iniciar app com login automático");

    try{

        let utilizador;

try{
    utilizador = await testarGraph(); // tenta usar sessão existente
}catch(e){
    console.log("Sem sessão, a iniciar login...");
    await login();
    return; // ⚠️ IMPORTANTE: parar aqui para evitar erro
}

    }catch(e){
        console.warn("Utilizador já autenticado ou erro silencioso");
    }

    const tabelaDashboard = document.getElementById("listaPedidos");
    if(tabelaDashboard){
        await carregarDashboard();
    }

    const selectAprovador = document.getElementById("aprovador1");
    if(selectAprovador){
        await carregarAprovadores();
    }

    const tabelaAprovacoes = document.getElementById("listaAprovacoes");
    if(tabelaAprovacoes){
        await carregarAprovacoes();
    }

});
window.guardarFatura = async function guardarFatura(){

    const fornecedor = document.getElementById("fornecedor").value;
    const numeroFatura = document.getElementById("numeroFatura").value;
    const numeroNormalizado = numeroFatura.toUpperCase().trim();
    const numeroInterno = await gerarNumeroInterno();
    const aprovador1 = document.getElementById("aprovador1")?.value || "";
const aprovador2 = document.getElementById("aprovador2")?.value || "";
if(!aprovador1){
alert("Tem de selecionar um aprovador.");
return;
}
const duplicado = await verificarFaturaDuplicada(numeroNormalizado);

if(duplicado){

    alert("Esta fatura já existe no sistema.");

    return;

}
    const ficheiro = document.getElementById("ficheiroPDF").files[0];
let pdfUrl = "";
let pdfNome = "";
let pdfId = "";
let pdfDownloadUrl = "";

if(ficheiro){

    const upload = await uploadPdfSharePoint(ficheiro);

pdfUrl = upload.webUrl;
pdfNome = upload.name;
pdfId = upload.id;
pdfDownloadUrl = upload.downloadUrl || "";
console.log("UPLOAD ID:", pdfId);
console.log("UPLOAD DOWNLOAD URL:", pdfDownloadUrl);

}
    const utilizador = await testarGraph();

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const body = {
        fields: {

            Title: fornecedor,
            
            NumeroInterno: numeroInterno,

            TipoDocumento: "Fatura",

            Fornecedor: fornecedor,

            NumeroFaturaOriginal: numeroFatura,

            NumeroFaturaNormalizado: numeroNormalizado,

            CriadoPorNome: utilizador.displayName,

            CriadoPorEmail: utilizador.mail || utilizador.userPrincipalName,

            DataCriacaoPedido: new Date().toISOString(),

            EstadoPedido: "Pendente",
            Aprovador1Email: aprovador1,
Aprovador2Email: aprovador2,
PdfUrl: pdfUrl,
PdfNomeFicheiro: pdfNome,
PdfDriveItemId: pdfId,
PdfDownloadUrl: pdfDownloadUrl
        }
    };

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items`,
        {
            method: "POST",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        }
    );

    const resultado = await resposta.json();

console.log("Resposta Graph:", resultado);

if(!resposta.ok){
    alert("Erro ao gravar no SharePoint");
    return;
}

alert("Fatura registada com sucesso");

window.location.href = "dashboard.html";

};
async function carregarPedido(){

    const params = new URLSearchParams(window.location.search);
    const id = params.get("id");

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const dados = await resp.json();

    const f = dados.fields;
/* TIMELINE */

const timeline = document.getElementById("timelinePedido");

if(timeline){

timeline.innerHTML = "";

/* criado */

timeline.innerHTML += `
<li class="timeline-criado">
Pedido criado por ${f.CriadoPorNome || "utilizador"}
</li>
`;

/* enviado */

timeline.innerHTML += `
<li>
Enviado para aprovação
</li>
`;

/* aprovado */

if(f.EstadoPedido === "Aprovado"){

timeline.innerHTML += `
<li class="timeline-aprovado">
Aprovado
</li>
`;

}

/* rejeitado */

if(f.EstadoPedido === "Rejeitado"){

timeline.innerHTML += `
<li class="timeline-rejeitado">
Rejeitado
</li>
`;

if(f.Observacoes){

timeline.innerHTML += `
<li>
Motivo: ${f.Observacoes}
</li>
`;

}

}

}
   document.getElementById("dadosPedido").innerHTML = `

<h3>${f.Fornecedor}</h3>

<p><b>Nº Interno:</b> ${f.NumeroInterno || "-"}</p>

<p><b>Nº Fatura:</b> ${f.NumeroFaturaOriginal}</p>

<p><b>Valor:</b> ${f.ValorDocumento} €</p>

<p><b>Estado:</b> ${badgeEstado(f.EstadoPedido)}</p>

`;

    document.getElementById("dadosPedido").innerHTML += `
<br>
<button onclick="abrirPdf('${f.PdfUrl}')">
Ver PDF da fatura
</button>
`;

}
if(window.location.pathname.includes("ver-pedido.html")){

testarGraph().then(()=>{
carregarPedido();
});

}
async function atualizarEstadoPedido(novoEstado, comentario=""){

const params = new URLSearchParams(window.location.search);
const id = params.get("id");

const token = await getAccessToken();

const site = await obterSiteApp();
const siteId = site.id;

const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

/* buscar pedido */

const respPedido = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`,
{
headers:{ Authorization:"Bearer "+token }
}
);

const dados = await respPedido.json();
const f = dados.fields;

/* carimbar PDF */

const pdfCarimbado = await carimbarPdf(id, novoEstado, comentario);

/* upload novo PDF */

const nomeFicheiro = `fatura_${Date.now()}.pdf`;

const ficheiro = new File(
    [pdfCarimbado],
    nomeFicheiro,
    { type: "application/pdf" }
);

const upload = await uploadPdfSharePoint(ficheiro);

const novoPdfUrl = upload.webUrl;
const novoNome = upload.name;
const novoPdfId = upload.id;
const novoPdfDownloadUrl = upload.downloadUrl || "";

/* atualizar estado + novo PDF */

const body = {
EstadoPedido: novoEstado,
Observacoes: comentario,
PdfUrl: novoPdfUrl,
PdfNomeFicheiro: novoNome,
PdfDriveItemId: novoPdfId,
PdfDownloadUrl: novoPdfDownloadUrl
};

const resp = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}/fields`,
{
method:"PATCH",
headers:{
Authorization:"Bearer "+token,
"Content-Type":"application/json"
},
body:JSON.stringify(body)
}
);

if(!resp.ok){
alert("Erro ao atualizar estado");
return;
}

alert("Pedido "+novoEstado);

window.location.href="dashboard.html";

}
function badgeEstado(estado){

if(estado === "Pendente"){
return '<span class="badge pendente">Pendente</span>';
}

if(estado === "Aprovado"){
return '<span class="badge aprovado">Aprovado</span>';
}

if(estado === "Rejeitado"){
return '<span class="badge rejeitado">Rejeitado</span>';
}

return estado;

}
async function aprovarSelecionados(){

const selecionados = document.querySelectorAll(".checkPedido:checked");

for(const check of selecionados){

await atualizarEstadoPedidoPorId(check.value,"Aprovado");

}

alert("Pedidos aprovados");

location.reload();

}
async function rejeitarSelecionados(){

const selecionados = document.querySelectorAll(".checkPedido:checked");

for(const check of selecionados){

await atualizarEstadoPedidoPorId(check.value,"Rejeitado");

}

alert("Pedidos rejeitados");

location.reload();

}

const checkTodos = document.getElementById("checkTodos");

if(checkTodos){

checkTodos.addEventListener("change", function(){

const checks = document.querySelectorAll(".checkPedido");

checks.forEach(c => {

c.checked = this.checked;

});

});

}
function diasParaVencimento(data){

if(!data) return null;

const hoje = new Date();
const venc = new Date(data);

const diff = venc - hoje;

return Math.ceil(diff / (1000*60*60*24));

}
function dividirTexto(texto, maxChars){

const palavras = texto.split(" ");
const linhas = [];
let linha = "";

for(let p of palavras){

if((linha + p).length > maxChars){
    linhas.push(linha.trim());
    linha = p + " ";
}else{
    linha += p + " ";
}

}

if(linha) linhas.push(linha.trim());

return linhas;

}
async function carimbarPdf(itemId, estado, comentarioExtra = ""){

const { PDFDocument, rgb, StandardFonts } = PDFLib;

const token = await getAccessToken();

const site = await obterSiteApp();
const siteId = site.id;

const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

/* ir buscar dados do item */
const respItem = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${itemId}?expand=fields`,
{
    headers: {
        Authorization: "Bearer " + token
    }
});

if(!respItem.ok){
    alert("Erro ao obter dados do pedido");
    throw new Error("Erro ao obter item");
}

const dataItem = await respItem.json();

const pdfDownloadUrl = dataItem.fields.PdfDownloadUrl;
const pdfDriveItemId = dataItem.fields.PdfDriveItemId;
const comentario = comentarioExtra || dataItem.fields.Observacoes || "";
console.log("PDF DOWNLOAD URL:", pdfDownloadUrl);
console.log("PDF DRIVE ITEM ID:", pdfDriveItemId);

let urlPdfFinal = pdfDownloadUrl;

if(!urlPdfFinal && pdfDriveItemId){
    urlPdfFinal = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${pdfDriveItemId}/content`;
}

if(!urlPdfFinal){
    alert("Este pedido não tem PdfDownloadUrl nem PdfDriveItemId gravados.");
    throw new Error("URL PDF em falta");
}

/* buscar PDF bruto */
const resp = await fetch(urlPdfFinal, {
    headers: {
        Authorization: "Bearer " + token
    }
});

if(!resp.ok){
    alert("Erro ao carregar PDF");
    throw new Error("Erro PDF");
}

const bytes = await resp.arrayBuffer();

/* carregar PDF */
const pdfDoc = await PDFDocument.load(bytes);

const pages = pdfDoc.getPages();
const page = pages[0];

const font = await pdfDoc.embedFont(StandardFonts.HelveticaBold);

const utilizador = await testarGraph();

let texto =
estado + "\n" +
utilizador.displayName + "\n" +
new Date().toLocaleString("pt-PT");

if(estado === "Rejeitado" && comentario){
    texto += "\nMotivo: " + comentario;
}

/* escrever no PDF */
const { width, height } = page.getSize();
const corPrincipal = estado === "Aprovado"
? rgb(0,0.6,0)
: rgb(0.8,0,0);
/* posição do carimbo */
const boxWidth = 260;
const margin = 20;

/* texto dividido */
let linhas = [];

const partesTexto = texto.split("\n");

partesTexto.forEach(linha => {

    if(linha.includes("Motivo:")){
        
        const textoLimpo = linha.replace("Motivo: ", "");
        const partes = dividirTexto(textoLimpo, 35);

        linhas.push("Motivo:");

        partes.forEach(p => {
            linhas.push(p);
        });

    } else {
        linhas.push(linha);
    }

});

/* AGORA SIM: calcular altura */
const baseHeight = 60;
const alturaLinha = 16;
const boxHeight = baseHeight + (linhas.length * alturaLinha);
/* posição correta */

const x = width - boxWidth - margin;
const y = height - boxHeight - margin;

/* fundo */
page.drawRectangle({
x, y,
width: boxWidth,
height: boxHeight,
color: rgb(1,1,1),
opacity: 0.7
});

/* borda */
page.drawRectangle({
x, y,
width: boxWidth,
height: boxHeight,
borderWidth: 2,
borderColor: corPrincipal
});
/* escrever linhas */
linhas.forEach((linha, i) => {
page.drawText(linha,{
    x: x + 10,
    y: y + boxHeight - 20 - (i * 16),
    size: i === 0 ? 18 : 11,
    font: font,
    color: i === 0 ? corPrincipal : rgb(0.2,0.2,0.2)
});
});

/* guardar */
const pdfFinal = await pdfDoc.save();

return pdfFinal;

}

function atualizarDataHora(){

const agora = new Date();

const data = agora.toLocaleDateString("pt-PT",{
weekday:"short",
day:"2-digit",
month:"short",
year:"numeric"
});

const hora = agora.toLocaleTimeString("pt-PT",{
hour:"2-digit",
minute:"2-digit"
});

const el = document.getElementById("dataHora");

if(el){
el.innerText = data + " • " + hora;
}

}

setInterval(atualizarDataHora,1000);
atualizarDataHora();
if(typeof lucide !== "undefined"){
lucide.createIcons();
}
function mostrarAreaRejeicao(){

document.getElementById("areaRejeicao").style.display = "flex";

}
async function confirmarRejeicao(){

const comentario = document
.getElementById("Observacoes")
.value
.trim();

if(!comentario){

alert("Tem de escrever um comentário para rejeitar.");

return;

}

await atualizarEstadoPedido("Rejeitado", comentario);

}
function rejeitarPedido(){

mostrarAreaRejeicao();

}
async function aprovarPedido(){

await atualizarEstadoPedido("Aprovado");

}
function mostrarSegundoAprovador(){
  document.getElementById("segundoAprovadorBox").style.display = "block";
}
async function obterAprovadores(){

const token = await getAccessToken();
const site = await obterSiteApp();

const siteId = site.id;

/* MUITO IMPORTANTE: nome da lista */
const listaNome = "AprovadoresApp";

const resp = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items?expand=fields`,
{
headers:{ Authorization:"Bearer " + token }
}
);

const data = await resp.json();

console.log("Aprovadores:", data);

return data.value.map(item => ({
nome: item.fields.NomeAprovador,
email: item.fields.EmailAprovador
}));

}
async function carregarAprovadores(){

const aprovadores = await obterAprovadores();

const select1 = document.getElementById("aprovador1");
const select2 = document.getElementById("aprovador2");

if(!select1) return;

select1.innerHTML = `<option value="">Selecionar</option>`;

if(select2){
select2.innerHTML = `<option value="">Selecionar</option>`;
}

aprovadores.forEach(a => {

const opt1 = document.createElement("option");
opt1.value = a.email;
opt1.textContent = a.nome;
select1.appendChild(opt1);

if(select2){
const opt2 = document.createElement("option");
opt2.value = a.email;
opt2.textContent = a.nome;
select2.appendChild(opt2);
}

});

}
async function carregarAprovacoes(){

const perfil = await obterPerfilUtilizador();

const btnFatura = document.getElementById("btnNovaFatura");
const btnDespesa = document.getElementById("btnNovaDespesa");
const menuAdmin = document.getElementById("menuAdmin");

if(btnFatura && btnDespesa){

if(perfil === "Admin"){
btnFatura.style.display = "inline-block";
btnDespesa.style.display = "inline-block";
if(menuAdmin) menuAdmin.style.display = "flex";
}
else if(perfil === "GestorFaturas"){
btnFatura.style.display = "inline-block";
btnDespesa.style.display = "inline-block";
if(menuAdmin) menuAdmin.style.display = "none";
}
else{
btnFatura.style.display = "none";
btnDespesa.style.display = "inline-block";
if(menuAdmin) menuAdmin.style.display = "none";
}

}

const utilizador = await testarGraph();
const email = utilizador.mail || utilizador.userPrincipalName;

const lista = await obterListaPedidosSegura();

const pendentesParaMim = lista.filter(p => {

const f = p.fields;

return f.EstadoPedido === "Pendente" &&
(
f.Aprovador1Email === email ||
f.Aprovador2Email === email
);

});

document.getElementById("totalAprovacoesPendentes").innerText = pendentesParaMim.length;

const tabela = document.getElementById("listaAprovacoes");
tabela.innerHTML = "";

pendentesParaMim.forEach(p => {

const f = p.fields;

const linha = document.createElement("tr");

linha.innerHTML = `
<td>${f.NumeroInterno || ""}</td>
<td>${f.Fornecedor || ""}</td>
<td>${f.NumeroFaturaOriginal || ""}</td>
<td>${badgeEstado(f.EstadoPedido)}</td>
<td>

${
f.TipoDocumento === "KMS"
? `<button class="btn-ver" onclick="verPdfKM('${p.id}')">📄 PDF</button>`
: `<button class="btn-ver" onclick="window.open('${f.PdfUrl}','_blank')">📄 PDF</button>`
}
📄 PDF
</button>

<button class="btn-aprovar" onclick="aprovarDireto('${p.id}')">
✔
</button>

<button class="btn-rejeitar" onclick="mostrarRejeicaoLinha('${p.id}')">
✖
</button>

<div id="rej-${p.id}" class="box-rejeicao-linha">
<textarea id="obs-${p.id}" placeholder="Motivo da rejeição..."></textarea>

<button onclick="confirmarRejeicaoLinha('${p.id}')">
Confirmar
</button>
</div>

</td>
`;

tabela.appendChild(linha);

});

const pesquisa = document.getElementById("pesquisaAprovacoes");

if(pesquisa){

pesquisa.addEventListener("keyup", function(){

const termo = this.value.toLowerCase();

document.querySelectorAll("#listaAprovacoes tr").forEach(linha => {

linha.style.display = linha.innerText.toLowerCase().includes(termo) ? "" : "none";

});

});

}

}
async function aprovarDireto(id){

await atualizarEstadoPedidoPorId(id, "Aprovado");

alert("Pedido aprovado");

location.reload();

}
function mostrarRejeicaoLinha(id){

document.getElementById("rej-" + id).style.display = "block";

}
async function confirmarRejeicaoLinha(id){

const obs = document.getElementById("obs-" + id).value.trim();

if(!obs){
alert("Tem de escrever um motivo.");
return;
}

await atualizarEstadoPedidoPorId(id, "Rejeitado", obs);

alert("Pedido rejeitado");

location.reload();

}
async function atualizarEstadoPedidoPorId(id, novoEstado, comentario=""){

const token = await getAccessToken();
const site = await obterSiteApp();
const siteId = site.id;
const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

/* buscar pedido */
const respPedido = await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}?expand=fields`,
{
headers:{ Authorization:"Bearer "+token }
}
);

const dados = await respPedido.json();
const f = dados.fields;

/* carimbar PDF */
    console.log("PDF ID:", f.PdfDriveItemId);
const pdfCarimbado = await carimbarPdf(id, novoEstado, comentario);

/* criar novo ficheiro */
const ficheiro = new File(
[pdfCarimbado],
`fatura_${id}_${Date.now()}.pdf`,
{ type: "application/pdf" }
);

/* upload */
const upload = await uploadPdfSharePoint(ficheiro);
const novoPdfUrl = upload.webUrl;
const novoPdfId = upload.id;
const novoPdfDownloadUrl = upload.downloadUrl || "";
const novoPdfNome = upload.name;
/* update SharePoint */
await fetch(
`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items/${id}/fields`,
{
method:"PATCH",
headers:{
Authorization:"Bearer "+token,
"Content-Type":"application/json"
},
body: JSON.stringify({
EstadoPedido: novoEstado,
Observacoes: comentario,
PdfUrl: novoPdfUrl,
PdfNomeFicheiro: novoPdfNome,
PdfDriveItemId: novoPdfId,
PdfDownloadUrl: novoPdfDownloadUrl
})
}
);

}
