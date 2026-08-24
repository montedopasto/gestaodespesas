/* =============================
   TABELA KM
============================= */

let notaCorrecaoAtual = null;

function notaTemCorrecao(campos){
    return campos.EmCorrecao === true ||
        campos.EmCorrecao === 1 ||
        String(campos.EmCorrecao).toLowerCase() === "true";
}

function addLinhaKM(){

    const tbody = document.getElementById("linhasKM");

    const tr = document.createElement("tr");

    tr.innerHTML = `
    <td><input type="date" class="data"></td>
    <td><input type="text" class="origem"></td>
    <td><input type="text" class="destino"></td>
    <td><input type="text" class="justificacao"></td>
    <td><input type="number" class="kms" oninput="calcularKM()"></td>
    <td><button onclick="removerLinha(this)">X</button></td>
    `;

    tbody.appendChild(tr);
if (window.lucide) {
    lucide.createIcons();
}
}


/* =============================
   REMOVER LINHA
============================= */

function removerLinha(btn){
    btn.closest("tr").remove();
    calcularKM();
}


/* =============================
   CALCULAR TOTAIS
============================= */

function calcularKM(){

    let totalKMs = 0;

    document.querySelectorAll(".kms").forEach(input => {
        totalKMs += Number(input.value) || 0;
    });

    const elTotal = document.getElementById("totalKMs");
    if(elTotal){
        elTotal.innerText = totalKMs;
    }

    const valorKM = Number(document.getElementById("valorKM")?.value) || 0;

    const totalFinal = totalKMs * valorKM;

    const elFinal = document.getElementById("totalFinalKM");
    if(elFinal){
        elFinal.innerText = totalFinal.toFixed(2) + " €";
    }

}
/* =============================
   GUARDAR DESPESA KM
============================= */

async function configurarCampoCriarEmNomeDe(){

    const campo = document.getElementById("campoCriarEmNomeDe");
    if(!campo) return;

    try{
        const perfil = await obterPerfilUtilizador();
        campo.style.display =
            perfil === "Admin" || perfil === "GestorFaturas"
                ? "block"
                : "none";
    }catch(erro){
        console.error("Não foi possível validar o campo 'em nome de':", erro);
        campo.style.display = "none";
    }
}

async function obterNomeColaboradorDaNota(utilizador){

    const perfil = await obterPerfilUtilizador();
    const podeCriarEmNomeDe =
        perfil === "Admin" || perfil === "GestorFaturas";

    if(!podeCriarEmNomeDe){
        return utilizador.displayName;
    }

    const nomeIndicado =
        document.getElementById("nomeColaborador")?.value.trim() || "";

    return nomeIndicado || utilizador.displayName;
}

async function obterTodasNotasDespesa(token, siteId){

    let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/NotasDespesa/items?expand=fields`;
    const items = [];

    while(url){
        const resp = await fetch(url, {
            headers:{ Authorization:"Bearer " + token }
        });

        if(!resp.ok){
            throw new Error("Não foi possível obter a numeração das notas.");
        }

        const data = await resp.json();
        items.push(...(data.value || []));
        url = data["@odata.nextLink"] || "";
    }

    return items;
}

async function gerarNumeroNota(token, siteId){

    const ano = new Date().getFullYear();
    const prefixo = `ND-${ano}-`;
    const items = await obterTodasNotasDespesa(token, siteId);

    const ultimoNumero = items.reduce((maior, item) => {
        const numero = item.fields?.NumeroNota || "";
        if(!numero.startsWith(prefixo)) return maior;

        const sequencia = Number(numero.slice(prefixo.length));
        return Number.isInteger(sequencia) && sequencia > maior
            ? sequencia
            : maior;
    }, 0);

    return prefixo + String(ultimoNumero + 1).padStart(6, "0");
}

function escaparHtmlEmail(valor){
    return String(valor ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatarEuroEmail(valor){
    return Number(valor || 0).toLocaleString("pt-PT", {
        style:"currency",
        currency:"EUR"
    });
}

function construirEmailBase(titulo, conteudo){
    return `
        <div style="font-family:Arial,sans-serif;color:#1e293b;line-height:1.55;max-width:640px">
            <div style="background:#166534;color:white;padding:18px 22px;border-radius:10px 10px 0 0">
                <h2 style="margin:0;font-size:20px">${escaparHtmlEmail(titulo)}</h2>
            </div>
            <div style="border:1px solid #dbe4dc;border-top:0;padding:22px;border-radius:0 0 10px 10px">
                ${conteudo}
                <p style="margin:24px 0 0;color:#64748b;font-size:12px">Mensagem enviada automaticamente pela aplicação Gestão de Despesas.</p>
            </div>
        </div>`;
}

async function notificarAprovadoresNota(campos){
    const destinatarios = [
        campos.Aprovador1Email,
        campos.Aprovador2Email
    ].filter(Boolean);

    const tipo = campos.TipoDocumento === "KMS"
        ? "Deslocação em KMs"
        : "Outras despesas";
    const link = "https://montedopasto.github.io/gestaodespesas/pages/aprovacoes-despesas.html";
    const html = construirEmailBase("Nova nota para aprovação", `
        <p>Foi submetida uma nova nota de despesa que aguarda a sua decisão.</p>
        <table style="border-collapse:collapse;width:100%">
            <tr><td style="padding:6px 0"><b>Número</b></td><td>${escaparHtmlEmail(campos.NumeroNota)}</td></tr>
            <tr><td style="padding:6px 0"><b>Colaborador</b></td><td>${escaparHtmlEmail(campos.CriadoPorNome)}</td></tr>
            <tr><td style="padding:6px 0"><b>Tipo</b></td><td>${escaparHtmlEmail(tipo)}</td></tr>
            <tr><td style="padding:6px 0"><b>Valor</b></td><td>${formatarEuroEmail(campos.TotalRecebido)}</td></tr>
        </table>
        <p style="margin-top:22px"><a href="${link}" style="background:#166534;color:white;text-decoration:none;padding:11px 18px;border-radius:7px;display:inline-block">Abrir pedido</a></p>
    `);

    await enviarEmailGraph(
        destinatarios,
        `Nota de despesa ${campos.NumeroNota} para aprovação`,
        html
    );
}

async function notificarAutorDecisao(campos, estado, justificacao, decisor){
    const aprovado = estado === "Aprovado";
    const titulo = aprovado
        ? "Nota de despesa aprovada"
        : "Nota de despesa recusada";
    const cor = aprovado ? "#166534" : "#b91c1c";
    const justificacaoHTML = !aprovado
        ? `<p style="background:#fef2f2;border-left:4px solid #b91c1c;padding:12px"><b>Justificação:</b><br>${escaparHtmlEmail(justificacao)}</p>`
        : "";
    const html = construirEmailBase(titulo, `
        <p>A sua nota de despesa foi <b style="color:${cor}">${aprovado ? "aprovada" : "recusada"}</b>.</p>
        <table style="border-collapse:collapse;width:100%">
            <tr><td style="padding:6px 0"><b>Número</b></td><td>${escaparHtmlEmail(campos.NumeroNota || "-")}</td></tr>
            <tr><td style="padding:6px 0"><b>Valor</b></td><td>${formatarEuroEmail(campos.TotalRecebido)}</td></tr>
            <tr><td style="padding:6px 0"><b>Decisão por</b></td><td>${escaparHtmlEmail(decisor)}</td></tr>
        </table>
        ${justificacaoHTML}
    `);

    await enviarEmailGraph(
        [campos.CriadoPorEmail],
        `Nota de despesa ${campos.NumeroNota || ""} ${aprovado ? "aprovada" : "recusada"}`.trim(),
        html
    );
}

async function guardarDespesaKM(){

    const utilizador = await testarGraph();
    const nomeColaborador = await obterNomeColaboradorDaNota(utilizador);
    const token = await getAccessToken();
    const site = await obterSiteApp();

    const siteId = site.id;
    const numeroNota = await gerarNumeroNota(token, siteId);

    const linhas = [];

    const rows = document.querySelectorAll("#linhasKM tr");

for(const tr of rows){

    const data = tr.querySelector(".data")?.value || "";
    const origem = tr.querySelector(".origem")?.value.trim() || "";
    const destino = tr.querySelector(".destino")?.value.trim() || "";
    const justificacao = tr.querySelector(".justificacao")?.value.trim() || "";
    const kms = Number(tr.querySelector(".kms")?.value) || 0;

    if(!data || !origem || !destino || !justificacao || kms <= 0){
    continue; // ignora só essa linha
}

    linhas.push({
        data,
        origem,
        destino,
        justificacao,
        kms
    });

}

    if(linhas.length === 0){
        alert("Tem de inserir pelo menos uma linha.");
        return;
    }

    /* totais */
    let totalKMs = 0;
    linhas.forEach(l => totalKMs += l.kms);
const matriculaVeiculo =
    document.getElementById("matriculaVeiculo")
    ?.value
    .trim();

if(!matriculaVeiculo){
    alert("Tem de indicar a matrícula do veículo.");
    return;
}
    const valorKM = Number(document.getElementById("valorKM").value) || 0;
    const totalRecebido = totalKMs * valorKM;

    /* JSON */
    const linhasJSON = JSON.stringify(linhas);

    const listaNome = "NotasDespesa";
const aprovador1 = document.getElementById("aprovador1")?.value || "";
const aprovador2 = document.getElementById("aprovador2")?.value || "";

if(!aprovador1){
    alert("Tem de selecionar um aprovador.");
    return;
}
    const body = {
        fields: {
    Title: numeroNota + " - Nota KM",
    NumeroNota: numeroNota,
    TipoDocumento: "KMS",
    CriadoPorNome: nomeColaborador,
    CriadoPorEmail: utilizador.mail || utilizador.userPrincipalName,

    MatriculaVeiculo: matriculaVeiculo,

    TotalKMs: totalKMs,
    ValorPorKM: valorKM,
    TotalRecebido: totalRecebido,
    LinhasJSON: linhasJSON,

    Estado: "Pendente",

    Aprovador1Email: aprovador1,
    Aprovador2Email: aprovador2,
}
    };

    const resp = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items`,
    {
        method: "POST",
        headers: {
            Authorization: "Bearer " + token,
            "Content-Type": "application/json"
        },
        body: JSON.stringify(body)
    }
);

const text = await resp.text();
console.log("RESPOSTA:", text);

if(!resp.ok){
    alert("Erro ao guardar nota de despesa");
    return;
}

let avisoEmail = "";
try{
    await notificarAprovadoresNota(body.fields);
}catch(erro){
    console.error("Nota guardada, mas o email falhou:", erro);
    avisoEmail = "\n\nA nota foi guardada, mas o email ao aprovador não foi enviado. " + erro.message;
}

alert("✅ Nota de despesa guardada com sucesso!" + avisoEmail);

window.location.href = "dashboard.html";

}
async function carregarAprovadoresDespesa(){

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
async function obterAprovadores(){

    const token = await getAccessToken();
    const site = await obterSiteApp();

    const siteId = site.id;

    const listaNome = "AprovadoresApp";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();

    console.log("APROVADORES RAW:", data); // 👈 IMPORTANTE

    return data.value.map(item => ({
        nome: item.fields.NomeAprovador,
        email: item.fields.EmailAprovador
    }));

}
   
/* =============================
   APROVAÇÕES DESPESAS
============================= */

async function carregarAprovacoesDespesas(){

    const utilizador = await testarGraph();
    const token = await getAccessToken();
    const site = await obterSiteApp();
    const siteId = site.id;

    const listaNome = "NotasDespesa";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();
    const items = data.value || [];

    const emailUser = utilizador.mail || utilizador.userPrincipalName;

    const filtrados = items.filter(item => {

        const f = item.fields;

        return (
            (f.Aprovador1Email === emailUser || f.Aprovador2Email === emailUser)
            && f.Estado === "Pendente"
        );

    });

    const tbody = document.getElementById("tabelaAprovacoesDespesas");

    tbody.innerHTML = "";

    filtrados.forEach(item => {

        const f = item.fields;

        const tr = document.createElement("tr");

      tr.innerHTML = `
    <td>${f.NumeroNota || "-"}</td>
    <td>${new Date(f.Created).toLocaleDateString("pt-PT")}</td>
    <td>${f.CriadoPorNome}</td>
    <td>${Number(f.TotalRecebido).toFixed(2)} €</td>
    <td>
        <button onclick="verDetalheKM('${item.id}')" class="btn-icon" title="Ver detalhe">
            <i data-lucide="file-text"></i>
        </button>

        <button onclick="aprovarDespesa('${item.id}')" class="btn-icon btn-aprovar" title="Aprovar">
            <i data-lucide="check"></i>
        </button>

        <button onclick="rejeitarDespesa('${item.id}')" class="btn-icon btn-rejeitar" title="Rejeitar">
            <i data-lucide="x"></i>
        </button>
    </td>
`;

tbody.appendChild(tr);

if (window.lucide) {
    lucide.createIcons();
}
    });

}
async function aprovarDespesa(id){

    await atualizarEstadoDespesa(id, "Aprovado", "");
}

async function rejeitarDespesa(id){

    const justificacao = prompt("Indique a justificação da rejeição:");

    if(justificacao === null){
        return; // utilizador cancelou
    }

    if(!justificacao.trim()){
        alert("Tem de indicar uma justificação para rejeitar.");
        return;
    }

    await atualizarEstadoDespesa(id, "Rejeitado", justificacao.trim());
}

async function atualizarEstadoDespesa(id, estado, justificacao = ""){

    const token = await getAccessToken();
    const site = await obterSiteApp();
    const siteId = site.id;

    const listaNome = "NotasDespesa";

    const utilizador = await testarGraph();

    const itemResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items/${id}?expand=fields`,
        { headers:{ Authorization:"Bearer " + token } }
    );

    if(!itemResp.ok){
        alert("Não foi possível confirmar o estado atual da nota.");
        return;
    }

    const itemAtual = await itemResp.json();
    const camposAtuais = itemAtual.fields || {};

    if(camposAtuais.Estado !== "Pendente"){
        alert("Esta nota já foi decidida por outro utilizador.");
        carregarAprovacoesDespesas();
        return;
    }

const body = {
    Estado: estado,
    AprovadoPorNome: utilizador.displayName,
    AprovadoPorEmail: utilizador.mail || utilizador.userPrincipalName
};

    if(estado === "Rejeitado"){
        body.JustificacaoRejeicao = justificacao;
    }

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaNome}/items/${id}/fields`,
        {
            method: "PATCH",
            headers: {
                Authorization: "Bearer " + token,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
        }
    );

    if(!resp.ok){
        alert("Erro ao atualizar o estado.");
        return;
    }

    let avisoEmail = "";
    try{
        await notificarAutorDecisao(
            camposAtuais,
            estado,
            justificacao,
            utilizador.displayName
        );
    }catch(erro){
        console.error("Estado atualizado, mas o email falhou:", erro);
        avisoEmail = "\n\nA decisão foi guardada, mas o email ao autor não foi enviado. " + erro.message;
    }

    alert("Estado atualizado: " + estado + avisoEmail);

    carregarAprovacoesDespesas();
}
/* =============================
   DASHBOARD DESPESAS
============================= */

async function carregarDashboardDespesas(){

    const utilizador = await testarGraph();
    const perfil = await obterPerfilUtilizador();
    const token = await getAccessToken();
    const site = await obterSiteApp();
    const siteId = site.id;

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/NotasDespesa/items?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();
    const todosOsItems = data.value || [];

    const email = String(
        utilizador.mail || utilizador.userPrincipalName || ""
    ).trim().toLowerCase();
    const podeVerTodas = perfil === "Admin" || perfil === "GestorFaturas";

    const items = podeVerTodas
        ? todosOsItems
        : todosOsItems.filter(item => {
            const campos = item.fields || {};
            const emailSubmissor = String(campos.CriadoPorEmail || "").trim().toLowerCase();
            const aprovador1 = String(campos.Aprovador1Email || "").trim().toLowerCase();
            const aprovador2 = String(campos.Aprovador2Email || "").trim().toLowerCase();

            return emailSubmissor === email || aprovador1 === email || aprovador2 === email;
        });

    let total = items.length;
    let pendentes = 0;
    let aprovados = 0;
    let rejeitados = 0;
    let meusPendentes = 0;

    items.forEach(i => {

        const f = i.fields;

        if(f.Estado === "Pendente"){
            pendentes++;

            if(
                String(f.Aprovador1Email || "").trim().toLowerCase() === email ||
                String(f.Aprovador2Email || "").trim().toLowerCase() === email
            ){
                meusPendentes++;
            }
        }

        if(f.Estado === "Aprovado") aprovados++;
        if(f.Estado === "Rejeitado") rejeitados++;

    });

    document.getElementById("totalPedidos").innerText = total;
    document.getElementById("pendentes").innerText = pendentes;
    document.getElementById("aprovados").innerText = aprovados;
    document.getElementById("rejeitados").innerText = rejeitados;
    document.getElementById("meusPendentes").innerText = meusPendentes;
/* =============================
   TABELA
============================= */

const tabela = document.getElementById("tabelaDespesas");

if(!tabela) return;

tabela.innerHTML = "";
items.sort((a,b) => {

    return new Date(b.fields.Created)
        - new Date(a.fields.Created);

});
items.forEach(item => {

    const f = item.fields;
    const emCorrecao = notaTemCorrecao(f);

    const linha = document.createElement("tr");

    const aprovadores = [
        f.Aprovador1Email,
        f.Aprovador2Email
    ]
    .filter(a => a)
    .map(a => a.split("@")[0])
    .join(" / ");

    linha.innerHTML = `
    <td>${f.NumeroNota || "-"}</td>
    <td>${new Date(f.Created).toLocaleDateString("pt-PT")}</td>
    <td>${f.CriadoPorNome}</td>
    <td>${Number(f.TotalRecebido).toFixed(2)} €</td>
    <td>${aprovadores || "-"}</td>
    <td>

    <span style="
        padding:6px 12px;
        border-radius:999px;
        font-weight:700;
        font-size:13px;

        ${
            emCorrecao
            ?
            `
            background:#fff7ed;
            color:#b45309;
            `
            :
            f.Estado === "Aprovado"
            ?
            `
            background:#e8f5e9;
            color:#2e7d32;
            `
            :
            f.Estado === "Rejeitado"
            ?
            `
            background:#ffebee;
            color:#c62828;
            `
            :
            `
            background:#fff8e1;
            color:#b08900;
            `
        }
    ">

        ${emCorrecao ? "Correção solicitada" : f.Estado}

    </span>

</td>
    <td>
        <button onclick="verDetalheKM('${item.id}')" class="btn-icon" title="Ver detalhe">
            <i data-lucide="file-text"></i>
        </button>
    </td>
`;
linha.style.cursor = "pointer";
    /* abrir PDF ao clicar */
    linha.onclick = () => {
    verDetalheKM(item.id);
};

    tabela.appendChild(linha);
if (window.lucide) {
    lucide.createIcons();
}
});
}
window.fecharModalKM = function(){
    document.getElementById("modalKM").style.display = "none";
    notaCorrecaoAtual = null;
}

window.verDetalheKM = async function(id){
document.getElementById("modalKM").dataset.id = id;
    const token = await getAccessToken();
    const site = await obterSiteApp();
    const siteId = site.id;

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/NotasDespesa/items/${id}?expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const data = await resp.json();
    const f = data.fields;
    const utilizador = await testarGraph();
    const emailAtual = (utilizador.mail || utilizador.userPrincipalName || "").toLowerCase();
    const emailSubmissor = String(f.CriadoPorEmail || "").toLowerCase();
    const emCorrecao = notaTemCorrecao(f);
    const podeCorrigir = emCorrecao && emailAtual === emailSubmissor;
    const linhas = JSON.parse(f.LinhasJSON || "[]");

    notaCorrecaoAtual = podeCorrigir
        ? { id:String(id), campos:f, linhas }
        : null;
const zonaEstado = document.getElementById("zonaEstadoPedido");
const carimbo = document.getElementById("carimboEstado");
const caixaJustificacao = document.getElementById("justificacaoRejeicao");
const textoJustificacao = document.getElementById("textoJustificacaoRejeicao");

if(zonaEstado && carimbo && caixaJustificacao && textoJustificacao){

    zonaEstado.style.display = "block";
    caixaJustificacao.style.display = "none";
    textoJustificacao.innerText = "";

    if(emCorrecao){
        carimbo.innerText = "↩ CORREÇÃO SOLICITADA";
        carimbo.style.background = "#b45309";
        carimbo.style.transform = "rotate(-3deg)";
        carimbo.style.boxShadow = "0 4px 10px rgba(0,0,0,0.2)";
    }
    else if(f.Estado === "Aprovado"){

    const dataHora = new Date(f.Modified).toLocaleString("pt-PT");

    carimbo.innerHTML = `
    ✔ APROVADO
    <div style="font-size:11px; margin-top:4px; opacity:0.9;">
        ${dataHora}
    </div>
`;

carimbo.style.transform = "rotate(-5deg)";
carimbo.style.boxShadow = "0 4px 10px rgba(0,0,0,0.2)";

    carimbo.style.background = "#2e7d32";
}
     
    else if(f.Estado === "Rejeitado"){

    const dataHora = new Date(f.Modified).toLocaleString("pt-PT");

    carimbo.innerHTML = `
        ✖ REJEITADO
        <div style="font-size:11px; margin-top:4px; opacity:0.9;">
            ${dataHora}
        </div>
    `;

    carimbo.style.background = "#c62828";
    carimbo.style.transform = "rotate(-5deg)";
    carimbo.style.boxShadow = "0 4px 10px rgba(0,0,0,0.2)";

    if(f.JustificacaoRejeicao && f.JustificacaoRejeicao.trim()){
        caixaJustificacao.style.display = "block";
        textoJustificacao.innerText = f.JustificacaoRejeicao;
    }
}
    else{
        carimbo.innerText = f.Estado || "PENDENTE";
        carimbo.style.background = "#b08900";
    }
}

console.log("LINHAS RAW:", f.LinhasJSON);
const avisoCorrecao = emCorrecao ? `
    <div class="aviso-correcao-valores">
        <b>Esta nota foi devolvida para corrigir os valores.</b><br>
        ${escaparHtmlEmail(f.MotivoCorrecao || "")}
        ${podeCorrigir ? "<br><small>Apenas os campos destacados podem ser alterados.</small>" : ""}
    </div>` : "";
if(f.TipoDocumento === "DESPESA"){

    let html = `

    ${avisoCorrecao}
    <p><b>Número:</b> ${f.NumeroNota || "-"}</p>
    <p><b>Total:</b> ${Number(f.TotalRecebido).toFixed(2)} €</p>

    <br>

    <table style="width:100%">

        <tr>
            <th>Data</th>
            <th>Rubrica</th>
            <th>Descrição</th>
            <th>Valor</th>
            <th>Fatura</th>
        </tr>
    `;

    linhas.forEach((l, indice) => {

        html += `

        <tr>

            <td>${l.data}</td>

            <td>${l.rubrica}</td>

            <td>${l.descricao}</td>

            <td>${podeCorrigir
                ? `<input class="campo-valor-correcao valor-linha-correcao" type="number" min="0.01" step="0.01" data-indice="${indice}" value="${Number(l.valor).toFixed(2)}">`
                : `${Number(l.valor).toFixed(2)} €`}</td>

            <td>

                ${l.ficheiroUrl ? `

                    <a href="${l.ficheiroUrl}"
                       target="_blank"
                       style="
                       background:#2563eb;
                       color:white;
                       padding:6px 10px;
                       border-radius:6px;
                       text-decoration:none;
                       font-size:12px;
                    ">
                        Abrir
                    </a>

                ` : "-"}

            </td>

        </tr>
        `;
    });

    html += `</table>${podeCorrigir ? `<button id="btnGuardarCorrecao" class="btn-guardar-correcao" onclick="guardarCorrecaoValores()">Guardar valores corrigidos</button>` : ""}`;

    document.getElementById("conteudoKM").innerHTML = html;

    document.getElementById("modalKM").style.display = "block";

    return;
}
    let html = `
    ${avisoCorrecao}
    <p><b>Número:</b> ${f.NumeroNota || "-"}</p>
    <p><b>Matrícula:</b> ${f.MatriculaVeiculo || "-"}</p>
    <p><b>Total KMs:</b> ${f.TotalKMs}</p>
    <p><b>Valor/KM:</b> ${podeCorrigir
        ? `<input id="valorKmCorrecao" class="campo-valor-correcao" type="number" min="0.01" step="0.01" value="${Number(f.ValorPorKM).toFixed(2)}">`
        : `${f.ValorPorKM} €`}</p>
    <p><b>Total:</b> ${Number(f.TotalRecebido).toFixed(2)} €</p>

    <br>

    <table style="width:100%">
        <tr>
            <th>Data</th>
            <th>Origem</th>
            <th>Destino</th>
            <th>Justificação</th>
            <th>KMs</th>
        </tr>
`;

    linhas.forEach((l, indice) => {
        html += `
            <tr>
                <td>${l.data}</td>
                <td>${l.origem}</td>
                <td>${l.destino}</td>
                <td>${l.justificacao}</td>
                <td>${podeCorrigir
                    ? `<input class="campo-valor-correcao kms-linha-correcao" type="number" min="0.01" step="0.01" data-indice="${indice}" value="${Number(l.kms)}">`
                    : l.kms}</td>
            </tr>
        `;
    });

    html += `</table>${podeCorrigir ? `<button id="btnGuardarCorrecao" class="btn-guardar-correcao" onclick="guardarCorrecaoValores()">Guardar valores corrigidos</button>` : ""}`;

    document.getElementById("conteudoKM").innerHTML = html;
    document.getElementById("modalKM").style.display = "block";
}

window.guardarCorrecaoValores = async function(){
    if(!notaCorrecaoAtual) return;
    if(!confirm("Guardar os valores corrigidos e devolver a nota para Pagamentos?")) return;

    const botao = document.getElementById("btnGuardarCorrecao");
    botao.disabled = true;
    botao.textContent = "A guardar...";

    try{
        const utilizador = await testarGraph();
        const emailAtual = (utilizador.mail || utilizador.userPrincipalName || "").toLowerCase();
        const token = await getAccessToken();
        const site = await obterSiteApp();
        const url = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/NotasDespesa/items/${notaCorrecaoAtual.id}`;
        const atualResp = await fetch(`${url}?expand=fields`, {
            headers:{ Authorization:"Bearer " + token }
        });
        if(!atualResp.ok) throw new Error(await atualResp.text());
        const atual = await atualResp.json();
        const camposAtuais = atual.fields;

        if(camposAtuais.Estado !== "Aprovado" || !notaTemCorrecao(camposAtuais)){
            throw new Error("Esta nota já não está disponível para correção.");
        }
        if(String(camposAtuais.CriadoPorEmail || "").toLowerCase() !== emailAtual){
            throw new Error("Apenas quem submeteu a nota pode corrigir os valores.");
        }

        const linhas = JSON.parse(camposAtuais.LinhasJSON || "[]");
        const alteracoes = {
            EmCorrecao:false
        };

        if(camposAtuais.TipoDocumento === "DESPESA"){
            document.querySelectorAll(".valor-linha-correcao").forEach(input => {
                const indice = Number(input.dataset.indice);
                const valor = Number(input.value);
                if(!Number.isInteger(indice) || !linhas[indice] || !Number.isFinite(valor) || valor <= 0){
                    throw new Error("Todos os valores devem ser superiores a zero.");
                }
                linhas[indice].valor = valor;
            });
            alteracoes.TotalRecebido = linhas.reduce((soma, linha) => soma + Number(linha.valor || 0), 0);
        }else{
            document.querySelectorAll(".kms-linha-correcao").forEach(input => {
                const indice = Number(input.dataset.indice);
                const kms = Number(input.value);
                if(!Number.isInteger(indice) || !linhas[indice] || !Number.isFinite(kms) || kms <= 0){
                    throw new Error("Todos os KMs devem ser superiores a zero.");
                }
                linhas[indice].kms = kms;
            });
            const valorKM = Number(document.getElementById("valorKmCorrecao")?.value);
            if(!Number.isFinite(valorKM) || valorKM <= 0){
                throw new Error("O valor por KM deve ser superior a zero.");
            }
            const totalKMs = linhas.reduce((soma, linha) => soma + Number(linha.kms || 0), 0);
            alteracoes.TotalKMs = totalKMs;
            alteracoes.ValorPorKM = valorKM;
            alteracoes.TotalRecebido = totalKMs * valorKM;
        }

        alteracoes.LinhasJSON = JSON.stringify(linhas);
        const resp = await fetch(`${url}/fields`, {
            method:"PATCH",
            headers:{
                Authorization:"Bearer " + token,
                "Content-Type":"application/json"
            },
            body:JSON.stringify(alteracoes)
        });
        if(!resp.ok) throw new Error(await resp.text());

        let avisoEmail = "";
        if(camposAtuais.DevolvidoPorEmail){
            try{
                const html = construirEmailBase("Valores corrigidos", `
                    <p>Os valores da nota <b>${escaparHtmlEmail(camposAtuais.NumeroNota || "-")}</b> foram corrigidos.</p>
                    <p><b>Novo total:</b> ${formatarEuroEmail(alteracoes.TotalRecebido)}</p>
                    <p>A nota regressou diretamente à área de Pagamentos.</p>
                `);
                await enviarEmailGraph(
                    [camposAtuais.DevolvidoPorEmail],
                    `Nota ${camposAtuais.NumeroNota || ""} corrigida`.trim(),
                    html
                );
            }catch(erro){
                console.error("Valores corrigidos, mas o email falhou:", erro);
                avisoEmail = "\n\nOs valores foram guardados, mas o email ao responsável pelo pagamento não foi enviado.";
            }
        }

        notaCorrecaoAtual = null;
        fecharModalKM();
        await carregarDashboardDespesas();
        alert("Valores corrigidos. A nota voltou para Pagamentos." + avisoEmail);
    }catch(erro){
        console.error("Erro ao guardar correção:", erro);
        alert(erro.message || "Não foi possível guardar a correção.");
        botao.disabled = false;
        botao.textContent = "Guardar valores corrigidos";
    }
};
window.verPreviewKM = function(){

    const rows = document.querySelectorAll("#linhasKM tr:not(:first-child)");

    const linhas = [];

    let totalKMs = 0;

    rows.forEach(tr => {

        const data = tr.querySelector(".data")?.value || "";
        const origem = tr.querySelector(".origem")?.value || "";
        const destino = tr.querySelector(".destino")?.value || "";
        const justificacao = tr.querySelector(".justificacao")?.value || "";
        const kms = Number(tr.querySelector(".kms")?.value) || 0;

        if(data || origem || destino || justificacao || kms){
            linhas.push({ data, origem, destino, justificacao, kms });
            totalKMs += kms;
        }

    });

    const valorKM = Number(document.getElementById("valorKM")?.value) || 0;
    const totalRecebido = totalKMs * valorKM;

    let html = `
        <p><b>Total KMs:</b> ${totalKMs}</p>
        <p><b>Valor/KM:</b> ${valorKM} €</p>
        <p><b>Total:</b> ${totalRecebido.toFixed(2)} €</p>

        <br>

        <table style="width:100%">
            <tr>
                <th>Data</th>
                <th>Origem</th>
                <th>Destino</th>
                <th>Justificação</th>
                <th>KMs</th>
            </tr>
    `;

    linhas.forEach(l => {
        html += `
            <tr>
                <td>${l.data}</td>
                <td>${l.origem}</td>
                <td>${l.destino}</td>
                <td>${l.justificacao}</td>
                <td>${l.kms}</td>
            </tr>
        `;
    });

    html += `</table>`;

    document.getElementById("conteudoKM").innerHTML = html;
    document.getElementById("modalKM").style.display = "block";
}
window.downloadPDF = async function(){

    const { jsPDF } = window.jspdf;
    const { PDFDocument } = PDFLib;

    const modal = document.getElementById("modalKM");

    const id = modal.dataset.id;

    if(!id){
        alert("Erro: sem ID");
        return;
    }

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/NotasDespesa/items/${id}?expand=fields`,
        {
            headers:{
                Authorization:"Bearer " + token
            }
        }
    );

    const data = await resp.json();

    const f = data.fields;

    const linhas =
        JSON.parse(f.LinhasJSON || "[]");

    /* =========================================
       PDF PRINCIPAL
    ========================================= */

    const pdf = new jsPDF("p","mm","a4");
   const logoBase64 = await carregarImagemBase64(
    "../assets/logo-monte-do-pasto.png"
);
   const paginaLargura = 210;

pdf.setFillColor(34, 139, 34);
pdf.rect(0, 0, paginaLargura, 28, "F");
pdf.addImage(
    logoBase64,
    "PNG",
    155,
    4,
    30,
    23
);
pdf.setTextColor(255,255,255);
pdf.setFontSize(22);
pdf.text("NOTA DE DESPESA", 15, 18);

pdf.setTextColor(0,0,0);

pdf.setDrawColor(220,220,220);
pdf.roundedRect(10, 35, 190, 80, 3, 3);

pdf.setFontSize(11);

    let y = 45;

    function novaLinha(texto, espacamento = 7){

        if(y > 270){
            pdf.addPage();
            y = 15;
        }

        pdf.text(String(texto), 15, y);

        y += espacamento;
    }

    pdf.setFontSize(18);

    novaLinha("Nota de Despesa", 10);

    pdf.setFontSize(11);

    pdf.setFont(undefined, "bold");
    novaLinha("Número:");
    pdf.setFont(undefined, "normal");
    novaLinha(f.NumeroNota || "-");

    pdf.setFont(undefined, "bold");
novaLinha("Submetido por:");
pdf.setFont(undefined, "normal");
novaLinha(f.CriadoPorNome || "-");

pdf.setFont(undefined, "bold");
novaLinha("Aprovado por:");
pdf.setFont(undefined, "normal");
novaLinha(f.AprovadoPorNome || "-");

pdf.setFont(undefined, "bold");
novaLinha("Estado:");

if(f.Estado === "Aprovado"){

    pdf.setFillColor(46,125,50);

}
else if(f.Estado === "Rejeitado"){

    pdf.setFillColor(198,40,40);

}
else{

    pdf.setFillColor(180,137,0);

}

pdf.roundedRect(15, y, 45, 10, 2, 2, "F");

pdf.setTextColor(255,255,255);

pdf.setFont(undefined, "bold");

pdf.text(
    (f.Estado || "-").toUpperCase(),
    20,
    y + 6.5
);

pdf.setTextColor(0,0,0);

pdf.setFont(undefined, "normal");

y += 16;

pdf.setFont(undefined, "bold");
novaLinha("Data/Hora:");
pdf.setFont(undefined, "normal");
novaLinha(new Date(f.Modified).toLocaleString("pt-PT"));

    if(f.JustificacaoRejeicao){

        novaLinha("");
        novaLinha("Justificação:");
        novaLinha(f.JustificacaoRejeicao);

    }

    novaLinha("");

    novaLinha(
        "Total: " +
        Number(f.TotalRecebido || 0).toFixed(2) +
        " €"
    );

    novaLinha("");

    for(const [index, l] of linhas.entries()){

        pdf.setFontSize(12);

        novaLinha(
            "Despesa " + (index + 1),
            8
        );

        pdf.setFontSize(10);

        novaLinha("Data: " + (l.data || "-"));
        novaLinha("Rubrica: " + (l.rubrica || "-"));
        novaLinha("Descrição: " + (l.descricao || "-"));

        novaLinha(
            "Valor: " +
            Number(l.valor || 0).toFixed(2) +
            " €"
        );

        if(l.ficheiroNome){

            novaLinha(
                "Documento: " +
                l.ficheiroNome
            );

        }

        novaLinha("");

        /* =========================================
           IMAGENS
        ========================================= */

        if(l.ficheiroDownloadUrl){

            try{

                const response =
                    await fetch(
                        l.ficheiroDownloadUrl
                    );

                const blob =
                    await response.blob();

                /* =============================
                   IMAGENS
                ============================= */

                if(blob.type.startsWith("image/")){

                    const reader =
                        new FileReader();

                    const base64 =
                        await new Promise(resolve => {

                            reader.onloadend =
                                () => resolve(reader.result);

                            reader.readAsDataURL(blob);

                        });

                    if(y > 180){

                        pdf.addPage();

                        y = 15;

                    }

                    pdf.addImage(
                        base64,
                        "JPEG",
                        15,
                        y,
                        90,
                        60
                    );

                    y += 70;

                }

            }catch(e){

                console.log(
                    "Erro imagem:",
                    e
                );

            }

        }

    }

    /* =========================================
       EXPORTAR PDF BASE
    ========================================= */

    const pdfBlob =
        pdf.output("blob");

    const mergedPdf =
        await PDFDocument.create();

    const baseBytes =
        await pdfBlob.arrayBuffer();

    const basePdf =
        await PDFDocument.load(baseBytes);

    const basePages =
        await mergedPdf.copyPages(
            basePdf,
            basePdf.getPageIndices()
        );

    basePages.forEach(page => {
        mergedPdf.addPage(page);
    });

    /* =========================================
       ANEXAR PDFs
    ========================================= */

    for(const l of linhas){

        if(!l.ficheiroDownloadUrl)
            continue;

        try{

            const response =
                await fetch(
                    l.ficheiroDownloadUrl
                );

            const blob =
                await response.blob();

            if(blob.type === "application/pdf"){

                const pdfBytes =
                    await blob.arrayBuffer();

                const anexoPdf =
                    await PDFDocument.load(
                        pdfBytes
                    );

                const paginas =
                    await mergedPdf.copyPages(
                        anexoPdf,
                        anexoPdf.getPageIndices()
                    );

                paginas.forEach(p => {
                    mergedPdf.addPage(p);
                });

            }

        }catch(e){

            console.log(
                "Erro PDF:",
                e
            );

        }

    }

    /* =========================================
       DOWNLOAD FINAL
    ========================================= */

    const finalBytes =
        await mergedPdf.save();

    const finalBlob =
        new Blob(
            [finalBytes],
            { type: "application/pdf" }
        );

    const url =
        URL.createObjectURL(finalBlob);

    const a =
        document.createElement("a");

    a.href = url;

    a.download = "Nota_Despesa_Final.pdf";

    a.click();

};
/* =============================
   LINHAS OUTRAS DESPESAS
============================= */

function addLinhaDespesa(){

    const tbody = document.getElementById("linhasDespesas");

    const tr = document.createElement("tr");

    tr.innerHTML = `
<td>
    <input type="date" class="dataDespesa">
</td>

    <td>
        <select class="rubrica">

            <option value="">Selecionar</option>

            <option value="Alojamento">Alojamento</option>

            <option value="Transporte">Transporte</option>

            <option value="Combustível">Combustível</option>

            <option value="Refeições">Refeições</option>

            <option value="Telefone">Telefone</option>

            <option value="Outros">Outros</option>

        </select>
    </td>

    <td>
        <input type="text" class="descricao">
    </td>

    <td>
        <input type="number"
               class="valorDespesa"
               step="0.01"
               oninput="calcularTotalDespesas()">
    </td>

    <td>
        <input type="file"
               class="ficheiroDespesa"
               accept=".pdf,image/*">
    </td>

    <td>
        <button onclick="removerLinhaDespesa(this)">
            X
        </button>
    </td>

    `;

    tbody.appendChild(tr);

}


/* =============================
   REMOVER LINHA DESPESA
============================= */

function removerLinhaDespesa(btn){

    btn.closest("tr").remove();

    calcularTotalDespesas();

}


/* =============================
   TOTAL OUTRAS DESPESAS
============================= */

function calcularTotalDespesas(){

    let total = 0;

    document.querySelectorAll(".valorDespesa").forEach(input => {

        total += Number(input.value) || 0;

    });

    document.getElementById("totalDespesas").innerText =
        total.toFixed(2) + " €";

}
/* =============================
   SEGUNDO APROVADOR DESPESA
============================= */

function mostrarSegundoAprovadorDespesa(){

    document.getElementById(
        "segundoAprovadorDespesaBox"
    ).style.display = "block";

}
/* =============================
   APROVADORES OUTRAS DESPESAS
============================= */

async function carregarAprovadoresOutrasDespesas(){

    const aprovadores = await obterAprovadores();

    const select1 = document.getElementById("aprovador1Despesa");
    const select2 = document.getElementById("aprovador2Despesa");

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
/* =============================
   GUARDAR OUTRAS DESPESAS
============================= */

async function guardarOutrasDespesas(){

    const utilizador = await testarGraph();
    const nomeColaborador = await obterNomeColaboradorDaNota(utilizador);

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;
    const numeroNota = await gerarNumeroNota(token, siteId);

    const rows =
        document.querySelectorAll("#linhasDespesas tr");

    const linhas = [];

    let indiceAnexo = 0;

    for(const tr of rows){

        const data =
            tr.querySelector(".dataDespesa")?.value || "";

        const rubrica =
            tr.querySelector(".rubrica")?.value || "";

        const descricao =
            tr.querySelector(".descricao")?.value || "";

        const valor =
            Number(
                tr.querySelector(".valorDespesa")?.value
            ) || 0;
      const ficheiro =
    tr.querySelector(".ficheiroDespesa")?.files[0];
        if(!data || !rubrica || !descricao || valor <= 0){
            continue;
        }

        let ficheiroInfo = null;

if(ficheiro){

    indiceAnexo++;

    ficheiroInfo =
        await uploadFicheiroDespesa(ficheiro, numeroNota, indiceAnexo);

}

linhas.push({

    data,
    rubrica,
    descricao,
    valor,

    ficheiroNome:
        ficheiroInfo?.nome || "",

    ficheiroUrl:
        ficheiroInfo?.url || "",

    ficheiroDownloadUrl:
        ficheiroInfo?.downloadUrl || ""

});

    }

    if(linhas.length === 0){

        alert("Tem de inserir pelo menos uma despesa.");

        return;
    }

    let total = 0;

    linhas.forEach(l => {
        total += l.valor;
    });

    const aprovador1 =
        document.getElementById("aprovador1Despesa")?.value || "";

    const aprovador2 =
        document.getElementById("aprovador2Despesa")?.value || "";

    if(!aprovador1){

        alert("Tem de selecionar um aprovador.");

        return;
    }

    const body = {

        fields: {

            Title:
                numeroNota + " - Despesa",

            NumeroNota: numeroNota,

            TipoDocumento: "DESPESA",

            CriadoPorNome:
                nomeColaborador,

            CriadoPorEmail:
                utilizador.mail ||
                utilizador.userPrincipalName,

            Estado: "Pendente",

            TotalRecebido: total,

            LinhasJSON: JSON.stringify(linhas),

            Aprovador1Email: aprovador1,

            Aprovador2Email: aprovador2

        }

    };

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/NotasDespesa/items`,
        {
            method:"POST",

            headers:{
                Authorization:"Bearer " + token,
                "Content-Type":"application/json"
            },

            body: JSON.stringify(body)
        }
    );

    if(!resp.ok){

        alert("Erro ao guardar despesa");

        return;
    }

    let avisoEmail = "";
    try{
        await notificarAprovadoresNota(body.fields);
    }catch(erro){
        console.error("Despesa guardada, mas o email falhou:", erro);
        avisoEmail = "\n\nA despesa foi guardada, mas o email ao aprovador não foi enviado. " + erro.message;
    }

    alert("✅ Despesa guardada com sucesso!" + avisoEmail);

    window.location.href = "dashboard.html";

}
/* =============================
   UPLOAD FICHEIRO SHAREPOINT
============================= */

async function uploadFicheiroDespesa(file, numeroNota, indice){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const partesNome = file.name.split(".");
    const extensao = partesNome.length > 1
        ? "." + partesNome.pop().toLowerCase().replace(/[^a-z0-9]/g, "")
        : "";
    const nomeBase = partesNome.join(".")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-zA-Z0-9_-]+/g, "-")
        .replace(/^-+|-+$/g, "")
        .slice(0, 80) || "anexo";

    const nomeFinal =
        `${numeroNota}-${String(indice).padStart(2, "0")}-${Date.now()}-${nomeBase}${extensao}`;

    const uploadResp = await fetch(

        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/AnexosDespesas/${nomeFinal}:/content`,

        {
            method:"PUT",

            headers:{
                Authorization:"Bearer " + token
            },

            body:file
        }
    );

    if(!uploadResp.ok){

        throw new Error("Erro upload ficheiro");

    }

    const uploadData = await uploadResp.json();

   return {

    nome: nomeFinal,

    url: uploadData.webUrl,

    downloadUrl:
        uploadData["@microsoft.graph.downloadUrl"]

};

}
async function carregarImagemBase64(url){

    const response = await fetch(url);

    const blob = await response.blob();

    return await new Promise(resolve => {

        const reader = new FileReader();

        reader.onloadend = () => {

            resolve(reader.result);

        };

        reader.readAsDataURL(blob);

    });

}
