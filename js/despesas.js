/* =============================
   TABELA KM
============================= */

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

async function guardarDespesaKM(){

    const utilizador = await testarGraph();
    const token = await getAccessToken();
    const site = await obterSiteApp();

    const siteId = site.id;

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
    Title: "Nota KM - " + new Date().toLocaleDateString("pt-PT"),
    TipoDocumento: "KMS",
    CriadoPorNome: utilizador.displayName,
    CriadoPorEmail: utilizador.mail || utilizador.userPrincipalName,

    // ❌ REMOVIDO DataCriacao

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

alert("✅ Nota de despesa guardada com sucesso!");

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

    alert("Estado atualizado: " + estado);

    carregarAprovacoesDespesas();
}
/* =============================
   DASHBOARD DESPESAS
============================= */

async function carregarDashboardDespesas(){

    const utilizador = await testarGraph();
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
    const items = data.value || [];

    const email = utilizador.mail || utilizador.userPrincipalName;

    let total = items.length;
    let pendentes = 0;
    let aprovados = 0;
    let rejeitados = 0;
    let meusPendentes = 0;

    items.forEach(i => {

        const f = i.fields;

        if(f.Estado === "Pendente"){
            pendentes++;

            if(f.Aprovador1Email === email || f.Aprovador2Email === email){
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

items.forEach(item => {

    const f = item.fields;

    const linha = document.createElement("tr");

    const aprovadores = [
        f.Aprovador1Email,
        f.Aprovador2Email
    ]
    .filter(a => a)
    .map(a => a.split("@")[0])
    .join(" / ");

    linha.innerHTML = `
    <td>${new Date(f.Created).toLocaleDateString("pt-PT")}</td>
    <td>${f.CriadoPorNome}</td>
    <td>${Number(f.TotalRecebido).toFixed(2)} €</td>
    <td>${aprovadores || "-"}</td>
    <td>
        <span class="estado ${f.Estado}">
            ${f.Estado}
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
const zonaEstado = document.getElementById("zonaEstadoPedido");
const carimbo = document.getElementById("carimboEstado");
const caixaJustificacao = document.getElementById("justificacaoRejeicao");
const textoJustificacao = document.getElementById("textoJustificacaoRejeicao");

if(zonaEstado && carimbo && caixaJustificacao && textoJustificacao){

    zonaEstado.style.display = "block";
    caixaJustificacao.style.display = "none";
    textoJustificacao.innerText = "";

    if(f.Estado === "Aprovado"){

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
const linhas = JSON.parse(f.LinhasJSON || "[]");
if(f.TipoDocumento === "DESPESA"){

    let html = `

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

    linhas.forEach(l => {

        html += `

        <tr>

            <td>${l.data}</td>

            <td>${l.rubrica}</td>

            <td>${l.descricao}</td>

            <td>${Number(l.valor).toFixed(2)} €</td>

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

    html += `</table>`;

    document.getElementById("conteudoKM").innerHTML = html;

    document.getElementById("modalKM").style.display = "block";

    return;
}
    let html = `
        <p><b>Total KMs:</b> ${f.TotalKMs}</p>
        <p><b>Valor/KM:</b> ${f.ValorPorKM} €</p>
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
   const paginaLargura = 210;

pdf.setFillColor(34, 139, 34);
pdf.rect(0, 0, paginaLargura, 28, "F");

pdf.setTextColor(255,255,255);
pdf.setFontSize(22);
pdf.text("NOTA DE DESPESA", 15, 18);

pdf.setTextColor(0,0,0);

pdf.setDrawColor(220,220,220);
pdf.roundedRect(10, 35, 190, 45, 3, 3);

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

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const rows =
        document.querySelectorAll("#linhasDespesas tr");

    const linhas = [];

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

    ficheiroInfo =
        await uploadFicheiroDespesa(ficheiro);

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
                "Despesa - " +
                new Date().toLocaleDateString("pt-PT"),

            TipoDocumento: "DESPESA",

            CriadoPorNome:
                utilizador.displayName,

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

    alert("✅ Despesa guardada com sucesso!");

    window.location.href = "dashboard.html";

}
/* =============================
   UPLOAD FICHEIRO SHAREPOINT
============================= */

async function uploadFicheiroDespesa(file){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const nomeFinal =
        Date.now() + "_" + file.name;

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
