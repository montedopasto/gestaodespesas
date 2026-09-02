const EMAIL_AUTORIZADO_RELATORIO = "jose.almanso@montedopasto.pt";

function mostrarEstadoRelatorio(mensagem, tipo = ""){
    const estado = document.getElementById("estadoRelatorio");
    if(!estado) return;
    estado.textContent = mensagem;
    estado.className = "estado-relatorio" + (tipo ? " " + tipo : "");
}

function dataISOHoje(){
    const agora = new Date();
    const deslocamento = agora.getTimezoneOffset() * 60000;
    return new Date(agora.getTime() - deslocamento).toISOString().slice(0, 10);
}

function dataISOPRimeiroDiaMes(){
    const agora = new Date();
    const mes = String(agora.getMonth() + 1).padStart(2, "0");
    return `${agora.getFullYear()}-${mes}-01`;
}

async function validarAcessoRelatorio(){
    const utilizador = await testarGraph();
    const email = String(utilizador.mail || utilizador.userPrincipalName || "").trim().toLowerCase();

    if(email !== EMAIL_AUTORIZADO_RELATORIO){
        alert("Não tem autorização para aceder a este relatório.");
        window.location.replace("dashboard.html");
        return null;
    }

    const nome = document.getElementById("utilizador");
    if(nome) nome.textContent = utilizador.displayName || email;
    return utilizador;
}

async function obterTodasNotasParaRelatorio(){
    const token = await getAccessToken();
    const site = await obterSiteApp();
    let url = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/NotasDespesa/items?expand=fields&$top=999`;
    const notas = [];

    while(url){
        const resposta = await fetch(url, {
            headers:{ Authorization:"Bearer " + token }
        });

        if(!resposta.ok){
            throw new Error("Não foi possível carregar as despesas.");
        }

        const dados = await resposta.json();
        notas.push(...(dados.value || []));
        url = dados["@odata.nextLink"] || null;
    }

    return notas;
}

function lerLinhasNota(campos){
    try{
        const linhas = JSON.parse(campos.LinhasJSON || "[]");
        return Array.isArray(linhas) ? linhas : [];
    }catch(erro){
        console.warn("Linhas inválidas na nota", campos.NumeroNota, erro);
        return [];
    }
}

function formatarDataHoraExcel(valor){
    if(!valor) return "";
    const data = new Date(valor);
    return Number.isNaN(data.getTime()) ? "" : data.toLocaleString("pt-PT");
}

function prepararDadosRelatorio(notas, dataInicio, dataFim){
    const todas = [];
    const outras = [];
    const kms = [];

    notas.forEach(item => {
        const f = item.fields || {};
        const tipoKm = String(f.TipoDocumento || "").toUpperCase() === "KMS";

        lerLinhasNota(f).forEach((linha, indice) => {
            const dataDespesa = String(linha.data || "").slice(0, 10);
            if(!dataDespesa || dataDespesa < dataInicio || dataDespesa > dataFim) return;

            const comum = {
                "Número da nota": f.NumeroNota || "",
                "Data da despesa": dataDespesa,
                "Data de submissão": formatarDataHoraExcel(item.createdDateTime || f.Created),
                "Colaborador": f.CriadoPorNome || "",
                "Submetido por (email)": f.CriadoPorEmail || "",
                "Tipo": tipoKm ? "Deslocação em KM" : "Outras despesas",
                "Estado da aprovação": f.Estado || "",
                "Estado do pagamento": f.EstadoPagamento || "Por pagar",
                "Aprovador principal": f.Aprovador1Email || "",
                "Segundo aprovador": f.Aprovador2Email || "",
                "Aprovado por": f.AprovadoPorNome || "",
                "Total da nota (€)": Number(f.TotalRecebido || 0)
            };

            if(tipoKm){
                const registoKm = {
                    ...comum,
                    "Linha": indice + 1,
                    "Matrícula": f.MatriculaVeiculo || "",
                    "Origem": linha.origem || "",
                    "Destino": linha.destino || "",
                    "Justificação": linha.justificacao || "",
                    "KMs": Number(linha.kms || 0),
                    "Valor por KM (€)": Number(f.ValorPorKM || 0),
                    "Valor da linha (€)": Number(linha.kms || 0) * Number(f.ValorPorKM || 0)
                };
                kms.push(registoKm);
                todas.push(registoKm);
            }else{
                const registoDespesa = {
                    ...comum,
                    "Linha": indice + 1,
                    "Rubrica": linha.rubrica || "",
                    "Descrição": linha.descricao || "",
                    "N.º da fatura": linha.numeroFatura || linha.qrNumeroDocumento || "",
                    "Valor da linha (€)": Number(linha.valor || 0),
                    "Nome do anexo": linha.ficheiroNome || "",
                    "Ligação ao anexo": linha.ficheiroUrl || ""
                };
                outras.push(registoDespesa);
                todas.push(registoDespesa);
            }
        });
    });

    return { todas, outras, kms };
}

function ajustarFolhaExcel(folha, dados){
    if(!dados.length) return;
    const colunas = Object.keys(dados[0]);
    folha["!cols"] = colunas.map(nome => ({
        wch: Math.min(45, Math.max(12, nome.length + 2))
    }));
    folha["!autofilter"] = { ref:folha["!ref"] };
}

function adicionarFolha(livro, nome, dados){
    const conteudo = dados.length ? dados : [{ "Resultado":"Sem despesas no período selecionado" }];
    const folha = XLSX.utils.json_to_sheet(conteudo);
    ajustarFolhaExcel(folha, conteudo);
    XLSX.utils.book_append_sheet(livro, folha, nome);
}

async function exportarRelatorioDespesas(){
    const botao = document.getElementById("btnExportarRelatorio");
    const dataInicio = document.getElementById("dataInicioRelatorio")?.value || "";
    const dataFim = document.getElementById("dataFimRelatorio")?.value || "";

    if(!dataInicio || !dataFim){
        mostrarEstadoRelatorio("Indique a data inicial e a data final.", "erro");
        return;
    }
    if(dataInicio > dataFim){
        mostrarEstadoRelatorio("A data inicial não pode ser posterior à data final.", "erro");
        return;
    }
    if(typeof XLSX === "undefined"){
        mostrarEstadoRelatorio("Não foi possível iniciar a exportação para Excel.", "erro");
        return;
    }

    botao.disabled = true;
    mostrarEstadoRelatorio("A preparar o relatório...", "a-carregar");

    try{
        if(!await validarAcessoRelatorio()) return;
        const notas = await obterTodasNotasParaRelatorio();
        const dados = prepararDadosRelatorio(notas, dataInicio, dataFim);

        if(!dados.todas.length){
            mostrarEstadoRelatorio("Não existem despesas no período selecionado.", "erro");
            return;
        }

        const livro = XLSX.utils.book_new();
        adicionarFolha(livro, "Todas", dados.todas);
        adicionarFolha(livro, "Outras despesas", dados.outras);
        adicionarFolha(livro, "KMs", dados.kms);
        XLSX.writeFile(livro, `Relatorio_Despesas_${dataInicio}_a_${dataFim}.xlsx`);

        mostrarEstadoRelatorio(
            `Relatório exportado com ${dados.todas.length} linha(s): ${dados.outras.length} despesa(s) e ${dados.kms.length} deslocação(ões).`,
            "sucesso"
        );
    }catch(erro){
        console.error("Erro ao exportar relatório:", erro);
        mostrarEstadoRelatorio(erro.message || "Não foi possível exportar o relatório.", "erro");
    }finally{
        botao.disabled = false;
    }
}

window.addEventListener("load", async () => {
    document.getElementById("dataInicioRelatorio").value = dataISOPRimeiroDiaMes();
    document.getElementById("dataFimRelatorio").value = dataISOHoje();
    try{
        await validarAcessoRelatorio();
    }catch(erro){
        console.error("Erro ao validar acesso ao relatório:", erro);
        alert("Não foi possível validar o acesso ao relatório.");
        window.location.replace("dashboard.html");
    }
});
