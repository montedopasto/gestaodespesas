let pagamentosAprovados = [];
let filtroPagamentoAtual = "por-pagar";
let pagamentoSelecionadoId = null;

function escaparHTML(valor){
    return String(valor ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function despesaEstaPaga(campos){
    return campos.Pago === true ||
        campos.Pago === 1 ||
        String(campos.Pago).toLowerCase() === "true" ||
        campos.EstadoPagamento === "Pago";
}

function formatarEuro(valor){
    return Number(valor || 0).toLocaleString("pt-PT", {
        style: "currency",
        currency: "EUR"
    });
}

function formatarData(valor, incluirHora = false){
    if(!valor) return "-";
    const data = new Date(valor);
    if(Number.isNaN(data.getTime())) return "-";
    return incluirHora
        ? data.toLocaleString("pt-PT")
        : data.toLocaleDateString("pt-PT");
}

function obterLinhasDespesa(valor){
    try{
        const linhas = JSON.parse(valor || "[]");
        return Array.isArray(linhas) ? linhas : [];
    }catch(erro){
        console.warn("LinhasJSON inválido na despesa:", erro);
        return [];
    }
}

async function validarAcessoPagamentos(){
    const perfil = await obterPerfilUtilizador();
    const autorizado = perfil === "Admin" || perfil === "GestorFaturas";

    if(!autorizado){
        alert("Não tem permissão para aceder aos pagamentos.");
        window.location.replace("dashboard.html");
    }

    return autorizado;
}

async function carregarPagamentos(){
    if(!await validarAcessoPagamentos()) return;

    const tabela = document.getElementById("tabelaPagamentos");
    tabela.innerHTML = '<tr><td colspan="6" class="pagamentos-vazio">A carregar...</td></tr>';

    try{
        const token = await getAccessToken();
        const site = await obterSiteApp();
        const resp = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/NotasDespesa/items?expand=fields`,
            { headers:{ Authorization:"Bearer " + token } }
        );

        if(!resp.ok){
            throw new Error(await resp.text());
        }

        const data = await resp.json();
        pagamentosAprovados = (data.value || [])
            .filter(item => item.fields.Estado === "Aprovado")
            .sort((a,b) => new Date(b.fields.Created) - new Date(a.fields.Created));

        atualizarResumoPagamentos();
        renderizarPagamentos();
    }catch(erro){
        console.error("Erro ao carregar pagamentos:", erro);
        tabela.innerHTML = '<tr><td colspan="6" class="pagamentos-vazio">Não foi possível carregar os pagamentos.</td></tr>';
    }
}

function atualizarResumoPagamentos(){
    const agora = new Date();
    const porPagar = pagamentosAprovados.filter(i => !despesaEstaPaga(i.fields));
    const pagasMes = pagamentosAprovados.filter(i => {
        if(!despesaEstaPaga(i.fields) || !i.fields.DataPagamento) return false;
        const data = new Date(i.fields.DataPagamento);
        return data.getMonth() === agora.getMonth() &&
            data.getFullYear() === agora.getFullYear();
    });

    document.getElementById("totalPorPagar").textContent = porPagar.length;
    document.getElementById("valorPorPagar").textContent =
        formatarEuro(porPagar.reduce((soma, i) => soma + Number(i.fields.TotalRecebido || 0), 0));
    document.getElementById("pagasEsteMes").textContent = pagasMes.length;
    document.getElementById("valorPagoEsteMes").textContent =
        formatarEuro(pagasMes.reduce((soma, i) => soma + Number(i.fields.TotalRecebido || 0), 0));
}

function alterarFiltroPagamentos(filtro, botao){
    filtroPagamentoAtual = filtro;
    document.querySelectorAll(".filtro-pagamento").forEach(b => b.classList.remove("ativo"));
    botao.classList.add("ativo");
    renderizarPagamentos();
}

function renderizarPagamentos(){
    const pesquisa = document.getElementById("pesquisaPagamentos").value.trim().toLowerCase();
    const filtrados = pagamentosAprovados.filter(item => {
        const f = item.fields;
        const paga = despesaEstaPaga(f);
        const correspondeEstado = filtroPagamentoAtual === "todas" ||
            (filtroPagamentoAtual === "pagas" ? paga : !paga);
        const texto = `${f.CriadoPorNome || ""} ${f.TipoDocumento || ""} ${f.Title || ""}`.toLowerCase();
        return correspondeEstado && texto.includes(pesquisa);
    });

    const tabela = document.getElementById("tabelaPagamentos");
    if(!filtrados.length){
        tabela.innerHTML = '<tr><td colspan="6" class="pagamentos-vazio">Não existem despesas neste filtro.</td></tr>';
        return;
    }

    tabela.innerHTML = filtrados.map(item => {
        const f = item.fields;
        const paga = despesaEstaPaga(f);
        return `
            <tr>
                <td>${formatarData(f.Created)}</td>
                <td>${escaparHTML(f.CriadoPorNome || "-")}</td>
                <td>${escaparHTML(f.TipoDocumento === "KMS" ? "KMs" : "Despesa")}</td>
                <td>${formatarEuro(f.TotalRecebido)}</td>
                <td><span class="estado-pagamento ${paga ? "pago" : "por-pagar"}">${paga ? "Pago" : "Por pagar"}</span></td>
                <td><button class="btn-icon" title="Ver detalhe" onclick="abrirDetalhePagamento('${item.id}')"><i data-lucide="file-text"></i></button></td>
            </tr>`;
    }).join("");

    if(window.lucide) lucide.createIcons();
}

async function abrirDetalhePagamento(id){
    pagamentoSelecionadoId = id;

    try{
        const token = await getAccessToken();
        const site = await obterSiteApp();
        const resp = await fetch(
            `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/NotasDespesa/items/${id}?expand=fields`,
            { headers:{ Authorization:"Bearer " + token } }
        );

        if(!resp.ok) throw new Error(await resp.text());

        const item = await resp.json();
        const f = item.fields;
        const linhas = obterLinhasDespesa(f.LinhasJSON);
        const isKM = f.TipoDocumento === "KMS";

        let detalhe = `
            <div class="detalhe-resumo">
                <p><b>Colaborador:</b> ${escaparHTML(f.CriadoPorNome || "-")}</p>
                <p><b>Data:</b> ${formatarData(f.Created)}</p>
                <p><b>Tipo:</b> ${isKM ? "KMs" : "Despesa"}</p>
                ${isKM ? `<p><b>Matrícula:</b> ${escaparHTML(f.MatriculaVeiculo || "-")}</p>` : ""}
                <p><b>Total:</b> ${formatarEuro(f.TotalRecebido)}</p>
            </div>`;

        if(isKM){
            detalhe += `<table><thead><tr><th>Data</th><th>Origem</th><th>Destino</th><th>Justificação</th><th>KMs</th></tr></thead><tbody>`;
            detalhe += linhas.map(l => `<tr><td>${escaparHTML(l.data)}</td><td>${escaparHTML(l.origem)}</td><td>${escaparHTML(l.destino)}</td><td>${escaparHTML(l.justificacao)}</td><td>${escaparHTML(l.kms)}</td></tr>`).join("");
        }else{
            detalhe += `<table><thead><tr><th>Data</th><th>Rubrica</th><th>Descrição</th><th>Valor</th><th>Fatura</th></tr></thead><tbody>`;
            detalhe += linhas.map(l => `<tr><td>${escaparHTML(l.data)}</td><td>${escaparHTML(l.rubrica)}</td><td>${escaparHTML(l.descricao)}</td><td>${formatarEuro(l.valor)}</td><td>${l.ficheiroUrl ? `<a href="${escaparHTML(l.ficheiroUrl)}" target="_blank" rel="noopener">Abrir</a>` : "-"}</td></tr>`).join("");
        }
        detalhe += "</tbody></table>";

        document.getElementById("detalhePagamento").innerHTML = detalhe;
        renderizarBlocoPagamento(f);
        document.getElementById("modalPagamento").style.display = "block";
    }catch(erro){
        console.error("Erro ao abrir pagamento:", erro);
        alert("Não foi possível abrir o detalhe da despesa.");
    }
}

function renderizarBlocoPagamento(f){
    const bloco = document.getElementById("blocoPagamento");
    if(despesaEstaPaga(f)){
        bloco.innerHTML = `
            <h3>Pagamento</h3>
            <div class="pagamento-confirmado">
                <strong>✓ PAGO</strong>
                <p><b>Pago por:</b> ${escaparHTML(f.PagoPorNome || "-")}</p>
                <p><b>Email:</b> ${escaparHTML(f.PagoPorEmail || "-")}</p>
                <p><b>Data:</b> ${formatarData(f.DataPagamento, true)}</p>
                <p><b>Notas:</b> ${escaparHTML(f.NotasPagamento || "-")}</p>
            </div>`;
        return;
    }

    bloco.innerHTML = `
        <h3>Pagamento</h3>
        <label for="notasPagamento">Notas <span>(opcional)</span></label>
        <textarea id="notasPagamento" rows="4" maxlength="2000" placeholder="Adicionar notas sobre o pagamento..."></textarea>
        <button id="btnConfirmarPagamento" class="btn-confirmar-pagamento" onclick="confirmarPagamento()">Confirmar pagamento</button>`;
}

async function confirmarPagamento(){
    if(!pagamentoSelecionadoId) return;
    if(!confirm("Tem a certeza que pretende confirmar o pagamento desta nota de despesa?")) return;

    const botao = document.getElementById("btnConfirmarPagamento");
    const notas = document.getElementById("notasPagamento").value.trim();
    botao.disabled = true;
    botao.textContent = "A confirmar...";

    try{
        if(!await validarAcessoPagamentos()) return;

        const utilizador = await testarGraph();
        const token = await getAccessToken();
        const site = await obterSiteApp();
        const url = `https://graph.microsoft.com/v1.0/sites/${site.id}/lists/NotasDespesa/items/${pagamentoSelecionadoId}`;

        const atualResp = await fetch(`${url}?expand=fields`, {
            headers:{ Authorization:"Bearer " + token }
        });
        if(!atualResp.ok) throw new Error(await atualResp.text());
        const atual = await atualResp.json();

        if(atual.fields.Estado !== "Aprovado"){
            throw new Error("A despesa deixou de estar aprovada e não pode ser paga.");
        }
        if(despesaEstaPaga(atual.fields)){
            throw new Error("Este pagamento já foi confirmado por outro utilizador.");
        }

        const resp = await fetch(`${url}/fields`, {
            method:"PATCH",
            headers:{
                Authorization:"Bearer " + token,
                "Content-Type":"application/json"
            },
            body:JSON.stringify({
                EstadoPagamento:"Pago",
                Pago:true,
                PagoPorNome:utilizador.displayName,
                PagoPorEmail:utilizador.mail || utilizador.userPrincipalName,
                DataPagamento:new Date().toISOString(),
                NotasPagamento:notas
            })
        });

        if(!resp.ok) throw new Error(await resp.text());

        fecharModalPagamento();
        await carregarPagamentos();
        alert("Pagamento confirmado com sucesso.");
    }catch(erro){
        console.error("Erro ao confirmar pagamento:", erro);
        alert(erro.message || "Não foi possível confirmar o pagamento.");
        botao.disabled = false;
        botao.textContent = "Confirmar pagamento";
    }
}

function fecharModalPagamento(){
    document.getElementById("modalPagamento").style.display = "none";
    pagamentoSelecionadoId = null;
}

window.addEventListener("load", carregarPagamentos);
