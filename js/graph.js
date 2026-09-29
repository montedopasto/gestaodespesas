async function getAccessToken() {

    await promessaRetornoLogin;

    let account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];

    if(!account){
        account = await redirecionarParaLoginSeNecessario();
    }

    if(!account){
        throw new Error("A iniciar sessão Microsoft...");
    }

    msalInstance.setActiveAccount(account);

    const request = {
        scopes: ["User.Read"],
        account: account
    };

    let response;

    try{
        response = await msalInstance.acquireTokenSilent(request);
    }catch(erro){
        console.warn("A sessão precisa de ser renovada:", erro);
        guardarDestinoDepoisDoLogin(window.location.href);
        await msalInstance.acquireTokenRedirect({
            ...request,
            redirectStartPage: window.location.href
        });
        throw new Error("A renovar a sessão Microsoft...");
    }

    return response.accessToken;

}

async function getEmailAccessToken(){

    const account = msalInstance.getAllAccounts()[0];
    if(!account){
        throw new Error("Sessão Microsoft não encontrada.");
    }

    try{
        const response = await msalInstance.acquireTokenSilent({
            scopes:["Mail.Send", "Mail.Send.Shared"],
            account:account
        });

        return response.accessToken;
    }catch(erro){
        console.error("Permissão Mail.Send indisponível:", erro);
        throw new Error(
            "A conta não tem autorização para enviar emails. Termine a sessão e volte a entrar depois de configurar a permissão Mail.Send."
        );
    }
}

async function enviarEmailGraph(destinatarios, assunto, conteudoHTML){

    const remetente = "gestaodespesas@montedopasto.pt";

    const emails = [...new Set(
        (destinatarios || [])
            .map(email => String(email || "").trim().toLowerCase())
            .filter(email => email.includes("@"))
    )];

    if(!emails.length){
        throw new Error("O email do destinatário não está definido.");
    }

    const token = await getEmailAccessToken();
    const resp = await fetch(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        {
            method:"POST",
            headers:{
                Authorization:"Bearer " + token,
                "Content-Type":"application/json"
            },
            body:JSON.stringify({
                message:{
                    subject:assunto,
                    from:{
                        emailAddress:{
                            name:"App Gestão de Despesas",
                            address:remetente
                        }
                    },
                    body:{
                        contentType:"HTML",
                        content:conteudoHTML
                    },
                    toRecipients:emails.map(email => ({
                        emailAddress:{ address:email }
                    }))
                },
                saveToSentItems:true
            })
        }
    );

    if(!resp.ok){
        const detalhe = await resp.text();
        console.error("Erro Graph sendMail:", detalhe);
        throw new Error("O Microsoft 365 recusou o envio do email.");
    }
}


async function testarGraph(){

    const token = await getAccessToken();

    const resposta = await fetch(
        "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,userPrincipalName,otherMails",
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterSiteApp(){

    const token = await getAccessToken();

    const resposta = await fetch(
        "https://graph.microsoft.com/v1.0/sites/montedopastopt.sharepoint.com:/sites/AppRegistoFaturas",
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPedidos(){

    const token = await getAccessToken();

    const siteId = "montedopastopt.sharepoint.com,309b2348-8df0-4dbe-9d3b2348-8df0-4dbe-945126c5bec7,3a90922f-7a65-44d9-ae1e-ef11c749a820";

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/PedidosAprovacao/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterListas(){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPedidosFaturas(){

    const token = await getAccessToken();

    const site = await obterSiteApp();

    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    return dados;

}
async function obterPerfilUtilizador(){

    const token = await getAccessToken();

    const utilizador = await testarGraph();

    const normalizarEmail = valor => String(valor || "").trim().toLowerCase();
    const emailsUtilizador = new Set([
        normalizarEmail(utilizador.mail),
        normalizarEmail(utilizador.userPrincipalName),
        ...(Array.isArray(utilizador.otherMails)
            ? utilizador.otherMails.map(normalizarEmail)
            : [])
    ].filter(Boolean));

    const site = await obterSiteApp();

    const siteId = site.id;

    let url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/UtilizadoresApp/items?expand=fields&$top=999`;
    const lista = [];

    while(url){
        const resposta = await fetch(url, {
            headers: { Authorization:"Bearer " + token }
        });

        if(!resposta.ok){
            throw new Error("Não foi possível consultar o perfil do utilizador.");
        }

        const dados = await resposta.json();
        lista.push(...(dados.value || []));
        url = dados["@odata.nextLink"] || null;
    }

    function obterCampo(fields, nome){
        const chave = Object.keys(fields || {}).find(
            item => item.toLowerCase() === nome.toLowerCase()
        );
        return chave ? fields[chave] : undefined;
    }

    function extrairEmail(valor){
        if(typeof valor === "string") return normalizarEmail(valor);
        if(valor && typeof valor === "object"){
            return normalizarEmail(
                valor.Email || valor.email || valor.LookupValue || valor.lookupValue
            );
        }
        return "";
    }

    let encontrado = lista.find(u => {
        const fields = u.fields || {};
        const emailRegisto = extrairEmail(obterCampo(fields, "Email"));
        const tituloRegisto = extrairEmail(obterCampo(fields, "Title"));
        return emailsUtilizador.has(emailRegisto) || emailsUtilizador.has(tituloRegisto);
    });

    if(!encontrado){
        const locaisUtilizador = new Set(
            [...emailsUtilizador].map(email => email.split("@")[0]).filter(Boolean)
        );
        const candidatos = lista.filter(u => {
            const fields = u.fields || {};
            const emailRegisto = extrairEmail(obterCampo(fields, "Email"));
            const tituloRegisto = extrairEmail(obterCampo(fields, "Title"));
            return [emailRegisto, tituloRegisto].some(email =>
                email && locaisUtilizador.has(email.split("@")[0])
            );
        });
        if(candidatos.length === 1) encontrado = candidatos[0];
    }

    if(encontrado){
        const valorPerfil = obterCampo(encontrado.fields || {}, "Perfil");
        const perfil = String(
            typeof valorPerfil === "object"
                ? valorPerfil?.LookupValue || valorPerfil?.Value || ""
                : valorPerfil || ""
        ).trim().toLowerCase();
        const perfis = {
            admin:"Admin",
            gestorfaturas:"GestorFaturas",
            utilizador:"Utilizador",
            registador:"Registador"
        };
        return perfis[perfil] || "Utilizador";
    }

    return "Utilizador";

}

/* Mostra o módulo financeiro apenas aos perfis autorizados. */
async function configurarMenuPagamentos(){

    const menu = document.getElementById("menuPagamentos");
    if(!menu) return;

    try{
        const perfil = await obterPerfilUtilizador();
        menu.style.display =
            perfil === "Admin" || perfil === "GestorFaturas"
                ? "flex"
                : "none";
    }catch(erro){
        console.error("Não foi possível validar o acesso a Pagamentos:", erro);
        menu.style.display = "none";
    }
}

const EMAIL_RELATORIO_DESPESAS = "jose.almanso@montedopasto.pt";

async function configurarMenuRelatorioDespesas(){
    const menu = document.getElementById("menuRelatorioDespesas");
    if(!menu) return;

    try{
        const utilizador = await testarGraph();
        const email = String(utilizador.mail || utilizador.userPrincipalName || "").toLowerCase();
        menu.style.display = email === EMAIL_RELATORIO_DESPESAS ? "flex" : "none";
    }catch(erro){
        console.error("Não foi possível validar o acesso ao relatório:", erro);
        menu.style.display = "none";
    }
}

async function configurarRestricoesRegistador(){
    try{
        const perfil = await obterPerfilUtilizador();
        if(perfil !== "Registador") return;

        document.querySelectorAll(
            "#menuDashboard, #menuAprovacoesDespesas, #menuAprovacoes, .btn-aprovar, .btn-rejeitar"
        ).forEach(elemento => {
            elemento.style.display = "none";
        });

        const pagina = window.location.pathname.split("/").pop().toLowerCase();
        if(pagina && pagina !== "nova-despesa.html"){
            window.location.replace("nova-despesa.html");
        }

    }catch(erro){
        console.error("Não foi possível aplicar as restrições do perfil Registador:", erro);
    }
}

window.addEventListener("load", configurarMenuPagamentos);
window.addEventListener("load", configurarMenuRelatorioDespesas);
window.addEventListener("load", configurarRestricoesRegistador);
async function uploadPdfSharePoint(ficheiro){

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    // obter drives do site
    const drivesResp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
        {
            headers: { Authorization: "Bearer " + token }
        }
    );

    const drives = await drivesResp.json();

    // encontrar biblioteca DocumentosAprovacao
    const drive = drives.value.find(d => d.name === "DocumentosAprovacao");

    if(!drive){
        throw new Error("Biblioteca DocumentosAprovacao não encontrada");
    }

    const driveId = drive.id;

    const uploadUrl =
        `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${ficheiro.name}:/content`;

    const uploadResp = await fetch(uploadUrl,{
        method: "PUT",
        headers: {
            Authorization: "Bearer " + token,
            "Content-Type": ficheiro.type
        },
        body: ficheiro
    });

    const resultado = await uploadResp.json();

    console.log("Upload PDF:", resultado);

    return resultado;
}
async function verificarFaturaDuplicada(numeroNormalizado){

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items?$expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const dados = await resp.json();

    const lista = dados.value || [];

    const existe = lista.some(item =>
        item.fields.NumeroFaturaNormalizado === numeroNormalizado
    );

    return existe;

}
async function gerarNumeroInterno(){

    const token = await getAccessToken();

    const site = await obterSiteApp();
    const siteId = site.id;

    const listaId = "5baaca12-aaf0-4e67-b094-20ed3487f7e9";

    const resp = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listaId}/items?$expand=fields`,
        {
            headers:{ Authorization:"Bearer " + token }
        }
    );

    const dados = await resp.json();

    const lista = dados.value || [];

    const ano = new Date().getFullYear();

    const numeros = lista
        .map(i => i.fields.NumeroInterno)
        .filter(n => n && n.includes(ano));

    let ultimo = 0;

    numeros.forEach(n => {

        const partes = n.split("-");
        const seq = parseInt(partes[2]);

        if(seq > ultimo){
            ultimo = seq;
        }

    });

    const novo = ultimo + 1;

    const numeroFormatado = String(novo).padStart(3,"0");

    return `FRL-${ano}-${numeroFormatado}`;

}
