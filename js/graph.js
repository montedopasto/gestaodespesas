async function getAccessToken() {

    const account = msalInstance.getAllAccounts()[0];

    const request = {
        scopes: ["User.Read"],
        account: account
    };

    const response = await msalInstance.acquireTokenSilent(request);

    return response.accessToken;

}


async function testarGraph(){

    const token = await getAccessToken();

    const resposta = await fetch(
        "https://graph.microsoft.com/v1.0/me",
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

    const email = utilizador.mail || utilizador.userPrincipalName;

    const site = await obterSiteApp();

    const siteId = site.id;

    const resposta = await fetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/UtilizadoresApp/items?expand=fields`,
        {
            headers: {
                Authorization: "Bearer " + token
            }
        }
    );

    const dados = await resposta.json();

    const lista = dados.value;

    const encontrado = lista.find(u => u.fields.Email === email);

    if(encontrado){
        return encontrado.fields.Perfil;
    }

    return "Utilizador";

}
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
