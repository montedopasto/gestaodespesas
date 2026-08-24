const msalConfig = {
    auth: {
        clientId: "81a4b1c0-13eb-4c3e-bb82-283fa7d52334",
        authority: "https://login.microsoftonline.com/ee417351-ea90-41e0-9147-5ea6ab38ea49",
        redirectUri: "https://montedopasto.github.io/gestaodespesas/"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const scopesAplicacao = ["User.Read", "Mail.Send", "Mail.Send.Shared"];
const promessaRetornoLogin = msalInstance.handleRedirectPromise()
    .then(response => {
        if(response?.account){
            msalInstance.setActiveAccount(response.account);
        }
        return response;
    })
    .catch(error => {
        console.error("Erro ao processar o regresso do login:", error);
        return null;
    });

function paginaAtualProtegida(){
    return window.location.pathname.includes("/pages/");
}

function guardarDestinoDepoisDoLogin(destino){
    const url = new URL(destino || window.location.href, window.location.origin);

    if(url.origin !== window.location.origin){
        return;
    }

    sessionStorage.setItem(
        "gestaoDespesasDestinoLogin",
        url.pathname + url.search + url.hash
    );
}

function obterDestinoDepoisDoLogin(){
    const guardado = sessionStorage.getItem("gestaoDespesasDestinoLogin");
    sessionStorage.removeItem("gestaoDespesasDestinoLogin");

    if(guardado?.startsWith("/gestaodespesas/")){
        return guardado;
    }

    return paginaAtualProtegida()
        ? window.location.pathname + window.location.search + window.location.hash
        : "/gestaodespesas/pages/dashboard.html";
}

async function login(destino) {

    const loginRequest = {
        scopes: scopesAplicacao
    };

    try {

        await promessaRetornoLogin;

        if(destino || paginaAtualProtegida()){
            guardarDestinoDepoisDoLogin(destino || window.location.href);
        }

        const response = await msalInstance.loginPopup({
            ...loginRequest,
            prompt: "select_account"
        });

        msalInstance.setActiveAccount(response.account);

        console.log("Login efetuado:", response.account);

        window.location.href = obterDestinoDepoisDoLogin();

    } catch (error) {

        console.error("Erro no login:", error);

    }

}

async function redirecionarParaLoginSeNecessario(){
    await promessaRetornoLogin;

    const conta = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
    if(conta){
        msalInstance.setActiveAccount(conta);
        return conta;
    }

    guardarDestinoDepoisDoLogin(window.location.href);

    await msalInstance.loginRedirect({
        scopes: scopesAplicacao,
        redirectStartPage: window.location.href
    });

    return null;
}
