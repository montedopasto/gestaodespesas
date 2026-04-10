const msalConfig = {
    auth: {
        clientId: "81a4b1c0-13eb-4c3e-bb82-283fa7d52334",
        authority: "https://login.microsoftonline.com/ee417351-ea90-41e0-9147-5ea6ab38ea49",
        redirectUri: "https://montedopasto.github.io/gestaodespesas/"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

async function login() {

    const loginRequest = {
        scopes: ["User.Read"]
    };

    try {

        const response = await msalInstance.loginPopup({
    scopes: ["User.Read"],
    prompt: "select_account" // 🔥 força seleção limpa
});

        console.log("Login efetuado:", response.account);

        window.location.href = "/gestaodespesas/pages/dashboard.html";

    } catch (error) {

        console.error("Erro no login:", error);

    }

}
