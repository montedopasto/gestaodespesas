const msalConfig = {
    auth: {
        clientId: "81693cb9-9ffb-41d0-b0ff-41dbf29990eb",
        authority: "https://login.microsoftonline.com/ee417351-ea90-41e0-9147-5ea6ab38ea49",
        redirectUri: "https://jaca000.github.io/aprovacao-faturas/"
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

        window.location.href = "/aprovacao-faturas/pages/dashboard.html";

    } catch (error) {

        console.error("Erro no login:", error);

    }

}
