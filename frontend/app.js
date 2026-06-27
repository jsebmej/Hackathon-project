const backendUrl = "http://127.0.0.1:8000";

const clientIdInput = document.querySelector("#clientIdInput");
const initGoogleButton = document.querySelector("#initGoogleButton");
const logoutButton = document.querySelector("#logoutButton");
const googleButton = document.querySelector("#googleButton");
const authStatus = document.querySelector("#authStatus");
const checkMeButton = document.querySelector("#checkMeButton");
const chatForm = document.querySelector("#chatForm");
const messageInput = document.querySelector("#messageInput");
const sendButton = document.querySelector("#sendButton");
const messagesEl = document.querySelector("#messages");

let idToken = localStorage.getItem("google_id_token") || "";
let chatHistory = [];

clientIdInput.value = localStorage.getItem("google_client_id") || "";

function setStatus(message, type = "") {
  authStatus.textContent = message;
  authStatus.className = `status ${type}`.trim();
}

function setAuthenticated(token) {
  idToken = token;
  localStorage.setItem("google_id_token", token);
  checkMeButton.disabled = false;
  messageInput.disabled = false;
  sendButton.disabled = false;
  logoutButton.disabled = false;
  setStatus("Token recibido de Google", "ok");
}

function setLoggedOut() {
  idToken = "";
  localStorage.removeItem("google_id_token");
  checkMeButton.disabled = true;
  messageInput.disabled = true;
  sendButton.disabled = true;
  logoutButton.disabled = true;
  setStatus("Sin autenticar");
}

function addMessage(role, content) {
  const item = document.createElement("div");
  item.className = `message ${role}`;
  item.textContent = content;
  messagesEl.appendChild(item);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function requireToken() {
  if (!idToken) {
    throw new Error("Primero inicia sesion con Google.");
  }
}

async function apiFetch(path, options = {}) {
  requireToken();

  const response = await fetch(`${backendUrl}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${idToken}`,
      ...(options.headers || {}),
    },
  });

  const text = await response.text();
  const data = text ? JSON.parse(text) : null;

  if (!response.ok) {
    throw new Error(data?.detail || `Error HTTP ${response.status}`);
  }

  return data;
}

initGoogleButton.addEventListener("click", () => {
  const clientId = clientIdInput.value.trim();

  if (!clientId) {
    setStatus("Pega tu Google Client ID primero", "error");
    return;
  }

  if (!window.google?.accounts?.id) {
    setStatus("Google Identity Services aun no cargo. Intenta de nuevo.", "error");
    return;
  }

  localStorage.setItem("google_client_id", clientId);
  googleButton.innerHTML = "";

  window.google.accounts.id.initialize({
    client_id: clientId,
    callback: (response) => setAuthenticated(response.credential),
  });

  window.google.accounts.id.renderButton(googleButton, {
    theme: "outline",
    size: "large",
    width: 260,
  });

  setStatus("Login de Google listo", "ok");
});

logoutButton.addEventListener("click", () => {
  setLoggedOut();
  addMessage("system", "Sesion cerrada en esta pagina.");
});

checkMeButton.addEventListener("click", async () => {
  try {
    const data = await apiFetch("/auth/me");
    addMessage("system", `Backend valido a: ${data.user.email || data.user.sub}`);
  } catch (error) {
    setStatus(error.message, "error");
    addMessage("system", error.message);
  }
});

chatForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const message = messageInput.value.trim();
  if (!message) {
    return;
  }

  messageInput.value = "";
  addMessage("user", message);
  sendButton.disabled = true;

  try {
    const data = await apiFetch("/chat", {
      method: "POST",
      body: JSON.stringify({
        message,
        history: chatHistory,
      }),
    });

    chatHistory.push({ role: "user", content: message });
    chatHistory.push({ role: "assistant", content: data.answer });
    addMessage("assistant", data.answer);
  } catch (error) {
    addMessage("system", error.message);
  } finally {
    sendButton.disabled = false;
  }
});

if (idToken) {
  setAuthenticated(idToken);
}
