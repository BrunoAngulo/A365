const REPORT_BASE_URL = "https://applinde.a365.com.pe/reportes/reporte_matriz";

function normalizeSession(value = "") {
  const text = String(value).trim();
  const match = text.match(/(?:^|;\s*)PHPSESSID=([^;\s]+)/i) || text.match(/PHPSESSID=([^;\s]+)/i);
  return (match?.[1] || text).replace(/^["']|["']$/g, "").trim();
}

function isValidDate(value = "") {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value));
}

function isValidSession(value = "") {
  return /^[A-Za-z0-9]+$/.test(String(value));
}

function setSessionCookie(sessionId) {
  return new Promise((resolve, reject) => {
    chrome.cookies.set(
      {
        url: "https://applinde.a365.com.pe/",
        name: "PHPSESSID",
        value: sessionId,
        path: "/",
      },
      (cookie) => {
        const error = chrome.runtime.lastError;

        if (error) {
          reject(new Error(error.message));
          return;
        }

        resolve(cookie);
      },
    );
  });
}

async function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;

  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }

  return btoa(binary);
}

function looksLikeLoginPage(buffer, contentType) {
  if (!String(contentType).toLowerCase().includes("text/html")) {
    return false;
  }

  const preview = new TextDecoder()
    .decode(buffer.slice(0, 5000))
    .toLowerCase();
  const hasReportTable = preview.includes("<table") && (
    preview.includes("subject_email") ||
    preview.includes("date_email") ||
    preview.includes("fecha_asignacion") ||
    preview.includes("usuario asignado")
  );

  return !hasReportTable && (
    preview.includes("login") ||
    preview.includes("password") ||
    preview.includes("contraseña") ||
    preview.includes("iniciar ses") ||
    preview.includes("phpsessid")
  );
}

async function downloadMatrixReport(payload) {
  const startDate = String(payload?.startDate || "").trim();
  const endDate = String(payload?.endDate || "").trim();
  const sessionId = normalizeSession(payload?.sessionId);

  if (!isValidDate(startDate) || !isValidDate(endDate)) {
    throw new Error("Selecciona fechas validas.");
  }

  if (!isValidSession(sessionId)) {
    throw new Error("Ingresa un PHPSESSID valido.");
  }

  await setSessionCookie(sessionId);

  const response = await fetch(`${REPORT_BASE_URL}/${startDate}/${endDate}/`, {
    cache: "no-store",
    redirect: "follow",
    credentials: "include",
    headers: {
      accept: "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;q=0.8,*/*;q=0.7",
      "accept-language": "es-419,es;q=0.9",
    },
  });

  if (!response.ok) {
    throw new Error(`No pude descargar el reporte. Codigo ${response.status}.`);
  }

  const contentType = response.headers.get("content-type") || "application/octet-stream";
  const buffer = await response.arrayBuffer();

  if (looksLikeLoginPage(buffer, contentType)) {
    throw new Error("A365 devolvio login o sesion expirada. Vuelve a copiar el PHPSESSID.");
  }

  return {
    dataBase64: await arrayBufferToBase64(buffer),
    contentType,
  };
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type !== "A365_MATRIX_REPORT_REQUEST") {
    return false;
  }

  downloadMatrixReport(message.payload)
    .then((report) => sendResponse({ ok: true, ...report }))
    .catch((error) => sendResponse({
      ok: false,
      error: error instanceof Error ? error.message : "No pude descargar el reporte.",
    }));

  return true;
});
