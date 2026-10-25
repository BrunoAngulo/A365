const REPORT_BASE_URL = "https://applinde.a365.com.pe/reportes/reporte_matriz";
const INBOUND_ORIGIN = "https://tmkreport.a365.com.pe:7443";
const INBOUND_LOGIN_PAGE = `${INBOUND_ORIGIN}/pe/intranet/index.php`;
const INBOUND_LOGIN = `${INBOUND_ORIGIN}/pe/intranet/ti_login_acceso.php`;
const INBOUND_REPORT_PAGE = `${INBOUND_ORIGIN}/pe/intranet/aai_claro_asistencia_reporte_inbound.php`;
const INBOUND_GENERATE = `${INBOUND_ORIGIN}/pe/intranet/aai_claro_asistencia_vicidial_reporte_operativo_acumulada_inbound_detalle_calls.php`;
const INBOUND_DOWNLOAD = `${INBOUND_ORIGIN}/pe/intranet/aai_claro_asistencia_reporte_descarga_excel.php`;
const INBOUND_QUEUES = [
  "INLINDECLI", "INLINDEINDCEN", "INLINDEINDDES", "INLINDEINDNOR", "INLINDEINDORI", "INLINDEINDSUR",
  "INLINDEMEDCEN", "INLINDEMEDDES", "INLINDEMEDNOR", "INLINDEMEDORI", "INLINDEMEDSUR", "INLINDESAC",
  "INLINDETERA", "INLINDEURGCEN", "INLINDEURGDES", "INLINDEURGNORTE", "INLINDEURGORI", "INLINDEURGSUR",
];

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

function extractInboundCsrf(page) {
  return page.match(/data-csrf\s*=\s*["']([^"']+)["']/i)?.[1] || "";
}

function extractInboundCampaign(page) {
  const patterns = [
    /(?:name|id)\s*=\s*["']vm_camp["'][^>]*value\s*=\s*["']([^"']+)["']/i,
    /value\s*=\s*["']([^"']+)["'][^>]*(?:name|id)\s*=\s*["']vm_camp["']/i,
    /["']vm_camp["']\s*:\s*["']([^"']+)["']/i,
    /(?:var\s+)?vm_camp\s*[:=]\s*["']([^"']+)["']/i,
    /\b(dmFf[A-Za-z0-9+/=]{32,})\b/,
  ];
  for (const pattern of patterns) {
    const value = page.match(pattern)?.[1]?.trim();
    if (value && value.length >= 24) return value;
  }
  return "";
}

function inboundHeaders(extra = {}) {
  return {
    "accept-language": "es-PE,es;q=0.9",
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/156.0.0.0 Safari/537.36",
    ...extra,
  };
}

function wait(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

async function downloadInboundReport(payload) {
  const startDate = String(payload?.startDate || "").trim();
  const endDate = String(payload?.endDate || "").trim();
  const token = String(payload?.token || "");
  if (!isValidDate(startDate) || !isValidDate(endDate) || startDate > endDate) throw new Error("Selecciona fechas válidas.");
  if (!token) throw new Error("Ingresa tu token de A365.");

  const loginPageResponse = await fetch(INBOUND_LOGIN_PAGE, { cache: "no-store", credentials: "include", headers: inboundHeaders({ accept: "text/html" }) });
  const loginCsrf = extractInboundCsrf(await loginPageResponse.text());
  if (!loginPageResponse.ok || !loginCsrf) throw new Error("No pude preparar el inicio de sesión de A365.");

  const loginPayload = new URLSearchParams({ usernamePost: "psalasb", passwordPost: token, extensionPost: "", csrfPost: loginCsrf });
  const loginResponse = await fetch(INBOUND_LOGIN, {
    method: "POST", cache: "no-store", credentials: "include",
    headers: inboundHeaders({ accept: "application/json, text/plain, */*", "content-type": "application/x-www-form-urlencoded; charset=UTF-8", origin: INBOUND_ORIGIN, "x-requested-with": "XMLHttpRequest" }),
    body: loginPayload,
  });
  if (!loginResponse.ok) throw new Error("No pude iniciar sesión en A365.");

  const reportPageResponse = await fetch(INBOUND_REPORT_PAGE, { cache: "no-store", credentials: "include", headers: inboundHeaders({ accept: "text/html" }) });
  const campaign = extractInboundCampaign(await reportPageResponse.text());
  if (!reportPageResponse.ok || !campaign) throw new Error("A365 no devolvió la configuración de campaña.");

  const generationPayload = new URLSearchParams({
    vm_camp: campaign, vm_fecha_ini: startDate, vm_fecha_fin: endDate, vm_busqueda: "2", vm_colas: INBOUND_QUEUES.join(","),
    vm_duracion_prom: "undefined", vm_ns_deseado: "undefined", vm_tiempo: "undefined",
    vm_nom_archivo: "Reporte Inbound - Detalle de llamadas.xls", vm_ruta_archivo: "reportes/",
  });
  const generationResponse = await fetch(INBOUND_GENERATE, {
    method: "POST", cache: "no-store", credentials: "include",
    headers: inboundHeaders({ accept: "*/*", "content-type": "application/x-www-form-urlencoded; charset=UTF-8", origin: INBOUND_ORIGIN, referer: INBOUND_REPORT_PAGE, "x-requested-with": "XMLHttpRequest" }),
    body: generationPayload,
  });
  if (!generationResponse.ok) throw new Error(`A365 no pudo generar el reporte (${generationResponse.status}).`);

  const downloadQuery = new URLSearchParams({ vm_ruta: "reportes/", vm_nom_archivo: "Reporte Inbound - Detalle de llamadas.xls" });
  for (let attempt = 0; attempt < 5; attempt += 1) {
    const response = await fetch(`${INBOUND_DOWNLOAD}?${downloadQuery}`, { cache: "no-store", credentials: "include", headers: inboundHeaders({ accept: "*/*", referer: INBOUND_REPORT_PAGE }) });
    if (response.ok) {
      const buffer = await response.arrayBuffer();
      if (buffer.byteLength) return { dataBase64: await arrayBufferToBase64(buffer), contentType: response.headers.get("content-type") || "application/vnd.ms-excel" };
    }
    await wait(1200);
  }
  throw new Error("A365 generó el reporte, pero aún no está disponible para descargar.");
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message?.type === "A365_INBOUND_REPORT_REQUEST") {
    downloadInboundReport(message.payload)
      .then((report) => sendResponse({ ok: true, ...report }))
      .catch((error) => sendResponse({ ok: false, error: error instanceof Error ? error.message : "No pude descargar el reporte inbound." }));
    return true;
  }

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
