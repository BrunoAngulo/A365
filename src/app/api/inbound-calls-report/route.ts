const portalOrigin = "https://tmkreport.a365.com.pe:7443";
const loginPageUrl = `${portalOrigin}/pe/intranet/index.php`;
const loginUrl = `${portalOrigin}/pe/intranet/ti_login_acceso.php`;
const inboundUsername = "psalasb";
const reportPageUrl = `${portalOrigin}/pe/intranet/aai_claro_asistencia_reporte_inbound.php`;
const reportUrl = `${portalOrigin}/pe/intranet/aai_claro_asistencia_vicidial_reporte_operativo_acumulada_inbound_detalle_calls.php`;
const downloadUrl = `${portalOrigin}/pe/intranet/aai_claro_asistencia_reporte_descarga_excel.php`;
const inboundQueues = [
  "INLINDECLI", "INLINDEINDCEN", "INLINDEINDDES", "INLINDEINDNOR", "INLINDEINDORI", "INLINDEINDSUR",
  "INLINDEMEDCEN", "INLINDEMEDDES", "INLINDEMEDNOR", "INLINDEMEDORI", "INLINDEMEDSUR", "INLINDESAC",
  "INLINDETERA", "INLINDEURGCEN", "INLINDEURGDES", "INLINDEURGNORTE", "INLINDEURGORI", "INLINDEURGSUR",
];

export const runtime = "nodejs";
export const maxDuration = 60;

type ReportRequest = {
  startDate?: string;
  endDate?: string;
  sessionId?: string;
  password?: string;
};

function normalizeSession(value: string) {
  const trimmed = value.trim();
  const cookieMatch = trimmed.match(/(?:^|;\s*)PHPSESSID=([^;\s]+)/i) ?? trimmed.match(/PHPSESSID=([^;\s]+)/i);
  return (cookieMatch?.[1] ?? trimmed).replace(/^["']|["']$/g, "").trim();
}

function isValidDate(value: string) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function isValidSession(value: string) {
  // A365 entrega sesiones con guiones y caracteres codificados para cookie, por
  // ejemplo "%2C". Se conservan codificados al reenviarlos al portal.
  return /^(?:[A-Za-z0-9_-]|%[0-9A-Fa-f]{2})+$/.test(value);
}

function sessionFromSetCookie(value: string | null) {
  const session = value?.match(/PHPSESSID=([^;\s,]+)/i)?.[1] ?? "";
  return normalizeSession(session);
}

function extractLoginCsrf(page: string) {
  return page.match(/data-csrf\s*=\s*["']([^"']+)["']/i)?.[1]
    ?? page.match(/name\s*=\s*["']csrfPost["'][^>]*value\s*=\s*["']([^"']+)["']/i)?.[1]
    ?? page.match(/value\s*=\s*["']([^"']+)["'][^>]*name\s*=\s*["']csrfPost["']/i)?.[1]
    ?? "";
}

function extractCampaignToken(page: string) {
  const patterns = [
    /(?:name|id)\s*=\s*["']vm_camp["'][^>]*value\s*=\s*["']([^"']+)["']/i,
    /value\s*=\s*["']([^"']+)["'][^>]*(?:name|id)\s*=\s*["']vm_camp["']/i,
    /(?:name|id)\s*=\s*["'](?:sel_)?camp(?:ana)?["'][^>]*value\s*=\s*["']([^"']+)["']/i,
    /value\s*=\s*["']([^"']+)["'][^>]*(?:name|id)\s*=\s*["'](?:sel_)?camp(?:ana)?["']/i,
    /["']vm_camp["']\s*:\s*["']([^"']+)["']/i,
    /(?:var\s+)?vm_camp\s*[:=]\s*["']([^"']+)["']/i,
    /\b(dmFf[A-Za-z0-9+/=]{32,})\b/,
  ];

  for (const pattern of patterns) {
    const value = page.match(pattern)?.[1]?.trim();
    if (value && value.length >= 24) return value;
  }

  return null;
}

function internalScriptUrls(page: string) {
  const urls = new Set<string>();
  for (const match of page.matchAll(/<script[^>]+src\s*=\s*["']([^"']+)["']/gi)) {
    try {
      const url = new URL(match[1], reportPageUrl);
      if (url.origin === portalOrigin) urls.add(url.href);
    } catch {
      // Ignora referencias de script malformadas del portal.
    }
  }
  return Array.from(urls).slice(0, 12);
}

function looksLikeLoginPage(content: ArrayBuffer, contentType: string) {
  if (!contentType.toLowerCase().includes("text/html")) return false;
  const preview = new TextDecoder().decode(content.slice(0, 5000)).toLowerCase();
  return (preview.includes("login") || preview.includes("password") || preview.includes("iniciar ses") || preview.includes("phpsessid")) && !preview.includes("<table");
}

function looksLikeVicidialDatabaseError(content: ArrayBuffer) {
  const preview = new TextDecoder().decode(content.slice(0, 20000)).toLowerCase();
  return preview.includes("conexion a la bd vicidial") || preview.includes("conexión a la bd vicidial");
}

function wait(milliseconds: number) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

export async function POST(request: Request) {
  const body = await request.json().catch(() => null) as ReportRequest | null;
  const startDate = body?.startDate?.trim() ?? "";
  const endDate = body?.endDate?.trim() ?? "";
  let sessionId = normalizeSession(body?.sessionId ?? "");
  const username = inboundUsername;
  const password = body?.password ?? "";

  if (!isValidDate(startDate) || !isValidDate(endDate) || startDate > endDate) {
    return Response.json({ error: "Selecciona un rango de fechas válido." }, { status: 400 });
  }
  if (!sessionId && !password) {
    return Response.json({ error: "Ingresa tu token de A365." }, { status: 400 });
  }

  try {
    let loginCampaignToken: string | null = null;
    if (!sessionId) {
      const loginPageResponse = await fetch(loginPageUrl, {
        cache: "no-store",
        headers: {
          accept: "text/html,application/xhtml+xml",
          "accept-language": "es-PE,es;q=0.9",
          "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
        },
      });
      const loginPage = await loginPageResponse.text();
      const loginCsrf = extractLoginCsrf(loginPage);
      const loginSessionId = sessionFromSetCookie(loginPageResponse.headers.get("set-cookie"));
      if (!loginPageResponse.ok || !loginCsrf) {
        return Response.json({ error: "No pude preparar el inicio de sesión de A365." }, { status: 502 });
      }
      const loginPayload = new URLSearchParams({
        usernamePost: username,
        passwordPost: password,
        extensionPost: "",
        csrfPost: loginCsrf,
      });
      const loginResponse = await fetch(loginUrl, {
        method: "POST",
        cache: "no-store",
        headers: {
          accept: "application/json, text/plain, */*",
          "accept-language": "es-PE,es;q=0.9",
          "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
          origin: portalOrigin,
          ...(loginSessionId ? { cookie: `PHPSESSID=${loginSessionId}` } : {}),
          "x-requested-with": "XMLHttpRequest",
          "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
        },
        body: loginPayload,
      });
      const loginContent = await loginResponse.text();
      sessionId = sessionFromSetCookie(loginResponse.headers.get("set-cookie")) || loginSessionId;
      loginCampaignToken = extractCampaignToken(loginContent);
      if (!loginResponse.ok || !isValidSession(sessionId)) {
        return Response.json({ error: "No pude iniciar sesión en A365. Verifica tu usuario y contraseña." }, { status: 401 });
      }
    }
    if (!isValidSession(sessionId)) {
      return Response.json({ error: "Ingresa un PHPSESSID válido." }, { status: 400 });
    }

    const commonHeaders = {
      "accept-language": "es-PE,es;q=0.9",
      cookie: `PHPSESSID=${sessionId}`,
      "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
    };

    // vm_camp es generado por A365 para la sesión autenticada. Nunca se recibe
    // del navegador ni se guarda: se lee justo antes de solicitar el reporte.
    const pageResponse = await fetch(reportPageUrl, {
      cache: "no-store",
      headers: { ...commonHeaders, accept: "text/html,application/xhtml+xml" },
    });
    if (!pageResponse.ok) {
      return Response.json({ error: `No pude abrir el portal de A365 (${pageResponse.status}).` }, { status: 502 });
    }

    const reportPage = await pageResponse.text();
    let campaignToken = loginCampaignToken ?? extractCampaignToken(reportPage);
    if (!campaignToken) {
      for (const scriptUrl of internalScriptUrls(reportPage)) {
        const scriptResponse = await fetch(scriptUrl, {
          cache: "no-store",
          headers: { ...commonHeaders, accept: "text/javascript, application/javascript, */*", referer: reportPageUrl },
        });
        if (!scriptResponse.ok) continue;
        campaignToken = extractCampaignToken(await scriptResponse.text());
        if (campaignToken) break;
      }
    }
    if (!campaignToken) {
      return Response.json({ error: "La sesión no devolvió la configuración de campaña. Verifica que el PHPSESSID siga activo y que tengas acceso al reporte inbound." }, { status: 401 });
    }

    const payload = new URLSearchParams({
      vm_camp: campaignToken,
      vm_fecha_ini: startDate,
      vm_fecha_fin: endDate,
      vm_busqueda: "2",
      vm_colas: inboundQueues.join(","),
      vm_duracion_prom: "undefined",
      vm_ns_deseado: "undefined",
      vm_tiempo: "undefined",
      vm_nom_archivo: "Reporte Inbound - Detalle de llamadas.xls",
      vm_ruta_archivo: "reportes/",
    });

    const generationResponse = await fetch(reportUrl, {
      method: "POST",
      cache: "no-store",
      headers: {
        ...commonHeaders,
        accept: "*/*",
        "content-type": "application/x-www-form-urlencoded; charset=UTF-8",
        origin: portalOrigin,
        referer: reportPageUrl,
        "x-requested-with": "XMLHttpRequest",
      },
      body: payload,
    });

    if (!generationResponse.ok) {
      return Response.json({ error: `A365 no pudo generar el reporte (${generationResponse.status}).` }, { status: 502 });
    }

    // El generador de A365 responde vacío: deja el XLS en una ruta temporal y la
    // interfaz web lo obtiene con una segunda solicitud al endpoint de descarga.
    const downloadParams = new URLSearchParams({
      vm_ruta: "reportes/",
      vm_nom_archivo: "Reporte Inbound - Detalle de llamadas.xls",
    });
    let report: ArrayBuffer | null = null;
    let contentType = "application/vnd.ms-excel";

    for (let attempt = 0; attempt < 4; attempt += 1) {
      const downloadResponse = await fetch(`${downloadUrl}?${downloadParams}`, {
        cache: "no-store",
        headers: {
          ...commonHeaders,
          accept: "*/*",
          referer: reportPageUrl,
        },
      });

      if (downloadResponse.ok) {
        const candidate = await downloadResponse.arrayBuffer();
        if (candidate.byteLength > 0) {
          report = candidate;
          contentType = downloadResponse.headers.get("content-type") ?? contentType;
          break;
        }
      }

      if (attempt < 3) await wait(1000);
    }

    if (!report) {
      return Response.json({ error: "A365 aceptó la solicitud, pero aún no generó el archivo descargable. Inténtalo nuevamente en unos segundos." }, { status: 502 });
    }

    if (looksLikeLoginPage(report, contentType)) {
      return Response.json({ error: "A365 devolvió una página de inicio de sesión. Copia un PHPSESSID vigente." }, { status: 401 });
    }
    if (looksLikeVicidialDatabaseError(report)) {
      return Response.json({ error: "A365/Vicidial reporta una falla de conexión con su base de datos. El reporte no se generó: espera unos minutos e inténtalo de nuevo; si persiste, repórtalo al administrador de A365." }, { status: 503 });
    }

    return new Response(report, {
      headers: {
        "content-type": contentType,
        "content-disposition": `attachment; filename="reporte-inbound-${startDate}-${endDate}.xls"`,
      },
    });
  } catch (error) {
    console.error("[inbound-calls-report] A365 request failed", error instanceof Error ? error.message : String(error));
    return Response.json({ error: "No pude conectar con A365 desde este equipo. Verifica tu red, VPN o que el portal esté disponible." }, { status: 502 });
  }
}
