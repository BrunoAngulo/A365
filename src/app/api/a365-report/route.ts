const reportBaseUrl = "https://applinde.a365.com.pe/reportes/reporte_matriz";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 30;

type ReportRequest = {
  startDate?: string;
  endDate?: string;
  sessionId?: string;
};

function isValidDate(value: string) {
  return /^\d{4}-\d{2}-\d{2}$/.test(value);
}

function isValidSession(value: string) {
  return /^[A-Za-z0-9]+$/.test(value);
}

function normalizeSession(value: string) {
  const trimmed = value.trim();
  const cookieMatch = trimmed.match(/(?:^|;\s*)PHPSESSID=([^;\s]+)/i) ?? trimmed.match(/PHPSESSID=([^;\s]+)/i);
  return (cookieMatch?.[1] ?? trimmed).replace(/^["']|["']$/g, "").trim();
}

function decodePreview(report: ArrayBuffer) {
  return new TextDecoder().decode(report.slice(0, 5000)).toLowerCase();
}

function looksLikeLoginPage(report: ArrayBuffer, contentType: string) {
  if (!contentType.toLowerCase().includes("text/html")) {
    return false;
  }

  const preview = decodePreview(report);
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

export async function POST(request: Request) {
  let body: ReportRequest;

  try {
    body = (await request.json()) as ReportRequest;
  } catch {
    return Response.json({ error: "Solicitud invalida." }, { status: 400 });
  }

  const startDate = body.startDate?.trim() ?? "";
  const endDate = body.endDate?.trim() ?? "";
  const sessionId = normalizeSession(body.sessionId ?? "");

  if (!isValidDate(startDate) || !isValidDate(endDate)) {
    return Response.json({ error: "Selecciona fechas validas." }, { status: 400 });
  }

  if (!isValidSession(sessionId)) {
    return Response.json({ error: "Ingresa un PHPSESSID valido." }, { status: 400 });
  }

  const upstreamUrl = `${reportBaseUrl}/${startDate}/${endDate}/`;
  let upstreamResponse: Response;

  try {
    upstreamResponse = await fetch(upstreamUrl, {
      cache: "no-store",
      redirect: "follow",
      headers: {
        accept: "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;q=0.8,*/*;q=0.7",
        "accept-language": "es-419,es;q=0.9",
        cookie: `PHPSESSID=${sessionId}`,
        referer: "https://applinde.a365.com.pe/reportes/",
        "sec-fetch-dest": "document",
        "sec-fetch-mode": "navigate",
        "sec-fetch-site": "same-origin",
        "upgrade-insecure-requests": "1",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
      },
    });
  } catch (error) {
    console.error("[a365-report] upstream fetch failed", {
      message: error instanceof Error ? error.message : String(error),
      range: `${startDate}..${endDate}`,
    });

    return Response.json(
      { error: "Vercel no pudo conectarse con A365. Puede ser bloqueo de red/IP del portal." },
      { status: 502 },
    );
  }

  if (!upstreamResponse.ok) {
    console.error("[a365-report] upstream bad status", {
      status: upstreamResponse.status,
      statusText: upstreamResponse.statusText,
      contentType: upstreamResponse.headers.get("content-type"),
      range: `${startDate}..${endDate}`,
    });

    return Response.json(
      { error: `No pude descargar el reporte. Codigo ${upstreamResponse.status}.` },
      { status: upstreamResponse.status },
    );
  }

  const report = await upstreamResponse.arrayBuffer();
  const contentType = upstreamResponse.headers.get("content-type") ?? "application/octet-stream";

  if (looksLikeLoginPage(report, contentType)) {
    console.error("[a365-report] upstream returned login page", {
      contentType,
      range: `${startDate}..${endDate}`,
    });

    return Response.json(
      { error: "A365 devolvio login o sesion expirada. En Vercel la sesion puede fallar porque la peticion sale desde otra IP." },
      { status: 401 },
    );
  }

  const contentDisposition =
    upstreamResponse.headers.get("content-disposition") ??
    `attachment; filename="reporte-matriz-${startDate}-${endDate}.xlsx"`;

  return new Response(report, {
    headers: {
      "content-disposition": contentDisposition,
      "content-type": contentType,
    },
  });
}
