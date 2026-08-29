const reportBaseUrl = "https://applinde.a365.com.pe/reportes/reporte_matriz";

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

export async function POST(request: Request) {
  let body: ReportRequest;

  try {
    body = (await request.json()) as ReportRequest;
  } catch {
    return Response.json({ error: "Solicitud invalida." }, { status: 400 });
  }

  const startDate = body.startDate?.trim() ?? "";
  const endDate = body.endDate?.trim() ?? "";
  const sessionId = body.sessionId?.trim() ?? "";

  if (!isValidDate(startDate) || !isValidDate(endDate)) {
    return Response.json({ error: "Selecciona fechas validas." }, { status: 400 });
  }

  if (!isValidSession(sessionId)) {
    return Response.json({ error: "Ingresa un PHPSESSID valido." }, { status: 400 });
  }

  const upstreamUrl = `${reportBaseUrl}/${startDate}/${endDate}/`;
  const upstreamResponse = await fetch(upstreamUrl, {
    headers: {
      accept: "text/html,application/xhtml+xml,application/xml;q=0.9,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;q=0.8,*/*;q=0.7",
      cookie: `PHPSESSID=${sessionId}`,
      referer: "https://applinde.a365.com.pe/reportes/",
      "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
    },
  });

  if (!upstreamResponse.ok) {
    return Response.json(
      { error: `No pude descargar el reporte. Codigo ${upstreamResponse.status}.` },
      { status: upstreamResponse.status },
    );
  }

  const report = await upstreamResponse.arrayBuffer();
  const contentType = upstreamResponse.headers.get("content-type") ?? "application/octet-stream";
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
