import { NextResponse } from "next/server";

const queues = ["INLINDEINDCEN", "INLINDEINDDES", "INLINDEINDNOR", "INLINDEINDORI", "INLINDEINDSUR", "INLINDEMEDCEN", "INLINDEMEDDES", "INLINDEMEDNOR", "INLINDEMEDORI", "INLINDEMEDSUR", "INLINDEURGCEN", "INLINDEURGDES", "INLINDEURGNORTE", "INLINDEURGORI", "INLINDEURGSUR", "INLINDECLI", "INLINDESAC", "INLINDETERA"];

export async function POST(request: Request) {
  const body = await request.json().catch(() => null) as { token?: string; startDate?: string; endDate?: string } | null;
  if (!body?.token || !body.startDate || !body.endDate) return NextResponse.json({ error: "Faltan token y fechas" }, { status: 400 });
  const url = new URL("https://tmkreport15.a365.com.pe/api/inbound-metrics");
  url.searchParams.set("connectionId", "a0bb59bd-b685-4ad5-8129-03781adb7d6b");
  url.searchParams.set("campaignId", "LINDESAC");
  url.searchParams.set("startDate", body.startDate);
  url.searchParams.set("endDate", body.endDate);
  queues.forEach((queue) => url.searchParams.append("queues[]", queue));
  const headers = { accept: "application/json", authorization: `Bearer ${body.token}` };
  const response = await fetch(url, { headers, cache: "no-store" });
  if (!response.ok) return NextResponse.json({ error: `La API respondió ${response.status}` }, { status: response.status });
  // Algunas respuestas del proveedor llegan vacías (o con HTML de error) aunque
  // el status sea 2xx. Leer como texto evita que Response.json() rompa la ruta.
  const responseText = await response.text();
  if (!responseText.trim()) {
    return NextResponse.json({ error: "La API devolvió una respuesta vacía." }, { status: 502 });
  }
  let payload: { data?: { nivelesServicio?: Array<{ range?: string; count?: number; percentage?: number }>; resumen?: Array<{ fecha?: string; llamadasRecibidas?: number; llamadasHorario?: number; atendidas?: number; abandonadas?: number; nivelAtencion?: number; nivelAbandono?: number; promedioEspera?: number; asesores?: number }> } };
  try {
    payload = JSON.parse(responseText) as typeof payload;
  } catch {
    return NextResponse.json({ error: "La API devolvió un formato no válido." }, { status: 502 });
  }
  const cumulative = payload.data?.nivelesServicio ?? [];
  const total = Number(cumulative.at(-1)?.count ?? 0);
  const rows = cumulative.map((item, index) => ({ range: index === 0 ? "0-5s" : item.range === ">30s" ? ">30s" : `${Number(cumulative[index - 1]?.range?.replace(/\D/g, "") || 0)}-${Number(item.range?.replace(/\D/g, "") || 0)}s`, count: index === 0 ? Number(item.count ?? 0) : item.range === ">30s" ? Math.max(0, total - Number(cumulative[index - 1]?.count ?? 0)) : Math.max(0, Number(item.count ?? 0) - Number(cumulative[index - 1]?.count ?? 0)) })).map((item) => ({ ...item, percentage: total ? (item.count / total) * 100 : 0 }));
  const daily = (payload.data?.resumen ?? []).map((item) => ({ fecha: item.fecha ?? "", recibidas: Number(item.llamadasRecibidas ?? 0), horario: Number(item.llamadasHorario ?? 0), atendidas: Number(item.atendidas ?? 0), abandonadas: Number(item.abandonadas ?? 0), nivelAtencion: Number(item.nivelAtencion ?? 0), nivelAbandono: Number(item.nivelAbandono ?? 0), promedioEspera: Number(item.promedioEspera ?? 0), asesores: Number(item.asesores ?? 0), atendidasMenos20: 0 }));

  // nivelesServicio es acumulado para todo el rango. Consultamos una vez por
  // mes para que la serie mensual conserve el conteo correcto de <20 segundos.
  const start = new Date(`${body.startDate}T00:00:00`);
  const end = new Date(`${body.endDate}T00:00:00`);
  const months: Array<{ key: string; start: string; end: string }> = [];
  for (const cursor = new Date(start.getFullYear(), start.getMonth(), 1); cursor <= end; cursor.setMonth(cursor.getMonth() + 1)) {
    const monthStart = new Date(Math.max(start.getTime(), cursor.getTime()));
    const monthEnd = new Date(Math.min(end.getTime(), new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0).getTime()));
    months.push({ key: cursor.toISOString().slice(0, 7), start: monthStart.toISOString().slice(0, 10), end: monthEnd.toISOString().slice(0, 10) });
  }
  const under20ByMonth = new Map<string, number>();
  await Promise.all(months.map(async ({ key, start: monthStart, end: monthEnd }) => {
    const monthUrl = new URL("https://tmkreport15.a365.com.pe/api/inbound-metrics");
    monthUrl.searchParams.set("connectionId", "a0bb59bd-b685-4ad5-8129-03781adb7d6b");
    monthUrl.searchParams.set("campaignId", "LINDESAC");
    monthUrl.searchParams.set("startDate", monthStart);
    monthUrl.searchParams.set("endDate", monthEnd);
    queues.forEach((queue) => monthUrl.searchParams.append("queues[]", queue));
    try {
      const monthResponse = await fetch(monthUrl, { headers, cache: "no-store" });
      if (!monthResponse.ok) return;
      const text = await monthResponse.text();
      if (!text.trim()) return;
      const monthPayload = JSON.parse(text) as { data?: { nivelesServicio?: Array<{ range?: string; count?: number }> } };
      const level = monthPayload.data?.nivelesServicio?.find((item) => String(item.range).replace(/\s/g, "").toLowerCase() === "<20s");
      if (level) under20ByMonth.set(key, Number(level.count ?? 0));
    } catch {
      // El resto de métricas sigue disponible aunque un mes no responda.
    }
  }));
  daily.forEach((item, index) => {
    const month = item.fecha.slice(0, 7);
    item.atendidasMenos20 = index === daily.findIndex((candidate) => candidate.fecha.slice(0, 7) === month) ? under20ByMonth.get(month) ?? 0 : 0;
  });
  return NextResponse.json({ rows, daily });
}
