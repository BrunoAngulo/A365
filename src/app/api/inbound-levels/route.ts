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
  const response = await fetch(url, { headers: { accept: "application/json", authorization: `Bearer ${body.token}` }, cache: "no-store" });
  if (!response.ok) return NextResponse.json({ error: `La API respondió ${response.status}` }, { status: response.status });
  const payload = await response.json() as { data?: { nivelesServicio?: Array<{ range?: string; count?: number; percentage?: number }>; resumen?: Array<{ fecha?: string; llamadasRecibidas?: number; llamadasHorario?: number; atendidas?: number; abandonadas?: number; nivelAtencion?: number; nivelAbandono?: number; promedioEspera?: number; asesores?: number }> } };
  const cumulative = payload.data?.nivelesServicio ?? [];
  const total = Number(cumulative.at(-1)?.count ?? 0);
  const rows = cumulative.map((item, index) => ({ range: index === 0 ? "0-5s" : item.range === ">30s" ? ">30s" : `${Number(cumulative[index - 1]?.range?.replace(/\D/g, "") || 0)}-${Number(item.range?.replace(/\D/g, "") || 0)}s`, count: index === 0 ? Number(item.count ?? 0) : item.range === ">30s" ? Math.max(0, total - Number(cumulative[index - 1]?.count ?? 0)) : Math.max(0, Number(item.count ?? 0) - Number(cumulative[index - 1]?.count ?? 0)) })).map((item) => ({ ...item, percentage: total ? (item.count / total) * 100 : 0 }));
  const daily = (payload.data?.resumen ?? []).map((item) => ({ fecha: item.fecha ?? "", recibidas: Number(item.llamadasRecibidas ?? 0), horario: Number(item.llamadasHorario ?? 0), atendidas: Number(item.atendidas ?? 0), abandonadas: Number(item.abandonadas ?? 0), nivelAtencion: Number(item.nivelAtencion ?? 0), nivelAbandono: Number(item.nivelAbandono ?? 0), promedioEspera: Number(item.promedioEspera ?? 0), asesores: Number(item.asesores ?? 0) }));
  return NextResponse.json({ rows, daily });
}
