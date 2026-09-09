"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, FormEvent, MouseEvent as ReactMouseEvent, ReactNode, RefObject } from "react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import { readSheet } from "read-excel-file/browser";
import type { Row } from "read-excel-file/browser";
import {
  CalendarDays,
  Clock3,
  ClipboardList,
  Download,
  ArrowDownAZ,
  FileSpreadsheet,
  Hash,
  Info,
  ListFilter,
  Mail,
  Maximize2,
  Phone,
  Timer,
  Upload,
  UserRound,
  X,
} from "lucide-react";
import { toPng } from "html-to-image";
import styles from "./page.module.css";

type RawRow = Record<string, unknown>;

declare global {
  interface Window {
    __TAURI_INTERNALS__?: unknown;
  }
}

type TauriMatrixReport = {
  dataBase64: string;
  contentType: string;
  fileName: string;
};

type ExtensionMatrixResponse = {
  source?: string;
  type?: string;
  requestId?: string;
  ok?: boolean;
  error?: string;
  dataBase64?: string;
  contentType?: string;
};

type MetricRow = {
  id: number;
  date: string;
  week: string;
  month: string;
  hour: string;
  phone: string;
  user: string;
  campaignId: string;
  statusName: string;
  listName: string;
  direction: string;
  durationSeconds: number;
  leadId: string;
  waitSeconds: number | null;
  raw: Record<string, string>;
};

type ChartPoint = {
  name: string;
  total: number;
};

type UserPerformancePoint = { name: string; received: number; attended: number; attentionRate: number };

type FilterField = "date" | "week" | "month" | "hour" | "statusName" | "campaignId" | "user" | "listName";
type ChartKind = "date" | "hour" | "status" | "campaign";
type DashboardView = "calls" | "matrix" | "sales" | "errors" | "performance";
type TimelineScale = "30m" | "1h" | "day";

type ChartClickState = {
  activeLabel?: string | number;
  activePayload?: Array<{ payload?: ChartPoint; name?: string | number; value?: string | number }>;
};

type DashboardSummary = {
  total: number;
  byDate: ChartPoint[];
  byWeek: ChartPoint[];
  byMonth: ChartPoint[];
  byHour: ChartPoint[];
  byStatus: ChartPoint[];
  byListName: ChartPoint[];
  byCampaign: ChartPoint[];
  byUser: ChartPoint[];
  uniquePhones: number;
  uniqueUsers: number;
  uniqueCampaigns: number;
  uniqueStatuses: number;
  attendedCalls: number;
};

type MatrixRow = {
  id: number;
  subjectEmail: string;
  dateEmail: string;
  fromEmail: string;
  tipificacion: string;
  motivo: string;
  fechaAsignacion: string;
  usuarioAsignado: string;
  estadoRegistro: string;
  fechaRegistro: string;
  nroMov: string;
  idEmail: string;
  minutesToAssign: number | null;
  minutesToRegister: number | null;
  minutesTotal: number | null;
  raw: Record<string, string>;
};

type MatrixSummary = {
  total: number;
  avgToAssign: number | null;
  avgToRegister: number | null;
  avgTotal: number | null;
  completed: number;
  byTipificacion: ChartPoint[];
  byEstado: ChartPoint[];
  byUser: ChartPoint[];
  avgByAgent: TimeAveragePoint[];
  avgByDay: TimeAveragePoint[];
  avgByMonth: TimeAveragePoint[];
};

type IncidentRow = {
  id: number;
  startTime: string;
  endTime: string;
  email: string;
  agent: string;
  branch: string;
  incidentDate: string;
  customerCode: string;
  customerName: string;
  inconsistencyType: string;
  channel: string;
  subject: string;
  detail: string;
  impact: string;
  raw: Record<string, string>;
};

type IncidentSummary = {
  total: number;
  uniqueAgents: number;
  uniqueTypes: number;
  uniqueBranches: number;
  byAgent: ChartPoint[];
  byType: ChartPoint[];
  byBranch: ChartPoint[];
  byMonth: ChartPoint[];
};

type AgentIndicatorRow = {
  agent: string;
  calls: number;
  emails: number;
  errors: number;
  attended: number;
  productivity: number;
  quality: number;
  effectiveness: number;
};

type PerformanceSummary = {
  totalCalls: number;
  totalEmails: number;
  totalAttended: number;
  totalErrors: number;
  productivity: number;
  quality: number;
  effectiveness: number;
  byAgent: AgentIndicatorRow[];
};

type TimeAveragePoint = {
  name: string;
  asignacion: number;
  resolucion: number;
  total: number;
  registros: number;
};

type TimelineBlock = {
  id: string;
  agent: string;
  left: number;
  width: number;
  subject: string;
  range: string;
  detail: string;
  tipificacion: string;
  estado: string;
  duration: number | null;
  idEmail: string;
  color: string;
};

type AgentTimeline = {
  agent: string;
  blocks: TimelineBlock[];
};

const detailRowLimit = 500;
const filterFields: FilterField[] = ["date", "week", "month", "hour", "statusName", "campaignId", "user", "listName"];
const todayInputValue = formatLocalDate(new Date());
const slaStartMinutes = 7 * 60;
const slaEndMinutes = 23 * 60 + 1;
const slaWindowMinutes = slaEndMinutes - slaStartMinutes;
const timelineScales: Array<{ value: TimelineScale; label: string }> = [
  { value: "30m", label: "30 min" },
  { value: "1h", label: "1 hora" },
  { value: "day", label: "Todo el dia" },
];

const chartColors = [
  "#1d4ed8",
  "#3b82f6",
  "#0f766e",
  "#0891b2",
  "#475569",
  "#64748b",
  "#0284c7",
  "#0e7490",
  "#334155",
  "#0369a1",
  "#1e40af",
  "#166534",
];

const heatmapDays = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];
const heatmapStartHour = 7;
const heatmapEndHour = 23;
const persistedCallsKey = "a365-i1-calls-v1";
const persistedDashboardKey = "a365-dashboard-state-v1";

type PersistedCalls = { fileName: string; columns: string[]; rows: MetricRow[] };
type PersistedDashboardState = {
  matrixRows: MatrixRow[];
  matrixColumns: string[];
  matrixRangeLabel: string;
  matrixFileName: string;
  dashboardFocus: Partial<Record<FilterField, string>>;
  selectedStatusNames: string[];
  statusFocus: string | null;
  callStartDate: string;
  callEndDate: string;
  activeView: DashboardView;
  timelineDay: string;
  timelineScale: TimelineScale;
  showSlowResolutionOnly: boolean;
  emailHourChronological: boolean;
};

type SalesRow = {
  id: number;
  year: string;
  month: string;
  day: string;
  date: string;
  advancedSale: string;
  company: string;
  region: string;
  branch: string;
  seller: string;
  sellerMerchandise: string;
  clientCode: string;
  clientName: string;
  ruc: string;
  documentType: string;
  productCode: string;
  productDescription: string;
  unit: string;
  className: string;
  group: string;
  business: string;
  industry: string;
  application: string;
  branchLine: string;
  volumeSold: number;
  conversionFactor: number;
  tons: number;
  amountPEN: number;
  amountUSD: number;
  raw: Record<string, string>;
};

type SalesTab = "summary" | "evolution" | "branches" | "products" | "clients" | "industries" | "advanced";

const salesColumnAliases = {
  year: ["año", "ano", "year"], month: ["mes", "month"], day: ["dia", "día", "day"], advancedSale: ["venta adelantada?", "venta adelantada", "venta adelantada ?"], company: ["compañia", "compania", "company"], region: ["región", "region"], branch: ["sucursal"], seller: ["nombre vendedor"], sellerMerchandise: ["nom.vend.mercaderia", "nom vend mercaderia"], clientCode: ["cod.clte", "cod clte", "codigo cliente"], clientName: ["nombre del cliente"], ruc: ["ruc"], documentType: ["tipo documento"], productCode: ["cod.producto", "cod producto"], productDescription: ["descripcion", "descripción"], unit: ["unidad"], className: ["clase"], group: ["grupo"], business: ["negocio"], industry: ["industria"], application: ["aplicación", "aplicacion"], branchLine: ["rama"], volumeSold: ["volumen vendido"], conversionFactor: ["factor conversion a kg.", "factor conversión a kg", "factor conversion a kg"], tons: ["volumen en tonelada"], amountPEN: ["importe vendido s/.", "importe vendido s/", "importe vendido soles"], amountUSD: ["importe vendido us$", "importe vendido usd", "importe vendido us$."] as string[],
};

function salesNumber(value: unknown) {
  const text = valueAsText(value).replace(/[^0-9,.-]/g, "").trim();
  if (!text) return 0;
  const normalized = text.includes(",") && text.includes(".") ? text.replace(/\./g, "").replace(",", ".") : text.replace(",", ".");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function salesDisplayName(value: string) {
  return value.replace(/^\s*\d+\s*-\s*/, "").trim() || "Sin dato";
}

function mapSalesRows(sheetRows: Row[]) {
  const rawRows = rowsFromSheetRows(sheetRows);
  if (!rawRows.length) return [];
  const headers = Object.keys(rawRows[0]);
  const columns = Object.fromEntries(Object.entries(salesColumnAliases).map(([key, aliases]) => [key, findColumn(headers, aliases)])) as Record<string, string | undefined>;
  const required = ["year", "month", "day", "region", "branch", "clientCode", "productCode", "className", "amountUSD"];
  const missing = required.filter((key) => !columns[key]);
  if (missing.length) throw new Error(`Faltan columnas de RESUMEN: ${missing.join(", ")}`);
  return rawRows.map((row, index) => {
    const text = (key: string) => valueAsText(columns[key] ? row[columns[key] as string] : "");
    const year = text("year");
    const month = text("month").padStart(2, "0");
    const day = text("day").padStart(2, "0");
    const date = `${year}-${month}-${day}`;
    return { id: index + 1, year, month, day, date, advancedSale: text("advancedSale"), company: text("company"), region: text("region"), branch: text("branch") || "Sin sucursal", seller: text("seller") || "Sin vendedor", sellerMerchandise: text("sellerMerchandise"), clientCode: text("clientCode") || "Sin cliente", clientName: text("clientName") || "Sin cliente", ruc: text("ruc"), documentType: text("documentType") || "Sin documento", productCode: text("productCode") || "Sin producto", productDescription: text("productDescription") || "Sin producto", unit: text("unit") || "Sin unidad", className: text("className") || "Sin clase", group: text("group") || "Sin grupo", business: text("business") || "Sin negocio", industry: text("industry") || "Sin industria", application: text("application") || "Sin aplicación", branchLine: text("branchLine"), volumeSold: salesNumber(columns.volumeSold ? row[columns.volumeSold] : ""), conversionFactor: salesNumber(columns.conversionFactor ? row[columns.conversionFactor] : ""), tons: salesNumber(columns.tons ? row[columns.tons] : ""), amountPEN: salesNumber(columns.amountPEN ? row[columns.amountPEN] : ""), amountUSD: salesNumber(columns.amountUSD ? row[columns.amountUSD] : ""), raw: headers.reduce<Record<string, string>>((record, header) => { record[header] = valueAsText(row[header]); return record; }, {}) } satisfies SalesRow;
  }).filter((row) => row.year && row.month && row.day);
}

function salesCountDistinct(rows: SalesRow[], key: keyof SalesRow) {
  return new Set(rows.map((row) => String(row[key])).filter(Boolean)).size;
}

function salesGrouped(rows: SalesRow[], key: keyof SalesRow, valueKey: keyof SalesRow = "amountUSD") {
  const map = new Map<string, number>();
  rows.forEach((row) => { const name = String(row[key] || "Sin dato"); map.set(name, (map.get(name) ?? 0) + Number(row[valueKey] ?? 0)); });
  return Array.from(map.entries()).map(([name, total]) => ({ name, total })).sort((a, b) => key === "month" || key === "date" ? a.name.localeCompare(b.name) : b.total - a.total);
}

function salesPurchaseFrequency(rows: SalesRow[]) {
  const byClient = new Map<string, { name: string; dates: number[]; total: number }>();
  rows.forEach((row) => {
    const key = row.clientCode || row.clientName || "Sin cliente";
    const item = byClient.get(key) ?? { name: row.clientName || key, dates: [], total: 0 };
    const timestamp = Date.parse(`${row.date}T00:00:00Z`);
    if (Number.isFinite(timestamp)) item.dates.push(timestamp);
    item.total += row.amountUSD;
    byClient.set(key, item);
  });
  return Array.from(byClient.values()).map((item) => {
    const dates = Array.from(new Set(item.dates)).sort((a, b) => a - b);
    const intervals = dates.slice(1).map((date, index) => (date - dates[index]) / 86_400_000);
    return { name: item.name, total: intervals.length ? intervals.reduce((sum, days) => sum + days, 0) / intervals.length : 0, purchases: dates.length, salesTotal: item.total };
  }).filter((item) => item.purchases > 1).sort((a, b) => b.salesTotal - a.salesTotal).slice(0, 10);
}

function formatSalesUSD(value: number) {
  const absolute = Math.abs(value);
  if (absolute >= 1_000_000) return `USD ${(value / 1_000_000).toFixed(1)} M`;
  if (absolute >= 1_000) return `USD ${(value / 1_000).toFixed(1)} k`;
  return `USD ${value.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}
function formatSalesPEN(value: number) { return `S/ ${value.toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`; }

function readPersistedCalls(): PersistedCalls | null {
  if (typeof window === "undefined") return null;
  try {
    const parsed = JSON.parse(window.localStorage.getItem(persistedCallsKey) ?? "null") as Partial<PersistedCalls> | null;
    return parsed && Array.isArray(parsed.rows) && parsed.rows.length
      ? { fileName: parsed.fileName || "Archivo recuperado", columns: parsed.columns ?? Object.keys(parsed.rows[0]?.raw ?? {}), rows: parsed.rows }
      : null;
  } catch {
    return null;
  }
}

function readPersistedDashboard(): Partial<PersistedDashboardState> | null {
  if (typeof window === "undefined") return null;
  try {
    return JSON.parse(window.localStorage.getItem(persistedDashboardKey) ?? "null") as Partial<PersistedDashboardState> | null;
  } catch {
    return null;
  }
}
const heatmapHours = Array.from(
  { length: heatmapEndHour - heatmapStartHour + 1 },
  (_, offset) => heatmapStartHour + offset,
);

type HeatmapCell = { day: string; hour: number; value: number; count: number };

function heatmapDayIndex(value: string) {
  const date = parseDateTimeValue(value);
  if (!date) return null;
  const day = date.getDay();
  return day === 0 ? 6 : day - 1;
}

function heatmapHour(value: string) {
  const match = valueAsText(value).match(/(?:T|\s|^)(\d{1,2})(?::\d{2})?/);
  if (!match) return null;
  const hour = Number(match[1]);
  return hour >= 0 && hour <= 23 ? hour : null;
}

function buildHeatmapCells(rows: Array<{ date: string; hour?: string; value?: number }>) {
  const cells = heatmapDays.flatMap((day) =>
    heatmapHours.map((hour) => ({ day, hour, value: 0, count: 0 })),
  );
  rows.forEach((row) => {
    const dayIndex = heatmapDayIndex(row.date);
    const hour = heatmapHour(row.hour ?? row.date);
    if (dayIndex === null || hour === null || hour < heatmapStartHour || hour > heatmapEndHour) return;
    const cell = cells[dayIndex * heatmapHours.length + hour - heatmapStartHour];
    cell.value += row.value ?? 1;
    cell.count += 1;
  });
  return cells;
}

function buildCallHeatmap(rows: MetricRow[]) {
  return buildHeatmapCells(rows.map((row) => ({ date: row.date, hour: row.hour })));
}

function buildEmailHeatmap(rows: MatrixRow[]) {
  return buildHeatmapCells(rows.flatMap((row) => {
    const date = parseDateTimeValue(row.dateEmail);
    if (!date) return [];
    return [{
      date: row.dateEmail,
      hour: `${String(date.getHours()).padStart(2, "0")}:00`,
      value: 1,
    }];
  }));
}

function HeatmapGrid({ cells, formatter }: { cells: HeatmapCell[]; formatter: (value: number) => string }) {
  const max = Math.max(...cells.map((cell) => cell.value), 0);
  return (
    <div className={styles.heatmapWrap}>
      <div className={styles.heatmapGrid}>
        <div className={styles.heatmapCorner}>Día / hora</div>
        {heatmapHours.map((hour) => <span key={hour} className={styles.heatmapAxis}>{String(hour).padStart(2, "0")}</span>)}
        {heatmapDays.flatMap((day, dayIndex) => [
          <span key={`${day}-label`} className={styles.heatmapRowLabel}>{day}</span>,
          ...cells.slice(dayIndex * heatmapHours.length, dayIndex * heatmapHours.length + heatmapHours.length).map((cell) => {
            const intensity = max ? 0.08 + (cell.value / max) * 0.92 : 0.08;
            const ratio = max ? cell.value / max : 0;
            const rgb = ratio >= 0.66 ? "220, 38, 38" : ratio >= 0.33 ? "249, 115, 22" : "250, 204, 21";
            const style = { "--heatmap-alpha": intensity, "--heatmap-value": cell.value, backgroundColor: `rgba(${rgb}, ${0.2 + ratio * 0.8})` } as CSSProperties;
            return <div key={`${cell.day}-${cell.hour}`} className={styles.heatmapCell} style={{ ...style, color: ratio >= 0.55 ? "#ffffff" : "#0f172a" }} title={`${cell.day} ${String(cell.hour).padStart(2, "0")}:00 · ${formatter(cell.value)}`}><span>{cell.value.toLocaleString("es-PE")}</span></div>;
          }),
        ])}
      </div>
      <div className={styles.heatmapLegend}><span>Menor</span><i /><span>Mayor</span></div>
    </div>
  );
}

const columnAliases = {
  date: ["date", "fecha", "created_at", "createdat", "call_date", "call date", "fecha_creacion", "created"],
  hour: ["hour", "hora", "time", "call_time", "call time", "created_time", "created time"],
  phone: ["phone", "telefono", "teléfono", "mobile", "celular", "number", "phone_number", "phone number", "numero", "número"],
  user: ["user", "usuario", "agent", "asesor", "owner", "ejecutivo", "username", "user_name", "user name"],
  campaignId: ["campaign_id", "campaign id", "campaign", "campaña", "campana", "id_campana", "id campaña"],
  statusName: ["status_name", "status name", "estado_nombre", "nombre_estado"],
  listName: ["list_name", "list name", "lista", "nombre_lista", "nombre lista"],
  direction: ["direction", "call_direction", "call direction", "direccion", "dirección"],
  durationSeconds: ["length_in_sec", "length in sec", "talk_time", "talk time"],
  leadId: ["lead_id", "lead id"],
  waitSeconds: ["wait_time", "queue_time", "answer_time", "ring_time", "answer_seconds", "queue_seconds", "time_to_answer"],
};

type EmailMonthPoint = { name: string; received: number; attended: number; withinSla: number; nda: number; nds: number; tmoMinutes: number | null };
type EmailSummary = { received: number; attended: number; withinSla: number; nda: number; nds: number; tmoMinutes: number | null; byMonth: EmailMonthPoint[]; byUser: ChartPoint[]; byTipificacion: ChartPoint[]; byMotivo: ChartPoint[]; byHour: ChartPoint[]; period: string };

const EMAIL_SLA_MINUTES = 60;

type MonthlyCallPoint = {
  name: string;
  received: number;
  attended: number;
};

type ServiceMonthlyPoint = {
  name: string;
  received: number;
  attended: number;
  attendedUnder20: number;
  abandoned: number;
  attentionRate: number;
  serviceRate: number | null;
  abandonmentRate: number;
  averageDurationSeconds: number;
};

type ServiceSummary = {
  attentionRate: number;
  serviceRate: number | null;
  abandonmentRate: number;
  averageDurationSeconds: number;
  byMonth: ServiceMonthlyPoint[];
};

const matrixColumnAliases = {
  subjectEmail: ["subject_email", "subject email", "asunto"],
  dateEmail: ["date_email", "date email", "fecha email", "fecha de email"],
  fromEmail: ["from_email", "from email", "correo origen", "remitente"],
  tipificacion: ["tipificacion", "tipificación"],
  motivo: ["motivo"],
  fechaAsignacion: ["fecha_asignacion", "fecha asignacion", "fecha asignación"],
  usuarioAsignado: ["usuario asignado", "usuario_asignado", "asesor asignado"],
  estadoRegistro: ["estado de registro", "estado_registro"],
  fechaRegistro: ["fecha de registro", "fecha_registro"],
  nroMov: ["nro_mov", "nro mov", "numero movimiento"],
  idEmail: ["id_email", "id email"],
};

const incidentColumnAliases = {
  startTime: ["hora de inicio", "inicio", "start time"],
  endTime: ["hora de finalizacion", "hora de finalización", "finalizacion", "finalización", "end time"],
  email: ["correo electronico", "correo electrónico", "email"],
  agent: ["nombre", "agente", "usuario", "asesor"],
  branch: ["a que sucursal pertenece el cliente", "sucursal", "plaza"],
  incidentDate: ["fecha de incidente", "fecha incidente"],
  customerCode: ["codigo de cliente", "código de cliente", "cod cliente"],
  customerName: ["razon social del cliente", "razón social del cliente", "cliente"],
  inconsistencyType: ["tipo de inconsistencia", "inconsistencia", "observacion", "observación"],
  channel: ["medio de comunicacion", "medio de comunicación", "canal"],
  subject: ["asunto del correo", "asunto"],
  detail: ["detalle de la incidencia", "breve explicacion", "breve explicación", "detalle"],
  impact: ["impacto"],
};

function normalizeKey(value: string) {
  return value
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function findColumn(headers: string[], aliases: string[]) {
  const normalizedHeaders = headers.map((header) => ({
    original: header,
    normalized: normalizeKey(header),
  }));
  const normalizedAliases = aliases.map(normalizeKey);

  return (
    normalizedHeaders.find((header) => normalizedAliases.includes(header.normalized))?.original ??
    normalizedHeaders.find((header) =>
      normalizedAliases.some((alias) => header.normalized.includes(alias)),
    )?.original
  );
}

function valueAsText(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value).trim();
}

function secondsFromValue(value: unknown) {
  const seconds = Number(valueAsText(value).replace(",", "."));
  return Number.isFinite(seconds) && seconds >= 0 ? seconds : 0;
}

function formatDateTimeMinute(value: unknown) {
  const text = valueAsText(value);
  const parsed = parseDateParts(text, "");

  if (parsed.date === "Sin fecha") {
    return text;
  }

  const [year, month, day] = parsed.date.split("-");
  const timeMatch = text.match(/(?:T|\s)(\d{1,2}):(\d{2})/);
  const hourMinute = timeMatch
    ? `${String(Number(timeMatch[1])).padStart(2, "0")}:${timeMatch[2]}`
    : parsed.hour.replace(":00", ":00");

  return `${day}/${month}/${year} ${hourMinute}`;
}

function displayCellValue(column: string, value: string) {
  const normalizedColumn = normalizeKey(column);

  if (normalizedColumn === "call date" || normalizedColumn.includes("fecha") || normalizedColumn.includes("date")) {
    return formatDateTimeMinute(value);
  }

  return value;
}

function formatLocalDate(value: Date) {
  return [
    value.getFullYear(),
    String(value.getMonth() + 1).padStart(2, "0"),
    String(value.getDate()).padStart(2, "0"),
  ].join("-");
}

function parseIsoDate(dateValue: string) {
  const match = dateValue.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    return null;
  }

  return new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
}

function weekFromDate(dateValue: string) {
  const date = parseIsoDate(dateValue);

  if (!date) {
    return "Sin semana";
  }

  const day = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((date.getTime() - yearStart.getTime()) / 86400000 + 1) / 7);

  return `${date.getUTCFullYear()}-S${String(week).padStart(2, "0")}`;
}

function monthFromDate(dateValue: string) {
  return dateValue.match(/^\d{4}-\d{2}/)?.[0] ?? "Sin mes";
}

function parseDateParts(dateValue: string, hourValue: string) {
  const merged = [dateValue, hourValue].filter(Boolean).join(" ");
  const literalDateTime = merged.match(
    /(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s]+(\d{1,2})(?::(\d{2}))?)?/,
  );

  if (literalDateTime) {
    return {
      date: [
        literalDateTime[1],
        String(Number(literalDateTime[2])).padStart(2, "0"),
        String(Number(literalDateTime[3])).padStart(2, "0"),
      ].join("-"),
      hour: literalDateTime[4] ? `${String(Number(literalDateTime[4])).padStart(2, "0")}:00` : "Sin hora",
    };
  }

  const dayFirstDateTime = merged.match(
    /(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?:[T\s]+(\d{1,2})(?::(\d{2}))?)?/,
  );

  if (dayFirstDateTime) {
    const year = dayFirstDateTime[3].length === 2 ? `20${dayFirstDateTime[3]}` : dayFirstDateTime[3];

    return {
      date: [
        year,
        String(Number(dayFirstDateTime[2])).padStart(2, "0"),
        String(Number(dayFirstDateTime[1])).padStart(2, "0"),
      ].join("-"),
      hour: dayFirstDateTime[4] ? `${String(Number(dayFirstDateTime[4])).padStart(2, "0")}:00` : "Sin hora",
    };
  }

  const parsed = new Date(merged);

  if (!Number.isNaN(parsed.getTime())) {
    return {
      date: formatLocalDate(parsed),
      hour: `${String(parsed.getHours()).padStart(2, "0")}:00`,
    };
  }

  const dateMatch = dateValue.match(/(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})/);
  const hourMatch = hourValue.match(/(\d{1,2}):?(\d{2})?\s*(am|pm)?/i);

  return {
    date: dateMatch?.[0]?.replaceAll("/", "-") ?? "Sin fecha",
    hour: hourMatch ? `${String(Number(hourMatch[1])).padStart(2, "0")}:00` : "Sin hora",
  };
}

function mapRows(rows: RawRow[]) {
  if (!rows.length) {
    return [];
  }

  const headers = Object.keys(rows[0]);
  const columns = {
    date: findColumn(headers, columnAliases.date),
    hour: findColumn(headers, columnAliases.hour),
    phone: findColumn(headers, columnAliases.phone),
    user: findColumn(headers, columnAliases.user),
    campaignId: findColumn(headers, columnAliases.campaignId),
    statusName: findColumn(headers, columnAliases.statusName),
    listName: findColumn(headers, columnAliases.listName),
    direction: findColumn(headers, columnAliases.direction),
    durationSeconds: findColumn(headers, columnAliases.durationSeconds),
    leadId: findColumn(headers, columnAliases.leadId),
    waitSeconds: findColumn(headers, columnAliases.waitSeconds),
  };

  return rows.map((row, index) => {
    const rawDate = valueAsText(columns.date ? row[columns.date] : "");
    const rawHour = valueAsText(columns.hour ? row[columns.hour] : "");
    const { date, hour } = parseDateParts(rawDate, rawHour);
    const week = weekFromDate(date);
    const month = monthFromDate(date);

    return {
      id: index + 1,
      date,
      week,
      month,
      hour,
      phone: valueAsText(columns.phone ? row[columns.phone] : "") || "Sin telefono",
      user: valueAsText(columns.user ? row[columns.user] : "") || "Sin usuario",
      campaignId: valueAsText(columns.campaignId ? row[columns.campaignId] : "") || "Sin campaign_id",
      statusName: valueAsText(columns.statusName ? row[columns.statusName] : "") || "Sin status_name",
      listName: valueAsText(columns.listName ? row[columns.listName] : "") || "Sin list_name",
      direction: valueAsText(columns.direction ? row[columns.direction] : "") || "Sin direction",
      durationSeconds: secondsFromValue(columns.durationSeconds ? row[columns.durationSeconds] : ""),
      leadId: valueAsText(columns.leadId ? row[columns.leadId] : ""),
      waitSeconds: columns.waitSeconds ? secondsFromValue(row[columns.waitSeconds]) : null,
      raw: headers.reduce<Record<string, string>>((record, header) => {
        record[header] = valueAsText(row[header]);
        return record;
      }, {}),
    };
  });
}

function parseDelimitedText(text: string) {
  const cleanText = text.replace(/^\uFEFF/, "").trim();
  const firstLine = cleanText.split(/\r?\n/)[0] ?? "";
  const delimiter = [",", ";", "\t", "|"].sort(
    (a, b) => firstLine.split(b).length - firstLine.split(a).length,
  )[0];

  const rows: string[][] = [];
  let current = "";
  let row: string[] = [];
  let quoted = false;

  for (let index = 0; index < cleanText.length; index += 1) {
    const char = cleanText[index];
    const nextChar = cleanText[index + 1];

    if (char === '"' && nextChar === '"') {
      current += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && nextChar === "\n") {
        index += 1;
      }
      row.push(current.trim());
      rows.push(row);
      row = [];
      current = "";
    } else {
      current += char;
    }
  }

  if (current || row.length) {
    row.push(current.trim());
    rows.push(row);
  }

  const headers = rows.shift()?.map((header, index) => header || `columna_${index + 1}`) ?? [];
  return rows
    .filter((line) => line.some(Boolean))
    .map((line) =>
      headers.reduce<RawRow>((record, header, index) => {
        record[header] = line[index] ?? "";
        return record;
      }, {}),
    );
}

function rowsFromSpreadsheet(sheetRows: Row[]) {
  const headers = (sheetRows.shift() ?? []).map((header, index) =>
    valueAsText(header) || `columna_${index + 1}`,
  );

  return sheetRows
    .filter((line) => line.some((cell) => valueAsText(cell)))
    .map((line) =>
      headers.reduce<RawRow>((record, header, index) => {
        record[header] = line[index] ?? "";
        return record;
      }, {}),
    );
}

function rowsFromSheetRows(sheetRows: Row[]) {
  const headers = (sheetRows[0] ?? []).map((header, index) =>
    valueAsText(header) || `columna_${index + 1}`,
  );

  return sheetRows
    .slice(1)
    .filter((line) => line.some((cell) => valueAsText(cell)))
    .map((line) =>
      headers.reduce<RawRow>((record, header, index) => {
        record[header] = line[index] ?? "";
        return record;
      }, {}),
    );
}

function dateForGrouping(value: string) {
  const date = parseDateTimeValue(value);
  return date ? formatLocalDate(date) : "Sin fecha";
}

function monthForGrouping(value: string) {
  const date = dateForGrouping(value);
  return date === "Sin fecha" ? "Sin mes" : date.slice(0, 7);
}

function parseHtmlTableRows(text: string) {
  if (typeof DOMParser === "undefined" || !/<table[\s>]/i.test(text)) {
    return [];
  }

  const document = new DOMParser().parseFromString(text, "text/html");
  const table = document.querySelector("table");

  if (!table) {
    return [];
  }

  const lines = Array.from(table.querySelectorAll("tr")).map((tr) =>
    Array.from(tr.querySelectorAll("th,td")).map((cell) => cell.textContent?.trim() ?? ""),
  );
  const headers = lines.shift()?.map((header, index) => header || `columna_${index + 1}`) ?? [];

  return lines
    .filter((line) => line.some(Boolean))
    .map((line) =>
      headers.reduce<RawRow>((record, header, index) => {
        record[header] = line[index] ?? "";
        return record;
      }, {}),
    );
}

function parseDateTimeValue(value: string) {
  const text = valueAsText(value);
  const match = text.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})[T\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?/);

  if (match) {
    const date = new Date(
      Number(match[1]),
      Number(match[2]) - 1,
      Number(match[3]),
      Number(match[4]),
      Number(match[5]),
      Number(match[6] ?? 0),
    );

    return Number.isNaN(date.getTime()) ? null : date;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function setTime(date: Date, hours: number, minutes: number) {
  const nextDate = new Date(date);
  nextDate.setHours(hours, minutes, 0, 0);
  return nextDate;
}

function slaMinutesBetween(start: string, end: string) {
  const startDate = parseDateTimeValue(start);
  const endDate = parseDateTimeValue(end);

  if (!startDate || !endDate || endDate < startDate) {
    return null;
  }

  let totalMinutes = 0;
  const cursor = new Date(startDate);
  cursor.setHours(0, 0, 0, 0);

  const lastDay = new Date(endDate);
  lastDay.setHours(0, 0, 0, 0);

  while (cursor <= lastDay) {
    const activeStart = setTime(cursor, 7, 0);
    const activeEnd = setTime(cursor, 23, 1);
    const rangeStart = startDate > activeStart ? startDate : activeStart;
    const rangeEnd = endDate < activeEnd ? endDate : activeEnd;

    if (rangeEnd > rangeStart) {
      totalMinutes += Math.round((rangeEnd.getTime() - rangeStart.getTime()) / 60000);
    }

    cursor.setDate(cursor.getDate() + 1);
  }

  return totalMinutes;
}

function signedMinutesBetween(start: string, end: string) {
  const startDate = parseDateTimeValue(start);
  const endDate = parseDateTimeValue(end);

  if (!startDate || !endDate) return null;
  return Math.round((endDate.getTime() - startDate.getTime()) / 60000);
}

function averageMinutes(values: Array<number | null>) {
  const validValues = values.filter((value): value is number => typeof value === "number");

  if (!validValues.length) {
    return null;
  }

  return Math.round(validValues.reduce((sum, value) => sum + value, 0) / validValues.length);
}

function formatMinutes(value: number | null) {
  if (value === null) {
    return "Sin data";
  }

  const hours = Math.floor(value / 60);
  const minutes = value % 60;

  if (!hours) {
    return `${minutes} min`;
  }

  return `${hours} h ${String(minutes).padStart(2, "0")} min`;
}

function percentage(numerator: number, denominator: number) {
  return denominator > 0 ? (numerator / denominator) * 100 : 0;
}

function formatPercent(value: number) {
  return `${value.toFixed(value >= 10 ? 1 : 2)}%`;
}

function formatEmailTmo(value: number | null) {
  return value === null ? "N/D" : `${value.toFixed(2)} min`;
}

function formatCallDuration(seconds: number) {
  const safeSeconds = Math.max(0, Math.round(seconds));
  return `${String(Math.floor(safeSeconds / 60)).padStart(2, "0")}:${String(safeSeconds % 60).padStart(2, "0")}`;
}

function isTauriRuntime() {
  return typeof window !== "undefined" && "__TAURI_INTERNALS__" in window;
}

function base64ToUint8Array(value: string) {
  const binary = window.atob(value);
  const bytes = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
}

function mapMatrixRows(rows: RawRow[]) {
  if (!rows.length) {
    return [];
  }

  const headers = Object.keys(rows[0]);
  const columns = {
    subjectEmail: findColumn(headers, matrixColumnAliases.subjectEmail),
    dateEmail: findColumn(headers, matrixColumnAliases.dateEmail),
    fromEmail: findColumn(headers, matrixColumnAliases.fromEmail),
    tipificacion: findColumn(headers, matrixColumnAliases.tipificacion),
    motivo: findColumn(headers, matrixColumnAliases.motivo),
    fechaAsignacion: findColumn(headers, matrixColumnAliases.fechaAsignacion),
    usuarioAsignado: findColumn(headers, matrixColumnAliases.usuarioAsignado),
    estadoRegistro: findColumn(headers, matrixColumnAliases.estadoRegistro),
    fechaRegistro: findColumn(headers, matrixColumnAliases.fechaRegistro),
    nroMov: findColumn(headers, matrixColumnAliases.nroMov),
    idEmail: findColumn(headers, matrixColumnAliases.idEmail),
  };

  return rows.map((row, index) => {
    const dateEmail = valueAsText(columns.dateEmail ? row[columns.dateEmail] : "");
    const fechaAsignacion = valueAsText(columns.fechaAsignacion ? row[columns.fechaAsignacion] : "");
    const fechaRegistro = valueAsText(columns.fechaRegistro ? row[columns.fechaRegistro] : "");

    return {
      id: index + 1,
      subjectEmail: valueAsText(columns.subjectEmail ? row[columns.subjectEmail] : "") || "Sin asunto",
      dateEmail,
      fromEmail: valueAsText(columns.fromEmail ? row[columns.fromEmail] : "") || "Sin remitente",
      tipificacion: valueAsText(columns.tipificacion ? row[columns.tipificacion] : "") || "Sin tipificacion",
      motivo: valueAsText(columns.motivo ? row[columns.motivo] : "") || "Sin motivo",
      fechaAsignacion,
      usuarioAsignado: valueAsText(columns.usuarioAsignado ? row[columns.usuarioAsignado] : "") || "Sin usuario",
      estadoRegistro: valueAsText(columns.estadoRegistro ? row[columns.estadoRegistro] : "") || "Sin estado",
      fechaRegistro,
      nroMov: valueAsText(columns.nroMov ? row[columns.nroMov] : ""),
      idEmail: valueAsText(columns.idEmail ? row[columns.idEmail] : ""),
      minutesToAssign: slaMinutesBetween(dateEmail, fechaAsignacion),
      minutesToRegister: slaMinutesBetween(fechaAsignacion, fechaRegistro),
      minutesTotal: slaMinutesBetween(dateEmail, fechaRegistro),
      raw: headers.reduce<Record<string, string>>((record, header) => {
        record[header] = valueAsText(row[header]);
        return record;
      }, {}),
    };
  });
}

function countMatrixBy(rows: MatrixRow[], key: keyof MatrixRow) {
  const map = new Map<string, number>();
  rows.forEach((row) => addToCount(map, String(row[key] || "Sin dato")));
  return pointsFromMap(map);
}

function countMatrixByCaseInsensitive(rows: MatrixRow[], key: keyof MatrixRow) {
  const groups = new Map<string, ChartPoint>();
  rows.forEach((row) => {
    const original = String(row[key] || "Sin dato").trim() || "Sin dato";
    const normalized = normalizeKey(original) || "sin dato";
    const current = groups.get(normalized);
    if (current) {
      current.total += 1;
    } else {
      groups.set(normalized, { name: original, total: 1 });
    }
  });
  return Array.from(groups.values())
    .filter((point) => normalizeKey(point.name) !== "sin dato")
    .sort((a, b) => b.total - a.total);
}

function dateKeyFromDateTime(value: string) {
  const date = parseDateTimeValue(value);
  return date ? formatLocalDate(date) : "Sin fecha";
}

function monthKeyFromDateTime(value: string) {
  const dateKey = dateKeyFromDateTime(value);
  return dateKey === "Sin fecha" ? "Sin mes" : dateKey.slice(0, 7);
}

function formatTime(value: Date) {
  return `${String(value.getHours()).padStart(2, "0")}:${String(value.getMinutes()).padStart(2, "0")}`;
}

function timelineDaysFromRows(rows: MatrixRow[]) {
  const days = new Set<string>();

  rows.forEach((row) => {
    const assigned = parseDateTimeValue(row.fechaAsignacion);
    const registered = parseDateTimeValue(row.fechaRegistro);

    if (!assigned || !registered || registered < assigned) {
      return;
    }

    const cursor = new Date(assigned);
    cursor.setHours(0, 0, 0, 0);
    const lastDay = new Date(registered);
    lastDay.setHours(0, 0, 0, 0);

    while (cursor <= lastDay) {
      days.add(formatLocalDate(cursor));
      cursor.setDate(cursor.getDate() + 1);
    }
  });

  return Array.from(days).sort((a, b) => a.localeCompare(b));
}

function buildAgentTimeline(rows: MatrixRow[], day: string): AgentTimeline[] {
  const [year, month, date] = day.split("-").map(Number);
  const dayStart = new Date(year, month - 1, date);
  const activeStart = setTime(dayStart, 7, 0);
  const activeEnd = setTime(dayStart, 23, 1);
  const groups = new Map<string, TimelineBlock[]>();

  rows.forEach((row) => {
    const assigned = parseDateTimeValue(row.fechaAsignacion);
    const registered = parseDateTimeValue(row.fechaRegistro);

    if (!assigned || !registered || registered <= assigned) {
      return;
    }

    const start = assigned > activeStart ? assigned : activeStart;
    const end = registered < activeEnd ? registered : activeEnd;

    if (end <= start) {
      return;
    }

    const agent = row.usuarioAsignado || "Sin usuario";
    const block: TimelineBlock = {
      id: `${row.id}-${day}`,
      agent,
      left: ((start.getTime() - activeStart.getTime()) / 60000 / slaWindowMinutes) * 100,
      width: Math.max(0.8, ((end.getTime() - start.getTime()) / 60000 / slaWindowMinutes) * 100),
      subject: row.subjectEmail,
      range: `${formatTime(start)} - ${formatTime(end)}`,
      detail: `${row.subjectEmail} | ${row.tipificacion} | ${formatMinutes(row.minutesToRegister)} | ID ${row.idEmail || row.id}`,
      tipificacion: row.tipificacion,
      estado: row.estadoRegistro,
      duration: row.minutesToRegister,
      idEmail: row.idEmail || String(row.id),
      color: colorForLabel(agent),
    };
    const agentBlocks = groups.get(agent);

    if (agentBlocks) {
      agentBlocks.push(block);
    } else {
      groups.set(agent, [block]);
    }
  });

  return Array.from(groups.entries())
    .map(([agent, blocks]) => ({
      agent,
      blocks: blocks.sort((a, b) => a.left - b.left),
    }))
    .sort((a, b) => b.blocks.length - a.blocks.length);
}

function timelineMarks(scale: TimelineScale) {
  const step = scale === "30m" ? 30 : scale === "1h" ? 60 : 120;
  const marks: string[] = [];

  for (let minute = slaStartMinutes; minute <= slaEndMinutes; minute += step) {
    const hour = Math.floor(minute / 60);
    const minutes = minute % 60;
    marks.push(`${String(hour).padStart(2, "0")}:${String(minutes).padStart(2, "0")}`);
  }

  if (marks.at(-1) !== "23:01") {
    marks.push("23:01");
  }

  return marks;
}

function averageOrZero(total: number, count: number) {
  return count ? Math.round(total / count) : 0;
}

function averageMatrixTimes(rows: MatrixRow[], groupBy: (row: MatrixRow) => string, sortBy: "name" | "total" = "total") {
  const groups = new Map<
    string,
    {
      assignTotal: number;
      assignCount: number;
      resolveTotal: number;
      resolveCount: number;
      totalTime: number;
      totalCount: number;
      registros: number;
    }
  >();

  rows.forEach((row) => {
    const key = groupBy(row) || "Sin dato";
    const group =
      groups.get(key) ??
      {
        assignTotal: 0,
        assignCount: 0,
        resolveTotal: 0,
        resolveCount: 0,
        totalTime: 0,
        totalCount: 0,
        registros: 0,
      };

    group.registros += 1;

    if (row.minutesToAssign !== null) {
      group.assignTotal += row.minutesToAssign;
      group.assignCount += 1;
    }

    if (row.minutesToRegister !== null) {
      group.resolveTotal += row.minutesToRegister;
      group.resolveCount += 1;
    }

    if (row.minutesTotal !== null) {
      group.totalTime += row.minutesTotal;
      group.totalCount += 1;
    }

    groups.set(key, group);
  });

  const points = Array.from(groups.entries())
    .map(([name, group]) => ({
      name,
      asignacion: averageOrZero(group.assignTotal, group.assignCount),
      resolucion: averageOrZero(group.resolveTotal, group.resolveCount),
      total: averageOrZero(group.totalTime, group.totalCount),
      registros: group.registros,
    }))
    .filter((point) => point.name !== "Sin dato" && point.name !== "Sin fecha" && point.name !== "Sin mes");

  return sortBy === "name"
    ? points.sort((a, b) => a.name.localeCompare(b.name))
    : points.sort((a, b) => b.total - a.total);
}

function buildMatrixSummary(rows: MatrixRow[]): MatrixSummary {
  return {
    total: rows.length,
    avgToAssign: averageMinutes(rows.map((row) => row.minutesToAssign)),
    avgToRegister: averageMinutes(rows.map((row) => row.minutesToRegister)),
    avgTotal: averageMinutes(rows.map((row) => row.minutesTotal)),
    completed: rows.filter((row) => normalizeKey(row.estadoRegistro).includes("terminado")).length,
    byTipificacion: countMatrixBy(rows, "tipificacion"),
    byEstado: countMatrixBy(rows, "estadoRegistro"),
    byUser: countMatrixByCaseInsensitive(rows, "usuarioAsignado"),
    avgByAgent: averageMatrixTimes(rows, (row) => row.usuarioAsignado),
    avgByDay: averageMatrixTimes(rows, (row) => dateKeyFromDateTime(row.dateEmail), "name"),
    avgByMonth: averageMatrixTimes(rows, (row) => monthKeyFromDateTime(row.dateEmail), "name"),
  };
}

function mapIncidentRows(rows: RawRow[]) {
  if (!rows.length) {
    return [];
  }

  const headers = Object.keys(rows[0]);
  const columns = {
    startTime: findColumn(headers, incidentColumnAliases.startTime),
    endTime: findColumn(headers, incidentColumnAliases.endTime),
    email: findColumn(headers, incidentColumnAliases.email),
    agent: findColumn(headers, incidentColumnAliases.agent),
    branch: findColumn(headers, incidentColumnAliases.branch),
    incidentDate: findColumn(headers, incidentColumnAliases.incidentDate),
    customerCode: findColumn(headers, incidentColumnAliases.customerCode),
    customerName: findColumn(headers, incidentColumnAliases.customerName),
    inconsistencyType: findColumn(headers, incidentColumnAliases.inconsistencyType),
    channel: findColumn(headers, incidentColumnAliases.channel),
    subject: findColumn(headers, incidentColumnAliases.subject),
    detail: findColumn(headers, incidentColumnAliases.detail),
    impact: findColumn(headers, incidentColumnAliases.impact),
  };

  return rows.map((row, index) => ({
    id: Number(valueAsText(row.Id ?? row.id)) || index + 1,
    startTime: valueAsText(columns.startTime ? row[columns.startTime] : ""),
    endTime: valueAsText(columns.endTime ? row[columns.endTime] : ""),
    email: valueAsText(columns.email ? row[columns.email] : "") || "Sin correo",
    agent: valueAsText(columns.agent ? row[columns.agent] : "") || "Sin agente",
    branch: valueAsText(columns.branch ? row[columns.branch] : "") || "Sin sucursal",
    incidentDate: valueAsText(columns.incidentDate ? row[columns.incidentDate] : "") || "Sin fecha",
    customerCode: valueAsText(columns.customerCode ? row[columns.customerCode] : ""),
    customerName: valueAsText(columns.customerName ? row[columns.customerName] : "") || "Sin cliente",
    inconsistencyType: valueAsText(columns.inconsistencyType ? row[columns.inconsistencyType] : "") || "Sin tipo",
    channel: valueAsText(columns.channel ? row[columns.channel] : "") || "Sin canal",
    subject: valueAsText(columns.subject ? row[columns.subject] : "") || "Sin asunto",
    detail: valueAsText(columns.detail ? row[columns.detail] : "") || "Sin detalle",
    impact: valueAsText(columns.impact ? row[columns.impact] : "") || "Sin impacto",
    raw: headers.reduce<Record<string, string>>((record, header) => {
      record[header] = valueAsText(row[header]);
      return record;
    }, {}),
  }));
}

function countIncidentsBy(rows: IncidentRow[], key: keyof IncidentRow, sortBy: "name" | "total" = "total") {
  const map = new Map<string, number>();
  rows.forEach((row) => addToCount(map, String(row[key] || "Sin dato")));
  return pointsFromMap(map, sortBy);
}

function buildIncidentSummary(rows: IncidentRow[]): IncidentSummary {
  const monthMap = new Map<string, number>();
  rows.forEach((row) => addToCount(monthMap, monthForGrouping(row.incidentDate)));

  return {
    total: rows.length,
    uniqueAgents: new Set(rows.map((row) => row.agent).filter(Boolean)).size,
    uniqueTypes: new Set(rows.map((row) => row.inconsistencyType).filter(Boolean)).size,
    uniqueBranches: new Set(rows.map((row) => row.branch).filter(Boolean)).size,
    byAgent: countIncidentsBy(rows, "agent"),
    byType: countIncidentsBy(rows, "inconsistencyType"),
    byBranch: countIncidentsBy(rows, "branch"),
    byMonth: pointsFromMap(monthMap, "name"),
  };
}

function incrementAgent(map: Map<string, number>, agent: string, amount = 1) {
  const key = agent || "Sin agente";
  map.set(key, (map.get(key) ?? 0) + amount);
}

function buildPerformanceSummary(callRows: MetricRow[], emailRows: MatrixRow[], errorRows: IncidentRow[]): PerformanceSummary {
  const callCounts = new Map<string, number>();
  const emailCounts = new Map<string, number>();
  const errorCounts = new Map<string, number>();

  callRows.forEach((row) => incrementAgent(callCounts, row.user));
  emailRows.forEach((row) => incrementAgent(emailCounts, row.usuarioAsignado));
  errorRows.forEach((row) => incrementAgent(errorCounts, row.agent));

  const totalCalls = callRows.length;
  const totalEmails = emailRows.length;
  const totalAttended = totalCalls + totalEmails;
  const totalErrors = errorRows.length;
  const agents = new Set([...callCounts.keys(), ...emailCounts.keys(), ...errorCounts.keys()]);
  const byAgent = Array.from(agents)
    .map((agent) => {
      const calls = callCounts.get(agent) ?? 0;
      const emails = emailCounts.get(agent) ?? 0;
      const errors = errorCounts.get(agent) ?? 0;
      const attended = calls + emails;
      const productivity = percentage(attended, totalAttended);
      const quality = percentage(errors, emails);

      return {
        agent,
        calls,
        emails,
        errors,
        attended,
        productivity,
        quality,
        effectiveness: (productivity * quality) / 100,
      };
    })
    .sort((a, b) => b.attended - a.attended || b.errors - a.errors || a.agent.localeCompare(b.agent));

  const productivity = percentage(totalAttended, totalAttended);
  const quality = percentage(totalErrors, totalEmails);

  return {
    totalCalls,
    totalEmails,
    totalAttended,
    totalErrors,
    productivity,
    quality,
    effectiveness: (productivity * quality) / 100,
    byAgent,
  };
}

function addToCount(map: Map<string, number>, value: string) {
  const label = value || "Sin dato";
  map.set(label, (map.get(label) ?? 0) + 1);
}

function pointsFromMap(map: Map<string, number>, sortBy: "name" | "total" = "total") {
  const points = Array.from(map.entries())
    .map(([name, total]) => ({ name, total }))
    .filter((point) => point.name !== "Sin dato");

  return sortBy === "name"
    ? points.sort((a, b) => a.name.localeCompare(b.name))
    : points.sort((a, b) => b.total - a.total);
}

function hourPointsFromMap(map: Map<string, number>) {
  const hours = Array.from(map.keys())
    .map((hour) => Number(hour.slice(0, 2)))
    .filter((hour) => Number.isInteger(hour) && hour >= 0 && hour <= 23);

  if (!hours.length) {
    return [];
  }

  const firstHour = Math.min(...hours);
  const lastHour = Math.max(...hours);

  return Array.from({ length: lastHour - firstHour + 1 }, (_, offset) => {
    const index = firstHour + offset;
    const name = `${String(index).padStart(2, "0")}:00`;

    return {
      name,
      total: map.get(name) ?? 0,
    };
  });
}

function matrixEmailKey(row: MatrixRow) {
  return row.idEmail.trim() || `row-${row.id}`;
}

function isEmailAttended(row: MatrixRow) {
  const status = normalizeKey(row.estadoRegistro);
  return ["terminado", "completado", "atendido", "resuelto", "cerrado"].some((value) => status.includes(value));
}

function buildEmailSummary(rows: MatrixRow[]): EmailSummary {
  const unique = new Map<string, MatrixRow>();
  rows.forEach((row) => {
    const key = matrixEmailKey(row);
    const previous = unique.get(key);
    if (!previous || (!previous.fechaRegistro && row.fechaRegistro)) unique.set(key, row);
  });
  const calls = Array.from(unique.values());
  const attended = calls.filter(isEmailAttended);
  const withinSla = calls.filter((row) => {
    const registrationMinutes = signedMinutesBetween(row.dateEmail, row.fechaRegistro);
    return registrationMinutes !== null && registrationMinutes <= EMAIL_SLA_MINUTES;
  });
  const validTmo = attended.map((row) => row.minutesTotal).filter((value): value is number => value !== null && value >= 0);
  const monthMap = new Map<string, MatrixRow[]>();
  calls.forEach((row) => {
    const month = monthKeyFromDateTime(row.dateEmail);
    monthMap.set(month, [...(monthMap.get(month) ?? []), row]);
  });
  const byMonth = Array.from(monthMap.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([name, monthRows]) => {
    const monthAttended = monthRows.filter(isEmailAttended);
    const monthSla = monthRows.filter((row) => {
      const registrationMinutes = signedMinutesBetween(row.dateEmail, row.fechaRegistro);
      return registrationMinutes !== null && registrationMinutes <= EMAIL_SLA_MINUTES;
    });
    const monthTmo = monthAttended.map((row) => row.minutesTotal).filter((value): value is number => value !== null && value >= 0);
    return { name, received: monthRows.length, attended: monthAttended.length, withinSla: monthSla.length, nda: percentage(monthAttended.length, monthRows.length), nds: percentage(monthSla.length, monthRows.length), tmoMinutes: monthTmo.length ? monthTmo.reduce((sum, value) => sum + value, 0) / monthTmo.length : null };
  });
  const points = (key: keyof MatrixRow) => key === "usuarioAsignado" ? countMatrixByCaseInsensitive(calls, key) : countMatrixBy(calls, key);
  const hours = new Map<string, number>();
  calls.forEach((row) => { const date = parseDateTimeValue(row.dateEmail); const hour = date ? `${String(date.getHours()).padStart(2, "0")}:00` : "Sin hora"; hours.set(hour, (hours.get(hour) ?? 0) + 1); });
  return { received: calls.length, attended: attended.length, withinSla: withinSla.length, nda: percentage(attended.length, calls.length), nds: percentage(withinSla.length, calls.length), tmoMinutes: validTmo.length ? validTmo.reduce((sum, value) => sum + value, 0) / validTmo.length : null, byMonth, byUser: points("usuarioAsignado"), byTipificacion: points("tipificacion"), byMotivo: points("motivo").filter((point) => point.name !== "Sin motivo"), byHour: pointsFromMap(hours) , period: byMonth.length ? `${byMonth[0].name} a ${byMonth[byMonth.length - 1].name}` : "Sin periodo" };
}

function isInboundCall(row: MetricRow) {
  return row.direction === "Sin direction" || normalizeKey(row.direction) === "inbound";
}

function isAttendedCall(row: MetricRow) {
  return row.durationSeconds > 0;
}

function matchesAttendedDefinition(row: MetricRow, statusNames: string[]) {
  return statusNames.length ? statusNames.includes(row.statusName) : isAttendedCall(row);
}

function countAttendedCalls(rows: MetricRow[], statusNames: string[] = []) {
  return rows.filter((row) => matchesAttendedDefinition(row, statusNames)).length;
}

function buildMonthlyCallPoints(rows: MetricRow[], statusNames: string[], focusedStatusName?: string | null): MonthlyCallPoint[] {
  const months = new Map<string, MonthlyCallPoint>();
  rows.forEach((row) => {
    if (!months.has(row.month)) months.set(row.month, { name: row.month, received: 0, attended: 0 });
  });

  rows.forEach((row) => {
    if (focusedStatusName && row.statusName !== focusedStatusName) return;
    const point = months.get(row.month);
    if (!point) return;
    point.received += 1;

    if (matchesAttendedDefinition(row, statusNames)) {
      point.attended += 1;
    }

  });

  return Array.from(months.values()).sort((a, b) => a.name.localeCompare(b.name));
}

function buildServiceSummary(rows: MetricRow[], statusNames: string[] = []): ServiceSummary {
  const months = new Map<string, { received: number; attended: number; attendedUnder20: number; durationSeconds: number; waitAvailable: boolean }>();
  let attended = 0;
  let attendedUnder20 = 0;
  let durationSeconds = 0;

  rows.forEach((row) => {
    const month = months.get(row.month) ?? { received: 0, attended: 0, attendedUnder20: 0, durationSeconds: 0, waitAvailable: false };
    month.received += 1;
    month.waitAvailable ||= row.waitSeconds !== null;

    if (matchesAttendedDefinition(row, statusNames)) {
      month.attended += 1;
      month.durationSeconds += row.durationSeconds;
      attended += 1;
      durationSeconds += row.durationSeconds;

      if (row.waitSeconds !== null && row.waitSeconds <= 20) {
        month.attendedUnder20 += 1;
        attendedUnder20 += 1;
      }
    }

    months.set(row.month, month);
  });

  const byMonth = Array.from(months.entries())
    .map(([name, month]) => {
      const abandoned = month.received - month.attended;
      return {
        name,
        received: month.received,
        attended: month.attended,
        attendedUnder20: month.attendedUnder20,
        abandoned,
        attentionRate: percentage(month.attended, month.received),
        serviceRate: month.waitAvailable ? percentage(month.attendedUnder20, month.received) : null,
        abandonmentRate: percentage(abandoned, month.received),
        averageDurationSeconds: month.attended ? month.durationSeconds / month.attended : 0,
      };
    })
    .sort((a, b) => a.name.localeCompare(b.name));
  const received = rows.length;
  const abandoned = received - attended;

  return {
    attentionRate: percentage(attended, received),
    serviceRate: rows.some((row) => row.waitSeconds !== null) ? percentage(attendedUnder20, received) : null,
    abandonmentRate: percentage(abandoned, received),
    averageDurationSeconds: attended ? durationSeconds / attended : 0,
    byMonth,
  };
}

function buildDashboardSummary(rows: MetricRow[]): DashboardSummary {
  const counts = {
    date: new Map<string, number>(),
    week: new Map<string, number>(),
    month: new Map<string, number>(),
    hour: new Map<string, number>(),
    statusName: new Map<string, number>(),
    campaignId: new Map<string, number>(),
    user: new Map<string, number>(),
    listName: new Map<string, number>(),
  } satisfies Record<FilterField, Map<string, number>>;
  const unique = {
    phone: new Set<string>(),
    user: new Set<string>(),
    campaignId: new Set<string>(),
    statusName: new Set<string>(),
  } satisfies Record<"phone" | "user" | "campaignId" | "statusName", Set<string>>;

  rows.forEach((row) => {
    filterFields.forEach((field) => addToCount(counts[field], row[field]));
    unique.phone.add(row.phone);
    unique.user.add(row.user);
    unique.campaignId.add(row.campaignId);
    unique.statusName.add(row.statusName);
  });

  return {
    total: rows.length,
    byDate: pointsFromMap(counts.date, "name"),
    byWeek: pointsFromMap(counts.week, "name"),
    byMonth: pointsFromMap(counts.month, "name"),
    byHour: hourPointsFromMap(counts.hour),
    byStatus: pointsFromMap(counts.statusName),
    byListName: pointsFromMap(counts.listName),
    byCampaign: pointsFromMap(counts.campaignId),
    byUser: pointsFromMap(counts.user),
    uniquePhones: unique.phone.size,
    uniqueUsers: unique.user.size,
    uniqueCampaigns: unique.campaignId.size,
    uniqueStatuses: unique.statusName.size,
    attendedCalls: countAttendedCalls(rows),
  };
}

function StatCard({
  icon,
  label,
  value,
  comparison,
}: {
  icon: ReactNode;
  label: string;
  value: string | number;
  comparison?: ReactNode;
}) {
  return (
    <div className={styles.statCard}>
      <div className={styles.statIcon}>{icon}</div>
      <span>{label}</span>
      <strong>{value}</strong>
      {comparison ? <small className={styles.statComparison}>{comparison}</small> : null}
    </div>
  );
}

function EmptyChart({ children }: { children: ReactNode }) {
  return <div className={styles.emptyChart}>{children}</div>;
}

function chartClickValue(state: ChartClickState) {
  const payloadName = state.activePayload?.[0]?.payload?.name;

  if (payloadName) {
    return payloadName;
  }

  if (state.activeLabel !== undefined && state.activeLabel !== null) {
    return String(state.activeLabel);
  }

  const fallback = state.activePayload?.[0]?.name ?? state.activePayload?.[0]?.value;
  return fallback === undefined || fallback === null ? undefined : String(fallback);
}

function chartPointName(entry: unknown) {
  const point = entry as Partial<ChartPoint>;
  return point.name ? String(point.name) : undefined;
}

function chartColor(index: number) {
  return chartColors[index % chartColors.length];
}

function colorForLabel(label: string) {
  const total = Array.from(label).reduce((sum, char) => sum + char.charCodeAt(0), 0);
  return chartColor(total);
}

function stopInsideClick(event: ReactMouseEvent<HTMLElement>) {
  event.stopPropagation();
}

function matrixUserNameFromChartEntry(entry: unknown) {
  const point = entry as { name?: unknown; payload?: Partial<ChartPoint> };
  return point.payload?.name ? String(point.payload.name) : point.name ? String(point.name) : undefined;
}

function rowsForDashboardFocus(rows: MetricRow[], focus: Partial<Record<FilterField, string>>, excludedField?: FilterField) {
  return rows.filter((row) => Object.entries(focus).every(([field, value]) => field === excludedField || !value || row[field as FilterField] === value));
}

function buildDurationDistribution(rows: MetricRow[]): ChartPoint[] {
  const buckets = [
    { name: "0 s", test: (seconds: number) => seconds === 0 },
    { name: "1–20 s", test: (seconds: number) => seconds > 0 && seconds <= 20 },
    { name: "21–60 s", test: (seconds: number) => seconds > 20 && seconds <= 60 },
    { name: "61–180 s", test: (seconds: number) => seconds > 60 && seconds <= 180 },
    { name: "181–300 s", test: (seconds: number) => seconds > 180 && seconds <= 300 },
    { name: ">300 s", test: (seconds: number) => seconds > 300 },
  ];
  return buckets.map((bucket) => ({ name: bucket.name, total: rows.filter((row) => bucket.test(row.durationSeconds)).length }));
}

function buildUserPerformance(rows: MetricRow[]): UserPerformancePoint[] {
  const groups = new Map<string, { received: number; attended: number }>();
  rows.forEach((row) => {
    const group = groups.get(row.user) ?? { received: 0, attended: 0 };
    group.received += 1;
    if (isAttendedCall(row)) group.attended += 1;
    groups.set(row.user, group);
  });
  return Array.from(groups.entries()).map(([name, group]) => ({ ...group, name, attentionRate: percentage(group.attended, group.received) })).sort((a, b) => b.received - a.received);
}

const chartInfoByTitle: Record<string, string> = {
  "Llamadas por mes": "Campos: call_date (para obtener el mes), status_name (filtro opcional) y length_in_sec (atendida si > 0). Fórmula: contar registros recibidos y atendidos por mes; campaign_id LINDESAC se excluye antes del conteo.",
  "Llamadas recibidas por hora": "Campos: call_date y hour (hora normalizada). Fórmula: contar registros agrupados por hora; no usa lead_id.",
  "Estados de llamada": "Campo: status_name. Fórmula: contar registros por estado. Al hacer clic en un estado, sólo se enfoca visualmente este dashboard; las demás categorías permanecen visibles y atenuadas.",
  "Campañas": "Campo: campaign_id. Fórmula: contar registros agrupados por campaña; se excluye LINDESAC.",
  "Mapa de calor I1": "Campos: call_date y hora. Cuenta llamadas por día de semana y hora, mostrando sólo 07:00–23:00.",
  "Calidad del servicio": "Campos: lead_id, call_date y tiempo de espera válido. NA = atendidas / recibidas. NS = atendidas con espera ≤20 segundos; si no existe espera válida, muestra N/D.",
  "TMO por mes": "Campo: length_in_sec. TMO = promedio de duración de llamadas atendidas, agrupado por mes.",
  "Abandono por mes": "Campos: lead_id y call_date. Abandonadas = recibidas únicas − atendidas únicas.",
};

function ChartPanel({
  title,
  meta,
  children,
  onExpand,
  onExport,
  actions,
  info,
  comparison,
}: {
  title: string;
  meta: string;
  children: ReactNode;
  onExpand?: () => void;
  onExport?: () => void;
  actions?: ReactNode;
  info?: string;
  comparison?: ReactNode;
}) {
  const panelInfo = info ?? chartInfoByTitle[title] ?? `Dashboard ${title}. Los datos se calculan dinámicamente desde el archivo cargado.`;

  return (
    <article className={styles.panel} onClick={stopInsideClick}>
      <div className={styles.panelHeader}>
        <div>
          <h2>{title}</h2>
          <span>{meta}</span>
          {comparison}
        </div>
        {actions || onExport || onExpand || info ? (
          <div className={styles.headerActions}>
            {actions}
            <details className={styles.chartInfo}><summary aria-label={`Ver fórmula de ${title}`} title="Ver fórmula y campos"><Info size={17} aria-hidden="true" /></summary><p>{panelInfo}</p></details>
            {onExport ? (
              <button className={styles.iconButton} type="button" onClick={onExport} aria-label={`Exportar ${title} como imagen`} title={`Exportar ${title} como imagen`}>
                <Download size={17} aria-hidden="true" />
              </button>
            ) : null}
            {onExpand ? (
              <button className={styles.iconButton} type="button" onClick={onExpand} aria-label={`Ampliar ${title}`} title={`Ampliar ${title}`}>
                <Maximize2 size={17} aria-hidden="true" />
              </button>
            ) : null}
          </div>
        ) : null}
      </div>
      {children}
    </article>
  );
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const incidentInputRef = useRef<HTMLInputElement>(null);
  const matrixInputRef = useRef<HTMLInputElement>(null);
  const dateChartRef = useRef<HTMLDivElement>(null);
  const hourChartRef = useRef<HTMLDivElement>(null);
  const statusChartRef = useRef<HTMLDivElement>(null);
  const campaignChartRef = useRef<HTMLDivElement>(null);
  const [rows, setRows] = useState<MetricRow[]>([]);
  const [fileName, setFileName] = useState("Sin archivo cargado");
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [dashboardFocus, setDashboardFocus] = useState<Partial<Record<FilterField, string>>>({});
  const [selectedStatusNames, setSelectedStatusNames] = useState<string[]>([]);
  const [statusFocus, setStatusFocus] = useState<string | null>(null);
  const [expandedChart, setExpandedChart] = useState<ChartKind | null>(null);
  const [callStartDate, setCallStartDate] = useState("");
  const [callEndDate, setCallEndDate] = useState("");
  const [activeView, setActiveView] = useState<DashboardView>("calls");
  const [reportStartDate, setReportStartDate] = useState(todayInputValue);
  const [reportEndDate, setReportEndDate] = useState(todayInputValue);
  const [reportSessionId, setReportSessionId] = useState("");
  const [reportError, setReportError] = useState("");
  const [isDownloadingReport, setIsDownloadingReport] = useState(false);
  const [matrixRows, setMatrixRows] = useState<MatrixRow[]>([]);
  const [matrixColumns, setMatrixColumns] = useState<string[]>([]);
  const [matrixRangeLabel, setMatrixRangeLabel] = useState("");
  const [matrixFileName, setMatrixFileName] = useState("Sin archivo local");
  const [extensionReady, setExtensionReady] = useState(false);
  const [timelineDay, setTimelineDay] = useState("");
  const [timelineScale, setTimelineScale] = useState<TimelineScale>("day");
  const [emailHourChronological, setEmailHourChronological] = useState(false);
  const [showSlowResolutionOnly, setShowSlowResolutionOnly] = useState(false);
  const [persistenceReady, setPersistenceReady] = useState(false);
  const [matrixDayDetail, setMatrixDayDetail] = useState<string | null>(null);
  const [matrixUserFocus, setMatrixUserFocus] = useState<string | null>(null);
  const [hoveredTimelineId, setHoveredTimelineId] = useState<string | null>(null);
  const [incidentRows, setIncidentRows] = useState<IncidentRow[]>([]);
  const [incidentColumns, setIncidentColumns] = useState<string[]>([]);
  const [incidentFileName, setIncidentFileName] = useState("Sin archivo cargado");
  const [incidentError, setIncidentError] = useState("");
  const salesInputRef = useRef<HTMLInputElement>(null);
  const [salesRows, setSalesRows] = useState<SalesRow[]>([]);
  const [salesFileName, setSalesFileName] = useState("Sin archivo cargado");
  const [salesError, setSalesError] = useState("");
  const [salesTab, setSalesTab] = useState<SalesTab>("summary");
  const [salesFilters, setSalesFilters] = useState<Record<string, string>>({});
  const [salesCompareRegions, setSalesCompareRegions] = useState(false);
  const [accessChecked, setAccessChecked] = useState(false);
  const [accessGranted, setAccessGranted] = useState(false);
  const [accessPassword, setAccessPassword] = useState("");

  const dateFilteredRows = useMemo(
    () => rows.filter((row) => (!callStartDate && !callEndDate) || (row.date !== "Sin fecha" && (!callStartDate || row.date >= callStartDate) && (!callEndDate || row.date <= callEndDate))),
    [callEndDate, callStartDate, rows],
  );
  const fullSummary = useMemo(() => buildDashboardSummary(rows), [rows]);
  const visibleRows = useMemo(
    () => dateFilteredRows.filter((row) => Object.entries(dashboardFocus).every(([field, value]) => !value || row[field as FilterField] === value)),
    [dashboardFocus, dateFilteredRows],
  );
  const currentSummary = useMemo(
    () => ({
      ...buildDashboardSummary(visibleRows),
      attendedCalls: countAttendedCalls(visibleRows, selectedStatusNames),
    }),
    [selectedStatusNames, visibleRows],
  );
  const monthChartRows = useMemo(() => rowsForDashboardFocus(dateFilteredRows, { ...dashboardFocus, statusName: undefined }, "month"), [dashboardFocus, dateFilteredRows]);
  const hourChartRows = useMemo(() => rowsForDashboardFocus(dateFilteredRows, dashboardFocus, "hour"), [dashboardFocus, dateFilteredRows]);
  const statusChartRows = useMemo(() => rowsForDashboardFocus(dateFilteredRows, dashboardFocus, "statusName"), [dashboardFocus, dateFilteredRows]);
  const campaignChartRows = useMemo(() => rowsForDashboardFocus(dateFilteredRows, dashboardFocus, "campaignId"), [dashboardFocus, dateFilteredRows]);
  const userChartRows = useMemo(() => rowsForDashboardFocus(dateFilteredRows, dashboardFocus, "user"), [dashboardFocus, dateFilteredRows]);
  const campaignChartSummary = useMemo(() => buildDashboardSummary(campaignChartRows), [campaignChartRows]);
  const statusChartSummary = useMemo(() => buildDashboardSummary(statusChartRows), [statusChartRows]);
  const durationDistribution = useMemo(() => buildDurationDistribution(visibleRows), [visibleRows]);
  const userPerformance = useMemo(() => buildUserPerformance(userChartRows), [userChartRows]);
  const monthlyCallData = useMemo(
    () => buildMonthlyCallPoints(monthChartRows, statusFocus ? [] : selectedStatusNames, statusFocus),
    [monthChartRows, selectedStatusNames, statusFocus],
  );
  const serviceSummary = useMemo(() => buildServiceSummary(visibleRows, selectedStatusNames), [selectedStatusNames, visibleRows]);
  const callMonthComparison = (valueFor: (month: MonthlyCallPoint | ServiceMonthlyPoint) => number | null) => {
    const months = serviceSummary.byMonth.length >= 2 ? serviceSummary.byMonth : monthlyCallData;
    if (months.length < 2) return null;
    const previous = valueFor(months[months.length - 2]);
    const current = valueFor(months[months.length - 1]);
    if (previous === null || current === null || previous === 0) return "N/D";
    const change = ((current - previous) / previous) * 100;
    return `${change >= 0 ? "↑" : "↓"} ${Math.abs(change).toFixed(1)}%`;
  };
  const callComparisonBadge = (label: string, valueFor: (month: MonthlyCallPoint | ServiceMonthlyPoint) => number | null) => {
    const value = callMonthComparison(valueFor);
    return value === null ? null : <span className={styles.comparisonBadge}>VS mes anterior · {label}: {value}</span>;
  };
  const byHour = useMemo(() => buildDashboardSummary(hourChartRows).byHour, [hourChartRows]);
  const allStatus = fullSummary.byStatus;
  const byStatus = statusChartSummary.byStatus;
  const statusChartField: FilterField = "statusName";
  const statusChartData = byStatus;
  const allStatusChartData = allStatus;
  const byCampaign = campaignChartSummary.byCampaign;
  const callHeatmap = useMemo(() => buildCallHeatmap(visibleRows), [visibleRows]);
  const slowResolutionRows = useMemo(
    () => matrixRows.filter((row) => row.minutesToRegister !== null && row.minutesToRegister > 20),
    [matrixRows],
  );
  const visibleMatrixRows = showSlowResolutionOnly ? slowResolutionRows : matrixRows;
  const matrixSummary = useMemo(() => buildMatrixSummary(visibleMatrixRows), [visibleMatrixRows]);
  const emailSummary = useMemo(() => buildEmailSummary(visibleMatrixRows), [visibleMatrixRows]);
  const emailMonthComparison = (label: string, valueFor: (month: EmailMonthPoint) => number | null) => {
    const months = emailSummary.byMonth.filter((month) => month.name !== "Sin mes");
    if (months.length < 2) return null;
    const previous = valueFor(months[months.length - 2]);
    const current = valueFor(months[months.length - 1]);
    if (previous === null || current === null || previous === 0) {
      return <span className={styles.comparisonBadge}>VS {months[months.length - 2].name} → {months[months.length - 1].name} · {label}: N/D</span>;
    }
    const change = ((current - previous) / previous) * 100;
    return <span className={`${styles.comparisonBadge} ${change >= 0 ? styles.comparisonUp : styles.comparisonDown}`}>VS {months[months.length - 2].name} → {months[months.length - 1].name} · {label}: {change >= 0 ? "↑" : "↓"} {Math.abs(change).toFixed(1)}%</span>;
  };
  const emailMonthDelta = (valueFor: (month: EmailMonthPoint) => number | null) => {
    const months = emailSummary.byMonth.filter((month) => month.name !== "Sin mes");
    if (months.length < 2) return "N/D";
    const previous = valueFor(months[months.length - 2]);
    const current = valueFor(months[months.length - 1]);
    if (previous === null || current === null || previous === 0) return "N/D";
    const change = ((current - previous) / previous) * 100;
    return `${change >= 0 ? "↑" : "↓"} ${Math.abs(change).toFixed(1)}%`;
  };
  const emailVolumeComparison = emailMonthComparison("volumen", (month) => month.received);
  const emailHourData = useMemo(
    () => emailHourChronological
      ? [...emailSummary.byHour].sort((a, b) => {
          const hourA = Number.parseInt(a.name.slice(0, 2), 10);
          const hourB = Number.parseInt(b.name.slice(0, 2), 10);
          if (!Number.isNaN(hourA) && !Number.isNaN(hourB)) return hourA - hourB;
          if (!Number.isNaN(hourA)) return -1;
          if (!Number.isNaN(hourB)) return 1;
          return a.name.localeCompare(b.name);
        })
      : emailSummary.byHour,
    [emailHourChronological, emailSummary.byHour],
  );
  const emailHeatmap = useMemo(() => buildEmailHeatmap(visibleMatrixRows), [visibleMatrixRows]);
  const filteredMatrixRows = useMemo(
    () => matrixUserFocus ? visibleMatrixRows.filter((row) => normalizeKey(row.usuarioAsignado) === normalizeKey(matrixUserFocus)) : visibleMatrixRows,
    [matrixUserFocus, visibleMatrixRows],
  );
  const matrixDetailRows = useMemo(() => filteredMatrixRows.slice(0, detailRowLimit), [filteredMatrixRows]);
  const hiddenMatrixRows = Math.max(0, filteredMatrixRows.length - matrixDetailRows.length);
  const matrixDayDetailRows = useMemo(
    () => matrixDayDetail ? filteredMatrixRows.filter((row) => dateKeyFromDateTime(row.dateEmail) === matrixDayDetail) : [],
    [filteredMatrixRows, matrixDayDetail],
  );
  const timelineDays = useMemo(() => timelineDaysFromRows(visibleMatrixRows), [visibleMatrixRows]);
  const selectedTimelineDay = timelineDays.includes(timelineDay) ? timelineDay : timelineDays[0] ?? "";
  const timelineRows = useMemo(
    () => (selectedTimelineDay ? buildAgentTimeline(visibleMatrixRows, selectedTimelineDay) : []),
    [visibleMatrixRows, selectedTimelineDay],
  );
  const activeTimelineBlock =
    timelineRows.flatMap((row) => row.blocks).find((block) => block.id === hoveredTimelineId) ??
    timelineRows[0]?.blocks[0];
  const incidentSummary = useMemo(() => buildIncidentSummary(incidentRows), [incidentRows]);
  const incidentDetailRows = useMemo(() => incidentRows.slice(0, detailRowLimit), [incidentRows]);
  const hiddenIncidentRows = Math.max(0, incidentRows.length - incidentDetailRows.length);
  const incidentAgentChartHeight = Math.max(280, Math.min(760, incidentSummary.byAgent.length * 38));
  const performanceSummary = useMemo(() => buildPerformanceSummary(rows, matrixRows, incidentRows), [incidentRows, matrixRows, rows]);
  const performanceChartHeight = Math.max(320, Math.min(760, performanceSummary.byAgent.length * 44));
  const campaignChartHeight = Math.max(280, Math.min(760, byCampaign.length * 38));
  const matrixUserChartHeight = Math.max(280, Math.min(620, matrixSummary.byUser.length * 38));
  const salesFilteredRows = useMemo(() => salesRows.filter((row) => Object.entries(salesFilters).every(([key, value]) => !value || String(row[key as keyof SalesRow] ?? "") === value)), [salesFilters, salesRows]);
  const salesComparisonRows = useMemo(() => salesRows.filter((row) => Object.entries(salesFilters).every(([key, value]) => key === "region" || !value || String(row[key as keyof SalesRow] ?? "") === value)), [salesFilters, salesRows]);
  const salesRegions = useMemo(() => Array.from(new Set(salesComparisonRows.map((row) => row.region).filter(Boolean))).sort((a, b) => a.localeCompare(b)).slice(0, 2), [salesComparisonRows]);
  const salesMonths = useMemo(() => {
    const map = new Map<string, { name: string; total: number; tons: number; records: number; clients: Set<string>; products: Set<string> }>();
    salesFilteredRows.forEach((row) => { const item = map.get(row.date.slice(0, 7)) ?? { name: row.date.slice(0, 7), total: 0, tons: 0, records: 0, clients: new Set<string>(), products: new Set<string>() }; item.total += row.amountUSD; item.tons += row.tons; item.records += 1; item.clients.add(row.clientCode); item.products.add(row.productCode); map.set(item.name, item); });
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name)).map((item) => ({ name: item.name, total: item.total, tons: item.tons, records: item.records, clients: item.clients.size, products: item.products.size }));
  }, [salesFilteredRows]);
  const salesTotalUSD = salesFilteredRows.reduce((sum, row) => sum + row.amountUSD, 0);
  const salesBranches = useMemo(() => salesGrouped(salesFilteredRows, "branch"), [salesFilteredRows]);
  const salesClasses = useMemo(() => salesGrouped(salesFilteredRows, "className").map((point) => ({ ...point, name: salesDisplayName(point.name) })), [salesFilteredRows]);
  const salesProducts = useMemo(() => salesGrouped(salesFilteredRows, "productDescription").slice(0, 10), [salesFilteredRows]);
  const salesClients = useMemo(() => salesGrouped(salesFilteredRows, "clientName").slice(0, 5), [salesFilteredRows]);
  const salesFrequency = useMemo(() => salesPurchaseFrequency(salesFilteredRows), [salesFilteredRows]);
  const salesIndustries = useMemo(() => salesGrouped(salesFilteredRows, "industry").slice(0, 10).map((point) => ({ ...point, name: salesDisplayName(point.name) })), [salesFilteredRows]);
  const statusTotal = statusChartSummary.total;
  const statusColorFor = (name: string) => chartColor(Math.max(0, allStatusChartData.findIndex((item) => item.name === name)));
  const statusFilterLabel = selectedStatusNames.length
    ? `${selectedStatusNames.length} de ${allStatus.length} seleccionados`
    : "Todos los estados";

  const chartRefs: Record<ChartKind, RefObject<HTMLDivElement | null>> = {
    date: dateChartRef,
    hour: hourChartRef,
    status: statusChartRef,
    campaign: campaignChartRef,
  };

  useEffect(() => {
    const hasAccessCookie = document.cookie.split(";").some((cookie) => cookie.trim().startsWith("a365_access=1"));
    setAccessGranted(hasAccessCookie);
    setAccessChecked(true);
  }, []);

  useEffect(() => {
    const saved = readPersistedCalls();
    const savedDashboard = readPersistedDashboard();
    window.setTimeout(() => {
      if (saved) {
        setRows(saved.rows);
        setFileName(saved.fileName);
      }
      if (savedDashboard) {
        if (Array.isArray(savedDashboard.matrixRows) && savedDashboard.matrixRows.length) setMatrixRows(savedDashboard.matrixRows);
        if (Array.isArray(savedDashboard.matrixColumns)) setMatrixColumns(savedDashboard.matrixColumns);
        if (typeof savedDashboard.matrixRangeLabel === "string") setMatrixRangeLabel(savedDashboard.matrixRangeLabel);
        if (typeof savedDashboard.matrixFileName === "string") setMatrixFileName(savedDashboard.matrixFileName);
        if (savedDashboard.dashboardFocus) setDashboardFocus(savedDashboard.dashboardFocus);
        if (Array.isArray(savedDashboard.selectedStatusNames)) setSelectedStatusNames(savedDashboard.selectedStatusNames);
        if (typeof savedDashboard.statusFocus === "string" || savedDashboard.statusFocus === null) setStatusFocus(savedDashboard.statusFocus);
        if (typeof savedDashboard.callStartDate === "string") setCallStartDate(savedDashboard.callStartDate);
        if (typeof savedDashboard.callEndDate === "string") setCallEndDate(savedDashboard.callEndDate);
        if (savedDashboard.activeView === "calls" || savedDashboard.activeView === "matrix" || savedDashboard.activeView === "sales" || savedDashboard.activeView === "errors" || savedDashboard.activeView === "performance") setActiveView(savedDashboard.activeView);
        if (typeof savedDashboard.timelineDay === "string") setTimelineDay(savedDashboard.timelineDay);
        if (savedDashboard.timelineScale === "30m" || savedDashboard.timelineScale === "1h" || savedDashboard.timelineScale === "day") setTimelineScale(savedDashboard.timelineScale);
        if (typeof savedDashboard.showSlowResolutionOnly === "boolean") setShowSlowResolutionOnly(savedDashboard.showSlowResolutionOnly);
        if (typeof savedDashboard.emailHourChronological === "boolean") setEmailHourChronological(savedDashboard.emailHourChronological);
      }
      setPersistenceReady(true);
    }, 0);
  }, []);

  useEffect(() => {
    if (!persistenceReady) return;
    try {
      window.localStorage.setItem(persistedDashboardKey, JSON.stringify({
        matrixRows,
        matrixColumns,
        matrixRangeLabel,
        matrixFileName,
        dashboardFocus,
        selectedStatusNames,
        statusFocus,
        callStartDate,
        callEndDate,
        activeView,
        timelineDay,
        timelineScale,
        showSlowResolutionOnly,
        emailHourChronological,
      } satisfies PersistedDashboardState));
    } catch {
      // La persistencia es opcional; el dashboard sigue funcionando si el almacenamiento no tiene espacio.
    }
  }, [activeView, callEndDate, callStartDate, dashboardFocus, emailHourChronological, matrixColumns, matrixFileName, matrixRangeLabel, matrixRows, persistenceReady, selectedStatusNames, showSlowResolutionOnly, statusFocus, timelineDay, timelineScale]);

  useEffect(() => {
    function handleExtensionMessage(event: MessageEvent) {
      if (event.source !== window || event.data?.source !== "a365-extension") {
        return;
      }

      if (event.data.type === "A365_EXTENSION_READY") {
        setExtensionReady(true);
      }
    }

    window.addEventListener("message", handleExtensionMessage);
    window.postMessage({ source: "a365-dashboard", type: "A365_EXTENSION_PING" }, window.location.origin);

    return () => window.removeEventListener("message", handleExtensionMessage);
  }, []);

  function setFilter(field: FilterField, _label: string, value: string) {
    if (field === "statusName") {
      toggleStatusName(value);
      return;
    }

    setDashboardFocus((current) => ({
      ...current,
      [field]: current[field] === value ? undefined : value,
    }));
  }

  function toggleStatusName(statusName: string) {
    setSelectedStatusNames((current) => {
      const baseline = current.length ? current : allStatus.map((item) => item.name);
      return baseline.includes(statusName)
        ? baseline.filter((name) => name !== statusName)
        : [...baseline, statusName];
    });
  }

  async function exportChart(kind: ChartKind, title: string) {
    const element = chartRefs[kind].current;

    if (!element) {
      return;
    }

    const dataUrl = await toPng(element, {
      backgroundColor: "#ffffff",
      pixelRatio: 2,
      cacheBust: true,
    });
    const link = document.createElement("a");
    link.download = `${title.toLowerCase().replace(/\s+/g, "-")}.png`;
    link.href = dataUrl;
    link.click();
  }

  function pointIsSelected(field: FilterField, point: ChartPoint) {
    if (field === "statusName") return statusFocus === point.name;
    return dashboardFocus[field] === point.name;
  }

  function pointLabel(field: FilterField, point: ChartPoint) {
    if (pointIsSelected(field, point)) {
      return "En foco";
    }

    if (field === "statusName") {
      return formatPercent(percentage(point.total, statusTotal));
    }

    return String(point.total);
  }

  function focusStatus(name: string) {
    const next = statusFocus === name ? null : name;
    setStatusFocus(next);
    setDashboardFocus((filters) => ({ ...filters, statusName: next ?? undefined }));
  }

  function isDimmed(field: FilterField, name: string) {
    const focus = dashboardFocus[field];
    return Boolean(focus && focus !== name);
  }

  async function loadFile(file: File) {
    setError("");
    setFileName(file.name);

    try {
      const extension = file.name.split(".").pop()?.toLowerCase();
      let rawRows: RawRow[];

      if (extension === "xlsx" || extension === "xls") {
        rawRows = rowsFromSpreadsheet(await readSheet(file));
      } else {
        rawRows = parseDelimitedText(await file.text());
      }

      const mappedRows = mapRows(rawRows).filter(
        (row) => isInboundCall(row) && !row.campaignId.toUpperCase().includes("LINDESAC"),
      );
      if (!mappedRows.length) {
        throw new Error("No encontre llamadas inbound válidas después de excluir las campañas LINDESAC.");
      }

      setRows(mappedRows);
      setSelectedStatusNames([]);
      try {
        window.localStorage.setItem(persistedCallsKey, JSON.stringify({ fileName: file.name, columns: Object.keys(rawRows[0] ?? {}), rows: mappedRows }));
      } catch {
        // Archivos muy grandes pueden superar la cuota; se mantienen disponibles durante la sesión.
      }
    } catch (currentError) {
      setError(currentError instanceof Error ? currentError.message : "No pude leer el archivo.");
    }
  }

  function handleFiles(files: FileList | null) {
    const file = files?.[0];
    if (file) {
      void loadFile(file);
    }
  }

  async function loadIncidentFile(file: File) {
    setIncidentError("");
    setIncidentFileName(file.name);

    try {
      const extension = file.name.split(".").pop()?.toLowerCase();
      let rawRows: RawRow[];

      if (extension === "xlsx" || extension === "xls") {
        let sheetRows: Row[];

        try {
          sheetRows = await readSheet(file, "BASE");
        } catch {
          sheetRows = await readSheet(file);
        }

        rawRows = rowsFromSheetRows(sheetRows);
      } else {
        rawRows = parseDelimitedText(await file.text());
      }

      const mappedRows = mapIncidentRows(rawRows);

      if (!mappedRows.length) {
        throw new Error("No encontre incidencias validas en el archivo.");
      }

      setIncidentRows(mappedRows);
      setIncidentColumns(Object.keys(rawRows[0] ?? {}));
    } catch (currentError) {
      setIncidentError(currentError instanceof Error ? currentError.message : "No pude leer el archivo de incidencias.");
    }
  }

  function handleIncidentFiles(files: FileList | null) {
    const file = files?.[0];
    if (file) {
      void loadIncidentFile(file);
    }
  }

  function renderMatrixBarChart(data: ChartPoint[], color: string, height = 280) {
    return data.length ? (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart data={data} margin={{ left: 0, right: 16, top: 8, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar dataKey="total" name="Registros" fill={color} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Sin datos para graficar</EmptyChart>
    );
  }

  function renderMatrixTimeChart(data: TimeAveragePoint[], height = Math.max(280, Math.min(720, data.length * 42)), onPointClick?: (name: string) => void) {
    return data.length ? (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart
            data={data}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
            onClick={onPointClick ? (state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) onPointClick(value);
            } : undefined}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip
              formatter={(value, name) => [
                typeof value === "number" ? formatMinutes(value) : value,
                name === "asignacion" ? "Asignacion" : name === "resolucion" ? "Resolucion" : name,
              ]}
            />
            <Legend />
            <Bar dataKey="asignacion" name="Asignacion" fill="#2563a8" radius={[0, 4, 4, 0]} />
            <Bar dataKey="resolucion" name="Resolucion" fill="#0f766e" radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Sin tiempos para graficar</EmptyChart>
    );
  }

  function renderMatrixUserChart() {
    return matrixSummary.byUser.length ? (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={matrixUserChartHeight}>
          <BarChart
            data={matrixSummary.byUser}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
            onClick={(state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) setMatrixUserFocus((current) => current === value ? null : value);
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip />
            <Bar dataKey="total" name="Registros" fill="#2563a8" radius={[0, 4, 4, 0]} onClick={(entry) => { const value = matrixUserNameFromChartEntry(entry); if (value) setMatrixUserFocus((current) => current === value ? null : value); }}>
              {matrixSummary.byUser.map((entry) => <Cell key={entry.name} fill="#2563a8" opacity={matrixUserFocus && normalizeKey(matrixUserFocus) !== normalizeKey(entry.name) ? 0.2 : 1} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Sin usuarios para graficar</EmptyChart>
    );
  }

  function renderIncidentAgentChart() {
    return incidentSummary.byAgent.length ? (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={incidentAgentChartHeight}>
          <BarChart
            data={incidentSummary.byAgent}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f3d8e4" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={150} />
            <Tooltip />
            <Bar dataKey="total" name="Observaciones" fill="#2563a8" radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Sin agentes para graficar</EmptyChart>
    );
  }

  function renderPerformanceChart() {
    return performanceSummary.byAgent.length ? (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={performanceChartHeight}>
          <BarChart
            data={performanceSummary.byAgent}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f3d8e4" />
            <XAxis
              type="number"
              tick={{ fontSize: 12 }}
              tickFormatter={(value) => `${Number(value).toFixed(0)}%`}
            />
            <YAxis type="category" dataKey="agent" tick={{ fontSize: 12 }} width={150} />
            <Tooltip
              formatter={(value, name) => [
                typeof value === "number" ? formatPercent(value) : value,
                name === "productivity" ? "Productividad" : name === "quality" ? "Calidad" : "Efectividad",
              ]}
              labelFormatter={(label) => `Agente: ${label}`}
            />
            <Legend />
            <Bar dataKey="productivity" name="Productividad" fill="#2563a8" radius={[0, 4, 4, 0]} />
            <Bar dataKey="quality" name="Calidad" fill="#0f766e" radius={[0, 4, 4, 0]} />
            <Bar dataKey="effectiveness" name="Efectividad" fill="#475569" radius={[0, 4, 4, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Carga llamadas, correos o errores para ver indicadores</EmptyChart>
    );
  }

  function renderAgentTimeline() {
    const hourMarks = timelineMarks(timelineScale);

    return (
      <article className={styles.panel}>
        <div className={styles.panelHeader}>
          <div>
            <h2>Timeline de gestión por agente</h2>
            <span>Asignación y resolución por correo · selecciona un día y escala temporal</span>
          </div>
          {timelineDays.length ? (
            <select
              className={styles.timelineSelect}
              value={selectedTimelineDay}
              onChange={(event) => setTimelineDay(event.target.value)}
              aria-label="Dia del timeline"
            >
              {timelineDays.map((day) => (
                <option key={day} value={day}>
                  {day}
                </option>
              ))}
            </select>
          ) : null}
        </div>
        <div className={styles.timelineToolbar}>
          {timelineScales.map((scale) => (
            <button
              key={scale.value}
              type="button"
              className={timelineScale === scale.value ? styles.timelineScaleActive : ""}
              onClick={() => setTimelineScale(scale.value)}
            >
              {scale.label}
            </button>
          ))}
        </div>

        {timelineRows.length ? (
          <div className={styles.timelineLayout}>
            <div className={`${styles.timelineWrap} ${styles[`timelineScale_${timelineScale}`]}`}>
              <div className={styles.timelineHours} aria-hidden="true">
                <span>Agente</span>
                <div>
                  {hourMarks.map((hour) => (
                    <em key={hour}>{hour}</em>
                  ))}
                </div>
              </div>
              <div className={styles.timelineRows}>
                {timelineRows.map((row) => (
                  <div className={styles.timelineRow} key={row.agent}>
                    <strong title={row.agent}>
                      <i style={{ backgroundColor: colorForLabel(row.agent) }} />
                      {row.agent}
                    </strong>
                    <div className={styles.timelineTrack}>
                      {row.blocks.map((block) => (
                        <button
                          key={block.id}
                          type="button"
                          className={`${styles.timelineBlock} ${activeTimelineBlock?.id === block.id ? styles.timelineBlockActive : ""}`}
                          style={{
                            left: `${block.left}%`,
                            width: `${block.width}%`,
                            backgroundColor: block.color,
                          }}
                          onMouseEnter={() => setHoveredTimelineId(block.id)}
                          onFocus={() => setHoveredTimelineId(block.id)}
                          aria-label={block.detail}
                        />
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
            <aside className={styles.timelineDetail}>
              {activeTimelineBlock ? (
                <>
                  <span>Detalle del correo</span>
                  <strong>{activeTimelineBlock.subject}</strong>
                  <dl>
                    <div>
                      <dt>Agente</dt>
                      <dd>{activeTimelineBlock.agent}</dd>
                    </div>
                    <div>
                      <dt>Horario</dt>
                      <dd>{activeTimelineBlock.range}</dd>
                    </div>
                    <div>
                      <dt>Resolucion SLA</dt>
                      <dd>{formatMinutes(activeTimelineBlock.duration)}</dd>
                    </div>
                    <div>
                      <dt>Tipificacion</dt>
                      <dd>{activeTimelineBlock.tipificacion}</dd>
                    </div>
                    <div>
                      <dt>Estado</dt>
                      <dd>{activeTimelineBlock.estado}</dd>
                    </div>
                    <div>
                      <dt>ID email</dt>
                      <dd>{activeTimelineBlock.idEmail}</dd>
                    </div>
                  </dl>
                </>
              ) : null}
            </aside>
          </div>
        ) : (
          <EmptyChart>Sin correos resueltos en horario SLA para este dia</EmptyChart>
        )}
      </article>
    );
  }

  async function requestMatrixReportBlob() {
    if (isTauriRuntime()) {
      const { invoke } = await import("@tauri-apps/api/core");
      const report = await invoke<TauriMatrixReport>("download_matrix_report", {
        startDate: reportStartDate,
        endDate: reportEndDate,
        sessionId: reportSessionId,
      });

      return new Blob([base64ToUint8Array(report.dataBase64)], {
        type: report.contentType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
    }

    if (extensionReady) {
      return requestMatrixReportFromExtension();
    }

    const response = await fetch("/api/a365-report", {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: JSON.stringify({
          startDate: reportStartDate,
          endDate: reportEndDate,
          sessionId: reportSessionId,
        }),
      });

    if (!response.ok) {
      const payload = (await response.json().catch(() => null)) as { error?: string } | null;
      throw new Error(payload?.error ?? "No pude descargar el reporte.");
    }

    return response.blob();
  }

  function requestMatrixReportFromExtension() {
    return new Promise<Blob>((resolve, reject) => {
      const requestId = crypto.randomUUID();
      const timeout = window.setTimeout(() => {
        window.removeEventListener("message", handleResponse);
        reject(new Error("La extension no respondio. Recarga la web o revisa que este instalada."));
      }, 45000);

      function handleResponse(event: MessageEvent<ExtensionMatrixResponse>) {
        if (
          event.source !== window ||
          event.data?.source !== "a365-extension" ||
          event.data.type !== "A365_MATRIX_REPORT_RESPONSE" ||
          event.data.requestId !== requestId
        ) {
          return;
        }

        window.clearTimeout(timeout);
        window.removeEventListener("message", handleResponse);

        if (!event.data.ok || !event.data.dataBase64) {
          reject(new Error(event.data.error ?? "La extension no pudo descargar el reporte."));
          return;
        }

        resolve(new Blob([base64ToUint8Array(event.data.dataBase64)], {
          type: event.data.contentType || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }));
      }

      window.addEventListener("message", handleResponse);
      window.postMessage(
        {
          source: "a365-dashboard",
          type: "A365_MATRIX_REPORT_REQUEST",
          requestId,
          payload: {
            startDate: reportStartDate,
            endDate: reportEndDate,
            sessionId: reportSessionId,
          },
        },
        window.location.origin,
      );
    });
  }

  async function applyMatrixBlob(blob: Blob, label: string) {
    const file = new File([blob], `reporte-matriz-${reportStartDate}-${reportEndDate}.xlsx`, {
      type: blob.type || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    let rawRows: RawRow[] = [];

    try {
      rawRows = rowsFromSpreadsheet(await readSheet(file));
    } catch {
      rawRows = parseHtmlTableRows(await blob.text());
    }

    const mappedRows = mapMatrixRows(rawRows);

    if (!mappedRows.length) {
      throw new Error("La respuesta no trajo filas validas para la matriz.");
    }

    setMatrixRows(mappedRows);
    setMatrixColumns(Object.keys(rawRows[0] ?? {}));
    setMatrixRangeLabel(label);
    setTimelineDay(timelineDaysFromRows(mappedRows)[0] ?? "");
    setShowSlowResolutionOnly(false);
    setHoveredTimelineId(null);
  }

  async function loadMatrixReport(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setReportError("");
    setIsDownloadingReport(true);

    try {
      await applyMatrixBlob(await requestMatrixReportBlob(), `${reportStartDate} a ${reportEndDate}`);
    } catch (currentError) {
      setReportError(currentError instanceof Error ? currentError.message : "No pude leer el reporte.");
    } finally {
      setIsDownloadingReport(false);
    }
  }

  async function loadMatrixFile(file: File) {
    setReportError("");
    setMatrixFileName(file.name);
    setIsDownloadingReport(true);

    try {
      await applyMatrixBlob(file, file.name);
    } catch (currentError) {
      setReportError(currentError instanceof Error ? currentError.message : "No pude leer el archivo de matriz.");
    } finally {
      setIsDownloadingReport(false);
    }
  }

  function handleMatrixFiles(files: FileList | null) {
    const file = files?.[0];
    if (file) {
      void loadMatrixFile(file);
    }
  }

  function renderDateChart(height = 280) {
    return monthlyCallData.length ? (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart
            data={monthlyCallData}
            margin={{ left: 0, right: 16, top: 8, bottom: 0 }}
            onClick={(state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) {
                setFilter("month", "Mes", value);
              }
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Legend />
            <Bar isAnimationActive={false} dataKey="received" name="Llamadas recibidas" fill="#1d4ed8" radius={[4, 4, 0, 0]}>
              {monthlyCallData.map((entry) => <Cell key={`received-${entry.name}`} fill="#1d4ed8" opacity={isDimmed("month", entry.name) ? 0.2 : 1} />)}
            </Bar>
            <Bar isAnimationActive={false} dataKey="attended" name="Llamadas atendidas" fill="#0f766e" radius={[4, 4, 0, 0]}>
              {monthlyCallData.map((entry) => <Cell key={`attended-${entry.name}`} fill="#0f766e" opacity={isDimmed("month", entry.name) ? 0.2 : 1} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : (
      <EmptyChart>Sin meses para graficar</EmptyChart>
    );
  }

  async function loadSalesFile(file: File) {
    setSalesError("");
    try {
      const sheetRows = await readSheet(file, "RESUMEN");
      const mappedRows = mapSalesRows(sheetRows);
      if (!mappedRows.length) throw new Error("La hoja RESUMEN no contiene registros válidos.");
      setSalesRows(mappedRows);
      setSalesFileName(file.name);
      setSalesFilters({});
      setSalesTab("summary");
    } catch (currentError) {
      setSalesError(currentError instanceof Error ? currentError.message : "No pude leer la hoja RESUMEN.");
      setSalesRows([]);
    }
  }

  function handleSalesFiles(files: FileList | null) {
    const file = files?.[0];
    if (file) void loadSalesFile(file);
  }

  function renderMatrixMonthlyTimeChart(data: TimeAveragePoint[]) {
    return data.length ? (
      <ResponsiveContainer width="100%" height={300}>
        <BarChart data={data} margin={{ left: 8, right: 16, top: 8, bottom: 8 }}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey="name" />
          <YAxis allowDecimals={false} tickFormatter={(value) => `${value} min`} />
          <Tooltip formatter={(value, name) => [typeof value === "number" ? formatMinutes(value) : value, name === "asignacion" ? "Asignación" : "Resolución"]} />
          <Legend />
          <Bar dataKey="asignacion" name="Asignación" fill="#2563a8" radius={[4, 4, 0, 0]} />
          <Bar dataKey="resolucion" name="Resolución" fill="#0f766e" radius={[4, 4, 0, 0]} />
        </BarChart>
      </ResponsiveContainer>
    ) : <EmptyChart>Sin tiempos para graficar</EmptyChart>;
  }

  function renderEmailPieChart(data: ChartPoint[]) {
    const total = data.reduce((sum, point) => sum + point.total, 0);
    return data.length ? (
      <div className={styles.statusChartLayout}>
        <div className={styles.pieFrame}>
          <ResponsiveContainer width="100%" height={300}>
            <PieChart>
              <Pie isAnimationActive={false} data={data} dataKey="total" nameKey="name" innerRadius={64} outerRadius={104} paddingAngle={2}>
                {data.map((entry, index) => <Cell key={entry.name} fill={chartColor(index)} />)}
              </Pie>
              <Tooltip formatter={(value) => [formatPercent(percentage(Number(value), total)), "Porcentaje"]} />
            </PieChart>
          </ResponsiveContainer>
          <div className={styles.pieCenter}>
            <span>Total</span>
            <strong>{total.toLocaleString("es-PE")}</strong>
          </div>
        </div>
        <div className={styles.statusLegend}>
          {data.map((item, index) => (
            <div key={item.name} className={styles.legendItem}>
              <span style={{ backgroundColor: chartColor(index) }} /><strong>{item.name}</strong><em>{formatPercent(percentage(item.total, total))}</em>
            </div>
          ))}
        </div>
      </div>
    ) : <EmptyChart>Sin datos para graficar</EmptyChart>;
  }

  function renderEmailMonthlyChart() {
    return <ResponsiveContainer width="100%" height={300}><BarChart data={emailSummary.byMonth}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis allowDecimals={false} /><Tooltip /><Bar isAnimationActive={false} dataKey="received" name="Correos recibidos" fill="#12355b" /></BarChart></ResponsiveContainer>;
  }

  function renderEmailQualityChart() {
    return <ResponsiveContainer width="100%" height={300}><LineChart data={emailSummary.byMonth}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis domain={[0, 100]} tickFormatter={(value) => `${value}%`} /><Tooltip formatter={(value) => [`${Number(value).toFixed(1)}%`, "Porcentaje"]} /><Legend /><Line isAnimationActive={false} type="monotone" dataKey="nda" name="NDA" stroke="#12355b" strokeWidth={3} /><Line isAnimationActive={false} type="monotone" dataKey="nds" name="NDS" stroke="#f97316" strokeWidth={3} /></LineChart></ResponsiveContainer>;
  }

  function renderEmailTmoChart() {
    return <ResponsiveContainer width="100%" height={300}><BarChart data={emailSummary.byMonth}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis tickFormatter={(value) => `${Number(value).toFixed(0)} min`} /><Tooltip formatter={(value) => [formatEmailTmo(Number(value)), "TMO"]} /><Bar isAnimationActive={false} dataKey="tmoMinutes" name="TMO" fill="#2563a8" /></BarChart></ResponsiveContainer>;
  }

  function renderEmailHourChart() {
    return <ResponsiveContainer width="100%" height={300}><BarChart data={emailHourData}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis allowDecimals={false} /><Tooltip /><Bar isAnimationActive={false} dataKey="total" name="Correos" fill="#0f766e" /></BarChart></ResponsiveContainer>;
  }

  function renderHourChart(height = 280) {
    return (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart
            data={byHour}
            margin={{ left: 0, right: 16, top: 8, bottom: 0 }}
            onClick={(state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) {
                setFilter("hour", "Hora", value);
              }
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar
              isAnimationActive={false}
              dataKey="total"
              name="Llamadas recibidas"
              fill="#2563a8"
              radius={[4, 4, 0, 0]}
              onClick={(entry) => {
                const value = chartPointName(entry);
                if (value) {
                  setFilter("hour", "Hora", value);
                }
              }}
            >
              {byHour.map((entry) => <Cell key={entry.name} fill="#2563a8" opacity={isDimmed("hour", entry.name) ? 0.2 : 1} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    );
  }

  function renderServiceQualityChart() {
    return serviceSummary.byMonth.length ? (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={300}>
          <LineChart data={serviceSummary.byMonth} margin={{ left: 0, right: 16, top: 8, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis domain={[0, 100]} tickFormatter={(value) => `${value}%`} tick={{ fontSize: 12 }} />
            <Tooltip formatter={(value) => [formatPercent(Number(value)), "Porcentaje"]} />
            <Legend />
            <Line isAnimationActive={false} type="monotone" dataKey="attentionRate" name="Nivel de atención" stroke="#1d4ed8" strokeWidth={3} dot={{ r: 4 }} />
          </LineChart>
        </ResponsiveContainer>
      </div>
    ) : <EmptyChart>Sin meses para calcular calidad</EmptyChart>;
  }

  function renderServiceDurationChart() {
    return serviceSummary.byMonth.length ? (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={serviceSummary.byMonth} margin={{ left: 0, right: 16, top: 8, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis tickFormatter={(value) => formatCallDuration(Number(value))} tick={{ fontSize: 12 }} />
            <Tooltip formatter={(value) => [formatCallDuration(Number(value)), "TMO"]} />
            <Bar dataKey="averageDurationSeconds" name="TMO" fill="#1d4ed8" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : <EmptyChart>Sin llamadas atendidas para calcular TMO</EmptyChart>;
  }

  function renderAbandonmentChart() {
    return serviceSummary.byMonth.length ? (
      <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={serviceSummary.byMonth} margin={{ left: 0, right: 16, top: 8, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar dataKey="abandoned" name="Llamadas abandonadas" fill="#f97316" radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>
    ) : <EmptyChart>Sin meses para calcular abandono</EmptyChart>;
  }

  function renderCampaignChart(height = campaignChartHeight) {
    return (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart
            data={byCampaign}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
            onClick={(state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) {
                setFilter("campaignId", "Campañas", value);
              }
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#f8dbe8" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip />
            <Bar
              isAnimationActive={false}
              dataKey="total"
              name="Llamadas recibidas"
              fill="#0f766e"
              radius={[0, 4, 4, 0]}
              onClick={(entry) => {
                const value = chartPointName(entry);
                if (value) {
                  setFilter("campaignId", "Campañas", value);
                }
              }}
            >
              {byCampaign.map((entry) => <Cell key={entry.name} fill="#0f766e" opacity={isDimmed("campaignId", entry.name) ? 0.2 : 1} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    );
  }

  function renderDurationDistribution() {
    return <ResponsiveContainer width="100%" height={280}><BarChart data={durationDistribution}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis allowDecimals={false} /><Tooltip /><Bar isAnimationActive={false} dataKey="total" name="Llamadas" fill="#2563a8" radius={[4, 4, 0, 0]} /></BarChart></ResponsiveContainer>;
  }

  function renderUserPerformance() {
    return <ResponsiveContainer width="100%" height={Math.max(280, Math.min(620, userPerformance.length * 34))}><BarChart data={userPerformance} layout="vertical" margin={{ left: 24, right: 20 }} onClick={(state) => { const value = chartClickValue(state as ChartClickState); if (value) setFilter("user", "Usuario", value); }}><CartesianGrid strokeDasharray="3 3" /><XAxis type="number" allowDecimals={false} /><YAxis type="category" dataKey="name" width={110} /><Tooltip /><Legend /><Bar isAnimationActive={false} dataKey="received" name="Recibidas" fill="#1d4ed8" onClick={(entry) => { const value = chartPointName(entry); if (value) setFilter("user", "Usuario", value); }}>{userPerformance.map((entry) => <Cell key={`received-${entry.name}`} fill="#1d4ed8" opacity={isDimmed("user", entry.name) ? 0.2 : 1} />)}</Bar><Bar isAnimationActive={false} dataKey="attended" name="Atendidas" fill="#0f766e" onClick={(entry) => { const value = chartPointName(entry); if (value) setFilter("user", "Usuario", value); }}>{userPerformance.map((entry) => <Cell key={`attended-${entry.name}`} fill="#0f766e" opacity={isDimmed("user", entry.name) ? 0.2 : 1} />)}</Bar></BarChart></ResponsiveContainer>;
  }

  function renderRawCallRows() {
    const rawColumns = Object.keys(visibleRows[0]?.raw ?? {});
    const previewRows = visibleRows.slice(0, 100);
    return (
      <div className={styles.rawDataBlock}>
        <div className={styles.rawDataHeader}><strong>Datos crudos filtrados</strong><span>{visibleRows.length.toLocaleString("es-PE")} filas{visibleRows.length > previewRows.length ? ` · mostrando ${previewRows.length}` : ""}</span></div>
        <div className={styles.rawDataScroll}>
          <table className={styles.rawDataTable}>
            <thead><tr>{rawColumns.map((column) => <th key={column}>{column}</th>)}</tr></thead>
            <tbody>{previewRows.map((row) => <tr key={row.id}>{rawColumns.map((column) => <td key={`${row.id}-${column}`}>{row.raw[column] ?? ""}</td>)}</tr>)}</tbody>
          </table>
        </div>
      </div>
    );
  }

  function renderStatusChart(height = 280) {
    return (
      <div className={styles.statusChartLayout}>
        <div className={styles.pieFrame}>
          <ResponsiveContainer width="100%" height={height}>
            <PieChart>
              <Pie isAnimationActive={false} data={statusChartData} dataKey="total" nameKey="name" innerRadius={height > 300 ? 82 : 56} outerRadius={height > 300 ? 150 : 98} paddingAngle={2} onClick={(entry) => focusStatus(String((entry as unknown as ChartPoint).name))}>
                {statusChartData.map((entry) => <Cell key={entry.name} fill={statusColorFor(entry.name)} opacity={statusFocus && statusFocus !== entry.name ? 0.2 : 1} style={statusFocus && statusFocus !== entry.name ? { filter: "grayscale(1)" } : undefined} />)}
              </Pie>
              <Tooltip formatter={(value) => [formatPercent(percentage(Number(value), statusTotal)), "Porcentaje"]} />
            </PieChart>
          </ResponsiveContainer>
          <div className={styles.pieCenter}>
            <span>{statusFocus ? "Seleccionado" : "Total"}</span>
            <strong>{(statusFocus ? statusChartData.find((item) => item.name === statusFocus)?.total ?? 0 : statusTotal).toLocaleString("es-PE")}</strong>
          </div>
        </div>
        <div className={styles.statusLegend}>
          {statusChartData.map((item) => (
            <button key={item.name} type="button" className={`${styles.legendItem} ${pointIsSelected(statusChartField, item) ? styles.focusedItem : ""}`} onClick={() => focusStatus(item.name)} style={statusFocus && statusFocus !== item.name ? { opacity: 0.42, filter: "grayscale(1)" } : undefined}>
              <span style={{ backgroundColor: statusColorFor(item.name) }} /><strong>{item.name}</strong><em>{pointLabel(statusChartField, item)}</em>
            </button>
          ))}
        </div>
      </div>
    );
  }

  function renderSalesBar(data: ChartPoint[], color = "#2563a8") {
    return data.length ? <ResponsiveContainer width="100%" height={Math.max(280, Math.min(560, data.length * 34))}><BarChart data={data} layout="vertical" margin={{ left: 24, right: 24 }}><CartesianGrid strokeDasharray="3 3" /><XAxis type="number" tickFormatter={(value) => formatSalesUSD(Number(value))} /><YAxis type="category" dataKey="name" width={150} tick={{ fontSize: 11 }} /><Tooltip formatter={(value) => formatSalesUSD(Number(value))} /><Bar dataKey="total" name="Venta USD" fill={color} radius={[0, 4, 4, 0]} /></BarChart></ResponsiveContainer> : <EmptyChart>Sin datos para graficar</EmptyChart>;
  }

  function renderSalesPie(data: ChartPoint[]) {
    const total = data.reduce((sum, item) => sum + item.total, 0);
    return data.length ? <div className={styles.statusChartLayout}><div className={styles.pieFrame}><ResponsiveContainer width="100%" height={300}><PieChart><Pie data={data} dataKey="total" nameKey="name" innerRadius={64} outerRadius={105} paddingAngle={2}>{data.map((item, index) => <Cell key={item.name} fill={chartColor(index)} />)}</Pie><Tooltip formatter={(value) => `${formatSalesUSD(Number(value))} · ${formatPercent(percentage(Number(value), total))}`} /></PieChart></ResponsiveContainer><div className={styles.pieCenter}><span>Total</span><strong>{formatSalesUSD(total)}</strong></div></div><div className={styles.statusLegend}>{data.map((item, index) => <div key={item.name} className={styles.legendItem}><span style={{ backgroundColor: chartColor(index) }} /><strong>{salesDisplayName(item.name)}</strong><em>{formatPercent(percentage(item.total, total))}</em></div>)}</div></div> : <EmptyChart>Sin datos para graficar</EmptyChart>;
  }

  function renderSalesContent() {
    const monthly = salesMonths.map((item) => ({ name: item.name, total: item.total }));
    const active = salesTab === "summary";
    return (
      <>
        <section className={styles.reportPanel} onClick={stopInsideClick}>
          <div><small className={styles.reportKicker}>Indicador 03 · Análisis comercial</small><h2>I3 - Análisis comercial / Inside</h2><span className={styles.reportDescription}>Hoja utilizada: RESUMEN · Archivo: {salesFileName} · {salesRows.length.toLocaleString("es-PE")} registros</span></div>
          <div className={styles.inlineUpload} onClick={() => salesInputRef.current?.click()} role="button" tabIndex={0} onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") salesInputRef.current?.click(); }}><FileSpreadsheet size={18} /><span>{salesFileName}</span><strong>Cargar Excel de ventas</strong><input ref={salesInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={(event) => handleSalesFiles(event.target.files)} /></div>
          {salesError ? <strong className={styles.error}>{salesError}</strong> : null}
        </section>
        {salesRows.length ? <>
          <section className={styles.filterBar} onClick={stopInsideClick}><div className={styles.filterBarHeader}><strong>Filtros comerciales</strong><div className={styles.filterActions}><button type="button" className={styles.chartActionButton} onClick={() => setSalesCompareRegions((current) => !current)}>{salesCompareRegions ? "Salir de VS" : "Abrir VS"}</button><button type="button" className={styles.chartActionButton} onClick={() => setSalesFilters({})}>Limpiar filtros</button></div></div><div className={styles.filterGrid}>{(["year", "month", "region", "branch", "clientName", "className", "group", "business", "industry", "application", "productDescription", "unit"] as Array<keyof SalesRow>).map((key) => { const values = Array.from(new Set(salesRows.map((row) => String(row[key] ?? "")).filter(Boolean))).sort((a, b) => a.localeCompare(b)); return <label key={key}>{key === "clientName" ? "Cliente" : key === "className" ? "Clase" : key === "productDescription" ? "Producto" : key}<select value={salesFilters[key] ?? ""} onChange={(event) => setSalesFilters((current) => ({ ...current, [key]: event.target.value }))}><option value="">Todos</option>{values.map((value) => <option key={value} value={value}>{salesDisplayName(value)}</option>)}</select></label>; })}</div></section>
          <nav className={styles.subNav} aria-label="Secciones I3">{([ ["summary", "Resumen ejecutivo"], ["evolution", "Evolución"], ["clients", "Clientes"] ] as Array<[SalesTab, string]>).map(([value, label]) => <button key={value} type="button" className={salesTab === value ? styles.topNavActive : ""} onClick={() => setSalesTab(value)}>{label}</button>)}</nav>
          {salesCompareRegions && salesRegions.length === 2 ? <section className={styles.regionCompareSection}><div className={styles.regionCompareTitle}><div><small>Comparador</small><h2>VS · Comparación por región</h2></div><span>Dos regiones con los mismos indicadores y gráficos</span></div><div className={styles.regionCompareGrid}>{salesRegions.map((region) => { const rows = salesComparisonRows.filter((row) => row.region === region); const total = rows.reduce((sum, row) => sum + row.amountUSD, 0); return <article className={styles.regionCompareCard} key={region}><div className={styles.regionCompareHeader}><h3>{salesDisplayName(region)}</h3><span>{rows.length.toLocaleString("es-PE")} registros</span></div><div className={styles.statsGrid}><StatCard icon={<Hash size={18} />} label="Venta total (USD)" value={formatSalesUSD(total)} /><StatCard icon={<UserRound size={18} />} label="Clientes únicos" value={salesCountDistinct(rows, "clientCode")} /><StatCard icon={<Timer size={18} />} label="Ticket promedio (USD)" value={formatSalesUSD(rows.length ? total / rows.length : 0)} /><StatCard icon={<Timer size={18} />} label="Venta promedio por cliente" value={formatSalesUSD(salesCountDistinct(rows, "clientCode") ? total / salesCountDistinct(rows, "clientCode") : 0)} /></div><ChartPanel title="Ventas mensuales" meta="SUM Importe Vendido US$"><ResponsiveContainer width="100%" height={250}><LineChart data={salesGrouped(rows, "month")}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis tickFormatter={(value) => formatSalesUSD(Number(value))} /><Tooltip formatter={(value) => formatSalesUSD(Number(value))} /><Line type="monotone" dataKey="total" name="Venta USD" stroke="#12355b" strokeWidth={3} dot /></LineChart></ResponsiveContainer></ChartPanel><ChartPanel title="Top 5 clientes" meta="Venta USD">{renderSalesBar(salesGrouped(rows, "clientName").slice(0, 5), "#475569")}</ChartPanel><ChartPanel title="Venta por negocio" meta="Venta USD">{renderSalesBar(salesGrouped(rows, "business").slice(0, 10), "#7c3aed")}</ChartPanel></article>; })}</div></section> : null}\n          {active && !salesCompareRegions ? <><section className={styles.statsGrid}><StatCard icon={<Hash size={18} />} label="Venta total" value={formatSalesUSD(salesTotalUSD)} /><StatCard icon={<UserRound size={18} />} label="Clientes únicos" value={salesCountDistinct(salesFilteredRows, "clientCode")} /><StatCard icon={<Timer size={18} />} label="Ticket promedio" value={formatSalesUSD(salesFilteredRows.length ? salesTotalUSD / salesFilteredRows.length : 0)} /><StatCard icon={<Timer size={18} />} label="Venta promedio por cliente" value={formatSalesUSD(salesCountDistinct(salesFilteredRows, "clientCode") ? salesTotalUSD / salesCountDistinct(salesFilteredRows, "clientCode") : 0)} /></section><div className={styles.dashboardGrid}><ChartPanel title="Evolución mensual de ventas (USD)" meta="SUM Importe Vendido US$"><ResponsiveContainer width="100%" height={300}><LineChart data={monthly}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis tickFormatter={(value) => formatSalesUSD(Number(value))} /><Tooltip formatter={(value) => formatSalesUSD(Number(value))} /><Line type="monotone" dataKey="total" name="Venta USD" stroke="#12355b" strokeWidth={3} dot /></LineChart></ResponsiveContainer></ChartPanel><ChartPanel title="Venta por sucursal" meta="Ordenado de mayor a menor">{renderSalesBar(salesBranches)}</ChartPanel><ChartPanel title="Participación por clase" meta="Porcentaje sobre venta filtrada">{renderSalesPie(salesClasses)}</ChartPanel><ChartPanel title="Top 10 productos por venta" meta="SUM Importe Vendido US$">{renderSalesBar(salesProducts, "#0f766e")}</ChartPanel><ChartPanel title="Top 5 mejores clientes" meta="SUM Importe Vendido US$">{renderSalesBar(salesClients, "#475569")}</ChartPanel><ChartPanel title="Venta por industria" meta="Top 10">{renderSalesBar(salesIndustries, "#0891b2")}</ChartPanel></div></> : null}
          {salesTab === "evolution" ? <div className={styles.dashboardGrid}><ChartPanel title="Venta diaria (USD)" meta="Fecha = Año + Mes + Dia"><ResponsiveContainer width="100%" height={320}><LineChart data={salesGrouped(salesFilteredRows, "date")}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis tickFormatter={(value) => formatSalesUSD(Number(value))} /><Tooltip formatter={(value) => formatSalesUSD(Number(value))} /><Line type="monotone" dataKey="total" name="Venta USD" stroke="#12355b" dot={false} /></LineChart></ResponsiveContainer></ChartPanel><ChartPanel title="Evolución del volumen vendido" meta="SUM Volumen en Tonelada"><ResponsiveContainer width="100%" height={320}><LineChart data={salesMonths}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis /><Tooltip /><Line type="monotone" dataKey="tons" name="Toneladas" stroke="#0f766e" /></LineChart></ResponsiveContainer></ChartPanel><ChartPanel title="Registros y clientes activos por mes" meta="Filas y COUNT DISTINCT Cod.Clte"><ResponsiveContainer width="100%" height={320}><BarChart data={salesMonths}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis /><Tooltip /><Legend /><Bar dataKey="records" name="Registros" fill="#2563a8" /><Bar dataKey="clients" name="Clientes" fill="#0f766e" /></BarChart></ResponsiveContainer></ChartPanel></div> : null}
          {salesTab === "branches" ? <div className={styles.dashboardGrid}><ChartPanel title="Venta por sucursal" meta="SUM Importe Vendido US$">{renderSalesBar(salesBranches)}</ChartPanel><ChartPanel title="Participación por sucursal" meta="Porcentaje">{renderSalesPie(salesBranches)}</ChartPanel></div> : null}
          {salesTab === "products" ? <div className={styles.dashboardGrid}><ChartPanel title="Top 10 productos por venta" meta="Producto + descripción">{renderSalesBar(salesProducts, "#0f766e")}</ChartPanel><ChartPanel title="Venta por clase" meta="Venta y participación">{renderSalesPie(salesClasses)}</ChartPanel><ChartPanel title="Venta por unidad" meta="SUM Importe Vendido US$">{renderSalesBar(salesGrouped(salesFilteredRows, "unit"), "#0891b2")}</ChartPanel></div> : null}
          {salesTab === "clients" ? <div className={styles.dashboardGrid}><ChartPanel title="Top 5 mejores clientes" meta="Cliente + venta USD">{renderSalesBar(salesClients, "#475569")}</ChartPanel><ChartPanel title="Clientes activos por mes" meta="COUNT DISTINCT Cod.Clte"><ResponsiveContainer width="100%" height={300}><LineChart data={salesMonths}><CartesianGrid strokeDasharray="3 3" /><XAxis dataKey="name" /><YAxis /><Tooltip /><Line type="monotone" dataKey="clients" name="Clientes" stroke="#12355b" /></LineChart></ResponsiveContainer></ChartPanel></div> : null}
          {salesTab === "industries" ? <div className={styles.dashboardGrid}><ChartPanel title="Distribución de venta por industria" meta="Top 10">{renderSalesBar(salesIndustries, "#0891b2")}</ChartPanel><ChartPanel title="Venta por aplicación" meta="Top 10">{renderSalesBar(salesGrouped(salesFilteredRows, "application").slice(0, 10), "#0f766e")}</ChartPanel><ChartPanel title="Venta por negocio" meta="Top 10">{renderSalesBar(salesGrouped(salesFilteredRows, "business").slice(0, 10), "#2563a8")}</ChartPanel></div> : null}
          {salesTab === "advanced" ? <div className={styles.dashboardGrid}><ChartPanel title="Matriz sucursal vs clase" meta="Venta USD"><EmptyChart>Vista de matriz pendiente de dimensiones compatibles</EmptyChart></ChartPanel><ChartPanel title="Venta por documento" meta="Cantidad de registros y venta">{renderSalesBar(salesGrouped(salesFilteredRows, "documentType"), "#2563a8")}</ChartPanel><ChartPanel title="Venta por grupo" meta="Participación comercial">{renderSalesBar(salesGrouped(salesFilteredRows, "group"), "#0f766e")}</ChartPanel></div> : null}
          {active && !salesCompareRegions ? <section className={styles.dashboardGrid}><ChartPanel title="Frecuencia de compra · Top 10 clientes" meta="Promedio de días entre compras consecutivas · ordenado por venta"><div className={styles.frequencyChart}>{salesFrequency.length ? <div className={styles.frequencyList}>{salesFrequency.map((item) => <div className={styles.frequencyRow} key={item.name}><div><strong>{salesDisplayName(item.name)}</strong><span>{item.purchases} compras · {formatSalesUSD(item.salesTotal)} acumulado</span></div><b>{item.total.toFixed(0)} días</b></div>)}</div> : <EmptyChart>No hay clientes con más de una compra en el periodo filtrado</EmptyChart>}</div></ChartPanel></section> : null}
          {!salesCompareRegions ? <article className={styles.panel}>
            <div className={styles.panelHeader}><div><h2>Detalle de registros</h2><span>{salesFilteredRows.length.toLocaleString("es-PE")} filas filtradas · mostrando hasta 200</span></div></div>
            <div className={styles.tableWrap}><table><thead><tr><th>Fecha</th><th>Sucursal</th><th>Cliente</th><th>Producto</th><th>Clase</th><th>Grupo</th><th>Negocio</th><th>Industria</th><th>Aplicación</th><th>Volumen</th><th>Toneladas</th><th>Importe S/.</th><th>Importe USD</th></tr></thead><tbody>{salesFilteredRows.slice(0, 200).map((row) => <tr key={row.id}><td>{row.date}</td><td>{row.branch}</td><td>{row.clientName}</td><td>{row.productDescription}</td><td>{salesDisplayName(row.className)}</td><td>{row.group}</td><td>{row.business}</td><td>{row.industry}</td><td>{row.application}</td><td>{row.volumeSold.toLocaleString("es-PE")}</td><td>{row.tons.toLocaleString("es-PE", { maximumFractionDigits: 3 })}</td><td>{formatSalesPEN(row.amountPEN)}</td><td>{formatSalesUSD(row.amountUSD)}</td></tr>)}</tbody></table></div>
          </article> : null}
        </> : null}
      </>
    );
  }

  function expandedContent() {
    if (expandedChart === "date") {
      return renderDateChart(520);
    }
    if (expandedChart === "hour") {
      return renderHourChart(520);
    }
    if (expandedChart === "status") {
      return renderStatusChart(520);
    }
    if (expandedChart === "campaign") {
      return renderCampaignChart(Math.max(520, Math.min(1200, byCampaign.length * 38)));
    }
    return null;
  }

  function submitAccess(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (accessPassword !== "Palabra!Aleatoria2026#") {
      setAccessPassword("");
      return;
    }
    document.cookie = "a365_access=1; max-age=2592000; path=/; SameSite=Lax";
    setAccessGranted(true);
    setAccessPassword("");
  }

  return (
    <main className={styles.page}>
      {accessChecked && !accessGranted ? <div className={styles.accessBackdrop}><form className={styles.accessModal} onSubmit={submitAccess} autoComplete="off"><div className={styles.accessBrand}>A365</div><h2>Acceso al dashboard</h2><p>Ingresa la contraseña para continuar.</p><label htmlFor="dashboard-access-password">Contraseña</label><input id="dashboard-access-password" type="password" value={accessPassword} onChange={(event) => setAccessPassword(event.target.value)} autoFocus /><button type="submit">Ingresar</button></form></div> : null}
      <section className={styles.header} onClick={stopInsideClick}>
        <div>
          <h1>Control operativo</h1>
        </div>
        <a
          className={styles.downloadAppLink}
          href="/a365-extension.zip"
          download
        >
          <Download size={18} aria-hidden="true" />
          <span>Descargar extension local</span>
          <strong>Chrome/Edge</strong>
        </a>
      </section>

      <nav className={styles.topNav} onClick={stopInsideClick} aria-label="Navegacion de indicadores">
        <button
          type="button"
          className={activeView === "calls" ? styles.topNavActive : ""}
          onClick={() => setActiveView("calls")}
        >
          I1 - Llamadas inbound
        </button>
        <button
          type="button"
          className={activeView === "matrix" ? styles.topNavActive : ""}
          onClick={() => setActiveView("matrix")}
        >
          I2 - Correos atendidos
        </button>
        <button
          type="button"
          className={activeView === "sales" ? styles.topNavActive : ""}
          onClick={() => setActiveView("sales")}
        >
          I3 - Análisis comercial
        </button>
        <button
          type="button"
          className={activeView === "errors" ? styles.topNavActive : ""}
          onClick={() => setActiveView("errors")}
        >
          I4 - Tasa de error
        </button>
        <button
          type="button"
          className={activeView === "performance" ? styles.topNavActive : ""}
          onClick={() => setActiveView("performance")}
        >
          I5 - Rendimiento operativo
        </button>
      </nav>

      {activeView === "calls" ? (
        <>
          <section className={styles.reportPanel} onClick={stopInsideClick}>
            <div>
              <small className={styles.reportKicker}>Indicador 01 · Canal inbound</small>
              <h2>I1 - Llamadas inbound</h2>
              <span className={styles.reportDescription}>Volumen, atención, campañas y horarios · LINDESAC excluida · <a href="/MANUAL_I1.txt" download>Descargar manual I1</a></span>
            </div>
            <div
              className={`${styles.inlineUpload} ${isDragging ? styles.dropzoneActive : ""}`}
              onClick={() => inputRef.current?.click()}
              onDragEnter={(event) => {
                event.preventDefault();
                setIsDragging(true);
              }}
              onDragOver={(event) => event.preventDefault()}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(event) => {
                event.preventDefault();
                setIsDragging(false);
                handleFiles(event.dataTransfer.files);
              }}
              role="button"
              tabIndex={0}
              onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === " ") {
                  inputRef.current?.click();
                }
              }}
            >
              <Upload size={18} aria-hidden="true" />
              <span>{fileName}</span>
              <strong>Cargar archivo</strong>
              <input
                ref={inputRef}
                type="file"
                accept=".txt,.csv,.tsv,.xlsx,.xls"
                onChange={(event) => handleFiles(event.target.files)}
              />
            </div>
            <div className={styles.callDateFilter}>
              <span>Filtrar intervalo de fechas</span>
              <label>Desde<input type="date" value={callStartDate} onChange={(event) => setCallStartDate(event.target.value)} /></label>
              <label>Hasta<input type="date" value={callEndDate} onChange={(event) => setCallEndDate(event.target.value)} /></label>
              {(callStartDate || callEndDate) ? <button type="button" onClick={() => { setCallStartDate(""); setCallEndDate(""); }}>Limpiar</button> : null}
            </div>
          </section>

          {error ? <div className={styles.error} onClick={stopInsideClick}>{error}</div> : null}

          {rows.length ? (
            <section className={styles.fileBar} onClick={stopInsideClick}>
              <FileSpreadsheet size={18} aria-hidden="true" />
              <span>{fileName}</span>
              <strong>
                {visibleRows.length.toLocaleString("es-PE")} de {dateFilteredRows.length.toLocaleString("es-PE")} registros
              </strong>
            </section>
          ) : (
            <section className={styles.emptyState} onClick={stopInsideClick}>
              <Phone size={22} aria-hidden="true" />
              <strong>I1 - Llamadas inbound listo</strong>
              <span>Carga un TXT, CSV o Excel para empezar sin datos de ejemplo.</span>
            </section>
          )}
        </>
      ) : null}

      {activeView === "matrix" ? (
        <>
          <section className={styles.reportPanel} onClick={stopInsideClick}>
            <div>
              <h2>Matriz A365</h2>
              <span>
                {extensionReady
                  ? "Extension conectada: la descarga saldra desde tu navegador local."
                  : "Ingresa tu PHPSESSID y el rango de fechas. En Vercel instala la extension local."}
              </span>
            </div>
            <form className={styles.reportForm} onSubmit={loadMatrixReport}>
              <label>
                Desde
                <input type="date" value={reportStartDate} onChange={(event) => setReportStartDate(event.target.value)} required />
              </label>
              <label>
                Hasta
                <input type="date" value={reportEndDate} onChange={(event) => setReportEndDate(event.target.value)} required />
              </label>
              <label>
                PHPSESSID
                <input
                  type="password"
                  value={reportSessionId}
                  onChange={(event) => setReportSessionId(event.target.value)}
                  placeholder="PHPSESSID=... o solo el valor"
                  autoComplete="off"
                  required
                />
              </label>
              <button type="submit" disabled={isDownloadingReport}>
                <Download size={16} aria-hidden="true" />
                {isDownloadingReport ? "Analizando" : "Leer matriz"}
              </button>
            </form>
            <div
              className={styles.inlineUpload}
              onClick={() => matrixInputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === " ") {
                  matrixInputRef.current?.click();
                }
              }}
            >
              <FileSpreadsheet size={18} aria-hidden="true" />
              <span>{matrixFileName}</span>
              <strong>Cargar Excel local</strong>
              <input
                ref={matrixInputRef}
                type="file"
                accept=".xlsx,.xls,.csv,.txt,.tsv"
                onChange={(event) => handleMatrixFiles(event.target.files)}
              />
            </div>
            {reportError ? <strong>{reportError}</strong> : null}
          </section>

          {matrixRows.length ? (
            <section className={styles.matrixSection} onClick={stopInsideClick}>
          <div className={styles.matrixHeader}>
            <div>
              <span className={styles.eyebrow}>Reporte matriz</span>
              <h2>Correos 2026 · desempeño operativo</h2>
              <p>{emailSummary.period} · SLA configurado: {EMAIL_SLA_MINUTES} min · {matrixRangeLabel}</p>
            </div>
            <div className={styles.matrixActions}>
              <button
                type="button"
                className={showSlowResolutionOnly ? styles.matrixActionActive : ""}
                onClick={() => {
                  setShowSlowResolutionOnly((current) => !current);
                  setHoveredTimelineId(null);
                }}
              >
                <ListFilter size={16} aria-hidden="true" />
                &gt; 20 min resolucion
              </button>
              <strong>
                {visibleMatrixRows.length.toLocaleString("es-PE")} de {matrixRows.length.toLocaleString("es-PE")} registros
              </strong>
            </div>
          </div>

          <div className={styles.statsGrid}>
            <StatCard icon={<Mail size={18} />} label="Correos recibidos" value={emailSummary.received.toLocaleString("es-PE")} />
            <StatCard icon={<UserRound size={18} />} label="NDA" value={formatPercent(emailSummary.nda)} />
            <StatCard icon={<Clock3 size={18} />} label="NDS" value={formatPercent(emailSummary.nds)} />
            <StatCard icon={<Timer size={18} />} label="TMO promedio" value={formatEmailTmo(emailSummary.tmoMinutes)} />
            <StatCard icon={<Hash size={18} />} label="Correos atendidos" value={emailSummary.attended.toLocaleString("es-PE")} />
          </div>

          {renderAgentTimeline()}

          <div className={styles.matrixGrid}>
            <ChartPanel title="Evolución mensual de correos recibidos" meta="COUNT DISTINCT id_email" comparison={emailVolumeComparison} info="Campos: date_email e id_email. Agrupa por año-mes y cuenta correos únicos.">
              {renderEmailMonthlyChart()}
            </ChartPanel>
            <ChartPanel title="Evolución de indicadores de calidad" meta="NDA vs NDS · fechas de registro" comparison={<>{emailMonthComparison("NDA", (month) => month.nda)}{emailMonthComparison("NDS", (month) => month.nds)}</>} info={`NDA = correos con estado atendido/finalizado / correos recibidos. NDS = (fecha de registro − date_email) ≤ ${EMAIL_SLA_MINUTES} minutos / correos recibidos. Los tiempos negativos se consideran válidos; solo se excluyen los mayores a ${EMAIL_SLA_MINUTES} minutos.`}>
              {renderEmailQualityChart()}
            </ChartPanel>
            <ChartPanel title="Evolución del TMO" meta="Desde date_email hasta fecha de registro · horario 07:00–23:01" comparison={emailMonthComparison("TMO", (month) => month.tmoMinutes)} info="Campos: date_email y fecha de registro. TMO por correo = fecha de registro − date_email, contando solo el horario operativo de 07:00 a 23:01. Los correos recibidos después de las 23:00 empiezan a contar desde las 07:00 del día siguiente. El TMO mensual es el promedio de los correos atendidos con fechas válidas.">
              {renderEmailTmoChart()}
            </ChartPanel>
            <ChartPanel title="Correos por hora de recepción" meta="Hora extraída de date_email" comparison={emailVolumeComparison} actions={<button className={styles.chartActionButton} type="button" onClick={() => setEmailHourChronological((value) => !value)} title={emailHourChronological ? "Mostrar horas por volumen" : "Ordenar de 00:00 a 23:00"}><ArrowDownAZ size={15} aria-hidden="true" />{emailHourChronological ? "Por volumen" : "Ordenar por hora"}</button>}>
              {renderEmailHourChart()}
            </ChartPanel>
            <ChartPanel title="Correos por tipificación" meta="Porcentaje · id_email" comparison={emailVolumeComparison} info="Campo: tipificacion. Cuenta id_email únicos por tipificación y calcula el porcentaje sobre el total filtrado. Se muestra en gráfico circular.">
              {renderEmailPieChart(emailSummary.byTipificacion)}
            </ChartPanel>
            <ChartPanel title="Correos por motivo" meta="Porcentaje · id_email" comparison={emailVolumeComparison} info="Campo: motivo. Cuenta id_email únicos por motivo y calcula el porcentaje sobre el total filtrado. Los registros sin motivo se excluyen del gráfico, pero permanecen en los datos originales.">
              {renderEmailPieChart(emailSummary.byMotivo)}
            </ChartPanel>
            <ChartPanel title="Correos por usuario" meta="usuario asignado · id_email" comparison={emailVolumeComparison}>
              {renderMatrixBarChart(emailSummary.byUser, "#475569")}
            </ChartPanel>
            <ChartPanel
              title="Mapa de calor I2"
              meta="Correos recibidos por día y hora"
              comparison={emailVolumeComparison}
            >
              <HeatmapGrid
                cells={emailHeatmap}
                formatter={(value) => `${value.toLocaleString("es-PE")} correos recibidos`}
              />
            </ChartPanel>
            <ChartPanel
              title="Promedio por agente"
              meta={`${matrixSummary.avgByAgent.length} agentes`}
            >
              {renderMatrixTimeChart(matrixSummary.avgByAgent)}
            </ChartPanel>
            <ChartPanel
              title="Promedio por dia"
              meta={`${matrixSummary.avgByDay.length} dias · clic para ver datos`}
              info="Agrupa por fecha de date_email y muestra los tiempos promedio de asignación y resolución. Haz clic en un día para abrir las filas crudas de ese día."
            >
              {renderMatrixTimeChart(matrixSummary.avgByDay, undefined, setMatrixDayDetail)}
            </ChartPanel>
            <ChartPanel
              title="Promedio por mes"
              meta={`${matrixSummary.avgByMonth.length} meses · asignación vs resolución`}
              info="Campos: date_email, fecha_asignacion y fecha de registro. Eje X: año-mes de date_email. Eje Y: minutos promedio. Las barras de asignación y resolución se muestran separadas."
            >
              {renderMatrixMonthlyTimeChart(matrixSummary.avgByMonth)}
            </ChartPanel>
            <ChartPanel
              title="Estado de registro"
              meta={`${matrixSummary.byEstado.length} estados`}
            >
              {renderMatrixBarChart(matrixSummary.byEstado, "#0f766e")}
            </ChartPanel>
          </div>

          <div className={styles.matrixUserDetailLayout}>
            <ChartPanel
              title="Usuario asignado"
              meta={`${matrixSummary.byUser.length} usuarios`}
            >
              {renderMatrixUserChart()}
            </ChartPanel>
            <article className={styles.panel}>
            <div className={styles.panelHeader}>
              <div>
                <h2>Detalle matriz</h2>
                <span>
                  {filteredMatrixRows.length.toLocaleString("es-PE")} de {visibleMatrixRows.length.toLocaleString("es-PE")} registros
                  {matrixUserFocus ? ` · usuario: ${matrixUserFocus}` : ""}
                  {hiddenMatrixRows ? `, mostrando ${matrixDetailRows.length.toLocaleString("es-PE")} para mantener fluidez` : ""}
                </span>
              </div>
              {matrixUserFocus ? <button className={styles.chartActionButton} type="button" onClick={() => setMatrixUserFocus(null)}>Mostrar todos los usuarios</button> : null}
            </div>
            <div className={styles.tableWrap}>
              <table>
                <thead>
                  <tr>
                    {matrixColumns.map((column) => <th key={column}>{column}</th>)}
                    <th>SLA a asignacion</th>
                    <th>SLA a resolucion</th>
                    <th>SLA total</th>
                  </tr>
                </thead>
                <tbody>
                  {matrixDetailRows.map((row) => (
                    <tr key={row.id}>
                      {matrixColumns.map((column) => (
                        <td key={`${row.id}-${column}`}>{displayCellValue(column, row.raw[column])}</td>
                      ))}
                      <td>{formatMinutes(row.minutesToAssign)}</td>
                      <td>{formatMinutes(row.minutesToRegister)}</td>
                      <td>{formatMinutes(row.minutesTotal)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            </article>
          </div>

          <article className={styles.panel}>
            <div className={styles.panelHeader}><div><h2>Tabla resumen del periodo</h2><span>Valores mensuales con id_email único · variación contra el mes anterior</span></div></div>
            <div className={styles.tableWrap}>
              <table>
                <thead><tr><th>Indicador</th>{emailSummary.byMonth.map((month) => <th key={month.name}>{month.name}</th>)}<th>Vs. mes anterior</th></tr></thead>
                <tbody>
                  <tr><td>Correos recibidos</td>{emailSummary.byMonth.map((month) => <td key={month.name}>{month.received.toLocaleString("es-PE")}</td>)}<td>{emailMonthDelta((month) => month.received)}</td></tr>
                  <tr><td>NDA</td>{emailSummary.byMonth.map((month) => <td key={month.name}>{formatPercent(month.nda)}</td>)}<td>{emailMonthDelta((month) => month.nda)}</td></tr>
                  <tr><td>NDS</td>{emailSummary.byMonth.map((month) => <td key={month.name}>{formatPercent(month.nds)}</td>)}<td>{emailMonthDelta((month) => month.nds)}</td></tr>
                  <tr><td>TMO</td>{emailSummary.byMonth.map((month) => <td key={month.name}>{formatEmailTmo(month.tmoMinutes)}</td>)}<td>{emailMonthDelta((month) => month.tmoMinutes)}</td></tr>
                </tbody>
              </table>
            </div>
          </article>
            </section>
          ) : (
            <section className={styles.emptyState} onClick={stopInsideClick}>
              <Timer size={22} aria-hidden="true" />
              <strong>I2 - Correos atendidos listo</strong>
              <span>Completa fechas y sesion para ver tiempos de asignacion y resolucion.</span>
            </section>
          )}
        </>
      ) : null}

      {activeView === "sales" ? renderSalesContent() : null}

      {activeView === "errors" ? (
        <>
          {incidentError ? <div className={styles.error} onClick={stopInsideClick}>{incidentError}</div> : null}

          <section className={styles.reportPanel} onClick={stopInsideClick}>
            <div>
              <h2>Tasa de error</h2>
              <span>Carga el Excel de incidencias para cuantificar errores por agente y por tipo.</span>
            </div>
            <div
              className={styles.inlineUpload}
              onClick={() => incidentInputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={(event) => {
                if (event.key === "Enter" || event.key === " ") {
                  incidentInputRef.current?.click();
                }
              }}
            >
              <ClipboardList size={18} aria-hidden="true" />
              <span>{incidentFileName}</span>
              <strong>Cargar Excel</strong>
              <input
                ref={incidentInputRef}
                type="file"
                accept=".xlsx,.xls,.csv,.txt,.tsv"
                onChange={(event) => handleIncidentFiles(event.target.files)}
              />
            </div>
          </section>

          {incidentRows.length ? (
            <section className={styles.matrixSection} onClick={stopInsideClick}>
              <div className={styles.matrixHeader}>
                <div>
                  <span className={styles.eyebrow}>Tasa de error</span>
                  <h2>Errores levantados</h2>
                  <p>Resumen desde la hoja BASE del registro de incidencias</p>
                </div>
                <strong>{incidentRows.length.toLocaleString("es-PE")} errores</strong>
              </div>

              <div className={styles.statsGrid}>
                <StatCard icon={<ClipboardList size={18} />} label="Errores" value={incidentSummary.total.toLocaleString("es-PE")} />
                <StatCard icon={<UserRound size={18} />} label="Agentes con error" value={incidentSummary.uniqueAgents.toLocaleString("es-PE")} />
                <StatCard icon={<Hash size={18} />} label="Tipos de error" value={incidentSummary.uniqueTypes.toLocaleString("es-PE")} />
                <StatCard icon={<CalendarDays size={18} />} label="Meses" value={incidentSummary.byMonth.length.toLocaleString("es-PE")} />
                <StatCard icon={<FileSpreadsheet size={18} />} label="Sucursales" value={incidentSummary.uniqueBranches.toLocaleString("es-PE")} />
              </div>

              <div className={styles.matrixGrid}>
                <ChartPanel title="Por agente" meta={`${incidentSummary.byAgent.length} agentes`}>
                  {renderIncidentAgentChart()}
                </ChartPanel>
                <ChartPanel title="Por tipo de inconsistencia" meta={`${incidentSummary.byType.length} tipos`}>
                  {renderMatrixBarChart(incidentSummary.byType, "#2563a8")}
                </ChartPanel>
                <ChartPanel title="Por sucursal" meta={`${incidentSummary.byBranch.length} sucursales`}>
                  {renderMatrixBarChart(incidentSummary.byBranch, "#0f766e")}
                </ChartPanel>
                <ChartPanel title="Por mes" meta={`${incidentSummary.byMonth.length} meses`}>
                  {renderMatrixBarChart(incidentSummary.byMonth, "#475569")}
                </ChartPanel>
              </div>

              <article className={styles.panel}>
                <div className={styles.panelHeader}>
                  <div>
                    <h2>Detalle de incidencias</h2>
                    <span>
                      {incidentRows.length.toLocaleString("es-PE")} registros
                      {hiddenIncidentRows ? `, mostrando ${incidentDetailRows.length.toLocaleString("es-PE")} para mantener fluidez` : ""}
                    </span>
                  </div>
                </div>
                <div className={styles.tableWrap}>
                  <table>
                    <thead>
                      <tr>
                        {incidentColumns.map((column) => (
                          <th key={column}>{column}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {incidentDetailRows.map((row) => (
                        <tr key={row.id}>
                          {incidentColumns.map((column) => (
                            <td key={`${row.id}-${column}`}>{displayCellValue(column, row.raw[column])}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </article>
            </section>
          ) : (
            <section className={styles.emptyState} onClick={stopInsideClick}>
              <ClipboardList size={22} aria-hidden="true" />
              <strong>Tasa de error lista</strong>
              <span>Carga el archivo Registro de Incidencias para ver errores por agente y tipo.</span>
            </section>
          )}
        </>
      ) : null}

      {activeView === "performance" ? (
        <section className={styles.matrixSection} onClick={stopInsideClick}>
          <div className={styles.matrixHeader}>
            <div>
              <span className={styles.eyebrow}>I4 - Rendimiento operativo</span>
              <h2>Productividad, calidad y efectividad</h2>
              <p>Calculado con llamadas inbound, correos matriz y errores levantados.</p>
            </div>
            <strong>{performanceSummary.byAgent.length.toLocaleString("es-PE")} agentes</strong>
          </div>

          <div className={styles.statsGrid}>
            <StatCard icon={<Phone size={18} />} label="Llamadas" value={performanceSummary.totalCalls.toLocaleString("es-PE")} />
            <StatCard icon={<Mail size={18} />} label="Correos" value={performanceSummary.totalEmails.toLocaleString("es-PE")} />
            <StatCard icon={<ClipboardList size={18} />} label="Errores" value={performanceSummary.totalErrors.toLocaleString("es-PE")} />
            <StatCard icon={<Hash size={18} />} label="Productividad" value={formatPercent(performanceSummary.productivity)} />
            <StatCard icon={<Timer size={18} />} label="Calidad" value={formatPercent(performanceSummary.quality)} />
            <StatCard icon={<UserRound size={18} />} label="Efectividad" value={formatPercent(performanceSummary.effectiveness)} />
          </div>

          <div className={styles.matrixGrid}>
            <ChartPanel
              title="Indicadores por agente"
              meta="Productividad, calidad y efectividad"
            >
              {renderPerformanceChart()}
            </ChartPanel>
            <ChartPanel
              title="Volumen por agente"
              meta={`${performanceSummary.totalAttended.toLocaleString("es-PE")} atenciones`}
            >
              {renderMatrixBarChart(
                performanceSummary.byAgent.map((row) => ({ name: row.agent, total: row.attended })),
                "#0f766e",
                Math.max(280, Math.min(620, performanceSummary.byAgent.length * 42)),
              )}
            </ChartPanel>
          </div>

          <article className={styles.panel}>
            <div className={styles.panelHeader}>
              <div>
                <h2>Detalle por agente</h2>
                <span>Llamadas + correos, errores, calidad y efectividad</span>
              </div>
            </div>
            <div className={styles.tableWrap}>
              <table>
                <thead>
                  <tr>
                    <th>Agente</th>
                    <th>Llamadas</th>
                    <th>Correos</th>
                    <th>Atenciones</th>
                    <th>Errores</th>
                    <th>Productividad</th>
                    <th>Calidad</th>
                    <th>Efectividad</th>
                  </tr>
                </thead>
                <tbody>
                  {performanceSummary.byAgent.map((row) => (
                    <tr key={row.agent}>
                      <td>{row.agent}</td>
                      <td>{row.calls.toLocaleString("es-PE")}</td>
                      <td>{row.emails.toLocaleString("es-PE")}</td>
                      <td>{row.attended.toLocaleString("es-PE")}</td>
                      <td>{row.errors.toLocaleString("es-PE")}</td>
                      <td>{formatPercent(row.productivity)}</td>
                      <td>{row.emails ? formatPercent(row.quality) : "Sin correos"}</td>
                      <td>{formatPercent(row.effectiveness)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </article>
        </section>
      ) : null}

      {activeView === "calls" && rows.length ? (
        <section className={styles.statusFilter} onClick={stopInsideClick}>
          <details>
            <summary>
              <span>Filtrar llamadas atendidas por Estado de llamada</span>
              <strong>{statusFilterLabel}</strong>
            </summary>
            <div className={styles.statusFilterContent}>
              <div className={styles.statusFilterActions}>
                <button type="button" onClick={() => setSelectedStatusNames(allStatus.map((item) => item.name))}>
                  Seleccionar todos
                </button>
                <button type="button" onClick={() => setSelectedStatusNames([])}>
                  Mostrar todos
                </button>
              </div>
              <div className={styles.statusFilterOptions}>
                {allStatus.map((item) => (
                  <label key={item.name}>
                    <input
                      type="checkbox"
                      checked={!selectedStatusNames.length || selectedStatusNames.includes(item.name)}
                      onChange={() => toggleStatusName(item.name)}
                    />
                    <span>{item.name}</span>
                    <em>{item.total.toLocaleString("es-PE")}</em>
                  </label>
                ))}
              </div>
            </div>
          </details>
        </section>
      ) : null}

      {activeView === "calls" && rows.length ? (
        <>
      <section className={styles.statsGrid} onClick={stopInsideClick}>
        <StatCard icon={<Hash size={18} />} label="Llamadas recibidas" value={currentSummary.total.toLocaleString("es-PE")} comparison={callMonthComparison((month) => month.received)} />
        <StatCard icon={<Phone size={18} />} label="Llamadas atendidas" value={currentSummary.attendedCalls.toLocaleString("es-PE")} comparison={callMonthComparison((month) => month.attended)} />
        <StatCard icon={<UserRound size={18} />} label="Nivel de atención" value={formatPercent(percentage(currentSummary.attendedCalls, currentSummary.total))} comparison={callMonthComparison((month) => percentage(month.attended, month.received))} />
        <StatCard icon={<Phone size={18} />} label="Abandono" value={formatPercent(percentage(Math.max(0, currentSummary.total - currentSummary.attendedCalls), currentSummary.total))} comparison={callMonthComparison((month) => percentage(Math.max(0, month.received - month.attended), month.received))} />
        <StatCard icon={<Timer size={18} />} label="TMO promedio" value={formatCallDuration(serviceSummary.averageDurationSeconds)} comparison={callMonthComparison((month) => "averageDurationSeconds" in month ? month.averageDurationSeconds : null)} />
      </section>

      <section className={styles.dashboardGrid}>
        <ChartPanel
          title="Llamadas por mes"
          meta={`${monthlyCallData.length} meses`}
          comparison={callComparisonBadge("volumen", (month) => month.received)}
          onExpand={() => setExpandedChart("date")}
          onExport={() => void exportChart("date", "Llamadas por mes")}
        >
          <div ref={dateChartRef} className={styles.exportFrame}>
            {renderDateChart()}
          </div>
        </ChartPanel>

        <ChartPanel
          title="Llamadas recibidas por hora"
          meta={`${byHour.length} horas`}
          comparison={callComparisonBadge("volumen", (month) => month.received)}
          onExpand={() => setExpandedChart("hour")}
          onExport={() => void exportChart("hour", "Llamadas recibidas por hora")}
        >
          <div ref={hourChartRef} className={styles.exportFrame}>
            {renderHourChart()}
          </div>
        </ChartPanel>

        <ChartPanel
          title="Categorías"
          meta={`${statusChartData.length} categorías`}
          comparison={callComparisonBadge("volumen", (month) => month.received)}
          onExpand={() => setExpandedChart("status")}
          onExport={() => void exportChart("status", "Categorías")}
          info="Campo: status_name. Cuenta las llamadas por categoría. Al seleccionar una categoría, las demás permanecen visibles y atenuadas; el filtro se combina con los demás dashboards."
        >
          <div ref={statusChartRef} className={styles.exportFrame}>{renderStatusChart()}</div>
        </ChartPanel>

        <ChartPanel
          title="Campañas"
          meta={`${byCampaign.length} campañas`}
          comparison={callComparisonBadge("volumen", (month) => month.received)}
          onExpand={() => setExpandedChart("campaign")}
          onExport={() => void exportChart("campaign", "Campañas")}
        >
          <div ref={campaignChartRef} className={styles.exportFrame}>
            {renderCampaignChart()}
          </div>
        </ChartPanel>
      </section>

      <section className={`${styles.dashboardGrid} ${styles.fullWidthDashboard}`} onClick={stopInsideClick}>
        <ChartPanel title="Mapa de calor I1" meta="Llamadas por día y hora · 07:00 a 23:00" comparison={callComparisonBadge("volumen", (month) => month.received)}>
          <HeatmapGrid cells={callHeatmap} formatter={(value) => `${value.toLocaleString("es-PE")} llamadas`} />
        </ChartPanel>
      </section>

      <section className={styles.serviceSection} onClick={stopInsideClick}>
        <div className={styles.dashboardGrid}>
          <ChartPanel title="Nivel de atención por mes" meta="NA: atendidas / recibidas · Meta ≥ 90%" comparison={callComparisonBadge("NA", (month) => percentage(month.attended, month.received))} info="Campos: length_in_sec y call_date. Llamada atendida = length_in_sec > 0. Nivel de Atención (NA) = llamadas atendidas / llamadas recibidas × 100. Meta operativa: NA ≥ 90%.">
            {renderServiceQualityChart()}
          </ChartPanel>
          <ChartPanel title="TMO por mes" meta="Tiempo medio de operación" comparison={callComparisonBadge("TMO", (month) => "averageDurationSeconds" in month ? month.averageDurationSeconds : null)}>
            {renderServiceDurationChart()}
          </ChartPanel>
          <ChartPanel title="Abandono por mes" meta="Recibidas menos atendidas" comparison={callComparisonBadge("abandono", (month) => percentage(Math.max(0, month.received - month.attended), month.received))}>
            {renderAbandonmentChart()}
          </ChartPanel>
          <ChartPanel title="Distribución de duración" meta="Segundos por llamada" comparison={callComparisonBadge("volumen", (month) => month.received)} info="Campo: length_in_sec. Agrupa las llamadas en rangos de duración para identificar concentración operativa.">
            {renderDurationDistribution()}
          </ChartPanel>
          <ChartPanel title="Rendimiento por usuario" meta="Recibidas y atendidas" comparison={callComparisonBadge("volumen", (month) => month.received)} info="Campos: user y length_in_sec. Compara llamadas recibidas contra atendidas por usuario.">
            {renderUserPerformance()}
          </ChartPanel>
          <ChartPanel title="Datos crudos filtrados" meta={`${visibleRows.length.toLocaleString("es-PE")} filas`} info="Muestra las filas originales que alimentan los gráficos, respetando todos los filtros activos.">
            {renderRawCallRows()}
          </ChartPanel>
        </div>
      </section>
        </>
      ) : null}

      {activeView === "calls" && rows.length && expandedChart ? (
        <div className={styles.modalBackdrop} onClick={() => setExpandedChart(null)} role="presentation">
          <section className={styles.modal} onClick={stopInsideClick} role="dialog" aria-modal="true" aria-label="Grafico ampliado">
            <div className={styles.modalHeader}>
              <div>
                <h2>
                  {expandedChart === "date" ? "Llamadas por mes" : null}
                  {expandedChart === "hour" ? "Llamadas recibidas por hora" : null}
                  {expandedChart === "status" ? "Categorías" : null}
                  {expandedChart === "campaign" ? "Campañas" : null}
                </h2>
                <span>Haz clic en una parte del grafico para filtrar toda la vista</span>
              </div>
              <button className={styles.iconButton} type="button" onClick={() => setExpandedChart(null)} aria-label="Cerrar modal">
                <X size={18} aria-hidden="true" />
              </button>
            </div>
            {expandedContent()}
          </section>
        </div>
      ) : null}

      {activeView === "matrix" && matrixDayDetail ? (
        <div className={styles.modalBackdrop} onClick={() => setMatrixDayDetail(null)} role="presentation">
          <section className={styles.modal} onClick={stopInsideClick} role="dialog" aria-modal="true" aria-label={`Datos crudos del ${matrixDayDetail}`}>
            <div className={styles.modalHeader}>
              <div>
                <h2>Datos crudos · {matrixDayDetail}</h2>
                <span>{matrixDayDetailRows.length.toLocaleString("es-PE")} correos filtrados por date_email</span>
              </div>
              <button className={styles.iconButton} type="button" onClick={() => setMatrixDayDetail(null)} aria-label="Cerrar datos del día">
                <X size={18} aria-hidden="true" />
              </button>
            </div>
            <div className={styles.rawDataScroll}>
              <table className={styles.rawDataTable}>
                <thead><tr>{(matrixColumns.length ? matrixColumns : Object.keys(matrixDayDetailRows[0]?.raw ?? {})).map((column) => <th key={column}>{column}</th>)}</tr></thead>
                <tbody>{matrixDayDetailRows.map((row) => <tr key={row.id}>{(matrixColumns.length ? matrixColumns : Object.keys(row.raw)).map((column) => <td key={`${row.id}-${column}`}>{row.raw[column] ?? ""}</td>)}</tr>)}</tbody>
              </table>
            </div>
          </section>
        </div>
      ) : null}
    </main>
  );
}
