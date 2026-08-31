"use client";

import { useMemo, useRef, useState, useTransition } from "react";
import type { FormEvent, MouseEvent as ReactMouseEvent, ReactNode, RefObject } from "react";
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
  FileSpreadsheet,
  Hash,
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
  raw: Record<string, string>;
};

type ChartPoint = {
  name: string;
  total: number;
};

type FilterField = "date" | "week" | "month" | "hour" | "statusName" | "campaignId" | "user";
type ChartKind = "date" | "hour" | "status" | "campaign";
type DateMode = "week" | "month" | "total";
type DashboardView = "calls" | "matrix" | "errors" | "performance";
type TimelineScale = "30m" | "1h" | "day";

type ActiveFilter = {
  field: FilterField;
  label: string;
  value: string;
} | null;

type ChartClickState = {
  activeLabel?: string | number;
  activePayload?: Array<{ payload?: ChartPoint; name?: string | number; value?: string | number }>;
};

type FilterIndex = Record<FilterField, Map<string, MetricRow[]>>;

type DashboardSummary = {
  total: number;
  byDate: ChartPoint[];
  byWeek: ChartPoint[];
  byMonth: ChartPoint[];
  byHour: ChartPoint[];
  byStatus: ChartPoint[];
  byCampaign: ChartPoint[];
  byUser: ChartPoint[];
  uniquePhones: number;
  uniqueUsers: number;
  uniqueCampaigns: number;
  uniqueStatuses: number;
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
const filterFields: FilterField[] = ["date", "week", "month", "hour", "statusName", "campaignId", "user"];
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
  "#2563eb",
  "#10b981",
  "#f59e0b",
  "#ef4444",
  "#7c3aed",
  "#0f766e",
  "#db2777",
  "#475569",
  "#14b8a6",
  "#ea580c",
  "#4f46e5",
  "#16a34a",
];

const demoRows: MetricRow[] = [
  { id: 1, date: "2026-08-24", week: "2026-S35", month: "2026-08", hour: "09:00", phone: "+51 987 123 456", user: "Maria", campaignId: "CMP-101", statusName: "Contactado", raw: { call_date: "2026-08-24T09:12:00", phone_number_dialed: "+51 987 123 456", user: "Maria", campaign_id: "CMP-101", status_name: "Contactado" } },
  { id: 2, date: "2026-08-24", week: "2026-S35", month: "2026-08", hour: "10:00", phone: "+51 956 443 201", user: "Luis", campaignId: "CMP-101", statusName: "No contesta", raw: { call_date: "2026-08-24T10:20:00", phone_number_dialed: "+51 956 443 201", user: "Luis", campaign_id: "CMP-101", status_name: "No contesta" } },
  { id: 3, date: "2026-08-25", week: "2026-S35", month: "2026-08", hour: "11:00", phone: "+51 999 800 120", user: "Ana", campaignId: "CMP-204", statusName: "Venta", raw: { call_date: "2026-08-25T11:03:00", phone_number_dialed: "+51 999 800 120", user: "Ana", campaign_id: "CMP-204", status_name: "Venta" } },
  { id: 4, date: "2026-08-25", week: "2026-S35", month: "2026-08", hour: "15:00", phone: "+51 944 320 111", user: "Maria", campaignId: "CMP-204", statusName: "Contactado", raw: { call_date: "2026-08-25T15:42:00", phone_number_dialed: "+51 944 320 111", user: "Maria", campaign_id: "CMP-204", status_name: "Contactado" } },
  { id: 5, date: "2026-08-26", week: "2026-S35", month: "2026-08", hour: "16:00", phone: "+51 933 740 812", user: "Diego", campaignId: "CMP-101", statusName: "Pendiente", raw: { call_date: "2026-08-26T16:18:00", phone_number_dialed: "+51 933 740 812", user: "Diego", campaign_id: "CMP-101", status_name: "Pendiente" } },
  { id: 6, date: "2026-08-27", week: "2026-S35", month: "2026-08", hour: "09:00", phone: "+51 970 662 114", user: "Ana", campaignId: "CMP-330", statusName: "Venta", raw: { call_date: "2026-08-27T09:55:00", phone_number_dialed: "+51 970 662 114", user: "Ana", campaign_id: "CMP-330", status_name: "Venta" } },
];

const demoColumns = Object.keys(demoRows[0].raw);

const columnAliases = {
  date: ["date", "fecha", "created_at", "createdat", "call_date", "call date", "fecha_creacion", "created"],
  hour: ["hour", "hora", "time", "call_time", "call time", "created_time", "created time"],
  phone: ["phone", "telefono", "teléfono", "mobile", "celular", "number", "phone_number", "phone number", "numero", "número"],
  user: ["user", "usuario", "agent", "asesor", "owner", "ejecutivo", "username", "user_name", "user name"],
  campaignId: ["campaign_id", "campaign id", "campaign", "campaña", "campana", "id_campana", "id campaña"],
  statusName: ["status_name", "status name", "estado_nombre", "nombre_estado"],
};

const hiddenDetailColumns = new Set(
  [
    "status",
    "user",
    "vendor_lead_code",
    "source_id",
    "list_id",
    "gmt_offset_now",
    "phone_code",
    "phone_number",
    "title",
    "first_name",
    "middle_initial",
    "last_name",
    "address1",
    "address2",
    "address3",
    "city",
    "state",
    "province",
    "postal_code",
    "country_code",
    "gender",
    "date_of_birth",
    "alt_phone",
    "email",
    "security_phrase",
    "comments",
    "length_in_sec",
    "user_group",
    "alt_dial",
    "rank",
    "owner",
    "lead_id",
    "list_name",
    "list_description",
  ].map(normalizeKey),
);

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

function visibleDetailColumns(columns: string[]) {
  const visibleColumns = columns.filter((column) => !hiddenDetailColumns.has(normalizeKey(column)));
  return visibleColumns.length ? visibleColumns : columns;
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
    byUser: countMatrixBy(rows, "usuarioAsignado"),
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

function createFilterIndex() {
  return filterFields.reduce<FilterIndex>((index, field) => {
    index[field] = new Map();
    return index;
  }, {} as FilterIndex);
}

function buildFilterIndex(rows: MetricRow[]) {
  const index = createFilterIndex();

  rows.forEach((row) => {
    filterFields.forEach((field) => {
      const key = row[field] || "Sin dato";
      const rowsForKey = index[field].get(key);

      if (rowsForKey) {
        rowsForKey.push(row);
      } else {
        index[field].set(key, [row]);
      }
    });
  });

  return index;
}

function rowsForFilter(rows: MetricRow[], index: FilterIndex, activeFilter: ActiveFilter) {
  if (!activeFilter) {
    return rows;
  }

  return index[activeFilter.field].get(activeFilter.value) ?? [];
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

function buildDashboardSummary(rows: MetricRow[]): DashboardSummary {
  const counts = {
    date: new Map<string, number>(),
    week: new Map<string, number>(),
    month: new Map<string, number>(),
    hour: new Map<string, number>(),
    statusName: new Map<string, number>(),
    campaignId: new Map<string, number>(),
    user: new Map<string, number>(),
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
    byCampaign: pointsFromMap(counts.campaignId),
    byUser: pointsFromMap(counts.user),
    uniquePhones: unique.phone.size,
    uniqueUsers: unique.user.size,
    uniqueCampaigns: unique.campaignId.size,
    uniqueStatuses: unique.statusName.size,
  };
}

function StatCard({
  icon,
  label,
  value,
}: {
  icon: ReactNode;
  label: string;
  value: string | number;
}) {
  return (
    <div className={styles.statCard}>
      <div className={styles.statIcon}>{icon}</div>
      <span>{label}</span>
      <strong>{value}</strong>
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

function ChartPanel({
  title,
  meta,
  children,
  onExpand,
  onExport,
}: {
  title: string;
  meta: string;
  children: ReactNode;
  onExpand?: () => void;
  onExport?: () => void;
}) {
  return (
    <article className={styles.panel} onClick={stopInsideClick}>
      <div className={styles.panelHeader}>
        <div>
          <h2>{title}</h2>
          <span>{meta}</span>
        </div>
        {onExport || onExpand ? (
          <div className={styles.headerActions}>
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
  const dateChartRef = useRef<HTMLDivElement>(null);
  const hourChartRef = useRef<HTMLDivElement>(null);
  const statusChartRef = useRef<HTMLDivElement>(null);
  const campaignChartRef = useRef<HTMLDivElement>(null);
  const [rows, setRows] = useState<MetricRow[]>(demoRows);
  const [columns, setColumns] = useState<string[]>(demoColumns);
  const [fileName, setFileName] = useState("Datos demo");
  const [error, setError] = useState("");
  const [isDragging, setIsDragging] = useState(false);
  const [activeFilter, setActiveFilter] = useState<ActiveFilter>(null);
  const [expandedChart, setExpandedChart] = useState<ChartKind | null>(null);
  const [dateMode, setDateMode] = useState<DateMode>("week");
  const [activeView, setActiveView] = useState<DashboardView>("calls");
  const [reportStartDate, setReportStartDate] = useState(todayInputValue);
  const [reportEndDate, setReportEndDate] = useState(todayInputValue);
  const [reportSessionId, setReportSessionId] = useState("");
  const [reportError, setReportError] = useState("");
  const [isDownloadingReport, setIsDownloadingReport] = useState(false);
  const [matrixRows, setMatrixRows] = useState<MatrixRow[]>([]);
  const [matrixColumns, setMatrixColumns] = useState<string[]>([]);
  const [matrixRangeLabel, setMatrixRangeLabel] = useState("");
  const [timelineDay, setTimelineDay] = useState("");
  const [timelineScale, setTimelineScale] = useState<TimelineScale>("day");
  const [showSlowResolutionOnly, setShowSlowResolutionOnly] = useState(false);
  const [hoveredTimelineId, setHoveredTimelineId] = useState<string | null>(null);
  const [incidentRows, setIncidentRows] = useState<IncidentRow[]>([]);
  const [incidentColumns, setIncidentColumns] = useState<string[]>([]);
  const [incidentFileName, setIncidentFileName] = useState("Sin archivo cargado");
  const [incidentError, setIncidentError] = useState("");
  const [isFiltering, startFilterTransition] = useTransition();

  const filterIndex = useMemo(() => buildFilterIndex(rows), [rows]);
  const fullSummary = useMemo(() => buildDashboardSummary(rows), [rows]);
  const detailColumns = useMemo(() => visibleDetailColumns(columns), [columns]);
  const visibleRows = useMemo(() => rowsForFilter(rows, filterIndex, activeFilter), [activeFilter, filterIndex, rows]);
  const currentSummary = useMemo(
    () => (activeFilter ? buildDashboardSummary(visibleRows) : fullSummary),
    [activeFilter, fullSummary, visibleRows],
  );
  const byDate = dateMode === "week" ? currentSummary.byWeek : dateMode === "month" ? currentSummary.byMonth : currentSummary.byDate;
  const byHour = currentSummary.byHour;
  const allStatus = fullSummary.byStatus;
  const byStatus = currentSummary.byStatus;
  const byCampaign = currentSummary.byCampaign;
  const byUser = currentSummary.byUser;
  const detailStatuses = byStatus.slice(0, 4);
  const detailCampaigns = byCampaign.slice(0, 4);
  const detailUsers = byUser.slice(0, 4);
  const detailRows = useMemo(() => visibleRows.slice(0, detailRowLimit), [visibleRows]);
  const hiddenDetailRows = Math.max(0, visibleRows.length - detailRows.length);
  const slowResolutionRows = useMemo(
    () => matrixRows.filter((row) => row.minutesToRegister !== null && row.minutesToRegister > 20),
    [matrixRows],
  );
  const visibleMatrixRows = showSlowResolutionOnly ? slowResolutionRows : matrixRows;
  const matrixSummary = useMemo(() => buildMatrixSummary(visibleMatrixRows), [visibleMatrixRows]);
  const matrixDetailRows = useMemo(() => visibleMatrixRows.slice(0, detailRowLimit), [visibleMatrixRows]);
  const hiddenMatrixRows = Math.max(0, visibleMatrixRows.length - matrixDetailRows.length);
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
  const userChartMax = byUser[0]?.total ?? 1;
  const statusTotal = currentSummary.total;
  const statusColorFor = (name: string) => chartColor(Math.max(0, allStatus.findIndex((item) => item.name === name)));

  const chartRefs: Record<ChartKind, RefObject<HTMLDivElement | null>> = {
    date: dateChartRef,
    hour: hourChartRef,
    status: statusChartRef,
    campaign: campaignChartRef,
  };

  function clearFilters() {
    startFilterTransition(() => {
      setActiveFilter(null);
    });
  }

  function setFilter(field: FilterField, label: string, value: string) {
    startFilterTransition(() => {
      setActiveFilter((current) => {
        if (current?.field === field && current.value === value) {
          return null;
        }

        return { field, label, value };
      });
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
    return activeFilter?.field === field && activeFilter.value === point.name;
  }

  function pointLabel(field: FilterField, point: ChartPoint) {
    if (pointIsSelected(field, point)) {
      return "En foco";
    }

    return String(point.total);
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

      const mappedRows = mapRows(rawRows);
      if (!mappedRows.length) {
        throw new Error("No encontre filas validas en el archivo.");
      }

      setRows(mappedRows);
      setColumns(Object.keys(rawRows[0] ?? {}));
      setActiveFilter(null);
      setDateMode("week");
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
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
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

  function renderMatrixTimeChart(data: TimeAveragePoint[], height = Math.max(280, Math.min(720, data.length * 42))) {
    return data.length ? (
      <div className={styles.scrollChart}>
        <ResponsiveContainer width="100%" height={height}>
          <BarChart
            data={data}
            layout="vertical"
            margin={{ left: 16, right: 20, top: 8, bottom: 0 }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip
              formatter={(value, name) => [
                typeof value === "number" ? formatMinutes(value) : value,
                name === "asignacion" ? "Asignacion" : name === "resolucion" ? "Resolucion" : name,
              ]}
            />
            <Legend />
            <Bar dataKey="asignacion" name="Asignacion" fill="#2563eb" radius={[0, 4, 4, 0]} />
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
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip />
            <Bar dataKey="total" name="Registros" fill="#7c3aed" radius={[0, 4, 4, 0]} />
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
            <Bar dataKey="total" name="Observaciones" fill="#ec4899" radius={[0, 4, 4, 0]} />
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
            <Bar dataKey="productivity" name="Productividad" fill="#ec4899" radius={[0, 4, 4, 0]} />
            <Bar dataKey="quality" name="Calidad" fill="#0f766e" radius={[0, 4, 4, 0]} />
            <Bar dataKey="effectiveness" name="Efectividad" fill="#f59e0b" radius={[0, 4, 4, 0]} />
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
            <h2>Timeline por agente</h2>
            <span>Ocupacion por correo desde asignacion hasta resolucion</span>
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
                        >
                          <span>{block.range}</span>
                        </button>
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

  async function loadMatrixReport(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setReportError("");
    setIsDownloadingReport(true);

    try {
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

      const blob = await response.blob();
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
      setMatrixRangeLabel(`${reportStartDate} a ${reportEndDate}`);
      setTimelineDay(timelineDaysFromRows(mappedRows)[0] ?? "");
      setShowSlowResolutionOnly(false);
      setHoveredTimelineId(null);
    } catch (currentError) {
      setReportError(currentError instanceof Error ? currentError.message : "No pude leer el reporte.");
    } finally {
      setIsDownloadingReport(false);
    }
  }

  function renderDateChart(height = 280) {
    const dateFilterField = dateMode === "week" ? "week" : dateMode === "month" ? "month" : "date";
    const dateFilterLabel = dateMode === "week" ? "Semana" : dateMode === "month" ? "Mes" : "Fecha";

    return byDate.length ? (
      <div className={styles.chartStack}>
        <div className={styles.segmentedControl}>
          {(["week", "month", "total"] as DateMode[]).map((mode) => (
            <button
              key={mode}
              type="button"
              className={dateMode === mode ? styles.segmentedActive : ""}
              onClick={() => setDateMode(mode)}
            >
              {mode === "week" ? "Semana" : mode === "month" ? "Mes" : "Total"}
            </button>
          ))}
        </div>
        <div className={styles.chartShell}>
        <ResponsiveContainer width="100%" height={height}>
          <LineChart
            data={byDate}
            margin={{ left: 0, right: 16, top: 8, bottom: 0 }}
            onClick={(state) => {
              const value = chartClickValue(state as ChartClickState);
              if (value) {
                setFilter(dateFilterField, dateFilterLabel, value);
              }
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Line
              type="monotone"
              dataKey="total"
              name="Registros"
              stroke="#2563eb"
              strokeWidth={3}
              dot={{ r: 4 }}
            />
          </LineChart>
        </ResponsiveContainer>
        </div>
      </div>
    ) : (
      <EmptyChart>Sin fechas para graficar</EmptyChart>
    );
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
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
            <XAxis dataKey="name" tick={{ fontSize: 12 }} />
            <YAxis allowDecimals={false} tick={{ fontSize: 12 }} />
            <Tooltip />
            <Bar
              dataKey="total"
              name="Registros"
              fill="#10b981"
              radius={[4, 4, 0, 0]}
              onClick={(entry) => {
                const value = chartPointName(entry);
                if (value) {
                  setFilter("hour", "Hora", value);
                }
              }}
            />
          </BarChart>
        </ResponsiveContainer>
      </div>
    );
  }

  function renderStatusChart(height = 280) {
    return (
      <div className={styles.statusChartLayout}>
        <div className={styles.pieFrame}>
          <ResponsiveContainer width="100%" height={height}>
            <PieChart>
              <Pie
                data={byStatus}
                dataKey="total"
                nameKey="name"
                innerRadius={height > 300 ? 82 : 56}
                outerRadius={height > 300 ? 150 : 98}
                paddingAngle={2}
                onClick={(entry) => setFilter("statusName", "Status name", String((entry as unknown as ChartPoint).name))}
              >
                {byStatus.map((entry) => (
                  <Cell key={entry.name} fill={statusColorFor(entry.name)} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
          <div className={styles.pieCenter}>
            <span>Total</span>
            <strong>{statusTotal.toLocaleString("es-PE")}</strong>
          </div>
        </div>
        <div className={styles.statusLegend}>
          {byStatus.map((item) => (
            <button
              key={item.name}
              type="button"
              className={`${styles.legendItem} ${pointIsSelected("statusName", item) ? styles.focusedItem : ""}`}
              onClick={() => setFilter("statusName", "Status name", item.name)}
              title={`${item.name}: ${item.total} registros`}
            >
              <span style={{ backgroundColor: statusColorFor(item.name) }} />
              <strong>{item.name}</strong>
              <em>{pointLabel("statusName", item)}</em>
            </button>
          ))}
        </div>
      </div>
    );
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
                setFilter("campaignId", "Campaign ID", value);
              }
            }}
          >
            <CartesianGrid strokeDasharray="3 3" stroke="#dbe4f0" />
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12 }} />
            <YAxis type="category" dataKey="name" tick={{ fontSize: 12 }} width={128} />
            <Tooltip />
            <Bar
              dataKey="total"
              name="Registros"
              fill="#f59e0b"
              radius={[0, 4, 4, 0]}
              onClick={(entry) => {
                const value = chartPointName(entry);
                if (value) {
                  setFilter("campaignId", "Campaign ID", value);
                }
              }}
            />
          </BarChart>
        </ResponsiveContainer>
      </div>
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

  return (
    <main className={styles.page} onClick={clearFilters}>
      <section className={styles.header} onClick={stopInsideClick}>
        <div>
          <span className={styles.eyebrow}>A365 indicadores</span>
          <h1>Dashboard de campañas</h1>
          <p>
            Arrastra un TXT, CSV o Excel y convierte fechas, horas, usuarios,
            telefonos, campaign_id y status_name en indicadores listos para revisar.
          </p>
        </div>

        <div
          className={`${styles.dropzone} ${isDragging ? styles.dropzoneActive : ""}`}
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
          <Upload size={24} aria-hidden="true" />
          <strong>Arrastra tu archivo aqui</strong>
          <span>Tambien puedes hacer clic para seleccionarlo</span>
          <input
            ref={inputRef}
            type="file"
            accept=".txt,.csv,.tsv,.xlsx,.xls"
            onChange={(event) => handleFiles(event.target.files)}
          />
        </div>
      </section>

      <nav className={styles.topNav} onClick={stopInsideClick} aria-label="Navegacion de indicadores">
        <button
          type="button"
          className={activeView === "calls" ? styles.topNavActive : ""}
          onClick={() => setActiveView("calls")}
        >
          Indicadores 1
        </button>
        <button
          type="button"
          className={activeView === "matrix" ? styles.topNavActive : ""}
          onClick={() => setActiveView("matrix")}
        >
          Indicadores 2
        </button>
        <button
          type="button"
          className={activeView === "errors" ? styles.topNavActive : ""}
          onClick={() => setActiveView("errors")}
        >
          Tasa de error
        </button>
        <button
          type="button"
          className={activeView === "performance" ? styles.topNavActive : ""}
          onClick={() => setActiveView("performance")}
        >
          Indicadores 4
        </button>
      </nav>

      {activeView === "calls" ? (
        <>
          {error ? <div className={styles.error} onClick={stopInsideClick}>{error}</div> : null}

          <section className={styles.fileBar} onClick={stopInsideClick}>
            <FileSpreadsheet size={18} aria-hidden="true" />
            <span>{fileName}</span>
            <strong>
              {visibleRows.length.toLocaleString("es-PE")} de {rows.length.toLocaleString("es-PE")} registros
            </strong>
          </section>
        </>
      ) : null}

      {activeView === "matrix" ? (
        <>
          <section className={styles.reportPanel} onClick={stopInsideClick}>
            <div>
              <h2>Matriz A365</h2>
              <span>Ingresa tu PHPSESSID y el rango de fechas para leer el reporte en indicadores.</span>
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
                  placeholder="Pega tu sesion"
                  autoComplete="off"
                  required
                />
              </label>
              <button type="submit" disabled={isDownloadingReport}>
                <Download size={16} aria-hidden="true" />
                {isDownloadingReport ? "Analizando" : "Leer matriz"}
              </button>
            </form>
            {reportError ? <strong>{reportError}</strong> : null}
          </section>

          {matrixRows.length ? (
            <section className={styles.matrixSection} onClick={stopInsideClick}>
          <div className={styles.matrixHeader}>
            <div>
              <span className={styles.eyebrow}>Reporte matriz</span>
              <h2>Tiempos SLA de atencion</h2>
              <p>{matrixRangeLabel} · horario muerto 23:01 a 06:59</p>
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
            <StatCard icon={<Mail size={18} />} label="Emails" value={matrixSummary.total.toLocaleString("es-PE")} />
            <StatCard icon={<Timer size={18} />} label="SLA a asignacion" value={formatMinutes(matrixSummary.avgToAssign)} />
            <StatCard icon={<Clock3 size={18} />} label="SLA a resolucion" value={formatMinutes(matrixSummary.avgToRegister)} />
            <StatCard icon={<Timer size={18} />} label="SLA total" value={formatMinutes(matrixSummary.avgTotal)} />
            <StatCard icon={<Hash size={18} />} label="Terminados" value={matrixSummary.completed.toLocaleString("es-PE")} />
          </div>

          {renderAgentTimeline()}

          <div className={styles.matrixGrid}>
            <ChartPanel
              title="Promedio por agente"
              meta={`${matrixSummary.avgByAgent.length} agentes`}
            >
              {renderMatrixTimeChart(matrixSummary.avgByAgent)}
            </ChartPanel>
            <ChartPanel
              title="Promedio por dia"
              meta={`${matrixSummary.avgByDay.length} dias`}
            >
              {renderMatrixTimeChart(matrixSummary.avgByDay)}
            </ChartPanel>
            <ChartPanel
              title="Promedio por mes"
              meta={`${matrixSummary.avgByMonth.length} meses`}
            >
              {renderMatrixTimeChart(matrixSummary.avgByMonth, Math.max(280, Math.min(520, matrixSummary.avgByMonth.length * 54)))}
            </ChartPanel>
            <ChartPanel
              title="Tipificacion"
              meta={`${matrixSummary.byTipificacion.length} tipos`}
            >
              {renderMatrixBarChart(matrixSummary.byTipificacion, "#2563eb")}
            </ChartPanel>
            <ChartPanel
              title="Estado de registro"
              meta={`${matrixSummary.byEstado.length} estados`}
            >
              {renderMatrixBarChart(matrixSummary.byEstado, "#0f766e")}
            </ChartPanel>
            <ChartPanel
              title="Usuario asignado"
              meta={`${matrixSummary.byUser.length} usuarios`}
            >
              {renderMatrixUserChart()}
            </ChartPanel>
          </div>

          <article className={styles.panel}>
            <div className={styles.panelHeader}>
              <div>
                <h2>Detalle matriz</h2>
                <span>
                  {visibleMatrixRows.length.toLocaleString("es-PE")} registros
                  {hiddenMatrixRows ? `, mostrando ${matrixDetailRows.length.toLocaleString("es-PE")} para mantener fluidez` : ""}
                </span>
              </div>
            </div>
            <div className={styles.tableWrap}>
              <table>
                <thead>
                  <tr>
                    {matrixColumns.map((column) => (
                      <th key={column}>{column}</th>
                    ))}
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
            </section>
          ) : (
            <section className={styles.emptyState} onClick={stopInsideClick}>
              <Timer size={22} aria-hidden="true" />
              <strong>Indicadores 2 listos para leer la matriz</strong>
              <span>Completa fechas y sesion para ver tiempos de asignacion y resolucion.</span>
            </section>
          )}
        </>
      ) : null}

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
                  {renderMatrixBarChart(incidentSummary.byType, "#ec4899")}
                </ChartPanel>
                <ChartPanel title="Por sucursal" meta={`${incidentSummary.byBranch.length} sucursales`}>
                  {renderMatrixBarChart(incidentSummary.byBranch, "#db2777")}
                </ChartPanel>
                <ChartPanel title="Por mes" meta={`${incidentSummary.byMonth.length} meses`}>
                  {renderMatrixBarChart(incidentSummary.byMonth, "#be185d")}
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
              <span className={styles.eyebrow}>Indicadores 4</span>
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
                "#db2777",
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

      {activeView === "calls" && activeFilter ? (
        <section className={styles.filterBar} onClick={stopInsideClick}>
          <span>
            Vista filtrada por {activeFilter.label}: <strong>{activeFilter.value}</strong>
            {isFiltering ? <em>Actualizando vista...</em> : null}
          </span>
        </section>
      ) : null}

      {activeView === "calls" ? (
        <>
      <section className={styles.statsGrid} onClick={stopInsideClick}>
        <StatCard icon={<Hash size={18} />} label="Registros" value={currentSummary.total.toLocaleString("es-PE")} />
        <StatCard icon={<Phone size={18} />} label="Telefonos" value={currentSummary.uniquePhones.toLocaleString("es-PE")} />
        <StatCard icon={<UserRound size={18} />} label="Usuarios" value={currentSummary.uniqueUsers.toLocaleString("es-PE")} />
        <StatCard icon={<CalendarDays size={18} />} label="Campañas" value={currentSummary.uniqueCampaigns.toLocaleString("es-PE")} />
        <StatCard icon={<Clock3 size={18} />} label="Estados" value={currentSummary.uniqueStatuses.toLocaleString("es-PE")} />
      </section>

      <section className={styles.dashboardGrid}>
        <ChartPanel
          title="Registros por fecha"
          meta={`${byDate.length} fechas`}
          onExpand={() => setExpandedChart("date")}
          onExport={() => void exportChart("date", "Registros por fecha")}
        >
          <div ref={dateChartRef} className={styles.exportFrame}>
            {renderDateChart()}
          </div>
        </ChartPanel>

        <ChartPanel
          title="Registros por hora"
          meta={`${byHour.length} horas`}
          onExpand={() => setExpandedChart("hour")}
          onExport={() => void exportChart("hour", "Registros por hora")}
        >
          <div ref={hourChartRef} className={styles.exportFrame}>
            {renderHourChart()}
          </div>
        </ChartPanel>

        <ChartPanel
          title="Status name"
          meta={`${byStatus.length} tipos`}
          onExpand={() => setExpandedChart("status")}
          onExport={() => void exportChart("status", "Status name")}
        >
          <div ref={statusChartRef} className={styles.exportFrame}>
            {renderStatusChart()}
          </div>
        </ChartPanel>

        <ChartPanel
          title="Campaign ID"
          meta={`${byCampaign.length} campañas`}
          onExpand={() => setExpandedChart("campaign")}
          onExport={() => void exportChart("campaign", "Campaign ID")}
        >
          <div ref={campaignChartRef} className={styles.exportFrame}>
            {renderCampaignChart()}
          </div>
        </ChartPanel>
      </section>

      <section className={styles.bottomGrid} onClick={stopInsideClick}>
        <article className={styles.panel}>
          <div className={styles.panelHeader}>
            <div>
              <h2>Usuarios</h2>
              <span>{byUser.length} usuarios</span>
            </div>
          </div>
          <div className={styles.rankList}>
            {byUser.map((item, index) => (
              <button
                className={`${styles.rankItem} ${pointIsSelected("user", item) ? styles.rankItemActive : ""}`}
                key={item.name}
                type="button"
                onClick={() => setFilter("user", "Usuario", item.name)}
              >
                <span>{index + 1}</span>
                <strong>{item.name}</strong>
                <div>
                  <i style={{ width: `${Math.max(8, (item.total / userChartMax) * 100)}%` }} />
                </div>
                <em>{pointLabel("user", item)}</em>
              </button>
            ))}
          </div>
        </article>

        <article className={styles.panel}>
          <div className={styles.panelHeader}>
            <div>
              <h2>{activeFilter ? `Detalle de ${activeFilter.value}` : "Detalle completo"}</h2>
              <span>
                {visibleRows.length.toLocaleString("es-PE")} registros
                {hiddenDetailRows ? `, mostrando ${detailRows.length.toLocaleString("es-PE")} para mantener fluidez` : ""}
              </span>
            </div>
          </div>
          {activeFilter ? (
            <div className={styles.userSummary}>
              <div>
                <span>Status name</span>
                <strong>{detailStatuses.map((item) => `${item.name}: ${item.total}`).join(" | ")}</strong>
              </div>
              <div>
                <span>Campaign ID</span>
                <strong>{detailCampaigns.map((item) => `${item.name}: ${item.total}`).join(" | ")}</strong>
              </div>
              <div>
                <span>Usuarios</span>
                <strong>{detailUsers.map((item) => `${item.name}: ${item.total}`).join(" | ")}</strong>
              </div>
            </div>
          ) : null}
          <div className={styles.tableWrap}>
            <table>
              <thead>
                <tr>
                  {detailColumns.map((column) => (
                    <th key={column}>{column}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {detailRows.map((row) => (
                  <tr key={row.id}>
                    {detailColumns.map((column) => (
                      <td key={`${row.id}-${column}`}>{displayCellValue(column, row.raw[column])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </article>
      </section>
        </>
      ) : null}

      {activeView === "calls" && expandedChart ? (
        <div className={styles.modalBackdrop} onClick={() => setExpandedChart(null)} role="presentation">
          <section className={styles.modal} onClick={stopInsideClick} role="dialog" aria-modal="true" aria-label="Grafico ampliado">
            <div className={styles.modalHeader}>
              <div>
                <h2>
                  {expandedChart === "date" ? "Registros por fecha" : null}
                  {expandedChart === "hour" ? "Registros por hora" : null}
                  {expandedChart === "status" ? "Status name" : null}
                  {expandedChart === "campaign" ? "Campaign ID" : null}
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
    </main>
  );
}
