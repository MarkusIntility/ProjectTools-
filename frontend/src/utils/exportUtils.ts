import { jsPDF } from "jspdf";
import { autoTable } from "jspdf-autotable";
import * as XLSX from "xlsx";
import type {
  Project,
  RiskMatrix,
  ProjectPlan,
  OppgaveListe,
  Runbook,
  MeetingPlan,
  CommunicationPlan,
} from "../api/client";
import type { PlannerData } from "../auth/plannerService";

// ── Constants ─────────────────────────────────────────────────────────────────

const NAVY: [number, number, number] = [15, 25, 65];
const YELLOW: [number, number, number] = [255, 228, 96];
const ALT_ROW: [number, number, number] = [246, 248, 252];
const HEAD_TEXT = 255;

// ── Helpers ───────────────────────────────────────────────────────────────────

function safeName(str: string): string {
  return str.replace(/[<>:"/\\|?*]/g, "_").slice(0, 60);
}

function fmtDate(iso: string | null | undefined): string {
  if (!iso) return "–";
  return new Date(iso).toLocaleDateString("nb-NO", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function fmtDateTime(iso: string | null | undefined): string {
  if (!iso) return "–";
  return new Date(iso).toLocaleString("nb-NO", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

async function loadLogo(): Promise<string | null> {
  try {
    const res = await fetch("/intility-logo.png");
    if (!res.ok) return null;
    const blob = await res.blob();
    return new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  } catch {
    return null;
  }
}

// ── PDF header ────────────────────────────────────────────────────────────────

interface PdfHeaderProps {
  projectName: string;
  projectManager: string | null;
  documentTitle: string;
}

async function addPdfHeader(doc: jsPDF, props: PdfHeaderProps): Promise<number> {
  const W = doc.internal.pageSize.getWidth();
  const STRIP_H = 40;

  doc.setFillColor(...NAVY);
  doc.rect(0, 0, W, STRIP_H, "F");

  const logo = await loadLogo();
  if (logo) {
    doc.addImage(logo, "PNG", 10, (STRIP_H - 14) / 2, 50, 14);
  } else {
    doc.setFontSize(17);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...YELLOW);
    doc.text("intility", 10, STRIP_H / 2 + 5);
  }

  doc.setFontSize(13);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(255, 255, 255);
  doc.text(props.projectName, W - 10, 14, { align: "right" });

  doc.setFontSize(8.5);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(190, 205, 230);
  let metaY = props.projectManager ? 23 : 27;
  if (props.projectManager) {
    doc.text(`Prosjektleder: ${props.projectManager}`, W - 10, metaY, { align: "right" });
    metaY += 7;
  }
  doc.text(
    `Eksportert: ${new Date().toLocaleDateString("nb-NO", { day: "numeric", month: "long", year: "numeric" })}`,
    W - 10,
    metaY,
    { align: "right" }
  );

  const titleY = STRIP_H + 13;
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(20, 20, 40);
  doc.text(props.documentTitle, 14, titleY);

  doc.setDrawColor(210, 218, 235);
  doc.setLineWidth(0.4);
  doc.line(14, titleY + 4, W - 14, titleY + 4);

  return titleY + 10; // returns startY for first table
}

function addFooters(doc: jsPDF) {
  const n = doc.getNumberOfPages();
  const W = doc.internal.pageSize.getWidth();
  const H = doc.internal.pageSize.getHeight();
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    doc.setFontSize(7.5);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(160, 170, 185);
    doc.text("ProjectTools – Intility", 14, H - 8);
    doc.text(`Side ${i} av ${n}`, W - 14, H - 8, { align: "right" });
  }
}

function savePdf(doc: jsPDF, filename: string) {
  const blob = doc.output("blob");
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// ── Risk Matrix ───────────────────────────────────────────────────────────────

export async function exportRiskMatrixPdf(matrix: RiskMatrix, project: Project) {
  try {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: matrix.title,
  });

  const STATUS: Record<string, string> = { open: "Åpen", mitigated: "Mitigert", closed: "Lukket" };

  autoTable(doc, {
    startY,
    head: [["Beskrivelse", "Fase", "Fagområde", "Risiko eier", "S", "K", "Score", "Tiltak", "Ansvarlig (tiltak)", "Restrisiko", "Status"]],
    body: matrix.risks.map((r) => [
      r.description,
      r.fase ?? "–",
      r.fagomrade ?? "–",
      r.risk_owner ?? "–",
      r.probability,
      r.consequence,
      r.risk_score,
      r.mitigation ?? "–",
      r.owner ?? "–",
      r.residual_score != null ? r.residual_score : "–",
      STATUS[r.status] ?? r.status,
    ]),
    styles: { fontSize: 7.5, cellPadding: 2 },
    headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold", fontSize: 8 },
    alternateRowStyles: { fillColor: ALT_ROW },
    columnStyles: {
      4: { halign: "center", cellWidth: 8 },
      5: { halign: "center", cellWidth: 8 },
      6: { halign: "center", cellWidth: 12 },
      9: { halign: "center", cellWidth: 16 },
    },
  });

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(matrix.title)}_risikomatrise.pdf`);
  } catch (err) { console.error("[exportRiskMatrixPdf]", err); throw err; }
}

export function exportRiskMatrixExcel(matrix: RiskMatrix, project: Project) {
  const STATUS: Record<string, string> = { open: "Åpen", mitigated: "Mitigert", closed: "Lukket" };
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Beskrivelse", "Fase", "Fagområde", "Risiko eier", "Sannsynlighet", "Konsekvens", "Score", "Tiltak", "Ansvarlig (tiltak)", "Restrisiko", "Status"],
    ...matrix.risks.map((r) => [
      r.description, r.fase ?? "", r.fagomrade ?? "", r.risk_owner ?? "",
      r.probability, r.consequence, r.risk_score,
      r.mitigation ?? "", r.owner ?? "",
      r.residual_score ?? "",
      STATUS[r.status] ?? r.status,
    ]),
  ]);
  ws["!cols"] = [40, 12, 18, 20, 14, 14, 8, 35, 20, 14, 12].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Risikomatrise");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(matrix.title)}_risikomatrise.xlsx`);
}

// ── Project Plan ──────────────────────────────────────────────────────────────

export async function exportProjectPlanPdf(
  plan: ProjectPlan,
  project: Project,
  plannerData: PlannerData | null
) {
  try {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: plan.title,
  });

  if (plan.source === "own") {
    autoTable(doc, {
      startY,
      head: [["Oppgave", "Fase/Bucket", "Start", "Frist", "Ansvarlig", "% Ferdig"]],
      body: plan.tasks.map((t) => [
        t.name,
        t.bucket ?? "–",
        fmtDate(t.start_date),
        fmtDate(t.end_date),
        t.responsible ?? "–",
        `${t.percent_complete}%`,
      ]),
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
      columnStyles: { 5: { halign: "center", cellWidth: 20 } },
    });
  } else if (plannerData) {
    autoTable(doc, {
      startY,
      head: [["Oppgave", "Bucket", "Forfallsdato", "% Ferdig"]],
      body: plannerData.tasks.map((t) => [
        t.title,
        plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "–",
        fmtDate(t.dueDateTime ?? null),
        `${t.percentComplete}%`,
      ]),
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
      columnStyles: { 3: { halign: "center", cellWidth: 20 } },
    });
  } else {
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 120);
    doc.text("Ekstern plan — data hentes fra Planner/Smartsheet.", 14, startY + 6);
  }

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(plan.title)}_prosjektplan.pdf`);
  } catch (err) { console.error("[exportProjectPlanPdf]", err); throw err; }
}

export function exportProjectPlanExcel(
  plan: ProjectPlan,
  project: Project,
  plannerData: PlannerData | null
) {
  const wb = XLSX.utils.book_new();
  let rows: (string | number)[][];

  if (plan.source === "own") {
    rows = [
      ["Oppgave", "Fase/Bucket", "Start", "Frist", "Ansvarlig", "% Ferdig"],
      ...plan.tasks.map((t) => [
        t.name, t.bucket ?? "", fmtDate(t.start_date), fmtDate(t.end_date),
        t.responsible ?? "", t.percent_complete,
      ]),
    ];
  } else if (plannerData) {
    rows = [
      ["Oppgave", "Bucket", "Forfallsdato", "% Ferdig"],
      ...plannerData.tasks.map((t) => [
        t.title, plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "", fmtDate(t.dueDateTime ?? null), t.percentComplete,
      ]),
    ];
  } else {
    rows = [["Ekstern plan — ingen lokal data tilgjengelig"]];
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [45, 20, 16, 16, 22, 12].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Prosjektplan");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(plan.title)}_prosjektplan.xlsx`);
}

// ── Oppgaveliste ──────────────────────────────────────────────────────────────

const OPPGAVE_STATUS: Record<string, string> = {
  not_started: "Ikke startet",
  in_progress: "Pågående",
  done: "Ferdig",
};

export async function exportOppgavePdf(
  liste: OppgaveListe,
  project: Project,
  plannerData: PlannerData | null
) {
  try {
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: liste.title,
  });

  if (liste.source === "own") {
    autoTable(doc, {
      startY,
      head: [["Oppgave", "Ansvarlig", "Forfallsdato", "Status", "Beskrivelse"]],
      body: liste.oppgaver.map((o) => [
        o.name,
        o.responsible ?? "–",
        fmtDate(o.due_date),
        OPPGAVE_STATUS[o.status] ?? o.status,
        o.description ?? "–",
      ]),
      styles: { fontSize: 8.5, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
      columnStyles: { 0: { cellWidth: 42 }, 1: { cellWidth: 35 }, 2: { cellWidth: 28 }, 3: { cellWidth: 24 }, 4: { cellWidth: "auto" } },
    });
  } else if (plannerData) {
    autoTable(doc, {
      startY,
      head: [["Oppgave", "Bucket", "Forfallsdato", "% Ferdig"]],
      body: plannerData.tasks.map((t) => [
        t.title, plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "–", fmtDate(t.dueDateTime ?? null), `${t.percentComplete}%`,
      ]),
      styles: { fontSize: 9, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
    });
  } else {
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 120);
    doc.text("Ekstern liste — data hentes fra Planner/Smartsheet.", 14, startY + 6);
  }

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(liste.title)}_oppgaver.pdf`);
  } catch (err) { console.error("[exportOppgavePdf]", err); throw err; }
}

export function exportOppgaveExcel(
  liste: OppgaveListe,
  project: Project,
  plannerData: PlannerData | null
) {
  const wb = XLSX.utils.book_new();
  let rows: (string | number)[][];

  if (liste.source === "own") {
    rows = [
      ["Oppgave", "Ansvarlig", "Forfallsdato", "Status", "Beskrivelse"],
      ...liste.oppgaver.map((o) => [
        o.name, o.responsible ?? "", fmtDate(o.due_date), OPPGAVE_STATUS[o.status] ?? o.status, o.description ?? "",
      ]),
    ];
  } else if (plannerData) {
    rows = [
      ["Oppgave", "Bucket", "Forfallsdato", "% Ferdig"],
      ...plannerData.tasks.map((t) => [
        t.title, plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "", fmtDate(t.dueDateTime ?? null), t.percentComplete,
      ]),
    ];
  } else {
    rows = [["Ekstern liste — ingen lokal data tilgjengelig"]];
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [42, 22, 16, 16, 40].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Oppgaver");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(liste.title)}_oppgaver.xlsx`);
}

// ── Runbook ───────────────────────────────────────────────────────────────────

const ACTIVITY_STATUS: Record<string, string> = {
  not_started: "Ikke startet",
  in_progress: "Pågående",
  done: "Ferdig",
  cancelled: "Kansellert",
};

export async function exportRunbookPdf(
  runbook: Runbook,
  project: Project,
  plannerData: PlannerData | null
) {
  try {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: runbook.title,
  });

  if (runbook.source === "own") {
    autoTable(doc, {
      startY,
      head: [["Aktivitet", "Fase", "Status", "Start", "Frist", "Ansvarlig"]],
      body: runbook.activities.map((a) => [
        a.name,
        a.phase ?? "–",
        ACTIVITY_STATUS[a.status] ?? a.status,
        fmtDate(a.start_date),
        fmtDate(a.end_date),
        a.responsible ?? "–",
      ]),
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
    });
  } else if (plannerData) {
    autoTable(doc, {
      startY,
      head: [["Aktivitet", "Bucket", "Forfallsdato", "% Ferdig"]],
      body: plannerData.tasks.map((t) => [
        t.title, plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "–", fmtDate(t.dueDateTime ?? null), `${t.percentComplete}%`,
      ]),
      styles: { fontSize: 8, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
    });
  } else {
    doc.setFontSize(10);
    doc.setTextColor(100, 100, 120);
    doc.text("Ekstern runbook — data hentes fra Planner/Smartsheet.", 14, startY + 6);
  }

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(runbook.title)}_runbook.pdf`);
  } catch (err) { console.error("[exportRunbookPdf]", err); throw err; }
}

export function exportRunbookExcel(
  runbook: Runbook,
  project: Project,
  plannerData: PlannerData | null
) {
  const wb = XLSX.utils.book_new();
  let rows: (string | number)[][];

  if (runbook.source === "own") {
    rows = [
      ["Aktivitet", "Fase", "Status", "Start", "Frist", "Ansvarlig"],
      ...runbook.activities.map((a) => [
        a.name, a.phase ?? "", ACTIVITY_STATUS[a.status] ?? a.status,
        fmtDate(a.start_date), fmtDate(a.end_date), a.responsible ?? "",
      ]),
    ];
  } else if (plannerData) {
    rows = [
      ["Aktivitet", "Bucket", "Forfallsdato", "% Ferdig"],
      ...plannerData.tasks.map((t) => [
        t.title, plannerData.buckets.find((b) => b.id === t.bucketId)?.name ?? "", fmtDate(t.dueDateTime ?? null), t.percentComplete,
      ]),
    ];
  } else {
    rows = [["Ekstern runbook — ingen lokal data tilgjengelig"]];
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [45, 20, 18, 16, 16, 25].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Runbook");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(runbook.title)}_runbook.xlsx`);
}

// ── Meeting Plan ──────────────────────────────────────────────────────────────

export async function exportMeetingPlanPdf(plan: MeetingPlan, project: Project) {
  try {
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: plan.title,
  });

  autoTable(doc, {
    startY,
    head: [["Møte", "Dato og tid", "Formål"]],
    body: plan.meetings.map((m) => [
      m.title,
      fmtDateTime(m.date),
      m.purpose ?? "–",
    ]),
    styles: { fontSize: 9, cellPadding: 2 },
    headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
    alternateRowStyles: { fillColor: ALT_ROW },
    columnStyles: { 0: { cellWidth: 55 }, 1: { cellWidth: 45 }, 2: { cellWidth: "auto" } },
  });

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(plan.title)}_møteplan.pdf`);
  } catch (err) { console.error("[exportMeetingPlanPdf]", err); throw err; }
}

export function exportMeetingPlanExcel(plan: MeetingPlan, project: Project) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Møte", "Dato og tid", "Formål"],
    ...plan.meetings.map((m) => [m.title, fmtDateTime(m.date), m.purpose ?? ""]),
  ]);
  ws["!cols"] = [35, 28, 50].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Møteplan");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(plan.title)}_møteplan.xlsx`);
}

// ── Communication Plan ────────────────────────────────────────────────────────

export async function exportCommPlanPdf(plan: CommunicationPlan, project: Project) {
  try {
  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const startY = await addPdfHeader(doc, {
    projectName: project.name,
    projectManager: project.project_manager,
    documentTitle: plan.title,
  });

  autoTable(doc, {
    startY,
    head: [["Interessent", "Budskap", "Kanal", "Frekvens", "Ansvarlig"]],
    body: plan.entries.map((e) => [
      e.stakeholder,
      e.message,
      e.channel,
      e.frequency,
      e.responsible,
    ]),
    styles: { fontSize: 8.5, cellPadding: 2 },
    headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
    alternateRowStyles: { fillColor: ALT_ROW },
  });

  addFooters(doc);
  savePdf(doc, `${safeName(project.name)}_${safeName(plan.title)}_kommunikasjonsplan.pdf`);
  } catch (err) { console.error("[exportCommPlanPdf]", err); throw err; }
}

export function exportCommPlanExcel(plan: CommunicationPlan, project: Project) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ["Interessent", "Budskap", "Kanal", "Frekvens", "Ansvarlig"],
    ...plan.entries.map((e) => [e.stakeholder, e.message, e.channel, e.frequency, e.responsible]),
  ]);
  ws["!cols"] = [25, 50, 18, 18, 22].map((wch) => ({ wch }));
  XLSX.utils.book_append_sheet(wb, ws, "Kommunikasjonsplan");
  XLSX.writeFile(wb, `${safeName(project.name)}_${safeName(plan.title)}_kommunikasjonsplan.xlsx`);
}

// ── Dashboard (PDF only) ──────────────────────────────────────────────────────

interface DashboardExportData {
  project: Project;
  openRisksCount: number;
  highestRiskScore: number;
  topRisks: Array<{ description: string; risk_score: number; fagomrade: string | null; risk_owner: string | null }>;
  doneOppgaver: number;
  totalOppgaver: number;
  nextMeetingTitle: string | null;
  nextMeetingDate: string | null;
  leveranserPct: number | null;
  planTasksDone: number;
  planTasksTotal: number;
  primaryPlanTitle: string | null;
  primaryOppgaveTitle: string | null;
  primaryMeetingTitle: string | null;
  primaryMatrixTitle: string | null;
}

function riskLevelLabel(score: number): string {
  if (score <= 6) return "Lav";
  if (score <= 14) return "Middels";
  return "Høy";
}

export async function exportDashboardPdf(data: DashboardExportData) {
  try {
  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const W = doc.internal.pageSize.getWidth();

  const startY = await addPdfHeader(doc, {
    projectName: data.project.name,
    projectManager: data.project.project_manager,
    documentTitle: "Prosjektstatusrapport",
  });

  let y = startY + 4;

  // ── KPI boxes ──────────────────────────────────────────────────────────────
  const kpis = [
    { label: "Åpne risikoer", value: String(data.openRisksCount), sub: data.openRisksCount > 0 ? riskLevelLabel(data.highestRiskScore) + " risiko" : "Ingen åpne risikoer" },
    { label: "Ferdige oppgaver", value: String(data.doneOppgaver), sub: data.totalOppgaver > 0 ? `av ${data.totalOppgaver} totalt` : "Ingen oppgaver" },
    { label: "Neste møte", value: data.nextMeetingDate ? new Date(data.nextMeetingDate).toLocaleDateString("nb-NO", { day: "numeric", month: "short" }) : "–", sub: data.nextMeetingTitle ?? "Ingen planlagte møter" },
    { label: "Leveranser", value: data.leveranserPct !== null ? `${data.leveranserPct}%` : "–", sub: data.planTasksTotal > 0 ? `${data.planTasksDone} av ${data.planTasksTotal} ferdig` : "Ingen prosjektplan" },
  ];

  const BOX_W = (W - 28 - 9) / 4;
  const BOX_H = 28;
  let bx = 14;
  for (const kpi of kpis) {
    doc.setFillColor(246, 248, 252);
    doc.setDrawColor(210, 218, 235);
    doc.setLineWidth(0.3);
    doc.roundedRect(bx, y, BOX_W, BOX_H, 2, 2, "FD");

    doc.setFontSize(17);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(...NAVY);
    doc.text(kpi.value, bx + BOX_W / 2, y + 11, { align: "center" });

    doc.setFontSize(7);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(60, 70, 90);
    doc.text(kpi.label.toUpperCase(), bx + BOX_W / 2, y + 18, { align: "center" });

    doc.setFontSize(6.5);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(130, 140, 160);
    doc.text(kpi.sub, bx + BOX_W / 2, y + 23, { align: "center", maxWidth: BOX_W - 4 });

    bx += BOX_W + 3;
  }

  y += BOX_H + 10;

  // ── Top risks ──────────────────────────────────────────────────────────────
  if (data.topRisks.length > 0) {
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(20, 20, 40);
    doc.text(`Topp risikoer${data.primaryMatrixTitle ? ` — ${data.primaryMatrixTitle}` : ""}`, 14, y);
    y += 3;

    autoTable(doc, {
      startY: y,
      head: [["Beskrivelse", "Fagområde", "Risiko eier", "Score", "Nivå"]],
      body: data.topRisks.map((r) => [
        r.description,
        r.fagomrade ?? "–",
        r.risk_owner ?? "–",
        r.risk_score,
        riskLevelLabel(r.risk_score),
      ]),
      styles: { fontSize: 8.5, cellPadding: 2 },
      headStyles: { fillColor: NAVY, textColor: HEAD_TEXT, fontStyle: "bold" },
      alternateRowStyles: { fillColor: ALT_ROW },
      columnStyles: {
        3: { halign: "center", cellWidth: 16 },
        4: { halign: "center", cellWidth: 20 },
      },
      margin: { left: 14, right: 14 },
    });

    y = (doc as jsPDF & { lastAutoTable: { finalY: number } }).lastAutoTable.finalY + 10;
  }

  // ── Primær ressurser ───────────────────────────────────────────────────────
  const resources = [
    { label: "Primær prosjektplan", value: data.primaryPlanTitle ?? "–" },
    { label: "Primær oppgaveliste", value: data.primaryOppgaveTitle ?? "–" },
    { label: "Primær møteplan", value: data.primaryMeetingTitle ?? "–" },
    { label: "Primær risikomatrise", value: data.primaryMatrixTitle ?? "–" },
  ].filter((r) => r.value !== "–");

  if (resources.length > 0) {
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(20, 20, 40);
    doc.text("Primære ressurser", 14, y);
    y += 3;

    autoTable(doc, {
      startY: y,
      body: resources.map((r) => [r.label, r.value]),
      styles: { fontSize: 9, cellPadding: 2 },
      alternateRowStyles: { fillColor: ALT_ROW },
      columnStyles: { 0: { fontStyle: "bold", cellWidth: 60 } },
      margin: { left: 14, right: 14 },
    });
  }

  addFooters(doc);
  savePdf(doc, `${safeName(data.project.name)}_statusrapport.pdf`);
  } catch (err) { console.error("[exportDashboardPdf]", err); throw err; }
}
