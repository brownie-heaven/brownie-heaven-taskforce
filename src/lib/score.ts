import { ANNA_NAGAR_HANDOVER_FROM } from "./outletOwner";
// ── Scoring constants (must stay identical to the Leaderboard) ──────────────
const SEASON_START = "2026-06-12";
const RATING_START = "2026-06-18";
const STARTING_POINTS: Record<string, number> = { ahila: 630, nilani: 430, vishnu: 30 };
const PTS_REPORT = 20;
const PTS_DAILY = 10;
const PTS_DAILY_PENALTY = 5;
const PTS_TARGET_MET = 30;
const PTS_TARGET_MISS = 0;
const PTS_TASK = 5;
const PTS_RATING = 100;
const PTS_RATING_FAIL = 50;
const RATING_THRESHOLD = 4.5;
const BACKFILL_PENALTY = 30;

// The scoring roster (non-Owner staff are ranked; Owner is excluded).
export const SCORE_STAFF: { id: string; name: string; role: string }[] = [
  { id: "nishant", name: "Nishant Vijayakumar", role: "Owner" },
  { id: "arun", name: "Arun Kumar", role: "Manager" },
  { id: "nilani", name: "Nilani Nallamuthu", role: "HR" },
  { id: "gowtham", name: "Gowtham", role: "Purchase Manager" },
  { id: "vishnu", name: "Vishnu", role: "Asst. Ops Manager" },
  { id: "ahila", name: "Ahila", role: "Custom Cakes & Asst Ops" },
  { id: "bharani", name: "Bharani", role: "Auditor" },
];

export type ScoreRow = {
  id: string; name: string; role: string;
  myReports: number; myLate: number; dailyPoints: number;
  outlets: number; outletLate: number; targetMet: number; targetMiss: number;
  tasks: number; ratingPoints: number; backfills: number; adjustments: number;
  points: number; missing: boolean; off: boolean; dailyToday: string;
};

export type ScoreData = {
  reports: any[]; tasks: any[]; outlets: any[]; adjustments: any[]; offs: any[];
};

export type ScoreResult = { rows: ScoreRow[]; arun: ScoreRow | null; worst: ScoreRow | null; best: ScoreRow | null };

const pad = (n: number) => String(n).padStart(2, "0");

// Window boundaries for a given month, clamped to the season start.
export function scoreWindow(y: number, m: number) {
  const monthStart = `${y}-${pad(m + 1)}-01`;
  const startDate = SEASON_START > monthStart ? SEASON_START : monthStart;
  const startISO = new Date(startDate + "T00:00:00").toISOString();
  const endISO = new Date(y, m + 1, 1).toISOString();
  const endDate = m === 11 ? `${y + 1}-01-01` : `${y}-${pad(m + 2)}-01`;
  return { startDate, startISO, endISO, endDate };
}

// ── Pure computation (identical math to the Leaderboard) ────────────────────
export function scoreFromData(y: number, m: number, now: Date, data: ScoreData): ScoreResult {
  const { startDate, startISO, endISO } = scoreWindow(y, m);
  const isCurrentMonth = y === now.getFullYear() && m === now.getMonth();

  const map: Record<string, ScoreRow> = {};
  SCORE_STAFF.filter((s) => s.role !== "Owner").forEach((s) => {
    map[s.id] = { id: s.id, name: s.name, role: s.role, myReports: 0, myLate: 0, dailyPoints: 0, outlets: 0, outletLate: 0, targetMet: 0, targetMiss: 0, tasks: 0, ratingPoints: 0, backfills: 0, adjustments: 0, points: 0, missing: false, off: false, dailyToday: "" };
  });

  const _offMonth = new Set((data.offs || []).map((o: any) => o.staff_id + "_" + o.off_date));

  const dayMap: Record<string, { sid: string; at: string; late: boolean; backfill: boolean }> = {};
  (data.reports || []).forEach((r: any) => {
    if (!map[r.staff_id]) return;
    if (r.no_points) return;
    const day = r.report_date || r.submitted_at.split("T")[0];
    if (_offMonth.has(r.staff_id + "_" + day)) return;
    const key = r.staff_id + "|" + day;
    if (!dayMap[key] || r.submitted_at < dayMap[key].at) {
      dayMap[key] = { sid: r.staff_id, at: r.submitted_at, late: !!r.is_late, backfill: !!r.is_backfill };
    }
  });
  Object.values(dayMap).forEach((d) => {
    const row = map[d.sid]; if (!row) return;
    row.myReports++;
    if (d.late || d.backfill) row.myLate++;
    row.dailyPoints += (d.backfill || d.late) ? -PTS_DAILY_PENALTY : PTS_DAILY;
  });

  (data.tasks || []).forEach((t: any) => {
    const row = map[t.assigned_to]; if (!row) return;
    const d = t.completed_at || t.created_at;
    if (d && d >= startISO && d < endISO) row.tasks++;
  });

  (data.outlets || []).forEach((o: any) => {
    if (o.no_points) return;
    if (_offMonth.has(o.staff_id + "_" + o.report_date)) return;
    const total = (Number(o.shop_sales_value) || 0) + (Number(o.swiggy_sales_value) || 0) + (Number(o.zomato_sales_value) || 0);
    const tgt = Number(o.target) || 0;
    const row = map[o.staff_id];
    if (!row) return;
    if (o.is_backfill) { row.backfills++; return; }
    row.outlets++;
    if (o.is_late) { row.outletLate++; return; }
    if (tgt > 0) { if (total >= tgt) row.targetMet++; else row.targetMiss++; }
  });

  const lastDay = new Date(y, m + 1, 0).getDate();
  const weeks = [
    { day: 7, reward: false },
    { day: 15, reward: true },
    { day: 22, reward: false },
    { day: lastDay, reward: true },
  ].filter((wk) => `${y}-${pad(m + 1)}-${pad(wk.day)}` >= RATING_START);
  const todayStr = now.toISOString().split("T")[0];
  const byOutlet: Record<string, any[]> = {};
  (data.outlets || []).forEach((o: any) => {
    if (o.no_points || o.is_backfill) return;
    if (_offMonth.has(o.staff_id + "_" + o.report_date)) return;
    if (o.bh_google_rating === null || o.bh_google_rating === undefined) return;
    (byOutlet[o.outlet_id] = byOutlet[o.outlet_id] || []).push(o);
  });
  Object.values(byOutlet).forEach((reps: any[]) => {
    reps.sort((a, b) => (a.report_date < b.report_date ? -1 : 1));
    weeks.forEach((wk) => {
      const cp = `${y}-${pad(m + 1)}-${pad(wk.day)}`;
      if (cp > todayStr) return;
      const upto = reps.filter((r) => r.report_date <= cp && r.report_date >= startDate);
      if (!upto.length) return;
      const latest = upto[upto.length - 1];
      const ok = Number(latest.bh_google_rating) >= RATING_THRESHOLD;
      const pts = ok ? (wk.reward ? PTS_RATING : 0) : -PTS_RATING_FAIL;
      if (pts !== 0 && map[latest.staff_id]) map[latest.staff_id].ratingPoints += pts;
    });
  });

  (data.adjustments || []).forEach((a: any) => {
    const row = map[a.staff_id]; if (!row) return;
    row.adjustments += a.points;
  });

 const _applyStart = (y * 12 + m) <= (2026 * 12 + 6); // June(5)+July(6) 2026 only; 0 from Aug onward
  Object.values(map).forEach((row) => {
    row.points =
      (_applyStart ? (STARTING_POINTS[row.id] || 0) : 0) +
      row.dailyPoints +
      (row.outlets - row.outletLate) * PTS_REPORT +
      row.targetMet * PTS_TARGET_MET - row.targetMiss * PTS_TARGET_MISS +
      row.tasks * PTS_TASK +
      row.ratingPoints -
      row.backfills * BACKFILL_PENALTY +
      row.adjustments;
  });

  const all = Object.values(map);
  const arunRow = all.find((r) => r.id === "arun") || null;

  if (isCurrentMonth) {
    const _todayStr2 = now.toISOString().split("T")[0];
    const _afterNoon = now.getHours() >= 12;
    const _after10 = now.getHours() >= 22;
    const _filedDaily = new Set<string>();
    (data.reports || []).forEach((r: any) => { const d = r.report_date || (r.submitted_at ? r.submitted_at.split("T")[0] : ""); if (d === _todayStr2) _filedDaily.add(r.staff_id); });
    const _outletFiled = new Set<string>();
    (data.outlets || []).forEach((o: any) => { if (o.report_date === _todayStr2) _outletFiled.add(o.outlet_id); });
   const _anStephen = _todayStr2 >= ANNA_NAGAR_HANDOVER_FROM;
   const _MGR_OUTLETS: Record<string, string[]> = { nilani: [], vishnu: ["velachery", "perumbakkam", "tambaram", "porur", ...(_anStephen ? [] : ["anna_nagar"]), "vadapalani"], stephen: _anStephen ? ["anna_nagar"] : [], ahila: ["adayar", "bsr_mall", "besant_nagar", "pallavaram", "ra_puram"] };
    all.forEach((r) => { r.off = _offMonth.has(r.id + "_" + _todayStr2); const md = _after10 && !_filedDaily.has(r.id); const mo = _afterNoon && (_MGR_OUTLETS[r.id] || []).some((o) => !_outletFiled.has(o)); r.missing = !r.off && (md || mo); r.dailyToday = r.off ? "off" : (_filedDaily.has(r.id) ? "done" : (_after10 ? "missed" : "pending")); });
  } else {
    all.forEach((r) => { r.off = false; r.missing = false; r.dailyToday = ""; });
  }

  const rows = all.filter((r) => r.id !== "bharani").sort((a, b) => b.points - a.points);
  const worst = rows.length ? rows[rows.length - 1] : null;
  const best = rows.length ? rows[0] : null;
  return { rows, arun: arunRow, worst, best };
}

// ── Fetch + compute for a month (defaults to the current month) ─────────────
export async function computeScores(year?: number, month?: number, now: Date = new Date()): Promise<ScoreResult> {
  const y = year ?? now.getFullYear();
  const m = month ?? now.getMonth();
  const { startDate, startISO, endISO, endDate } = scoreWindow(y, m);
  const { supabase } = await import("@/lib/supabase");

  const [repRes, taskRes, outRes, adjRes, offRes] = await Promise.all([
    supabase.from("reports").select("staff_id,is_late,no_points,submitted_at,report_date,is_backfill").gte("submitted_at", startISO).lt("submitted_at", endISO),
    supabase.from("tasks").select("assigned_to,completed_at,created_at").eq("status", "completed"),
    supabase.from("outlet_reports").select("staff_id,outlet_id,is_late,bh_google_rating,report_date,shop_sales_value,swiggy_sales_value,zomato_sales_value,target,no_points,is_backfill").gte("report_date", startDate).lt("report_date", endDate),
    supabase.from("point_adjustments").select("staff_id,points"),
    supabase.from("day_off").select("staff_id,off_date").gte("off_date", startDate).lt("off_date", endDate),
  ]);

  return scoreFromData(y, m, now, {
    reports: repRes.data || [], tasks: taskRes.data || [], outlets: outRes.data || [],
    adjustments: adjRes.data || [], offs: offRes.data || [],
  });
}
