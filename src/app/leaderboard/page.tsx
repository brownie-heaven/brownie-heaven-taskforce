"use client";
import { useEffect, useState } from "react";
import { useRouter } from "next/navigation";
import { createClient } from "@supabase/supabase-js";
import { ANNA_NAGAR_HANDOVER_FROM } from "@/lib/outletOwner";

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
);

type Staff = { id: string; name: string; role: string };

const ALL_STAFF: Staff[] = [
  { id: "nishant", name: "Nishant Vijayakumar", role: "Owner" },
  { id: "arun", name: "Arun Kumar", role: "Manager" },
  { id: "nilani", name: "Nilani Nallamuthu", role: "HR" },
  { id: "gowtham", name: "Gowtham", role: "Purchase Manager" },
  { id: "vishnu", name: "Vishnu", role: "Asst. Ops Manager" },
  { id: "ahila", name: "Ahila", role: "Custom Cakes & Asst Ops" },
  { id: "bharani", name: "Bharani", role: "Auditor" },
];

const SEASON_START = "2026-06-12";
const STARTING_POINTS: Record<string, number> = { ahila: 630, nilani: 430, vishnu: 30 };
const PTS_REPORT = 20;
const PTS_DAILY = 10;
const PTS_DAILY_PENALTY = 5;
const PTS_LATE_PENALTY = 20;
const PTS_TARGET_MET = 30;
const PTS_TARGET_MISS = 0;
const PTS_TASK = 5;
const PTS_RATING = 100;
const PTS_RATING_FAIL = 50;
const RATING_THRESHOLD = 4.5;
const BACKFILL_PENALTY = 30;
const ARUN_TARGET = 19000;

type Row = {
  id: string; name: string; role: string;
  myReports: number; myLate: number; dailyPoints: number;
  outlets: number; outletLate: number; targetMet: number; targetMiss: number;
  tasks: number; ratingPoints: number; backfills: number; adjustments: number; points: number; missing: boolean; off: boolean;
};

export default function LeaderboardPage() {
  const router = useRouter();
  const [user, setUser] = useState<Staff | null>(null);
  const [rows, setRows] = useState<Row[]>([]);
  const [arun, setArun] = useState<Row | null>(null);
  const [loading, setLoading] = useState(true);
  const [giveStaff, setGiveStaff] = useState("");
  const [givePoints, setGivePoints] = useState("");
  const [giveReason, setGiveReason] = useState("");
  const [giving, setGiving] = useState(false);

  const now = new Date();

  // ── Month being viewed (defaults to the current month) ─────────────────
  const [viewYear, setViewYear] = useState(now.getFullYear());
  const [viewMonth, setViewMonth] = useState(now.getMonth()); // 0-indexed

  const isCurrentMonth = viewYear === now.getFullYear() && viewMonth === now.getMonth();
  const monthLabel = new Date(viewYear, viewMonth, 1)
    .toLocaleString("en-IN", { month: "long", year: "numeric" }).toUpperCase();

  // Build the list of selectable months: from the season-start month up to the current month.
  const monthOptions: { y: number; m: number; label: string }[] = [];
  {
    const seasonStart = new Date(SEASON_START + "T00:00:00");
    let d = new Date(seasonStart.getFullYear(), seasonStart.getMonth(), 1);
    const cur = new Date(now.getFullYear(), now.getMonth(), 1);
    while (d <= cur) {
      monthOptions.push({
        y: d.getFullYear(),
        m: d.getMonth(),
        label: d.toLocaleString("en-IN", { month: "long", year: "numeric" }),
      });
      d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
    }
    monthOptions.reverse(); // most recent first
  }

  // Auth check (runs once)
  useEffect(() => {
    const stored = localStorage.getItem("currentUser");
    if (!stored) { router.push("/"); return; }
    let parsed;
    try {
      parsed = JSON.parse(stored);
      if (typeof parsed === "string") { localStorage.removeItem("currentUser"); router.push("/"); return; }
    } catch { localStorage.removeItem("currentUser"); router.push("/"); return; }
    setUser(parsed);
  }, [router]);

  // Fetch whenever the user is known OR the viewed month changes
  useEffect(() => {
    if (user) fetchScores();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [user, viewYear, viewMonth]);

  const givePointsFn = async () => {
    if (!giveStaff || !givePoints) return;
    setGiving(true);
    const { error } = await supabase.from("point_adjustments").insert({ staff_id: giveStaff, points: parseInt(givePoints) || 0, reason: giveReason || null });
    setGiving(false);
    if (error) { alert("Error: " + error.message); return; }
    setGiveStaff(""); setGivePoints(""); setGiveReason("");
    fetchScores();
  };

  const fetchScores = async () => {
    setLoading(true);
    // Use the VIEWED month rather than always the current month.
    const y = viewYear;
    const m = viewMonth;
    const pad = (n: number) => String(n).padStart(2, "0");
    const monthStart = `${y}-${pad(m + 1)}-01`;
    const startDate = SEASON_START > monthStart ? SEASON_START : monthStart;
    const startISO = new Date(startDate + "T00:00:00").toISOString();
    const endISO = new Date(y, m + 1, 1).toISOString();
    const endDate = m === 11 ? `${y + 1}-01-01` : `${y}-${pad(m + 2)}-01`;

    const [repRes, taskRes, outRes, adjRes, offRes] = await Promise.all([
      supabase.from("reports").select("staff_id,is_late,no_points,submitted_at,report_date,is_backfill")
        .gte("submitted_at", startISO).lt("submitted_at", endISO),
      supabase.from("tasks").select("assigned_to,completed_at,created_at")
        .eq("status", "completed"),
      supabase.from("outlet_reports").select("staff_id,outlet_id,is_late,bh_google_rating,report_date,shop_sales_value,swiggy_sales_value,zomato_sales_value,target,no_points,is_backfill")
        .gte("report_date", startDate).lt("report_date", endDate),
      supabase.from("point_adjustments").select("staff_id,points"),
      supabase.from("day_off").select("staff_id,off_date").gte("off_date", startDate).lt("off_date", endDate),
    ]);

    const map: Record<string, Row> = {};
    ALL_STAFF.filter(s => s.role !== "Owner").forEach(s => {
      map[s.id] = { id: s.id, name: s.name, role: s.role, myReports: 0, myLate: 0, dailyPoints: 0, outlets: 0, outletLate: 0, targetMet: 0, targetMiss: 0, tasks: 0, ratingPoints: 0, backfills: 0, adjustments: 0, points: 0, missing: false, off: false };
    });

    const _offMonth = new Set((offRes.data || []).map((o: any) => o.staff_id + "_" + o.off_date));
    const dayMap: Record<string, { sid: string; at: string; late: boolean; backfill: boolean }> = {};
    (repRes.data || []).forEach((r: any) => {
      if (!map[r.staff_id]) return;
      if (r.no_points) return;
      const day = r.report_date || r.submitted_at.split("T")[0];
      if (_offMonth.has(r.staff_id + "_" + day)) return;
      const key = r.staff_id + "|" + day;
      if (!dayMap[key] || r.submitted_at < dayMap[key].at) {
        dayMap[key] = { sid: r.staff_id, at: r.submitted_at, late: !!r.is_late, backfill: !!r.is_backfill };
      }
    });
    Object.values(dayMap).forEach(d => {
      const row = map[d.sid]; if (!row) return;
      row.myReports++;
      if (d.late || d.backfill) row.myLate++;
      row.dailyPoints += (d.backfill || d.late) ? -PTS_DAILY_PENALTY : PTS_DAILY;
    });

    // Tasks: only count those completed within the viewed month's window
    (taskRes.data || []).forEach((t: any) => {
      const row = map[t.assigned_to]; if (!row) return;
      const d = t.completed_at || t.created_at;
      if (d && d >= startISO && d < endISO) row.tasks++;
    });

   (outRes.data || []).forEach((o: any) => {
      if (o.no_points) return;
      if (_offMonth.has(o.staff_id + "_" + o.report_date)) return;
      const total = (Number(o.shop_sales_value) || 0) + (Number(o.swiggy_sales_value) || 0) + (Number(o.zomato_sales_value) || 0);
      const tgt = Number(o.target) || 0;
      const credit = (row: Row | undefined, rollup: boolean) => {
        if (!row) return;
        if (o.is_backfill) { if (!rollup) row.backfills++; return; }
       row.outlets++;
        if (o.is_late) { row.outletLate++; return; }
        if (tgt > 0) { if (total >= tgt) row.targetMet++; else row.targetMiss++; }
      };
      credit(map[o.staff_id], false);
    });

    const lastDay = new Date(y, m + 1, 0).getDate();
    const RATING_START = "2026-06-18"; // ratings begin 3rd week of June; no retroactive deductions
    const weeks = [
      { day: 7, reward: false },
      { day: 15, reward: true },
      { day: 22, reward: false },
      { day: lastDay, reward: true },
    ].filter(wk => `${y}-${pad(m + 1)}-${pad(wk.day)}` >= RATING_START);
    // Gate rating checkpoints so we never count a checkpoint that is still in the future.
    const todayStr = now.toISOString().split("T")[0];
    const byOutlet: Record<string, any[]> = {};
    (outRes.data || []).forEach((o: any) => {
      if (o.no_points || o.is_backfill) return;
      if (_offMonth.has(o.staff_id + "_" + o.report_date)) return;
      if (o.bh_google_rating === null || o.bh_google_rating === undefined) return;
      (byOutlet[o.outlet_id] = byOutlet[o.outlet_id] || []).push(o);
    });
    Object.values(byOutlet).forEach((reps: any[]) => {
      reps.sort((a, b) => (a.report_date < b.report_date ? -1 : 1));
      weeks.forEach(wk => {
        const cp = `${y}-${pad(m + 1)}-${pad(wk.day)}`;
        if (cp > todayStr) return;
        const upto = reps.filter(r => r.report_date <= cp && r.report_date >= startDate);
        if (!upto.length) return;
        const latest = upto[upto.length - 1];
        const ok = Number(latest.bh_google_rating) >= RATING_THRESHOLD;
        const pts = ok ? (wk.reward ? PTS_RATING : 0) : -PTS_RATING_FAIL;
        if (pts !== 0 && map[latest.staff_id]) map[latest.staff_id].ratingPoints += pts;
      });
    });

    (adjRes.data || []).forEach((a: any) => {
      const row = map[a.staff_id]; if (!row) return;
      row.adjustments += a.points;
    });

   const _applyStart = (viewYear * 12 + viewMonth) <= (2026 * 12 + 6); // June+July 2026 only
    Object.values(map).forEach(row => {
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
    const arunRow = all.find(r => r.id === "arun");
    const arunOwn = arunRow ? arunRow.points : 0;
    const teamTotal = all.reduce((s, r) => (r.id === "arun" ? s : s + r.points), 0) + arunOwn;
    if (arunRow) arunRow.points = teamTotal;

    // ── Live "not filed / off today" flags only make sense for the current month ──
    if (isCurrentMonth) {
      const _now2 = new Date();
      const _todayStr2 = _now2.toISOString().split("T")[0];
      const _afterNoon = _now2.getHours() >= 12;
      const _after10 = _now2.getHours() >= 22;
      const _filedDaily = new Set<string>();
      (repRes.data || []).forEach((r: any) => { const d = r.report_date || (r.submitted_at ? r.submitted_at.split("T")[0] : ""); if (d === _todayStr2) _filedDaily.add(r.staff_id); });
      const _outletFiled = new Set<string>();
      (outRes.data || []).forEach((o: any) => { if (o.report_date === _todayStr2) _outletFiled.add(o.outlet_id); });
     const _anStephen = _todayStr2 >= ANNA_NAGAR_HANDOVER_FROM;
     const _MGR_OUTLETS: Record<string, string[]> = { nilani: [], vishnu: ["velachery","perumbakkam","tambaram","porur",...(_anStephen ? [] : ["anna_nagar"]),"vadapalani"], stephen: _anStephen ? ["anna_nagar"] : [], ahila: ["royapettah","adayar","bsr_mall","besant_nagar","pallavaram","ra_puram"] };
      all.forEach(r => { r.off = _offMonth.has(r.id + "_" + _todayStr2); const md = _after10 && !_filedDaily.has(r.id); const mo = _afterNoon && (_MGR_OUTLETS[r.id] || []).some(o => !_outletFiled.has(o)); r.missing = !r.off && (md || mo); });
    } else {
      // Historic month: no real-time flags
      all.forEach(r => { r.off = false; r.missing = false; });
    }

    setArun(arunRow || null);
    setRows(all.filter(r => r.id !== "arun").sort((a, b) => b.points - a.points));
    setLoading(false);
  };

  const C = {
    bg: "#0a0a0a", panel: "#141414", border: "#262626",
    text: "#ffffff", muted: "#888888", accent: "#facc15",
  };
  const page: React.CSSProperties = { background: C.bg, color: C.text, minHeight: "100vh", fontFamily: "monospace", padding: "20px" };

  if (loading) {
    return <div style={{ ...page, display: "flex", alignItems: "center", justifyContent: "center", color: C.accent }}>Loading scores…</div>;
  }

  const isOwner = user?.role === "Owner" || user?.role === "Manager";
  const me = rows.find(r => r.id === user?.id);
  const myRank = rows.findIndex(r => r.id === user?.id) + 1;
  const OUTLET_MANAGERS = ["nilani", "vishnu", "ahila"];
  const noOutlet = (id: string) => !OUTLET_MANAGERS.includes(id);
  const cashFor = (p: number, id: string) => { if (noOutlet(id)) return 0; return p >= 6800 ? 3000 : p >= 6400 ? 2000 : p <= 5600 ? -500 : 0; };
  const cashLabel = (p: number, id: string) => { const c = cashFor(p, id); return c > 0 ? `+₹${c}` : c < 0 ? `-₹${Math.abs(c)}` : "—"; };
  const cashColor = (p: number, id: string) => { const c = cashFor(p, id); return c > 0 ? "#22c55e" : c < 0 ? "#ef4444" : C.muted; };

  const th: React.CSSProperties = { textAlign: "right", padding: "10px 12px", color: C.muted, fontSize: "12px", borderBottom: `1px solid ${C.border}`, textTransform: "uppercase" };
  const td: React.CSSProperties = { textAlign: "right", padding: "12px", borderBottom: `1px solid ${C.border}` };

  return (
    <div style={page}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px", flexWrap: "wrap", gap: "8px" }}>
        <h1 style={{ margin: 0, fontSize: "22px", color: C.accent, letterSpacing: "1px" }}>POINTS &amp; REWARDS</h1>
        <button onClick={() => router.push("/dashboard")}
          style={{ background: "transparent", color: C.accent, border: `1px solid ${C.accent}`, padding: "8px 14px", cursor: "pointer", fontFamily: "monospace" }}>
          ← DASHBOARD
        </button>
      </div>

      {/* ── Month selector: view any month back to the season start ── */}
      <div style={{ display: "flex", alignItems: "center", gap: "10px", marginBottom: "18px", flexWrap: "wrap" }}>
        <span style={{ color: C.muted, fontSize: "12px", textTransform: "uppercase", letterSpacing: "1px" }}>Viewing</span>
        <select
          value={`${viewYear}-${viewMonth}`}
          onChange={(e) => { const [yy, mm] = e.target.value.split("-").map(Number); setViewYear(yy); setViewMonth(mm); }}
          style={{ background: "#000", color: C.accent, border: `1px solid ${C.accent}`, padding: "6px 10px", fontFamily: "monospace", fontSize: "13px", cursor: "pointer" }}
        >
          {monthOptions.map(opt => (
            <option key={`${opt.y}-${opt.m}`} value={`${opt.y}-${opt.m}`}>
              {opt.label}{opt.y === now.getFullYear() && opt.m === now.getMonth() ? " (current)" : ""}
            </option>
          ))}
        </select>
        {!isCurrentMonth && (
          <>
            <span style={{ background: "#3f3f46", color: "#fbbf24", fontSize: "10px", fontWeight: "bold", padding: "3px 8px", borderRadius: "3px" }}>
              PAST MONTH · READ ONLY
            </span>
            <button onClick={() => { setViewYear(now.getFullYear()); setViewMonth(now.getMonth()); }}
              style={{ background: "transparent", color: C.muted, border: `1px solid ${C.border}`, padding: "5px 10px", cursor: "pointer", fontFamily: "monospace", fontSize: "12px" }}>
              ← back to current
            </button>
          </>
        )}
      </div>

      <div style={{ background: C.panel, border: `1px solid ${C.accent}`, padding: "12px 16px", marginBottom: "22px", fontSize: "13px", lineHeight: 1.7 }}>
      💰 Monthly cash incentive(outlet managers) — <span style={{ color: "#22c55e", fontWeight: "bold" }}>6800+ = ₹3000</span> · <span style={{ color: "#22c55e", fontWeight: "bold" }}>6400+ = ₹2000</span> · <span style={{ color: "#ef4444", fontWeight: "bold" }}>5600 or below = -₹500</span>
      </div>

      {arun && (
        <div style={{ background: C.panel, border: `1px solid ${C.accent}`, padding: "20px", marginBottom: "26px" }}>
          <div style={{ color: C.muted, fontSize: "12px", textTransform: "uppercase", letterSpacing: "1px" }}>Arun — Team Total (everyone's points) · {monthLabel}</div>
          <div style={{ fontSize: "40px", color: C.accent, fontWeight: "bold", lineHeight: 1.2 }}>{arun.points} <span style={{ fontSize: "16px", color: C.muted }}>/ {ARUN_TARGET}</span></div>
          <div style={{ background: "#000", height: "10px", borderRadius: "5px", overflow: "hidden", margin: "12px 0" }}>
            <div style={{ height: "100%", width: `${Math.min(100, Math.max(0, (arun.points / ARUN_TARGET) * 100))}%`, background: arun.points >= ARUN_TARGET ? "#22c55e" : C.accent }}></div>
          </div>
          <div style={{ color: arun.points >= ARUN_TARGET ? "#22c55e" : C.muted, fontSize: "13px" }}>
            {arun.points >= ARUN_TARGET ? "🎉 Target hit!" : `${ARUN_TARGET - arun.points} points to go`}
          </div>
        </div>
      )}

      {!isOwner && me && (
        <div style={{ background: C.panel, border: `1px solid ${C.border}`, padding: "20px", marginBottom: "26px" }}>
          <div style={{ color: C.muted, fontSize: "12px", textTransform: "uppercase" }}>Your Score · {monthLabel}</div>
          <div style={{ fontSize: "48px", color: C.accent, fontWeight: "bold", lineHeight: 1.1 }}>{me.points}</div>
         <div style={{ color: C.muted, marginBottom: "10px" }}>Rank #{myRank} of {rows.length}</div>
          <div style={{ marginBottom: "16px", fontSize: "15px" }}>{isCurrentMonth ? "This month's cash: " : "Cash: "}<span style={{ color: cashColor(me.points, me.id), fontWeight: "bold" }}>{cashLabel(me.points, me.id)}</span></div>
          <div style={{ fontSize: "13px", lineHeight: 1.9 }}>
            {((viewYear * 12 + viewMonth) <= (2026 * 12 + 6)) && (STARTING_POINTS[me.id] || 0) > 0 && <div>Starting credit: {STARTING_POINTS[me.id]}</div>}
            <div>Daily reports: {me.myReports} filed → {me.dailyPoints >= 0 ? "+" : ""}{me.dailyPoints} (on-time +{PTS_DAILY}, late/back-dated -{PTS_DAILY_PENALTY})</div>
            {me.myLate > 0 && <div>Daily after cut-off: {me.myLate} × 0 = 0</div>}
            <div>Outlet reports (on time): {me.outlets - me.outletLate} × {PTS_REPORT} = {(me.outlets - me.outletLate) * PTS_REPORT}</div>
            {me.outletLate > 0 && <div>Outlet after cut-off: {me.outletLate} × 0 = 0</div>}
            <div>Targets met: {me.targetMet} × {PTS_TARGET_MET} = {me.targetMet * PTS_TARGET_MET}</div>
            <div>Targets missed: {me.targetMiss} (no bonus, kept the +20)</div>
            <div>Tasks: {me.tasks} × {PTS_TASK} = {me.tasks * PTS_TASK}</div>
            <div>Rating (15th & month-end): {me.ratingPoints >= 0 ? "+" : ""}{me.ratingPoints}</div>
            {me.backfills > 0 && <div>Back-dated entries: {me.backfills} × -{BACKFILL_PENALTY} = {-me.backfills * BACKFILL_PENALTY}</div>}
            {me.adjustments !== 0 && <div>Adjustments: {me.adjustments >= 0 ? "+" : ""}{me.adjustments}</div>}
          </div>
        </div>
      )}

      {user?.role === "Owner" && isCurrentMonth && (
        <div style={{ background: C.panel, border: `1px solid ${C.border}`, padding: "16px", marginBottom: "22px" }}>
          <div style={{ color: C.muted, fontSize: "12px", textTransform: "uppercase", marginBottom: "10px", letterSpacing: "1px" }}>Give Points</div>
          <div style={{ display: "flex", gap: "8px", flexWrap: "wrap" }}>
            <select value={giveStaff} onChange={(e) => setGiveStaff(e.target.value)} style={{ background: "#000", color: C.text, border: `1px solid ${C.border}`, padding: "8px", fontFamily: "monospace" }}>
              <option value="">Select staff…</option>
              {rows.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
            </select>
            <input type="number" placeholder="Points (+/-)" value={givePoints} onChange={(e) => setGivePoints(e.target.value)} style={{ background: "#000", color: C.text, border: `1px solid ${C.border}`, padding: "8px", width: "120px", fontFamily: "monospace" }} />
            <input type="text" placeholder="Reason (optional)" value={giveReason} onChange={(e) => setGiveReason(e.target.value)} style={{ background: "#000", color: C.text, border: `1px solid ${C.border}`, padding: "8px", flex: "1", minWidth: "140px", fontFamily: "monospace" }} />
            <button onClick={givePointsFn} disabled={giving} style={{ background: C.accent, color: "#000", border: "none", padding: "8px 16px", fontWeight: "bold", cursor: "pointer", fontFamily: "monospace" }}>{giving ? "..." : "GIVE"}</button>
          </div>
        </div>
      )}

     {rows.length > 0 && (
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: "13px" }}>
            <thead>
              <tr>
                <th style={{ ...th, textAlign: "left" }}>#</th>
                <th style={{ ...th, textAlign: "left" }}>Staff</th>
                <th style={th}>Reports</th>
                <th style={th}>Tasks</th>
                <th style={th}>Rating</th>
                <th style={{ ...th, color: C.accent }}>Points</th>
                <th style={th}>Cash</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((r, i) => (
                <tr key={r.id} style={i === 0 ? { background: "rgba(250,204,21,0.08)" } : {}}>
                  <td style={{ ...td, textAlign: "left", color: C.accent }}>{i + 1}</td>
                  <td style={{ ...td, textAlign: "left" }}>
                    {r.name}{r.off && <span style={{ marginLeft: "8px", background: "#3f3f46", color: "#fbbf24", fontSize: "9px", fontWeight: "bold", padding: "2px 6px", borderRadius: "3px", verticalAlign: "middle" }}>OFF TODAY</span>}{r.missing && <span style={{ marginLeft: "8px", background: "#ef4444", color: "#fff", fontSize: "9px", fontWeight: "bold", padding: "2px 6px", borderRadius: "3px", verticalAlign: "middle" }}>🚩 NOT FILED</span>}
                    <div style={{ color: C.muted, fontSize: "11px" }}>{r.role}</div>
                  </td>
                  <td style={td}>{r.myReports + r.outlets}</td>
                  <td style={td}>{r.tasks}</td>
                 <td style={td}>{r.ratingPoints >= 0 ? "+" : ""}{r.ratingPoints}</td>
                  <td style={{ ...td, color: C.accent, fontWeight: "bold", fontSize: "16px" }}>{r.points}</td>
                 <td style={{ ...td, color: cashColor(r.points, r.id), fontWeight: "bold" }}>{cashLabel(r.points, r.id)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginTop: "26px", color: C.muted, fontSize: "12px", lineHeight: 1.8 }}>
       Daily report +{PTS_DAILY} / late or back-dated -{PTS_DAILY_PENALTY} · Outlet report = {PTS_REPORT} · After cut-off = 0 · Target met = +{PTS_TARGET_MET} / miss = -{PTS_TARGET_MISS} · Task = {PTS_TASK} · Rating weekly: below 4.5 = -{PTS_RATING_FAIL}/wk · maintain 4.5 = +{PTS_RATING} at 15th & month-end
      </div>
    </div>
  );
}
