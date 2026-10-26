"use client";
import { useEffect, useState } from "react";
import { supabase } from "@/lib/supabase";
import supabaseStock from "@/lib/supabaseStock";
import { fetchRealFoodCostPct, REAL_FOOD_COST_WINDOW_DAYS } from "@/lib/realFoodCost";
import { fetchRealCommissionPct, REAL_COMMISSION_WINDOW_DAYS } from "@/lib/realCommission";
import { computeMonthEnd, fetchProjectionRows } from "@/lib/monthEndProjection";

type Staff = { id: string; name: string; role: string; outlets?: string[] };

const OUTLETS = ["royapettah","adayar","bsr_mall","velachery","ra_puram","anna_nagar","pallavaram","vadapalani","besant_nagar","perumbakkam","tambaram","porur"];
const OUTLET_NAMES: Record<string, string> = {
  royapettah: "Royapettah", adayar: "Adyar", bsr_mall: "BSR Mall", velachery: "Velachery",
  ra_puram: "RA Puram", anna_nagar: "Anna Nagar", pallavaram: "Pallavaram", vadapalani: "Vadapalani",
  besant_nagar: "Besant Nagar", perumbakkam: "Perumbakkam", tambaram: "Tambaram", porur: "Porur",
};
const OUTLET_TARGETS: Record<string, number> = {
  royapettah: 80000, adayar: 18333, bsr_mall: 35000, ra_puram: 21667, anna_nagar: 51667,
  porur: 50000, perumbakkam: 13333, tambaram: 20000, velachery: 23333, pallavaram: 23333, vadapalani: 23333, besant_nagar: 11667,
};
const MONTHLY_TARGET = 8750000; // ₹87.5 L company-wide
const DUTY_STAFF = [
  { id: "arun", name: "Arun" }, { id: "nilani", name: "Nilani" }, { id: "vishnu", name: "Vishnu" },
  { id: "ahila", name: "Ahila" }, { id: "gowtham", name: "Gowtham" }
];
const OUTLET_OWNER: Record<string, string> = {
  ra_puram: "ahila", anna_nagar: "vishnu", pallavaram: "ahila", vadapalani: "vishnu",
  velachery: "vishnu", perumbakkam: "vishnu", tambaram: "vishnu", porur: "vishnu",
  royapettah: "stephen", adayar: "ahila", bsr_mall: "ahila", besant_nagar: "ahila",
};
const reviewPoints = (rating: number, valid: boolean) => {
  let p = 0;
  if (rating === 5) p += 5; else if (rating === 4) p += 3; else if (rating >= 1 && rating <= 2) p -= 5;
  if (valid) p -= 10;
  return p;
};
const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
const lakh = (n: number) => "₹" + (n / 100000).toFixed(2) + " L";

  export default function CommandCentre({ user }: { user: Staff }) {
  // India date (IST), so this agrees with the Sales Forecast tab at any hour.
  const today = new Date(Date.now() + 5.5 * 60 * 60 * 1000).toISOString().split("T")[0];
  const _now0 = new Date();
  const _fmtDate = (dd: Date) => `${dd.getFullYear()}-${String(dd.getMonth() + 1).padStart(2, "0")}-${String(dd.getDate()).padStart(2, "0")}`;
  const _defaultDate = _now0.getDate() === 1 ? _fmtDate(new Date(_now0.getFullYear(), _now0.getMonth(), 0)) : today;
  const [date, setDate] = useState(_defaultDate);
  const [out, setOut] = useState<any[]>([]);
  const [month, setMonth] = useState<any[]>([]);
  const [daily, setDaily] = useState<any[]>([]);
  const [offRows, setOffRows] = useState<string[]>([]);
  const [stFixed, setStFixed] = useState<Record<string, any>>({});
  const [stTargets, setStTargets] = useState<Record<string, { a?: number; b?: number }>>({});
  const [stMonthSales, setStMonthSales] = useState<Record<string, { net: number; online: number }>>({});
  // Projected month-end sales from the shared calculation (same as Sales Forecast tab).
  const [projFc, setProjFc] = useState<number | null>(null);
  const [revs, setRevs] = useState<any[]>([]);
  const [payouts, setPayouts] = useState<any[]>([]);
  const [atlasResults, setAtlasResults] = useState<any[]>([]);
  const [openTasks, setOpenTasks] = useState<any[]>([]);
  const [staffNames, setStaffNames] = useState<Record<string, string>>({});
  const [yesterdayRows, setYesterdayRows] = useState<any[]>([]);
  const [prevMonthRows, setPrevMonthRows] = useState<any[]>([]);
  const [monthReviews, setMonthReviews] = useState<any[]>([]);
  const [tableCounts, setTableCounts] = useState<{ name: string; count: number | null; source: string }[]>([]);
  const [realFoodCostPct, setRealFoodCostPct] = useState<number | null>(null);
  const [foodCostWindow, setFoodCostWindow] = useState<{ from: string; to: string } | null>(null);
  const cogsRate = (realFoodCostPct ?? 29.4) / 100;
  // Real, payout-data-backed online commission % (trailing 60 days,
  // company-wide) — replaces the flat 50% assumption. Falls back to 50% if
  // there isn't enough settled Swiggy/Zomato payout data yet in the window.
  const [realCommissionPct, setRealCommissionPct] = useState<number | null>(null);
  const [commissionWindow, setCommissionWindow] = useState<{ from: string; to: string } | null>(null);
  const commissionRate = (realCommissionPct ?? 50) / 100;
  const fmtWin = (w: { from: string; to: string } | null) => w ? `${new Date(w.from + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}–${new Date(w.to + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}` : "";
  const [loading, setLoading] = useState(true);

  const d0 = new Date(date + "T00:00:00");
  const monthStart = `${d0.getFullYear()}-${String(d0.getMonth() + 1).padStart(2, "0")}-01`;
  const daysInMonth = new Date(d0.getFullYear(), d0.getMonth() + 1, 0).getDate();
  const dayOfMonth = d0.getDate();
  // Reports are filed the day AFTER the sales day (report_date = sales day + 1).
  // So month-to-date sales = reports filed from the 2nd of the month up to the selected date.
  const salesDays = Math.max(dayOfMonth - 1, 0);
  const monthFiledFrom = `${monthStart.slice(0, 8)}02`;
  const yDate = new Date(d0.getTime() - 86400000);
  const yDateStr = `${yDate.getFullYear()}-${String(yDate.getMonth() + 1).padStart(2, "0")}-${String(yDate.getDate()).padStart(2, "0")}`;
  const prevMonthDate = new Date(d0.getFullYear(), d0.getMonth() - 1, 1);
  const prevMonthStart = `${prevMonthDate.getFullYear()}-${String(prevMonthDate.getMonth() + 1).padStart(2, "0")}-01`;
  const prevMonthEnd = `${prevMonthDate.getFullYear()}-${String(prevMonthDate.getMonth() + 1).padStart(2, "0")}-${String(new Date(prevMonthDate.getFullYear(), prevMonthDate.getMonth() + 1, 0).getDate()).padStart(2, "0")}`;

  useEffect(() => {
    (async () => {
      setLoading(true);
     const [o, mo, d, r, p, off, st, atlasRaw, tasksRaw, staffRaw, yRaw, pmRaw, mrRaw] = await Promise.all([
        supabase.from("outlet_reports").select("*").eq("report_date", date),
        supabase.from("outlet_reports").select("outlet_id,report_date,shop_sales_value,swiggy_sales_value,zomato_sales_value,swiggy_sales_count,zomato_sales_count").gte("report_date", monthFiledFrom).lte("report_date", date),
        supabase.from("reports").select("staff_id,report_date,is_late,is_backfill,no_points").eq("report_date", date),
        supabase.from("outlet_reviews").select("*").eq("report_date", date),
        supabase.from("outlet_payouts").select("outlet_id,platform,period_start,amount_transferable,net_payout").order("period_start", { ascending: false }).limit(60),
        supabase.from("day_off").select("staff_id").eq("off_date", date),
        supabase.from("sales_target").select("outlet_id,line_items").eq("brand", "BH"),
        supabase.from("atlas_monthly_results").select("*").eq("month", monthStart.slice(0, 7)),
        supabase.from("tasks").select("id,title,description,assigned_to,priority,status,due_at,outlet_id").neq("status", "completed").order("due_at", { ascending: true }),
        supabase.from("staff").select("id,name"),
        supabase.from("outlet_reports").select("outlet_id,shop_sales_value,swiggy_sales_value,zomato_sales_value").eq("report_date", yDateStr),
        supabase.from("outlet_reports").select("outlet_id,shop_sales_value,swiggy_sales_value,zomato_sales_value").gte("report_date", `${prevMonthStart.slice(0, 8)}02`).lte("report_date", monthStart),
        supabase.from("outlet_reviews").select("outlet_id,rating,valid_complaint,refund_given,report_date").gte("report_date", monthStart).lte("report_date", date),
      ]);
      setOut(o.data || []); setMonth(mo.data || []); setDaily(d.data || []); setRevs(r.data || []); setPayouts(p.data || []); setOffRows((off.data || []).map((x: any) => x.staff_id));
      setAtlasResults(atlasRaw.data || []);
      setOpenTasks(tasksRaw.data || []);
      setYesterdayRows(yRaw.data || []);
      setPrevMonthRows(pmRaw.data || []);
      setMonthReviews(mrRaw.data || []);
      const sn: Record<string, string> = {};
      (staffRaw.data || []).forEach((s: any) => { sn[s.id] = s.name; });
      setStaffNames(sn);
      const fm: Record<string, any> = {}; const sm: Record<string, { net: number; online: number }> = {};
      const tm: Record<string, { a?: number; b?: number }> = {};
      (st.data || []).forEach((row: any) => { fm[row.outlet_id] = (row.line_items || {}).fixed || {}; tm[row.outlet_id] = (row.line_items || {}).targets || {}; });
      (mo.data || []).forEach((row: any) => {
        const oid = row.outlet_id;
        if (!sm[oid]) sm[oid] = { net: 0, online: 0 };
        sm[oid].net += Number(row.shop_sales_value) || 0;
        sm[oid].online += (Number(row.swiggy_sales_value) || 0) + (Number(row.zomato_sales_value) || 0);
      });
      setStFixed(fm); setStMonthSales(sm); setStTargets(tm);
      try {
        const _projRows = await fetchProjectionRows(supabase, date);
        setProjFc(computeMonthEnd(_projRows, OUTLETS, date).projected);
      } catch (e) { console.error("month-end projection failed", e); setProjFc(null); }
      setLoading(false);
    })();
  }, [date, monthStart]);

  // Data volume check — not part of the main dashboard data, just a proactive
  // heads-up so a growing table's row count is visible before it silently
  // causes the same kind of truncation bug Purchase & Vendors and Outlet
  // Health hit (Supabase caps any single query at ~1000 rows by default).
  // This only counts rows — it does not know whether any specific query
  // against a table is protected with batching or not; that's still a coding
  // discipline, not something a dashboard widget can verify on its own.
  useEffect(() => {
    (async () => {
      const [orCount, taskCount, revCount, payoutCount, plCount] = await Promise.all([
        supabase.from("outlet_reports").select("id", { count: "exact", head: true }),
        supabase.from("tasks").select("id", { count: "exact", head: true }),
        supabase.from("outlet_reviews").select("id", { count: "exact", head: true }),
        supabase.from("outlet_payouts").select("id", { count: "exact", head: true }),
        supabaseStock.from("purchase_ledger").select("id", { count: "exact", head: true }),
      ]);
      setTableCounts([
        { name: "outlet_reports", count: orCount.count, source: "TASKFORCE" },
        { name: "tasks", count: taskCount.count, source: "TASKFORCE" },
        { name: "outlet_reviews", count: revCount.count, source: "TASKFORCE" },
        { name: "outlet_payouts", count: payoutCount.count, source: "TASKFORCE" },
        { name: "purchase_ledger", count: plCount.count, source: "Production Stock" },
      ]);
    })();
  }, []);

  // Real food cost % (trailing 30 days, company-wide) — replaces the flat
  // 29.4% assumption previously used in the modules below. Same shared
  // calculation as Outlet P&L and Channel P&L, so all three stay consistent.
  useEffect(() => {
    fetchRealFoodCostPct().then(({ pct, windowFrom, windowTo }) => { setRealFoodCostPct(pct); setFoodCostWindow({ from: windowFrom, to: windowTo }); }).catch((err) => console.error("real food cost fetch failed", err));
  }, []);
  useEffect(() => {
    fetchRealCommissionPct().then(({ pct, windowFrom, windowTo }) => { setRealCommissionPct(pct); setCommissionWindow({ from: windowFrom, to: windowTo }); }).catch((err) => console.error("real commission fetch failed", err));
  }, []);

  const n = (v: any) => Number(v) || 0;
  const sum = (rows: any[], k: string) => rows.reduce((s, r) => s + n(r[k]), 0);

  // today
  const tShop = sum(out, "shop_sales_value"), tSw = sum(out, "swiggy_sales_value"), tZo = sum(out, "zomato_sales_value");
  const tTotal = tShop + tSw + tZo;
  const swC = sum(out, "swiggy_sales_count"), zoC = sum(out, "zomato_sales_count");

  // Module 1: "Sales, gross margin, contribution and estimated EBITDA for today,
  // yesterday and month-to-date" — same cost model as Outlet P&L (29.4% COGS, 5%
  // wastage, 50% online commission), applied company-wide. EBITDA here means
  // contribution minus fixed costs — this business has no interest/depreciation
  // line to separate out, so EBITDA and "net profit" are effectively the same
  // figure; shown as EBITDA to match his terminology.
  const companyFinancials = (totalSales: number, onlineSales: number, fixedTotal: number) => {
    const grossMargin = totalSales * (1 - cogsRate);
    const commission = commissionRate * onlineSales;
    const contribution = totalSales - totalSales * cogsRate - totalSales * 0.05 - commission;
    const ebitda = contribution - fixedTotal;
    return { sales: totalSales, grossMargin, grossMarginPct: totalSales > 0 ? (grossMargin / totalSales) * 100 : 0, contribution, ebitda };
  };
  const totalFixedAllOutlets = OUTLETS.reduce((s, o) => {
    const f = stFixed[o] || {};
    const _abs = (v: any) => Math.abs(Number(v) || 0);
    return s + _abs(f.staff) + _abs(f.rent) + _abs(f.eb) + _abs(f.transport) + 0.2 * _abs(f.rent) + _abs(f.pest) + _abs(f.water) + _abs(f.airtel);
  }, 0);
  const dailyFixedShare = daysInMonth > 0 ? totalFixedAllOutlets / daysInMonth : 0;
  const yShop = sum(yesterdayRows, "shop_sales_value"), ySw = sum(yesterdayRows, "swiggy_sales_value"), yZo = sum(yesterdayRows, "zomato_sales_value");
  const yTotal = yShop + ySw + yZo, yOnline = ySw + yZo;
  const todayFin = companyFinancials(tTotal, tSw + tZo, dailyFixedShare);
  const yesterdayFin = companyFinancials(yTotal, yOnline, dailyFixedShare);

  // month to date
  const _stVals = Object.values(stMonthSales);
  const mShop = _stVals.reduce((s, v) => s + (v.net || 0), 0);
  const mOnline = _stVals.reduce((s, v) => s + (v.online || 0), 0);
  const mtd = mShop + mOnline;
  // Fixed costs are monthly, but MTD sales only cover the days elapsed — so charge the same share of fixed costs.
  const fixedFrac = daysInMonth > 0 ? salesDays / daysInMonth : 0;
  const mtdFin = companyFinancials(mtd, mOnline, totalFixedAllOutlets * fixedFrac);
  const runRate = salesDays > 0 ? mtd / salesDays : 0;
  // Same shared calculation as the Sales Forecast tab. Falls back to the simple daily-average only if it could not be loaded.
  const projected = projFc ?? runRate * daysInMonth;
  const daysLeft = Math.max(daysInMonth - salesDays, 0);
  const required = daysLeft > 0 ? (MONTHLY_TARGET - mtd) / daysLeft : 0;
  const targetPct = MONTHLY_TARGET > 0 ? (mtd / MONTHLY_TARGET) * 100 : 0;
  const shortfall = projected - MONTHLY_TARGET;
  const onTrack = projected >= MONTHLY_TARGET;
  const offlineRatio = mShop > 0 ? (mOnline / mShop) : 0;

  const pnl = OUTLETS.map(o => {
    const rows = month.filter((r: any) => r.outlet_id === o);
    const _ms = stMonthSales[o] || { net: 0, online: 0 };
    const oNet = _ms.net;
    const oOnline = _ms.online;
    const f = stFixed[o] || {};
    const _abs = (v: any) => Math.abs(Number(v) || 0);
    const fixed = (_abs(f.staff) + _abs(f.rent) + _abs(f.eb) + _abs(f.transport) + 0.2 * _abs(f.rent) + _abs(f.pest) + _abs(f.water) + _abs(f.airtel)) * fixedFrac;
    const oTotal = oNet + oOnline;
    const comm = commissionRate * oOnline;
    const contribution = oTotal - cogsRate * oTotal - 0.05 * oTotal - comm;
    const netProfit = contribution - fixed;
    return { o, name: OUTLET_NAMES[o] || o, net: oNet, online: oOnline, fixed, comm, contribution, netProfit, reported: rows.length };
  });
  // Profit budget vs actual — "Target A" is the monthly profit goal you
  // already set per outlet in Outlet P&L (it's currently only used to work
  // out required sales); this compares it against actual profit made so
  // far this month. Same-period-last-year isn't included — the app doesn't
  // have a full year of history yet, so that comparison isn't real.
  const budgetRows = pnl.map((p) => {
    const budget = Number(stTargets[p.o]?.a) || 0;
    const hasBudget = budget > 0;
    const variance = p.netProfit - budget;
    const variancePct = hasBudget ? (variance / budget) * 100 : null;
    return { ...p, budget, hasBudget, variance, variancePct };
  }).filter((p) => p.hasBudget);
  const budgetTotalActual = budgetRows.reduce((s, p) => s + p.netProfit, 0);
  const budgetTotalPlanned = budgetRows.reduce((s, p) => s + p.budget, 0);
  const _EXCLUDE: string[] = [];
  const _hasCosts = (o: string) => { const f = stFixed[o] || {}; return (Math.abs(Number(f.rent) || 0) + Math.abs(Number(f.staff) || 0) + Math.abs(Number(f.pest) || 0) + Math.abs(Number(f.airtel) || 0)) > 0; };
  const _complete = pnl.filter(p => !_EXCLUDE.includes(p.o) && _hasCosts(p.o) && (p.net > 0 || p.online > 0));
  const _incompleteCount = pnl.filter(p => !_EXCLUDE.includes(p.o) && (p.net > 0 || p.online > 0) && !_hasCosts(p.o)).length;
  const totalProfit = _complete.reduce((s, p) => s + p.netProfit, 0);
  const bleeders = _complete.filter(p => p.netProfit < 0).sort((a, b) => a.netProfit - b.netProfit);
  const worstPnl = bleeders[0];
  const noFixedCount = OUTLETS.filter(o => { const f = stFixed[o] || {}; return !((Number(f.staff) || 0) + (Number(f.rent) || 0) + (Number(f.pest) || 0)); }).length;
  // These 3 outlets are 100% online cloud kitchens — no dine-in counter exists,
  // so "too online-dependent" isn't a problem to fix, it's the business model.
  // Keep this list in sync with the one in the Auto Reviews page (page.tsx).
  const CLOUD_KITCHENS = ["Pallavaram", "Velachery", "Vadapalani"];
  const isCloudKitchen = (p: any) => !!p && CLOUD_KITCHENS.includes(p.name);
  const whyBleed = (p: any) => {
    if (!p) return "";
    if (p.comm > p.contribution + p.fixed) {
      return isCloudKitchen(p)
        ? `aggregator commission (${(commissionRate * 100).toFixed(0)}% on online) is eating most of the margin — expected for a cloud kitchen, but it means volume/pricing has to work harder to cover it.`
        : `aggregator commission (${(commissionRate * 100).toFixed(0)}% on online) is the killer — too online-dependent.`;
    }
    if (p.fixed > p.contribution) return "fixed costs (rent/staff) outweigh what sales bring in — rent is high or sales too low to cover it.";
    return "sales are simply too low this month to cover its costs.";
  };
  const fixBleed = (p: any) => {
    if (!p) return "";
    if (p.comm > p.contribution + p.fixed) {
      return isCloudKitchen(p)
        ? "No dine-in to shift to here (cloud kitchen) — push order volume and AOV (combos, upsells) to outrun the commission, and see if the platform commission slab can be renegotiated."
        : "Shift mix toward dine-in/takeaway (commission-free) and cut discounting on the apps.";
    }
    if (p.fixed > p.contribution) return "Drive volume hard (footfall + online) to cover fixed costs, or review the cost base for that site.";
    return "Push both channels — promotions, visibility, counter upsell — to lift the topline.";
  };

  // Module 2: "Outlet ranking based on sales growth, profitability, food cost,
  // customer ratings and operational compliance." Food cost is deliberately
  // shown as informational only, not ranked — it's currently a flat 29.4%
  // assumption applied identically to every outlet, so ranking on it would be
  // meaningless (every outlet would tie). The other four are real, computed
  // per-outlet figures.
  const prevMonthByOutlet: Record<string, number> = {};
  prevMonthRows.forEach((r: any) => { const t = n(r.shop_sales_value) + n(r.swiggy_sales_value) + n(r.zomato_sales_value); prevMonthByOutlet[r.outlet_id] = (prevMonthByOutlet[r.outlet_id] || 0) + t; });
  const filedDaysByOutlet: Record<string, number> = {};
  month.forEach((r: any) => { filedDaysByOutlet[r.outlet_id] = (filedDaysByOutlet[r.outlet_id] || 0) + 1; });
  const ratingsByOutlet: Record<string, { sum: number; count: number }> = {};
  monthReviews.forEach((r: any) => { if (!ratingsByOutlet[r.outlet_id]) ratingsByOutlet[r.outlet_id] = { sum: 0, count: 0 }; ratingsByOutlet[r.outlet_id].sum += n(r.rating); ratingsByOutlet[r.outlet_id].count += 1; });

  const outletRanking = pnl.map(p => {
    const prevTotal = prevMonthByOutlet[p.o] || 0;
    const curTotal = p.net + p.online;
    const growthPct = prevTotal > 0 ? ((curTotal - prevTotal) / prevTotal) * 100 : null;
    const profitabilityPct = curTotal > 0 ? (p.netProfit / curTotal) * 100 : null;
    const rt = ratingsByOutlet[p.o];
    const avgRating = rt && rt.count > 0 ? rt.sum / rt.count : null;
    const compliancePct = salesDays > 0 ? Math.min(100, ((filedDaysByOutlet[p.o] || 0) / salesDays) * 100) : 0;
    return { ...p, growthPct, profitabilityPct, avgRating, ratingCount: rt?.count || 0, compliancePct };
  });

  // Module 3: "Top revenue opportunities, profit leakages and operational risks
  // ranked by financial impact" — three distinct lists, each sorted by ₹ size.
  const opportunities = [...outletRanking].filter(p => p.netProfit > 0).sort((a, b) => b.netProfit - a.netProfit).slice(0, 3);
  const leakages = [...outletRanking].filter(p => p.netProfit < 0).sort((a, b) => a.netProfit - b.netProfit).slice(0, 5);
  const complianceRisks = outletRanking.filter(p => p.compliancePct < 80).sort((a, b) => a.compliancePct - b.compliancePct);

  const filedDaily = new Set(daily.filter(x => !x.no_points).map(x => x.staff_id));

  // Module 4: "Real-time alerts for sales decline, stock-outs, high wastage,
  // staffing problems and unresolved complaints." Sales decline and unresolved
  // complaints are real, computed alerts. Stock-outs, wastage and staffing have
  // no data source yet (need Inventory & Food Cost and People & Labour, not yet
  // built) — shown honestly as locked, not faked with placeholder numbers.
  const salesDeclineAlerts = OUTLETS.map(o => {
    const r = out.find((x: any) => x.outlet_id === o);
    const tot = r ? n(r.shop_sales_value) + n(r.swiggy_sales_value) + n(r.zomato_sales_value) : 0;
    const tgt = (r && n(r.target)) || OUTLET_TARGETS[o] || 0;
    const pct = tgt > 0 ? (tot / tgt) * 100 : 100;
    return { o, name: OUTLET_NAMES[o] || o, tot, tgt, pct, filed: !!r };
  }).filter(x => x.filed && x.pct < 50).sort((a, b) => a.pct - b.pct);
  const unresolvedComplaints = monthReviews.filter((r: any) => r.valid_complaint && !r.refund_given);
  const filedOutlets = new Set(out.map(r => r.outlet_id));

  const pts: Record<string, number> = {};
  daily.forEach(x => { if (x.no_points) return; pts[x.staff_id] = (pts[x.staff_id] || 0) + (x.is_backfill || x.is_late ? -5 : 10); });
  out.forEach(r => {
    const owner = r.outlet_id === "royapettah" ? (r.staff_id || OUTLET_OWNER[r.outlet_id]) : OUTLET_OWNER[r.outlet_id]; if (!owner) return;
    let p = 0;
    if (r.no_points) p = 0; else if (r.is_backfill) p = -30; else if (r.is_late) p = 0;
    else { const t = n(r.target) || OUTLET_TARGETS[r.outlet_id]; const tot = n(r.shop_sales_value) + n(r.swiggy_sales_value) + n(r.zomato_sales_value); p = 20 + (t > 0 && tot >= t ? 30 : 0); }
    pts[owner] = (pts[owner] || 0) + p;
  });
  revs.forEach(rv => { const owner = rv.outlet_id === "royapettah" ? (rv.staff_id || OUTLET_OWNER[rv.outlet_id]) : (OUTLET_OWNER[rv.outlet_id] || rv.staff_id); pts[owner] = (pts[owner] || 0) + reviewPoints(n(rv.rating), rv.valid_complaint); });

  const latestPay: Record<string, any> = {};
  payouts.forEach(p => { const k = p.outlet_id + "_" + p.platform; if (!latestPay[k]) latestPay[k] = p; });

  const reportedByOutletMtd: Record<string, number> = {};
  (month as any[]).forEach((r: any) => {
    const g = (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
    reportedByOutletMtd[r.outlet_id] = (reportedByOutletMtd[r.outlet_id] || 0) + g;
  });
  const atlasMap: Record<string, any> = {};
  atlasResults.forEach((r: any) => { atlasMap[r.outlet_id] = r; });
  const ATLAS_THRESHOLD = 500;
  const atlasReconRows = OUTLETS.map(o => {
    const ar = atlasMap[o];
    const rg = reportedByOutletMtd[o] || 0;
    const ag: number | null = ar ? (Number(ar.atlas_gross) || 0) : null;
    const an: number | null = ar ? (Number(ar.atlas_net) || 0) : null;
    const al: number | null = ar ? (Number(ar.atlas_lost) || 0) : null;
    const diff: number | null = ag !== null ? ag - rg : null;
    return { outlet_id: o, atlasGross: ag, atlasNet: an, atlasLost: al, reportedGross: rg, diff };
  });

 const loadHtml2Pdf = (): Promise<any> => new Promise((resolve, reject) => {
    const w = window as any;
    if (w.html2pdf) return resolve(w.html2pdf);
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js";
    s.onload = () => resolve((window as any).html2pdf);
    s.onerror = () => reject(new Error("pdf lib failed"));
    document.body.appendChild(s);
  });
  const rsF = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
  const lk = (n: number) => "₹" + (n / 100000).toFixed(2) + "L";
const downloadPDF = async () => {
    let html2pdf: any;
    try { html2pdf = await loadHtml2Pdf(); } catch { alert("Could not load the PDF tool — check your connection and retry."); return; }
    const making = totalProfit >= 0;
    const sorted = [...pnl].sort((a, b) => b.netProfit - a.netProfit);
    const maxTot = Math.max(...sorted.map(p => p.net + p.online), 1);
    const dateStr = new Date(date + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" });
    const sd = new Date(date + "T00:00:00"); sd.setDate(sd.getDate() - 1);
    const sdStr = sd.toLocaleDateString("en-IN", { day: "numeric", month: "short" });

    const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", brown: "#5C3A22", soft: "#8A6A4A", gold: "#C8901E", gold2: "#E0A52E", line: "#EADBC2", green: "#2E7D32", red: "#C62828" };

    const rowHtml = sorted.map(p => {
      const tot = p.net + p.online;
      const has = (p.net > 0 || p.online > 0);
      const dot = !has ? "⚪️" : (p.netProfit >= 0 ? "🟢" : "🔴");
      const bar = Math.min(100, (tot / maxTot) * 100);
      const col = p.netProfit >= 0 ? C.green : C.red;
      return `<tr>
        <td style="padding:8px 12px;border-bottom:1px solid ${C.line};font-weight:600;color:${C.ink};font-size:12px">${dot}&nbsp;${p.name}</td>
        <td style="padding:8px 12px;border-bottom:1px solid ${C.line}">
          <div style="font-size:11px;color:${C.soft};margin-bottom:3px">${rsF(tot)}</div>
          <div style="height:7px;background:${C.line};border-radius:4px;overflow:hidden"><div style="height:100%;width:${bar}%;background:linear-gradient(90deg,${C.gold2},${C.gold})"></div></div>
        </td>
        <td style="padding:8px 12px;border-bottom:1px solid ${C.line};text-align:right;font-weight:800;color:${col};font-size:12px;white-space:nowrap">${p.netProfit >= 0 ? "+" : ""}${rsF(p.netProfit)}</td>
      </tr>`;
    }).join("");

    const kpi = (emoji: string, label: string, value: string, sub: string) => `
      <div style="flex:1;background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:16px 18px">
        <div style="font-size:11px;letter-spacing:1px;text-transform:uppercase;color:${C.soft};margin-bottom:6px">${emoji}&nbsp;${label}</div>
        <div style="font-size:24px;font-weight:800;color:${C.ink};line-height:1">${value}</div>
        <div style="font-size:11px;color:${C.soft};margin-top:5px">${sub}</div>
      </div>`;

    const drain = worstPnl ? `
      <div style="background:#FBEAE7;border:1px solid #E8C0B8;border-radius:14px;padding:16px 18px;margin-top:14px">
        <div style="font-size:13px;font-weight:800;color:${C.red};margin-bottom:6px">🩸 Biggest drain — ${worstPnl.name} (${rsF(worstPnl.netProfit)})</div>
        <div style="font-size:12px;color:${C.brown};margin-bottom:4px"><b>Why:</b> ${whyBleed(worstPnl)}</div>
        <div style="font-size:12px;color:${C.gold}"><b>✅ Fix:</b> ${fixBleed(worstPnl)}</div>
      </div>` : "";

    const offPct = mtd > 0 ? (mShop / mtd) * 100 : 0;
    const onPct = 100 - offPct;
    const tgtPct = Math.min(100, targetPct);

    const html = `
    <div style="width:794px;background:${C.bg};font-family:'Segoe UI',Helvetica,Arial,sans-serif;color:${C.ink};box-sizing:border-box">
      <div style="background:linear-gradient(135deg,${C.ink},${C.brown});padding:22px 32px;display:flex;justify-content:space-between;align-items:center">
        <div>
          <div style="font-size:22px;font-weight:800;color:#FFF6E5">🍫 Brownie Heaven</div>
          <div style="font-size:12px;color:${C.gold2};letter-spacing:2px;text-transform:uppercase">Money Report — are we making or losing?</div>
        </div>
        <div style="text-align:right;color:#E8D5BC;font-size:11px">📅 ${dateStr}<br><span style="color:#B89A78">Sales data: ${sdStr} (prev. day)</span></div>
      </div>

      <div style="padding:24px 32px">
        <div style="background:${making ? "#E9F5EA" : "#FBEAE7"};border:2px solid ${making ? C.green : C.red};border-radius:18px;padding:22px 26px;text-align:center;margin-bottom:18px">
          <div style="font-size:13px;letter-spacing:1px;text-transform:uppercase;color:${C.soft}">Profit / Loss · Month to date</div>
          <div style="font-size:40px;font-weight:900;color:${making ? C.green : C.red};margin:6px 0">${making ? "🎉 Making" : "⚠️ Losing"} ${rsF(Math.abs(totalProfit))}</div>
          <div style="font-size:12px;color:${C.soft}">net this month · based on ${_complete.length} of 12 outlets with full cost data${_incompleteCount > 0 ? ` · ${_incompleteCount} pending` : ""}</div>
        </div>

        <div style="display:flex;gap:14px;margin-bottom:18px">
          ${kpi("📅", "Today's sales", rsF(tTotal), `${dayOfMonth}/${daysInMonth} reported`)}
          ${kpi("💰", "Month to date", lk(mtd), `${salesDays} days`)}
          ${kpi("📈", "Projected end", lk(projected), onTrack ? "on target" : "below target")}
        </div>

        <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:16px 18px;margin-bottom:14px">
          <div style="display:flex;justify-content:space-between;font-size:12px;color:${C.brown};margin-bottom:8px"><span>🎯 Achieved ${lk(mtd)}</span><span>Target ${lk(MONTHLY_TARGET)}</span></div>
          <div style="height:14px;background:${C.line};border-radius:8px;overflow:hidden"><div style="height:100%;width:${tgtPct}%;background:linear-gradient(90deg,${C.gold2},${C.gold})"></div></div>
          <div style="font-size:11px;color:${C.soft};margin-top:7px">${targetPct.toFixed(1)}% of target · ${lk(Math.max(MONTHLY_TARGET - mtd, 0))} to go · ⚡ ${rsF(runRate)}/day current pace</div>
        </div>

        ${drain}

        <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:16px 18px;margin-top:14px">
          <div style="font-size:12px;font-weight:700;color:${C.ink};margin-bottom:8px">📱 Channel mix · 🛍️ Offline ${offPct.toFixed(0)}% &nbsp;vs&nbsp; 📱 Online ${onPct.toFixed(0)}%</div>
          <div style="height:14px;border-radius:8px;overflow:hidden;display:flex">
            <div style="width:${offPct}%;background:${C.gold}"></div><div style="width:${onPct}%;background:${C.ink}"></div>
          </div>
          <div style="font-size:11px;color:${C.soft};margin-top:7px">For every ₹1 walk-in, ₹${offlineRatio > 0 ? (1 / offlineRatio).toFixed(1) : "—"} comes from online (${(commissionRate * 100).toFixed(0)}% app commission territory, real, this + last calendar month).</div>
        </div>

        <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:18px 20px;margin-top:14px;text-align:center">
          <div style="font-size:12px;font-weight:700;color:${C.ink};margin-bottom:10px">🎯 Monthly target gauge</div>
          <div style="font-size:30px;font-weight:900;color:${onTrack ? C.green : C.gold}">${targetPct.toFixed(0)}%</div>
          <div style="height:16px;background:${C.line};border-radius:9px;overflow:hidden;margin:10px 0"><div style="height:100%;width:${tgtPct}%;background:linear-gradient(90deg,${C.gold2},${C.gold})"></div></div>
          <div style="font-size:11px;color:${C.soft}">${lk(mtd)} of ${lk(MONTHLY_TARGET)} · projected ${lk(projected)} ${onTrack ? "✅ on track" : "⚠️ below target"}</div>
        </div>
      </div>

      <div style="padding:8px 32px 28px;page-break-before:always">
        <div style="font-size:16px;font-weight:800;color:${C.ink};margin:8px 0 12px">📊 Profit / Loss by outlet</div>
        <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:14px;overflow:hidden">
          <thead><tr style="background:${C.ink}">
            <th style="padding:9px 12px;text-align:left;color:#FFF6E5;font-size:11px;letter-spacing:1px">OUTLET</th>
            <th style="padding:9px 12px;text-align:left;color:#FFF6E5;font-size:11px;letter-spacing:1px">SALES (MTD)</th>
            <th style="padding:9px 12px;text-align:right;color:#FFF6E5;font-size:11px;letter-spacing:1px">NET P/L</th>
          </tr></thead>
          <tbody>${rowHtml}</tbody>
        </table>
        <div style="text-align:center;background:${making ? "#E9F5EA" : "#FBEAE7"};border:1px solid ${making ? C.green : C.red};border-radius:12px;padding:14px;margin-top:14px;font-size:16px;font-weight:800;color:${making ? C.green : C.red}">
          ${making ? "🎉" : "⚠️"} TOTAL: ${making ? "Making" : "Losing"} ${rsF(Math.abs(totalProfit))} net this month
        </div>
       <div style="font-size:16px;font-weight:800;color:${C.ink};margin:22px 0 12px">🎯 Sales vs target by outlet (month)</div>
        <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:14px 18px">
          ${sorted.map(p => {
            const tot = p.net + p.online;
            const mt = (OUTLET_TARGETS[p.o] || 0) * daysInMonth;
            const pct = mt > 0 ? Math.min(120, (tot / mt) * 100) : 0;
            const hit = mt > 0 && tot >= mt;
            const barCol = hit ? C.green : (pct >= 60 ? C.gold : C.red);
            return `<div style="margin-bottom:11px">
              <div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:3px"><span style="color:${C.ink};font-weight:600">${hit ? "✅" : ""} ${p.name}</span><span style="color:${C.soft}">${rsF(tot)} / ${mt > 0 ? rsF(mt) : "—"}${mt > 0 ? ` · ${((tot / mt) * 100).toFixed(0)}%` : ""}</span></div>
              <div style="height:9px;background:${C.line};border-radius:5px;overflow:hidden"><div style="height:100%;width:${Math.min(100, pct)}%;background:${barCol}"></div></div>
            </div>`;
          }).join("")}
        </div>
      <div style="font-size:16px;font-weight:800;color:${C.ink};margin:22px 0 12px">📆 Per-day — ${sdStr} actual vs daily target</div>
        <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:14px 18px">
          ${OUTLETS.map((o: string) => {
            const r = out.find((x: any) => x.outlet_id === o);
            const day = r ? (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0) : 0;
            const dt = (r && Number(r.target)) ? Number(r.target) : (OUTLET_TARGETS[o] || 0);
            const filed = !!r;
            const pct = dt > 0 ? Math.min(120, (day / dt) * 100) : 0;
            const hit = dt > 0 && day >= dt;
            const barCol = !filed ? C.line : (hit ? C.green : (pct >= 60 ? C.gold : C.red));
            return `<div style="margin-bottom:11px">
              <div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:3px"><span style="color:${C.ink};font-weight:600">${!filed ? "⚪️" : (hit ? "✅" : "🔴")} ${OUTLET_NAMES[o] || o}</span><span style="color:${C.soft}">${filed ? rsF(day) : "not filed"}${dt > 0 ? ` / ${rsF(dt)}${filed ? ` · ${((day / dt) * 100).toFixed(0)}%` : ""}` : ""}</span></div>
              <div style="height:9px;background:${C.line};border-radius:5px;overflow:hidden"><div style="height:100%;width:${Math.min(100, pct)}%;background:${barCol}"></div></div>
            </div>`;
          }).join("")}
        </div>
      ${(() => {
          const ra = atlasReconRows.filter((r: any) => r.atlasGross !== null);
          if (ra.length === 0) return "";
          const tAG = ra.reduce((s: number, r: any) => s + (r.atlasGross || 0), 0);
          const tRG = ra.reduce((s: number, r: any) => s + (r.reportedGross || 0), 0);
          const tAN = ra.reduce((s: number, r: any) => s + (r.atlasNet || 0), 0);
          const tAL = ra.reduce((s: number, r: any) => s + (r.atlasLost || 0), 0);
          const rows = ra.map((r: any) => {
            const gap = (r.atlasGross || 0) - (r.reportedGross || 0);
            const flag = Math.abs(gap) > 500;
            return `<tr>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};font-weight:600;color:${C.ink};font-size:11px">${flag ? "🔴" : "🟢"} ${OUTLET_NAMES[r.outlet_id] || r.outlet_id}</td>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:${C.brown}">${rsF(r.atlasGross || 0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:${C.soft}">${rsF(r.reportedGross || 0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;font-weight:700;color:${flag ? C.red : C.green}">${gap >= 0 ? "+" : ""}${rsF(gap)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:#1565C0">${rsF(r.atlasNet || 0)}</td>
              <td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:${C.red}">${(r.atlasLost || 0) > 0 ? rsF(r.atlasLost) : "—"}</td>
            </tr>`;
          }).join("");
          return `
        <div style="font-size:16px;font-weight:800;color:${C.ink};margin:22px 0 12px;page-break-before:always">🔍 Atlas reconciliation — honesty check</div>
        <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:14px;overflow:hidden">
          <thead><tr style="background:${C.ink}">
            <th style="padding:9px 10px;text-align:left;color:#FFF6E5;font-size:10px;letter-spacing:.5px">OUTLET</th>
            <th style="padding:9px 10px;text-align:right;color:#FFF6E5;font-size:10px;letter-spacing:.5px">ATLAS GROSS</th>
            <th style="padding:9px 10px;text-align:right;color:#FFF6E5;font-size:10px;letter-spacing:.5px">STAFF GROSS</th>
            <th style="padding:9px 10px;text-align:right;color:#FFF6E5;font-size:10px;letter-spacing:.5px">GAP</th>
            <th style="padding:9px 10px;text-align:right;color:#FFF6E5;font-size:10px;letter-spacing:.5px">ATLAS NET</th>
            <th style="padding:9px 10px;text-align:right;color:#FFF6E5;font-size:10px;letter-spacing:.5px">LOST REV.</th>
          </tr></thead>
          <tbody>${rows}
            <tr style="background:#FFF4DD">
              <td style="padding:9px 10px;font-weight:800;color:${C.ink};font-size:11px">TOTAL</td>
              <td style="padding:9px 10px;text-align:right;font-weight:800;color:${C.ink};font-size:11px">${rsF(tAG)}</td>
              <td style="padding:9px 10px;text-align:right;font-weight:700;color:${C.soft};font-size:11px">${rsF(tRG)}</td>
              <td style="padding:9px 10px;text-align:right;font-weight:800;color:${Math.abs(tAG - tRG) > 500 ? C.red : C.green};font-size:11px">${(tAG - tRG) >= 0 ? "+" : ""}${rsF(tAG - tRG)}</td>
              <td style="padding:9px 10px;text-align:right;font-weight:800;color:#1565C0;font-size:11px">${rsF(tAN)}</td>
              <td style="padding:9px 10px;text-align:right;font-weight:800;color:${C.red};font-size:11px">${rsF(tAL)}</td>
            </tr>
          </tbody>
        </table>
        <div style="font-size:10px;color:${C.soft};margin-top:8px">🔴 = gap over ₹500 between Atlas and staff-reported gross · Atlas Net = founder-verified revenue · Lost Rev. = platform cancellations.</div>`;
        })()}
        <div style="text-align:center;font-size:10px;color:${C.soft};margin-top:18px">🍫 Brownie Heaven · Generated ${dateStr} · 🟢 profit · 🔴 loss · ⚪️ no data yet · monthly target = daily × ${daysInMonth} · per-day = ${sdStr} actual</div>
      </div>
    </div>`;
         const holder = document.createElement("div");
    holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0";
    holder.innerHTML = html;
    document.body.appendChild(holder);
    try {
      await html2pdf().set({
        margin: 0,
        filename: "BrownieHeaven_Report_" + date + ".pdf",
        image: { type: "jpeg", quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true, backgroundColor: C.bg },
        jsPDF: { unit: "pt", format: "a4", orientation: "portrait" },
        pagebreak: { mode: ["css", "legacy"] },
      }).from(holder.firstElementChild).save();
    } finally {
      document.body.removeChild(holder);
    }
  };
  const Hero = ({ label, value, sub, accent }: any) => (
    <div className="flex-1 min-w-[150px] bg-gradient-to-b from-surface to-surface border border-line p-5">
      <p className="text-[10px] font-mono text-text-muted uppercase tracking-widest mb-2">{label}</p>
      <p className={`text-2xl md:text-3xl font-black ${accent || "text-text"}`}>{value}</p>
      {sub && <p className="text-[11px] text-text-muted mt-1">{sub}</p>}
    </div>
  );
  const Card = ({ title, children, right }: any) => (
    <div className="bg-surface border border-line p-5 mb-5">
      <div className="flex justify-between items-center mb-4"><p className="text-[10px] font-mono text-text-muted uppercase tracking-widest">{title}</p>{right}</div>
      {children}
    </div>
  );

  const dayLabel = d0.toLocaleDateString("en-IN", { weekday: "long", day: "numeric", month: "long", year: "numeric" });

  return (
    <div>
      <div className="flex flex-wrap justify-between items-end gap-3 mb-6">
        <div>
          <h2 className="text-2xl md:text-3xl font-black tracking-tight">Command Centre</h2>
          <p className="text-[11px] font-mono text-text-muted uppercase tracking-widest mt-1">{dayLabel}</p>
        </div>
       <div className="flex items-center gap-2">
          <button onClick={downloadPDF} className="bg-accent text-accent-ink font-bold text-[10px] px-4 py-2.5 uppercase tracking-widest hover:opacity-90">📄 Download PDF</button>
          <input type="date" max={today} value={date} onChange={e => setDate(e.target.value)} className="bg-canvas border border-line text-text px-3 py-2 focus:outline-none focus:border-accent text-sm font-mono" />
        </div>
      </div>

      <p className="text-[11px] font-mono text-text-muted mb-5 border-l-2 border-accent/40 pl-3">📊 Outlet sales reflect the <span className="text-accent">previous day&apos;s</span> business — figures below are {new Date(d0.getTime() - 86400000).toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}&apos;s sales, filed this morning.</p>

      {loading ? <p className="text-text-faint font-mono text-sm">Loading…</p> : (
        <>
          <div className="flex flex-wrap gap-3 mb-5">
            <Hero label="Today's sales" value={inr(tTotal)} sub={`${out.length}/12 reported`} accent="text-accent" />
            <Hero label="Month to date" value={lakh(mtd)} sub={`${salesDays} days`} />
            <Hero label="Projected month-end" value={lakh(projected)} sub={onTrack ? "on track" : "below target"} accent={onTrack ? "text-green-400" : "text-red-400"} />
          </div>

          <Card title="1. Sales, gross margin, contribution &amp; EBITDA">
            <div className="overflow-x-auto">
              <table className="w-full text-xs font-mono whitespace-nowrap">
                <thead>
                  <tr className="text-text-muted uppercase tracking-widest text-[10px] border-b border-line">
                    <th className="text-left py-2 pr-3">Metric</th>
                    <th className="text-right py-2 pl-3">Today</th>
                    <th className="text-right py-2 pl-3">Yesterday</th>
                    <th className="text-right py-2 pl-3">Month-to-date</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="border-b border-line/40"><td className="py-2 pr-3 text-text">Sales</td><td className="py-2 pl-3 text-right text-text font-bold">{inr(todayFin.sales)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(yesterdayFin.sales)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(mtdFin.sales)}</td></tr>
                  <tr className="border-b border-line/40"><td className="py-2 pr-3 text-text">Gross margin <span className="text-text-faint">(@70.6%)</span></td><td className="py-2 pl-3 text-right text-text">{inr(todayFin.grossMargin)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(yesterdayFin.grossMargin)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(mtdFin.grossMargin)}</td></tr>
                  <tr className="border-b border-line/40"><td className="py-2 pr-3 text-text">Contribution</td><td className="py-2 pl-3 text-right text-text">{inr(todayFin.contribution)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(yesterdayFin.contribution)}</td><td className="py-2 pl-3 text-right text-text-muted">{inr(mtdFin.contribution)}</td></tr>
                  <tr><td className="py-2 pr-3 text-text font-bold">EBITDA (est.)</td><td className={`py-2 pl-3 text-right font-bold ${todayFin.ebitda >= 0 ? "text-green-400" : "text-red-400"}`}>{inr(todayFin.ebitda)}</td><td className={`py-2 pl-3 text-right font-bold ${yesterdayFin.ebitda >= 0 ? "text-green-400" : "text-red-400"}`}>{inr(yesterdayFin.ebitda)}</td><td className={`py-2 pl-3 text-right font-bold ${mtdFin.ebitda >= 0 ? "text-green-400" : "text-red-400"}`}>{inr(mtdFin.ebitda)}</td></tr>
                </tbody>
              </table>
            </div>
            <p className="text-[10px] text-text-faint mt-3">Uses a real {(cogsRate * 100).toFixed(1)}% COGS (from actual purchase data, {fmtWin(foodCostWindow)}, company-wide — not yet per-outlet), 5% wastage, {(commissionRate * 100).toFixed(1)}% online commission (real, {fmtWin(commissionWindow)}). Both reset to the 1st of the month — COGS is this month to date; commission also includes last month since Swiggy/Zomato payouts settle slowly. EBITDA = contribution minus fixed costs; this business has no separate interest/depreciation line to strip out. Month-to-date fixed costs are counted for {salesDays} of {daysInMonth} days (the days of sales included).</p>
          </Card>

          <Card title="Budget vs Actual — profit (month)">
            {budgetRows.length === 0 ? (
              <p className="text-xs text-text-faint">No outlet has a monthly profit goal (&quot;Target A&quot;) set yet — set one in Outlet P&amp;L to see it compared here.</p>
            ) : (
              <>
                <div className="flex items-baseline gap-3 mb-4">
                  <span className={`text-2xl font-black ${budgetTotalActual >= budgetTotalPlanned ? "text-green-400" : "text-red-400"}`}>{inr(budgetTotalActual)}</span>
                  <span className="text-[11px] font-mono text-text-muted">actual vs {inr(budgetTotalPlanned)} planned ({budgetRows.length} outlet{budgetRows.length === 1 ? "" : "s"} with a goal set)</span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs font-mono whitespace-nowrap">
                    <thead>
                      <tr className="text-text-muted uppercase tracking-widest text-[10px] border-b border-line">
                        <th className="text-left py-2 pr-3">Outlet</th>
                        <th className="text-right py-2 pl-3">Planned (Target A)</th>
                        <th className="text-right py-2 pl-3">Actual (MTD)</th>
                        <th className="text-right py-2 pl-3">Variance</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...budgetRows].sort((a, b) => (a.variancePct ?? 0) - (b.variancePct ?? 0)).map((r) => (
                        <tr key={r.o} className="border-b border-line/40">
                          <td className="py-1.5 pr-3 text-text">{r.name}</td>
                          <td className="py-1.5 pl-3 text-right text-text-muted">{inr(r.budget)}</td>
                          <td className="py-1.5 pl-3 text-right text-text">{inr(r.netProfit)}</td>
                          <td className={`py-1.5 pl-3 text-right font-bold ${r.variance >= 0 ? "text-green-400" : "text-red-400"}`}>{r.variance >= 0 ? "+" : ""}{inr(r.variance)} {r.variancePct !== null ? `(${r.variancePct >= 0 ? "+" : ""}${r.variancePct.toFixed(0)}%)` : ""}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <p className="text-[10px] text-text-faint mt-3">Planned figure is each outlet&apos;s &quot;Target A&quot; profit goal from Outlet P&amp;L. Outlets with no goal set aren&apos;t shown — set one there to track it here.</p>
              </>
            )}
          </Card>

          <Card title="2. Outlet ranking — growth, profitability, ratings, compliance">
            <div className="overflow-x-auto">
              <table className="w-full text-xs font-mono whitespace-nowrap">
                <thead>
                  <tr className="text-text-muted uppercase tracking-widest text-[10px] border-b border-line">
                    <th className="text-left py-2 pr-3">Outlet</th>
                    <th className="text-right py-2 pl-3">Sales growth (MoM)</th>
                    <th className="text-right py-2 pl-3">Profitability</th>
                    <th className="text-right py-2 pl-3">Rating</th>
                    <th className="text-right py-2 pl-3">Compliance</th>
                  </tr>
                </thead>
                <tbody>
                  {[...outletRanking].sort((a, b) => (b.profitabilityPct ?? -999) - (a.profitabilityPct ?? -999)).map((r) => (
                    <tr key={r.o} className="border-b border-line/40">
                      <td className="py-1.5 pr-3 text-text">{r.name}</td>
                      <td className={`py-1.5 pl-3 text-right ${r.growthPct == null ? "text-text-faint" : r.growthPct >= 0 ? "text-green-400" : "text-red-400"}`}>{r.growthPct == null ? "—" : `${r.growthPct >= 0 ? "+" : ""}${r.growthPct.toFixed(0)}%`}</td>
                      <td className={`py-1.5 pl-3 text-right ${r.profitabilityPct == null ? "text-text-faint" : r.profitabilityPct >= 0 ? "text-green-400" : "text-red-400"}`}>{r.profitabilityPct == null ? "—" : `${r.profitabilityPct.toFixed(1)}%`}</td>
                      <td className="py-1.5 pl-3 text-right text-text">{r.avgRating == null ? "—" : `${r.avgRating.toFixed(1)}★ (${r.ratingCount})`}</td>
                      <td className={`py-1.5 pl-3 text-right ${r.compliancePct >= 90 ? "text-green-400" : r.compliancePct >= 70 ? "text-yellow-400" : "text-red-400"}`}>{r.compliancePct.toFixed(0)}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <p className="text-[10px] text-text-faint mt-3">Food cost isn't ranked here — it's a real {(cogsRate * 100).toFixed(1)}% figure now (not a guess), but still company-wide and applied the same way to every outlet, so ranking on it wouldn't show anything outlet-specific yet. Needs per-outlet Inventory &amp; Food Cost data first.</p>
          </Card>

          <Card title="3. Top opportunities, profit leakages &amp; operational risks">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div>
                <p className="text-[10px] font-mono text-green-400 uppercase tracking-widest mb-2">Opportunities</p>
                {opportunities.length === 0 ? <p className="text-xs text-text-faint">None with full cost data yet.</p> : opportunities.map((o) => (
                  <div key={o.o} className="text-xs py-1.5 border-t border-line/60"><span className="text-text">{o.name}</span><br /><span className="text-green-400 font-bold">+{inr(o.netProfit)}</span></div>
                ))}
              </div>
              <div>
                <p className="text-[10px] font-mono text-red-400 uppercase tracking-widest mb-2">Profit leakages</p>
                {leakages.length === 0 ? <p className="text-xs text-text-faint">No outlet in the red.</p> : leakages.map((o) => (
                  <div key={o.o} className="text-xs py-1.5 border-t border-line/60"><span className="text-text">{o.name}</span><br /><span className="text-red-400 font-bold">{inr(o.netProfit)}</span> <span className="text-text-faint">— {whyBleed(o)}</span></div>
                ))}
              </div>
              <div>
                <p className="text-[10px] font-mono text-yellow-400 uppercase tracking-widest mb-2">Operational risks</p>
                {complianceRisks.length === 0 ? <p className="text-xs text-text-faint">All outlets filing reliably.</p> : complianceRisks.slice(0, 5).map((o) => (
                  <div key={o.o} className="text-xs py-1.5 border-t border-line/60"><span className="text-text">{o.name}</span><br /><span className="text-yellow-400 font-bold">{o.compliancePct.toFixed(0)}% reporting</span> <span className="text-text-faint">this month — data below this point is unreliable</span></div>
                ))}
              </div>
            </div>
          </Card>

          <Card title="4. Real-time alerts">
            <div className="space-y-3">
              <div>
                <p className="text-[10px] font-mono text-red-400 uppercase tracking-widest mb-1.5">📉 Sales decline (today, &lt;50% of target)</p>
                {salesDeclineAlerts.length === 0 ? <p className="text-xs text-text-faint">No outlet below half its target today.</p> : salesDeclineAlerts.map((a) => (
                  <p key={a.o} className="text-xs text-text">{a.name} — {inr(a.tot)} of {inr(a.tgt)} target ({a.pct.toFixed(0)}%)</p>
                ))}
              </div>
              <div>
                <p className="text-[10px] font-mono text-red-400 uppercase tracking-widest mb-1.5">💬 Unresolved complaints (this month)</p>
                {unresolvedComplaints.length === 0 ? <p className="text-xs text-text-faint">Nothing outstanding.</p> : unresolvedComplaints.slice(0, 5).map((c: any, i: number) => (
                  <p key={i} className="text-xs text-text">{OUTLET_NAMES[c.outlet_id] || c.outlet_id} — {c.rating}★, no refund logged yet</p>
                ))}
              </div>
              <div className="grid grid-cols-3 gap-2 pt-2">
                <div className="bg-surface-2/60 border border-dashed border-line p-2 text-center"><p className="text-[10px] text-text-faint">📦 Stock-outs</p><p className="text-[10px] text-text-faint mt-1">🔒 Needs Inventory data</p></div>
                <div className="bg-surface-2/60 border border-dashed border-line p-2 text-center"><p className="text-[10px] text-text-faint">🗑️ High wastage</p><p className="text-[10px] text-text-faint mt-1">🔒 Needs Inventory data</p></div>
                <div className="bg-surface-2/60 border border-dashed border-line p-2 text-center"><p className="text-[10px] text-text-faint">👥 Staffing problems</p><p className="text-[10px] text-text-faint mt-1">🔒 Needs Labour data</p></div>
              </div>
            </div>
          </Card>

          <Card title="5. Today's action list" right={<span className="text-[10px] font-mono text-text-faint">{openTasks.length} open</span>}>
            {openTasks.length === 0 ? (
              <p className="text-text-faint text-xs">Nothing open — all caught up.</p>
            ) : (
              <div className="space-y-1.5">
                {openTasks.slice(0, 15).map((t: any) => {
                  const overdue = t.due_at && new Date(t.due_at) < new Date() && t.status !== "completed";
                  const dueStr = t.due_at ? new Date(t.due_at).toLocaleDateString("en-IN", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }) : "no deadline";
                  return (
                    <div key={t.id} className="flex justify-between items-start text-xs py-1.5 border-t border-line/60 gap-3">
                      <div className="min-w-0">
                        <p className="text-text truncate">{t.title}{t.outlet_id ? <span className="text-text-faint"> · {OUTLET_NAMES[t.outlet_id] || t.outlet_id}</span> : ""}</p>
                        <p className="text-[10px] text-text-muted">{staffNames[t.assigned_to] || t.assigned_to} · {t.priority}{overdue ? <span className="text-red-400"> · overdue</span> : ""}</p>
                      </div>
                      <span className={`text-[10px] font-mono whitespace-nowrap ${overdue ? "text-red-400" : "text-text-muted"}`}>{dueStr}</span>
                    </div>
                  );
                })}
                {openTasks.length > 15 && <p className="text-[10px] text-text-faint pt-1">+{openTasks.length - 15} more — see Tasks tab for the full list.</p>}
              </div>
            )}
          </Card>

          <Card title="📊 Data volume — Supabase health">
            <p className="text-[10px] text-text-faint mb-3">Any single query pulls at most ~1000 rows unless it's specifically built to fetch in batches — a table crossing that size is worth double-checking. Outlet Health and Purchase &amp; Vendors already had this bug and are now fixed.</p>
            <div className="grid grid-cols-2 md:grid-cols-5 gap-2">
              {tableCounts.map((t) => {
                const risk = t.count == null ? "unknown" : t.count >= 1000 ? "high" : t.count >= 700 ? "watch" : "ok";
                const color = risk === "high" ? "text-red-400 border-red-500/30" : risk === "watch" ? "text-yellow-400 border-yellow-500/30" : "text-green-400 border-line";
                return (
                  <div key={t.name} className={`bg-surface-2/60 border p-2 text-center ${color}`}>
                    <p className="text-[9px] font-mono text-text-faint uppercase truncate">{t.name}</p>
                    <p className={`text-sm font-bold ${color.split(" ")[0]}`}>{t.count ?? "—"}</p>
                    <p className="text-[9px] text-text-faint">{t.source}</p>
                  </div>
                );
              })}
            </div>
          </Card>

          <p className="text-[10px] font-mono text-text-faint uppercase tracking-widest mt-8 mb-3">Supporting detail</p>

          <Card title="Profit / Loss — month to date (all outlets)">
            <div className="flex items-baseline gap-3 mb-2">
              <span className={`text-3xl font-black ${totalProfit >= 0 ? "text-green-400" : "text-red-400"}`}>{totalProfit >= 0 ? "Making " : "Losing "}{inr(Math.abs(totalProfit))}</span>
              <span className="text-[11px] font-mono text-text-muted">net this month</span>
            </div>
           <p className="text-[11px] text-text-muted mb-2">Based on {_complete.length} of {OUTLETS.length} outlets with full cost data{_incompleteCount > 0 ? ` · ${_incompleteCount} excluded (no P&L uploaded)` : ""}.</p>
            {worstPnl ? (
              <p className="text-xs text-text mb-1"><span className="text-red-400 font-bold">Bleeding most: {worstPnl.name}</span> ({inr(worstPnl.netProfit)}) — {whyBleed(worstPnl)}</p>
            ) : <p className="text-xs text-green-400 mb-1">No outlet is in the red this month.</p>}
            {worstPnl && <p className="text-xs text-text-muted mb-2"><span className="text-yellow-400">Fix:</span> {fixBleed(worstPnl)}</p>}
            {bleeders.length > 1 && <p className="text-[11px] font-mono text-text-muted">Also in red: {bleeders.slice(1, 4).map(b => `${b.name} (${inr(b.netProfit)})`).join(", ")}</p>}
            {noFixedCount > 0 && <p className="text-[10px] text-orange-400 mt-2">⚠ Fixed costs not entered for {noFixedCount} outlet(s) in Outlet P&amp;L — their profit is overstated until you add rent/staff/etc.</p>}
          </Card>

          <Card title={`Monthly target · ${d0.toLocaleDateString("en-IN", { month: "long" })}`}>
            <div className="flex justify-between text-xs mb-2"><span className="text-text-muted">Achieved {lakh(mtd)}</span><span className="text-text-muted">Target {lakh(MONTHLY_TARGET)}</span></div>
            <div className="h-3 bg-canvas border border-line overflow-hidden mb-1"><div className="h-full bg-accent" style={{ width: `${Math.min(targetPct, 100)}%` }} /></div>
            <p className="text-[11px] text-text-muted mb-4">{targetPct.toFixed(1)}% of target · {lakh(Math.max(MONTHLY_TARGET - mtd, 0))} to go</p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-3 border-t border-line">
              <div><p className="text-[10px] text-text-muted uppercase">Current run rate</p><p className="text-base font-bold">{lakh(runRate)}<span className="text-[10px] text-text-muted">/day</span></p></div>
             <div><p className="text-[10px] text-text-muted uppercase">Required run rate</p><p className={`text-base font-bold ${required > runRate ? "text-red-400" : "text-green-400"}`}>{daysLeft <= 2 ? "Month ending" : <>{lakh(required)}<span className="text-[10px] text-text-muted">/day</span></>}</p></div>
             <div><p className="text-[10px] text-text-muted uppercase">Gap / day</p><p className={`text-base font-bold ${required > runRate ? "text-red-400" : "text-green-400"}`}>{daysLeft <= 2 ? "—" : (required > runRate ? lakh(required - runRate) : "On pace")}</p></div>
              <div><p className="text-[10px] text-text-muted uppercase">Proj. shortfall</p><p className={`text-base font-bold ${shortfall >= 0 ? "text-green-400" : "text-red-400"}`}>{shortfall >= 0 ? "+" : "−"}{lakh(Math.abs(shortfall))}</p></div>
            </div>
            <p className="text-[11px] text-text-muted mt-3 pt-3 border-t border-line">Channel mix (MTD): Offline {lakh(mShop)} · Online {lakh(mOnline)} — for every ₹1 walk-in, <span className="text-yellow-400 font-bold">₹{offlineRatio.toFixed(1)}</span> online.</p>
          </Card>
<Card title={`Sales by channel · ${new Date(date + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" })}`} right={<span className="text-[10px] font-mono text-text-faint">{dayOfMonth}/{daysInMonth}</span>}>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
              <div><p className="text-[10px] text-text-muted uppercase">Shop</p><p className="text-lg font-bold">{inr(tShop)}</p></div>
              <div><p className="text-[10px] text-text-muted uppercase">Swiggy</p><p className="text-lg font-bold text-orange-400">{inr(tSw)}</p></div>
              <div><p className="text-[10px] text-text-muted uppercase">Zomato</p><p className="text-lg font-bold text-red-400">{inr(tZo)}</p></div>
              <div><p className="text-[10px] text-text-muted uppercase">Orders Sw / Zo</p><p className="text-lg font-bold">{swC} / {zoC}</p></div>
            </div>
          </Card>

          <Card title={`Sales vs target by outlet · ${new Date(date + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}`}>
            <div className="space-y-2 mb-2">
              {OUTLETS.map(o => {
                const r = out.find((x: any) => x.outlet_id === o);
                const tot = r ? (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0) : 0;
                const tgt = r ? (Number(r.target) || OUTLET_TARGETS[o]) : OUTLET_TARGETS[o];
                const maxV = Math.max(...OUTLETS.map(oo => { const rr = out.find((x: any) => x.outlet_id === oo); const tt = rr ? (Number(rr.shop_sales_value) || 0) + (Number(rr.swiggy_sales_value) || 0) + (Number(rr.zomato_sales_value) || 0) : 0; return Math.max(tt, Number(rr?.target) || OUTLET_TARGETS[oo] || 0); }), 1);
                const hit = tgt > 0 && tot >= tgt;
                const salesPct = Math.min((tot / maxV) * 100, 100);
                const tgtPct = Math.min((tgt / maxV) * 100, 100);
                return (
                  <div key={o}>
                    <div className="flex justify-between text-[11px] mb-0.5"><span className="text-text">{OUTLET_NAMES[o]}</span><span className="font-mono text-text-muted">{inr(tot)}{tgt > 0 ? ` / ${inr(tgt)}` : ""}</span></div>
                    <div className="relative h-4 bg-canvas border border-line">
                      <div className={`absolute top-0 left-0 h-full ${!r ? "bg-zinc-700" : hit ? "bg-green-500" : "bg-orange-500"}`} style={{ width: `${salesPct}%` }} />
                      {tgt > 0 && <div className="absolute top-0 h-full w-0.5 bg-yellow-300" style={{ left: `${tgtPct}%` }} title="target" />}
                    </div>
                  </div>
                );
              })}
            </div>
            <p className="text-[10px] font-mono text-text-faint mb-4">▮ green = hit · ▮ orange = below · ▮ grey = not reported · | yellow line = target</p>
          <div className="hidden">
            <div className="space-y-1">
              {OUTLETS.map(o => {
                const r = out.find(x => x.outlet_id === o);
                if (!r) return <div key={o} className="flex justify-between text-xs py-1.5 border-t border-line/60"><span className="text-text-muted">{OUTLET_NAMES[o]}</span><span className="text-text-faint">not reported</span></div>;
                const tot = n(r.shop_sales_value) + n(r.swiggy_sales_value) + n(r.zomato_sales_value);
                const t = n(r.target) || OUTLET_TARGETS[o]; const hit = t > 0 && tot >= t;
                return <div key={o} className="flex justify-between items-center text-xs py-1.5 border-t border-line/60"><span className="text-text">{OUTLET_NAMES[o]}</span><span className="font-mono">{inr(tot)} {t > 0 && <span className={hit ? "text-green-400 ml-1" : "text-red-400 ml-1"}>{hit ? "✓" : "✗"}</span>}</span></div>;
              })}
            </div>
          </div>
          </Card>

          <Card title="Reporting status (today)">
            <p className="text-[10px] text-text-muted uppercase mb-1">Daily reports</p>
            <p className="text-xs mb-3">{DUTY_STAFF.map(s => { const isOff = offRows.includes(s.id); return <span key={s.id} className={isOff ? "text-text-muted mr-3 inline-block" : filedDaily.has(s.id) ? "text-green-400 mr-3 inline-block" : "text-red-400 mr-3 inline-block"}>{isOff ? "🌙" : filedDaily.has(s.id) ? "✓" : "✗"} {s.name}{isOff ? " (off)" : ""}</span>; })}</p>
            <p className="text-[10px] text-text-muted uppercase mb-1">Outlet reports</p>
            <p className="text-xs">{OUTLETS.map(o => <span key={o} className={filedOutlets.has(o) ? "text-green-400 mr-3 inline-block" : "text-red-400 mr-3 inline-block"}>{filedOutlets.has(o) ? "✓" : "✗"} {OUTLET_NAMES[o]}</span>)}</p>
          </Card>

          <Card title="Reviews & complaints (today)">
            {revs.length === 0 ? <p className="text-text-faint text-xs">No reviews logged.</p> : (
              <div className="space-y-1">
                <p className="text-xs text-text-muted mb-2">{revs.length} review(s) · {revs.filter(r => r.valid_complaint).length} valid complaint(s) · {revs.filter(r => r.refund_given).length} refunded</p>
                {revs.map(rv => { const p = reviewPoints(n(rv.rating), rv.valid_complaint); return (
                  <div key={rv.id} className="flex justify-between text-xs py-1 border-t border-line/60">
                    <span className="text-text">{OUTLET_NAMES[rv.outlet_id] || rv.outlet_id} · {rv.platform} · {rv.rating}★{rv.valid_complaint ? " · complaint" : ""}</span>
                    <span className={p >= 0 ? "text-green-400 font-bold" : "text-red-400 font-bold"}>{p >= 0 ? "+" : ""}{p}</span>
                  </div>
                ); })}
              </div>
            )}
          </Card>

          <Card title={`Points earned · ${d0.toLocaleDateString("en-IN", { day: "numeric", month: "short" })}`}>
            <div className="space-y-1">
              {DUTY_STAFF.map(s => (
                <div key={s.id} className="flex justify-between text-sm py-1.5 border-t border-line/60">
                  <span className="text-text">{s.name}</span>
                  <span className={`font-bold font-mono ${(pts[s.id] || 0) >= 0 ? "text-yellow-400" : "text-red-400"}`}>{(pts[s.id] || 0) >= 0 ? "+" : ""}{pts[s.id] || 0}</span>
                </div>
              ))}
            </div>
            <p className="text-[10px] text-text-faint mt-2">Daily +10 / late −5 · outlet +20 (+30 target) · reviews. Season totals on Leaderboard.</p>
          </Card>

          <Card title="Payout — latest week per outlet">
            <div className="space-y-1">
              {OUTLETS.map(o => {
                const sw = latestPay[o + "_swiggy"], zo = latestPay[o + "_zomato"];
                if (!sw && !zo) return null;
                return <div key={o} className="flex justify-between text-xs py-1 border-t border-line/60"><span className="text-text">{OUTLET_NAMES[o]}</span><span className="font-mono text-text-muted">{sw ? `Sw ${inr(n(sw.amount_transferable))}` : ""} {zo ? `Zo ${inr(n(zo.net_payout))}` : ""}</span></div>;
              })}
              {OUTLETS.every(o => !latestPay[o + "_swiggy"] && !latestPay[o + "_zomato"]) && <p className="text-text-faint text-xs">No payout data yet.</p>}
            </div>
          </Card>

          <Card title={`Atlas Reconciliation — ${d0.toLocaleDateString("en-IN", { month: "long", year: "numeric" })}`}>
            {atlasResults.length === 0 ? (
              <p className="text-text-faint text-xs font-mono">No Atlas data saved for {d0.toLocaleDateString("en-IN", { month: "long", year: "numeric" })} yet — run Reconciliation for this month and save to dashboard.</p>
            ) : (
              <>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs font-mono whitespace-nowrap">
                    <thead>
                      <tr className="text-text-muted uppercase tracking-widest text-[10px] border-b border-line">
                        <th className="text-left py-2 pr-3">Outlet</th>
                        <th className="text-right py-2 pl-3">Atlas Gross</th>
                        <th className="text-right py-2 pl-3">Reported</th>
                        <th className="text-right py-2 pl-3">Diff</th>
                        <th className="text-center py-2 pl-3">Flag</th>
                        <th className="text-right py-2 pl-3">Atlas Net</th>
                        <th className="text-right py-2 pl-3">Lost Rev.</th>
                      </tr>
                    </thead>
                    <tbody>
                      {atlasReconRows.map(r => {
                        const hasAtlas = r.atlasGross !== null;
                        const ok = hasAtlas && Math.abs(r.diff!) <= ATLAS_THRESHOLD;
                        return (
                          <tr key={r.outlet_id} className="border-b border-line/40">
                            <td className="py-1.5 pr-3 text-text">{OUTLET_NAMES[r.outlet_id]}</td>
                            <td className="py-1.5 pl-3 text-right">{hasAtlas ? inr(r.atlasGross!) : <span className="text-text-faint">—</span>}</td>
                            <td className="py-1.5 pl-3 text-right text-text-muted">{inr(r.reportedGross)}</td>
                            <td className={`py-1.5 pl-3 text-right font-semibold ${!hasAtlas ? "text-text-faint" : ok ? "text-text-muted" : r.diff! < 0 ? "text-red-400" : "text-yellow-400"}`}>
                              {!hasAtlas ? "—" : `${r.diff! >= 0 ? "+" : "−"}${inr(Math.abs(r.diff!))}`}
                            </td>
                            <td className="py-1.5 pl-3 text-center">
                              {hasAtlas && (
                                <span className={`text-[10px] uppercase tracking-widest border px-1.5 py-0.5 ${ok ? "border-green-500/40 text-green-400" : "border-red-500/40 text-red-400"}`}>
                                  {ok ? "✓" : "⚠"}
                                </span>
                              )}
                            </td>
                            <td className="py-1.5 pl-3 text-right text-blue-400">{hasAtlas ? inr(r.atlasNet!) : <span className="text-text-faint">—</span>}</td>
                            <td className="py-1.5 pl-3 text-right text-red-400">{hasAtlas && r.atlasLost! > 0 ? inr(r.atlasLost!) : <span className="text-text-faint">—</span>}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      {(() => {
                        const tAG = atlasReconRows.reduce((s, r) => s + (r.atlasGross ?? 0), 0);
                        const tRG = atlasReconRows.reduce((s, r) => s + r.reportedGross, 0);
                        const tD = tAG - tRG;
                        return (
                          <tr className="border-t border-line">
                            <td className="py-2 text-text-muted font-bold text-[10px] uppercase tracking-widest">Total</td>
                            <td className="py-2 pl-3 text-right font-bold text-text">{inr(tAG)}</td>
                            <td className="py-2 pl-3 text-right font-bold text-text">{inr(tRG)}</td>
                            <td className={`py-2 pl-3 text-right font-bold ${tD < 0 ? "text-red-400" : "text-yellow-400"}`}>{tD >= 0 ? "+" : "−"}{inr(Math.abs(tD))}</td>
                            <td></td>
                            <td className="py-2 pl-3 text-right font-bold text-blue-400">{inr(atlasReconRows.reduce((s, r) => s + (r.atlasNet ?? 0), 0))}</td>
                            <td className="py-2 pl-3 text-right font-bold text-red-400">{inr(atlasReconRows.reduce((s, r) => s + (r.atlasLost ?? 0), 0))}</td>
                          </tr>
                        );
                      })()}
                    </tfoot>
                  </table>
                </div>
                <p className="text-[10px] font-mono text-text-faint mt-2">Atlas Gross vs staff-reported MTD (shop + Swiggy + Zomato) · ⚠ = gap &gt; {inr(ATLAS_THRESHOLD)} · Atlas Net = verified founder figure · Lost Rev. = platform cancellations</p>
              </>
            )}
          </Card>
        </>
      )}
    </div>
  );
}
