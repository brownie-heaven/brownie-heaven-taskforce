"use client";
import { useEffect, useState, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import { useRouter } from "next/navigation";
import { createClient } from "@supabase/supabase-js";
import { celebrate } from "../celebrate";
import { computeScores, type ScoreRow } from "@/lib/score";
import { computeMonthEnd, type OutletMonthEnd } from "@/lib/monthEndProjection";
import PayoutTab from "./PayoutTab";
import FounderDashboard from "./FounderDashboard";
import CommandCentre from "./CommandCentre";
import MyOutletsDashboard from "./MyOutletsDashboard";
import TeamDashboards from "./TeamDashboards";
import MessagesTab from "./MessagesTab";
import ActiveStatusTab from "./ActiveStatusTab";
import HelpTab from "./HelpTab";
import ThemePicker from "@/components/ThemePicker";
import BrownieLoader from "@/components/BrownieLoader";
import { getStoredBrownieMode } from "@/lib/theme";
import ReconciliationTab from "./ReconciliationTab";
import supabaseStock from "@/lib/supabaseStock";
import { buildSwiggyUrl } from "@/lib/swiggyLiveOffers";
import { fetchRealFoodCostPct, REAL_FOOD_COST_WINDOW_DAYS } from "@/lib/realFoodCost";
import { fetchRealCommissionPct, REAL_COMMISSION_WINDOW_DAYS } from "@/lib/realCommission";
import { OUTLET_ID_TO_STOCK_NAME } from "@/lib/outletMap";
import { FOOD_COST_MAP } from "@/lib/foodCosts";
import { matchItemCost } from "@/lib/itemPerfCosts";
import { useActivityHeartbeat } from "@/lib/useActivityHeartbeat";
import ActivityToastStack from "@/components/ActivityToastStack";
import PushRegister from "@/components/PushRegister";
import NudgeToast from "@/components/NudgeToast";
import { parseSwiggyPayoutEmail } from "@/lib/swiggyPayoutParser";


const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL!,
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY!
);

type Staff = { id: string; name: string; role: string; report_time: string | null; outlets?: string[] };
type Task = { id: string; title: string; description: string; status: string; priority: string; due_at: string; assigned_to: string; assigned_by: string; outlet_id: string | null };
type Report = { id: string; staff_id: string; content: string; submitted_at: string; is_late: boolean; report_data: Record<string, string>; staff_role: string };
type OutletReport = {
  id: string; staff_id: string; outlet_id: string; report_date: string;
  shop_sales_count: number; shop_sales_value: number;
  swiggy_sales_count: number; swiggy_sales_value: number;
  zomato_sales_count: number; zomato_sales_value: number;
  target: number; swiggy_live: boolean; zomato_live: boolean;
  discount_running: string; discount_rate_good: boolean; discount_given?: number;
  unavailable_items: string; expiry_count: number; expiry_items: string;
  complimentary_count: number; complimentary_items: string; complimentary_reason: string;
issues: string; action_taken: string; submitted_at: string; is_late: boolean; is_edited: boolean;
bh_google_rating: number; bh_swiggy_rating: number; bh_zomato_rating: number;
cbh_google_rating: number; cbh_swiggy_rating: number; cbh_zomato_rating: number;
icbh_google_rating: number; icbh_swiggy_rating: number; icbh_zomato_rating: number;
// Brand-wise sales (from BRAND_SPLIT_FROM onwards). The combined shop/swiggy/zomato fields above stay as BH + CBH totals.
bh_shop_sales_count?: number | null; bh_shop_sales_value?: number | null;
bh_swiggy_sales_count?: number | null; bh_swiggy_sales_value?: number | null;
bh_zomato_sales_count?: number | null; bh_zomato_sales_value?: number | null;
cbh_shop_sales_count?: number | null; cbh_shop_sales_value?: number | null;
cbh_swiggy_sales_count?: number | null; cbh_swiggy_sales_value?: number | null;
cbh_zomato_sales_count?: number | null; cbh_zomato_sales_value?: number | null;
};
// Outlet reports for dates on/after this are entered as separate BH and CBH sales. Older dates keep the single combined entry.
const BRAND_SPLIT_FROM = "2026-10-09";
// Entry limits (typo guard) apply only to report dates on/after this, so older and backfilled reports are never blocked.
const ENTRY_LIMITS_FROM = "2026-10-26";
// Card rows showing BH vs CBH sales for a report (empty for older reports that only have the combined total).
const brandSalesItems = (r: any): { label: string; value: string; color?: string }[] => {
  if (!r || r.bh_shop_sales_value == null) return [];
  const tot = (b: string) => ["shop", "swiggy", "zomato"].reduce((a, ch) => a + (Number(r[`${b}_${ch}_sales_value`]) || 0), 0);
  const cnt = (b: string) => ["shop", "swiggy", "zomato"].reduce((a, ch) => a + (Number(r[`${b}_${ch}_sales_count`]) || 0), 0);
  const line = (b: string) => `Shop ₹${Number(r[`${b}_shop_sales_value`]) || 0} · Swiggy ₹${Number(r[`${b}_swiggy_sales_value`]) || 0} · Zomato ₹${Number(r[`${b}_zomato_sales_value`]) || 0}`;
  return [
    { label: "BH Total Sales", value: `₹${tot("bh")} (${cnt("bh")} orders)` },
    { label: "BH Split", value: line("bh") },
    { label: "CBH Total Sales", value: `₹${tot("cbh")} (${cnt("cbh")} orders)` },
    { label: "CBH Split", value: line("cbh") },
  ];
};
const BRAND_SALES_KEYS =["shop", "swiggy", "zomato"].flatMap((ch) => ["bh", "cbh"].flatMap((b) => [`${b}_${ch}_sales_count`, `${b}_${ch}_sales_value`]));

const ALL_STAFF = [
  { id: "nishant", name: "Nishant Vijayakumar", role: "Owner", report_time: null, outlets: [] },
  { id: "arun", name: "Arun Kumar", role: "Manager", report_time: "22:00", outlets: [] },
  { id: "nilani", name: "Nilani Nallamuthu", role: "HR", report_time: "22:00", outlets: [] },
  { id: "gowtham", name: "Gowtham", role: "Purchase Manager", report_time: "22:00", outlets: [] },
  { id: "vishnu", name: "Vishnu", role: "Asst. Ops Manager", report_time: "22:00", outlets: ["velachery","perumbakkam","tambaram","porur","anna_nagar","vadapalani"] },
  { id: "ahila", name: "Ahila", role: "Custom Cakes & Asst Ops", report_time: "22:00", outlets: ["adayar","bsr_mall","pallavaram","ra_puram","besant_nagar"] },
  { id: "stephen", name: "Stephen", role: "Outlet Manager", report_time: null, outlets: ["royapettah"] },
  { id: "niranjana", name: "Niranjana", role: "Founder's Office", report_time: null, outlets: [] },
  { id: "rafiq", name: "Rafiq", role: "Head Chef", report_time: null, outlets: [] },
  { id: "ajay", name: "Ajay", role: "Financial Analyst", report_time: null, outlets: [] },
];
// Former staff: no longer in the app, but their old reports/tasks still show their name.
const FORMER_STAFF: Record<string, string> = { bharani: "Bharani" };

const OUTLETS = ["royapettah","adayar","bsr_mall","velachery","ra_puram","anna_nagar","pallavaram","vadapalani","besant_nagar","perumbakkam","tambaram","porur"];
const OUTLET_NAMES: Record<string, string> = {
  royapettah: "Royapettah",
  adayar: "Adyar",
  bsr_mall: "BSR Mall",
  velachery: "Velachery",
  ra_puram: "RA Puram",
  anna_nagar: "Anna Nagar",
  pallavaram: "Pallavaram",
  vadapalani: "Vadapalani",
  besant_nagar: "Besant Nagar",
  perumbakkam: "Perumbakkam",
  tambaram: "Tambaram",
  porur: "Porur",
};
const OUTLET_TARGETS: Record<string, string> = {
  royapettah: "86667",
  adayar: "21667",
  bsr_mall: "35000",
  ra_puram: "21667",
  anna_nagar: "51667",
  porur: "50000",
  perumbakkam: "13333",
  tambaram: "20000",
  velachery: "23333",
  pallavaram: "23333",
  vadapalani: "23333",
  besant_nagar: "11667",
};

// Official FY 2026-27 monthly targets by brand (BH = Brownie Heaven, CBH = Cakes by BH, ICBH = Ice Creams by BH).
// Vadapalani & Besant Nagar not on the official sheet yet -> no brand split; they keep their default daily target above.
const BRAND_TARGETS: Record<string, { bh: number; cbh: number; icbh: number; total: number }> = {
  royapettah:  { bh: 1500000, cbh: 700000, icbh: 200000, total: 2400000 },
  anna_nagar:  { bh: 800000,  cbh: 600000, icbh: 150000, total: 1550000 },
  porur:       { bh: 750000,  cbh: 600000, icbh: 150000, total: 1500000 },
  bsr_mall:    { bh: 600000,  cbh: 300000, icbh: 150000, total: 1050000 },
  tambaram:    { bh: 300000,  cbh: 200000, icbh: 100000, total: 600000 },
  ra_puram:    { bh: 350000,  cbh: 200000, icbh: 100000, total: 650000 },
  adayar:      { bh: 300000,  cbh: 150000, icbh: 100000, total: 550000 },
  perumbakkam: { bh: 300000,  cbh: 100000, icbh: 0,      total: 400000 },
  pallavaram:  { bh: 350000,  cbh: 250000, icbh: 100000, total: 700000 },
  velachery:   { bh: 350000,  cbh: 250000, icbh: 100000, total: 700000 },
};

// Base monthly totals (current, in effect through July 2026).
const MONTHLY_BASE: Record<string, number> = {
  royapettah: 2400000, adayar: 550000, bsr_mall: 1050000, velachery: 700000,
  ra_puram: 650000, anna_nagar: 1550000, pallavaram: 700000, vadapalani: 699990,
  besant_nagar: 350010, perumbakkam: 400000, tambaram: 600000, porur: 1500000,
};
// Scheduled target changes. Each applies from the given month (YYYY-MM) onward. Latest matching wins.
const TARGET_UPDATES: { from: string; monthly: Record<string, number> }[] = [
 { from: "2026-08", monthly: { royapettah: 2400000, adayar: 600000, bsr_mall: 1050000, velachery: 700000, ra_puram: 650000, anna_nagar: 1800000, pallavaram: 700000, vadapalani: 700000, besant_nagar: 500000, perumbakkam: 550000, tambaram: 650000, porur: 1600000 } },
 { from: "2026-09", monthly: { royapettah: 2600000, adayar: 650000, bsr_mall: 1050000, velachery: 700000, ra_puram: 650000, anna_nagar: 1800000, pallavaram: 700000, vadapalani: 1050000, besant_nagar: 500000, perumbakkam: 550000, tambaram: 650000, porur: 1600000 } },
];
// Target changes Arun makes from the Outlet Targets screen land in the
// outlet_target_updates table instead of here — this array stays as the
// historical seed, and RUNTIME_TARGET_UPDATES (populated from that table on
// load) is merged on top of it below. Each entry works exactly like the
// hardcoded ones above: "from this month (YYYY-MM) onward, use this value" —
// so a change made effective November never touches October or earlier.
let RUNTIME_TARGET_UPDATES: { from: string; monthly: Record<string, number> }[] = [];
function monthlyTargetFor(oid: string, ym: string): number {
  let v = MONTHLY_BASE[oid] || 0;
  const all = [...TARGET_UPDATES, ...RUNTIME_TARGET_UPDATES].sort((a, b) => a.from.localeCompare(b.from));
  for (const u of all) { if (ym >= u.from && u.monthly[oid] != null) v = u.monthly[oid]; }
  return v;
}
function dailyTargetFor(oid: string, ym: string): number {
  return Math.round(monthlyTargetFor(oid, ym) / 30);
}

function computeCeoData(repRows: any[], monthRep: any[], win: string) {
  const winDays = win === "1" ? 1 : win === "30" ? 30 : 7;
  const duty = (ALL_STAFF as any[]).filter((s) => s.report_time);
  const acct = duty.map((s) => {
    const mine = repRows.filter((r) => r.staff_id === s.id);
    const filed = mine.length, late = mine.filter((r) => r.is_late).length, onTime = filed - late, missed = Math.max(0, winDays - filed);
    let tag = "", tone = "gray";
    if (filed === 0) { tag = "nothing filed 👀 hello?"; tone = "red"; }
    else if (missed === 0 && late === 0) { tag = `${filed}/${winDays} on time 🌟 the reliable one`; tone = "green"; }
    else if (missed >= Math.ceil(winDays * 0.5)) { tag = `skipped ${missed}/${winDays} days 👻 ghost mode`; tone = "red"; }
    else if (late >= Math.ceil(filed * 0.4)) { tag = `late ${late}/${filed} times 😴 loves the snooze button`; tone = "amber"; }
    else { tag = `${onTime} on time · ${late} late · ${missed} missed`; tone = "gray"; }
    return { id: s.id, name: s.name.split(" ")[0], filed, late, missed, tag, tone, score: missed * 2 + late };
  }).sort((a, b) => a.score - b.score);
  const ym = new Date().toISOString().slice(0, 7);
  const _d = new Date();
  const daysInMonth = new Date(_d.getFullYear(), _d.getMonth() + 1, 0).getDate();
  const daysElapsed = Math.max(1, _d.getDate() - 1);
  const perOutlet = OUTLETS.map((o) => {
    const sales = monthRep.filter((r) => r.outlet_id === o).reduce((a, r) => a + (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0), 0);
    const tgt = monthlyTargetFor(o, ym);
    return { o, name: OUTLET_NAMES[o] || o, sales, tgt, pace: tgt > 0 ? (sales / tgt) / (daysElapsed / daysInMonth) : 0, pct: tgt > 0 ? sales / tgt * 100 : 0 };
  });
  const monthSales = perOutlet.reduce((a, r) => a + r.sales, 0), monthTgt = perOutlet.reduce((a, r) => a + r.tgt, 0);
  const salesPct = monthTgt > 0 ? monthSales / monthTgt * 100 : 0, timePct = daysInMonth > 0 ? daysElapsed / daysInMonth * 100 : 0, onTrack = salesPct >= timePct;
  const sorted = [...perOutlet].filter((r) => r.tgt > 0 && r.sales > 0).sort((a, b) => a.pace - b.pace);
  const drag = sorted[0], hero = sorted[sorted.length - 1];
  const ideas: string[] = [];
  const notFiling = acct.filter((a) => a.filed === 0);
  const lateOnes = acct.filter((a) => a.filed > 0 && a.late >= Math.ceil(a.filed * 0.4));
  const missingOnes = acct.filter((a) => a.filed > 0 && a.missed >= Math.ceil(winDays * 0.5));
  if (notFiling.length >= 2) ideas.push(`${notFiling.length} people filed nothing this period — reporting discipline is slipping. A group reminder or a quick accountability chat would help.`);
  else if (notFiling.length === 1) ideas.push(`${notFiling[0].name} filed nothing — follow up directly; could be a real issue or just slacking.`);
  lateOnes.forEach((a) => ideas.push(`${a.name} is chronically late (${a.late}/${a.filed}). A fixed reminder time or a word in person usually fixes it.`));
  missingOnes.forEach((a) => ideas.push(`${a.name} keeps skipping reports — worth understanding why before it becomes a habit.`));
  if (!onTrack) ideas.push(`Month is behind pace (${salesPct.toFixed(0)}% sales vs ${timePct.toFixed(0)}% of the month gone). A weekend push, a combo offer, or leaning on the strong outlets could close the gap.`);
  if (drag && drag.pct < 15) ideas.push(`${drag.name} is badly behind (${drag.pct.toFixed(0)}% of target). A direct call to the manager or a local promo could lift it.`);
  if (hero && hero.pct >= 90) ideas.push(`${hero.name} is doing great (${hero.pct.toFixed(0)}%). Find out what's working there and copy it to weaker outlets.`);
  if (ideas.length === 0) ideas.push(`Everything looks healthy — team is filing and the month's on pace. Keep the momentum. 🎉`);
  return { winDays, acct, perOutlet, monthSales, monthTgt, salesPct, timePct, onTrack, drag, hero, ideas, ym, daysElapsed, daysInMonth };
}
function computeCeoCustom(staffRows: any[], outletRows: any[], from: string, to: string, outletsSel: string[]) {
  const rows = outletRows;
  const days = Math.max(1, Math.round((new Date(to).getTime() - new Date(from).getTime()) / 86400000) + 1);
  const duty = (ALL_STAFF as any[]).filter((s) => s.report_time);
  const acct = duty.map((s) => {
    const mine = staffRows.filter((r) => r.staff_id === s.id);
    const filed = mine.length, late = mine.filter((r) => r.is_late).length, onTime = filed - late, missed = Math.max(0, days - filed);
    let tag = "";
    if (filed === 0) tag = "nothing filed 👀 hello?";
    else if (missed === 0 && late === 0) tag = `${filed}/${days} on time 🌟 the reliable one`;
    else if (missed >= Math.ceil(days * 0.5)) tag = `skipped ${missed}/${days} days 👻 ghost mode`;
    else if (late >= Math.ceil(filed * 0.4)) tag = `late ${late}/${filed} times 😴 loves the snooze button`;
    else tag = `${onTime} on time · ${late} late · ${missed} missed`;
    return { id: s.id, name: s.name.split(" ")[0], filed, late, missed, tag };
  });
  const activeOutlets = outletsSel.length ? outletsSel : OUTLETS;
  const ymFrom = from.slice(0, 7);
  const perOutlet = activeOutlets.map((o) => {
    const oRows = rows.filter((r) => r.outlet_id === o);
    const sales = oRows.reduce((a, r) => a + (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0), 0);
    const tgt = dailyTargetFor(o, ymFrom) * days;
    return { o, name: OUTLET_NAMES[o] || o, sales, tgt, pct: tgt > 0 ? (sales / tgt) * 100 : 0 };
  });
  const periodSales = perOutlet.reduce((a, r) => a + r.sales, 0), periodTgt = perOutlet.reduce((a, r) => a + r.tgt, 0);
  const salesPct = periodTgt > 0 ? (periodSales / periodTgt) * 100 : 0;
  const sortedByPct = [...perOutlet].filter((r) => r.tgt > 0 && r.sales > 0).sort((a, b) => a.pct - b.pct);
  const drag = sortedByPct[0], hero = sortedByPct[sortedByPct.length - 1];
  const ideas: string[] = [];
  const notFiling = acct.filter((a) => a.filed === 0);
  if (notFiling.length >= 2) ideas.push(`${notFiling.length} people filed nothing this period — reporting discipline is slipping.`);
  else if (notFiling.length === 1) ideas.push(`${notFiling[0].name} filed nothing this period — worth a follow-up.`);
  if (periodTgt > 0 && salesPct < 90) ideas.push(`Sales are at ${salesPct.toFixed(0)}% of target for this period.`);
  if (drag && drag.pct < 40) ideas.push(`${drag.name} is badly behind (${drag.pct.toFixed(0)}% of target) for this period.`);
  if (hero && hero.pct >= 90) ideas.push(`${hero.name} is doing great (${hero.pct.toFixed(0)}%) this period.`);
  if (ideas.length === 0) ideas.push(`Everything looks healthy for this period. 🎉`);
  return { days, acct, perOutlet, periodSales, periodTgt, salesPct, drag, hero, ideas };
}
const REPORT_FIELDS: Record<string, { label: string; key: string; type?: string }[]> = {
  arun: [
    { label: "Total Sales (Rs)", key: "total_sales" },
    { label: "Target (Rs)", key: "target" },
    { label: "Achievement %", key: "achievement" },
    { label: "Best Outlet", key: "best_outlet" },
    { label: "Worst Outlet", key: "worst_outlet" },
    { label: "Swiggy Sales (Rs)", key: "swiggy_sales" },
    { label: "Zomato Sales (Rs)", key: "zomato_sales" },
    { label: "Shop Sales (Rs)", key: "shop_sales" },
    { label: "Cake Sales", key: "cake_sales" },
    { label: "Ice Cream Sales", key: "ice_cream_sales" },
    { label: "Complaints Today", key: "complaints" },
    { label: "Negative Reviews", key: "negative_reviews" },
    { label: "Stock-out Issues", key: "stock_out" },
    { label: "Staff Issues", key: "staff_issues" },
    { label: "Top Issue 1", key: "issue_1" },
    { label: "Top Issue 2", key: "issue_2" },
    { label: "Top Issue 3", key: "issue_3" },
    { label: "Tomorrow Action 1", key: "action_1" },
    { label: "Tomorrow Action 2", key: "action_2" },
    { label: "Tomorrow Action 3", key: "action_3" },
  ],
 nilani: [
    { label: "Replacement Required (Yes/No)", key: "replacement_required" },
    { label: "Training Conducted (Yes/No)", key: "training_conducted" },
    { label: "Training Topic", key: "training_topic" },
    { label: "Outlets Checked/Visited", key: "outlets_checked" },
    { label: "Staff Issues", key: "staff_issues" },
    { label: "Warnings Issued", key: "warnings_issued" },
    { label: "New Candidates Contacted", key: "candidates_contacted" },
    { label: "Interviews Scheduled", key: "interviews_scheduled" },
    { label: "Top Issue 1", key: "issue_1" },
    { label: "Top Issue 2", key: "issue_2" },
    { label: "Top Issue 3", key: "issue_3" },
    { label: "Action Needed 1", key: "action_1" },
    { label: "Action Needed 2", key: "action_2" },
    { label: "Action Needed 3", key: "action_3" },
  ],
  gowtham: [
    { label: "Total Purchase Value (Rs)", key: "purchase_value" },
    { label: "Emergency Purchase Value (Rs)", key: "emergency_purchase" },
    { label: "Purchase as per Plan (Yes/No)", key: "purchase_as_planned" },
    { label: "Supplier Issues", key: "supplier_issues" },
    { label: "Material Rejected", key: "material_rejected" },
    { label: "Stock Shortage", key: "stock_shortage" },
    { label: "Dispatch Delay", key: "dispatch_delay" },
    { label: "High-value Stock Variance (Rs)", key: "stock_variance" },
    { label: "Tomorrow Purchase Requirement (Rs)", key: "tomorrow_purchase" },
    { label: "Top Issue 1", key: "issue_1" },
    { label: "Top Issue 2", key: "issue_2" },
    { label: "Top Issue 3", key: "issue_3" },
    { label: "Action Needed 1", key: "action_1" },
    { label: "Action Needed 2", key: "action_2" },
    { label: "Action Needed 3", key: "action_3" },
  ],
  vishnu: [
    { label: "Total Enquiries Received", key: "total_enquiries" },
    { label: "Orders Confirmed", key: "orders_confirmed" },
    { label: "Conversion %", key: "conversion_pct" },
    { label: "WhatsApp Sales Value (Rs)", key: "whatsapp_sales" },
    { label: "Pending Payments (Rs)", key: "pending_payments" },
    { label: "Custom Cake Orders", key: "custom_cake_orders" },
    { label: "Bulk Enquiries", key: "bulk_enquiries" },
    { label: "Complaints Handled", key: "complaints_handled" },
    { label: "Reviews Requested", key: "reviews_requested" },
    { label: "Reviews Received", key: "reviews_received" },
    { label: "Top Revenue Opportunity 1", key: "opportunity_1" },
    { label: "Top Revenue Opportunity 2", key: "opportunity_2" },
    { label: "Top Revenue Opportunity 3", key: "opportunity_3" },
    { label: "Pending Follow-up 1", key: "followup_1" },
    { label: "Pending Follow-up 2", key: "followup_2" },
    { label: "Pending Follow-up 3", key: "followup_3" },
  ],
  
   ahila: [
    { label: "Total No Of Customizations", key: "total_cakes" },
    { label: "Revenue", key: "revenue" },
    { label: "Confirmed & Placed Today", key: "confirmed_placed" },
    { label: "Pending & Following Leads", key: "pending_following" },
    { label: "Customisations Done Today", key: "cake_orders_today" },
    { label: "Orders Delivered on Time", key: "orders_on_time" },
    { label: "Cake Complaints", key: "cake_complaints" },
    { label: "Cake Wastage/Damage", key: "cake_wastage" },
    { label: "Royapettah Sales/Ops Issue", key: "royapettah_issue" },
    { label: "Swiggy/Zomato Issues", key: "swiggy_zomato_issues" },
    { label: "Negative Reviews", key: "negative_reviews" },
    { label: "Product Unavailable Issues", key: "unavailable_issues" },
    { label: "Tomorrow's Cake Orders", key: "tomorrow_cake_orders" },
    { label: "Top Issue 1", key: "issue_1" },
    { label: "Top Issue 2", key: "issue_2" },
    { label: "Top Issue 3", key: "issue_3" },
    { label: "Action Needed 1", key: "action_1" },
    { label: "Action Needed 2", key: "action_2" },
    { label: "Action Needed 3", key: "action_3" },
  ],
  bharani: [
    { label: "Outlets Audited", key: "outlets_audited" },
    { label: "Total Wastage (Rs)", key: "total_wastage" },
    { label: "Stock Mismatch Found (Yes/No)", key: "stock_mismatch" },
    { label: "Cash Reconciliation Status", key: "cash_reconciliation" },
    { label: "Exceptions Found", key: "exceptions_found" },
    { label: "High-value Discrepancy (Rs)", key: "discrepancy_value" },
    { label: "Outlets with Issues", key: "outlets_with_issues" },
    { label: "Top Issue 1", key: "issue_1" },
    { label: "Top Issue 2", key: "issue_2" },
    { label: "Top Issue 3", key: "issue_3" },
    { label: "Action Needed 1", key: "action_1" },
    { label: "Action Needed 2", key: "action_2" },
    { label: "Action Needed 3", key: "action_3" },
  ],
};

function parseMoney(raw: string): number | null {
  const s = (raw || "").replace(/\*/g, "").replace(/₹/g, "").replace(/,/g, "").trim();
  if (!s || /negligible/i.test(s) || /^[-—\s]+$/.test(s)) return null;
  const m = s.match(/([\d.]+)\s*(cr|crore|l|lakh|k)?/i);
  if (!m) return null;
  let v = parseFloat(m[1]); if (isNaN(v)) return null;
  const u = (m[2] || "").toLowerCase();
  if (u === "cr" || u === "crore") v *= 1e7; else if (u === "l" || u === "lakh") v *= 1e5; else if (u === "k") v *= 1e3;
  return Math.round(v);
}
function parseAreaTable(text: string): { rows: { area: string; competitor: string; their: number | null; our: number | null }[] } | null {
  const lines = (text || "").split("\n").map(l => l.trim()).filter(l => l.startsWith("|"));
  if (lines.length < 2) return null;
  const cells = (l: string) => l.split("|").slice(1, -1).map(c => c.trim());
  const header = cells(lines[0]);
  const norm = header.map(h => h.replace(/\*/g, "").toLowerCase().trim());
  const areaIdx = norm.findIndex(h => h.includes("area") || h.includes("locality"));
  const ourIdx = norm.findIndex(h => h.includes("brownie heaven") || h === "bh" || h === "ours" || h === "us");
  if (areaIdx < 0 || ourIdx < 0) return null;
  const skip = /gap|note|verdict|diff|remark|rank|^$/i;
  const compCols = header.map((_, i) => i).filter(i => i !== areaIdx && i !== ourIdx && !skip.test(norm[i]));
  if (compCols.length === 0) return null;
  const nameOf = (i: number) => header[i].replace(/\*/g, "").replace(/gmv/i, "").trim() || "Unknown";
  const rows: { area: string; competitor: string; their: number | null; our: number | null }[] = [];
  for (let li = 1; li < lines.length; li++) {
    const c = cells(lines[li]);
    if (!c[areaIdx] || /^[-:\s]+$/.test(c[areaIdx])) continue;
    const area = c[areaIdx].replace(/\*/g, "").trim();
    if (!area) continue;
    const our = parseMoney(c[ourIdx] || "");
    for (const ci of compCols) {
      const their = parseMoney(c[ci] || "");
      if (their == null && our == null) continue;
      rows.push({ area, competitor: nameOf(ci), their, our });
    }
  }
  return rows.length ? { rows } : null;
}
function parseNum(raw: string): number | null {
  const s = (raw || "").replace(/[^0-9.]/g, "");
  if (!s) return null;
  const n = Math.round(parseFloat(s));
  return isNaN(n) ? null : n;
}
function parseProductTable(text: string): { product: string; gmv: number | null; orders: number | null; units: number | null; areas: number | null }[] | null {
  const lines = (text || "").split("\n").map(l => l.trim()).filter(l => l.startsWith("|"));
  if (lines.length < 2) return null;
  const cells = (l: string) => l.split("|").slice(1, -1).map(c => c.trim());
  const header = cells(lines[0]);
  const norm = header.map(h => h.replace(/\*/g, "").toLowerCase().trim());
  const pIdx = norm.findIndex(h => h.includes("product") || h.includes("item"));
  if (pIdx < 0) return null;
  const gIdx = norm.findIndex(h => h.includes("gmv") || h.includes("sales") || h.includes("value"));
  const oIdx = norm.findIndex(h => h.includes("order"));
  const uIdx = norm.findIndex(h => h.includes("unit"));
  const aIdx = norm.findIndex(h => h.includes("area"));
  const rows: { product: string; gmv: number | null; orders: number | null; units: number | null; areas: number | null }[] = [];
  for (let i = 1; i < lines.length; i++) {
    const c = cells(lines[i]);
    const product = (c[pIdx] || "").replace(/\*/g, "").trim();
    if (!product || /^[-:\s]+$/.test(product)) continue;
    rows.push({ product, gmv: gIdx >= 0 ? parseMoney(c[gIdx] || "") : null, orders: oIdx >= 0 ? parseNum(c[oIdx] || "") : null, units: uIdx >= 0 ? parseNum(c[uIdx] || "") : null, areas: aIdx >= 0 ? parseNum(c[aIdx] || "") : null });
  }
  return rows.length ? rows : null;
}

export default function DashboardPage() {
    useActivityHeartbeat(typeof window !== "undefined" ? localStorage.getItem("tf_session_id") : null);
  const router = useRouter();
  const [user, setUser] = useState<Staff | null>(null);
  const [msgUnreadCount, setMsgUnreadCount] = useState(0);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [reports, setReports] = useState<Report[]>([]);
  const [loading, setLoading] = useState(true);
  const [showModal, setShowModal] = useState(false);
  const [activeTab, setActiveTab] = useState<"tasks" | "my_report" | "all_reports" | "analytics" | "outlet_reports" | "owner_outlets" | "history" | "attendance" | "sales_target" | "payout" | "reconciliation" | "competition" | "item_perf" | "ceo_report" | "fines" | "niranjana_report" | "pnl" | "contribution_margins" | "net_realisation" | "cash_flow" | "sales_forecast" | "cheques" | "auto_reviews" | "purchase_vendors" | "team_dashboards" | "notify" | "messages" | "active_status" | "help" | "production" | "ops_audits" | "outlet_targets" | "admin_dq">("tasks");
  // Brownie mode's tab-switch loader — bumping this tick is what tells
  // BrownieLoader to play its whole->broken animation + crack sound. Only
  // does anything when the viewer has Brownie mode turned on (personal,
  // per-device choice from the theme picker); everyone else's tab switches
  // stay exactly as quiet and instant as before.
  const [brownieTick, setBrownieTick] = useState(0);
  const fireBrownieTransition = () => { if (getStoredBrownieMode()) setBrownieTick((t) => t + 1); };
  // Real, purchase-data-backed food cost % (trailing 30 days, company-wide) —
  // replaces the flat 29.4% assumption in Outlet P&L, Channel P&L and Command
  // Centre. Falls back to 29.4% if there's no purchase/revenue data yet in the
  // window, so nothing breaks while Purchase & Vendors data is still thin.
  const [realFoodCostPct, setRealFoodCostPct] = useState<number | null>(null);
  const [realFoodCostWindow, setRealFoodCostWindow] = useState<{ from: string; to: string } | null>(null);
  useEffect(() => {
    fetchRealFoodCostPct().then(({ pct, windowFrom, windowTo }) => {
      setRealFoodCostPct(pct);
      setRealFoodCostWindow({ from: windowFrom, to: windowTo });
    }).catch((err) => console.error("real food cost fetch failed", err));
  }, []);
  const cogsRate = (realFoodCostPct ?? 29.4) / 100;
  const fmtWin = (w: { from: string; to: string } | null) => w ? `${new Date(w.from + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}–${new Date(w.to + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}` : "";
  // Real, payout-data-backed online commission % (trailing 60 days,
  // company-wide) — replaces the flat 50% assumption in Outlet P&L, Channel
  // P&L and Command Centre. Falls back to 50% if there isn't enough settled
  // Swiggy/Zomato payout data yet in the window.
  const [realCommissionPct, setRealCommissionPct] = useState<number | null>(null);
  const [realCommissionWindow, setRealCommissionWindow] = useState<{ from: string; to: string } | null>(null);
  useEffect(() => {
    fetchRealCommissionPct().then(({ pct, windowFrom, windowTo }) => {
      setRealCommissionPct(pct);
      setRealCommissionWindow({ from: windowFrom, to: windowTo });
    }).catch((err) => console.error("real commission fetch failed", err));
  }, []);
  const commissionRate = (realCommissionPct ?? 50) / 100;
  const RANGE_PRESETS = [
    { id: "yesterday", label: "Yesterday" },
    { id: "last7", label: "Last 7 days" },
    { id: "last30", label: "Last 30 days" },
    { id: "mtd", label: "Month to date" },
    { id: "lastmonth", label: "Last month" },
    { id: "custom", label: "Custom" },
  ];
  // Separate preset list just for the Outlet Health period comparison — adds explicit
  // week/month options useful for "this week vs last week" or "this month vs last month"
  // style comparisons. Kept separate from RANGE_PRESETS so Download Custom Report is untouched.
  const COMPARE_PRESETS = [
    { id: "thisweek", label: "This week" },
    { id: "lastweek", label: "Last week" },
    { id: "last7", label: "Last 7 days" },
    { id: "mtd", label: "This month (so far)" },
    { id: "lastmonth", label: "Last month" },
    { id: "prevmonth", label: "Month before last" },
    { id: "custom", label: "Custom" },
  ];
  const [outletRangeSel, setOutletRangeSel] = useState<Record<string, { preset: string; from?: string; to?: string }>>({});
  const getOutletSel = (o: string) => outletRangeSel[o] || { preset: "last30" };
  const resolveOutletRange = (sel: { preset: string; from?: string; to?: string }): { from: string; to: string; label: string } => {
    const today = new Date();
    const iso = (d: Date) => { const y = d.getFullYear(); const m = String(d.getMonth() + 1).padStart(2, "0"); const dd = String(d.getDate()).padStart(2, "0"); return `${y}-${m}-${dd}`; };
    if (sel.preset === "yesterday") { const y = new Date(today); y.setDate(y.getDate() - 1); return { from: iso(y), to: iso(y), label: "Yesterday" }; }
    if (sel.preset === "last7") { const f = new Date(today); f.setDate(f.getDate() - 6); return { from: iso(f), to: iso(today), label: "Last 7 days" }; }
    if (sel.preset === "last30") { const f = new Date(today); f.setDate(f.getDate() - 29); return { from: iso(f), to: iso(today), label: "Last 30 days" }; }
    if (sel.preset === "mtd") { const f = new Date(today.getFullYear(), today.getMonth(), 2); return { from: iso(f), to: iso(today), label: "Month to date" }; }
    if (sel.preset === "lastmonth") { const f = new Date(today.getFullYear(), today.getMonth() - 1, 2); const t = new Date(today.getFullYear(), today.getMonth(), 1); return { from: iso(f), to: iso(t), label: "Last month" }; }
    if (sel.preset === "prevmonth") { const f = new Date(today.getFullYear(), today.getMonth() - 2, 2); const t = new Date(today.getFullYear(), today.getMonth() - 1, 1); return { from: iso(f), to: iso(t), label: "Month before last" }; }
    if (sel.preset === "thisweek") { const dow = (today.getDay() + 6) % 7; const f = new Date(today); f.setDate(f.getDate() - dow); return { from: iso(f), to: iso(today), label: "This week" }; }
    if (sel.preset === "lastweek") { const dow = (today.getDay() + 6) % 7; const thisMon = new Date(today); thisMon.setDate(thisMon.getDate() - dow); const t = new Date(thisMon); t.setDate(t.getDate() - 1); const f = new Date(t); f.setDate(f.getDate() - 6); return { from: iso(f), to: iso(t), label: "Last week" }; }
    const shiftDay = (dateStr: string, days: number) => { const d = new Date(dateStr + "T00:00:00"); d.setDate(d.getDate() + days); return iso(d); };
    const rawFrom = sel.from || iso(today), rawTo = sel.to || iso(today);
    const f = shiftDay(rawFrom, 1), t = shiftDay(rawTo, 1);
    return { from: f, to: t, label: `${rawFrom} → ${rawTo}` };
  };
  const fetchPerOutletReports = async (outlets: string[]) => {
    const results = await Promise.all(outlets.map(async (o) => {
      const r = resolveOutletRange(getOutletSel(o));
      const { data } = await supabase.from("outlet_reports").select("*").eq("outlet_id", o).gte("report_date", r.from).lte("report_date", r.to).order("report_date", { ascending: true });
      return (data || []).map((row: any) => ({ ...row, _rangeLabel: r.label }));
    }));
    return results.flat();
  };
   const salesDateOf = (reportDate: string) => {
    const d = new Date(reportDate + "T00:00:00");
    d.setDate(d.getDate() - 1);
    const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), dd = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dd}`;
  };
  const buildRangeRows = (rows: any[]) => rows.map(r => {
    const shop = Number(r.shop_sales_value) || 0, sw = Number(r.swiggy_sales_value) || 0, zo = Number(r.zomato_sales_value) || 0;
    return { Date: salesDateOf(r.report_date), Outlet: (typeof OUTLET_NAMES !== "undefined" ? (OUTLET_NAMES as any)[r.outlet_id] : r.outlet_id) || r.outlet_id, RangeLabel: r._rangeLabel || "", Shop: shop, Swiggy: sw, Zomato: zo, Total: shop + sw + zo, "Shop Orders": Number(r.shop_sales_count) || 0, "Swiggy Orders": Number(r.swiggy_sales_count) || 0, "Zomato Orders": Number(r.zomato_sales_count) || 0, Target: Number(r.target) || 0, Discount: Number(r.discount_given) || 0, Late: r.is_late ? "Yes" : "No", Issues: r.issues || "" };
  });
  const downloadRangeExcel = async () => {
    setRepBusy(true);
    try {
      const targetOutlets = repOutlets.length ? repOutlets : (canAssign ? OUTLETS : (user?.outlets || []));
      const rows = await fetchPerOutletReports(targetOutlets);
      if (rows.length === 0) { alert("No reports found for the chosen range(s)/outlets."); setRepBusy(false); return; }
      const data = buildRangeRows(rows);
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(data);
      XLSX.utils.book_append_sheet(wb, ws, "Daily Rows");
      const byOutlet: Record<string, any> = {};
      data.forEach(d => { const k = d.Outlet; if (!byOutlet[k]) byOutlet[k] = { Outlet: k, Range: d.RangeLabel, Days: 0, Shop: 0, Swiggy: 0, Zomato: 0, Total: 0 }; byOutlet[k].Days++; byOutlet[k].Shop += d.Shop; byOutlet[k].Swiggy += d.Swiggy; byOutlet[k].Zomato += d.Zomato; byOutlet[k].Total += d.Total; });
      const summary = Object.values(byOutlet).map((s: any) => ({ ...s, "Avg/Day": Math.round(s.Total / s.Days) }));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), "Summary");
     XLSX.writeFile(wb, `OutletReports_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (e: any) { alert("Export failed: " + (e?.message || "error")); }
    setRepBusy(false);
  };
  const loadH2P = (): Promise<any> => new Promise((res, rej) => { const w = window as any; if (w.html2pdf) return res(w.html2pdf); const s = document.createElement("script"); s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"; s.onload = () => res((window as any).html2pdf); s.onerror = () => rej(new Error("pdf lib failed")); document.body.appendChild(s); });
 const downloadRangePDF = async () => {
    setRepBusy(true);
    window.scrollTo(0, 0); await new Promise((r) => setTimeout(r, 50));
    let h2p: any; try { h2p = await loadH2P(); } catch { alert("Could not load the PDF tool."); setRepBusy(false); return; }
    try {
      const targetOutlets = repOutlets.length ? repOutlets : (canAssign ? OUTLETS : (user?.outlets || []));
      const rows = await fetchPerOutletReports(targetOutlets);
      if (rows.length === 0) { alert("No reports found for the chosen range(s)/outlets."); setRepBusy(false); return; }
      const data = buildRangeRows(rows);
      const rs = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
      const _ym = new Date().toISOString().slice(0, 7);
      const TGT_BY_NAME: Record<string, number> = {};
      const MTGT_BY_NAME: Record<string, number> = {};
      Object.keys(MONTHLY_BASE).forEach(oid => { const nm = (OUTLET_NAMES as any)[oid] || oid; MTGT_BY_NAME[nm] = monthlyTargetFor(oid, _ym); TGT_BY_NAME[nm] = dailyTargetFor(oid, _ym); });
      const dayTgt = (d: any) => (TGT_BY_NAME[d.Outlet] || 0);
      const pctCol = (pc: number) => pc >= 100 ? C.green : (pc >= 60 ? C.gold : C.red);
      const byO: Record<string, any> = {};
      data.forEach(d => { if (!byO[d.Outlet]) byO[d.Outlet] = { Outlet: d.Outlet, Range: d.RangeLabel, Days: 0, Shop: 0, Swiggy: 0, Zomato: 0, Total: 0, Target: 0, Discount: 0 }; const b = byO[d.Outlet]; b.Days++; b.Shop += d.Shop; b.Swiggy += d.Swiggy; b.Zomato += d.Zomato; b.Total += d.Total; b.Discount += (d.Discount || 0); });
      const summ = Object.values(byO) as any[];
      const grand = summ.reduce((s, x) => s + x.Total, 0);
      const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", gold: "#C8901E", line: "#EADBC2", green: "#2E7D32", red: "#C62828" };
      const avgPerDay = (s: any) => s.Days > 0 ? s.Total / s.Days : 0;
      const comment = (perDay: number) => {
        if (perDay >= 80000) return ["🔥", "On fire! Crushing it 🤑", C.green];
        if (perDay >= 50000) return ["💪", "Strong stuff, keep going 👏", C.green];
        if (perDay >= 30000) return ["🙂", "Doing okay, room to grow 📈", C.gold];
        if (perDay >= 15000) return ["😬", "A bit slow… push harder 🏃", C.gold];
        if (perDay > 0)      return ["🥲", "Yikes, needs serious love 🚑", C.red];
        return ["💤", "Fast asleep — wake up! ⏰", C.soft];
      };
      const _pm = new Date(); const daysInMonth = new Date(_pm.getFullYear(), _pm.getMonth() + 1, 0).getDate();
      const gPerDay = summ.length > 0 ? summ.reduce((a: number, s: any) => a + avgPerDay(s), 0) : 0;
      const gVerdict = gPerDay >= 80000 * summ.length ? "🎉 Money machine go brrr! We're cooking 🧑‍🍳" : gPerDay >= 40000 * Math.max(summ.length, 1) ? "😎 Solid month, brownies are selling 🍫" : "😅 Could be tastier — let's hustle next month 💸";
      const sumRows = summ.map(s => { const [emo] = comment(avgPerDay(s));const stgt = MTGT_BY_NAME[s.Outlet] || 0; const proj = s.Days > 1 ? Math.round(s.Total / (s.Days - 1) * daysInMonth) : s.Total; const dpc = s.Total > 0 ? (s.Discount / s.Total) * 100 : 0; return `<tr><td style="padding:7px 10px;border-bottom:1px solid ${C.line};font-weight:600;color:${C.ink};font-size:12px">${emo} ${s.Outlet}<br/><span style="font-size:9px;font-weight:400;color:${C.soft}">${s.Range}</span></td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${s.Days}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${rs(s.Shop)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${rs(s.Swiggy)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${rs(s.Zomato)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-weight:800;color:${C.ink};font-size:12px">${rs(s.Total)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${stgt > 0 ? rs(stgt) : "-"}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;font-weight:800;color:${stgt > 0 ? pctCol((s.Total / stgt) * 100) : C.soft}">${stgt > 0 ? ((s.Total / stgt) * 100).toFixed(0) + "%" : "-"}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-weight:700;font-size:12px;color:${stgt > 0 ? pctCol((proj / stgt) * 100) : C.ink}">${rs(proj)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;font-weight:700;color:${dpc >= 15 ? C.red : dpc >= 8 ? C.gold : C.ink}">${s.Discount > 0 ? dpc.toFixed(1) + "%" : "-"}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${s.Discount > 0 ? rs(s.Discount) : "-"}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.ink}">${rs(Math.max(0, s.Total - s.Discount) * 0.05)}</td></tr>`; }).join("");
      const gT = summ.reduce((a: any, s: any) => { const stgt = MTGT_BY_NAME[s.Outlet] || 0; const proj = s.Days > 1 ? Math.round(s.Total / (s.Days - 1) * daysInMonth) : s.Total; a.Shop += s.Shop; a.Swiggy += s.Swiggy; a.Zomato += s.Zomato; a.Total += s.Total; a.Target += stgt; a.Proj += proj; a.Disc += (s.Discount || 0); return a; }, { Shop: 0, Swiggy: 0, Zomato: 0, Total: 0, Target: 0, Proj: 0, Disc: 0 });
      const totalRow = `<tr style="background:${C.line};border-top:2px solid ${C.ink}"><td style="padding:9px 10px;color:${C.ink};font-size:12px;font-weight:900">TOTAL</td><td style="padding:9px 10px;text-align:right;color:${C.soft};font-size:12px">—</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${rs(gT.Shop)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${rs(gT.Swiggy)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${rs(gT.Zomato)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:13px;font-weight:900">${rs(gT.Total)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${rs(gT.Target)}</td><td style="padding:9px 10px;text-align:right;font-size:12px;font-weight:900;color:${gT.Target > 0 ? pctCol((gT.Total / gT.Target) * 100) : C.soft}">${gT.Target > 0 ? ((gT.Total / gT.Target) * 100).toFixed(0) + "%" : "-"}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:13px;font-weight:900">${rs(gT.Proj)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${gT.Total > 0 ? ((gT.Disc / gT.Total) * 100).toFixed(1) + "%" : "-"}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:700">${rs(gT.Disc)}</td><td style="padding:9px 10px;text-align:right;color:${C.ink};font-size:12px;font-weight:900">${rs(Math.max(0, gT.Total - gT.Disc) * 0.05)}</td></tr>`;
      const noteCards = summ.map(s => { const [emo, msg, col] = comment(avgPerDay(s)); return `<div style="display:flex;align-items:center;gap:8px;background:${C.card};border:1px solid ${C.line};border-left:4px solid ${col};border-radius:8px;padding:8px 12px;margin-bottom:7px"><span style="font-size:18px">${emo}</span><span style="font-size:12px;color:${C.ink};font-weight:600;min-width:120px">${s.Outlet}</span><span style="font-size:12px;color:${col};font-style:italic">${msg}</span></div>`; }).join("");
      const dayRowHtml = (d: any) => `<tr><td style="padding:3px 8px;border-bottom:1px solid ${C.line};font-size:10px;color:${C.soft}">${d.Date}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${rs(d.Shop)}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${rs(d.Swiggy)}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${rs(d.Zomato)}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-weight:700;color:${C.ink};font-size:10px">${rs(d.Total)}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${dayTgt(d) > 0 ? rs(dayTgt(d)) : "-"}</td><td style="padding:3px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;font-weight:700;color:${dayTgt(d) > 0 ? pctCol((d.Total / dayTgt(d)) * 100) : C.soft}">${dayTgt(d) > 0 ? ((d.Total / dayTgt(d)) * 100).toFixed(0) + "%" : "-"}</td></tr>`;
      const dayTablesByOutlet = summ.map((s: any) => {
      const outletRows = [...data].filter((d: any) => d.Outlet === s.Outlet).sort((a: any, b: any) => (a.Date as string).localeCompare(b.Date));
        return `<div style="margin-bottom:16px;page-break-inside:avoid;break-inside:avoid;"><div style="font-size:12px;font-weight:800;margin:8px 0 4px;color:${C.ink}">${s.Outlet} — ${s.Range}</div><table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden;page-break-inside:avoid;break-inside:avoid;"><thead><tr style="background:${C.ink}"><th style="padding:5px 8px;text-align:left;color:#FFF6E5;font-size:9px">DATE</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">SHOP</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">SWIGGY</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">ZOMATO</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">TOTAL</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">TARGET</th><th style="padding:5px 8px;text-align:right;color:#FFF6E5;font-size:9px">%</th></tr></thead><tbody>${outletRows.map(dayRowHtml).join("")}</tbody></table></div>`;
      }).join("");    
      const html = `<div style="width:1120px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink}">
        <div style="background:linear-gradient(135deg,${C.ink},#5C3A22);padding:22px 32px">
          <div style="font-size:22px;font-weight:800;color:#FFF6E5">🍫 Brownie Heaven — Outlet Reports</div>
          <div style="font-size:12px;color:#E0A52E;letter-spacing:1px">📅 Custom per-outlet ranges · ${repOutlets.length ? repOutlets.length + " outlet(s)" : "all outlets"}</div>
        </div>
        <div style="padding:24px 32px">
          <div style="background:${C.card};border:1px solid ${C.line};border-radius:14px;padding:18px;text-align:center;margin-bottom:18px">
            <div style="font-size:12px;color:${C.soft};text-transform:uppercase;letter-spacing:1px">💰 Total sales in range</div>
            <div style="font-size:34px;font-weight:900;color:${C.ink}">${rs(grand)}</div>
          <div style="font-size:11px;color:${C.soft}">${data.length} daily reports across ${summ.length} outlet(s)</div>
            <div style="margin-top:10px;font-size:14px;font-weight:700;color:${C.ink}">${gVerdict}</div>
          </div>
          <div style="font-size:15px;font-weight:800;margin:6px 0 10px">📊 Summary by outlet — with the honest verdict 👀</div>
        <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:12px;overflow:hidden;margin-bottom:14px">
            <thead><tr style="background:${C.ink}"><th style="padding:8px 10px;text-align:left;color:#FFF6E5;font-size:10px">OUTLET</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">DAYS</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">SHOP</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">SWIGGY</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">ZOMATO</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">TOTAL</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">TARGET</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">%</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">PROJECTION</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">DISC %</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">DISCOUNT</th><th style="padding:8px 10px;text-align:right;color:#FFF6E5;font-size:10px">TAXES</th></tr></thead>
       <tbody>${sumRows}${totalRow}</tbody>
          </table>
         ${(() => {
            if (summ.length === 0) return "";
            const tShop = summ.reduce((a, s) => a + s.Shop, 0);
            const tSw = summ.reduce((a, s) => a + s.Swiggy, 0);
            const tZo = summ.reduce((a, s) => a + s.Zomato, 0);
            const tAll = tShop + tSw + tZo || 1;
            const R = 60, CX = 75, CY = 75, SW = 26, CIRC = 2 * Math.PI * R;
            let acc = 0;
            const segs = [[tShop, "#FACC15"], [tSw, "#FB923C"], [tZo, "#EF4444"]].map(([v, c]: any) => { const frac = v / tAll; const len = frac * CIRC; const off = -acc * CIRC; acc += frac; return `<circle cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="${c}" stroke-width="${SW}" stroke-dasharray="${len} ${CIRC - len}" stroke-dashoffset="${off}" transform="rotate(-90 ${CX} ${CY})"></circle>`; }).join("");
            const leg = (c: string, n: string, v: number) => `<div style="display:flex;align-items:center;gap:7px;margin-bottom:5px"><span style="width:11px;height:11px;background:${c};border-radius:2px;display:inline-block"></span><span style="font-size:12px;color:${C.ink};font-weight:600;min-width:62px">${n}</span><span style="font-size:12px;color:${C.soft}">${rs(v)} · ${((v / tAll) * 100).toFixed(0)}%</span></div>`;
            const CLOUD = ["Pallavaram", "Velachery", "Vadapalani"];
            const dineIn = summ.filter(s => !CLOUD.includes(s.Outlet));
            const byTotal = [...summ].sort((a, b) => b.Total - a.Total);
            const star = byTotal[0], slug = byTotal[byTotal.length - 1];
            const onShare = (s: any) => s.Total > 0 ? ((s.Swiggy + s.Zomato) / s.Total) * 100 : 0;
            const onlinePool = dineIn.length > 0 ? dineIn : summ;
            const mostOnline = [...onlinePool].sort((a, b) => onShare(b) - onShare(a))[0];
            const bestShop = [...onlinePool].sort((a, b) => b.Shop - a.Shop)[0];
            const multi = summ.length > 1;
            const card = (emoji: string, title: string, name: string, val: string, quip: string, accent: string) => `<div style="flex:1;min-width:150px;background:${C.card};border:1px solid ${C.line};border-top:4px solid ${accent};border-radius:12px;padding:13px 15px"><div style="font-size:10px;letter-spacing:1px;text-transform:uppercase;color:${C.soft};margin-bottom:5px">${emoji} ${title}</div><div style="font-size:15px;font-weight:800;color:${C.ink}">${name}</div><div style="font-size:12px;font-weight:700;color:${accent};margin:2px 0 4px">${val}</div><div style="font-size:10px;color:${C.soft};font-style:italic">${quip}</div></div>`;
            const awards = multi ? `
              <div style="display:flex;flex-wrap:wrap;gap:10px">
                ${card("🏆", "Star Outlet", star.Outlet, rs(star.Total), "The MVP carrying the team 💪", C.gold)}
                ${card("🐌", "Needs Help", slug.Outlet, rs(slug.Total), "Send backup… and a hug 🫂", "#C62828")}
                ${card("📱", "Most Online", mostOnline.Outlet, onShare(mostOnline).toFixed(0) + "% online", "Living that delivery life 🛵", "#3B82F6")}
                ${card("🏪", "Best Walk-in", bestShop.Outlet, rs(bestShop.Shop), "People show up here 🚶", "#2E7D32")}
              </div>` : `
              <div style="display:flex;flex-wrap:wrap;gap:10px">
                ${card("📱", "Online share", star.Outlet, onShare(star).toFixed(0) + "%", onShare(star) > 60 ? "Delivery is the bread & butter 🛵" : "Nice walk-in balance 🚶", "#3B82F6")}
                ${card("🏪", "Walk-in sales", star.Outlet, rs(star.Shop), "Counter's doing work 💪", "#2E7D32")}
                ${card("🛵", "Delivery sales", star.Outlet, rs(star.Swiggy + star.Zomato), "Swiggy + Zomato combined 📦", "#FB923C")}
              </div>`;
            return `
            <div style="display:flex;gap:18px;align-items:flex-start;margin:18px 0;page-break-inside:avoid">
            <div style="text-align:center">
              <div style="font-size:13px;font-weight:800;color:${C.ink};margin-bottom:6px">📱 Channel mix</div>
              <svg width="150" height="150" viewBox="0 0 150 150">
                <circle cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="#EADBC2" stroke-width="${SW}"></circle>
                ${segs}
                <text x="${CX}" y="${CY - 4}" text-anchor="middle" style="font-size:14px;font-weight:800;fill:${C.ink}">${rs(tAll)}</text>
                <text x="${CX}" y="${CY + 12}" text-anchor="middle" style="font-size:8px;fill:${C.soft};letter-spacing:1px">TOTAL</text>
              </svg>
              <div style="margin-top:10px;text-align:left">
                ${leg("#FACC15", "Shop", tShop)}
                ${leg("#FB923C", "Swiggy", tSw)}
                ${leg("#EF4444", "Zomato", tZo)}
              </div>
            </div>
            <div style="flex:1">
              <div style="font-size:15px;font-weight:800;margin-bottom:10px;color:${C.ink}">🏅 ${multi ? "Awards — hall of fame &amp; shame" : "Outlet snapshot"}</div>
              ${awards}
            </div>
          </div>`;
          })()}
          ${(() => {
            const DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            const full: Record<string, string> = { Sun: "Sunday", Mon: "Monday", Tue: "Tuesday", Wed: "Wednesday", Thu: "Thursday", Fri: "Friday", Sat: "Saturday" };
            const tot = [0, 0, 0, 0, 0, 0, 0], cnt = [0, 0, 0, 0, 0, 0, 0];
            data.forEach(d => { const sd = new Date(d.Date + "T00:00:00"); sd.setDate(sd.getDate() - 1); const wd = sd.getDay(); tot[wd] += d.Total; cnt[wd]++; });
            const avg = tot.map((t, i) => cnt[i] > 0 ? t / cnt[i] : 0);
            const maxAvg = Math.max(...avg, 1);
            const order = [1, 2, 3, 4, 5, 6, 0];
            const bestI = avg.indexOf(Math.max(...avg));
            const nz = avg.map((v, i) => ({ v, i })).filter(x => x.v > 0).sort((a, b) => a.v - b.v);
            const worstI = nz.length ? nz[0].i : bestI;
            const bars = order.map(i => { const pct = Math.max(2, (avg[i] / maxAvg) * 100); const best = i === bestI && avg[i] > 0; return `<div style="margin-bottom:9px"><div style="font-size:11px;font-weight:700;color:${best ? C.gold : C.ink};margin-bottom:3px">${best ? "🏆 " : ""}${full[DOW[i]]} — ${rs(avg[i])}/day</div><div style="height:14px;background:${C.line};border-radius:7px;overflow:hidden"><div style="height:14px;width:${pct}%;background:${best ? C.gold : "#D9C3A0"}"></div></div></div>`; }).join("");
            const quip = (bestI === 0 || bestI === 6) ? "Weekends are the goldmine 🤑 — staff up Fri–Sun!" : "Midweek is quietly carrying the month 💪 — most people bet on weekends.";
            return `
          <div style="font-size:15px;font-weight:800;margin:18px 0 10px;color:${C.ink}">📅 Which day sells best? <span style="font-size:11px;font-weight:400;color:${C.soft}">(avg sales per weekday)</span></div>
          <div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px 18px;margin-bottom:18px;page-break-inside:avoid">
            ${bars}
            <div style="margin-top:10px;font-size:12px;color:${C.ink};font-weight:700">🏆 Best: ${full[DOW[bestI]]} (${rs(avg[bestI])}/day) &nbsp;·&nbsp; 😴 Slowest: ${full[DOW[worstI]]} (${rs(avg[worstI])}/day)</div>
            <div style="font-size:11px;color:${C.soft};font-style:italic;margin-top:4px">${quip}</div>
          </div>
          <div style="font-size:13px;font-weight:800;margin:4px 0 8px;color:${C.ink}">💬 The honest verdict 👀</div>
          ${noteCards}`;
          })()}
          <div style="font-size:15px;font-weight:800;margin:18px 0 10px">🎯 Monthly targets by brand</div>
          <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:12px;overflow:hidden">
            <thead><tr style="background:${C.ink}"><th style="padding:7px 8px;text-align:left;color:#FFF6E5;font-size:9px">OUTLET</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">BH</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">CBH</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">ICBH</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">MONTHLY</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">DAILY</th></tr></thead>
            <tbody>${OUTLETS.map((o) => { const bt = (BRAND_TARGETS as any)[o]; const daily = parseFloat(OUTLET_TARGETS[o] || "0") || 0; const monthly = bt ? bt.total : daily * 30; const nm = (OUTLET_NAMES as any)[o] || o; return `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;color:${C.ink}">${nm}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${bt ? rs(bt.bh) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${bt ? rs(bt.cbh) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${bt && bt.icbh ? rs(bt.icbh) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;font-weight:700;color:${C.ink}">${rs(monthly)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${rs(daily)}</td></tr>`; }).join("")}</tbody>
          </table>
          <div style="font-size:9px;color:${C.soft};margin:6px 2px 0">BH = Brownie Heaven · CBH = Cakes by Brownie Heaven · ICBH = Ice Creams by Brownie Heaven. Vadapalani & Besant Nagar pending official brand split — showing current default.</div>
          <div style="font-size:15px;font-weight:800;margin:18px 0 10px">🗓️ Daily detail — by outlet's own range</div>
          ${dayTablesByOutlet}
          <div style="text-align:center;font-size:10px;color:${C.soft};margin-top:18px">🍫 Brownie Heaven · generated ${new Date().toISOString().slice(0, 10)}</div>
        </div></div>`;
      const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.innerHTML = html; document.body.appendChild(holder);
      try { await h2p().set({ margin: 0, filename: `OutletReports_${new Date().toISOString().slice(0, 10)}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: "#FAF3E7" }, jsPDF: { unit: "pt", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
      finally { document.body.removeChild(holder); }
    } catch (e: any) { alert("Export failed: " + (e?.message || "error")); }
    setRepBusy(false);
  };
  const [expandedOutlet, setExpandedOutlet] = useState<string | null>(null);
  const _today = new Date().toISOString().split("T")[0];
  const _mStart = _today.slice(0, 8) + "01";
  const [repOutlets, setRepOutlets] = useState<string[]>([]);
  const [repQuickRange, setRepQuickRange] = useState("last30");
  const [repQuickFrom, setRepQuickFrom] = useState("");
  const [repQuickTo, setRepQuickTo] = useState("");
  const [repCustomizePerOutlet, setRepCustomizePerOutlet] = useState(false);
  const activeRepOutlets = () => repOutlets.length ? repOutlets : (canAssign ? OUTLETS : (user?.outlets || []));
  const applyQuickRange = (presetId: string) => {
    setRepQuickRange(presetId);
    const updated = { ...outletRangeSel };
    activeRepOutlets().forEach((o) => { updated[o] = presetId === "custom" ? { preset: "custom", from: repQuickFrom, to: repQuickTo } : { preset: presetId }; });
    setOutletRangeSel(updated);
  };
  const applyQuickCustom = (from: string, to: string) => {
    setRepQuickFrom(from); setRepQuickTo(to);
    const updated = { ...outletRangeSel };
    activeRepOutlets().forEach((o) => { updated[o] = { preset: "custom", from, to }; });
    setOutletRangeSel(updated);
  };
  useEffect(() => {
    if (repCustomizePerOutlet) return;
    const updated = { ...outletRangeSel };
    let changed = false;
    activeRepOutlets().forEach((o) => { if (!updated[o]) { updated[o] = repQuickRange === "custom" ? { preset: "custom", from: repQuickFrom, to: repQuickTo } : { preset: repQuickRange }; changed = true; } });
    if (changed) setOutletRangeSel(updated);
    /* eslint-disable-next-line react-hooks/exhaustive-deps */
  }, [repOutlets]);
  const [repBusy, setRepBusy] = useState(false);
  const [outletFilter, setOutletFilter] = useState("all");
  const [reportData, setReportData] = useState<Record<string, string>>({});
  const [reportSubmitting, setReportSubmitting] = useState(false);
  const [reportOffDay, setReportOffDay] = useState(false);
  const [offToday, setOffToday] = useState<string[]>([]);
  const [attendanceData, setAttendanceData] = useState({ present: "", absent: "", late: "", absent_names: "", late_names: "", overtime_hours: "" });
  const [attendanceSubmitting, setAttendanceSubmitting] = useState(false);
  const [todayAttendance, setTodayAttendance] = useState<any>(null);
  const [attendanceDate, setAttendanceDate] = useState<string>(() => { const d = new Date(); d.setDate(d.getDate() - 1); return d.toISOString().split("T")[0]; });
  const [attendanceWeekOvertime, setAttendanceWeekOvertime] = useState<number | null>(null);
  // Staff registry for People & Labour — Nilani registers names under one of
  // three departments (Production / Sales & Outlet Staff / Housekeeping),
  // same soft-delete pattern as kitchen_chefs (active:false keeps history).
  const LABOUR_DEPARTMENTS = ["Production", "Sales & Outlet Staff", "Housekeeping"] as const;
  const [kLaborStaff, setKLaborStaff] = useState<any[]>([]);
  const [kLaborNewName, setKLaborNewName] = useState<Record<string, string>>({});
  const [kLaborEditId, setKLaborEditId] = useState<string | null>(null);
  const [kLaborEditName, setKLaborEditName] = useState("");
  const fetchLaborStaff = async () => { const { data } = await supabase.from("labour_staff").select("*").eq("active", true).order("name"); setKLaborStaff(data || []); };
  const addLaborStaff = async (dept: string) => {
    const name = (kLaborNewName[dept] || "").trim();
    if (!name) return;
    const { error } = await supabase.from("labour_staff").insert({ name, department: dept });
    if (error) { alert("Failed: " + error.message); return; }
    setKLaborNewName((prev) => ({ ...prev, [dept]: "" }));
    fetchLaborStaff();
  };
  const saveLaborStaffEdit = async () => {
    if (!kLaborEditId || !kLaborEditName.trim()) { setKLaborEditId(null); return; }
    const { error } = await supabase.from("labour_staff").update({ name: kLaborEditName.trim() }).eq("id", kLaborEditId);
    if (error) { alert("Failed: " + error.message); return; }
    setKLaborEditId(null); setKLaborEditName(""); fetchLaborStaff();
  };
  const deleteLaborStaff = async (id: string, name: string) => {
    if (!confirm(`Remove ${name} from the staff list?`)) return;
    const { error } = await supabase.from("labour_staff").update({ active: false }).eq("id", id);
    if (error) { alert("Failed: " + error.message); return; }
    fetchLaborStaff();
  };
  useEffect(() => { if (user?.role === "HR" || user?.role === "Owner") fetchLaborStaff(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [user]);
  const [salesTargets, setSalesTargets] = useState<Record<string, any>>({});
  const [stEditing, setStEditing] = useState<string | null>(null);
  const [stDate, setStDate] = useState<string>(() => new Date(Date.now() - 86400000).toISOString().split("T")[0]);
  // Outlet P&L lets you browse to ANY past date via the date picker — the
  // global cogsRate/commissionRate above always reflect the real current
  // month, so browsing to e.g. 29 August was showing September's rate
  // labeled "real, 1 Sept-30 Sept" next to August's numbers, which is wrong.
  // These two track whichever month stDate is currently pointing at instead,
  // and are what Outlet P&L's own calculations use.
  const [stCogsPct, setStCogsPct] = useState<number | null>(null);
  const [stCogsWindow, setStCogsWindow] = useState<{ from: string; to: string } | null>(null);
  const [stCommissionPct, setStCommissionPct] = useState<number | null>(null);
  const [stCommissionWindow, setStCommissionWindow] = useState<{ from: string; to: string } | null>(null);
  const stDateYm = stDate.slice(0, 7); // "YYYY-MM" — only refetch when the month actually changes
  useEffect(() => {
    const [y, m] = stDateYm.split("-").map(Number);
    fetchRealFoodCostPct({ year: y, month: m }).then(({ pct, windowFrom, windowTo }) => {
      setStCogsPct(pct);
      setStCogsWindow({ from: windowFrom, to: windowTo });
    }).catch((err) => console.error("outlet P&L real food cost fetch failed", err));
    fetchRealCommissionPct({ year: y, month: m }).then(({ pct, windowFrom, windowTo }) => {
      setStCommissionPct(pct);
      setStCommissionWindow({ from: windowFrom, to: windowTo });
    }).catch((err) => console.error("outlet P&L real commission fetch failed", err));
  }, [stDateYm]);
  const stCogsRate = (stCogsPct ?? 29.4) / 100;
  const stCommissionRate = (stCommissionPct ?? 50) / 100;
  // Plain labels that say whether the % is real or just the fallback guess,
  // so a placeholder is never presented as if it were measured data.
  const stCommissionSrc = stCommissionPct === null ? "placeholder — no real payout data for this period" : `real ${fmtWin(stCommissionWindow)}`;
  const stCogsSrc = stCogsPct === null ? "placeholder — no real purchase data for this period" : `real ${fmtWin(stCogsWindow)}`;
  const [stFiles, setStFiles] = useState<Record<string, { mis?: File; pnl?: File }>>({});
  const [stUpload, setStUpload] = useState<Record<string, any>>({});
  const [stUpBusy, setStUpBusy] = useState<string>("");
  const [stUpMsg, setStUpMsg] = useState<Record<string, string>>({});
  const [stEditValues, setStEditValues] = useState<Record<string, string>>({});
  const [stSaving, setStSaving] = useState(false);
  const [stActuals, setStActuals] = useState<Record<string, { daily: Record<string, { shop: number; online: number }>; monthNet: number; monthOnline: number }>>({});
  const reportDateForSalesDate = (salesDate: string) => {
    const d = new Date(salesDate + "T00:00:00");
    d.setDate(d.getDate() + 1);
    const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), dd = String(d.getDate()).padStart(2, "0");
    return `${y}-${m}-${dd}`;
  };
  const fetchSalesTargetActuals = async () => {
    const monthStart = stDate.slice(0, 7) + "-01";
    const [yy, mm] = stDate.slice(0, 7).split("-").map(Number);
    const lastDay = new Date(yy, mm, 0).getDate();
    const monthEnd = stDate.slice(0, 7) + "-" + String(lastDay).padStart(2, "0");
    const reportFrom = reportDateForSalesDate(monthStart);
    const reportTo = reportDateForSalesDate(monthEnd);
    const outletsToFetch = canAssign ? OUTLETS : (user?.outlets || []);
    if (!outletsToFetch.length) return;
    const { data } = await supabase.from("outlet_reports").select("outlet_id, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value").in("outlet_id", outletsToFetch).gte("report_date", reportFrom).lte("report_date", reportTo);
    const map: Record<string, { daily: Record<string, { shop: number; online: number }>; monthNet: number; monthOnline: number }> = {};
    (data || []).forEach((r: any) => {
      const salesDate = salesDateOf(r.report_date);
      if (!map[r.outlet_id]) map[r.outlet_id] = { daily: {}, monthNet: 0, monthOnline: 0 };
      const shop = Number(r.shop_sales_value) || 0;
      const online = (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
      map[r.outlet_id].daily[salesDate] = { shop, online };
      map[r.outlet_id].monthNet += shop;
      map[r.outlet_id].monthOnline += online;
    });
    setStActuals(map);
  };
  useEffect(() => { if (activeTab === "sales_target") fetchSalesTargetActuals(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, stDate]);
  const [todayReport, setTodayReport] = useState<Report | null>(null);
  const [overdueTask, setOverdueTask] = useState<Task | null>(null);
  const [forceAckReason, setForceAckReason] = useState("");
  const [taskTitle, setTaskTitle] = useState("");
  const [taskDesc, setTaskDesc] = useState("");
  const [taskAssignee, setTaskAssignee] = useState("arun");
  const [taskPriority, setTaskPriority] = useState("medium");
  const [taskDueHours, setTaskDueHours] = useState("4");
  const [taskOutlet, setTaskOutlet] = useState("");
  const [submitting, setSubmitting] = useState(false);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [showPinModal, setShowPinModal] = useState(false);
  const [newPin, setNewPin] = useState("");
  const [pinMsg, setPinMsg] = useState("");
  const [selectedReport, setSelectedReport] = useState<Report | null>(null);
  const [reportHistoryDate, setReportHistoryDate] = useState<string>(new Date().toISOString().split("T")[0]);
  const [reportByDate, setReportByDate] = useState<Report | null>(null);
  const [activeOutlet, setActiveOutlet] = useState<string>("");
  const [outletReports, setOutletReports] = useState<Record<string, OutletReport>>({});
  const [outletReportData, setOutletReportData] = useState<Record<string, string>>({});
  const [outletSubmitting, setOutletSubmitting] = useState(false);
  const [outletEntryDate, setOutletEntryDate] = useState<string>(() => new Date().toISOString().split("T")[0]);
  const [outletWasOff, setOutletWasOff] = useState(false);
  const [dipFlag, setDipFlag] = useState<null | { outlet: string; todayTotal: number; prevTotal: number; pct: number; signals: string[] }>(null);
  const [reviews, setReviews] = useState<any[]>([]);
  const [revForm, setRevForm] = useState<{ platform: string; rating: string; valid: boolean; refund: boolean; note: string }>({ platform: "Swiggy", rating: "5", valid: false, refund: false, note: "" });
  const [revSaving, setRevSaving] = useState(false);
  const [targetCheck, setTargetCheck] = useState<any[] | null>(null);
  const [targetReaction, setTargetReaction] = useState(false);
  const [anFrom, setAnFrom] = useState<string>(() => { const d = new Date(); d.setDate(d.getDate() - 29); return d.toISOString().slice(0, 10); });
  const [anTo, setAnTo] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [anRows, setAnRows] = useState<any[]>([]);
  const [anCard, setAnCard] = useState<string>("revenue_received");
  const [anLoading, setAnLoading] = useState(false);
  const [outletHealthData, setOutletHealthData] = useState<any[]>([]);
  const [outletHealthLoading, setOutletHealthLoading] = useState(false);
  const [outletHealthSel, setOutletHealthSel] = useState(OUTLETS[0]);
  const [outletHealthPdfBusy, setOutletHealthPdfBusy] = useState(false);
  const [outletHealthXlsxBusy, setOutletHealthXlsxBusy] = useState(false);
  const [outletDeepDive, setOutletDeepDive] = useState<any>(null);
  const [itemPerfContext, setItemPerfContext] = useState<any>(null);
  const [outletDeepDiveLoading, setOutletDeepDiveLoading] = useState(false);
  const [comparePresetA, setComparePresetA] = useState("none");
  const [comparePresetB, setComparePresetB] = useState("none");
  const [quickCompareResult, setQuickCompareResult] = useState<any>(null);
  const [quickCompareBusy, setQuickCompareBusy] = useState(false);
  const [compareAFrom, setCompareAFrom] = useState("");
  const [compareATo, setCompareATo] = useState("");
  const [compareBFrom, setCompareBFrom] = useState("");
  const [compareBTo, setCompareBTo] = useState("");
  const [scoreRows, setScoreRows] = useState<ScoreRow[]>([]);

  useEffect(() => {
    if (!user) return;
    let cancelled = false;
    computeScores().then((res) => { if (!cancelled) { setScoreRows(res.rows); } }).catch(() => {});
    return () => { cancelled = true; };
  }, [user]);

  const [compRows, setCompRows] = useState<any[]>([]);
  const [compSaving, setCompSaving] = useState(false);
  const [compForm, setCompForm] = useState({ competitor: "", area: "", address: "", sales_value: "", our_sales: "", our_outlet_id: "", period_label: "", period_date: new Date().toISOString().slice(0, 10), note: "" });
  const [compPaste, setCompPaste] = useState("");
  const [compPasteComp, setCompPasteComp] = useState("");
  const [compPasteLabel, setCompPasteLabel] = useState("");
  const [compPasteDate, setCompPasteDate] = useState(new Date().toISOString().slice(0, 10));
  const [compParsed, setCompParsed] = useState<{ area: string; competitor: string; their: number | null; our: number | null }[] | null>(null);
  const [compSavingBulk, setCompSavingBulk] = useState(false);
  const [compView, setCompView] = useState<"entry" | "products" | "insights">("entry");
  const [prodSide, setProdSide] = useState<"them" | "us">("them");
  const [prodComp, setProdComp] = useState("");
  const [prodLabel, setProdLabel] = useState("");
  const [prodDate, setProdDate] = useState(new Date().toISOString().slice(0, 10));
  const [prodPaste, setProdPaste] = useState("");
  const [prodParsed, setProdParsed] = useState<{ product: string; gmv: number | null; orders: number | null; units: number | null; areas: number | null }[] | null>(null);
  const [prodSaving, setProdSaving] = useState(false);
  const [compProducts, setCompProducts] = useState<any[]>([]);
  const [compHeadlineOn, setCompHeadlineOn] = useState(true);
  const [autoReviewsChecking, setAutoReviewsChecking] = useState(false);
  const [autoReviewsResult, setAutoReviewsResult] = useState<any>(null);
  const [autoReviewsList, setAutoReviewsList] = useState<any[]>([]);
  const fetchAutoReviews = async () => {
    const { data } = await supabase
      .from("outlet_reviews")
      .select("*")
      .eq("platform", "Google")
      .order("created_at", { ascending: false })
      .limit(20);
    setAutoReviewsList(data || []);
  };
  const checkNewReviews = async () => {
    setAutoReviewsChecking(true);
    setAutoReviewsResult(null);
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 65000); // safety cap
    try {
      const res = await fetch("/api/check-reviews", { signal: controller.signal });
      const rawText = await res.text();
      let json: any;
      try {
        json = JSON.parse(rawText);
      } catch {
        // The server returned something that isn't JSON — usually the platform's own
        // timeout page, not our route's response. Show a plain message instead of a
        // confusing "not valid JSON" parse error.
        setAutoReviewsResult({
          success: false,
          error: `The check took too long and the server cut it off before finishing (this can happen with a very large backlog). Click Check Now again — it picks up where it left off, since already-checked emails are remembered.`,
        });
        return;
      }
      setAutoReviewsResult(json);
      fetchAutoReviews();
    } catch (err: any) {
      const timedOut = err.name === "AbortError";
      setAutoReviewsResult({
        success: false,
        error: timedOut ? "Timed out after 65 seconds — the inbox check is taking too long. Try again, or there may be too many unread emails to process at once." : (err.message || String(err)),
      });
    } finally {
      clearTimeout(timeoutId);
      setAutoReviewsChecking(false);
    }
  };
  const [fineStaff, setFineStaff] = useState<string[]>([]);
  const [fineReason, setFineReason] = useState("1-star review");
  const [fineAmount, setFineAmount] = useState("50");
  const [fineDate, setFineDate] = useState(new Date().toISOString().split("T")[0]);
  const [fineOutlet, setFineOutlet] = useState("");
  const [fines, setFines] = useState<any[]>([]);
  const [fineBusy, setFineBusy] = useState(false);
    const fetchFines = async () => { const { data } = await supabase.from("fines").select("*").order("created_at", { ascending: false }).limit(50); setFines(data || []); };
  const [auditReports, setAuditReports] = useState<any[]>([]);
  const [auditReportsLoading, setAuditReportsLoading] = useState(false);
  const fetchAuditReports = async () => {
    setAuditReportsLoading(true);
    const { data } = await supabase.from("reports").select("report_date, report_data, submitted_at").eq("staff_id", "bharani").order("report_date", { ascending: false }).limit(60);
    setAuditReports(data || []);
    setAuditReportsLoading(false);
  };
  // Wastage & Complimentary — lives inside Item Performance (above the Atlas
  // upload section), reading the same expiry/complimentary boxes Ahila/Vishnu
  // already fill in on the daily Outlet Report. Kept simple on purpose: a
  // plain day-by-day log for the current month, plus a month-end per-product
  // read built only from real counts (how many separate days it showed up),
  // never invented. Scoped to the current calendar month — expiry/
  // complimentary only became compulsory 3 Oct, so there's no older history
  // worth looking back at yet.
  const [wastageRows, setWastageRows] = useState<any[]>([]);
  const [wastageLoading, setWastageLoading] = useState(false);
  const fetchWastage = async () => {
    setWastageLoading(true);
    const monthStart = new Date().toISOString().slice(0, 7) + "-01";
    const today = new Date().toISOString().slice(0, 10);
    const { data } = await supabase.from("outlet_reports")
      .select("outlet_id, report_date, expiry_count, expiry_items, complimentary_count, complimentary_items, complimentary_reason")
      .gte("report_date", monthStart).lte("report_date", today)
      .order("report_date", { ascending: false });
    setWastageRows(data || []);
    setWastageLoading(false);
  };
  const _isBlank = (v: any) => { const s = String(v ?? "").trim().toLowerCase(); return ["", "none", "nil", "na", "n/a", "-", "no", "nope", "0"].includes(s); };
  const wastageExpiryEntries = wastageRows.filter((r) => (Number(r.expiry_count) || 0) > 0 || !_isBlank(r.expiry_items));
  // ── Admin & Data Quality (Owner / Founder's Office) ──────────────────────────
  // Real checks only, on the last 7 full days. Nothing here is estimated: each
  // number is a count of actual rows (or missing rows) in the database.
  const [dqLoading, setDqLoading] = useState(false);
  const [dq, setDq] = useState<any>(null);
  const fetchDq = async () => {
    setDqLoading(true);
    const istNow = new Date(Date.now() + 5.5 * 60 * 60 * 1000);
    const iso = (d: Date) => d.toISOString().slice(0, 10);
    const today = iso(istNow);
    const days: string[] = [];
    for (let i = 7; i >= 1; i--) days.push(iso(new Date(istNow.getTime() - i * 86400000)));
    const from = days[0];
    const [repRes, prodRes, upRes, brandRes] = await Promise.all([
      supabase.from("outlet_reports").select("outlet_id, report_date, staff_id, is_late, is_backfill, shop_sales_value, swiggy_sales_value, zomato_sales_value, expiry_items, complimentary_items, complimentary_reason").gte("report_date", from).lte("report_date", today),
      supabase.from("kitchen_production").select("prod_date, qty").gte("prod_date", from).lte("prod_date", today),
      supabase.from("item_perf_uploads").select("label, created_at").order("created_at", { ascending: false }).limit(1),
      supabase.from("outlet_reports").select("outlet_id, report_date, bh_shop_sales_value").gte("report_date", BRAND_SPLIT_FROM).lte("report_date", today),
    ]);
    const reps = ((repRes.data as any[]) || []);
    const byKey: Record<string, any[]> = {};
    reps.forEach((r) => { const k = r.outlet_id + "|" + r.report_date; if (!byKey[k]) byKey[k] = []; byKey[k].push(r); });
    const matrix = OUTLETS.map((o) => ({
      outlet: o,
      cells: days.map((d) => {
        const rs = byKey[o + "|" + d] || [];
        if (rs.length === 0) return "missing";
        if (rs.some((r) => !r.is_late && !r.is_backfill)) return "ok";
        return rs.some((r) => r.is_late) ? "late" : "backfill";
      }),
    }));
    const nm = (o: string) => OUTLET_NAMES[o] || o.replace(/_/g, " ");
    const missing: string[] = [];
    matrix.forEach((m) => m.cells.forEach((c, i) => { if (c === "missing") missing.push(`${nm(m.outlet)} · ${days[i]}`); }));
    const inWindow = reps.filter((r) => r.report_date <= days[days.length - 1]);
    const late = inWindow.filter((r) => r.is_late).map((r) => `${nm(r.outlet_id)} · ${r.report_date}`);
    const backfilled = inWindow.filter((r) => r.is_backfill).map((r) => `${nm(r.outlet_id)} · ${r.report_date}`);
    const dupes = Object.keys(byKey).filter((k) => byKey[k].length > 1).map((k) => { const [o, d] = k.split("|"); return `${nm(o)} · ${d} (${byKey[k].length} reports)`; });
    const zero = reps.filter((r) => ((Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0)) === 0 && !r.is_backfill).map((r) => `${nm(r.outlet_id)} · ${r.report_date}`);
    const _empty = (v: any) => String(v ?? "").trim() === "";
    const blankReq = reps.filter((r) => r.report_date >= "2026-10-03" && (_empty(r.expiry_items) || _empty(r.complimentary_items) || _empty(r.complimentary_reason))).map((r) => `${nm(r.outlet_id)} · ${r.report_date}`);
    const prodRows = ((prodRes.data as any[]) || []);
    const prodDays = new Set(prodRows.map((p) => p.prod_date));
    const prodMissing = days.filter((d) => !prodDays.has(d));
    const brandColsMissing = !!brandRes.error;
    const brandMissing = brandColsMissing ? [] : ((brandRes.data as any[]) || []).filter((r) => r.bh_shop_sales_value === null || r.bh_shop_sales_value === undefined).map((r) => `${nm(r.outlet_id)} · ${r.report_date}`);
    const filedToday = new Set(reps.filter((r) => r.report_date === today).map((r) => r.outlet_id)).size;
    setDq({
      days, today, matrix, filedToday, missing, late, backfilled, dupes, zero, blankReq, prodMissing,
      brandColsMissing, brandMissing,
      latestUpload: ((upRes.data as any[]) || [])[0] || null,
      fetchError: repRes.error?.message || prodRes.error?.message || null,
      checkedAt: new Date().toLocaleString("en-IN"),
    });
    setDqLoading(false);
  };
  useEffect(() => { if (activeTab === "admin_dq" && (user?.role === "Owner" || user?.role === "Founder's Office")) fetchDq(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, user]);
  const wastageComplimentaryEntries = wastageRows.filter((r) => (Number(r.complimentary_count) || 0) > 0 || !_isBlank(r.complimentary_items) || !_isBlank(r.complimentary_reason));
  // Per-product month-end read. Splits the free-typed field on commas (staff
  // often list several items in one box, e.g. "Tiramisu -1, Butterscotch -2")
  // and strips trailing "-1"/"x2" quantity markers so the same product
  // mentioned on different days groups together. Counts DIFFERENT DAYS a
  // product showed up (not units, so one big batch on one day doesn't
  // outweigh a product that's genuinely recurring trouble) — the verdict
  // word is a plain bucket off that real count, and the outlet breakdown
  // underneath is the same real rows, just grouped by outlet instead of date.
  const wastageProductBreakdown = (entries: any[], field: string) => {
    const data: Record<string, { label: string; days: Set<string>; outlets: Record<string, Set<string>> }> = {};
    entries.forEach((r) => {
      const outletName = OUTLET_NAMES[r.outlet_id] || r.outlet_id;
      String(r[field] || "").split(",").map((s) => s.trim()).filter((s) => s && !_isBlank(s)).forEach((piece) => {
        const label = piece.replace(/[\s-]*[x-]\s*\d+\s*$/i, "").trim() || piece;
        const key = label.toLowerCase();
        if (!data[key]) data[key] = { label, days: new Set(), outlets: {} };
        data[key].days.add(r.report_date);
        if (!data[key].outlets[outletName]) data[key].outlets[outletName] = new Set();
        data[key].outlets[outletName].add(r.report_date);
      });
    });
    return Object.values(data).map((d) => {
      const n = d.days.size;
      const verdict = n >= 4 ? { label: "Recurring — worth checking", color: "text-red-400" } : n >= 2 ? { label: "Happened again", color: "text-yellow-400" } : { label: "One-off so far", color: "text-zinc-400" };
      const outlets = Object.entries(d.outlets).map(([o, dates]) => ({ outlet: o, days: dates.size })).sort((a, b) => b.days - a.days);
      return { name: d.label, days: n, outlets, ...verdict };
    }).sort((a, b) => b.days - a.days);
  };
  // ── Outlet manager review: last full week + month-to-date, own outlets only ──
  // Reports are filed the day AFTER the sales day (report_date = filing date),
  // so a Mon–Sun sales week = report dates Tue..next Mon. The week shown is the
  // most recent one that has ended, and it only rolls forward every Monday.
  const [ovRows, setOvRows] = useState<any[]>([]);
  const [ovLoading, setOvLoading] = useState(false);
  const [ovView, setOvView] = useState<"week" | "mtd">("week");
  const [ovOpen, setOvOpen] = useState<string | null>(null);
  const ovAddDays = (s: string, n: number) => { const d = new Date(s + "T00:00:00Z"); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
  const ovDates = () => {
    const today = new Date(Date.now() + 5.5 * 60 * 60 * 1000).toISOString().slice(0, 10);
    const dow = new Date(today + "T00:00:00Z").getUTCDay(); // 0 = Sun
    const lastMonday = ovAddDays(today, -((dow + 6) % 7));
    return { today, lastMonday, wkFrom: ovAddDays(lastMonday, -6), wkTo: lastMonday, prevFrom: ovAddDays(lastMonday, -13), prevTo: ovAddDays(lastMonday, -7), monthStart: today.slice(0, 7) + "-01" };
  };
  const fetchOv = async () => {
    const outs = ((user?.outlets || []) as string[]);
    if (!outs.length) return;
    setOvLoading(true);
    const D = ovDates();
    const from = D.prevFrom < D.monthStart ? D.prevFrom : D.monthStart;
    const { data } = await supabase.from("outlet_reports")
      .select("outlet_id, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value, expiry_count, expiry_items, complimentary_count, complimentary_items, complimentary_reason, bh_google_rating")
      .in("outlet_id", outs).gte("report_date", from).lte("report_date", D.today);
    setOvRows(data || []);
    setOvLoading(false);
  };
  useEffect(() => { if (activeTab === "outlet_reports" && ((user?.outlets || []) as string[]).length > 0 && !(user?.role === "Owner" || user?.role === "Manager")) fetchOv(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, user]);
  const [wastageExpanded, setWastageExpanded] = useState<string | null>(null);
  const saveFine = async () => {
    if (fineStaff.length === 0) { alert("Pick at least one person to fine."); return; }
    const amt = Number(fineAmount) || 0;
    setFineBusy(true);
      for (const sid of fineStaff) {
      const st = (ALL_STAFF as any[]).find((x) => x.id === sid);
      await supabase.from("fines").insert({ staff_id: sid, staff_name: st?.name || sid, reason: fineReason.trim() || null, amount: amt, outlet: fineOutlet || null, fine_date: fineDate, entered_by: user?.id || null });
      await supabase.from("point_adjustments").insert({ staff_id: sid, points: -amt, reason: `Fine: ${fineReason.trim() || "review"}` });
      fetch("/api/push/send", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ staff_ids: [sid], title: "⚖️ Fine issued", body: `₹${amt} — ${fineReason.trim() || "review"}`, tag: "fine" }),
      }).catch(() => {});
    }
    setFineBusy(false);
    setFineStaff([]);
    fetchFines();
  };
  useEffect(() => { if (user && (user.role === "Owner" || user.role === "Manager" || user.role === "Founder's Office")) fetchFines(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [user]);
   const [nrEntries, setNrEntries] = useState<any[]>([]);
  const [nrNewContent, setNrNewContent] = useState("");
  const [nrNewDate, setNrNewDate] = useState(new Date().toISOString().slice(0, 10));
  const [nrSaving, setNrSaving] = useState(false);
  const fetchNrEntries = async () => { const { data } = await supabase.from("niranjana_log").select("*").order("entry_date", { ascending: false }); setNrEntries(data || []); };
  const saveNrEntry = async () => {
    if (!nrNewContent.trim()) { alert("Write something first."); return; }
    setNrSaving(true);
    const { error } = await supabase.from("niranjana_log").insert({ entry_date: nrNewDate, content: nrNewContent.trim() });
    setNrSaving(false);
    if (error) { alert("Save failed: " + error.message); return; }
    setNrNewContent("");
    fetchNrEntries();
  };
  useEffect(() => { if (activeTab === "niranjana_report") fetchNrEntries(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);
  const [pnlFrom, setPnlFrom] = useState<string>(() => new Date().toISOString().slice(0, 8) + "01");
  const [pnlTo, setPnlTo] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [pnlRows, setPnlRows] = useState<any[]>([]);
  const [pnlLoading, setPnlLoading] = useState(false);
  const [pnlExpanded, setPnlExpanded] = useState<string | null>(null);
  const fetchPnl = async () => {
    setPnlLoading(true);
    const { data: salesRows } = await supabase.from("outlet_reports").select("outlet_id, shop_sales_value, swiggy_sales_value, zomato_sales_value").gte("report_date", pnlFrom).lte("report_date", pnlTo);
    const { data: targetRows } = await supabase.from("sales_target").select("outlet_id, brand, line_items").eq("brand", "BH");
       const fixedByOutlet: Record<string, any> = {};
    (targetRows || []).forEach((t: any) => {
      const f = t.line_items?.fixed || {};
      const ab = (v: any) => Math.abs(Number(v) || 0);
      const staff = ab(f.staff), rent = ab(f.rent), eb = ab(f.eb), transport = ab(f.transport), rm = 0.2 * rent, pest = ab(f.pest), water = ab(f.water), airtel = ab(f.airtel);
      fixedByOutlet[t.outlet_id] = { staff, rent, eb, transport, rm, pest, water, airtel, total: staff + rent + eb + transport + rm + pest + water + airtel };
    });
    const byOutlet: Record<string, { shop: number; swiggy: number; zomato: number }> = {};
    (salesRows || []).forEach((r: any) => {
      if (!byOutlet[r.outlet_id]) byOutlet[r.outlet_id] = { shop: 0, swiggy: 0, zomato: 0 };
      byOutlet[r.outlet_id].shop += Number(r.shop_sales_value) || 0;
      byOutlet[r.outlet_id].swiggy += Number(r.swiggy_sales_value) || 0;
      byOutlet[r.outlet_id].zomato += Number(r.zomato_sales_value) || 0;
    });
    const rows = OUTLETS.filter((o) => byOutlet[o]).map((o) => {
      const b = byOutlet[o];
      const chan = (sales: number, isOnline: boolean) => {
        const cogs = sales * cogsRate;
        const wastage = sales * 0.05;
        const commission = isOnline ? sales * commissionRate : 0;
        const contrib = sales - cogs - wastage - commission;
        return { sales, cogs, wastage, commission, contrib, margin: sales > 0 ? (contrib / sales) * 100 : 0 };
      };
      const shop = chan(b.shop, false), swiggy = chan(b.swiggy, true), zomato = chan(b.zomato, true);
      const totalSales = b.shop + b.swiggy + b.zomato;
      const totalContrib = shop.contrib + swiggy.contrib + zomato.contrib;
           const fixedBreakdown = fixedByOutlet[o] || { staff: 0, rent: 0, eb: 0, transport: 0, rm: 0, pest: 0, water: 0, airtel: 0, total: 0 };
      const fixed = fixedBreakdown.total;
      const netProfit = totalContrib - fixed;
      return { oid: o, name: OUTLET_NAMES[o] || o, shop, swiggy, zomato, totalSales, totalContrib, fixed, fixedBreakdown, netProfit, netMargin: totalSales > 0 ? (netProfit / totalSales) * 100 : 0 };
    });
    setPnlRows(rows);
    setPnlLoading(false);
  };
   useEffect(() => { if (activeTab === "tasks" && user?.role === "Financial Analyst") fetchPnl(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, user, pnlFrom, pnlTo]);
  const [cfWeeks, setCfWeeks] = useState<any[]>([]);
  const [cfLoading, setCfLoading] = useState(false);
  const fetchCashFlowForecast = async () => {
    setCfLoading(true);
    const today = new Date();
    const start = new Date(today); start.setDate(start.getDate() - 42); // last 6 weeks of history
    const { data: salesRows } = await supabase.from("outlet_reports").select("report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value").gte("report_date", start.toISOString().slice(0, 10));
    const { data: targetRows } = await supabase.from("sales_target").select("outlet_id, line_items").eq("brand", "BH");
    let totalWeeklyFixed = 0;
    (targetRows || []).forEach((t: any) => {
      const f = t.line_items?.fixed || {};
      const ab = (v: any) => Math.abs(Number(v) || 0);
      const monthly = ab(f.staff) + ab(f.rent) + ab(f.eb) + ab(f.transport) + 0.2 * ab(f.rent) + ab(f.pest) + ab(f.water) + ab(f.airtel);
      totalWeeklyFixed += monthly / 4.33;
    });
    // bucket historical sales into weeks (Mon-Sun) for a trend — shop and
    // online kept separate so commission (online-only) can be worked out
    // per week, not just total sales.
    const weekOf = (d: Date) => { const x = new Date(d); const day = x.getDay() || 7; x.setDate(x.getDate() - day + 1); return x.toISOString().slice(0, 10); };
    const byWeek: Record<string, { shop: number; online: number }> = {};
    (salesRows || []).forEach((r: any) => {
      const wk = weekOf(new Date(r.report_date));
      const shop = Number(r.shop_sales_value) || 0;
      const online = (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
      if (!byWeek[wk]) byWeek[wk] = { shop: 0, online: 0 };
      byWeek[wk].shop += shop;
      byWeek[wk].online += online;
    });
    const weekKeys = Object.keys(byWeek).sort();
    const recentWeeks = weekKeys.slice(-5, -1); // exclude current partial week
    const avgWeeklyShop = recentWeeks.length ? recentWeeks.reduce((a, k) => a + byWeek[k].shop, 0) / recentWeeks.length : 0;
    const avgWeeklyOnline = recentWeeks.length ? recentWeeks.reduce((a, k) => a + byWeek[k].online, 0) / recentWeeks.length : 0;
    // Real profit/EBITDA math — same real COGS % and real commission % used
    // everywhere else in the app (Outlet P&L, Command Centre), not a separate
    // guess. Previously this only subtracted fixed costs from sales, which
    // wasn't a real profit number at all.
    const buildRow = (week: string, shop: number, online: number, isForecast: boolean) => {
      const sales = shop + online;
      const cogs = cogsRate * sales;
      const wastage = 0.05 * sales;
      const commission = commissionRate * online;
      const contribution = sales - cogs - wastage - commission;
      const netProfit = contribution - totalWeeklyFixed;
      return { week, sales, cogs, wastage, commission, contribution, fixed: totalWeeklyFixed, netProfit, isForecast };
    };
    const historyRows = weekKeys.slice(-5).map((k) => buildRow(k, byWeek[k].shop, byWeek[k].online, false));
    const forecastRows: any[] = [];
    let lastMonday = new Date(weekOf(today));
    for (let i = 1; i <= 4; i++) {
      const wk = new Date(lastMonday); wk.setDate(wk.getDate() + i * 7);
      forecastRows.push(buildRow(wk.toISOString().slice(0, 10), avgWeeklyShop, avgWeeklyOnline, true));
    }
    setCfWeeks([...historyRows, ...forecastRows]);
    setCfLoading(false);
  };
  useEffect(() => { if (activeTab === "cash_flow") fetchCashFlowForecast(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  // ── Sales Forecast: outlet × channel × day ────────────────────────────
  // Real-data-only, from Nishant's Forecasting spec point 1 ("sales
  // forecasts by outlet, channel, category, product, day and hour").
  // What's actually buildable from what's in the system today:
  //   - outlet: yes — outlet_reports is per-outlet.
  //   - channel: yes — shop/Swiggy/Zomato are separate columns.
  //   - day: yes — daily rows going back months.
  //   - category / product: NOT buildable yet — Item Performance data is a
  //     single point-in-time upload each time, not a dated daily series, so
  //     there's no history to project a trend from at that level.
  //   - hour: NOT buildable — no hourly sales data exists anywhere.
  // Method: for each outlet+channel, average the same day-of-week over the
  // last ~10 weeks (day-of-week seasonality — Saturdays aren't like
  // Tuesdays), then nudge that by a short-term trend (last 14 days vs the
  // 14 before that, capped to ±30% so one unusual fortnight can't swing the
  // forecast wildly). This is a straightforward, explainable projection —
  // not a black-box model — so it's easy to sanity-check against what
  // actually happens.
  const [sfLoading, setSfLoading] = useState(false);
  const [sfOutlets, setSfOutlets] = useState<any[]>([]);
  const [sfMonthEnd, setSfMonthEnd] = useState<any[]>([]);
  const [sfMonthLabel, setSfMonthLabel] = useState("");
  const [sfDaysRemaining, setSfDaysRemaining] = useState(0);

  const fetchSalesForecast = async () => {
    setSfLoading(true);
    const todayIst = new Date(Date.now() + 5.5 * 60 * 60 * 1000);
    const todayStr = todayIst.toISOString().slice(0, 10);
    const historyStart = new Date(todayIst); historyStart.setDate(historyStart.getDate() - 70); // 10 weeks of history
    const historyStartStr = historyStart.toISOString().slice(0, 10);

    const SFBATCH = 1000;
    let all: any[] = [];
    let offset = 0;
    while (true) {
      const { data, error } = await supabase.from("outlet_reports").select("outlet_id, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value").gte("report_date", historyStartStr).lte("report_date", todayStr).order("report_date", { ascending: true }).range(offset, offset + SFBATCH - 1);
      if (error) { console.error("sales forecast fetch failed", error); break; }
      all = all.concat(data || []);
      if (!data || data.length < SFBATCH) break;
      offset += SFBATCH;
    }

    type DowBuckets = { shop: number[][]; swiggy: number[][]; zomato: number[][] };
    const perOutlet: Record<string, DowBuckets> = {};
    OUTLETS.forEach((o) => { perOutlet[o] = { shop: [[], [], [], [], [], [], []], swiggy: [[], [], [], [], [], [], []], zomato: [[], [], [], [], [], [], []] }; });

    // exclude today — it's a partial day and would drag the average down
    all.forEach((r: any) => {
      if (r.report_date === todayStr) return;
      const b = perOutlet[r.outlet_id];
      if (!b) return;
      const dow = new Date(r.report_date + "T00:00:00Z").getUTCDay();
      b.shop[dow].push(Number(r.shop_sales_value) || 0);
      b.swiggy[dow].push(Number(r.swiggy_sales_value) || 0);
      b.zomato[dow].push(Number(r.zomato_sales_value) || 0);
    });

    const avg = (arr: number[]) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);

    const trendFactor: Record<string, number> = {};
    OUTLETS.forEach((o) => {
      const rows = all.filter((r) => r.outlet_id === o && r.report_date !== todayStr).sort((a, b) => a.report_date.localeCompare(b.report_date));
      const totalOf = (rs: any[]) => rs.reduce((s, r) => s + (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0), 0);
      const last14 = totalOf(rows.slice(-14));
      const prev14 = totalOf(rows.slice(-28, -14));
      let f = prev14 > 0 ? last14 / prev14 : 1;
      f = Math.max(0.7, Math.min(1.3, f));
      trendFactor[o] = f;
    });

    const next7: string[] = [];
    for (let i = 1; i <= 7; i++) { const d = new Date(todayIst); d.setDate(d.getDate() + i); next7.push(d.toISOString().slice(0, 10)); }

    const outletForecasts = OUTLETS.map((o) => {
      const hasHistory = perOutlet[o].shop.some((a) => a.length > 0) || perOutlet[o].swiggy.some((a) => a.length > 0) || perOutlet[o].zomato.some((a) => a.length > 0);
      const days = next7.map((dateStr) => {
        const dow = new Date(dateStr + "T00:00:00Z").getUTCDay();
        const shop = avg(perOutlet[o].shop[dow]) * trendFactor[o];
        const swiggy = avg(perOutlet[o].swiggy[dow]) * trendFactor[o];
        const zomato = avg(perOutlet[o].zomato[dow]) * trendFactor[o];
        return { date: dateStr, shop, swiggy, zomato, total: shop + swiggy + zomato };
      });
      const weekTotal = days.reduce((s, d) => s + d.total, 0);
      return { oid: o, name: OUTLET_NAMES[o] || o, days, weekTotal, trendPct: (trendFactor[o] - 1) * 100, hasHistory };
    });

    // Month-end sales projection, per outlet: real month-to-date + forecast
    // for the remaining days. This is the SAME shared calculation the Command
    // Centre uses (src/lib/monthEndProjection.ts), so both show one number.
    const monthEndCalc = computeMonthEnd(all as any, OUTLETS, todayStr);
    const monthEndByOutlet: Record<string, OutletMonthEnd> = {};
    monthEndCalc.perOutlet.forEach((x) => { monthEndByOutlet[x.oid] = x; });

    // Per-outlet monthly fixed costs — same source and math as the Outlet
    // P&L tab (sales_target.line_items.fixed), so the numbers here line up
    // with what's shown there. These figures are already monthly, so no
    // /4.33 weekly conversion like the company-wide Profit Forecast does.
    const { data: sfTargetRows } = await supabase.from("sales_target").select("outlet_id, line_items").eq("brand", "BH");
    const sfFixedByOutlet: Record<string, number> = {};
    (sfTargetRows || []).forEach((t: any) => {
      const f = t.line_items?.fixed || {};
      const ab = (v: any) => Math.abs(Number(v) || 0);
      const staff = ab(f.staff), rent = ab(f.rent), eb = ab(f.eb), transport = ab(f.transport), rm = 0.2 * rent, pest = ab(f.pest), water = ab(f.water), airtel = ab(f.airtel);
      sfFixedByOutlet[t.outlet_id] = staff + rent + eb + transport + rm + pest + water + airtel;
    });

    const monthEnd = OUTLETS.map((o) => {
      const _me = monthEndByOutlet[o] || { oid: o, mtdShop: 0, mtdOnline: 0, remShop: 0, remOnline: 0 };
      const { mtdShop, mtdOnline, remShop, remOnline } = _me;
      const mtd = mtdShop + mtdOnline;
      const remaining = remShop + remOnline;
      const projShop = mtdShop + remShop;
      const projOnline = mtdOnline + remOnline;
      const projSales = projShop + projOnline;
      // Real P&L math — same real COGS %/commission % used everywhere else
      // in the app (Outlet P&L, Command Centre, Profit Forecast), applied
      // to the projected month-end sales, then this outlet's own real
      // monthly fixed costs subtracted to get a projected net profit.
      const cogs = cogsRate * projSales;
      const wastage = 0.05 * projSales;
      const commission = commissionRate * projOnline;
      const contribution = projSales - cogs - wastage - commission;
      const fixed = sfFixedByOutlet[o] || 0;
      const netProfit = contribution - fixed;
      return { oid: o, name: OUTLET_NAMES[o] || o, mtd, remaining, projected: projSales, cogs, wastage, commission, contribution, fixed, netProfit, hasFixedData: !!sfFixedByOutlet[o] };
    });

    setSfOutlets(outletForecasts);
    setSfMonthEnd(monthEnd);
    setSfMonthLabel(todayIst.toLocaleDateString("en-IN", { month: "long", year: "numeric" }));
    setSfDaysRemaining(monthEndCalc.daysRemaining);
    setSfLoading(false);
  };
  useEffect(() => { if (activeTab === "sales_forecast") fetchSalesForecast(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  // ── Production & Purchase Forecast (point 2 of Forecasting) ───────────
  // What Nishant asked for was "ingredient requirement, production
  // requirement and purchase requirement forecasts". A true ingredient
  // forecast (e.g. "you'll need 42kg flour next week") needs a recipe/
  // bill-of-materials — how much of each raw ingredient goes into one unit
  // of each product — and that doesn't exist anywhere in the system yet
  // (the only per-product cost data, FOOD_COST_MAP, is a single landed ₹
  // cost per product, not an ingredient breakdown). That piece has to wait
  // until that recipe data exists (Rafiq's side).
  // What IS real and buildable right now, from data already being logged:
  //   - Production requirement: kitchen_production logs actual units made
  //     per product per day — forecast next week's production the same way
  //     as sales (day-of-week average, trend-adjusted).
  //   - Purchase requirement: purchase_ledger logs actual raw-material
  //     purchases (qty + unit + ₹) per product per day — forecast next
  //     week's purchase quantity the same way, at a weekly grain since
  //     purchases are lumpy (not every ingredient is bought daily).
  const [pfLoading, setPfLoading] = useState(false);
  const [pfProducts, setPfProducts] = useState<any[]>([]);
  const [pfCategories, setPfCategories] = useState<any[]>([]);
  const [purchLoading, setPurchLoading] = useState(false);
  const [purchForecast, setPurchForecast] = useState<any[]>([]);

  const fetchProductionForecast = async () => {
    setPfLoading(true);
    const todayIst = new Date(Date.now() + 5.5 * 60 * 60 * 1000);
    const todayStr = todayIst.toISOString().slice(0, 10);
    const historyStart = new Date(todayIst); historyStart.setDate(historyStart.getDate() - 70);
    const { data: products } = await supabase.from("kitchen_products").select("*").eq("active", true).order("sort_order");
    const { data: prodRows } = await supabase.from("kitchen_production").select("prod_date, flavour, qty").gte("prod_date", historyStart.toISOString().slice(0, 10)).lte("prod_date", todayStr).order("prod_date", { ascending: true });

    const norm = (s: string) => (s || "").trim().toLowerCase();
    const byProduct: Record<string, number[][]> = {}; // name -> [dow][values]
    (products || []).forEach((p: any) => { byProduct[norm(p.name)] = [[], [], [], [], [], [], []]; });
    (prodRows || []).forEach((r: any) => {
      if (r.prod_date === todayStr) return; // exclude partial day
      const key = norm(r.flavour);
      if (!byProduct[key]) return; // custom/off-menu item made once — not enough history to forecast, skip
      const dow = new Date(r.prod_date + "T00:00:00Z").getUTCDay();
      byProduct[key][dow].push(Number(r.qty) || 0);
    });
    const avg = (arr: number[]) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);
    const trendFactor: Record<string, number> = {};
    (products || []).forEach((p: any) => {
      const key = norm(p.name);
      const rows = (prodRows || []).filter((r: any) => norm(r.flavour) === key && r.prod_date !== todayStr).sort((a: any, b: any) => a.prod_date.localeCompare(b.prod_date));
      const totalOf = (rs: any[]) => rs.reduce((s, r) => s + (Number(r.qty) || 0), 0);
      const last14 = totalOf(rows.slice(-14));
      const prev14 = totalOf(rows.slice(-28, -14));
      let f = prev14 > 0 ? last14 / prev14 : 1;
      f = Math.max(0.7, Math.min(1.3, f));
      trendFactor[key] = f;
    });

    const next7: string[] = [];
    for (let i = 1; i <= 7; i++) { const d = new Date(todayIst); d.setDate(d.getDate() + i); next7.push(d.toISOString().slice(0, 10)); }

    const productForecasts = (products || []).map((p: any) => {
      const key = norm(p.name);
      const hasHistory = byProduct[key]?.some((a) => a.length > 0);
      const weekQty = next7.reduce((s, dateStr) => {
        const dow = new Date(dateStr + "T00:00:00Z").getUTCDay();
        return s + avg(byProduct[key][dow]) * (trendFactor[key] ?? 1);
      }, 0);
      return { id: p.id, name: p.name, category: p.category || "Other", weekQty, hasHistory };
    }).filter((p: any) => p.hasHistory);

    const catTotals: Record<string, number> = {};
    productForecasts.forEach((p: any) => { catTotals[p.category] = (catTotals[p.category] || 0) + p.weekQty; });
    const categories = Object.entries(catTotals).map(([category, weekQty]) => ({ category, weekQty })).sort((a, b) => b.weekQty - a.weekQty);

    setPfProducts(productForecasts.sort((a: any, b: any) => b.weekQty - a.weekQty));
    setPfCategories(categories);
    setPfLoading(false);
  };
  useEffect(() => { if (activeTab === "sales_forecast") fetchProductionForecast(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  const fetchPurchaseForecast = async () => {
    setPurchLoading(true);
    const today = new Date();
    const historyStart = new Date(today); historyStart.setDate(historyStart.getDate() - 70);
    const PBATCH = 1000;
    let all: any[] = [];
    let offset = 0;
    while (true) {
      const { data, error } = await supabaseStock.from("purchase_ledger").select("date, product, category, qty, unit, amount").gte("date", historyStart.toISOString().slice(0, 10)).order("date", { ascending: true }).range(offset, offset + PBATCH - 1);
      if (error) { console.error("purchase forecast fetch failed", error); break; }
      all = all.concat(data || []);
      if (!data || data.length < PBATCH) break;
      offset += PBATCH;
    }

    const weekOf = (d: Date) => { const x = new Date(d); const day = x.getDay() || 7; x.setDate(x.getDate() - day + 1); return x.toISOString().slice(0, 10); };
    const key = (product: string, unit: string) => `${(product || "Unknown").trim()}__${(unit || "").trim()}`;
    const byProduct: Record<string, { product: string; category: string; unit: string; weeks: Record<string, { qty: number; amount: number }> }> = {};
    all.forEach((r: any) => {
      const k = key(r.product, r.unit);
      if (!byProduct[k]) byProduct[k] = { product: r.product || "Unknown", category: r.category || "Uncategorised", unit: r.unit || "", weeks: {} };
      const wk = weekOf(new Date(r.date));
      if (!byProduct[k].weeks[wk]) byProduct[k].weeks[wk] = { qty: 0, amount: 0 };
      byProduct[k].weeks[wk].qty += Number(r.qty) || 0;
      byProduct[k].weeks[wk].amount += Number(r.amount) || 0;
    });

    const results = Object.values(byProduct).map((p) => {
      const weekKeys = Object.keys(p.weeks).sort();
      const recentWeeks = weekKeys.slice(-5, -1); // exclude current partial week, use the 4 before it
      if (recentWeeks.length < 2) return null; // not enough purchase history for this item to forecast
      const avgQty = recentWeeks.reduce((s, k) => s + p.weeks[k].qty, 0) / recentWeeks.length;
      const avgAmount = recentWeeks.reduce((s, k) => s + p.weeks[k].amount, 0) / recentWeeks.length;
      const last2 = recentWeeks.slice(-2).reduce((s, k) => s + p.weeks[k].qty, 0);
      const prev2 = recentWeeks.slice(-4, -2).reduce((s, k) => s + p.weeks[k].qty, 0);
      let trend = prev2 > 0 ? last2 / prev2 : 1;
      trend = Math.max(0.7, Math.min(1.3, trend));
      return { product: p.product, category: p.category, unit: p.unit, forecastQty: avgQty * trend, forecastAmount: avgAmount * trend, trendPct: (trend - 1) * 100 };
    }).filter((x): x is NonNullable<typeof x> => !!x).sort((a, b) => b.forecastAmount - a.forecastAmount);

    setPurchForecast(results);
    setPurchLoading(false);
  };
  useEffect(() => { if (activeTab === "sales_forecast") fetchPurchaseForecast(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  // --- Purchase & Vendors (reads from the separate Production Stock database) ---
  const [pvFrom, setPvFrom] = useState<string>(() => { const d = new Date(); d.setDate(d.getDate() - 60); return d.toISOString().slice(0, 10); });
  const [pvTo, setPvTo] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [pvRows, setPvRows] = useState<any[]>([]);
  const [pvVendorSel, setPvVendorSel] = useState<string>("");
  // --- Approved discount % per outlet+brand, per week (Nishant's request) ---
  const [discWeekStart, setDiscWeekStart] = useState<string>(() => {
    const d = new Date();
    const day = d.getDay(); // 0=Sun
    const diffToMonday = day === 0 ? 6 : day - 1;
    d.setDate(d.getDate() - diffToMonday);
    return d.toISOString().slice(0, 10);
  });
  const [discTargets, setDiscTargets] = useState<Record<string, number>>({}); // key: outlet_brand
  const [discSaving, setDiscSaving] = useState<Record<string, boolean>>({});
  const fetchDiscountTargets = async () => {
    const { data } = await supabase.from("outlet_discount_targets").select("outlet_id,brand,approved_pct").eq("week_start", discWeekStart);
    const m: Record<string, number> = {};
    (data || []).forEach((r: any) => { m[`${r.outlet_id}_${r.brand}`] = Number(r.approved_pct); });
    setDiscTargets(m);
  };
  useEffect(() => { if (activeTab === "net_realisation") fetchDiscountTargets(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, discWeekStart]);
  const saveDiscountTarget = async (oid: string, brand: string, pct: number) => {
    const key = `${oid}_${brand}`;
    setDiscSaving((m) => ({ ...m, [key]: true }));
    const { error } = await supabase.from("outlet_discount_targets").upsert(
      { outlet_id: oid, brand, week_start: discWeekStart, approved_pct: pct, set_by: user?.id || null, updated_at: new Date().toISOString() },
      { onConflict: "outlet_id,brand,week_start" }
    );
    setDiscSaving((m) => ({ ...m, [key]: false }));
    if (error) { alert("Save failed: " + error.message); return; }
    setDiscTargets((m) => ({ ...m, [key]: pct }));
  };
  const [pvRevenueRows, setPvRevenueRows] = useState<any[]>([]);
  const [pvItemPerfCategories, setPvItemPerfCategories] = useState<{ label: string; rows: { category: string; revenue: number }[] } | null>(null);
  const [pvLoading, setPvLoading] = useState(false);
  const fetchPurchaseVendors = async () => {
    setPvLoading(true);
    // Supabase enforces a server-side "Max Rows" cap (usually 1000) on every
    // query, regardless of any .limit() requested client-side — so a single
    // request can silently come back truncated with no error. Fetching in
    // batches of 1000 and looping until a batch comes back short sidesteps
    // that cap entirely, however large the table grows in the future.
    const BATCH = 1000;
    let all: any[] = [];
    let offset = 0;
    while (true) {
      const { data, error } = await supabaseStock
        .from("purchase_ledger")
        .select("date,agency,product,category,qty,unit,rate,amount")
        .gte("date", pvFrom)
        .lte("date", pvTo)
        .order("date", { ascending: true })
        .range(offset, offset + BATCH - 1);
      if (error) { console.error(error); setPvRows(all); setPvLoading(false); return; }
      all = all.concat(data || []);
      if (!data || data.length < BATCH) break;
      offset += BATCH;
    }
    setPvRows(all);

    // Real revenue for the same date range — company-wide, batched for the
    // same reason as above (12 outlets × a wide range can exceed 1000 rows).
    let revAll: any[] = [];
    let revOffset = 0;
    while (true) {
      const { data: revBatch, error: revErr } = await supabase
        .from("outlet_reports")
        .select("shop_sales_value,swiggy_sales_value,zomato_sales_value")
        .gte("report_date", pvFrom)
        .lte("report_date", pvTo)
        .range(revOffset, revOffset + BATCH - 1);
      if (revErr) { console.error(revErr); break; }
      revAll = revAll.concat(revBatch || []);
      if (!revBatch || revBatch.length < BATCH) break;
      revOffset += BATCH;
    }
    setPvRevenueRows(revAll);

    // Latest Item Performance upload's category revenue — for the category-
    // level cost-vs-revenue check. This is whatever period that upload covers,
    // not necessarily the same window as the purchase date range above —
    // shown with its own label so the two periods are never conflated.
    const { data: uploads } = await supabase.from("item_perf_uploads").select("id,label").order("created_at", { ascending: false }).limit(1);
    if (uploads && uploads.length > 0) {
      const { data: ipRows } = await supabase.from("item_perf_rows").select("category,net_revenue").eq("upload_id", uploads[0].id);
      const byCategory: Record<string, number> = {};
      (ipRows || []).forEach((r: any) => { const c = r.category || "Uncategorised"; byCategory[c] = (byCategory[c] || 0) + (Number(r.net_revenue) || 0); });
      setPvItemPerfCategories({ label: uploads[0].label, rows: Object.entries(byCategory).map(([category, revenue]) => ({ category, revenue })) });
    } else {
      setPvItemPerfCategories(null);
    }

    setPvLoading(false);
  };
  useEffect(() => { if (activeTab === "purchase_vendors") fetchPurchaseVendors(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, pvFrom, pvTo]);

  const pvAnalysis = (() => {
    const totalSpend = pvRows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const byVendor: Record<string, number> = {};
    const byCategory: Record<string, number> = {};
    pvRows.forEach((r) => {
      byVendor[r.agency] = (byVendor[r.agency] || 0) + (Number(r.amount) || 0);
      byCategory[r.category || "Uncategorised"] = (byCategory[r.category || "Uncategorised"] || 0) + (Number(r.amount) || 0);
    });
    const vendorRanked = Object.entries(byVendor).sort((a, b) => b[1] - a[1]);
    const categoryRanked = Object.entries(byCategory).sort((a, b) => b[1] - a[1]);

    // Price-creep: for each (agency + product) pair bought more than once in this
    // window, compare the earliest and latest rate. Flags a real jump, not noise —
    // requires at least 2 purchases and a meaningful (>=10%) increase.
    const pairKey = (r: any) => `${r.agency}__${r.product}`;
    const byPair: Record<string, any[]> = {};
    pvRows.forEach((r) => { const k = pairKey(r); if (!byPair[k]) byPair[k] = []; byPair[k].push(r); });
    const priceCreep = Object.entries(byPair)
      .map(([, rows]) => {
        const sorted = [...rows].sort((a, b) => a.date.localeCompare(b.date));
        if (sorted.length < 2) return null;
        const first = sorted[0], last = sorted[sorted.length - 1];
        const firstRate = Number(first.rate) || 0, lastRate = Number(last.rate) || 0;
        if (firstRate <= 0) return null;
        const changePct = ((lastRate - firstRate) / firstRate) * 100;
        return { agency: first.agency, product: first.product, firstRate, lastRate, firstDate: first.date, lastDate: last.date, changePct, purchaseCount: sorted.length };
      })
      .filter((x): x is NonNullable<typeof x> => !!x && x.changePct >= 10)
      .sort((a, b) => b.changePct - a.changePct);

    // Vendor comparison: same product bought from 2+ different vendors in this
    // window — shows who's cheaper, using each vendor's average rate.
    const byProduct: Record<string, Record<string, number[]>> = {};
    pvRows.forEach((r) => {
      if (!byProduct[r.product]) byProduct[r.product] = {};
      if (!byProduct[r.product][r.agency]) byProduct[r.product][r.agency] = [];
      byProduct[r.product][r.agency].push(Number(r.rate) || 0);
    });
    const vendorComparison = Object.entries(byProduct)
      .map(([product, vendors]) => {
        const vendorAvgs = Object.entries(vendors).map(([agency, rates]) => ({ agency, avgRate: rates.reduce((s, v) => s + v, 0) / rates.length, count: rates.length }));
        if (vendorAvgs.length < 2) return null;
        vendorAvgs.sort((a, b) => a.avgRate - b.avgRate);
        const cheapest = vendorAvgs[0], priciest = vendorAvgs[vendorAvgs.length - 1];
        const gapPct = cheapest.avgRate > 0 ? ((priciest.avgRate - cheapest.avgRate) / cheapest.avgRate) * 100 : 0;
        return { product, vendorAvgs, cheapest, priciest, gapPct };
      })
      .filter((x): x is NonNullable<typeof x> => !!x && x.gapPct >= 10)
      .sort((a, b) => b.gapPct - a.gapPct);

    // Food cost % for the date range selected on THIS page — compared against
    // the app-wide figure (trailing 30 days, fixed window) now actually used in
    // Outlet P&L / Channel P&L / Command Centre. These two numbers can
    // legitimately differ if you pick a different range here than the
    // app-wide rolling window — that's expected, not a bug.
    const totalRevenue = pvRevenueRows.reduce((s, r) => s + (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0), 0);
    const selectedRangeFoodCostPct = totalRevenue > 0 ? (totalSpend / totalRevenue) * 100 : null;
    const assumedFoodCostPct = realFoodCostPct ?? 29.4;
    const modelGapPct = selectedRangeFoodCostPct != null ? selectedRangeFoodCostPct - assumedFoodCostPct : null;

    // Category-level: real ingredient spend vs Item Performance's category
    // revenue. Matched case-insensitively/trimmed — anything that doesn't
    // match on either side is listed explicitly, never silently dropped,
    // since a wording mismatch (e.g. "Ice cream" vs "Ice Cream") would
    // otherwise make a real category vanish without any sign it happened.
    const norm = (s: string) => (s || "").trim().toLowerCase();
    let categoryMatch: { category: string; revenue: number; spend: number; costPct: number }[] = [];
    let unmatchedSpendCategories: string[] = [];
    let unmatchedRevenueCategories: string[] = [];
    if (pvItemPerfCategories) {
      const revByNorm: Record<string, { label: string; revenue: number }> = {};
      pvItemPerfCategories.rows.forEach((r) => { revByNorm[norm(r.category)] = { label: r.category, revenue: r.revenue }; });
      const spendByNorm: Record<string, { label: string; spend: number }> = {};
      categoryRanked.forEach(([cat, amt]) => { spendByNorm[norm(cat)] = { label: cat, spend: amt }; });
      const allKeys = new Set([...Object.keys(revByNorm), ...Object.keys(spendByNorm)]);
      allKeys.forEach((k) => {
        const rev = revByNorm[k], sp = spendByNorm[k];
        if (rev && sp) {
          categoryMatch.push({ category: rev.label, revenue: rev.revenue, spend: sp.spend, costPct: rev.revenue > 0 ? (sp.spend / rev.revenue) * 100 : 0 });
        } else if (sp && !rev) {
          unmatchedSpendCategories.push(sp.label);
        } else if (rev && !sp) {
          unmatchedRevenueCategories.push(rev.label);
        }
      });
      categoryMatch.sort((a, b) => b.spend - a.spend);
    }

    return { totalSpend, vendorRanked, categoryRanked, priceCreep, vendorComparison, totalRevenue, selectedRangeFoodCostPct, assumedFoodCostPct, modelGapPct, categoryMatch, unmatchedSpendCategories, unmatchedRevenueCategories };
  })();

  // Drill-down into one selected vendor: everything they sold, whether their
  // pricing on each item beats or loses to the best price seen elsewhere.
  const pvVendorDrilldown = (() => {
    if (!pvVendorSel) return null;
    const vendorRows = pvRows.filter((r) => r.agency === pvVendorSel);
    if (vendorRows.length === 0) return null;
    const totalSpent = vendorRows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const byItem: Record<string, { qty: number; amount: number; rates: number[] }> = {};
    vendorRows.forEach((r) => {
      if (!byItem[r.product]) byItem[r.product] = { qty: 0, amount: 0, rates: [] };
      byItem[r.product].qty += Number(r.qty) || 0;
      byItem[r.product].amount += Number(r.amount) || 0;
      byItem[r.product].rates.push(Number(r.rate) || 0);
    });
    const items = Object.entries(byItem).map(([product, d]) => {
      const avgRate = d.rates.reduce((s, v) => s + v, 0) / d.rates.length;
      // How does this vendor's price for this item compare to the cheapest
      // other vendor selling the same item, in the same date range?
      const otherRates = pvRows.filter((r) => r.product === product && r.agency !== pvVendorSel).map((r) => Number(r.rate) || 0);
      const cheapestOther = otherRates.length ? Math.min(...otherRates) : null;
      const vsCheapestPct = cheapestOther && cheapestOther > 0 ? ((avgRate - cheapestOther) / cheapestOther) * 100 : null;
      return { product, qty: d.qty, amount: d.amount, avgRate, cheapestOther, vsCheapestPct };
    }).sort((a, b) => b.amount - a.amount);
    const rankPosition = pvAnalysis.vendorRanked.findIndex(([agency]) => agency === pvVendorSel) + 1;
    return { totalSpent, items, rankPosition, totalVendors: pvAnalysis.vendorRanked.length };
  })();

  const [nrFrom, setNrFrom] = useState<string>(() => { const d = new Date(); d.setDate(d.getDate() - 60); return d.toISOString().slice(0, 10); });
  const [nrTo, setNrTo] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [nrRows, setNrRows] = useState<any[]>([]);
  const [nrLoading, setNrLoading] = useState(false);
  const fetchNetRealisation = async () => {
    setNrLoading(true);
    const { data: payouts } = await supabase.from("outlet_payouts").select("*").gte("period_start", nrFrom).lte("period_end", nrTo);
    const { data: repRows } = await supabase.from("outlet_reports").select("outlet_id, report_date, zomato_sales_value").gte("report_date", nrFrom).lte("report_date", nrTo);
    const rows = (payouts || []).map((p: any) => {
      if (p.platform === "swiggy") {
        const gross = Number(p.customer_payable) || 0;
        const net = Number(p.amount_transferable) || 0;
        return { outletId: p.outlet_id, outlet: OUTLET_NAMES[p.outlet_id] || p.outlet_id, brand: p.brand || "—", platform: "Swiggy", periodStart: p.period_start, periodEnd: p.period_end, gross, net, pct: gross > 0 ? (net / gross) * 100 : null, verified: true };
      } else {
        const gross = (repRows || []).filter((r: any) => r.outlet_id === p.outlet_id && r.report_date >= p.period_start && r.report_date <= p.period_end).reduce((a: number, r: any) => a + (Number(r.zomato_sales_value) || 0), 0);
        const net = Number(p.net_payout) || 0;
        return { outletId: p.outlet_id, outlet: OUTLET_NAMES[p.outlet_id] || p.outlet_id, brand: p.brand || "—", platform: "Zomato", periodStart: p.period_start, periodEnd: p.period_end, gross, net, pct: gross > 0 ? (net / gross) * 100 : null, verified: false };
      }
    }).sort((a: any, b: any) => (a.periodStart < b.periodStart ? 1 : -1));
    setNrRows(rows);
    setNrLoading(false);
  };
  useEffect(() => { if (activeTab === "net_realisation") fetchNetRealisation(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, nrFrom, nrTo]);

  // --- Swiggy payout upload (feeds outlet_payouts, which fetchNetRealisation reads) ---
  const [autoPayoutChecking, setAutoPayoutChecking] = useState(false);
  const [autoPayoutResult, setAutoPayoutResult] = useState<any>(null);
  const [manualAssignPick, setManualAssignPick] = useState<Record<number, { outlet: string; brand: string }>>({});
  const [manualAssignBusy, setManualAssignBusy] = useState<Record<number, boolean>>({});
  const [manualAssignDone, setManualAssignDone] = useState<Record<number, boolean>>({});
  const saveManualPayout = async (item: any, idx: number) => {
    const pick = manualAssignPick[idx];
    if (!pick?.outlet || !pick?.brand) { alert("Pick both an outlet and a brand first."); return; }
    setManualAssignBusy((m) => ({ ...m, [idx]: true }));
    const parsedData = item.parsed || parseSwiggyPayoutEmail(item.fullText || "");
    const num = (s: any) => { const v = parseFloat(String(s ?? "").replace(/,/g, "")); return isNaN(v) ? null : v; };
    const int = (s: any) => { const v = parseInt(String(s ?? "").replace(/,/g, "")); return isNaN(v) ? null : v; };
    const payload: any = {
      outlet_id: pick.outlet,
      brand: pick.brand,
      platform: "swiggy",
      period_start: parsedData.period_start,
      period_end: parsedData.period_end,
      total_orders: int(parsedData.total_orders),
      customer_payable: num(parsedData.customer_payable),
      swiggy_service_fee: num(parsedData.swiggy_service_fee),
      other_charges_refund: num(parsedData.other_charges_refund),
      govt_taxes: num(parsedData.govt_taxes),
      amount_transferable: num(parsedData.amount_transferable),
      next_payout_cycle: parsedData.next_payout_cycle || null,
      next_payout_date: parsedData.next_payout_date || null,
      entry_method: "manual_assign",
      updated_at: new Date().toISOString(),
    };
    if (!payload.period_start || !payload.period_end) {
      alert("Couldn't find the payout period dates in this email — can't save without them.");
      setManualAssignBusy((m) => ({ ...m, [idx]: false }));
      return;
    }
    const { data: existing } = await supabase.from("outlet_payouts").select("id").eq("outlet_id", pick.outlet).eq("brand", pick.brand).eq("platform", "swiggy").eq("period_start", payload.period_start).eq("period_end", payload.period_end).maybeSingle();
    const { error } = existing
      ? await supabase.from("outlet_payouts").update(payload).eq("id", existing.id)
      : await supabase.from("outlet_payouts").insert(payload);
    setManualAssignBusy((m) => ({ ...m, [idx]: false }));
    if (error) { alert("Save failed: " + error.message); return; }
    setManualAssignDone((m) => ({ ...m, [idx]: true }));
    // If what we just saved falls outside the currently viewed date range, widen
    // the range to include it — otherwise it'd save correctly but the box would
    // still show "No data yet", which looks like a failure when it isn't one.
    if (payload.period_start < nrFrom) setNrFrom(payload.period_start);
    if (payload.period_end > nrTo) setNrTo(payload.period_end);
    fetchNetRealisation();
  };
  const checkNewPayouts = async () => {
    setAutoPayoutChecking(true);
    setAutoPayoutResult(null);
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 65000);
    try {
      const res = await fetch("/api/check-payouts", { signal: controller.signal });
      const rawText = await res.text();
      let json: any;
      try { json = JSON.parse(rawText); }
      catch {
        setAutoPayoutResult({ success: false, error: "The check took too long and the server cut it off before finishing. Click Check Now again — it picks up where it left off." });
        return;
      }
      setAutoPayoutResult(json);
      fetchNetRealisation();
    } catch (err: any) {
      const timedOut = err.name === "AbortError";
      setAutoPayoutResult({ success: false, error: timedOut ? "Timed out after 65 seconds — try again." : (err.message || String(err)) });
    } finally {
      clearTimeout(timeoutId);
      setAutoPayoutChecking(false);
    }
  };
  const [spOutlet, setSpOutlet] = useState("");
  const [spPeriodFrom, setSpPeriodFrom] = useState(() => { const d = new Date(); d.setDate(d.getDate() - 6); return d.toISOString().slice(0, 10); });
  const [spPeriodTo, setSpPeriodTo] = useState(() => new Date().toISOString().slice(0, 10));
  const [spFile, setSpFile] = useState<File | null>(null);
  const [spBrand, setSpBrand] = useState<"BH" | "CBH" | "ICBH">("BH");
  const [spParsed, setSpParsed] = useState<{ gross: number; net: number } | null>(null);
  const [spBusy, setSpBusy] = useState(false);
  const [spMsg, setSpMsg] = useState("");

  // Reads the "Payout Breakup" sheet of a Swiggy monthly/weekly annexure export.
  // Finds "Total Customer Paid" (gross) and "Net Payout" (net) by label, taking the
  // value in the column right after the label — this is the "Delivered Orders"
  // column, which is what should feed realisation (not cancelled/other columns).
  const parseSwiggyPayoutFile = async (file: File): Promise<{ gross: number; net: number } | null> => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheetName = wb.SheetNames.find((n) => n.toLowerCase().includes("payout breakup")) || wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, blankrows: false }) as any[][];
    const findValue = (label: string): number | null => {
      for (const row of rows) {
        const idx = row.findIndex((c: any) => typeof c === "string" && c.trim().toLowerCase() === label.toLowerCase());
        if (idx >= 0 && idx + 1 < row.length) {
          const v = row[idx + 1];
          const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[,₹\s]/g, ""));
          if (!isNaN(n)) return n;
        }
      }
      return null;
    };
    const gross = findValue("Total Customer Paid");
    const net = findValue("Net Payout");
    if (gross == null || net == null) return null;
    return { gross, net };
  };

  const onSpFile = async (e: any) => {
    const f = e.target.files?.[0]; if (!f) return;
    setSpFile(f); setSpMsg(""); setSpParsed(null);
    try {
      const result = await parseSwiggyPayoutFile(f);
      if (!result) { setSpMsg("Couldn't find \"Total Customer Paid\" or \"Net Payout\" in a \"Payout Breakup\" sheet — check this is the right file."); return; }
      setSpParsed(result);
    } catch (err: any) { setSpMsg("Couldn't read the file: " + (err?.message || "")); }
    e.target.value = "";
  };

  const saveSwiggyPayout = async () => {
    if (!spOutlet) { alert("Pick which outlet this file is for."); return; }
    if (!spParsed) { alert("Upload and parse a file first."); return; }
    setSpBusy(true);
    const { error } = await supabase.from("outlet_payouts").insert({
      outlet_id: spOutlet,
      brand: spBrand,
      platform: "swiggy",
      period_start: spPeriodFrom,
      period_end: spPeriodTo,
      customer_payable: spParsed.gross,
      amount_transferable: spParsed.net,
    });
    setSpBusy(false);
    if (error) { setSpMsg("Save failed: " + error.message); return; }
    setSpFile(null); setSpParsed(null); setSpOutlet("");
    fetchNetRealisation();
  };

  const [cmProductRows, setCmProductRows] = useState<any[]>([]);
  const [cmLoading, setCmLoading] = useState(false);
  const fetchContributionMargins = async () => {
    setCmLoading(true);
    const { data: latestUpload } = await supabase.from("item_perf_uploads").select("id").order("created_at", { ascending: false }).limit(1).maybeSingle();
    if (!latestUpload) { setCmProductRows([]); setCmLoading(false); return; }
    const { data: rows } = await supabase.from("item_perf_rows").select("name, net_revenue, units_sold").eq("upload_id", latestUpload.id);
    const withMargin = (rows || []).map((r: any) => {
      const unitCost = (FOOD_COST_MAP as any)[r.name];
      const revenue = Number(r.net_revenue) || 0;
      const units = Number(r.units_sold) || 0;
      if (unitCost === undefined) return { name: r.name, revenue, units, cost: null, margin: null, marginPct: null };
      const cost = unitCost * units;
      const margin = revenue - cost;
      const marginPct = revenue > 0 ? (margin / revenue) * 100 : null;
      return { name: r.name, revenue, units, cost, margin, marginPct };
    });
    setCmProductRows(withMargin);
    setCmLoading(false);
  };
  useEffect(() => { if (activeTab === "tasks" && user?.role === "Financial Analyst") fetchContributionMargins(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, user]);
  const [pnlPdfBusy, setPnlPdfBusy] = useState(false);
  const downloadPnlPDF = async () => {
    if (pnlRows.length === 0) { alert("No data to export for this range."); return; }
    setPnlPdfBusy(true);
    const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", amber: "#C8901E" };
    const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
    const totals = pnlRows.reduce((a, r) => ({ shop: a.shop + r.shop.sales, swiggy: a.swiggy + r.swiggy.sales, zomato: a.zomato + r.zomato.sales, sales: a.sales + r.totalSales, contrib: a.contrib + r.totalContrib, fixed: a.fixed + r.fixed, net: a.net + r.netProfit }), { shop: 0, swiggy: 0, zomato: 0, sales: 0, contrib: 0, fixed: 0, net: 0 });
    const rowsHtml = pnlRows.map((r) => `<tr><td style="padding:6px 8px;border-bottom:1px solid ${C.line};font-size:10px;font-weight:600">${r.name}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px">${inr(r.shop.sales)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px">${inr(r.swiggy.sales)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px">${inr(r.zomato.sales)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;font-weight:700">${inr(r.totalSales)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${inr(r.totalContrib)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${inr(r.fixed)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;font-weight:700;color:${r.netProfit >= 0 ? C.green : C.red}">${inr(r.netProfit)}</td><td style="padding:6px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;font-weight:700;color:${r.netMargin >= 0 ? C.green : C.red}">${r.netMargin.toFixed(1)}%</td></tr>`).join("");
    const totalRow = `<tr style="background:${C.line};border-top:2px solid ${C.ink}"><td style="padding:8px;font-size:10px;font-weight:900">TOTAL</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:700">${inr(totals.shop)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:700">${inr(totals.swiggy)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:700">${inr(totals.zomato)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:900">${inr(totals.sales)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:700">${inr(totals.contrib)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:700">${inr(totals.fixed)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:900;color:${totals.net >= 0 ? C.green : C.red}">${inr(totals.net)}</td><td style="padding:8px;text-align:right;font-size:10px;font-weight:900;color:${totals.net >= 0 ? C.green : C.red}">${totals.sales > 0 ? ((totals.net / totals.sales) * 100).toFixed(1) + "%" : "-"}</td></tr>`;
    const html = `<div style="width:1000px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px"><div style="font-size:22px;font-weight:900">Brownie Heaven — Channel P&amp;L</div><div style="font-size:11px;color:${C.soft};margin-bottom:16px">${pnlFrom} to ${pnlTo} · real fixed costs from Outlet P&amp;L · ${(cogsRate * 100).toFixed(1)}% COGS (real, month to date) · 5% wastage · ${(commissionRate * 100).toFixed(1)}% online commission (real, this + last calendar month)</div><table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden"><thead><tr style="background:${C.ink}"><th style="padding:8px;text-align:left;color:#FFF6E5;font-size:9px">OUTLET</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">SHOP</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">SWIGGY</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">ZOMATO</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">TOTAL SALES</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">CONTRIBUTION</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">FIXED COSTS</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">NET PROFIT</th><th style="padding:8px;text-align:right;color:#FFF6E5;font-size:9px">NET %</th></tr></thead><tbody>${rowsHtml}${totalRow}</tbody></table><div style="font-size:9px;color:${C.soft};margin-top:12px">Generated ${new Date().toISOString().split("T")[0]}</div></div>`;
    const lib = await loadH2P();
    window.scrollTo(0, 0); await new Promise((r) => setTimeout(r, 50));
    const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0"; holder.innerHTML = html; document.body.appendChild(holder);
    try { await lib().set({ margin: 0, filename: `PnL_${pnlFrom}_to_${pnlTo}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "landscape" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
    finally { document.body.removeChild(holder); }
    setPnlPdfBusy(false);
  };
  const [chqSuppliers, setChqSuppliers] = useState<any[]>([]);
  const [chqRows, setChqRows] = useState<any[]>([]);
  const [chqLoading, setChqLoading] = useState(false);
  const fetchChequeSuppliers = async () => { const { data } = await supabase.from("cheque_suppliers").select("*").order("name"); setChqSuppliers(data || []); };
  const fetchCheques = async () => { setChqLoading(true); const { data } = await supabase.from("cheques").select("*").order("date_issued", { ascending: false }); setChqRows(data || []); setChqLoading(false); };
  useEffect(() => { if (activeTab === "cheques") { fetchChequeSuppliers(); fetchCheques(); } /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);
  const [chqSupplierSel, setChqSupplierSel] = useState("");
  const [chqNewSupplierName, setChqNewSupplierName] = useState("");
  const [chqAmount, setChqAmount] = useState("");
  const [chqDate, setChqDate] = useState(new Date().toISOString().slice(0, 10));
  const [chqPhotoFront, setChqPhotoFront] = useState<File | null>(null);
  const [chqPhotoBack, setChqPhotoBack] = useState<File | null>(null);
  const [chqNotes, setChqNotes] = useState("");
  const [chqSaving, setChqSaving] = useState(false);
  const uploadChequePhoto = async (file: File, label: string) => {
    const ext = file.name.split(".").pop();
    const path = `${Date.now()}_${label}.${ext}`;
    const { error } = await supabase.storage.from("cheque-photos").upload(path, file);
    if (error) { alert("Photo upload failed: " + error.message); return null; }
    const { data } = supabase.storage.from("cheque-photos").getPublicUrl(path);
    return data.publicUrl;
  };
  const saveCheque = async () => {
    if (!chqSupplierSel) { alert("Pick a supplier."); return; }
    if (chqSupplierSel === "__new__" && !chqNewSupplierName.trim()) { alert("Enter the new supplier's name."); return; }
    if (!chqAmount || Number(chqAmount) <= 0) { alert("Enter a valid amount."); return; }
    setChqSaving(true);
    let supplierId = chqSupplierSel;
    let supplierName = chqSuppliers.find((s) => s.id === chqSupplierSel)?.name || "";
    if (chqSupplierSel === "__new__") {
      const { data: newSup, error: supErr } = await supabase.from("cheque_suppliers").insert({ name: chqNewSupplierName.trim() }).select().single();
      if (supErr) { alert("Could not add supplier: " + supErr.message); setChqSaving(false); return; }
      supplierId = newSup.id;
      supplierName = newSup.name;
    }
    let frontUrl = null, backUrl = null;
    if (chqPhotoFront) frontUrl = await uploadChequePhoto(chqPhotoFront, "front");
    if (chqPhotoBack) backUrl = await uploadChequePhoto(chqPhotoBack, "back");
    const { error } = await supabase.from("cheques").insert({
      supplier_id: supplierId, supplier_name: supplierName, amount: Number(chqAmount), date_issued: chqDate,
      photo_front_url: frontUrl, photo_back_url: backUrl, notes: chqNotes.trim() || null, entered_by: user?.id || null,
    });
    setChqSaving(false);
    if (error) { alert("Save failed: " + error.message); return; }
    setChqSupplierSel(""); setChqNewSupplierName(""); setChqAmount(""); setChqPhotoFront(null); setChqPhotoBack(null); setChqNotes("");
    fetchChequeSuppliers(); fetchCheques();
  };
  const toggleChequeStatus = async (c: any) => {
    const newStatus = c.status === "cleared" ? "pending" : "cleared";
    await supabase.from("cheques").update({ status: newStatus }).eq("id", c.id);
    fetchCheques();
  };
  const [ceoWin, setCeoWin] = useState("7");
  const [ceoRepRows, setCeoRepRows] = useState<any[]>([]);
  const [ceoMonthRep, setCeoMonthRep] = useState<any[]>([]);
  const [ceoPopupOpen, setCeoPopupOpen] = useState(false);
  const [unackedFines, setUnackedFines] = useState<any[]>([]);
  const [fineAckOpen, setFineAckOpen] = useState(false);
  const [fineAckBusy, setFineAckBusy] = useState(false);
  const [marginMonth, setMarginMonth] = useState("");
  const [marginMonths, setMarginMonths] = useState<string[]>([]);
  const [marginData, setMarginData] = useState<any[]>([]);
  const [marginLoading, setMarginLoading] = useState(false);
  const [ceoPushSending, setCeoPushSending] = useState(false);
  const [ceoCustomOpen, setCeoCustomOpen] = useState(false);
  const [ceoCustomFrom, setCeoCustomFrom] = useState("");
  const [ceoCustomTo, setCeoCustomTo] = useState("");
  const [ceoCustomOutlets, setCeoCustomOutlets] = useState<string[]>([]);
  const [ceoCustomBusy, setCeoCustomBusy] = useState(false);
  const [ipUploads, setIpUploads] = useState<any[]>([]);
  const [ipSel, setIpSel] = useState<string>("");
  const [ipRows, setIpRows] = useState<any[]>([]);
  const [ipParsed, setIpParsed] = useState<any[] | null>(null);
  const [ipLabel, setIpLabel] = useState("");
  const [ipDays, setIpDays] = useState("30");
  const [ipBusy, setIpBusy] = useState(false);
  const [ipView, setIpView] = useState<"insights" | "data">("insights");
  const [ipPrevRows, setIpPrevRows] = useState<any[]>([]);
  const [compPeriod, setCompPeriod] = useState("");
  const fetchCompRows = async () => {
    const { data } = await supabase.from("competitor_sales").select("*").order("period_date", { ascending: false }).order("created_at", { ascending: false });
    setCompRows(data || []);
  };
  const saveComp = async () => {
    if (!compForm.competitor.trim()) { alert("Competitor name is required"); return; }
    setCompSaving(true);
    const { error } = await supabase.from("competitor_sales").insert({ competitor: compForm.competitor.trim(), area: compForm.area.trim() || null, address: compForm.address.trim() || null, sales_value: compForm.sales_value ? Number(compForm.sales_value) : null, our_sales: compForm.our_sales ? Number(compForm.our_sales) : null, our_outlet_id: compForm.our_outlet_id || null, period_label: compForm.period_label.trim() || null, period_date: compForm.period_date, note: compForm.note.trim() || null, entered_by: user?.id || null });
    setCompSaving(false);
    if (error) { alert("Save failed: " + error.message); return; }
   setCompForm({ competitor: "", area: "", address: "", sales_value: "", our_sales: "", our_outlet_id: "", period_label: "", period_date: new Date().toISOString().slice(0, 10), note: "" });
    fetchCompRows();
  };
  const parseComp = () => {
    const res = parseAreaTable(compPaste);
    if (!res || res.rows.length === 0) { alert("Couldn't read a table. Need an Area column, a Brownie Heaven column, and one or more competitor columns."); return; }
    setCompParsed(res.rows);
  };
  const saveParsedComp = async () => {
    if (!compParsed || !compParsed.length) return;
    setCompSavingBulk(true);
    const payload = compParsed.map(r => ({ competitor: (r.competitor || compPasteComp).trim() || "Unknown", area: r.area, sales_value: r.their, our_sales: r.our, period_label: compPasteLabel.trim() || null, period_date: compPasteDate, entered_by: user?.id || null }));
    const { error } = await supabase.from("competitor_sales").insert(payload);
    setCompSavingBulk(false);
    if (error) { alert("Save failed: " + error.message); return; }
    setCompPaste(""); setCompParsed(null);
    fetchCompRows();
  };
  const fetchCompHeadline = async () => { const { data } = await supabase.from("app_settings").select("value").eq("key", "comp_headline_on").maybeSingle(); setCompHeadlineOn(data ? data.value !== "false" : true); };
  const fetchCeoData = async (win: string) => {
    const winDays = win === "1" ? 1 : win === "30" ? 30 : 7;
    const end = new Date(Date.now() - 86400000);
    const start = new Date(end.getTime() - (winDays - 1) * 86400000);
    const sISO = start.toISOString().split("T")[0], eISO = end.toISOString().split("T")[0];
    const { data: reps } = await supabase.from("reports").select("staff_id, is_late, submitted_at").gte("submitted_at", sISO).lte("submitted_at", eISO + "T23:59:59");
    setCeoRepRows(reps || []);
    const ym = new Date().toISOString().slice(0, 7);
    const { data: mrep } = await supabase.from("outlet_reports").select("outlet_id, shop_sales_value, swiggy_sales_value, zomato_sales_value, report_date").gte("report_date", `${ym}-01`).lte("report_date", eISO);
    setCeoMonthRep(mrep || []);
  };
  const fetchMargin = async (month: string) => {
    setMarginLoading(true);
    try {
      const res = await fetch(`/api/margin?month=${month}`);
      const json = await res.json();
      setMarginData(json.success ? json.margins : []);
    } finally { setMarginLoading(false); }
  };
  const fetchMarginMonths = async () => {
    const { data } = await supabase.from("atlas_monthly_results").select("month");
    const uniq = Array.from(new Set((data || []).map((r: any) => r.month))).sort().reverse() as string[];
    setMarginMonths(uniq);
    if (uniq.length) { setMarginMonth(uniq[0]); fetchMargin(uniq[0]); }
  };

  const downloadCeoPDF = async () => {
    const ceo = computeCeoData(ceoRepRows, ceoMonthRep, ceoWin);
    const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", amber: "#C8901E" };
    const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
    const hh = new Date().getHours();
    const greet = hh < 12 ? "Good morning" : hh < 17 ? "Good afternoon" : "Good evening";
    const acctRows = ceo.acct.map((a) => `<div style="margin:3px 0;font-size:12px"><b>${a.name}</b> — ${a.tag}</div>`).join("") + `<div style="margin:3px 0;font-size:12px;color:${C.soft}"><b>Niranjana</b> — Founder's Office, MIA from daily reports (but she built this thing)</div>`;
    const ideaRows = ceo.ideas.map((i) => `<li style="margin:5px 0;font-size:12px">${i}</li>`).join("");
    const marginRows = marginData.map((m: any) => `<tr><td style="padding:4px 6px;font-size:11px">${m.outletName}</td><td style="padding:4px 6px;font-size:11px;text-align:right">${inr(m.salesNet)}</td><td style="padding:4px 6px;font-size:11px;text-align:right;color:${C.soft}">${inr(m.cogs)}</td><td style="padding:4px 6px;font-size:11px;text-align:right;font-weight:700;color:${m.marginPercent == null ? C.soft : m.marginPercent < 40 ? C.red : m.marginPercent < 60 ? C.amber : C.green}">${m.marginPercent == null ? "—" : m.marginPercent.toFixed(1) + "%"}</td></tr>`).join("");
    const marginCard = marginData.length > 0 ? `<div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:14px;font-weight:800;margin-bottom:8px">💰 Gross margin by outlet — ${marginMonth}</div><table style="width:100%;border-collapse:collapse"><thead><tr style="text-align:left"><th style="font-size:10px;color:${C.soft};padding:4px 6px">OUTLET</th><th style="font-size:10px;color:${C.soft};padding:4px 6px;text-align:right">SALES</th><th style="font-size:10px;color:${C.soft};padding:4px 6px;text-align:right">COST</th><th style="font-size:10px;color:${C.soft};padding:4px 6px;text-align:right">MARGIN %</th></tr></thead><tbody>${marginRows}</tbody></table></div>` : "";
    const html = `<div style="width:794px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px"><div style="font-size:21px;font-weight:900">${greet} — Brownie Heaven CEO Ops Brief</div><div style="font-size:11px;color:${C.soft};margin-bottom:16px">${ceo.ym} · day ${ceo.daysElapsed} of ${ceo.daysInMonth} · window: ${ceoWin === "1" ? "yesterday" : "last " + ceo.winDays + " days"}</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:14px;font-weight:800;margin-bottom:8px">Who's on top of it — and who's slipping</div>${acctRows}</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:14px;font-weight:800;margin-bottom:8px">Are we going to make the month?</div><div style="font-size:20px;font-weight:900">${inr(ceo.monthSales)}</div><div style="font-size:12px;color:${ceo.onTrack ? C.green : C.amber}">${ceo.salesPct.toFixed(0)}% of ${inr(ceo.monthTgt)} target · ${ceo.onTrack ? "on pace" : "behind pace"} (time elapsed ${ceo.timePct.toFixed(0)}%)</div>${ceo.drag ? `<div style="font-size:12px;color:${C.red};margin-top:6px">Dragging: ${ceo.drag.name} — ${ceo.drag.pct.toFixed(0)}% of target</div>` : ""}${ceo.hero ? `<div style="font-size:12px;color:${C.green}">Carrying: ${ceo.hero.name} — ${ceo.hero.pct.toFixed(0)}% of target</div>` : ""}</div>${marginCard}<div style="background:${C.card};border:2px solid ${C.amber};border-radius:12px;padding:16px"><div style="font-size:14px;font-weight:800;margin-bottom:8px;color:${C.amber}">Ideas to act on</div><ul style="margin:0;padding-left:18px">${ideaRows}</ul></div></div>`;
    const lib = await loadH2P();
    window.scrollTo(0, 0); await new Promise((r) => setTimeout(r, 50));
    const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0"; holder.innerHTML = html; document.body.appendChild(holder);
    try { await lib().set({ margin: 0, filename: `CEO_Brief_${new Date().toISOString().split("T")[0]}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
    finally { document.body.removeChild(holder); }
  };
  const openCeoCustom = () => {
    const today = new Date().toISOString().slice(0, 10);
    const weekAgo = new Date(Date.now() - 6 * 86400000).toISOString().slice(0, 10);
    setCeoCustomFrom(weekAgo); setCeoCustomTo(today); setCeoCustomOutlets([]);
    setCeoCustomOpen(true);
  };
  const downloadCeoPDFCustom = async () => {
    if (!ceoCustomFrom || !ceoCustomTo) { alert("Pick both a from and to date."); return; }
    setCeoCustomBusy(true);
    try {
      const { data: staffRows, error: staffErr } = await supabase.from("reports").select("staff_id, is_late, submitted_at").gte("submitted_at", ceoCustomFrom).lte("submitted_at", ceoCustomTo + "T23:59:59");
      if (staffErr) throw staffErr;
      const { data: outletRows, error: outletErr } = await supabase.from("outlet_reports").select("*").gte("report_date", ceoCustomFrom).lte("report_date", ceoCustomTo);
      if (outletErr) throw outletErr;
      const rows = outletRows || [];
      const cd = computeCeoCustom(staffRows || [], rows, ceoCustomFrom, ceoCustomTo, ceoCustomOutlets);
      const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", amber: "#C8901E" };
      const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
      const acctRows = cd.acct.map((a) => `<div style="margin:3px 0;font-size:12px"><b>${a.name}</b> — ${a.tag}</div>`).join("");
      const ideaRows = cd.ideas.map((i) => `<li style="margin:5px 0;font-size:12px">${i}</li>`).join("");
      const outletRowsHtml = cd.perOutlet.map((o) => `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:11px">${o.name}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px">${inr(o.sales)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:${C.soft}">${o.tgt > 0 ? inr(o.tgt) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;font-weight:700;color:${o.tgt > 0 ? (o.pct >= 100 ? C.green : o.pct >= 60 ? C.amber : C.red) : C.soft}">${o.tgt > 0 ? o.pct.toFixed(0) + "%" : "-"}</td></tr>`).join("");
      // The real differentiator: true margin for THIS EXACT period, computed live from Stock dispatch
      // quantities x actual recipe cost — not a monthly snapshot. Works for any date range, any outlets.
      let stockCard = "";
      try {
        const stockOutletIds = ceoCustomOutlets.length ? ceoCustomOutlets : OUTLETS;
        const stockOutletNames = stockOutletIds.map((o) => (OUTLET_ID_TO_STOCK_NAME as any)[o]).filter(Boolean);
        const { data: dispatchRows } = await supabaseStock.from("outlet_supply_ledger").select("outlet, product, qty, amount").gte("date", ceoCustomFrom).lte("date", ceoCustomTo).in("outlet", stockOutletNames);
        if (dispatchRows && dispatchRows.length) {
          const byOutlet: Record<string, { cost: number; dispatchValue: number; uncosted: Set<string>; products: Record<string, number> }> = {};
          const productTotals: Record<string, number> = {};
          dispatchRows.forEach((d: any) => {
            if (!byOutlet[d.outlet]) byOutlet[d.outlet] = { cost: 0, dispatchValue: 0, uncosted: new Set(), products: {} };
            const b = byOutlet[d.outlet];
            const unitCost = (FOOD_COST_MAP as any)[d.product];
            b.dispatchValue += Number(d.amount) || 0;
            if (unitCost !== undefined) {
              const c = unitCost * (Number(d.qty) || 0);
              b.cost += c;
              b.products[d.product] = (b.products[d.product] || 0) + c;
              productTotals[d.product] = (productTotals[d.product] || 0) + c;
            } else {
              b.uncosted.add(d.product);
            }
          });
          const marginRowsLive = cd.perOutlet.map((o: any) => {
            const stockName = (OUTLET_ID_TO_STOCK_NAME as any)[o.o];
            const b = stockName ? byOutlet[stockName] : undefined;
            if (!b) return null;
            const margin = o.sales - b.cost;
            const marginPct = o.sales > 0 ? (margin / o.sales) * 100 : null;
            const ratio = o.sales > 0 ? b.dispatchValue / o.sales : 0;
            return { name: o.name, sales: o.sales, cost: b.cost, margin, marginPct, flag: ratio > 1.5, uncostedCount: b.uncosted.size };
          }).filter(Boolean) as any[];
          const totalSales = marginRowsLive.reduce((a, r) => a + r.sales, 0);
          const totalCost = marginRowsLive.reduce((a, r) => a + r.cost, 0);
          const totalMarginPct = totalSales > 0 ? ((totalSales - totalCost) / totalSales) * 100 : 0;
          const mRows = marginRowsLive.map((r: any) => `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:11px">${r.flag ? "🚩 " : ""}${r.name}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px">${inr(r.sales)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;color:${C.soft}">${inr(r.cost)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:11px;font-weight:700;color:${r.marginPct == null ? C.soft : r.marginPct < 40 ? C.red : r.marginPct < 60 ? C.amber : C.green}">${r.marginPct == null ? "—" : r.marginPct.toFixed(1) + "%"}</td></tr>`).join("");
          const topProductsHtml = Object.entries(productTotals).sort((a: any, b: any) => b[1] - a[1]).slice(0, 5).map(([pname, val]: any) => `<tr><td style="padding:4px 8px;border-bottom:1px solid ${C.line};font-size:10px">${pname}</td><td style="padding:4px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${inr(val as number)}</td></tr>`).join("");
          const uncostedTotal = marginRowsLive.reduce((a, r) => a + r.uncostedCount, 0);
          stockCard = `<div style="background:${C.card};border:2px solid ${C.amber};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:15px;font-weight:900;margin-bottom:2px">📦 True margin for this exact period — live from Stock</div><div style="font-size:9px;color:${C.soft};margin-bottom:10px">Sales minus actual ingredient cost (dispatch quantity × real recipe cost), computed fresh for ${ceoCustomFrom} → ${ceoCustomTo}. This is not available anywhere else in TASKFORCE IQ.</div><div style="font-size:22px;font-weight:900;color:${totalMarginPct < 40 ? C.red : totalMarginPct < 60 ? C.amber : C.green}">${totalMarginPct.toFixed(1)}% overall margin</div><div style="font-size:11px;color:${C.soft};margin-bottom:10px">${inr(totalSales)} sales − ${inr(totalCost)} true food cost</div><table style="width:100%;border-collapse:collapse;margin-bottom:10px"><thead><tr style="text-align:left"><th style="font-size:9px;color:${C.soft};padding:4px 8px">OUTLET</th><th style="font-size:9px;color:${C.soft};padding:4px 8px;text-align:right">SALES</th><th style="font-size:9px;color:${C.soft};padding:4px 8px;text-align:right">FOOD COST</th><th style="font-size:9px;color:${C.soft};padding:4px 8px;text-align:right">MARGIN %</th></tr></thead><tbody>${mRows}</tbody></table>${topProductsHtml ? `<div style="font-size:11px;font-weight:700;margin-bottom:4px">Top 5 costliest products this period</div><table style="width:100%;border-collapse:collapse">${topProductsHtml}</table>` : ""}${uncostedTotal > 0 ? `<div style="font-size:9px;color:${C.soft};margin-top:8px;font-style:italic">${uncostedTotal} outlet-product combos have no recipe cost yet, so true cost is a touch higher than shown.</div>` : ""}${marginRowsLive.some((r: any) => r.flag) ? `<div style="font-size:9px;color:${C.red};margin-top:4px">🚩 = dispatched stock value is running well above sales for that outlet — worth a look.</div>` : ""}</div>`;
        }
      } catch { /* stock connection optional — skip on failure */ }
      const CEO_QUOTES = [
        "Revenue is vanity, margin is sanity, cash is reality.",
        "What gets measured gets managed. — Peter Drucker",
        "The best time to fix a leak was yesterday. The next best time is today.",
        "Small daily improvements are the key to staggering long-term results.",
        "A bakery runs on butter, sugar, and someone actually reading the numbers. 🍫",
        "Amateurs talk revenue. Professionals talk margin.",
      ];
      const quote = CEO_QUOTES[Math.floor(Math.random() * CEO_QUOTES.length)];
      const quoteCard = `<div style="text-align:center;font-size:11px;font-style:italic;color:${C.soft};margin:14px 0;padding:10px;border-top:1px dashed ${C.line};border-bottom:1px dashed ${C.line}">"${quote}"</div>`;
      const html = `<div style="width:794px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px"><div style="font-size:21px;font-weight:900">Brownie Heaven — Custom CEO Report</div><div style="font-size:11px;color:${C.soft};margin-bottom:16px">${ceoCustomFrom} → ${ceoCustomTo} · ${ceoCustomOutlets.length ? ceoCustomOutlets.length + " outlet(s)" : "all outlets"}</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:14px;font-weight:800;margin-bottom:8px">Who's on top of it — and who's slipping</div>${acctRows}</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;margin-bottom:14px"><div style="font-size:14px;font-weight:800;margin-bottom:8px">Sales for this period</div><div style="font-size:20px;font-weight:900">${inr(cd.periodSales)}</div><div style="font-size:12px;color:${cd.salesPct >= 90 ? C.green : C.amber}">${cd.periodTgt > 0 ? cd.salesPct.toFixed(0) + "% of " + inr(cd.periodTgt) + " target" : "No target data for this selection"}</div><table style="width:100%;border-collapse:collapse;margin-top:10px"><thead><tr style="text-align:left"><th style="font-size:10px;color:${C.soft};padding:4px 8px">OUTLET</th><th style="font-size:10px;color:${C.soft};padding:4px 8px;text-align:right">SALES</th><th style="font-size:10px;color:${C.soft};padding:4px 8px;text-align:right">TARGET</th><th style="font-size:10px;color:${C.soft};padding:4px 8px;text-align:right">%</th></tr></thead><tbody>${outletRowsHtml}</tbody></table></div>${stockCard}${quoteCard}<div style="background:${C.card};border:2px solid ${C.amber};border-radius:12px;padding:16px"><div style="font-size:14px;font-weight:800;margin-bottom:8px;color:${C.amber}">Ideas to act on</div><ul style="margin:0;padding-left:18px">${ideaRows}</ul></div></div>`;
      const lib = await loadH2P();
      window.scrollTo(0, 0); await new Promise((r) => setTimeout(r, 50));
      const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0"; holder.innerHTML = html; document.body.appendChild(holder);
      try { await lib().set({ margin: 0, filename: `CEO_Custom_${ceoCustomFrom}_to_${ceoCustomTo}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
      finally { document.body.removeChild(holder); }
      setCeoCustomOpen(false);
    } catch (e: any) { alert("Failed to generate: " + (e?.message || "error")); }
    setCeoCustomBusy(false);
  };
  useEffect(() => {
    if (activeTab === "ceo_report" && marginMonths.length === 0) { fetchMarginMonths(); }
    /* eslint-disable-next-line react-hooks/exhaustive-deps */
  }, [activeTab]);
  useEffect(() => {
    if (user && (user.role === "Owner" || user.role === "Founder's Office")) {
      fetchCeoData(ceoWin);
      const today = new Date().toISOString().split("T")[0];
      let showPopup = localStorage.getItem("ceo_popup_seen") !== today;
      (async () => {
        const { data: pushRow } = await supabase.from("app_settings").select("value").eq("key", "ceo_push_ts").maybeSingle();
        const pushTs = pushRow?.value;
        if (pushTs && pushTs !== localStorage.getItem("ceo_push_seen_ts")) {
          showPopup = true;
          localStorage.setItem("ceo_push_seen_ts", pushTs);
        }
        if (showPopup) { setCeoPopupOpen(true); localStorage.setItem("ceo_popup_seen", today); }
      })();
    }
    /* eslint-disable-next-line react-hooks/exhaustive-deps */
  }, [user]);
  useEffect(() => {
    if (!user) return;
    (async () => {
      const { data } = await supabase.from("fines").select("*").eq("staff_id", user.id).eq("acknowledged", false).order("created_at", { ascending: true });
      if (data && data.length > 0) { setUnackedFines(data); setFineAckOpen(true); }
    })();
    /* eslint-disable-next-line react-hooks/exhaustive-deps */
  }, [user]);
  const acknowledgeFines = async () => {
    setFineAckBusy(true);
    const ids = unackedFines.map((f) => f.id);
    await supabase.from("fines").update({ acknowledged: true }).in("id", ids);
    setFineAckBusy(false);
    setFineAckOpen(false);
    setUnackedFines([]);
  };
  const toggleCompHeadline = async () => { const nv = !compHeadlineOn; setCompHeadlineOn(nv); await supabase.from("app_settings").upsert({ key: "comp_headline_on", value: nv ? "true" : "false", updated_at: new Date().toISOString() }, { onConflict: "key" }); };
  const pushCeoPopupToNishant = async () => {
    setCeoPushSending(true);
    try {
      await supabase.from("app_settings").upsert({ key: "ceo_push_ts", value: new Date().toISOString(), updated_at: new Date().toISOString() }, { onConflict: "key" });
      alert("Pushed — Nishant will see the popup next time he opens TASKFORCE IQ (or right away if he already has it open).");
    } finally { setCeoPushSending(false); }
  };
  const fetchCompProducts = async () => {
    const { data } = await supabase.from("competitor_products").select("*").order("gmv", { ascending: false, nullsFirst: false });
    setCompProducts(data || []);
  };
  const parseProd = () => {
    const rows = parseProductTable(prodPaste);
    if (!rows || rows.length === 0) { alert("Couldn't read a product table. Need a header with Product and GMV columns."); return; }
    setProdParsed(rows);
  };
  const saveProd = async () => {
    if (!prodParsed || !prodParsed.length) return;
    if (prodSide === "them" && !prodComp.trim()) { alert("Enter the competitor name for their products."); return; }
    setProdSaving(true);
    const payload = prodParsed.map((r, i) => ({ side: prodSide, competitor: prodSide === "them" ? prodComp.trim() : null, product: r.product, gmv: r.gmv, orders: r.orders, units: r.units, areas: r.areas, rank: i + 1, period_label: prodLabel.trim() || null, period_date: prodDate, entered_by: user?.id || null }));
    const { error } = await supabase.from("competitor_products").insert(payload);
    setProdSaving(false);
    if (error) { alert("Save failed: " + error.message); return; }
    setProdPaste(""); setProdParsed(null);
    fetchCompProducts();
  };
  const compFmtL = (n: number) => n >= 100000 ? "₹" + (n / 100000).toFixed(2) + "L" : "₹" + Math.round(n).toLocaleString("en-IN");
  const compPeriods = Array.from(new Set(compRows.map((r: any) => r.period_label || r.period_date))).filter(Boolean) as string[];
  const compActivePeriod = compPeriod || compPeriods[0] || "";
  const compPeriodRows = compRows.filter((r: any) => (r.period_label || r.period_date) === compActivePeriod).map((r: any) => { const their = Number(r.sales_value) || 0; const our = r.our_sales == null ? null : Number(r.our_sales); return { area: r.area as string, competitor: (r.competitor || "Unknown") as string, their, our: our as number | null, gap: their - (our || 0) }; });
  const compExpansion = compPeriodRows.filter(r => r.their > (r.our || 0) && (r.our == null || r.our === 0)).sort((a, b) => b.gap - a.gap);
  const compDefend = compPeriodRows.filter(r => r.our != null && r.our > 0 && r.their > r.our).sort((a, b) => b.gap - a.gap);
  const compWinning = compPeriodRows.filter(r => (r.our || 0) >= r.their).sort((a, b) => a.gap - b.gap);
  const compTop = compExpansion[0] || compDefend[0] || compWinning[0] || null;
  const compFunny = (() => {
    if (!compTop) return { playful: "", pro: "" };
    const g = compFmtL(Math.abs(compTop.gap)); const area = compTop.area; const competitor = compTop.competitor;
    if (compExpansion[0] === compTop) return { playful: `${competitor} is printing money in ${area} — ${g} ahead and we haven't even shown up yet 😬🍫`, pro: `${competitor} leads ${area} by ${g} in an area where we have no real presence — worth a look.` };
    if (compDefend[0] === compTop) return { playful: `${competitor} is out-baking us in ${area} by ${g}. Time to fire up the ovens 🔥`, pro: `We trail ${competitor} in ${area} by ${g} — a focus area to defend.` };
    return { playful: `${area} is all ours — ${g} clear of ${competitor}. Chef's kiss 🧑‍🍳`, pro: `We lead ${competitor} in ${area} by ${g}.` };
  })();
  const downloadCompPDF = async () => {
    if (compPeriodRows.length === 0) { alert("No competitor data for this period."); return; }
    const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", orange: "#C2410C" };
    const rs = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
    type Row = { area: string; competitor: string; their: number; our: number | null; gap: number };
    const section = (title: string, subtitle: string, rows: Row[], color: string) => {
      if (!rows.length) return "";
      const body = rows.map(r => `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;color:${C.ink}">${r.area}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;color:${C.soft}">${r.competitor}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;text-align:right">${rs(r.their)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;text-align:right;color:${C.soft}">${r.our != null ? rs(r.our) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;text-align:right;font-weight:700;color:${color}">${rs(Math.abs(r.gap))}</td></tr>`).join("");
      return `<div style="margin-top:20px"><div style="font-size:14px;font-weight:800;color:${color}">${title}</div><div style="font-size:10px;color:${C.soft};margin-bottom:6px">${subtitle}</div><table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden"><thead><tr style="background:${C.ink}"><th style="padding:7px 8px;text-align:left;color:#FFF6E5;font-size:9px">AREA</th><th style="padding:7px 8px;text-align:left;color:#FFF6E5;font-size:9px">COMPETITOR</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">THEM</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">US</th><th style="padding:7px 8px;text-align:right;color:#FFF6E5;font-size:9px">GAP</th></tr></thead><tbody>${body}</tbody></table></div>`;
    };
    const html = `<div style="width:794px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px"><div style="font-size:22px;font-weight:900">Brownie Heaven — Competition Report</div><div style="font-size:11px;color:${C.soft};margin-bottom:16px">${compActivePeriod}</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:18px;font-size:17px;font-weight:700;line-height:1.5">${compFunny.playful}</div>${section("Expansion candidates", "They lead where we're absent — worth a look", compExpansion, C.orange)}${section("Defend", "We're present but trailing", compDefend, C.red)}${section("Winning", "We're ahead", compWinning, C.green)}<div style="font-size:9px;color:${C.soft};margin-top:18px">Competitor figures are estimates shared by the owner — a directional signal, not a verdict.</div></div>`;
    const lib = await loadH2P();
    const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.innerHTML = html; document.body.appendChild(holder);
    try { await lib().set({ margin: 0, filename: `Competition_${(compActivePeriod || "report").replace(/[^a-z0-9]+/gi, "_")}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
    finally { document.body.removeChild(holder); }
  };
  useEffect(() => { if (activeTab === "competition") { fetchCompRows(); fetchCompProducts(); } /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);
  useEffect(() => { if (user) { fetchCompRows(); fetchCompProducts(); fetchCompHeadline(); } /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [user]);
  const fetchIpRows = async (uploadId: string) => { const { data } = await supabase.from("item_perf_rows").select("*").eq("upload_id", uploadId).order("net_revenue", { ascending: false, nullsFirst: false }); setIpRows(data || []); };
  const fetchIpUploads = async () => { const { data } = await supabase.from("item_perf_uploads").select("*").order("created_at", { ascending: false }); setIpUploads(data || []); if (data && data.length) { setIpSel((cur) => cur || data[0].id); if (!ipSel) fetchIpRows(data[0].id); } };
  const [ipDeleteBusy, setIpDeleteBusy] = useState(false);
  const deleteIpUpload = async () => {
    if (!ipSel) return;
    const up = ipUploads.find((u) => u.id === ipSel);
    if (!confirm(`Delete "${up?.label || "this upload"}" and all its ${up?.row_count || ""} item rows? This can't be undone.`)) return;
    setIpDeleteBusy(true);
    await supabase.from("item_perf_rows").delete().eq("upload_id", ipSel);
    await supabase.from("item_perf_uploads").delete().eq("id", ipSel);
    setIpSel(""); setIpRows([]);
    await fetchIpUploads();
    setIpDeleteBusy(false);
  };
  useEffect(() => { if (activeTab === "item_perf") { fetchIpUploads(); fetchWastage(); } /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);
  // Whichever upload sits immediately before the selected one (by upload date) becomes
  // the comparison period for the pricing/discount-discipline section — no extra picker.
  useEffect(() => {
    if (!ipSel || ipUploads.length < 2) { setIpPrevRows([]); return; }
    const idx = ipUploads.findIndex((u) => u.id === ipSel);
    const prevUpload = idx >= 0 ? ipUploads[idx + 1] : null;
    if (!prevUpload) { setIpPrevRows([]); return; }
    (async () => {
      const { data } = await supabase.from("item_perf_rows").select("name, net_revenue, units_sold, avg_price").eq("upload_id", prevUpload.id);
      setIpPrevRows(data || []);
    })();
    /* eslint-disable-next-line react-hooks/exhaustive-deps */
  }, [ipSel, ipUploads]);
  const parseItemFile = async (file: File) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json: any[] = XLSX.utils.sheet_to_json(ws, { defval: "" });
    const num = (v: any) => { const n = parseFloat(String(v).replace(/[^0-9.\-]/g, "")); return isNaN(n) ? null : n; };
    const pick = (r: any, keys: string[]) => { for (const k of Object.keys(r)) { const nk = k.replace(/^\ufeff/, "").trim().toLowerCase(); if (keys.some((x) => nk === x || nk.includes(x))) return r[k]; } return ""; };
    // "units sold"/"revenue"/etc were the generic names originally guessed for
    // this export. The real UrbanPiper "Items summary across all locations"
    // report uses "Total Quantity" and "Total Amount" instead \u2014 added here so
    // that exact file parses correctly (was previously reading these as blank,
    // showing 0 for every item, e.g. Chocolate Truffle Cake, Milo Madness).
    // "Total Amount" is treated as net revenue (pre-tax; Total Taxes and Total
    // Charges are separate columns in that export) \u2014 flag if that's not right.
    return json.map((r) => {
      const name = String(pick(r, ["name"])).trim();
      const units_sold = num(pick(r, ["units sold", "units", "total quantity", "quantity"]));
      const net_revenue = num(pick(r, ["net revenue", "revenue", "total amount", "amount"]));
      const avg_price_raw = num(pick(r, ["avg price", "average price"]));
      const avg_price = avg_price_raw != null ? avg_price_raw : (units_sold && net_revenue != null && units_sold > 0 ? Math.round((net_revenue / units_sold) * 100) / 100 : null);
      return { name, category: String(pick(r, ["category"])).trim(), net_revenue, units_sold, avg_price, lost_orders: num(pick(r, ["lost orders"])), avg_orders_day: num(pick(r, ["avg orders", "orders / day", "orders/day"])) };
    }).filter((r) => r.name);
  };
  const onIpFile = async (e: any) => {
    const f = e.target.files?.[0]; if (!f) return;
    try { const rows = await parseItemFile(f); if (!rows.length) { alert("No item rows found. Check the export format."); return; } setIpParsed(rows); } catch (err: any) { alert("Couldn't read the file: " + (err?.message || "")); }
    e.target.value = "";
  };
  const saveIp = async () => {
    if (!ipParsed || !ipParsed.length) return;
    setIpBusy(true);
    const { data: up, error: e1 } = await supabase.from("item_perf_uploads").insert({ label: ipLabel.trim() || `Last ${ipDays} days`, period_days: Number(ipDays) || null, uploaded_by: user?.id || null, row_count: ipParsed.length }).select().single();
    if (e1 || !up) { setIpBusy(false); alert("Save failed: " + (e1?.message || "")); return; }
    const payload = ipParsed.map((r) => ({ upload_id: up.id, ...r }));
    const { error: e2 } = await supabase.from("item_perf_rows").insert(payload);
    setIpBusy(false);
    if (e2) { alert("Rows save failed: " + e2.message); return; }
    setIpParsed(null); setIpLabel("");
    await fetchIpUploads(); setIpSel(up.id); fetchIpRows(up.id);
  };
  // Real price-vs-demand classification, extracted so both the Item Performance tab and
  // Outlet Health use the exact same analysis — not two versions that can drift apart.
   const classifyItemPerf = (rawRows: any[]) => {
    const rows = rawRows.map((r: any) => ({ name: r.name as string, category: r.category as string, rev: Number(r.net_revenue) || 0, units: Number(r.units_sold) || 0, price: Number(r.avg_price) || 0, lost: Number(r.lost_orders) || 0, opd: Number(r.avg_orders_day) || 0 }));
    if (!rows.length) return null;
    const med = (arr: number[]) => { const a = [...arr].sort((x, y) => x - y); return a.length ? a[Math.floor(a.length / 2)] : 0; };
    const mU = med(rows.map((r) => r.units)), mP = med(rows.map((r) => r.price)), mR = med(rows.map((r) => r.rev));
    const stars = [...rows].filter((r) => r.units >= mU).sort((a, b) => b.rev - a.rev).slice(0, 5);
    const moneyLeft = [...rows].filter((r) => r.lost > 0).sort((a, b) => b.lost - a.lost).slice(0, 5);
    const byCat: Record<string, typeof rows> = {};
    rows.forEach((r) => { (byCat[r.category] = byCat[r.category] || []).push(r); });
    let suspects: typeof rows = [];
    Object.values(byCat).forEach((items) => { if (items.length < 3) return; const cP = med(items.map((r) => r.price)), cU = med(items.map((r) => r.units)); items.forEach((r) => { if (r.units >= 10 && r.price > cP && r.units < cU) suspects.push(r); }); });
    suspects = suspects.sort((a, b) => (b.lost - a.lost) || (b.price - a.price)).slice(0, 5);
    const sweet = [...rows].filter((r) => r.price <= mP && r.units >= mU).sort((a, b) => b.units - a.units).slice(0, 5);
    const dead = [...rows].filter((r) => r.units >= 3 && r.units < mU && r.rev < mR).sort((a, b) => a.units - b.units).slice(0, 5);
    const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
    const top = stars[0], s0 = suspects[0], w0 = sweet[0], d0 = dead[0];
    const funny = {
      headline: top ? `Top performer: ${top.name} at ${inr(top.price)} — ${top.units.toLocaleString("en-IN")} units sold this window. Highest-volume item; hold the current price.` : "",
      suspect: s0 ? `Priced above category median with below-median demand: ${s0.name} at ${inr(s0.price)}${s0.lost ? `, ${s0.lost} lost order${s0.lost === 1 ? "" : "s"}` : ""}. Worth reviewing the price.` : "",
      sweet: w0 ? `Strong performer: ${w0.name} at ${inr(w0.price)} — at/below median price with ${w0.units.toLocaleString("en-IN")} units sold. Pricing is working as intended here.` : "",
      dead: d0 ? `Underperforming: ${d0.name} — only ${d0.units} units sold this window. Candidate for a price review, menu repositioning, or discontinuation.` : "",
    };

    // Revenue concentration (Pareto/ABC) — which items actually carry the menu.
    const totalRev = rows.reduce((s, r) => s + r.rev, 0);
    const byRevDesc = [...rows].sort((a, b) => b.rev - a.rev);
    let cum = 0;
    const tiered = byRevDesc.map((r) => {
      const share = totalRev > 0 ? (r.rev / totalRev) * 100 : 0;
      cum += share;
      const tier = cum <= 70 ? "A" : cum <= 90 ? "B" : "C";
      return { ...r, share, cumShare: cum, tier };
    });
    const abc = (["A", "B", "C"] as const).reduce((acc, t) => {
      const items = tiered.filter((r) => r.tier === t);
      acc[t] = { count: items.length, share: items.reduce((s, r) => s + r.share, 0) };
      return acc;
    }, {} as Record<"A" | "B" | "C", { count: number; share: number }>);

    // Real margin, matched from the actual costing sheet — never guessed for
    // items that don't confidently match (Assorted Boxes, combos, etc).
    const margin = rows
      .map((r) => {
        const m = matchItemCost(r.name);
        if (!m || r.price <= 0) return null;
        const marginRs = r.price - m.cost;
        const marginPct = (marginRs / r.price) * 100;
        return { name: r.name, price: r.price, cost: m.cost, matchedName: m.matchedName, method: m.method, marginRs, marginPct };
      })
      .filter((x): x is NonNullable<typeof x> => !!x)
      .sort((a, b) => a.marginPct - b.marginPct);
    const marginRevCovered = margin.reduce((s, m) => { const row = rows.find((r) => r.name === m.name); return s + (row ? row.rev : 0); }, 0);
    const marginCoverage = { matched: margin.length, total: rows.length, revSharePct: totalRev > 0 ? (marginRevCovered / totalRev) * 100 : 0 };

    // Classic menu-engineering quadrant: margin % (profitability) x units sold
    // (popularity) — a different lens from the price-vs-demand one above. Only
    // covers items that matched the costing sheet, since margin is real cost
    // data here, never guessed. Split at the median of THIS matched set (not
    // the whole menu) so the two axes are fairly compared against each other.
    const mvRows = margin
      .map((m) => { const row = rows.find((r) => r.name === m.name); return row ? { ...m, units: row.units, rev: row.rev } : null; })
      .filter((x): x is NonNullable<typeof x> => !!x);
    const mvMedianMargin = med(mvRows.map((r) => r.marginPct));
    const mvMedianUnits = med(mvRows.map((r) => r.units));
    const mvStars = mvRows.filter((r) => r.marginPct >= mvMedianMargin && r.units >= mvMedianUnits).sort((a, b) => b.rev - a.rev);
    const mvWorkhorses = mvRows.filter((r) => r.marginPct < mvMedianMargin && r.units >= mvMedianUnits).sort((a, b) => b.units - a.units);
    const mvPuzzles = mvRows.filter((r) => r.marginPct >= mvMedianMargin && r.units < mvMedianUnits).sort((a, b) => b.marginPct - a.marginPct);
    const mvDogs = mvRows.filter((r) => r.marginPct < mvMedianMargin && r.units < mvMedianUnits).sort((a, b) => a.marginPct - b.marginPct);
    const marginQuadrant = { rows: mvRows, medianMarginPct: mvMedianMargin, medianUnits: mvMedianUnits, stars: mvStars, workhorses: mvWorkhorses, puzzles: mvPuzzles, dogs: mvDogs };

    return { rows, stars, moneyLeft, suspects, sweet, dead, funny, inr, tiered, abc, margin, marginCoverage, marginQuadrant };
  };

  // Price-vs-volume read between two uploads: if price rose and volume held or grew,
  // that's demand absorbing the increase — no discount needed. If price fell while
  // volume jumped, that's the shape of a discount or platform promo — worth checking
  // the real margin before repeating it. Matched by name since UrbanPiper's own
  // item IDs aren't in this export.
  const compareItemPeriods = (curRows: any[], prevRows: any[]) => {
    if (!curRows?.length || !prevRows?.length) return null;
    const prevByName: Record<string, any> = {};
    prevRows.forEach((r: any) => { prevByName[r.name] = r; });
    const curByName: Record<string, boolean> = {};
    const out: any[] = [];
    curRows.forEach((r: any) => {
      curByName[r.name] = true;
      const p = prevByName[r.name];
      const curPrice = Number(r.avg_price) || 0, curUnits = Number(r.units_sold) || 0;
      if (!p) { out.push({ name: r.name, curPrice, curUnits, bucket: "new" }); return; }
      const prevPrice = Number(p.avg_price) || 0, prevUnits = Number(p.units_sold) || 0;
      if (prevPrice <= 0 || prevUnits <= 0) return;
      const priceChg = ((curPrice - prevPrice) / prevPrice) * 100;
      const unitsChg = ((curUnits - prevUnits) / prevUnits) * 100;
      let bucket = "mixed";
      if (priceChg > 2 && unitsChg >= -2) bucket = "priceLedGrowth";
      else if (priceChg < -2 && unitsChg > 5) bucket = "possibleDiscount";
      else if (Math.abs(priceChg) <= 2 && unitsChg < -10) bucket = "decliningStable";
      else if (priceChg > 2 && unitsChg < -10) bucket = "priceResistance";
      else if (Math.abs(priceChg) <= 2 && Math.abs(unitsChg) <= 5) bucket = "stable";
      out.push({ name: r.name, curPrice, curUnits, prevPrice, prevUnits, priceChg, unitsChg, bucket });
    });
    const dropped = prevRows.filter((p: any) => !curByName[p.name]).map((p: any) => ({ name: p.name, prevUnits: Number(p.units_sold) || 0 }));
    return { rows: out, dropped };
  };
  const ipStats = classifyItemPerf(ipRows);
  const ipComparison = compareItemPeriods(ipRows, ipPrevRows);
  const downloadIpPDF = async () => {
    if (!ipStats) { alert("No data to report."); return; }
    window.scrollTo(0, 0);
    await new Promise((r) => setTimeout(r, 50));
    const up = ipUploads.find((u) => u.id === ipSel);
    const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", orange: "#C2410C", gold: "#C8901E" };
    const inr = ipStats.inr;
    const tbl = (title: string, sub: string, rows: any[], color: string) => rows.length ? `<div style="margin-top:18px"><div style="font-size:14px;font-weight:800;color:${color}">${title}</div><div style="font-size:10px;color:${C.soft};margin-bottom:6px">${sub}</div><table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden"><thead><tr style="background:${C.ink}"><th style="padding:6px 8px;text-align:left;color:#FFF6E5;font-size:9px">ITEM</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">AVG</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">UNITS</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">LOST</th></tr></thead><tbody>${rows.map((r) => `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:10px;color:${C.ink}">${r.name}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px">${inr(r.price)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px">${r.units.toLocaleString("en-IN")}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:10px;color:${C.soft}">${r.lost}</td></tr>`).join("")}</tbody></table></div>` : "";
    const html = `<div style="width:794px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px"><div style="font-size:22px;font-weight:900">Brownie Heaven — Item Performance</div><div style="font-size:11px;color:${C.soft};margin-bottom:14px">${up?.label || ""} · ${ipStats.rows.length} items</div><div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px;font-size:16px;font-weight:700;line-height:1.5">${ipStats.funny.headline}</div>${ipStats.funny.suspect ? `<div style="margin-top:8px;font-size:13px;color:${C.orange}">${ipStats.funny.suspect}</div>` : ""}${ipStats.funny.sweet ? `<div style="margin-top:4px;font-size:13px;color:${C.green}">${ipStats.funny.sweet}</div>` : ""}${ipStats.funny.dead ? `<div style="margin-top:4px;font-size:13px;color:${C.red}">${ipStats.funny.dead}</div>` : ""}${tbl("Stars", "The workhorses — protect these", ipStats.stars, C.green)}${tbl("Priced-too-high suspects", "High price, shy demand — worth a rethink", ipStats.suspects, C.orange)}${tbl("Sweet-spot winners", "Great price, flying off shelves", ipStats.sweet, C.green)}${tbl("Money left on the table", "People wanted it, didn't get it", ipStats.moneyLeft, C.gold)}${tbl("Dead weight", "Barely moving — rework or retire?", ipStats.dead, C.red)}<div style="font-size:9px;color:${C.soft};margin-top:16px">Directional signals from one export, not final verdicts. Confirm price effects across two windows before changing prices.</div></div>`;
    const lib = await loadH2P();
    const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0"; holder.innerHTML = html; document.body.appendChild(holder);
    try { await lib().set({ margin: 0, filename: `ItemPerformance_${(up?.label || "report").replace(/[^a-z0-9]+/gi, "_")}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
    finally { document.body.removeChild(holder); }
  };

    useEffect(() => {
    if (activeTab !== "analytics") return;
    let cancelled = false;
    (async () => {
      setAnLoading(true);
      const { data } = await supabase.from("outlet_reports")
        .select("outlet_id,shop_sales_count,shop_sales_value,swiggy_sales_count,swiggy_sales_value,zomato_sales_count,zomato_sales_value,discount_given,report_date")
        .gte("report_date", anFrom).lte("report_date", anTo);
      if (cancelled) return;
      setAnRows(data || []);
      setAnLoading(false);
    })();
    return () => { cancelled = true; };
  }, [activeTab, anFrom, anTo]);

  const fetchOutletHealth = async () => {
    setOutletHealthLoading(true);
    // Same silent-truncation risk as Purchase & Vendors: this query has no end
    // date, so it grows every day — already past Supabase's 1000-row cap as of
    // September 2026. Batched fetch removes the ceiling entirely.
    const BATCH = 1000;
    let data: any[] = [];
    let offset = 0;
    while (true) {
      const { data: batch, error } = await supabase.from("outlet_reports").select("outlet_id, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value").gte("report_date", "2026-06-01").range(offset, offset + BATCH - 1);
      if (error) { console.error(error); break; }
      data = data.concat(batch || []);
      if (!batch || batch.length < BATCH) break;
      offset += BATCH;
    }
    const byOutletMonth: Record<string, Record<string, number>> = {};
    (data || []).forEach((r: any) => {
      const oid = r.outlet_id;
      const ym = r.report_date.slice(0, 7);
      const total = (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
      if (!byOutletMonth[oid]) byOutletMonth[oid] = {};
      byOutletMonth[oid][ym] = (byOutletMonth[oid][ym] || 0) + total;
    });
    const nowYm = new Date().toISOString().slice(0, 7);
    const result = OUTLETS.map((oid) => {
      const months = byOutletMonth[oid] || {};
      const sortedMonths = Object.keys(months).sort();
      const thisMonthTotal = months[nowYm] || 0;
      const tgt = monthlyTargetFor(oid, nowYm);
      const pct = tgt > 0 ? (thisMonthTotal / tgt) * 100 : 0;
      const prevMonths = sortedMonths.filter((m) => m !== nowYm);
      const lastMonth = prevMonths[prevMonths.length - 1];
      const lastMonthTotal = lastMonth ? months[lastMonth] : null;
      let trendPct: number | null = null;
      if (lastMonthTotal && lastMonthTotal > 0) trendPct = ((thisMonthTotal - lastMonthTotal) / lastMonthTotal) * 100;
      let health = "No data";
      if (tgt > 0 && thisMonthTotal > 0) health = pct >= 90 ? "Strong" : pct >= 60 ? "On track" : pct >= 30 ? "Needs attention" : "Struggling";
      let trendLabel = "steady";
      if (trendPct != null) trendLabel = trendPct > 10 ? "growing" : trendPct < -10 ? "declining" : "steady";
      return { oid, name: OUTLET_NAMES[oid] || oid, thisMonthTotal, pct, health, trendPct, trendLabel, monthly: sortedMonths.map((m) => ({ month: m, total: months[m] })) };
    });
    setOutletHealthData(result);
    setOutletHealthLoading(false);
  };
   useEffect(() => { if ((activeTab === "analytics" || activeTab === "owner_outlets") && outletHealthData.length === 0) fetchOutletHealth(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  // Builds the full daily series + monthly aggregates + rule-based insights for one
  // outlet. Shared by the in-app deep-dive view and the PDF export so both stay in sync.
  // Fetches real item-level sales data (from actual Prime/billing exports, via the
  // Item Performance feature) — the latest upload and the one before it, so we can flag
  // real top sellers, real availability/lost-order problems, and real trend shifts.
  // Not tied to a specific outlet (company-wide), but genuinely current, not researched.
  const fetchItemPerfContext = async () => {
    const { data: uploads } = await supabase.from("item_perf_uploads").select("*").order("created_at", { ascending: false }).limit(2);
    if (!uploads || uploads.length === 0) { setItemPerfContext({ available: false }); return; }
    const { data: latestRows } = await supabase.from("item_perf_rows").select("name, category, net_revenue, units_sold, avg_price, lost_orders, avg_orders_day").eq("upload_id", uploads[0].id);
    const rows = latestRows || [];
    const classified = classifyItemPerf(rows);

    let gainers: any[] = [], decliners: any[] = [];
    if (uploads.length > 1) {
      const { data: prevRows } = await supabase.from("item_perf_rows").select("name, net_revenue").eq("upload_id", uploads[1].id);
      const prevByName: Record<string, number> = {};
      (prevRows || []).forEach((r: any) => { if (r.net_revenue != null) prevByName[r.name] = r.net_revenue; });
      const changes = rows.filter((r: any) => r.net_revenue != null && prevByName[r.name] != null && prevByName[r.name] > 0).map((r: any) => ({ name: r.name, cur: r.net_revenue, prev: prevByName[r.name], chg: ((r.net_revenue - prevByName[r.name]) / prevByName[r.name]) * 100 }));
      gainers = changes.filter((c) => c.chg > 15).sort((a, b) => b.chg - a.chg).slice(0, 2);
      decliners = changes.filter((c) => c.chg < -15).sort((a, b) => a.chg - b.chg).slice(0, 2);
    }
    setItemPerfContext({ available: !!classified, label: uploads[0].label, periodDays: uploads[0].period_days, classified, gainers, decliners, prevLabel: uploads[1]?.label || null });
  };
  useEffect(() => { if (activeTab === "owner_outlets" && itemPerfContext === null) fetchItemPerfContext(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab]);

  const computeOutletAnalysis = (oid: string, rawRows: any[], itemPerf: any) => {
    const daily = [...rawRows].sort((a, b) => (a.report_date as string).localeCompare(b.report_date));
    const dailyRows = daily.map((r: any) => {
      const ym = r.report_date.slice(0, 7);
      const shop = Number(r.shop_sales_value) || 0, swiggy = Number(r.swiggy_sales_value) || 0, zomato = Number(r.zomato_sales_value) || 0;
      const total = shop + swiggy + zomato;
      const target = dailyTargetFor(oid, ym);
      const discount = Number(r.discount_given) || 0;
      return { date: r.report_date, shop, swiggy, zomato, total, target, discount, pct: target > 0 ? (total / target) * 100 : 0 };
    });
    const byMonth: Record<string, { shop: number; swiggy: number; zomato: number }> = {};
    dailyRows.forEach((r) => {
      const ym = r.date.slice(0, 7);
      if (!byMonth[ym]) byMonth[ym] = { shop: 0, swiggy: 0, zomato: 0 };
      byMonth[ym].shop += r.shop; byMonth[ym].swiggy += r.swiggy; byMonth[ym].zomato += r.zomato;
    });
    const months = Object.keys(byMonth).sort();
    const nowYm = new Date().toISOString().slice(0, 7);
    const thisM = byMonth[nowYm] || { shop: 0, swiggy: 0, zomato: 0 };
    const prevYm = months.filter((m) => m !== nowYm).pop();
    const lastM = prevYm ? byMonth[prevYm] : null;
    const chg = (cur: number, prev: number | undefined) => prev == null || prev === 0 ? null : ((cur - prev) / prev) * 100;
    const shopChg = chg(thisM.shop, lastM?.shop), swiggyChg = chg(thisM.swiggy, lastM?.swiggy), zomatoChg = chg(thisM.zomato, lastM?.zomato);
    const monthTotals = months.map((m) => ({ m, total: byMonth[m].shop + byMonth[m].swiggy + byMonth[m].zomato, target: monthlyTargetFor(oid, m) }));

    const thisMonthRows = dailyRows.filter((r) => r.date.slice(0, 7) === nowYm);
    const thisMonthDays = thisMonthRows.filter((r) => r.target > 0);
    const daysHit = thisMonthDays.filter((r) => r.total >= r.target).length;
    const hitRate = thisMonthDays.length > 0 ? (daysHit / thisMonthDays.length) * 100 : null;

    // Weekday pattern — which day of the week performs best/worst, this month
    const dowNames = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    const dowTotals = [0, 0, 0, 0, 0, 0, 0], dowCounts = [0, 0, 0, 0, 0, 0, 0];
    thisMonthRows.forEach((r) => { const dow = new Date(r.date + "T00:00:00").getDay(); dowTotals[dow] += r.total; dowCounts[dow]++; });
    const dowAvg = dowTotals.map((t, i) => dowCounts[i] > 0 ? t / dowCounts[i] : 0);
    const validDow = dowAvg.map((v, i) => ({ v, i })).filter((x) => dowCounts[x.i] > 0);
    const bestDow = validDow.length ? validDow.reduce((a, b) => (b.v > a.v ? b : a)) : null;
    const worstDow = validDow.length ? validDow.reduce((a, b) => (b.v < a.v ? b : a)) : null;

    // Channel balance — this month
    const tAll = thisM.shop + thisM.swiggy + thisM.zomato || 1;
    const onlineShare = ((thisM.swiggy + thisM.zomato) / tAll) * 100;
    const shopShare = (thisM.shop / tAll) * 100;

    // Discount drag — this month
    const discountTotal = thisMonthRows.reduce((a, r) => a + r.discount, 0);
    const discountPct = tAll > 0 ? (discountTotal / tAll) * 100 : 0;

    const insights: string[] = [];
    const channelName = (v: number) => v === shopChg ? "Shop" : v === swiggyChg ? "Swiggy" : "Zomato";
    const validChgs = [shopChg, swiggyChg, zomatoChg].filter((v) => v != null) as number[];
    if (validChgs.length) {
      const worst = Math.min(...validChgs), best = Math.max(...validChgs);
      if (worst < -5) {
        const nm = channelName(worst);
        const advice = nm === "Swiggy" || nm === "Zomato"
          ? `Check three things on ${nm}: menu photos and descriptions are complete and appetizing, negative reviews are kept under 2% (a couple of bad ratings can quietly sink ranking), and kitchen prep time is fast — ${nm}'s own ranking algorithm favours restaurants with lower prep times. Running a paid ad slot during your peak order hours can also recover lost visibility fast.`
          : "Check walk-in footfall at peak hours, whether a nearby competitor has opened or is running a stronger offer, and whether staffing dips are causing wait times to build up.";
        insights.push(`${nm} dropped ${Math.abs(worst).toFixed(1)}% vs last month. ${advice}`);
      }
      if (best > 5 && best !== worst) { const nm = channelName(best); insights.push(`${nm} grew ${best.toFixed(1)}% vs last month — find out what worked (a running offer, a new combo, better ratings) and repeat it on the other channels too.`); }
    }
    if (monthTotals.length >= 3) { const last3 = monthTotals.slice(-3); if (last3[2].total < last3[1].total && last3[1].total < last3[0].total) insights.push(`Sales have fallen for two months straight (${last3[0].m} → ${last3[2].m}) — this isn't a one-off dip, worth digging into what changed structurally (menu, pricing, a competitor, staffing).`); }
    const tgtNow = monthlyTargetFor(oid, nowYm);
    const pctNow = tgtNow > 0 ? (thisM.shop + thisM.swiggy + thisM.zomato) / tgtNow * 100 : 0;
    if (pctNow > 0 && pctNow < 60) insights.push(`Only ${pctNow.toFixed(0)}% of this month's target hit so far — at this pace the outlet will fall well short unless the remaining days pick up sharply.`);
    if (pctNow >= 90) insights.push(`Tracking at ${pctNow.toFixed(0)}% of target — on pace to hit or beat the number this month. Keep doing what's working rather than changing course.`);
    if (hitRate != null && hitRate < 40) insights.push(`Only hit the daily target on ${daysHit} of ${thisMonthDays.length} days this month (${hitRate.toFixed(0)}%) — the real problem is consistency, not the average. A few big days are propping up a lot of weak ones.`);
    if (hitRate != null && hitRate >= 80) insights.push(`Hit the daily target on ${daysHit} of ${thisMonthDays.length} days (${hitRate.toFixed(0)}%) — a genuinely consistent month, which matters more for planning than any single big day.`);
    if (bestDow && worstDow && bestDow.i !== worstDow.i && dowCounts[bestDow.i] >= 2 && dowCounts[worstDow.i] >= 2) {
      insights.push(`${dowNames[bestDow.i]}s are consistently the strongest day (avg ₹${Math.round(bestDow.v).toLocaleString("en-IN")}), while ${dowNames[worstDow.i]}s lag behind (avg ₹${Math.round(worstDow.v).toLocaleString("en-IN")}). A targeted offer or a 'Meal of the Day' combo on ${dowNames[worstDow.i]}s specifically — rather than a blanket discount every day — usually lifts the weak day without giving away margin on days that already sell well.`);
    }
    if (onlineShare >= 75) insights.push(`Online (Swiggy + Zomato) carries ${onlineShare.toFixed(0)}% of sales — walk-in is a small slice. A strong offline presence (visible signage, in-store promos, local awareness) actually feeds online credibility too, so it's worth not neglecting the shop-front experience even if most orders arrive through an app.`);
    if (shopShare >= 60) insights.push(`Shop (walk-in) carries ${shopShare.toFixed(0)}% of sales — online is underused here. Getting listed on Swiggy POP or a similar budget-meal slot, and making sure the online menu has clear photos and combo pricing, is usually the fastest way to pick up incremental online orders without cannibalising walk-in.`);
    if (discountPct >= 15) insights.push(`Discounts this month equal ${discountPct.toFixed(1)}% of sales — on the high side. Blanket discounts every day quietly erode margin; a targeted offer (specific day, specific item, or first-time-customer only) usually drives similar order volume for a fraction of the giveaway.`);
    if (insights.length === 0) insights.push("No sharp swings this month — performance is holding steady across channels. Worth using this calm stretch to test one small change (a new combo, a weekday offer) since there's no crisis pulling focus elsewhere.");

    const thisMonthTotal = thisM.shop + thisM.swiggy + thisM.zomato;
    const lastMonthTotal = lastM ? lastM.shop + lastM.swiggy + lastM.zomato : null;
    const monthTarget = monthlyTargetFor(oid, nowYm);
    const monthPct = monthTarget > 0 ? (thisMonthTotal / monthTarget) * 100 : 0;
    const monthTrendPct = lastMonthTotal != null && lastMonthTotal > 0 ? ((thisMonthTotal - lastMonthTotal) / lastMonthTotal) * 100 : null;
    let health = "No data";
    if (monthTarget > 0 && thisMonthTotal > 0) health = monthPct >= 90 ? "Strong" : monthPct >= 60 ? "On track" : monthPct >= 30 ? "Needs attention" : "Struggling";
    let trendLabel = "steady";
    if (monthTrendPct != null) trendLabel = monthTrendPct > 10 ? "growing" : monthTrendPct < -10 ? "declining" : "steady";

    // Built from the same real price-vs-demand classification used in the Item Performance
    // tab (classifyItemPerf) — not a separate, simpler version. Same engine, same voice.
    const playbook: string[] = [];
    if (itemPerf && itemPerf.available && itemPerf.classified) {
      const c = itemPerf.classified;
      if (c.funny.headline) playbook.push(`🏆 Stars (${itemPerf.label}): ${c.funny.headline}`);
      if (c.funny.sweet) playbook.push(`💰 Sweet-spot winner: ${c.funny.sweet}`);
      if (c.funny.suspect) playbook.push(`⚠️ Priced-too-high suspect: ${c.funny.suspect}`);
      if (c.funny.dead) playbook.push(`💀 Dead weight: ${c.funny.dead}`);
      if (c.moneyLeft.length) {
        const names = c.moneyLeft.slice(0, 2).map((r: any) => `${r.name} (${r.lost} lost orders)`).join(", ");
        playbook.push(`📉 Money left on the table: ${names} — people wanted these and didn't get them. Usually a stock/availability problem, not a demand problem.`);
      }
      if (itemPerf.gainers.length) {
        const names = itemPerf.gainers.map((r: any) => `${r.name} (+${r.chg.toFixed(0)}%)`).join(", ");
        playbook.push(`📈 Genuinely trending up vs the previous period (${itemPerf.prevLabel} → ${itemPerf.label}): ${names}. Worth understanding what changed and whether it can be repeated.`);
      }
      if (itemPerf.decliners.length) {
        const names = itemPerf.decliners.map((r: any) => `${r.name} (${r.chg.toFixed(0)}%)`).join(", ");
        playbook.push(`📉 Genuinely declining vs the previous period: ${names}. Worth checking whether this is a quality dip, menu fatigue, or just seasonal.`);
      }
      if (playbook.length === 0) playbook.push(`Item Performance data exists (${itemPerf.label}) but nothing stood out sharply this period — no item swung hard enough in either direction to flag.`);
    } else {
      playbook.push("No Item Performance data uploaded yet — once a Prime/billing export is uploaded under the Item Performance tab, this section will show real Stars, Suspects, and Dead-weight items instead of this note.");
    }
    // Structural points — about how the business is set up, not which items are trending,
    // so these stay valid even as the specific bestsellers change over time.
    playbook.push("Brownie Heaven runs three brands from largely the same kitchen — Brownie Heaven, Cakes by Brownie Heaven, and Ice Cream by Brownie Heaven. If one brand is under-target here while another is strong, cross-promoting the weaker brand's items on the stronger brand's listing (same kitchen, same delivery slot) is close to a free lever.");
    playbook.push("A recurring leak flagged in past reconciliation work: a meaningful share of lost revenue (roughly ₹80K/month company-wide, historically) comes from platform cancellations after the order was already acknowledged — a kitchen/ops timing issue, not a demand issue.");

    return { dailyRows, byMonth, months, thisM, lastM, shopChg, swiggyChg, zomatoChg, monthTotals, hitRate, daysHit, daysWithTarget: thisMonthDays.length, bestDow, worstDow, dowNames, onlineShare, shopShare, discountPct, insights, playbook, thisMonthTotal, monthTarget, monthPct, monthTrendPct, health, trendLabel };
  };

  const fetchOutletDeepDive = async (oid: string) => {
    setOutletDeepDiveLoading(true);
    const { data } = await supabase.from("outlet_reports").select("report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value, discount_given").eq("outlet_id", oid).gte("report_date", "2026-06-01");
    setOutletDeepDive(computeOutletAnalysis(oid, data || [], itemPerfContext));
    setOutletDeepDiveLoading(false);
  };
  useEffect(() => { if (activeTab === "owner_outlets" && outletHealthSel) fetchOutletDeepDive(outletHealthSel); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [activeTab, outletHealthSel, itemPerfContext]);

  // Returns the A-vs-B period comparison if all four dates are filled in, else null.
  // Used by the PDF/Excel exports — there's no live on-screen result for this anymore.
  // Resolves a range preset (same list as Download Custom Report) into concrete
  // from/to dates for Period A or B. "custom" leaves the dates as-is for manual editing.
  const applyComparePreset = (which: "A" | "B", presetId: string) => {
    const setPreset = which === "A" ? setComparePresetA : setComparePresetB;
    const setFrom = which === "A" ? setCompareAFrom : setCompareBFrom;
    const setTo = which === "A" ? setCompareATo : setCompareBTo;
    setPreset(presetId);
    setQuickCompareResult(null);
    if (presetId === "none") { setFrom(""); setTo(""); return; }
    if (presetId === "custom") return;
    const r = resolveOutletRange({ preset: presetId });
    setFrom(r.from); setTo(r.to);
  };

  const fetchPeriodComparison = async () => {
    if (!compareAFrom || !compareATo || !compareBFrom || !compareBTo) return null;
    const { data: aRows } = await supabase.from("outlet_reports").select("shop_sales_value, swiggy_sales_value, zomato_sales_value").eq("outlet_id", outletHealthSel).gte("report_date", compareAFrom).lte("report_date", compareATo);
    const { data: bRows } = await supabase.from("outlet_reports").select("shop_sales_value, swiggy_sales_value, zomato_sales_value").eq("outlet_id", outletHealthSel).gte("report_date", compareBFrom).lte("report_date", compareBTo);
    const sum = (rows: any[]) => (rows || []).reduce((acc, r) => ({ shop: acc.shop + (Number(r.shop_sales_value) || 0), swiggy: acc.swiggy + (Number(r.swiggy_sales_value) || 0), zomato: acc.zomato + (Number(r.zomato_sales_value) || 0) }), { shop: 0, swiggy: 0, zomato: 0 });
    const a = sum(aRows || []), b = sum(bRows || []);
    const aTotal = a.shop + a.swiggy + a.zomato, bTotal = b.shop + b.swiggy + b.zomato;
    const pctChg = (curV: number, prevV: number) => prevV === 0 ? null : ((curV - prevV) / prevV) * 100;
    return { a, b, aTotal, bTotal, totalChg: pctChg(bTotal, aTotal), shopChg: pctChg(b.shop, a.shop), swiggyChg: pctChg(b.swiggy, a.swiggy), zomatoChg: pctChg(b.zomato, a.zomato) };
  };

  // Quick on-screen compare — one compact result line, not a table. Full detail stays PDF/Excel-only.
  const runQuickCompare = async () => {
    if (!compareAFrom || !compareATo || !compareBFrom || !compareBTo) { alert("Pick both Period A and Period B first."); return; }
    setQuickCompareBusy(true);
    setQuickCompareResult(await fetchPeriodComparison());
    setQuickCompareBusy(false);
  };

    const downloadOutletHealthPDF = async () => {
    const outletName = OUTLET_NAMES[outletHealthSel] || outletHealthSel;
    if (!outletDeepDive) { alert("No data for this outlet yet."); return; }
    setOutletHealthPdfBusy(true);
    try {
      const dd = outletDeepDive;
      const cmp = await fetchPeriodComparison();
      const C = { bg: "#FAF3E7", card: "#FFFDF8", ink: "#3E2415", soft: "#8A6A4A", line: "#EADBC2", green: "#2E7D32", red: "#C62828", amber: "#C8901E" };
      const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
      const CH_COLS = ["#FACC15", "#FB923C", "#EF4444"];

      // Donut for a Shop/Swiggy/Zomato split — same segmented-circle technique used
      // for the on-screen channel donuts, rendered as raw SVG for the PDF capture.
      const donut = (shop: number, swiggy: number, zomato: number, size = 108) => {
        const total = shop + swiggy + zomato || 1;
        const R = size * 0.34, CX = size / 2, CY = size / 2, SW = size * 0.17, CIRC = 2 * Math.PI * R;
        let acc = 0;
        const segs = [shop, swiggy, zomato].map((v, i) => {
          const frac = v / total, len = frac * CIRC, off = -acc * CIRC;
          acc += frac;
          return { len, gap: CIRC - len, off, col: CH_COLS[i] };
        });
        const circles = segs.map((s) => `<circle cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="${s.col}" stroke-width="${SW}" stroke-dasharray="${s.len.toFixed(2)} ${s.gap.toFixed(2)}" stroke-dashoffset="${s.off.toFixed(2)}" transform="rotate(-90 ${CX} ${CY})"></circle>`).join("");
        return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${CX}" cy="${CY}" r="${R}" fill="none" stroke="${C.line}" stroke-width="${SW}"></circle>${circles}</svg>`;
      };
      const donutLegend = (shop: number, swiggy: number, zomato: number) => {
        const total = shop + swiggy + zomato || 1;
        const rows = [["Shop", shop], ["Swiggy", swiggy], ["Zomato", zomato]] as [string, number][];
        return rows.map(([name, v], i) => `<div style="display:flex;align-items:center;gap:5px;font-size:9px;color:${C.soft};margin-top:2px"><span style="width:8px;height:8px;border-radius:2px;background:${CH_COLS[i]};display:inline-block"></span>${name} ${((v / total) * 100).toFixed(0)}%</div>`).join("");
      };
      const donutCard = (title: string, shop: number, swiggy: number, zomato: number) => `
        <div style="text-align:center">
          <div style="font-size:10px;font-weight:700;color:${C.ink};margin-bottom:6px">${title}</div>
          ${donut(shop, swiggy, zomato)}
          <div style="margin-top:6px;text-align:left;display:inline-block">${donutLegend(shop, swiggy, zomato)}</div>
        </div>`;

      const chRow = (name: string, cur: number, prev: number | undefined, pct: number | null) => {
        const col = pct == null ? C.soft : pct >= 0 ? C.green : C.red;
        const arrow = pct == null ? "" : pct >= 0 ? "Up" : "Down";
        return `<tr><td style="padding:7px 10px;border-bottom:1px solid ${C.line};font-size:12px;font-weight:600">${name}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px">${inr(cur)}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;color:${C.soft}">${prev != null ? inr(prev) : "—"}</td><td style="padding:7px 10px;border-bottom:1px solid ${C.line};text-align:right;font-size:12px;font-weight:700;color:${col}">${pct == null ? "—" : arrow + " " + Math.abs(pct).toFixed(1) + "%"}</td></tr>`;
      };
      const channelRows = chRow("Shop", dd.thisM.shop, dd.lastM?.shop, dd.shopChg) + chRow("Swiggy", dd.thisM.swiggy, dd.lastM?.swiggy, dd.swiggyChg) + chRow("Zomato", dd.thisM.zomato, dd.lastM?.zomato, dd.zomatoChg);
      const chartW = 700, chartH = 220, padL = 60, padR = 20, padT = 20, padB = 34;
      const plotW = chartW - padL - padR, plotH = chartH - padT - padB;
      const maxV = Math.max(...dd.monthTotals.map((x: any) => x.total), 1);
      const stepX = dd.monthTotals.length > 1 ? plotW / (dd.monthTotals.length - 1) : 0;
      const pts = dd.monthTotals.map((x: any, i: number) => {
        const px = padL + i * stepX;
        const py = padT + plotH - (x.total / maxV) * plotH;
        return { px, py, ...x };
      });
      const linePath = pts.map((p: any, i: number) => (i === 0 ? "M" : "L") + p.px.toFixed(1) + " " + p.py.toFixed(1)).join(" ");
      const areaPath = linePath + ` L ${pts[pts.length - 1]?.px.toFixed(1) || padL} ${padT + plotH} L ${padL} ${padT + plotH} Z`;
      const gridLines = [0, 0.25, 0.5, 0.75, 1].map((f) => {
        const y = padT + plotH * f;
        const val = Math.round(maxV * (1 - f));
        return `<line x1="${padL}" y1="${y}" x2="${chartW - padR}" y2="${y}" stroke="${C.line}" stroke-width="1"></line><text x="${padL - 8}" y="${y + 4}" text-anchor="end" style="font-size:9px;fill:${C.soft}">${val >= 100000 ? (val / 100000).toFixed(1) + "L" : val}</text>`;
      }).join("");
      const dots = pts.map((p: any) => `<circle cx="${p.px}" cy="${p.py}" r="4" fill="${C.amber}" stroke="${C.card}" stroke-width="2"></circle><text x="${p.px}" y="${chartH - 8}" text-anchor="middle" style="font-size:9px;fill:${C.soft}">${p.m.slice(5)}</text>`).join("");
      const chartSvg = `<svg width="${chartW}" height="${chartH}" viewBox="0 0 ${chartW} ${chartH}">
        <defs><linearGradient id="areaFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="${C.amber}" stop-opacity="0.35"/><stop offset="100%" stop-color="${C.amber}" stop-opacity="0"/></linearGradient></defs>
        ${gridLines}
        <path d="${areaPath}" fill="url(#areaFill)"></path>
        <path d="${linePath}" fill="none" stroke="${C.amber}" stroke-width="2.5"></path>
        ${dots}
      </svg>`;

      const insightRows = dd.insights.map((s: string) => `<li style="margin:6px 0;font-size:12px;line-height:1.5">${s}</li>`).join("");
      const playbookRows = dd.playbook.map((s: string) => `<li style="margin:6px 0;font-size:11px;line-height:1.5;color:${C.soft}">${s}</li>`).join("");

      const monthNames: Record<string, string> = { "01": "January", "02": "February", "03": "March", "04": "April", "05": "May", "06": "June", "07": "July", "08": "August", "09": "September", "10": "October", "11": "November", "12": "December" };
      const byMonthRows: Record<string, any[]> = {};
      dd.dailyRows.forEach((r: any) => { const ym = r.date.slice(0, 7); if (!byMonthRows[ym]) byMonthRows[ym] = []; byMonthRows[ym].push(r); });
      const dailyTablesByMonth = Object.keys(byMonthRows).sort().reverse().map((ym) => {
        const rows = [...byMonthRows[ym]].reverse();
        const rowsHtml = rows.map((r: any) => `<tr><td style="padding:5px 8px;border-bottom:1px solid ${C.line};font-size:9px">${r.date}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px">${inr(r.shop)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px">${inr(r.swiggy)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px">${inr(r.zomato)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px;font-weight:700">${inr(r.total)}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px;color:${C.soft}">${r.target > 0 ? inr(r.target) : "-"}</td><td style="padding:5px 8px;border-bottom:1px solid ${C.line};text-align:right;font-size:9px;font-weight:700;color:${r.target === 0 ? C.soft : r.pct >= 100 ? C.green : r.pct >= 70 ? C.amber : C.red}">${r.target > 0 ? r.pct.toFixed(0) + "%" : "-"}</td></tr>`).join("");
        const [y, m] = ym.split("-");
        return `<div style="margin-bottom:16px;page-break-inside:avoid;break-inside:avoid;">
          <div style="font-size:12px;font-weight:800;margin:8px 0 4px;color:${C.ink}">${monthNames[m]} ${y} (${rows[0].date.slice(8)}–${rows[rows.length - 1].date.slice(8)})</div>
          <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden;page-break-inside:avoid;break-inside:avoid;">
            <thead><tr style="background:${C.ink}"><th style="padding:6px 8px;text-align:left;color:#FFF6E5;font-size:9px">DATE</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">SHOP</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">SWIGGY</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">ZOMATO</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">TOTAL</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">TARGET</th><th style="padding:6px 8px;text-align:right;color:#FFF6E5;font-size:9px">%</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>`;
      }).join("");

           const monthDonuts = `<div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px 18px;margin-bottom:18px;display:flex;justify-content:space-around;flex-wrap:wrap;gap:12px;page-break-inside:avoid">
        ${donutCard("This month", dd.thisM.shop, dd.thisM.swiggy, dd.thisM.zomato)}
        ${dd.lastM ? donutCard("Last month", dd.lastM.shop, dd.lastM.swiggy, dd.lastM.zomato) : ""}
      </div>`;

      const periodDonuts = cmp ? `<div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px 18px;margin-bottom:12px;display:flex;justify-content:space-around;flex-wrap:wrap;gap:12px;page-break-inside:avoid">
        ${donutCard(`Period A (${compareAFrom} – ${compareATo})`, cmp.a.shop, cmp.a.swiggy, cmp.a.zomato)}
        ${donutCard(`Period B (${compareBFrom} – ${compareBTo})`, cmp.b.shop, cmp.b.swiggy, cmp.b.zomato)}
      </div>` : "";

      const html = `<div style="width:794px;background:${C.bg};font-family:'Segoe UI',Arial,sans-serif;color:${C.ink};padding:34px">
        <div style="font-size:22px;font-weight:900">Brownie Heaven — Outlet Health: ${outletName}</div>
        <div style="font-size:11px;color:${C.soft};margin-bottom:16px">Since June launch · generated ${new Date().toISOString().split("T")[0]}</div>
        <div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:18px;margin-bottom:16px">
          <div style="font-size:13px;color:${C.soft}">This month</div>
          <div style="font-size:28px;font-weight:900">${inr(dd.thisMonthTotal)}</div>
          <div style="font-size:13px;color:${dd.health === "Strong" ? C.green : dd.health === "Struggling" ? C.red : C.amber};font-weight:700">${dd.health}${dd.monthPct > 0 ? ` · ${dd.monthPct.toFixed(0)}% of target` : ""}${dd.hitRate != null ? ` · hit target ${dd.daysHit}/${dd.daysWithTarget} days (${dd.hitRate.toFixed(0)}%)` : ""}</div>
        </div>
        ${monthDonuts}
        <div style="background:${C.card};border:2px solid ${C.amber};border-radius:12px;padding:18px;margin-bottom:18px;page-break-inside:avoid">
          <div style="font-size:15px;font-weight:800;margin-bottom:8px;color:${C.amber}">📈 Business Insights &amp; What To Do About It</div>
          <ul style="margin:0;padding-left:18px">${insightRows}</ul>
        </div>
        <div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:18px;margin-bottom:18px;page-break-inside:avoid">
          <div style="font-size:13px;font-weight:800;margin-bottom:4px;color:${C.ink}">🍫 Real sellers &amp; item-level signals</div>
          <div style="font-size:10px;color:${C.soft};margin-bottom:8px">From actual Item Performance uploads (Prime/billing exports) and standing operational structure — not researched or generic.</div>
          <ul style="margin:0;padding-left:18px">${playbookRows}</ul>
        </div>
        <div style="font-size:14px;font-weight:800;margin-bottom:8px">Revenue trend — month by month</div>
        <div style="background:${C.card};border:1px solid ${C.line};border-radius:12px;padding:16px 18px;margin-bottom:18px">
          ${chartSvg}
        </div>
        <div style="font-size:14px;font-weight:800;margin-bottom:8px">Did we gain or lose? — vs last month</div>
        <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden;margin-bottom:18px;page-break-inside:avoid">
          <thead><tr style="background:${C.ink}"><th style="padding:7px 10px;text-align:left;color:#FFF6E5;font-size:9px">CHANNEL</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">THIS MONTH</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">LAST MONTH</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">CHANGE</th></tr></thead>
          <tbody>${channelRows}</tbody>
        </table>
          ${cmp ? `<div style="font-size:14px;font-weight:800;margin-bottom:8px">Period comparison — ${compareAFrom} to ${compareATo} vs ${compareBFrom} to ${compareBTo}</div>
        ${periodDonuts}
        <table style="width:100%;border-collapse:collapse;background:${C.card};border:1px solid ${C.line};border-radius:10px;overflow:hidden;margin-bottom:18px;page-break-inside:avoid">
          <thead><tr style="background:${C.ink}"><th style="padding:7px 10px;text-align:left;color:#FFF6E5;font-size:9px">CHANNEL</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">PERIOD A</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">PERIOD B</th><th style="padding:7px 10px;text-align:right;color:#FFF6E5;font-size:9px">CHANGE</th></tr></thead>
          <tbody>
            ${chRow("Shop", cmp.b.shop, cmp.a.shop, cmp.shopChg)}
            ${chRow("Swiggy", cmp.b.swiggy, cmp.a.swiggy, cmp.swiggyChg)}
            ${chRow("Zomato", cmp.b.zomato, cmp.a.zomato, cmp.zomatoChg)}
            <tr style="background:${C.line}"><td style="padding:8px 10px;font-size:12px;font-weight:900">TOTAL</td><td style="padding:8px 10px;text-align:right;font-size:12px;font-weight:700">${inr(cmp.aTotal)}</td><td style="padding:8px 10px;text-align:right;font-size:12px;font-weight:900">${inr(cmp.bTotal)}</td><td style="padding:8px 10px;text-align:right;font-size:12px;font-weight:900;color:${cmp.totalChg == null ? C.soft : cmp.totalChg >= 0 ? C.green : C.red}">${cmp.totalChg == null ? "—" : (cmp.totalChg >= 0 ? "Up " : "Down ") + Math.abs(cmp.totalChg).toFixed(1) + "%"}</td></tr>
          </tbody>
        </table>` : ""}
        <div style="font-size:14px;font-weight:800;margin-bottom:8px">Full daily detail — by month</div>
        ${dailyTablesByMonth}
      </div>`;

      const lib = await loadH2P();
      window.scrollTo(0, 0); await new Promise((r) => setTimeout(r, 50));
      const holder = document.createElement("div"); holder.style.position = "fixed"; holder.style.left = "-9999px"; holder.style.top = "0"; holder.innerHTML = html; document.body.appendChild(holder);
      try { await lib().set({ margin: 0, filename: `OutletHealth_${outletName.replace(/\s+/g, "_")}.pdf`, image: { type: "jpeg", quality: 0.97 }, html2canvas: { scale: 2, backgroundColor: C.bg }, jsPDF: { unit: "pt", format: "a4", orientation: "portrait" }, pagebreak: { mode: ["css", "legacy"] } }).from(holder.firstElementChild).save(); }
      finally { document.body.removeChild(holder); }
    } catch (e: any) {
      alert("Failed to generate: " + (e?.message || "error"));
    }
    setOutletHealthPdfBusy(false);
  };


  const downloadOutletHealthExcel = async () => {
    const outletName = OUTLET_NAMES[outletHealthSel] || outletHealthSel;
    if (!outletDeepDive) { alert("No data for this outlet yet."); return; }
    setOutletHealthXlsxBusy(true);
    try {
      const cmp = await fetchPeriodComparison();
      const dailySheet = outletDeepDive.dailyRows.map((r: any) => ({
        Date: r.date, Shop: r.shop, Swiggy: r.swiggy, Zomato: r.zomato, Total: r.total, Target: r.target,
        "%": r.target > 0 ? Math.round(r.pct) + "%" : "-",
      }));
      const monthlySheet = outletDeepDive.monthTotals.map((m: any) => ({
        Month: m.m, Total: Math.round(m.total), Target: m.target, "%": m.target > 0 ? Math.round((m.total / m.target) * 100) + "%" : "-",
      }));
      const insightsSheet = [
        ...outletDeepDive.insights.map((s: string) => ({ Type: "This outlet, this month", Insight: s })),
        ...outletDeepDive.playbook.map((s: string) => ({ Type: "Brownie Heaven playbook", Insight: s })),
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(insightsSheet), "Insights");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(monthlySheet), "Monthly Summary");
      if (cmp) {
        const compareSheet = [
          { Channel: "Shop", [`Period A (${compareAFrom} to ${compareATo})`]: cmp.a.shop, [`Period B (${compareBFrom} to ${compareBTo})`]: cmp.b.shop, "Change %": cmp.shopChg == null ? "-" : cmp.shopChg.toFixed(1) + "%" },
          { Channel: "Swiggy", [`Period A (${compareAFrom} to ${compareATo})`]: cmp.a.swiggy, [`Period B (${compareBFrom} to ${compareBTo})`]: cmp.b.swiggy, "Change %": cmp.swiggyChg == null ? "-" : cmp.swiggyChg.toFixed(1) + "%" },
          { Channel: "Zomato", [`Period A (${compareAFrom} to ${compareATo})`]: cmp.a.zomato, [`Period B (${compareBFrom} to ${compareBTo})`]: cmp.b.zomato, "Change %": cmp.zomatoChg == null ? "-" : cmp.zomatoChg.toFixed(1) + "%" },
          { Channel: "TOTAL", [`Period A (${compareAFrom} to ${compareATo})`]: cmp.aTotal, [`Period B (${compareBFrom} to ${compareBTo})`]: cmp.bTotal, "Change %": cmp.totalChg == null ? "-" : cmp.totalChg.toFixed(1) + "%" },
        ];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(compareSheet), "Period Comparison");
      }
      XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(dailySheet), "Daily (by month, latest last)");
      XLSX.writeFile(wb, `OutletHealth_${outletName.replace(/\s+/g, "_")}.xlsx`);
    } catch (e: any) {
      alert("Failed to generate: " + (e?.message || "error"));
    }
    setOutletHealthXlsxBusy(false);
  };

  const anAgg = (() => {
    const ch = { shop: { c: 0, v: 0 }, swiggy: { c: 0, v: 0 }, zomato: { c: 0, v: 0 } };
    const byOutlet: Record<string, number> = {};
    let discountTotal = 0;
    anRows.forEach((r: any) => {
      ch.shop.c += Number(r.shop_sales_count) || 0; ch.shop.v += Number(r.shop_sales_value) || 0;
      ch.swiggy.c += Number(r.swiggy_sales_count) || 0; ch.swiggy.v += Number(r.swiggy_sales_value) || 0;
      ch.zomato.c += Number(r.zomato_sales_count) || 0; ch.zomato.v += Number(r.zomato_sales_value) || 0;
      discountTotal += Number(r.discount_given) || 0;
      const rowTotal = (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
      if (r.outlet_id) byOutlet[r.outlet_id] = (byOutlet[r.outlet_id] || 0) + rowTotal;
    });
    const totalV = ch.shop.v + ch.swiggy.v + ch.zomato.v;
    const totalC = ch.shop.c + ch.swiggy.c + ch.zomato.c;
    const outletRanked = Object.entries(byOutlet).map(([oid, v]) => ({ oid, name: OUTLET_NAMES[oid] || oid, v })).sort((a, b) => b.v - a.v);
    return { ch, totalV, totalC, discountTotal, outletRanked };
  })();
  const anINR = (v: number) => "₹" + Math.round(v || 0).toLocaleString("en-IN");
  const anPct = (part: number, whole: number) => (whole > 0 ? Math.round((part / whole) * 100) : 0);
  const [outletHistoryDate, setOutletHistoryDate] = useState<string>(new Date().toISOString().split("T")[0]);
  const [lastOutletRatings, setLastOutletRatings] = useState<Record<string, OutletReport>>({});
  const [allOutletReports, setAllOutletReports] = useState<OutletReport[]>([]);
  const [historyDate, setHistoryDate] = useState<string>(new Date().toISOString().split("T")[0]);
  const [historyReports, setHistoryReports] = useState<Report[]>([]);
  const [historyOutletReports, setHistoryOutletReports] = useState<OutletReport[]>([]);
  const [historyLoading, setHistoryLoading] = useState(false);

  useEffect(() => {
    const stored = localStorage.getItem("currentUser");
    if (!stored) { router.push("/"); return; }
    let parsed: any;
    try {
      parsed = JSON.parse(stored);
      if (typeof parsed === "string") { localStorage.removeItem("currentUser"); router.push("/"); return; }
    } catch { localStorage.removeItem("currentUser"); router.push("/"); return; }
    const _fresh = ALL_STAFF.find(s => s.id === parsed.id);
    if (_fresh) parsed = { ...parsed, outlets: (_fresh as any).outlets, role: _fresh.role };
    setUser(parsed);
    fetchDayOff(parsed.id);
    fetchTasks(parsed);
    fetchReports(parsed);
   fetchAttendance(parsed, new Date(Date.now() - 86400000).toISOString().split("T")[0]);
   fetchAttendanceWeekOvertime(parsed);
   fetchSalesTargets(parsed);
  fetchOutletReports(parsed);
        fetchLastOutletRatings(parsed);
   if (parsed.role === "Owner" || parsed.role === "Manager") fetchAllOutletReports();
    fetchMsgUnread(parsed);
  }, [router]);

  // Open straight to a specific tab when a notification link asks for it
  // (e.g. a push notification's "Reply" deep link into Messages).
  useEffect(() => {
    if (typeof window === "undefined") return;
    const params = new URLSearchParams(window.location.search);
    const tab = params.get("tab");
    if (tab === "messages") setActiveTab("messages");
  }, []);

 useEffect(() => { if (activeOutlet) fetchReviews(activeOutlet, outletEntryDate); else setReviews([]); }, [activeOutlet, outletEntryDate]);
  useEffect(() => {
    if (!user) return;
    const interval = setInterval(() => { fetchTasks(user); fetchReports(user); fetchMsgUnread(user); }, 30000);
    return () => clearInterval(interval);
  }, [user]);

  const fetchMsgUnread = async (u: Staff) => {
    const filterId = u.role === "Owner" ? "nishant" : u.id;
    const { count } = await supabase.from("messages").select("id", { count: "exact", head: true }).eq("to_id", filterId).is("read_at", null);
    setMsgUnreadCount(count || 0);
  };

  const fetchTasks = async (u: Staff) => {
    setLoading(true);
    let query = supabase.from("tasks").select("*").order("created_at", { ascending: false });
   if (u.role !== "Owner" && u.role !== "Manager") {
  const staffOutlets = (u as Staff & { outlets?: string[] }).outlets || [];
  if (staffOutlets.length > 0) {
    query = query.or(`assigned_to.eq.${u.id},outlet_id.in.(${staffOutlets.join(",")})`);
  } else {
    query = query.eq("assigned_to", u.id);
  }
}
    const { data } = await query;
    setTasks(data || []);
    setLoading(false);
    if (u.role !== "Owner") {
      const overdue = (data || []).find((t: Task) => t.status !== "completed" && new Date(t.due_at) < new Date());
     if (overdue) { setOverdueTask(overdue); playAlert(); }
    }
  };
  const exportCSV = () => {
  if (historyOutletReports.length === 0) { alert("No outlet reports for this date."); return; }
  const headers = ["Outlet", "Manager", "Date", "Shop Sales Value", "Shop Sales Count", "Swiggy Value", "Swiggy Count", "Zomato Value", "Zomato Count", "Total Sales", "BH Sales", "CBH Sales", "Target", "Swiggy Live", "Zomato Live", "Discount Running", "Discount Rate Good", "Unavailable Items", "Expiry Count", "Expiry Items", "Complimentary Count", "Complimentary Product Name", "Complimentary Reason", "Issues", "Action Taken", "Submitted At", "Late"];
  const rows = historyOutletReports.map(r => {
    const manager = ALL_STAFF.find(s => (s.outlets as string[]).includes(r.outlet_id))?.name || "—";
    const total = Number(r.shop_sales_value) + Number(r.swiggy_sales_value) + Number(r.zomato_sales_value);
    return [
      OUTLET_NAMES[r.outlet_id] || r.outlet_id.replace(/_/g, " "),
      manager,
      r.report_date,
      r.shop_sales_value,
      r.shop_sales_count,
      r.swiggy_sales_value,
      r.swiggy_sales_count,
      r.zomato_sales_value,
      r.zomato_sales_count,
      total,
      (r as any).bh_shop_sales_value == null ? "" : ["shop", "swiggy", "zomato"].reduce((a, ch) => a + (Number((r as any)[`bh_${ch}_sales_value`]) || 0), 0),
      (r as any).cbh_shop_sales_value == null ? "" : ["shop", "swiggy", "zomato"].reduce((a, ch) => a + (Number((r as any)[`cbh_${ch}_sales_value`]) || 0), 0),
      r.target,
      r.swiggy_live ? "Yes" : "No",
      r.zomato_live ? "Yes" : "No",
      r.discount_running || "",
      r.discount_rate_good ? "Yes" : "No",
      r.unavailable_items || "",
      r.expiry_count,
      r.expiry_items || "",
      r.complimentary_count,
      r.complimentary_items || "",
      r.complimentary_reason || "",
      r.issues || "",
      r.action_taken || "",
      new Date(r.submitted_at).toLocaleString("en-IN"),
      r.is_late ? "Yes" : "No",
    ];
  });
  const csv = [headers, ...rows].map(r => r.map(v => `"${v}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `brownie-heaven-outlet-reports-${historyDate}.csv`;
  a.click();
  URL.revokeObjectURL(url);
};
  const fetchHistoryReports = async (date: string) => {
  setHistoryLoading(true);
  const start = `${date}T00:00:00.000Z`;
  const end = `${date}T23:59:59.999Z`;
  const [{ data: reps }, { data: outletReps }] = await Promise.all([
    supabase.from("reports").select("*").gte("submitted_at", start).lte("submitted_at", end).order("submitted_at", { ascending: false }),
    supabase.from("outlet_reports").select("*").eq("report_date", date).order("submitted_at", { ascending: false }),
  ]);
  setHistoryReports(reps || []);
  setHistoryOutletReports(outletReps || []);
  setHistoryLoading(false);
};
  const fetchAllOutletReportsByDate = async (date: string) => {
  const { data } = await supabase
    .from("outlet_reports")
    .select("*")
    .eq("report_date", date)
    .order("submitted_at", { ascending: false });
  setAllOutletReports(data || []);
};
 const fetchAllOutletReports = async (date?: string) => {
  const today = date || new Date().toISOString().split("T")[0];
  const { data } = await supabase
    .from("outlet_reports")
    .select("*")
    .eq("report_date", today)
    .order("submitted_at", { ascending: false });
  setAllOutletReports(data || []);
};
  const fetchOutletReportsByDate = async (date: string, u: Staff | null = user) => {
  if (!u) return;
  const _isMgr = u.role === "Owner" || u.role === "Manager";
  let _q = supabase.from("outlet_reports").select("*").eq("report_date", date);
  if (!_isMgr) _q = _q.eq("staff_id", u.id);
  const { data } = await _q;
  const map: Record<string, OutletReport> = {};
  (data || []).forEach((r: OutletReport) => { map[r.outlet_id] = r; });
  setOutletReports(map);
};
  const fetchLastOutletRatings = async (u: Staff) => {
  const { data } = await supabase
    .from("outlet_reports")
    .select("outlet_id, bh_google_rating, bh_swiggy_rating, bh_zomato_rating, cbh_google_rating, cbh_swiggy_rating, cbh_zomato_rating, icbh_google_rating, icbh_swiggy_rating, icbh_zomato_rating")
    .eq("staff_id", u.id)
    .not("bh_google_rating", "is", null)
    .order("submitted_at", { ascending: false });
  const map: Record<string, OutletReport> = {};
 (data || []).forEach((r: any) => {
    if (!map[r.outlet_id]) map[r.outlet_id] = r;
  });
 
  setLastOutletRatings(map);
};
const fetchOutletReports = async (u: Staff) => {
  const today = new Date().toISOString().split("T")[0];
  const _isMgr = u.role === "Owner" || u.role === "Manager";
  let _q = supabase.from("outlet_reports").select("*").eq("report_date", today);
  if (!_isMgr) _q = _q.eq("staff_id", u.id);
  const { data } = await _q;
  const map: Record<string, OutletReport> = {};
  (data || []).forEach((r: OutletReport) => { map[r.outlet_id] = r; });
  setOutletReports(map);
  return map;
};
  const fetchReportByDate = async (date: string) => {
  if (!user) return;
  const start = `${date}T00:00:00.000Z`;
  const end = `${date}T23:59:59.999Z`;
  const { data } = await supabase
    .from("reports")
    .select("*")
    .eq("staff_id", user.id)
    .gte("submitted_at", start)
    .lte("submitted_at", end)
    .single();
  setReportByDate(data || null);
};
  const fetchReports = async (u: Staff) => {
    let query = supabase.from("reports").select("*").order("submitted_at", { ascending: false });
    if (u.role !== "Owner") query = query.eq("staff_id", u.id);
    const { data } = await query;
    setReports(data || []);
    const today = new Date().toDateString();
    const mine = (data || []).find((r: Report) => r.staff_id === u.id && new Date(r.submitted_at).toDateString() === today);
    setTodayReport(mine || null);
  };

 const fetchDayOff = async (uid: string) => {
    const today = new Date().toISOString().split("T")[0];
    const { data } = await supabase.from("day_off").select("staff_id").eq("off_date", today);
    const ids = (data || []).map((r: any) => r.staff_id);
    setOffToday(ids);
    setReportOffDay(ids.includes(uid));
  };
  const toggleOffDay = async () => {
    if (!user) return;
    const today = new Date().toISOString().split("T")[0];
    if (reportOffDay) {
      await supabase.from("day_off").delete().eq("staff_id", user.id).eq("off_date", today);
    } else {
      await supabase.from("day_off").insert({ staff_id: user.id, off_date: today });
    }
    fetchDayOff(user.id);
  };
 const submitReport = async () => {
    if (!user) return;
    const _fields = REPORT_FIELDS[user.id] || [];
    const _missing = _fields.filter((f) => {
      if (user.id === "arun" && (f.key === "achievement" || f.key === "target")) return false;
      return !reportData[f.key] || !String(reportData[f.key]).trim();
    });
    if (_missing.length) { alert("Please fill all fields before submitting.\n\nMissing: " + _missing.map((f) => f.label).join(", ")); return; }
    setReportSubmitting(true);
    const _today = new Date().toISOString().split("T")[0];
    const _date = reportHistoryDate || _today;
    const _isBackfill = _date < _today;
    const deadline = new Date();
    deadline.setHours(22, 0, 0, 0);
    const isLate = !_isBackfill && new Date() > deadline;
    const { data: _existing } = await supabase.from("reports").select("id").eq("staff_id", user.id).eq("report_date", _date);
    const _isEdit = (_existing?.length || 0) > 0;
    const finalData: Record<string, string> = { ...reportData };
    if (user.id === "arun") {
      if (!finalData.target || !String(finalData.target).trim()) finalData.target = "299666";
      const _ts = parseFloat(finalData.total_sales || "0");
      const _tg = parseFloat(finalData.target || "299666");
      finalData.achievement = _tg ? (_ts / _tg * 100).toFixed(1) + "%" : "";
    }
    const content = Object.entries(finalData).map(([k, v]) => `${k}: ${v}`).join(", ");
    const { data, error } = await supabase.from("reports").insert({
      staff_id: user.id,
      content,
      report_date: _date,
      is_backfill: _isBackfill,
      is_late: reportOffDay ? false : isLate,
      submitted_at: new Date().toISOString(),
      report_data: finalData,
      staff_role: user.role,
      no_points: reportOffDay,
    }).select().single();
    setReportSubmitting(false);
    if (error) { alert("Error: " + error.message); return; }
    if (_date === _today) setTodayReport(data);
    if (!reportOffDay) {
      if (_isEdit) celebrate(0, "Report updated — no extra points");
      else if (_isBackfill) celebrate(-5, "Back-dated report — -5");
      else if (isLate) celebrate(-5, "After 10 PM — -5");
      else celebrate(10);
    }
    setReportData({});
    fetchReports(user);
    if (_date !== _today) fetchReportByDate(_date);
  };

const runTargetCheck = async (u: Staff) => {
    const outlets = (u as Staff & { outlets?: string[] }).outlets || [];
    if (!outlets.length) return;
    const yest = new Date(Date.now() - 86400000).toISOString().split("T")[0];
    const { data } = await supabase.from("outlet_reports").select("outlet_id,shop_sales_value,swiggy_sales_value,zomato_sales_value").in("outlet_id", outlets).eq("report_date", yest);
    const byOutlet: Record<string, number> = {};
    (data || []).forEach((r: any) => { byOutlet[r.outlet_id] = (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0); });
    const results = outlets.map((oid) => {
      const target = parseFloat(OUTLET_TARGETS[oid] || "");
      const actual = byOutlet[oid];
      const name = OUTLET_NAMES[oid] || oid;
      if (!target || isNaN(target)) return { oid, name, status: "notarget" };
      if (actual === undefined) return { oid, name, status: "noentry", target };
      return { oid, name, status: actual >= target ? "win" : "miss", target, actual };
    });
    setTargetCheck(results);
  };

  const fetchSalesTargets = async (u: Staff) => {
    const isViewer = u.role === "Owner" || u.role === "Manager";
    const outlets = (u as Staff & { outlets?: string[] }).outlets || [];
    if (!isViewer && !outlets.length) return;
    let query = supabase.from("sales_target").select("*");
    if (!isViewer) query = query.in("outlet_id", outlets);
    const { data } = await query;
    const map: Record<string, any> = {};
    (data || []).forEach((row: any) => {
      if (!map[row.outlet_id]) map[row.outlet_id] = {};
      map[row.outlet_id][row.brand] = row.line_items;
    });
    setSalesTargets(map);
  };

 const saveSalesTarget = async (outletId: string, brand: string, li: any) => {
    setStSaving(true);
    const num = (k: string, fb: number) => { const v = stEditValues[k]; return v !== undefined && v !== "" ? (parseFloat(String(v).replace(/,/g, "")) || 0) : fb; };
   const f = li?.fixed || {}; const t = li?.targets || {};
    const mk = stDate.slice(0, 7);
    const updated = {
      sales: { ...(li?.sales || {}), [stDate]: { net: num("net", Number(li?.sales?.[stDate]?.net) || 0), online: num("online", Number(li?.sales?.[stDate]?.online) || 0) } },
      monthly: { ...(li?.monthly || {}), [mk]: { net: num("mnet", Number(li?.monthly?.[mk]?.net) || 0), online: num("monline", Number(li?.monthly?.[mk]?.online) || 0) } },
      fixed: { staff: num("staff", Number(f.staff) || 0), rent: num("rent", Number(f.rent) || 0), eb: num("eb", Number(f.eb) || 0), transport: num("transport", Number(f.transport) || 0), pest: num("pest", Number(f.pest) || 0), water: num("water", Number(f.water) || 0), airtel: num("airtel", Number(f.airtel) || 0) },
      targets: { a: num("a", Number(t.a) || 0), b: num("b", Number(t.b) || 0) },
    };
    const { error } = await supabase.from("sales_target").upsert({ outlet_id: outletId, brand, line_items: updated, updated_at: new Date().toISOString() }, { onConflict: "outlet_id,brand" });
    setStSaving(false);
    if (error) { alert("Error: " + error.message); return; }
    setStEditing(null); setStEditValues({});
    if (user) fetchSalesTargets(user);
  };

  const fetchAttendance = async (u: Staff, date: string) => {
    const { data } = await supabase.from("attendance").select("*").eq("staff_id", u.id).eq("attendance_date", date).maybeSingle();
    setTodayAttendance(data || null);
    if (!data) setAttendanceData({ present: "", absent: "", late: "", absent_names: "", late_names: "", overtime_hours: "" });
  };

  // Rolling 7-day overtime total (People & Labour ask #1: overtime monitoring).
  // Company-wide, since attendance is logged as one daily company-wide count,
  // not per-employee — matches the grain the rest of this form already uses.
  const fetchAttendanceWeekOvertime = async (u: Staff) => {
    const since = new Date(); since.setDate(since.getDate() - 6);
    const { data } = await supabase.from("attendance").select("overtime_hours").eq("staff_id", u.id).gte("attendance_date", since.toISOString().slice(0, 10));
    setAttendanceWeekOvertime((data || []).reduce((s: number, r: any) => s + (Number(r.overtime_hours) || 0), 0));
  };

  const submitAttendance = async () => {
    if (!user) return;
    setAttendanceSubmitting(true);
    const { data, error } = await supabase.from("attendance").upsert({
      staff_id: user.id,
      attendance_date: attendanceDate,
     present: parseInt(attendanceData.present) || 0,
      absent: parseInt(attendanceData.absent) || 0,
      late: parseInt(attendanceData.late) || 0,
      absent_names: attendanceData.absent_names.trim() || null,
      late_names: attendanceData.late_names.trim() || null,
      overtime_hours: parseFloat(attendanceData.overtime_hours) || 0,
      submitted_at: new Date().toISOString(),
    }, { onConflict: "staff_id,attendance_date" }).select().single();
    setAttendanceSubmitting(false);
    if (error) { alert("Error: " + error.message); return; }
    setTodayAttendance(data);
    setAttendanceData({ present: "", absent: "", late: "", absent_names: "", late_names: "", overtime_hours: "" });
    fetchAttendanceWeekOvertime(user);
  };

  const assignTask = async () => {
    if (!taskTitle.trim() || !user) return;
    setSubmitting(true);
    const dueAt = new Date(Date.now() + parseFloat(taskDueHours) * 60 * 60 * 1000).toISOString();
    const { error } = await supabase.from("tasks").insert({
      title: taskTitle.trim(), description: taskDesc.trim(),
      assigned_to: taskAssignee, assigned_by: user.id,
      priority: taskPriority, status: "assigned", due_at: dueAt,
      outlet_id: taskOutlet || null,
    });
    setSubmitting(false);
    if (error) { alert("Error: " + error.message); return; }
    setShowModal(false);
    setTaskTitle(""); setTaskDesc(""); setTaskPriority("medium"); setTaskDueHours("4"); setTaskOutlet("");
    fetchTasks(user);
  };

  const updateStatus = async (taskId: string, status: string) => {
    await supabase.from("tasks").update({ status, ...(status === "completed" ? { completed_at: new Date().toISOString() } : {}) }).eq("id", taskId);
    if (user) fetchTasks(user);
  };

  const deleteTask = async (taskId: string) => {
    if (!confirm("Delete this task?")) return;
    await supabase.from("tasks").delete().eq("id", taskId);
    if (user) fetchTasks(user);
  };

 const updatePin = async () => {
  if (!newPin || newPin.length < 4) { setPinMsg("PIN must be at least 4 digits."); return; }
  if (!user) return;
  const { error } = await supabase
    .from("staff")
    .update({ pin: newPin })
    .eq("id", user.id)
    .select()
    .limit(1);
  if (error) { setPinMsg("Error: " + error.message); return; }
  setPinMsg("PIN updated successfully!");
  setNewPin("");
  setTimeout(() => { setShowPinModal(false); setPinMsg(""); }, 1500);
};
 const editOutletReport = async (outletId: string) => {
  const r = outletReports[outletId];
  if (!r) return;
  setOutletReportData({
    target: String(r.target),
    shop_sales_count: String(r.shop_sales_count),
    shop_sales_value: String(r.shop_sales_value),
    swiggy_sales_count: String(r.swiggy_sales_count),
    swiggy_sales_value: String(r.swiggy_sales_value),
    discount_given: String(r.discount_given || ""),
    zomato_sales_count: String(r.zomato_sales_count),
    zomato_sales_value: String(r.zomato_sales_value),
    ...Object.fromEntries(BRAND_SALES_KEYS.map((k) => [k, (r as any)[k] === null || (r as any)[k] === undefined ? "" : String((r as any)[k])])),
    swiggy_live: r.swiggy_live ? "yes" : "no",
    zomato_live: r.zomato_live ? "yes" : "no",
    discount_running: r.discount_running || "",
    discount_rate_good: r.discount_rate_good ? "yes" : "no",
    unavailable_items: r.unavailable_items || "",
    expiry_count: String(r.expiry_count),
    expiry_items: r.expiry_items || "",
    complimentary_count: String(r.complimentary_count),
    complimentary_items: r.complimentary_items || "",
    complimentary_reason: r.complimentary_reason || "",
   bh_google_rating: String(r.bh_google_rating || ""),
  bh_swiggy_rating: String(r.bh_swiggy_rating || ""),
  bh_zomato_rating: String(r.bh_zomato_rating || ""),
  cbh_google_rating: String(r.cbh_google_rating || ""),
  cbh_swiggy_rating: String(r.cbh_swiggy_rating || ""),
  cbh_zomato_rating: String(r.cbh_zomato_rating || ""),
 icbh_google_rating: String(r.icbh_google_rating || ""),
 icbh_swiggy_rating: String(r.icbh_swiggy_rating || ""),
 icbh_zomato_rating: String(r.icbh_zomato_rating || ""),
    issues: r.issues || "",
    action_taken: r.action_taken || "",
    is_edited: "true",
    editing_id: r.id,
  });
  setOutletReports(prev => {
    const updated = { ...prev };
    delete updated[outletId];
    return updated;
  });
};
const parseFileRows = async (file: File): Promise<any[][]> => {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  let rows: any[][] = [];
  wb.SheetNames.forEach((sn) => { const rr = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false }) as any[][]; rows = rows.concat(rr); });
  return rows;
};
const parseFileSheets = async (file: File): Promise<Record<string, any[][]>> => {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const out: Record<string, any[][]> = {};
  wb.SheetNames.forEach((sn) => { out[sn] = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, blankrows: false }) as any[][]; });
  return out;
};
const extractMisTotals = (rows: any[][]): { net: number | null; swiggy: number | null; zomato: number | null } => {
  let hr = -1;
  for (let i = 0; i < rows.length; i++) { const r = rows[i] || []; if (r.some((c: any) => typeof c === "string" && /net sales \(after discount/i.test(c))) { hr = i; break; } }
  if (hr < 0) return { net: null, swiggy: null, zomato: null };
  const hdr: string[] = (rows[hr] || []).map((c: any) => (typeof c === "string" ? c : ""));
  const find = (re: RegExp, from = 0) => { for (let c = from; c < hdr.length; c++) { if (re.test(hdr[c])) return c; } return -1; };
  const shopCol = find(/net sales \(after discount/i);
  const swGross = find(/gross sales swiggy/i);
  const zoGross = find(/gross sales zomato/i);
  const swNet = swGross >= 0 ? find(/net sales after discount/i, swGross + 1) : -1;
  const zoNet = zoGross >= 0 ? find(/net sales after discount/i, zoGross + 1) : -1;
  let tr = -1;
  for (let i = hr + 1; i < rows.length; i++) { const r = rows[i] || []; if (typeof r[0] === "string" && /^\s*total\s*$/i.test(r[0])) { tr = i; break; } }
  if (tr < 0) return { net: null, swiggy: null, zomato: null };
  const trow = rows[tr] || [];
  const numAt = (c: number): number | null => { if (c < 0) return null; const v = trow[c]; if (typeof v === "number" && !isNaN(v)) return v; const n = parseFloat(String(v).replace(/[,₹\s]/g, "")); return isNaN(n) ? null : n; };
  const net = numAt(shopCol), sw = numAt(swNet), zo = numAt(zoNet);
  return { net, swiggy: sw, zomato: zo };
};
const findVal = (rows: any[][], regex: RegExp): number | null => {
  for (const row of rows) {
    if (!row) continue;
    let li = -1;
    for (let i = 0; i < row.length; i++) { if (typeof row[i] === "string" && row[i].trim()) { li = i; break; } }
    if (li < 0) continue;
    if (regex.test(String(row[li]))) {
      for (let j = li + 1; j < row.length; j++) {
        const v = row[j];
        if (typeof v === "number" && !isNaN(v)) return v;
        if (typeof v === "string" && v.trim()) { const n = parseFloat(v.replace(/[,₹\s]/g, "")); if (!isNaN(n)) return n; }
      }
    }
  }
  return null;
};
const stExtractOutlet = async (oid: string, brand: string) => {
  const key = oid + "_" + brand;
  const f = stFiles[key] || {};
  if (!f.mis && !f.pnl) { setStUpMsg(m => ({ ...m, [key]: "Pick at least one file." })); return; }
  setStUpBusy(key); setStUpMsg(m => ({ ...m, [key]: "Reading..." }));
  try {
    const e: any = {};
   if (f.mis) {
      const sheets = await parseFileSheets(f.mis);
      const monthName = new Date(stDate + "T00:00:00").toLocaleString("en-US", { month: "long" });
      const snames = Object.keys(sheets);
      const sname = snames.find(s => s.toLowerCase() === monthName.toLowerCase()) || snames.find(s => s.toLowerCase().includes(monthName.toLowerCase()));
      const mis = extractMisTotals(sname ? sheets[sname] : []);
      e.net = mis.net; e.swiggy = mis.swiggy; e.zomato = mis.zomato;
    }
    if (f.pnl) {
      const rows = await parseFileRows(f.pnl);
      e.pest = findVal(rows, /pest/i); e.water = findVal(rows, /water/i); e.airtel = findVal(rows, /airtel|wifi|internet|broadband/i);
      if (brand === "BH") { e.rent = findVal(rows, /rent/i); e.staff = findVal(rows, /salar|staff|wage/i); e.eb = findVal(rows, /electric|power|\beb\b/i); e.transport = findVal(rows, /transport|convey/i); }
    }
    setStUpload(u => ({ ...u, [key]: e })); setStUpMsg(m => ({ ...m, [key]: "" }));
  } catch (err: any) { setStUpMsg(m => ({ ...m, [key]: "Could not read: " + (err?.message || "bad file") })); }
  setStUpBusy("");
};
const stApplyOutlet = async (oid: string, brand: string) => {
  const key = oid + "_" + brand;
  const e = stUpload[key]; if (!e) return;
  setStUpBusy(key);
  const { data: cur } = await supabase.from("sales_target").select("line_items").eq("outlet_id", oid).eq("brand", brand).single();
  const li: any = cur?.line_items || { sales: {}, fixed: {}, targets: {}, monthly: {} };
  const num = (a: any, b: any) => (a != null ? a : (b ?? 0));
  const mk = stDate.slice(0, 7);
  const isSharedFixedUp = brand !== "BH";
  const updated = {
    sales: li.sales || {},
    monthly: { ...(li.monthly || {}), [mk]: { net: e.net || 0, online: (e.swiggy || 0) + (e.zomato || 0) } },
    fixed: { staff: isSharedFixedUp ? 0 : num(e.staff, li.fixed?.staff), rent: isSharedFixedUp ? 0 : num(e.rent, li.fixed?.rent), eb: isSharedFixedUp ? 0 : num(e.eb, li.fixed?.eb), transport: isSharedFixedUp ? 0 : num(e.transport, li.fixed?.transport), pest: num(e.pest, li.fixed?.pest), water: num(e.water, li.fixed?.water), airtel: num(e.airtel, li.fixed?.airtel) },
    targets: li.targets || { a: 0, b: 0 },
  };
  const { error } = await supabase.from("sales_target").upsert({ outlet_id: oid, brand, line_items: updated, updated_at: new Date().toISOString() }, { onConflict: "outlet_id,brand" });
  setStUpBusy("");
  if (error) { setStUpMsg(m => ({ ...m, [key]: "Error: " + error.message })); return; }
  setStUpMsg(m => ({ ...m, [key]: "✓ Saved to Outlet P&L (" + brand + ")." }));
  setStUpload(u => ({ ...u, [key]: null }));
  if (user) fetchSalesTargets(user);
};

const reviewPoints = (rating: number, valid: boolean) => {
  let p = 0;
  if (rating === 5) p += 5; else if (rating === 4) p += 3; else if (rating >= 1 && rating <= 2) p -= 5;
  if (valid) p -= 10;
  return p;
};
const fetchReviews = async (outlet: string, date: string) => {
  if (!outlet) { setReviews([]); return; }
  const { data } = await supabase.from("outlet_reviews").select("*").eq("outlet_id", outlet).eq("report_date", date).order("created_at", { ascending: false });
  setReviews(data || []);
};
const saveReview = async () => {
  if (!user || !activeOutlet) return;
  const _owner = ALL_STAFF.find(s => (s.outlets as string[] | undefined)?.includes(activeOutlet));
  setRevSaving(true);
  const { error } = await supabase.from("outlet_reviews").insert({
    outlet_id: activeOutlet,
    staff_id: _owner ? _owner.id : user.id,
    report_date: outletEntryDate,
    platform: revForm.platform,
    rating: parseInt(revForm.rating) || null,
    valid_complaint: revForm.valid,
    refund_given: revForm.refund,
    note: revForm.note || null,
  });
  setRevSaving(false);
  if (error) { alert("Error: " + error.message); return; }
  setRevForm({ platform: "Swiggy", rating: "5", valid: false, refund: false, note: "" });
  fetchReviews(activeOutlet, outletEntryDate);
};
const deleteReview = async (id: string) => {
  await supabase.from("outlet_reviews").delete().eq("id", id);
  fetchReviews(activeOutlet, outletEntryDate);
};
const submitOutletReport = async () => {
  if (!user || !activeOutlet) return;
  const _brandSplit = outletEntryDate >= BRAND_SPLIT_FROM;
  const _req: { k: string; label: string }[] = _brandSplit
    ? ["bh", "cbh"].flatMap((b) => ["shop", "swiggy", "zomato"].flatMap((ch) => [
        { k: `${b}_${ch}_sales_count`, label: `${b.toUpperCase()} ${ch[0].toUpperCase() + ch.slice(1)} Orders Count (enter 0 if none)` },
        { k: `${b}_${ch}_sales_value`, label: `${b.toUpperCase()} ${ch[0].toUpperCase() + ch.slice(1)} Sales Value (enter 0 if none)` },
      ]))
    : [
    { k: "shop_sales_count", label: "Shop Sales Count" },
    { k: "shop_sales_value", label: "Shop Sales Value" },
    { k: "swiggy_sales_count", label: "Swiggy Sales Count" },
    { k: "swiggy_sales_value", label: "Swiggy Sales Value" },
    { k: "zomato_sales_count", label: "Zomato Sales Count" },
    { k: "zomato_sales_value", label: "Zomato Sales Value" },
  ];
  // Expiry/complimentary tracking becomes compulsory starting 3 Oct 2026 —
  // gated on the date the report is FOR (not the date it's being filed),
  // so reports being backfilled for earlier dates aren't blocked by a rule
  // that didn't exist yet when that day happened.
  const EXPIRY_FIELDS_REQUIRED_FROM = "2026-10-03";
  if (outletEntryDate >= EXPIRY_FIELDS_REQUIRED_FROM) {
    _req.push(
      { k: "expiry_items", label: "Expired Product Name (type 'None' if nothing expired)" },
      { k: "complimentary_items", label: "Complimentary Product Name (type 'None' if none given)" },
      { k: "complimentary_reason", label: "Reason (type 'None' if not applicable)" },
    );
  }
  const _miss = _req.filter((f) => !(outletReportData as any)[f.k] || !String((outletReportData as any)[f.k]).trim());
  if (_miss.length) { alert("Please fill all required fields before submitting.\n\nMissing: " + _miss.map((f) => f.label).join(", ")); return; }
  setOutletSubmitting(true);
 const deadline = new Date();
  deadline.setHours(12, 0, 0, 0);
  const _afterNoon = new Date() > deadline;
  const d = outletReportData;
  const newRating = parseFloat(d.bh_google_rating) || 0;
  const todayStr = new Date().toISOString().split("T")[0];
  const { data: prevRows } = await supabase.from("outlet_reports")
    .select("bh_google_rating, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value, swiggy_sales_count, zomato_sales_count").eq("outlet_id", activeOutlet).lt("report_date", todayStr)
    .order("report_date", { ascending: false }).limit(1);
  const prevRating = prevRows && prevRows[0] ? Number(prevRows[0].bh_google_rating) || 0 : 0;
  const earnedBonus = newRating > 4.5 && newRating > prevRating;
  const isBackfill = outletEntryDate < new Date().toISOString().split("T")[0];
  const isLate = !isBackfill && _afterNoon;
const _clean = (v: any) => String(v ?? "").replace(/[^0-9.]/g, "");
  // BH + CBH entered separately -> the combined columns (used by every other tab) are their sum, so nothing downstream breaks.
  const _bn = (k: string) => parseFloat(_clean((d as any)[k])) || 0;
  const _bi = (k: string) => parseInt(_clean((d as any)[k])) || 0;
  const _brandCols: Record<string, number> = {};
  if (_brandSplit) BRAND_SALES_KEYS.forEach((k) => { _brandCols[k] = k.endsWith("_count") ? _bi(k) : _bn(k); });
  const _sum = (ch: string, kind: "count" | "value") => (_brandCols[`bh_${ch}_sales_${kind}`] || 0) + (_brandCols[`cbh_${ch}_sales_${kind}`] || 0);
  const payload = {
    ..._brandCols,
    shop_sales_count: _brandSplit ? _sum("shop", "count") : parseInt(_clean(d.shop_sales_count)) || 0,
    shop_sales_value: _brandSplit ? _sum("shop", "value") : parseFloat(_clean(d.shop_sales_value)) || 0,
    swiggy_sales_count: _brandSplit ? _sum("swiggy", "count") : parseInt(_clean(d.swiggy_sales_count)) || 0,
    swiggy_sales_value: _brandSplit ? _sum("swiggy", "value") : parseFloat(_clean(d.swiggy_sales_value)) || 0,
    discount_given: parseFloat(_clean(d.discount_given)) || 0,
    zomato_sales_count: _brandSplit ? _sum("zomato", "count") : parseInt(_clean(d.zomato_sales_count)) || 0,
    zomato_sales_value: _brandSplit ? _sum("zomato", "value") : parseFloat(_clean(d.zomato_sales_value)) || 0,
    target: parseFloat(_clean(d.target)) || 0,
    swiggy_live: (d.swiggy_live || "yes").toLowerCase() === "yes",
    zomato_live: (d.zomato_live || "yes").toLowerCase() === "yes",
    discount_running: d.discount_running || "",
    discount_rate_good: (d.discount_rate_good || "yes").toLowerCase() === "yes",
    unavailable_items: d.unavailable_items || "",
    expiry_count: parseInt(d.expiry_count) || 0,
    expiry_items: d.expiry_items || "",
    complimentary_count: parseInt(d.complimentary_count) || 0,
    complimentary_items: d.complimentary_items || "",
    complimentary_reason: d.complimentary_reason || "",
    issues: d.issues || "",
    action_taken: d.action_taken || "",
    bh_google_rating: parseFloat(d.bh_google_rating) || null,
bh_swiggy_rating: parseFloat(d.bh_swiggy_rating) || null,
bh_zomato_rating: parseFloat(d.bh_zomato_rating) || null,
cbh_google_rating: parseFloat(d.cbh_google_rating) || null,
cbh_swiggy_rating: parseFloat(d.cbh_swiggy_rating) || null,
cbh_zomato_rating: parseFloat(d.cbh_zomato_rating) || null,
icbh_google_rating: parseFloat(d.icbh_google_rating) || null,
icbh_swiggy_rating: parseFloat(d.icbh_swiggy_rating) || null,
icbh_zomato_rating: parseFloat(d.icbh_zomato_rating) || null,
    is_late: isLate,
   is_edited: d.is_edited === "true",
    rating_bonus: isBackfill ? false : earnedBonus,
    is_backfill: isBackfill,
    no_points: isBackfill && outletWasOff,
  };
  // ── Entry limits (typo guard) ──────────────────────────────────────────
  // Each outlet has a daily maximum per channel (shop / swiggy / zomato) for
  // orders and for sales value, stored in the outlet_entry_limits table. A
  // number above the limit is almost always an extra-zero typo, so the report
  // is NOT saved until it is fixed. Only applies to report dates from
  // ENTRY_LIMITS_FROM. If the table is missing, empty, or can't be read,
  // nothing is blocked (fails open) so reports are never stuck by this check.
  if (outletEntryDate >= ENTRY_LIMITS_FROM) {
    const { data: _lim, error: _limErr } = await supabase.from("outlet_entry_limits").select("channel, max_orders, max_value").eq("outlet_id", activeOutlet);
    if (!_limErr && _lim && _lim.length) {
      const _over: string[] = [];
      _lim.forEach((l: any) => {
        const ch = String(l.channel);
        const cnt = Number((payload as any)[`${ch}_sales_count`]) || 0;
        const val = Number((payload as any)[`${ch}_sales_value`]) || 0;
        const nm = ch.charAt(0).toUpperCase() + ch.slice(1);
        if (l.max_orders != null && cnt > Number(l.max_orders)) _over.push(`${nm} orders: you entered ${cnt}, the limit is ${l.max_orders}`);
        if (l.max_value != null && val > Number(l.max_value)) _over.push(`${nm} sales value: you entered ₹${val.toLocaleString("en-IN")}, the limit is ₹${Number(l.max_value).toLocaleString("en-IN")}`);
      });
      if (_over.length) {
        setOutletSubmitting(false);
        alert("Report NOT submitted — some numbers look too high:\n\n" + _over.join("\n") + "\n\nPlease check for an extra zero and correct it. If the numbers are really correct, ask Niranjana or Nishant to raise this outlet's limit.");
        return;
      }
    }
  }
  let error;
  if (d.editing_id) {
    const result = await supabase.from("outlet_reports").update(payload).eq("id", d.editing_id);
    error = result.error;
  } else {
   const _owner = ALL_STAFF.find(s => (s.outlets as string[] | undefined)?.includes(activeOutlet));
    const result = await supabase.from("outlet_reports").upsert({
  ...payload,
  staff_id: _owner ? _owner.id : user.id,
  outlet_id: activeOutlet,
  report_date: outletEntryDate,
}, { onConflict: "staff_id,outlet_id,report_date" });
error = result.error;
  }
  setOutletSubmitting(false);
 if (error) { alert("Error: " + error.message); return; }
setOutletReportData({});
if (isBackfill) { if (!outletWasOff) celebrate(-30); }
else if (isLate) { celebrate(0, "After 12 PM cut-off — 0 points"); }
else {
  const _total = (Number(payload.shop_sales_value) || 0) + (Number(payload.swiggy_sales_value) || 0) + (Number(payload.zomato_sales_value) || 0);
  const _tgt = Number(payload.target) || 0;
  if (_tgt > 0 && _total < _tgt) celebrate(20, "Boo boo! Target missed — only +20 for submitting 😭");
  else if (_tgt > 0) celebrate(50);
  else celebrate(20);
}
await new Promise(resolve => setTimeout(resolve, 500));
const _today2 = new Date().toISOString().split("T")[0];
if (outletEntryDate === _today2) await fetchOutletReports(user);
else await fetchOutletReportsByDate(outletEntryDate);
};
  const playAlert = () => {
  try {
    const ctx = new (window.AudioContext || (window as any).webkitAudioContext)();
    const oscillator = ctx.createOscillator();
    const gainNode = ctx.createGain();
    oscillator.connect(gainNode);
    gainNode.connect(ctx.destination);
    oscillator.type = "square";
    oscillator.frequency.setValueAtTime(880, ctx.currentTime);
    oscillator.frequency.setValueAtTime(440, ctx.currentTime + 0.1);
    oscillator.frequency.setValueAtTime(880, ctx.currentTime + 0.2);
    gainNode.gain.setValueAtTime(0.3, ctx.currentTime);
    gainNode.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.4);
    oscillator.start(ctx.currentTime);
    oscillator.stop(ctx.currentTime + 0.4);
  } catch (e) {
    console.log("Audio not supported");
  }
};
  const submitForceAck = async (action: "complete" | "reason") => {
    if (!overdueTask) return;
    if (action === "reason") {
      if (forceAckReason.trim().length < 20) { alert("Please provide at least 20 characters explaining the delay."); return; }
      await supabase.from("tasks").update({ delay_reason: forceAckReason.trim(), status: "overdue" }).eq("id", overdueTask.id);
    } else {
      await supabase.from("tasks").update({ status: "completed", completed_at: new Date().toISOString() }).eq("id", overdueTask.id);
    }
    setOverdueTask(null);
    setForceAckReason("");
    if (user) fetchTasks(user);
  };

  const total = tasks.length;
  const completed = tasks.filter(t => t.status === "completed").length;
  const inProgress = tasks.filter(t => t.status === "in_progress").length;
  const overdue = tasks.filter(t => t.status !== "completed" && new Date(t.due_at) < new Date()).length;
  const rate = total > 0 ? Math.round(completed / total * 100) : 0;
  const canAssign = user?.role === "Owner" || user?.role === "Manager";
  const hasOutlets = (user?.outlets?.length || 0) > 0;
  const isFO = user?.role === "Founder's Office";
  const isOwner = user?.role === "Owner";
  const isArun = user?.id === "arun";
  // Outlet Targets — Arun-only screen to change an outlet's monthly target
  // from a chosen month onward, without rewriting past months. Each change
  // is INSERTED as a new row (never updated/overwritten), so the full
  // history stays intact; monthlyTargetFor/dailyTargetFor above pick
  // whichever change is the latest one that applies to the month being
  // looked at. this is the same mechanism that already correctly separates
  // August (24L) from September (26L) for Royapettah — just made editable
  // from the app instead of a code change.
  const [otRows, setOtRows] = useState<any[]>([]);
  const [otOutlet, setOtOutlet] = useState(OUTLETS[0]);
  const [otMonth, setOtMonth] = useState(() => new Date().toISOString().slice(0, 7));
  const [otValue, setOtValue] = useState("");
  const [otSaving, setOtSaving] = useState(false);
  const fetchOutletTargetUpdates = async () => {
    const { data } = await supabase.from("outlet_target_updates").select("*").order("effective_from", { ascending: true });
    const rows = data || [];
    setOtRows(rows);
    RUNTIME_TARGET_UPDATES = rows.map((r: any) => ({ from: r.effective_from, monthly: { [r.outlet_id]: Number(r.monthly_target) } }));
  };
  useEffect(() => { fetchOutletTargetUpdates(); /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, []);
  const saveOutletTargetUpdate = async () => {
    const val = Number(String(otValue).replace(/,/g, "")) || 0;
    const curYm = new Date().toISOString().slice(0, 7);
    if (!val) { alert("Enter the new monthly target."); return; }
    if (otMonth < curYm) { alert("Can't backdate a target change to a month that's already closed — pick the current month or a future one."); return; }
    setOtSaving(true);
    const { error } = await supabase.from("outlet_target_updates").insert({ outlet_id: otOutlet, effective_from: otMonth, monthly_target: val, created_by: user?.id || null });
    setOtSaving(false);
    if (error) { alert("Failed: " + error.message); return; }
    setOtValue("");
    fetchOutletTargetUpdates();
  };
  const canUploadItemPerf = true;
  const canViewItemPerf = canAssign || canUploadItemPerf;
  const hasReportDuty = user?.role !== "Owner" && user?.role !== "Founder's Office" && user?.role !== "Head Chef" && user?.role !== "Financial Analyst" && user?.role !== "Outlet Manager";
  const [kChefs, setKChefs] = useState<any[]>([]);
  const [kRows, setKRows] = useState<any[]>([]);
  const [kDate, setKDate] = useState(new Date().toISOString().slice(0, 10));
  const [kNewChef, setKNewChef] = useState("");
  // Simple daily submission, organised exactly the way Rafiq actually works:
  // half kg cakes, 1kg cakes, pastries, corporate orders, hotel orders — each
  // with its own flavour/description + count, so he just fills in whatever
  // he made that day and submits once. No chef picking required here (that's
  // a separate, optional roster below) — this is meant to be fast on a phone.
  const K_CATEGORIES = ["Half kg Cake", "1 kg Cake", "Pastries", "Corporate Order", "Hotel Order"] as const;
  // 5 separate boxes, one per fixed category. Inside each non-order box,
  // every product already on the master list (kitchen_products) is printed
  // as its own row with just a count field — exactly like Rafiq's paper
  // sheet — so there's nothing to pick, only numbers to fill in. "+ Add
  // another flavour" below that is only for something NOT already printed
  // (a one-off, not yet on the master list); "+ Add new product" actually
  // adds it to the master list so it's pre-printed here from then on.
  const [kMasterQty, setKMasterQty] = useState<Record<string, string>>({}); // productId -> qty
  type KCatRow = { flavour: string; qty: string };
  const emptyKCatRows = (): Record<string, KCatRow[]> => Object.fromEntries(K_CATEGORIES.map((c) => [c, [] as KCatRow[]]));
  const [kCatRows, setKCatRows] = useState<Record<string, KCatRow[]>>(emptyKCatRows());
  const addKCatRow = (cat: string) => setKCatRows((prev) => ({ ...prev, [cat]: [...(prev[cat] || []), { flavour: "", qty: "" }] }));
  const removeKCatRow = (cat: string, idx: number) => setKCatRows((prev) => ({ ...prev, [cat]: (prev[cat] || []).filter((_, i) => i !== idx) }));
  const updateKCatRow = (cat: string, idx: number, field: keyof KCatRow, value: string) => setKCatRows((prev) => ({ ...prev, [cat]: (prev[cat] || []).map((r, i) => (i === idx ? { ...r, [field]: value } : r)) }));
  // Inline "add a new product to the master list" form, one open at a time.
  const [kAddProductCat, setKAddProductCat] = useState<string | null>(null);
  const [kNewProductName, setKNewProductName] = useState("");
  const [kAddingProduct, setKAddingProduct] = useState(false);
  const [kDailySubmitting, setKDailySubmitting] = useState(false);
  const [kChefEditId, setKChefEditId] = useState<string | null>(null);
  const [kChefEditName, setKChefEditName] = useState("");
  const [kProducts, setKProducts] = useState<any[]>([]);
  const [kTgtEdits, setKTgtEdits] = useState<Record<string, string>>({});
  const [kShowTargets, setKShowTargets] = useState(false);
  const fetchKProducts = async () => { const { data } = await supabase.from("kitchen_products").select("*").eq("active", true).order("sort_order"); setKProducts(data || []); };
  // Adds a brand-new flavour to the master list, permanently, under the
  // given category — from then on it's pre-printed in that category's box
  // instead of needing "+ Add another flavour" every single day.
  const addKProduct = async (cat: string) => {
    const name = kNewProductName.trim();
    if (!name) return;
    setKAddingProduct(true);
    const maxSort = kProducts.reduce((m, pr) => Math.max(m, Number(pr.sort_order) || 0), 0);
    const { error } = await supabase.from("kitchen_products").insert({ name, category: cat, active: true, sort_order: maxSort + 1, target_qty: 0 });
    setKAddingProduct(false);
    if (error) { alert("Failed: " + error.message); return; }
    setKNewProductName("");
    setKAddProductCat(null);
    fetchKProducts();
  };
  const saveKTargets = async () => { for (const pr of kProducts) { const v = kTgtEdits[pr.id]; if (v !== undefined && v !== String(pr.target_qty ?? "")) { await supabase.from("kitchen_products").update({ target_qty: Number(v) || 0 }).eq("id", pr.id); } } setKTgtEdits({}); fetchKProducts(); };
  const kProdVsTarget = kProducts.map((pr) => { const made = kRows.filter((r) => (r.flavour || "").trim().toLowerCase() === (pr.name || "").trim().toLowerCase()).reduce((sm, r) => sm + (Number(r.qty) || 0), 0); return { name: pr.name as string, category: (pr.category || "Other") as string, made, target: (pr.target_qty || 0) as number }; });
  const kCats = Array.from(new Set(kProducts.map((pr) => pr.category || "Other")));
  // Flavour suggestions for the entry form, scoped to each box's own
  // category — so the "Half kg Cake" box only suggests half-kg flavours
  // (in the same order as the master product list), not every product
  // mixed together. Falls back to showing everything if a category has
  // no matching products yet, so the field never ends up empty.
  const kProductsForCat = (cat: string) => {
    const norm = (s: string) => (s || "").trim().toLowerCase();
    const matched = kProducts.filter((pr) => norm(pr.category) === norm(cat));
    return matched.length > 0 ? matched : kProducts;
  };
  const kMasterNames = new Set(kProducts.map((pr) => (pr.name || "").trim().toLowerCase()));
  const kCustom = kRows.filter((r) => !kMasterNames.has((r.flavour || "").trim().toLowerCase()));
  const fetchKChefs = async () => { const { data } = await supabase.from("kitchen_chefs").select("*").eq("active", true).order("name"); setKChefs(data || []); };
  const fetchKProduction = async (d: string) => { const { data } = await supabase.from("kitchen_production").select("*").eq("prod_date", d).order("created_at"); setKRows(data || []); };
  // Daily Production Plan (Production ask #1/#2: a forecast-driven plan per
  // category, compared against what's actually logged that day). Forecasts
  // at the CATEGORY level (the "station" field Rafiq's form already saves —
  // Half kg Cake / 1kg Cake / Pastries / Corporate / Hotel), not per flavour,
  // because flavours are free-typed text and won't match day to day. Same
  // day-of-week-average + trend method used everywhere else in Forecasting.
  const [kPlanLoading, setKPlanLoading] = useState(false);
  const [kPlan, setKPlan] = useState<Record<string, { planned: number; hasHistory: boolean }>>({});
  const fetchKPlan = async (dateStr: string) => {
    setKPlanLoading(true);
    const target = new Date(dateStr + "T00:00:00Z");
    const dow = target.getUTCDay();
    const historyStart = new Date(target); historyStart.setUTCDate(historyStart.getUTCDate() - 70);
    const { data: rows } = await supabase.from("kitchen_production").select("prod_date, station, qty").gte("prod_date", historyStart.toISOString().slice(0, 10)).lt("prod_date", dateStr).in("station", K_CATEGORIES as unknown as string[]);
    const dailyTotals: Record<string, Record<string, number>> = {};
    K_CATEGORIES.forEach((c) => { dailyTotals[c] = {}; });
    (rows || []).forEach((r: any) => {
      const cat = r.station as string;
      if (!dailyTotals[cat]) return;
      dailyTotals[cat][r.prod_date] = (dailyTotals[cat][r.prod_date] || 0) + (Number(r.qty) || 0);
    });
    const result: Record<string, { planned: number; hasHistory: boolean }> = {};
    K_CATEGORIES.forEach((cat) => {
      const dates = Object.keys(dailyTotals[cat]).sort();
      const sameDow = dates.filter((d) => new Date(d + "T00:00:00Z").getUTCDay() === dow);
      const hasHistory = sameDow.length >= 2;
      const avg = sameDow.length ? sameDow.reduce((s, d) => s + dailyTotals[cat][d], 0) / sameDow.length : 0;
      const last14 = dates.slice(-14).reduce((s, d) => s + dailyTotals[cat][d], 0);
      const prev14 = dates.slice(-28, -14).reduce((s, d) => s + dailyTotals[cat][d], 0);
      let trend = prev14 > 0 ? last14 / prev14 : 1;
      trend = Math.max(0.7, Math.min(1.3, trend));
      result[cat] = { planned: Math.round(avg * trend), hasHistory };
    });
    setKPlan(result);
    setKPlanLoading(false);
  };
  const kActualByCat: Record<string, number> = Object.fromEntries(K_CATEGORIES.map((c) => [c, kRows.filter((r) => r.station === c).reduce((s, r) => s + (Number(r.qty) || 0), 0)]));
  useEffect(() => { if (user?.role === "Head Chef" || user?.role === "Owner" || isFO) { fetchKChefs(); fetchKProduction(kDate); fetchKProducts(); fetchKPlan(kDate); } /* eslint-disable-next-line react-hooks/exhaustive-deps */ }, [user]);
  const addKChef = async () => { if (!kNewChef.trim()) return; const { error } = await supabase.from("kitchen_chefs").insert({ name: kNewChef.trim() }); if (error) { alert("Failed: " + error.message); return; } setKNewChef(""); fetchKChefs(); };
  const saveKChefEdit = async () => {
    if (!kChefEditId || !kChefEditName.trim()) { setKChefEditId(null); return; }
    const { error } = await supabase.from("kitchen_chefs").update({ name: kChefEditName.trim() }).eq("id", kChefEditId);
    if (error) { alert("Failed: " + error.message); return; }
    setKChefEditId(null); setKChefEditName(""); fetchKChefs();
  };
  // Soft delete (active = false) rather than removing the row — past
  // production entries already carry the chef's name directly, so this
  // doesn't touch history, it just takes the name off the list going
  // forward (and out of new entries).
  const deleteKChef = async (id: string, name: string) => {
    if (!confirm(`Remove ${name} from the chef list?`)) return;
    const { error } = await supabase.from("kitchen_chefs").update({ active: false }).eq("id", id);
    if (error) { alert("Failed: " + error.message); return; }
    fetchKChefs();
  };
  // One submit for the whole day, across all 5 categories — matches how
  // Rafiq actually works: fill in whatever was made, submit once.
  const submitKDailyProduction = async () => {
    const toInsert: any[] = [];
    // Pre-printed master-list rows — one per product with a count typed in.
    kProducts.forEach((pr) => {
      const qty = Number(kMasterQty[pr.id]) || 0;
      if (qty <= 0) return;
      toInsert.push({ prod_date: kDate, chef_id: null, chef_name: null, flavour: pr.name, qty, station: pr.category || "Other", entered_by: user?.id || null });
    });
    // Extra ad-hoc rows (not on the master list yet) + order-category rows.
    K_CATEGORIES.forEach((cat) => {
      (kCatRows[cat] || []).forEach((r) => {
        const qty = Number(r.qty) || 0;
        const flavour = r.flavour.trim();
        if (!flavour && qty <= 0) return; // skip genuinely empty rows
        toInsert.push({ prod_date: kDate, chef_id: null, chef_name: null, flavour: flavour || null, qty, station: cat, entered_by: user?.id || null });
      });
    });
    if (toInsert.length === 0) { alert("Fill in at least one item before submitting."); return; }
    setKDailySubmitting(true);
    const { error } = await supabase.from("kitchen_production").insert(toInsert);
    setKDailySubmitting(false);
    if (error) { alert("Save failed: " + error.message); return; }
    setKMasterQty({});
    setKCatRows(emptyKCatRows());
    fetchKProduction(kDate);
  };
  const delKProd = async (id: string) => { await supabase.from("kitchen_production").delete().eq("id", id); fetchKProduction(kDate); };
  const kByChef = kChefs.map((c) => ({ name: c.name, total: kRows.filter((r) => r.chef_id === c.id).reduce((sm, r) => sm + (Number(r.qty) || 0), 0) })).sort((a, b) => b.total - a.total);
  const kDayTotal = kRows.reduce((sm, r) => sm + (Number(r.qty) || 0), 0);
  const kMax = Math.max(1, ...kByChef.map((c) => c.total));
  const kAvg = kByChef.length ? Math.round(kDayTotal / kByChef.length) : 0;
  const reportFields = user ? REPORT_FIELDS[user.id] || [] : [];
  const reportInput = (f: { label: string; key: string }) => {
    if (user && user.id === "arun" && f.key === "achievement") {
      const ts = parseFloat(reportData.total_sales || "0");
      const tg = parseFloat(reportData.target || "299666");
      const pct = tg ? (ts / tg * 100).toFixed(1) + "%" : "";
      return <input type="text" readOnly value={pct} className="w-full bg-zinc-900 border border-zinc-800 text-yellow-400 px-3 py-2 text-sm cursor-not-allowed" placeholder="—" />;
    }
    const val = (user && user.id === "arun" && f.key === "target") ? (reportData.target ?? "299666") : (reportData[f.key] || "");
    return <input type="text" value={val} onChange={(e) => setReportData(prev => ({ ...prev, [f.key]: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="—" />;
  };

  if (!user) return (
    <div className="min-h-screen bg-black flex items-center justify-center">
      <p className="text-zinc-500 font-mono text-sm uppercase tracking-widest">Loading...</p>
    </div>
  );

  return (
    <div className="min-h-screen bg-canvas text-text flex">
      <ThemePicker />
      {/* Account bar — always visible top-right, no need to open/scroll the
          sidebar just to see who's logged in or to change PIN / exit. */}
      <div className="fixed top-20 right-4 z-50 flex items-center gap-1.5 bg-surface border border-line rounded-full pl-1 pr-2 py-1 shadow-lg">
        <div className="w-7 h-7 bg-accent text-accent-ink flex items-center justify-center font-bold text-[11px] rounded-full shrink-0">
          {user.name.split(" ").map((n: string) => n[0]).join("").slice(0, 2)}
        </div>
        <button onClick={() => setShowPinModal(true)} className="text-[10px] font-mono text-text-faint uppercase hover:text-accent transition-colors px-1.5">PIN</button>
        <button onClick={() => { localStorage.removeItem("currentUser"); router.push("/"); }} className="text-[10px] font-mono text-text-faint uppercase hover:text-red-500 transition-colors px-1.5">Exit</button>
      </div>
      <BrownieLoader tick={brownieTick} />
       {(user?.role === "Owner" || (user as any)?.isFO) && <ActivityToastStack />}
      {user && <PushRegister staffId={user.id} />}
      {isFO && <NudgeToast />}
      {sidebarOpen && <div className="fixed inset-0 bg-black/50 z-30 md:hidden" onClick={() => setSidebarOpen(false)} />}
      {targetCheck && targetCheck.length > 0 && !targetReaction && (() => {
        const wins = targetCheck.filter((r: any) => r.status === "win").length;
        const misses = targetCheck.filter((r: any) => r.status === "miss").length;
        const header = misses > 0 ? "Yesterday had some gaps — let's go today! 🔥" : (wins > 0 ? "Yesterday was a win! 👏🎉" : "Fresh start today 🚀");
        const fmt = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
        return (
          <div className="fixed inset-0 bg-black/70 z-[60] flex items-center justify-center p-4" onClick={() => setTargetCheck(null)}>
            <div className="bg-[#131316] border border-yellow-400 max-w-md w-full p-6" onClick={(e) => e.stopPropagation()}>
              <h3 className="text-xl font-black mb-4">{header}</h3>
              <div className="space-y-2 mb-5">
                {targetCheck.map((r: any) => (
                  <div key={r.oid} className="text-sm font-mono border-b border-zinc-800 pb-2">
                    {r.status === "win" && <span className="text-green-400">🎉 {r.name}: {fmt(r.actual)} vs {fmt(r.target)} — smashed it!</span>}
                    {r.status === "miss" && <span className="text-yellow-400">💪 {r.name}: {fmt(r.actual)} vs {fmt(r.target)} — bit short, today's the bounce-back!</span>}
                    {r.status === "noentry" && <span className="text-zinc-500">🤔 {r.name}: no entry logged yesterday</span>}
                    {r.status === "notarget" && <span className="text-blue-400">🚀 {r.name}: new outlet — no target yet, every sale counts!</span>}
                  </div>
                ))}
              </div>
             <button onClick={() => setTargetReaction(true)} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-5 py-2.5 uppercase w-full">Next →</button>
            </div>
          </div>
      );
      })()}
      {targetCheck && targetReaction && (() => {
        const misses = targetCheck.filter((r: any) => r.status === "miss").length;
        const wins = targetCheck.filter((r: any) => r.status === "win").length;
        const bad = misses > 0;
        const neutral = wins === 0 && misses === 0;
        const close = () => { setTargetCheck(null); setTargetReaction(false); };
        return (
          <div className="fixed inset-0 bg-black/80 z-[60] flex items-center justify-center p-4" onClick={close}>
            <style>{`@keyframes bhShake{0%,100%{transform:translateX(0) rotate(0)}20%{transform:translateX(-10px) rotate(-6deg)}40%{transform:translateX(10px) rotate(6deg)}60%{transform:translateX(-8px) rotate(-4deg)}80%{transform:translateX(8px) rotate(4deg)}}@keyframes bhPop{0%{transform:scale(0)}60%{transform:scale(1.3)}100%{transform:scale(1)}}`}</style>
            <div className={`max-w-sm w-full p-8 text-center border-2 ${bad ? "bg-[#1a1010] border-red-500" : (neutral ? "bg-[#101418] border-blue-400" : "bg-[#101a12] border-green-400")}`} onClick={(e) => e.stopPropagation()}>
              <div className="text-8xl mb-3" style={{ animation: bad ? "bhShake 0.4s ease-in-out 3" : "bhPop 0.5s ease-out", display: "inline-block" }}>{bad ? "🤚💥" : (neutral ? "🚀" : "👏")}</div>
              {!bad && !neutral && <div className="text-3xl mb-3">🎉 🎊 🎉</div>}
              <h3 className="text-2xl font-black mb-2">{bad ? "SLAP! 🖐️😵‍💫 Target said NO" : (neutral ? "Fresh start, let's roll! 🚀" : "BOOM! You SMASHED it! 👏")}</h3>
              <p className="text-sm font-mono text-zinc-400 mb-6">{bad ? "That's a wake-up smack 😤 Shake it off — today you hit back twice as hard!" : (neutral ? "No target pressure — every sale's a bonus." : "Target crushed. Keep that energy going today!")}</p>
              <button onClick={close} className={`${bad ? "bg-red-500" : (neutral ? "bg-blue-400" : "bg-green-400")} text-black font-bold tracking-widest text-xs px-6 py-3 uppercase w-full`}>Let's go →</button>
            </div>
          </div>
        );
      })()}
      <button onClick={() => setSidebarOpen(!sidebarOpen)} className="fixed top-4 left-4 z-50 md:hidden bg-zinc-900 border border-zinc-700 p-2 text-white">☰</button>

      <aside className={`fixed inset-y-0 left-0 z-40 w-60 bg-surface border-r border-line flex flex-col shrink-0 transition-transform duration-200 md:relative md:translate-x-0 ${sidebarOpen ? "translate-x-0" : "-translate-x-full"}`}>
        <div className="px-6 py-6 border-b border-line">
          <h1 className="text-xl font-black tracking-tight">TASK<span className="text-accent">FORCE IQ</span></h1>
          <p className="text-[10px] font-mono text-text-faint uppercase tracking-widest mt-1">Brownie Heaven</p>
        </div>
        <nav className="flex-1 min-h-0 overflow-y-auto overscroll-contain px-3 py-4" style={{ WebkitOverflowScrolling: "touch", touchAction: "pan-y" }}>
          <p className="text-[10px] font-mono text-text-faint uppercase tracking-widest px-3 pb-2">Workspace</p>
          <div onClick={() => { fireBrownieTransition(); setActiveTab("tasks"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "tasks" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
            <span>▣</span> Dashboard
          </div>
          {!isOwner && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("messages"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "messages" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>💬</span> Messages
              {msgUnreadCount > 0 && <span className="ml-auto w-2 h-2 bg-accent rounded-full"></span>}
            </div>
          )}
          {(canAssign || hasOutlets) && (
              <div onClick={() => { fireBrownieTransition(); setActiveTab("owner_outlets"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "owner_outlets" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🏪</span> Outlet Reports
              </div>
          )}
         {((user.outlets && user.outlets.length > 0) || canAssign) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("sales_target"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "sales_target" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🎯</span> Outlet P&amp;L
            </div>
          )}
                  {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("purchase_vendors"); setSidebarOpen(false); fetchPurchaseVendors(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "purchase_vendors" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🛒</span> Purchase &amp; Vendors
            </div>
          )}
          {(isFO || ["nishant","arun","vishnu","ahila","nilani"].includes(user?.id ?? "")) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("auto_reviews"); setSidebarOpen(false); fetchAutoReviews(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "auto_reviews" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>⭐</span> Auto Reviews
            </div>
          )}
        <div onClick={() => { router.push("/leaderboard"); }} className="flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors text-text-muted hover:text-text">
            <span>🏆</span> Leaderboard
          </div>
         {["nishant","arun","nilani","vishnu","ahila"].includes(user?.id ?? "") && (
            <div onClick={() => { router.push("/orders-race"); }} className="flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors text-text-muted hover:text-text">
              <span>⚔️</span> Orders Race
            </div>
          )}
         {(canAssign || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("analytics"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "analytics" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
             <span>◬</span> Analytics
            </div>
          )}
          {false && (canAssign || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("competition"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "competition" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🥊</span> Competition
            </div>
          )}
          {canViewItemPerf && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("item_perf"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "item_perf" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📈</span> Item Performance
            </div>
          )}
          {(isOwner || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("admin_dq"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "admin_dq" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🛡</span> Admin &amp; Data Quality
            </div>
          )}
                    {canAssign && (
            <>
             <div onClick={() => { fireBrownieTransition(); setActiveTab("all_reports"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "all_reports" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
                <span>📋</span> Reports
              </div>
            </>
          )}
         {((user.outlets && user.outlets.length > 0) || canAssign) && (
  <div onClick={() => { fireBrownieTransition(); setActiveTab("outlet_reports"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "outlet_reports" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
    <span>🏪</span> Outlets
  {Object.keys(outletReports).length < (user.outlets?.length || 0) && <span className="ml-auto w-2 h-2 bg-accent rounded-full"></span>}
  </div>
)}
          {canAssign && (
            <>
              <div onClick={() => { fireBrownieTransition(); setActiveTab("history"); setSidebarOpen(false); fetchHistoryReports(historyDate); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "history" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
             <span>📅</span> History
             </div>
            </>
          )}
                    {isOwner && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("team_dashboards"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "team_dashboards" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>👥</span> Team Dashboards
            </div>
          )}
          {(isOwner || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("active_status"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "active_status" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🟢</span> Active Status
            </div>
          )}
          {isOwner && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("messages"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "messages" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>💬</span> Messages
              {msgUnreadCount > 0 && <span className="ml-auto w-2 h-2 bg-accent rounded-full"></span>}
            </div>
          )}
                  {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("net_realisation"); setSidebarOpen(false); fetchNetRealisation(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "net_realisation" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🧾</span> Net Realisation
            </div>
          )}
                {user?.role === "Financial Analyst" && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("pnl"); setSidebarOpen(false); fetchPnl(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "pnl" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>💹</span> Channel P&amp;L
            </div>
          )}
                   {user?.role === "Financial Analyst" && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("contribution_margins"); setSidebarOpen(false); fetchContributionMargins(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "contribution_margins" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📐</span> Contribution Margins
            </div>
          )}
          {(user?.role === "Financial Analyst" || user?.role === "Owner" || user?.role === "Founder's Office") && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("cash_flow"); setSidebarOpen(false); fetchCashFlowForecast(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "cash_flow" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📉</span> Profit Forecast
            </div>
          )}
          {(user?.role === "Financial Analyst" || user?.role === "Owner" || user?.role === "Founder's Office") && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("sales_forecast"); setSidebarOpen(false); fetchSalesForecast(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "sales_forecast" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📈</span> Sales Forecast
            </div>
          )}
                {user?.role === "Financial Analyst" && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("payout"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "payout" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>💰</span> Payout
            </div>
          )}
                    {false && user?.role === "Owner" && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("reconciliation"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "reconciliation" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>⚖️</span> Reconciliation
            </div>
          )}
          {hasReportDuty && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("my_report"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "my_report" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📋</span> My Report
             {!todayReport && <span className="ml-auto w-2 h-2 bg-accent rounded-full"></span>}
            </div>
          )}
          {(user.role === "HR" || user?.role === "Owner") && (
                <div onClick={() => { fireBrownieTransition(); setActiveTab("attendance"); setSidebarOpen(false); const staffRef = (user?.role === "Owner" ? { id: "nilani" } : user) as Staff; fetchAttendance(staffRef, attendanceDate); fetchAttendanceWeekOvertime(staffRef); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "attendance" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>👥</span> People &amp; Labour
            </div>
          )}
          {(user?.role === "Head Chef" || user?.role === "Owner" || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("production"); setSidebarOpen(false); fetchKChefs(); fetchKProduction(kDate); fetchKProducts(); fetchKPlan(kDate); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "production" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🏭</span> Production
            </div>
          )}
          {(canAssign || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("ops_audits"); setSidebarOpen(false); fetchAuditReports(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "ops_audits" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🔍</span> Operations &amp; Audits
            </div>
          )}
          {isArun && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("outlet_targets"); setSidebarOpen(false); fetchOutletTargetUpdates(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "outlet_targets" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>🎯</span> Outlet Targets
            </div>
          )}
                            {(canAssign || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("fines"); setSidebarOpen(false); fetchFines(); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "fines" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>⚖️</span> Fines
            </div>
          )}
                   {(isOwner || isFO) && (
            <div onClick={() => { fireBrownieTransition(); setActiveTab("niranjana_report"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "niranjana_report" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
              <span>📝</span> Niranjana's Report
            </div>
          )}
          <div onClick={() => { fireBrownieTransition(); setActiveTab("help"); setSidebarOpen(false); }} className={`flex items-center gap-3 px-3 py-2.5 text-sm font-medium cursor-pointer transition-colors ${activeTab === "help" ? "text-text bg-surface-2 border-l-2 border-accent" : "text-text-muted hover:text-text"}`}>
            <span>🛟</span> Help
          </div>
        </nav>
      </aside>

      <main className="flex-1 px-4 py-4 md:px-8 md:py-8 overflow-auto">
        {ceoPopupOpen && (isOwner || isFO) && (
          <div style={{ position: "fixed", inset: 0, zIndex: 50, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem" }}>
            <div className="bg-zinc-900 border border-zinc-700 max-w-md w-full p-6 relative">
              <button onClick={() => setCeoPopupOpen(false)} className="absolute top-3 right-3 text-zinc-500 hover:text-white text-lg">✕</button>
              <p className="text-2xl font-black mb-1">{(() => { const h = new Date().getHours(); return h < 12 ? "Good morning ☕" : h < 17 ? "Good afternoon 🌤️" : "Good evening 🌙"; })()}</p>
              <p className="text-sm text-zinc-400 mb-4">{user?.name?.split(" ")[0]}, here's your Brownie Heaven morning brief.</p>
              <p className="text-sm text-zinc-300 mb-5">A quick look at who's on top of their game, who's slipping, and how the month is tracking.</p>
              <button onClick={() => { setCeoPopupOpen(false); setActiveTab("ceo_report"); fetchCeoData(ceoWin); }} className="w-full bg-yellow-400 text-black px-4 py-2.5 text-sm font-semibold hover:bg-yellow-300 transition-colors">Wanna go through today's report? →</button>
              <button onClick={() => setCeoPopupOpen(false)} className="w-full mt-2 text-zinc-500 hover:text-white text-sm py-2">Maybe later — take me to the dashboard</button>
            </div>
          </div>
        )}

        {fineAckOpen && unackedFines.length > 0 && (
          <div style={{ position: "fixed", inset: 0, zIndex: 60, background: "rgba(0,0,0,0.75)", display: "flex", alignItems: "center", justifyContent: "center", padding: "1rem" }}>
            <div className="bg-zinc-900 border border-red-500 max-w-md w-full p-6">
              <p className="text-2xl font-black mb-1 text-red-400">⚠️ You've been fined</p>
              <p className="text-sm text-zinc-400 mb-4">{unackedFines.length === 1 ? "This needs your acknowledgment before you continue." : `${unackedFines.length} fines need your acknowledgment before you continue.`}</p>
              <div className="space-y-2 mb-5 max-h-64 overflow-y-auto">
                {unackedFines.map((f) => (
                  <div key={f.id} className="border border-zinc-800 p-3">
                    <div className="flex justify-between items-baseline">
                      <span className="text-sm font-semibold">{f.reason || "No reason given"}</span>
                      <span className="font-mono text-red-400">−₹{Number(f.amount).toLocaleString("en-IN")}</span>
                    </div>
                    <div className="text-xs text-zinc-500 mt-1">{f.fine_date}{f.outlet ? ` · ${OUTLET_NAMES[f.outlet] || f.outlet}` : ""}</div>
                  </div>
                ))}
              </div>
              <button onClick={acknowledgeFines} disabled={fineAckBusy} className="w-full bg-red-500 text-white px-4 py-2.5 text-sm font-semibold hover:bg-red-400 disabled:opacity-50 transition-colors">{fineAckBusy ? "…" : "I understand"}</button>
            </div>
          </div>
        )}

       {activeTab === "tasks" && (isOwner || isFO) && (
         <div className="mb-4 flex justify-end">
           <button onClick={() => { fireBrownieTransition(); setActiveTab("ceo_report"); fetchCeoData(ceoWin); }} className="text-[11px] font-mono uppercase tracking-widest text-yellow-400 hover:text-yellow-300 border border-yellow-400/40 px-3 py-2 transition-colors">
             📊 CEO Report →
           </button>
         </div>
       )}
       {activeTab === "tasks" && user && user.role === "Founder's Office" && <FounderDashboard user={user} />}
       {activeTab === "tasks" && user && user.role === "Owner" && <CommandCentre user={user} />}
       {activeTab === "tasks" && user && (user.role === "Asst. Ops Manager" || user.role === "Custom Cakes & Asst Ops" || user.role === "Outlet Manager") && <div className="mb-8"><MyOutletsDashboard user={user} /></div>}
        {((activeTab === "tasks" && user?.role === "Head Chef") || (activeTab === "production" && (user?.role === "Head Chef" || user?.role === "Owner" || isFO))) && (
          <div>
            <div className="flex justify-between items-end mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl font-black tracking-tight">Kitchen Operations</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Daily production & workload · {user.name.split(" ")[0]}</p>
              </div>
              <input type="date" value={kDate} onChange={(e) => { setKDate(e.target.value); fetchKProduction(e.target.value); fetchKPlan(e.target.value); }} className="bg-black border border-zinc-800 text-white px-4 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors font-mono text-sm" />
            </div>

            {/* Dashboard-only: the entry form lives in the Production tab, so here
                Rafiq gets a clear pointer + today's filing status instead of a
                second copy of the same form. */}
            {activeTab === "tasks" && user?.role === "Head Chef" && (() => {
              const istToday = new Date(Date.now() + 5.5 * 60 * 60 * 1000).toISOString().slice(0, 10);
              const viewingToday = kDate === istToday;
              const filed = viewingToday && kRows.length > 0;
              return (
                <div className={`mb-8 border p-4 sm:p-5 ${filed ? "border-green-500/40 bg-green-500/5" : "border-yellow-400/50 bg-yellow-400/5"}`}>
                  <p className={`text-[11px] font-mono uppercase tracking-widest mb-1.5 ${filed ? "text-green-400" : "text-yellow-400"}`}>
                    {viewingToday ? (filed ? `✓ Today's production filed — ${kDayTotal} units` : "⚠ Today's production not filed yet") : `Viewing ${kDate}`}
                  </p>
                  <p className="text-sm text-zinc-200 mb-3">
                    Rafiq, fill in your daily production from the <span className="font-semibold text-white">Production</span> tab in the left menu — that's where the count boxes are. This Dashboard just shows the summary.
                  </p>
                  <button onClick={() => { fireBrownieTransition(); setActiveTab("production"); setSidebarOpen(false); }} className="bg-yellow-400 text-black px-5 py-2.5 text-sm font-bold hover:bg-yellow-300 transition-colors">
                    Go to Production →
                  </button>
                </div>
              );
            })()}

            {activeTab === "production" && user?.role === "Head Chef" && (
            <div className="mb-8 border border-zinc-800 p-4 sm:p-5">
              <p className="text-sm font-semibold mb-1">Today's Production</p>
              <p className="text-xs text-zinc-500 mb-4">Every known flavour is already listed under its category — just type the count next to whatever was made. Use "+ Add another flavour" only for something not listed yet, or "+ Add new product" to put it on the list permanently.</p>
              <div className="space-y-6">
                {K_CATEGORIES.map((cat) => {
                  const isOrder = cat === "Corporate Order" || cat === "Hotel Order";
                  const catProducts = kProductsForCat(cat);
                  return (
                    <div key={cat}>
                      <p className="text-[11px] font-mono uppercase tracking-widest text-yellow-400 mb-2">{cat}</p>

                      {/* Pre-printed master-list rows — just a count field each, no picking. */}
                      {!isOrder && (
                        <div className="space-y-1.5 mb-2">
                          {catProducts.length === 0 && <p className="text-xs text-zinc-600">No products on the list yet — add one below.</p>}
                          {catProducts.map((pr) => (
                            <div key={pr.id} className="flex items-center gap-2">
                              <span className="flex-1 text-sm text-zinc-200">{pr.name}</span>
                              <input
                                type="number"
                                value={kMasterQty[pr.id] || ""}
                                onChange={(e) => setKMasterQty((prev) => ({ ...prev, [pr.id]: e.target.value }))}
                                placeholder="0"
                                className="w-20 bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm shrink-0"
                              />
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Extra rows — ad-hoc flavours (non-order) or free-text orders. */}
                      {(kCatRows[cat] || []).length > 0 && (
                        <div className="space-y-2 mb-2">
                          {(kCatRows[cat] || []).map((row, idx) => (
                            <div key={idx} className="flex flex-wrap gap-2 items-center">
                              <input
                                type="text"
                                value={row.flavour}
                                onChange={(e) => updateKCatRow(cat, idx, "flavour", e.target.value)}
                                placeholder={isOrder ? "Description (e.g. Taj Hotel order)" : "Flavour name"}
                                className="flex-1 min-w-[140px] bg-black border border-yellow-400 text-white px-3 py-2.5 focus:outline-none transition-colors text-sm"
                              />
                              <input
                                type="number"
                                value={row.qty}
                                onChange={(e) => updateKCatRow(cat, idx, "qty", e.target.value)}
                                placeholder="Count"
                                className="w-20 bg-black border border-zinc-800 text-white px-3 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors text-sm shrink-0"
                              />
                              <button onClick={() => removeKCatRow(cat, idx)} className="text-zinc-500 hover:text-red-500 px-2 py-2.5 border border-zinc-800 hover:border-red-500 transition-colors shrink-0">✕</button>
                            </div>
                          ))}
                        </div>
                      )}

                      {/* Inline "add new product" form. */}
                      {kAddProductCat === cat && !isOrder && (
                        <div className="flex flex-wrap gap-2 items-center mb-2 bg-black/40 p-2 border border-yellow-400/40">
                          <input
                            type="text"
                            value={kNewProductName}
                            onChange={(e) => setKNewProductName(e.target.value)}
                            placeholder="New product name"
                            autoFocus
                            className="flex-1 min-w-[140px] bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm"
                          />
                          <button onClick={() => addKProduct(cat)} disabled={kAddingProduct || !kNewProductName.trim()} className="text-[11px] font-mono uppercase bg-yellow-400 text-black px-3 py-2 hover:bg-yellow-300 disabled:opacity-50 transition-colors shrink-0">{kAddingProduct ? "Adding…" : "Add to list"}</button>
                          <button onClick={() => { setKAddProductCat(null); setKNewProductName(""); }} className="text-[11px] font-mono uppercase text-zinc-500 hover:text-white px-2 shrink-0">Cancel</button>
                        </div>
                      )}

                      <div className="flex gap-4">
                        <button onClick={() => addKCatRow(cat)} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-yellow-400 transition-colors">+ Add another {isOrder ? "order" : "flavour"}</button>
                        {!isOrder && kAddProductCat !== cat && (
                          <button onClick={() => { setKAddProductCat(cat); setKNewProductName(""); }} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-yellow-400 transition-colors">+ Add new product</button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
              <button onClick={submitKDailyProduction} disabled={kDailySubmitting} className="mt-5 block w-full sm:w-auto bg-yellow-400 text-black px-6 py-3 text-sm font-bold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{kDailySubmitting ? "Submitting…" : "Submit Today's Production"}</button>
            </div>
            )}

            <div className="mb-8 border border-zinc-800 p-4 sm:p-5 max-w-3xl">
              <p className="text-sm font-semibold mb-1">Daily Production Plan · {kDate}</p>
              <p className="text-[11px] text-zinc-500 mb-4">Planned = the typical amount made on this day of the week, adjusted for the recent trend. Needs about 2 weeks of logged history per category before it can show a number — until then it just shows what's been made so far.</p>
              {kPlanLoading ? <p className="text-sm text-zinc-500">Loading…</p> : (
                <div className="space-y-3">
                  {K_CATEGORIES.map((cat) => {
                    const plan = kPlan[cat];
                    const actual = kActualByCat[cat] || 0;
                    if (!plan || !plan.hasHistory) {
                      return (
                        <div key={cat} className="flex justify-between text-sm text-zinc-500">
                          <span>{cat}</span>
                          <span className="font-mono text-xs">Not enough history yet · made so far: {actual}</span>
                        </div>
                      );
                    }
                    const pct = plan.planned > 0 ? Math.round((actual / plan.planned) * 100) : 0;
                    const col = pct >= 90 ? "bg-green-500" : pct >= 50 ? "bg-yellow-400" : "bg-red-500";
                    return (
                      <div key={cat}>
                        <div className="flex justify-between text-sm mb-0.5">
                          <span>{cat}</span>
                          <span className="font-mono text-zinc-300">{actual} / ~{plan.planned} planned{plan.planned > 0 ? ` · ${pct}%` : ""}</span>
                        </div>
                        <div className="h-2 bg-zinc-800"><div className={`h-full ${col}`} style={{ width: `${Math.min(100, pct)}%` }} /></div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 max-w-5xl">
              <div className="border border-zinc-800 p-4">
                <div className="flex justify-between items-baseline mb-3">
                  <p className="text-sm font-semibold">Workload balance</p>
                  <p className="text-[11px] font-mono text-zinc-500">Day total: {kDayTotal} · avg {kAvg}/chef</p>
                </div>
                {kByChef.length === 0 ? <p className="text-sm text-zinc-600">No chefs yet.</p> : (
                  <div className="space-y-2">
                    {kByChef.map((c) => { const over = kAvg > 0 && c.total > kAvg * 1.5; const under = kAvg > 0 && c.total > 0 && c.total < kAvg * 0.5; const w = Math.round(c.total / kMax * 100); return (
                      <div key={c.name}>
                        <div className="flex justify-between text-sm mb-0.5"><span>{c.name}</span><span className={`font-mono ${over ? "text-red-400" : "text-zinc-300"}`}>{c.total}{over ? " ⚠" : ""}</span></div>
                        <div className="h-2 bg-zinc-800"><div className={`h-full ${over ? "bg-red-500" : under ? "bg-zinc-600" : "bg-yellow-400"}`} style={{ width: `${w}%` }} /></div>
                      </div>
                    ); })}
                  </div>
                )}
                <p className="text-[10px] text-zinc-600 mt-3">⚠ = more than 1.5× the day's average — consider rebalancing.</p>
              </div>

              <div className="border border-zinc-800 p-4">
                <p className="text-sm font-semibold mb-3">Today's log · {kDate}</p>
                {kRows.length === 0 ? <p className="text-sm text-zinc-600">Nothing logged yet.</p> : (
                  <div className="space-y-3">
                    {[...K_CATEGORIES, ...Array.from(new Set(kRows.map((r) => r.station).filter((s) => s && !(K_CATEGORIES as readonly string[]).includes(s))))].map((cat) => {
                      const rows = kRows.filter((r) => r.station === cat);
                      if (rows.length === 0) return null;
                      const catTotal = rows.reduce((s, r) => s + (Number(r.qty) || 0), 0);
                      return (
                        <div key={cat}>
                          <div className="flex justify-between items-baseline mb-1"><p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{cat}</p><span className="text-[10px] font-mono text-zinc-500">{catTotal}</span></div>
                          <div className="space-y-1">
                            {rows.map((r) => (
                              <div key={r.id} className="flex items-baseline gap-2 text-sm border-b border-zinc-900 pb-1">
                                <span className="flex-1 text-zinc-300">{r.flavour || "—"}{r.chef_name ? ` · ${r.chef_name}` : ""}</span>
                                <span className="font-mono text-yellow-400">{r.qty}</span>
                                <button onClick={() => delKProd(r.id)} className="text-[10px] font-mono uppercase px-2 py-0.5 border border-zinc-700 hover:border-red-500 hover:text-red-500 transition-colors">✕</button>
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            </div>

            <div className="mt-6 border border-zinc-800 p-4 max-w-3xl">
              <div className="flex justify-between items-baseline mb-3">
                <p className="text-sm font-semibold">Production vs target · {kDate}</p>
                <button onClick={() => setKShowTargets(v => !v)} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-white">{kShowTargets ? "Done" : "Set targets"}</button>
              </div>
              {kProducts.length === 0 ? <p className="text-sm text-zinc-600">No products yet.</p> : (
                <div className="space-y-4">
                  {kCats.map((cat) => { const items = kProdVsTarget.filter((pr) => pr.category === cat); const catMade = items.reduce((sm, pr) => sm + pr.made, 0); const catTgt = items.reduce((sm, pr) => sm + pr.target, 0); return (
                    <div key={cat}>
                      <div className="flex justify-between items-baseline mb-1.5"><p className="text-[11px] font-mono text-zinc-400 uppercase tracking-widest">{cat}</p><span className="font-mono text-xs text-zinc-500">{catMade}{catTgt > 0 ? ` / ${catTgt}` : ""}</span></div>
                      <div className="space-y-1.5">
                        {items.map((pr) => { const pct = pr.target > 0 ? Math.round(pr.made / pr.target * 100) : 0; const col = pr.target === 0 ? "bg-zinc-600" : pct >= 100 ? "bg-green-500" : pct >= 60 ? "bg-yellow-400" : "bg-red-500"; return (
                          <div key={pr.name}>
                            <div className="flex justify-between text-sm mb-0.5"><span>{pr.name}</span><span className="font-mono text-zinc-300">{pr.made}{pr.target > 0 ? ` / ${pr.target} · ${pct}%` : ""}</span></div>
                            <div className="h-2 bg-zinc-800"><div className={`h-full ${col}`} style={{ width: `${pr.target > 0 ? Math.min(100, pct) : (pr.made > 0 ? 100 : 0)}%` }} /></div>
                          </div>
                        ); })}
                      </div>
                    </div>
                  ); })}
                  {kCustom.length > 0 && (
                    <div>
                      <p className="text-[11px] font-mono text-orange-400 uppercase tracking-widest mb-1.5">Customised / special orders</p>
                      <div className="space-y-1">
                        {kCustom.map((r) => (
                          <div key={r.id} className="flex justify-between text-sm"><span>{r.flavour || "—"}{r.chef_name ? ` · ${r.chef_name}` : ""}</span><span className="font-mono text-orange-300">{r.qty}</span></div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
              )}
              {kShowTargets && (
                <div className="mt-4 pt-4 border-t border-zinc-800">
                  <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Daily target per product</p>
                  <div className="space-y-2">
                    {kProducts.map((pr) => (
                      <div key={pr.id} className="flex items-center gap-2 text-sm">
                        <span className="flex-1">{pr.name}</span>
                        <input type="number" value={kTgtEdits[pr.id] ?? String(pr.target_qty ?? 0)} onChange={(e) => setKTgtEdits(m => ({ ...m, [pr.id]: e.target.value }))} className="w-24 bg-black border border-zinc-800 text-white px-2 py-1 focus:outline-none focus:border-yellow-400 text-sm" />
                      </div>
                    ))}
                  </div>
                  <button onClick={saveKTargets} className="mt-3 bg-yellow-400 text-black px-4 py-1.5 text-sm font-semibold hover:bg-yellow-300 transition-colors">Save targets</button>
                </div>
              )}
            </div>

            <div className="mt-6 border border-zinc-800 p-4 max-w-md">
              <p className="text-sm font-semibold mb-3">Chefs</p>
              {kChefs.length === 0 ? (
                <p className="text-xs text-zinc-500 mb-3">None yet.</p>
              ) : (
                <div className="space-y-1.5 mb-4">
                  {kChefs.map((c) => (
                    <div key={c.id} className="flex items-center gap-2">
                      {kChefEditId === c.id ? (
                        <>
                          <input type="text" value={kChefEditName} onChange={(e) => setKChefEditName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && saveKChefEdit()} className="flex-1 bg-black border border-yellow-400 text-white px-2 py-1.5 focus:outline-none text-sm" autoFocus />
                          <button onClick={saveKChefEdit} className="text-[11px] font-mono uppercase text-green-400 hover:text-green-300 px-2">Save</button>
                          <button onClick={() => { setKChefEditId(null); setKChefEditName(""); }} className="text-[11px] font-mono uppercase text-zinc-500 hover:text-white px-2">Cancel</button>
                        </>
                      ) : (
                        <>
                          <span className="flex-1 text-sm">{c.name}</span>
                          <button onClick={() => { setKChefEditId(c.id); setKChefEditName(c.name); }} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-yellow-400 px-2">Edit</button>
                          <button onClick={() => deleteKChef(c.id, c.name)} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-red-500 px-2">Delete</button>
                        </>
                      )}
                    </div>
                  ))}
                </div>
              )}
              <div className="flex gap-2">
                <input type="text" value={kNewChef} onChange={(e) => setKNewChef(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addKChef()} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm flex-1" placeholder="Add a chef (e.g. Riyas)" />
                <button onClick={addKChef} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors shrink-0">Add</button>
              </div>
            </div>
          </div>
       )}
                    {activeTab === "tasks" && user?.role === "Financial Analyst" && (
          <div className="flex flex-col items-center justify-center py-24 text-center max-w-lg mx-auto">
            <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-[0.3em] mb-3">{user.role}</p>
            <h2 className="text-4xl md:text-5xl font-black tracking-tight mb-3">Welcome, {user.name.split(" ")[0]}</h2>
            <p className="text-sm text-zinc-500 mb-8">Use the sidebar — Channel P&amp;L and Contribution Margins are live. Still building:</p>
            <div className="text-left w-full space-y-3 text-sm text-zinc-400 border border-zinc-800 p-5">
              <p>Weekly cash-flow forecasting</p>
              <p>Swiggy/Zomato net-realisation analysis</p>
            </div>
          </div>
       )}
       {activeTab === "pnl" && user?.role === "Financial Analyst" && (
          <div>
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">Channel P&amp;L</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Real fixed costs from Outlet P&amp;L · {(cogsRate * 100).toFixed(1)}% COGS (real, {fmtWin(realFoodCostWindow)}) · 5% wastage · {(commissionRate * 100).toFixed(1)}% online commission (real, {fmtWin(realCommissionWindow)}) · both reset to the 1st of the month (COGS is month-to-date; commission also includes last month, since payouts settle slowly)</p>
              </div>
              <div className="flex gap-2">
                <input type="date" value={pnlFrom} onChange={(e) => setPnlFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                <input type="date" value={pnlTo} onChange={(e) => setPnlTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                <button onClick={downloadPnlPDF} disabled={pnlPdfBusy} className="bg-yellow-400 text-black font-bold text-xs px-4 py-2 uppercase tracking-widest disabled:opacity-50 hover:opacity-90 transition-opacity">{pnlPdfBusy ? "Generating…" : "Download Report"}</button>
              </div>
            </div>
            {pnlLoading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : pnlRows.length === 0 ? (
              <p className="text-sm text-zinc-500">No sales data for this range.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                      <th className="py-2 pr-3">Outlet</th>
                      <th className="py-2 pr-3 text-right">Shop</th>
                      <th className="py-2 pr-3 text-right">Swiggy</th>
                      <th className="py-2 pr-3 text-right">Zomato</th>
                      <th className="py-2 pr-3 text-right">Total Sales</th>
                      <th className="py-2 pr-3 text-right">Contribution</th>
                      <th className="py-2 pr-3 text-right">Fixed Costs</th>
                      <th className="py-2 pr-3 text-right">Net Profit</th>
                      <th className="py-2 text-right">Net %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {pnlRows.map((r) => (
                      <Fragment key={r.oid}>
                        <tr className="border-b border-zinc-900 cursor-pointer hover:bg-zinc-900" onClick={() => setPnlExpanded(pnlExpanded === r.oid ? null : r.oid)}>
                          <td className="py-2 pr-3 font-semibold">{r.name} <span className="text-zinc-600 text-xs">{pnlExpanded === r.oid ? "▲" : "▼"}</span></td>
                          <td className="py-2 pr-3 text-right font-mono text-xs text-zinc-400">₹{Math.round(r.shop.sales).toLocaleString("en-IN")}<br /><span className="text-zinc-600">{r.shop.margin.toFixed(0)}% margin</span></td>
                          <td className="py-2 pr-3 text-right font-mono text-xs text-zinc-400">₹{Math.round(r.swiggy.sales).toLocaleString("en-IN")}<br /><span className="text-zinc-600">{r.swiggy.margin.toFixed(0)}% margin</span></td>
                          <td className="py-2 pr-3 text-right font-mono text-xs text-zinc-400">₹{Math.round(r.zomato.sales).toLocaleString("en-IN")}<br /><span className="text-zinc-600">{r.zomato.margin.toFixed(0)}% margin</span></td>
                          <td className="py-2 pr-3 text-right font-mono font-semibold">₹{Math.round(r.totalSales).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.totalContrib).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.fixed).toLocaleString("en-IN")}</td>
                          <td className={`py-2 pr-3 text-right font-mono font-bold ${r.netProfit >= 0 ? "text-green-400" : "text-red-500"}`}>₹{Math.round(r.netProfit).toLocaleString("en-IN")}</td>
                          <td className={`py-2 text-right font-mono font-bold ${r.netMargin >= 0 ? "text-green-400" : "text-red-500"}`}>{r.netMargin.toFixed(1)}%</td>
                        </tr>
                        {pnlExpanded === r.oid && (
                          <tr className="bg-zinc-900/40">
                            <td colSpan={9} className="py-3 px-3">
                              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Fixed costs — {r.name} (from Outlet P&amp;L)</p>
                              <div className="grid grid-cols-2 md:grid-cols-4 gap-2 text-xs font-mono">
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Staff</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.staff).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Rent</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.rent).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Electricity</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.eb).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Transport</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.transport).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">R&amp;M (20% rent)</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.rm).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Pest control</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.pest).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Water</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.water).toLocaleString("en-IN")}</span></div>
                                <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-500">Airtel/WiFi</span><span className="text-zinc-300">₹{Math.round(r.fixedBreakdown.airtel).toLocaleString("en-IN")}</span></div>
                              </div>
                            </td>
                          </tr>
                        )}
                      </Fragment>
                    ))}
                  </tbody>
                </table>
                <p className="text-[10px] text-zinc-600 mt-3">Fixed costs pulled from what's entered in Outlet P&amp;L for each outlet (BH brand). Outlets with no fixed-cost entry yet show ₹0 there — worth flagging to whoever owns that outlet.</p>
              </div>
            )}
          </div>
       )}
       {activeTab === "net_realisation" && (user?.role === "Financial Analyst" || user?.role === "Owner") && (
          <div>
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">Net Realisation</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">What Swiggy/Zomato actually pay vs the order value</p>
              </div>
              <div className="flex gap-2 items-center">
                <input type="date" value={nrFrom} onChange={(e) => setNrFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                <input type="date" value={nrTo} onChange={(e) => setNrTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                <button
                  onClick={async () => {
                    const { count } = await supabase.from("outlet_payouts").select("id", { count: "exact", head: true });
                    if (!count) { alert("Nothing to clear — outlet_payouts is already empty."); return; }
                    if (!confirm(`This will permanently delete all ${count} payout row(s) — Swiggy and Zomato, every outlet, every period. Use this to reset while testing. Continue?`)) return;
                    const { error } = await supabase.from("outlet_payouts").delete().not("id", "is", null);
                    if (error) { alert("Clear failed: " + error.message); return; }
                    fetchNetRealisation();
                  }}
                  className="text-xs text-zinc-600 hover:text-red-400 border border-zinc-800 hover:border-red-500/40 px-3 py-2 transition-colors"
                  title="Delete all payout data — for resetting while testing"
                >
                  🗑 Clear test data
                </button>
              </div>
            </div>

            {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
              <div className="mb-6 bg-[#131316] border border-zinc-800 p-4">
                <div className="flex justify-between items-center mb-1">
                  <p className="text-sm font-semibold">📢 Approved discount % — per outlet, per week</p>
                  <input type="date" value={discWeekStart} onChange={(e) => setDiscWeekStart(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
                </div>
                <p className="text-[10px] text-zinc-600 mb-3">What discount % each outlet is actually authorized to run on Swiggy this week — the number to compare against what Swiggy really applies, once real discount data is wired in. Week starting {discWeekStart}.</p>
                <div className="overflow-x-auto">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-zinc-500 uppercase tracking-widest text-[10px] border-b border-zinc-800">
                        <th className="text-left py-2 pr-3">Outlet</th>
                        <th className="text-center py-2 px-2">BH</th>
                        <th className="text-center py-2 px-2">CBH</th>
                        <th className="text-center py-2 px-2">ICBH</th>
                      </tr>
                    </thead>
                    <tbody>
                      {OUTLETS.map((oid) => (
                        <tr key={oid} className="border-b border-zinc-800/40">
                          <td className="py-1.5 pr-3 text-zinc-300">{OUTLET_NAMES[oid] || oid}</td>
                          {(["BH", "CBH", "ICBH"] as const).map((brand) => {
                            const key = `${oid}_${brand}`;
                            const val = discTargets[key];
                            return (
                              <td key={brand} className="py-1.5 px-2 text-center">
                                <input
                                  type="number"
                                  defaultValue={val ?? ""}
                                  key={`${key}_${val ?? "empty"}`}
                                  placeholder="—"
                                  onBlur={(e) => {
                                    const n = parseFloat(e.target.value);
                                    if (!isNaN(n) && n !== val) saveDiscountTarget(oid, brand, n);
                                  }}
                                  className={`w-16 bg-black border text-center px-1 py-1 text-xs focus:outline-none focus:border-yellow-400 ${discSaving[key] ? "border-yellow-400" : "border-zinc-800"}`}
                                />
                                <span className="text-zinc-600">%</span>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <p className="text-[10px] text-zinc-600 mt-3">Excel-based real-discount comparison still waiting on a confirmed sample. The live check below is a working alternative — reads Swiggy's actual public page for each outlet right now.</p>
              </div>
            )}

            {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
              <div className="mb-6 bg-[#131316] border border-zinc-800 p-4">
                <p className="text-sm font-semibold mb-1">📋 Manual check list — all outlet Swiggy pages</p>
                <p className="text-[10px] text-zinc-600 mb-4">Click, glance at the offer banner, click the next one. Opens each real Swiggy page in a new tab — same pages the automated checker above couldn't get past.</p>
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                  {OUTLETS.map((oid) => (
                    <div key={oid} className="border border-zinc-800 p-3">
                      <p className="text-xs font-bold text-zinc-200 mb-2">{OUTLET_NAMES[oid] || oid}</p>
                      <div className="flex flex-wrap gap-2">
                        {(["BH", "CBH", "ICBH"] as const).map((brand) => {
                          const built = buildSwiggyUrl(oid, brand);
                          if (!built) return <span key={brand} className="text-[10px] text-zinc-700 border border-zinc-800 px-2 py-1">{brand}: no ID</span>;
                          return (
                            <a key={brand} href={built.url} target="_blank" rel="noopener noreferrer" className="text-[10px] text-yellow-400 hover:text-yellow-300 border border-zinc-800 hover:border-yellow-400/50 px-2 py-1 transition-colors">
                              {brand} ↗
                            </a>
                          );
                        })}
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
              <div className="mb-6 flex items-center gap-3 relative">
                <button onClick={checkNewPayouts} disabled={autoPayoutChecking} className="bg-yellow-400 text-black px-4 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">
                  {autoPayoutChecking ? "Checking…" : "🔄 Check Swiggy inbox"}
                </button>
                {autoPayoutResult && (
                  autoPayoutResult.success ? (
                    <span className="text-xs text-green-400">✅ Saved {autoPayoutResult.inserted.length}{autoPayoutResult.skipped.length > 0 ? ` · ${autoPayoutResult.skipped.length} not matched yet` : ""}</span>
                  ) : (
                    <span className="text-xs text-red-400">⚠️ {autoPayoutResult.error}</span>
                  )
                )}
                {autoPayoutResult?.success && (autoPayoutResult.sampleSkips?.length > 0 || autoPayoutResult.skipReasons) && (
                  <details className="text-xs text-zinc-600">
                    <summary className="cursor-pointer hover:text-zinc-400">Advanced</summary>
                    <div className="absolute mt-2 bg-neutral-900 border border-zinc-800 p-3 z-10 max-w-2xl max-h-[70vh] overflow-y-auto">
                      {autoPayoutResult.timedOut && <p className="text-yellow-400 mb-1">⏱️ Ran out of time this batch — ~{autoPayoutResult.remaining} still unchecked. Check again.</p>}
                      {autoPayoutResult.skipReasons && Object.keys(autoPayoutResult.skipReasons).length > 0 && (
                        <ul className="text-zinc-400 space-y-0.5 mb-2">
                          {Object.entries(autoPayoutResult.skipReasons).map(([reason, count]: any) => <li key={reason}>• {count} skipped — {reason}</li>)}
                        </ul>
                      )}
                      {autoPayoutResult.sampleSkips?.map((s: any, i: number) => {
                        const preview = s.parsed || (s.fullText ? parseSwiggyPayoutEmail(s.fullText) : null);
                        return (
                        <div key={i} className="bg-black/40 p-2 mb-2 text-[11px] text-zinc-500">
                          <p className="text-zinc-300">{s.reason}</p>
                          {s.restId && <p>Rest ID: {s.restId}</p>}
                          {s.subject && <p>Subject: {s.subject}</p>}
                          {preview && (
                            <div className="mt-2 grid grid-cols-2 gap-1.5">
                              <div className="bg-[#131316] border border-zinc-800 p-1.5">
                                <p className="text-[9px] font-mono text-zinc-600 uppercase">Period</p>
                                <p className="text-zinc-200">{preview.period_start || "?"} → {preview.period_end || "?"}</p>
                              </div>
                              <div className="bg-[#131316] border border-zinc-800 p-1.5">
                                <p className="text-[9px] font-mono text-zinc-600 uppercase">Orders</p>
                                <p className="text-zinc-200">{preview.total_orders ?? "—"}</p>
                              </div>
                              <div className="bg-[#131316] border border-zinc-800 p-1.5">
                                <p className="text-[9px] font-mono text-zinc-600 uppercase">Order Value</p>
                                <p className="text-zinc-200">{preview.customer_payable ? `₹${Number(preview.customer_payable).toLocaleString("en-IN")}` : "—"}</p>
                              </div>
                              <div className="bg-[#131316] border border-zinc-800 p-1.5">
                                <p className="text-[9px] font-mono text-zinc-600 uppercase">Actually Paid</p>
                                <p className="text-zinc-200">{preview.amount_transferable ? `₹${Number(preview.amount_transferable).toLocaleString("en-IN")}` : "—"}</p>
                              </div>
                            </div>
                          )}
                          {s.fullText && (
                            <details className="mt-1">
                              <summary className="text-[10px] text-zinc-600 cursor-pointer hover:text-zinc-400">Show raw email (if the numbers above look wrong)</summary>
                              <pre className="text-zinc-500 mt-1 whitespace-pre-wrap font-mono text-[10px] max-h-48 overflow-y-auto border border-zinc-800 p-2">{s.fullText}</pre>
                            </details>
                          )}
                          {(s.reason === "no Rest ID found in email" || s.reason === "Rest ID not in outlet map yet") && (
                            <div className="mt-2 pt-2 border-t border-zinc-800 flex flex-wrap items-center gap-2">
                              {manualAssignDone[i] ? (
                                <span className="text-green-400">✓ Saved</span>
                              ) : (
                                <>
                                  <select value={manualAssignPick[i]?.outlet || ""} onChange={(e) => setManualAssignPick((m) => ({ ...m, [i]: { outlet: e.target.value, brand: m[i]?.brand || "" } }))} className="bg-black border border-zinc-800 text-white px-2 py-1 text-[10px]">
                                    <option value="">Outlet…</option>
                                    {OUTLETS.map((o) => <option key={o} value={o}>{OUTLET_NAMES[o] || o}</option>)}
                                  </select>
                                  <select value={manualAssignPick[i]?.brand || ""} onChange={(e) => setManualAssignPick((m) => ({ ...m, [i]: { outlet: m[i]?.outlet || "", brand: e.target.value } }))} className="bg-black border border-zinc-800 text-white px-2 py-1 text-[10px]">
                                    <option value="">Brand…</option>
                                    <option value="BH">BH</option>
                                    <option value="CBH">CBH</option>
                                    <option value="ICBH">ICBH</option>
                                  </select>
                                  <button onClick={() => saveManualPayout(s, i)} disabled={manualAssignBusy[i]} className="bg-yellow-400 text-black px-2 py-1 text-[10px] font-semibold hover:bg-yellow-300 disabled:opacity-50">
                                    {manualAssignBusy[i] ? "Saving…" : "Save to this box"}
                                  </button>
                                </>
                              )}
                            </div>
                          )}
                        </div>
                        );
                      })}
                    </div>
                  </details>
                )}
              </div>
            )}

            {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
              <div className="mb-8 grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                {OUTLETS.map((oid) => {
                  const renderPlatformRow = (platform: "Swiggy" | "Zomato") => (
                    <div>
                      <p className="text-[9px] font-mono text-zinc-600 uppercase tracking-widest mb-1.5">{platform}</p>
                      <div className="grid grid-cols-3 gap-2">
                        {(["BH", "CBH", "ICBH"] as const).map((brand) => {
                          const matches = nrRows.filter((r: any) => r.outletId === oid && r.brand === brand && r.platform === platform);
                          const latest = matches.length ? matches.reduce((a: any, b: any) => (a.periodEnd > b.periodEnd ? a : b)) : null;
                          if (!latest) {
                            return (
                              <div key={brand} className="bg-black/30 border border-dashed border-zinc-800 p-2 text-center">
                                <p className="text-[9px] font-mono text-zinc-600 uppercase tracking-widest mb-1">{brand}</p>
                                <p className="text-[10px] text-zinc-700">No data yet</p>
                              </div>
                            );
                          }
                          const pctColor = latest.pct == null ? "text-zinc-600" : latest.pct < 50 ? "text-red-500" : latest.pct < 65 ? "text-yellow-400" : "text-green-400";
                          return (
                            <div key={brand} className={`bg-[#131316] border p-2 ${platform === "Swiggy" ? "border-orange-500/20" : "border-red-500/20"}`}>
                              <p className={`text-[9px] font-mono uppercase tracking-widest mb-1 ${platform === "Swiggy" ? "text-orange-400" : "text-red-400"}`}>{brand}</p>
                              <p className="text-[9px] text-zinc-600 mb-1">{latest.periodStart.slice(5)} → {latest.periodEnd.slice(5)}</p>
                              <p className="text-xs font-mono text-zinc-300">₹{Math.round(latest.gross).toLocaleString("en-IN")}</p>
                              <p className="text-xs font-mono text-white font-bold">₹{Math.round(latest.net).toLocaleString("en-IN")}</p>
                              <p className={`text-xs font-mono font-bold ${pctColor}`}>{latest.pct == null ? "—" : latest.pct.toFixed(0) + "%"}</p>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                  return (
                    <div key={oid} className="border border-zinc-800 p-4 space-y-3">
                      <p className="text-sm font-bold uppercase tracking-widest">{OUTLET_NAMES[oid] || oid}</p>
                      {renderPlatformRow("Swiggy")}
                      {renderPlatformRow("Zomato")}
                    </div>
                  );
                })}
              </div>
            )}

            {(user?.role === "Financial Analyst" || user?.role === "Owner") && (
              <div className="mb-8 border border-zinc-800 p-5 max-w-2xl">
                <p className="text-sm font-semibold mb-1">📥 Upload Swiggy payout report</p>
                <p className="text-xs text-zinc-500 mb-4">Manual alternative to the auto-checker above — the weekly/monthly annexure Excel Swiggy emails. Reads the "Payout Breakup" sheet, pulls Total Customer Paid (gross) and Net Payout (net) from the Delivered Orders column.</p>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                  <div>
                    <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Outlet</label>
                    <select value={spOutlet} onChange={(e) => setSpOutlet(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1">
                      <option value="">— which outlet is this for —</option>
                      {OUTLETS.map((o) => <option key={o} value={o}>{OUTLET_NAMES[o] || o}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Brand</label>
                    <select value={spBrand} onChange={(e) => setSpBrand(e.target.value as any)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1">
                      <option value="BH">Brownie Heaven (BH)</option>
                      <option value="CBH">Cakes by Brownie Heaven (CBH)</option>
                      <option value="ICBH">Ice Cream by Brownie Heaven (ICBH)</option>
                    </select>
                  </div>
                  <div className="flex gap-2">
                    <div className="flex-1">
                      <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period start</label>
                      <input type="date" value={spPeriodFrom} onChange={(e) => setSpPeriodFrom(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" />
                    </div>
                    <div className="flex-1">
                      <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period end</label>
                      <input type="date" value={spPeriodTo} onChange={(e) => setSpPeriodTo(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" />
                    </div>
                  </div>
                </div>
                <label className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 cursor-pointer transition-colors inline-block">Choose file<input type="file" accept=".xlsx,.xls" onChange={onSpFile} className="hidden" /></label>
                {spFile && <span className="text-xs text-zinc-500 ml-3">{spFile.name}</span>}
                {spMsg && <p className="text-xs text-yellow-400 mt-2">{spMsg}</p>}
                {spParsed && (
                  <div className="mt-4 bg-black/30 border border-zinc-800 p-3">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Found — check before saving</p>
                    <div className="grid grid-cols-2 gap-2 text-sm mb-3">
                      <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-400">Total Customer Paid (gross)</span><span className="font-mono text-green-400">₹{spParsed.gross.toLocaleString("en-IN")}</span></div>
                      <div className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-400">Net Payout (net)</span><span className="font-mono text-green-400">₹{spParsed.net.toLocaleString("en-IN")}</span></div>
                    </div>
                    <p className="text-xs text-zinc-500 mb-3">Realisation: {((spParsed.net / spParsed.gross) * 100).toFixed(1)}%</p>
                    <button onClick={saveSwiggyPayout} disabled={spBusy} className="bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{spBusy ? "Saving…" : "Save to Net Realisation"}</button>
                  </div>
                )}
              </div>
            )}

            {nrLoading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : nrRows.length === 0 ? (
              <p className="text-sm text-zinc-500">No payouts entered yet for this range — add them in the Payout tab first.</p>
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                      <th className="py-2 pr-3">Outlet</th>
                      <th className="py-2 pr-3">Brand</th>
                      <th className="py-2 pr-3">Platform</th>
                      <th className="py-2 pr-3">Period</th>
                      <th className="py-2 pr-3 text-right">Order Value</th>
                      <th className="py-2 pr-3 text-right">Actually Paid</th>
                      <th className="py-2 text-right">Realisation %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {nrRows.map((r, i) => (
                      <tr key={i} className="border-b border-zinc-900">
                        <td className="py-2 pr-3 font-semibold">{r.outlet}</td>
                        <td className="py-2 pr-3 font-mono text-xs text-zinc-400">{r.brand}</td>
                        <td className="py-2 pr-3">
                          <span className={`font-mono text-[10px] uppercase px-1.5 py-0.5 ${r.platform === "Swiggy" ? "bg-orange-500/10 text-orange-400" : "bg-red-500/10 text-red-400"}`}>{r.platform}</span>
                        </td>
                        <td className="py-2 pr-3 font-mono text-xs text-zinc-400">{r.periodStart} → {r.periodEnd}</td>
                        <td className="py-2 pr-3 text-right font-mono">₹{Math.round(r.gross).toLocaleString("en-IN")}</td>
                        <td className="py-2 pr-3 text-right font-mono">₹{Math.round(r.net).toLocaleString("en-IN")}</td>
                        <td className={`py-2 text-right font-mono font-bold ${r.pct == null ? "text-zinc-600" : r.pct < 50 ? "text-red-500" : r.pct < 65 ? "text-yellow-400" : "text-green-400"}`}>{r.pct == null ? "—" : r.pct.toFixed(1) + "%"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <p className="text-[10px] text-zinc-600 mt-3">Swiggy: both figures come straight from Swiggy's own statement, fully verified. Zomato: "Order Value" is the amount staff reported in Outlet Reports (Zomato's statement doesn't include a gross figure), so Zomato's % is directionally useful but not platform-verified like Swiggy's.</p>
              </div>
            )}
          </div>
       )}
        {activeTab === "purchase_vendors" && (user?.role === "Financial Analyst" || user?.role === "Owner") && (
          <div>
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">Purchase &amp; Vendors</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Company-wide — from Production Stock's purchase records</p>
              </div>
              <div className="flex gap-2">
                <input type="date" value={pvFrom} onChange={(e) => setPvFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                <input type="date" value={pvTo} onChange={(e) => setPvTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
              </div>
            </div>

            {pvLoading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : pvRows.length === 0 ? (
              <p className="text-sm text-zinc-500">No purchase records in this range.</p>
            ) : (
              <>
                <div className="flex flex-wrap gap-3 mb-6">
                  <div className="bg-[#131316] border border-zinc-800 p-4 min-w-[160px]">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Total spend</p>
                    <p className="text-xl font-black text-yellow-400">₹{Math.round(pvAnalysis.totalSpend).toLocaleString("en-IN")}</p>
                    <p className="text-[10px] text-zinc-600 mt-1">{pvRows.length} purchase entries · {pvRows[0]?.date} to {pvRows[pvRows.length - 1]?.date}</p>
                  </div>
                  <div className="bg-[#131316] border border-zinc-800 p-4 min-w-[160px]">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Vendors</p>
                    <p className="text-xl font-black">{pvAnalysis.vendorRanked.length}</p>
                  </div>
                  <div className="bg-[#131316] border border-red-500/30 p-4 min-w-[160px]">
                    <p className="text-[10px] font-mono text-red-400 uppercase tracking-widest mb-1">Price jumps (≥10%)</p>
                    <p className="text-xl font-black text-red-400">{pvAnalysis.priceCreep.length}</p>
                  </div>
                  <div className={`bg-[#131316] border p-4 min-w-[160px] ${pvAnalysis.modelGapPct == null ? "border-zinc-800" : Math.abs(pvAnalysis.modelGapPct) >= 5 ? "border-orange-500/40" : "border-zinc-800"}`}>
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Food cost % (this range)</p>
                    <p className="text-xl font-black">{pvAnalysis.selectedRangeFoodCostPct == null ? "—" : `${pvAnalysis.selectedRangeFoodCostPct.toFixed(1)}%`}</p>
                    <p className="text-[10px] text-zinc-600 mt-1">vs. {pvAnalysis.assumedFoodCostPct.toFixed(1)}% used app-wide (this calendar month to date)</p>
                  </div>
                </div>

                <div className={`border p-4 mb-6 ${pvAnalysis.modelGapPct != null && Math.abs(pvAnalysis.modelGapPct) >= 5 ? "bg-orange-950/20 border-orange-500/40" : "bg-[#131316] border-zinc-800"}`}>
                  <p className="text-[10px] font-mono text-zinc-400 uppercase tracking-widest mb-1">💰 This range vs. the app-wide figure</p>
                  {pvAnalysis.selectedRangeFoodCostPct == null ? (
                    <p className="text-xs text-zinc-600">No revenue data in this date range to compare against.</p>
                  ) : (
                    <>
                      <p className="text-xs text-zinc-300 mb-1">Actual purchase spend was <span className="font-bold text-white">₹{Math.round(pvAnalysis.totalSpend).toLocaleString("en-IN")}</span> against <span className="font-bold text-white">₹{Math.round(pvAnalysis.totalRevenue).toLocaleString("en-IN")}</span> revenue in this window — a food cost of <span className="font-bold text-yellow-400">{pvAnalysis.selectedRangeFoodCostPct.toFixed(1)}%</span> for this specific date range.</p>
                      <p className="text-xs text-zinc-400">Outlet P&amp;L, Channel P&amp;L and Command Centre now use a real figure too — <span className="font-bold">{pvAnalysis.assumedFoodCostPct.toFixed(1)}%</span>, based on this calendar month to date rather than this page's selected range. {pvAnalysis.modelGapPct != null && (
                        Math.abs(pvAnalysis.modelGapPct) < 2
                          ? <span className="text-green-400">Close to each other — expected, since both are real numbers now.</span>
                          : pvAnalysis.modelGapPct > 0
                            ? <span className="text-yellow-400">This range is running {pvAnalysis.modelGapPct.toFixed(1)} points higher than the rolling figure — likely just this period being pricier, or a shorter/different window than 30 days.</span>
                            : <span className="text-yellow-400">This range is running {Math.abs(pvAnalysis.modelGapPct).toFixed(1)} points lower than the rolling figure — likely just this period being cheaper, or a shorter/different window than 30 days.</span>
                      )}</p>
                      <p className="text-[10px] text-zinc-600 mt-2">Purchases aren't outlet-tagged yet, and not every ingredient purchased converts to revenue in the exact same window it was bought (some gets used later) — so some gap between any two windows is expected, not a bug.</p>
                    </>
                  )}
                </div>

                {pvItemPerfCategories && pvAnalysis.categoryMatch.length > 0 && (
                  <div className="bg-[#131316] border border-zinc-800 p-4 mb-6">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">📦 Category cost ratio — ingredient spend vs. item revenue</p>
                    <p className="text-[10px] text-zinc-600 mb-3">Revenue from Item Performance's latest upload ("{pvItemPerfCategories.label}") · spend from the purchase date range above — two different windows, shown together as a directional check, not a perfect match.</p>
                    <div className="overflow-x-auto">
                      <table className="w-full text-xs font-mono whitespace-nowrap">
                        <thead>
                          <tr className="text-zinc-500 uppercase tracking-widest text-[10px] border-b border-zinc-800">
                            <th className="text-left py-2 pr-3">Category</th>
                            <th className="text-right py-2 pl-3">Item revenue</th>
                            <th className="text-right py-2 pl-3">Ingredient spend</th>
                            <th className="text-right py-2 pl-3">Cost ratio</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pvAnalysis.categoryMatch.map((c) => (
                            <tr key={c.category} className="border-b border-zinc-800/40">
                              <td className="py-1.5 pr-3 text-zinc-300">{c.category}</td>
                              <td className="py-1.5 pl-3 text-right text-zinc-400">₹{Math.round(c.revenue).toLocaleString("en-IN")}</td>
                              <td className="py-1.5 pl-3 text-right text-zinc-400">₹{Math.round(c.spend).toLocaleString("en-IN")}</td>
                              <td className={`py-1.5 pl-3 text-right font-bold ${c.costPct >= 40 ? "text-red-400" : c.costPct >= 30 ? "text-yellow-400" : "text-green-400"}`}>{c.costPct.toFixed(1)}%</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {(pvAnalysis.unmatchedSpendCategories.length > 0 || pvAnalysis.unmatchedRevenueCategories.length > 0) && (
                      <p className="text-[10px] text-orange-400 mt-3">⚠ Not shown above (no match on the other side): {[...pvAnalysis.unmatchedSpendCategories.map((c: string) => `"${c}" (spend only)`), ...pvAnalysis.unmatchedRevenueCategories.map((c: string) => `"${c}" (revenue only)`)].join(", ")} — likely just a naming difference between the two systems, worth checking rather than assuming it's zero.</p>
                    )}
                  </div>
                )}

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-6">
                  <div className="bg-[#131316] border border-zinc-800 p-4">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Spend by vendor <span className="text-zinc-700 normal-case">— click one to drill in</span></p>
                    <div className="space-y-1.5">
                      {pvAnalysis.vendorRanked.slice(0, 12).map(([agency, amt]) => (
                        <button key={agency} onClick={() => setPvVendorSel(agency === pvVendorSel ? "" : agency)} className={`w-full flex justify-between text-xs py-1 border-t border-zinc-800/60 text-left transition-colors ${agency === pvVendorSel ? "bg-yellow-400/10" : "hover:bg-zinc-900"}`}>
                          <span className={agency === pvVendorSel ? "text-yellow-400 font-bold" : "text-zinc-300"}>{agency}</span>
                          <span className="font-mono text-zinc-400">₹{Math.round(amt).toLocaleString("en-IN")}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                  <div className="bg-[#131316] border border-zinc-800 p-4">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Spend by category</p>
                    <div className="space-y-1.5">
                      {pvAnalysis.categoryRanked.slice(0, 12).map(([cat, amt]) => (
                        <div key={cat} className="flex justify-between text-xs py-1 border-t border-zinc-800/60">
                          <span className="text-zinc-300">{cat}</span>
                          <span className="font-mono text-zinc-400">₹{Math.round(amt).toLocaleString("en-IN")}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>

                {pvVendorDrilldown && (
                  <div className="bg-[#131316] border border-yellow-400/40 p-4 mb-6">
                    <div className="flex justify-between items-start mb-1">
                      <p className="text-sm font-bold text-yellow-400">{pvVendorSel}</p>
                      <button onClick={() => setPvVendorSel("")} className="text-[10px] text-zinc-600 hover:text-white">✕ close</button>
                    </div>
                    <p className="text-[11px] text-zinc-500 mb-3">₹{Math.round(pvVendorDrilldown.totalSpent).toLocaleString("en-IN")} total · #{pvVendorDrilldown.rankPosition} of {pvVendorDrilldown.totalVendors} vendors by spend in this range</p>
                    <div className="overflow-x-auto">
                      <table className="w-full text-xs font-mono whitespace-nowrap">
                        <thead>
                          <tr className="text-zinc-500 uppercase tracking-widest text-[10px] border-b border-zinc-800">
                            <th className="text-left py-2 pr-3">Item</th>
                            <th className="text-right py-2 pl-3">Qty bought</th>
                            <th className="text-right py-2 pl-3">Avg rate</th>
                            <th className="text-right py-2 pl-3">Spent</th>
                            <th className="text-right py-2 pl-3">vs. cheapest elsewhere</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pvVendorDrilldown.items.map((it, i) => (
                            <tr key={i} className="border-b border-zinc-800/40">
                              <td className="py-1.5 pr-3 text-zinc-300">{it.product}</td>
                              <td className="py-1.5 pl-3 text-right text-zinc-400">{it.qty}</td>
                              <td className="py-1.5 pl-3 text-right text-zinc-300">₹{it.avgRate.toFixed(2)}</td>
                              <td className="py-1.5 pl-3 text-right text-white font-bold">₹{Math.round(it.amount).toLocaleString("en-IN")}</td>
                              <td className={`py-1.5 pl-3 text-right ${it.vsCheapestPct == null ? "text-zinc-700" : it.vsCheapestPct <= 0 ? "text-green-400" : it.vsCheapestPct >= 10 ? "text-red-400 font-bold" : "text-yellow-400"}`}>
                                {it.vsCheapestPct == null ? "only vendor" : `${it.vsCheapestPct >= 0 ? "+" : ""}${it.vsCheapestPct.toFixed(0)}%`}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <p className="text-[10px] text-zinc-600 mt-3">"vs. cheapest elsewhere" compares this vendor's average rate to the lowest rate any other vendor charged for the same item, in this same date range. Green = cheapest or tied. "Only vendor" = nobody else supplied this item in this window, so there's nothing to compare against.</p>
                  </div>
                )}

                <div className="bg-[#131316] border border-red-500/20 p-4 mb-6">
                  <p className="text-[10px] font-mono text-red-400 uppercase tracking-widest mb-1">📈 Price creep — same vendor, rising rate</p>
                  <p className="text-[10px] text-zinc-600 mb-3">Comparing the first and last purchase of each item from each vendor in this window — flags a genuine ≥10% rise, not day-to-day noise.</p>
                  {pvAnalysis.priceCreep.length === 0 ? (
                    <p className="text-xs text-zinc-600">Nothing flagged in this range.</p>
                  ) : (
                    <div className="space-y-2">
                      {pvAnalysis.priceCreep.slice(0, 15).map((p, i) => (
                        <div key={i} className="flex justify-between items-center text-xs py-1.5 border-t border-zinc-800/60">
                          <div>
                            <span className="text-zinc-200">{p.product}</span>
                            <span className="text-zinc-600"> · {p.agency}</span>
                          </div>
                          <div className="text-right">
                            <span className="font-mono text-zinc-400">₹{p.firstRate.toFixed(2)} → ₹{p.lastRate.toFixed(2)}</span>
                            <span className="font-mono text-red-400 font-bold ml-2">+{p.changePct.toFixed(0)}%</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <div className="bg-[#131316] border border-yellow-500/20 p-4">
                  <p className="text-[10px] font-mono text-yellow-400 uppercase tracking-widest mb-1">⚖️ Vendor price comparison — same item, different vendors</p>
                  <p className="text-[10px] text-zinc-600 mb-3">Only items bought from 2+ vendors in this window, where the gap between cheapest and priciest is ≥10%.</p>
                  {pvAnalysis.vendorComparison.length === 0 ? (
                    <p className="text-xs text-zinc-600">No comparable items in this range.</p>
                  ) : (
                    <div className="space-y-2">
                      {pvAnalysis.vendorComparison.slice(0, 15).map((v, i) => (
                        <div key={i} className="text-xs py-1.5 border-t border-zinc-800/60">
                          <p className="text-zinc-200 mb-1">{v.product} <span className="text-yellow-400 font-bold">({v.gapPct.toFixed(0)}% gap)</span></p>
                          <p className="text-[10px] text-zinc-500">{v.vendorAvgs.map((va) => `${va.agency}: ₹${va.avgRate.toFixed(2)}`).join(" · ")}</p>
                        </div>
                      ))}
                    </div>
                  )}
                </div>

                <p className="text-[10px] text-zinc-600 mt-4">Company-wide — purchases in Production Stock aren't tagged to a specific outlet yet, so this can't be broken down per outlet. Invoice number/photo tracking exists in the source table but isn't populated yet.</p>
              </>
            )}
          </div>
        )}
        {activeTab === "cash_flow" && (user?.role === "Financial Analyst" || user?.role === "Owner" || user?.role === "Founder's Office") && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl md:text-3xl font-black tracking-tight">Weekly Profit Forecast</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Real COGS/commission applied · not just sales in, fixed costs out</p>
            </div>
            {(realCommissionPct === null || new Date().getDate() <= 5) && (
              <div className="mb-6 border border-yellow-500/40 bg-yellow-500/5 p-4">
                <p className="text-[11px] font-mono text-yellow-400 uppercase tracking-widest mb-1.5">⚠ Not reliable yet — data gap, not a real prediction</p>
                <ul className="text-xs text-zinc-300 space-y-1 list-disc list-inside">
                  {realCommissionPct === null && (
                    <li>No real Swiggy/Zomato payout records found for this period — commission is showing a 50% placeholder guess, not your actual rate.</li>
                  )}
                  {new Date().getDate() <= 5 && (
                    <li>We're only a few days into this month, so the ingredient-cost % is based on very little purchase data so far and will settle down as the month fills in.</li>
                  )}
                </ul>
              </div>
            )}
            {cfLoading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : cfWeeks.length === 0 ? (
              <p className="text-sm text-zinc-500">Not enough sales history yet to build a forecast.</p>
            ) : (
              <>
                {(() => {
                  const forecastTotal = cfWeeks.filter((w) => w.isForecast).reduce((s, w) => s + w.netProfit, 0);
                  return (
                    <div className="bg-yellow-400/5 border border-yellow-400/30 p-4 max-w-2xl mb-6">
                      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Projected profit, next 4 weeks</p>
                      <p className={`text-2xl font-black ${forecastTotal >= 0 ? "text-green-400" : "text-red-500"}`}>₹{Math.round(forecastTotal).toLocaleString("en-IN")}</p>
                    </div>
                  );
                })()}
                <div className="overflow-x-auto mb-6">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                        <th className="py-2 pr-3">Week of</th>
                        <th className="py-2 pr-3 text-right">Sales</th>
                        <th className="py-2 pr-3 text-right">COGS</th>
                        <th className="py-2 pr-3 text-right">Wastage</th>
                        <th className="py-2 pr-3 text-right">Commission</th>
                        <th className="py-2 pr-3 text-right">Fixed Costs</th>
                        <th className="py-2 text-right">Net Profit</th>
                      </tr>
                    </thead>
                    <tbody>
                      {cfWeeks.map((w, i) => (
                        <tr key={i} className={`border-b border-zinc-900 ${w.isForecast ? "bg-yellow-400/5" : ""}`}>
                          <td className="py-2 pr-3 font-mono text-xs">{w.week}{w.isForecast && <span className="ml-2 text-[9px] text-yellow-400 uppercase">Forecast</span>}</td>
                          <td className="py-2 pr-3 text-right font-mono">₹{Math.round(w.sales).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(w.cogs).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(w.wastage).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(w.commission).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(w.fixed).toLocaleString("en-IN")}</td>
                          <td className={`py-2 text-right font-mono font-bold ${w.netProfit >= 0 ? "text-green-400" : "text-red-500"}`}>₹{Math.round(w.netProfit).toLocaleString("en-IN")}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="bg-yellow-400/5 border border-yellow-400/30 p-4 max-w-2xl">
                  <p className="text-xs text-zinc-300 leading-relaxed">The 4 highlighted weeks are projected from your last month's actual sales trend (shop and online kept separate, since commission only applies to online) — not guaranteed, just where things are headed if nothing changes. COGS and commission use the same real, trailing % as the rest of the app. This doesn't yet include supplier cheque payments, since that data isn't live in the system — once the Cheque Ledger is built, those outflows will factor in here too and the forecast will get sharper.</p>
                </div>
              </>
            )}
          </div>
       )}
       {activeTab === "sales_forecast" && (user?.role === "Financial Analyst" || user?.role === "Owner" || user?.role === "Founder's Office") && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl md:text-3xl font-black tracking-tight">Sales Forecast</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Per outlet · shop / Swiggy / Zomato · next 7 days + month-end</p>
            </div>
            {sfLoading ? (
              <p className="text-sm text-zinc-500">Loading…</p>
            ) : sfOutlets.length === 0 ? (
              <p className="text-sm text-zinc-500">Not enough sales history yet to build a forecast.</p>
            ) : (
              <>
                {(() => {
                  const totalProjSales = sfMonthEnd.reduce((s, r) => s + r.projected, 0);
                  const totalNetProfit = sfMonthEnd.reduce((s, r) => s + r.netProfit, 0);
                  return (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                      <div className="bg-yellow-400/5 border border-yellow-400/30 p-4">
                        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{sfMonthLabel} — projected company sales</p>
                        <p className="text-2xl font-black text-green-400">₹{Math.round(totalProjSales).toLocaleString("en-IN")}</p>
                        <p className="text-[11px] text-zinc-500 mt-1">₹{Math.round(sfMonthEnd.reduce((s, r) => s + r.mtd, 0)).toLocaleString("en-IN")} actual so far + ₹{Math.round(sfMonthEnd.reduce((s, r) => s + r.remaining, 0)).toLocaleString("en-IN")} forecast for the remaining {sfDaysRemaining} day{sfDaysRemaining === 1 ? "" : "s"}</p>
                      </div>
                      <div className="bg-yellow-400/5 border border-yellow-400/30 p-4">
                        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{sfMonthLabel} — projected company net profit</p>
                        <p className={`text-2xl font-black ${totalNetProfit >= 0 ? "text-green-400" : "text-red-500"}`}>₹{Math.round(totalNetProfit).toLocaleString("en-IN")}</p>
                        <p className="text-[11px] text-zinc-500 mt-1">After real COGS, wastage, commission and each outlet's own fixed costs</p>
                      </div>
                    </div>
                  );
                })()}

                <h3 className="text-lg font-bold mb-1">Month-end P&amp;L projection, by outlet</h3>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Actual month-to-date + forecast for the rest of {sfMonthLabel} · real COGS/commission rates applied</p>
                <div className="overflow-x-auto mb-8">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                        <th className="py-2 pr-3">Outlet</th>
                        <th className="py-2 pr-3 text-right">Projected sales</th>
                        <th className="py-2 pr-3 text-right">COGS</th>
                        <th className="py-2 pr-3 text-right">Wastage</th>
                        <th className="py-2 pr-3 text-right">Commission</th>
                        <th className="py-2 pr-3 text-right">Fixed Costs</th>
                        <th className="py-2 text-right">Net Profit</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...sfMonthEnd].sort((a, b) => b.netProfit - a.netProfit).map((r) => (
                        <tr key={r.oid} className="border-b border-zinc-900">
                          <td className="py-2 pr-3">{r.name}{!r.hasFixedData && <span className="ml-1.5 text-[9px] text-yellow-500 uppercase">no fixed cost set</span>}</td>
                          <td className="py-2 pr-3 text-right font-mono">₹{Math.round(r.projected).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.cogs).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.wastage).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.commission).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(r.fixed).toLocaleString("en-IN")}</td>
                          <td className={`py-2 text-right font-mono font-bold ${r.netProfit >= 0 ? "text-green-400" : "text-red-500"}`}>₹{Math.round(r.netProfit).toLocaleString("en-IN")}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <h3 className="text-lg font-bold mb-1">Next 7 days, by outlet and channel</h3>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Based on the same day-of-week over the last ~10 weeks, adjusted for the recent trend</p>
                <div className="space-y-6">
                  {[...sfOutlets].sort((a, b) => b.weekTotal - a.weekTotal).map((o) => (
                    <div key={o.oid} className="border border-zinc-800 p-4">
                      <div className="flex items-center justify-between mb-2">
                        <p className="font-bold">{o.name}</p>
                        <p className={`text-[11px] font-mono ${o.trendPct >= 0 ? "text-green-400" : "text-red-500"}`}>{o.trendPct >= 0 ? "+" : ""}{o.trendPct.toFixed(0)}% recent trend</p>
                      </div>
                      {!o.hasHistory ? (
                        <p className="text-xs text-zinc-500">No sales history yet for this outlet.</p>
                      ) : (
                        <div className="overflow-x-auto">
                          <table className="w-full text-xs">
                            <thead>
                              <tr className="text-left text-[9px] font-mono text-zinc-500 uppercase border-b border-zinc-900">
                                <th className="py-1.5 pr-3">Date</th>
                                <th className="py-1.5 pr-3 text-right">Shop</th>
                                <th className="py-1.5 pr-3 text-right">Swiggy</th>
                                <th className="py-1.5 pr-3 text-right">Zomato</th>
                                <th className="py-1.5 text-right">Total</th>
                              </tr>
                            </thead>
                            <tbody>
                              {o.days.map((d: any) => (
                                <tr key={d.date} className="border-b border-zinc-900/60">
                                  <td className="py-1.5 pr-3 font-mono text-zinc-400">{new Date(d.date + "T00:00:00Z").toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short" })}</td>
                                  <td className="py-1.5 pr-3 text-right font-mono text-zinc-400">₹{Math.round(d.shop).toLocaleString("en-IN")}</td>
                                  <td className="py-1.5 pr-3 text-right font-mono text-zinc-400">₹{Math.round(d.swiggy).toLocaleString("en-IN")}</td>
                                  <td className="py-1.5 pr-3 text-right font-mono text-zinc-400">₹{Math.round(d.zomato).toLocaleString("en-IN")}</td>
                                  <td className="py-1.5 text-right font-mono font-bold">₹{Math.round(d.total).toLocaleString("en-IN")}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  ))}
                </div>

                <div className="bg-yellow-400/5 border border-yellow-400/30 p-4 mt-6 max-w-2xl mb-10">
                  <p className="text-xs text-zinc-300 leading-relaxed">Sales are forecast by outlet, channel and day, from real history — the same day-of-week over the last ~10 weeks, nudged by whether that outlet has been trending up or down lately. The month-end P&amp;L applies the same real COGS % and commission % used everywhere else in the app, and each outlet's own fixed costs from Outlet P&amp;L's targets — so "no fixed cost set" means that outlet's fixed costs haven't been filled in yet, and its net profit here will read too high until they are. This does not yet forecast by category or product (the item-level data we get is a one-time snapshot each upload, not a dated daily series to project from), by hour (no hourly sales data exists in the system yet), or include staffing or scenario planning — those are separate pieces still to come.</p>
                </div>

                <div className="border-t border-zinc-800 pt-8 mb-8">
                  <h2 className="text-2xl md:text-3xl font-black tracking-tight">Production Forecast</h2>
                  <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Next 7 days · units to bake, by product</p>
                </div>
                {pfLoading ? (
                  <p className="text-sm text-zinc-500 mb-10">Loading…</p>
                ) : pfProducts.length === 0 ? (
                  <p className="text-sm text-zinc-500 mb-10">Not enough logged production history yet (Kitchen Production tab) to forecast — needs a few weeks of daily entries per product.</p>
                ) : (
                  <>
                    <h3 className="text-lg font-bold mb-1">By category</h3>
                    <div className="overflow-x-auto mb-6">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                            <th className="py-2 pr-3">Category</th>
                            <th className="py-2 text-right">Forecast units, next 7 days</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pfCategories.map((c) => (
                            <tr key={c.category} className="border-b border-zinc-900">
                              <td className="py-2 pr-3">{c.category}</td>
                              <td className="py-2 text-right font-mono font-bold">{Math.round(c.weekQty).toLocaleString("en-IN")}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <h3 className="text-lg font-bold mb-1">By product</h3>
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Only products with enough production history to forecast are shown</p>
                    <div className="overflow-x-auto mb-10">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                            <th className="py-2 pr-3">Product</th>
                            <th className="py-2 pr-3">Category</th>
                            <th className="py-2 text-right">Forecast units, next 7 days</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pfProducts.map((p: any) => (
                            <tr key={p.id} className="border-b border-zinc-900">
                              <td className="py-2 pr-3">{p.name}</td>
                              <td className="py-2 pr-3 text-zinc-500 text-xs">{p.category}</td>
                              <td className="py-2 text-right font-mono">{Math.round(p.weekQty).toLocaleString("en-IN")}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                )}

                <div className="border-t border-zinc-800 pt-8 mb-8">
                  <h2 className="text-2xl md:text-3xl font-black tracking-tight">Purchase Forecast</h2>
                  <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Next week · quantity and ₹ to buy, by raw material</p>
                </div>
                {purchLoading ? (
                  <p className="text-sm text-zinc-500">Loading…</p>
                ) : purchForecast.length === 0 ? (
                  <p className="text-sm text-zinc-500">Not enough purchase history yet to forecast — needs a few weeks of logged purchases per item.</p>
                ) : (
                  <div className="overflow-x-auto mb-6">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                          <th className="py-2 pr-3">Item</th>
                          <th className="py-2 pr-3">Category</th>
                          <th className="py-2 pr-3 text-right">Forecast qty, next week</th>
                          <th className="py-2 pr-3 text-right">Forecast ₹, next week</th>
                          <th className="py-2 text-right">Trend</th>
                        </tr>
                      </thead>
                      <tbody>
                        {purchForecast.map((p, i) => (
                          <tr key={i} className="border-b border-zinc-900">
                            <td className="py-2 pr-3">{p.product}</td>
                            <td className="py-2 pr-3 text-zinc-500 text-xs">{p.category}</td>
                            <td className="py-2 pr-3 text-right font-mono">{p.forecastQty.toLocaleString("en-IN", { maximumFractionDigits: 1 })} {p.unit}</td>
                            <td className="py-2 pr-3 text-right font-mono text-zinc-400">₹{Math.round(p.forecastAmount).toLocaleString("en-IN")}</td>
                            <td className={`py-2 text-right font-mono text-xs ${p.trendPct >= 0 ? "text-green-400" : "text-red-500"}`}>{p.trendPct >= 0 ? "+" : ""}{p.trendPct.toFixed(0)}%</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                <div className="bg-yellow-400/5 border border-yellow-400/30 p-4 max-w-2xl">
                  <p className="text-xs text-zinc-300 leading-relaxed">Production is forecast from real Kitchen Production log entries (same day-of-week average over the last ~10 weeks, trend-adjusted). Purchases are forecast from real Purchase Ledger entries, averaged over the last ~4 full weeks and trend-adjusted, at a weekly grain since purchases don't happen every day. Neither of these is a true ingredient/recipe forecast — that would need a bill-of-materials (how much flour, sugar etc. goes into one unit of each product), which doesn't exist in the system yet. What's shown here is the closest real-data equivalent: what's actually been produced and purchased, projected forward.</p>
                </div>
              </>
            )}
          </div>
       )}
       {activeTab === "contribution_margins" && user?.role === "Financial Analyst" && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl md:text-3xl font-black tracking-tight">Contribution Margins</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Outlet and product-level profitability</p>
            </div>
            <div>
              <h3 className="text-xl font-black tracking-tight mb-1">Outlet Contribution Margins</h3>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Ranked worst to best · same range as the P&amp;L tab</p>
              {pnlRows.length === 0 ? (
                <p className="text-sm text-zinc-500">No data yet — open the P&amp;L tab once first to load a date range.</p>
              ) : (
                <div className="space-y-1.5 max-w-2xl">
                  {[...pnlRows].sort((a, b) => a.netMargin - b.netMargin).map((r) => {
                    const col = r.netMargin >= 10 ? "text-green-400" : r.netMargin >= 0 ? "text-yellow-400" : "text-red-500";
                    const barW = Math.min(100, Math.max(2, Math.abs(r.netMargin) * 3));
                    return (
                      <div key={r.oid} className="flex items-center gap-3 text-sm">
                        <span className="w-28 shrink-0 font-medium">{r.name}</span>
                        <div className="flex-1 h-2 bg-zinc-800"><div className={`h-full ${r.netMargin >= 0 ? "bg-green-500" : "bg-red-500"}`} style={{ width: `${barW}%` }} /></div>
                        <span className={`font-mono w-14 text-right ${col}`}>{r.netMargin.toFixed(1)}%</span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="mt-10 pt-6 border-t border-zinc-800">
              <h3 className="text-xl font-black tracking-tight mb-1">Product Contribution Margins</h3>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-1">From the latest Item Performance upload</p>
              <p className="text-xs text-zinc-500 mb-4">Chain-wide, not split by outlet — Item Performance data doesn't carry an outlet field, so this shows overall product profitability, not per-outlet.</p>
              {cmLoading ? (
                <p className="text-sm text-zinc-500">Loading…</p>
              ) : cmProductRows.length === 0 ? (
                <p className="text-sm text-zinc-500">No Item Performance upload yet — upload one from the Item Performance tab first.</p>
              ) : (
                <div className="overflow-x-auto max-w-3xl">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                        <th className="py-2 pr-3">Product</th>
                        <th className="py-2 pr-3 text-right">Revenue</th>
                        <th className="py-2 pr-3 text-right">Units</th>
                        <th className="py-2 pr-3 text-right">Cost</th>
                        <th className="py-2 text-right">Margin %</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...cmProductRows].sort((a, b) => b.revenue - a.revenue).slice(0, 20).map((p, i) => (
                        <tr key={i} className="border-b border-zinc-900">
                          <td className="py-2 pr-3">{p.name}</td>
                          <td className="py-2 pr-3 text-right font-mono">₹{Math.round(p.revenue).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">{p.units.toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3 text-right font-mono text-zinc-400">{p.cost != null ? "₹" + Math.round(p.cost).toLocaleString("en-IN") : "—"}</td>
                          <td className={`py-2 text-right font-mono font-bold ${p.marginPct == null ? "text-zinc-600" : p.marginPct < 40 ? "text-red-500" : p.marginPct < 60 ? "text-yellow-400" : "text-green-400"}`}>{p.marginPct == null ? "no cost data" : p.marginPct.toFixed(1) + "%"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <p className="text-[10px] text-zinc-600 mt-2">"no cost data" means that product isn't in the recipe cost map yet — margin can't be calculated until it is.</p>
                </div>
              )}
            </div>

          </div>
       )}
       {activeTab === "tasks" && user?.role !== "Founder's Office" && user?.role !== "Owner" && user?.role !== "Head Chef" && user?.role !== "Financial Analyst" && (
         <div>
            {canAssign && compTop && (compHeadlineOn || user?.role === "Owner") && (
              <div className="mb-6 border border-zinc-800 bg-zinc-900/40 px-5 py-4">
                <div className="flex justify-between items-start gap-4">
                  <div className="flex-1">
                    {compHeadlineOn ? (
                      <>
                        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Competition watch · {compActivePeriod}</p>
                        <p className="text-base md:text-lg">{compFunny.playful}</p>
                      </>
                    ) : (
                      <p className="text-[11px] font-mono text-zinc-600 uppercase tracking-widest">Competition headline is hidden for everyone</p>
                    )}
                  </div>
                  {user?.role === "Owner" && (
                    <button onClick={toggleCompHeadline} className="shrink-0 text-[10px] font-mono uppercase px-3 py-1.5 border border-zinc-700 hover:border-yellow-400 hover:text-yellow-400 transition-colors">{compHeadlineOn ? "Hide for all" : "Show for all"}</button>
                  )}
                </div>
              </div>
            )}
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">{canAssign ? "Command Center" : "My Tasks"}</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Welcome back, {user.name.split(" ")[0]}</p>
              </div>
              {canAssign && (
                <button onClick={() => setShowModal(true)} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-4 py-3 hover:opacity-90 transition-opacity uppercase">+ Assign Task</button>
              )}
           </div>
            {canAssign && offToday.length > 0 && (
              <div className="bg-[#131316] border border-zinc-800 px-5 py-3 mb-6">
                <span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">🌙 Off today: </span>
                <span className="text-sm text-yellow-400">{offToday.map(id => ALL_STAFF.find(s => s.id === id)?.name || id).join(", ")}</span>
              </div>
            )}
            
                {user.role !== "Owner" && (
              <div className="flex items-center justify-between bg-[#131316] border border-zinc-800 px-5 py-4 mb-6">
                <div>
                  <p className="text-sm font-semibold">Off day today</p>
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mt-0.5">Reports you submit won't earn or lose points</p>
                </div>
               <button onClick={toggleOffDay} className={`relative w-12 h-6 rounded-full transition-colors ${reportOffDay ? "bg-yellow-400" : "bg-zinc-700"}`}>
                  <span className={`absolute top-0.5 left-0.5 w-5 h-5 bg-black rounded-full transition-transform ${reportOffDay ? "translate-x-6" : ""}`}></span>
                </button>
              </div>
            )}
           {tasks.filter(t => t.assigned_to === user.id && t.status !== "completed").length > 0 && (
              <div className="bg-yellow-400/5 border border-yellow-400/40 p-5 mb-6">
                <p className="text-[11px] font-mono text-yellow-400 uppercase tracking-widest mb-3">📌 {tasks.filter(t => t.assigned_to === user.id && t.status !== "completed").length} task(s) assigned to you</p>
                <div className="space-y-2">
                  {tasks.filter(t => t.assigned_to === user.id && t.status !== "completed").map(t => {
                    const isOverdue = new Date(t.due_at) < new Date();
                    return (
                      <div key={t.id} className="flex flex-wrap gap-2 items-center bg-black/30 border border-zinc-800 px-3 py-2.5">
                        <div className={`w-2 h-2 rounded-full shrink-0 ${t.priority === "critical" ? "bg-red-500" : t.priority === "high" ? "bg-orange-400" : t.priority === "medium" ? "bg-yellow-400" : "bg-zinc-600"}`} />
                        <div className="flex-1 min-w-0">
                          <p className="font-semibold text-sm">{t.title}</p>
                          {t.description && <p className="text-[11px] text-zinc-400 mt-0.5">{t.description}</p>}
                          <p className="text-[10px] font-mono text-zinc-500 mt-0.5">From {ALL_STAFF.find(s => s.id === t.assigned_by)?.name || FORMER_STAFF[t.assigned_by] || t.assigned_by} · Due {t.due_at ? new Date(t.due_at).toLocaleString("en-IN", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "—"}{isOverdue ? " · OVERDUE" : ""}{t.outlet_id ? ` · ${OUTLET_NAMES[t.outlet_id] || t.outlet_id}` : ""}</p>
                        </div>
                        <div className="flex gap-2">
                          {t.status === "assigned" && <button onClick={() => updateStatus(t.id, "in_progress")} className="text-[10px] font-mono uppercase px-2 py-1 border border-zinc-700 hover:border-yellow-400 hover:text-yellow-400 transition-colors">Start</button>}
                          <button onClick={() => updateStatus(t.id, "completed")} className="text-[10px] font-mono uppercase px-2 py-1 border border-zinc-700 hover:border-green-400 hover:text-green-400 transition-colors">Done</button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
           {canAssign && (
              <div className="flex gap-2 flex-wrap mb-6">
                {["all", ...OUTLETS].map(o => (
                  <button key={o} onClick={() => setOutletFilter(o)} className={`font-mono text-[10px] uppercase tracking-widest px-3 py-1.5 border transition-colors ${outletFilter === o ? "border-yellow-400 text-yellow-400" : "border-zinc-700 text-zinc-500 hover:border-zinc-500"}`}>
                    {o === "all" ? "All" : (OUTLET_NAMES[o] || o.replace(/_/g, " "))}
                  </button>
                ))}
              </div>
            )}
            {canAssign && (
              <div className="bg-[#131316] border border-zinc-800 mb-6">
                <div className="px-5 py-3 border-b border-zinc-800">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Staff Status — Today (report time)</p>
                </div>
                {ALL_STAFF.filter(s => s.id !== "nishant").map(s => {
                  const staffTasks = tasks.filter(t => t.assigned_to === s.id);
                  const staffOverdue = staffTasks.filter(t => t.status !== "completed" && new Date(t.due_at) < new Date()).length;
                  const staffCompleted = staffTasks.filter(t => t.status === "completed").length;
                 const todayReport = reports.find(r => r.staff_id === s.id && new Date(r.submitted_at).toDateString() === new Date().toDateString());
                  const hasReport = !!todayReport;
                  return (
                    <div key={s.id} className="grid grid-cols-[1fr_56px_56px_56px_130px] gap-2 items-center px-5 py-3 border-b border-zinc-800 last:border-0">
                      <div>
                        <p className="font-semibold text-sm">{s.name}</p>
                        <p className="text-[10px] font-mono text-zinc-600 uppercase tracking-widest">{s.role}</p>
                      </div>
                      <div className="text-center">
                        <p className="font-mono text-sm font-bold">{staffTasks.length}</p>
                        <p className="text-[9px] font-mono text-zinc-600 uppercase">Tasks</p>
                      </div>
                      <div className="text-center">
                        <p className={`font-mono text-sm font-bold ${staffOverdue > 0 ? "text-red-500" : "text-zinc-500"}`}>{staffOverdue}</p>
                        <p className="text-[9px] font-mono text-zinc-600 uppercase">Late</p>
                      </div>
                      <div className="text-center">
                        <p className="font-mono text-sm font-bold text-green-400">{staffCompleted}</p>
                        <p className="text-[9px] font-mono text-zinc-600 uppercase">Done</p>
                      </div>
                      <div className="text-center">
                        <span className={`font-mono text-[10px] uppercase tracking-widest px-2 py-1 ${hasReport ? (todayReport!.is_late ? "bg-red-500/10 text-red-500" : "bg-green-400/10 text-green-400") : "bg-yellow-400/10 text-yellow-400"}`}>
                          {hasReport ? `✓ ${new Date(todayReport!.submitted_at).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })}${todayReport!.is_late ? " late" : ""}` : "Pending"}
                        </span>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
           {canAssign && (
              <div className="bg-[#131316] border border-zinc-800 mb-6">
                <div className="px-5 py-3 border-b border-zinc-800">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Outlet Reports — Today{outletFilter !== "all" ? ` · ${OUTLET_NAMES[outletFilter] || outletFilter.replace(/_/g, " ")}` : ""}</p>
                </div>
                {(outletFilter === "all" ? OUTLETS : [outletFilter]).map(o => {
                  const rep = allOutletReports.find(r => r.outlet_id === o);
                  const mgr = ALL_STAFF.find(s => (s.outlets as string[]).includes(o));
                  const oTotal = rep ? (Number(rep.shop_sales_value) || 0) + (Number(rep.swiggy_sales_value) || 0) + (Number(rep.zomato_sales_value) || 0) : 0;
                  const oTgt = rep ? Number(rep.target) || 0 : 0;
                  const hit = oTgt > 0 && oTotal >= oTgt;
                  return (
                    <div key={o} className="border-b border-zinc-800 last:border-0 px-5 py-4">
                      <div className="flex items-center justify-between mb-2">
                        <div>
                          <p className="font-semibold text-sm uppercase tracking-widest">{OUTLET_NAMES[o] || o.replace(/_/g, " ")}</p>
                          <p className="text-[10px] font-mono text-zinc-600">{mgr?.name || "—"}</p>
                        </div>
                        {rep ? (
                          <span className={`font-mono text-[10px] uppercase tracking-widest px-2 py-1 ${rep.is_late ? "bg-red-500/10 text-red-500" : "bg-green-400/10 text-green-400"}`}>✓ {new Date(rep.submitted_at).toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" })}{rep.is_late ? " late" : ""}</span>
                        ) : (
                          <span className="font-mono text-[10px] uppercase tracking-widest px-2 py-1 bg-yellow-400/10 text-yellow-400">Pending</span>
                        )}
                      </div>
                      {rep && (
                        <div className="text-[11px] font-mono text-zinc-400 space-y-1">
                          <div className="flex flex-wrap gap-x-5 gap-y-1">
                            <span>Shop ₹{rep.shop_sales_value} ({rep.shop_sales_count})</span>
                            <span>Swiggy ₹{rep.swiggy_sales_value} ({rep.swiggy_sales_count})</span>
                            <span>Zomato ₹{rep.zomato_sales_value} ({rep.zomato_sales_count})</span>
                            <span className="text-white">Total ₹{oTotal}</span>
                            {oTgt > 0 && <span className={hit ? "text-green-400" : "text-red-500"}>Target ₹{oTgt} {hit ? "✓ hit" : "✗ miss"}</span>}
                          </div>
                          {(rep.bh_google_rating || rep.expiry_count || rep.issues || rep.action_taken) && (
                            <div className="flex flex-wrap gap-x-5 gap-y-1 text-zinc-500">
                              {rep.bh_google_rating ? <span>Google {rep.bh_google_rating}</span> : null}
                              {rep.expiry_count ? <span>Expiry {rep.expiry_count}</span> : null}
                              {rep.issues ? <span>Issues: {rep.issues}</span> : null}
                              {rep.action_taken ? <span>Action: {rep.action_taken}</span> : null}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-px bg-zinc-800 border border-zinc-800 mb-6">
              {[
                { label: "Total Tasks", value: total, sub: "assigned", color: "" },
                { label: "Completed", value: completed, sub: `${rate}% rate`, color: "text-green-400" },
                { label: "In Progress", value: inProgress, sub: "active", color: "text-yellow-400" },
                { label: "Overdue", value: overdue, sub: overdue > 0 ? "action needed" : "all clear", color: overdue > 0 ? "text-red-500" : "" },
              ].map((s) => (
                <div key={s.label} className="bg-[#131316] p-4">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">{s.label}</p>
                  <p className={`text-3xl font-black tracking-tight ${s.color}`}>{s.value}</p>
                  <p className="text-[11px] font-mono text-zinc-600 mt-1">{s.sub}</p>
                </div>
              ))}
            </div>
            {loading ? (
              <div className="bg-[#131316] border border-zinc-800 p-10 text-center">
                <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">Loading...</p>
              </div>
            ) : tasks.filter(t => (outletFilter === "all" || t.outlet_id === outletFilter) && t.status !== "completed").length === 0 ? (
              <div className="bg-[#131316] border border-zinc-800 p-10 text-center">
                <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">No tasks yet</p>
              </div>
            ) : (
              <div className="bg-[#131316] border border-zinc-800">
                {tasks.filter(t => (outletFilter === "all" || t.outlet_id === outletFilter) && t.status !== "completed").map((t) => {
                  const assigneeName = ALL_STAFF.find(s => s.id === t.assigned_to)?.name || FORMER_STAFF[t.assigned_to] || t.assigned_to;
                  const isOverdue = t.status !== "completed" && new Date(t.due_at) < new Date();
                  return (
                    <div key={t.id} className={`flex flex-wrap gap-2 items-center px-4 py-3 border-b border-zinc-800 last:border-0 hover:bg-zinc-900 transition-colors ${isOverdue ? "border-l-2 border-l-red-500" : ""}`}>
                      <div className={`w-2 h-2 rounded-full shrink-0 ${t.priority === "critical" ? "bg-red-500" : t.priority === "high" ? "bg-orange-400" : t.priority === "medium" ? "bg-yellow-400" : "bg-zinc-600"}`} />
                      <div className="flex-1 min-w-0">
                        <p className="font-semibold text-sm">{t.title}</p>
                        <p className="text-[11px] font-mono text-zinc-500 mt-0.5">{assigneeName} · {t.due_at ? new Date(t.due_at).toLocaleString("en-IN", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" }) : "No deadline"}{t.outlet_id ? ` · ${OUTLET_NAMES[t.outlet_id] || t.outlet_id.replace(/_/g, " ")}` : ""}</p>
                      </div>
                      <span className={`font-mono text-[10px] uppercase tracking-widest px-2 py-1 ${t.status === "completed" ? "bg-green-400/10 text-green-400" : isOverdue ? "bg-red-500/10 text-red-500" : t.status === "in_progress" ? "bg-yellow-400/10 text-yellow-400" : "bg-zinc-800 text-zinc-500"}`}>
                        {isOverdue && t.status !== "completed" ? "overdue" : t.status.replace("_", " ")}
                      </span>
                      <div className="flex gap-2">
                        {t.status !== "completed" && (
                          <>
                            {t.status === "assigned" && <button onClick={() => updateStatus(t.id, "in_progress")} className="text-[10px] font-mono uppercase px-2 py-1 border border-zinc-700 hover:border-yellow-400 hover:text-yellow-400 transition-colors">Start</button>}
                            <button onClick={() => updateStatus(t.id, "completed")} className="text-[10px] font-mono uppercase px-2 py-1 border border-zinc-700 hover:border-green-400 hover:text-green-400 transition-colors">Done</button>
                          </>
                        )}
                        {user.role === "Owner" && <button onClick={() => deleteTask(t.id)} className="text-[10px] font-mono uppercase px-2 py-1 border border-zinc-700 hover:border-red-500 hover:text-red-500 transition-colors">✕</button>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

      {activeTab === "sales_target" && (() => {
            // Shared by both the compact Today snapshot table and the detailed
            // per-brand view below — same formulas, same source data, computed
            // once so the two views can never quietly disagree with each other.
            const computeBrandSnapshot = (oid: string, brand: "BH" | "CBH" | "ICBH") => {
              const li = salesTargets[oid]?.[brand];
              if (!li) return null;
              const _outletActuals = stActuals[oid] || { daily: {}, monthNet: 0, monthOnline: 0 };
              const _sales = li.sales || {};
              const dayNet = brand === "BH" ? (_outletActuals.daily[stDate]?.shop || 0) : (Number(_sales[stDate]?.net) || 0);
              const dayOnline = brand === "BH" ? (_outletActuals.daily[stDate]?.online || 0) : (Number(_sales[stDate]?.online) || 0);
              const f = li.fixed || {};
              const isSharedFixed = brand !== "BH"; // CBH & ICBH share BH's rent/staff/EB/transport
              const _ab = (v: any) => Math.abs(Number(v) || 0);
              const fStaff = isSharedFixed ? 0 : _ab(f.staff);
              const fRent = isSharedFixed ? 0 : _ab(f.rent);
              const fEb = isSharedFixed ? 0 : _ab(f.eb);
              const fTransport = isSharedFixed ? 0 : _ab(f.transport);
              const rm = 0.2 * fRent;
              const totalFixed = fStaff + fRent + fEb + fTransport + rm + _ab(f.pest) + _ab(f.water) + _ab(f.airtel);
              const stYear = Number(stDate.slice(0, 4)), stMonthNum = Number(stDate.slice(5, 7));
              const daysInThisMonth = new Date(stYear, stMonthNum, 0).getDate();
              const todayTotalSales = dayNet + dayOnline;
              const todayCogs = stCogsRate * todayTotalSales, todayWastage = 0.05 * todayTotalSales, todayComm = stCommissionRate * dayOnline;
              const todayContrib = todayTotalSales - todayCogs - todayWastage - todayComm;
              const todayFixedShare = totalFixed / daysInThisMonth;
              const todayNetProfit = todayContrib - todayFixedShare;
              const todayMargin = todayTotalSales ? todayNetProfit / todayTotalSales : 0;
              return { dayNet, dayOnline, todayTotalSales, todayNetProfit, todayMargin };
            };
            return (
          <div>
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">Outlet P&amp;L</h2>
            <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Net &amp; Online filled per day · P&amp;L sums the whole month · fixed costs &amp; targets edit-once</p>
              </div>
              <div className="text-right">
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Sales day</label>
                <input type="date" max={new Date().toISOString().split("T")[0]} value={stDate} onChange={(e) => { setStDate(e.target.value); setStEditing(null); setStEditValues({}); }} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
              </div>
            </div>
            {(canAssign ? OUTLETS : (user.outlets || [])).map((oid: string) => {
              const dayLblTop = new Date(stDate + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" });
              const snapshots = (["BH", "CBH", "ICBH"] as const).map((b) => ({ brand: b, s: computeBrandSnapshot(oid, b) }));
              return (
              <div key={oid} className="mb-8">
                <div className="flex flex-wrap justify-between items-center gap-2 mb-3">
                  <h3 className="text-lg font-bold">{OUTLET_NAMES[oid] || oid}</h3>
                  <span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{dayLblTop} snapshot</span>
                </div>
                <div className="grid grid-cols-3 gap-2 mb-5">
                  {snapshots.map(({ brand, s }) => (
                    <div key={brand} className="bg-[#131316] border border-zinc-800 p-3">
                      <p className="text-[10px] font-mono text-yellow-400 uppercase tracking-widest mb-1.5">{brand}</p>
                      {!s ? (
                        <p className="text-[11px] text-zinc-700">Not set up</p>
                      ) : (
                        <>
                          <p className="text-[10px] text-zinc-500">Sales</p>
                          <p className="text-sm font-mono text-zinc-200 mb-1">₹{Math.round(s.todayTotalSales).toLocaleString("en-IN")}</p>
                          <p className="text-[10px] text-zinc-500">Net profit</p>
                          <p className={`text-sm font-mono font-bold ${s.todayNetProfit >= 0 ? "text-green-400" : "text-red-400"}`}>₹{Math.round(s.todayNetProfit).toLocaleString("en-IN")} <span className="text-[10px] text-zinc-600">({(s.todayMargin * 100).toFixed(0)}%)</span></p>
                        </>
                      )}
                    </div>
                  ))}
                </div>
              {(["BH", "CBH", "ICBH"] as const).map((brand) => {
                  const li = salesTargets[oid]?.[brand];
                  if (!li) return null;
                  const key = `${oid}_${brand}`;
                  const editing = stEditing === key;
                  const mk = stDate.slice(0, 7);
                  const ml = new Date(stDate + "T00:00:00").toLocaleString("en-IN", { month: "short" });
                  const dayLbl = new Date(stDate + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" });
                  const _outletActuals = stActuals[oid] || { daily: {}, monthNet: 0, monthOnline: 0 };
                  const _sales = li.sales || {};
                  const _mKeys = Object.keys(_sales).filter(d => d.length === 10 && d.startsWith(mk));
                  const _moNet = li.monthly?.[mk]?.net;
                  const _moOnline = li.monthly?.[mk]?.online;
                  const _dNet = _mKeys.reduce((s, d) => s + (Number(_sales[d]?.net) || 0), 0);
                  const _dOnline = _mKeys.reduce((s, d) => s + (Number(_sales[d]?.online) || 0), 0);
                  const dayNet = brand === "BH" ? (_outletActuals.daily[stDate]?.shop || 0) : (Number(_sales[stDate]?.net) || 0);
                  const dayOnline = brand === "BH" ? (_outletActuals.daily[stDate]?.online || 0) : (Number(_sales[stDate]?.online) || 0);
                  const net = brand === "BH" ? _outletActuals.monthNet : ((Number(_moNet) || 0) + _dNet);
                  const online = brand === "BH" ? _outletActuals.monthOnline : ((Number(_moOnline) || 0) + _dOnline);
                  const f = li.fixed || {}; const t = li.targets || {};
                  const isSharedFixed = brand !== "BH";
                  const _ab = (v: any) => Math.abs(Number(v) || 0);
                  const fStaff = isSharedFixed ? 0 : _ab(f.staff);
                  const fRent = isSharedFixed ? 0 : _ab(f.rent);
                  const fEb = isSharedFixed ? 0 : _ab(f.eb);
                  const fTransport = isSharedFixed ? 0 : _ab(f.transport);
                 const totalSales = net + online;
                  const cogs = stCogsRate * totalSales, wastage = 0.05 * totalSales, comm = stCommissionRate * online;
                  const contrib = totalSales - cogs - wastage - comm;
                  const rm = 0.2 * fRent;
                  const totalFixed = fStaff+fRent+fEb+fTransport+rm+_ab(f.pest)+_ab(f.water)+_ab(f.airtel);
                  const netProfit = contrib - totalFixed;
                  const cMargin = totalSales ? contrib / totalSales : 0;
                  const nMargin = totalSales ? netProfit / totalSales : 0;
                  const cmSame = cMargin > 0 ? cMargin : 0.156;
                  const cmDine = 0.656;
                  // Today's P&L — same cost logic as the monthly one, but using just
                  // today's sales, and the monthly fixed cost divided by however many
                  // days are actually in this month (28/29/30/31), not a flat 30.
                  const stYear = Number(stDate.slice(0, 4)), stMonthNum = Number(stDate.slice(5, 7));
                  const daysInThisMonth = new Date(stYear, stMonthNum, 0).getDate();
                  const todayTotalSales = dayNet + dayOnline;
                  const todayCogs = stCogsRate * todayTotalSales, todayWastage = 0.05 * todayTotalSales, todayComm = stCommissionRate * dayOnline;
                  const todayContrib = todayTotalSales - todayCogs - todayWastage - todayComm;
                  const todayFixedShare = totalFixed / daysInThisMonth;
                  const todayNetProfit = todayContrib - todayFixedShare;
                  const todayMargin = todayTotalSales ? todayNetProfit / todayTotalSales : 0;
                  const ta = Number(t.a) || 0, tb = Number(t.b) || 0;
                  const req = (p: number, cm: number) => cm > 0 ? (totalFixed + p) / cm : 0;
                  const m = (n: number) => Math.round(n).toLocaleString("en-IN");
                 const inp = (k: string, fb: number) => editing
                    ? <input key={`${key}_${k}`} type="number" value={stEditValues[k] !== undefined ? stEditValues[k] : (fb ? String(fb) : "")} onChange={(e) => setStEditValues(prev => ({ ...prev, [k]: e.target.value }))} className="w-24 bg-black border border-zinc-700 text-white px-2 py-1 text-right focus:outline-none focus:border-yellow-400" placeholder="0" />
                    : <span>{m(fb)}</span>;
                  const row = (label: string, val: any, opts?: { neg?: boolean; bold?: boolean; fixedCost?: boolean }) => (
                    <tr key={label} className="border-t border-zinc-800/60">
                      <td className={`px-4 py-2 ${opts?.bold ? "text-white font-bold" : "text-zinc-300"}`}>{label}</td>
                      <td className={`px-4 py-2 text-right font-mono ${opts?.fixedCost ? "text-blue-400 font-bold" : opts?.neg ? "text-red-400" : opts?.bold ? "text-yellow-400 font-bold" : ""}`}>{val}</td>
                    </tr>
                  );
                  return (
                    <div key={brand} className="bg-[#131316] border border-zinc-800 mb-5 overflow-x-auto">
                      <div className="flex justify-between items-center px-4 py-3 border-b border-zinc-800">
                        <span className="font-mono text-xs uppercase tracking-widest text-yellow-400">{brand}</span>
                        {editing ? (
                          <div className="flex gap-2">
                           <button onClick={() => saveSalesTarget(oid, brand, li)} disabled={stSaving} className="bg-yellow-400 text-black font-bold text-[10px] px-3 py-1.5 uppercase tracking-widest disabled:opacity-50">{stSaving ? "Saving..." : "Save"}</button>
                            <button onClick={() => setStEditValues(prev => ({ ...prev, net: "0", online: "0" }))} className="text-[10px] font-mono text-red-400 uppercase tracking-widest border border-red-900 px-2 py-1.5 hover:bg-red-950">Clear day</button>
                            <button onClick={() => { setStEditing(null); setStEditValues({}); }} className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest px-2">Cancel</button>
                          </div>
                        ) : (
                          (canAssign || (user.outlets || []).includes(oid)) ? <button onClick={() => { setStEditing(key); setStEditValues({}); }} className="text-[10px] font-mono text-zinc-400 uppercase tracking-widest border border-zinc-700 px-3 py-1.5 hover:border-yellow-400 hover:text-yellow-400">Edit {dayLbl}</button> : <span className="text-[10px] font-mono text-zinc-600 uppercase tracking-widest">View only</span>
                        )}
                      </div>
                      <table className="w-full text-sm">
                        <thead><tr className="text-[10px] font-mono text-zinc-500 uppercase"><th className="text-left px-4 py-2">Line item</th><th className="text-right px-4 py-2">{ml}</th></tr></thead>
                        <tbody>
                         {row(`Net Sales (excl GST) · ${dayLbl}`, inp("net", dayNet))}
                          {row(`Online Sales (Swiggy+Zomato) · ${dayLbl}`, inp("online", dayOnline))}
                          <tr key="_tpldiv" className="border-t border-zinc-800"><td colSpan={2} className="px-4 pt-3 pb-1 text-[10px] font-mono text-yellow-400 uppercase tracking-widest">Today's P&amp;L · {dayLbl} (fixed cost ÷ {daysInThisMonth} days this month)</td></tr>
                          {row("Today's Total Sales", m(todayTotalSales), { bold: true })}
                          {row(`Today · Less: COGS @ ${(stCogsRate * 100).toFixed(1)}% (${stCogsSrc})`, m(todayCogs), { neg: true })}
                          {row("Today · Less: Wastage @ 5%", m(todayWastage), { neg: true })}
                          {row(`Today · Less: Commission @ ${(stCommissionRate * 100).toFixed(1)}% (online, ${stCommissionSrc})`, m(todayComm), { neg: true })}
                          {row("Today's Contribution", m(todayContrib), { bold: true })}
                          {row(`Less: Fixed cost share (monthly ÷ ${daysInThisMonth})`, m(todayFixedShare), { fixedCost: true })}
                          {row("TODAY'S NET PROFIT / (LOSS)", m(todayNetProfit), { bold: true })}
                          {row("   Today's margin %", (todayMargin * 100).toFixed(1) + "%")}
                          <tr key="_pldiv" className="border-t border-zinc-800"><td colSpan={2} className="px-4 pt-3 pb-1 text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Month-to-date P&amp;L · {ml}</td></tr>
                          {row(`Net Sales — ${ml} total ${editing ? "✏️ (whole-month override)" : ""}`, editing ? inp("mnet", Number(_moNet) || 0) : m(net))}
                          {row(`Online Sales — ${ml} total ${editing ? "✏️ (whole-month override)" : ""}`, editing ? inp("monline", Number(_moOnline) || 0) : m(online))}
                          {row("Total Sales (shop + online)", m(totalSales), { bold: true })}
                          {row(`Month · Less: COGS (food cost) @ ${(stCogsRate * 100).toFixed(1)}% of total (${stCogsSrc})`, m(cogs), { neg: true })}
                          {row("Month · Less: Wastage @ 5% of total", m(wastage), { neg: true })}
                          {row(`Month · Less: Commission @ ${(stCommissionRate * 100).toFixed(1)}% (online, ${stCommissionSrc})`, m(comm), { neg: true })}
                          {row("Contribution (before fixed)", m(contrib), { bold: true })}
                          {row("   Contribution margin %", (cMargin * 100).toFixed(1) + "%")}
                          {row("Less: Staff salaries", isSharedFixed ? <span className="text-zinc-600">0</span> : inp("staff", Number(f.staff) || 0), { fixedCost: !isSharedFixed })}
                          {row("Less: Rent", isSharedFixed ? <span className="text-zinc-600">0</span> : inp("rent", Number(f.rent) || 0), { fixedCost: !isSharedFixed })}
                          {row("Less: Electricity / EB", isSharedFixed ? <span className="text-zinc-600">0</span> : inp("eb", Number(f.eb) || 0), { fixedCost: !isSharedFixed })}
                          {row("Less: Transport", isSharedFixed ? <span className="text-zinc-600">0</span> : inp("transport", Number(f.transport) || 0), { fixedCost: !isSharedFixed })}
                          {row("Less: Repair & Maintenance (20% of rent)", m(rm), { fixedCost: true })}
                          {row("Less: Pest control", inp("pest", Number(f.pest) || 0), { fixedCost: true })}
                          {row("Less: Water", inp("water", Number(f.water) || 0), { fixedCost: true })}
                          {row("Less: Airtel / WiFi", inp("airtel", Number(f.airtel) || 0), { fixedCost: true })}
                          {row("NET PROFIT / (LOSS)", m(netProfit), { bold: true })}
                          {row("   Net margin %", (nMargin * 100).toFixed(1) + "%")}
                        </tbody>
                      </table>
                      <div className="px-4 py-3 border-t border-zinc-800">
                        <p className="font-mono text-[10px] uppercase tracking-widest text-zinc-500 mb-2">Sales needed to hit target profit</p>
                        <div className="mb-3 text-xs text-zinc-400 flex gap-4 items-center">Target A: {inp("a", ta)} &nbsp; Target B: {inp("b", tb)}</div>
                        <table className="w-full text-xs">
                          <thead><tr className="text-[10px] font-mono text-zinc-500 uppercase"><th className="text-left px-2 py-1">Goal</th><th className="text-right px-2 py-1">Same mix</th><th className="text-right px-2 py-1">Via dine-in</th></tr></thead>
                          <tbody>
                            <tr className="border-t border-zinc-800/60"><td className="px-2 py-1 text-zinc-300">Breakeven (₹0)</td><td className="px-2 py-1 text-right font-mono">{m(req(0, cmSame))}</td><td className="px-2 py-1 text-right font-mono">{m(req(0, cmDine))}</td></tr>
                            <tr className="border-t border-zinc-800/60"><td className="px-2 py-1 text-zinc-300">Profit = A ({m(ta)})</td><td className="px-2 py-1 text-right font-mono">{m(req(ta, cmSame))}</td><td className="px-2 py-1 text-right font-mono">{m(req(ta, cmDine))}</td></tr>
                           <tr className="border-t border-zinc-800/60"><td className="px-2 py-1 text-zinc-300">Profit = B ({m(tb)})</td><td className="px-2 py-1 text-right font-mono">{m(req(tb, cmSame))}</td><td className="px-2 py-1 text-right font-mono">{m(req(tb, cmDine))}</td></tr>
                          </tbody>
                        </table>
                      </div>
                    </div>
                  );
                })}
             {(canAssign || (user.outlets || []).includes(oid)) && (["BH", "CBH", "ICBH"] as const).map((brand) => { const key = oid + "_" + brand; const u = stUpload[key]; const busy = stUpBusy === key; const msg = stUpMsg[key]; const isSharedFixed = brand !== "BH"; return (
                  <div key={key} className="border border-zinc-800 bg-black/20 p-4 mt-3">
                    <p className="text-[11px] font-bold uppercase tracking-widest mb-1"><span className="text-yellow-400">{brand}</span> · 📥 Upload P&amp;L / MIS — {OUTLET_NAMES[oid] || oid}</p>
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">MIS → Net + Swiggy + Zomato · P&amp;L → {isSharedFixed ? "Pest, Water, Airtel only (shares rent/staff with BH)" : "fixed costs"} · saved for the month</p>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                      <div><label className="text-[10px] font-mono text-zinc-500 uppercase block mb-1">{brand} MIS (.xlsx)</label><input type="file" accept=".xlsx,.xls" onChange={e => { const file = e.target.files?.[0]; setStFiles(s => ({ ...s, [key]: { ...s[key], mis: file } })); }} className="text-xs text-zinc-400 w-full" /></div>
                      <div><label className="text-[10px] font-mono text-zinc-500 uppercase block mb-1">{brand} P&amp;L (.xlsx)</label><input type="file" accept=".xlsx,.xls" onChange={e => { const file = e.target.files?.[0]; setStFiles(s => ({ ...s, [key]: { ...s[key], pnl: file } })); }} className="text-xs text-zinc-400 w-full" /></div>
                    </div>
                    <button onClick={() => stExtractOutlet(oid, brand)} disabled={busy} className="bg-zinc-700 text-white font-bold text-[10px] px-4 py-2 uppercase tracking-widest disabled:opacity-50 mb-2">{busy ? "Reading..." : "Extract"}</button>
                    {msg && <p className="text-xs text-yellow-400 mb-2">{msg}</p>}
                    {u && (
                      <div className="bg-black/30 border border-zinc-800 p-3 mb-2">
                        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Found — check, then apply (red = not found)</p>
                        <div className="grid grid-cols-2 md:grid-cols-3 gap-2 text-xs">
                          {((isSharedFixed ? [["Net Sales", u.net], ["Swiggy", u.swiggy], ["Zomato", u.zomato], ["Pest", u.pest], ["Water", u.water], ["Airtel", u.airtel]] : [["Net Sales", u.net], ["Swiggy", u.swiggy], ["Zomato", u.zomato], ["Rent", u.rent], ["Staff", u.staff], ["Electricity", u.eb], ["Transport", u.transport], ["Pest", u.pest], ["Water", u.water], ["Airtel", u.airtel]]) as [string, any][]).map(([k, v]) => (
                            <div key={k} className="flex justify-between bg-black/40 px-2 py-1"><span className="text-zinc-400">{k}</span><span className={v == null ? "text-red-400" : "text-green-400 font-mono"}>{v == null ? "not found" : Math.round(v).toLocaleString("en-IN")}</span></div>
                          ))}
                        </div>
                        <button onClick={() => stApplyOutlet(oid, brand)} disabled={busy} className="bg-yellow-400 text-black font-bold text-[10px] px-4 py-2 uppercase tracking-widest disabled:opacity-50 mt-3">Apply to {brand}</button>
                      </div>
                    )}
                  </div>
                ); })}
              </div>
              );
            })}
          </div>
            );
          })()}
      {activeTab === "payout" && user && user.role === "Financial Analyst" && <PayoutTab user={user} />}
      {activeTab === "reconciliation" && user && <ReconciliationTab />}
       {activeTab === "attendance" && (
          <div>
            <div className="flex justify-between items-start mb-6 pb-5 border-b border-zinc-800">
              <div>
                <h2 className="text-2xl md:text-3xl font-black tracking-tight">People &amp; Labour</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">{user?.role === "Owner" ? "Filed by Nilani · staff count & overtime" : "Today's staff count & overtime"}</p>
              </div>
              {attendanceWeekOvertime !== null && (
                <div className="text-right">
                  <p className="text-2xl font-black">{attendanceWeekOvertime.toFixed(1)}<span className="text-sm font-normal text-zinc-500"> hrs</span></p>
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Overtime · last 7 days</p>
                </div>
              )}
            </div>
           <div className="bg-[#131316] border border-zinc-800 p-6 max-w-md">
              <div className="mb-5">
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Date</label>
               <input type="date" value={attendanceDate} onChange={(e) => { setAttendanceDate(e.target.value); const staffRef = (user?.role === "Owner" ? { id: "nilani" } : user) as Staff; if (user) fetchAttendance(staffRef, e.target.value); }} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" />
                <p className="text-[11px] font-mono text-yellow-400/80 uppercase tracking-widest mt-1.5">
                  {(() => {
                    const today = new Date().toISOString().split("T")[0];
                    const y = new Date(Date.now() - 86400000).toISOString().split("T")[0];
                    const pretty = new Date(attendanceDate + "T00:00:00").toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
                    const tag = attendanceDate === today ? "Today" : attendanceDate === y ? "Yesterday" : null;
                    return tag ? `${tag} • ${pretty}` : pretty;
                  })()}
                </p>
              </div>
              {todayAttendance ? (
                <div>
                  <p className="text-green-400 font-mono text-sm uppercase tracking-widest mb-4">✓ Submitted for today</p>
                  <div className="grid grid-cols-4 gap-3 text-center">
                    <div><p className="text-3xl font-black">{todayAttendance.present}</p><p className="text-[10px] font-mono text-zinc-500 uppercase mt-1">Present</p></div>
                    <div><p className="text-3xl font-black">{todayAttendance.absent}</p><p className="text-[10px] font-mono text-zinc-500 uppercase mt-1">Absent</p></div>
                    <div><p className="text-3xl font-black">{todayAttendance.late}</p><p className="text-[10px] font-mono text-zinc-500 uppercase mt-1">Late</p></div>
                    <div><p className="text-3xl font-black">{Number(todayAttendance.overtime_hours || 0).toFixed(1)}</p><p className="text-[10px] font-mono text-zinc-500 uppercase mt-1">OT hrs</p></div>
                  </div>
                  {(todayAttendance.absent_names || todayAttendance.late_names) && (
                    <div className="mt-4 space-y-2 text-sm">
                      {todayAttendance.absent_names && <p><span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Absent:</span> {todayAttendance.absent_names}</p>}
                      {todayAttendance.late_names && <p><span className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Late:</span> {todayAttendance.late_names}</p>}
                    </div>
                  )}
                  {user?.role !== "Owner" && (
                    <button onClick={() => { setAttendanceData({ present: String(todayAttendance.present), absent: String(todayAttendance.absent), late: String(todayAttendance.late), absent_names: todayAttendance.absent_names || "", late_names: todayAttendance.late_names || "", overtime_hours: String(todayAttendance.overtime_hours || "") }); setTodayAttendance(null); }} className="mt-5 text-[10px] font-mono text-zinc-500 uppercase tracking-widest hover:text-yellow-400">Edit</button>
                  )}
                </div>
              ) : user?.role === "Owner" ? (
                <p className="text-sm text-zinc-500">Nilani hasn't filed this date yet.</p>
              ) : (
                <div className="space-y-4">
                  {[{ k: "present", l: "Total Staff Present" }, { k: "absent", l: "Total Absent" }, { k: "late", l: "Total Late" }].map(f => (
                    <div key={f.k}>
                      <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{f.l}</label>
                      <input type="number" value={(attendanceData as any)[f.k]} onChange={(e) => setAttendanceData(prev => ({ ...prev, [f.k]: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="0" />
                    </div>
                  ))}
                  <div>
                    <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Absent — Names</label>
                    <textarea value={attendanceData.absent_names} onChange={(e) => setAttendanceData(prev => ({ ...prev, absent_names: e.target.value }))} rows={2} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="e.g. Ravi, Priya" />
                  </div>
                  <div>
                    <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Late — Names</label>
                    <textarea value={attendanceData.late_names} onChange={(e) => setAttendanceData(prev => ({ ...prev, late_names: e.target.value }))} rows={2} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="e.g. Kumar" />
                  </div>
                  <div>
                    <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Total Overtime Hours (all staff, today)</label>
                    <input type="number" step="0.5" value={attendanceData.overtime_hours} onChange={(e) => setAttendanceData(prev => ({ ...prev, overtime_hours: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="0" />
                  </div>
                  <button onClick={submitAttendance} disabled={attendanceSubmitting} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-6 py-3 hover:opacity-90 transition-opacity uppercase disabled:opacity-50">
                    {attendanceSubmitting ? "Submitting..." : "Submit Attendance →"}
                  </button>
                </div>
              )}
            </div>

            <div className="mt-6 border border-zinc-800 p-4 sm:p-5 max-w-2xl">
              <p className="text-sm font-semibold mb-1">Staff Registry</p>
              <p className="text-[11px] text-zinc-500 mb-4">{user?.role === "Owner" ? "Everyone Nilani has registered, by department." : "Register everyone by department — Production, Sales & Outlet Staff, Housekeeping."}</p>
              <div className="space-y-5">
                {LABOUR_DEPARTMENTS.map((dept) => {
                  const people = kLaborStaff.filter((s) => s.department === dept);
                  return (
                    <div key={dept} className="border-t border-zinc-900 pt-4 first:border-t-0 first:pt-0">
                      <p className="text-[11px] font-mono text-yellow-400 uppercase tracking-widest mb-2">{dept} <span className="text-zinc-600">· {people.length}</span></p>
                      {people.length === 0 ? (
                        <p className="text-xs text-zinc-600 mb-2">None registered yet.</p>
                      ) : (
                        <div className="space-y-1.5 mb-3">
                          {people.map((p) => (
                            <div key={p.id} className="flex items-center gap-2">
                              {kLaborEditId === p.id ? (
                                <>
                                  <input type="text" value={kLaborEditName} onChange={(e) => setKLaborEditName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && saveLaborStaffEdit()} className="flex-1 bg-black border border-yellow-400 text-white px-2 py-1.5 focus:outline-none text-sm" autoFocus />
                                  <button onClick={saveLaborStaffEdit} className="text-[11px] font-mono uppercase text-green-400 hover:text-green-300 px-2">Save</button>
                                  <button onClick={() => { setKLaborEditId(null); setKLaborEditName(""); }} className="text-[11px] font-mono uppercase text-zinc-500 hover:text-white px-2">Cancel</button>
                                </>
                              ) : (
                                <>
                                  <span className="flex-1 text-sm">{p.name}</span>
                                  {user?.role !== "Owner" && (
                                    <>
                                      <button onClick={() => { setKLaborEditId(p.id); setKLaborEditName(p.name); }} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-yellow-400 px-2">Edit</button>
                                      <button onClick={() => deleteLaborStaff(p.id, p.name)} className="text-[11px] font-mono uppercase text-zinc-400 hover:text-red-500 px-2">Delete</button>
                                    </>
                                  )}
                                </>
                              )}
                            </div>
                          ))}
                        </div>
                      )}
                      {user?.role !== "Owner" && (
                        <div className="flex gap-2">
                          <input type="text" value={kLaborNewName[dept] || ""} onChange={(e) => setKLaborNewName((prev) => ({ ...prev, [dept]: e.target.value }))} onKeyDown={(e) => e.key === "Enter" && addLaborStaff(dept)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm flex-1" placeholder={`Add a name (e.g. ${dept === "Production" ? "Suresh" : dept === "Housekeeping" ? "Mala" : "Priya"})`} />
                          <button onClick={() => addLaborStaff(dept)} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors shrink-0">Add</button>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
        {(activeTab === "my_report" || activeTab === "all_reports") && (
      
          <div>
           <div className="flex justify-between items-end mb-6 pb-5 border-b border-zinc-800">
        <div>
    <h2 className="text-2xl font-black tracking-tight">{canAssign ? "All Reports" : "Daily Report"}</h2>
    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">
      {canAssign ? "Staff submissions overview" : `Due by ${ALL_STAFF.find(s => s.id === user.id)?.report_time || "--:--"} daily`}
    </p>
  </div>
  {hasReportDuty && (
    <input
      type="date"
      value={reportHistoryDate}
      onChange={(e) => { setReportHistoryDate(e.target.value); fetchReportByDate(e.target.value); }}
      className="bg-black border border-zinc-800 text-white px-4 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors font-mono text-sm"
    />
  )}
</div>
           {(activeTab === "my_report" || activeTab === "all_reports") && hasReportDuty && (
  <div className="mb-8">
    {reportHistoryDate !== new Date().toISOString().split("T")[0] ? (
      <div className="bg-[#131316] border border-zinc-800 p-6">
        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">
          Report for {new Date(reportHistoryDate).toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" })}
        </p>
        {reportByDate ? (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {reportFields.map(f => (
              <div key={f.key} className="bg-black/30 px-3 py-2">
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                <p className="text-sm text-white mt-1">{reportByDate.report_data?.[f.key] || "—"}</p>
              </div>
            ))}
          </div>
       ) : (
          <div>
            <p className="text-[10px] font-mono text-yellow-400 uppercase tracking-widest mb-4">Back-dated entry — counts as -5</p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
              {reportFields.map(f => (
                <div key={f.key}>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{f.label}</label>
                  {reportInput(f)}
                </div>
              ))}
            </div>
            <button onClick={submitReport} disabled={reportSubmitting} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-6 py-3 hover:opacity-90 transition-opacity uppercase disabled:opacity-50">
              {reportSubmitting ? "Submitting..." : "Submit Back-dated Report →"}
            </button>
          </div>
        )}
    ) : null
  </div>

    ) : todayReport ? (
  <div className="bg-green-400/5 border border-green-400/30 p-6">
    <div className="flex items-center justify-between mb-4">
      <div className="flex items-center gap-3">
        <span className="text-green-400 font-mono text-xs uppercase tracking-widest">✓ Today's report submitted</span>
        {todayReport.is_late && <span className="text-red-500 font-mono text-[10px] uppercase bg-red-500/10 px-2 py-0.5">Late</span>}
      </div>
      <button
        onClick={async () => {
          if (!user) return;
          const data: Record<string, string> = {};
          if (todayReport.report_data) {
            reportFields.forEach(f => { data[f.key] = todayReport.report_data[f.key] || ""; });
          }
          await supabase.from("reports").delete().eq("id", todayReport.id);
          setTodayReport(null);
          setReportData(data);
        }}
        className="font-mono text-[10px] uppercase tracking-widest px-3 py-1.5 border border-zinc-700 hover:border-yellow-400 hover:text-yellow-400 transition-colors"
      >
        ✏ Edit
      </button>
    </div>
                    {todayReport.report_data && (
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                        {reportFields.map(f => (
                          <div key={f.key} className="bg-black/30 px-3 py-2">
                            <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                            <p className="text-sm text-white mt-1">{todayReport.report_data[f.key] || "—"}</p>
                          </div>
                        ))}
                      </div>
                    )}
                    <p className="text-[11px] font-mono text-zinc-600 mt-4">{new Date(todayReport.submitted_at).toLocaleString("en-IN")}</p>
                  </div>
                ) : (
                  <div className="bg-[#131316] border border-zinc-800 p-6">
                    <div className="flex items-center justify-between mb-6">
                      <p className="text-sm font-bold uppercase tracking-widest">Submit Today's Report</p>
                      <span className="text-yellow-400 font-mono text-xs">Due: {ALL_STAFF.find(s => s.id === user.id)?.report_time}</span>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
                      {reportFields.map(f => (
                        <div key={f.key}>
                          <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{f.label}</label>
                          {reportInput(f)}
                        </div>
                      ))}
                    </div>
                    <button onClick={submitReport} disabled={reportSubmitting} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-6 py-3 hover:opacity-90 transition-opacity uppercase disabled:opacity-50">
                      {reportSubmitting ? "Submitting..." : "Submit Report →"}
                    </button>
                  </div>
                )}
              </div>
            )}
            <div className="bg-[#131316] border border-zinc-800">
              {reports.length === 0 ? (
                <div className="p-10 text-center">
                  <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">No reports yet</p>
                </div>
              ) : reports.map((r) => {
                const staffName = ALL_STAFF.find(s => s.id === r.staff_id)?.name || FORMER_STAFF[r.staff_id] || r.staff_id;
                const staffFields = REPORT_FIELDS[r.staff_id] || [];
                return (
                  <div key={r.id} className="border-b border-zinc-800 last:border-0">
                    <div className="flex items-center justify-between px-6 py-4 cursor-pointer hover:bg-zinc-900 transition-colors" onClick={() => setSelectedReport(selectedReport?.id === r.id ? null : r)}>
                      <div className="flex items-center gap-3">
                        <span className="font-semibold text-sm">{canAssign ? staffName : "My Report"}</span>
                        {r.is_late && <span className="text-red-500 font-mono text-[10px] uppercase bg-red-500/10 px-2 py-0.5">Late</span>}
                      </div>
                      <div className="flex items-center gap-3">
                        <span className="text-[11px] font-mono text-zinc-500">{new Date(r.submitted_at).toLocaleString("en-IN", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
                        <span className="text-zinc-500">{selectedReport?.id === r.id ? "▲" : "▼"}</span>
                      </div>
                    </div>
                    {selectedReport?.id === r.id && r.report_data && (
                      <div className="px-6 pb-6 grid grid-cols-1 md:grid-cols-2 gap-3">
                        {staffFields.map(f => (
                          <div key={f.key} className="bg-black/30 px-3 py-2">
                            <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                            <p className="text-sm text-white mt-1">{r.report_data[f.key] || "—"}</p>
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {activeTab === "owner_outlets" && (
  <div>
   <div className="flex justify-between items-end mb-6 pb-5 border-b border-zinc-800">
  <div>
    <h2 className="text-2xl font-black tracking-tight">Outlet Reports</h2>
    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">{canAssign ? "All 12 outlets" : `Your ${(user.outlets || []).length} outlets`} — daily tracker</p>
  </div>
  <input
    type="date"
    value={historyDate}
   onChange={(e) => { setHistoryDate(e.target.value); fetchAllOutletReports(e.target.value); }}
    className="bg-black border border-zinc-800 text-white px-4 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors font-mono text-sm"
  />
</div>
    <div className="bg-[#131316] border border-zinc-800 p-5 mb-6">
    <div className="flex items-center justify-between mb-1">
      <p className="text-sm font-bold uppercase tracking-widest">📊 Outlet Health</p>
      <div className="flex gap-2">
        <button onClick={downloadOutletHealthExcel} disabled={outletHealthXlsxBusy || !outletDeepDive} className="px-3 py-1.5 bg-green-600 hover:bg-green-500 text-white transition-colors disabled:opacity-50 text-[10px] font-mono uppercase tracking-widest">
          {outletHealthXlsxBusy ? "…" : "↓ Excel"}
        </button>
        <button onClick={downloadOutletHealthPDF} disabled={outletHealthPdfBusy} className="px-3 py-1.5 bg-yellow-400 hover:bg-yellow-300 text-black transition-colors disabled:opacity-50 text-[10px] font-mono uppercase tracking-widest">
          {outletHealthPdfBusy ? "…" : "↓ PDF"}
        </button>
      </div>
    </div>
       <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Pick an outlet, optionally compare two periods · full breakdown, trend &amp; verdict inside the download</p>
    <div className="mb-4">
      <div className="flex flex-wrap gap-2">
        {OUTLETS.map((o) => { const on = outletHealthSel === o; return (
          <button key={o} onClick={() => setOutletHealthSel(o)} className={`text-[11px] px-3 py-1.5 border font-mono uppercase tracking-wide transition-colors ${on ? "bg-yellow-400 text-black border-yellow-400" : "bg-black text-zinc-400 border-zinc-800 hover:border-zinc-600"}`}>{OUTLET_NAMES[o] || o}</button>
        ); })}
      </div>
    </div>
    <div className="mb-2">    

    
      <label className="text-[10px] font-mono text-zinc-500 uppercase block mb-2">Compare two periods (optional — leave as "none" to skip)</label>
      <div className="flex flex-wrap items-center gap-4">
        <div className="flex items-center gap-2">
          <span className="text-[9px] text-zinc-600">A:</span>
          <select value={comparePresetA} onChange={(e) => applyComparePreset("A", e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400">
            <option value="none">None</option>
            {COMPARE_PRESETS.map(p => <option key={p.id} value={p.id}>{p.label}</option>)}
          </select>
          {comparePresetA === "custom" && (
            <>
              <input type="date" value={compareAFrom} onChange={(e) => setCompareAFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
              <span className="text-zinc-600 text-xs">to</span>
              <input type="date" value={compareATo} onChange={(e) => setCompareATo(e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
            </>
          )}
        </div>
        <div className="flex items-center gap-2">
          <span className="text-[9px] text-zinc-600">B:</span>
          <select value={comparePresetB} onChange={(e) => applyComparePreset("B", e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400">
            <option value="none">None</option>
            {COMPARE_PRESETS.map(p => <option key={p.id} value={p.id}>{p.label}</option>)}
          </select>
          {comparePresetB === "custom" && (
            <>
              <input type="date" value={compareBFrom} onChange={(e) => setCompareBFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
              <span className="text-zinc-600 text-xs">to</span>
              <input type="date" value={compareBTo} onChange={(e) => setCompareBTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
            </>
          )}
        </div>
        <button onClick={runQuickCompare} disabled={quickCompareBusy} className="px-3 py-1.5 border border-zinc-700 text-zinc-300 hover:border-yellow-400 hover:text-yellow-400 transition-colors text-[10px] font-mono uppercase tracking-widest disabled:opacity-50">{quickCompareBusy ? "…" : "Compare"}</button>
      </div>
      {quickCompareResult && (
        <p className="text-xs text-zinc-400 mt-2 font-mono">
          A: ₹{Math.round(quickCompareResult.aTotal).toLocaleString("en-IN")} → B: ₹{Math.round(quickCompareResult.bTotal).toLocaleString("en-IN")}
          <span className={`ml-2 font-semibold ${quickCompareResult.totalChg == null ? "text-zinc-600" : quickCompareResult.totalChg >= 0 ? "text-green-400" : "text-red-400"}`}>{quickCompareResult.totalChg == null ? "—" : `${quickCompareResult.totalChg >= 0 ? "▲" : "▼"} ${Math.abs(quickCompareResult.totalChg).toFixed(1)}%`}</span>
          <span className="text-zinc-600 ml-2">· Shop {quickCompareResult.shopChg == null ? "—" : `${quickCompareResult.shopChg >= 0 ? "▲" : "▼"}${Math.abs(quickCompareResult.shopChg).toFixed(0)}%`} · Swiggy {quickCompareResult.swiggyChg == null ? "—" : `${quickCompareResult.swiggyChg >= 0 ? "▲" : "▼"}${Math.abs(quickCompareResult.swiggyChg).toFixed(0)}%`} · Zomato {quickCompareResult.zomatoChg == null ? "—" : `${quickCompareResult.zomatoChg >= 0 ? "▲" : "▼"}${Math.abs(quickCompareResult.zomatoChg).toFixed(0)}%`}</span>
        </p>
      )}
    </div>
  </div>

  <div className="bg-[#131316] border border-zinc-800 p-5 mb-6">
    <p className="text-sm font-bold uppercase tracking-widest mb-1">📥 Download custom report</p>
    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Pick a date range + outlets · Excel or PDF · daily rows + summary</p>
    <div className="mb-4">
      <div className="flex items-center justify-between mb-2">
        <label className="text-[10px] font-mono text-zinc-500 uppercase">Outlets {repOutlets.length === 0 ? "(all)" : `(${repOutlets.length})`}</label>
        <div className="flex gap-2">
          <button onClick={() => setRepOutlets([...OUTLETS])} className="text-[10px] font-mono text-yellow-400 uppercase">All</button>
          <button onClick={() => setRepOutlets([])} className="text-[10px] font-mono text-zinc-500 uppercase">Clear</button>
        </div>
      </div>
      <div className="flex flex-wrap gap-2">
        {(canAssign ? OUTLETS : (user.outlets || [])).map(o => { const on = repOutlets.includes(o); return (
          <button key={o} onClick={() => setRepOutlets(on ? repOutlets.filter(x => x !== o) : [...repOutlets, o])} className={`text-[11px] px-3 py-1.5 border font-mono uppercase tracking-wide transition-colors ${on ? "bg-yellow-400 text-black border-yellow-400" : "bg-black text-zinc-400 border-zinc-800 hover:border-zinc-600"}`}>{OUTLET_NAMES[o] || o}</button>
        ); })}
      </div>
    </div>
    <div className="mb-4">
      <label className="text-[10px] font-mono text-zinc-500 uppercase block mb-2">Range</label>
      <div className="flex flex-wrap items-center gap-2">
        <select value={repQuickRange} onChange={e => applyQuickRange(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm font-mono focus:outline-none focus:border-yellow-400">
          {RANGE_PRESETS.map(p => <option key={p.id} value={p.id}>{p.label}</option>)}
        </select>
        {repQuickRange === "custom" && !repCustomizePerOutlet && (
          <>
            <input type="date" value={repQuickFrom} onChange={e => applyQuickCustom(e.target.value, repQuickTo)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
            <span className="text-zinc-600 text-xs">to</span>
            <input type="date" value={repQuickTo} onChange={e => applyQuickCustom(repQuickFrom, e.target.value)} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
          </>
        )}
        <button onClick={() => setRepCustomizePerOutlet(v => !v)} className="text-[11px] font-mono text-yellow-400 uppercase underline ml-1">{repCustomizePerOutlet ? "Use one range for all" : "Customize per outlet"}</button>
      </div>
      {repCustomizePerOutlet && (
        <div className="space-y-2 mt-3">
          {activeRepOutlets().map(o => {
            const sel = getOutletSel(o);
            return (
              <div key={o} className="flex flex-wrap items-center gap-2">
                <span className="text-xs font-mono w-28 shrink-0 text-zinc-300">{OUTLET_NAMES[o] || o}</span>
                <select value={sel.preset} onChange={e => setOutletRangeSel({ ...outletRangeSel, [o]: { ...sel, preset: e.target.value } })} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400">
                  {RANGE_PRESETS.map(p => <option key={p.id} value={p.id}>{p.label}</option>)}
                </select>
                {sel.preset === "custom" && (
                  <>
                    <input type="date" value={sel.from || ""} onChange={e => setOutletRangeSel({ ...outletRangeSel, [o]: { ...sel, from: e.target.value } })} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
                    <span className="text-zinc-600 text-xs">to</span>
                    <input type="date" value={sel.to || ""} onChange={e => setOutletRangeSel({ ...outletRangeSel, [o]: { ...sel, to: e.target.value } })} className="bg-black border border-zinc-800 text-white px-2 py-1.5 text-xs font-mono focus:outline-none focus:border-yellow-400" />
                  </>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
    <div className="flex gap-3">
      <button onClick={downloadRangeExcel} disabled={repBusy} className="bg-green-600 text-white font-bold text-xs px-5 py-2.5 uppercase tracking-widest disabled:opacity-50 hover:opacity-90">{repBusy ? "Working…" : "⬇ Excel"}</button>
      <button onClick={downloadRangePDF} disabled={repBusy} className="bg-yellow-400 text-black font-bold text-xs px-5 py-2.5 uppercase tracking-widest disabled:opacity-50 hover:opacity-90">{repBusy ? "Working…" : "⬇ PDF"}</button>
    </div>
  </div>
   {(() => {
      const cols = ["#FACC15", "#FB923C", "#EF4444"];
      const chLabels = [["Shop", cols[0]], ["Swiggy", cols[1]], ["Zomato", cols[2]]] as [string, string][];
      return (
        <div className="mb-6">
          <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Tap an outlet to open its full report ↓</p>
        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3">
            {(canAssign ? OUTLETS : (user.outlets || [])).map(o => {
              const r = allOutletReports.find(x => x.outlet_id === o);
              const shop = r ? Number(r.shop_sales_value) || 0 : 0;
              const sw = r ? Number(r.swiggy_sales_value) || 0 : 0;
              const zo = r ? Number(r.zomato_sales_value) || 0 : 0;
              const tot = shop + sw + zo;
              const filed = !!r;
              const sel = expandedOutlet === o;
              const R = 34, CX = 44, CY = 44, SW = 12, CIRC = 2 * Math.PI * R;
              let acc = 0;
              const segs = tot > 0 ? [shop, sw, zo].map((v, i) => { const frac = v / tot; const len = frac * CIRC; const off = -acc * CIRC; acc += frac; return { len, gap: CIRC - len, off, col: cols[i] }; }) : [];
              return (
                <div key={o} onClick={() => setExpandedOutlet(sel ? null : o)} className={`cursor-pointer bg-[#131316] border ${sel ? "border-yellow-400" : filed ? "border-green-400/30" : "border-zinc-800"} p-3 hover:border-yellow-400/50 transition-colors`}>
                  <div className="flex items-center gap-3">
                    <svg width="56" height="56" viewBox="0 0 88 88" className="shrink-0">
                      <circle cx={CX} cy={CY} r={R} fill="none" stroke="#27272a" strokeWidth={SW} />
                      {segs.map((s, i) => (<circle key={i} cx={CX} cy={CY} r={R} fill="none" stroke={s.col} strokeWidth={SW} strokeDasharray={`${s.len} ${s.gap}`} strokeDashoffset={s.off} transform={`rotate(-90 ${CX} ${CY})`} />))}
                    </svg>
                    <div className="min-w-0">
                      <p className="text-xs font-bold text-white truncate">{OUTLET_NAMES[o] || o}</p>
                      <p className="text-[11px] font-mono text-zinc-400">{filed ? `₹${(tot / 1000).toFixed(1)}k` : "—"}</p>
                      <p className={`text-[9px] font-mono uppercase tracking-widest ${!filed ? "text-zinc-600" : r.is_late ? "text-red-400" : "text-green-400"}`}>{!filed ? "not filed" : r.is_late ? "late" : "on time"}</p>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
          <div className="flex gap-4 mt-3 text-[10px] font-mono text-zinc-500">
            {chLabels.map(([n, c]) => (<span key={n} className="flex items-center gap-1"><span className="w-2.5 h-2.5 rounded-sm" style={{ background: c }}></span>{n}</span>))}
          </div>
        </div>
      );
    })()}
   <div className="grid grid-cols-1 gap-4">
      {!expandedOutlet && <p className="text-center text-sm text-zinc-600 py-10">Tap an outlet donut above to open its full report.</p>}
      {OUTLETS.filter(o => o === expandedOutlet).map(o => {
        const report = allOutletReports.find(r => r.outlet_id === o);
      const manager = ALL_STAFF.find(s => (s.outlets as string[]).includes(o));
        return (
          <div key={o} className={`bg-[#131316] border ${report ? "border-green-400/30" : "border-zinc-800"} p-5`}>
            <div className="flex items-center justify-between mb-4">
              <div>
               <p className="font-bold text-sm uppercase tracking-widest">{OUTLET_NAMES[o]}</p>
                <p className="text-[10px] font-mono text-zinc-500 mt-0.5">{manager?.name || "—"}</p>
              </div>
              {report ? (
                <span className="font-mono text-[10px] uppercase tracking-widest px-2 py-1 bg-green-400/10 text-green-400">
                 ✓ Submitted {report.is_late ? "· Late" : "· On Time"}{report.is_edited ? " · Edited" : ""}
                 </span>
              ) : (
                <span className="font-mono text-[10px] uppercase tracking-widest px-2 py-1 bg-yellow-400/10 text-yellow-400">Pending</span>
              )}
            </div>
            {report && (
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                {[
                  { label: "Shop Sales", value: `₹${report.shop_sales_value} (${report.shop_sales_count})` },
                { label: "Shop AOV", value: report.shop_sales_count > 0 ? `₹${Math.round(Number(report.shop_sales_value) / Number(report.shop_sales_count))}` : "—" },
                { label: "Swiggy", value: `₹${report.swiggy_sales_value} (${report.swiggy_sales_count})` },
                { label: "Swiggy AOV", value: report.swiggy_sales_count > 0 ? `₹${Math.round(Number(report.swiggy_sales_value) / Number(report.swiggy_sales_count))}` : "—" },
                { label: "Zomato", value: `₹${report.zomato_sales_value} (${report.zomato_sales_count})` },
                { label: "Zomato AOV", value: report.zomato_sales_count > 0 ? `₹${Math.round(Number(report.zomato_sales_value) / Number(report.zomato_sales_count))}` : "—" },
                { label: "Total Sales", value: `₹${Number(report.shop_sales_value) + Number(report.swiggy_sales_value) + Number(report.zomato_sales_value)}` },
                { label: "Total AOV", value: (() => { const tv = Number(report.shop_sales_value) + Number(report.swiggy_sales_value) + Number(report.zomato_sales_value); const tc = Number(report.shop_sales_count) + Number(report.swiggy_sales_count) + Number(report.zomato_sales_count); return tc > 0 ? `₹${Math.round(tv/tc)}` : "—"; })() },
                ...brandSalesItems(report),
                 { label: "Target", value: `₹${report.target}` },
                  { label: "Swiggy Live", value: report.swiggy_live ? "✓ Yes" : "✗ No", color: report.swiggy_live ? "text-green-400" : "text-red-500" },
                  { label: "Zomato Live", value: report.zomato_live ? "✓ Yes" : "✗ No", color: report.zomato_live ? "text-green-400" : "text-red-500" },
                  { label: "Discount", value: report.discount_running || "—" },
                  { label: "Expiry", value: report.expiry_count > 0 ? `${report.expiry_count} items` : "None", color: report.expiry_count > 0 ? "text-red-500" : "" },
                  { label: "Complimentary", value: report.complimentary_count > 0 ? `${report.complimentary_count}` : "None" },
                  { label: "Issues", value: report.issues || "—", color: report.issues ? "text-yellow-400" : "" },
                  { label: "Action Taken", value: report.action_taken || "—" },
                ].map(f => (
                  <div key={f.label} className="bg-black/30 px-3 py-2">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                    <p className={`text-sm mt-1 ${f.color || "text-white"}`}>{f.value}</p>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  </div>
)}
{activeTab === "outlet_reports" && (
  <div>
    <div className="flex justify-between items-end mb-6 pb-5 border-b border-zinc-800">
  <div>
    <h2 className="text-2xl font-black tracking-tight">Outlet Reports</h2>
    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Daily tracker — fill for each outlet</p>
  </div>
  <input
    type="date"
    value={outletHistoryDate}
    onChange={(e) => { setOutletHistoryDate(e.target.value); setOutletEntryDate(e.target.value); setOutletWasOff(false); fetchOutletReportsByDate(e.target.value); }}
    className="bg-black border border-zinc-800 text-white px-4 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors font-mono text-sm"
  />
</div>

    {/* Outlet manager review — last full week + month to date, own outlets only */}
    {!canAssign && ((user.outlets || []) as string[]).length > 0 && (() => {
      const D = ovDates();
      const outs = (user.outlets || []) as string[];
      const tot = (r: any) => (Number(r.shop_sales_value) || 0) + (Number(r.swiggy_sales_value) || 0) + (Number(r.zomato_sales_value) || 0);
      const inRange = (o: string, a: string, b: string) => { const seen = new Set<string>(); return ovRows.filter((r) => { if (r.outlet_id !== o || r.report_date < a || r.report_date > b || seen.has(r.report_date)) return false; seen.add(r.report_date); return true; }); };
      const money = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
      const fmtDay = (reportDate: string) => new Date(ovAddDays(reportDate, -1) + "T00:00:00Z").toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", timeZone: "UTC" });
      const entriesOf = (rows: any[], f: "expiry" | "comp") => f === "expiry"
        ? rows.filter((r) => (Number(r.expiry_count) || 0) > 0 || !_isBlank(r.expiry_items))
        : rows.filter((r) => (Number(r.complimentary_count) || 0) > 0 || !_isBlank(r.complimentary_items) || !_isBlank(r.complimentary_reason));
      const prodOf = (rows: any[], f: "expiry" | "comp") => { const e = entriesOf(rows, f); return { days: e.length, top: wastageProductBreakdown(e, f === "expiry" ? "expiry_items" : "complimentary_items").slice(0, 3) }; };
      const verdict = (p: number | null) => p === null ? { t: "No data", c: "text-zinc-500 border-zinc-700" } : p >= 100 ? { t: "On target", c: "text-green-400 border-green-400/50" } : p >= 85 ? { t: "Slightly behind", c: "text-yellow-400 border-yellow-400/50" } : { t: "Behind", c: "text-red-400 border-red-400/50" };
      const avgRating = (rows: any[]) => { const v = rows.map((r) => Number(r.bh_google_rating)).filter((x) => x > 0); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
      const wkLabel = `${fmtDay(D.wkFrom)} – ${fmtDay(D.wkTo)}`;
      const ProdLine = ({ label, p }: { label: string; p: { days: number; top: any[] } }) => (
        <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">{label}:</span> {p.days === 0 ? "none reported" : `${p.days} day${p.days > 1 ? "s" : ""}${p.top.length ? " — " + p.top.map((t: any) => `${t.name} (${t.days}d)`).join(", ") : ""}`}</p>
      );
      return (
        <div className="mb-8 border border-zinc-800 p-4">
          <div className="flex justify-between items-start gap-3 flex-wrap mb-3">
            <div>
              <p className="text-sm font-semibold">Your outlets — review</p>
              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mt-0.5">{ovView === "week" ? `Sales week ${wkLabel} · updates every Monday` : `${new Date(D.today + "T00:00:00Z").toLocaleDateString("en-IN", { month: "long", year: "numeric", timeZone: "UTC" })} · month to date`}</p>
            </div>
            <div className="flex gap-1">
              {([["week", "Last week"], ["mtd", "Month to date"]] as const).map(([k, l]) => (
                <button key={k} onClick={() => setOvView(k)} className={`font-mono text-[10px] uppercase tracking-widest px-3 py-1.5 border ${ovView === k ? "border-yellow-400 text-yellow-400" : "border-zinc-700 text-zinc-500"}`}>{l}</button>
              ))}
            </div>
          </div>
          {ovLoading ? <p className="text-xs text-zinc-500">Loading…</p> : (
            <div className="divide-y divide-zinc-800 border border-zinc-800">
              {outs.map((o) => {
                const name = OUTLET_NAMES[o] || o.replace(/_/g, " ");
                const key = ovView + o;
                const open = ovOpen === key;
                if (ovView === "week") {
                  const wk = inRange(o, D.wkFrom, D.wkTo), pv = inRange(o, D.prevFrom, D.prevTo);
                  const sales = wk.reduce((a, r) => a + tot(r), 0), pSales = pv.reduce((a, r) => a + tot(r), 0);
                  const n = wk.length, pn = pv.length;
                  const daily = dailyTargetFor(o, ovAddDays(D.lastMonday, -3).slice(0, 7));
                  const pct = daily > 0 && n > 0 ? sales / (daily * n) * 100 : null;
                  const avg = n ? sales / n : 0, pAvg = pn ? pSales / pn : 0;
                  const change = pAvg > 0 && n > 0 ? (avg - pAvg) / pAvg * 100 : null;
                  const v = verdict(pct);
                  const sorted = [...wk].sort((a, b) => tot(b) - tot(a));
                  const mix = (f: string) => sales > 0 ? Math.round(wk.reduce((a, r) => a + (Number(r[f]) || 0), 0) / sales * 100) : 0;
                  const rNow = avgRating(wk), rPrev = avgRating(pv);
                  return (
                    <div key={key} className="px-3 py-2.5">
                      <div onClick={() => setOvOpen(open ? null : key)} className="flex items-center gap-3 cursor-pointer flex-wrap">
                        <span className="text-xs font-semibold w-28">{name}</span>
                        <span className="text-xs font-mono">{n ? money(sales) : "—"}</span>
                        <span className={`text-[10px] font-mono uppercase tracking-widest px-2 py-0.5 border ${v.c}`}>{v.t}{pct !== null ? ` · ${Math.round(pct)}%` : ""}</span>
                        {change !== null && <span className={`text-[11px] font-mono ${change >= 0 ? "text-green-400" : "text-red-400"}`}>{change >= 0 ? "▲" : "▼"} {Math.abs(Math.round(change))}% vs prev week</span>}
                        <span className="ml-auto text-zinc-600 text-xs">{open ? "−" : "+"}</span>
                      </div>
                      {open && (
                        <div className="mt-2 pl-1 space-y-1">
                          <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Reports filed:</span> {n} of 7 days{n < 7 ? " (missing days are not counted against the target)" : ""}</p>
                          <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Target:</span> {daily > 0 ? `${money(daily)}/day (monthly target ÷ 30) · this week ${money(daily * n)} for ${n} day${n === 1 ? "" : "s"}` : "not set"}</p>
                          {pn > 0 && <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Previous week:</span> {money(pSales)} over {pn} day{pn === 1 ? "" : "s"} ({money(pAvg)}/day)</p>}
                          {sorted.length > 1 && <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Best day:</span> {fmtDay(sorted[0].report_date)} {money(tot(sorted[0]))} · <span className="text-zinc-500">Weakest:</span> {fmtDay(sorted[sorted.length - 1].report_date)} {money(tot(sorted[sorted.length - 1]))}</p>}
                          {sales > 0 && <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Channel mix:</span> Shop {mix("shop_sales_value")}% · Swiggy {mix("swiggy_sales_value")}% · Zomato {mix("zomato_sales_value")}%</p>}
                          <ProdLine label={`Expiry (prev week: ${prodOf(pv, "expiry").days} day${prodOf(pv, "expiry").days === 1 ? "" : "s"})`} p={prodOf(wk, "expiry")} />
                          <ProdLine label={`Complimentary (prev week: ${prodOf(pv, "comp").days} day${prodOf(pv, "comp").days === 1 ? "" : "s"})`} p={prodOf(wk, "comp")} />
                          {rNow !== null && <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">BH Google rating (avg):</span> ⭐ {rNow.toFixed(2)}{rPrev !== null ? ` (prev week ${rPrev.toFixed(2)})` : ""}</p>}
                        </div>
                      )}
                    </div>
                  );
                }
                // Month to date
                const mr = inRange(o, D.monthStart, D.today);
                const n = mr.length, sales = mr.reduce((a, r) => a + tot(r), 0);
                const ym = D.today.slice(0, 7);
                const mt = monthlyTargetFor(o, ym);
                const dim = new Date(Number(ym.slice(0, 4)), Number(ym.slice(5, 7)), 0).getDate();
                const expected = mt * n / dim;
                const pace = expected > 0 ? sales / expected * 100 : null;
                const achieved = mt > 0 ? sales / mt * 100 : null;
                const left = Math.max(1, dim - n);
                const need = mt > sales ? (mt - sales) / left : 0;
                const v = verdict(pace);
                return (
                  <div key={key} className="px-3 py-2.5">
                    <div onClick={() => setOvOpen(open ? null : key)} className="flex items-center gap-3 cursor-pointer flex-wrap">
                      <span className="text-xs font-semibold w-28">{name}</span>
                      <span className="text-xs font-mono">{n ? money(sales) : "—"}{mt > 0 ? ` / ${money(mt)}` : ""}</span>
                      <span className={`text-[10px] font-mono uppercase tracking-widest px-2 py-0.5 border ${v.c}`}>{v.t}{pace !== null ? ` · ${Math.round(pace)}% of pace` : ""}</span>
                      <span className="ml-auto text-zinc-600 text-xs">{open ? "−" : "+"}</span>
                    </div>
                    {mt > 0 && (
                      <div className="mt-2 h-1.5 bg-zinc-800"><div className={`h-1.5 ${pace !== null && pace >= 100 ? "bg-green-400" : pace !== null && pace >= 85 ? "bg-yellow-400" : "bg-red-400"}`} style={{ width: `${Math.min(100, achieved || 0)}%` }} /></div>
                    )}
                    {open && (
                      <div className="mt-2 pl-1 space-y-1">
                        <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Achieved so far:</span> {achieved !== null ? `${achieved.toFixed(1)}% of the monthly target` : "no target set"} · {n} report{n === 1 ? "" : "s"} filed</p>
                        {expected > 0 && <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">Should be at by now:</span> {money(expected)} ({n}/{dim} days of the month reported)</p>}
                        <p className="text-[11px] text-zinc-400"><span className="text-zinc-500">To reach the target:</span> {mt <= 0 ? "—" : need === 0 ? "target reached 🎉" : `about ${money(need)} per remaining day (${left} days)`}</p>
                        <ProdLine label="Expiry this month" p={prodOf(mr, "expiry")} />
                        <ProdLine label="Complimentary this month" p={prodOf(mr, "comp")} />
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
          <p className="text-[10px] text-zinc-600 mt-2">Weekly target = monthly target ÷ 30 per day. "Pace" compares sales with the share of the month already reported. Reports are counted by the date they were filed.</p>
        </div>
      );
    })()}

    {/* Outlet selector tabs */}
    <div className="flex gap-2 flex-wrap mb-6">
     {(canAssign ? OUTLETS : (user.outlets || [])).map(o => {
        const submitted = !!outletReports[o];
        return (
         <button key={o} onClick={() => { setActiveOutlet(o); const lastRatings = lastOutletRatings[o]; setOutletReportData({ target: String(dailyTargetFor(o, new Date().toISOString().slice(0, 7))), bh_google_rating: lastRatings ? String(lastRatings.bh_google_rating || "") : "", bh_swiggy_rating: lastRatings ? String(lastRatings.bh_swiggy_rating || "") : "", bh_zomato_rating: lastRatings ? String(lastRatings.bh_zomato_rating || "") : "", cbh_google_rating: lastRatings ? String(lastRatings.cbh_google_rating || "") : "", cbh_swiggy_rating: lastRatings ? String(lastRatings.cbh_swiggy_rating || "") : "", cbh_zomato_rating: lastRatings ? String(lastRatings.cbh_zomato_rating || "") : "", icbh_google_rating: lastRatings ? String(lastRatings.icbh_google_rating || "") : "", icbh_swiggy_rating: lastRatings ? String(lastRatings.icbh_swiggy_rating || "") : "", icbh_zomato_rating: lastRatings ? String(lastRatings.icbh_zomato_rating || "") : "" }); }}
            className={`font-mono text-[10px] uppercase tracking-widest px-4 py-2 border transition-colors relative ${activeOutlet === o ? "border-yellow-400 text-yellow-400" : "border-zinc-700 text-zinc-500 hover:border-zinc-500"}`}>
           {OUTLET_NAMES[o] || o.replace(/_/g, " ")}
            {submitted && <span className="ml-2 text-green-400">✓</span>}
          </button>
        );
      })}
    </div>

    {!activeOutlet && (
      <div className="bg-[#131316] border border-zinc-800 p-10 text-center">
        <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">Select an outlet above to fill today's report</p>
      </div>
    )}
 {activeOutlet && outletReports[activeOutlet] && (
  <div className="bg-green-400/5 border border-green-400/30 p-6 mb-4">
    <div className="flex items-center justify-between mb-4">
      <div className="flex items-center gap-2">
<p className="text-green-400 font-mono text-xs uppercase tracking-widest">✓ Report for {OUTLET_NAMES[activeOutlet] || activeOutlet.replace(/_/g, " ")} — {(() => { const d = new Date(outletHistoryDate + "T00:00:00"); d.setDate(d.getDate() - 1); return d.toLocaleDateString("en-IN", { day: "numeric", month: "long", year: "numeric" }); })()} <span className="text-zinc-500">(sales day)</span></p>
  {outletReports[activeOutlet]?.is_edited && (
    <span className="font-mono text-[10px] uppercase tracking-widest px-2 py-0.5 bg-yellow-400/10 text-yellow-400">Edited</span>
  )}
</div>
      <button
        onClick={() => editOutletReport(activeOutlet)}
        className="font-mono text-[10px] uppercase tracking-widest px-3 py-1.5 border border-zinc-700 hover:border-yellow-400 hover:text-yellow-400 transition-colors"
      >
        ✏ Edit
      </button>
    </div>
    <div className="grid grid-cols-2 md:grid-cols-3 gap-3">
      {[
        { label: "Shop Sales", value: `₹${outletReports[activeOutlet].shop_sales_value} (${outletReports[activeOutlet].shop_sales_count} orders)` },
        { label: "Shop AOV", value: outletReports[activeOutlet].shop_sales_count > 0 ? `₹${Math.round(Number(outletReports[activeOutlet].shop_sales_value) / Number(outletReports[activeOutlet].shop_sales_count))}` : "—" },
        { label: "Swiggy", value: `₹${outletReports[activeOutlet].swiggy_sales_value} (${outletReports[activeOutlet].swiggy_sales_count} orders)` },
        { label: "Swiggy AOV", value: outletReports[activeOutlet].swiggy_sales_count > 0 ? `₹${Math.round(Number(outletReports[activeOutlet].swiggy_sales_value) / Number(outletReports[activeOutlet].swiggy_sales_count))}` : "—" },
        { label: "Zomato", value: `₹${outletReports[activeOutlet].zomato_sales_value} (${outletReports[activeOutlet].zomato_sales_count} orders)` },
        { label: "Zomato AOV", value: outletReports[activeOutlet].zomato_sales_count > 0 ? `₹${Math.round(Number(outletReports[activeOutlet].zomato_sales_value) / Number(outletReports[activeOutlet].zomato_sales_count))}` : "—" },
        { label: "Total Sales", value: `₹${Number(outletReports[activeOutlet].shop_sales_value) + Number(outletReports[activeOutlet].swiggy_sales_value) + Number(outletReports[activeOutlet].zomato_sales_value)}`, color: "text-yellow-400" },
        { label: "Total AOV", value: (() => { const totalVal = Number(outletReports[activeOutlet].shop_sales_value) + Number(outletReports[activeOutlet].swiggy_sales_value) + Number(outletReports[activeOutlet].zomato_sales_value); const totalCount = Number(outletReports[activeOutlet].shop_sales_count) + Number(outletReports[activeOutlet].swiggy_sales_count) + Number(outletReports[activeOutlet].zomato_sales_count); return totalCount > 0 ? `₹${Math.round(totalVal / totalCount)}` : "—"; })(), color: "text-yellow-400" },
        ...brandSalesItems(outletReports[activeOutlet]),
        { label: "Target", value: `₹${outletReports[activeOutlet].target}` },
        { label: "Swiggy Live", value: outletReports[activeOutlet].swiggy_live ? "✓ Yes" : "✗ No", color: outletReports[activeOutlet].swiggy_live ? "text-green-400" : "text-red-500" },
        { label: "Zomato Live", value: outletReports[activeOutlet].zomato_live ? "✓ Yes" : "✗ No", color: outletReports[activeOutlet].zomato_live ? "text-green-400" : "text-red-500" },
        { label: "Discount Running", value: outletReports[activeOutlet].discount_running || "—" },
        { label: "Expiry Items", value: `${outletReports[activeOutlet].expiry_count} — ${outletReports[activeOutlet].expiry_items || "—"}` },
        { label: "Complimentary", value: `${outletReports[activeOutlet].complimentary_count} — ${outletReports[activeOutlet].complimentary_items || "—"} (${outletReports[activeOutlet].complimentary_reason || "—"})` },
        { label: "BH Google", value: outletReports[activeOutlet].bh_google_rating ? `⭐ ${outletReports[activeOutlet].bh_google_rating}` : "—" },
        { label: "BH Swiggy", value: outletReports[activeOutlet].bh_swiggy_rating ? `⭐ ${outletReports[activeOutlet].bh_swiggy_rating}` : "—" },
        { label: "BH Zomato", value: outletReports[activeOutlet].bh_zomato_rating ? `⭐ ${outletReports[activeOutlet].bh_zomato_rating}` : "—" },
        { label: "CBH Swiggy", value: outletReports[activeOutlet].cbh_swiggy_rating ? `⭐ ${outletReports[activeOutlet].cbh_swiggy_rating}` : "—" },
        { label: "CBH Zomato", value: outletReports[activeOutlet].cbh_zomato_rating ? `⭐ ${outletReports[activeOutlet].cbh_zomato_rating}` : "—" },
        { label: "ICBH Swiggy", value: outletReports[activeOutlet].icbh_swiggy_rating ? `⭐ ${outletReports[activeOutlet].icbh_swiggy_rating}` : "—" },
        { label: "ICBH Zomato", value: outletReports[activeOutlet].icbh_zomato_rating ? `⭐ ${outletReports[activeOutlet].icbh_zomato_rating}` : "—" },
        { label: "Issues", value: outletReports[activeOutlet].issues || "—" },
        { label: "Action Taken", value: outletReports[activeOutlet].action_taken || "—" },
      ].map(f => (
        <div key={f.label} className="bg-black/30 px-3 py-2">
          <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
          <p className={`text-sm mt-1 ${f.color || "text-white"}`}>{f.value}</p>
        </div>
      ))}
    </div>
  </div>
)}

    {activeOutlet && !outletReports[activeOutlet] && (
      <div className="bg-[#131316] border border-zinc-800 p-6">
        <div className="flex items-center justify-between mb-6">
          <p className="text-sm font-bold uppercase tracking-widest">{OUTLET_NAMES[activeOutlet] || activeOutlet.replace(/_/g, " ")} — Yesterday's Sales Report</p>
          <span className="text-yellow-400 font-mono text-xs">Due: 12:00 PM today</span>
        </div>
        <p className="text-[11px] font-mono text-zinc-400 mb-5 -mt-3">📋 You're filing <span className="text-yellow-400">yesterday's sales</span> ({new Date(Date.now() - 86400000).toLocaleDateString("en-IN", { day: "numeric", month: "short" })}) — recorded under today's date, due by 12 noon.</p>
        {outletEntryDate >= BRAND_SPLIT_FROM && (
          <div className="mb-4 border border-yellow-400/40 bg-yellow-400/5 p-3 text-[11px] font-mono text-yellow-400 leading-relaxed">
            Enter <b>BH (Brownie Heaven)</b> and <b>CBH (Cakes by Brownie Heaven)</b> sales separately. Fields marked * are compulsory — type 0 if there were no sales.<br />
            Also compulsory: <b>Expired Product Name</b>, <b>Complimentary Product Name</b> and <b>Reason</b> — type "None" if nothing applies. The report cannot be submitted without them.
          </div>
        )}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-6">
         {[
  { label: "Yesterday's Target (Rs)", key: "target" },
  ...(outletEntryDate >= BRAND_SPLIT_FROM
    ? ["bh", "cbh"].flatMap((b) => [["shop", "Shop"], ["swiggy", "Swiggy"], ["zomato", "Zomato"]].flatMap(([ch, chName]) => [
        { label: `${b.toUpperCase()} — ${chName} Orders Count *`, key: `${b}_${ch}_sales_count`, hint: "Enter 0 if none" },
        { label: `${b.toUpperCase()} — ${chName} Sales Value (Rs) *`, key: `${b}_${ch}_sales_value`, hint: "Enter 0 if none" },
      ]))
    : [
  { label: "Shop Sales — Orders Count", key: "shop_sales_count" },
  { label: "Shop Sales — Value (Rs)", key: "shop_sales_value" },
  { label: "Swiggy Orders Count", key: "swiggy_sales_count" },
  { label: "Swiggy Sales Value (Rs)", key: "swiggy_sales_value" },
  { label: "Zomato Orders Count", key: "zomato_sales_count" },
  { label: "Zomato Sales Value (Rs)", key: "zomato_sales_value" },
  ]),
  { label: "Discount/Offer Running", key: "discount_running" },
  { label: "Discount Given (Rs)", key: "discount_given" },
  { label: "Unavailable Items", key: "unavailable_items" },
  { label: "Expiry Items Count", key: "expiry_count" },
  { label: `Expired Product Name${outletEntryDate >= "2026-10-03" ? " *" : ""}`, key: "expiry_items", hint: "Type 'None' if nothing expired" },
  { label: "Complimentary Given (count)", key: "complimentary_count" },
  { label: `Complimentary Product Name${outletEntryDate >= "2026-10-03" ? " *" : ""}`, key: "complimentary_items", hint: "Type 'None' if none given" },
  { label: `Reason${outletEntryDate >= "2026-10-03" ? " *" : ""}`, key: "complimentary_reason", hint: "Reason for the expiry / complimentary item — type 'None' if not applicable" },
  { label: "BH — Google Rating", key: "bh_google_rating" },
  { label: "BH — Swiggy Rating", key: "bh_swiggy_rating" },
  { label: "BH — Zomato Rating", key: "bh_zomato_rating" },
  { label: "CBH — Swiggy Rating", key: "cbh_swiggy_rating" },
  { label: "CBH — Zomato Rating", key: "cbh_zomato_rating" },
  { label: "ICBH — Swiggy Rating", key: "icbh_swiggy_rating" },
  { label: "ICBH — Zomato Rating", key: "icbh_zomato_rating" },
  { label: "Issues Today", key: "issues" },
  { label: "Action Taken", key: "action_taken" },
].map((f: { label: string; key: string; hint?: string }) => (
  <div key={f.key}>
    <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{f.label}</label>
    <input
      type="text"
      value={outletReportData[f.key] || ""}
      onChange={(e) => setOutletReportData(prev => ({ ...prev, [f.key]: e.target.value }))}
      className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm"
      placeholder={f.hint || "—"}
    />
  </div>
))}

{/* Yes/No dropdowns */}
{[
  { label: "Swiggy Live?", key: "swiggy_live" },
  { label: "Zomato Live?", key: "zomato_live" },
  { label: "Discount Rate Good?", key: "discount_rate_good" },
].map(f => (
  <div key={f.key}>
    <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">{f.label}</label>
    <select
      value={outletReportData[f.key] || "yes"}
      onChange={(e) => setOutletReportData(prev => ({ ...prev, [f.key]: e.target.value }))}
      className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm"
    >
      <option value="yes">Yes</option>
      <option value="no">No</option>
    </select>
  </div>
))}
        </div>
        <div className="mb-4">
          <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Report Date</label>
         <input type="date" max={new Date().toISOString().split("T")[0]} value={outletEntryDate} onChange={(e) => { setOutletEntryDate(e.target.value); setOutletHistoryDate(e.target.value); setOutletWasOff(false); fetchOutletReportsByDate(e.target.value); }} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" />
          {outletEntryDate < new Date().toISOString().split("T")[0] && (
            <div className="mt-3">
              <p className="text-[11px] font-mono text-zinc-400 uppercase tracking-widest mb-1.5">Were you off on this day?</p>
              <div className="flex gap-2">
                <button type="button" onClick={() => setOutletWasOff(true)} className={`text-[11px] font-mono uppercase tracking-widest px-3 py-1.5 border transition-colors ${outletWasOff ? "border-green-400 text-green-400" : "border-zinc-700 text-zinc-500"}`}>Yes, I was off</button>
                <button type="button" onClick={() => setOutletWasOff(false)} className={`text-[11px] font-mono uppercase tracking-widest px-3 py-1.5 border transition-colors ${!outletWasOff ? "border-red-400 text-red-400" : "border-zinc-700 text-zinc-500"}`}>No</button>
              </div>
              <p className={`text-[11px] font-mono uppercase tracking-widest mt-2 ${outletWasOff ? "text-green-400" : "text-red-400"}`}>{outletWasOff ? "✓ Off day — no points, no penalty" : "⚠️ Back-dated — −30 penalty"}</p>
            </div>
          )}
        </div>
        <button onClick={submitOutletReport} disabled={outletSubmitting} className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-6 py-3 hover:opacity-90 transition-opacity uppercase disabled:opacity-50">
          {outletSubmitting ? "Submitting..." : `Submit ${OUTLET_NAMES[activeOutlet] || activeOutlet.replace(/_/g, " ")} Report →`}
        </button>
        {(user.outlets || []).includes(activeOutlet) && (
          <div className="mt-8 border-t border-zinc-800 pt-6">
            <p className="text-sm font-bold uppercase tracking-widest mb-1">Swiggy / Zomato Reviews</p>
            <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">5★ +5 · 4★ +3 · under 2★ −5 · valid complaint −10 (stacks) · for {new Date(outletEntryDate + "T00:00:00").toLocaleDateString("en-IN", { day: "numeric", month: "short" })}</p>
            <div className="bg-black/30 border border-zinc-800 p-4 mb-3 grid grid-cols-1 md:grid-cols-2 gap-3">
              <select value={revForm.platform} onChange={e => setRevForm(p => ({ ...p, platform: e.target.value }))} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm"><option>Swiggy</option><option>Zomato</option></select>
              <select value={revForm.rating} onChange={e => setRevForm(p => ({ ...p, rating: e.target.value }))} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm"><option value="5">5 ★</option><option value="4">4 ★</option><option value="3">3 ★</option><option value="2">2 ★</option><option value="1">1 ★</option></select>
              <input type="text" value={revForm.note} onChange={e => setRevForm(p => ({ ...p, note: e.target.value }))} placeholder="Note (optional)" className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm md:col-span-2" />
              <label className="flex items-center gap-2 text-xs text-zinc-300"><input type="checkbox" checked={revForm.valid} onChange={e => setRevForm(p => ({ ...p, valid: e.target.checked }))} /> Valid complaint (our mistake) −10</label>
              <label className="flex items-center gap-2 text-xs text-zinc-300"><input type="checkbox" checked={revForm.refund} onChange={e => setRevForm(p => ({ ...p, refund: e.target.checked }))} /> Refund given</label>
            </div>
            <button onClick={saveReview} disabled={revSaving} className="bg-yellow-400 text-black font-bold text-[10px] px-4 py-2 uppercase tracking-widest disabled:opacity-50 mb-4">{revSaving ? "Adding..." : "+ Add Review"}</button>
            {reviews.length > 0 && (
              <div className="space-y-2">
                {reviews.map(rv => { const pts = reviewPoints(Number(rv.rating), rv.valid_complaint); return (
                  <div key={rv.id} className="flex items-center justify-between bg-black/30 border border-zinc-800 px-3 py-2 text-xs">
                    <div className="text-zinc-300">{rv.platform} · {rv.rating}★{rv.valid_complaint ? " · valid complaint" : ""}{rv.refund_given ? " · refunded" : ""}{rv.note ? ` · ${rv.note}` : ""}</div>
                    <div className="flex items-center gap-3">
                      <span className={pts >= 0 ? "text-green-400 font-bold" : "text-red-400 font-bold"}>{pts >= 0 ? "+" : ""}{pts}</span>
                      <button onClick={() => deleteReview(rv.id)} className="text-zinc-600 hover:text-red-400">✕</button>
                    </div>
                  </div>
                ); })}
              </div>
            )}
          </div>
        )}
      </div>
    )}
  </div>
)}
        {activeTab === "history" && (
  <div>
    <div className="flex justify-between items-end mb-6 pb-5 border-b border-zinc-800">
  <div>
    <h2 className="text-2xl font-black tracking-tight">History</h2>
    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">View reports by date</p>
  </div>
 <div className="flex gap-2">
  <button
    onClick={exportCSV}
    className="bg-zinc-800 border border-zinc-700 text-white font-bold tracking-widest text-xs px-5 py-3 hover:border-yellow-400 hover:text-yellow-400 transition-colors uppercase"
  >
    ↓ CSV
  </button>
  <button
    onClick={() => window.print()}
    className="bg-yellow-400 text-black font-bold tracking-widest text-xs px-5 py-3 hover:opacity-90 transition-opacity uppercase"
  >
    ↓ PDF
  </button>
</div>
</div>
    <div className="flex items-center gap-4 mb-8">
      <input
        type="date"
        value={historyDate}
        onChange={(e) => { setHistoryDate(e.target.value); fetchHistoryReports(e.target.value); }}
        className="bg-black border border-zinc-800 text-white px-4 py-2.5 focus:outline-none focus:border-yellow-400 transition-colors font-mono text-sm"
      />
      <span className="text-zinc-500 font-mono text-xs uppercase tracking-widest">{historyLoading ? "Loading..." : `${historyReports.length} reports · ${historyOutletReports.length} outlet reports`}</span>
    </div>

    {/* Staff Reports */}
    <div className="mb-8">
      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Staff Reports</p>
      <div className="bg-[#131316] border border-zinc-800">
        {historyReports.length === 0 ? (
          <div className="p-8 text-center">
            <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">No reports for this date</p>
          </div>
        ) : historyReports.map((r) => {
          const staffName = ALL_STAFF.find(s => s.id === r.staff_id)?.name || FORMER_STAFF[r.staff_id] || r.staff_id;
          const staffFields = REPORT_FIELDS[r.staff_id] || [];
          return (
            <div key={r.id} className="border-b border-zinc-800 last:border-0">
              <div className="flex items-center justify-between px-6 py-4 cursor-pointer hover:bg-zinc-900 transition-colors" onClick={() => setSelectedReport(selectedReport?.id === r.id ? null : r)}>
                <div className="flex items-center gap-3">
                  <span className="font-semibold text-sm">{staffName}</span>
                  {r.is_late && <span className="text-red-500 font-mono text-[10px] uppercase bg-red-500/10 px-2 py-0.5">Late</span>}
                </div>
                <div className="flex items-center gap-3">
                  <span className="text-[11px] font-mono text-zinc-500">{new Date(r.submitted_at).toLocaleString("en-IN", { hour: "2-digit", minute: "2-digit" })}</span>
                  <span className="text-zinc-500">{selectedReport?.id === r.id ? "▲" : "▼"}</span>
                </div>
              </div>
              {selectedReport?.id === r.id && r.report_data && (
                <div className="px-6 pb-6 grid grid-cols-1 md:grid-cols-2 gap-3">
                  {staffFields.map(f => (
                    <div key={f.key} className="bg-black/30 px-3 py-2">
                      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                      <p className="text-sm text-white mt-1">{r.report_data[f.key] || "—"}</p>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>

    {/* Outlet Reports */}
    <div>
      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Outlet Reports</p>
      <div className="grid grid-cols-1 gap-4">
        {historyOutletReports.length === 0 ? (
          <div className="bg-[#131316] border border-zinc-800 p-8 text-center">
            <p className="text-zinc-600 font-mono text-sm uppercase tracking-widest">No outlet reports for this date</p>
          </div>
        ) : historyOutletReports.map((r) => {
          const staffName = ALL_STAFF.find(s => s.id === r.staff_id)?.name || FORMER_STAFF[r.staff_id] || r.staff_id;
          return (
            <div key={r.id} className="bg-[#131316] border border-zinc-800 p-5">
              <div className="flex items-center justify-between mb-4">
                <div>
                  <p className="font-bold text-sm uppercase tracking-widest">{OUTLET_NAMES[r.outlet_id] || r.outlet_id.replace(/_/g, " ")}</p>
                  <p className="text-[10px] font-mono text-zinc-500 mt-0.5">{staffName}</p>
                </div>
                <div className="flex items-center gap-2">
                  {r.is_late && <span className="font-mono text-[10px] uppercase px-2 py-1 bg-red-500/10 text-red-500">Late</span>}
                  <span className="text-[11px] font-mono text-zinc-500">{new Date(r.submitted_at).toLocaleString("en-IN", { hour: "2-digit", minute: "2-digit" })}</span>
                </div>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                {[
                  { label: "Shop Sales", value: `₹${r.shop_sales_value} (${r.shop_sales_count})` },
                  { label: "Swiggy", value: `₹${r.swiggy_sales_value} (${r.swiggy_sales_count})` },
                  { label: "Zomato", value: `₹${r.zomato_sales_value} (${r.zomato_sales_count})` },
                  { label: "Target", value: `₹${r.target}` },
                  { label: "Swiggy Live", value: r.swiggy_live ? "✓ Yes" : "✗ No", color: r.swiggy_live ? "text-green-400" : "text-red-500" },
                  { label: "Zomato Live", value: r.zomato_live ? "✓ Yes" : "✗ No", color: r.zomato_live ? "text-green-400" : "text-red-500" },
                  { label: "Expiry", value: r.expiry_count > 0 ? `${r.expiry_count} — ${r.expiry_items}` : "None", color: r.expiry_count > 0 ? "text-red-500" : "" },
                  { label: "Issues", value: r.issues || "—", color: r.issues ? "text-yellow-400" : "" },
                ].map(f => (
                  <div key={f.label} className="bg-black/30 px-3 py-2">
                    <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{f.label}</p>
                    <p className={`text-sm mt-1 ${f.color || "text-white"}`}>{f.value}</p>
                  </div>
                ))}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  </div>
)}
          {activeTab === "team_dashboards" && isOwner && (
          <TeamDashboards staffList={ALL_STAFF.filter((s) => ["vishnu", "ahila", "arun", "nilani"].includes(s.id))} />
        )}
        {activeTab === "active_status" && (isOwner || isFO) && (
          <ActiveStatusTab staffList={ALL_STAFF.filter((s) => s.id !== "nishant")} />
        )}
        {activeTab === "messages" && isOwner && user && (
          <MessagesTab mode="owner" currentUserId={user.id} currentUserName={user.name} staffList={ALL_STAFF.filter((s) => s.id !== "nishant")} />
        )}
        {activeTab === "messages" && !isOwner && user && (
          <MessagesTab mode="staff" currentUserId={user.id} currentUserName={user.name} />
        )}
        {activeTab === "help" && <HelpTab canEdit={isOwner || isFO} />}
        {activeTab === "auto_reviews" && (isFO || ["nishant","arun","vishnu","ahila","nilani"].includes(user?.id ?? "")) && (
          <div>
            <div className="mb-6 flex items-center justify-between flex-wrap gap-3">
              <div>
                <h2 className="text-3xl font-black tracking-tight">⭐ Auto Reviews</h2>
                <p className="text-sm text-zinc-500 mt-1">Google and Zomato reviews auto-pulled into TASKFORCE IQ. Swiggy coming next.</p>
              </div>
              <button
                onClick={checkNewReviews}
                disabled={autoReviewsChecking}
                className="bg-yellow-400 text-black px-4 py-2 text-sm font-semibold hover:bg-yellow-300 transition-colors disabled:opacity-50"
              >
                {autoReviewsChecking ? "Checking..." : "🔄 Check Now"}
              </button>
            </div>

            {autoReviewsResult && (
              <div className={`border p-4 mb-6 text-sm ${autoReviewsResult.success ? "border-green-700 bg-green-950/30 text-green-300" : "border-red-700 bg-red-950/30 text-red-300"}`}>
                {autoReviewsResult.success ? (
                  <>
                    <p className="font-semibold mb-1">
                      ✅ Checked {autoReviewsResult.candidateCount ?? autoReviewsResult.inserted.length + autoReviewsResult.skipped.length} review-shaped emails — found {autoReviewsResult.inserted.length} new review{autoReviewsResult.inserted.length === 1 ? "" : "s"}
                      {autoReviewsResult.skipped.length > 0 ? `, skipped ${autoReviewsResult.skipped.length}` : ""}
                    </p>
                    {autoReviewsResult.timedOut && (
                      <p className="text-xs text-yellow-400 mt-1">⏱️ Ran out of time before finishing this batch — roughly {autoReviewsResult.remaining} still unchecked. Click "Check Now" again to continue; it picks up right where it left off.</p>
                    )}
                    {autoReviewsResult.skipReasons && Object.keys(autoReviewsResult.skipReasons).length > 0 && (
                      <ul className="text-xs text-zinc-400 mt-2 space-y-0.5">
                        {Object.entries(autoReviewsResult.skipReasons).map(([reason, count]: any) => (
                          <li key={reason}>• {count} skipped — {reason}</li>
                        ))}
                      </ul>
                    )}
                    {autoReviewsResult.sampleSkips && autoReviewsResult.sampleSkips.length > 0 && (
                      <details className="mt-2">
                        <summary className="text-xs text-zinc-500 cursor-pointer hover:text-zinc-300">Show example skipped emails</summary>
                        <div className="mt-2 space-y-2">
                          {autoReviewsResult.sampleSkips.map((s: any, i: number) => (
                            <div key={i} className="bg-black/40 p-2 text-[11px] text-zinc-400">
                              <p className="text-zinc-300">{s.reason}</p>
                              {s.subject && <p>Subject: {s.subject}</p>}
                              {s.outletNameRaw && <p>Outlet text found: "{s.outletNameRaw}"</p>}
                              {s.textPreview && <pre className="text-zinc-500 mt-1 whitespace-pre-wrap font-mono text-[10px]">{s.textPreview}</pre>}
                              {s.error && <p className="text-red-400">{s.error}</p>}
                            </div>
                          ))}
                        </div>
                      </details>
                    )}
                  </>
                ) : (
                  <p className="font-semibold">⚠️ Error: {autoReviewsResult.error}</p>
                )}
              </div>
            )}

            {autoReviewsList.length === 0 ? (
              <div className="bg-zinc-900 border border-zinc-800 p-6 max-w-xl">
                <p className="text-sm text-zinc-400">No auto-pulled reviews yet. Click "Check Now" to pull in any new Google reviews sitting in the inbox.</p>
              </div>
            ) : (
              <div className="space-y-3">
                {autoReviewsList.map((r) => (
                  <div key={r.id} className="bg-zinc-900 border border-zinc-800 p-4">
                    <div className="flex items-center justify-between mb-1">
                      <span className="font-semibold">{OUTLET_NAMES[r.outlet_id] || r.outlet_id}</span>
                      <span className="text-yellow-400">{"★".repeat(r.rating || 0)}{"☆".repeat(5 - (r.rating || 0))}</span>
                    </div>
                    <p className="text-sm text-zinc-400">{r.note}</p>
                    <div className="flex items-center justify-between mt-2">
                      <p className="text-xs text-zinc-600">{r.report_date}</p>
                      {r.review_url && (
                        <a href={r.review_url} target="_blank" rel="noopener noreferrer" className="text-xs text-yellow-400 hover:text-yellow-300 transition-colors">
                          View full review ↗
                        </a>
                      )}
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
        {activeTab === "ceo_report" && (isOwner || isFO) && (
          <div>
            <div className="mb-6">
              <h2 className="text-3xl font-black tracking-tight">{(() => { const h = new Date().getHours(); return h < 12 ? "Good morning ☕" : h < 17 ? "Good afternoon 🌤️" : "Good evening 🌙"; })()}</h2>
              <p className="text-sm text-zinc-500 mt-1">{user?.name?.split(" ")[0]} — here's the team and the month at a glance.</p>
            </div>
            <div className="flex gap-2 mb-6">
              {[{ k: "1", l: "Yesterday" }, { k: "7", l: "Last 7 days" }, { k: "30", l: "Last 30 days" }].map((w) => (
                <button key={w.k} onClick={() => { setCeoWin(w.k); fetchCeoData(w.k); }} className={`px-3 py-1.5 text-sm font-semibold transition-colors ${ceoWin === w.k ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>{w.l}</button>
              ))}
            </div>
            {(() => {
              const ceo = computeCeoData(ceoRepRows, ceoMonthRep, ceoWin);
              const inr = (n: number) => "₹" + Math.round(n).toLocaleString("en-IN");
              const toneC: Record<string, string> = { red: "text-red-400", green: "text-green-400", amber: "text-yellow-400", gray: "text-zinc-300" };
              return (
                <>
                  <div className="flex justify-end gap-2 mb-2 max-w-3xl">
                    {isFO && <button onClick={pushCeoPopupToNishant} disabled={ceoPushSending} className="bg-yellow-400 text-black px-4 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{ceoPushSending ? "Pushing…" : "📣 Push to Nishant"}</button>}
                    <button onClick={downloadCeoPDF} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors">Download report (PDF)</button>
                    <button onClick={() => ceoCustomOpen ? setCeoCustomOpen(false) : openCeoCustom()} className="bg-black border border-yellow-400 text-yellow-400 px-4 py-2 text-sm font-semibold hover:bg-yellow-400 hover:text-black transition-colors">🗓️ Custom date &amp; outlets</button>
                  </div>
                  {ceoCustomOpen && (
                    <div className="mb-4 max-w-3xl bg-[#131316] border border-zinc-800 p-4">
                      <div className="flex flex-wrap gap-3 mb-3">
                        <div><label className="text-[10px] font-mono text-zinc-500 uppercase block mb-1">From</label><input type="date" value={ceoCustomFrom} onChange={e => setCeoCustomFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm font-mono focus:outline-none focus:border-yellow-400" /></div>
                        <div><label className="text-[10px] font-mono text-zinc-500 uppercase block mb-1">To</label><input type="date" value={ceoCustomTo} onChange={e => setCeoCustomTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm font-mono focus:outline-none focus:border-yellow-400" /></div>
                      </div>
                      <div className="mb-3">
                        <div className="flex items-center justify-between mb-2">
                          <label className="text-[10px] font-mono text-zinc-500 uppercase">Outlets {ceoCustomOutlets.length === 0 ? "(all)" : `(${ceoCustomOutlets.length})`}</label>
                          <button onClick={() => setCeoCustomOutlets([])} className="text-[10px] font-mono text-zinc-500 uppercase">Reset to all</button>
                        </div>
                        <div className="flex flex-wrap gap-2">
                          {OUTLETS.map(o => { const on = ceoCustomOutlets.includes(o); return (
                            <button key={o} onClick={() => setCeoCustomOutlets(on ? ceoCustomOutlets.filter(x => x !== o) : [...ceoCustomOutlets, o])} className={`text-[11px] px-3 py-1.5 border font-mono uppercase tracking-wide transition-colors ${on ? "bg-yellow-400 text-black border-yellow-400" : "bg-black text-zinc-400 border-zinc-800 hover:border-zinc-600"}`}>{OUTLET_NAMES[o] || o}</button>
                          ); })}
                        </div>
                      </div>
                      <div className="flex gap-2">
                        <button onClick={downloadCeoPDFCustom} disabled={ceoCustomBusy} className="bg-yellow-400 text-black font-bold text-xs px-5 py-2.5 uppercase tracking-widest disabled:opacity-50 hover:opacity-90">{ceoCustomBusy ? "Generating…" : "Generate PDF"}</button>
                        <button onClick={() => setCeoCustomOpen(false)} className="text-zinc-500 text-xs font-mono uppercase px-3">Cancel</button>
                      </div>
                    </div>
                  )}
                  <div className="mb-6 border border-zinc-800 p-5 max-w-3xl">
                    <p className="text-sm font-bold mb-1">👀 Who's on top of it — and who's slipping</p>
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Report filing · {ceoWin === "1" ? "yesterday" : `last ${ceo.winDays} days`}</p>
                    <div className="space-y-2">
                      {ceo.acct.map((a) => (
                        <div key={a.name} className="flex items-baseline gap-2 text-sm">
                          <span className="font-medium w-24">{a.name}</span>
                          <span className={`flex-1 ${toneC[a.tone]}`}>{a.tag}</span>
                        </div>
                      ))}
                      <div className="flex items-baseline gap-2 text-sm pt-2 border-t border-zinc-900">
                        <span className="font-medium w-24">Niranjana</span>
                        <span className="flex-1 text-zinc-500">Founder's Office — MIA from daily reports 😎 but she built this thing, so we'll let it slide</span>
                      </div>
                    </div>
                  </div>
                  <div className="mb-6 border border-zinc-800 p-5 max-w-3xl">
                    <p className="text-sm font-bold mb-1">📈 Are we going to make the month?</p>
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-4">{ceo.ym} · day {ceo.daysElapsed} of {ceo.daysInMonth}</p>
                    <div className="flex items-end justify-between mb-3 flex-wrap gap-3">
                      <div>
                        <p className="text-3xl font-black">{inr(ceo.monthSales)}</p>
                        <p className={`text-sm font-mono ${ceo.onTrack ? "text-green-400" : "text-yellow-400"}`}>{ceo.salesPct.toFixed(0)}% of {inr(ceo.monthTgt)} target · {ceo.onTrack ? "on pace ✅" : "behind pace ⚠️"}</p>
                      </div>
                      <div className="text-right text-xs text-zinc-500 font-mono">
                        <p>Time elapsed: {ceo.timePct.toFixed(0)}%</p>
                        <p>Sales so far: {ceo.salesPct.toFixed(0)}%</p>
                      </div>
                    </div>
                    {ceo.drag && <p className="text-sm text-red-400">🔻 Dragging: <span className="font-medium">{ceo.drag.name}</span> — {ceo.drag.pct.toFixed(0)}% of target</p>}
                    {ceo.hero && <p className="text-sm text-green-400 mt-0.5">🔺 Carrying: <span className="font-medium">{ceo.hero.name}</span> — {ceo.hero.pct.toFixed(0)}% of target</p>}
                    <button onClick={() => setActiveTab("owner_outlets")} className="mt-4 text-xs font-mono uppercase text-zinc-400 hover:text-yellow-400 transition-colors">See full sales →</button>
                  </div>
                  <div className="mb-6 border-2 border-yellow-500/30 bg-yellow-500/5 p-5 max-w-3xl">
                    <p className="text-sm font-bold mb-3 text-yellow-400">💡 Ideas to act on</p>
                    <ul className="space-y-2 list-disc list-inside">
                      {ceo.ideas.map((idea, i) => <li key={i} className="text-sm text-white">{idea}</li>)}
                    </ul>
                  </div>
                  <div className="mb-6 border border-zinc-800 p-5 max-w-3xl">
                    <div className="flex items-center justify-between mb-1">
                      <p className="text-sm font-bold">💰 Gross margin by outlet</p>
                      {marginMonths.length > 0 && (
                        <select value={marginMonth} onChange={(e) => { setMarginMonth(e.target.value); fetchMargin(e.target.value); }} className="bg-black border border-zinc-800 text-white px-2 py-1 text-xs focus:outline-none focus:border-yellow-400">
                          {marginMonths.map((m) => <option key={m} value={m}>{m}</option>)}
                        </select>
                      )}
                    </div>
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Sales − true food cost · dispatch-based</p>
                    {marginLoading ? (
                      <p className="text-sm text-zinc-500">Loading…</p>
                    ) : marginData.length === 0 ? (
                      <p className="text-sm text-zinc-500">No data for this month yet.</p>
                    ) : (
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead>
                            <tr className="text-left text-[11px] font-mono text-zinc-500 uppercase">
                              <th className="py-1">Outlet</th>
                              <th className="py-1 text-right">Sales</th>
                              <th className="py-1 text-right">Food cost</th>
                              <th className="py-1 text-right">Margin</th>
                              <th className="py-1 text-right">%</th>
                            </tr>
                          </thead>
                          <tbody>
                            {[...marginData].sort((a: any, b: any) => (a.marginPercent ?? 0) - (b.marginPercent ?? 0)).map((m: any) => (
                              <tr key={m.outletId} className="border-t border-zinc-800">
                                <td className="py-1.5">{m.outletName}</td>
                                <td className="py-1.5 text-right font-mono">{inr(m.salesNet)}</td>
                                <td className="py-1.5 text-right font-mono text-zinc-400">{inr(m.cogs)}</td>
                                <td className={`py-1.5 text-right font-mono ${m.grossMargin < 0 ? "text-red-400" : "text-zinc-300"}`}>{inr(m.grossMargin)}</td>
                                <td className={`py-1.5 text-right font-mono ${m.marginPercent == null ? "text-zinc-600" : m.marginPercent < 40 ? "text-red-400" : m.marginPercent < 60 ? "text-yellow-400" : "text-green-400"}`}>{m.marginPercent == null ? "—" : m.marginPercent.toFixed(1) + "%"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <p className="text-[10px] text-zinc-600 mt-3">A few products (combo boxes, some newer items) don't have a confirmed recipe cost yet, so their cost is left out — true margin is a little lower than shown. Franchise outlets aren't included.</p>
                      </div>
                    )}
                  </div>
                  <p className="text-[10px] text-zinc-600 max-w-3xl">Ideas are rule-based for now — the smarter AI-written version is the next phase. Purchases, stock and wastage aren't tracked yet.</p>
                </>
              );
            })()}
          </div>
        )}

                 {activeTab === "niranjana_report" && (isOwner || isFO) && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Niranjana's Report</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">What's been built, day by day</p>
            </div>
            {isFO && (
              <div className="mb-8 border border-zinc-800 p-5 max-w-3xl">
                <p className="text-sm font-semibold mb-3">Add today's entry</p>
                <div className="flex flex-wrap gap-3 mb-3">
                  <input type="date" value={nrNewDate} onChange={(e) => setNrNewDate(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" />
                </div>
                <textarea value={nrNewContent} onChange={(e) => setNrNewContent(e.target.value)} rows={4} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm" placeholder="What did you build/fix today?" />
                <button onClick={saveNrEntry} disabled={nrSaving} className="mt-3 bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{nrSaving ? "Saving…" : "Push to report"}</button>
              </div>
            )}
            <div className="max-w-3xl space-y-4 text-sm text-zinc-300 leading-relaxed">
              {nrEntries.length === 0 ? <p className="text-zinc-600">No entries yet.</p> : nrEntries.map((e) => (
                <p key={e.id}><span className="text-yellow-400 font-mono text-xs mr-2">{new Date(e.entry_date + "T00:00:00").toLocaleDateString("en-IN", { month: "short", day: "numeric" })}</span>{e.content}</p>
              ))}
            </div>
          </div>
        )}
                {activeTab === "cheques" && (canAssign || isFO) && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Cheque Ledger</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Every cheque, tracked pending → cleared</p>
            </div>

            <div className="mb-8 border border-zinc-800 p-5 max-w-2xl">
              <p className="text-sm font-semibold mb-4">Add a cheque</p>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                <div>
                  <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Supplier</label>
                  <select value={chqSupplierSel} onChange={(e) => setChqSupplierSel(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1">
                    <option value="">— pick —</option>
                    {chqSuppliers.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
                    <option value="__new__">+ Add new supplier</option>
                  </select>
                </div>
                {chqSupplierSel === "__new__" && (
                  <div>
                    <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">New supplier name</label>
                    <input type="text" value={chqNewSupplierName} onChange={(e) => setChqNewSupplierName(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="Supplier name" />
                  </div>
                )}
                <div>
                  <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Amount (₹)</label>
                  <input type="number" value={chqAmount} onChange={(e) => setChqAmount(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="e.g. 25000" />
                </div>
                <div>
                  <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Date issued</label>
                  <input type="date" value={chqDate} onChange={(e) => setChqDate(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" />
                </div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3 mb-3">
                <div>
                  <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest block mb-1">Cheque photo — front</label>
                  <input type="file" accept="image/*" onChange={(e) => setChqPhotoFront(e.target.files?.[0] || null)} className="text-xs text-zinc-400 w-full" />
                </div>
                <div>
                  <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest block mb-1">Cheque photo — back (optional)</label>
                  <input type="file" accept="image/*" onChange={(e) => setChqPhotoBack(e.target.files?.[0] || null)} className="text-xs text-zinc-400 w-full" />
                </div>
              </div>
              <div className="mb-3">
                <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Notes (optional)</label>
                <input type="text" value={chqNotes} onChange={(e) => setChqNotes(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="What's this for" />
              </div>
              <button onClick={saveCheque} disabled={chqSaving} className="bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{chqSaving ? "Saving…" : "Save Cheque"}</button>
            </div>

            <div className="max-w-4xl">
              <p className="text-sm font-semibold mb-3">Register</p>
              {chqLoading ? (
                <p className="text-sm text-zinc-500">Loading…</p>
              ) : chqRows.length === 0 ? (
                <p className="text-sm text-zinc-500">No cheques logged yet.</p>
              ) : (
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="text-left text-[10px] font-mono text-zinc-500 uppercase border-b border-zinc-800">
                        <th className="py-2 pr-3">Date</th>
                        <th className="py-2 pr-3">Supplier</th>
                        <th className="py-2 pr-3 text-right">Amount</th>
                        <th className="py-2 pr-3">Photo</th>
                        <th className="py-2 pr-3">Status</th>
                        <th className="py-2">Notes</th>
                      </tr>
                    </thead>
                    <tbody>
                      {chqRows.map((c) => (
                        <tr key={c.id} className="border-b border-zinc-900">
                          <td className="py-2 pr-3 font-mono text-xs">{c.date_issued}</td>
                          <td className="py-2 pr-3 font-semibold">{c.supplier_name}</td>
                          <td className="py-2 pr-3 text-right font-mono">₹{Math.round(c.amount).toLocaleString("en-IN")}</td>
                          <td className="py-2 pr-3">
                            {c.photo_front_url ? <a href={c.photo_front_url} target="_blank" rel="noreferrer" className="text-yellow-400 underline text-xs">View</a> : <span className="text-zinc-600 text-xs">—</span>}
                          </td>
                          <td className="py-2 pr-3">
                            <button onClick={() => toggleChequeStatus(c)} className={`font-mono text-[10px] uppercase px-2 py-1 ${c.status === "cleared" ? "bg-green-400/10 text-green-400" : "bg-yellow-400/10 text-yellow-400"}`}>{c.status}</button>
                          </td>
                          <td className="py-2 text-xs text-zinc-400">{c.notes || "—"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        )}
        {activeTab === "fines" && (canAssign || isFO) && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Fines</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Log a fine · it also hits their leaderboard score</p>
            </div>
            <div className="mb-8 border border-zinc-800 p-5 max-w-3xl">
              <p className="text-sm font-semibold mb-1">Issue a fine</p>
              <p className="text-xs text-zinc-500 mb-4">Pick who's fined (e.g. for a 1-star review), set the reason and amount. Each selected person is fined the same amount.</p>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Who</p>
              <div className="flex flex-wrap gap-2 mb-4">
                {(ALL_STAFF as any[]).filter((s) => s.id !== "nishant").map((s) => { const on = fineStaff.includes(s.id); return (
                  <button key={s.id} onClick={() => setFineStaff((pp) => on ? pp.filter((x) => x !== s.id) : [...pp, s.id])} className={`px-3 py-1.5 text-sm font-semibold transition-colors ${on ? "bg-red-500 text-white" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>{s.name.split(" ")[0]}</button>
                ); })}
              </div>
              <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Reason</label><input type="text" value={fineReason} onChange={(e) => setFineReason(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="1-star review" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Amount (₹)</label><input type="number" value={fineAmount} onChange={(e) => setFineAmount(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Date</label><input type="date" value={fineDate} onChange={(e) => setFineDate(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Outlet (optional)</label><select value={fineOutlet} onChange={(e) => setFineOutlet(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1"><option value="">—</option>{OUTLETS.map((o) => <option key={o} value={o}>{OUTLET_NAMES[o] || o}</option>)}</select></div>
              </div>
              <button onClick={saveFine} disabled={fineBusy} className="mt-4 bg-red-500 text-white px-5 py-2 text-sm font-semibold hover:bg-red-400 disabled:opacity-50 transition-colors">{fineBusy ? "Saving…" : `Fine ${fineStaff.length || 0} ${fineStaff.length === 1 ? "person" : "people"}`}</button>
            </div>
            <div className="max-w-3xl">
              <p className="text-sm font-semibold mb-3">Recent fines</p>
              {fines.length === 0 ? <p className="text-sm text-zinc-600">No fines yet.</p> : (
                <div className="space-y-1.5">
                  {fines.map((f) => (
                    <div key={f.id} className="flex items-baseline gap-3 text-sm border-b border-zinc-900 pb-1.5">
                      <span className="font-medium w-24">{f.staff_name}</span>
                      <span className="flex-1 text-zinc-400">{f.reason}{f.outlet ? ` · ${OUTLET_NAMES[f.outlet] || f.outlet}` : ""}</span>
                      <span className="font-mono text-red-400">−₹{Number(f.amount).toLocaleString("en-IN")}</span>
                      <span className="font-mono text-zinc-600 text-xs">{f.fine_date}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "ops_audits" && (canAssign || isFO) && (
          <div>
            <div className="mb-6 pb-5 border-b border-line">
              <h2 className="text-2xl font-black tracking-tight text-text">Operations &amp; Audits</h2>
              <p className="text-[11px] font-mono text-text-muted uppercase tracking-widest mt-1">Bharani&apos;s daily audit reports</p>
            </div>
            {auditReportsLoading ? (
              <p className="text-sm text-text-faint">Loading…</p>
            ) : auditReports.length === 0 ? (
              <p className="text-sm text-text-faint">No audit reports submitted yet.</p>
            ) : (
              <div className="max-w-3xl space-y-4">
                {auditReports.map((r, i) => {
                  const d = r.report_data || {};
                  const notes = [d.issue_1, d.issue_2, d.issue_3, d.action_1, d.action_2, d.action_3].filter((x) => x && String(x).trim() && String(x).trim().toLowerCase() !== "none" && String(x).trim() !== "-").join(" · ");
                  const isEmpty = (v: any) => v === undefined || v === null || String(v).trim() === "" || String(v).trim() === "-" || String(v).trim().toLowerCase() === "none";
                  const fields = [
                    { label: "Outlets audited", value: d.outlets_audited },
                    { label: "Wastage", value: d.total_wastage, money: true },
                    { label: "Stock mismatch", value: d.stock_mismatch, flag: !isEmpty(d.stock_mismatch) && String(d.stock_mismatch).toLowerCase() !== "no" },
                    { label: "Cash reconciliation", value: d.cash_reconciliation },
                    { label: "Exceptions found", value: d.exceptions_found },
                    { label: "Discrepancy", value: d.discrepancy_value, money: true },
                    { label: "Outlets w/ issues", value: d.outlets_with_issues },
                  ].filter((f) => !isEmpty(f.value));
                  const dateLabel = new Date(r.report_date + "T00:00:00").toLocaleDateString("en-IN", { weekday: "short", day: "numeric", month: "short", year: "numeric" });
                  return (
                    <div key={i} className="bg-surface border border-line p-5">
                      <p className="text-sm font-bold text-text mb-3">{dateLabel}</p>
                      {fields.length === 0 ? (
                        <p className="text-xs text-text-faint mb-3">No structured figures logged for this audit.</p>
                      ) : (
                        <div className="grid grid-cols-2 md:grid-cols-3 gap-x-6 gap-y-3 mb-4">
                          {fields.map((f) => (
                            <div key={f.label}>
                              <p className="text-[10px] font-mono text-text-muted uppercase tracking-widest mb-0.5">{f.label}</p>
                              <p className={`text-sm font-semibold ${f.flag ? "text-red-400" : "text-text"}`}>{f.money ? `₹${f.value}` : String(f.value)}</p>
                            </div>
                          ))}
                        </div>
                      )}
                      {notes && (
                        <div className="pt-3 border-t border-line">
                          <p className="text-[10px] font-mono text-text-muted uppercase tracking-widest mb-1.5">Notes</p>
                          <p className="text-sm text-text-muted leading-relaxed">{notes}</p>
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {activeTab === "outlet_targets" && isArun && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl md:text-3xl font-black tracking-tight">Outlet Targets</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Change an outlet's monthly target from a chosen month onward — earlier months are never touched</p>
            </div>

            <div className="bg-[#131316] border border-zinc-800 p-5 max-w-lg mb-8">
              <p className="text-sm font-semibold mb-4">New target change</p>
              <div className="space-y-3">
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Outlet</label>
                  <select value={otOutlet} onChange={(e) => setOtOutlet(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm">
                    {OUTLETS.map((o) => <option key={o} value={o}>{OUTLET_NAMES[o] || o}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">Apply from (month)</label>
                  <input type="month" min={new Date().toISOString().slice(0, 7)} value={otMonth} onChange={(e) => setOtMonth(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                  <p className="text-[10px] text-zinc-600 mt-1">Current now: ₹{monthlyTargetFor(otOutlet, new Date().toISOString().slice(0, 7)).toLocaleString("en-IN")}/month. Only this month or a future month can be picked — past months stay exactly as they were.</p>
                </div>
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">New monthly target (₹)</label>
                  <input type="text" value={otValue} onChange={(e) => setOtValue(e.target.value)} placeholder="e.g. 2800000" className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm font-mono" />
                </div>
                <button onClick={saveOutletTargetUpdate} disabled={otSaving} className="bg-yellow-400 text-black font-bold text-xs px-6 py-3 uppercase tracking-widest hover:opacity-90 disabled:opacity-50 transition-opacity">{otSaving ? "Saving…" : "Apply from this month →"}</button>
              </div>
            </div>

            <div className="max-w-2xl">
              <p className="text-sm font-semibold mb-3">History of changes</p>
              {otRows.length === 0 ? (
                <p className="text-xs text-zinc-600">No changes made from this screen yet — outlets are still on their original targets.</p>
              ) : (
                <div className="space-y-1.5">
                  {[...otRows].reverse().map((r: any) => (
                    <div key={r.id} className="flex justify-between text-sm border-b border-zinc-900 pb-1.5">
                      <span>{OUTLET_NAMES[r.outlet_id] || r.outlet_id} <span className="text-zinc-600">· from {r.effective_from}</span></span>
                      <span className="font-mono text-yellow-400">₹{Number(r.monthly_target).toLocaleString("en-IN")}/mo</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === "admin_dq" && (isOwner || isFO) && (
          <div>
            <div className="mb-6 pb-5 border-b border-zinc-800 flex justify-between items-end gap-4 flex-wrap">
              <div>
                <h2 className="text-2xl font-black tracking-tight">Admin &amp; Data Quality</h2>
                <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Are the numbers complete? · last 7 days · real checks only</p>
              </div>
              <button onClick={fetchDq} disabled={dqLoading} className="bg-yellow-400 text-black font-bold text-[10px] px-4 py-2 uppercase tracking-widest disabled:opacity-50">{dqLoading ? "Checking…" : "↻ Re-check"}</button>
            </div>

            {!dq || dqLoading ? (
              <p className="text-xs text-zinc-500">Running checks…</p>
            ) : (
              <div className="space-y-8 max-w-4xl">
                {dq.fetchError && <p className="text-xs text-red-400 border border-red-400/40 p-3">Some data could not be read: {dq.fetchError}</p>}
                {dq.brandColsMissing && <p className="text-xs text-red-400 border border-red-400/40 p-3">The BH / CBH sales columns are not in the database yet — run <b>brand_sales_columns.sql</b>, otherwise outlet reports from {BRAND_SPLIT_FROM} cannot be saved.</p>}

                {/* Summary tiles */}
                {(() => {
                  const tiles = [
                    { label: "Filed today", value: `${dq.filedToday}/${OUTLETS.length}`, note: "outlets (due 12 noon)", bad: false },
                    { label: "Missing reports", value: String(dq.missing.length), note: "outlet-days, last 7", bad: dq.missing.length > 0 },
                    { label: "Late reports", value: String(dq.late.length), note: "after 12 noon", bad: dq.late.length > 0 },
                    { label: "Production gaps", value: String(dq.prodMissing.length), note: "days with no entry", bad: dq.prodMissing.length > 0 },
                  ];
                  return (
                    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
                      {tiles.map((t) => (
                        <div key={t.label} className="border border-zinc-800 p-3">
                          <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{t.label}</p>
                          <p className={`text-2xl font-black mt-1 ${t.bad ? "text-red-400" : "text-green-400"}`}>{t.value}</p>
                          <p className="text-[10px] text-zinc-500 mt-0.5">{t.note}</p>
                        </div>
                      ))}
                    </div>
                  );
                })()}

                {/* Outlet report completeness */}
                <div>
                  <p className="text-sm font-semibold mb-1">Outlet report completeness</p>
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">✓ on time · L late · B back-filled · ✗ missing (an outlet's off day also shows ✗)</p>
                  <div className="overflow-x-auto border border-zinc-800">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-zinc-500 text-[10px] font-mono uppercase tracking-widest">
                          <th className="text-left px-3 py-2">Outlet</th>
                          {dq.days.map((d: string) => <th key={d} className="px-2 py-2 text-center">{d.slice(8)}/{d.slice(5, 7)}</th>)}
                        </tr>
                      </thead>
                      <tbody>
                        {dq.matrix.map((m: any) => (
                          <tr key={m.outlet} className="border-t border-zinc-800">
                            <td className="px-3 py-1.5">{OUTLET_NAMES[m.outlet] || m.outlet}</td>
                            {m.cells.map((c: string, i: number) => (
                              <td key={i} className={`px-2 py-1.5 text-center font-bold ${c === "ok" ? "text-green-400" : c === "late" ? "text-yellow-400" : c === "backfill" ? "text-orange-400" : "text-red-400"}`}>
                                {c === "ok" ? "✓" : c === "late" ? "L" : c === "backfill" ? "B" : "✗"}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Issue list */}
                <div>
                  <p className="text-sm font-semibold mb-3">Data-quality checks</p>
                  <div className="border border-zinc-800 divide-y divide-zinc-800">
                    {[
                      { title: "Missing outlet reports", items: dq.missing, hint: "No report filed for that outlet and day." },
                      { title: "Late outlet reports", items: dq.late, hint: "Filed after the 12 noon cut-off." },
                      { title: "Back-filled reports", items: dq.backfilled, hint: "Entered for a past date, after that day." },
                      { title: "Duplicate reports (same outlet, same day)", items: dq.dupes, hint: "Totals may be double-counted — check which one is right." },
                      { title: "Reports with zero sales", items: dq.zero, hint: "Shop + Swiggy + Zomato are all 0 — a mistake or a closed day?" },
                      { title: "Compulsory fields left empty (from 3 Oct)", items: dq.blankReq, hint: "Expired product, complimentary product or reason is empty." },
                      { title: "Reports missing the BH / CBH split", items: dq.brandMissing, hint: `Expected from ${BRAND_SPLIT_FROM}.` },
                      { title: "Days with no production entry", items: dq.prodMissing.map((d: string) => d), hint: "Head Chef has not filed production for that day." },
                    ].map((c) => (
                      <details key={c.title} className="px-3 py-2">
                        <summary className="cursor-pointer flex items-center gap-2 text-xs">
                          <span className={`w-2 h-2 rounded-full ${c.items.length === 0 ? "bg-green-400" : "bg-red-400"}`} />
                          <span className="flex-1">{c.title}</span>
                          <span className={`font-mono ${c.items.length === 0 ? "text-green-400" : "text-red-400"}`}>{c.items.length === 0 ? "OK" : c.items.length}</span>
                        </summary>
                        {c.items.length > 0 && (
                          <div className="mt-2 pl-4 text-[11px] text-zinc-400">
                            <p className="text-zinc-500 mb-1">{c.hint}</p>
                            <ul className="list-disc pl-4 space-y-0.5">{c.items.slice(0, 40).map((it: string, i: number) => <li key={i}>{it}</li>)}</ul>
                            {c.items.length > 40 && <p className="mt-1 text-zinc-500">…and {c.items.length - 40} more</p>}
                          </div>
                        )}
                      </details>
                    ))}
                  </div>
                </div>

                {/* Data freshness */}
                <div>
                  <p className="text-sm font-semibold mb-3">Data freshness</p>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs">
                    <div className="border border-zinc-800 p-3">
                      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Latest Item Performance (Atlas) upload</p>
                      <p className="mt-1">{dq.latestUpload ? `${dq.latestUpload.label || "Untitled"} · ${new Date(dq.latestUpload.created_at).toLocaleDateString("en-IN")}` : "No upload yet"}</p>
                    </div>
                    <div className="border border-zinc-800 p-3">
                      <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">Last checked</p>
                      <p className="mt-1">{dq.checkedAt}</p>
                    </div>
                  </div>
                </div>

                {/* Users & roles */}
                <div>
                  <p className="text-sm font-semibold mb-1">Users &amp; roles</p>
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-3">As set up in the app · {ALL_STAFF.length} users</p>
                  <div className="overflow-x-auto border border-zinc-800">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-zinc-500 text-[10px] font-mono uppercase tracking-widest">
                          <th className="text-left px-3 py-2">Name</th>
                          <th className="text-left px-3 py-2">Role</th>
                          <th className="text-left px-3 py-2">Outlets</th>
                          <th className="text-left px-3 py-2">Daily report</th>
                        </tr>
                      </thead>
                      <tbody>
                        {ALL_STAFF.map((s) => (
                          <tr key={s.id} className="border-t border-zinc-800">
                            <td className="px-3 py-1.5">{s.name}</td>
                            <td className="px-3 py-1.5">{s.role}</td>
                            <td className="px-3 py-1.5">{(s.outlets as string[]).length ? (s.outlets as string[]).map((o) => OUTLET_NAMES[o] || o).join(", ") : "—"}</td>
                            <td className="px-3 py-1.5">{s.report_time ? `by ${s.report_time}` : "—"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* Honest list of what is NOT covered yet */}
                <div className="border border-zinc-800 p-4">
                  <p className="text-sm font-semibold mb-2">Not covered yet</p>
                  <ul className="list-disc pl-5 text-[11px] text-zinc-500 space-y-0.5">
                    <li>Audit log of who changed what</li>
                    <li>Backup and data-refresh status</li>
                    <li>Live health of Swiggy / Zomato / POS / accounting / attendance integrations</li>
                    <li>Duplicate-transaction and unit-mismatch checks on purchases and stock</li>
                    <li>Editing users and permissions from this screen (today they are set in the code)</li>
                  </ul>
                </div>
              </div>
            )}
          </div>
        )}

        {activeTab === "item_perf" && (
          <div>
            <div className="mb-8 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Item Performance</h2>
              <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">UrbanPiper item export · price vs demand</p>
            </div>

            {/* Wastage & Complimentary — this month, from the daily Outlet Report.
                Compact by design: a ranked list per product, click one to see
                which outlets it's actually happening at. */}
            <div className="mb-8 border border-zinc-800 p-4 max-w-xl">
              <div className="flex justify-between items-baseline mb-3">
                <p className="text-sm font-semibold">Wastage &amp; Complimentary</p>
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest">{new Date().toLocaleDateString("en-IN", { month: "short", year: "numeric" })}</p>
              </div>
              {wastageLoading ? (
                <p className="text-xs text-zinc-500">Loading…</p>
              ) : wastageRows.length === 0 ? (
                <p className="text-xs text-zinc-500">Nothing logged yet this month.</p>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-6 gap-y-4">
                  {[
                    { title: "Expiry", entries: wastageExpiryEntries, field: "expiry_items", sectionKey: "exp" },
                    { title: "Complimentary", entries: wastageComplimentaryEntries, field: "complimentary_items", sectionKey: "comp" },
                  ].map(({ title, entries, field, sectionKey }) => {
                    const rows = wastageProductBreakdown(entries, field);
                    return (
                      <div key={sectionKey}>
                        <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1.5">{title}</p>
                        {rows.length === 0 ? <p className="text-xs text-zinc-600">None yet.</p> : (
                          <div className="space-y-0.5">
                            {rows.map((v) => {
                              const expandKey = `${sectionKey}:${v.name}`;
                              const isOpen = wastageExpanded === expandKey;
                              return (
                                <div key={v.name}>
                                  <button onClick={() => setWastageExpanded(isOpen ? null : expandKey)} className="w-full flex justify-between items-center text-left py-1 hover:bg-zinc-900 transition-colors px-1 -mx-1">
                                    <span className="text-sm capitalize text-zinc-300 truncate pr-2">{v.name}</span>
                                    <span className={`text-[10px] font-mono shrink-0 ${v.color}`}>{v.days}d {isOpen ? "▲" : "▼"}</span>
                                  </button>
                                  {isOpen && (
                                    <div className="pl-1 pb-1.5 pt-0.5 border-l border-zinc-800 ml-1">
                                      <p className={`text-[10px] font-mono mb-1 ${v.color}`}>{v.label}</p>
                                      {v.outlets.map((o) => (
                                        <div key={o.outlet} className="flex justify-between text-xs text-zinc-400 pl-2"><span>{o.outlet}</span><span className="font-mono">{o.days}d</span></div>
                                      ))}
                                    </div>
                                  )}
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {canUploadItemPerf && (
              <div className="mb-8 border border-zinc-800 p-5 max-w-3xl">
                <p className="text-sm font-semibold mb-1">Upload item export</p>
                <p className="text-xs text-zinc-500 mb-4">CSV or Excel from UrbanPiper Prime (revenue by item). Pick the window it covers.</p>
                <div className="flex flex-wrap gap-3 items-end">
                  <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Window</label><select value={ipDays} onChange={(e) => setIpDays(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-36"><option value="7">Last 7 days</option><option value="10">Last 10 days</option><option value="30">Last 30 days</option></select></div>
                  <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Label (optional)</label><input type="text" value={ipLabel} onChange={(e) => setIpLabel(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-44" placeholder="1–30 Jul" /></div>
                  <label className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 cursor-pointer transition-colors">Choose file<input type="file" accept=".csv,.xlsx,.xls" onChange={onIpFile} className="hidden" /></label>
                </div>
                {ipParsed && (
                  <div className="mt-5">
                    <p className="text-xs text-zinc-500 mb-2">Parsed {ipParsed.length} items · preview (top 8 by revenue):</p>
                    <table className="w-full text-sm">
                      <thead><tr className="text-left text-[11px] font-mono text-zinc-500 uppercase"><th className="py-1">Item</th><th className="py-1 text-right">Revenue</th><th className="py-1 text-right">Units</th><th className="py-1 text-right">Avg ₹</th><th className="py-1 text-right">Lost</th></tr></thead>
                      <tbody>{[...ipParsed].sort((a, b) => (b.net_revenue || 0) - (a.net_revenue || 0)).slice(0, 8).map((r, i) => (
                        <tr key={i} className="border-t border-zinc-800"><td className="py-1.5">{r.name}</td><td className="py-1.5 text-right font-mono">{r.net_revenue != null ? "₹" + Number(r.net_revenue).toLocaleString("en-IN") : "—"}</td><td className="py-1.5 text-right font-mono text-zinc-400">{r.units_sold ?? "—"}</td><td className="py-1.5 text-right font-mono text-zinc-400">{r.avg_price != null ? "₹" + r.avg_price : "—"}</td><td className="py-1.5 text-right font-mono text-zinc-400">{r.lost_orders ?? "—"}</td></tr>
                      ))}</tbody>
                    </table>
                    <button onClick={saveIp} disabled={ipBusy} className="mt-4 bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{ipBusy ? "Saving…" : `Save ${ipParsed.length} items`}</button>
                  </div>
                )}
              </div>
            )}

            <div className="mb-4 flex items-center gap-3 flex-wrap">
              <label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Upload</label>
              <select value={ipSel} onChange={(e) => { setIpSel(e.target.value); fetchIpRows(e.target.value); }} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 text-sm">
                {ipUploads.length === 0 ? <option value="">No uploads yet</option> : ipUploads.map((u) => <option key={u.id} value={u.id}>{u.label} · {new Date(u.created_at).toLocaleDateString("en-IN")} · {u.row_count} items</option>)}
              </select>
              {ipSel && (
                <button onClick={deleteIpUpload} disabled={ipDeleteBusy} className="text-[10px] font-mono text-red-400 uppercase tracking-widest border border-red-900 px-3 py-2 hover:bg-red-950 disabled:opacity-50">
                  {ipDeleteBusy ? "…" : "🗑 Clear data"}
                </button>
              )}
            </div>
            <div className="flex gap-2 mb-6">
              <button onClick={() => setIpView("insights")} className={`px-4 py-2 text-sm font-semibold transition-colors ${ipView === "insights" ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Insights</button>
              <button onClick={() => setIpView("data")} className={`px-4 py-2 text-sm font-semibold transition-colors ${ipView === "data" ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Data</button>
            </div>
            {ipRows.length === 0 ? <p className="text-sm text-zinc-500">No data yet. {canUploadItemPerf ? "Upload an export above." : "Waiting for an upload."}</p> : ipView === "insights" ? (
                           <div>
                {ipStats && (<>
                  <div className="flex items-center justify-between mb-4 max-w-3xl gap-3 flex-wrap">
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">{(ipUploads.find((u) => u.id === ipSel)?.label) || ""} · {ipStats.rows.length} items · {ipStats.inr(ipStats.rows.reduce((s: number, r: any) => s + r.rev, 0))}</p>
                    <button onClick={downloadIpPDF} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors">Download report (PDF)</button>
                  </div>

                  <div className="mb-6 border border-zinc-800 p-5 max-w-3xl">
                    <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Headline</p>
                    <p className="text-lg md:text-xl mb-3">{ipStats.funny.headline}</p>
                    {ipStats.funny.suspect && <p className="text-sm text-orange-400 mb-1">{ipStats.funny.suspect}</p>}
                    {ipStats.funny.sweet && <p className="text-sm text-green-400 mb-1">{ipStats.funny.sweet}</p>}
                    {ipStats.funny.dead && <p className="text-sm text-red-400">{ipStats.funny.dead}</p>}
                  </div>

                  {(() => {
                    const aTier = ipStats.abc.A, bTier = ipStats.abc.B, cTier = ipStats.abc.C;
                    const top8 = ipStats.tiered.slice(0, 8);
                    const tierColor = (t: string) => t === "A" ? "#FACC15" : t === "B" ? "#C89A2E" : "#52525b";
                    return (
                      <div className="mb-8 max-w-3xl">
                        <p className="text-sm font-bold uppercase tracking-widest mb-1">Where the revenue actually comes from</p>
                        <p className="text-xs text-zinc-500 mb-4"><b className="text-zinc-300">{aTier.count} items</b> ({(aTier.count / ipStats.rows.length * 100).toFixed(0)}% of the menu) bring in <b className="text-zinc-300">{aTier.share.toFixed(0)}% of revenue</b>. Protect those before anything else — a price or availability problem here costs far more than the same problem on a C-tier item.</p>
                        <div className="flex h-6 mb-3 gap-0.5">
                          <div style={{ flexBasis: `${aTier.share}%`, background: tierColor("A") }} className="flex items-center justify-center text-[10px] font-bold text-black">{aTier.share >= 8 ? `A ${aTier.share.toFixed(0)}%` : ""}</div>
                          <div style={{ flexBasis: `${bTier.share}%`, background: tierColor("B") }} className="flex items-center justify-center text-[10px] font-bold text-black">{bTier.share >= 8 ? `B ${bTier.share.toFixed(0)}%` : ""}</div>
                          <div style={{ flexBasis: `${cTier.share}%`, background: tierColor("C") }} className="flex items-center justify-center text-[10px] font-bold text-white">{cTier.share >= 8 ? `C ${cTier.share.toFixed(0)}%` : ""}</div>
                        </div>
                        <div className="space-y-1">
                          {top8.map((r: any, i: number) => (
                            <div key={i} className="flex items-baseline gap-3 text-xs">
                              <span className="w-4 text-zinc-600 font-mono">{r.tier}</span>
                              <span className="flex-1 text-zinc-300">{r.name}</span>
                              <span className="font-mono text-zinc-500">{r.share.toFixed(1)}%</span>
                              <span className="font-mono text-zinc-400 w-24 text-right">{ipStats.inr(r.rev)}</span>
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })()}

                  {[
                    { title: "🏆 Stars", sub: "The workhorses — protect these", rows: ipStats.stars, border: "border-green-500/30", explain: (r: any) => `${r.units.toLocaleString("en-IN")} sold at ${ipStats!.inr(r.price)} — above-median demand and one of the biggest revenue lines on the menu. The only real risk here is running out of stock.` },
                    { title: "👀 Priced-too-high suspects", sub: "High price, shy demand — worth a rethink", rows: ipStats.suspects, border: "border-orange-500/30", explain: (r: any) => `Priced above the median for its category but selling below the category's median volume${r.lost ? `, and ${r.lost} lost orders on top` : ""}. Worth testing a small price step-down or a combo to see if demand was actually price-sensitive here.` },
                    { title: "💚 Sweet-spot winners", sub: "Great price, flying off shelves", rows: ipStats.sweet, border: "border-green-500/30", explain: (r: any) => `At or below the median price and still moving ${r.units.toLocaleString("en-IN")} units — this is the price doing its job. Leave it alone rather than "optimising" it.` },
                    { title: "💸 Money left on the table", sub: "People wanted it, didn't get it", rows: ipStats.moneyLeft, border: "border-yellow-500/30", explain: (r: any) => `${r.lost} lost order${r.lost === 1 ? "" : "s"} — people tried to order this and couldn't. Usually a stock/availability problem on the platform, not a demand problem, so it's normally the cheapest fix on this whole page.` },
                    { title: "💀 Dead weight", sub: "Barely moving — rework or retire?", rows: ipStats.dead, border: "border-red-500/30", explain: (r: any) => `Only ${r.units} sold this window, and below-median revenue to match. Either this needs a menu-photo/description refresh, a price cut, or it's honestly not earning its spot on the menu.` },
                  ].map((b) => (
                    <div key={b.title} className={`mb-4 border ${b.border} p-4 max-w-3xl`}>
                      <p className="text-sm font-semibold">{b.title}</p>
                      <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-2">{b.sub}</p>
                      {b.rows.length === 0 ? <p className="text-sm text-zinc-600">None.</p> : (
                        <div className="space-y-3">
                          {b.rows.map((r: any, i: number) => (
                            <div key={i} className="text-sm">
                              <div className="flex items-baseline gap-3">
                                <span className="flex-1 font-semibold">{r.name}</span>
                                <span className="font-mono text-zinc-400">{ipStats!.inr(r.price)}</span>
                                <span className="font-mono text-zinc-300 w-16 text-right">{r.units.toLocaleString("en-IN")}u</span>
                                <span className="font-mono text-zinc-500 w-16 text-right">{r.lost} lost</span>
                              </div>
                              <p className="text-xs text-zinc-500 mt-0.5">{b.explain(r)}</p>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  ))}

                  {(() => {
                    if (!ipComparison) {
                      return (
                        <div className="mb-8 max-w-3xl border border-dashed border-zinc-700 p-5 text-sm text-zinc-500">
                          <p className="font-semibold text-zinc-300 mb-1">Pricing &amp; discount discipline</p>
                          <p>Upload one more period — this section will then check every item against your own rule of thumb: if price went up and sales held or grew, it flags "don't discount this." If a price drop lines up with a sales jump, it flags a possible discount worth double-checking against real margin.</p>
                        </div>
                      );
                    }
                    const buckets: Record<string, any[]> = { priceLedGrowth: [], possibleDiscount: [], priceResistance: [], decliningStable: [] };
                    ipComparison.rows.forEach((r: any) => { if (buckets[r.bucket]) buckets[r.bucket].push(r); });
                    Object.values(buckets).forEach((arr) => arr.sort((a, b) => b.curPrice * b.curUnits - a.curPrice * a.curUnits));
                    const pct = (n: number) => `${n >= 0 ? "+" : ""}${n.toFixed(1)}%`;
                    const rowsFor = (arr: any[], color: string, sentence: (r: any) => string) => arr.slice(0, 6).map((r: any, i: number) => (
                      <div key={i} className={`border-l-2 ${color} bg-zinc-900/60 pl-3 py-2 mb-2 text-sm`}>
                        <span className="font-semibold">{r.name}</span> — {sentence(r)}
                      </div>
                    ));
                    const priceLed = rowsFor(buckets.priceLedGrowth, "border-green-500", (r) => `price moved ${pct(r.priceChg)}, units moved ${pct(r.unitsChg)}. Demand absorbed the increase — don't discount this one.`);
                    const possibleDisc = rowsFor(buckets.possibleDiscount, "border-orange-400", (r) => `price fell ${pct(r.priceChg)} while units jumped ${pct(r.unitsChg)}. Looks like a discount or platform offer ran here — worth checking it actually made margin.`);
                    const resistance = rowsFor(buckets.priceResistance, "border-red-500", (r) => `price went up ${pct(r.priceChg)} and units dropped ${pct(r.unitsChg)}. The increase wasn't absorbed — reconsider it or pair it with something that justifies the new price.`);
                    const declining = rowsFor(buckets.decliningStable, "border-orange-400", (r) => `price barely moved (${pct(r.priceChg)}) but units fell ${pct(r.unitsChg)}. Not a pricing problem — something else is pulling demand away.`);
                    const dropped = ipComparison.dropped.slice(0, 6);
                    const nothing = !priceLed.length && !possibleDisc.length && !resistance.length && !declining.length && !dropped.length;
                    return (
                      <div className="mb-8 max-w-3xl">
                        <p className="text-sm font-bold uppercase tracking-widest mb-1">Pricing &amp; discount discipline</p>
                        <p className="text-xs text-zinc-500 mb-4">If price went up and people still bought it, don't discount it. If a price drop lines up with a volume jump, that's worth double-checking.</p>
                        {nothing && <p className="text-sm text-zinc-600">Nothing swung hard enough between the two uploads to flag.</p>}
                        {!!priceLed.length && <><p className="text-[11px] font-mono text-green-400 uppercase tracking-widest mb-1">Priced up, demand held — don't discount</p>{priceLed}</>}
                        {!!possibleDisc.length && <><p className="text-[11px] font-mono text-orange-400 uppercase tracking-widest mb-1 mt-3">Possible discounting — verify the margin</p>{possibleDisc}</>}
                        {!!resistance.length && <><p className="text-[11px] font-mono text-red-400 uppercase tracking-widest mb-1 mt-3">Price increase not absorbed</p>{resistance}</>}
                        {!!declining.length && <><p className="text-[11px] font-mono text-orange-400 uppercase tracking-widest mb-1 mt-3">Falling without a price change</p>{declining}</>}
                        {!!dropped.length && (
                          <div className="border-l-2 border-red-500 bg-zinc-900/60 pl-3 py-2 mt-3 text-sm">
                            <span className="font-semibold">{dropped.length} item{dropped.length > 1 ? "s" : ""} sold last period, nothing this period:</span> {dropped.map((d: any) => d.name).join(", ")}. Worth checking whether these were pulled off the menu on purpose.
                          </div>
                        )}
                      </div>
                    );
                  })()}

                  {(() => {
                    if (!ipStats.margin.length) {
                      return (
                        <div className="mb-8 max-w-3xl border border-dashed border-zinc-700 p-5 text-sm text-zinc-500">
                          <p className="font-semibold text-zinc-300 mb-1">Cost &amp; margin</p>
                          <p>None of this upload's items matched the costing sheet by name — check src/lib/itemPerfCosts.ts covers this menu, or that item names here are close to the ones in Master Pricing.</p>
                        </div>
                      );
                    }
                    const thin = ipStats.margin.filter((m: any) => m.marginPct < 15);
                    const avg = ipStats.margin.reduce((s: number, m: any) => s + m.marginPct, 0) / ipStats.margin.length;
                    return (
                      <div className="mb-8 max-w-3xl">
                        <p className="text-sm font-bold uppercase tracking-widest mb-1">Cost &amp; margin</p>
                        <p className="text-xs text-zinc-500 mb-4">Real landed cost from Master Pricing, matched to {ipStats.marginCoverage.matched} of {ipStats.marginCoverage.total} items on this upload ({ipStats.marginCoverage.revSharePct.toFixed(0)}% of revenue) — average margin across those: <b className="text-zinc-300">{avg.toFixed(1)}%</b>.</p>
                        {thin.length === 0 ? (
                          <p className="text-sm text-green-400">No thin margins in what matched — everything costed is clearing at least 15%.</p>
                        ) : thin.slice(0, 8).map((m: any, i: number) => (
                          <div key={i} className="border-l-2 border-red-500 bg-zinc-900/60 pl-3 py-2 mb-2 text-sm">
                            <span className="font-semibold">{m.name}</span> — sells at {ipStats!.inr(m.price)}, costs {ipStats!.inr(m.cost)} to make. That's {m.marginPct.toFixed(1)}% margin{m.marginPct < 0 ? " — currently sold at a loss" : ""}. If cocoa, Maida or butter has moved recently, this is one of the first prices worth revisiting.
                          </div>
                        ))}
                      </div>
                    );
                  })()}

                  {(() => {
                    const mq = ipStats.marginQuadrant;
                    if (!mq.rows.length) {
                      return (
                        <div className="mb-8 max-w-3xl border border-dashed border-zinc-700 p-5 text-sm text-zinc-500">
                          <p className="font-semibold text-zinc-300 mb-1">Menu engineering — margin × volume</p>
                          <p>No items matched the costing sheet, so this classic profitability-vs-popularity view can't be built yet. See "Cost & margin" above.</p>
                        </div>
                      );
                    }
                    const quads = [
                      { key: "stars", title: "🌟 Stars", sub: "High margin, high volume — protect & feature these", rows: mq.stars, border: "border-green-500/30", explain: (r: any) => `${r.marginPct.toFixed(0)}% margin, ${r.units.toLocaleString("en-IN")} sold. Your most profitable big-sellers — the menu's actual engine.` },
                      { key: "workhorses", title: "🐴 Workhorses", sub: "Low margin, high volume — popular but thin", rows: mq.workhorses, border: "border-yellow-500/30", explain: (r: any) => `${r.units.toLocaleString("en-IN")} sold but only ${r.marginPct.toFixed(0)}% margin. People love it, but it barely pays — a small price nudge here does more than one on a slow item.` },
                      { key: "puzzles", title: "🧩 Puzzles", sub: "High margin, low volume — underexposed", rows: mq.puzzles, border: "border-blue-500/30", explain: (r: any) => `${r.marginPct.toFixed(0)}% margin but only ${r.units.toLocaleString("en-IN")} sold. Profitable when it sells — the problem is visibility/positioning, not the recipe or price.` },
                      { key: "dogs", title: "🐶 Dogs", sub: "Low margin, low volume — reconsider", rows: mq.dogs, border: "border-red-500/30", explain: (r: any) => `Only ${r.marginPct.toFixed(0)}% margin and ${r.units.toLocaleString("en-IN")} sold. Doing the least amount of work on the menu — rework the recipe/price, or retire it.` },
                    ];
                    return (
                      <div className="mb-8">
                        <p className="text-sm font-bold uppercase tracking-widest mb-1">Menu engineering — margin × volume</p>
                        <p className="text-xs text-zinc-500 mb-4">A different lens from the price/demand view above — this one classifies by REAL profitability (landed cost vs price) against how much it actually sells, split at this upload's own median ({mq.medianMarginPct.toFixed(0)}% margin, {mq.medianUnits.toLocaleString("en-IN")} units). Only the {mq.rows.length} items matched to the costing sheet are included.</p>
                        {quads.map((b) => (
                          <div key={b.key} className={`mb-4 border ${b.border} p-4 max-w-3xl`}>
                            <p className="text-sm font-semibold">{b.title}</p>
                            <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-2">{b.sub}</p>
                            {b.rows.length === 0 ? <p className="text-sm text-zinc-600">None.</p> : (
                              <div className="space-y-3">
                                {b.rows.slice(0, 6).map((r: any, i: number) => (
                                  <div key={i} className="text-sm">
                                    <div className="flex items-baseline gap-3">
                                      <span className="flex-1 font-semibold">{r.name}</span>
                                      <span className="font-mono text-zinc-400">{ipStats!.inr(r.price)}</span>
                                      <span className="font-mono text-zinc-300 w-16 text-right">{r.units.toLocaleString("en-IN")}u</span>
                                      <span className="font-mono text-zinc-500 w-14 text-right">{r.marginPct.toFixed(0)}%</span>
                                    </div>
                                    <p className="text-xs text-zinc-500 mt-0.5">{b.explain(r)}</p>
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                        ))}
                      </div>
                    );
                  })()}

                  {(() => {
                    const paras: string[] = [];
                    const aTier = ipStats.abc.A;
                    paras.push(`The menu is carried by a small core: ${aTier.count} items (${aTier.share.toFixed(0)}% of revenue) do most of the work. That's the first place to look when deciding what to feature, restock aggressively, or protect from a stock-out — a problem on one of these costs far more than the same problem on a C-tier item.`);
                    if (ipComparison) {
                      const n = ipComparison.rows.filter((r: any) => r.bucket === "priceLedGrowth").length;
                      const d = ipComparison.rows.filter((r: any) => r.bucket === "possibleDiscount").length;
                      if (n) paras.push(`${n} item${n > 1 ? "s" : ""} took a price increase and kept or grew volume since the last upload — the textbook case for leaving well alone rather than discounting.`);
                      if (d) paras.push(`${d} item${d > 1 ? "s" : ""} show the opposite shape — price down, volume up — worth checking those against real margin before running the same offer again.`);
                    } else {
                      paras.push(`This read is from a single upload, so it can say what's selling well right now but not yet whether a price move is being absorbed or resisted. Upload a second period and this section sharpens considerably.`);
                    }
                    if (ipStats.margin.length) {
                      const thinCount = ipStats.margin.filter((m: any) => m.marginPct < 15).length;
                      if (thinCount) paras.push(`On cost: ${thinCount} item${thinCount > 1 ? "s are" : " is"} running under 15% margin against real landed cost. If ingredient prices have moved, the fix is usually a small price adjustment on exactly these items, not a blanket increase across the menu.`);
                    }
                    return (
                      <div className="mb-4 max-w-3xl border border-zinc-800 bg-zinc-900/40 p-5">
                        <p className="text-sm font-bold uppercase tracking-widest mb-3">Improvement actions</p>
                        {paras.map((p, i) => <p key={i} className="text-sm text-zinc-300 leading-relaxed mb-3 last:mb-0">{p}</p>)}
                      </div>
                    );
                  })()}
                </>)}
              </div>
            ) : (
              <div className="overflow-x-auto max-w-5xl border border-zinc-800">
                <table className="w-full text-sm">
                  <thead><tr className="text-left text-[11px] font-mono text-zinc-500 uppercase bg-zinc-900"><th className="py-2 px-3">Item</th><th className="py-2 px-3">Category</th><th className="py-2 px-3 text-right">Revenue</th><th className="py-2 px-3 text-right">Units</th><th className="py-2 px-3 text-right">Avg ₹</th><th className="py-2 px-3 text-right">Lost</th><th className="py-2 px-3 text-right">Orders/day</th></tr></thead>
                  <tbody>{ipRows.map((r) => (
                    <tr key={r.id} className="border-t border-zinc-800"><td className="py-1.5 px-3">{r.name}</td><td className="py-1.5 px-3 text-zinc-500 text-xs">{r.category}</td><td className="py-1.5 px-3 text-right font-mono">{r.net_revenue != null ? "₹" + Number(r.net_revenue).toLocaleString("en-IN") : "—"}</td><td className="py-1.5 px-3 text-right font-mono text-zinc-400">{r.units_sold ?? "—"}</td><td className="py-1.5 px-3 text-right font-mono text-zinc-400">{r.avg_price != null ? "₹" + r.avg_price : "—"}</td><td className="py-1.5 px-3 text-right font-mono text-zinc-400">{r.lost_orders ?? "—"}</td><td className="py-1.5 px-3 text-right font-mono text-zinc-400">{r.avg_orders_day ?? "—"}</td></tr>
                  ))}</tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {activeTab === "competition" && (
          <div>
            <div className="mb-8 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Competition</h2>
             <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Rival bakery sales · expansion signals</p>
            </div>

            <div className="flex gap-2 mb-6">
              <button onClick={() => setCompView("entry")} className={`px-4 py-2 text-sm font-semibold transition-colors ${compView === "entry" ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Entry</button>
              <button onClick={() => setCompView("products")} className={`px-4 py-2 text-sm font-semibold transition-colors ${compView === "products" ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Products</button>
              <button onClick={() => setCompView("insights")} className={`px-4 py-2 text-sm font-semibold transition-colors ${compView === "insights" ? "bg-yellow-400 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Insights</button>
            </div>

            {compView === "products" && (
              <div>
                <div className="mb-8 border border-zinc-800 p-5 max-w-3xl">
                  <p className="text-sm font-semibold mb-1">Paste a top-products table</p>
                  <p className="text-xs text-zinc-500 mb-4">Rank | Product | GMV | Orders | Units | Areas. Pick whose list it is first.</p>
                  <div className="flex gap-2 mb-3">
                    <button onClick={() => setProdSide("them")} className={`px-3 py-1.5 text-sm font-semibold transition-colors ${prodSide === "them" ? "bg-orange-500 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Their products</button>
                    <button onClick={() => setProdSide("us")} className={`px-3 py-1.5 text-sm font-semibold transition-colors ${prodSide === "us" ? "bg-green-500 text-black" : "bg-zinc-900 text-zinc-400 hover:text-white"}`}>Our products</button>
                  </div>
                  <textarea value={prodPaste} onChange={(e) => setProdPaste(e.target.value)} rows={6} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm font-mono" placeholder="| Rank | Product | GMV | Orders | Units | Areas |" />
                  <div className="flex flex-wrap gap-3 mt-3 items-end">
                    {prodSide === "them" && <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Competitor</label><input type="text" value={prodComp} onChange={(e) => setProdComp(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-44" placeholder="Brownie Studio" /></div>}
                    <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period label</label><input type="text" value={prodLabel} onChange={(e) => setProdLabel(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-40" placeholder="June 2026" /></div>
                    <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period date</label><input type="date" value={prodDate} onChange={(e) => setProdDate(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" /></div>
                    <button onClick={parseProd} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors">Parse</button>
                  </div>
                  {prodParsed && (
                    <div className="mt-5">
                      <table className="w-full text-sm">
                        <thead><tr className="text-left text-[11px] font-mono text-zinc-500 uppercase"><th className="py-1">#</th><th className="py-1">Product</th><th className="py-1 text-right">GMV</th><th className="py-1 text-right">Orders</th><th className="py-1 text-right">Units</th><th className="py-1 text-right">Areas</th></tr></thead>
                        <tbody>
                          {prodParsed.map((r, i) => (
                            <tr key={i} className="border-t border-zinc-800">
                              <td className="py-1.5 text-zinc-500">{i + 1}</td>
                              <td className="py-1.5">{r.product}</td>
                              <td className="py-1.5 text-right font-mono">{r.gmv != null ? compFmtL(r.gmv) : "—"}</td>
                              <td className="py-1.5 text-right font-mono text-zinc-400">{r.orders != null ? r.orders.toLocaleString("en-IN") : "—"}</td>
                              <td className="py-1.5 text-right font-mono text-zinc-400">{r.units != null ? r.units.toLocaleString("en-IN") : "—"}</td>
                              <td className="py-1.5 text-right font-mono text-zinc-400">{r.areas != null ? r.areas : "—"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <button onClick={saveProd} disabled={prodSaving} className="mt-4 bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{prodSaving ? "Saving…" : `Save ${prodParsed.length} products (${prodSide === "them" ? "theirs" : "ours"})`}</button>
                    </div>
                  )}
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 max-w-4xl">
                  {[{ side: "them", title: "🟠 Their top sellers", color: "text-orange-400" }, { side: "us", title: "🟢 Our top sellers", color: "text-green-400" }].map((col) => {
                    const rows = compProducts.filter((r) => r.side === col.side).sort((a, b) => (Number(b.gmv) || 0) - (Number(a.gmv) || 0)).slice(0, 15);
                    return (
                      <div key={col.side} className="border border-zinc-800 p-4">
                        <p className={`text-sm font-semibold mb-3 ${col.color}`}>{col.title}</p>
                        {rows.length === 0 ? <p className="text-sm text-zinc-600">Nothing saved yet.</p> : (
                          <div className="space-y-1.5">
                            {rows.map((r, i) => (
                              <div key={r.id} className="flex items-baseline gap-2 text-sm">
                                <span className="text-zinc-600 w-5">{i + 1}</span>
                                <span className="flex-1">{r.product}</span>
                                <span className="font-mono text-zinc-300">{r.gmv != null ? compFmtL(Number(r.gmv)) : "—"}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {compView === "insights" && (
              <div>
                <div className="mb-6 flex items-center gap-3 flex-wrap">
                  {compPeriods.length > 0 && (<><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period</label>
                    <select value={compActivePeriod} onChange={(e) => setCompPeriod(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm focus:outline-none focus:border-yellow-400">
                      {compPeriods.map((pl) => <option key={pl} value={pl}>{pl}</option>)}
                    </select></>)}
                  {compPeriodRows.length > 0 && <button onClick={downloadCompPDF} className="ml-auto bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors">Download report (PDF)</button>}
                </div>
                {compPeriodRows.length === 0 ? (
                  <p className="text-sm text-zinc-500">No competitor data yet. Add some in the Entry tab.</p>
                ) : (
                  <>
                    {compTop && (
                      <div className="mb-8 border border-zinc-800 p-5 max-w-2xl">
                        <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-3">Headline</p>
                        <p className="text-lg md:text-xl">{compFunny.playful}</p>
                      </div>
                    )}
                    {[
                      { title: "🟠 Expansion candidates", sub: "They lead where we're absent", rows: compExpansion, color: "text-orange-400", border: "border-orange-500/30" },
                      { title: "🔴 Defend", sub: "We're here but trailing", rows: compDefend, color: "text-red-400", border: "border-red-500/30" },
                      { title: "🟢 Winning", sub: "We're ahead", rows: compWinning, color: "text-green-400", border: "border-green-500/30" },
                    ].map((b) => (
                      <div key={b.title} className={`mb-6 border ${b.border} p-4 max-w-3xl`}>
                        <p className="text-sm font-semibold">{b.title}</p>
                        <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mb-3">{b.sub}</p>
                        {b.rows.length === 0 ? <p className="text-sm text-zinc-600">None this period.</p> : (
                          <div className="space-y-1.5">
                            {b.rows.map((r, i) => (
                              <div key={i} className="flex flex-wrap gap-x-4 items-baseline text-sm">
                                <span className="font-medium w-32">{r.area}</span>
                                <span className="text-zinc-400 w-36">{r.competitor}</span>
                                <span className="font-mono text-zinc-300">them {compFmtL(r.their)}</span>
                                <span className="font-mono text-zinc-500">us {r.our != null ? compFmtL(r.our) : "—"}</span>
                                <span className={`font-mono ml-auto ${b.color}`}>{compFmtL(Math.abs(r.gap))}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    ))}
                  </>
                )}
              </div>
            )}

            {compView === "entry" && (<>
            <div className="mb-8 border border-zinc-800 p-5 max-w-3xl">
              <p className="text-sm font-semibold mb-1">Paste sir&apos;s area comparison table</p>
              <p className="text-xs text-zinc-500 mb-4">Paste the &quot;Area | … GMV | Brownie Heaven GMV | Gap&quot; table. We parse the rows and compute the gap — you just review and save.</p>
              <textarea value={compPaste} onChange={(e) => setCompPaste(e.target.value)} rows={6} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm font-mono" placeholder="| Area | Brownie Studio GMV | Brownie Heaven GMV | Gap |" />
              <div className="flex flex-wrap gap-3 mt-3 items-end">
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Competitor (fallback)</label><input type="text" value={compPasteComp} onChange={(e) => setCompPasteComp(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-44" placeholder="only if table has no name" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period label</label><input type="text" value={compPasteLabel} onChange={(e) => setCompPasteLabel(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1 w-40" placeholder="June 2026" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period date</label><input type="date" value={compPasteDate} onChange={(e) => setCompPasteDate(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" /></div>
                <button onClick={parseComp} className="bg-zinc-800 text-white px-4 py-2 text-sm font-semibold hover:bg-zinc-700 transition-colors">Parse</button>
              </div>
              {compParsed && (
                <div className="mt-5">
                  <table className="w-full text-sm">
                    <thead><tr className="text-left text-[11px] font-mono text-zinc-500 uppercase"><th className="py-1">Area</th><th className="py-1">Competitor</th><th className="py-1 text-right">Them</th><th className="py-1 text-right">Us</th><th className="py-1 text-right">Gap</th></tr></thead>
                    <tbody>
                      {compParsed.map((r, i) => { const gap = (r.their || 0) - (r.our || 0); return (
                        <tr key={i} className="border-t border-zinc-800">
                          <td className="py-1.5">{r.area}</td>
                          <td className="py-1.5">{r.competitor}</td>
                          <td className="py-1.5 text-right font-mono">{r.their != null ? "₹" + r.their.toLocaleString("en-IN") : "—"}</td>
                          <td className="py-1.5 text-right font-mono text-zinc-400">{r.our != null ? "₹" + r.our.toLocaleString("en-IN") : "—"}</td>
                          <td className={`py-1.5 text-right font-mono ${gap > 0 ? "text-orange-400" : "text-green-400"}`}>{gap > 0 ? "they +" : "we +"}₹{Math.abs(gap).toLocaleString("en-IN")}</td>
                        </tr>); })}
                    </tbody>
                  </table>
                  <button onClick={saveParsedComp} disabled={compSavingBulk} className="mt-4 bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{compSavingBulk ? "Saving…" : `Save ${compParsed.length} rows`}</button>
                </div>
              )}
            </div>

            <div className="mb-8 border border-zinc-800 p-5 max-w-2xl">
              <p className="text-sm font-semibold mb-4">Log a competitor&apos;s numbers</p>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Competitor *</label><input type="text" value={compForm.competitor} onChange={(e) => setCompForm(pp => ({ ...pp, competitor: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="The Brownie Studio" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Area / locality</label><input type="text" value={compForm.area} onChange={(e) => setCompForm(pp => ({ ...pp, area: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="Perambur" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Specific address (optional)</label><input type="text" value={compForm.address} onChange={(e) => setCompForm(pp => ({ ...pp, address: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="123 Main Rd, Perambur" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Our outlet in this area</label><select value={compForm.our_outlet_id} onChange={(e) => setCompForm(pp => ({ ...pp, our_outlet_id: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1"><option value="">— none (we&apos;re not here) —</option>{OUTLETS.map((o) => <option key={o} value={o}>{(OUTLET_NAMES as any)[o] || o}</option>)}</select></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Their sales (₹)</label><input type="number" value={compForm.sales_value} onChange={(e) => setCompForm(pp => ({ ...pp, sales_value: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="800000" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Our sales here (₹, optional)</label><input type="number" value={compForm.our_sales} onChange={(e) => setCompForm(pp => ({ ...pp, our_sales: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="—" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period label</label><input type="text" value={compForm.period_label} onChange={(e) => setCompForm(pp => ({ ...pp, period_label: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="July 2026" /></div>
                <div><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Period date</label><input type="date" value={compForm.period_date} onChange={(e) => setCompForm(pp => ({ ...pp, period_date: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" /></div>
                <div className="md:col-span-2"><label className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest">Note</label><input type="text" value={compForm.note} onChange={(e) => setCompForm(pp => ({ ...pp, note: e.target.value }))} className="w-full bg-black border border-zinc-800 text-white px-3 py-2 focus:outline-none focus:border-yellow-400 transition-colors text-sm mt-1" placeholder="Sir&apos;s comment / context" /></div>
              </div>
              <button onClick={saveComp} disabled={compSaving} className="mt-4 bg-yellow-400 text-black px-5 py-2 text-sm font-semibold hover:bg-yellow-300 disabled:opacity-50 transition-colors">{compSaving ? "Saving…" : "Save entry"}</button>
            </div>

            <div className="max-w-3xl">
              <p className="text-sm font-semibold mb-3">Recent entries</p>
              {compRows.length === 0 ? (
                <p className="text-sm text-zinc-500">No entries yet.</p>
              ) : (
                <div className="space-y-2">
                  {compRows.slice(0, 10).map((r) => (
                    <div key={r.id} className="border border-zinc-800 px-4 py-3 text-sm flex flex-wrap gap-x-4 gap-y-1 items-baseline">
                      <span className="font-semibold">{r.competitor}</span>
                      {r.area && <span className="text-zinc-400">{r.area}</span>}
                      {r.sales_value != null && <span className="text-yellow-400 font-mono">₹{Number(r.sales_value).toLocaleString("en-IN")}</span>}
                      {r.our_sales != null && <span className="text-zinc-500 font-mono">us: ₹{Number(r.our_sales).toLocaleString("en-IN")}</span>}
                      {!r.our_outlet_id && <span className="text-[10px] font-mono text-orange-400 uppercase">no outlet here</span>}
                      {r.period_label && <span className="text-zinc-600 text-xs">{r.period_label}</span>}
                      {r.note && <span className="text-zinc-500 text-xs italic w-full">{r.note}</span>}
                    </div>
                  ))}
                </div>
              )}
            </div>
            </>)}
          </div>
        )}

        {activeTab === "analytics" && (
          <div>
            <div className="mb-8 pb-5 border-b border-zinc-800">
              <h2 className="text-2xl font-black tracking-tight">Analytics</h2>
                        <p className="text-[11px] font-mono text-zinc-500 uppercase tracking-widest mt-1">Sales performance · channel mix</p>
            </div>

            {/* Atlas-style overview grid — click a square to jump to its detail below */}
            <div className="mb-10 space-y-6">
              {[
                { heading: "💰 Revenue", accent: "#FACC15", cards: [
                  { id: "an-revenue", title: "Revenue Received", value: anINR(anAgg.totalV), sub: `${anRows.length} report-days` },
                  { id: "an-channel", title: "Revenue by Channel", value: `${anPct(anAgg.ch.swiggy.v + anAgg.ch.zomato.v, anAgg.totalV)}% online`, sub: "Shop / Swiggy / Zomato split" },
                  { id: "an-outlet", title: "Revenue by Outlet", value: anAgg.outletRanked[0] ? anAgg.outletRanked[0].name : "—", sub: anAgg.outletRanked[0] ? `top of ${anAgg.outletRanked.length} · ${anINR(anAgg.outletRanked[0].v)} · see full list ↓` : "no data" },
                ]},
                { heading: "🛒 Orders", accent: "#60A5FA", cards: [
                  { id: "an-revenue", title: "Orders Received", value: anAgg.totalC.toLocaleString("en-IN"), sub: "across all channels" },
                  { id: "an-channel", title: "Average Order Value", value: anINR(anAgg.totalC ? anAgg.totalV / anAgg.totalC : 0), sub: "blended, all channels" },
                  { id: "an-channel", title: "Orders by Channel", value: `${anAgg.ch.shop.c}/${anAgg.ch.swiggy.c}/${anAgg.ch.zomato.c}`, sub: "Shop / Swiggy / Zomato" },
                ]},
                { heading: "⚙️ Operations", accent: "#FB923C", cards: [
                  { id: "an-discount", title: "Discounts Given", value: anINR(anAgg.discountTotal), sub: `${anPct(anAgg.discountTotal, anAgg.totalV)}% of revenue` },
                  { id: "an-funnel", title: "Order Completion Funnel", value: "🔒 Needs Atlas export", sub: "Pre-Ack / Post-Ack cancellations", locked: true },
                  { id: "an-lost", title: "Lost Orders", value: "🔒 Needs Atlas export", sub: "not visible to staff entry", locked: true },
                ]},
              ].map((section) => (
                <div key={section.heading}>
                  <p className="text-[10px] font-mono uppercase tracking-widest mb-2" style={{ color: section.accent }}>{section.heading}</p>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    {section.cards.map((c, i) => (
                      <button
                        key={c.title + i}
                        onClick={() => { setAnCard(c.id); document.getElementById(c.id)?.scrollIntoView({ behavior: "smooth", block: "start" }); }}
                        style={{ borderTopColor: c.locked ? "#78716c" : section.accent, borderTopWidth: 3, borderTopStyle: c.locked ? "dashed" : "solid" }}
                        className={`text-left bg-[#131316] border border-zinc-800 p-4 transition-colors ${c.locked ? "opacity-70 hover:border-amber-700/50" : "hover:border-zinc-600"}`}
                      >
                        <p className="text-[9px] font-mono text-zinc-500 uppercase tracking-widest mb-1.5">{c.title}</p>
                        <p className={`font-black tracking-tight mb-1 ${c.locked ? "text-sm text-amber-500/80" : "text-lg"}`}>{c.value}</p>
                        <p className="text-[10px] text-zinc-600">{c.sub}</p>
                      </button>
                    ))}
                  </div>
                </div>
              ))}
            </div>

            {/* Sales performance (date range) */}
            <div id="an-revenue" className="mb-10">
              <div className="flex flex-wrap items-end gap-3 mb-6">
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">From</label>
                  <input type="date" value={anFrom} onChange={(e) => setAnFrom(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm focus:outline-none focus:border-yellow-400 transition-colors" />
                </div>
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-1">To</label>
                  <input type="date" value={anTo} onChange={(e) => setAnTo(e.target.value)} className="bg-black border border-zinc-800 text-white px-3 py-2 text-sm focus:outline-none focus:border-yellow-400 transition-colors" />
                </div>
                <span className="text-[10px] font-mono text-zinc-600 pb-2">{anLoading ? "Loading…" : `${anRows.length} report-days`}</span>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Total Sales</p>
                  <p className="text-2xl font-black tracking-tight">{anINR(anAgg.totalV)}</p>
                </div>
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Total Orders</p>
                  <p className="text-2xl font-black tracking-tight">{anAgg.totalC.toLocaleString("en-IN")}</p>
                </div>
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Avg Order Value</p>
                  <p className="text-2xl font-black tracking-tight">{anINR(anAgg.totalC ? anAgg.totalV / anAgg.totalC : 0)}</p>
                </div>
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Delivery Share</p>
                  <p className="text-2xl font-black tracking-tight">{anPct(anAgg.ch.swiggy.v + anAgg.ch.zomato.v, anAgg.totalV)}%</p>
                </div>
              </div>

              <div id="an-channel" className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {([["shop", "Shop / Store", "text-blue-400", "bg-blue-400"], ["swiggy", "Swiggy", "text-orange-400", "bg-orange-400"], ["zomato", "Zomato", "text-red-400", "bg-red-400"]] as const).map(([key, label, tc, bc]) => {
                  const c = anAgg.ch[key];
                  const share = anPct(c.v, anAgg.totalV);
                  const aov = c.c ? c.v / c.c : 0;
                  return (
                    <div key={key} className="bg-[#131316] border border-zinc-800 p-5">
                      <div className="flex justify-between items-center mb-3">
                        <p className={`text-[10px] font-mono uppercase tracking-widest ${tc}`}>{label}</p>
                        <span className="font-mono text-xs text-zinc-500">{share}%</span>
                      </div>
                      <p className="text-xl font-black tracking-tight mb-2">{anINR(c.v)}</p>
                      <div className="h-2 bg-zinc-800 border border-zinc-700 mb-3"><div className={`h-full ${bc}`} style={{ width: `${share}%` }} /></div>
                      <div className="flex justify-between text-xs font-mono text-zinc-500">
                        <span>{c.c.toLocaleString("en-IN")} orders</span>
                        <span>AOV {anINR(aov)}</span>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div id="an-outlet" className="mb-10">
              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Revenue by Outlet — this range, ranked</p>
              {anAgg.outletRanked.length === 0 ? (
                <p className="text-sm text-zinc-600">No data for this range.</p>
              ) : (
                <div className="space-y-2">
                  {anAgg.outletRanked.map((o, i) => {
                    const share = anPct(o.v, anAgg.totalV);
                    return (
                      <div key={o.oid} className="flex items-center gap-3 bg-[#131316] border border-zinc-800 p-3">
                        <span className="text-xs font-mono text-zinc-600 w-5">{i + 1}</span>
                        <span className="text-sm flex-1">{o.name}</span>
                        <div className="h-2 bg-zinc-800 border border-zinc-700 w-32"><div className="h-full bg-yellow-400" style={{ width: `${share}%` }} /></div>
                        <span className="text-xs font-mono text-zinc-500 w-12 text-right">{share}%</span>
                        <span className="text-sm font-mono w-28 text-right">{anINR(o.v)}</span>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            <div id="an-discount" className="mb-10">
              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">Discounts Given — this range</p>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 max-w-2xl">
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Total Discounts</p>
                  <p className="text-xl font-black tracking-tight">{anINR(anAgg.discountTotal)}</p>
                </div>
                <div className="bg-[#131316] border border-zinc-800 p-5">
                  <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">% of Revenue</p>
                  <p className="text-xl font-black tracking-tight">{anPct(anAgg.discountTotal, anAgg.totalV)}%</p>
                </div>
              </div>
            </div>

            <div id="an-funnel" className="mb-10">
              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">🔒 Order Completion Funnel</p>
              <div className="bg-[#131316] border border-zinc-800 p-5 max-w-xl">
                <p className="text-sm text-zinc-400">This needs order-level data (Received → Completed → Cancelled Pre-Ack / Post-Ack) that only exists inside Atlas/UrbanPiper — staff daily entry only captures totals, not individual order outcomes. Once an Atlas export upload is built (same pattern as Item Performance), this section will show a real funnel here.</p>
              </div>
            </div>

            <div id="an-lost" className="mb-10">
              <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-4">🔒 Lost Orders</p>
              <div className="bg-[#131316] border border-zinc-800 p-5 max-w-xl">
                <p className="text-sm text-zinc-400">Same limitation as above — lost/cancelled order counts aren't visible to staff at entry time, only inside Atlas. Will populate once Atlas exports can be uploaded.</p>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div className="bg-[#131316] border border-zinc-800 p-6">
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-6">Performance by Staff · this month</p>
                {scoreRows.length === 0 && <p className="text-zinc-600 font-mono text-xs">Loading scores…</p>}
                {(() => {
                  const pmax = Math.max(...scoreRows.map(x => x.points), 1);
                  const pmin = Math.min(...scoreRows.map(x => x.points), 0);
                  const span = (pmax - pmin) || 1;
               return scoreRows.map((r, i) => {
                    const w = Math.max(3, Math.round((r.points - pmin) / span * 100));
                    return (
                      <div key={r.id} className="mb-4">
                        <div className="flex justify-between mb-1">
                          <span className="text-sm font-medium">{i + 1}. {r.name.split(" ")[0]}</span>
                         <span className="font-mono text-xs text-zinc-500">{r.dailyToday === "done" && <span className="text-green-400 mr-2">report ✓</span>}{r.dailyToday === "missed" && <span className="text-red-500 mr-2">report ✗</span>}{r.dailyToday === "pending" && <span className="text-zinc-500 mr-2">report …</span>}{r.dailyToday === "off" && <span className="text-zinc-600 mr-2">off</span>}{r.points} pts</span>
                        </div>
                        <div className="h-2 bg-zinc-800 border border-zinc-700">
                          <div className="h-full transition-all bg-zinc-500" style={{ width: `${w}%` }} />
                        </div>
                      </div>
                    );
                  });
                })()}
              </div>
              <div className="bg-[#131316] border border-zinc-800 p-6">
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-6">Task Breakdown</p>
                {[
                  { label: "Total Assigned", value: tasks.length, color: "text-white" },
                  { label: "Completed", value: tasks.filter(t => t.status === "completed").length, color: "text-green-400" },
                  { label: "In Progress", value: tasks.filter(t => t.status === "in_progress").length, color: "text-yellow-400" },
                  { label: "Not Started", value: tasks.filter(t => t.status === "assigned").length, color: "text-zinc-400" },
                  { label: "Overdue", value: tasks.filter(t => t.status !== "completed" && new Date(t.due_at) < new Date()).length, color: "text-red-500" },
                ].map(s => (
                  <div key={s.label} className="flex justify-between items-center py-3 border-b border-zinc-800 last:border-0">
                    <span className="text-sm text-zinc-400">{s.label}</span>
                    <span className={`font-mono font-bold text-lg ${s.color}`}>{s.value}</span>
                  </div>
                ))}
              </div>
              <div className="bg-[#131316] border border-zinc-800 p-6">
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-6">Reports Today</p>
                {ALL_STAFF.filter(s => s.id !== "nishant").map(s => {
                  const todayStr = new Date().toDateString();
                  const report = reports.find(r => r.staff_id === s.id && new Date(r.submitted_at).toDateString() === todayStr);
                  return (
                    <div key={s.id} className="flex justify-between items-center py-3 border-b border-zinc-800 last:border-0">
                      <span className="text-sm font-medium">{s.name.split(" ")[0]}</span>
                      <span className={`font-mono text-[10px] uppercase tracking-widest px-2 py-1 ${report ? report.is_late ? "bg-red-500/10 text-red-500" : "bg-green-400/10 text-green-400" : "bg-zinc-800 text-zinc-500"}`}>
                        {report ? report.is_late ? "Late" : "On Time" : "Pending"}
                      </span>
                    </div>
                  );
                })}
              </div>
              <div className="bg-[#131316] border border-zinc-800 p-6">
                <p className="text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-6">Tasks by Outlet</p>
                {OUTLETS.map(o => {
                  const count = tasks.filter(t => t.outlet_id === o).length;
                  return (
                    <div key={o} className="flex justify-between items-center py-2 border-b border-zinc-800 last:border-0">
                      <span className="text-sm text-zinc-400 capitalize">{OUTLET_NAMES[o] || o.replace(/_/g, " ")}</span>
                      <span className="font-mono text-sm font-bold">{count}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

      </main>

      {showModal && (
        <div className="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[500px] bg-[#131316] border border-zinc-800 p-6 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-center mb-6 pb-4 border-b border-zinc-800">
              <h3 className="text-xl font-black tracking-tight">Assign New Task</h3>
              <button onClick={() => setShowModal(false)} className="text-zinc-500 hover:text-white text-xl">✕</button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Task Title *</label>
                <input value={taskTitle} onChange={(e) => setTaskTitle(e.target.value)} placeholder="e.g. Check outlet stock" className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm" />
              </div>
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Description</label>
                <textarea value={taskDesc} onChange={(e) => setTaskDesc(e.target.value)} placeholder="Optional details..." rows={2} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm resize-none" />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Assign To</label>
                  <select value={taskAssignee} onChange={(e) => setTaskAssignee(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm">
                    {ALL_STAFF.filter(s => s.id !== "nishant").map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                </div>
                <div>
                  <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Priority</label>
                  <select value={taskPriority} onChange={(e) => setTaskPriority(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm">
                    <option value="low">Low</option>
                    <option value="medium">Medium</option>
                    <option value="high">High</option>
                    <option value="critical">Critical</option>
                  </select>
                </div>
              </div>
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Outlet (Optional)</label>
                <select value={taskOutlet} onChange={(e) => setTaskOutlet(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm">
                  <option value="">No specific outlet</option>
                  {OUTLETS.map(o => <option key={o} value={o}>{OUTLET_NAMES[o] || o.replace(/_/g, " ")}</option>)}
                </select>
              </div>
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Due In</label>
                <select value={taskDueHours} onChange={(e) => setTaskDueHours(e.target.value)} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm">
                  <option value="1">1 hour</option>
                  <option value="2">2 hours</option>
                  <option value="4">4 hours</option>
                  <option value="8">8 hours</option>
                  <option value="24">1 day</option>
                </select>
              </div>
            </div>
            <div className="flex gap-3 justify-end mt-6 pt-5 border-t border-zinc-800">
              <button onClick={() => setShowModal(false)} className="px-5 py-2.5 border border-zinc-700 text-sm font-medium hover:border-zinc-500 transition-colors">Cancel</button>
              <button onClick={assignTask} disabled={submitting || !taskTitle.trim()} className="px-5 py-2.5 bg-yellow-400 text-black font-bold text-sm uppercase tracking-widest hover:opacity-90 transition-opacity disabled:opacity-50">
                {submitting ? "Assigning..." : "Assign Task"}
              </button>
            </div>
          </div>
        </div>
      )}

      {showPinModal && (
        <div className="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="w-full max-w-[400px] bg-[#131316] border border-zinc-800 p-8">
            <div className="flex justify-between items-center mb-6 pb-4 border-b border-zinc-800">
              <h3 className="text-xl font-black tracking-tight">Change PIN</h3>
              <button onClick={() => { setShowPinModal(false); setPinMsg(""); setNewPin(""); }} className="text-zinc-500 hover:text-white text-xl">✕</button>
            </div>
            <div className="space-y-4">
              <div>
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">New PIN (min 4 digits)</label>
                <input type="password" value={newPin} onChange={(e) => setNewPin(e.target.value)} placeholder="Enter new PIN" className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-yellow-400 transition-colors text-sm" />
              </div>
              {pinMsg && <p className={`font-mono text-xs uppercase ${pinMsg.includes("success") ? "text-green-400" : "text-red-500"}`}>{pinMsg}</p>}
              <button onClick={updatePin} className="w-full bg-yellow-400 text-black font-bold tracking-widest text-sm py-3 hover:opacity-90 transition-opacity uppercase">Update PIN</button>
            </div>
          </div>
        </div>
      )}

      {overdueTask && user?.role !== "Owner" && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4" style={{ background: "rgba(20,0,0,0.97)" }}>
          <div className="w-full max-w-[560px] bg-[#131316] border-2 border-red-500">
            <div className="bg-red-500 px-6 py-4 text-center font-mono text-xs font-bold uppercase tracking-widest text-white animate-pulse">
              ⚠ Action Required · Overdue Task ⚠
            </div>
            <div className="p-6">
              <h3 className="text-2xl font-black mb-2">You have an overdue task</h3>
              <p className="text-zinc-400 text-sm mb-4 leading-relaxed">This task is past its deadline. You must either complete it or provide a reason.</p>
              <div className="bg-black border border-zinc-800 p-4 mb-4">
                <p className="font-bold text-base mb-1">{overdueTask.title}</p>
                <p className="font-mono text-xs text-red-500 uppercase tracking-widest">
                  Overdue by {Math.round((Date.now() - new Date(overdueTask.due_at).getTime()) / 60000)} min · {overdueTask.priority}
                </p>
              </div>
              <div className="mb-4">
                <label className="block text-[10px] font-mono text-zinc-500 uppercase tracking-widest mb-2">Reason for delay (min 20 characters)</label>
                <textarea value={forceAckReason} onChange={(e) => setForceAckReason(e.target.value)} placeholder="Explain why this task is overdue..." rows={3} className="w-full bg-black border border-zinc-800 text-white px-4 py-3 focus:outline-none focus:border-red-500 transition-colors text-sm resize-none" />
              </div>
              <div className="flex gap-3 justify-end pt-4 border-t border-zinc-800">
                <button onClick={() => submitForceAck("reason")} className="px-5 py-2.5 border border-zinc-700 text-sm font-medium hover:border-zinc-500 transition-colors">Submit Reason</button>
                <button onClick={() => submitForceAck("complete")} className="px-5 py-2.5 bg-yellow-400 text-black font-bold text-sm uppercase tracking-widest hover:opacity-90 transition-opacity">Mark Complete</button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
