// ONE shared month-end sales projection, used by BOTH the Command Centre and
// the Sales Forecast tab, so the two always show the same number.
//
// Method (per outlet, per channel):
//   1. Sales so far = real reports filed from the 2nd of the month up to the
//      "as of" date. Reports are filed the day AFTER the sales day
//      (report_date = sales day + 1), so the report dated the 1st belongs to
//      last month and is NOT counted.
//   2. Each remaining sales day is estimated from the average of the same
//      weekday over the last ~10 weeks (a Sunday uses past Sundays).
//   3. That is multiplied by a short-term trend factor: last 14 reports vs the
//      14 before, limited to between 0.7 and 1.3.
//   4. Projected month-end = sales so far + all remaining days, all outlets.
// The "as of" report date itself is left out of the weekday averages (it may
// still be only partly filed) but IS counted in sales so far.

export type ProjRow = {
  outlet_id: string;
  report_date: string;
  shop_sales_value?: any;
  swiggy_sales_value?: any;
  zomato_sales_value?: any;
};

export type OutletMonthEnd = {
  oid: string;
  mtdShop: number;
  mtdOnline: number;
  remShop: number;
  remOnline: number;
};

export type MonthEndProjection = {
  perOutlet: OutletMonthEnd[];
  mtd: number;
  remaining: number;
  projected: number;
  daysRemaining: number;
};

export const PROJECTION_HISTORY_DAYS = 70;

const DAY = 86400000;
const toMs = (s: string) => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10));
const toStr = (ms: number) => new Date(ms).toISOString().slice(0, 10);

export const projectionHistoryStart = (asOf: string) => toStr(toMs(asOf) - PROJECTION_HISTORY_DAYS * DAY);

// Fetches the rows the projection needs (history window up to asOf), paging
// past the 1000-row limit.
export async function fetchProjectionRows(supabase: any, asOf: string): Promise<ProjRow[]> {
  const from = projectionHistoryStart(asOf);
  const BATCH = 1000;
  let all: ProjRow[] = [];
  let offset = 0;
  while (true) {
    const { data, error } = await supabase
      .from("outlet_reports")
      .select("outlet_id, report_date, shop_sales_value, swiggy_sales_value, zomato_sales_value")
      .gte("report_date", from)
      .lte("report_date", asOf)
      .order("report_date", { ascending: true })
      .range(offset, offset + BATCH - 1);
    if (error) throw error;
    all = all.concat((data || []) as ProjRow[]);
    if (!data || data.length < BATCH) break;
    offset += BATCH;
  }
  return all;
}

export function computeMonthEnd(rows: ProjRow[], outlets: string[], asOf: string): MonthEndProjection {
  const n = (v: any) => Number(v) || 0;
  const asOfMs = toMs(asOf);
  const d = new Date(asOfMs);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth();
  const monthFiledFrom = `${y}-${String(m + 1).padStart(2, "0")}-02`;
  const nextMonthFirstMs = Date.UTC(y, m + 1, 1);

  // Remaining report dates: the day after asOf up to the 1st of next month
  // (the report for the month's last sales day is filed on the 1st).
  const remainingDates: string[] = [];
  for (let t = asOfMs + DAY; t <= nextMonthFirstMs; t += DAY) remainingDates.push(toStr(t));

  const avg = (arr: number[]) => (arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0);
  const perOutletOut: OutletMonthEnd[] = [];

  outlets.forEach((o) => {
    const mine = rows.filter((r) => r.outlet_id === o);
    // weekday buckets from history, leaving out the as-of date
    const shop: number[][] = [[], [], [], [], [], [], []];
    const swiggy: number[][] = [[], [], [], [], [], [], []];
    const zomato: number[][] = [[], [], [], [], [], [], []];
    const hist = mine.filter((r) => r.report_date !== asOf).sort((a, b) => a.report_date.localeCompare(b.report_date));
    hist.forEach((r) => {
      const dow = new Date(r.report_date + "T00:00:00Z").getUTCDay();
      shop[dow].push(n(r.shop_sales_value));
      swiggy[dow].push(n(r.swiggy_sales_value));
      zomato[dow].push(n(r.zomato_sales_value));
    });
    const totalOf = (rs: ProjRow[]) => rs.reduce((s, r) => s + n(r.shop_sales_value) + n(r.swiggy_sales_value) + n(r.zomato_sales_value), 0);
    const last14 = totalOf(hist.slice(-14));
    const prev14 = totalOf(hist.slice(-28, -14));
    let trend = prev14 > 0 ? last14 / prev14 : 1;
    trend = Math.max(0.7, Math.min(1.3, trend));

    // sales so far: reports filed from the 2nd of the month up to asOf
    const mtdRows = mine.filter((r) => r.report_date >= monthFiledFrom && r.report_date <= asOf);
    const mtdShop = mtdRows.reduce((s, r) => s + n(r.shop_sales_value), 0);
    const mtdOnline = mtdRows.reduce((s, r) => s + n(r.swiggy_sales_value) + n(r.zomato_sales_value), 0);

    let remShop = 0;
    let remOnline = 0;
    remainingDates.forEach((ds) => {
      const dow = new Date(ds + "T00:00:00Z").getUTCDay();
      remShop += avg(shop[dow]) * trend;
      remOnline += (avg(swiggy[dow]) + avg(zomato[dow])) * trend;
    });

    perOutletOut.push({ oid: o, mtdShop, mtdOnline, remShop, remOnline });
  });

  const mtd = perOutletOut.reduce((s, r) => s + r.mtdShop + r.mtdOnline, 0);
  const remaining = perOutletOut.reduce((s, r) => s + r.remShop + r.remOnline, 0);
  return { perOutlet: perOutletOut, mtd, remaining, projected: mtd + remaining, daysRemaining: remainingDates.length };
}
