// Anna Nagar handover: Vishnu looked after it up to and including 2026-10-27.
// From 2026-10-28 onwards Stephen does. Dates BEFORE the handover date keep
// pointing at Vishnu, so old reports, points and history are never changed.
export const ANNA_NAGAR_HANDOVER_FROM = "2026-10-28";

// Today's date in India (IST), as YYYY-MM-DD.
export const istToday = (): string =>
  new Date(Date.now() + 5.5 * 60 * 60 * 1000).toISOString().split("T")[0];

// Staff id responsible for Anna Nagar on a given date.
export const annaNagarOwnerOn = (date: string): "vishnu" | "stephen" =>
  date >= ANNA_NAGAR_HANDOVER_FROM ? "stephen" : "vishnu";

// Who owns an outlet on a date. Only Anna Nagar changes; everything else
// returns `base` (the existing owner) untouched.
export const ownerOn = (outletId: string, date: string | null | undefined, base: string | undefined): string | undefined =>
  outletId === "anna_nagar" ? annaNagarOwnerOn(date || istToday()) : base;
