import { ownerOn, istToday } from "./outletOwner";
// Matches the outlet name that appears in a review email (e.g. "Brownie Heaven Royapettah",
// "Brownie Heaven, Pallavaram") to your internal outlet_id, and finds who owns that outlet.
// Kept separate from outletMap.ts, which is for the Stock/margin integration, not reviews.

export const REVIEW_OUTLET_DISPLAY_NAMES: Record<string, string> = {
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

// Same ownership as ALL_STAFF in dashboard/page.tsx — kept in sync manually since
// this file is used server-side (API route) and shouldn't import the client dashboard.
const STAFF_BY_OUTLET: Record<string, string> = {
  velachery: "vishnu", perumbakkam: "vishnu", tambaram: "vishnu", porur: "vishnu",
  anna_nagar: "vishnu", vadapalani: "vishnu",
  royapettah: "stephen", adayar: "ahila", bsr_mall: "ahila", pallavaram: "ahila", ra_puram: "ahila",
  besant_nagar: "ahila",
};

// Finds the outlet_id whose display name appears inside the raw outlet name text
// from the email. Case-insensitive, ignores punctuation like the comma Zomato adds.
//
// Some platforms use the real street/area name for a listing instead of
// TASKFORCE's own outlet name — confirmed via the Swiggy ID sheets ("Thoraipakkam"
// = BSR Mall, "Adambakkam" = Velachery) and now seen the same way in a real
// Zomato review email too ("Brownie Heaven, Adambakkam").
const OUTLET_NAME_ALIASES: Record<string, string> = {
  thoraipakkam: "bsr_mall",
  adambakkam: "velachery",
};

export function matchReviewOutletId(rawOutletName: string): string | null {
  const cleaned = rawOutletName.toLowerCase().replace(/[,.]/g, " ").replace(/\s+/g, " ").trim();

  for (const [alias, outletId] of Object.entries(OUTLET_NAME_ALIASES)) {
    if (cleaned.includes(alias)) return outletId;
  }

  for (const [outletId, displayName] of Object.entries(REVIEW_OUTLET_DISPLAY_NAMES)) {
    if (cleaned.includes(displayName.toLowerCase())) {
      return outletId;
    }
  }

  // Some review emails just say "Brownie Heaven" with no location at all — this
  // appears to be Royapettah's default Google listing name (it's the flagship
  // outlet). Only apply this fallback for the bare name, not a partial/garbled one.
  if (cleaned === "brownie heaven") return "royapettah";

  return null;
}

export function staffIdForReviewOutlet(outletId: string): string {
  return ownerOn(outletId, istToday(), STAFF_BY_OUTLET[outletId]) || "niranjana"; // fallback: Founder's Office
}
