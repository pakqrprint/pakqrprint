// Pak QR Print - Supabase config (file storage)
// ============================================================
// Supabase FREE plan (muft, koi card nahi): 1GB file storage.
// Files customer yahan upload karta hai; dukaan ka PC (print agent)
// public URL se download karke print karta hai.
//
// NOTE (zaroori): ye anon key purani (legacy) format ki hai — Supabase
// is ko 2027 ke shuru mein band kar dega. Us se pehle Supabase
// dashboard > Settings > API se nayi "publishable key" le kar yahan
// badal dein (aur GitHub par dobara upload kar dein). Tab tak ye
// key theek chalti hai.
//
// Ye key PUBLIC hoti hai (site ki JS mein hoti hai) — is mein koi
// password nahi hota. Bucket "pakqrprint-files" ki policy sirf
// upload (INSERT) aur parhna (SELECT) deti hai; delete/update nahi.
// Is liye purani files bucket se khud delete nahi hoti — Supabase
// dashboard se waqt par saaf karni hongi (Hidayat dekhein).
// ============================================================
window.PAKQR_SUPABASE_CONFIG = {
  url: "https://jdenvuhnvnpgydkpoyie.supabase.co",
  anonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkZW52dWhudm5wZ3lka3BveWllIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTM5NTM4NDQsImV4cCI6MjEwOTUyOTg0NH0.MIBOn2E1fdCogbW9gArKKHASdTTKJG6I2KqbNaBBkpk",
  bucket: "pakqrprint-files"
};
