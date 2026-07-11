/* ============================================================
   ShipX Billing — configuration
   Fill these in AFTER you create your Supabase project:
   Supabase dashboard → Project Settings → API
     • Project URL   → SUPABASE_URL
     • anon public key (Project API keys) → SUPABASE_ANON_KEY
   The anon key is safe to expose in the browser because Row-Level
   Security (see supabase-schema.sql) blocks anyone who isn't logged in.
   ============================================================ */
window.SHIPX_CONFIG = {
  SUPABASE_URL:      "PASTE_YOUR_SUPABASE_PROJECT_URL_HERE",
  SUPABASE_ANON_KEY: "PASTE_YOUR_SUPABASE_ANON_KEY_HERE"
};
