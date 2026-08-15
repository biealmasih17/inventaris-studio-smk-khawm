// supabase-config.example.js
// Copy this file to 'supabase-config.js' and fill the values with your real Supabase project URL and ANON key.
// IMPORTANT: Do NOT commit supabase-config.js with real keys into source control. Add it to .gitignore.

// Example (replace with your real values locally or via CI):
const SUPABASE_URL = 'https://your-project.supabase.co';
const SUPABASE_ANON_KEY = 'your-anon-key-here';

// Initialize Supabase client (available as global `supabaseClient`):
if (typeof supabase !== 'undefined' && typeof supabase.createClient === 'function') {
  window.supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
} else {
  // supabase library not loaded yet; create initializer for Code.js to call later
  window.__SUPABASE_CONFIG = { url: SUPABASE_URL, key: SUPABASE_ANON_KEY };
}
