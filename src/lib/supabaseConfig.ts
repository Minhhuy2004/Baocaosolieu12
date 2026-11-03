const defaultSupabaseUrl = "https://kybdhsflbcwqzepdcheo.supabase.co"
const defaultSupabasePublishableKey =
  "sb_publishable_DdzrSJw-_S1DunvImz0nQg_kvG57hDi"

export const supabaseUrl =
  import.meta.env.VITE_SUPABASE_URL?.trim() || defaultSupabaseUrl

export const supabaseAnonKey =
  import.meta.env.VITE_SUPABASE_ANON_KEY?.trim() ||
  defaultSupabasePublishableKey

export const isSupabaseConfigured = Boolean(supabaseUrl && supabaseAnonKey)
