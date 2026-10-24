import { createClient } from "@supabase/supabase-js"

type AuthStorageName = "localStorage" | "sessionStorage"

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY

export const isSupabaseConfigured = Boolean(supabaseUrl && supabaseAnonKey)

let preferredStorage: AuthStorageName = "sessionStorage"
let storageWasBlocked = false
const memoryStorage = new Map<string, string>()

const authStorage = {
  getItem(key: string) {
    const storageNames: AuthStorageName[] =
      preferredStorage === "localStorage"
        ? ["localStorage", "sessionStorage"]
        : ["sessionStorage", "localStorage"]

    for (const storageName of storageNames) {
      try {
        const value = window[storageName].getItem(key)
        if (value !== null) return value
      } catch (error) {
        storageWasBlocked = true
        console.warn(`Unable to read Supabase session from ${storageName}.`, error)
      }
    }

    return memoryStorage.get(key) ?? null
  },
  setItem(key: string, value: string) {
    try {
      window[preferredStorage].setItem(key, value)
      memoryStorage.delete(key)
    } catch (error) {
      storageWasBlocked = true
      console.warn(
        `Unable to persist the Supabase session in ${preferredStorage}.`,
        error,
      )
      memoryStorage.set(key, value)
    }
  },
  removeItem(key: string) {
    for (const storageName of ["localStorage", "sessionStorage"] as const) {
      try {
        window[storageName].removeItem(key)
      } catch (error) {
        storageWasBlocked = true
        console.warn(`Unable to clear Supabase session from ${storageName}.`, error)
      }
    }
    memoryStorage.delete(key)
  },
}

export const supabase = isSupabaseConfigured
  ? createClient(supabaseUrl, supabaseAnonKey, {
      auth: {
        storage: authStorage,
        persistSession: true,
        autoRefreshToken: true,
        detectSessionInUrl: true,
        flowType: "pkce",
      },
    })
  : null

export function setAuthRememberPreference(remember: boolean) {
  preferredStorage = remember ? "localStorage" : "sessionStorage"
  storageWasBlocked = false
}

export function consumeAuthStorageWarning() {
  const wasBlocked = storageWasBlocked
  storageWasBlocked = false
  return wasBlocked
}
