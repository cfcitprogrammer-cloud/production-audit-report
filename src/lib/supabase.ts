// src/lib/supabase.ts
import { createClient } from "@supabase/supabase-js"

const supabaseUrl = "https://fvurfovcggcamkushrnz.supabase.co"
const supabaseAnonKey = "sb_publishable_I0k8sa855o-Aud0hL3Gg1w_2i8Y4RGy"

if (!supabaseUrl || !supabaseAnonKey) {
  throw new Error("Missing Supabase environment variables")
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey)
