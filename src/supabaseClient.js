import { createClient } from "@supabase/supabase-js";

// Bu kalit "publishable" (ochiq) kalit — brauzerda ko'rinishi xavfsiz,
// chunki ma'lumotlar bazasi RLS (Row Level Security) orqali himoyalangan.
const SUPABASE_URL = "https://akmdcdiovharnzxvpthh.supabase.co";
const SUPABASE_KEY = "sb_publishable_KxeHqepLclrwkVMdnq5Wbg_scs7C9Z0";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
