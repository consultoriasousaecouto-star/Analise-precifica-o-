import { createClient } from '@supabase/supabase-js';

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const supabaseConfigurado = !!(url && anonKey);

export const supabase = supabaseConfigurado ? createClient(url, anonKey) : null;
