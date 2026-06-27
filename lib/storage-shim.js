import { supabase } from "@/integrations/supabase/client";

export function installStorageShim() {
  if (typeof window === "undefined") return;
  if (window.storage && window.storage.__mkisSupabase) return;

  window.storage = {
    __mkisSupabase: true,
    async get(key) {
      try {
        const { data, error } = await supabase.from("kv_store").select("value").eq("key", key).maybeSingle();
        if (error || !data) return null;
        return { value: data.value };
      } catch { return null; }
    },
    async set(key, value) {
      try {
        const { error } = await supabase.from("kv_store").upsert({ key, value: String(value), updated_at: new Date().toISOString() });
        return !error;
      } catch { return false; }
    },
    async remove(key) {
      try { await supabase.from("kv_store").delete().eq("key", key); return true; }
      catch { return false; }
    },
  };
}
