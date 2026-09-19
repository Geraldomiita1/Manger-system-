// Provides window.storage.{get,set} backed by localStorage so the MKIS
// component runs identically outside the Lovable host environment.
// Build: 2026.09.19.1302
export function installStorageShim() {
  if (typeof window === "undefined") return;
  if (window.storage && window.storage.__mkisShim) return;
  const prefix = "mkis_shared::";
  window.storage = {
    __mkisShim: true,
    async get(key) {
      try {
        const value = window.localStorage.getItem(prefix + key);
        return value == null ? null : { value };
      } catch { return null; }
    },
    async set(key, value) {
      // Deliberately no try/catch here -- if localStorage.setItem throws (most
      // commonly QuotaExceededError once accumulated marks/attendance/reports
      // data grows large enough), that failure MUST propagate as a rejected
      // promise. saveShared()/queueKeySave() rely on that to know a write
      // didn't actually land, so they can retry it instead of wrongly marking
      // it saved. Swallowing the error here was exactly the bug: it told the
      // app every save succeeded even when the browser had just silently
      // refused to store it, so that data quietly vanished on the next reload.
      window.localStorage.setItem(prefix + key, String(value));
      return true;
    },
    async remove(key) {
      try { window.localStorage.removeItem(prefix + key); } catch {}
      return true;
    },
  };
}
