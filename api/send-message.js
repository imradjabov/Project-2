import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://akmdcdiovharnzxvpthh.supabase.co";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { shopId, debtorId, text } = req.body || {};
    if (!shopId || !debtorId || !text || !text.trim()) {
      return res.status(400).json({ error: "Ma'lumot yetarli emas" });
    }
    if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ error: "Server sozlanmagan" });
    }

    const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

    const { data: shop } = await supabase.from("shops").select("bot_token").eq("id", shopId).single();
    if (!shop || !shop.bot_token) return res.status(400).json({ error: "Bot ulanmagan" });

    const { data: debtor } = await supabase
      .from("debtors").select("telegram_chat_id, name")
      .eq("id", debtorId).eq("shop_id", shopId).single();
    if (!debtor || !debtor.telegram_chat_id) {
      return res.status(400).json({ error: "Bu mijoz botga hali ulanmagan" });
    }

    const tgRes = await fetch(`https://api.telegram.org/bot${shop.bot_token}/sendMessage`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ chat_id: debtor.telegram_chat_id, text: text.trim() }),
    });
    const tgData = await tgRes.json();
    if (!tgData.ok) return res.status(400).json({ error: tgData.description || "Telegram xatoligi" });

    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(500).json({ error: String(e && e.message ? e.message : e) });
  }
}
