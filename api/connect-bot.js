import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://akmdcdiovharnzxvpthh.supabase.co";

export default async function handler(req, res) {
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { shopId } = req.body || {};
    if (!shopId) return res.status(400).json({ error: "shopId required" });

    if (!process.env.SUPABASE_SERVICE_ROLE_KEY) {
      return res.status(500).json({ error: "Server sozlanmagan: SUPABASE_SERVICE_ROLE_KEY yo'q" });
    }

    const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);
    const { data: shop, error } = await supabase.from("shops").select("bot_token").eq("id", shopId).single();
    if (error || !shop || !shop.bot_token) {
      return res.status(400).json({ error: "Bot tokeni topilmadi" });
    }

    const host = req.headers["x-forwarded-host"] || req.headers.host;
    const webhookUrl = `https://${host}/api/telegram/${shop.bot_token}`;

    const tgRes = await fetch(
      `https://api.telegram.org/bot${shop.bot_token}/setWebhook?url=${encodeURIComponent(webhookUrl)}`
    );
    const tgData = await tgRes.json();

    if (!tgData.ok) {
      return res.status(400).json({ error: tgData.description || "Telegram xatoligi" });
    }

    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(500).json({ error: String(e && e.message ? e.message : e) });
  }
}
