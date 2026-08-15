import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://akmdcdiovharnzxvpthh.supabase.co";

function fmtMoney(n) {
  return Math.round(n || 0).toLocaleString("uz-UZ").replace(/,/g, " ") + " so'm";
}
function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const p = (x) => String(x).padStart(2, "0");
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`;
}

async function tg(token, method, payload) {
  await fetch(`https://api.telegram.org/bot${token}/${method}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
}

export default async function handler(req, res) {
  // Always answer 200 quickly so Telegram doesn't retry-storm us.
  if (req.method !== "POST") return res.status(200).json({ ok: true });

  const { token } = req.query;

  try {
    const supabase = createClient(SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);
    const { data: shop } = await supabase.from("shops").select("*").eq("bot_token", token).maybeSingle();
    if (!shop) return res.status(200).json({ ok: true });

    const update = req.body || {};

    if (update.message) {
      const chatId = update.message.chat.id;
      const text = (update.message.text || "").trim();

      if (text === "/start") {
        await tg(token, "sendMessage", {
          chat_id: chatId,
          text: `Assalomu alaykum! "${shop.name}" do'koniga xush kelibsiz.\n\nIltimos, sizga berilgan ID raqamingizni yuboring.`,
        });
      } else if (/^\d+$/.test(text)) {
        const { data: debtor } = await supabase
          .from("debtors").select("*")
          .eq("shop_id", shop.id).eq("debtor_number", parseInt(text, 10)).maybeSingle();

        if (debtor) {
          await supabase.from("debtors").update({ telegram_chat_id: chatId }).eq("id", debtor.id);
          await tg(token, "sendMessage", {
            chat_id: chatId,
            text: `Xush kelibsiz, ${debtor.name}!`,
            reply_markup: { inline_keyboard: [[{ text: "💰 Qarzlar", callback_data: "show_debts" }]] },
          });
        } else {
          await tg(token, "sendMessage", {
            chat_id: chatId,
            text: "Bunday ID raqami topilmadi. Iltimos, do'kon egasidan to'g'ri ID raqamini so'rang.",
          });
        }
      } else {
        await tg(token, "sendMessage", { chat_id: chatId, text: "Iltimos /start bosing yoki ID raqamingizni yuboring." });
      }
    }

    if (update.callback_query) {
      const chatId = update.callback_query.message.chat.id;
      const data = update.callback_query.data;
      await tg(token, "answerCallbackQuery", { callback_query_id: update.callback_query.id });

      if (data === "show_debts") {
        const { data: debtor } = await supabase
          .from("debtors").select("*")
          .eq("shop_id", shop.id).eq("telegram_chat_id", chatId).maybeSingle();

        if (debtor) {
          const dueText = debtor.due_date ? `\nMuddat: ${fmtDate(debtor.due_date)}` : "";
          await tg(token, "sendMessage", {
            chat_id: chatId,
            text: `📋 ${debtor.name}\nQarz summasi: ${fmtMoney(debtor.amount)}${dueText}`,
          });
        } else {
          await tg(token, "sendMessage", {
            chat_id: chatId,
            text: "Sizning ID raqamingiz aniqlanmadi. Iltimos /start bosib qayta urinib ko'ring.",
          });
        }
      }
    }

    return res.status(200).json({ ok: true });
  } catch (e) {
    return res.status(200).json({ ok: true });
  }
}
