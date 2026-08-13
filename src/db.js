import { supabase } from "./supabaseClient";

// ---------- row <-> camelCase mappers ----------
const mapShop = (r) => ({
  id: r.id, name: r.name, productType: r.product_type || "", phone: r.phone || "",
  address: r.address || "", login: r.login, pass: r.pass,
});
const mapSeller = (r) => ({ id: r.id, name: r.name });
const mapDebtor = (r) => ({
  id: r.id, name: r.name, phone: r.phone || "", amount: Number(r.amount),
  dueDate: r.due_date, colorOverride: r.color_override, sellerId: r.seller_id,
  createdAt: r.created_at,
});
const mapHistory = (r) => ({
  id: r.id, kind: r.kind, debtorName: r.debtor_name, amount: Number(r.amount),
  sellerName: r.seller_name || "", note: r.note || "",
  date: fmtDateFromISO(r.created_at), time: fmtTimeFromISO(r.created_at),
});
const pad2 = (n) => String(n).padStart(2, "0");
function fmtDateFromISO(iso) { const d = new Date(iso); return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`; }
function fmtTimeFromISO(iso) { const d = new Date(iso); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; }

// ---------- LOAD EVERYTHING ----------
export async function loadAll() {
  const [{ data: adminRows, error: e1 }, { data: shopRows, error: e2 }, { data: sellerRows, error: e3 }, { data: debtorRows, error: e4 }, { data: historyRows, error: e5 }] =
    await Promise.all([
      supabase.from("admin_account").select("*").limit(1),
      supabase.from("shops").select("*").order("created_at", { ascending: false }),
      supabase.from("sellers").select("*"),
      supabase.from("debtors").select("*"),
      supabase.from("history_log").select("*").order("created_at", { ascending: false }),
    ]);
  if (e1 || e2 || e3 || e4 || e5) throw (e1 || e2 || e3 || e4 || e5);

  const admin = adminRows[0] ? { id: adminRows[0].id, login: adminRows[0].login, pass: adminRows[0].pass } : null;

  const shops = (shopRows || []).map((s) => {
    const sellers = (sellerRows || []).filter((x) => x.shop_id === s.id).map(mapSeller);
    const debtors = (debtorRows || []).filter((x) => x.shop_id === s.id).map(mapDebtor);
    const historyLog = (historyRows || []).filter((x) => x.shop_id === s.id).map(mapHistory);
    return { ...mapShop(s), sellers, debtors, historyLog };
  });

  return { admin, shops };
}

// ---------- ADMIN ----------
export async function updateAdminCredentials(adminId, login, pass) {
  const { error } = await supabase.from("admin_account").update({ login, pass }).eq("id", adminId);
  if (error) throw error;
}

// ---------- SHOPS (admin panel) ----------
export async function createShop({ name, productType, phone, address, login, pass }) {
  const { data, error } = await supabase.from("shops")
    .insert({ name, product_type: productType, phone, address, login, pass })
    .select().single();
  if (error) throw error;
  return { ...mapShop(data), sellers: [], debtors: [], historyLog: [] };
}

export async function updateShopInfo(shopId, { name, productType, phone, address, login, pass }) {
  const patch = { name, product_type: productType, phone, address, login };
  if (pass) patch.pass = pass;
  const { error } = await supabase.from("shops").update(patch).eq("id", shopId);
  if (error) throw error;
}

export async function deleteShopRow(shopId) {
  const { error } = await supabase.from("shops").delete().eq("id", shopId);
  if (error) throw error;
}

export async function isLoginTaken(login, excludeShopId) {
  let q = supabase.from("shops").select("id").eq("login", login);
  if (excludeShopId) q = q.neq("id", excludeShopId);
  const { data, error } = await q;
  if (error) throw error;
  return (data || []).length > 0;
}

// ---------- LOGIN ----------
export async function tryLoginQuery(login, pass) {
  const { data: adminMatch } = await supabase.from("admin_account").select("*").eq("login", login).eq("pass", pass).maybeSingle();
  if (adminMatch) return { role: "admin", admin: { id: adminMatch.id, login: adminMatch.login, pass: adminMatch.pass } };

  const { data: shopMatch } = await supabase.from("shops").select("*").eq("login", login).eq("pass", pass).maybeSingle();
  if (shopMatch) {
    const [{ data: sellerRows }, { data: debtorRows }, { data: historyRows }] = await Promise.all([
      supabase.from("sellers").select("*").eq("shop_id", shopMatch.id),
      supabase.from("debtors").select("*").eq("shop_id", shopMatch.id),
      supabase.from("history_log").select("*").eq("shop_id", shopMatch.id).order("created_at", { ascending: false }),
    ]);
    return {
      role: "owner",
      shop: {
        ...mapShop(shopMatch),
        sellers: (sellerRows || []).map(mapSeller),
        debtors: (debtorRows || []).map(mapDebtor),
        historyLog: (historyRows || []).map(mapHistory),
      },
    };
  }
  return null;
}

// ---------- SELLERS ----------
export async function addSellerRow(shopId, name) {
  const { data, error } = await supabase.from("sellers").insert({ shop_id: shopId, name }).select().single();
  if (error) throw error;
  return mapSeller(data);
}
export async function removeSellerRow(sellerId) {
  const { error } = await supabase.from("sellers").delete().eq("id", sellerId);
  if (error) throw error;
}

// ---------- DEBTORS ----------
export async function createDebtorRow(shopId, { name, phone, amount, dueDate, sellerId }) {
  const { data, error } = await supabase.from("debtors")
    .insert({ shop_id: shopId, name, phone, amount, due_date: dueDate, seller_id: sellerId || null })
    .select().single();
  if (error) throw error;
  return mapDebtor(data);
}

export async function addAmountToDebtor(debtorId, newAmount, dueDate) {
  const patch = { amount: newAmount };
  if (dueDate) patch.due_date = dueDate;
  const { data, error } = await supabase.from("debtors").update(patch).eq("id", debtorId).select().single();
  if (error) throw error;
  return mapDebtor(data);
}

export async function updateDebtorProfileRow(debtorId, { name, phone, amount, dueDate, colorOverride }) {
  const { data, error } = await supabase.from("debtors")
    .update({ name, phone, amount, due_date: dueDate, color_override: colorOverride })
    .eq("id", debtorId).select().single();
  if (error) throw error;
  return mapDebtor(data);
}

export async function payDebtorRow(debtorId, remainingAmount) {
  if (remainingAmount <= 0) {
    const { error } = await supabase.from("debtors").delete().eq("id", debtorId);
    if (error) throw error;
    return null; // fully paid, removed
  }
  const { data, error } = await supabase.from("debtors").update({ amount: remainingAmount }).eq("id", debtorId).select().single();
  if (error) throw error;
  return mapDebtor(data);
}

// ---------- HISTORY ----------
export async function addHistoryRow(shopId, { kind, debtorName, amount, sellerName, note, createdAt }) {
  const payload = { shop_id: shopId, kind, debtor_name: debtorName, amount, seller_name: sellerName || "", note: note || "" };
  if (createdAt) payload.created_at = createdAt;
  const { data, error } = await supabase.from("history_log").insert(payload).select().single();
  if (error) throw error;
  return mapHistory(data);
}

// ---------- BULK IMPORT (from Excel) ----------
export async function bulkImportDebtors(shopId, existingSellers, rows) {
  // rows: [{ name, amount, createdAt(ISO), sellerName }]
  const existingByName = new Map(existingSellers.map((s) => [s.name.trim().toLowerCase(), s.id]));
  const neededNames = [...new Set(rows.map((r) => (r.sellerName || "").trim()).filter((n) => n && !existingByName.has(n.toLowerCase())))];

  let createdSellers = [];
  if (neededNames.length > 0) {
    const { data: newSellers, error: sErr } = await supabase
      .from("sellers").insert(neededNames.map((name) => ({ shop_id: shopId, name }))).select();
    if (sErr) throw sErr;
    createdSellers = newSellers || [];
    createdSellers.forEach((s) => existingByName.set(s.name.trim().toLowerCase(), s.id));
  }

  const debtorInsertRows = rows.map((r) => ({
    shop_id: shopId, name: r.name, phone: "", amount: r.amount, due_date: null,
    seller_id: r.sellerName ? existingByName.get(r.sellerName.trim().toLowerCase()) || null : null,
    created_at: r.createdAt,
  }));
  const { error: dErr } = await supabase.from("debtors").insert(debtorInsertRows);
  if (dErr) throw dErr;

  const historyInsertRows = rows.map((r) => ({
    shop_id: shopId, kind: "added", debtor_name: r.name, amount: r.amount,
    seller_name: r.sellerName || "", note: "Excel orqali import qilindi", created_at: r.createdAt,
  }));
  const { error: hErr } = await supabase.from("history_log").insert(historyInsertRows);
  if (hErr) throw hErr;

  return { imported: rows.length, newSellers: createdSellers.length };
}
