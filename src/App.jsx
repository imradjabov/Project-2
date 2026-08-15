import React, { useState, useEffect, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import {
  loadAll, createShop, updateShopInfo, deleteShopRow, updateAdminCredentials,
  addSellerRow, removeSellerRow, createDebtorRow, addAmountToDebtor,
  updateDebtorProfileRow, payDebtorRow, addHistoryRow, bulkImportDebtors,
  updateDebtorNumberRow, updateShopBotToken, assignMissingIds,
} from "./db";

const pad2 = (n) => String(n).padStart(2, "0");
const fmt = (n) => Math.round(n || 0).toLocaleString("uz-UZ").replace(/,/g, " ") + " so'm";
const fmtDate = (iso) => { if (!iso) return ""; const d = new Date(iso); return `${pad2(d.getDate())}.${pad2(d.getMonth() + 1)}.${d.getFullYear()}`; };
const fmtTime = (iso) => { if (!iso) return ""; const d = new Date(iso); return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`; };
const daysUntil = (iso) => { if (!iso) return 999; const due = new Date(iso); const now = new Date(); due.setHours(0, 0, 0, 0); now.setHours(0, 0, 0, 0); return Math.round((due - now) / 86400000); };
const autoColor = (dueIso) => { const d = daysUntil(dueIso); if (d > 7) return "green"; if (d >= 1) return "yellow"; if (d >= -4) return "red"; return "black"; };
const ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

function exportDebtorsXlsx(shop, periodKey) {
  const now = new Date();
  const list = shop.debtors.filter((d) => {
    const created = new Date(d.createdAt);
    if (periodKey === "daily") return created.toDateString() === now.toDateString();
    if (periodKey === "monthly") return created.getFullYear() === now.getFullYear() && created.getMonth() === now.getMonth();
    if (periodKey === "yearly") return created.getFullYear() === now.getFullYear();
    return true; // "all"
  });
  const header = ["ISM FAMILYA", "QARZ MIQDORI ", "QARZ OLINGAN kun/oy/yil", "QARZ BERGAN SOTUVCHI"];
  const rows = list.map((d) => ({
    "ISM FAMILYA": d.name,
    "QARZ MIQDORI ": d.amount,
    "QARZ OLINGAN kun/oy/yil": fmtDate(d.createdAt),
    "QARZ BERGAN SOTUVCHI": shop.sellers.find((s) => s.id === d.sellerId)?.name || "",
  }));
  const ws = XLSX.utils.json_to_sheet(rows, { header });
  ws["!cols"] = [{ wch: 24 }, { wch: 16 }, { wch: 22 }, { wch: 22 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Qarzdorlar");
  const label = { daily: "kunlik", monthly: "oylik", yearly: "yillik", all: "butun_tarix" }[periodKey];
  XLSX.writeFile(wb, `qarzdorlar_${label}_${fmtDate(now.toISOString())}.xlsx`);
}

// ---------- shared design tokens ----------
const bgGrad = "linear-gradient(135deg, #0A1A2E 0%, #123B4D 28%, #1B4B5A 52%, #1E5A63 72%, #2E8C7C 100%)";
const glass = "rgba(255,255,255,0.10)";
const glassBorder = "rgba(255,255,255,0.22)";
const text = "#EAF3F1";
const textSoft = "rgba(234,243,241,0.62)";
const textFaint = "rgba(234,243,241,0.4)";
const accentGrad = "linear-gradient(90deg, #1F9C8C 0%, #35C99A 100%)";
const accent = "#2FBF9E";
const danger = "#E0554B";
const warn = "#E8B94A";
const blackTag = "#7C8792";
const glassPanel = "rgba(9,22,27,0.55)";
const tagColor = { green: accent, yellow: warn, red: danger, black: blackTag };
const tagLabel = { green: "Yashil", yellow: "Sariq", red: "Qizil", black: "Qora" };

const inputStyle = {
  width: "100%", boxSizing: "border-box", padding: "10px 12px", borderRadius: 9,
  border: `1px solid ${glassBorder}`, background: "rgba(255,255,255,0.05)", color: text,
  fontSize: 13, marginBottom: 8, outline: "none",
};

function FontLoad() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=Space+Mono:wght@400;700&display=swap');
      .qk-root { font-family: 'Space Grotesk', sans-serif; }
      .qk-serif { font-family: 'Space Grotesk', sans-serif; font-weight: 600; letter-spacing: -0.01em; }
      .qk-mono { font-family: 'Space Mono', monospace; }
      .qk-input::placeholder { color: rgba(234,243,241,0.35); }
      @keyframes qkFade { 0%{opacity:0; transform:translateY(4px);} 10%{opacity:1; transform:translateY(0);} 90%{opacity:1;} 100%{opacity:0; transform:translateY(-4px);} }

      html, body, #root { height: 100%; }
      body {
        margin: 0;
        background: radial-gradient(circle at 50% 0%, #16383D 0%, #08181A 70%);
        min-height: 100dvh;
      }
      #root {
        min-height: 100dvh;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 32px 16px;
        box-sizing: border-box;
      }
      .qk-shell {
        width: 100%;
        max-width: 400px;
        min-height: 660px;
        border-radius: 26px;
        box-shadow: 0 40px 80px -20px rgba(0,0,0,0.6), 0 0 0 1px rgba(255,255,255,0.06);
      }
      @media (max-width: 480px) {
        #root { padding: 0; }
        .qk-shell { max-width: 100%; min-height: 100dvh; border-radius: 0; box-shadow: none; }
      }
    `}</style>
  );
}

function SellerSelect({ sellers, value, onChange }) {
  if (sellers.length === 0) return <p style={{ fontSize: 11, color: textSoft, margin: "0 0 8px" }}>Avval "Sotuvchi" bo'limidan sotuvchi qo'shing.</p>;
  return (
    <select value={value} onChange={(e) => onChange(e.target.value)} style={{ ...inputStyle, appearance: "auto" }}>
      {sellers.map((s) => <option key={s.id} value={s.id} style={{ color: "#000" }}>{s.name}</option>)}
    </select>
  );
}

// ================= MODALS =================

function AddDebtModal({ shop, refreshAll, onClose }) {
  const [q, setQ] = useState("");
  const [mode, setMode] = useState("search");
  const [target, setTarget] = useState(null);
  const [amount, setAmount] = useState("");
  const [dueDate, setDueDate] = useState("");
  const [sellerId, setSellerId] = useState(shop.sellers[0]?.id || "");
  const [nName, setNName] = useState("");
  const [nPhone, setNPhone] = useState("");
  const [nAmount, setNAmount] = useState("");
  const [nDue, setNDue] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const matches = (shop.debtors || [])
    .filter((d) => !q.trim() || d.name.toLowerCase().includes(q.trim().toLowerCase()))
    .slice(0, 30);

  const sellerName = (id) => shop.sellers.find((s) => s.id === id)?.name || "";

  const saveAddToExisting = async () => {
    const amt = parseInt(amount.replace(/\D/g, ""), 10);
    if (!amt) return setErr("Summani kiriting");
    setBusy(true);
    try {
      await addAmountToDebtor(target.id, target.amount + amt, dueDate || undefined);
      await addHistoryRow(shop.id, { kind: "added", debtorName: target.name, amount: amt, sellerName: sellerName(sellerId) });
      await refreshAll();
      onClose();
    } catch (e) { setErr("Xatolik yuz berdi, qayta urinib ko'ring"); setBusy(false); }
  };

  const saveNew = async () => {
    if (!nName.trim()) return setErr("Ism familiyani kiriting");
    const amt = parseInt(nAmount.replace(/\D/g, ""), 10);
    if (!amt) return setErr("Qarz summasini kiriting");
    if (!nDue) return setErr("Muddatni kiriting");
    setBusy(true);
    try {
      await createDebtorRow(shop.id, { name: nName.trim(), phone: nPhone.trim(), amount: amt, dueDate: nDue, sellerId });
      await addHistoryRow(shop.id, { kind: "added", debtorName: nName.trim(), amount: amt, sellerName: sellerName(sellerId) });
      await refreshAll();
      onClose();
    } catch (e) { setErr("Xatolik yuz berdi, qayta urinib ko'ring"); setBusy(false); }
  };

  return (
    <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "flex-end", borderRadius: 16, zIndex: 20 }}>
      <div style={{ width: "100%", maxHeight: "88%", overflowY: "auto", background: glassPanel, backdropFilter: "blur(28px)", WebkitBackdropFilter: "blur(28px)", borderRadius: "20px 20px 0 0", padding: "18px 18px 22px", border: `1px solid ${glassBorder}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <p className="qk-serif" style={{ fontSize: 16, fontWeight: 600, color: text, margin: 0 }}>Yangi qarz yaratish</p>
          <button onClick={onClose} style={{ background: "none", border: "none", color: textSoft, fontSize: 18, cursor: "pointer" }}>✕</button>
        </div>

        {mode !== "newCustomer" && (
          <>
            <input className="qk-input" value={q} onChange={(e) => { setQ(e.target.value); setErr(""); }} placeholder="🔍 Ism familiya orqali qidirish" style={inputStyle} />
            <p style={{ fontSize: 11, color: textFaint, margin: "0 0 6px" }}>Mavjud qarzdorlar:</p>
            <div style={{ maxHeight: 180, overflowY: "auto", marginBottom: 10 }}>
              {matches.length === 0 && <p style={{ fontSize: 12, color: textFaint, margin: "0 0 8px" }}>Hech kim topilmadi.</p>}
              {matches.map((m) => (
                <button key={m.id} onClick={() => { setTarget(m); setMode("addExisting"); setErr(""); }} style={{ width: "100%", textAlign: "left", display: "flex", justifyContent: "space-between", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 10, padding: "10px 12px", marginBottom: 6, cursor: "pointer" }}>
                  <span style={{ fontSize: 13, color: text }}>{m.name}</span>
                  <span className="qk-mono" style={{ fontSize: 12, color: danger }}>{fmt(m.amount)}</span>
                </button>
              ))}
            </div>
            <button onClick={() => { setMode("newCustomer"); setNName(q); setErr(""); }} style={{ width: "100%", padding: "11px", borderRadius: 9, border: `1px dashed ${glassBorder}`, background: "none", color: accent, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
              + Yangi mijoz sifatida qo'shish
            </button>
          </>
        )}

        {mode === "addExisting" && target && (
          <div style={{ marginTop: 14, borderTop: `1px solid ${glassBorder}`, paddingTop: 14 }}>
            <p style={{ fontSize: 13, color: textSoft, margin: "0 0 10px" }}>
              <strong style={{ color: text }}>{target.name}</strong> uchun qo'shimcha qarz — joriy: <span className="qk-mono">{fmt(target.amount)}</span>
            </p>
            <input className="qk-input" value={amount} onChange={(e) => setAmount(e.target.value.replace(/\D/g, ""))} placeholder="Qo'shimcha summa" inputMode="numeric" style={inputStyle} />
            <input className="qk-input" type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} style={inputStyle} />
            <SellerSelect sellers={shop.sellers} value={sellerId} onChange={setSellerId} />
            {err && <p style={{ fontSize: 12, color: danger, margin: "4px 0" }}>{err}</p>}
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <button onClick={() => setMode("search")} style={{ flex: 1, padding: "11px", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Orqaga</button>
              <button disabled={busy} onClick={saveAddToExisting} style={{ flex: 1, padding: "11px", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "Saqlash"}</button>
            </div>
          </div>
        )}

        {mode === "newCustomer" && (
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "rgba(47,191,158,0.12)", border: `1px solid ${accent}`, borderRadius: 9, padding: "9px 12px", marginBottom: 10 }}>
              <span style={{ fontSize: 12, color: textSoft }}>Mijoz ID raqami</span>
              <span className="qk-mono" style={{ fontSize: 12, color: accent, fontWeight: 700 }}>saqlanganda avtomatik beriladi</span>
            </div>
            <input className="qk-input" value={nName} onChange={(e) => setNName(e.target.value)} placeholder="Ism familiya" style={inputStyle} />
            <input className="qk-input" value={nPhone} onChange={(e) => setNPhone(e.target.value)} placeholder="Telefon raqami" style={inputStyle} />
            <input className="qk-input" value={nAmount} onChange={(e) => setNAmount(e.target.value.replace(/\D/g, ""))} placeholder="Qarz summasi" inputMode="numeric" style={inputStyle} />
            <input className="qk-input" type="date" value={nDue} onChange={(e) => setNDue(e.target.value)} style={inputStyle} />
            <SellerSelect sellers={shop.sellers} value={sellerId} onChange={setSellerId} />
            {err && <p style={{ fontSize: 12, color: danger, margin: "4px 0" }}>{err}</p>}
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <button onClick={() => setMode("search")} style={{ flex: 1, padding: "11px", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Orqaga</button>
              <button disabled={busy} onClick={saveNew} style={{ flex: 1, padding: "11px", borderRadius: 9, border: "none", background: danger, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "Saqlash"}</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function PayDebtModal({ shop, refreshAll, onClose }) {
  const [q, setQ] = useState("");
  const [target, setTarget] = useState(null);
  const [payAmount, setPayAmount] = useState("");
  const [sellerId, setSellerId] = useState(shop.sellers[0]?.id || "");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const list = (q.trim() ? shop.debtors.filter((d) => d.name.toLowerCase().includes(q.trim().toLowerCase())) : shop.debtors)
    .slice().sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const grouped = {};
  list.forEach((d) => { const key = fmtDate(d.createdAt); if (!grouped[key]) grouped[key] = []; grouped[key].push(d); });

  const sellerName = (id) => shop.sellers.find((s) => s.id === id)?.name || "";

  const savePay = async () => {
    const amt = parseInt(payAmount.replace(/\D/g, ""), 10);
    if (!amt) return setErr("To'lov summasini kiriting");
    setBusy(true);
    try {
      const remaining = amt >= target.amount ? 0 : target.amount - amt;
      await payDebtorRow(target.id, remaining);
      await addHistoryRow(shop.id, { kind: "paid", debtorName: target.name, amount: amt, sellerName: sellerName(sellerId), note: remaining <= 0 ? "To'liq qoplandi" : "Qisman to'landi" });
      await refreshAll();
      onClose();
    } catch (e) { setErr("Xatolik yuz berdi, qayta urinib ko'ring"); setBusy(false); }
  };

  return (
    <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "flex-end", borderRadius: 16, zIndex: 20 }}>
      <div style={{ width: "100%", maxHeight: "88%", overflowY: "auto", background: glassPanel, backdropFilter: "blur(28px)", WebkitBackdropFilter: "blur(28px)", borderRadius: "20px 20px 0 0", padding: "18px 18px 22px", border: `1px solid ${glassBorder}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <p className="qk-serif" style={{ fontSize: 16, fontWeight: 600, color: text, margin: 0 }}>Qarz qaytarish</p>
          <button onClick={onClose} style={{ background: "none", border: "none", color: textSoft, fontSize: 18, cursor: "pointer" }}>✕</button>
        </div>

        {!target ? (
          <>
            <input className="qk-input" value={q} onChange={(e) => setQ(e.target.value)} placeholder="🔍 Ism familiya orqali qidirish" style={inputStyle} />
            {list.length === 0 && <p style={{ fontSize: 12, color: textFaint }}>Qarzdorlar yo'q.</p>}
            {Object.keys(grouped).map((dateKey) => (
              <div key={dateKey} style={{ marginBottom: 10 }}>
                <p style={{ fontSize: 11, color: textFaint, margin: "8px 0 6px" }}>{dateKey}</p>
                {grouped[dateKey].map((d) => (
                  <button key={d.id} onClick={() => setTarget(d)} style={{ width: "100%", textAlign: "left", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 10, padding: "10px 12px", marginBottom: 6, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between" }}>
                      <span style={{ fontSize: 13, color: text, fontWeight: 600 }}>{d.name}</span>
                      <span className="qk-mono" style={{ fontSize: 12, color: danger }}>{fmt(d.amount)}</span>
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", marginTop: 3 }}>
                      <span style={{ fontSize: 11, color: textFaint }}>{fmtTime(d.createdAt)}</span>
                      <span style={{ fontSize: 11, color: textFaint }}>muddat: {fmtDate(d.dueDate)}</span>
                    </div>
                  </button>
                ))}
              </div>
            ))}
          </>
        ) : (
          <div>
            <p style={{ fontSize: 13, color: textSoft, margin: "0 0 10px" }}>
              <strong style={{ color: text }}>{target.name}</strong> — joriy qarz: <span className="qk-mono">{fmt(target.amount)}</span>
            </p>
            <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
              <input className="qk-input" value={payAmount} onChange={(e) => setPayAmount(e.target.value.replace(/\D/g, ""))} placeholder="To'lov summasi" inputMode="numeric" style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
              <button onClick={() => setPayAmount(String(target.amount))} title="To'liq summani qo'yish" style={{ width: 42, borderRadius: 9, border: `1px solid ${glassBorder}`, background: "rgba(232,185,74,0.15)", color: warn, fontSize: 18, cursor: "pointer" }}>🪙</button>
            </div>
            <p style={{ fontSize: 10, color: textFaint, margin: "0 0 10px" }}>🪙 — qarzning to'liq summasini avtomatik qo'yadi</p>
            <SellerSelect sellers={shop.sellers} value={sellerId} onChange={setSellerId} />
            {err && <p style={{ fontSize: 12, color: danger, margin: "4px 0" }}>{err}</p>}
            <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
              <button onClick={() => setTarget(null)} style={{ flex: 1, padding: "11px", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Orqaga</button>
              <button disabled={busy} onClick={savePay} style={{ flex: 1, padding: "11px", borderRadius: 9, border: "none", background: accent, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "Saqlash"}</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function DebtorProfile({ d, shop, refreshAll, onBack }) {
  const [name, setName] = useState(d.name);
  const [phone, setPhone] = useState(d.phone || "");
  const [amount, setAmount] = useState(String(d.amount));
  const [due, setDue] = useState(d.dueDate || "");
  const [color, setColor] = useState(d.colorOverride || autoColor(d.dueDate));
  const [saved, setSaved] = useState(false);
  const [busy, setBusy] = useState(false);

  const save = async () => {
    setBusy(true);
    try {
      await updateDebtorProfileRow(d.id, { name: name.trim(), phone: phone.trim(), amount: parseInt(amount, 10) || 0, dueDate: due, colorOverride: color });
      await refreshAll();
      setSaved(true);
      setTimeout(() => setSaved(false), 1500);
    } finally { setBusy(false); }
  };

  const sellerName = shop.sellers.find((s) => s.id === d.sellerId)?.name;

  return (
    <div>
      <button onClick={onBack} style={{ background: "none", border: "none", color: accent, fontSize: 12, fontWeight: 600, cursor: "pointer", padding: 0, marginBottom: 14 }}>← Ro'yxatga</button>
      {d.debtorNumber != null && (
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "rgba(47,191,158,0.12)", border: `1px solid ${accent}`, borderRadius: 9, padding: "9px 12px", marginBottom: 10 }}>
          <span style={{ fontSize: 12, color: textSoft }}>Mijoz ID raqami</span>
          <span className="qk-mono" style={{ fontSize: 14, color: accent, fontWeight: 700 }}>{d.debtorNumber}</span>
        </div>
      )}
      <input className="qk-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="Ism familiya" style={inputStyle} />
      <input className="qk-input" value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefon raqami" style={inputStyle} />
      <input className="qk-input" value={amount} onChange={(e) => setAmount(e.target.value.replace(/\D/g, ""))} placeholder="Qarz summasi" inputMode="numeric" style={inputStyle} />
      <input className="qk-input" type="date" value={due} onChange={(e) => setDue(e.target.value)} style={inputStyle} />
      {sellerName && <p style={{ fontSize: 11, color: textFaint, margin: "0 0 10px" }}>Qo'shgan sotuvchi: {sellerName}</p>}
      <p style={{ fontSize: 12, color: textSoft, margin: "6px 0 8px" }}>Holat belgisi</p>
      <div style={{ display: "flex", gap: 8, marginBottom: 18 }}>
        {["green", "yellow", "red", "black"].map((c) => (
          <button key={c} onClick={() => setColor(c)} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: color === c ? `2px solid ${tagColor[c]}` : `1px solid ${glassBorder}`, background: color === c ? `${tagColor[c]}22` : "none", color: tagColor[c], fontSize: 11, fontWeight: 700, cursor: "pointer" }}>
            {tagLabel[c]}
          </button>
        ))}
      </div>
      {saved && <p style={{ fontSize: 12, color: accent, margin: "0 0 8px" }}>✓ Saqlandi</p>}
      <button disabled={busy} onClick={save} style={{ width: "100%", padding: "12px", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "Saqlash"}</button>
    </div>
  );
}

function parseDateFlexible(val) {
  if (val instanceof Date && !isNaN(val)) return val.toISOString();
  if (typeof val === "number") {
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    if (!isNaN(d)) return d.toISOString();
  }
  if (typeof val === "string") {
    const m = val.trim().match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/);
    if (m) {
      const [, d, mo, y] = m;
      const dt = new Date(`${y}-${mo.padStart(2, "0")}-${d.padStart(2, "0")}T00:00:00`);
      if (!isNaN(dt)) return dt.toISOString();
    }
  }
  return new Date().toISOString();
}
function parseAmountFlexible(val) {
  const n = parseFloat(String(val).replace(/[^\d.]/g, ""));
  return isNaN(n) ? 0 : Math.round(n);
}
function normalizeKey(k) { return String(k || "").trim().toLowerCase(); }
function findHeaderKey(keys, patterns) {
  return keys.find((k) => { const nk = normalizeKey(k); return patterns.some((p) => nk.includes(p)); });
}

function ImportDebtorsModal({ shop, refreshAll, onClose }) {
  const [rows, setRows] = useState(null); // parsed preview rows
  const [fileName, setFileName] = useState("");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [done, setDone] = useState(null);

  const handleFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setErr(""); setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (json.length === 0) {
        setErr("Fayl bo'sh yoki ma'lumot topilmadi.");
        setRows(null);
        return;
      }
      const sampleKeys = Object.keys(json[0]);
      const nameKey = findHeaderKey(sampleKeys, ["ism", "familiya"]);
      const amountKey = findHeaderKey(sampleKeys, ["miqdor", "summa", "narx", "pul"]);
      const dateKey = findHeaderKey(sampleKeys, ["olingan", "sana", "kun"]);
      const sellerKey = findHeaderKey(sampleKeys, ["sotuvchi"]);

      if (!nameKey || !amountKey) {
        setErr(`Ustunlarni aniqlab bo'lmadi. Faylda topilgan ustunlar: "${sampleKeys.join('", "')}". Kamida ism-familiya va qarz miqdori ustunlari bo'lishi kerak.`);
        setRows(null);
        return;
      }

      const parsed = json.map((row) => ({
        name: String(row[nameKey] || "").trim(),
        amount: parseAmountFlexible(row[amountKey] || 0),
        createdAt: dateKey ? parseDateFlexible(row[dateKey]) : new Date().toISOString(),
        sellerName: sellerKey ? String(row[sellerKey] || "").trim() : "",
      })).filter((r) => r.name && r.amount > 0);

      if (parsed.length === 0) setErr("Faylda to'g'ri to'ldirilgan qatorlar topilmadi — ism va qarz miqdori bo'sh emasligini tekshiring.");
      setRows(parsed);
    } catch (e) {
      setErr("Faylni o'qib bo'lmadi. .xlsx formatida ekanini tekshiring.");
      setRows(null);
    }
  };

  const confirmImport = async () => {
    if (!rows || rows.length === 0) return;
    setBusy(true); setErr("");
    try {
      const result = await bulkImportDebtors(shop.id, shop.sellers, rows);
      await refreshAll();
      setDone(result);
    } catch (e) {
      setErr("Saqlashda xatolik yuz berdi, qayta urinib ko'ring.");
    } finally { setBusy(false); }
  };

  return (
    <div style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "flex-end", borderRadius: 16, zIndex: 20 }}>
      <div style={{ width: "100%", maxHeight: "88%", overflowY: "auto", background: glassPanel, backdropFilter: "blur(28px)", WebkitBackdropFilter: "blur(28px)", borderRadius: "20px 20px 0 0", padding: "18px 18px 22px", border: `1px solid ${glassBorder}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
          <p className="qk-serif" style={{ fontSize: 16, fontWeight: 600, color: text, margin: 0 }}>Excel'dan qarzdorlarni import qilish</p>
          <button onClick={onClose} style={{ background: "none", border: "none", color: textSoft, fontSize: 18, cursor: "pointer" }}>✕</button>
        </div>

        {done ? (
          <div style={{ textAlign: "center", padding: "20px 0" }}>
            <p style={{ fontSize: 30, margin: "0 0 10px" }}>✅</p>
            <p style={{ fontSize: 14, color: text, margin: "0 0 6px" }}>{done.imported} ta qarzdor muvaffaqiyatli qo'shildi</p>
            {done.newSellers > 0 && <p style={{ fontSize: 12, color: textFaint, margin: "0 0 16px" }}>{done.newSellers} ta yangi sotuvchi ham avtomatik yaratildi</p>}
            <button onClick={onClose} style={{ padding: "11px 24px", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>Yopish</button>
          </div>
        ) : !rows ? (
          <div>
            <p style={{ fontSize: 12, color: textSoft, margin: "0 0 14px" }}>
              Eski qarzdorlar ro'yxatini shablon bo'yicha (ISM FAMILYA, QARZ MIQDORI, QARZ OLINGAN kun/oy/yil, QARZ BERGAN SOTUVCHI) to'ldirilgan .xlsx faylni tanlang.
            </p>
            <label style={{ display: "block", border: `2px dashed ${glassBorder}`, borderRadius: 14, padding: "30px 16px", textAlign: "center", cursor: "pointer" }}>
              <input type="file" accept=".xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
              <p style={{ fontSize: 26, margin: "0 0 8px" }}>📤</p>
              <p style={{ fontSize: 13, color: accent, fontWeight: 600, margin: 0 }}>Fayl tanlash uchun bosing</p>
              <p style={{ fontSize: 11, color: textFaint, margin: "4px 0 0" }}>.xlsx yoki .xls</p>
            </label>
            {err && <p style={{ fontSize: 12, color: danger, margin: "12px 0 0" }}>{err}</p>}
          </div>
        ) : (
          <div>
            <p style={{ fontSize: 12, color: textSoft, margin: "0 0 4px" }}>{fileName}</p>
            <p style={{ fontSize: 13, color: text, fontWeight: 600, margin: "0 0 12px" }}>{rows.length} ta qator topildi — tekshirib chiqing:</p>
            <div style={{ maxHeight: 220, overflowY: "auto", marginBottom: 14, border: `1px solid ${glassBorder}`, borderRadius: 12 }}>
              {rows.map((r, i) => (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 12px", borderBottom: i < rows.length - 1 ? `1px solid ${glassBorder}` : "none" }}>
                  <div>
                    <p style={{ fontSize: 13, color: text, margin: "0 0 2px" }}>{r.name}</p>
                    <p style={{ fontSize: 11, color: textFaint, margin: 0 }}>{fmtDate(r.createdAt)}{r.sellerName ? ` · ${r.sellerName}` : ""}</p>
                  </div>
                  <p className="qk-mono" style={{ fontSize: 13, color: danger, margin: 0 }}>{fmt(r.amount)}</p>
                </div>
              ))}
            </div>
            {err && <p style={{ fontSize: 12, color: danger, margin: "0 0 10px" }}>{err}</p>}
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={() => { setRows(null); setFileName(""); }} style={{ flex: 1, padding: "12px", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Boshqa fayl</button>
              <button disabled={busy} onClick={confirmImport} style={{ flex: 1, padding: "12px", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "✓ Tasdiqlash"}</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function ReportPage({ shop, refreshAll }) {
  const [search, setSearch] = useState("");
  const [dateFilter, setDateFilter] = useState("");
  const [selected, setSelected] = useState(null);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showImportModal, setShowImportModal] = useState(false);
  const listRef = useRef(null);

  const sorted = [...shop.debtors].sort((a, b) => a.name.localeCompare(b.name));
  const filtered = sorted.filter((d) => {
    if (search.trim() && !d.name.toLowerCase().includes(search.trim().toLowerCase())) return false;
    if (dateFilter && fmtDate(d.createdAt) !== fmtDate(new Date(dateFilter).toISOString())) return false;
    return true;
  });
  const lettersPresent = new Set(sorted.map((d) => d.name.charAt(0).toUpperCase()));

  const scrollToLetter = (L) => {
    const el = document.getElementById(`letter-${L}`);
    if (el && listRef.current) listRef.current.scrollTop = el.offsetTop - 4;
  };

  if (selected) {
    const fresh = shop.debtors.find((x) => x.id === selected.id) || selected;
    return <DebtorProfile d={fresh} shop={shop} refreshAll={refreshAll} onBack={() => setSelected(null)} />;
  }

  let lastLetter = null;

  return (
    <div style={{ position: "relative", height: "100%" }}>
      <div style={{ display: "flex", height: "100%", paddingBottom: 44 }}>
        <div style={{ flex: 1, minWidth: 0, display: "flex", flexDirection: "column" }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            <div style={{ position: "relative", flexShrink: 0 }}>
              <input type="date" value={dateFilter} onChange={(e) => setDateFilter(e.target.value)} style={{ position: "absolute", inset: 0, opacity: 0, width: 40, height: 40, cursor: "pointer" }} />
              <div style={{ width: 40, height: 40, borderRadius: 9, background: dateFilter ? accent : glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>📅</div>
            </div>
            <input className="qk-input" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="🔍 Ism familiya orqali qidirish" style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
          </div>
          <div ref={listRef} style={{ flex: 1, overflowY: "auto", maxHeight: 330, marginTop: 8 }}>
            {filtered.length === 0 && <p style={{ fontSize: 12, color: textFaint, textAlign: "center", padding: "30px 0" }}>Qarzdorlar topilmadi.</p>}
            {filtered.map((d) => {
              const L = d.name.charAt(0).toUpperCase();
              const showHeader = L !== lastLetter;
              lastLetter = L;
              return (
                <div key={d.id} id={showHeader ? `letter-${L}` : undefined}>
                  {showHeader && <p style={{ fontSize: 11, color: accent, fontWeight: 700, margin: "10px 0 4px" }}>{L}</p>}
                  <button onClick={() => setSelected(d)} style={{ width: "100%", textAlign: "left", display: "flex", justifyContent: "space-between", alignItems: "center", background: "none", border: "none", borderBottom: `1px solid ${glassBorder}`, padding: "10px 2px", cursor: "pointer" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <span style={{ width: 8, height: 8, borderRadius: "50%", background: tagColor[d.colorOverride || autoColor(d.dueDate)], display: "inline-block", border: "1px solid rgba(255,255,255,0.3)" }} />
                      <span style={{ fontSize: 13, color: text }}>{d.name}</span>
                    </div>
                    <span className="qk-mono" style={{ fontSize: 12, color: danger }}>{fmt(d.amount)}</span>
                  </button>
                </div>
              );
            })}
          </div>
        </div>
        <div style={{ width: 18, display: "flex", flexDirection: "column", alignItems: "center", overflowY: "auto", marginLeft: 4, maxHeight: 340 }}>
          {ALPHABET.filter((L) => lettersPresent.has(L)).map((L) => (
            <button key={L} onClick={() => scrollToLetter(L)} style={{ background: "none", border: "none", color: textFaint, fontSize: 9, fontWeight: 700, cursor: "pointer", padding: "1px 0" }}>{L}</button>
          ))}
        </div>
      </div>

      <div style={{ position: "absolute", bottom: 0, left: 0 }}>
        <button
          onClick={() => setShowImportModal(true)}
          title="Excel'dan qarzdorlarni import qilish"
          style={{ width: 36, height: 36, borderRadius: 10, background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15, cursor: "pointer" }}
        >
          📤
        </button>
      </div>

      <div style={{ position: "absolute", bottom: 0, right: 0 }}>
        <button
          onClick={() => setShowExportMenu((v) => !v)}
          title="Ro'yxatni yuklab olish"
          style={{ width: 36, height: 36, borderRadius: 10, background: showExportMenu ? accent : glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, cursor: "pointer" }}
        >
          ⬇️
        </button>
        {showExportMenu && (
          <div style={{ position: "absolute", bottom: 44, right: 0, background: glassPanel, backdropFilter: "blur(24px)", WebkitBackdropFilter: "blur(24px)", border: `1px solid ${glassBorder}`, borderRadius: 14, padding: 6, width: 160, boxShadow: "0 10px 30px rgba(0,0,0,0.4)" }}>
            {[
              { key: "daily", label: "Kunlik" },
              { key: "monthly", label: "Oylik" },
              { key: "yearly", label: "Yillik" },
              { key: "all", label: "Butun tarix" },
            ].map((opt) => (
              <button
                key={opt.key}
                onClick={() => { exportDebtorsXlsx(shop, opt.key); setShowExportMenu(false); }}
                style={{ width: "100%", textAlign: "left", padding: "9px 10px", borderRadius: 8, border: "none", background: "none", color: text, fontSize: 13, cursor: "pointer" }}
              >
                {opt.label}
              </button>
            ))}
          </div>
        )}
      </div>

      {showImportModal && <ImportDebtorsModal shop={shop} refreshAll={refreshAll} onClose={() => setShowImportModal(false)} />}
    </div>
  );
}

function OwnerAnalytics({ shop }) {
  const debtors = shop.debtors;
  const totalDebt = debtors.reduce((s, d) => s + d.amount, 0);
  const counts = { green: 0, yellow: 0, red: 0, black: 0 };
  debtors.forEach((d) => { counts[d.colorOverride || autoColor(d.dueDate)]++; });
  const maxCount = Math.max(1, ...Object.values(counts));
  const upcoming = debtors.filter((d) => daysUntil(d.dueDate) <= 7 && daysUntil(d.dueDate) >= 0).sort((a, b) => daysUntil(a.dueDate) - daysUntil(b.dueDate));
  const overdue = debtors.filter((d) => daysUntil(d.dueDate) < 0);
  const overdueSum = overdue.reduce((s, d) => s + d.amount, 0);
  const topDebtors = [...debtors].sort((a, b) => b.amount - a.amount).slice(0, 5);
  const maxTop = Math.max(1, ...topDebtors.map((d) => d.amount));
  const avgDebt = debtors.length ? totalDebt / debtors.length : 0;

  const now = Date.now();
  const week = 7 * 86400000;
  const toTime = (dstr) => { const [dd, mm, yy] = dstr.split("."); return new Date(`${yy}-${mm}-${dd}T00:00:00`).getTime(); };
  const addedLast7 = (shop.historyLog || []).filter((h) => h.kind === "added" && now - toTime(h.date) <= week);
  const paidLast7 = (shop.historyLog || []).filter((h) => h.kind === "paid" && now - toTime(h.date) <= week);
  const addedSum7 = addedLast7.reduce((s, h) => s + h.amount, 0);
  const paidSum7 = paidLast7.reduce((s, h) => s + h.amount, 0);

  const Stat = ({ label, value, color }) => (
    <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "12px 14px", flex: 1, minWidth: 0 }}>
      <p style={{ fontSize: 10, color: textSoft, margin: "0 0 4px" }}>{label}</p>
      <p className="qk-mono" style={{ fontSize: 15, fontWeight: 700, color: color || text, margin: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{value}</p>
    </div>
  );

  return (
    <>
      <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 14, padding: "16px 18px", marginBottom: 12 }}>
        <p style={{ fontSize: 11, color: textSoft, margin: "0 0 4px" }}>UMUMIY QARZ MIQDORI</p>
        <p className="qk-mono" style={{ fontSize: 24, fontWeight: 700, color: accent, margin: 0 }}>{fmt(totalDebt)}</p>
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <Stat label="Jami mijozlar" value={debtors.length} />
        <Stat label="O'rtacha qarz" value={fmt(avgDebt)} />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <Stat label="Muddati o'tgan" value={`${overdue.length} ta`} color={danger} />
        <Stat label="O'tgan qarz summasi" value={fmt(overdueSum)} color={danger} />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 20 }}>
        <Stat label="So'nggi 7 kun berilgan" value={fmt(addedSum7)} color={danger} />
        <Stat label="So'nggi 7 kun qaytarilgan" value={fmt(paidSum7)} color={accent} />
      </div>

      <p style={{ fontSize: 13, fontWeight: 700, color: text, margin: "0 0 10px" }}>Eng yirik 5 qarzdor</p>
      {topDebtors.length === 0 ? <p style={{ fontSize: 12, color: textFaint, marginBottom: 18 }}>Ma'lumot yo'q.</p> : (
        <div style={{ marginBottom: 20 }}>
          {topDebtors.map((d) => (
            <div key={d.id} style={{ marginBottom: 9 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: textSoft, marginBottom: 3 }}>
                <span>{d.name}</span><span className="qk-mono">{fmt(d.amount)}</span>
              </div>
              <div style={{ height: 7, borderRadius: 4, background: "rgba(255,255,255,0.08)" }}>
                <div style={{ height: "100%", width: `${(d.amount / maxTop) * 100}%`, borderRadius: 4, background: accentGrad }} />
              </div>
            </div>
          ))}
        </div>
      )}

      <p style={{ fontSize: 13, fontWeight: 700, color: text, margin: "0 0 10px" }}>Rang bo'yicha qarzdorlar</p>
      <div style={{ marginBottom: 20 }}>
        {["green", "yellow", "red", "black"].map((c) => (
          <div key={c} style={{ marginBottom: 9 }}>
            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: textSoft, marginBottom: 3 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span style={{ width: 9, height: 9, borderRadius: "50%", background: tagColor[c], display: "inline-block", border: "1px solid rgba(255,255,255,0.3)" }} />
                {tagLabel[c]}
              </span>
              <span>{counts[c]} ta</span>
            </div>
            <div style={{ height: 7, borderRadius: 4, background: "rgba(255,255,255,0.08)" }}>
              <div style={{ height: "100%", width: `${(counts[c] / maxCount) * 100}%`, borderRadius: 4, background: tagColor[c] }} />
            </div>
          </div>
        ))}
      </div>

      <p style={{ fontSize: 13, fontWeight: 700, color: text, margin: "0 0 10px" }}>Muddati yaqin qolgan qarzdorlar</p>
      {upcoming.length === 0 ? <p style={{ fontSize: 12, color: textFaint }}>Yo'q.</p> : upcoming.map((d) => (
        <div key={d.id} style={{ display: "flex", justifyContent: "space-between", padding: "9px 0", borderBottom: `1px dashed ${glassBorder}` }}>
          <div>
            <p style={{ fontSize: 13, color: text, margin: "0 0 2px" }}>{d.name}</p>
            <p style={{ fontSize: 11, color: warn, margin: 0 }}>{daysUntil(d.dueDate)} kun qoldi</p>
          </div>
          <p className="qk-mono" style={{ fontSize: 13, color: text, margin: 0 }}>{fmt(d.amount)}</p>
        </div>
      ))}
    </>
  );
}

function HistoryPage({ historyLog }) {
  const [tab, setTab] = useState("paid");
  const list = historyLog.filter((h) => tab === "paid" ? h.kind === "paid" : h.kind === "added");
  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <button onClick={() => setTab("paid")} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: "none", background: tab === "paid" ? accent : "rgba(255,255,255,0.06)", color: tab === "paid" ? "#08221E" : textSoft, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>To'langan qarzlar</button>
        <button onClick={() => setTab("added")} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: "none", background: tab === "added" ? danger : "rgba(255,255,255,0.06)", color: tab === "added" ? "#fff" : textSoft, fontSize: 12, fontWeight: 700, cursor: "pointer" }}>Yangi / o'zgartirilgan</button>
      </div>
      {list.length === 0 && <p style={{ fontSize: 12, color: textFaint, textAlign: "center", padding: "20px 0" }}>Tarix bo'sh.</p>}
      {list.map((h) => (
        <div key={h.id} style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "11px 14px", marginBottom: 8 }}>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <p style={{ fontSize: 13, fontWeight: 600, color: text, margin: "0 0 2px" }}>{h.debtorName}</p>
            <p className="qk-mono" style={{ fontSize: 13, color: tab === "paid" ? accent : danger, margin: 0 }}>{fmt(h.amount)}</p>
          </div>
          <p style={{ fontSize: 11, color: textFaint, margin: 0 }}>{h.date} · {h.time}{h.sellerName ? ` · ${h.sellerName}` : ""}{h.note ? ` · ${h.note}` : ""}</p>
        </div>
      ))}
    </div>
  );
}

function SellersPage({ shop, refreshAll }) {
  const [showForm, setShowForm] = useState(false);
  const [name, setName] = useState("");
  const [busy, setBusy] = useState(false);

  const add = async () => {
    if (!name.trim()) return;
    setBusy(true);
    try { await addSellerRow(shop.id, name.trim()); await refreshAll(); setName(""); setShowForm(false); }
    finally { setBusy(false); }
  };
  const remove = async (id) => { await removeSellerRow(id); await refreshAll(); };

  return (
    <div>
      {showForm ? (
        <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: 14, marginBottom: 16 }}>
          <input className="qk-input" value={name} onChange={(e) => setName(e.target.value)} placeholder="Ism familiya" style={inputStyle} autoFocus />
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => { setShowForm(false); setName(""); }} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Bekor qilish</button>
            <button disabled={busy} onClick={add} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>Saqlash</button>
          </div>
        </div>
      ) : (
        <button onClick={() => setShowForm(true)} style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1px dashed ${glassBorder}`, background: "none", color: accent, fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 16 }}>+ Qo'shish</button>
      )}
      {shop.sellers.length === 0 && <p style={{ fontSize: 12, color: textFaint, textAlign: "center", padding: "12px 0" }}>Hali sotuvchi qo'shilmagan.</p>}
      {shop.sellers.map((s) => (
        <div key={s.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "11px 14px", marginBottom: 8 }}>
          <span style={{ fontSize: 13, color: text }}>{s.name}</span>
          <button onClick={() => remove(s.id)} style={{ background: "none", border: "1px solid rgba(224,85,75,0.4)", borderRadius: 8, color: danger, fontSize: 12, padding: "5px 8px", cursor: "pointer" }}>✕</button>
        </div>
      ))}
      <p style={{ fontSize: 11, color: textFaint, marginTop: 14 }}>Bu ro'yxatdagi sotuvchilar "Qarz qo'shish" va "Qarz qaytarish" oynalarida tanlanadi — kim qarz kiritgani yoki to'lov qabul qilgani Tarix bo'limida shu orqali bilinadi.</p>
    </div>
  );
}

// ================= OWNER PANEL =================

function OwnerPanel({ shop, refreshAll, onLogout }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [slideIdx, setSlideIdx] = useState(0);
  const [showAddModal, setShowAddModal] = useState(false);
  const [showPayModal, setShowPayModal] = useState(false);

  const debtors = shop.debtors;
  const totalDebt = useMemo(() => debtors.reduce((s, d) => s + d.amount, 0), [debtors]);

  useEffect(() => {
    if (debtors.length < 2) return;
    const t = setInterval(() => setSlideIdx((i) => (i + 1) % debtors.length), 3200);
    return () => clearInterval(t);
  }, [debtors.length]);

  const menuItems = [
    { key: "dashboard", label: "Dashboard", icon: "▦" },
    { key: "report", label: "Qarz hisoboti", icon: "📋" },
    { key: "analytics", label: "Analitika", icon: "📊" },
    { key: "history", label: "Tarix", icon: "🕒" },
    { key: "sellers", label: "Sotuvchi", icon: "🧑‍💼" },
    { key: "contact", label: "Admin bilan bog'lanish", icon: "💬" },
  ];
  const pageTitle = menuItems.find((m) => m.key === page)?.label || "Dashboard";
  const goto = (key) => { setPage(key); setSidebarOpen(false); };

  let body;
  if (page === "dashboard") {
    const cur = debtors[slideIdx % Math.max(debtors.length, 1)];
    body = (
      <>
        <div style={{ backgroundImage: accentGrad, borderRadius: 16, padding: "18px 20px", marginBottom: 12 }}>
          <p style={{ fontSize: 11, color: "rgba(8,34,30,0.75)", margin: "0 0 6px", fontWeight: 700, letterSpacing: 0.5 }}>UMUMIY QARZ MIQDORI</p>
          <p className="qk-mono" style={{ fontSize: 27, fontWeight: 700, color: "#08221E", margin: 0 }}>{fmt(totalDebt)}</p>
        </div>
        <div style={{ background: "rgba(255,255,255,0.03)", border: `1px solid ${glassBorder}`, borderRadius: 16, padding: "20px", marginBottom: 18, minHeight: 76, display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", textAlign: "center" }}>
          {debtors.length === 0 ? <p style={{ fontSize: 12, color: textFaint, margin: 0 }}>Hali qarzdorlar yo'q</p> : (
            <div key={slideIdx} style={{ animation: "qkFade 3.2s ease-in-out" }}>
              <p className="qk-serif" style={{ fontSize: 17, fontWeight: 600, color: text, margin: "0 0 4px" }}>{cur.name}</p>
              <p className="qk-mono" style={{ fontSize: 14, color: accent, margin: 0 }}>{fmt(cur.amount)}</p>
            </div>
          )}
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={() => setShowAddModal(true)} style={{ flex: 1, padding: "16px 0", borderRadius: 14, border: "none", background: danger, color: "#fff", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>+ Qarz qo'shish</button>
          <button onClick={() => setShowPayModal(true)} style={{ flex: 1, padding: "16px 0", borderRadius: 14, border: "none", background: accent, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>✓ Qarz qaytarish</button>
        </div>
      </>
    );
  } else if (page === "report") {
    body = <ReportPage shop={shop} refreshAll={refreshAll} />;
  } else if (page === "analytics") {
    body = <OwnerAnalytics shop={shop} />;
  } else if (page === "history") {
    body = <HistoryPage historyLog={shop.historyLog} />;
  } else if (page === "sellers") {
    body = <SellersPage shop={shop} refreshAll={refreshAll} />;
  } else if (page === "contact") {
    body = (
      <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 16, padding: "26px 20px", textAlign: "center" }}>
        <div style={{ width: 56, height: 56, borderRadius: "50%", background: accentGrad, margin: "0 auto 14px", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>💬</div>
        <p className="qk-serif" style={{ fontSize: 17, fontWeight: 600, color: text, margin: "0 0 4px" }}>Admin bilan bog'lanish</p>
        <p style={{ fontSize: 12, color: textSoft, margin: "0 0 20px" }}>Savol yoki muammo bo'lsa, quyidagi aloqa orqali murojaat qiling</p>
        <a href="https://t.me/imradjabov" target="_blank" rel="noreferrer" style={{ display: "flex", alignItems: "center", gap: 12, background: "rgba(255,255,255,0.05)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "13px 16px", marginBottom: 10, textDecoration: "none" }}>
          <span style={{ fontSize: 20 }}>✈️</span>
          <div style={{ textAlign: "left" }}><p style={{ fontSize: 11, color: textFaint, margin: "0 0 2px" }}>Telegram</p><p style={{ fontSize: 14, color: accent, fontWeight: 600, margin: 0 }}>@imradjabov</p></div>
        </a>
        <a href="tel:+998904692017" style={{ display: "flex", alignItems: "center", gap: 12, background: "rgba(255,255,255,0.05)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "13px 16px", textDecoration: "none" }}>
          <span style={{ fontSize: 20 }}>📞</span>
          <div style={{ textAlign: "left" }}><p style={{ fontSize: 11, color: textFaint, margin: "0 0 2px" }}>Telefon</p><p className="qk-mono" style={{ fontSize: 14, color: accent, fontWeight: 600, margin: 0 }}>+998 90 469 20 17</p></div>
        </a>
      </div>
    );
  }

  return (
    <>
      <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "16px 18px", borderBottom: `1px solid ${glassBorder}` }}>
        {page !== "dashboard" ? (
          <button onClick={() => goto("dashboard")} style={{ background: "none", border: "none", color: text, fontSize: 18, cursor: "pointer", padding: 0 }}>←</button>
        ) : (
          <button onClick={() => setSidebarOpen(true)} style={{ background: "none", border: "none", color: text, fontSize: 20, cursor: "pointer", padding: 0 }}>☰</button>
        )}
        <div style={{ width: 30, height: 30, borderRadius: 8, background: accentGrad, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 13, color: "#08221E" }}>qk</div>
        <div>
          <p style={{ fontSize: 15, fontWeight: 700, color: text, margin: 0 }}>{pageTitle}</p>
          {page === "dashboard" && <p style={{ fontSize: 10, color: textFaint, margin: 0 }}>{shop.name}</p>}
        </div>
      </div>

      <div style={{ padding: "18px", overflowY: "auto", flex: 1, position: "relative" }}>{body}</div>

      {showAddModal && <AddDebtModal shop={shop} refreshAll={refreshAll} onClose={() => setShowAddModal(false)} />}
      {showPayModal && <PayDebtModal shop={shop} refreshAll={refreshAll} onClose={() => setShowPayModal(false)} />}

      {sidebarOpen && (
        <div onClick={() => setSidebarOpen(false)} style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 5 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ position: "absolute", top: 0, left: 0, bottom: 0, width: 250, background: glassPanel, backdropFilter: "blur(28px)", WebkitBackdropFilter: "blur(28px)", borderRight: `1px solid ${glassBorder}`, padding: "20px 14px", zIndex: 6 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 22, padding: "0 6px" }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: accentGrad, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 14, color: "#08221E" }}>qk</div>
              <p style={{ fontSize: 14, fontWeight: 700, color: text, margin: 0 }}>{shop.name}</p>
            </div>
            {menuItems.map((m) => (
              <button key={m.key} onClick={() => goto(m.key)} style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, textAlign: "left", padding: "11px 10px", borderRadius: 10, border: "none", marginBottom: 4, cursor: "pointer", background: page === m.key ? "rgba(47,191,158,0.15)" : "none", color: page === m.key ? accent : textSoft, fontSize: 13, fontWeight: page === m.key ? 700 : 500 }}>
                <span style={{ fontSize: 15 }}>{m.icon}</span>{m.label}
              </button>
            ))}
            <div style={{ borderTop: `1px solid ${glassBorder}`, marginTop: 14, paddingTop: 14 }}>
              <button onClick={onLogout} style={{ width: "100%", textAlign: "left", padding: "11px 10px", borderRadius: 10, border: "none", background: "none", color: danger, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>⏻ Chiqish</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ================= ADMIN PANEL =================

function BotControlPage({ shops, refreshAll }) {
  const [selectedShopId, setSelectedShopId] = useState(null);
  const [showTokenForm, setShowTokenForm] = useState(false);
  const [tokenVal, setTokenVal] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [assigning, setAssigning] = useState(false);

  const shop = shops.find((s) => s.id === selectedShopId);
  const missingCount = shop ? shop.debtors.filter((d) => d.debtorNumber == null).length : 0;

  const runAssignMissing = async () => {
    setAssigning(true);
    try {
      const count = await assignMissingIds(shop.id);
      await refreshAll();
      setMsg(`✓ ${count} ta mijozga ID berildi`);
      setTimeout(() => setMsg(""), 3000);
    } catch (e) { setMsg("Xatolik yuz berdi"); }
    finally { setAssigning(false); }
  };

  if (!shop) {
    return (
      <>
        <p style={{ fontSize: 12, color: textSoft, margin: "0 0 12px" }}>Bot sozlamalarini ko'rish uchun do'konni tanlang</p>
        {shops.length === 0 && <p style={{ fontSize: 13, color: textSoft, textAlign: "center", padding: "40px 0" }}>Hali do'kon qo'shilmagan.</p>}
        {shops.map((s) => (
          <button key={s.id} onClick={() => setSelectedShopId(s.id)} style={{ width: "100%", textAlign: "left", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "12px 14px", marginBottom: 8, cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <div>
              <p style={{ fontSize: 14, fontWeight: 600, color: text, margin: "0 0 2px" }}>{s.name}</p>
              <p style={{ fontSize: 11, color: textFaint, margin: 0 }}>{s.botToken ? "🟢 Bot ulangan" : "⚪ Bot ulanmagan"}</p>
            </div>
          </button>
        ))}
      </>
    );
  }

  const saveToken = async () => {
    if (!tokenVal.trim()) return;
    setBusy(true);
    try {
      await updateShopBotToken(shop.id, tokenVal.trim());
      const res = await fetch("/api/connect-bot", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ shopId: shop.id }),
      });
      const result = await res.json();
      await refreshAll();
      setShowTokenForm(false); setTokenVal("");
      setMsg(result.ok ? "✓ Bot ulandi va faollashtirildi!" : `Token saqlandi, lekin ulanishda xatolik: ${result.error || ""}`);
      setTimeout(() => setMsg(""), 4000);
    } catch (e) { setMsg("Xatolik yuz berdi"); }
    finally { setBusy(false); }
  };

  const saveDebtorNumber = async (debtorId) => {
    const n = parseInt(editVal, 10);
    if (isNaN(n)) return setEditingId(null);
    try {
      await updateDebtorNumberRow(debtorId, n);
      await refreshAll();
      setEditingId(null); setMsg("");
    } catch (e) {
      setMsg(e && e.code === "23505" ? "Bu ID raqami boshqa mijozda band" : "Xatolik yuz berdi");
    }
  };

  return (
    <>
      <button onClick={() => setSelectedShopId(null)} style={{ background: "none", border: "none", color: accent, fontSize: 12, fontWeight: 600, cursor: "pointer", padding: 0, marginBottom: 14 }}>← Do'konlar ro'yxatiga</button>
      <p style={{ fontSize: 16, fontWeight: 700, color: text, margin: "0 0 2px" }}>{shop.name}</p>
      <p style={{ fontSize: 12, color: textSoft, margin: "0 0 16px" }}>{shop.botToken ? "🟢 Bot ulangan" : "⚪ Bot hali ulanmagan"}</p>

      {missingCount > 0 && (
        <button
          disabled={assigning}
          onClick={runAssignMissing}
          style={{ width: "100%", padding: "12px", borderRadius: 10, border: `1px solid ${accent}`, background: "rgba(47,191,158,0.12)", color: accent, fontSize: 13, fontWeight: 700, cursor: "pointer", marginBottom: 16, opacity: assigning ? 0.6 : 1 }}
        >
          {assigning ? "Beriyapti..." : `🔢 ID'si yo'q ${missingCount} ta mijozga avtomatik ID berish`}
        </button>
      )}

      {!showTokenForm ? (
        <button onClick={() => { setShowTokenForm(true); setTokenVal(shop.botToken || ""); }} style={{ width: "100%", padding: "13px", borderRadius: 10, border: "none", background: accentGrad, color: "#08221E", fontSize: 14, fontWeight: 700, cursor: "pointer", marginBottom: 20 }}>
          {shop.botToken ? "Bot tokenini o'zgartirish" : "🤖 Botni ulash"}
        </button>
      ) : (
        <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: 14, marginBottom: 20 }}>
          <p style={{ fontSize: 12, color: textSoft, margin: "0 0 8px" }}>BotFather'da yaratilgan botning tokenini kiriting:</p>
          <input className="qk-input" value={tokenVal} onChange={(e) => setTokenVal(e.target.value)} placeholder="123456789:AAExxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx" style={inputStyle} />
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={() => { setShowTokenForm(false); setTokenVal(""); }} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, cursor: "pointer" }}>Bekor qilish</button>
            <button disabled={busy} onClick={saveToken} style={{ flex: 1, padding: "10px 0", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>Saqlash</button>
          </div>
        </div>
      )}
      {msg && <p style={{ fontSize: 12, color: msg.startsWith("✓") ? accent : danger, margin: "-12px 0 16px" }}>{msg}</p>}

      <p style={{ fontSize: 13, fontWeight: 700, color: text, margin: "0 0 10px" }}>Qarzdorlar va ID raqamlari</p>
      {shop.debtors.length === 0 && <p style={{ fontSize: 12, color: textFaint }}>Hali qarzdor yo'q.</p>}
      {[...shop.debtors].sort((a, b) => (a.debtorNumber || 999) - (b.debtorNumber || 999)).map((d) => (
        <div key={d.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 10, padding: "9px 12px", marginBottom: 6 }}>
          <span style={{ fontSize: 13, color: text }}>{d.name}</span>
          {editingId === d.id ? (
            <div style={{ display: "flex", gap: 4, alignItems: "center" }}>
              <input value={editVal} onChange={(e) => setEditVal(e.target.value.replace(/\D/g, ""))} className="qk-mono" style={{ width: 70, padding: "4px 6px", borderRadius: 6, border: `1px solid ${accent}`, background: "rgba(255,255,255,0.08)", color: text, fontSize: 12 }} autoFocus />
              <button onClick={() => saveDebtorNumber(d.id)} style={{ background: "none", border: "none", color: accent, fontSize: 13, cursor: "pointer" }}>✓</button>
              <button onClick={() => setEditingId(null)} style={{ background: "none", border: "none", color: textFaint, fontSize: 13, cursor: "pointer" }}>✕</button>
            </div>
          ) : (
            <button onClick={() => { setEditingId(d.id); setEditVal(String(d.debtorNumber ?? "")); }} className="qk-mono" style={{ background: "none", border: `1px solid ${glassBorder}`, borderRadius: 6, padding: "3px 8px", color: accent, fontSize: 12, cursor: "pointer" }}>
              {d.debtorNumber != null ? d.debtorNumber : "— belgilash"}
            </button>
          )}
        </div>
      ))}
    </>
  );
}

function AdminPanel({ data, refreshAll, onLogout }) {
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [page, setPage] = useState("dashboard");
  const [showShopForm, setShowShopForm] = useState(false);
  const [editingShopId, setEditingShopId] = useState(null);
  const [f_name, setF_name] = useState(""); const [f_productType, setF_productType] = useState("");
  const [f_phone, setF_phone] = useState(""); const [f_address, setF_address] = useState("");
  const [f_login, setF_login] = useState(""); const [f_pass, setF_pass] = useState(""); const [f_pass2, setF_pass2] = useState("");
  const [formError, setFormError] = useState(""); const [busy, setBusy] = useState(false);
  const [selectedShopId, setSelectedShopId] = useState(null);
  const [s_login, setS_login] = useState(""); const [s_pass, setS_pass] = useState(""); const [s_pass2, setS_pass2] = useState(""); const [s_msg, setS_msg] = useState("");

  const shops = data.shops || [];
  const totals = useMemo(() => ({
    shopCount: shops.length,
    customerCount: shops.reduce((s, sh) => s + (sh.debtors || []).length, 0),
    totalDebt: shops.reduce((s, sh) => s + (sh.debtors || []).reduce((a, d) => a + d.amount, 0), 0),
  }), [shops]);

  const menuItems = [
    { key: "dashboard", label: "Dashboard", icon: "▦" },
    { key: "shops", label: "Ma'lumotlar va Avtorizatsiya", icon: "🏬" },
    { key: "analytics", label: "Analitika", icon: "📊" },
    { key: "bot", label: "Bot Nazorati", icon: "🤖" },
    { key: "settings", label: "Admin sozlamalari", icon: "⚙️" },
  ];
  const pageTitle = menuItems.find((m) => m.key === page)?.label || "Dashboard";
  const goto = (key) => { setPage(key); setSidebarOpen(false); setShowShopForm(false); setEditingShopId(null); setSelectedShopId(null); };

  const StatCard = ({ label, value }) => (
    <div style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 14, padding: "14px 16px", flex: 1, minWidth: 0 }}>
      <p style={{ fontSize: 11, color: textSoft, margin: "0 0 6px" }}>{label}</p>
      <p className="qk-mono" style={{ fontSize: 18, fontWeight: 600, color: text, margin: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{value}</p>
    </div>
  );

  const resetShopForm = () => { setF_name(""); setF_productType(""); setF_phone(""); setF_address(""); setF_login(""); setF_pass(""); setF_pass2(""); setEditingShopId(null); setFormError(""); };
  const openNewShopForm = () => { resetShopForm(); setShowShopForm(true); };
  const openEditShopForm = (s) => {
    setEditingShopId(s.id); setF_name(s.name); setF_productType(s.productType || ""); setF_phone(s.phone || ""); setF_address(s.address || "");
    setF_login(s.login || ""); setF_pass(""); setF_pass2(""); setFormError(""); setShowShopForm(true);
  };

  const saveShop = async () => {
    if (!f_name.trim()) return setFormError("Do'kon nomini kiriting");
    if (!f_login.trim()) return setFormError("Login kiriting");
    const loginTaken = shops.some((s) => s.login === f_login.trim() && s.id !== editingShopId) || (data.admin.login === f_login.trim());
    if (loginTaken) return setFormError("Bu login band, boshqasini tanlang");
    if (!editingShopId && !f_pass) return setFormError("Parol kiriting");
    if (f_pass && f_pass !== f_pass2) return setFormError("Parollar mos kelmadi");
    if (f_pass && f_pass.length < 4) return setFormError("Parol kamida 4 belgidan iborat bo'lsin");

    setBusy(true);
    try {
      if (editingShopId) {
        await updateShopInfo(editingShopId, { name: f_name.trim(), productType: f_productType.trim(), phone: f_phone.trim(), address: f_address.trim(), login: f_login.trim(), pass: f_pass });
      } else {
        await createShop({ name: f_name.trim(), productType: f_productType.trim(), phone: f_phone.trim(), address: f_address.trim(), login: f_login.trim(), pass: f_pass });
      }
      await refreshAll();
      setShowShopForm(false); resetShopForm();
    } catch (e) { setFormError("Xatolik yuz berdi, qayta urinib ko'ring"); }
    finally { setBusy(false); }
  };

  const deleteShop = async (id) => { await deleteShopRow(id); await refreshAll(); };

  let body;
  if (page === "dashboard") {
    body = (
      <>
        <div style={{ display: "flex", gap: 10, marginBottom: 18 }}>
          <StatCard label="Jami do'konlar" value={totals.shopCount} />
          <StatCard label="Jami mijozlar" value={totals.customerCount} />
        </div>
        <StatCard label="Umumiy qarzdorlik (barcha do'konlar)" value={fmt(totals.totalDebt)} />
        <div style={{ marginTop: 22, display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <p style={{ fontSize: 13, fontWeight: 700, color: text, margin: 0 }}>Do'konlar</p>
          <button onClick={() => goto("shops")} style={{ background: "none", border: "none", color: accent, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Barchasi →</button>
        </div>
        {shops.length === 0 && <p style={{ fontSize: 13, color: textSoft, textAlign: "center", padding: "16px 0" }}>Hali do'kon qo'shilmagan.</p>}
        {shops.slice(0, 4).map((s) => (
          <div key={s.id} style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "12px 14px", marginBottom: 8 }}>
            <p style={{ fontSize: 14, fontWeight: 600, color: text, margin: "0 0 2px" }}>{s.name}</p>
            <p style={{ fontSize: 12, color: textSoft, margin: 0 }}>{s.address || "Manzil kiritilmagan"}</p>
          </div>
        ))}
      </>
    );
  } else if (page === "shops") {
    if (showShopForm) {
      body = (
        <>
          <p style={{ fontSize: 13, fontWeight: 700, color: accent, margin: "0 0 10px" }}>Ma'lumotlar</p>
          <input className="qk-input" style={inputStyle} value={f_name} onChange={(e) => setF_name(e.target.value)} placeholder="Do'kon nomi" />
          <input className="qk-input" style={inputStyle} value={f_productType} onChange={(e) => setF_productType(e.target.value)} placeholder="Mahsulot turi" />
          <input className="qk-input" style={inputStyle} value={f_phone} onChange={(e) => setF_phone(e.target.value)} placeholder="Telefon raqami" />
          <input className="qk-input" style={inputStyle} value={f_address} onChange={(e) => setF_address(e.target.value)} placeholder="Manzil" />
          <p style={{ fontSize: 13, fontWeight: 700, color: accent, margin: "16px 0 10px" }}>Avtorizatsiya</p>
          <input className="qk-input" style={inputStyle} value={f_login} onChange={(e) => setF_login(e.target.value)} placeholder="Login yaratish" />
          <input className="qk-input" style={inputStyle} type="password" value={f_pass} onChange={(e) => setF_pass(e.target.value)} placeholder={editingShopId ? "Yangi parol (ixtiyoriy)" : "Parol yaratish"} />
          <input className="qk-input" style={{ ...inputStyle, marginBottom: 4 }} type="password" value={f_pass2} onChange={(e) => setF_pass2(e.target.value)} placeholder="Parolni tasdiqlash" />
          {formError && <p style={{ fontSize: 12, color: danger, margin: "8px 0 0" }}>{formError}</p>}
          <div style={{ display: "flex", gap: 8, marginTop: 18 }}>
            <button onClick={() => { setShowShopForm(false); resetShopForm(); }} style={{ flex: 1, padding: "11px 0", borderRadius: 9, border: `1px solid ${glassBorder}`, background: "none", color: textSoft, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>Bekor qilish</button>
            <button disabled={busy} onClick={saveShop} style={{ flex: 1, padding: "11px 0", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", opacity: busy ? 0.6 : 1 }}>{busy ? "Saqlanmoqda..." : "Saqlash"}</button>
          </div>
        </>
      );
    } else {
      body = (
        <>
          {shops.length === 0 ? <p style={{ fontSize: 13, color: textSoft, textAlign: "center", padding: "60px 0 20px" }}>Hali do'kon qo'shilmagan.</p> : shops.map((s) => (
            <div key={s.id} style={{ background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "12px 14px", marginBottom: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div>
                  <p style={{ fontSize: 14, fontWeight: 600, color: text, margin: "0 0 2px" }}>{s.name}</p>
                  {s.productType && <p style={{ fontSize: 12, color: textSoft, margin: "0 0 2px" }}>{s.productType}</p>}
                  <p style={{ fontSize: 12, color: textFaint, margin: 0 }}>{s.address || "Manzil kiritilmagan"}{s.phone ? ` · ${s.phone}` : ""}</p>
                  <p className="qk-mono" style={{ fontSize: 11, color: textFaint, margin: "4px 0 0" }}>login: {s.login}</p>
                </div>
                <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                  <button onClick={() => openEditShopForm(s)} style={{ background: "none", border: `1px solid ${glassBorder}`, borderRadius: 8, color: textSoft, fontSize: 12, padding: "6px 8px", cursor: "pointer" }}>✎</button>
                  <button onClick={() => deleteShop(s.id)} style={{ background: "none", border: "1px solid rgba(224,85,75,0.4)", borderRadius: 8, color: danger, fontSize: 12, padding: "6px 8px", cursor: "pointer" }}>✕</button>
                </div>
              </div>
            </div>
          ))}
          <button onClick={openNewShopForm} style={{ width: "100%", padding: "13px", borderRadius: 10, border: "none", background: accentGrad, color: "#08221E", fontSize: 14, fontWeight: 700, cursor: "pointer", marginTop: 14 }}>+ Do'kon qo'shish</button>
        </>
      );
    }
  } else if (page === "analytics") {
    const selShop = shops.find((s) => s.id === selectedShopId);
    if (!selShop) {
      body = shops.length === 0 ? <p style={{ fontSize: 13, color: textSoft, textAlign: "center", padding: "40px 0" }}>Hali do'kon qo'shilmagan.</p> : (
        <>
          <p style={{ fontSize: 12, color: textSoft, margin: "0 0 12px" }}>Analitikasini ko'rish uchun do'konni tanlang</p>
          {shops.map((s) => (
            <button key={s.id} onClick={() => setSelectedShopId(s.id)} style={{ width: "100%", textAlign: "left", background: glass, backdropFilter: "blur(20px)", WebkitBackdropFilter: "blur(20px)", border: `1px solid ${glassBorder}`, borderRadius: 12, padding: "12px 14px", marginBottom: 8, cursor: "pointer" }}>
              <p style={{ fontSize: 14, fontWeight: 600, color: text, margin: "0 0 2px" }}>{s.name}</p>
              <p style={{ fontSize: 12, color: textSoft, margin: 0 }}>{s.address || "Manzil kiritilmagan"}</p>
            </button>
          ))}
        </>
      );
    } else {
      body = (
        <>
          <button onClick={() => setSelectedShopId(null)} style={{ background: "none", border: "none", color: accent, fontSize: 12, fontWeight: 600, cursor: "pointer", padding: 0, marginBottom: 14 }}>← Do'konlar ro'yxatiga</button>
          <p style={{ fontSize: 16, fontWeight: 700, color: text, margin: "0 0 2px" }}>{selShop.name}</p>
          <p style={{ fontSize: 12, color: textSoft, margin: "0 0 16px" }}>{selShop.address || "Manzil kiritilmagan"}</p>
          <OwnerAnalytics shop={selShop} />
        </>
      );
    }
  } else if (page === "bot") {
    body = <BotControlPage shops={shops} refreshAll={refreshAll} />;
  } else if (page === "settings") {
    body = (
      <>
        <p style={{ fontSize: 13, fontWeight: 700, color: accent, margin: "0 0 10px" }}>Profil login va parolini tahrirlash</p>
        <input className="qk-input" style={inputStyle} value={s_login} onChange={(e) => { setS_login(e.target.value); setS_msg(""); }} placeholder={data.admin.login} />
        <input className="qk-input" style={inputStyle} type="password" value={s_pass} onChange={(e) => { setS_pass(e.target.value); setS_msg(""); }} placeholder="Yangi parol" />
        <input className="qk-input" style={inputStyle} type="password" value={s_pass2} onChange={(e) => { setS_pass2(e.target.value); setS_msg(""); }} placeholder="Yangi parolni tasdiqlash" />
        {s_msg && <p style={{ fontSize: 12, color: s_msg.startsWith("✓") ? accent : danger, margin: "4px 0 0" }}>{s_msg}</p>}
        <button
          onClick={async () => {
            const nextLogin = s_login.trim() || data.admin.login;
            const nextPass = s_pass ? s_pass : data.admin.pass;
            if (s_pass && s_pass !== s_pass2) return setS_msg("Parollar mos kelmadi");
            if (s_pass && s_pass.length < 4) return setS_msg("Parol kamida 4 belgidan iborat bo'lsin");
            try {
              await updateAdminCredentials(data.admin.id, nextLogin, nextPass);
              await refreshAll();
              setS_login(""); setS_pass(""); setS_pass2(""); setS_msg("✓ Saqlandi");
            } catch (e) { setS_msg("Xatolik yuz berdi"); }
          }}
          style={{ width: "100%", padding: "12px", borderRadius: 9, border: "none", background: accentGrad, color: "#08221E", fontSize: 13, fontWeight: 700, cursor: "pointer", marginTop: 10 }}
        >
          Saqlash
        </button>
      </>
    );
  }

  return (
    <>
      <div style={{ display: "flex", alignItems: "center", gap: 12, padding: "16px 18px", borderBottom: `1px solid ${glassBorder}` }}>
        {page !== "dashboard" ? (
          <button onClick={() => goto("dashboard")} style={{ background: "none", border: "none", color: text, fontSize: 18, cursor: "pointer", padding: 0 }}>←</button>
        ) : (
          <button onClick={() => setSidebarOpen(true)} style={{ background: "none", border: "none", color: text, fontSize: 20, cursor: "pointer", padding: 0 }}>☰</button>
        )}
        <div style={{ width: 30, height: 30, borderRadius: 8, background: accentGrad, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 13, color: "#08221E" }}>qk</div>
        <p style={{ fontSize: 15, fontWeight: 700, color: text, margin: 0 }}>{pageTitle}</p>
      </div>

      <div style={{ padding: "18px", overflowY: "auto", flex: 1 }}>{body}</div>

      {sidebarOpen && (
        <div onClick={() => setSidebarOpen(false)} style={{ position: "absolute", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 5 }}>
          <div onClick={(e) => e.stopPropagation()} style={{ position: "absolute", top: 0, left: 0, bottom: 0, width: 250, background: glassPanel, backdropFilter: "blur(28px)", WebkitBackdropFilter: "blur(28px)", borderRight: `1px solid ${glassBorder}`, padding: "20px 14px", zIndex: 6 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 22, padding: "0 6px" }}>
              <div style={{ width: 32, height: 32, borderRadius: 8, background: accentGrad, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 14, color: "#08221E" }}>qk</div>
              <p style={{ fontSize: 14, fontWeight: 700, color: text, margin: 0 }}>q.Kassa.uz</p>
            </div>
            {menuItems.map((m) => (
              <button key={m.key} onClick={() => goto(m.key)} style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, textAlign: "left", padding: "11px 10px", borderRadius: 10, border: "none", marginBottom: 4, cursor: "pointer", background: page === m.key ? "rgba(47,191,158,0.15)" : "none", color: page === m.key ? accent : textSoft, fontSize: 13, fontWeight: page === m.key ? 700 : 500 }}>
                <span style={{ fontSize: 15 }}>{m.icon}</span>{m.label}
              </button>
            ))}
            <div style={{ borderTop: `1px solid ${glassBorder}`, marginTop: 14, paddingTop: 14 }}>
              <button onClick={onLogout} style={{ width: "100%", textAlign: "left", padding: "11px 10px", borderRadius: 10, border: "none", background: "none", color: danger, fontSize: 13, fontWeight: 600, cursor: "pointer" }}>⏻ Chiqish</button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}

// ================= ROOT APP =================

export default function App() {
  const [phase, setPhase] = useState("loading"); // loading | login | admin | owner
  const [data, setData] = useState({ admin: null, shops: [] });
  const [currentShopId, setCurrentShopId] = useState(null);
  const [loadError, setLoadError] = useState(false);

  const [loginVal, setLoginVal] = useState("");
  const [passVal, setPassVal] = useState("");
  const [showPass, setShowPass] = useState(false);
  const [loginError, setLoginError] = useState(false);

  const refreshAll = async () => {
    try {
      const fresh = await loadAll();
      setData(fresh);
      setLoadError(false);
    } catch (e) { setLoadError(true); }
  };

  useEffect(() => {
    (async () => {
      try {
        const fresh = await loadAll();
        setData(fresh);
        setPhase("login");
      } catch (e) {
        setLoadError(true);
        setPhase("login");
      }
    })();
  }, []);

  const tryLogin = () => {
    if (!data.admin) return setLoginError(true);
    if (loginVal === data.admin.login && passVal === data.admin.pass) {
      setLoginError(false); setPhase("admin"); return;
    }
    const shop = data.shops.find((s) => s.login === loginVal && s.pass === passVal);
    if (shop) { setLoginError(false); setCurrentShopId(shop.id); setPhase("owner"); return; }
    setLoginError(true);
  };

  const handleLogout = () => { setPhase("login"); setLoginVal(""); setPassVal(""); setCurrentShopId(null); };

  const currentShop = data.shops.find((s) => s.id === currentShopId);

  return (
    <div className="qk-root qk-shell" style={{ background: bgGrad, position: "relative", overflow: "hidden", display: "flex", flexDirection: "column" }}>
      <FontLoad />

      {phase === "loading" && (
        <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>
          <p style={{ color: textSoft, fontSize: 13 }}>Yuklanmoqda...</p>
        </div>
      )}

      {phase === "login" && (
        <div style={{ flex: 1, display: "flex", alignItems: "center", justifyContent: "center", padding: "32px 20px", position: "relative", overflow: "hidden" }}>
          <div style={{ position: "absolute", top: -60, right: -60, width: 200, height: 200, borderRadius: "50%", background: "radial-gradient(circle, rgba(47,191,158,0.25), transparent 70%)" }} />
          <div style={{ width: "100%", background: glass, border: `1px solid ${glassBorder}`, borderRadius: 18, padding: "32px 26px", backdropFilter: "blur(24px)", WebkitBackdropFilter: "blur(24px)", position: "relative" }}>
            <div style={{ textAlign: "center", marginBottom: 22 }}>
              <div style={{ width: 52, height: 52, borderRadius: 14, background: accentGrad, margin: "0 auto 14px", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 20, color: "#08221E" }}>qk</div>
              <p style={{ fontSize: 19, fontWeight: 700, color: text, margin: "0 0 2px" }}>Xush kelibsiz!</p>
              <p style={{ fontSize: 13, color: textSoft, margin: 0 }}>q.Kassa.uz</p>
            </div>
            {loadError && <p style={{ fontSize: 12, color: danger, textAlign: "center", margin: "0 0 14px" }}>Bazaga ulanishda xatolik. Internetni tekshiring.</p>}
            <label style={{ fontSize: 12, color: textSoft, display: "block", marginBottom: 6 }}>Login</label>
            <input className="qk-input" value={loginVal} onChange={(e) => { setLoginVal(e.target.value); setLoginError(false); }} placeholder="login" style={{ ...inputStyle, borderColor: loginError ? danger : glassBorder, marginBottom: 14 }} />
            <label style={{ fontSize: 12, color: textSoft, display: "block", marginBottom: 6 }}>Parol</label>
            <div style={{ position: "relative", marginBottom: loginError ? 10 : 22 }}>
              <input className="qk-input" type={showPass ? "text" : "password"} value={passVal} onChange={(e) => { setPassVal(e.target.value); setLoginError(false); }} onKeyDown={(e) => e.key === "Enter" && tryLogin()} placeholder="parol" style={{ ...inputStyle, borderColor: loginError ? danger : glassBorder, padding: "10px 40px 10px 12px", marginBottom: 0 }} />
              <button onClick={() => setShowPass(!showPass)} style={{ position: "absolute", right: 10, top: "50%", transform: "translateY(-50%)", background: "none", border: "none", color: textFaint, cursor: "pointer", fontSize: 12 }}>{showPass ? "yashirish" : "ko'rish"}</button>
            </div>
            {loginError && <p style={{ fontSize: 12, color: danger, margin: "0 0 12px" }}>Login yoki parol noto'g'ri</p>}
            <button onClick={tryLogin} style={{ width: "100%", padding: "13px", borderRadius: 10, border: "none", background: accentGrad, color: "#08221E", fontSize: 14, fontWeight: 800, letterSpacing: 0.5, cursor: "pointer", textTransform: "uppercase" }}>Tizimga kirish →</button>
            <p style={{ textAlign: "center", fontSize: 11, color: textFaint, marginTop: 22, marginBottom: 0 }}>created by - imradjabov</p>
          </div>
        </div>
      )}

      {phase === "admin" && <AdminPanel data={data} refreshAll={refreshAll} onLogout={handleLogout} />}
      {phase === "owner" && currentShop && <OwnerPanel shop={currentShop} refreshAll={refreshAll} onLogout={handleLogout} />}
    </div>
  );
}
