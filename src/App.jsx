import { useState, useEffect } from "react";
import "./App.css";
import { generateDocx } from "./docxGenerator";

// ======= Supabase =======
const SUPABASE_URL = "https://bpkslzciqpbzmkktrymm.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJwa3NsemNpcXBiem1ra3RyeW1tIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzgwMzM4NTEsImV4cCI6MjA5MzYwOTg1MX0.aAwwF0nTGbp2cSa04f5WjGnqH_GHNsymV7CQ6ue-r4E";

const sbHeaders = {
  "apikey": SUPABASE_KEY,
  "Authorization": `Bearer ${SUPABASE_KEY}`,
  "Content-Type": "application/json",
};

async function dbLoadAll() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/nibondhon_data?select=date,rows&order=date.asc`, { headers: sbHeaders });
    if (!res.ok) return {};
    const data = await res.json();
    const obj = {};
    for (const r of data) obj[r.date] = r.rows;
    return obj;
  } catch { return {}; }
}

async function dbSave(date, rows) {
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/nibondhon_data`, {
      method: "POST",
      headers: { ...sbHeaders, "Prefer": "resolution=merge-duplicates" },
      body: JSON.stringify({ date, rows, updated_at: new Date().toISOString() }),
    });
  } catch {}
}

async function dbDelete(date) {
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/nibondhon_data?date=eq.${date}`, {
      method: "DELETE", headers: sbHeaders,
    });
  } catch {}
}

async function dbDeleteAll() {
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/nibondhon_data?id=gte.0`, {
      method: "DELETE", headers: sbHeaders,
    });
  } catch {}
}

// ======= Constants =======
const UNIONS = [
  { name: "উথলী ইউনিয়ন", jT: 30, mT: 13 },
  { name: "আন্দুলবাড়ীয়া ইউনিয়ন", jT: 84, mT: 19 },
  { name: "বাঁকা ইউনিয়ন", jT: 30, mT: 13 },
  { name: "সীমান্ত ইউনিয়ন", jT: 45, mT: 22 },
  { name: "হাসাদাহ ইউনিয়ন", jT: 31, mT: 10 },
  { name: "রায়পুর ইউনিয়ন", jT: 23, mT: 9 },
  { name: "মনোহরপুর ইউনিয়ন", jT: 26, mT: 1 },
  { name: "কেভিডে ইউনিয়ন", jT: 29, mT: 1 },
  { name: "জীবননগর পৌরসভা", jT: 84, mT: 19 },
];

const SMARAK = "স্মারক নং-০৫.৪৪.১৮৫৫.০০০.১৬.০০৩.২৫";

// ======= Helpers =======
const BN_DIGITS = { 0:"০",1:"১",2:"২",3:"৩",4:"৪",5:"৫",6:"৬",7:"৭",8:"৮",9:"৯" };
const toBangla = (n) => String(n).replace(/[0-9]/g, (d) => BN_DIGITS[d]);
const toEnglish = (s) => String(s).replace(/[০-৯]/g, (d) => "০১২৩৪৫৬৭৮৯".indexOf(d));
const toNum = (s) => parseInt(toEnglish(s || "0")) || 0;

const BN_MONTHS = ["জানুয়ারি","ফেব্রুয়ারি","মার্চ","এপ্রিল","মে","জুন","জুলাই","আগস্ট","সেপ্টেম্বর","অক্টোবর","নভেম্বর","ডিসেম্বর"];
const BN_BN_MONTHS = ["বৈশাখ","জ্যৈষ্ঠ","আষাঢ়","শ্রাবণ","ভাদ্র","আশ্বিন","কার্তিক","অগ্রহায়ণ","পৌষ","মাঘ","ফাল্গুন","চৈত্র"];

function getMasYear(dateStr) {
  if (!dateStr) return "";
  const [year, month] = dateStr.split("-").map(Number);
  return `${BN_MONTHS[month - 1]}/${toBangla(year % 100)}`;
}

const BN_MONTH_MAP = [
  [4,14,0],[5,15,1],[6,15,2],[7,16,3],[8,17,4],
  [9,17,5],[10,17,6],[11,16,7],[12,15,8],
  [1,14,9],[2,13,10],[3,14,11],
];

function toBanglaDate(dateStr) {
  if (!dateStr) return { bangla: "", english: "" };
  // new Date() ব্যবহার করলে timezone shift হয়, তাই manually parse করি
  const [year, month, day] = dateStr.split("-").map(Number);
  let bnMonth = 0, bnDay = 0, bnYear = 0;
  for (const [m, startDay, bMonth] of BN_MONTH_MAP) {
    if (month === m) {
      bnMonth = bMonth;
      bnDay = day - startDay + 1;
      if (bnDay <= 0) { bnMonth = (bMonth - 1 + 12) % 12; bnDay = day + (30 - startDay) + 1; }
      break;
    }
  }
  bnYear = month >= 4 ? year - 593 : year - 594;
  return {
    bangla: `${toBangla(bnDay)} ${BN_BN_MONTHS[bnMonth]}, ${toBangla(bnYear)}`,
    english: `${toBangla(day)} ${BN_MONTHS[month - 1]}, ${toBangla(year)}`,
  };
}

// এ পর্যন্ত = আগের দিনের এ পর্যন্ত + আজকের আজ
function calcCum(allData, selectedDate, rowIdx, todayField, cumField) {
  const allDates = Object.keys(allData).sort();
  const idx = allDates.indexOf(selectedDate);
  const todayRows = allData[selectedDate];
  const todayVal = todayRows?.[rowIdx] ? toNum(todayRows[rowIdx][todayField] || "0") : 0;
  let prevCum = 0;
  if (idx > 0) {
    for (let i = idx - 1; i >= 0; i--) {
      const prevRows = allData[allDates[i]];
      if (prevRows?.[rowIdx]) {
        const stored = prevRows[rowIdx][cumField];
        if (stored !== "" && stored !== undefined && stored !== null) {
          prevCum = toNum(stored); break;
        } else {
          prevCum = calcCum(allData, allDates[i], rowIdx, todayField, cumField); break;
        }
      }
    }
  }
  return prevCum + todayVal;
}

const EMPTY_ROW = () => ({ j45Today:"", j45Cum:"", j46Today:"", j46Cum:"", m45Today:"", m45Cum:"", m46Today:"", m46Cum:"" });

// ======= App =======
export default function App() {
  const [view, setView] = useState("entry");
  const [selectedDate, setSelectedDate] = useState("");
  const [smarakEnd, setSmarakEnd] = useState("১৪২");
  const [allData, setAllData] = useState({});
  const [savedDates, setSavedDates] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  useEffect(() => {
    dbLoadAll().then(data => { setAllData(data); setLoading(false); });
  }, []);

  useEffect(() => {
    setSavedDates(Object.keys(allData).sort().reverse());
  }, [allData]);

  const dateInfo = toBanglaDate(selectedDate);
  const masYear = getMasYear(selectedDate);
  const currentRows = allData[selectedDate] || UNIONS.map(EMPTY_ROW);

  const getCum = (rowIdx, todayField, cumField) => {
    const r = currentRows[rowIdx] || {};
    if (r[cumField] !== "" && r[cumField] !== undefined && r[cumField] !== null) return r[cumField];
    const auto = calcCum(allData, selectedDate, rowIdx, todayField, cumField);
    return auto > 0 ? toBangla(auto) : "";
  };

  const updateCell = async (rowIdx, field, val) => {
    const rows = currentRows.map((r, i) => i === rowIdx ? { ...r, [field]: val } : r);
    setAllData(prev => ({ ...prev, [selectedDate]: rows }));
    setSaving(true);
    await dbSave(selectedDate, rows);
    setSaving(false);
  };

  const sumField = (field) => currentRows.reduce((acc, r) => acc + toNum(r[field] || "0"), 0);

  const sumCum = (todayField, cumField) => {
    let total = 0;
    for (let i = 0; i < UNIONS.length; i++) {
      const r = currentRows[i] || {};
      if (r[cumField] !== "" && r[cumField] !== undefined && r[cumField] !== null) total += toNum(r[cumField]);
      else total += calcCum(allData, selectedDate, i, todayField, cumField);
    }
    return total;
  };

  const deleteDate = async (date) => {
    if (!window.confirm("এই তারিখের সব তথ্য মুছে ফেলবেন?")) return;
    const newData = { ...allData };
    delete newData[date];
    setAllData(newData);
    await dbDelete(date);
    if (date === selectedDate) setSelectedDate("");
  };

  const deleteAll = async () => {
    if (!window.confirm("সব তারিখের সব তথ্য মুছে ফেলবেন?")) return;
    setAllData({});
    await dbDeleteAll();
    setSelectedDate("");
  };

  const smarak = `${SMARAK}-${smarakEnd}`;
  const letterBody = `উপর্যুক্ত বিষয় ও সূত্রোক্ত স্মারকের পরিপ্রেক্ষিতে জীবননগর উপজেলার ০৮ (আট)টি ইউনিয়ন পরিষদের ও ০১ (এক)টি পৌরসভার ${masYear} মাসের কার্যদিবস অনুযায়ী ${dateInfo.english} তারিখের জন্ম ও মৃত্যু নিবন্ধন লক্ষমাত্রা অনুযায়ী তথ্য নির্ধারিত ছক মোতাবেক প্রস্তুতপূর্বক মহোদয়ের সদয় অবগতি ও প্রয়োজনীয় ব্যবস্থা গ্রহণের জন্য প্রেরণ করা হলো।`;

  if (loading) return <div className="loading-screen">ডেটা লোড হচ্ছে...</div>;

  return (
    <div className="app">
      {/* টপ নেভ */}
      <div className="top-nav">
        <div className="nav-title">নিবন্ধন তথ্য জেনারেটর — জীবননগর</div>
        <div className="nav-btns">
          {saving && <span className="saving-indicator">💾 সংরক্ষণ হচ্ছে...</span>}
          <button className={`nav-btn ${view === "entry" ? "active" : ""}`} onClick={() => setView("entry")}>তথ্য প্রবেশ</button>
          <button className={`nav-btn ${view === "preview" ? "active" : ""}`} onClick={() => setView("preview")}>প্রিভিউ ও ডাউনলোড</button>
        </div>
      </div>

      {view === "entry" && (
        <div className="entry-view">
          <div className="section-card">
            <div className="section-label">স্মারক ও তারিখ</div>
            <div className="fields-row">
              <div className="field-group">
                <label className="field-label">তারিখ নির্বাচন করুন</label>
                <input type="date" className="text-input" value={selectedDate} onChange={e => setSelectedDate(e.target.value)} />
              </div>
              <div className="field-group">
                <label className="field-label">বাংলা তারিখ (স্বয়ংক্রিয়)</label>
                <div className="auto-field">{dateInfo.bangla || "—"}</div>
              </div>
              <div className="field-group">
                <label className="field-label">ইংরেজি তারিখ (স্বয়ংক্রিয়)</label>
                <div className="auto-field">{dateInfo.english || "—"}</div>
              </div>
              <div className="field-group">
                <label className="field-label">মাস/বছর (স্বয়ংক্রিয়)</label>
                <div className="auto-field">{masYear || "—"}</div>
              </div>
              <div className="field-group">
                <label className="field-label">স্মারকের শেষাংশ</label>
                <input type="text" className="text-input" value={smarakEnd} onChange={e => setSmarakEnd(e.target.value)} placeholder="১৪২" />
              </div>
            </div>

            {savedDates.length > 0 && (
              <div className="field-group" style={{ marginTop: "12px" }}>
                <label className="field-label">সংরক্ষিত তারিখ দেখুন</label>
                <div className="saved-dates-row">
                  <select className="text-input" value={selectedDate} onChange={e => setSelectedDate(e.target.value)}>
                    <option value="">— তারিখ বেছে নিন —</option>
                    {savedDates.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                  {selectedDate && <button className="del-btn" onClick={() => deleteDate(selectedDate)}>এই দিনের তথ্য মুছুন</button>}
                  <button className="del-btn danger" onClick={deleteAll}>সব মুছুন</button>
                </div>
              </div>
            )}
          </div>

          {/* এন্ট্রি টেবিল */}
          <div className="section-card">
            <div className="section-label">
              তথ্য প্রবেশ &nbsp;|&nbsp;
              <span className="blue">আজ</span> = আজকের সংখ্যা &nbsp;|&nbsp;
              <span className="green">এ পর্যন্ত</span> = স্বয়ংক্রিয় (নিজে লিখলে override হবে)
            </div>
            <div className="table-wrapper">
              <table className="entry-table">
                <thead>
                  <tr>
                    <th rowSpan={3} className="union-th">ইউনিয়নের<br />নাম</th>
                    <th colSpan={6} className="janmo-head">জন্ম নিবন্ধনের তথ্য</th>
                    <th colSpan={6} className="mrityu-head">মৃত্যু নিবন্ধনের তথ্য</th>
                  </tr>
                  <tr>
                    <th rowSpan={2} className="target-th">মাসিক<br />টার্গেট</th>
                    <th colSpan={2}>৪৫ দিনের মধ্যে</th>
                    <th colSpan={2}>৪৬-৩৬৫ দিনের মধ্যে</th>
                    <th rowSpan={2} className="birth-total-th">মোট</th>
                    <th rowSpan={2} className="target-th">মাসিক<br />টার্গেট</th>
                    <th colSpan={2}>৪৫ দিনের মধ্যে</th>
                    <th colSpan={2}>৪৬-৩৬৫ দিনের মধ্যে</th>
                    <th rowSpan={2} className="birth-total-th">মোট</th>
                  </tr>
                  <tr>
                    <th className="today-th">আজ</th><th className="cum-th">এ পর্যন্ত</th>
                    <th className="today-th">আজ</th><th className="cum-th">এ পর্যন্ত</th>
                    <th className="today-th">আজ</th><th className="cum-th">এ পর্যন্ত</th>
                    <th className="today-th">আজ</th><th className="cum-th">এ পর্যন্ত</th>
                  </tr>
                </thead>
                <tbody>
                  {UNIONS.map((union, i) => {
                    const r = currentRows[i] || {};
                    const j45Cum = getCum(i,"j45Today","j45Cum");
                    const j46Cum = getCum(i,"j46Today","j46Cum");
                    const m45Cum = getCum(i,"m45Today","m45Cum");
                    const m46Cum = getCum(i,"m46Today","m46Cum");
                    const jTotal = toNum(j45Cum) + toNum(j46Cum);
                    const mTotal = toNum(m45Cum) + toNum(m46Cum);
                    return (
                      <tr key={i}>
                        <td className="union-name">{union.name}</td>
                        <td className="target-cell">{toBangla(union.jT)}</td>
                        <td><input className="num-input blue-input" value={r.j45Today||""} onChange={e=>updateCell(i,"j45Today",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input green-input" value={j45Cum} onChange={e=>updateCell(i,"j45Cum",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input blue-input" value={r.j46Today||""} onChange={e=>updateCell(i,"j46Today",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input green-input" value={j46Cum} onChange={e=>updateCell(i,"j46Cum",e.target.value)} disabled={!selectedDate}/></td>
                        <td className="total-cell">{jTotal>0?toBangla(jTotal):""}</td>
                        <td className="target-cell">{toBangla(union.mT)}</td>
                        <td><input className="num-input blue-input" value={r.m45Today||""} onChange={e=>updateCell(i,"m45Today",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input green-input" value={m45Cum} onChange={e=>updateCell(i,"m45Cum",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input blue-input" value={r.m46Today||""} onChange={e=>updateCell(i,"m46Today",e.target.value)} disabled={!selectedDate}/></td>
                        <td><input className="num-input green-input" value={m46Cum} onChange={e=>updateCell(i,"m46Cum",e.target.value)} disabled={!selectedDate}/></td>
                        <td className="total-cell">{mTotal>0?toBangla(mTotal):""}</td>
                      </tr>
                    );
                  })}
                  <tr className="total-row">
                    <td><strong>সর্বমোট</strong></td>
                    <td>{toBangla(UNIONS.reduce((a,u)=>a+u.jT,0))}</td>
                    <td>{toBangla(sumField("j45Today"))}</td>
                    <td>{toBangla(sumCum("j45Today","j45Cum"))}</td>
                    <td>{toBangla(sumField("j46Today"))}</td>
                    <td>{toBangla(sumCum("j46Today","j46Cum"))}</td>
                    <td>{toBangla(sumCum("j45Today","j45Cum")+sumCum("j46Today","j46Cum"))}</td>
                    <td>{toBangla(UNIONS.reduce((a,u)=>a+u.mT,0))}</td>
                    <td>{toBangla(sumField("m45Today"))}</td>
                    <td>{toBangla(sumCum("m45Today","m45Cum"))}</td>
                    <td>{toBangla(sumField("m46Today"))}</td>
                    <td>{toBangla(sumCum("m46Today","m46Cum"))}</td>
                    <td>{toBangla(sumCum("m45Today","m45Cum")+sumCum("m46Today","m46Cum"))}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            <div className="action-row">
              <button className="btn-primary" onClick={() => setView("preview")}>প্রিভিউ দেখুন →</button>
            </div>
          </div>
        </div>
      )}

      {view === "preview" && (
        <div className="preview-view">
          <div className="preview-actions">
            <button className="btn-secondary" onClick={() => setView("entry")}>← এডিট করুন</button>
            <button className="btn-primary" onClick={async () => {
              const { default: jsPDF } = await import('jspdf');
              const { default: html2canvas } = await import('html2canvas');
              const el = document.getElementById('printable');
              const canvas = await html2canvas(el, {
                scale: 2,
                useCORS: true,
                logging: false,
                width: el.scrollWidth,
                height: el.scrollHeight,
              });
              const imgData = canvas.toDataURL('image/png');
              const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
              const pdfW = pdf.internal.pageSize.getWidth();
              const pdfH = (canvas.height * pdfW) / canvas.width;
              pdf.addImage(imgData, 'PNG', 0, 0, pdfW, pdfH);
              pdf.save('nibondhon-report.pdf');
            }}>📥 PDF ডাউনলোড</button>
            <button className="btn-success" onClick={() => generateDocx({ UNIONS, currentRows, smarak, dateInfo, masYear, letterBody, toBangla, sumField })}>
              📄 Word (.docx) ডাউনলোড
            </button>
          </div>

          <div className="doc-preview" id="printable">
            <div className="doc-header">গণপ্রজাতন্ত্রী বাংলাদেশ সরকার</div>
            <div className="doc-sub">উপজেলা নির্বাহী অফিসারের কার্যালয়</div>
            <div className="doc-sub">জীবননগর, চুয়াডাঙ্গা</div>
            <div className="doc-sub small">www.jibannagar.chuadanga.gov.bd</div>

            <div className="doc-smarak-row">
              <div>{smarak}</div>
              <div className="doc-date-block">
                <div className="doc-date-bangla">তারিখ: {dateInfo.bangla}</div>
                <div className="doc-date-english">{dateInfo.english}</div>
              </div>
            </div>

            <p className="doc-bishoy"><strong>বিষয়:</strong> জন্ম ও মৃত্যু নিবন্ধনের লক্ষমাত্রা অনুযায়ী প্রতিদিনের তথ্য প্রেরণ</p>
            <p className="doc-sutro"><strong>সূত্র :</strong> জেলা প্রশাসক, চুয়াডাঙ্গা মহোদয়ের ২১/০৯/২০২২ তারিখের ০৫.৪৪.১৮০০.১০৬.১৪.০০৮.২২-৫০৩(৮২) নং স্মারক</p>
            <p className="doc-body">{letterBody}</p>
            <p className="doc-tbl-title"><strong>উপজেলার নাম: জীবননগর</strong></p>

            <table className="doc-table">
              <thead>
                <tr>
                  <th rowSpan={3} className="doc-th-union">ইউনিয়নের<br />নাম</th>
                  <th colSpan={6}>জন্ম নিবন্ধনের তথ্য</th>
                  <th colSpan={6}>মৃত্যু নিবন্ধনের তথ্য</th>
                </tr>
                <tr>
                  <th rowSpan={2}>মাসিক<br />টার্গেট</th>
                  <th colSpan={2}>৪৫ দিনের মধ্যে</th>
                  <th colSpan={2}>৪৬-৩৬৫ দিনের মধ্যে</th>
                  <th rowSpan={2}>মোট</th>
                  <th rowSpan={2}>মাসিক<br />টার্গেট</th>
                  <th colSpan={2}>৪৫ দিনের মধ্যে</th>
                  <th colSpan={2}>৪৬-৩৬৫ দিনের মধ্যে</th>
                  <th rowSpan={2}>মোট</th>
                </tr>
                <tr>
                  <th>আজ</th><th>এ পর্যন্ত</th>
                  <th>আজ</th><th>এ পর্যন্ত</th>
                  <th>আজ</th><th>এ পর্যন্ত</th>
                  <th>আজ</th><th>এ পর্যন্ত</th>
                </tr>
              </thead>
              <tbody>
                {UNIONS.map((union, i) => {
                  const r = currentRows[i] || {};
                  const j45Cum = getCum(i,"j45Today","j45Cum");
                  const j46Cum = getCum(i,"j46Today","j46Cum");
                  const m45Cum = getCum(i,"m45Today","m45Cum");
                  const m46Cum = getCum(i,"m46Today","m46Cum");
                  const jTotal = toNum(j45Cum)+toNum(j46Cum);
                  const mTotal = toNum(m45Cum)+toNum(m46Cum);
                  return (
                    <tr key={i}>
                      <td className="doc-union-name">{union.name}</td>
                      <td>{toBangla(union.jT)}</td>
                      <td>{r.j45Today||""}</td>
                      <td>{j45Cum}</td>
                      <td>{r.j46Today||""}</td>
                      <td>{j46Cum}</td>
                      <td>{jTotal>0?toBangla(jTotal):""}</td>
                      <td>{toBangla(union.mT)}</td>
                      <td>{r.m45Today||""}</td>
                      <td>{m45Cum}</td>
                      <td>{r.m46Today||""}</td>
                      <td>{m46Cum}</td>
                      <td>{mTotal>0?toBangla(mTotal):""}</td>
                    </tr>
                  );
                })}
                <tr className="doc-total-row">
                  <td><strong>সর্বমোট</strong></td>
                  <td>{toBangla(UNIONS.reduce((a,u)=>a+u.jT,0))}</td>
                  <td>{toBangla(sumField("j45Today"))}</td>
                  <td>{toBangla(sumCum("j45Today","j45Cum"))}</td>
                  <td>{toBangla(sumField("j46Today"))}</td>
                  <td>{toBangla(sumCum("j46Today","j46Cum"))}</td>
                  <td>{toBangla(sumCum("j45Today","j45Cum")+sumCum("j46Today","j46Cum"))}</td>
                  <td>{toBangla(UNIONS.reduce((a,u)=>a+u.mT,0))}</td>
                  <td>{toBangla(sumField("m45Today"))}</td>
                  <td>{toBangla(sumCum("m45Today","m45Cum"))}</td>
                  <td>{toBangla(sumField("m46Today"))}</td>
                  <td>{toBangla(sumCum("m46Today","m46Cum"))}</td>
                  <td>{toBangla(sumCum("m45Today","m45Cum")+sumCum("m46Today","m46Cum"))}</td>
                </tr>
              </tbody>
            </table>

            {/* স্বাক্ষর */}
            <div className="sig-section">
              <div className="sig-top-row">
                <div className="sig-right-block">
                  <div className="sig-name-spacer"></div>
                  <div className="sig-name">মোঃ আশরাফুল আলম রাসেল</div>
                  <div>উপজেলা নির্বাহী অফিসার</div>
                  <div>জীবননগর, চুয়াডাঙ্গা।</div>
                  <div>☎ ০২৪৭৭৭৮৯৭০১</div>
                  <div>unojibannagar@mopa.gov.bd</div>
                </div>
              </div>
              <div className="sig-bottom-row">
                <div className="sig-left-block">
                  <div>জেলা প্রশাসক</div>
                  <div>চুয়াডাঙ্গা।</div>
                  <div>দুঃআঃ উপ-পরিচালক, স্থানীয় সরকার, চুয়াডাঙ্গা।</div>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
