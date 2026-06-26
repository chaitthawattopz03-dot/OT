import { useState, useEffect, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, Legend,
} from "recharts";
import {
  Upload, Search, CheckSquare, Square, Download, FileSpreadsheet, Trash2,
  Pencil, Plus, X, ChevronUp, ChevronDown, Users, Clock, TrendingUp,
  TrendingDown, CalendarDays, AlertCircle, Check, Loader2,
} from "lucide-react";

const THAI_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
  "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];

const BLUE = ["#1d4ed8","#2563eb","#3b82f6","#60a5fa","#93c5fd","#0ea5e9","#0284c7","#1e40af"];

const DEFAULT_ROSTER = [
  "สมพร พรมมา","ประสาน ลาภเหลือ","สมยา เต็งศิริ","กฤษดา พรมศร","รัตนะ เส็งดี",
  "สมเกียรติ จงกลาง","ฐานลินต์ วินิจธนษรณ์","สมศักดิ์ วงษ์ประทัด","วันพิชิต สุจริตปฎิภาณ",
  "ณชานนท์ คำหวาน","บัณฑิต เดชขุนทด","ภควัต สิงหล","ขวัญชัย พวงวงษา","พิสิฏฐ์ สุวรรณสิทธิ์",
  "ชูเกียรติ สีแก้ว","ชินกร เต็งศิริ","วัชรากรณ์ ดิลกโกมล","ดลพักร เสาหงษ์","ปรัชญา รวยระรื่น",
  "กรีทาพล มีทาทอง","กฤษดา เสมาทอง","ธวัชชัย แป้นจันทร์","ณฐกรณ์ จันทาบุตร","ธรรมนูญ บุตรเลียบ",
  "ประติวัติ ศรีรักษา","สมศักดิ์ หวังกุลกลาง","ศุภชัย แสงกาวิน","ฉัตรชัย อดิสรณกุล","สหรัฐ ชื่นสว่าง",
  "สฤษดิ์ ภู่ห้อย","เธียรธรรม ประทุมวงษ์","สุเมธ พยับเดช","ฐนพัฒน์ บุญเถื่อน","ปภัสสิริย์ พรหมมินทร์",
  "จักรพันธ์ ภู่ขาว","นพณัฐ มณีรัตน์","ศิวกร ทิพย์สิงห์",
].map((name, i) => ({ id: `e${String(i + 1).padStart(2, "0")}`, name }));

function parseCSVFile(file) {
  return new Promise((resolve, reject) => {
    Papa.parse(file, {
      header: false,
      skipEmptyLines: false,
      complete: (res) => resolve(res.data),
      error: (err) => reject(err),
    });
  });
}

function toNum(v) {
  if (v === undefined || v === null) return 0;
  const n = parseFloat(String(v).trim().replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function processRows(rows, fileName) {
  let dateRangeStr = "";
  for (const r of rows) {
    if (r && r[1] && String(r[1]).trim() === "วันที่เลือก") {
      dateRangeStr = String(r[2] || "").trim();
      break;
    }
  }

  let headerIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    if (r.includes("ชื่อพนักงาน") && r.includes("วันที่") && r.includes("รหัสพนักงาน")) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) return { error: "ไม่พบข้อมูล OT" };

  const header = rows[headerIdx].map((h) => String(h || "").trim());
  const nameIdx = header.indexOf("ชื่อพนักงาน");
  const deptIdx = header.indexOf("ชื่อตำแหน่งงาน");
  const otIdx = [];
  header.forEach((h, i) => {
    if (h.startsWith("O.T.") && h.includes("ชั่วโมง")) otIdx.push(i);
  });
  if (otIdx.length === 0) return { error: "ไม่พบข้อมูล OT" };

  const dataRows = rows
    .slice(headerIdx + 1)
    .filter((r) => r && r.length > nameIdx && String(r[nameIdx] || "").trim() !== "");

  if (dataRows.length === 0) return { error: "No Data" };

  const employees = {};
  for (const r of dataRows) {
    const name = String(r[nameIdx]).trim();
    const dept = deptIdx >= 0 ? String(r[deptIdx] || "").trim() : "";
    let ot = 0;
    for (const idx of otIdx) ot += toNum(r[idx]);
    if (!employees[name]) employees[name] = { ot: 0, dept: dept, days: 0 };
    employees[name].ot += ot;
    employees[name].days += 1;
    if (dept) employees[name].dept = dept;
  }

  let monthKey, label;
  const m1 = dateRangeStr.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m1) {
    const mm = m1[2].padStart(2, "0");
    const yyyy = m1[3];
    monthKey = `${yyyy}-${mm}`;
    label = `${THAI_MONTHS[parseInt(mm, 10) - 1]} ${yyyy}`;
  } else {
    const dateIdx = header.indexOf("วันที่");
    const firstDate = dateIdx >= 0 ? String(dataRows[0][dateIdx] || "") : "";
    const m2 = firstDate.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m2) {
      const mm = m2[2].padStart(2, "0");
      const yyyy = m2[3];
      monthKey = `${yyyy}-${mm}`;
      label = `${THAI_MONTHS[parseInt(mm, 10) - 1]} ${yyyy}`;
    } else {
      monthKey = `unknown-${Date.now()}`;
      label = "ไม่ทราบเดือน";
    }
  }

  return {
    monthKey, label, dateRangeStr, employees, fileName,
    rowCount: dataRows.length,
    empCount: Object.keys(employees).length,
  };
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

export default function OTDashboard() {
  const [roster, setRoster] = useState([]);
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [monthsData, setMonthsData] = useState({});
  const [selectedMonths, setSelectedMonths] = useState(new Set());
  const [loading, setLoading] = useState(true);
  const [uploading, setUploading] = useState(false);
  const [toast, setToast] = useState(null);
  const [rosterSearch, setRosterSearch] = useState("");
  const [tableSearch, setTableSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("ทั้งหมด");
  const [otMin, setOtMin] = useState("");
  const [otMax, setOtMax] = useState("");
  const [sortAsc, setSortAsc] = useState(false);
  const [newName, setNewName] = useState("");
  const [editingId, setEditingId] = useState(null);
  const [editingValue, setEditingValue] = useState("");
  const [dragOver, setDragOver] = useState(false);
  const [page, setPage] = useState(1);
  const fileInputRef = useRef(null);
  const pageSize = 15;

  useEffect(() => {
    (async () => {
      let r = null;
      try {
        const res = await window.storage.get("roster", false);
        r = res ? JSON.parse(res.value) : null;
      } catch (e) { r = null; }
      if (!r || !r.length) {
        r = DEFAULT_ROSTER;
        try { await window.storage.set("roster", JSON.stringify(r), false); } catch (e) {}
      }
      setRoster(r);
      setSelectedIds(new Set(r.map((e) => e.id)));

      let idx = [];
      try {
        const res = await window.storage.get("months-index", false);
        idx = res ? JSON.parse(res.value) : [];
      } catch (e) { idx = []; }

      const md = {};
      for (const key of idx) {
        try {
          const res = await window.storage.get("month:" + key, false);
          if (res) md[key] = JSON.parse(res.value);
        } catch (e) {}
      }
      setMonthsData(md);
      setSelectedMonths(new Set(Object.keys(md)));
      setLoading(false);
    })();
  }, []);

  useEffect(() => { setPage(1); }, [tableSearch, deptFilter, otMin, otMax, sortAsc, selectedMonths, selectedIds]);

  const showToast = (msg, kind = "success") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 4000);
  };

  const persistRoster = async (next) => {
    setRoster(next);
    try { await window.storage.set("roster", JSON.stringify(next), false); } catch (e) {}
  };

  const handleFiles = async (files) => {
    const list = Array.from(files || []).filter((f) => f.name.toLowerCase().endsWith(".csv"));
    if (!list.length) {
      showToast("กรุณาเลือกไฟล์ .csv จากระบบ HR เท่านั้น", "error");
      return;
    }
    setUploading(true);
    for (const file of list) {
      try {
        const rows = await parseCSVFile(file);
        const result = processRows(rows, file.name);
        if (result.error) {
          showToast(`${file.name}: ${result.error}`, "error");
          continue;
        }
        const isUpdate = Object.prototype.hasOwnProperty.call(monthsData, result.monthKey);
        const payload = {
          label: result.label,
          dateRangeStr: result.dateRangeStr,
          fileName: result.fileName,
          uploadedAt: Date.now(),
          employees: result.employees,
        };
        await window.storage.set("month:" + result.monthKey, JSON.stringify(payload), false);

        let idx = [];
        try {
          const res = await window.storage.get("months-index", false);
          idx = res ? JSON.parse(res.value) : [];
        } catch (e) { idx = []; }
        if (!idx.includes(result.monthKey)) {
          idx.push(result.monthKey);
          await window.storage.set("months-index", JSON.stringify(idx), false);
        }

        setMonthsData((prev) => ({ ...prev, [result.monthKey]: payload }));
        setSelectedMonths((prev) => new Set([...prev, result.monthKey]));
        showToast(
          `${isUpdate ? "เขียนทับข้อมูลเดือนเดิม" : "เพิ่มข้อมูลเดือนใหม่"}: ${result.label} • พบ ${result.empCount} คน (${result.rowCount} แถว)`,
          "success"
        );
      } catch (e) {
        showToast(`${file.name}: เกิดข้อผิดพลาดในการอ่านไฟล์`, "error");
      }
    }
    setUploading(false);
  };

  const removeMonth = async (key) => {
    try { await window.storage.delete("month:" + key, false); } catch (e) {}
    let idx = [];
    try {
      const res = await window.storage.get("months-index", false);
      idx = res ? JSON.parse(res.value) : [];
    } catch (e) { idx = []; }
    idx = idx.filter((k) => k !== key);
    try { await window.storage.set("months-index", JSON.stringify(idx), false); } catch (e) {}
    setMonthsData((prev) => { const n = { ...prev }; delete n[key]; return n; });
    setSelectedMonths((prev) => { const n = new Set(prev); n.delete(key); return n; });
  };

  const toggleMonth = (key) => {
    setSelectedMonths((prev) => {
      const n = new Set(prev);
      if (n.has(key)) n.delete(key); else n.add(key);
      return n;
    });
  };

  const sortedMonthKeys = useMemo(() => Object.keys(monthsData).sort(), [monthsData]);

  const aggregated = useMemo(() => {
    const map = {};
    for (const key of selectedMonths) {
      const m = monthsData[key];
      if (!m) continue;
      for (const [name, info] of Object.entries(m.employees)) {
        if (!map[name]) map[name] = { ot: 0, dept: info.dept, days: 0 };
        map[name].ot += info.ot;
        map[name].days += info.days;
        if (info.dept) map[name].dept = info.dept;
      }
    }
    return map;
  }, [monthsData, selectedMonths]);

  const selectedEmployees = useMemo(() => roster.filter((e) => selectedIds.has(e.id)), [roster, selectedIds]);

  const filteredRoster = useMemo(() => {
    const q = rosterSearch.trim().toLowerCase();
    if (!q) return roster;
    return roster.filter((e) => e.name.toLowerCase().includes(q));
  }, [roster, rosterSearch]);

  const resultRowsAll = useMemo(() => {
    return selectedEmployees.map((e) => {
      const a = aggregated[e.name];
      return {
        id: e.id,
        name: e.name,
        dept: a && a.dept ? a.dept : "-",
        ot: a ? Math.round(a.ot * 100) / 100 : 0,
        days: a ? a.days : 0,
        hasData: !!a,
      };
    });
  }, [selectedEmployees, aggregated]);

  const deptOptions = useMemo(() => {
    const set = new Set(["ทั้งหมด"]);
    resultRowsAll.forEach((r) => { if (r.hasData && r.dept !== "-") set.add(r.dept); });
    return Array.from(set);
  }, [resultRowsAll]);

  const resultRows = useMemo(() => {
    let rows = [...resultRowsAll];
    if (tableSearch.trim()) {
      const q = tableSearch.trim().toLowerCase();
      rows = rows.filter((r) => r.name.toLowerCase().includes(q));
    }
    if (deptFilter !== "ทั้งหมด") rows = rows.filter((r) => r.dept === deptFilter);
    if (otMin !== "") rows = rows.filter((r) => r.ot >= parseFloat(otMin));
    if (otMax !== "") rows = rows.filter((r) => r.ot <= parseFloat(otMax));
    rows.sort((a, b) => (sortAsc ? a.ot - b.ot : b.ot - a.ot));
    return rows;
  }, [resultRowsAll, tableSearch, deptFilter, otMin, otMax, sortAsc]);

  const pagedRows = useMemo(() => {
    const start = (page - 1) * pageSize;
    return resultRows.slice(start, start + pageSize);
  }, [resultRows, page]);

  const totalPages = Math.max(1, Math.ceil(resultRows.length / pageSize));

  const summary = useMemo(() => {
    const withData = resultRows.filter((r) => r.hasData);
    const n = withData.length;
    const total = withData.reduce((s, r) => s + r.ot, 0);
    const avg = n ? total / n : 0;
    const max = n ? Math.max(...withData.map((r) => r.ot)) : 0;
    const min = n ? Math.min(...withData.map((r) => r.ot)) : 0;
    return { selected: resultRows.length, withData: n, total, avg, max, min };
  }, [resultRows]);

  const median = useMemo(() => {
    const vals = resultRows.filter((r) => r.hasData).map((r) => r.ot).sort((a, b) => a - b);
    if (!vals.length) return 0;
    const mid = Math.floor(vals.length / 2);
    return vals.length % 2 ? vals[mid] : (vals[mid - 1] + vals[mid]) / 2;
  }, [resultRows]);

  const top10 = useMemo(
    () => [...resultRows].filter((r) => r.hasData).sort((a, b) => b.ot - a.ot).slice(0, 10),
    [resultRows]
  );
  const bottom10 = useMemo(
    () => [...resultRows].filter((r) => r.hasData).sort((a, b) => a.ot - b.ot).slice(0, 10),
    [resultRows]
  );

  const barData = useMemo(() => {
    const withData = [...resultRows].filter((r) => r.hasData).sort((a, b) => b.ot - a.ot);
    return withData.slice(0, 20).map((r) => ({ name: r.name, ot: r.ot }));
  }, [resultRows]);

  const pieData = useMemo(() => {
    const byDept = {};
    resultRows.filter((r) => r.hasData).forEach((r) => {
      byDept[r.dept] = (byDept[r.dept] || 0) + r.ot;
    });
    const arr = Object.entries(byDept).map(([dept, ot]) => ({ name: dept, value: Math.round(ot * 100) / 100 }));
    arr.sort((a, b) => b.value - a.value);
    if (arr.length > 6) {
      const head = arr.slice(0, 6);
      const restVal = arr.slice(6).reduce((s, x) => s + x.value, 0);
      head.push({ name: "อื่นๆ", value: Math.round(restVal * 100) / 100 });
      return head;
    }
    return arr;
  }, [resultRows]);

  const lineData = useMemo(() => {
    return sortedMonthKeys.map((key) => {
      const m = monthsData[key];
      let total = 0;
      for (const e of selectedEmployees) {
        const info = m.employees[e.name];
        if (info) total += info.ot;
      }
      return { month: m.label, total: Math.round(total * 100) / 100 };
    });
  }, [sortedMonthKeys, monthsData, selectedEmployees]);

  const toggleSelect = (id) => {
    setSelectedIds((prev) => {
      const n = new Set(prev);
      if (n.has(id)) n.delete(id); else n.add(id);
      return n;
    });
  };
  const selectAll = () => setSelectedIds(new Set(filteredRoster.map((e) => e.id)));
  const clearAll = () => setSelectedIds(new Set());
  const sortAZ = () => {
    const next = [...roster].sort((a, b) => a.name.localeCompare(b.name, "th"));
    persistRoster(next);
  };
  const addEmployee = () => {
    const name = newName.trim();
    if (!name) return;
    const id = `e${Date.now()}`;
    const next = [...roster, { id, name }];
    persistRoster(next);
    setSelectedIds((prev) => new Set([...prev, id]));
    setNewName("");
  };
  const removeEmployee = (id) => {
    const next = roster.filter((e) => e.id !== id);
    persistRoster(next);
    setSelectedIds((prev) => { const n = new Set(prev); n.delete(id); return n; });
  };
  const startEdit = (e) => { setEditingId(e.id); setEditingValue(e.name); };
  const saveEdit = () => {
    const next = roster.map((e) => (e.id === editingId ? { ...e, name: editingValue.trim() || e.name } : e));
    persistRoster(next);
    setEditingId(null);
  };

  const exportCSV = () => {
    const header = ["ลำดับ", "ชื่อพนักงาน", "แผนก/ตำแหน่ง", "ชั่วโมง OT", "จำนวนวันที่มีข้อมูล"];
    const lines = [header.join(",")];
    resultRows.forEach((r, i) => {
      lines.push([i + 1, `"${r.name}"`, `"${r.dept}"`, r.ot, r.days].join(","));
    });
    const blob = new Blob(["\uFEFF" + lines.join("\n")], { type: "text/csv;charset=utf-8;" });
    downloadBlob(blob, `OT_Summary_${new Date().toISOString().slice(0, 10)}.csv`);
  };

  const exportExcel = () => {
    const data = resultRows.map((r, i) => ({
      "ลำดับ": i + 1,
      "ชื่อพนักงาน": r.name,
      "แผนก/ตำแหน่ง": r.dept,
      "ชั่วโมง OT": r.ot,
      "จำนวนวันที่มีข้อมูล": r.days,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "OT Summary");
    XLSX.writeFile(wb, `OT_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const onDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    handleFiles(e.dataTransfer.files);
  };

  const hasAnyMonth = sortedMonthKeys.length > 0;

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900" style={{ fontFamily: "'Sarabun','Noto Sans Thai',ui-sans-serif,system-ui,sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;500;600;700&family=Noto+Sans+Thai:wght@400;500;600;700&display=swap');`}</style>

      {toast && (
        <div className={`fixed top-4 right-4 z-50 max-w-sm rounded-xl shadow-lg px-4 py-3 text-sm font-medium border ${
          toast.kind === "error" ? "bg-red-50 text-red-700 border-red-200" : "bg-blue-50 text-blue-700 border-blue-200"
        }`}>
          {toast.msg}
        </div>
      )}

      <div className="max-w-7xl mx-auto px-4 sm:px-6 py-6">
        <div className="flex items-center justify-between mb-6">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">OT Employee Search &amp; Summary Dashboard</h1>
            <p className="text-gray-500 text-sm mt-1">อัปโหลดไฟล์ OT รายเดือน เลือกพนักงาน แล้วระบบจะรวมชั่วโมง OT ให้อัตโนมัติ — อัปโหลดเดือนเดิมซ้ำเพื่อเขียนทับข้อมูล</p>
          </div>
          <div className="hidden sm:flex items-center gap-2 bg-blue-600 text-white rounded-xl px-4 py-2 shadow-sm">
            <Clock size={18} />
            <span className="text-sm font-semibold">{sortedMonthKeys.length} เดือนที่บันทึกไว้</span>
          </div>
        </div>

        {/* Upload */}
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          className={`rounded-2xl border-2 border-dashed p-6 bg-white shadow-sm transition-colors ${dragOver ? "border-blue-500 bg-blue-50" : "border-gray-200"}`}
        >
          <div className="flex flex-col sm:flex-row items-center gap-4">
            <div className="flex items-center justify-center w-14 h-14 rounded-full bg-blue-100 text-blue-600 flex-shrink-0">
              {uploading ? <Loader2 size={26} className="animate-spin" /> : <Upload size={26} />}
            </div>
            <div className="flex-1 text-center sm:text-left">
              <p className="font-semibold text-gray-800">ลากไฟล์ OT_Comparing_by_Emp_*.csv มาวางที่นี่ หรือเลือกไฟล์</p>
              <p className="text-sm text-gray-500">ระบบจะตรวจหา Header อัตโนมัติ และจับคู่เดือนจากช่วงวันที่ในไฟล์ — รองรับเลือกหลายไฟล์พร้อมกัน</p>
            </div>
            <button
              onClick={() => fileInputRef.current?.click()}
              className="bg-blue-600 hover:bg-blue-700 text-white font-medium px-5 py-2.5 rounded-xl shadow-sm whitespace-nowrap"
            >
              เลือกไฟล์ CSV
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".csv"
              multiple
              className="hidden"
              onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }}
            />
          </div>

          {hasAnyMonth && (
            <div className="mt-5 pt-4 border-t border-gray-100">
              <div className="flex items-center gap-2 text-sm text-gray-500 mb-2">
                <CalendarDays size={16} />
                <span>เดือนที่มีข้อมูล (ติ๊กเพื่อรวมในสรุป — อัปโหลดไฟล์เดือนเดิมซ้ำจะเขียนทับอัตโนมัติ)</span>
              </div>
              <div className="flex flex-wrap gap-2">
                {sortedMonthKeys.map((key) => {
                  const m = monthsData[key];
                  const active = selectedMonths.has(key);
                  return (
                    <div
                      key={key}
                      className={`flex items-center gap-2 rounded-full px-3 py-1.5 text-sm font-medium border cursor-pointer ${
                        active ? "bg-blue-600 text-white border-blue-600" : "bg-gray-50 text-gray-600 border-gray-200"
                      }`}
                      onClick={() => toggleMonth(key)}
                    >
                      {active ? <Check size={14} /> : null}
                      <span>{m.label}</span>
                      <span className={active ? "text-blue-100" : "text-gray-400"}>· {Object.keys(m.employees).length} คน</span>
                      <button
                        onClick={(e) => { e.stopPropagation(); removeMonth(key); }}
                        className={active ? "text-blue-100 hover:text-white" : "text-gray-400 hover:text-gray-600"}
                        title="ลบเดือนนี้"
                      >
                        <X size={14} />
                      </button>
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 mt-6">
          {/* Roster panel */}
          <div className="lg:col-span-4">
            <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4 lg:sticky lg:top-6">
              <div className="flex items-center justify-between mb-3">
                <h2 className="font-semibold text-gray-800 flex items-center gap-2">
                  <Users size={18} className="text-blue-600" /> รายชื่อพนักงาน
                </h2>
                <span className="text-xs bg-gray-100 text-gray-600 rounded-full px-2 py-1 font-medium">
                  เลือก {selectedIds.size}/{roster.length}
                </span>
              </div>

              <div className="relative mb-3">
                <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                <input
                  value={rosterSearch}
                  onChange={(e) => setRosterSearch(e.target.value)}
                  placeholder="ค้นหาชื่อพนักงาน..."
                  className="w-full pl-9 pr-3 py-2 text-sm rounded-lg border border-gray-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
              </div>

              <div className="flex flex-wrap gap-2 mb-3">
                <button onClick={selectAll} className="text-xs font-medium bg-blue-50 text-blue-700 rounded-lg px-3 py-1.5 hover:bg-blue-100">เลือกทั้งหมด</button>
                <button onClick={clearAll} className="text-xs font-medium bg-gray-50 text-gray-600 rounded-lg px-3 py-1.5 hover:bg-gray-100">ล้างการเลือก</button>
                <button onClick={sortAZ} className="text-xs font-medium bg-gray-50 text-gray-600 rounded-lg px-3 py-1.5 hover:bg-gray-100">เรียง A-Z</button>
              </div>

              <div className="max-h-[420px] overflow-y-auto divide-y divide-gray-100 border border-gray-100 rounded-xl">
                {filteredRoster.length === 0 && (
                  <div className="p-4 text-sm text-gray-400 text-center">ไม่พบพนักงานที่ค้นหา</div>
                )}
                {filteredRoster.map((e) => {
                  const a = aggregated[e.name];
                  const checked = selectedIds.has(e.id);
                  const isEditing = editingId === e.id;
                  return (
                    <div key={e.id} className="flex items-center gap-2 px-3 py-2 hover:bg-gray-50">
                      <button onClick={() => toggleSelect(e.id)} className="text-blue-600 flex-shrink-0">
                        {checked ? <CheckSquare size={18} /> : <Square size={18} className="text-gray-300" />}
                      </button>
                      {isEditing ? (
                        <input
                          autoFocus
                          value={editingValue}
                          onChange={(ev) => setEditingValue(ev.target.value)}
                          onKeyDown={(ev) => ev.key === "Enter" && saveEdit()}
                          className="flex-1 text-sm border border-blue-300 rounded-md px-2 py-1 focus:outline-none"
                        />
                      ) : (
                        <button onClick={() => toggleSelect(e.id)} className="flex-1 text-left text-sm text-gray-700">
                          {e.name}
                          {a ? (
                            <span className="ml-2 text-xs text-gray-400">({a.ot.toFixed(1)} ชม.)</span>
                          ) : hasAnyMonth ? (
                            <span className="ml-2 text-xs text-amber-500">ไม่มีข้อมูล</span>
                          ) : null}
                        </button>
                      )}
                      {isEditing ? (
                        <button onClick={saveEdit} className="text-green-600 flex-shrink-0"><Check size={16} /></button>
                      ) : (
                        <button onClick={() => startEdit(e)} className="text-gray-300 hover:text-blue-600 flex-shrink-0"><Pencil size={14} /></button>
                      )}
                      <button onClick={() => removeEmployee(e.id)} className="text-gray-300 hover:text-red-500 flex-shrink-0"><Trash2 size={14} /></button>
                    </div>
                  );
                })}
              </div>

              <div className="flex gap-2 mt-3">
                <input
                  value={newName}
                  onChange={(e) => setNewName(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && addEmployee()}
                  placeholder="เพิ่มพนักงานใหม่..."
                  className="flex-1 text-sm border border-gray-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
                />
                <button onClick={addEmployee} className="bg-gray-900 text-white text-sm font-medium rounded-lg px-4 flex items-center gap-1 hover:bg-gray-800">
                  <Plus size={15} /> เพิ่ม
                </button>
              </div>
              <p className="text-xs text-gray-400 mt-2">* ชื่อพนักงานต้องตรงกับชื่อในไฟล์ CSV ทุกตัวอักษรจึงจะจับคู่ข้อมูล OT ได้ถูกต้อง</p>
            </div>
          </div>

          {/* Right: summary + charts + table */}
          <div className="lg:col-span-8 space-y-6">
            {!hasAnyMonth ? (
              <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-10 text-center">
                <AlertCircle className="mx-auto text-gray-300 mb-3" size={36} />
                <p className="text-gray-500 font-medium">ยังไม่มีข้อมูล OT — กรุณาอัปโหลดไฟล์ CSV ด้านบนเพื่อเริ่มสรุปผล</p>
              </div>
            ) : (
              <>
                {/* Summary cards */}
                <div className="grid grid-cols-2 sm:grid-cols-5 gap-3">
                  {[
                    { label: "Employee", value: summary.selected, suffix: "", icon: Users },
                    { label: "Total OT", value: summary.total.toFixed(1), suffix: " hrs", icon: Clock },
                    { label: "Average", value: summary.avg.toFixed(2), suffix: " hrs", icon: TrendingUp },
                    { label: "Highest", value: summary.max.toFixed(1), suffix: " hrs", icon: TrendingUp },
                    { label: "Lowest", value: summary.min.toFixed(1), suffix: " hrs", icon: TrendingDown },
                  ].map((c, i) => (
                    <div key={i} className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                      <div className="flex items-center gap-2 text-gray-400 text-xs font-medium mb-1">
                        <c.icon size={14} /> {c.label}
                      </div>
                      <div className="text-xl font-bold text-gray-900">{c.value}<span className="text-sm font-medium text-gray-400">{c.suffix}</span></div>
                    </div>
                  ))}
                </div>

                {/* Charts */}
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                    <h3 className="text-sm font-semibold text-gray-700 mb-2">
                      ชั่วโมง OT รายบุคคล {barData.length < resultRows.filter(r=>r.hasData).length ? `(แสดง ${barData.length} อันดับแรก)` : ""}
                    </h3>
                    <ResponsiveContainer width="100%" height={Math.max(220, barData.length * 26)}>
                      <BarChart data={barData} layout="vertical" margin={{ left: 10, right: 20 }}>
                        <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#eef2f7" />
                        <XAxis type="number" tick={{ fontSize: 11 }} />
                        <YAxis type="category" dataKey="name" width={110} tick={{ fontSize: 11 }} />
                        <Tooltip formatter={(v) => [`${v} ชม.`, "OT"]} />
                        <Bar dataKey="ot" fill="#2563eb" radius={[0, 6, 6, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                    <h3 className="text-sm font-semibold text-gray-700 mb-2">สัดส่วน OT ตามแผนก/ตำแหน่ง</h3>
                    <ResponsiveContainer width="100%" height={260}>
                      <PieChart>
                        <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={85} label={{ fontSize: 11 }}>
                          {pieData.map((_, i) => <Cell key={i} fill={BLUE[i % BLUE.length]} />)}
                        </Pie>
                        <Tooltip formatter={(v) => [`${v} ชม.`, "OT"]} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>

                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4 lg:col-span-2">
                    <h3 className="text-sm font-semibold text-gray-700 mb-2">แนวโน้ม OT รวมตามเดือน (ของพนักงานที่เลือกไว้)</h3>
                    <ResponsiveContainer width="100%" height={220}>
                      <LineChart data={lineData} margin={{ left: 10, right: 20 }}>
                        <CartesianGrid strokeDasharray="3 3" stroke="#eef2f7" />
                        <XAxis dataKey="month" tick={{ fontSize: 11 }} />
                        <YAxis tick={{ fontSize: 11 }} />
                        <Tooltip formatter={(v) => [`${v} ชม.`, "OT รวม"]} />
                        <Line type="monotone" dataKey="total" stroke="#1d4ed8" strokeWidth={2.5} dot={{ r: 4 }} />
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                {/* Filters */}
                <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                  <div className="flex flex-wrap items-center gap-3">
                    <div className="relative flex-1 min-w-[180px]">
                      <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input
                        value={tableSearch}
                        onChange={(e) => setTableSearch(e.target.value)}
                        placeholder="ค้นหาในผลลัพธ์..."
                        className="w-full pl-9 pr-3 py-2 text-sm rounded-lg border border-gray-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                    </div>
                    <select
                      value={deptFilter}
                      onChange={(e) => setDeptFilter(e.target.value)}
                      className="text-sm rounded-lg border border-gray-200 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      {deptOptions.map((d) => <option key={d} value={d}>{d}</option>)}
                    </select>
                    <input
                      type="number"
                      value={otMin}
                      onChange={(e) => setOtMin(e.target.value)}
                      placeholder="OT ขั้นต่ำ"
                      className="w-28 text-sm rounded-lg border border-gray-200 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <input
                      type="number"
                      value={otMax}
                      onChange={(e) => setOtMax(e.target.value)}
                      placeholder="OT สูงสุด"
                      className="w-28 text-sm rounded-lg border border-gray-200 px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
                    />
                    <div className="flex gap-2 ml-auto">
                      <button onClick={exportCSV} className="flex items-center gap-1.5 text-sm font-medium bg-gray-50 text-gray-700 rounded-lg px-3 py-2 hover:bg-gray-100">
                        <Download size={15} /> CSV
                      </button>
                      <button onClick={exportExcel} className="flex items-center gap-1.5 text-sm font-medium bg-green-50 text-green-700 rounded-lg px-3 py-2 hover:bg-green-100">
                        <FileSpreadsheet size={15} /> Excel
                      </button>
                    </div>
                  </div>
                </div>

                {/* Table */}
                <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-gray-50 text-gray-500 text-xs uppercase">
                        <th className="text-left px-4 py-3 font-semibold">#</th>
                        <th className="text-left px-4 py-3 font-semibold">Employee</th>
                        <th className="text-left px-4 py-3 font-semibold">Department / Position</th>
                        <th
                          className="text-right px-4 py-3 font-semibold cursor-pointer select-none"
                          onClick={() => setSortAsc((s) => !s)}
                        >
                          <span className="inline-flex items-center gap-1">
                            OT (hrs) {sortAsc ? <ChevronUp size={13} /> : <ChevronDown size={13} />}
                          </span>
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-100">
                      {pagedRows.length === 0 && (
                        <tr><td colSpan={4} className="text-center text-gray-400 py-8">No Data</td></tr>
                      )}
                      {pagedRows.map((r, i) => (
                        <tr key={r.id} className="hover:bg-gray-50">
                          <td className="px-4 py-2.5 text-gray-400">{(page - 1) * pageSize + i + 1}</td>
                          <td className="px-4 py-2.5 font-medium text-gray-800">{r.name}</td>
                          <td className="px-4 py-2.5 text-gray-500">{r.dept}</td>
                          <td className="px-4 py-2.5 text-right">
                            {r.hasData ? (
                              <span className="font-semibold text-gray-800">{r.ot.toFixed(1)}</span>
                            ) : (
                              <span className="text-xs text-amber-500 bg-amber-50 rounded-full px-2 py-0.5">ไม่มีข้อมูล</span>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  <div className="flex items-center justify-between px-4 py-3 border-t border-gray-100 text-sm text-gray-500">
                    <span>{resultRows.length} รายการ</span>
                    <div className="flex items-center gap-2">
                      <button disabled={page <= 1} onClick={() => setPage((p) => p - 1)} className="px-3 py-1 rounded-lg border border-gray-200 disabled:opacity-40">ก่อนหน้า</button>
                      <span>หน้า {page}/{totalPages}</span>
                      <button disabled={page >= totalPages} onClick={() => setPage((p) => p + 1)} className="px-3 py-1 rounded-lg border border-gray-200 disabled:opacity-40">ถัดไป</button>
                    </div>
                  </div>
                </div>

                {/* Statistics */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                    <h3 className="text-sm font-semibold text-gray-700 mb-3 flex items-center gap-2"><TrendingUp size={15} className="text-blue-600" /> Top 10 OT</h3>
                    <ol className="space-y-1.5 text-sm">
                      {top10.map((r, i) => (
                        <li key={r.id} className="flex justify-between">
                          <span className="text-gray-600">{i + 1}. {r.name}</span>
                          <span className="font-semibold text-gray-800">{r.ot.toFixed(1)}</span>
                        </li>
                      ))}
                      {top10.length === 0 && <li className="text-gray-400">No Data</li>}
                    </ol>
                  </div>
                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4">
                    <h3 className="text-sm font-semibold text-gray-700 mb-3 flex items-center gap-2"><TrendingDown size={15} className="text-blue-600" /> Bottom 10 OT</h3>
                    <ol className="space-y-1.5 text-sm">
                      {bottom10.map((r, i) => (
                        <li key={r.id} className="flex justify-between">
                          <span className="text-gray-600">{i + 1}. {r.name}</span>
                          <span className="font-semibold text-gray-800">{r.ot.toFixed(1)}</span>
                        </li>
                      ))}
                      {bottom10.length === 0 && <li className="text-gray-400">No Data</li>}
                    </ol>
                  </div>
                  <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-4 sm:col-span-2">
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 text-center">
                      <div><div className="text-xs text-gray-400 mb-1">Average</div><div className="font-bold text-gray-800">{summary.avg.toFixed(2)}</div></div>
                      <div><div className="text-xs text-gray-400 mb-1">Median</div><div className="font-bold text-gray-800">{median.toFixed(2)}</div></div>
                      <div><div className="text-xs text-gray-400 mb-1">Maximum</div><div className="font-bold text-gray-800">{summary.max.toFixed(1)}</div></div>
                      <div><div className="text-xs text-gray-400 mb-1">Minimum</div><div className="font-bold text-gray-800">{summary.min.toFixed(1)}</div></div>
                    </div>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
