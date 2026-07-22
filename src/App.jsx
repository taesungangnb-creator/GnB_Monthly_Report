import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  RadarChart, Radar, PolarGrid, PolarAngleAxis, PolarRadiusAxis,
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid, Legend, LabelList
} from "recharts";

// ---------- 고정 정의 ----------
const PART_DEFS = [
  { key: "partI", label: "Part I", kr: "Vocabulary" },
  { key: "partII", label: "Part II", kr: "Listening" },
  { key: "partIII", label: "Part III", kr: "Dictation" },
  { key: "partIV", label: "Part IV", kr: "Reading" },
  { key: "partV", label: "Part V", kr: "Writing" },
];

const PARTICIPATION_DEFS = [
  { key: "attendance", label: "Attendance", kr: "출석" },
  { key: "presentation", label: "Presentation Skill", kr: "발표력" },
  { key: "spontaneity", label: "Spontaneity", kr: "자발성" },
];
const BEHAVIOR_DEFS = [
  { key: "manners", label: "Manners", kr: "예의" },
  { key: "classAttitude", label: "Class Attitude", kr: "수업태도" },
  { key: "interpersonal", label: "Interpersonal Skill", kr: "교우관계" },
];
const HOMEWORK_DEFS = [
  { key: "textbookHw", label: "교재 및 쓰기 숙제", kr: "" },
  { key: "onlineHw", label: "Gplum 온라인 학습", kr: "" },
  { key: "accuracy", label: "정확성 및 완성도", kr: "" },
];

const TEXTBOOKS = ["(Gplum) Susie's Day 1"];

// 영역별 만점은 교재 기준 고정값입니다. (Part I~V는 시험 만점, 그 외 항목은 10점 만점)
const MAX_SCORES = { partI: 10, partII: 6, partIII: 6, partIV: 4, partV: 14 };
const DEFAULT_MAX = 10;
function getRowMax(key) {
  return MAX_SCORES[key] !== undefined ? MAX_SCORES[key] : DEFAULT_MAX;
}
function clamp(value, max) {
  const n = Number(value);
  if (Number.isNaN(n)) return 0;
  return Math.min(Math.max(n, 0), max);
}

function makeStudent(i) {
  const base = { id: i, name: "", comment: "" };
  PART_DEFS.forEach((p) => (base[p.key] = 0));
  PARTICIPATION_DEFS.forEach((p) => (base[p.key] = 0));
  BEHAVIOR_DEFS.forEach((p) => (base[p.key] = 0));
  HOMEWORK_DEFS.forEach((p) => (base[p.key] = 0));
  return base;
}

// 화면의 입력표와 동일한 행 구성 (라벨 / 만점 / 학생1 / 학생2 ...)
const ROW_DEFS = [
  ...PART_DEFS.map((p) => ({ label: `${p.label} (${p.kr})`, key: p.key, max: MAX_SCORES[p.key] })),
  ...PARTICIPATION_DEFS.map((d) => ({ label: `${d.label} (${d.kr})`, key: d.key, max: DEFAULT_MAX })),
  ...BEHAVIOR_DEFS.map((d) => ({ label: `${d.label} (${d.kr})`, key: d.key, max: DEFAULT_MAX })),
  ...HOMEWORK_DEFS.map((d) => ({ label: d.label, key: d.key, max: DEFAULT_MAX })),
];
const SECTION_BREAK_ROWS = ["Participation (참여도)", "Behavior (태도)", "Homework (숙제)"];
const COMMENT_ROW_LABEL = "Teacher Comments";

// 학생 데이터 배열 -> 화면 표와 동일한 구조의 엑셀 파일로 다운로드 (템플릿 다운로드 / 결과 다운로드 겸용)
function exportStudentsToExcel(form, students, filenameSuffix = "") {
  const aoa = [];
  aoa.push(["항목", "만점", ...students.map((s) => s.name || "")]);
  PART_DEFS.forEach((p) => aoa.push([`${p.label} (${p.kr})`, MAX_SCORES[p.key], ...students.map((s) => s[p.key])]));
  aoa.push(["Participation (참여도)", "", ...students.map(() => "")]);
  PARTICIPATION_DEFS.forEach((d) => aoa.push([`${d.label} (${d.kr})`, DEFAULT_MAX, ...students.map((s) => s[d.key])]));
  aoa.push(["Behavior (태도)", "", ...students.map(() => "")]);
  BEHAVIOR_DEFS.forEach((d) => aoa.push([`${d.label} (${d.kr})`, DEFAULT_MAX, ...students.map((s) => s[d.key])]));
  aoa.push(["Homework (숙제)", "", ...students.map(() => "")]);
  HOMEWORK_DEFS.forEach((d) => aoa.push([d.label, DEFAULT_MAX, ...students.map((s) => s[d.key])]));
  aoa.push([COMMENT_ROW_LABEL, "", ...students.map((s) => s.comment || "")]);

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"] = [{ wch: 28 }, { wch: 8 }, ...students.map(() => ({ wch: 16 }))];

  const infoRows = [
    { 항목: "담임교사", 값: form.teacher },
    { 항목: "Class명", 값: form.className },
    { 항목: "수업일자", 값: `${form.dateStart} ~ ${form.dateEnd}` },
    { 항목: "교재명", 값: form.textbook },
    { 항목: "학원명", 값: form.academyName },
    { 항목: "전화번호", 값: form.phone },
  ];
  const wsInfo = XLSX.utils.json_to_sheet(infoRows);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "성적입력");
  XLSX.utils.book_append_sheet(wb, wsInfo, "기본정보");

  const safeClass = (form.className || "성적표").replace(/[\\/:*?"<>|]/g, "");
  const filename = `${safeClass}${filenameSuffix}.xlsx`;
  XLSX.writeFile(wb, filename);
}

// 화면 표와 동일한 구조의 엑셀 파일 -> 학생 데이터 배열로 변환
function parseExcelFile(file, onSuccess, onError) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: "array" });
      const sheetName = wb.SheetNames.includes("성적입력") ? "성적입력" : wb.SheetNames[0];
      const ws = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      if (!aoa.length) throw new Error("시트에서 데이터를 찾을 수 없습니다.");

      const headerRow = aoa[0];
      const studentCount = Math.max(0, headerRow.length - 2);
      if (!studentCount) throw new Error("학생 열을 찾을 수 없습니다. 템플릿 형식을 확인해주세요.");

      const names = headerRow.slice(2);
      const students = names.map((n, i) => makeStudent(i + 1));
      students.forEach((s, i) => { s.name = String(names[i] ?? "").trim(); });

      const labelMap = {};
      ROW_DEFS.forEach((r) => { labelMap[r.label] = r; });

      for (let r = 1; r < aoa.length; r++) {
        const row = aoa[r];
        const label = String(row[0] ?? "").trim();
        if (!label || SECTION_BREAK_ROWS.includes(label)) continue;
        if (label === COMMENT_ROW_LABEL) {
          students.forEach((s, i) => { s.comment = String(row[2 + i] ?? ""); });
          continue;
        }
        const def = labelMap[label];
        if (!def) continue;
        students.forEach((s, i) => {
          const raw = row[2 + i];
          s[def.key] = clamp(raw === "" || raw === undefined ? 0 : raw, def.max);
        });
      }
      onSuccess(students);
    } catch (err) {
      onError(err);
    }
  };
  reader.onerror = () => onError(new Error("파일을 읽지 못했습니다."));
  reader.readAsArrayBuffer(file);
}

function grade100(pct) {
  if (pct >= 100) return { label: "Perfect", color: "#16a34a", bg: "#dcfce7" };
  if (pct >= 90) return { label: "Excellent", color: "#0891b2", bg: "#cffafe" };
  if (pct >= 80) return { label: "Very Good", color: "#2563eb", bg: "#dbeafe" };
  if (pct >= 70) return { label: "Good", color: "#7c3aed", bg: "#ede9fe" };
  if (pct >= 60) return { label: "Not Bad", color: "#ea580c", bg: "#ffedd5" };
  return { label: "Practice More", color: "#e11d48", bg: "#ffe4e6" };
}
function grade10(v) {
  if (v >= 10) return { label: "Perfect", color: "#16a34a", bg: "#dcfce7" };
  if (v >= 9) return { label: "Excellent", color: "#0891b2", bg: "#cffafe" };
  if (v >= 7) return { label: "Good", color: "#2563eb", bg: "#dbeafe" };
  if (v >= 6) return { label: "Not Bad", color: "#ea580c", bg: "#ffedd5" };
  return { label: "Practice More", color: "#e11d48", bg: "#ffe4e6" };
}
function gradeBadgeStyle(g) {
  return {
    display: "inline-block", width: 96, textAlign: "center",
    background: g.bg, color: g.color, padding: "4px 0",
    borderRadius: 999, fontSize: 10.5, fontWeight: 700,
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
    boxSizing: "border-box", lineHeight: 1.3,
  };
}

// ---------- 메인 앱 ----------
export default function App() {
  const [step, setStep] = useState(1);
  const [form, setForm] = useState({
    teacher: "",
    className: "",
    dateStart: "",
    dateEnd: "",
    academyName: "",
    phone: "",
    textbook: TEXTBOOKS[0],
  });
  const [studentCount, setStudentCount] = useState(7);
  const [students, setStudents] = useState(
    Array.from({ length: 7 }, (_, i) => makeStudent(i + 1))
  );
  const [reportIndex, setReportIndex] = useState(0);

  const totalMax = useMemo(
    () => Object.values(MAX_SCORES).reduce((a, b) => a + Number(b || 0), 0),
    []
  );

  function updateStudentCount(n) {
    n = Math.max(1, Math.min(20, n));
    setStudentCount(n);
    setStudents((prev) => {
      const arr = [...prev];
      if (n > arr.length) {
        for (let i = arr.length; i < n; i++) arr.push(makeStudent(i + 1));
      } else {
        arr.length = n;
      }
      return arr;
    });
  }

  function updateStudentField(idx, key, value) {
    setStudents((prev) => {
      const arr = [...prev];
      arr[idx] = { ...arr[idx], [key]: value };
      return arr;
    });
  }

  function replaceAllStudents(newStudents) {
    const trimmed = newStudents.slice(0, 20);
    setStudents(trimmed);
    setStudentCount(Math.max(1, trimmed.length));
  }

  // 반평균 계산
  const classAverages = useMemo(() => {
    const avg = {};
    PART_DEFS.forEach((p) => {
      const vals = students.map((s) => Number(s[p.key]) || 0);
      const mean = vals.reduce((a, b) => a + b, 0) / (vals.length || 1);
      avg[p.key] = MAX_SCORES[p.key] ? (mean / MAX_SCORES[p.key]) * 100 : 0;
    });
    return avg;
  }, [students]);

  return (
    <div className="app-shell" style={{ minHeight: "100vh", background: "#f3f4f6", fontFamily: "'Pretendard','Malgun Gothic',sans-serif" }}>
      <StepIndicator step={step} />
      {step === 1 && (
        <Step1
          form={form}
          setForm={setForm}
          onNext={() => setStep(2)}
        />
      )}
      {step === 2 && (
        <Step2
          form={form}
          studentCount={studentCount}
          updateStudentCount={updateStudentCount}
          students={students}
          updateStudentField={updateStudentField}
          replaceAllStudents={replaceAllStudents}
          onBack={() => setStep(1)}
          onNext={() => {
            setReportIndex(0);
            setStep(3);
          }}
        />
      )}
      {step === 3 && (
        <Step3
          form={form}
          totalMax={totalMax}
          students={students}
          classAverages={classAverages}
          reportIndex={reportIndex}
          setReportIndex={setReportIndex}
          onBack={() => setStep(2)}
        />
      )}
    </div>
  );
}

// ---------- 상단 스텝 표시 ----------
function StepIndicator({ step }) {
  const steps = ["기본정보 입력", "학생 성적 입력", "최종 성적표"];
  return (
    <div style={{ background: "#111827", padding: "14px 20px", display: "flex", justifyContent: "center", gap: 0 }} className="print-hide">
      {steps.map((s, i) => {
        const n = i + 1;
        const active = step === n;
        const done = step > n;
        return (
          <div key={s} style={{ display: "flex", alignItems: "center" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div
                style={{
                  width: 26, height: 26, borderRadius: "50%",
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: 13, fontWeight: 700,
                  background: active ? "#f43f5e" : done ? "#16a34a" : "#374151",
                  color: "#fff",
                }}
              >
                {done ? "✓" : n}
              </div>
              <span style={{ color: active ? "#fff" : "#9ca3af", fontSize: 13, fontWeight: active ? 700 : 500 }}>
                {s}
              </span>
            </div>
            {i < steps.length - 1 && (
              <div style={{ width: 40, height: 2, background: "#374151", margin: "0 14px" }} />
            )}
          </div>
        );
      })}
    </div>
  );
}

// ---------- STEP 1 ----------
function Step1({ form, setForm, onNext }) {
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));
  const canProceed = form.teacher && form.className && form.dateStart && form.dateEnd && form.academyName;

  return (
    <div style={{ maxWidth: 560, margin: "0 auto", padding: "40px 20px" }}>
      <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", overflow: "hidden" }}>
        <div style={{ background: "linear-gradient(90deg,#1d4ed8,#db2777)", padding: "22px 28px" }}>
          <div style={{ color: "#fff", fontSize: 12, letterSpacing: 2, opacity: 0.85 }}>GNB EDUCATION</div>
          <div style={{ color: "#fff", fontSize: 22, fontWeight: 800, marginTop: 4 }}>성적표 기본 정보 입력</div>
        </div>
        <div style={{ padding: "28px" }}>
          <Field label="담임교사" required>
            <input style={inputStyle} placeholder="예) Sophie" value={form.teacher} onChange={set("teacher")} />
          </Field>
          <Field label="Class명" required>
            <input style={inputStyle} placeholder="예) Harvard 2:00" value={form.className} onChange={set("className")} />
          </Field>
          <Field label="수업일자" required>
            <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
              <input type="date" style={inputStyle} value={form.dateStart} onChange={set("dateStart")} />
              <span style={{ color: "#9ca3af" }}>~</span>
              <input type="date" style={inputStyle} value={form.dateEnd} onChange={set("dateEnd")} />
            </div>
          </Field>
          <Field label="학원명" required>
            <input style={inputStyle} placeholder="예) 지앤비어학원 OOO캠퍼스" value={form.academyName} onChange={set("academyName")} />
          </Field>
          <Field label="전화번호">
            <input style={inputStyle} placeholder="예) 02-567-0582" value={form.phone} onChange={set("phone")} />
          </Field>
          <Field label="교재명" required>
            <select style={inputStyle} value={form.textbook} onChange={set("textbook")}>
              {TEXTBOOKS.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
            <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 6 }}>
              * 교재는 추후 엑셀 데이터 업로드로 계속 추가될 예정입니다.
            </div>
          </Field>

          <button
            disabled={!canProceed}
            onClick={onNext}
            style={{
              width: "100%", marginTop: 10, padding: "14px",
              borderRadius: 10, border: "none", fontSize: 15, fontWeight: 700,
              cursor: canProceed ? "pointer" : "not-allowed",
              background: canProceed ? "#111827" : "#d1d5db",
              color: "#fff",
            }}
          >
            확인 → 성적 입력표로 이동
          </button>
        </div>
      </div>
    </div>
  );
}

function Field({ label, required, children }) {
  return (
    <div style={{ marginBottom: 18 }}>
      <label style={{ display: "block", fontSize: 13, fontWeight: 700, color: "#374151", marginBottom: 6 }}>
        {label} {required && <span style={{ color: "#e11d48" }}>*</span>}
      </label>
      {children}
    </div>
  );
}
const inputStyle = {
  width: "100%", padding: "10px 12px", borderRadius: 8,
  border: "1px solid #d1d5db", fontSize: 14, boxSizing: "border-box",
  outline: "none",
};

// ---------- STEP 2 ----------
const LABEL_W = 176;
const MAXCOL_W = 60;
const STUDENT_COL_W = 84;
const COMMENT_ROW_W = 130;
const COMMENT_MAX_LEN = 220; // 최종 성적표 인쇄 시 코멘트 칸(고정 높이)에 딱 맞도록 제한

function Step2({ form, studentCount, updateStudentCount, students, updateStudentField, replaceAllStudents, onBack, onNext }) {
  const [uploadError, setUploadError] = useState("");
  const fileInputRef = React.useRef(null);

  function handleUploadClick() {
    setUploadError("");
    fileInputRef.current?.click();
  }
  function handleFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const ok = window.confirm("엑셀 파일의 데이터로 현재 입력표를 덮어씁니다. 계속할까요?");
    if (!ok) { e.target.value = ""; return; }
    parseExcelFile(
      file,
      (newStudents) => {
        replaceAllStudents(newStudents);
        setUploadError("");
        alert(`${newStudents.length}명의 성적 데이터를 불러왔습니다.`);
      },
      (err) => setUploadError(err.message || "업로드 중 오류가 발생했습니다.")
    );
    e.target.value = "";
  }

  const rowLabelStyle = {
    position: "sticky", left: 0, background: "#fecaca", zIndex: 2,
    padding: "7px 10px", fontSize: 12, fontWeight: 700, color: "#7f1d1d",
    borderRight: "2px solid #fff", borderBottom: "1px solid #fca5a5",
    whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", textAlign: "left",
  };
  const maxColStyle = {
    position: "sticky", left: LABEL_W, background: "#e5e7eb", zIndex: 2,
    padding: "7px 6px", fontSize: 12, fontWeight: 800, color: "#374151",
    borderRight: "2px solid #cbd5e1", borderBottom: "1px solid #d1d5db",
    textAlign: "center",
  };
  const groupLabelStyle = { ...rowLabelStyle, background: "#fca5a5", borderBottom: "1px solid #f87171" };
  const groupMaxStyle = { ...maxColStyle, background: "#f1f5f9", borderBottom: "1px solid #d1d5db" };

  const dataCellStyle = (i, isLast) => ({
    padding: 3, background: i % 2 === 0 ? "#fef9c3" : "#fef3c7",
    borderRight: isLast ? "none" : "2px solid #fbbf24",
    borderBottom: "1px solid #fde68a",
    textAlign: "center",
  });
  const emptyCellStyle = (i, isLast) => ({
    background: i % 2 === 0 ? "#fee2e2" : "#fecaca",
    borderRight: isLast ? "none" : "2px solid #fca5a5",
  });
  const cellInput = {
    width: "100%", maxWidth: 56, textAlign: "center", padding: "5px 2px", fontSize: 12,
    border: "1px solid #fbbf24", borderRadius: 4, background: "#fffbeb", boxSizing: "border-box",
  };

  return (
    <div style={{ maxWidth: 1280, margin: "0 auto", padding: "30px 20px" }}>
      <div style={{ background: "#fff", borderRadius: 16, boxShadow: "0 1px 3px rgba(0,0,0,0.1)", overflow: "hidden" }}>
        <div style={{ padding: "20px 26px", borderBottom: "1px solid #f1f5f9" }}>
          <div style={{ fontSize: 22, fontWeight: 800, color: "#111827", marginBottom: 10 }}>학생 성적 입력표</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: "8px 22px", fontSize: 13, color: "#374151" }}>
            <span><b>담임교사</b> {form.teacher}</span>
            <span><b>Class명</b> {form.className}</span>
            <span><b>수업일자</b> {form.dateStart} ~ {form.dateEnd}</span>
            <span><b>교재명</b> {form.textbook}</span>
            <span><b>학원명</b> {form.academyName}</span>
            {form.phone && <span><b>전화</b> {form.phone}</span>}
          </div>
        </div>

        <div style={{ padding: "16px 26px", display: "flex", flexWrap: "wrap", alignItems: "center", gap: 16, borderBottom: "1px solid #f1f5f9" }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>학생 수</span>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <button onClick={() => updateStudentCount(studentCount - 1)} style={stepperBtn}>-</button>
            <span style={{ minWidth: 24, textAlign: "center", fontWeight: 700 }}>{studentCount}</span>
            <button onClick={() => updateStudentCount(studentCount + 1)} style={stepperBtn}>+</button>
          </div>
          <span style={{ fontSize: 12, color: "#9ca3af" }}>‘만점’ 열은 교재 기준 고정값이며 수정할 수 없습니다. 입력값은 만점을 초과할 수 없습니다.</span>

          <div style={{ marginLeft: "auto", display: "flex", gap: 8, alignItems: "center" }}>
            <button onClick={() => exportStudentsToExcel(form, students, "_템플릿")} style={secondaryBtn}>📥 엑셀 템플릿 다운로드</button>
            <button onClick={handleUploadClick} style={secondaryBtn}>📤 엑셀로 일괄 입력</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} style={{ display: "none" }} />
          </div>
        </div>
        {uploadError && (
          <div style={{ padding: "0 26px 12px", color: "#dc2626", fontSize: 12 }}>⚠ {uploadError}</div>
        )}

        <div style={{ overflowX: "auto", padding: "18px 26px" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: LABEL_W }} />
              <col style={{ width: MAXCOL_W }} />
              {students.map((s) => <col key={s.id} style={{ width: STUDENT_COL_W }} />)}
            </colgroup>
            <thead>
              <tr>
                <th style={{ ...rowLabelStyle, background: "#111827", color: "#fff", zIndex: 3 }}>학생명</th>
                <th style={{ ...maxColStyle, background: "#111827", color: "#fff", zIndex: 3 }}>만점</th>
                {students.map((s, i) => (
                  <th key={s.id} style={{ padding: "4px 3px", background: "#111827", borderRight: i === students.length - 1 ? "none" : "2px solid #374151" }}>
                    <input
                      value={s.name}
                      onChange={(e) => updateStudentField(i, "name", e.target.value)}
                      placeholder={`학생명 입력`}
                      style={{ width: "100%", boxSizing: "border-box", textAlign: "center", padding: "4px 2px", fontSize: 12, border: "1px solid #4b5563", borderRadius: 4, color: "#111827", background: s.name ? "#fff" : "#f3f4f6" }}
                    />
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {PART_DEFS.map((p) => (
                <tr key={p.key}>
                  <td style={rowLabelStyle}>{p.label} <span style={{ fontWeight: 400 }}>({p.kr})</span></td>
                  <td style={maxColStyle}>{getRowMax(p.key)}</td>
                  {students.map((s, i) => (
                    <td key={s.id} style={dataCellStyle(i, i === students.length - 1)}>
                      <input
                        type="number" min={0} max={getRowMax(p.key)}
                        value={s[p.key]}
                        onChange={(e) => updateStudentField(i, p.key, clamp(e.target.value, getRowMax(p.key)))}
                        onFocus={(e) => e.target.select()}
                        style={cellInput}
                      />
                    </td>
                  ))}
                </tr>
              ))}

              <SectionRows title="Participation (참여도)" defs={PARTICIPATION_DEFS} students={students} update={updateStudentField} groupLabelStyle={groupLabelStyle} groupMaxStyle={groupMaxStyle} rowLabelStyle={rowLabelStyle} maxColStyle={maxColStyle} dataCellStyle={dataCellStyle} emptyCellStyle={emptyCellStyle} cellInput={cellInput} />
              <SectionRows title="Behavior (태도)" defs={BEHAVIOR_DEFS} students={students} update={updateStudentField} groupLabelStyle={groupLabelStyle} groupMaxStyle={groupMaxStyle} rowLabelStyle={rowLabelStyle} maxColStyle={maxColStyle} dataCellStyle={dataCellStyle} emptyCellStyle={emptyCellStyle} cellInput={cellInput} />
              <SectionRows title="Homework (숙제)" defs={HOMEWORK_DEFS} students={students} update={updateStudentField} groupLabelStyle={groupLabelStyle} groupMaxStyle={groupMaxStyle} rowLabelStyle={rowLabelStyle} maxColStyle={maxColStyle} dataCellStyle={dataCellStyle} emptyCellStyle={emptyCellStyle} cellInput={cellInput} />

              <tr>
                <td style={{ ...rowLabelStyle, background: "#fca5a5" }}>Teacher Comments</td>
                <td style={{ ...maxColStyle, background: "#f1f5f9" }}>-</td>
                {students.map((s, i) => (
                  <td key={s.id} style={{ padding: 3, background: i % 2 === 0 ? "#fef9c3" : "#fef3c7", borderRight: i === students.length - 1 ? "none" : "2px solid #fbbf24" }}>
                    <textarea
                      value={s.comment}
                      onChange={(e) => updateStudentField(i, "comment", e.target.value.slice(0, COMMENT_MAX_LEN))}
                      placeholder="코멘트 입력"
                      rows={5}
                      maxLength={COMMENT_MAX_LEN}
                      style={{ width: COMMENT_ROW_W, maxWidth: COMMENT_ROW_W, boxSizing: "border-box", padding: "6px 6px", fontSize: 11, lineHeight: 1.4, border: "1px solid #fbbf24", borderRadius: 4, background: "#fffbeb", resize: "vertical", fontFamily: "inherit" }}
                    />
                    <div style={{ width: COMMENT_ROW_W, textAlign: "right", fontSize: 9, color: (s.comment || "").length >= COMMENT_MAX_LEN ? "#dc2626" : "#9ca3af", marginTop: 2 }}>
                      {(s.comment || "").length} / {COMMENT_MAX_LEN}자
                    </div>
                  </td>
                ))}
              </tr>
            </tbody>
          </table>
        </div>

        <div style={{ padding: "18px 26px", display: "flex", justifyContent: "space-between", borderTop: "1px solid #f1f5f9" }}>
          <button onClick={onBack} style={secondaryBtn}>← 이전</button>
          <button onClick={onNext} style={primaryBtn}>확인 → 최종 성적표 보기</button>
        </div>
      </div>
    </div>
  );
}

function SectionRows({ title, defs, students, update, groupLabelStyle, groupMaxStyle, rowLabelStyle, maxColStyle, dataCellStyle, emptyCellStyle, cellInput }) {
  return (
    <>
      <tr>
        <td style={groupLabelStyle}>{title}</td>
        <td style={groupMaxStyle} />
        {students.map((s, i) => <td key={s.id} style={emptyCellStyle(i, i === students.length - 1)} />)}
      </tr>
      {defs.map((d) => (
        <tr key={d.key}>
          <td style={rowLabelStyle}>{d.label} {d.kr && <span style={{ fontWeight: 400 }}>({d.kr})</span>}</td>
          <td style={maxColStyle}>{getRowMax(d.key)}</td>
          {students.map((s, i) => (
            <td key={s.id} style={dataCellStyle(i, i === students.length - 1)}>
              <input
                type="number" min={0} max={getRowMax(d.key)}
                value={s[d.key]}
                onChange={(e) => update(i, d.key, clamp(e.target.value, getRowMax(d.key)))}
                onFocus={(e) => e.target.select()}
                style={cellInput}
              />
            </td>
          ))}
        </tr>
      ))}
    </>
  );
}

const stepperBtn = {
  width: 28, height: 28, borderRadius: 6, border: "1px solid #d1d5db",
  background: "#fff", fontSize: 16, fontWeight: 700, cursor: "pointer",
};
const primaryBtn = {
  padding: "12px 26px", borderRadius: 10, border: "none",
  background: "#111827", color: "#fff", fontWeight: 700, fontSize: 14, cursor: "pointer",
};
const secondaryBtn = {
  padding: "12px 26px", borderRadius: 10, border: "1px solid #d1d5db",
  background: "#fff", color: "#374151", fontWeight: 700, fontSize: 14, cursor: "pointer",
};

// ---------- STEP 3 ----------
function computeReportData(student, totalMax, classAverages) {
  const totalGot = PART_DEFS.reduce((sum, p) => sum + (Number(student[p.key]) || 0), 0);
  const totalPct = totalMax ? (totalGot / totalMax) * 100 : 0;
  const radarData = PART_DEFS.map((p) => {
    const pct = MAX_SCORES[p.key] ? (Number(student[p.key]) / MAX_SCORES[p.key]) * 100 : 0;
    return { subject: p.label, 득점: Math.round(pct), 반평균: Math.round(classAverages[p.key]), 기준: 80 };
  });
  return { totalGot, totalPct, radarData };
}

function Step3({ form, totalMax, students, classAverages, reportIndex, setReportIndex, onBack }) {
  const [printAll, setPrintAll] = useState(false);
  const student = students[reportIndex];
  const { totalGot, totalPct, radarData } = computeReportData(student, totalMax, classAverages);

  useEffect(() => {
    if (!printAll) return;
    const t = setTimeout(() => window.print(), 60);
    const revert = () => setPrintAll(false);
    window.addEventListener("afterprint", revert, { once: true });
    return () => {
      clearTimeout(t);
      window.removeEventListener("afterprint", revert);
    };
  }, [printAll]);

  return (
    <div className={`step3-wrapper${printAll ? " printing-all" : ""}`} style={{ maxWidth: 900, margin: "0 auto", padding: "24px 20px 60px" }}>
      <div className="print-hide" style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 14 }}>
        <button onClick={onBack} style={secondaryBtn}>← 입력표 수정</button>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <button
            disabled={reportIndex === 0}
            onClick={() => setReportIndex((i) => Math.max(0, i - 1))}
            style={{ ...stepperBtn, width: 34, height: 34, opacity: reportIndex === 0 ? 0.4 : 1 }}
          >‹</button>
          <select
            value={reportIndex}
            onChange={(e) => setReportIndex(Number(e.target.value))}
            style={{ ...inputStyle, width: "auto" }}
          >
            {students.map((s, i) => (
              <option key={s.id} value={i}>{i + 1}. {s.name || "(이름 미입력)"}</option>
            ))}
          </select>
          <button
            disabled={reportIndex === students.length - 1}
            onClick={() => setReportIndex((i) => Math.min(students.length - 1, i + 1))}
            style={{ ...stepperBtn, width: 34, height: 34, opacity: reportIndex === students.length - 1 ? 0.4 : 1 }}
          >›</button>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => exportStudentsToExcel(form, students, "_성적결과")} style={secondaryBtn}>📊 엑셀 다운로드</button>
          <button onClick={() => setPrintAll(true)} style={secondaryBtn}>🖨 전체 인쇄</button>
          <button onClick={() => window.print()} style={primaryBtn}>🖨 인쇄</button>
        </div>
      </div>

      <div className="single-report">
        <ReportCard form={form} totalMax={totalMax} student={student} totalGot={totalGot} totalPct={totalPct} radarData={radarData} classAverages={classAverages} />
      </div>

      {printAll && (
        <div className="all-reports-container">
          {students.map((st) => {
            const r = computeReportData(st, totalMax, classAverages);
            return (
              <div className="print-page" key={st.id}>
                <ReportCard form={form} totalMax={totalMax} student={st} totalGot={r.totalGot} totalPct={r.totalPct} radarData={r.radarData} classAverages={classAverages} />
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function ReportCard({ form, totalMax, student, totalGot, totalPct, radarData, classAverages }) {
  return (
    <div className="report-card" style={{ background: "#fff", borderRadius: 16, overflow: "hidden", boxShadow: "0 1px 3px rgba(0,0,0,0.12)" }}>
      <div className="report-header" style={{ padding: "20px 24px 6px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <span style={{ fontSize: 26, fontWeight: 900, color: "#111827" }}>GnB</span>
          <span style={{ fontSize: 18, fontWeight: 800, color: "#111827", marginLeft: 8 }}>Monthly Report</span>
        </div>
        <div style={{ display: "flex", gap: 10 }}>
          <SignatureBox label="Director's Signature" />
          <SignatureBox label="Parents' Signature" />
        </div>
      </div>

      <InfoRow form={form} student={student} />

      <div className="report-section" style={{ margin: "14px 24px 0", background: "#fef9c3", border: "1px solid #fde68a", borderRadius: 8, padding: "10px 16px", fontWeight: 800, color: "#78350f", fontSize: 14 }}>
        {form.textbook}
      </div>

      <div className="report-section" style={{ margin: "22px 24px 0" }}>
        <ScoreTable totalMax={totalMax} student={student} totalGot={totalGot} totalPct={totalPct} />
      </div>

      <div className="report-section" style={{ margin: "22px 24px 0" }}>
        <SectionHeader icon="📊" title="Test Result in Graph Form" />
        <div style={{ display: "flex", gap: 14, marginTop: 10, flexWrap: "wrap" }}>
          <div className="chart-box" style={{ flex: "1 1 380px", height: 220, background: "#f9fafb", borderRadius: 10, padding: 8 }}>
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={radarData} margin={{ top: 18, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} />
                <XAxis dataKey="subject" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 10 }} domain={[0, 100]} />
                <Tooltip />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="득점" fill="#16a34a" radius={[3, 3, 0, 0]}>
                  <LabelList dataKey="득점" position="top" style={{ fontSize: 10, fontWeight: 700, fill: "#166534" }} />
                </Bar>
                <Bar dataKey="반평균" fill="#2563eb" radius={[3, 3, 0, 0]}>
                  <LabelList dataKey="반평균" position="top" style={{ fontSize: 10, fontWeight: 700, fill: "#1e40af" }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="radar-box" style={{ flex: "1 1 260px", height: 220, background: "#f9fafb", borderRadius: 10, padding: 8 }}>
            <ResponsiveContainer width="100%" height="100%">
              <RadarChart data={radarData} outerRadius="72%">
                <PolarGrid />
                <PolarAngleAxis dataKey="subject" tick={{ fontSize: 10 }} />
                <PolarRadiusAxis domain={[0, 100]} tick={{ fontSize: 8 }} />
                <Radar name="기준점수" dataKey="기준" stroke="#9ca3af" fill="#9ca3af" fillOpacity={0.05} strokeDasharray="4 3" />
                <Radar name="반평균" dataKey="반평균" stroke="#2563eb" fill="#2563eb" fillOpacity={0.15} />
                <Radar name="득점" dataKey="득점" stroke="#16a34a" fill="#16a34a" fillOpacity={0.25} />
                <Legend wrapperStyle={{ fontSize: 10 }} />
              </RadarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      <div className="report-section" style={{ margin: "22px 24px 0" }}>
        <SectionHeader icon="📋" title="Class Performance" />
        <PerformanceTable student={student} />
      </div>

      <div className="report-section" style={{ margin: "22px 24px 24px" }}>
        <SectionHeader icon="📝" title="Teacher's Comments" />
        <div className="comments-box" style={{ marginTop: 8, minHeight: 60, border: "1px solid #e5e7eb", borderRadius: 10, padding: "12px 16px", fontSize: 14, color: "#111827", background: "#fafafa" }}>
          {student.comment || <span style={{ color: "#9ca3af" }}>입력된 코멘트가 없습니다.</span>}
        </div>
      </div>
    </div>
  );
}

function SignatureBox({ label }) {
  return (
    <div style={{ width: 90, textAlign: "center" }}>
      <div style={{ height: 36, border: "1px solid #d1d5db", borderRadius: 6, background: "#f9fafb" }} />
      <div style={{ fontSize: 9, color: "#6b7280", marginTop: 3 }}>{label}</div>
    </div>
  );
}

function InfoRow({ form, student }) {
  const cell = { padding: "8px 14px", fontSize: 12, borderRight: "1px solid #e5e7eb" };
  const label = { fontWeight: 700, color: "#6b7280", marginRight: 6 };
  return (
    <div style={{ margin: "6px 24px 0", border: "1px solid #e5e7eb", borderRadius: 8, display: "flex", flexWrap: "wrap", overflow: "hidden" }}>
      <div style={cell}><span style={label}>Date</span>{form.dateStart} ~ {form.dateEnd}</div>
      <div style={cell}><span style={label}>Teacher's Name</span>{form.teacher}</div>
      <div style={cell}><span style={label}>Class</span>{form.className}</div>
      <div style={{ ...cell, borderRight: "none" }}><span style={label}>Student's Name</span>{student.name || "-"}</div>
    </div>
  );
}

function ScoreTable({ totalMax, student, totalGot, totalPct }) {
  const th = { padding: "8px 10px", fontSize: 12, color: "#fff", textAlign: "center" };
  const td = { padding: "8px 10px", fontSize: 13, textAlign: "center", borderBottom: "1px solid #f1f5f9" };
  return (
    <table className="score-table" style={{ width: "100%", borderCollapse: "collapse" }}>
      <thead>
        <tr style={{ background: "#111827" }}>
          <th style={{ ...th, width: "14%" }}>영역</th>
          <th style={th}>문제유형</th>
          <th style={{ ...th, width: "16%" }}>득점/문항수</th>
          <th style={{ ...th, width: "10%" }}>점수</th>
          <th style={{ ...th, width: "16%" }}>평가</th>
        </tr>
      </thead>
      <tbody>
        {PART_DEFS.map((p, i) => {
          const got = Number(student[p.key]) || 0;
          const max = MAX_SCORES[p.key] || 1;
          const pct = Math.round((got / max) * 100);
          const g = grade100(pct);
          return (
            <tr key={p.key} style={{ background: i % 2 ? "#fafafa" : "#fff" }}>
              <td style={{ ...td, fontWeight: 700 }}>{p.label}</td>
              <td style={td}>{p.kr}</td>
              <td style={td}>{got} / {max}</td>
              <td style={{ ...td, fontWeight: 700 }}>{pct}</td>
              <td style={td}>
                <span className="grade-badge" style={gradeBadgeStyle(g)}>{g.label}</span>
              </td>
            </tr>
          );
        })}
        <tr style={{ background: "#fef3c7" }}>
          <td colSpan={2} style={{ ...td, fontWeight: 800, textAlign: "center" }}>Total</td>
          <td style={{ ...td, fontWeight: 800 }}>{totalGot} / {totalMax}</td>
          <td style={{ ...td, fontWeight: 800 }}>{Math.round(totalPct)}</td>
          <td style={td}>
            {(() => {
              const g = grade100(Math.round(totalPct));
              return <span className="grade-badge" style={gradeBadgeStyle(g)}>{g.label}</span>;
            })()}
          </td>
        </tr>
      </tbody>
      <caption style={{ captionSide: "bottom", textAlign: "left", fontSize: 10, color: "#9ca3af", padding: "6px 2px" }}>
        * 평가 안내: Perfect(100), Excellent(90~99), Very Good(80~89), Good(70~79), Not Bad(60~69), Practice More(59 이하)
      </caption>
    </table>
  );
}

function PerformanceTable({ student }) {
  const groups = [
    { title: "참여도", defs: PARTICIPATION_DEFS, color: "#8b5cf6" },
    { title: "태도", defs: BEHAVIOR_DEFS, color: "#0d9488" },
    { title: "숙제", defs: HOMEWORK_DEFS, color: "#f59e0b" },
  ];
  const td = { padding: "7px 10px", fontSize: 12, borderBottom: "1px solid #f1f5f9", verticalAlign: "middle", textAlign: "center" };
  return (
    <table className="perf-table" style={{ width: "100%", borderCollapse: "collapse", marginTop: 8 }}>
      <thead>
        <tr style={{ background: "#1f2937" }}>
          <th style={{ ...td, color: "#fff", width: "12%" }}>항목</th>
          <th style={{ ...td, color: "#fff" }}>세부항목</th>
          <th style={{ ...td, color: "#fff", width: "40%" }}>성취도</th>
          <th style={{ ...td, color: "#fff", width: "16%" }}>평가</th>
        </tr>
      </thead>
      <tbody>
        {groups.map((g) =>
          g.defs.map((d, di) => {
            const v = Number(student[d.key]) || 0;
            const grd = grade10(v);
            return (
              <tr key={d.key}>
                {di === 0 && (
                  <td style={{ ...td, fontWeight: 700, textAlign: "center", verticalAlign: "middle", background: "#f3f4f6" }} rowSpan={g.defs.length}>{g.title}</td>
                )}
                <td style={td}>{d.label} {d.kr && `(${d.kr})`}</td>
                <td style={td}>
                  <div style={{ background: "#f1f5f9", borderRadius: 6, height: 14, position: "relative" }}>
                    <div style={{ width: `${Math.min(100, (v / 10) * 100)}%`, background: g.color, height: 14, borderRadius: 6 }} />
                    <span style={{ position: "absolute", right: 6, top: -1, fontSize: 10, fontWeight: 700, color: "#374151" }}>{v}</span>
                  </div>
                </td>
                <td style={td}>
                  <span className="grade-badge" style={gradeBadgeStyle(grd)}>{grd.label}</span>
                </td>
              </tr>
            );
          })
        )}
      </tbody>
    </table>
  );
}

function SectionHeader({ icon, title }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, borderBottom: "2px solid #111827", paddingBottom: 6 }}>
      <span>{icon}</span>
      <span style={{ fontWeight: 800, fontSize: 15, color: "#111827" }}>{title}</span>
    </div>
  );
}
