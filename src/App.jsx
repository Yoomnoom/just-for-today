import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  Home as HomeIcon,
  Camera,
  BarChart3,
  Settings as SettingsIcon,
  ChevronLeft,
  Eye,
  EyeOff,
  Dumbbell,
  Check,
  Plus,
  X,
  Sun,
  Moon,
  Type,
  Aperture,
  Download,
  Info,
  Loader2,
} from "lucide-react";
import {
  LineChart,
  Line,
  BarChart,
  Bar,
  ComposedChart,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer,
  CartesianGrid,
} from "recharts";

/* ------------------------------------------------------------------ */
/*  상수 & 유틸                                                        */
/* ------------------------------------------------------------------ */

const STORAGE_KEY = "ifasting_app_v1";
const FASTING_HOURS = 16;
const MEAL_HOURS = 7;

// 노란 텍스트를 얹는 버튼·토스트는 라이트/다크 모드와 무관하게 항상 이 어두운 색을 배경으로 써서
// 다크모드에서 배경(c.ink)이 밝은 색으로 뒤집혀 노란 글씨가 안 보이는 문제를 막는다.
const BRAND_DARK = "#2A1B10";

const COLORS = {
  light: {
    bg: "#FFFDF7",
    card: "#FFFFFF",
    cardMuted: "#F6F4EC",
    ink: "#3A2317",
    inkSoft: "#8A7D6E",
    yellow: "#FEE500",
    yellowSoft: "#FFF4B8",
    line: "#EFE9D8",
    green: "#3FAE7A",
    danger: "#E2725B",
  },
  dark: {
    bg: "#1B1712",
    card: "#262019",
    cardMuted: "#332B20",
    ink: "#F5EFE3",
    inkSoft: "#B4A891",
    yellow: "#FEE500",
    yellowSoft: "#4A431F",
    line: "#3A3122",
    green: "#5FCB98",
    danger: "#F08C74",
  },
};

const FOOD_TEMPLATES = [
  { text: "탄수화물이 조금 많아요. 다음 식사엔 단백질을 조금 더 추가해보세요.", tag: "탄수화물 많음" },
  { text: "단백질과 채소가 충분해서 균형 잡힌 식사예요.", tag: "균형 좋음" },
  { text: "비타민이 부족해 보여요. 채소를 곁들이면 더 좋아요.", tag: "비타민 부족" },
  { text: "가공식품 비율이 높아요. 다음 끼니는 신선한 재료로 가볍게 가볼까요?", tag: "가공식품 많음" },
  { text: "지방과 탄수화물이 적절해요. 단백질만 조금 더하면 완벽해요.", tag: "단백질 보완" },
];

const EXERCISE_TYPES = ["걷기", "달리기", "헬스", "수영", "요가", "홈트"];

function resizeImage(dataUrl, maxDim = 1024, quality = 0.72) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      let { width, height } = img;
      if (width > height && width > maxDim) {
        height = Math.round((height * maxDim) / width);
        width = maxDim;
      } else if (height > maxDim) {
        width = Math.round((width * maxDim) / height);
        height = maxDim;
      }
      const canvas = document.createElement("canvas");
      canvas.width = width;
      canvas.height = height;
      canvas.getContext("2d").drawImage(img, 0, 0, width, height);
      resolve(canvas.toDataURL("image/jpeg", quality));
    };
    img.onerror = () => resolve(dataUrl);
    img.src = dataUrl;
  });
}

function pad(n) {
  return String(n).padStart(2, "0");
}
function todayStr(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function timeStr(d = new Date()) {
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function addDays(d, n) {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
}
function diffDaysStr(fromStr, toStr) {
  const a = new Date(fromStr + "T00:00:00");
  const b = new Date(toStr + "T00:00:00");
  return Math.round((b - a) / 86400000);
}
function getMealWindow(baseDate, mealStart) {
  const [h, m] = mealStart.split(":").map(Number);
  const start = new Date(baseDate);
  start.setHours(h, m, 0, 0);
  const end = new Date(start.getTime() + MEAL_HOURS * 3600 * 1000);
  return { start, end };
}
function getPhase(now, mealStart) {
  const todayW = getMealWindow(now, mealStart);
  const yestW = getMealWindow(addDays(now, -1), mealStart);
  if (now >= todayW.start && now < todayW.end) return { phase: "meal", window: todayW };
  if (now >= yestW.start && now < yestW.end) return { phase: "meal", window: yestW };
  let nextStart = todayW.start;
  if (now >= todayW.start) nextStart = getMealWindow(addDays(now, 1), mealStart).start;
  return { phase: "fasting", nextStart };
}
const MEAL_GRACE_MIN = 30;
/** 기록된 날짜·시간이 식사 가능 시간(앞뒤 30분 허용)에 해당하는지 판단.
 *  실시간 상태가 아니라, 사용자가 입력/수정한 시간 기준으로 계산한다. */
function isMealTimeKept(dateStr, timeStr, mealStart, graceMin = MEAL_GRACE_MIN) {
  const entryDate = new Date(`${dateStr}T${timeStr}:00`);
  const graceMs = graceMin * 60 * 1000;
  const inRange = (w) => entryDate >= new Date(w.start.getTime() - graceMs) && entryDate <= new Date(w.end.getTime() + graceMs);
  return inRange(getMealWindow(entryDate, mealStart)) || inRange(getMealWindow(addDays(entryDate, -1), mealStart));
}
/** 날짜별로 "직전 기록 대비 늘었는지/줄었는지" 방향만 계산한다. 실제 체중 수치가 아니라
 *  방향(▲/▼/-) 정보라서, 체중 가림 설정과 무관하게 항상 보여줘도 되는 정보로 취급한다. */
function buildWeightDirMap(weightLogs) {
  const sorted = weightLogs.slice().sort((a, b) => (a.date < b.date ? -1 : 1));
  const map = {};
  sorted.forEach((w, i) => {
    if (i === 0) {
      map[w.date] = { dir: "first" };
      return;
    }
    const prev = sorted[i - 1];
    const diff = Math.round((w.value - prev.value) * 10) / 10;
    const dir = diff > 0 ? "up" : diff < 0 ? "down" : "flat";
    map[w.date] = { dir, diff: Math.abs(diff) };
  });
  return map;
}
/** 실제 체중 수치가 아니라 "직전 기록 대비 몇 kg 늘었는지/줄었는지"만 담고 있어서,
 *  체중 가림 설정과 무관하게 항상 보여줘도 되는 정보로 취급한다. */
function WeightDirBadge({ info, c, size = 10, compact = false }) {
  if (!info) return null;
  if (info.dir === "first") {
    if (compact) return <span style={{ fontSize: size, color: c.inkSoft }}>●</span>;
    return (
      <span style={{ fontSize: size, color: c.inkSoft, fontWeight: 700, whiteSpace: "nowrap" }}>
        첫 기록이에요
      </span>
    );
  }
  const cfg = {
    up: { symbol: "▲", color: c.danger },
    down: { symbol: "▼", color: c.green },
    flat: { symbol: "－", color: c.inkSoft },
  }[info.dir];
  return (
    <span style={{ fontSize: size, color: cfg.color, fontWeight: 800, whiteSpace: "nowrap" }}>
      {cfg.symbol}{!compact && `${info.diff.toFixed(1)}kg`}
    </span>
  );
}

const FOOD_TAG_MACROS = {
  "탄수화물 많음": { 탄수화물: "높음", 단백질: "보통", 지방: "보통", "채소/비타민": "보통", score: 3 },
  "균형 좋음": { 탄수화물: "보통", 단백질: "높음", 지방: "보통", "채소/비타민": "높음", score: 5 },
  "비타민 부족": { 탄수화물: "보통", 단백질: "보통", 지방: "보통", "채소/비타민": "낮음", score: 3 },
  "가공식품 많음": { 탄수화물: "높음", 단백질: "낮음", 지방: "높음", "채소/비타민": "낮음", score: 2 },
  "단백질 보완": { 탄수화물: "보통", 단백질: "낮음", 지방: "보통", "채소/비타민": "보통", score: 3 },
};

function sparkline(values) {
  if (!values || values.length === 0) return "(데이터 없음)";
  const blocks = "▁▂▃▄▅▆▇█";
  const min = Math.min(...values);
  const max = Math.max(...values);
  const span = max - min || 1;
  return values.map((v) => blocks[Math.min(7, Math.max(0, Math.floor(((v - min) / span) * 7)))]).join("");
}

// 엑셀 REPT() 함수로 막대를 그 자리에서 계산하게 만드는 "살아있는" 수식 셀
function fmlBar(fracExpr) {
  const f = `IFERROR(${fracExpr},0)`;
  return { t: "str", f: `REPT("█",ROUND((${f})*10,0))&REPT("░",10-ROUND((${f})*10,0))&" "&TEXT((${f}),"0%")` };
}
function fmlNum(expr) {
  return { t: "n", f: expr };
}

/** 음식/체중/운동 기록을 바탕으로, 병원·트레이너·영양사에게 제출할 수 있는 수준의
 *  "건강 리포트" 형태 엑셀 워크북(6개 시트)을 만들어 blob URL을 반환한다.
 *
 *  설계 원칙: 대시보드(1번 시트)의 숫자·막대는 전부 엑셀 수식(AVERAGE, COUNTIF,
 *  REPT 등)으로 다른 시트를 직접 참조한다. 그래서 나중에 사용자가 체중/음식
 *  시트의 원본 데이터를 고쳐도 대시보드가 자동으로 다시 계산된다.
 *
 *  기술적 한계: 여기서 쓰는 xlsx 라이브러리(SheetJS 무료판)는 셀 배경색·글꼴
 *  스타일, 틀고정, 실제 도형 차트 삽입을 지원하지 않는다(유료판 전용 기능).
 *  그 대신 REPT() 수식 막대와 텍스트 스파크라인, 자동필터로 가독성을 높였다.
 *  range = { start: 'YYYY-MM-DD'|null, end: 'YYYY-MM-DD'|null } — null이면 전체 기간 */
function exportToExcel(data, range = null) {
  const inRange = (d) => !range || ((!range.start || d >= range.start) && (!range.end || d <= range.end));
  const dirMap = buildWeightDirMap(data.logs.weight);
  const today = todayStr();

  const foodInRange = data.logs.food.filter((f) => inRange(f.date)).sort((a, b) => (a.date + a.time < b.date + b.time ? -1 : 1));
  const weightInRange = data.logs.weight.filter((w) => inRange(w.date)).sort((a, b) => (a.date < b.date ? -1 : 1));
  const exerciseInRange = data.logs.exercise.filter((e) => inRange(e.date)).sort((a, b) => (a.date < b.date ? -1 : 1));

  // ---- 기간 전체 날짜 목록 (캘린더 시트 & 기록률 계산용) ----
  const periodStart = range?.start || (weightInRange[0]?.date || foodInRange[0]?.date || today);
  const periodEnd = range?.end || today;
  const allDates = [];
  for (let d = periodStart; d <= periodEnd && allDates.length < 3660; d = todayStr(addDays(new Date(d + "T00:00:00"), 1))) {
    allDates.push(d);
  }

  const logsByDate = {};
  const ensure = (d) => (logsByDate[d] ||= { food: [], weight: [], exercise: [] });
  foodInRange.forEach((f) => ensure(f.date).food.push(f));
  weightInRange.forEach((w) => ensure(w.date).weight.push(w));
  exerciseInRange.forEach((e) => ensure(e.date).exercise.push(e));

  const recentWeights = weightInRange.map((w) => w.value).slice(-60);
  const dailyFoodCounts = allDates.slice(-60).map((d) => logsByDate[d]?.food.length || 0);

  // ---- Sheet 2: 체중 기록 (행 4부터 데이터 시작) ----
  const s2Header = ["날짜", "체중(kg)", "직전 대비 증감", "사진 유무"];
  const s2Rows = weightInRange.map((w) => {
    const info = dirMap[w.date];
    const change = !info ? "" : info.dir === "first" ? "첫 기록" : `${info.dir === "up" ? "+" : info.dir === "down" ? "-" : "±"}${info.diff.toFixed(1)}kg`;
    return [w.date, w.value, change, w.image ? "있음" : "없음"];
  });
  const ws2 = XLSX.utils.aoa_to_sheet([
    ["체중 기록"],
    [],
    s2Header,
    ...(s2Rows.length ? s2Rows : [["기록 없음", "", "", ""]]),
  ]);
  ws2["!cols"] = [{ wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 10 }];
  if (s2Rows.length) {
    const lastRow = 3 + s2Rows.length;
    ws2["!autofilter"] = { ref: `A3:D${lastRow}` };
    for (let r = 4; r <= lastRow; r++) {
      const cell = ws2[`B${r}`];
      if (cell) cell.z = "0.0";
    }
  }

  // ---- Sheet 3: 음식 기록 (K열에 AI점수 숫자 컬럼 추가 — 대시보드 평균 계산용) ----
  const s3Header = ["날짜", "시간", "식사시간 준수", "AI 분석", "탄수화물", "단백질", "지방", "채소/비타민", "메모", "사진 유무", "AI점수(5점)"];
  const s3Rows = foodInRange.map((f) => {
    const macro = FOOD_TAG_MACROS[f.tag] || {};
    return [
      f.date,
      f.time,
      f.inMealWindow ? "Y" : "N",
      f.tag || "",
      macro.탄수화물 || "-",
      macro.단백질 || "-",
      macro.지방 || "-",
      macro["채소/비타민"] || "-",
      f.note || "",
      f.image ? "있음" : "없음",
      macro.score || "",
    ];
  });
  const ws3 = XLSX.utils.aoa_to_sheet([
    ["음식 기록"],
    [],
    s3Header,
    ...(s3Rows.length ? s3Rows : [["기록 없음"]]),
  ]);
  ws3["!cols"] = [{ wch: 12 }, { wch: 8 }, { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 24 }, { wch: 10 }, { wch: 12 }];
  if (s3Rows.length) ws3["!autofilter"] = { ref: `A3:K${3 + s3Rows.length}` };

  // ---- Sheet 4: 운동 기록 (다중 선택 시 종류별로 행 분리) ----
  const s4Header = ["날짜", "운동 종류", "운동 시간(분)", "강도", "메모"];
  const s4Rows = [];
  exerciseInRange.forEach((e) => {
    const types = e.types && e.types.length ? e.types : [e.type || "기타"];
    types.forEach((t) => s4Rows.push([e.date, t, e.minutes, e.intensity || "", e.note || ""]));
  });
  const ws4 = XLSX.utils.aoa_to_sheet([
    ["운동 기록"],
    [],
    s4Header,
    ...(s4Rows.length ? s4Rows : [["기록 없음"]]),
  ]);
  ws4["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 10 }, { wch: 24 }];
  if (s4Rows.length) ws4["!autofilter"] = { ref: `A3:E${3 + s4Rows.length}` };

  // ---- Sheet 5: 루틴 캘린더 (하루 한눈에 보기) ----
  const s5Header = ["날짜", "음식", "체중", "운동", "변화 기록"];
  const s5Rows = allDates.map((d) => {
    const info = logsByDate[d];
    return [d, info?.food.length ? "📷" : "", info?.weight.length ? "⚖️" : "", info?.exercise.length ? "🏃" : "", ""];
  });
  const ws5 = XLSX.utils.aoa_to_sheet([
    ["루틴 캘린더 — 이 시트만 봐도 하루하루의 생활 패턴이 보여요"],
    [],
    s5Header,
    ...s5Rows,
  ]);
  ws5["!cols"] = [{ wch: 14 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 10 }];
  if (s5Rows.length) ws5["!autofilter"] = { ref: `A3:E${3 + s5Rows.length}` };

  // ---- Sheet 1: 건강 리포트 (대시보드) — 아래 숫자·막대는 전부 다른 시트를 참조하는 엑셀 수식 ----
  const wLast = 3 + s2Rows.length;
  const fLast = 3 + s3Rows.length;
  const eLast = 3 + s4Rows.length;
  const cLast = 3 + s5Rows.length;
  const wRef = (col) => `'체중 기록'!${col}4:${col}${wLast}`;
  const fRef = (col) => `'음식 기록'!${col}4:${col}${fLast}`;
  const eRef = (col) => `'운동 기록'!${col}4:${col}${eLast}`;
  const cRef = (col) => `'루틴 캘린더'!${col}4:${col}${cLast}`;

  const s1 = [
    ["오늘부터 — 건강 리포트"],
    [],
    ["기록 기간", `${periodStart} ~ ${periodEnd}`],
    ["내보낸 날짜", today],
    [],
    ["■ 체중 요약"],
  ];
  if (s2Rows.length) {
    s1.push(["평균 체중(kg)", fmlNum(`ROUND(AVERAGE(${wRef("B")}),1)`)]);
    s1.push(["최저 체중(kg)", fmlNum(`ROUND(MIN(${wRef("B")}),1)`)]);
    s1.push(["최고 체중(kg)", fmlNum(`ROUND(MAX(${wRef("B")}),1)`)]);
    s1.push([
      "총 변화량(kg)",
      s2Rows.length > 1
        ? fmlNum(`ROUND(INDEX(${wRef("B")},COUNT(${wRef("B")}))-INDEX(${wRef("B")},1),1)`)
        : "0 (기록 1건)",
    ]);
  } else {
    s1.push(["평균 체중(kg)", "-"], ["최저 체중(kg)", "-"], ["최고 체중(kg)", "-"], ["총 변화량(kg)", "-"]);
  }
  s1.push(
    [],
    ["■ 루틴 요약"],
    ["설정된 공복시간", "16시간 (앱 설정 기준)"],
    ["식사시간 준수율", s3Rows.length ? fmlBar(`COUNTIF(${fRef("C")},"Y")/COUNTA(${fRef("A")})`) : "- (기록 없음)"],
    ["이 기간 기록률", s5Rows.length ? fmlBar(`SUMPRODUCT(--((${cRef("B")}<>"")+(${cRef("C")}<>"")>0))/COUNTA(${cRef("A")})`) : "- (기록 없음)"],
    [],
    ["■ 활동 요약"],
    ["음식 기록 수", s3Rows.length ? fmlNum(`COUNTA(${fRef("A")})`) : 0],
    ["운동 기록 수", s4Rows.length ? fmlNum(`COUNTA(${eRef("A")})`) : 0],
    ["체중 기록 수", s2Rows.length ? fmlNum(`COUNTA(${wRef("A")})`) : 0],
    ["변화 기록(몸사진) 수", "0 (준비 중인 기능)"],
    ["AI 영양 점수 평균(5점 만점)", s3Rows.length ? fmlNum(`ROUND(AVERAGE(${fRef("K")}),1)`) : "-"],
    ["총 획득 포인트(전체 누적)", data.points],
    [],
    ["■ 체중 변화 추이 (최근 60건, 텍스트 그래프)"],
    ["", sparkline(recentWeights)],
    ["■ 음식 기록 추이 (최근 60일, 하루 기록 수)"],
    ["", sparkline(dailyFoodCounts)],
    [],
    ["※ 위 숫자·막대는 대부분 실제 엑셀 함수(AVERAGE·COUNTIF·REPT 등)라서, 체중/음식 시트의 원본 값을 고치면 자동으로 다시 계산돼요."],
    ["※ 다만 이 라이브러리는 셀 배경색·틀고정·실제 도형 차트는 지원하지 않아, 그래프는 텍스트 막대로 대신했어요."],
    ["※ 정식 차트가 필요하면 해당 시트에서 데이터를 선택한 뒤 엑셀 메뉴의 '삽입 > 차트'를 눌러 직접 만드실 수 있어요."]
  );
  const ws1 = XLSX.utils.aoa_to_sheet(s1);
  ws1["!cols"] = [{ wch: 30 }, { wch: 45 }, { wch: 20 }];

  // ---- Sheet 6: AI 영양 분석 (기간 평균) ----
  const tagCounts = {};
  foodInRange.forEach((f) => { if (f.tag) tagCounts[f.tag] = (tagCounts[f.tag] || 0) + 1; });
  const total = foodInRange.length || 1;
  const scoreOf = (key) => {
    // 각 매크로별 "높음/보통/낮음" 비율을 5점 척도로 환산 (간이 추정치)
    const levels = foodInRange.map((f) => FOOD_TAG_MACROS[f.tag]?.[key]).filter(Boolean);
    if (!levels.length) return 3;
    const score = (l) => (l === "높음" ? 5 : l === "보통" ? 3.5 : 2);
    return Math.round((levels.reduce((s, l) => s + score(l), 0) / levels.length));
  };
  const proteinStars = "★".repeat(Math.max(1, Math.min(5, scoreOf("단백질")))) + "☆".repeat(5 - Math.max(1, Math.min(5, scoreOf("단백질"))));
  const veggieStars = "★".repeat(Math.max(1, Math.min(5, scoreOf("채소/비타민")))) + "☆".repeat(5 - Math.max(1, Math.min(5, scoreOf("채소/비타민"))));
  const carbLevel = scoreOf("탄수화물"); // 높음일수록 조절이 더 필요하다는 뜻이라 반대로 해석
  const carbStars = "★".repeat(Math.max(1, Math.min(5, 6 - carbLevel))) + "☆".repeat(5 - Math.max(1, Math.min(5, 6 - carbLevel)));

  const comments = [];
  if (foodInRange.length === 0) {
    comments.push("이 기간엔 음식 기록이 없어 분석할 내용이 없어요.");
  } else {
    if ((tagCounts["균형 좋음"] || 0) / total >= 0.3) comments.push("균형 잡힌 식사가 꾸준히 이어지고 있어요.");
    if ((tagCounts["단백질 보완"] || 0) / total >= 0.25) comments.push("단백질 섭취를 조금 더 늘리면 더 좋아질 거예요.");
    if ((tagCounts["비타민 부족"] || 0) / total >= 0.25) comments.push("채소·비타민 섭취를 더해보면 좋아요.");
    if ((tagCounts["가공식품 많음"] || 0) / total >= 0.25) comments.push("신선한 재료 위주의 식사 비중을 늘려보면 좋아요.");
    if ((tagCounts["탄수화물 많음"] || 0) / total >= 0.3) comments.push("탄수화물 비중이 조금 높은 편이니, 단백질과 채소를 더해 균형을 맞춰보세요.");
    if (comments.length === 0) comments.push("전반적으로 무난한 식습관을 유지하고 있어요.");
  }

  const ws6 = XLSX.utils.aoa_to_sheet([
    ["AI 영양 분석 (기간 평균)"],
    [],
    ["단백질", proteinStars],
    ["채소/비타민", veggieStars],
    ["탄수화물 조절", carbStars],
    [],
    ["AI 종합 코멘트"],
    ...comments.map((c) => [`· ${c}`]),
    [],
    ["※ 이 분석은 참고용 추정치이며, 비난이 아닌 방향 제안을 목적으로 항상 긍정적으로 작성됩니다."],
  ]);
  ws6["!cols"] = [{ wch: 16 }, { wch: 50 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "건강 리포트");
  XLSX.utils.book_append_sheet(wb, ws2, "체중 기록");
  XLSX.utils.book_append_sheet(wb, ws3, "음식 기록");
  XLSX.utils.book_append_sheet(wb, ws4, "운동 기록");
  XLSX.utils.book_append_sheet(wb, ws5, "루틴 캘린더");
  XLSX.utils.book_append_sheet(wb, ws6, "AI 영양 분석");

  const rangeSuffix = range ? `_${(range.start || "처음").replace(/-/g, "")}-${(range.end || today).replace(/-/g, "")}` : "_전체";
  const filename = `오늘부터_건강리포트${rangeSuffix}.xlsx`;

  try {
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    // 여기서 클릭을 자동으로 실행하지 않는다. 미리보기 환경(iframe)에 따라
    // 자바스크립트가 만든 클릭은 조용히 막히는 경우가 있어서, 대신 실제 화면에
    // <a> 링크를 띄우고 사용자가 직접 눌러야 다운로드/새 탭 열기가 확실히 동작한다.
    return { url, filename };
  } catch (e) {
    console.error("엑셀 파일 생성 실패", e);
    return null;
  }
}

function formatDur(ms) {
  if (ms < 0) ms = 0;
  const totalMin = Math.floor(ms / 60000);
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${h}시간 ${pad(m)}분`;
}
function fmtHM(d) {
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function uid() {
  return Math.random().toString(36).slice(2, 10);
}

/** 화면이 마운트되어 있는 동안에만 1초마다 갱신되는 현재 시각. 다른 화면(설정 등)의
 *  불필요한 리렌더링을 막기 위해 App 전역이 아니라 필요한 화면에서만 사용한다. */
function useNow(intervalMs = 1000) {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), intervalMs);
    return () => clearInterval(t);
  }, [intervalMs]);
  return now;
}

function defaultData() {
  return {
    onboarded: false,
    mealStart: "12:00",
    gender: "female",
    weightHidden: true,
    darkMode: false,
    fontScale: 1,
    points: 0,
    lastActiveDate: null,
    logs: { food: [], weight: [], exercise: [] },
  };
}

/* ------------------------------------------------------------------ */
/*  스토리지 (브라우저 localStorage — 실제 배포용)                        */
/* ------------------------------------------------------------------ */

async function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return { ...defaultData(), ...JSON.parse(raw) };
  } catch (e) {
    /* 최초 사용자: 키 없음, 또는 저장된 값이 손상됨 */
  }
  return defaultData();
}
async function saveData(data) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    console.error("저장 실패", e);
  }
}

/* ------------------------------------------------------------------ */
/*  메인 앱                                                            */
/* ------------------------------------------------------------------ */

export default function IntermittentFastingApp() {
  const [data, setData] = useState(null);
  const [screen, setScreen] = useState("home");
  const [toast, setToast] = useState(null);
  // 체중 화면과 통계 화면에서 공유하는 "가림 해제" 상태 - 한쪽에서 풀면 다른 쪽도 풀린 채로 보인다
  const [weightRevealed, setWeightRevealed] = useState(false);
  // 통계 화면 상태를 App으로 끌어올려, 화면을 나갔다 와도 유지되게 한다
  const [statsRangeKey, setStatsRangeKey] = useState("1w");
  const [statsPeriodOffset, setStatsPeriodOffset] = useState(0);
  const [statsCalendarOpen, setStatsCalendarOpen] = useState(false);
  const [statsSelectedDate, setStatsSelectedDate] = useState(null);
  const [statsCalYear, setStatsCalYear] = useState(new Date().getFullYear());
  const [statsCalMonth, setStatsCalMonth] = useState(new Date().getMonth());
  const toastTimer = useRef(null);
  const firstLoad = useRef(true);

  // 인트로: 로고·제작자·달리는 사람(로딩)을 한 화면에 함께 2.5초 정도 유지한다
  const [splashDone, setSplashDone] = useState(false);
  useEffect(() => {
    const t = setTimeout(() => setSplashDone(true), 2000);
    return () => clearTimeout(t);
  }, []);

  useEffect(() => {
    loadData().then((d) => setData(d));
  }, []);

  useEffect(() => {
    if (!data) return;
    if (firstLoad.current) {
      firstLoad.current = false;
      return;
    }
    saveData(data);
  }, [data]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    if (toastTimer.current) clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(null), 2200);
  }, []);

  const addPoints = useCallback(
    (amount, label) => {
      const today = todayStr();
      const gap = data?.lastActiveDate ? diffDaysStr(data.lastActiveDate, today) : 0;
      const returning = !!data?.lastActiveDate && gap >= 1;
      setData((prev) => ({ ...prev, points: prev.points + amount, lastActiveDate: today }));
      if (amount <= 0) {
        showToast(`✅ ${label}`);
        return;
      }
      showToast(
        returning
          ? `기록해주셔서 고마워요. 오늘도 하나의 루틴이 쌓였어요 · +${amount}P`
          : `✅ ${label} +${amount}P`
      );
    },
    [data?.lastActiveDate, showToast]
  );

  if (!splashDone || !data) {
    return <SplashScreen />;
  }

  if (!data.onboarded) {
    return (
      <Shell data={data}>
        <Onboarding
          data={data}
          onComplete={(mealStart, gender) =>
            setData((p) => ({ ...p, mealStart, gender, onboarded: true, lastActiveDate: todayStr() }))
          }
        />
      </Shell>
    );
  }

  return (
    <Shell data={data} toast={toast} footer={<BottomNav screen={screen} setScreen={setScreen} data={data} />}>
      {screen === "home" && (
        <HomeScreen data={data} setData={setData} addPoints={addPoints} goto={setScreen} />
      )}
      {screen === "recordHub" && <RecordHubScreen data={data} goto={setScreen} goBack={() => setScreen("home")} />}
      {screen === "food" && (
        <FoodScreen data={data} setData={setData} addPoints={addPoints} goBack={() => setScreen("recordHub")} goToWeight={() => setScreen("weight")} />
      )}
      {screen === "exercise" && (
        <ExerciseScreen data={data} setData={setData} addPoints={addPoints} goBack={() => setScreen("recordHub")} goToWeight={() => setScreen("weight")} />
      )}
      {screen === "weight" && (
        <WeightScreen data={data} setData={setData} addPoints={addPoints} goBack={() => setScreen("recordHub")} weightRevealed={weightRevealed} setWeightRevealed={setWeightRevealed} />
      )}
      {screen === "stats" && (
        <StatsScreen
          data={data}
          goBack={() => setScreen("home")}
          weightRevealed={weightRevealed}
          setWeightRevealed={setWeightRevealed}
          rangeKey={statsRangeKey}
          setRangeKey={setStatsRangeKey}
          periodOffset={statsPeriodOffset}
          setPeriodOffset={setStatsPeriodOffset}
          calendarOpen={statsCalendarOpen}
          setCalendarOpen={setStatsCalendarOpen}
          selectedDate={statsSelectedDate}
          setSelectedDate={setStatsSelectedDate}
          calYear={statsCalYear}
          setCalYear={setStatsCalYear}
          calMonth={statsCalMonth}
          setCalMonth={setStatsCalMonth}
        />
      )}
      {screen === "bodyPhoto" && <BodyPhotoScreen data={data} goBack={() => setScreen("recordHub")} />}
      {screen === "settings" && <SettingsScreen data={data} setData={setData} goBack={() => setScreen("home")} />}
    </Shell>
  );
}

/* ------------------------------------------------------------------ */
/*  Shell (앱 프레임 + 폰트/다크모드 적용)                               */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/*  인트로 (스플래시)                                                   */
/* ------------------------------------------------------------------ */

const JUICE_FRAMES = [
  "data:image/png;base64,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",
  "data:image/png;base64,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",
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAANwAAAC+CAYAAABeZmHQAABN40lEQVR4nO396Y8sS5LlB/5EVc3M3WO/21tyz6qsqix2s5uD+UCCBAbzNxMgQGCAmSE40zXsIQl0s5ZOJiuzMvNt996I8PDdzFRF+EHVzD3u8vJlVeZbIvzgOtyvhy9m5iYmoiJHjsARRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUccccQRRxxxxBFHHHHEEUcc8W2GfNMbcMR3D93yC4tdT10FvAsQAjK5PJ5LXwHHg3TEV0a3nZt3sLm7Y724ZRpq6spTVQ1+MkPOXxzPp98D901vwBHfHaSU8M2lOIHUd6S4I7VbuvUd7fwV/Re/srR8ad/0dn6bcbwiHfGVsLj+wswSDuV02kC35e7VS7xFqhBou47p2SWrrqc5u+Lk+Q+P59Y7cPRwR3wlpO0W6yOiib7dgiYuLs+oK0/wxqR2aLdFNLJd3hEXXxw93TsQvukNOOLbj+43/2gxKT4IqtCtt0RJeHpS6kltRERY75b0VMzXLd5X3/RmfytxNLgjfi+2Lz9nenFJVc/Y9i19v8OJst4sEW2Juy3OOdbbji46Xs9XvPj4B9/0Zn8rcQwpj/hS2O/+0e7+6VfsXn0G7RZpt1jfEhwYiZQS27ZlvdoSY2S5XKKqnJ2ef9Ob/q3E0cMd8eX47BO66y/YOOPs6QWVODZ9R0dEVTEMRehjh+HZdj1Xzz6kfv79Y9LkHTh6uCPeC1te2+vf/IoTSfSLW7rXr9B2g2hivV7Tti19FwFQg13bk1T4wQ9/8g1v+bcXR4M74v1YrVi8fs2z2Rl+s+Xu17+iv35NnTokdZgl+r5HkxHwxDbiXcXVT39+9G7vwTGkPOK9+Pwf/o5TN4V15JkKy1evsKCIe0qbIjqZor2iO0W6BLvE848+/KY3+1uNo4c74p2wz//RNotbTp8/o3MB6RMnauw+/5z+8y84855+scQng6SQIHaR733vmJ38Mhw93BHvhKbE+YvnNFXFRnesNwsm0ZGWO/r0mtn0lDolrA6kpLSdQl1z+r3vfdOb/q3G0cMd8U50bWT25BK9OqH5/gfo6QnbaEzdBL3bsvz1p1wkR9i2BIXFZs0HP/gBcnLsGvgyHA3uiHeiT5FUV+xqjzy7YPrRh+xCjRKootB9+gr/+pbJrsX1PZvdlh/+xV9805v9rcfR4L7lmM/nNp/Pv1Zeol2/MiqPm9S0zuiCZ/LxB/jnT1mqgqsJyVj89lNksWF3e8dsekL44EhY/n04ruG+Jtze3lqMkZTSeDMzRASzd9uTiLBerwH49NNPTUTG54f7GOO99xx+lgnsYovDU0kgOA/e4b3HY4gzqrrGzLi4uBIAe3lt3fIWNSMZeAn02rGbTmh++DGb1ZrV61fUXaBdb9m+umbtPH/+f/u///EP2gPE0eD+SHj9+rX1fZ/rUqqZhWGGcw4RYblc5hPde+q6Hp8fbpeXf/q1z/zmzroUiTESY0I10i8W7NoNv/zf/95OnSPtVpAS4aShs5bgA13bs9DI2QfPmSwXvLx9xZkYdTVhvlyzqQLf+4u/+lNv/oPA0eD+QLx+/dratqXrutFLOedYLpeEEKiqajQsEcE597UY01fB5ZOLd27H3fLGtN3i2x3L62tmswkg9CkyraZs04ZWFTNonj/n9PsfsfzF/85JqLi5ueEH/5f/K0ymX/PefDdxNLjfg08++cRUdQwD5/M53ntCCDRNQwj5EA6e6tBzQQ7xbm9v37sGG0LA4fW/D296w3et7/5QA784eyK6vbaUOj67fsXlk5+izpjUDbHrCSHQm7FdrXHOmD17Srq749VnX9CdnfC9n/8cuXj6rbiofNtxNLg3cH17Y7vdju12m7mCfY+I4L2nqiqcc6MHc26fc7q3djJ777rszee/qqEdvv/29nZcz31V47q7u7WLiyuZzxcGIOZQM6JrqSQhKMsvPmP58jP43oc4C5hTYlIgnyi7roek1LMZ0w8/YH59g784Z/qzI3fyq+JocMCr69d2d3fHZrPhk08+oWkaJpMJs9kMS3rP2ERkXKMdJizeldAA7hkl/OHe54+FISlyeXku6/nKEg7nLF88Ug9dz81vfsMMw+Y31M1T1IxdUvrYQ0x4U2LfErueNJnB5SVXf/5npLriN7/6O1NXc372hCdPnhy93XvwaA3u9m5u6/Wa6+trvvjiC6qq4vz8fDQsMyOllDN7QIzxnjEdho+Hz31b1mtvYr7Jnu1ydi4nl6f3ttHaW+PTT9H5nBezE3Yvv2A6nTA5ndJFRbse6zNhOfWRvuvY9Ik4O+HDv/hL7lKiA3arFf0m8k+//LXNzi94/uLqW3ksvkk8SoP75LNP7fXr13Rdx9nZ2RjmqSp934/h4hA+HmYTv22JkK8K9YKT95QfmiuZ/4//vZ2JELqWbrXGfGD64UeYd6CJ2G/pux2aEuqFxW6Ha2bMnjxn1Sdmp+dUoUN3kdV6wd12x3/421/Y06tzvvfxh9+pY/WnxKMyuC9evbTFYsF8Psc5R1VVYy1s8FaDoY0G57LxXV19t6/WLngu/eyd+2C7uf32v/tveeorJqq4ZKxe3pLwcHUF0hP7HX3q6VJE8KzbnucffZ8qTNHVkp3sEDNCcExmM3oCu67l9evX/PKXv7Tz83NevDjqVj4ag/un3/7Gbm5uAJhMJogIKSUgh4PBecQ7gvO44Kl8+M55sS+DleTHO3H9mvX1NU+aGhFPg7G8nbPpOs4mNbhI1I5IDrM1CV2f+N73foDHcTI9ZbtbYygJI4QKcTVd7Kl8zXa7Zbvd8qtf/cp+8pOfPJhj+s/BozC4f/iHX1jsEnVoEBE26w1mxunpKVWVM47BeXwV8OK4uNob2mKR0+7n599t47uqT9+7/Z/8p79lUgnOgaYe6VrONTG/fs2Ni0x+9BFt3LJTQ1vo257T03NevHhBGyONd9TTE2JKtDHRR8ViS43Q9x1VPaGqKro+8nd//5/s9PSUH/7ge9/p4/nPxYPmUt7c3tnf/u3fm6pmLxYCAxtkqKHVoaLygRCysb2Zpj8/v5TvurF9Gez1b62f33LW5ItRrz1t7CH2hJjYvfqC5SefcO5qfBQqX7PZtvzoz/8cDYGEYZKznZOmYTqbMZ1OqYPH+VzGWK/XmBmz2YwQAvP5nE8+/fxR6lY+aA/36aefEmPkyZMnqCpd348Uq9PT08IM8b83u7heZS93cvoADe/lNfFmzexshvZrwBNdgEqJqcfvIrvffsFlM+PcB+bdDhP46Kc/oW88rSogeM1X76urXAC/vr2x0HX4EPj888/x3jObzZhMJvR9z8uXL3n56tpePH9cBfMH6+FeX9/a7e0tp6enzGYzRIS2bYkxcnFxMV5tv0oq/+T0Uh6ksQHrT19xpp6qV2gVsVw7TOJIKXEiDrdYs/r17zi3QLfe8vH3v4+bNfTeY1XAfEAxkirzm8yqcSFQTyZcXFxQ1zVt26KqOOc4PT1FRPjtb3/7Te/+144Ha3Cffvrp+LhtW7zP9bRmMuHy6orLy3NJqpjAfHH3oMOb1frd7T02f2Wr333OB9UJYRuZJqFJ4KOifUTbHt1uqWNP92qOLtZYUn72139NK0LyDhOHc4FnT55L5Wu8CHe3c+tTpIs9oak5PT1lMpmMIb2qMpvNWC6XfPbZZw/62L+JB2tw8/mcyWSSQ8muG8sATVmrzO+WFkLg6uJSLs/fTep9E+vV3LbLuQ0h5ncFp+/rwt6sWN7OISXoE2KC6xW6joDhHaTYMfEVPim//sUvwQmTJ5dEBJzHmydYKf6bgRoO8OJwVV6xDKTugaUz/B5mxuvXr7+mo/DtwIM0uNv5wob6WjSlSzHXj4InNDWr7YbLizO5ujz/g8LEk9NLmZ59t8PLgXHSba6tXa+YnZ+waTeYRUg9llo09XinQMR5T3LgvOezV19w9eIZeMflyRN5Wp/Js5NzuTw9k+3N0iSC12xsTy8upXKezSZnhEMIY/IkpcRms8F7z3K5/IaPyNeLB2lwm81mDF82mw0DGdl7z2Qyoeu6b3oTvzGoZf6nT4m+2/H8px+xcB1dBZ12JO1Be0x72m4H3tF2HckJUlV89JMfI5O3uZLBeU6fX0pVVcwuz+Xm5rVtNyvubm4xM6bTKZvNZswSr9dr6roea6GPBQ/S4IYWmqG4HWNku92yXC5ZrVaoKr/+p9/a7z55HOuH+aucyLhe3lrqe9r21tY316QUqX72A9xHF3ze3rF1kU1sUVEWq1um0yavwyYnrHtlevWUq5/sOwOGtW+83ZmSD2VVeJrr9Zrb19djYmq9XueyQ9+PvNSUElVVcXNz8yh+B3igBqeaGf5N07DZbOi6buzAbtuWxWLBYrFgt9vx209+Z9e3D/sHv3yeaWlPz66kDoFKlbTZUAdH3K148bMfY9Oal/MbnPNIUqa+Yn23ILaGmuOz62v+i//yv0QuPpL1Jh+vYe0bribSXJ4IZCmIX/ziF7bdbplNTsY+whgjXdex2+2IMY7RBvCovNyDrMMd/qDnqrRty2qzpa5rTk5OCD6MC/jtdstms+GXv/ylTadTvve99zMg5qu1XZ6efKfWb4v52s4vT+TV3Z15p0y9sL29JS020DgW/ZYnZ2f8+M9/xv/x+TXrmyUnJtQRLsMpKxPWyWH1lBc/+wtsc2sye5tX+rvPv7C237HuWwgeYs4Oq8k9yYmBGD70FPbd7i1dloeMB2lwKaXxB3327Bnz+Zy+7zEzqqqirmu6rmO73WYhnapCBBaLBf/xP/5HG1p1Pv7443sn1nfN2ADOi+epKo/1PfX0mXzxt/+LnYQKL8akmbFaLDk9f8rHP/wpd7/8DTEmRI1tG5GTc/7Tr3/F07/8GfKv/s29/f/1p78zVOi6HfPFLYnSX4cgWEmQZA6niDCZZIoXMHo9gL7vv9Zj8k3iQRrcEFKq5h97KHSrKk1d5+bRvme72TCbTWi3u/HqG5xHY+LVFy/59//+39t0OqWeVMwmJ8xmE548efZ7jW65uDEzQcQ4O/9mmzFv5wuLpphGnEB3/Zl98X/8IxfTCWI53BOrQIXpz37O3csln//2Ey5mp1zHBZ9/8Qp7+oS/+m/+a+z1S3t1e8OrzQY3mbDcrMdoovKegJKi0pcQ0cFIqavreuw1HNTLBsrd0eC+4xjCFzMjxshsNsvd22YE78c0ddM0rNfLsaPbew9qo4echUDft2x3S27jNWaJ//n//z/ZbDbh/OScalJxOp3hvDA5qHV900Z2iKH0MV9c28X5M3n1y78156ATQ8yIKvh6wi6B9C3XTc3y9IyVONzHH/LXf/4zTn/wffyzC373+WeoOc6nZ9ikYRM7mrMJaKKKhhOPOiGZ0hWjmk6nTCaTsf45rKeBMbH1mELKb82J8cfC7Xxhf/d3f8d0OmU2mzEtYYxzjhACz57dN4b5za2t1+uxXjSsMVQ1p8hRRMDjQIwUla5v6duePvVMmwkhOKqqoao8IdS5J6yuEe85e/LtaL5cLq5tFjzX//RPOI2oNxQj9sp6t6WaTkh9z4lWXJ2d0nUds/MzFMGaCfPYoQonbspu1xEduFlDpz2180yTICnSqtGJkCqH9xVeHLUPY8TRti273W5UPdusl5ycnPDznz+OEVcPzsMNC3TgQAJBAYe8o+P58smVLOZ3Np1O6ft+5Ft672kmFR4DsayqKobDIXIGCCKMNb2hoLtbb+g1IZZQ4O//t//VTIQgHudlL6FH3ra6ru91kh82wgKcXe1D2NXy1g5FiFx5OHjzwbMP95gifUK1o13esG5bYteCJdQcSaGuKk5OAskbzcmU4AM3fc/p83Outy11qFktbplcXpL6RDJjMp1itafTRFPXSNSs/YJn0niq2qM+YALPzi/lZn5rse9JFnEi+CBIl1ksQ/LqseDBGZwTw4lRBYemHsiLdLPEe4S0OL+8kOVibj4IVe3vrTFMcxjgXMB7yVdtLzgXcA6en7+/E3y7nJumSLKIRiOlvoiwRizlJM62y+wLdK/0NRoQxi/+9n81k7wW+vR3v8IN4kSmOAQphZ1RKUwHg4OQlBPJ+67TQJci52enaBQSjuBr6DrEelwfEevpPFxeXpD6HdOmgui4vHzGRpWNKeqNnkjXZnqWqBBw+LrCBwfeIQ56MUyV169fmqWIl8xlTZqwFEHTeFE5lgW+47gv7DPoPt63tuVibmfnh+uu/Hi5mNsQ/qgqJH1LKMhKqtvsy6Og6dm7KWC6LTqVariTJ2LbG6N4pDe9VdQElOdJoIZoGj3Y4OXGdet4D06VJkEbW2RaMZWG4AJRBE9N30Zq52iXaza3L4n9mrW2RE2cTs/x7oTTixfsknHy9AmnTy/ZqNFrpJpOc3e8WjZ8EWZXl7K+m5uagUYorTsCILkTAdH9/3k8nm3AgzO4NwV/Dm+HMDOWi/m9EO1Q2m4g17rAPc8z3A+31Twbj4hwcvHVOJZuuveKcX1tMRXj1jR+vqqCRVzcIcT8nMXiCRNYHmefkt7bnsFT5oeBJBXbPuK2gjnBWU8VptSzkENf52lmDXWaYa1x3nuuP3+N3V7z6e1v+GT7H/jhz3/Oz54/BZdD3amvcL5BY+LJk7c9/OH2yHvSBH+oHudDwYMzuMvLS/l3/+7f2VcxOnhbwHXA217N3vn+08v7J9z6bm/E+aRL99aA4+cbgNL2CSMhKsD+fWICJmhO14CAkyqzgwGxItAawn3h2bKdgsfEYVZx4gRLLamPOOdopmcwO+FktoO+pdctu2WiXS7xfeJZaEhRWYvjX//1f8bP/5v/Cvf8I1lcf2EuOC7Pn8mru1tzCPPFnTktnQJAtGH9WCKL9/xOb168HgsenMEBbxnY4eNBo8R0X5Ad8C7Z8VSMYDhzRISzL5Fc+KpebrucG5I9WfYCQ3InPxIzknl8fcbkImdWt8u5vRmmbheZZrU/aQ/3PYd2k4vnEj//lcV2i4TArr9jd3ND227pd0tsvWDStZB63K7HJ+Gzl68RV/PX//bfIM8/ktXLz2253TC7uOBuM7e2U55eXsl8cWeHBmbsj6EApveN6dDA7nnmR4IHaXBv4k015DfxPs/2vr/9MfC+9d2bWMyXdjdf28XlibzrPdPfU/PblebT9etb4nZFPanZdR3LzYYkSu0S3XqN0w4XI3G7xdrELik/+Vc/hydP6F5/YR1wfnnJyemVzNd35r3n1e2NDXzKxfzOZPSuoAgOGEjN8P5o4jHhQZKX4W3jeJexfNlstvF5kfFmwJ/yNFkt83pwubix1fzaFnfXdn55JheXb1PK5us7m6/f3ak+X9/ZfLU0gMnJpXSf/Npuf/sJstoSdh1ut2ViyqkTfN9TuX1WlLpm3vfI00su/vqvoO/o1JAQUDVeXb+0lBJPzy5kGGQCOdNrh2G7OTD3lhc7erhHjvet7b4JnJ7l9eDAVFktb225mNtQRzwMZS9P3t+l/ubfdq9f09/eMvvgOZVEOu2JZdRWt93SeEdUpesjZ82MNuyonj2DZ0/oW8U1FTEpCXj+9IXM13f2en5rV2f7Bt7F/O6N0PY+3jSsx2Rkh3iQBjcY0GEh+PC5w6EcbyZEvmx99nVDFc4vLqWd31lz+X4Dm9/c2mXJFi5ubu28PNb53KTb8Y//w/+TD06mbG5ecTF5hpAw7TGDygmqCVWDULOMylzh3/7rf416Tx8EM3K3vMub8C5jPy/bt7i7Hdd0wzr58PgOBIFD+t2bA08eMh7snr7rCvqdC2Es/zyHxraer+5t/N3twqQUlQHOD9L07vJSuH7NqUZq7UF3tLsVah0uyMju1wShakgI26SEszNOn3+Inz0RdQ5zwu8pOd7f7C85vu86/o/J4B60h4OvnnZ+X9r/m8T5G17tdrmwDrhZLMcd6lV5dnkht8vFvRBvwPrT3zGzhFfFkVhvFtjJFOossGQ4ogi1BHw1YblY8uEPf0Z99cG98Pb3YfFGTfPNY/6utdyAx2RwD3ZP9SCcObyNPMOC74K3u10u7Ha5MCUnNkqJDpM8pAPKoMa7eUm6lNLH9Svb3t4yFcO0R8Rotxv63RZRw7lAMs1CQSoQKlqDH/75nwFve9OvgjeP7Zu3d71mkDB8DHiQBjesy961UP+uhJVDvRBAVBAVvDq8OsTyzavjyempzOcL88m4uriU5fX1SFlbff450veIGaoR74rm5HpH2rYEheACznnwgU3bcfr0islPfibxemHBvtwQBsM+ZOwYoGaYyltZynf9HkPv4mPBgza4Ae9LSx/+/duG8/NLGYzu8uJM9rcTcaY8OT+Ry4tcLnDa403pFnNzAz3s9ta++OQTvIAUupeoUavgtgldtti2pyE3n7rgWXY7PvppFglSgeZq+qXhpLO90cH7w/d3lQEOn3tMBvdg13DvC1+G+2/jmu1NnJ9fyuLu2u5X/wSTjvw8eAQnoG0P4ggSSTefWbe9o98t8MEROwUvpD4yUUE6pY1bUgTvKkyNXnv8tOH5j79P399aVX25ZudmPh+P7FcpBbzvgnfMUj4QvJcjmZ/4+jfonwsn4By4A6qaZe8ipiSU6fkTSURUI67K7TEicHZ2Ru+hQ6nME1qjUk9lAm2LrbfE1QpJkU274/TyisnFj6T7A7pl3N7q8rbxbv7kl4Xxj8ngHryH61NuJs2tLeVklYAoQE55C57zNwZ6rFdzcwfnx1elYv2xoVqhAg7l8vxCFne35ixzmBNCXxUjDA0Rw7xDNXE3v+N7P/oxu0aY//KWk11g6k5Y7hZ0ojg12HX0LtLHKTfbDf/mP/+3AJxMcmlhe72z6dPJO/d7KBNkonWpuw3hq2ouIpY6Z4wRcQEkC9EaDsSjZmP/4mPBg7203FsvCIgd7OrYd5bA3D2+H8BylSlWVsjE+jWb2vpubuu7/fpNxGdjm98Z5samA5GcmLhe3po6jxqYk3F4SZc6mDXUZ2e0bQfbNuuYtDuctZBa0mbNbrVkdnLC5fMPgL3Aq/fv33GV+8dFRHDvsJu31su4bHAHfzt6uO84BhaJquKo3s6MGbjhOTMQYT5f2OXBrAGzzHsX+Xrmwq3WQ6bP0GLoy9WtqQkXZzNZ381NKLHe2OljeIOkijkQL4SU2KxWBGe0saNuKiZX52w+v0a7HeITttugwaNJkUa4u1nx03/1c5jWtJtra/ss6lNfNl+638PFbIwqy20wRpXcBJveeM/h/aDq9VjwIC8tQ/Po3tAcasawu4MxpkMmOzC/ywXls9OBrfGHMSz+JTg9uZSz0ysxhP0NLs4uZbO4tYDhxguFkQQQwVuWW1BL+CDobod1O5BIrz07Vdxsip1OWWtP3+0ISYm7HanroTe0hw//7C+h7cCMOngGD/svhariLHd3v4/I/JgM7kHuqfeeruvyD6qDSM2+QLy/L2u94i3e16C6XM9N+JKxT39EnJ9evfUduT8u31vpc0tZRIygiohjqz3OeeJmi08JBJIpbZcQV1FdXXH3+oZq3VKhpL5Hqoab6yUvfvpncHpFu9vhvMck9/UtFnN737jlNzmnd/N3y8W/tw56cK1/TAb3ID3cUNdRcmE1WdYhGTCSZg0QfSvMgczcHxMD6FtrkT8Em0HDBFhv/nDPMT2/Kg3gOQeYvZ9DJbfAeASnYH0ktTsqAU09ZomoubdNZg1MPF2/I/VZJCmZcL1c8mf/2b+CNiHVjBShL2nKf4l3z8dXwRTB8FjWMLF9gKnkC+L5xfuFmB4aHuSlZegGeJtxMljQ0GVdjEw0yxzI/vqTi842esCzs3/+STE70DA5mb3bYwxDHt80+uFxU07KzXxhJ5fncrvMc96SM1xKTF2gXy7RmHCSmSVYQgz62NITaS5mtNeezXqDes86tkxfPMX/6Ef0UWkuL2V9e2dSkhjmvvouD+u3Pd7dKZCPZ7q3hntMeJAe7jDrZSYj02KsC1CuwEJZX+w1Qg4x0JMwd49q9SfZZqPk7wxHwpHGhtR72zRk+CSP+z29vBBVZYLHVpuxy1osIZpwRFJsadOO2fkEdxLY2Q71xrLd8sFf/AROG6RpePn6lZ1cXcjp5YXc3t0Y/2xVrWJsKK7c8opUuVcBMPe1rZG/LXiQHs57TyoNllKu0llWWxHnUO1JfZ8HS0g5ESQgpsznc7ssNbn3rV8Ocbe8sVxoNpwDcSXTKODUM2u+Gtt+enYp2+XcMLDB+xafMV/uLIrljJ8qi/nSouYM4ev5rQVVWGzpb5eEsxoJDolAUqzdMa2g22xYtWsmT06Y3wm99Wwk8YN//Vf0cUtf10jIokCX5xfivaDpyw1ueO18PjeR7Mn6bpf7DNVIqvR9nwd5zGpgr0FplmXoH1NJAB6owR3+iIfdAaK5QJsFXhWxxCC6k6/K9zl9X5Y0GHBxlg1qvbs1FcWZkcyyLPof6CCG4vr67mZI45DvI4ZhkrdZUMQchiIkgg+0L18y7RUT6CzL6llMVOLQNmEmqPN0AuH8nFcv7zh9/hGTjz+kn0yxBEECimO+Wtv5V5gUNOiZWDmORiphY5aHF1M09vgqjEaIKKb7MP8xJUzggRrc8CPmbNhBS07ug84y5ClToXAhZ//McJl+kjNuB2uPQ4Lu+zrCB3bGHwMnF3uvuLp7bck2eCnGZfnU1nLvQx5sP//iM86aBnETesvyq2IOT8Wu34E2JBSVKeHsGevPN/zlT/8S6hPWnSJJubp6KvPlzt5U2joMbU/fWMve3d1aSkUn07JROYoyYPFi3ntIeq/LG/Lfq6r+Yx227wQepMFVVSl2q2EuYZan4uDzojVpQjWisSfUDtOE84PUwv5zzIz57StTjTgX/tnyC9vd3KaTd7+3393aoKg8MPrzhUIRU7zbYC4RneINKlNQwwdQcXgTWK+IyxtSPGHKOQEjksNo14MkD1LTawQ3IeKoL57wo7/8KzApnjAbweXZnsrVztcW2ZUkkxuTKIPnn8/nZgesHSsdezJoU2ok9i2TST16v+G4DvdHD/cAMAz9U1W8y/w9tYjgcCVQM03E1OOTRzygHudtTJ4IHhUQ78jL/8Td4tqkJFHELPellWxoGc9BdTl7y7Aap2h3Y1nue0iNlySO3ub/D89bxCyhmkPHuOtBFESz7kjM25dCLgt4rdh+MafWLf2mZ9o/pwpG1IhJ1tUU3xBROvMEmbCNHc9/8GP4+GMw48SD10Q/v7FNUkQ8EwmodngHqZDfVGFxd22mkdvba1ONe95q4VSa5qgCVWLX59FfzmFJ72mcDBnjo8E9ADx58kT+5m/+Zt8UmRSLCfNDelqJscN3QlPVZbCE5oykvJGqFtk3F1iuJXmneHF4iViK+blh6ufiM8vxFMVQOmivMcvDPNA+j8FKEbOIpbYQqxW1Loe7mrLxaSqKxkrE8Gb4zjAnbCsP4iFmg5tWhqUe1mvqsymtE3pNIJHkIhEl+pwxjCS+/6OPICTo1tSugmCgLY0XwOHxOITkBBnGBpuShmNRdFSyJHsavZpZyt/R93TdDiuJkaRplGVXVRB/9HAPDWbZm6lqOdmBosFoKdH3faYeiRtfOyBn3RznZ1dyu1yYQ3OWk4RIHkdl1uNCj9MOYktM22xAo0EZvdroAUFxItlAzUAiXhiTH5hhmj2cmIEkkuUCdiDhVPMPZo5GKkwCtky47Zo6nGAusL19TV09hUlNShFcy06W9F6h7kmbDc3MOH9ew+I3kBJUDYQKvKfyFYpAEpAaTTMS2auXq0g+thrHY5w5qwfjskzH0V+DMlpWBjtQUSv0uyEaeSx4sAa3LyDnNc8wVSaz+hKmbb76co7ZPixyYmNPl7ee9d2n5ljjxfJQDW1J/YbUb9C0I7YrsBbRNhuNRDyGdyDiqS2QVbWKFywGB1rCyH17C5Zna2fjAywRJIH0OIv5PQKIo0JwGujbjlCGRzrnubt+xdWspq7P2PYtOCXGFguC85HO1kxPKqpJz2b+aa5DitBppHfg6wnJjCA1k+aMUJ3jXYP3FeYDZhV9ErpkxMLuNgQb5RQEU0XTjpS2Y3QwGJoOa9VS/D56uAeCoRYXYyR4yTUpFTYkQuNxtqHf9uw259TTKeLrksnMXsebJ/gIdgP9grTb0u7WxLgB3SH0OCKNKz1dNujp55DMJckZUtlTmQYp8PyfAybGGLKm0sip4/OWEuIMI+KC0ltErCYkgR1sbiMTNwPztH1H8Mby9efUac3F2QWvVjtm0wuWqzmp74gucvXiCtgRbEXtBLNEJYqakdrl3hOtoYsJJ4FYTXDVKb56QlNfUNUnWNWwij3JhF4nmAYkCaKKpjXd9oa6vtz3xInL8wvGmQpKVU++ztPiG8eDNbhhbLDL86bQVLoEJM9Z87RYapHYQuwx7wkCjVc8PXG3pVstSdtP8LZFUw4dG3rERYJTRHIImBMnHtSBVORcqEMk4Yx7g3PMSnFcXTa6oQ3IKKn1gSQlOdQcpcNLjSs4RD1ER9wqlZugbYCUtfwrjfTbDtwKNU+tjl2M+GSkvuNs1nB21kDcEqyFWOqRlmX0vBoq+5rZJEhJgBixrem4hXBOqk9I1ZTm7AyhwVxA1ZMsr5mJLWgpCUiZ5KOpBKSa9x8elZ4JPGCDCyHQtj2qAQOig1h4k14gWY9LO9JmyayZMRFQ3ZGWd3RpSWyXkHbQrjCLObvpEk40D0Qk5WY5y0MHh44DcIOOQ/4n48MRpoXYSwkl2f994HZi7q2BhWZZZQuycXfbhKeBBETDm8fH3CnRRSNFz+Tyir7bUsVEt+s4++CUJgRss8asI5GTHDYYtJUCO3nfdOQ/5m02VVLakrqa6AOxn0E4I1QfULkLKquJSWmjQzuoLpt3znkYQv6nT749StdfBx6swVVVxXYXiZZ5iqnMO5Tg8C7XrzT2aHuD6wL4JX03p4+vIK3AOrxFpr4aeYBI7u0yi+QcR1EuLtRdKYX1bHuF7TKyWO4Tk/Ve98GBYQ2nn2gx1JJyF0/SiBOHJoezgLOE4PN3qyEpEqIQekhdT9stOJnMqFPE+o5aIxdNBf2O1K9BEuISemBwA6wML0kp5W7ucqEKLqIkEh29wW41x/wpVkUsbPCcUZlRaUsg5SwwQ+Y3e3UnmfH52Ghd8IANrq5rYI1Jjwo500aN4HGWi8gp7WhsgW53bLueFBc4t6UJOakhSQnDIZK9UYgU2YbhxDkg4Y6MFtFCFtsXeQdGC+yZGMNnv5M1L6CpcD2DYElQgb5PBHEE7zHL01qFHu0TLiYaFfoY6XY9u+olMg24fsN57Zk6I20WuNCTrM8lCOyesUkphagI3jcl4dEhGhEizkFwjto5TurATpekTtF+iXKKr2pqWVK7LXV4uytgePzYwkl46AYnCiHm5k3JnsgnjxOHVwjWUknKB0FbAhuc9YRkOHWI81jq9h9a2n4yccoPVpM9kMtrQ9wYJeJEGJgYA0bDg2LEhd//Rl/AqFui5XNLOsaS0e06pnUOm6MoXjSL86SOFLfZ86UO6basbzY051O8T5yeNtCt2PV3+AkkepwrbTiliO8sS34p2eDGCwQOESMMjTjaZ2fsEhOLqEiRVYiIBaLeEUxzHdAoPYklQ+kCJpkD+tjwYPc4p5sVcbkwa5qv0qoVAgRv+EqZhC21RMR1GLtcgG4No0ZqX7qshzG/Mi7Thgc2LtJK8kPT3lqQQuzNGA1v/HMx2DdUq/aZy6w5GWPCmyssFSO2Hcl1uTjtlSSRZGsia9AtMbb0fYdZR9x1dEmYnVZIPGc53xJdhySwYFRNjagDHxDCfnqsCU6ETnucy97UO5+9tybQPG+ctkec4ase7zYgfV7S+jV15TBLxPSGzHw5DkcP94Aw/Jg5EViY7JpI0iOSqCqDGuoqoboitassBR4avHcQDfqIuZSzduIxMRyS++QKA8NKW8rguXLYmIqiFrn+J/dtKttSzqTIQSh6qPGYM5cG4nIWUWXohCPFlhR32RsFSOyILEmsSbImpi3ROhQl9VuW6yUhXLLbtmzTmul5g0aomwazHvMBh8vZUwEvnmExOWkGPRglaVY582LgSmuT1Fk7Uxxoj2mLVI6mUs5PZuxUDwri+bB6e3xjqgY8WIP78MMP5f/z//0b01gTvCCppbMdTRXY9StOXMT7RNIdkPbCQ33p/HbsaV7iS2toSTs6h8O/EXIxeisZY0bw4jLB5S0v9sYD29PBxtVOpsngnWB9l7M/Zjy5PCXutuxih/ZKn7Zs2te061va7YLY7dDC8uisp542IA4Drp5O8ROPVR7zDl8FTHzW6nRDkb7sPCl7fZepWLi8vxGHE5dnEuBzVJA02533uaafHFgukfQx0cdE13Xg8oWw73suLy//yL/6tx8P1uAAvK8Qq3AqmbdYMnGpkKmw3F0tkovLWajH5VS/5iJ4FAFy8yqU5IgZXaE25dDVsq3JYD/7BIsmLSRi3p0YgbG/7TAxA8UrltKDkGt6goBXXE02BG/UtdDMpmiraKyQNHAxjRCyanNoHGFSQeOhFvAuM0cImSI20K+lGFtpcHUO8JKzliaoGqqpiL16PAEvki8oXsb3puhJKdcjM7XufmgtIo+O1gUP3OAqH7DkEAlgPWpC0oqUGqKFHDqZ4okEYu4Ot8G1We4rUEGdIQiOAC4XiYPLnkDKGi2n5weDKZlJcaPHkkO39wZGeQd9h8a4DBxGLSI8md7lQvHCQXATqJKDVIGdgPV7j2mAxfx/nzCvoIKJB+cRl9P/+ybcsoYL2XiSSrkYeHBWDDDvsTOwmGt2IpbrdgqxF2JXES2geLqk9GqouJJIyuFk0zT/wl/4u4cHbXDee7o+YeYQIKqD5JhKjdkUo0HN585kybSqnMb3JZzMIWFOVjjMJSQFEJ87msX2epdDsRgY3iwYe12i9wwjHD1ieutvAmWdVIqIpPxdVjxYIUQjCfwO3A5Sh8Y8oiqbeGmXEc0hoMshpLjsWdXK3y1moxIr35MNzoemSCwbqn3OQg4XFvFUoQZVrNDouhjp+wpSQEPIbU4HHm7sFoCjwT00OA90bQnVImqQ1BOtJjFFmGEWEBwxRbxFjNx6I7mWgLhQIqXCILFiwOpRBPH5pLJhgoaVsGoofJfMIgwJlX1Xgg1xKG8YWflbXtJZNihJORlDTlzkzxXUYi5g02LWotJjbuhKoJQlUqa4haF/D8oVAzd6J0OJJeGqmAdRnzOYpW3JieFcyh4UMj0rRkiJFDu62NMmiGmC82eIy3MRhqTJkFwaOgiePXv2qFgm8MANLgSH0YFLuQk1AcTc4+UqxDWITSFt0NhCKsPmRQkulaxdoKQvSwvLfp3nSrhnkvLr8Nkwy0mbyxLcq7cNxXIY7PcgjJShJ6+8XrKmJsS8hnOW15aUJAYuP1e6FHCG17zwkpQNPpYaX0JzkXwoRWCI5rKFlQuMK8ZoLmWP7bW06fhsoC4CfTa4mCAq/TrluR2UzymlBVyNVBPeFIYbPPhjLAnAAze4ug6Y60tRepdXKH6LOY/63Ant0gmOLkduFnKbjW1IssonsTmcC7jR4KpsfL4CC0RAXA1FtSszTnJgalYKxaJleaRjJnLEweOxD67wLA0lGojk/jm1eFA6GMJJciu2HcpD5LR+wnA+T6nBHOJcCXXLxcF76AxxORzO5Y5U1qXlQhBycywpQddjaUvsWrRV6AVnM0wF8YKvPRYqkBmEKean9LgyOSeONRBx7tG15Qx40Hs9ZMGsnOhZg7IvV/Ga5GqcVuAafEmGyNCNrEDKxF6NicgOJbNPXPBIqDBX46sKcT3iK6BGpC6JjiKgUkLDnLigeBRA873jzahq0M7McuWuKBcbCadFqKd4PUqm9Z6wKnkslIkrhjgwbWzv3awf14FykCnM4WqCvh+zurHt0JTQPtfTxGJuQNIa0YrgBmaK4KQmuQaTBvNTrGoYAuih7DI8Pnq4B4jT01P0M8O0QlLAW25TETW0PiGmGQ6Ht8EzBVJZuzhmiHlUO4yhcTKC9GS2Vw77pGowJzjn8WECoYYqGzG+gslZrm0VGpPsawf5XmOZpZbGISMDO0VUEct6lEOiYpiaM2ZEvQPyOnSECeLBm+a8/pi1NHJIWNaVMUFXHqeeFCMpdqSURZbMDJIiUiFaFUUYUCeYU3LzqWLO40MFvsZcQ/RTtDlBmhkWAhLz9jonOOfo+56qelx9cAMetMH5qs4MEc2FasEhaVDFCqhrSsYxp8mteCbFIVYx9IdjPa5wAXMTadY2MevRNgudqvMgu0yRch7nAslVaFjm0M2FTDL2bs+wEIHKF8NLuFFcaBhLpZDifofswNAG9OW1Y97F3X9t6rJBDxIOsc+CSpqrkZaygQ9anaqajQwdrkLIQNCGMonVk3CI5KGQUuYciARUGsxNwE1QXx+UHIbNskc3ouoQD3qvL88v5G/+p/+fpWR458ZaWZYBACceJxW4CpE+62yoHxMS71OCH8ROKcFe/j8ks9x8KQ6RCOyAeT7hvBvJzyYHY7C0hIOST/yh8TMnRwzx+0RDvlDcr9V5uT/c8N42W76AmEpW1LLBi8YiOaGjPN9wMcGkfHeuxylDdrGE2lCSsLmjwJfm20zm9tnT+0AKVV67mdtT3QpEpHRzPD48aIODPF44RYU6/9BqOWVv4nG+QVydkyISyBKmma6UKU7F+KxojJCnvSC5IE6RTs8rMTeuybJIKyAJpz1OEmgpNZA75oaT3Jf0fU6e6FhIR3Wc/jNeJGzv+ax8f2aIDFnQA5grPW0NJnLQcV54jAfbUFr9MNtfZFx+eqSzWe5HyvolPh9HKccyezfJdC+pcb7BpEJdNTrroQMf8u8wmRxDygeJqmrYbDaF3eBLBzOY5l4vkZBT+kPrjbjsoYYUP9k4rZyBIntmhpkrrI37ArK5WqA4E1Cf11TiR4OTQg8TlNi1I8sji6iWOtUgslP4ilbqe3tNx2zAqkPzgiufPWyIQwUiXQ6VjaK5sq8HAiXR6TIR20qmEso2G+KH2mLI3y8ybguZGFckAiuUPGBEXMiZ22LAZoMMxSC2K4+y6A2PwODquma9XjNMz8n6inl4ryGo5fQ+JczLDsyPhpX7xDIDQ0v2UMyXE7skN4oNyPC4SIWrGc6yaqU3TzqgfuUchsO7hmFG3TA/zTmFVARoey0eZnSb483GFiGKRx6WcsVoSieD3AtD3fiKzLXOVLbB2MaheJQiOYZ5Xzy0L2waf3Bscp1OS9LJxKFkUoCau1f0HiAiXF5+tSEnDw0P3uD2oYuW9HUOo2JKRAFPJvCqMLbcjBCPOT+uX6S05wypgvFlw8nkijcoici8disno7mxu0A1ljVVkcvLVOrM8BgGRGY6J8FVY1ZzL4MwsPkdOmr5H/I08/bkjm24P/1VxscDx2V4Xgr7JAfGHpWiu1L2UcmGbYMxHrr14mEVyUZnvjSeZqZJnl4kbxnfY8ODN7iqqgpNK5JDsyzcHS3XjlSyIQhWXmeIKDYaXwkdC7tjv1oajGRgCA+ebrS28qoMKwToQeErKxVL4RgP1ar8ffuo0KFxqJXnLgche1oZ1kMu74eY7ddqrsiww7hu2mOfXNGD8DRrdlpZuxleUk6eDppIHF5YigeX+/OGkpXJrCW0POzyPiLjwRvc6elpUQmuSNrhqpouRcSFfMV2FWqCx4rceaEqWsqGZ670gRXNSTs4xaTUtgcZoRL1iZTeOcAN9TOF4WTP0uapeJgEZe02jDY+FPSx0oOWCcjcY5iMz1OGlJRqXJ5bMjzveXPw4TBeKrcM5c2y8tFKvujkHEnxYm6fNNlTXWAI0SnML1UlWk72JBNSSvR97jXsuo4QAiklTk5O/jg/7ncQD77l9snVmXjvySOFy0lcEgPJKP1ghWRLKmHUGyjrOB2SBiWsGksHkvvIRIa0/94bOHy55VEieZ03aPeXyafmSlOAFIqZ298017nkINPoDvzKPrS7P1hk2G41X3r/wri92asPSZTcWJrra5aNy5PD7yFf8gayEY7/GY0SJ+W45EztEE7GeF8W/bEmTOAReDjIiROzDc75PN2lXME1WfFYZV3DXpg1iC/0q7xeG9ZkAoW3WMI/tz8rc41qX2PLLTKlpifFazLoUebOchmyi6Wj/LB9ZfByyTI7JmPwgkOBu3wXgNr4eXlTPGrCkHjMmcf96jNHzFK82qHnzheWnAhyJQ9ZSiQyUNUEcX40tkxAK+EpIStUJCWmPCMuhEAeaJKYzWb//B/zO44H7+EgG1wevlhqSZZPjD4pSgCp9jJ3JRM4iukMBuSG1L6Ay6WCnL0s8ZRIDv/cwfucJ0kO5/JnFKkGyaGiIvk5AS1pfC1liOF+8FjDZzo8TmRsv3EHBuQOkhH3f9hxFTYmXYb9zcY4GK+UmoYU1s3QeOpHaYThGIyGJjDMZsiXE5/DSYWYbOyFO1zLvXjx4aPNmjwKDzednrDdvgLAuYCmclJEwapchxNcST7k7oBBUmHwZEPefzi983JqyOqFMes3lhOglBGsJFoGubxMqXKFz6hWXuFz8sTK2kwpkg+jd9uzTPb+tLBRZO8ND8nQOSfyPsZM2Te1PI5KSmlhyK4Wubxc+B8yl0M9z40XGzNXwshQ3uNBAmqOlKxc3MrxUH2UsgqHeBQGdzI74+YmS9o5F0iaa1f59PUIFeIDaNYuyeblGHiWY93L5fXf8AzjVLg0GpqxT3Ts+Y1axAt0lD7f/2VfArCSgsnlidwPZ1mBqNTthgxmoWCNWcaBDWIlqTEYWMmMcpgpfJORsl//jfdWtE0MIOQ1qrnsjZ3lOQGFezrU8bI3DGWt60jmy3HO8wUGD/eYw0l4JCHlRx8/F+dclu0eaFuWDQ0LYGGfPBCXjc/50S6GZEt+XNY3LmcvAUqOk3zC7nmYw2M3JD9KIiQzUNy95Igr954KsVASLRUmVTnB8wmvzpNESOJQtz/xs0ZJQMtr1JUaojMyBc0Y6GO5747StjMkWnJx3lsoaR4/rj+HYybDxYlq3DfvK5yETI9zAXEVztXkEkz2vIPBqeqjzlDCI/FwAOJqTLP38iM1xBPF6H2N6gynE5xrkdiNoaOK2xd6h27l4TNhH+YZ2D0pvJz4cGM4lvVI3LiWMUxzet6VOpg78Jsjw6pwGhmoZFJIxCVxEpwrJYvsAYci+Uj1Qou6npDe8HJawkNzZR06pP7Zf8xhiJqVlYFyARB8lpiQkHmTrsaoUSpSqcNhQlUudmb2aEnLAx6NwQXf0McWSR2eQOUb+r5nmRQmp1TpCuKaKXeYtZBinu8tQhKPI8//dmUSqCttPs4PQUIeioEZnj1DY0wqqIzG6UrWUnyRODAZ1Z0hZzNdrnYjRT1sr3/i9zW6kSI2vPHAPIS8PlM/NoY7rIR8ZS0qZC84vkfHpI+ieBNC0R9JPqLe0FCh3uUuCwkkFaq6AalIVqN+RrRAFx3qKipRUt+DCk01YTad/kl+3+8KHkVICXB5+YS2bbMorPaA4kPA6obeVfTulOjPSDQgAYLP2o2y12vcZyAtM0eclYKxjjWrQQ39fli5r3lJyWbCfil3WLd7q9+tPPeuYRjj5w5JjsP/c1BDc4J6KZnR/XZqEYfl4LnDLRhVpctn4IaSx74GKSLsYk9nHmlOoDohUo2zwA9DyqZpuHqyFw5ar+aPjoLyaDzcj370I/l//b//H5ZSj/O+zBwoXih5TGYEf4HKGnMtVAZxRSCU7GFmnuBT5jy6kupm6Ecbvil7IbHhalZqdmNmvqz7pAzwOEywjAkLHXMirni1t1A8rQxJE3n73FUBq7TwYIaQ83A7GetvbiAiD4wasvH2XssFwmdJBauyZ/aKOUUCxKTgKqrpKfgJphXBCaAEPN7n+tvF5RkAy8XcGLO2jwuPxsMBTKdTVqslvhJEMvNEE/RJgBkuXIA7x/w5hAkmuc3ESY0Tn088B+YT5g3zRnJKynKXJMdogIfG9uZtgMj+xB6fL9nC/NCPXvEPud37fCclNPa5d89JaUXypa4nB8wVwZnPyRvLSZHeQZTMTnHmCAjehGpYc4on+QZrTrFwSq8ViqfygcY7mkmVJ9F6ODs74+b6lZ2dX8qgU/nYvNyj8XAAHzx/wT/84u84PT9BJOFDkzNxVoNU4Ax1a3r6HHa6PJQRqhyOWV7HGVnyO6c+3BiPZSnyrLCV30ehge0Na0xKFM82rLpGlooVBa2hLCCCmT9oz+He+5G3CcLj/81wVjKosi9pjD1vBVK6xsU8vnQF+FLwjy73BVYKToyq1BsHaYWd1tBc4mbP6f0ZURucr6jE0acsYLRZLXj27BkXl0/k9eef2GZxa33RcHHOsVremohwcnr54Avij8rD/fgnPxPvK1arFdt2h5YZ1C7UKA3JpiR/Rs8prZsRq1Ni0T0xqxAalBrTGqwqt4OObRkITgcJEH37HHqXR9rfO8Zhj7ZfJ91//dvtMW+t63BZQsICwTzBPJVkzxXEMST5M4ulvMfltSneUG/lYjF0wOcQMXddaGHnzOg4xU8/wE2eof4UwgRfNTgHYkq325BS4tmzZ9xdv7aqqkaq1yEDxcwehbd7VB4O4MMPP+Lly5dUqTR2Tj11aMqcgZo6nBN1Rx93NK5H3BrN1WdMAz6dgvaodWBdOQETe60RV1pUcpimAxV/jBiH2l0pO9jAs7TxPq8M3Z5xZfvug/IhDMrMh//nwOhy8RzcQRozh5jlBHeF0SLsScciqPixOpCrlmVqrOTpPQnBXE3PhJYzZPIxNB/Syhk7regNkkZi16OxY7vdcnV1wWRSjx0DZobFlI1dh227v/kPFY/O4D7+6Id8/tlr+i6S4gozj0wd3uciLnJCTGe03RrY4Wnw0iEYTgOiEUwIWpIQQi4dHPARtXQQAFl1axwpPLBAhq66fSf20JyZHxclL3MlGXJgsV8Kd+Dpcuin5NJCzjaWJJEbOvryJ6vkNZ6WfZEiFiTi8Vq6vJ3LlxUJJNfQuVO2XBCaF/T+gm0fWLVdbs3pI323wxnE2PHBBz8dw8e+76nrbHxN04yy548lgfLoDO7p06fyD3//S/u7//Qf+fDDF6xWC5Z3c55cPufJ+RnbKDh/QjV7wWbRUtkKpy1ee5y1OHrEdmBtnnbqBkGd3DVuuJHU64zMWrFUmj2ztxGXewNyqLqXPR9oW9nzWJmmWuhe7E/Kfef2oU4IpQ4n4/pNSZgXRA6k+chDEYG83QNLRfJaTUealsdLyOKxahgeDTXRT+nkilg9p/fPie6UxV1LLxBN6bqOKuTv266WvPjgOdNJgxNou5bpdFq21dCUDrK2eV82i7yeA5iePbw13YM3uLa9tqZ5eu+He/HiA+aLW15ff87p6Yy26/jii89YLeZcXpxwPpvkE6F+xs18zqk/o3EdtThi7HAx4SziLKF9ySY6n09s73L7T1bzp5LCBCmaKsNayMlePu6eDN7hTVxRxtpv/pueQFXHGtxgrGMdDof43AkRzQ70RQaVLU9oJsReiaqI8zRVg4knJaPve6a+Au9IrqGlYRtP2PkLcJcwecJy3RPFEelAfJ47HjuwxHQ65eLsHFBi1LE9B7KBee9Jb3WkZzxEY4NHYHDgiO3cQrP/AZ88PZNPP/3cbm9eE/s8w3rbb1hvlmw2M+bThsmk4Xx6gTv9c+a7l8TNF9TaclrVnDQzGhykFq9k5S6TPNYq5nFO3mvuQDCXPcngkTLdPp/8DON9x/JzMczisWQQaj2kW72RYRy6sQ+wf5kQU+aFeoHghzLA/iTvVi2uCkx97nxPuw4la/9PJjOSKbtUsYkn9P4JXfWErrokyjmxd+xSj1SBFCPOBXyo6fqI83BxccF02pD6lpSyJCBpr9kyhL0P0rLeg8e0r2/h7//+7+0ff/WLPLjRjLZt6bZdXmdMZlycnXB+XhNkS0hLbPcS235Oo3dc1InTiRGsxxcl40x81IGpWEK1MBaei8QOFHUu1TjW6/ZirPe30VIeU/U+OBfuNa3CPjGTzCGuRgjjyGM5mLkNEEI1hsIyEKAL77JXR2cVW6b07jmpeYFNPiKGU3Z4OktlTrgRu6yBIpqlFabThhfPnzFtKlLqcS5Q13UhROeyRhYWGpS/3s7aPkQv9+B26Kvg7u7WLi6u5O52bv/b3/4H7m6uR2Ltpu3YbTuiGrtuS5Sei/Mpzy7OOK8doV9h2xt8f0PQW06alsa3THyickowj9c8LipHg4k0JkcUR55Bh+bMppiOXi2rIOdX+tKzE/WQY3m/a6E8eqsOd69EoPdZKiJZ399JAO8RF+gN1DwRSOZzt7YqnTZsucDNPsBPP8KqZ/T+nN5VECrEK06UdrfBYk6WxDZSVRUXl2dMJjXD/LmmmVJVFUo2NO+qvM1FzcuNqtT3yQKTk4dldI8gpLyP7W5uMUZWy1vT3vjxD37MP/aR+fyGvu8xSzTTmqCONkVmzZTVasvi5iWNCBezmovZFSf1CSbPmOstlW5odENFJKhSA7VUBA9iG0x0HMKxD6JcTqCIgGiejHOop2Lk6cIje//L5eVGDRW351SOxW+Xm2pzp7ojmSOaJ5mnjwGlYpuEVatsOgXXMDk5pTp9ytnVn0FziVTnqD/BWU1NyAkZEu1uBakMDjEl0dFUgcop3XZNVTdj02lKCeXAuN4o1ucN/yP90N9SPCqD22xem1lLCEK7jvSdcjI95cWzj0CNxeqOGCMuGKGacHF6Rup6TiZnMAGsR1Nivk4sfSD4c86mz/HW4myHpBWVrgi2pQ4dE+0587kqJ5bT8w4QZ7n1DkF1UO0CSsA1yCDIgS7kcCa+mTQ5tMF8Eu/Xe2bga4eZo8eRktAmYdcbrTo6rVm3QssEDaeE06ecffAxZ08+5OT0CX5ySs8sd8dbrjdWTtAE3TbR7Xq0z7PPrQyznFSOyivabQvTM7fjDGGscwF0PwXVDrZ94Jy+WdR/SHhUBuecI2ksV1ijCo7l7TWzJnB1cUZMLU6UZErXbREctQ9lMmg+WWRsx8lMkHmbqMKMxhtV0xLjCkt3WL8k6Jrb2BMsa/lDzMNUXSjGp0zCNI/yFSttrIe8SsMR0SIOm+cJvGF05soa0Y1ecdAYiQq7daJTo++UTh0qHl+f4ScXSHPG0x/9mOb0GbPzD3DTK0wmbJOwS44YK9QcEgTxEawlxR3ad/Rdh/Ud3sA7z1Y70EgzrfCWSL1SNdOD+d451PQ+F9ZTSqM47N7Y5N4F5KGFk/DIDG4yeSLt5ta0z3O81SmTmWe1WoF1XJzUNF5Zb1Y46RFXkXwost2Ac3lgRQnbTKCeNHm4oINEjVQniHuGI6Gpo9cdbdphqcdSBO1JcUfcbYj9ltqDtw7TPPgDyx7PCwgxD4pkGEhPuR+aSR3iJcsZRMMs64/kzuuA+QYLF7jJOSdPzrg4v2R2esn09IJmdoarZqirUAsk82w1j6tLOkj9JYITVBKWOmK/I7U7LLZ4TTSVYRpZ3d3R71Y0lc9rRpeJBA5GMrb4Up/ExhnqiUQYGDmWRx5jDzNZMuBRGRxAM7uS3Xpu2WYiITjq2hGbCiwiqYKmJnpHRNi0MRua86hGUsxroqqqxrVJVmx2JMlJAhd8VnM2MifRcuZSiHgM7zK5WUj0u1VpbI2geZabFJEhkZ7Ul9HH5JSLyIF6mDjwNc5XBN9Q1xPqakaoa4KvIUwIs+eoqwvdTEgmtMAOB8kjmkPEPMYqk69FFEfAO0XThl27o2+3eIymcoS6ot91dNst89tXOJS6qpg0dfZgSJkUGxAfcN6XdaQb12/54nXfu8HDrb8NeHQGBzlU2a3neYin1cAsG4rPP34ikbZb6BOXpzN2XUvb7nCAryq8z2uabtviXcjKXyXjJy4Xq4UsNddZ2Cc0qIiF3+gcBCdUFy/wMih0aR5fVdgnSKKq/EEypZysvozsHbQkXZ5zJ5JPbBVPX5pLbzYdI5GEgyZaF0pRnWxgzuXmXMs1Qo2ZdBxjRxCjqfL2auxYLRfMX79iPr/h8vwcCXl71Mi6DS4TwkPdICELNIVQE0IYjQ6Xj8sQogvC7PzqQRsbPFKDg/36YHP3yiaTzKP03hN8ha+ySrHZhtX8hqZpuDidICLEmOhjh0MIPhOMxVI2Mu0xdSTL8t6Ix/mqeMhsLKMqMpon+GQpMLw4vBPicNUvhOk27WXuVDS/3xyS8skqIlifM5820LtEcKWrPExmY9p98CKHbJZQOVKvxD6iqc+y8EnR1ENKTCeB4IXYtry6fskXn31Kt9txcjLlxbMnudui7J/icK6mmU7w9QRzFXXVIN7hfbj3Wnuj/jY9ffjGBg8+Cfv7sV3dFi5+HiDf9z1937Jdb9hu1mwWt2zXS9brLX1KuZ7U1Ll1xYykWXtRh0Pp8okVQsD5gNCMHimHoT4nC4o8QyY2H2bmiuoyCdFAJTPccF10RdSndFObk7HgfRiyueD302pSzCn8UhxX8nsGQaG+76FkEL1AXVU0weOdI6Dcvvwdt9dfcP3qNV3fcjKdcXp6SlVVJFNClT1XqCdUzYSqmVI3E6SQCabTk+LRhu0L9+tsZ4/D0AY8qp39MvSbucXU5XS1Zh5h323p10v67YbFYsHt3Zzlck2fElXVUDU1TT3FBg/msmEM9SbM4asJQ4eAFMMbjEFFqScNeUmTkwr3DM88Tie5J1uyLqZHcs8aAEoIByewszJw8uCmMe+P7geE7DsXYFIHqqqiqWq8GLvNmvntLdcvX7Fc3JK6NZV3TJuGyWRC0zSIH1qPhLqZEOoJzXRGqCf40OCrGlfCR+89oRjc4fpTRGgeQcPpm3h0O/xl2K1vbKBXpZQgtgRLxG6XaV9dx263Y75YcHMzZ7Fasl6vcZJpS/V0StM0Y0IlVA1qfmTrm8tGOIgGmXiSaaZUlZMxr8EYCctDRtSROcrOKNNqslGmlBA3sDOGedwDjcsIPg8SqXxJ9HiHD1KUnyPtesV2s2I1n7Na3tJvN4hafl0VOD1/Ml4InMtk6LquR0/vfI34CqlqnA/gKnCBUOfXofYGk+Rhpvu/Kh7tjn8Zuu3cVMlp+pLK1wRJe2KMxfslYoxstivW6zXz2wWLxYLtdotapAoNTdPQNBOqSfEO9TQLpwZP8PW4tlFcCUtzrS1hkPKaTSS32WA2MlZE9l3aOXkyhJWMXsX7LOXqNI2F5th3bLdrVos5i8Wc3XZDt9tQOaEJnklTMa0bJk3NpK7wVY36Bhdq6npCqKtxHZa7IzyhnqAmmPc5SeKrUbTWOUdwmTD9GL3Zu3A8CL8H7erGxBKq5PXQQP7VnKrv+xyGpr6l73u6Xct2u2azyan01d0tMUa6risFYCEUD+h9xXSajXDwiCEEJHiCeFRyNtNkv3a7VxiHMqw+5oxi6ui6vA6NMWIxsrmbo7HPz2kPOFwVaJqGUFXMZnk9Vtd13oaSTaxDhfOBqpmUutqBsfksApu9dk4OZd6oLxSyYRgInD6yNdrvw/Fg/AHYLudmmtn7YgaaskGU9ZGUQYrDCa4xEpxkz9Lu2O12bLfbMTxNfcd6vR1nqMUYidoXrY89az5Lyh0QktnXsFJKo8cbDGK8F+P89IQ6OKpJQ13XhKrCVXVJ3ASM4hHr/FyWLc8GJi5gQ3Z1YIW8sRYbdCqtPDeEn48l6/iH4nhQ/gVoV3MTy9NM8+w3yzodRTdEE7QplUbS/cy3pJEYu5zQkPz/ruvoukybSqlopJgj9XktpmoHSY8y4HEwDJF7oWQIgRACiNCrjYVnH/YliREuG1jl9wJF+TsEXCARxrFZ99tn7OD/eV16NLLfj+MB+iOg3wyGt1fsMtM8isrXedjG2NWdxpBUNZZZ34lhzLCMAqm5ryD43G9Gyvx8N84MGIxgEKJ9k9RcDKOqcxeeGaqxvE5HIx26EA6NZ/8ZeSjIMIH18HukZE1Pjkb2B+F4sP6ESJu59SndM4ZRgHXQFYn75tJhxtxeYkExiUXfhKKTcuCBDj7vTTn0cXpqH9/KEjL0fJdu8WEa0NnReP7kOB7gbwDtam6q2WtNr/Z6K7vbuR2ODDZRCClnJwuNa3LyRHbrvX5jjPHAQ71RFiAPH3GFhTIYcrKDOeTOMXvEafqvG8cD/S1Au5rbmx3bQxp9t5zb5A8k9G42Nwa5Zjc5eTJ+zmiI/n4n9W49t8dcGzviCNrN7b9IqLHfffn729Xc2tWNpc3c4vpf9l1HfHUcr2pfM2zz0sAjs6d/smPftrdWFY5l9J66ftt7pc3cxGJOvpy8e+0WuxsL9ZPjOfJHxKOaLfDtgB9pW38qNM2ViOSRw+8yNgA/u5Tkv8TY2rn5L1ELO+KII/6IsO2N2fblMdT8I+Po4R4xbrU1gLZ/19Qa4RG3Sx5xxL8M2r82271+y7CW9so26froyb4mHC9hjwTvmEgMgLfAzB8L3l8Xjgf6G4Zti3cRh0yOJ/5Dx9HDfUOwzWeGbSHeZIUfd/bu13VzgwiD5HmeBwy1kCeTHo30iCOOOOKII4444ogjjjjiiCOOOOKII4444ogjjjjiiCOOOOKII4444ogjjjjiiCOOOOKII4444ogjjjjiiCOOeBT4PwGtvWhEebSgaAAAAABJRU5ErkJggg==",
  "data:image/png;base64,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",
  "data:image/png;base64,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",
  "data:image/png;base64,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",
];

function SplashScreen() {
  const [frameIdx, setFrameIdx] = useState(0);
  useEffect(() => {
    const t = setInterval(() => setFrameIdx((i) => (i + 1) % JUICE_FRAMES.length), 330);
    return () => clearInterval(t);
  }, []);

  return (
    <div
      className="w-full h-screen flex flex-col items-center overflow-hidden"
      style={{ background: "#FFFDF7", fontFamily: "'Pretendard', -apple-system, 'Apple SD Gothic Neo', sans-serif" }}
    >
      <style>{`
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css');
        @keyframes splashFade { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .splash-fade { animation: splashFade 0.6s ease both; }
      `}</style>

      {/* 상단: 로고 (여백을 위로 밀어 화면 위쪽 1/3 지점에 위치) */}
      <div className="splash-fade text-center" style={{ marginTop: "18vh" }}>
        <div className="text-4xl font-extrabold tracking-tight mb-2" style={{ color: "#3A2317" }}>
          오늘부터
        </div>
        <div className="text-xs font-semibold" style={{ color: "#8A7D6E" }}>
          실패해도 괜찮아요, 오늘부터 다시
        </div>
      </div>

      {/* 하단: 컵 애니메이션 + 제작자 표시를 화면 아래쪽에 함께 배치 */}
      <div className="splash-fade text-center" style={{ marginTop: "auto", marginBottom: "8vh" }}>
        <div className="relative mx-auto mb-3" style={{ width: 110, height: 95 }}>
          <img
            src={JUICE_FRAMES[frameIdx]}
            alt=""
            className="absolute top-0 left-0 w-full h-full"
            style={{ objectFit: "contain" }}
          />
        </div>
        <div className="mx-auto mb-4" style={{ width: 32, height: 3, borderRadius: 2, background: "#FEE500" }} />
        <div className="text-[11px] font-bold tracking-[0.2em] mb-1" style={{ color: "#B4A891" }}>
          MADE BY
        </div>
        <div className="text-base font-extrabold" style={{ color: "#3A2317" }}>
          AYoom
        </div>
      </div>
    </div>
);
}

/* ------------------------------------------------------------------ */
/*  앱 프레임 (Shell)                                                   */
/* ------------------------------------------------------------------ */

function Shell({ data, toast, footer, children }) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const scale = data.fontScale || 1;

  // Tailwind의 text-xs/text-sm 등은 rem 단위라 부모 요소의 em 크기를 바꿔도 영향을 받지 않는다.
  // 실제로 글자 크기를 바꾸려면 rem의 기준이 되는 문서 루트(html)의 font-size 자체를 바꿔야 한다.
  useEffect(() => {
    document.documentElement.style.fontSize = `${16 * scale}px`;
    return () => {
      document.documentElement.style.fontSize = "16px";
    };
  }, [scale]);

  return (
    <div
      className="w-full h-screen flex items-stretch justify-center overflow-hidden"
      style={{ background: data.darkMode ? "#0f0d0a" : "#EDE7D6", fontFamily: "'Pretendard', -apple-system, 'Apple SD Gothic Neo', sans-serif" }}
    >
      <style>{`
        html, body, #root { height: 100%; margin: 0; }
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { display: none; }
        @keyframes riseUp { from { opacity:0; transform: translateY(8px);} to {opacity:1; transform:translateY(0);} }
        @keyframes toastIn { from {opacity:0; transform: translate(-50%, 10px);} to {opacity:1; transform: translate(-50%,0);} }
        .fade-in { animation: riseUp 0.35s ease both; }
      `}</style>
      <div
        className="relative w-full max-w-[430px] h-full overflow-hidden flex flex-col"
        style={{ background: c.bg, color: c.ink }}
      >
        {/* 스크롤되는 화면 콘텐츠 영역 - 하단바는 이 안에 포함되지 않는다 */}
        <div className="flex-1 overflow-y-auto relative min-h-0">{children}</div>
        {/* 하단바는 스크롤 영역 바깥, 별도의 고정 영역으로 렌더링된다 */}
        {footer}
        {toast && (
          <div
            className="absolute left-1/2 bottom-24 px-4 py-2.5 rounded-full text-sm font-semibold shadow-lg z-50"
            style={{ background: BRAND_DARK, color: c.yellow, animation: "toastIn 0.25s ease both", transform: "translateX(-50%)" }}
          >
            {toast}
          </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  온보딩                                                             */
/* ------------------------------------------------------------------ */

function Onboarding({ data, onComplete }) {
  const c = COLORS.light;
  const [mealStart, setMealStart] = useState("12:00");
  const [gender, setGender] = useState("female");
  const presets = ["08:00", "12:00", "17:30"];
  const window_ = getMealWindow(new Date(), mealStart);

  return (
    <div className="h-full flex flex-col px-7 pt-16 pb-8 fade-in" style={{ color: c.ink }}>
      <div
        className="w-14 h-14 rounded-2xl flex items-center justify-center text-2xl mb-6"
        style={{ background: c.yellow }}
      >
        🌤️
      </div>
      <h1 className="text-2xl font-extrabold mb-2 leading-snug">
        오늘부터, 다시
        <br />
        시작해볼까요?
      </h1>
      <p className="text-sm mb-8" style={{ color: c.inkSoft }}>
        식사를 시작할 시간만 정해주세요.
        <br />
        7시간 뒤 자동으로 식사 시간이 끝나요.
      </p>

      <div className="text-xs font-bold mb-2" style={{ color: c.inkSoft }}>
        식사 시작 시간
      </div>
      <div className="flex gap-2 mb-4">
        {presets.map((p) => (
          <button
            key={p}
            onClick={() => setMealStart(p)}
            className="flex-1 py-2.5 rounded-xl text-sm font-bold transition"
            style={{
              background: mealStart === p ? c.yellow : c.cardMuted,
              color: c.ink,
            }}
          >
            {p}
          </button>
        ))}
      </div>
      <input
        type="time"
        value={mealStart}
        onChange={(e) => setMealStart(e.target.value)}
        className="w-full py-3 px-4 rounded-xl text-base font-semibold mb-6 outline-none"
        style={{ background: c.cardMuted, color: c.ink, border: `1px solid ${c.line}` }}
      />

      <div
        className="rounded-2xl p-4 mb-8 text-sm font-semibold flex items-center justify-between"
        style={{ background: c.yellowSoft }}
      >
        <span>식사 가능 시간</span>
        <span>
          {fmtHM(window_.start)} ~ {fmtHM(window_.end)}
        </span>
      </div>

      <div className="text-xs font-bold mb-2" style={{ color: c.inkSoft }}>
        변화 기록 촬영 가이드용 성별
      </div>
      <div className="flex gap-2 mb-10">
        {[
          { key: "female", label: "여성" },
          { key: "male", label: "남성" },
        ].map((g) => (
          <button
            key={g.key}
            onClick={() => setGender(g.key)}
            className="flex-1 py-2.5 rounded-xl text-sm font-bold"
            style={{ background: gender === g.key ? c.yellow : c.cardMuted }}
          >
            {g.label}
          </button>
        ))}
      </div>

      <div
        className="rounded-2xl p-4 mb-8 text-xs leading-relaxed"
        style={{ background: c.cardMuted }}
      >
        <div className="font-bold mb-1">참고로 알려드려요</div>
        <span style={{ color: c.inkSoft }}>
          월·목요일엔 음식 사진을, 수·토요일엔 체중 기록을 특히 챙기도록 홈 화면에서 알려드려요. 다른 요일도 기록은 언제든 자유롭게 할 수 있어요.
        </span>
      </div>

      <button
        onClick={() => onComplete(mealStart, gender)}
        className="mt-auto w-full py-4 rounded-2xl font-extrabold text-base"
        style={{ background: BRAND_DARK, color: c.yellow }}
      >
        시작하기
      </button>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  홈 화면                                                            */
/* ------------------------------------------------------------------ */

function HomeScreen({ data, setData, addPoints, goto }) {
  const now = useNow();
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const day = now.getDay(); // 0 Sun ... 6 Sat
  const isFoodDay = day === 1 || day === 4;
  const isWeightDay = day === 3 || day === 6;
  const today = todayStr(now);

  const { phase, window: mealWindow, nextStart } = getPhase(now, data.mealStart);

  const foodToday = data.logs.food.filter((f) => f.date === today);
  const weightToday = data.logs.weight.filter((w) => w.date === today);
  const exerciseToday = data.logs.exercise.filter((e) => e.date === today);

  const mealKept = foodToday.some((f) => f.inMealWindow);

  const homeItems = [
    { key: "meal", label: "식사시간 지키기", short: "식사시간", icon: "⏰", done: mealKept, emphasize: false, onClick: () => goto("food") },
    { key: "food", label: "음식 기록", short: "음식", icon: "🍽️", done: foodToday.length > 0, emphasize: isFoodDay, onClick: () => goto("food") },
    { key: "exercise", label: "운동 기록", short: "운동", icon: "🏃", done: exerciseToday.length > 0, emphasize: false, onClick: () => goto("exercise") },
    { key: "weight", label: "체중 기록", short: "체중", icon: "⚖️", done: weightToday.length > 0, emphasize: isWeightDay, onClick: () => goto("weight") },
  ];

  // 진행률 & 표정
  let progress = 0;
  let face = "😴";
  let title = "";
  let subtitle = "";
  if (phase === "meal") {
    progress = Math.min(1, (now - mealWindow.start) / (MEAL_HOURS * 3600 * 1000));
    const remain = mealWindow.end - now;
    face = remain < 30 * 60 * 1000 ? "🙂" : "😊";
    title = "지금은 식사 가능 시간이에요";
    subtitle = `${formatDur(remain)} 후 식사 시간이 끝나요`;
  } else {
    const total = FASTING_HOURS * 3600 * 1000;
    const remain = nextStart - now;
    progress = Math.min(1, 1 - remain / total);
    face = remain < 60 * 60 * 1000 ? "🙂" : "😴";
    title = "지금은 공복 시간이에요";
    subtitle = `${formatDur(remain)} 후 식사가 가능해요`;
  }

  const r = 78;
  const circumference = 2 * Math.PI * r;
  const dash = circumference * progress;

  // 재방문 메시지 로직
  const gapDays = data.lastActiveDate ? diffDaysStr(data.lastActiveDate, today) : 0;
  const [dismissedMild, setDismissedMild] = useState(false);

  const handleRestart = () => {
    setData((p) => ({ ...p, lastActiveDate: today }));
    addPoints(10, "루틴 재시작");
  };

  const [showWeekInfo, setShowWeekInfo] = useState(false);

  return (
    <div className="px-6 pt-6 pb-24 fade-in">
      <div className="flex items-center justify-between mb-4">
        <button onClick={() => setShowWeekInfo(true)} className="text-left">
          <div className="text-xs font-bold flex items-center gap-1" style={{ color: c.inkSoft }}>
            {now.toLocaleDateString("ko-KR", { month: "long", day: "numeric", weekday: "long" })}
            <Info size={11} style={{ opacity: 0.6 }} />
          </div>
          <div className="text-lg font-extrabold mt-0.5" style={{ color: c.ink }}>오늘부터</div>
        </button>
        <div
          className="px-3 py-1.5 rounded-full text-xs font-extrabold flex items-center gap-1"
          style={{ background: c.yellow, color: "#3A2317" }}
        >
          ⭐ {data.points}P
        </div>
      </div>

      {gapDays >= 5 && (
        <div
          className="rounded-2xl p-4 mb-4 text-center"
          style={{ background: c.cardMuted }}
        >
          <div className="text-2xl mb-1">👋</div>
          <div className="font-extrabold mb-1 text-sm">다시 오셨네요. 오늘부터 다시 시작해볼까요?</div>
          <button
            onClick={handleRestart}
            className="mt-2 px-5 py-2 rounded-full font-bold text-xs"
            style={{ background: c.yellow, color: "#3A2317" }}
          >
            오늘부터 다시 시작하기
          </button>
        </div>
      )}

      {gapDays >= 1 && gapDays < 5 && !dismissedMild && (
        <div
          className="rounded-xl p-3 mb-4 flex items-center justify-between gap-3"
          style={{ background: c.cardMuted }}
        >
          <div className="text-xs font-semibold leading-snug">
            어제는 기록이 없었어요. <span style={{ color: c.inkSoft }}>괜찮아요, 오늘 다시 이어가요.</span>
          </div>
          <button onClick={() => setDismissedMild(true)} className="opacity-50 shrink-0">
            <X size={14} />
          </button>
        </div>
      )}

      {/* 단식 타이머 링 */}
      <div className="flex flex-col items-center py-1 mb-5">
        <div className="relative w-44 h-44 flex items-center justify-center">
          <svg width="176" height="176" className="-rotate-90 absolute">
            <circle cx="88" cy="88" r={r} stroke={c.line} strokeWidth="10" fill="none" />
            <circle
              cx="88"
              cy="88"
              r={r}
              stroke={c.yellow}
              strokeWidth="10"
              fill="none"
              strokeLinecap="round"
              strokeDasharray={`${dash} ${circumference}`}
              style={{ transition: "stroke-dasharray 0.6s ease" }}
            />
          </svg>
          {(isFoodDay || isWeightDay) && (
            <div
              className="absolute rounded-full flex items-center justify-center shadow"
              style={{ top: 2, right: 6, width: 26, height: 26, background: c.card, fontSize: 13 }}
            >
              {isFoodDay ? "📸" : "⚖️"}
            </div>
          )}
          <div className="flex flex-col items-center">
            <div className="text-4xl mb-1">{face}</div>
            <div className="text-[11px] font-bold" style={{ color: c.inkSoft }}>
              {phase === "meal" ? "식사 가능" : "공복 중"}
            </div>
            {(isFoodDay || isWeightDay) && (
              <div className="text-[11px] font-bold mt-0.5" style={{ color: "#D9A400" }}>
                {isFoodDay ? "오늘은 사진 찍는 날" : "오늘은 체중 재는 날"}
              </div>
            )}
          </div>
        </div>
        <div className="mt-3 text-center">
          <div className="font-extrabold text-sm">{title}</div>
          <div className="text-xs mt-0.5" style={{ color: c.inkSoft }}>
            {subtitle}
          </div>
        </div>
      </div>

      {/* 오늘의 루틴 */}
      <div className="text-[11px] font-bold mb-2 px-0.5" style={{ color: c.inkSoft }}>오늘의 루틴</div>
      <div className="grid grid-cols-2 gap-2">
        {homeItems.map((it) => (
          <button
            key={it.key}
            onClick={it.onClick}
            className="relative flex items-center gap-2.5 p-3 rounded-xl text-left"
            style={{ background: c.card, border: `1px solid ${c.line}` }}
          >
            <div
              className="w-6 h-6 rounded-full flex items-center justify-center shrink-0"
              style={{ background: it.done ? c.green : c.cardMuted }}
            >
              {it.done && <Check size={13} color="#fff" strokeWidth={3} />}
            </div>
            <span
              className="text-xs font-semibold flex-1"
              style={{ textDecoration: it.done ? "line-through" : "none", opacity: it.done ? 0.55 : 1 }}
            >
              {it.label}
            </span>
            {it.emphasize && (
              <span
                onClick={(e) => {
                  e.stopPropagation();
                  setShowWeekInfo(true);
                }}
                className="w-4 h-4 rounded-full flex items-center justify-center shrink-0 text-[9px] font-extrabold"
                style={{ background: c.yellowSoft, color: "#8A6A00" }}
              >
                ?
              </span>
            )}
          </button>
        ))}
      </div>

      {showWeekInfo && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center px-8"
          style={{ background: "rgba(0,0,0,0.4)" }}
          onClick={() => setShowWeekInfo(false)}
        >
          <div
            className="w-full max-w-xs rounded-2xl p-5"
            style={{ background: c.card }}
            onClick={(e) => e.stopPropagation()}
          >
            <div className="font-extrabold text-sm mb-3">요일별 루틴 안내</div>
            <div className="flex flex-col gap-2 mb-4">
              {[
                { d: "월요일", label: "🍽️ 음식 사진 강조" },
                { d: "화요일", label: "자유롭게 기록" },
                { d: "수요일", label: "⚖️ 체중 기록 강조" },
                { d: "목요일", label: "🍽️ 음식 사진 강조" },
                { d: "금요일", label: "자유롭게 기록" },
                { d: "토요일", label: "⚖️ 체중 기록 강조" },
                { d: "일요일", label: "자유롭게 기록" },
              ].map((row) => (
                <div key={row.d} className="flex items-center justify-between text-xs">
                  <span className="font-bold" style={{ color: c.inkSoft }}>{row.d}</span>
                  <span className="font-semibold">{row.label}</span>
                </div>
              ))}
            </div>
            <div className="text-[11px] mb-4" style={{ color: c.inkSoft }}>
              강조 요일이 아니어도 기록은 언제든 할 수 있어요. 다만 그 요일엔 홈에서 조금 더 챙기시라고 안내해드려요.
            </div>
            <button
              onClick={() => setShowWeekInfo(false)}
              className="w-full py-2.5 rounded-xl text-sm font-extrabold"
              style={{ background: BRAND_DARK, color: c.yellow }}
            >
              확인했어요
            </button>
          </div>
        </div>
      )}

      {/* 오늘 쌓은 기록 */}
      <div className="mt-4">
        <div className="text-[11px] font-bold mb-2" style={{ color: c.inkSoft }}>오늘 쌓은 기록</div>
        <div className="flex justify-between px-1">
          {homeItems.map((it) => (
            <div key={it.key} className="flex flex-col items-center gap-1">
              <div
                className="w-9 h-9 rounded-full flex items-center justify-center text-sm"
                style={{
                  background: it.done ? c.yellow : c.cardMuted,
                  border: it.done ? "none" : `1.5px dashed ${c.line}`,
                  opacity: it.done ? 1 : 0.5,
                }}
              >
                {it.icon}
              </div>
              <span className="text-[9px] font-bold" style={{ color: c.inkSoft }}>{it.short}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  기록 허브 (음식 · 운동 · 체중 · 변화 기록 진입점)                     */
/* ------------------------------------------------------------------ */

function RecordHubScreen({ data, goto, goBack }) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const now = new Date();
  const today = todayStr(now);
  const foodCount = data.logs.food.filter((f) => f.date === today).length;
  const exerciseCount = data.logs.exercise.filter((e) => e.date === today).length;
  const weightDone = data.logs.weight.some((w) => w.date === today);

  const cards = [
    { key: "food", icon: "🍽️", title: "음식 기록", desc: foodCount > 0 ? `오늘 ${foodCount}건 기록했어요` : "먹을 때마다 기록해요", onClick: () => goto("food") },
    { key: "exercise", icon: "🏃", title: "운동 기록", desc: exerciseCount > 0 ? `오늘 ${exerciseCount}건 기록했어요` : "여러 운동을 함께 기록할 수 있어요", onClick: () => goto("exercise") },
    { key: "weight", icon: "⚖️", title: "체중 기록", desc: weightDone ? "오늘 기록을 남겼어요" : "기상 직후가 가장 정확해요", onClick: () => goto("weight") },
    { key: "body", icon: "📷", title: "변화 기록", desc: "곧 추가될 예정이에요 · 하단 탭에서 미리보기", onClick: () => goto("bodyPhoto") },
  ];

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="기록" onBack={goBack} />
      <div className="px-6 pt-4 flex flex-col gap-3">
        {cards.map((card) => (
          <button
            key={card.key}
            onClick={card.onClick}
            disabled={card.disabled}
            className="flex items-center gap-4 p-4 rounded-2xl text-left"
            style={{ background: c.card, border: `1px solid ${c.line}`, opacity: card.disabled ? 0.5 : 1 }}
          >
            <div className="w-11 h-11 rounded-xl flex items-center justify-center text-xl shrink-0" style={{ background: c.cardMuted }}>
              {card.icon}
            </div>
            <div>
              <div className="text-sm font-extrabold mb-0.5">{card.title}</div>
              <div className="text-xs font-semibold" style={{ color: c.inkSoft }}>{card.desc}</div>
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  음식 기록 화면                                                     */
/* ------------------------------------------------------------------ */

function WeightLinkCard({ data, c, goToWeight, today }) {
  const doneToday = data.logs.weight.some((w) => w.date === today);
  return (
    <button
      onClick={goToWeight}
      className="w-full flex items-center gap-3 p-3.5 rounded-2xl mt-2 mb-8"
      style={{ background: c.cardMuted }}
    >
      <span className="text-xl">⚖️</span>
      <div className="flex-1 text-left">
        <div className="text-xs font-bold">체중도 기록해보세요</div>
        <div className="text-[11px]" style={{ color: c.inkSoft }}>
          {doneToday ? "오늘 이미 기록했어요" : "아직 오늘 기록 전이에요"}
        </div>
      </div>
      <ChevronLeft size={16} style={{ transform: "rotate(180deg)", color: c.inkSoft }} />
    </button>
  );
}

function FoodScreen({ data, setData, addPoints, goBack, goToWeight }) {
  const now = useNow();
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const [note, setNote] = useState("");
  const [time, setTime] = useState(timeStr(now));
  const [preview, setPreview] = useState(null);
  const [analyzing, setAnalyzing] = useState(false);
  const fileRef = useRef(null);

  const today = todayStr(now);
  const list = data.logs.food.filter((f) => f.date === today).slice().reverse();

  const handleFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async () => {
      const resized = await resizeImage(reader.result);
      setPreview(resized);
    };
    reader.readAsDataURL(file);
  };

  const deleteEntry = (id) => {
    setData((p) => ({ ...p, logs: { ...p.logs, food: p.logs.food.filter((f) => f.id !== id) } }));
  };

  const analyzeFood = async (image) => {
    if (!image) return null;
    try {
      const res = await fetch("/api/analyze-food", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ image }),
      });
      if (!res.ok) return null;
      const data = await res.json();
      return { text: data.comment, tag: data.tag };
    } catch {
      return null;
    }
  };

  const save = async () => {
    if (analyzing) return;
    setAnalyzing(true);
    const analyzed = await analyzeFood(preview);
    const tpl = analyzed || FOOD_TEMPLATES[Math.floor(Math.random() * FOOD_TEMPLATES.length)];
    const kept = isMealTimeKept(today, time, data.mealStart);
    const entry = {
      id: uid(),
      date: today,
      time,
      note,
      inMealWindow: kept,
      analysis: tpl.text,
      tag: tpl.tag,
      image: preview,
    };
    setData((p) => ({ ...p, logs: { ...p.logs, food: [...p.logs.food, entry] }, lastActiveDate: today }));
    addPoints(20, "음식 기록");
    if (kept) addPoints(5, "식사시간 지킴");
    setNote("");
    setPreview(null);
    setTime(timeStr(new Date()));
    setAnalyzing(false);
  };

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="음식 기록" onBack={goBack} />
      <div className="px-6 pt-4">
        <div
          onClick={() => fileRef.current?.click()}
          className="w-full h-44 rounded-2xl flex flex-col items-center justify-center gap-2 mb-4 cursor-pointer overflow-hidden"
          style={{ background: c.cardMuted, border: `1.5px dashed ${c.line}` }}
        >
          {preview ? (
            <img src={preview} className="w-full h-full object-cover" alt="음식 사진" />
          ) : (
            <>
              <Camera size={28} style={{ color: c.inkSoft }} />
              <span className="text-sm font-semibold" style={{ color: c.inkSoft }}>
                먹은 음식 사진 올리기
              </span>
            </>
          )}
        </div>
        <input ref={fileRef} type="file" accept="image/*" capture="environment" onChange={handleFile} className="hidden" />

        <div className="flex items-center justify-between mb-4">
          <span className="text-sm font-bold" style={{ color: c.inkSoft }}>촬영 시간</span>
          <input
            type="time"
            value={time}
            onChange={(e) => setTime(e.target.value)}
            className="px-3 py-1.5 rounded-lg text-sm font-semibold outline-none"
            style={{ background: c.cardMuted, color: c.ink }}
          />
        </div>

        <textarea
          value={note}
          onChange={(e) => setNote(e.target.value)}
          placeholder="메모 (선택)"
          rows={2}
          className="w-full p-3 rounded-xl text-sm outline-none mb-4 resize-none"
          style={{ background: c.cardMuted, color: c.ink }}
        />

        <div
          className="rounded-xl px-3 py-2 mb-6 text-xs font-bold text-center"
          style={{ background: isMealTimeKept(today, time, data.mealStart) ? c.yellowSoft : c.cardMuted, color: c.ink }}
        >
          {isMealTimeKept(today, time, data.mealStart)
            ? "식사 가능 시간 안에 먹은 걸로 기록돼요 (앞뒤 30분까지 괜찮아요)"
            : "이 시간은 식사 가능 시간 밖이지만, 기록은 언제든 가능해요"}
        </div>

        <button
          onClick={save}
          disabled={analyzing}
          className="w-full py-3.5 rounded-2xl font-extrabold text-sm mb-8 flex items-center justify-center gap-2 disabled:opacity-70"
          style={{ background: BRAND_DARK, color: c.yellow }}
        >
          {analyzing ? (
            <>
              <Loader2 size={16} className="animate-spin" />
              사진 분석 중...
            </>
          ) : (
            "기록 저장하기"
          )}
        </button>

        {list.length > 0 && (
          <>
            <div className="text-xs font-bold mb-3" style={{ color: c.inkSoft }}>
              오늘 기록 {list.length}건
            </div>
            <div className="flex flex-col gap-3 mb-8">
              {list.map((f) => (
                <div key={f.id} className="rounded-2xl p-3 flex gap-3" style={{ background: c.card, border: `1px solid ${c.line}` }}>
                  {f.image ? (
                    <img src={f.image} className="w-14 h-14 rounded-xl object-cover shrink-0" />
                  ) : (
                    <div className="w-14 h-14 rounded-xl shrink-0 flex items-center justify-center text-lg" style={{ background: c.cardMuted }}>
                      🍽️
                    </div>
                  )}
                  <div className="flex-1 min-w-0">
                    <div className="text-xs font-bold mb-1" style={{ color: c.inkSoft }}>{f.time}</div>
                    <div className="text-sm font-semibold leading-snug">{f.analysis}</div>
                  </div>
                  <button
                    onClick={() => {
                      if (window.confirm("이 기록을 삭제할까요?")) deleteEntry(f.id);
                    }}
                    className="shrink-0 w-7 h-7 rounded-full flex items-center justify-center"
                    style={{ background: c.cardMuted, color: c.inkSoft }}
                    aria-label="기록 삭제"
                  >
                    <X size={14} />
                  </button>
                </div>
              ))}
            </div>
          </>
        )}

        <WeightLinkCard data={data} c={c} goToWeight={goToWeight} today={today} />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  운동 기록 화면 (독립)                                               */
/* ------------------------------------------------------------------ */

function ExerciseScreen({ data, setData, addPoints, goBack, goToWeight }) {
  const now = useNow();
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const today = todayStr(now);
  const [types, setTypes] = useState([]);
  const [minutes, setMinutes] = useState(30);
  const list = data.logs.exercise.filter((e) => e.date === today);

  const toggleType = (t) => {
    setTypes((prev) => (prev.includes(t) ? prev.filter((x) => x !== t) : [...prev, t]));
  };

  const save = () => {
    if (types.length === 0) return;
    const isFirstToday = list.length === 0;
    const entry = { id: uid(), date: today, types, minutes, intensity: "보통" };
    setData((p) => ({ ...p, logs: { ...p.logs, exercise: [...p.logs.exercise, entry] }, lastActiveDate: today }));
    if (isFirstToday) {
      addPoints(10, "운동 기록");
    }
    setTypes([]);
    setMinutes(30);
  };

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="운동 기록" onBack={goBack} />
      <div className="px-6 pt-4">
        <div className="flex items-center gap-2 mb-3">
          <Dumbbell size={16} />
          <span className="text-sm font-bold">오늘 한 운동을 모두 선택해주세요</span>
        </div>
        <div className="flex gap-2 mb-4 flex-wrap">
          {EXERCISE_TYPES.map((t) => {
            const active = types.includes(t);
            return (
              <button
                key={t}
                onClick={() => toggleType(t)}
                className="px-3 py-1.5 rounded-full text-xs font-bold flex items-center gap-1"
                style={{ background: active ? c.yellow : c.cardMuted, color: active ? BRAND_DARK : c.ink }}
              >
                {active && <Check size={11} strokeWidth={3} />}
                {t}
              </button>
            );
          })}
        </div>
        <div className="flex items-center gap-3 mb-5">
          <input
            type="range"
            min={10}
            max={120}
            step={5}
            value={minutes}
            onChange={(e) => setMinutes(Number(e.target.value))}
            className="flex-1"
          />
          <span className="text-sm font-bold w-14 text-right">{minutes}분</span>
        </div>
        <button
          onClick={save}
          disabled={types.length === 0}
          className="w-full py-3.5 rounded-2xl text-sm font-extrabold mb-3"
          style={{ background: BRAND_DARK, color: c.yellow, opacity: types.length === 0 ? 0.4 : 1 }}
        >
          운동 기록 추가
        </button>
        <div className="text-[11px] text-center mb-6" style={{ color: c.inkSoft }}>
          {list.length === 0 ? "오늘 첫 기록에만 +10P가 붙어요" : "오늘은 이미 포인트를 받았어요 · 기록은 계속 쌓여요"}
        </div>

        {list.length > 0 && (
          <>
            <div className="text-xs font-bold mb-3" style={{ color: c.inkSoft }}>오늘 기록 {list.length}건</div>
            <div className="flex flex-col gap-2">
              {list.map((e) => (
                <div key={e.id} className="rounded-2xl p-3.5 text-sm font-semibold" style={{ background: c.card, border: `1px solid ${c.line}` }}>
                  🏃 {(e.types || [e.type]).join(", ")} · {e.minutes}분
                </div>
              ))}
            </div>
          </>
        )}

        <WeightLinkCard data={data} c={c} goToWeight={goToWeight} today={today} />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  변화 기록 (준비 중)                                                 */
/* ------------------------------------------------------------------ */

function BodyPhotoScreen({ data, goBack }) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  return (
    <div className="pb-10 fade-in h-full flex flex-col">
      <TopBar c={c} title="변화 기록" onBack={goBack} />
      <div className="flex-1 flex flex-col items-center justify-center px-8 text-center">
        <div className="w-16 h-16 rounded-2xl flex items-center justify-center text-3xl mb-4" style={{ background: c.cardMuted }}>
          📷
        </div>
        <div className="font-extrabold mb-2">곧 만나볼 수 있어요</div>
        <div className="text-sm leading-relaxed" style={{ color: c.inkSoft }}>
          {data.gender === "male" ? "남성" : "여성"} 실루엣 가이드에 맞춰 정면·측면·후면을 찍고,
          <br />
          날짜를 골라 빠르게 넘겨보는 타임랩스가 여기에 들어올 예정이에요.
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  체중 기록 화면                                                     */
/* ------------------------------------------------------------------ */

function WeightScreen({ data, setData, addPoints, goBack, weightRevealed, setWeightRevealed }) {
  const now = useNow();
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const today = todayStr(now);
  const todayEntry = data.logs.weight.find((w) => w.date === today);
  const [value, setValue] = useState(todayEntry ? String(todayEntry.value) : "");
  const sorted = data.logs.weight.slice().sort((a, b) => (a.date < b.date ? -1 : 1));
  const lastEntry = sorted[sorted.length - 1];
  const isVisible = !data.weightHidden || weightRevealed;
  const dirMap = useMemo(() => buildWeightDirMap(data.logs.weight), [data.logs.weight]);

  const save = () => {
    const v = parseFloat(value);
    if (!v) return;
    const isNew = !todayEntry;
    const entry = { id: todayEntry?.id || uid(), date: today, value: v };
    const others = data.logs.weight.filter((w) => w.date !== today);
    setData((p) => ({ ...p, logs: { ...p.logs, weight: [...others, entry] }, lastActiveDate: today }));
    if (isNew) {
      addPoints(20, "체중 기록");
    } else {
      addPoints(0, "체중을 수정했어요");
    }
  };

  // ---- 지난 날짜 백필용 달력 (월 이동 가능) ----
  const [calendarOpen, setCalendarOpen] = useState(false);
  const todayDate = new Date();
  const [calYear, setCalYear] = useState(todayDate.getFullYear());
  const [calMonth, setCalMonth] = useState(todayDate.getMonth());
  const isCurrentMonth = calYear === todayDate.getFullYear() && calMonth === todayDate.getMonth();
  const goPrevMonth = () => {
    if (calMonth === 0) { setCalYear((y) => y - 1); setCalMonth(11); } else { setCalMonth((m) => m - 1); }
  };
  const goNextMonth = () => {
    if (isCurrentMonth) return;
    if (calMonth === 11) { setCalYear((y) => y + 1); setCalMonth(0); } else { setCalMonth((m) => m + 1); }
  };
  const daysInMonth = new Date(calYear, calMonth + 1, 0).getDate();
  const firstWeekday = new Date(calYear, calMonth, 1).getDay();
  const weightByDate = useMemo(() => {
    const map = {};
    data.logs.weight.forEach((w) => (map[w.date] = w));
    return map;
  }, [data.logs.weight]);

  const [selectedDate, setSelectedDate] = useState(null);
  const [backfillValue, setBackfillValue] = useState("");
  const [backfillPreview, setBackfillPreview] = useState(null);
  const backfileRef = useRef(null);

  const openDate = (dateStr) => {
    if (selectedDate === dateStr) {
      setSelectedDate(null);
      return;
    }
    setSelectedDate(dateStr);
    const existing = weightByDate[dateStr];
    setBackfillValue(existing ? String(existing.value) : "");
    setBackfillPreview(existing?.image || null);
  };

  const handleBackfillFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => setBackfillPreview(reader.result);
    reader.readAsDataURL(file);
  };

  const saveBackfill = () => {
    if (!selectedDate) return;
    const v = parseFloat(backfillValue);
    const existedBefore = data.logs.weight.some((w) => w.date === selectedDate);
    const entry = { id: uid(), date: selectedDate, value: v || null, image: backfillPreview || null };
    if (!v && !backfillPreview) return;
    const others = data.logs.weight.filter((w) => w.date !== selectedDate);
    setData((p) => ({ ...p, logs: { ...p.logs, weight: [...others, entry] }, lastActiveDate: today }));
    if (selectedDate === today && !existedBefore) {
      addPoints(20, "체중 기록");
    } else if (selectedDate === today) {
      addPoints(0, "체중을 수정했어요");
    } else {
      addPoints(0, "지난 날짜의 기록을 채워 넣었어요");
    }
    setSelectedDate(null);
    setBackfillValue("");
    setBackfillPreview(null);
  };

  const cells = [];
  for (let i = 0; i < firstWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="체중 기록" onBack={goBack} />
      <div className="px-6 pt-4">
        <div className="text-xs font-bold mb-2 text-center" style={{ color: c.inkSoft }}>
          일어나서 화장실을 다녀온 후가 가장 정확해요
        </div>

        <div className="rounded-2xl p-6 mb-5 text-center" style={{ background: c.cardMuted }}>
          {lastEntry ? (
            <>
              <button
                onClick={() => data.weightHidden && setWeightRevealed((v) => !v)}
                className="text-3xl font-extrabold mb-1 flex items-center justify-center gap-2 mx-auto"
              >
                {!isVisible ? "••.•kg" : `${lastEntry.value}kg`}
                {data.weightHidden && (isVisible ? <Eye size={18} /> : <EyeOff size={18} />)}
              </button>
              {data.weightHidden && (
                <div className="text-[11px] font-semibold mb-1" style={{ color: c.inkSoft }}>
                  {isVisible ? "탭하면 다시 가려요 · 통계에서도 함께 보여요" : "탭하면 보여요"}
                </div>
              )}
              {dirMap[lastEntry.date] && (
                <div className="text-xs font-semibold flex items-center justify-center gap-1" style={{ color: c.inkSoft }}>
                  {dirMap[lastEntry.date].dir !== "first" && "직전 기록 대비"}
                  <WeightDirBadge info={dirMap[lastEntry.date]} c={c} size={11} />
                </div>
              )}
            </>
          ) : (
            <div className="text-sm font-semibold" style={{ color: c.inkSoft }}>
              아직 기록이 없어요
            </div>
          )}
        </div>

        <div className="flex gap-2 mb-4">
          <input
            type="number"
            step="0.1"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            placeholder="오늘 체중 (kg)"
            className="flex-1 px-4 py-3 rounded-xl text-base font-semibold outline-none"
            style={{ background: c.cardMuted, color: c.ink }}
          />
          <button onClick={save} className="px-5 rounded-xl font-extrabold text-sm" style={{ background: BRAND_DARK, color: c.yellow }}>
            {todayEntry ? "수정하기" : "저장"}
          </button>
        </div>
        {todayEntry && (
          <div className="text-[11px] font-semibold mb-2" style={{ color: c.inkSoft }}>
            오늘 이미 기록했어요 · 값을 고치면 오늘 기록이 업데이트돼요 (포인트는 다시 안 쌓여요)
          </div>
        )}

        <div className="text-xs mb-6" style={{ color: c.inkSoft }}>
          체중은 민감한 정보라 기본적으로 가려져 있어요. 늘었는지 줄었는지 방향(▲▼)은 가림과 상관없이 항상 볼 수 있어요.
        </div>

        <button
          onClick={() => setCalendarOpen((v) => !v)}
          className="w-full flex items-center justify-between p-3.5 rounded-2xl mb-3"
          style={{ background: c.card, border: `1px solid ${c.line}` }}
        >
          <span className="text-xs font-bold">깜빡하고 놓친 날짜, 나중에 채워 넣기</span>
          <span className="text-xs" style={{ color: c.inkSoft }}>{calendarOpen ? "▲" : "▼"}</span>
        </button>

        {calendarOpen && (
          <div className="rounded-2xl p-4 mb-6" style={{ background: c.card, border: `1px solid ${c.line}` }}>
            <div className="flex items-center justify-between mb-2">
              <button onClick={goPrevMonth} className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: c.cardMuted }}>
                <ChevronLeft size={14} />
              </button>
              <span className="text-xs font-bold">{calYear}년 {calMonth + 1}월</span>
              <button
                onClick={goNextMonth}
                disabled={isCurrentMonth}
                className="w-7 h-7 rounded-full flex items-center justify-center"
                style={{ background: c.cardMuted, opacity: isCurrentMonth ? 0.3 : 1 }}
              >
                <ChevronLeft size={14} style={{ transform: "rotate(180deg)" }} />
              </button>
            </div>
            <div className="grid grid-cols-7 gap-1 mb-1">
              {["일", "월", "화", "수", "목", "금", "토"].map((d) => (
                <div key={d} className="text-center text-[10px] font-bold py-1" style={{ color: c.inkSoft }}>
                  {d}
                </div>
              ))}
            </div>
            <div className="grid grid-cols-7 gap-1 mb-2">
              {cells.map((d, i) => {
                if (!d) return <div key={i} />;
                const dateStr = `${calYear}-${pad(calMonth + 1)}-${pad(d)}`;
                const entry = weightByDate[dateStr];
                const isFuture = dateStr > today;
                const isToday = dateStr === today;
                const isSelected = dateStr === selectedDate;
                return (
                  <button
                    key={i}
                    disabled={isFuture}
                    onClick={() => openDate(dateStr)}
                    className="aspect-square rounded-lg flex flex-col items-center justify-center gap-0.5"
                    style={{
                      background: isSelected ? BRAND_DARK : entry ? c.yellow : "transparent",
                      opacity: isFuture ? 0.3 : 1,
                      border: isToday && !isSelected ? `1.5px solid ${c.ink}` : "none",
                    }}
                  >
                    <span className="text-[11px] font-bold" style={{ color: isSelected ? c.yellow : entry ? BRAND_DARK : c.ink }}>{d}</span>
                    <div className="flex items-center gap-0.5">
                      {entry && <WeightDirBadge info={dirMap[dateStr]} c={c} size={8} compact />}
                      {entry?.image && <span style={{ fontSize: 6 }}>📷</span>}
                    </div>
                  </button>
                );
              })}
            </div>

            {selectedDate && (
              <div className="pt-3" style={{ borderTop: `1px solid ${c.line}` }}>
                <div className="text-xs font-bold mb-2" style={{ color: c.inkSoft }}>{selectedDate} 기록</div>
                <div className="flex gap-2 mb-2">
                  <input
                    type="number"
                    step="0.1"
                    value={backfillValue}
                    onChange={(e) => setBackfillValue(e.target.value)}
                    placeholder="체중 (kg)"
                    className="flex-1 px-3 py-2.5 rounded-xl text-sm font-semibold outline-none"
                    style={{ background: c.cardMuted, color: c.ink }}
                  />
                  <button
                    onClick={() => backfileRef.current?.click()}
                    className="w-11 h-11 rounded-xl flex items-center justify-center shrink-0 overflow-hidden"
                    style={{ background: c.cardMuted }}
                  >
                    {backfillPreview ? (
                      <img src={backfillPreview} className="w-full h-full object-cover" />
                    ) : (
                      <Camera size={16} style={{ color: c.inkSoft }} />
                    )}
                  </button>
                  <input ref={backfileRef} type="file" accept="image/*" onChange={handleBackfillFile} className="hidden" />
                </div>
                <button
                  onClick={saveBackfill}
                  className="w-full py-2.5 rounded-xl text-sm font-extrabold"
                  style={{ background: BRAND_DARK, color: c.yellow }}
                >
                  이 날짜로 저장
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  통계 화면                                                          */
/* ------------------------------------------------------------------ */

function StatsScreen({
  data,
  goBack,
  weightRevealed,
  setWeightRevealed,
  rangeKey,
  setRangeKey,
  periodOffset,
  setPeriodOffset,
  calendarOpen,
  setCalendarOpen,
  selectedDate,
  setSelectedDate,
  calYear,
  setCalYear,
  calMonth,
  setCalMonth,
}) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth();
  const today = todayStr(now);

  const logsByDate = useMemo(() => {
    const map = {};
    const ensure = (d) => (map[d] ||= { food: [], weight: [], exercise: [] });
    data.logs.food.forEach((f) => ensure(f.date).food.push(f));
    data.logs.weight.forEach((w) => ensure(w.date).weight.push(w));
    data.logs.exercise.forEach((e) => ensure(e.date).exercise.push(e));
    return map;
  }, [data.logs]);

  const dirMap = useMemo(() => buildWeightDirMap(data.logs.weight), [data.logs.weight]);

  // 요약 문구는 "오늘" 기준이 아니라, 지금 달력에서 보고 있는 달(calYear/calMonth) 기준으로 계산한다
  const isCurrentMonth = calYear === year && calMonth === month;
  const daysInBrowsedMonth = new Date(calYear, calMonth + 1, 0).getDate();
  const denomDays = isCurrentMonth ? now.getDate() : daysInBrowsedMonth;
  const browsedMonthLabel = new Date(calYear, calMonth, 1).toLocaleDateString("ko-KR", { month: "long" });
  const activeDates = Object.keys(logsByDate).filter((d) => {
    const dt = new Date(d + "T00:00:00");
    return dt.getFullYear() === calYear && dt.getMonth() === calMonth && (logsByDate[d].food.length || logsByDate[d].weight.length);
  });
  const successDays = activeDates.length;
  const rate = denomDays > 0 ? Math.round((successDays / denomDays) * 100) : 0;

  // ---- 몸무게 변화 그래프: 기간 선택 (기본 최근 1주, 최대 전체) + 기간 탐색 ----
  const RANGE_OPTIONS = [
    { key: "1w", label: "1주일", days: 7 },
    { key: "1m", label: "1개월", days: 30 },
    { key: "3m", label: "3개월", days: 90 },
    { key: "6m", label: "6개월", days: 182 },
    { key: "1y", label: "1년", days: 365 },
    { key: "all", label: "전체", days: null },
  ];
  const range = RANGE_OPTIONS.find((r) => r.key === rangeKey);
  const selectRange = (key) => {
    setRangeKey(key);
    setPeriodOffset(0);
  };
  const weightAsc = data.logs.weight.slice().sort((a, b) => (a.date < b.date ? -1 : 1));

  const windowEndDate = range.days ? addDays(now, -periodOffset * range.days) : now;
  const windowEndStr = todayStr(windowEndDate);
  const cutoffDate = range.days ? todayStr(addDays(windowEndDate, -range.days)) : null;
  const chartData = weightAsc
    .filter((w) => (!cutoffDate || w.date >= cutoffDate) && w.date <= windowEndStr)
    .map((w) => ({ date: w.date.slice(5), kg: w.value }));

  const canGoNext = range.days && periodOffset > 0;
  const canGoPrev = !!range.days;
  const periodLabel = range.days
    ? `${todayStr(addDays(windowEndDate, -range.days + 1)).slice(5)} ~ ${windowEndStr.slice(5)}`
    : "전체 기간";

  const chartAvg = chartData.length ? chartData.reduce((s, d) => s + d.kg, 0) / chartData.length : null;
  const weightTicks = useMemo(() => {
    if (!chartData.length) return [];
    const vals = chartData.map((d) => d.kg);
    const min = Math.floor(Math.min(...vals)) - 1;
    const max = Math.ceil(Math.max(...vals)) + 1;
    const arr = [];
    for (let v = min; v <= max; v++) arr.push(v);
    return arr;
  }, [chartData]);
  const chartNetChange = chartData.length > 1 ? Math.round((chartData[chartData.length - 1].kg - chartData[0].kg) * 10) / 10 : null;

  const dow = now.getDay(); // 0=일 ... 6=토
  const mondayOffset = dow === 0 ? 6 : dow - 1; // 오늘부터 이번 주 월요일까지 며칠 전인지
  const weekMonday = addDays(now, -mondayOffset);
  const weekDates = new Set(Array.from({ length: 7 }, (_, i) => todayStr(addDays(weekMonday, i))));
  const weekFood = data.logs.food.filter((f) => weekDates.has(f.date));
  let proteinScore = 3, veggieScore = 3, carbScore = 3;
  weekFood.forEach((f) => {
    if (f.tag === "균형 좋음") { proteinScore += 0.4; veggieScore += 0.4; carbScore += 0.2; }
    if (f.tag === "단백질 보완") proteinScore -= 0.4;
    if (f.tag === "비타민 부족") veggieScore -= 0.4;
    if (f.tag === "탄수화물 많음") carbScore -= 0.4;
    if (f.tag === "가공식품 많음") { veggieScore -= 0.2; carbScore -= 0.2; }
  });
  const clampStar = (v) => Math.max(1, Math.min(5, Math.round(v)));
  const stars = (n) => "★".repeat(n) + "☆".repeat(5 - n);
  const weeklyBalance = [
    { label: "단백질", value: clampStar(proteinScore) },
    { label: "채소", value: clampStar(veggieScore) },
    { label: "탄수화물 조절", value: clampStar(carbScore) },
  ];

  // ---- 기록 달력: 이번 달에 갇히지 않고 이전 달로 이동 가능 (상태는 App에서 관리해 화면 전환에도 유지됨) ----
  const goPrevMonth = () => {
    setSelectedDate(null);
    if (calMonth === 0) { setCalYear((y) => y - 1); setCalMonth(11); } else { setCalMonth((m) => m - 1); }
  };
  const goNextMonth = () => {
    if (isCurrentMonth) return;
    setSelectedDate(null);
    if (calMonth === 11) { setCalYear((y) => y + 1); setCalMonth(0); } else { setCalMonth((m) => m + 1); }
  };
  const daysInCalMonth = new Date(calYear, calMonth + 1, 0).getDate();
  const firstWeekday = new Date(calYear, calMonth, 1).getDay();
  const cells = [];
  for (let i = 0; i < firstWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInCalMonth; d++) cells.push(d);

  const selected = selectedDate ? logsByDate[selectedDate] || { food: [], weight: [], exercise: [] } : null;

  // ---- 엑셀 내보내기 기간 선택 ----
  const EXPORT_PRESETS = [
    { key: "all", label: "전체", days: null },
    { key: "1w", label: "1주일", days: 7 },
    { key: "1m", label: "1개월", days: 30 },
    { key: "3m", label: "3개월", days: 90 },
    { key: "6m", label: "6개월", days: 182 },
    { key: "1y", label: "1년", days: 365 },
  ];
  const [exportOpen, setExportOpen] = useState(false);
  const earliestRecordDate = useMemo(() => {
    const all = [...data.logs.food, ...data.logs.weight, ...data.logs.exercise].map((r) => r.date);
    return all.length ? all.reduce((min, d) => (d < min ? d : min)) : today;
  }, [data.logs, today]);
  const [customStart, setCustomStart] = useState(earliestRecordDate);
  const [customEnd, setCustomEnd] = useState(today);

  const getPresetRange = (preset) =>
    preset.days ? { start: todayStr(addDays(now, -preset.days)), end: today } : { start: earliestRecordDate, end: today };

  const applyPreset = (preset) => {
    const r = getPresetRange(preset);
    setCustomStart(r.start);
    setCustomEnd(r.end);
  };

  const [exportError, setExportError] = useState(false);
  const [exportFile, setExportFile] = useState(null); // { url, filename }
  const handleExport = () => {
    const range = customStart || customEnd ? { start: customStart || null, end: customEnd || null } : null;
    if (exportFile?.url) URL.revokeObjectURL(exportFile.url);
    const result = exportToExcel(data, range);
    setExportFile(result);
    setExportError(!result);
  };

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="통계" onBack={goBack} />
      <div className="px-6 pt-4">
        <button
          onClick={() => setExportOpen((v) => !v)}
          className="w-full flex items-center justify-center gap-2 py-2.5 rounded-xl text-xs font-bold mb-2"
          style={{ background: c.cardMuted, color: c.ink }}
        >
          <Download size={14} /> 엑셀로 내보내기 {exportOpen ? "▲" : "▼"}
        </button>

        {exportOpen && (
          <div className="rounded-2xl p-4 mb-4" style={{ background: c.card, border: `1px solid ${c.line}` }}>
            <div className="text-[11px] font-bold mb-2" style={{ color: c.inkSoft }}>내보낼 기간</div>
            <div className="flex gap-1.5 mb-3 flex-wrap">
              {EXPORT_PRESETS.map((p) => {
                const r = getPresetRange(p);
                const active = customStart === r.start && customEnd === r.end;
                return (
                  <button
                    key={p.key}
                    onClick={() => applyPreset(p)}
                    className="px-3 py-1.5 rounded-full text-[11px] font-bold"
                    style={{
                      background: active ? c.yellow : c.cardMuted,
                      color: active ? BRAND_DARK : c.inkSoft,
                    }}
                  >
                    {p.label}
                  </button>
                );
              })}
            </div>

            <div className="text-[11px] font-bold mb-2" style={{ color: c.inkSoft }}>또는 직접 기간 설정</div>
            <div className="flex items-center gap-2 mb-4">
              <input
                type="date"
                value={customStart}
                onChange={(e) => setCustomStart(e.target.value)}
                className="flex-1 px-2 py-2 rounded-lg text-xs font-semibold outline-none"
                style={{ background: c.cardMuted, color: c.ink }}
              />
              <span className="text-xs" style={{ color: c.inkSoft }}>~</span>
              <input
                type="date"
                value={customEnd}
                onChange={(e) => setCustomEnd(e.target.value)}
                className="flex-1 px-2 py-2 rounded-lg text-xs font-semibold outline-none"
                style={{ background: c.cardMuted, color: c.ink }}
              />
            </div>

            <button
              onClick={handleExport}
              className="w-full flex items-center justify-center gap-2 py-3 rounded-xl text-sm font-extrabold"
              style={{ background: BRAND_DARK, color: c.yellow }}
            >
              <Download size={15} /> 엑셀 파일 만들기
            </button>
            {exportError && (
              <div className="text-[11px] font-semibold text-center mt-2" style={{ color: c.danger }}>
                파일 생성에 실패했어요. 다시 시도해주세요.
              </div>
            )}
            {exportFile && (
              <div className="mt-3 pt-3" style={{ borderTop: `1px solid ${c.line}` }}>
                <div className="text-[11px] font-semibold text-center mb-2" style={{ color: c.inkSoft }}>
                  파일이 준비됐어요. 아래 버튼을 직접 눌러주세요
                </div>
                <div className="flex gap-2">
                  <a
                    href={exportFile.url}
                    download={exportFile.filename}
                    className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl text-xs font-extrabold"
                    style={{ background: c.yellow, color: BRAND_DARK }}
                  >
                    <Download size={13} /> 다운로드
                  </a>
                  <a
                    href={exportFile.url}
                    target="_blank"
                    rel="noreferrer"
                    className="flex-1 flex items-center justify-center gap-1.5 py-2.5 rounded-xl text-xs font-extrabold"
                    style={{ background: c.cardMuted, color: c.ink }}
                  >
                    새 탭에서 열기
                  </a>
                </div>
                <div className="text-[10px] text-center mt-2" style={{ color: c.inkSoft }}>
                  하나가 안 되면 다른 버튼을 눌러보세요
                </div>
              </div>
            )}
          </div>
        )}

        <button
          onClick={() => setCalendarOpen((v) => !v)}
          className="w-full rounded-2xl p-5 mb-3 text-center"
          style={{ background: c.yellowSoft }}
        >
          <div className="text-xs font-bold mb-1" style={{ color: c.ink }}>
            {isCurrentMonth
              ? `${browsedMonthLabel} 들어 ${denomDays}일 동안, 기록한 날은 ${successDays}일이에요`
              : `${browsedMonthLabel} 한 달(${denomDays}일) 중, 기록한 날은 ${successDays}일이에요`}
          </div>
          <div className="text-3xl font-extrabold mb-1">{rate}%</div>
          <div className="text-[11px] font-semibold leading-relaxed" style={{ color: c.inkSoft }}>
            연속 기록이 아니라 기록한 날의 비율이에요
            <br />
            탭해서 달력 보기 {calendarOpen ? "▲" : "▼"}
          </div>
        </button>

        {calendarOpen && (
          <div className="rounded-2xl p-4 mb-5" style={{ background: c.card, border: `1px solid ${c.line}` }}>
            <div className="flex items-center justify-between mb-2">
              <button onClick={goPrevMonth} className="w-7 h-7 rounded-full flex items-center justify-center" style={{ background: c.cardMuted }}>
                <ChevronLeft size={14} />
              </button>
              <span className="text-xs font-bold">{calYear}년 {calMonth + 1}월</span>
              <button
                onClick={goNextMonth}
                disabled={isCurrentMonth}
                className="w-7 h-7 rounded-full flex items-center justify-center"
                style={{ background: c.cardMuted, opacity: isCurrentMonth ? 0.3 : 1 }}
              >
                <ChevronLeft size={14} style={{ transform: "rotate(180deg)" }} />
              </button>
            </div>
            <div className="grid grid-cols-7 gap-1 mb-1">
              {["일", "월", "화", "수", "목", "금", "토"].map((d) => (
                <div key={d} className="text-center text-[10px] font-bold py-1" style={{ color: c.inkSoft }}>
                  {d}
                </div>
              ))}
            </div>
            {(() => {
              const weekRows = [];
              for (let i = 0; i < cells.length; i += 7) weekRows.push(cells.slice(i, i + 7));

              const renderDayCell = (d, key) => {
                if (!d) return <div key={key} />;
                const dateStr = `${calYear}-${pad(calMonth + 1)}-${pad(d)}`;
                const info = logsByDate[dateStr];
                const success = info && (info.food.length > 0 || info.weight.length > 0);
                const isFuture = dateStr > today;
                const isToday = dateStr === today;
                const isSelected = dateStr === selectedDate;
                const thumb = info?.food?.find((f) => f.image)?.image;
                return (
                  <button
                    key={key}
                    disabled={isFuture}
                    onClick={() => setSelectedDate(isSelected ? null : dateStr)}
                    className="aspect-square rounded-lg flex flex-col items-center justify-center gap-0.5 relative overflow-hidden"
                    style={{
                      background: isSelected
                        ? BRAND_DARK
                        : thumb
                        ? undefined
                        : success
                        ? c.yellow
                        : "transparent",
                      backgroundImage: !isSelected && thumb ? `linear-gradient(rgba(0,0,0,0.3), rgba(0,0,0,0.3)), url(${thumb})` : undefined,
                      backgroundSize: "cover",
                      backgroundPosition: "center",
                      opacity: isFuture ? 0.3 : 1,
                      border: isToday && !isSelected ? `1.5px solid ${c.ink}` : "none",
                    }}
                  >
                    <span
                      className="text-[11px] font-bold relative z-10"
                      style={{ color: isSelected || thumb ? "#fff" : success ? BRAND_DARK : c.ink }}
                    >
                      {d}
                    </span>
                    {info && (info.food.length > 0 || info.weight.length > 0 || info.exercise.length > 0) && (
                      <div className="flex items-center gap-0.5 relative z-10">
                        {info.food.length > 0 && !thumb && <span style={{ fontSize: 6 }}>🍽️</span>}
                        {info.weight.length > 0 && <WeightDirBadge info={dirMap[dateStr]} c={c} size={8} compact />}
                        {info.exercise.length > 0 && <span style={{ fontSize: 6 }}>🏃</span>}
                      </div>
                    )}
                  </button>
                );
              };

              return weekRows.map((row, ri) => {
                const rowHasSelected = row.some((d) => d && `${calYear}-${pad(calMonth + 1)}-${pad(d)}` === selectedDate);
                return (
                  <div key={ri}>
                    <div className="grid grid-cols-7 gap-1 mb-1">
                      {row.map((d, ci) => renderDayCell(d, `${ri}-${ci}`))}
                    </div>
                    {rowHasSelected && selected && (
                      <div className="rounded-xl p-3 mb-2 fade-in" style={{ background: c.cardMuted }}>
                        <div className="text-xs font-bold mb-2" style={{ color: c.inkSoft }}>{selectedDate} 기록</div>
                        {selected.food.length === 0 && selected.weight.length === 0 && selected.exercise.length === 0 && (
                          <div className="text-xs font-semibold" style={{ color: c.inkSoft }}>기록이 없어요</div>
                        )}
                        {selected.food.map((f) => (
                          <div key={f.id} className="flex items-center gap-2 mb-2 text-xs font-semibold">
                            {f.image ? (
                              <img src={f.image} className="w-8 h-8 rounded-lg object-cover" />
                            ) : (
                              <span>🍽️</span>
                            )}
                            <span>{f.time} · {f.tag}</span>
                          </div>
                        ))}
                        {selected.weight.map((w) => (
                          <div key={w.id} className="flex items-center gap-1.5 text-xs font-semibold mb-2">
                            <span>⚖️ {!data.weightHidden || weightRevealed ? `${w.value}kg` : "••.•kg (가림)"}</span>
                            <WeightDirBadge info={dirMap[w.date]} c={c} size={11} />
                          </div>
                        ))}
                        {selected.exercise.map((e) => (
                          <div key={e.id} className="text-xs font-semibold mb-2">
                            🏃 {(e.types || [e.type]).join(", ")} · {e.minutes}분
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              });
            })()}
          </div>
        )}

        <div className="flex items-center justify-between mb-2">
          <span className="text-xs font-bold" style={{ color: c.inkSoft }}>몸무게 변화</span>
          {data.weightHidden && weightRevealed && (
            <button
              onClick={() => setWeightRevealed(false)}
              className="flex items-center gap-1 text-[11px] font-bold px-2 py-1 rounded-full"
              style={{ background: c.cardMuted, color: c.inkSoft }}
            >
              <EyeOff size={12} /> 가리기
            </button>
          )}
        </div>
        <div className="flex gap-1.5 mb-3 flex-wrap">
          {RANGE_OPTIONS.map((r) => (
            <button
              key={r.key}
              onClick={() => selectRange(r.key)}
              className="px-3 py-1.5 rounded-full text-[11px] font-bold"
              style={{ background: rangeKey === r.key ? c.yellow : c.cardMuted, color: rangeKey === r.key ? BRAND_DARK : c.inkSoft }}
            >
              {r.label}
            </button>
          ))}
        </div>

        <div className="rounded-2xl p-4 mb-6" style={{ background: c.card, border: `1px solid ${c.line}` }}>
          {data.weightHidden && !weightRevealed ? (
            <button
              onClick={() => setWeightRevealed(true)}
              className="w-full h-36 flex flex-col items-center justify-center gap-1 text-sm font-semibold text-center px-4"
              style={{ color: c.inkSoft }}
            >
              <Eye size={18} />
              체중이 가려져 있어요
              <span className="text-[11px]">눌러서 그래프 보기</span>
            </button>
          ) : (
            <>
              {/* 기간 탐색 */}
              <div className="flex items-center justify-between mb-1">
                <button
                  onClick={() => canGoPrev && setPeriodOffset((p) => p + 1)}
                  disabled={!canGoPrev}
                  className="w-6 h-6 rounded-full flex items-center justify-center"
                  style={{ background: c.cardMuted, opacity: canGoPrev ? 1 : 0.3 }}
                >
                  <ChevronLeft size={13} />
                </button>
                <span className="text-[11px] font-bold" style={{ color: c.inkSoft }}>{periodLabel}</span>
                <button
                  onClick={() => canGoNext && setPeriodOffset((p) => p - 1)}
                  disabled={!canGoNext}
                  className="w-6 h-6 rounded-full flex items-center justify-center"
                  style={{ background: c.cardMuted, opacity: canGoNext ? 1 : 0.3 }}
                >
                  <ChevronLeft size={13} style={{ transform: "rotate(180deg)" }} />
                </button>
              </div>

              {/* 기간 요약 한 줄 */}
              {chartAvg !== null && (
                <div className="text-xs font-semibold text-center mb-2" style={{ color: c.inkSoft }}>
                  이 기간 평균 {chartAvg.toFixed(1)}kg
                  {chartNetChange !== null && ` · ${chartNetChange > 0 ? "+" : ""}${chartNetChange}kg 변화`}
                </div>
              )}

              {chartData.length > 1 ? (
                <ResponsiveContainer width="100%" height={160}>
                  <ComposedChart data={chartData} margin={{ top: 22, right: 8, left: 8, bottom: 0 }}>
                    <CartesianGrid stroke={c.line} strokeDasharray="3 3" vertical={false} />
                    <XAxis dataKey="date" tick={{ fontSize: 10, fill: c.inkSoft }} axisLine={false} tickLine={false} interval="preserveStartEnd" minTickGap={20} />
                    <YAxis
                      tick={{ fontSize: 10, fill: c.inkSoft }}
                      axisLine={false}
                      tickLine={false}
                      width={30}
                      allowDecimals={false}
                      ticks={weightTicks.length ? weightTicks : undefined}
                      domain={weightTicks.length ? [weightTicks[0], weightTicks[weightTicks.length - 1]] : ["dataMin - 1", "dataMax + 1"]}
                    />
                    <Tooltip
                      content={({ active, payload, label }) => {
                        if (!active || !payload || !payload.length) return null;
                        return (
                          <div
                            style={{
                              background: c.card,
                              border: `1px solid ${c.line}`,
                              borderRadius: 8,
                              padding: "6px 10px",
                              fontSize: 11,
                              fontWeight: 700,
                              color: c.ink,
                            }}
                          >
                            <div style={{ color: c.inkSoft, marginBottom: 2 }}>{label}</div>
                            <div>{payload[0].value}kg</div>
                          </div>
                        );
                      }}
                    />
                    <Bar dataKey="kg" fill={c.yellowSoft} radius={[4, 4, 0, 0]} barSize={16} />
                    <Line
                      type="monotone"
                      dataKey="kg"
                      stroke={c.green}
                      strokeWidth={2.5}
                      dot={{ r: 3 }}
                      label={(props) => {
                        const isLast = props.index === chartData.length - 1;
                        if (!isLast) return null;
                        return (
                          <text x={props.x} y={props.y - 12} textAnchor="middle" fontSize={11} fontWeight={800} fill={c.ink}>
                            {props.value}kg
                          </text>
                        );
                      }}
                    />
                  </ComposedChart>
                </ResponsiveContainer>
              ) : (
                <div className="h-36 flex items-center justify-center text-sm font-semibold" style={{ color: c.inkSoft }}>
                  이 기간엔 기록이 부족해요
                </div>
              )}
            </>
          )}
        </div>

        <div className="text-xs font-bold mb-2" style={{ color: c.inkSoft }}>
          이번 주 영양 균형 ({todayStr(weekMonday).slice(5).replace("-", "/")} ~ {todayStr(addDays(weekMonday, 6)).slice(5).replace("-", "/")})
        </div>
        <div className="rounded-2xl p-4 mb-6" style={{ background: c.card, border: `1px solid ${c.line}` }}>
          {weekFood.length === 0 ? (
            <div className="text-sm font-semibold text-center py-4" style={{ color: c.inkSoft }}>
              이번 주 음식 기록이 쌓이면 보여드릴게요
            </div>
          ) : (
            <div className="flex flex-col gap-2.5">
              {weeklyBalance.map((b) => (
                <div key={b.label} className="flex items-center justify-between">
                  <span className="text-sm font-semibold">{b.label}</span>
                  <span className="text-sm tracking-wide" style={{ color: c.yellow.replace("#FEE500", "#C9A400") }}>
                    {stars(b.value)}
                  </span>
                </div>
              ))}
              <div className="text-[11px] mt-1" style={{ color: c.inkSoft }}>
                기록 {weekFood.length}건 기준 · 정답은 없어요, 그냥 흐름만 살펴보세요
              </div>
            </div>
          )}
        </div>

        <div className="grid grid-cols-2 gap-3">
          <StatCard c={c} label="음식 기록 횟수" value={`${data.logs.food.length}회`} />

          <StatCard c={c} label="체중 기록 횟수" value={`${data.logs.weight.length}회`} />
          <StatCard c={c} label="운동 기록 횟수" value={`${data.logs.exercise.length}회`} />
          <StatCard c={c} label="누적 포인트" value={`${data.points}P`} />
        </div>
      </div>
    </div>
  );
}

function StatCard({ c, label, value }) {
  return (
    <div className="rounded-2xl p-4" style={{ background: c.card, border: `1px solid ${c.line}` }}>
      <div className="text-xs font-bold mb-1" style={{ color: c.inkSoft }}>{label}</div>
      <div className="text-lg font-extrabold">{value}</div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  설정 화면                                                          */
/* ------------------------------------------------------------------ */

function SettingsScreen({ data, setData, goBack }) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;

  const Row = ({ label, children }) => (
    <div className="flex items-center justify-between py-3.5" style={{ borderBottom: `1px solid ${c.line}` }}>
      <span className="text-sm font-semibold">{label}</span>
      {children}
    </div>
  );

  const Toggle = ({ value, onChange }) => (
    <button
      onClick={() => onChange(!value)}
      className="w-11 h-6 rounded-full relative transition"
      style={{ background: value ? c.yellow : c.line }}
    >
      <div
        className="w-5 h-5 rounded-full bg-white absolute top-0.5 transition-all"
        style={{ left: value ? "22px" : "2px" }}
      />
    </button>
  );

  return (
    <div className="pb-10 fade-in">
      <TopBar c={c} title="설정" onBack={goBack} />
      <div className="px-6 pt-2">
        <div className="text-xs font-bold mb-1 mt-4" style={{ color: c.inkSoft }}>식사 시간</div>
        <Row label="식사 시작 시간">
          <input
            type="time"
            value={data.mealStart}
            onChange={(e) => setData((p) => ({ ...p, mealStart: e.target.value }))}
            className="px-3 py-1.5 rounded-lg text-sm font-semibold outline-none"
            style={{ background: c.cardMuted, color: c.ink }}
          />
        </Row>

        <div className="text-xs font-bold mb-1 mt-6" style={{ color: c.inkSoft }}>개인정보</div>
        <Row label="체중 가림 처리">
          <Toggle value={data.weightHidden} onChange={(v) => setData((p) => ({ ...p, weightHidden: v }))} />
        </Row>
        <Row label="변화 기록 촬영 가이드 성별">
          <div className="flex gap-2">
            {[
              { key: "female", label: "여성" },
              { key: "male", label: "남성" },
            ].map((g) => (
              <button
                key={g.key}
                onClick={() => setData((p) => ({ ...p, gender: g.key }))}
                className="px-3 py-1.5 rounded-lg text-xs font-bold"
                style={{ background: data.gender === g.key ? c.yellow : c.cardMuted, color: data.gender === g.key ? BRAND_DARK : c.ink }}
              >
                {g.label}
              </button>
            ))}
          </div>
        </Row>

        <div className="text-xs font-bold mb-1 mt-6" style={{ color: c.inkSoft }}>화면 · 접근성</div>
        <Row label="다크 모드">
          <Toggle value={data.darkMode} onChange={(v) => setData((p) => ({ ...p, darkMode: v }))} />
        </Row>
        <Row label="글자 크기">
          <div className="flex gap-1.5">
            {[
              { v: 0.9, label: "작게" },
              { v: 1, label: "보통" },
              { v: 1.2, label: "크게" },
            ].map((f) => (
              <button
                key={f.v}
                onClick={() => setData((p) => ({ ...p, fontScale: f.v }))}
                className="px-2.5 py-1.5 rounded-lg text-xs font-bold"
                style={{ background: data.fontScale === f.v ? c.yellow : c.cardMuted, color: data.fontScale === f.v ? BRAND_DARK : c.ink }}
              >
                {f.label}
              </button>
            ))}
          </div>
        </Row>
        <Row label="언어">
          <span className="text-xs font-bold px-3 py-1.5 rounded-lg" style={{ background: c.cardMuted, color: c.inkSoft }}>
            한국어 (추후 확장)
          </span>
        </Row>

        <div className="text-xs font-bold mb-1 mt-6" style={{ color: c.inkSoft }}>계정</div>
        <Row label="로그인">
          <span className="text-xs font-bold px-3 py-1.5 rounded-lg" style={{ background: c.cardMuted, color: c.inkSoft, opacity: 0.6 }}>
            준비 중
          </span>
        </Row>

        <div
          className="mt-8 rounded-2xl p-4 text-xs leading-relaxed"
          style={{ background: c.cardMuted, color: c.inkSoft }}
        >
          현재 모든 기록은 이 기기에만 저장돼요. 나중에 로그인 기능이 추가되면 서버에 안전하게 백업하고 여러 기기에서 확인할 수 있게 될 예정이에요.
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  공통 컴포넌트                                                      */
/* ------------------------------------------------------------------ */

function TopBar({ c, title, onBack }) {
  return (
    <div className="flex items-center gap-2 px-4 pt-6 pb-2 sticky top-0 z-10" style={{ background: c.bg }}>
      <button onClick={onBack} className="w-9 h-9 rounded-full flex items-center justify-center" style={{ background: c.cardMuted }}>
        <ChevronLeft size={18} />
      </button>
      <span className="font-extrabold text-base">{title}</span>
    </div>
  );
}

function BottomNav({ screen, setScreen, data }) {
  const c = data.darkMode ? COLORS.dark : COLORS.light;
  const recordScreens = ["recordHub", "food", "exercise", "weight"];
  const sideItems = [
    { key: "recordHub", icon: Camera, label: "기록" },
    { key: "stats", icon: BarChart3, label: "통계" },
    { key: "bodyPhoto", icon: Aperture, label: "변화" },
    { key: "settings", icon: SettingsIcon, label: "설정" },
  ];

  return (
    <div
      className="relative z-10 shrink-0"
      style={{ background: c.card, borderTop: `1px solid ${c.line}` }}
    >
      {/* 박스 높이는 기록·통계·변화·설정 4개 기준으로만 결정된다 (홈은 여기 포함 안 됨) */}
      <div className="grid grid-cols-5 items-center px-1 pb-3 pt-3">
        <NavItem it={sideItems[0]} active={recordScreens.includes(screen)} onClick={() => setScreen(sideItems[0].key)} c={c} />
        <NavItem it={sideItems[1]} active={screen === sideItems[1].key} onClick={() => setScreen(sideItems[1].key)} c={c} />
        <div />
        <NavItem it={sideItems[2]} active={screen === sideItems[2].key} onClick={() => setScreen(sideItems[2].key)} c={c} />
        <NavItem it={sideItems[3]} active={screen === sideItems[3].key} onClick={() => setScreen(sideItems[3].key)} c={c} />
      </div>

      {/* 홈: 박스 위에 겹쳐서 떠 있는 형태 (absolute라 박스 높이 계산에 영향 없음) */}
      <button
        onClick={() => setScreen("home")}
        className="absolute left-1/2"
        style={{ top: 0, transform: "translate(-50%, -20%)", padding: 0 }}
      >
        <div style={{ width: 56 }}>
          <div
            className="rounded-full flex items-center justify-center shadow-lg"
            style={{ width: 56, height: 56, background: c.yellow, border: `4px solid ${c.card}` }}
          >
            <HomeIcon size={23} color="#3A2317" strokeWidth={2.5} />
          </div>
          <div
            className="font-extrabold"
            style={{ fontSize: 10, lineHeight: "10px", textAlign: "center", marginTop: 6, color: screen === "home" ? c.ink : c.inkSoft }}
          >
            홈
          </div>
        </div>
      </button>
    </div>
  );
}

function NavItem({ it, active, onClick, c }) {
  const Icon = it.icon;
  return (
    <button onClick={onClick} className="flex flex-col items-center justify-self-center gap-1">
      <Icon size={19} color={active ? c.ink : c.inkSoft} strokeWidth={active ? 2.5 : 2} />
      <span className="text-[10px] font-bold" style={{ color: active ? c.ink : c.inkSoft }}>
        {it.label}
      </span>
    </button>
  );
}
