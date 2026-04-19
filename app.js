pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const BASE_HEADERS = ["序号", "出境/入境", "出入境日期", "证件名称", "证件号码", "出入境口岸", "航班号"];

const state = {
  originalData: null,
  workingData: null,
  validated: false,
  baseTotalRows: 0,
  adjustmentSummary: "",
  highlightedDays: 0,
  shanghaiLevel: "master",
  adjustMode: "",
  earliestEntryDate: "",
  calcRafId: 0,
  calendarDayInfoMap: new Map(),
  activeCalendarDay: "",
};

const el = {
  beijingTime: document.getElementById("beijingTime"),
  systemTime: document.getElementById("systemTime"),
  pdfInput: document.getElementById("pdfInput"),
  dropzone: document.getElementById("dropzone"),
  stats: document.getElementById("stats"),
  status: document.getElementById("status"),
  adjustBox: document.getElementById("adjustBox"),
  adjustTitle: document.getElementById("adjustTitle"),
  adjustPrimaryItem: document.getElementById("adjustPrimaryItem"),
  adjustPrimaryText: document.getElementById("adjustPrimaryText"),
  adjustSecondaryItem: document.getElementById("adjustSecondaryItem"),
  adjustSecondaryText: document.getElementById("adjustSecondaryText"),
  adjustManualItem: document.getElementById("adjustManualItem"),
  manualExitDate: document.getElementById("manualExitDate"),
  btnAddManualExit: document.getElementById("btnAddManualExit"),
  manualExitWarn: document.getElementById("manualExitWarn"),
  btnAddVirtual: document.getElementById("btnAddVirtual"),
  btnUseLastEntry: document.getElementById("btnUseLastEntry"),
  btnReupload: document.getElementById("btnReupload"),
  btnExport: document.getElementById("btnExport"),
  btnRandomDemo: document.getElementById("btnRandomDemo"),
  calcSection: document.getElementById("calcSection"),
  startDate: document.getElementById("startDate"),
  endDate: document.getElementById("endDate"),
  credentialList: document.getElementById("credentialList"),
  portList: document.getElementById("portList"),
  calcOrderWarn: document.getElementById("calcOrderWarn"),
  calcOrderWarnText: document.getElementById("calcOrderWarnText"),
  resultBox: document.getElementById("resultBox"),
  beijingPanel: document.getElementById("beijingPanel"),
  beijingSummary: document.getElementById("beijingSummary"),
  beijingProgress: document.getElementById("beijingProgress"),
  shanghaiPanel: document.getElementById("shanghaiPanel"),
  shanghaiDegreeSwitch: document.getElementById("shanghaiDegreeSwitch"),
  shanghaiSummary: document.getElementById("shanghaiSummary"),
  shanghaiProgress: document.getElementById("shanghaiProgress"),
  carPanel: document.getElementById("carPanel"),
  carSummary: document.getElementById("carSummary"),
  carProgress: document.getElementById("carProgress"),
  totalDays: document.getElementById("totalDays"),
  calendar: document.getElementById("calendar"),
};

const SHANGHAI_REQUIREMENTS = {
  bachelor: 720,
  master: 180,
  doctor: 360,
};

const DEMO_CERT_TYPES = ["普通护照", "往来港澳通行证", "往来台湾通行证"];
const DEMO_PORTS = [
  "上海浦东国际机场",
  "北京大兴国际机场",
  "深圳湾口岸",
  "广州白云国际机场",
  "杭州萧山国际机场",
  "成都天府国际机场",
  "港珠澳大桥口岸",
  "拱北口岸",
  "福田口岸",
  "深圳口岸",
  "横琴口岸",
  "呜咦唔啊啊蹭蹭哇啦哇啦笨笨口岸"
];
const DEMO_AIRLINE_PREFIX = ["MU", "CA", "CZ", "FM", "HU", "9C", "MF", "ZH", "HO", "SC"];
const CALENDAR_WEEKDAY_ROW =
  "<div class=\"week\"><div class=\"wk\">一</div><div class=\"wk\">二</div><div class=\"wk\">三</div><div class=\"wk\">四</div><div class=\"wk\">五</div><div class=\"wk\">六</div><div class=\"wk\">日</div></div>";
let calendarTooltipEl = null;

function pad2(v) {
  return String(v).padStart(2, "0");
}

function formatDateKey(y, m, d) {
  return `${y}-${pad2(m)}-${pad2(d)}`;
}

function parseDateKey(dateText) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeText(dateText));
  if (!m) return null;
  const y = Number(m[1]);
  const mm = Number(m[2]);
  const dd = Number(m[3]);
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;
  return { y, m: mm, d: dd };
}

function utcDateFromDateKey(dateText) {
  const parts = parseDateKey(dateText);
  if (!parts) return null;
  const dt = new Date(Date.UTC(parts.y, parts.m - 1, parts.d));
  if (
    dt.getUTCFullYear() !== parts.y ||
    dt.getUTCMonth() + 1 !== parts.m ||
    dt.getUTCDate() !== parts.d
  ) {
    return null;
  }
  return dt;
}

function isValidDateKey(dateText) {
  return !!utcDateFromDateKey(dateText);
}

function formatDateKeyFromUTCDate(date) {
  return formatDateKey(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
}

function addDaysToDateKey(dateText, deltaDays) {
  const dt = utcDateFromDateKey(dateText);
  if (!dt) return "";
  dt.setUTCDate(dt.getUTCDate() + deltaDays);
  return formatDateKeyFromUTCDate(dt);
}

function daysInMonthByDateKey(year, month) {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function weekdayMon0ByDateKey(year, month, day) {
  const weekDaySun0 = new Date(Date.UTC(year, month - 1, day)).getUTCDay();
  return (weekDaySun0 + 6) % 7;
}

function getBeijingTodayKey() {
  const formatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const map = {};
  for (const part of formatter.formatToParts(new Date())) {
    if (part.type !== "literal") map[part.type] = part.value;
  }
  return `${map.year}-${map.month}-${map.day}`;
}

function formatLocalDateTime(date) {
  const y = date.getFullYear();
  const m = pad2(date.getMonth() + 1);
  const d = pad2(date.getDate());
  const h = pad2(date.getHours());
  const mm = pad2(date.getMinutes());
  return `${y}-${m}-${d} ${h}:${mm}`;
}

function formatDateTimeByTimezone(date, timeZone) {
  const formatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  const map = {};
  for (const part of formatter.formatToParts(date)) {
    if (part.type !== "literal") map[part.type] = part.value;
  }
  return `${map.year}-${map.month}-${map.day} ${map.hour}:${map.minute}`;
}

function updateTimeDisplays() {
  const now = new Date();
  if (el.beijingTime) {
    el.beijingTime.textContent = formatDateTimeByTimezone(now, "Asia/Shanghai");
  }
  if (el.systemTime) {
    el.systemTime.textContent = formatLocalDateTime(now);
  }
}

function normalizeText(v) {
  return (v ?? "").toString().replace(/\s+/g, " ").trim();
}

function toSerial(row) {
  const s = normalizeText(row["序号"]);
  return /^\d+$/.test(s) ? parseInt(s, 10) : Number.MAX_SAFE_INTEGER;
}

function setStatus(text, type = "") {
  el.status.textContent = text;
  el.status.className = type ? `status ${type}` : "status";
}

function escapeHtml(text) {
  return normalizeText(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatDateZh(dateText) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalizeText(dateText));
  if (!m) return normalizeText(dateText);
  return `${Number(m[1])}年${Number(m[2])}月${Number(m[3])}日`;
}

function formatDayRecordBrief(row) {
  if (!row) return "PDF无记录";
  const dateText = formatDateZh(normalizeText(row["出入境日期"]));
  const port = normalizePortName(row["出入境口岸"]) || normalizeText(row["出入境口岸"]) || "口岸未知";
  const certName = normalizeText(row["证件名称"]) || "证件";
  const certNum = normalizeText(row["证件号码"]);
  const certText = certNum ? `${certName}${certNum}` : certName;
  return `${dateText} ${port} ${certText}`;
}

function setCalcOrderWarn(text = "") {
  if (!el.calcOrderWarn || !el.calcOrderWarnText) return;
  const msg = normalizeText(text);
  if (!msg) {
    el.calcOrderWarn.classList.add("hidden");
    el.calcOrderWarnText.textContent = "";
    return;
  }
  el.calcOrderWarnText.textContent = msg;
  el.calcOrderWarn.classList.remove("hidden");
}

function setManualExitWarn(text = "") {
  if (!el.manualExitWarn) return;
  const msg = normalizeText(text);
  if (!msg) {
    el.manualExitWarn.textContent = "";
    el.manualExitWarn.classList.add("hidden");
    return;
  }
  el.manualExitWarn.textContent = msg;
  el.manualExitWarn.classList.remove("hidden");
}

function renderQuota(panelEl, summaryEl, progressEl, total, required) {
  if (!panelEl || !summaryEl || !progressEl) return;
  const ratio = required > 0 ? Math.min(total / required, 1) : 1;
  const achieved = total >= required;
  panelEl.classList.toggle("is-achieved", achieved);
  panelEl.classList.toggle("is-pending", !achieved);

  if (achieved) {
    summaryEl.textContent = `${total}天/${required}天，恭喜已达标！`;
  } else {
    summaryEl.textContent = `${total}天/${required}天，还需${required - total}天`;
  }

  progressEl.style.width = `${Math.max(0, Math.min(100, ratio * 100))}%`;
}

function updateQualificationPanels(totalDays) {
  const total = Math.max(0, Number(totalDays) || 0);
  const shRequired = SHANGHAI_REQUIREMENTS[state.shanghaiLevel] || SHANGHAI_REQUIREMENTS.bachelor;

  renderQuota(el.beijingPanel, el.beijingSummary, el.beijingProgress, total, 360);
  renderQuota(el.shanghaiPanel, el.shanghaiSummary, el.shanghaiProgress, total, shRequired);
  renderQuota(el.carPanel, el.carSummary, el.carProgress, total, 270);

  if (el.shanghaiDegreeSwitch) {
    Array.from(el.shanghaiDegreeSwitch.querySelectorAll("button[data-level]")).forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.level === state.shanghaiLevel);
    });
  }
}

function initQualificationControls() {
  if (!el.shanghaiDegreeSwitch) return;
  el.shanghaiDegreeSwitch.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-level]");
    if (!btn) return;
    const level = btn.dataset.level;
    if (!SHANGHAI_REQUIREMENTS[level]) return;
    state.shanghaiLevel = level;
    updateQualificationPanels(state.highlightedDays);
  });
}

function randomInt(min, max) {
  const low = Math.ceil(Number(min));
  const high = Math.floor(Number(max));
  if (!Number.isFinite(low) || !Number.isFinite(high) || high < low) return low;
  return Math.floor(Math.random() * (high - low + 1)) + low;
}

function randomPick(arr) {
  if (!Array.isArray(arr) || !arr.length) return "";
  return arr[randomInt(0, arr.length - 1)];
}

function createDemoCertNumber() {
  const prefix = randomPick(["E", "P", "H", "K"]);
  const body = String(randomInt(1000000, 9999999));
  return `${prefix}${body}`;
}

function createDemoFlightCode() {
  const prefix = randomPick(DEMO_AIRLINE_PREFIX);
  const serial = String(randomInt(100, 9999));
  return `${prefix}${serial}`;
}

function buildRandomDemoData() {
  const count = randomInt(1, 50) * 2;
  const maxSpanDays = randomInt(Math.max(1, count - 1), 365);
  const candidates = [];
  for (let i = 0; i <= maxSpanDays; i++) {
    candidates.push(i);
  }
  for (let i = candidates.length - 1; i > 0; i--) {
    const j = randomInt(0, i);
    const tmp = candidates[i];
    candidates[i] = candidates[j];
    candidates[j] = tmp;
  }
  const sortedOffsets = candidates
    .slice(0, count)
    .sort((a, b) => b - a);
  const certType = randomPick(DEMO_CERT_TYPES);
  const certNumber = createDemoCertNumber();
  const now = getBeijingTodayKey();
  const records = sortedOffsets.map((offset, index) => ({
    "序号": String(count - index),
    "出境/入境": index % 2 === 0 ? "出境" : "入境",
    "出入境日期": addDaysToDateKey(now, -offset),
    "证件名称": certType,
    "证件号码": certNumber,
    "出入境口岸": randomPick(DEMO_PORTS),
    "航班号": createDemoFlightCode(),
  }));

  return {
    "来源文件": "随机演示数据",
    "生成时间": formatDateTimeByTimezone(new Date(), "Asia/Shanghai"),
    "提取页数": 0,
    "表头": BASE_HEADERS,
    "总行数": records.length,
    "数据": records,
  };
}

function prepareDemoWorkingData(rawDemoData) {
  const copy = JSON.parse(JSON.stringify(rawDemoData));
  copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
  copy["总行数"] = copy["数据"].length;
  const base = copy["总行数"];

  const latest = getLatestRecord(copy["数据"]);
  if (latest && normalizeText(latest["出境/入境"]) === "出境") {
    copy["数据"].push(buildVirtualEntry(latest));
    copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
    copy["总行数"] = copy["数据"].length;
    return {
      dataset: copy,
      summary: `${base}+1（随机演示自动补齐今日入境记录）`,
    };
  }

  return {
    dataset: copy,
    summary: String(base),
  };
}

function normalizePortName(raw) {
  let text = normalizeText(raw)
    .replace(/\s+/g, "")
    .split(/本电子文件|制作日期|国家移民管理局|https?:\/\//)[0]
    .replace(/[。；].*$/, "")
    .trim();

  if (!text || text === "岸") return "";
  const matches = text.match(/[\u4e00-\u9fa5A-Za-z0-9]{2,}(?:口岸|机场)/g);
  if (!matches || !matches.length) return "";
  let port = matches.find((m) => m.length >= 3 && m.length <= 24) || matches[0];
  if (/^岸.+(?:口岸|机场)$/.test(port)) port = port.slice(1);
  port = port.replace(/口岸岸$/g, "口岸").trim();
  if (!port || port.length < 3 || port.length > 24) return "";
  return port;
}

function isLikelyCertNumber(token) {
  const t = normalizeText(token).toUpperCase();
  if (t.length < 5 || t.length > 20) return false;
  if (!/\d/.test(t)) return false;
  // 避免把典型航班号误判为证件号。
  if (/^[A-Z]{2,3}\d{2,5}[A-Z]?$/.test(t)) return false;
  return /^[A-Z0-9]+$/.test(t);
}

function extractCertNumber(raw) {
  const text = normalizeText(raw).toUpperCase();
  const tokens = text.match(/[A-Z0-9]{5,20}/g) || [];
  for (const token of tokens) {
    if (isLikelyCertNumber(token)) return token;
  }
  return "";
}

function normalizeFlightCode(raw) {
  const text = normalizeText(raw).toUpperCase();
  if (!text) return "";
  const token = text.replace(/\s+/g, "");
  if (!/^[A-Z0-9]{4,9}$/.test(token)) return "";
  if (!/[A-Z]/.test(token) || !/\d/.test(token)) return "";
  return token;
}

function isNoiseLine(line) {
  const t = normalizeText(line);
  if (!t) return true;
  return /本电子文件|制作日期|国家移民管理局|电子签名|验签|https?:\/\//.test(t);
}

function isRecordStartLine(line) {
  return /^\d+\s+(出境|入境)\s+\d{4}-\d{2}-\d{2}\b/.test(normalizeText(line));
}

function isPortPrefixLine(line) {
  const t = normalizeText(line);
  return t.length >= 3 && t.length <= 12 && /口$/.test(t) && !isRecordStartLine(t);
}

function isPortSuffixLine(line) {
  return normalizeText(line) === "岸";
}

function stitchWrappedPortLines(lines) {
  const result = [];

  for (let i = 0; i < lines.length; i++) {
    const current = normalizeText(lines[i]);
    if (!current) continue;

    // 处理：港珠澳大桥口 / [记录行] / 岸
    if (isPortPrefixLine(current) && i + 1 < lines.length && isRecordStartLine(lines[i + 1])) {
      let merged = normalizeText(lines[i + 1]);
      if (i + 2 < lines.length && isPortSuffixLine(lines[i + 2])) {
        merged = `${merged} ${current}${normalizeText(lines[i + 2])}`;
        i += 2;
      } else {
        merged = `${merged} ${current}`;
        i += 1;
      }
      result.push(merged);
      continue;
    }

    // 孤立碎片直接丢弃，避免污染
    if (isPortPrefixLine(current) || isPortSuffixLine(current)) {
      continue;
    }

    result.push(current);
  }

  return result;
}

function groupRowsByY(items) {
  const groups = [];
  for (const it of items) {
    const txt = normalizeText(it.str);
    if (!txt) continue;
    const x = it.transform[4];
    const y = it.transform[5];

    let g = null;
    for (const candidate of groups) {
      if (Math.abs(candidate.y - y) <= 2.4) {
        g = candidate;
        break;
      }
    }
    if (!g) {
      g = { y, parts: [] };
      groups.push(g);
    }
    g.parts.push({ x, txt });
  }

  groups.sort((a, b) => b.y - a.y);
  return groups;
}

function detectAnchors(firstPageItems) {
  const headerKeywords = ["序号", "出境/入境", "出入境日期", "证件名称", "证件号码", "出入境口岸", "航班号"];
  const rows = groupRowsByY(firstPageItems);

  let headerRow = null;
  for (const row of rows) {
    const joined = row.parts.map((p) => p.txt).join(" ");
    const hitCount = headerKeywords.filter((k) => joined.includes(k)).length;
    if (hitCount >= 4) {
      headerRow = row;
      break;
    }
  }

  if (!headerRow) {
    return [70, 160, 280, 410, 540, 680, 810];
  }

  const anchors = new Array(7).fill(null);
  for (const part of headerRow.parts) {
    const txt = part.txt;
    const idx = headerKeywords.findIndex((k) => txt.includes(k));
    if (idx >= 0 && anchors[idx] === null) {
      anchors[idx] = part.x;
    }
  }

  const fallback = [70, 160, 280, 410, 540, 680, 810];
  for (let i = 0; i < anchors.length; i++) {
    if (anchors[i] === null) anchors[i] = fallback[i];
  }
  return anchors;
}

function nearestAnchorIndex(x, anchors) {
  let minDiff = Number.MAX_SAFE_INTEGER;
  let idx = 0;
  for (let i = 0; i < anchors.length; i++) {
    const diff = Math.abs(x - anchors[i]);
    if (diff < minDiff) {
      minDiff = diff;
      idx = i;
    }
  }
  return idx;
}

function rowGroupToColumns(rowGroup, anchors) {
  const cols = new Array(7).fill("").map(() => []);
  const sortedParts = [...rowGroup.parts].sort((a, b) => a.x - b.x);

  for (const part of sortedParts) {
    const idx = nearestAnchorIndex(part.x, anchors);
    cols[idx].push(part.txt);
  }

  return cols.map((arr) => normalizeText(arr.join(" ")));
}

function columnsToRecord(cols) {
  const serial = normalizeText(cols[0]);
  const movement = normalizeText(cols[1]);
  const dateText = normalizeText(cols[2]);

  if (!/^\d+$/.test(serial)) return null;
  if (!(movement === "出境" || movement === "入境")) return null;
  if (!isValidDateKey(dateText)) return null;

  const certName = normalizeText(cols[3]).replace(/\s+/g, "");
  const certNum = extractCertNumber(cols[4]) || normalizeText(cols[4]).toUpperCase();
  const port = normalizePortName(cols[5]);
  const flight = normalizeFlightCode(cols[6]);

  if (!port || !certName || !certNum || !isLikelyCertNumber(certNum)) return null;

  return {
    "序号": serial,
    "出境/入境": movement,
    "出入境日期": dateText,
    "证件名称": certName,
    "证件号码": certNum,
    "出入境口岸": port,
    "航班号": flight,
  };
}

function parseRecordChunk(serial, movement, dateText, rawRest) {
  const rest = normalizeText(rawRest).replace(/\s+/g, " ").trim();
  if (!rest) return null;

  const certNum = extractCertNumber(rest);
  if (!certNum) return null;

  const certIdx = rest.toUpperCase().indexOf(certNum);
  if (certIdx < 0) return null;
  const certName = normalizeText(rest.slice(0, certIdx)).replace(/\s+/g, "");
  let tail = normalizeText(rest.slice(certIdx + certNum.length));

  let flight = "";
  const flightMatch = tail.match(/([A-Z0-9]{4,9})\s*$/i);
  if (flightMatch) {
    const normalized = normalizeFlightCode(flightMatch[1]);
    if (normalized) {
      flight = normalized;
      tail = normalizeText(tail.slice(0, tail.length - flightMatch[0].length));
    }
  }

  const port = normalizePortName(tail);
  if (!certName || !certNum || !port) return null;

  return {
    "序号": String(serial),
    "出境/入境": movement,
    "出入境日期": dateText,
    "证件名称": certName,
    "证件号码": certNum,
    "出入境口岸": port,
    "航班号": flight,
  };
}

function isContinuationColumns(cols) {
  const serial = normalizeText(cols[0]);
  const movement = normalizeText(cols[1]);
  const dateText = normalizeText(cols[2]);
  if (serial || movement || dateText) return false;

  const tail = normalizeText(`${cols[5]} ${cols[6]}`);
  if (!tail) return false;
  if (/本电子文件|制作日期|国家移民管理局|https?:\/\//.test(tail)) return false;
  return tail.length <= 8;
}

function parseRecordsByAnchors(items) {
  const anchors = detectAnchors(items);
  const rows = groupRowsByY(items);
  const records = [];
  let lastRecord = null;

  for (const row of rows) {
    const cols = rowGroupToColumns(row, anchors);
    const rec = columnsToRecord(cols);
    if (rec) {
      records.push(rec);
      lastRecord = rec;
      continue;
    }

    if (lastRecord && isContinuationColumns(cols)) {
      const mergedPort = normalizePortName(`${lastRecord["出入境口岸"]} ${cols[5]} ${cols[6]}`);
      if (mergedPort) lastRecord["出入境口岸"] = mergedPort;

      const mergedFlight = normalizeFlightCode(`${lastRecord["航班号"]} ${cols[6]}`);
      if (mergedFlight) lastRecord["航班号"] = mergedFlight;
    }
  }

  return records;
}

async function extractRecordsFromPdf(file) {
  const buffer = await file.arrayBuffer();
  const task = pdfjsLib.getDocument({ data: buffer });
  const pdf = await task.promise;

  const records = [];
  const dedupe = new Set();

  for (let p = 1; p <= pdf.numPages; p++) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    const addRecord = (rec) => {
      if (!rec) return;
      const key = `${rec["序号"]}|${rec["出境/入境"]}|${rec["出入境日期"]}|${rec["证件号码"]}`;
      if (dedupe.has(key)) return;
      dedupe.add(key);
      records.push(rec);
    };

    const rows = groupRowsByY(content.items);
    const lines = rows
      .map((r) => [...r.parts].sort((a, b) => a.x - b.x).map((x) => x.txt).join(" "))
      .filter((line) => !isNoiseLine(line));

    const stitchedLines = stitchWrappedPortLines(lines);

    const pageText = stitchedLines.join("\n");
    const pattern = /(\d+)\s+(出境|入境)\s+(\d{4}-\d{2}-\d{2})\s+([\s\S]*?)(?=(?:\n|\s)+\d+\s+(?:出境|入境)\s+\d{4}-\d{2}-\d{2}\b|$)/g;

    let match;
    while ((match = pattern.exec(pageText)) !== null) {
      const serial = parseInt(match[1], 10);
      const movement = match[2];
      const dateText = match[3];
      const rest = match[4];

      const rec = parseRecordChunk(serial, movement, dateText, rest);
      if (!rec) continue;
      addRecord(rec);
    }

    // 第二通道：按列锚点解析，增强对版式变化和字段长度变化的兼容性。
    const anchorRecords = parseRecordsByAnchors(content.items);
    for (const rec of anchorRecords) {
      addRecord(rec);
    }
  }

  records.sort((a, b) => toSerial(a) - toSerial(b));

  return {
    "来源文件": file.name,
    "生成时间": formatDateTimeByTimezone(new Date(), "Asia/Shanghai"),
    "提取页数": pdf.numPages,
    "表头": BASE_HEADERS,
    "总行数": records.length,
    "数据": records,
  };
}

function validateRecords(records) {
  const errors = [];
  const indexed = [];

  records.forEach((r, i) => {
    const serial = normalizeText(r["序号"]);
    if (!/^\d+$/.test(serial)) {
      errors.push(`第 ${i + 1} 条记录的序号不是有效数字：${serial}`);
      return;
    }
    indexed.push({ serial: Number(serial), row: r });
  });

  indexed.sort((a, b) => b.serial - a.serial);
  const sorted = indexed.map((x) => x.row);
  if (!sorted.length) return { valid: false, errors: ["没有可核验记录。"] };

  let prevDate = null;
  let prevType = null;
  for (const row of sorted) {
    const serial = normalizeText(row["序号"]);
    const typ = normalizeText(row["出境/入境"]);
    const d = normalizeText(row["出入境日期"]);

    if (!(typ === "出境" || typ === "入境")) {
      errors.push(`序号 ${serial} 的出入境字段无效：${typ}`);
      continue;
    }

    if (!isValidDateKey(d)) {
      errors.push(`序号 ${serial} 的日期格式无效：${d}`);
      continue;
    }

    const current = d;

    if (prevDate && current < prevDate) {
      errors.push(`序号 ${serial} 的日期 ${d} 早于上一条记录日期 ${prevDate}`);
    }

    if (prevType && prevType === typ) {
      errors.push(`出现连续相同状态：序号 ${serial} 与上一条记录均为 ${typ}`);
    }

    prevDate = current;
    prevType = typ;
  }

  return { valid: errors.length === 0, errors };
}

function getLatestRecord(records) {
  if (!records.length) return null;
  return records.reduce((a, b) => (toSerial(a) < toSerial(b) ? a : b));
}

function getEarliestRecord(records) {
  if (!records.length) return null;
  return records.reduce((a, b) => (toSerial(a) > toSerial(b) ? a : b));
}

function getLatestExitIndex(records) {
  let idx = -1;
  let serial = Number.MAX_SAFE_INTEGER;
  records.forEach((row, i) => {
    if (normalizeText(row["出境/入境"]) !== "出境") return;
    const s = toSerial(row);
    if (s < serial) {
      serial = s;
      idx = i;
    }
  });
  return idx;
}

function getEarliestEntryIndex(records) {
  let idx = -1;
  let serial = -1;
  records.forEach((row, i) => {
    if (normalizeText(row["出境/入境"]) !== "入境") return;
    const s = toSerial(row);
    if (s > serial) {
      serial = s;
      idx = i;
    }
  });
  return idx;
}

function buildVirtualEntry(latestRecord) {
  const serial = toSerial(latestRecord);
  return {
    "序号": String(serial > 1 ? serial - 1 : 0),
    "出境/入境": "入境",
    "出入境日期": getBeijingTodayKey(),
    "证件名称": latestRecord["证件名称"] || "",
    "证件号码": latestRecord["证件号码"] || "",
    "出入境口岸": latestRecord["出入境口岸"] || "",
    "航班号": "",
  };
}

function buildManualExitFromEarliestEntry(earliestRecord, exitDateText) {
  const serial = toSerial(earliestRecord);
  return {
    "序号": String(serial + 1),
    "出境/入境": "出境",
    "出入境日期": exitDateText,
    "证件名称": earliestRecord["证件名称"] || "",
    "证件号码": earliestRecord["证件号码"] || "",
    "出入境口岸": earliestRecord["出入境口岸"] || "",
    "航班号": "",
  };
}

function configureAdjustBoxForLatestExit(latestRecord) {
  const dateText = formatDateZh(normalizeText(latestRecord?.["出入境日期"]));
  state.adjustMode = "latest-exit";
  state.earliestEntryDate = "";

  if (el.adjustTitle) {
    el.adjustTitle.textContent = `检测到最近记录为${dateText}出境，无入境记录，请选择处理方式：`;
  }
  if (el.adjustPrimaryText) {
    el.adjustPrimaryText.textContent = "我还在境外/我刚入境系统还未同步数据";
  }
  if (el.btnAddVirtual) {
    el.btnAddVirtual.innerHTML = "（<i class=\"bi bi-check\"></i>推荐）计算截止到今日数据";
  }
  if (el.adjustSecondaryItem) {
    el.adjustSecondaryItem.classList.remove("hidden");
  }
  if (el.adjustSecondaryText) {
    el.adjustSecondaryText.textContent = "另一种选择";
  }
  if (el.btnUseLastEntry) {
    el.btnUseLastEntry.textContent = "计算截止到上次入境数据";
  }
  if (el.adjustManualItem) {
    el.adjustManualItem.classList.add("hidden");
  }
  if (el.manualExitDate) {
    el.manualExitDate.value = "";
    el.manualExitDate.removeAttribute("max");
  }
  setManualExitWarn("");
}

function configureAdjustBoxForEarliestEntry(earliestRecord) {
  const rawDate = normalizeText(earliestRecord?.["出入境日期"]);
  const dateText = formatDateZh(rawDate);
  state.adjustMode = "earliest-entry";
  state.earliestEntryDate = /^\d{4}-\d{2}-\d{2}$/.test(rawDate) ? rawDate : "";

  if (el.adjustTitle) {
    el.adjustTitle.textContent = `检测到最早记录为${dateText}入境，无更早出境记录，请选择处理方式：`;
  }
  if (el.adjustPrimaryText) {
    el.adjustPrimaryText.textContent = "忽略这条记录";
  }
  if (el.btnAddVirtual) {
    el.btnAddVirtual.textContent = "忽略最早入境记录";
  }
  if (el.adjustSecondaryItem) {
    el.adjustSecondaryItem.classList.add("hidden");
  }
  if (el.adjustManualItem) {
    el.adjustManualItem.classList.remove("hidden");
  }
  if (el.manualExitDate) {
    el.manualExitDate.value = state.earliestEntryDate;
    if (state.earliestEntryDate) {
      el.manualExitDate.max = state.earliestEntryDate;
    } else {
      el.manualExitDate.removeAttribute("max");
    }
  }
  setManualExitWarn("");
}

function renderStats(data) {
  const stats = [
    ["来源文件", data["来源文件"]],
    ["PDF总页数", String(data["提取页数"])],
    ["有效记录数", String(data["总行数"])],
  ];
  el.stats.innerHTML = stats
    .map(([k, v]) => `<article class="stat"><div class="k">${k}</div><div class="v">${v}</div></article>`)
    .join("");
}

function buildChecklist(container, items, onChange) {
  container.innerHTML = "";

  const field = container.parentElement;
  let actionRow = null;
  let selectAllBtn = null;
  let clearAllBtn = null;
  if (field) {
    const rowSelector = `.checklist-actions-row[data-for="${container.id}"]`;
    actionRow = field.querySelector(rowSelector);
    if (!actionRow) {
      actionRow = document.createElement("div");
      actionRow.className = "checklist-actions-row";
      actionRow.dataset.for = container.id;
      field.insertBefore(actionRow, container);
    }

    // Clean up legacy standalone button from old DOM structure.
    const legacySelector = `button.checklist-select-all-row[data-for="${container.id}"]`;
    const legacyBtn = field.querySelector(legacySelector);
    if (legacyBtn && legacyBtn.parentElement !== actionRow) {
      legacyBtn.remove();
    }

    selectAllBtn = actionRow.querySelector(".checklist-select-all");
    if (!selectAllBtn) {
      selectAllBtn = document.createElement("button");
      selectAllBtn.type = "button";
      selectAllBtn.className = "checklist-select-all";
      selectAllBtn.textContent = "全部选择";
      actionRow.appendChild(selectAllBtn);
    }

    clearAllBtn = actionRow.querySelector(".checklist-clear-all");
    if (!clearAllBtn) {
      clearAllBtn = document.createElement("button");
      clearAllBtn.type = "button";
      clearAllBtn.className = "checklist-clear-all";
      clearAllBtn.textContent = "全部取消";
      clearAllBtn.title = "全部取消选择";
      actionRow.appendChild(clearAllBtn);
    }
  }

  const list = document.createElement("div");
  list.className = "checklist-items";
  container.appendChild(list);

  items.forEach((it) => {
    const label = document.createElement("label");
    label.innerHTML = `<input type="checkbox" data-value="${encodeURIComponent(it.value)}" checked /> ${it.label}`;
    list.appendChild(label);
  });

  const itemInputs = Array.from(list.querySelectorAll("input[data-value]"));

  if (selectAllBtn) {
    selectAllBtn.disabled = itemInputs.length === 0;
    selectAllBtn.onclick = () => {
      itemInputs.forEach((it) => {
        it.checked = true;
      });
      if (typeof onChange === "function") onChange();
    };
  }

  if (clearAllBtn) {
    clearAllBtn.disabled = itemInputs.length === 0;
    clearAllBtn.onclick = () => {
      itemInputs.forEach((it) => {
        it.checked = false;
      });
      if (typeof onChange === "function") onChange();
    };
  }

  itemInputs.forEach((it) => {
    it.addEventListener("change", () => {
      if (typeof onChange === "function") onChange();
    });
  });
}

function getCheckedValuesStrict(container) {
  return Array.from(container.querySelectorAll('input[data-value]:checked')).map((x) => decodeURIComponent(x.dataset.value));
}

function getSelected(container) {
  const items = Array.from(container.querySelectorAll("input[data-value]"));
  const selected = items.filter((x) => x.checked).map((x) => decodeURIComponent(x.dataset.value));
  return selected.length ? selected : items.map((x) => decodeURIComponent(x.dataset.value));
}

function getCredentialKey(row) {
  return `${normalizeText(row["证件名称"])}|${normalizeText(row["证件号码"])}`;
}

function updatePortChecklistByCredentials(records) {
  const selectedCred = new Set(getCheckedValuesStrict(el.credentialList));
  const previousPortValues = new Set(
    Array.from(el.portList.querySelectorAll("input[data-value]")).map((x) => decodeURIComponent(x.dataset.value))
  );
  const previousSelectedPorts = new Set(getCheckedValuesStrict(el.portList));
  const ports = new Set();

  records.forEach((r) => {
    const key = getCredentialKey(r);
    if (!selectedCred.has(key)) return;
    const port = normalizePortName(r["出入境口岸"]);
    if (port) ports.add(port);
  });

  const portItems = Array.from(ports)
    .sort((a, b) => a.localeCompare(b, "zh-CN"))
    .map((x) => ({ value: x, label: x }));

  buildChecklist(el.portList, portItems, () => requestCalculation());

  const itemInputs = Array.from(el.portList.querySelectorAll('input[data-value]'));
  if (!itemInputs.length) return;

  let hasAnyChecked = false;
  itemInputs.forEach((it) => {
    const value = decodeURIComponent(it.dataset.value);
    if (!previousPortValues.has(value)) {
      // Newly appeared ports should be selected by default.
      it.checked = true;
      hasAnyChecked = true;
    } else if (previousSelectedPorts.has(value)) {
      it.checked = true;
      hasAnyChecked = true;
    } else {
      it.checked = false;
    }
  });

  if (!hasAnyChecked) {
    itemInputs.forEach((it) => {
      it.checked = true;
    });
  }
}

function setupCalc(dataset) {
  const records = dataset["数据"];
  const dates = records
    .map((r) => normalizeText(r["出入境日期"]))
    .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))
    .sort();
  if (dates.length) {
    el.startDate.value = dates[0];
    el.endDate.value = dates[dates.length - 1];
  }

  const credMap = new Map();
  records.forEach((r) => {
    const certType = normalizeText(r["证件名称"]);
    const certNum = normalizeText(r["证件号码"]);
    if (certType && certNum) {
      const key = `${certType}|${certNum}`;
      credMap.set(key, `${certType}${certNum}`);
    }
  });

  buildChecklist(
    el.credentialList,
    Array.from(credMap.entries())
      .sort((a, b) => a[1].localeCompare(b[1], "zh-CN"))
      .map(([value, label]) => ({ value, label })),
    () => {
      updatePortChecklistByCredentials(records);
      requestCalculation();
    }
  );

  updatePortChecklistByCredentials(records);

  el.calcSection.classList.remove("hidden");
  setCalcOrderWarn("");
  el.resultBox.classList.add("hidden");
  requestCalculation();
}

function calculateAbroad(records) {
  const sorted = [...records].sort((a, b) => toSerial(b) - toSerial(a));
  const abroad = new Map();
  let pendingExit = null;

  for (const row of sorted) {
    const typ = normalizeText(row["出境/入境"]);
    const d = normalizeText(row["出入境日期"]);
    if (!isValidDateKey(d)) continue;
    const cert = `${normalizeText(row["证件名称"])}${normalizeText(row["证件号码"])}`;

    if (typ === "出境") {
      pendingExit = { date: d, cert };
      continue;
    }

    if (typ !== "入境" || !pendingExit) continue;
    for (let key = pendingExit.date; key <= d; key = addDaysToDateKey(key, 1)) {
      if (!abroad.has(key)) abroad.set(key, new Set());
      abroad.get(key).add(pendingExit.cert);
    }
    pendingExit = null;
  }

  return abroad;
}

function findBoundaryRecordsByDate(dateKey, rawSortedRecords) {
  let latestExit = null;
  let nearestExit = null;
  let latestEntry = null;
  let nearestEntry = null;
  for (const row of rawSortedRecords) {
    const typ = normalizeText(row["出境/入境"]);
    const d = normalizeText(row["出入境日期"]);
    if (typ === "出境" && d <= dateKey) {
      latestExit = row;
    }
    if (!nearestExit && typ === "出境" && d >= dateKey) {
      nearestExit = row;
    }
    if (typ === "入境" && d <= dateKey) {
      latestEntry = row;
    }
    if (typ === "入境" && d >= dateKey) {
      nearestEntry = row;
      break;
    }
  }
  return { latestExit, nearestExit, latestEntry, nearestEntry };
}

function getFilteredRawPdfRecords(selectedCred, selectedPort) {
  if (!state.originalData || !Array.isArray(state.originalData["数据"])) return [];
  return state.originalData["数据"].filter((r) => {
    const d = normalizeText(r["出入境日期"]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return false;
    const c = `${normalizeText(r["证件名称"])}|${normalizeText(r["证件号码"])}`;
    if (!selectedCred.has(c)) return false;
    const p = normalizePortName(r["出入境口岸"]);
    if (!p || !selectedPort.has(p)) return false;
    return true;
  });
}

function buildVirtualBoundaryRecord(baseRow, movement, dateText, serialText) {
  return {
    "序号": String(serialText),
    "出境/入境": movement,
    "出入境日期": dateText,
    "证件名称": normalizeText(baseRow?.["证件名称"] || ""),
    "证件号码": normalizeText(baseRow?.["证件号码"] || ""),
    "出入境口岸": normalizeText(baseRow?.["出入境口岸"] || ""),
    "航班号": normalizeText(baseRow?.["航班号"] || ""),
  };
}

function validateSelectedOrder(records, startDate = "", endDate = "") {
  const hasRangeLimit = isValidDateKey(startDate) && isValidDateKey(endDate) && startDate <= endDate;
  const allSorted = [...records]
    .filter((r) => {
      const typ = normalizeText(r["出境/入境"]);
      const d = normalizeText(r["出入境日期"]);
      return (typ === "出境" || typ === "入境") && /^\d{4}-\d{2}-\d{2}$/.test(d);
    })
    .sort((a, b) => {
      const da = normalizeText(a["出入境日期"]);
      const db = normalizeText(b["出入境日期"]);
      if (da !== db) return da.localeCompare(db, "zh-CN");
      // 同一天按序号从大到小，尽量贴合真实发生顺序。
      return toSerial(b) - toSerial(a);
    });

  const inRangeRecords = hasRangeLimit
    ? allSorted.filter((r) => {
        const d = normalizeText(r["出入境日期"]);
        return d >= startDate && d <= endDate;
      })
    : allSorted;

  const sorted = [...inRangeRecords];
  if (!sorted.length) return "";

  if (hasRangeLimit) {
    const first = sorted[0];
    const firstType = normalizeText(first["出境/入境"]);
    if (firstType === "入境") {
      const firstIndex = allSorted.indexOf(first);
      let baseRow = first;
      for (let i = firstIndex - 1; i >= 0; i--) {
        if (normalizeText(allSorted[i]["出境/入境"]) === "出境") {
          baseRow = allSorted[i];
          break;
        }
      }
      sorted.unshift(buildVirtualBoundaryRecord(baseRow, "出境", startDate, "999999999"));
    }

    const last = sorted[sorted.length - 1];
    const lastType = normalizeText(last["出境/入境"]);
    if (lastType === "出境") {
      const lastIndex = allSorted.indexOf(last);
      let baseRow = last;
      for (let i = lastIndex + 1; i < allSorted.length; i++) {
        if (normalizeText(allSorted[i]["出境/入境"]) === "入境") {
          baseRow = allSorted[i];
          break;
        }
      }
      sorted.push(buildVirtualBoundaryRecord(baseRow, "入境", endDate, "0"));
    }

    sorted.sort((a, b) => {
      const da = normalizeText(a["出入境日期"]);
      const db = normalizeText(b["出入境日期"]);
      if (da !== db) return da.localeCompare(db, "zh-CN");
      return toSerial(b) - toSerial(a);
    });
  }

  let expected = "出境";
  let lastExitDate = "";

  for (const row of sorted) {
    const typ = normalizeText(row["出境/入境"]);
    const d = normalizeText(row["出入境日期"]);

    if (typ !== expected) {
      if (typ === "入境") {
        return `根据所勾选的数据，${formatDateZh(d)}入境，但无对应出境记录，请核查。`;
      }

      if (lastExitDate) {
        return `根据所勾选的数据，${formatDateZh(lastExitDate)}出境，无入境记录，${formatDateZh(d)}又出境，请核查。`;
      }

      return `根据所勾选的数据，${formatDateZh(d)}出境，但无对应入境记录，请核查。`;
    }

    if (typ === "出境") {
      lastExitDate = d;
      expected = "入境";
    } else {
      expected = "出境";
    }
  }

  if (expected === "入境" && lastExitDate) {
    return `根据所勾选的数据，${formatDateZh(lastExitDate)}出境后无入境记录，请核查。`;
  }

  return "";
}

function renderCalendar(start, end, abroadSet, rawPdfRecords) {
  el.calendar.innerHTML = "";
  state.calendarDayInfoMap = new Map();
  state.activeCalendarDay = "";
  hideCalendarTooltip();
  if (!isValidDateKey(start) || !isValidDateKey(end) || start > end) {
    el.totalDays.textContent = "日期范围无效，请重新选择。";
    return;
  }

  const inRangeAbroadDays = Array.from(abroadSet)
    .filter((d) => d >= start && d <= end)
    .sort((a, b) => a.localeCompare(b, "zh-CN"));
  const inRangeAbroadSet = new Set(inRangeAbroadDays);

  // 仅裁掉首尾没有在境外日期的月份；中间月份即使没有在境外日期也保留显示。
  let displayStartMonth = parseDateKey(start);
  let displayEndMonth = parseDateKey(end);
  if (inRangeAbroadDays.length) {
    displayStartMonth = parseDateKey(inRangeAbroadDays[0]);
    displayEndMonth = parseDateKey(inRangeAbroadDays[inRangeAbroadDays.length - 1]);
  }
  if (!displayStartMonth || !displayEndMonth) {
    el.totalDays.textContent = "日期范围无效，请重新选择。";
    return;
  }

  let highlightedDays = 0;

  let cursorYear = displayStartMonth.y;
  let cursorMonth = displayStartMonth.m;
  const endYear = displayEndMonth.y;
  const endMonth = displayEndMonth.m;
  const frag = document.createDocumentFragment();
  const rawSorted = [...(rawPdfRecords || [])].sort((a, b) => {
    const da = normalizeText(a["出入境日期"]);
    const db = normalizeText(b["出入境日期"]);
    if (da !== db) return da.localeCompare(db, "zh-CN");
    return toSerial(b) - toSerial(a);
  });

  while (cursorYear < endYear || (cursorYear === endYear && cursorMonth <= endMonth)) {
    const y = cursorYear;
    const m = cursorMonth;
    const days = daysInMonthByDateKey(y, m);

    const card = document.createElement("article");
    card.className = "month";
    const html = [`<h4>${y}年${pad2(m)}月</h4>`, CALENDAR_WEEKDAY_ROW, "<div class=\"week\">"];
    const lead = weekdayMon0ByDateKey(y, m, 1);
    for (let i = 0; i < lead; i++) {
      html.push("<div class=\"day blank\"></div>");
    }

    for (let d = 1; d <= days; d++) {
      const key = formatDateKey(y, m, d);
      const inRange = key >= start && key <= end;
      const inAbroad = inRangeAbroadSet.has(key);
      const { latestExit, nearestExit, latestEntry, nearestEntry } = findBoundaryRecordsByDate(key, rawSorted);
      state.calendarDayInfoMap.set(key, {
        date: key,
        inRange,
        inAbroad,
        abroadExitText: formatDayRecordBrief(latestExit),
        abroadEntryText: formatDayRecordBrief(nearestEntry),
        domesticEntryText: formatDayRecordBrief(latestEntry),
        domesticExitText: formatDayRecordBrief(nearestExit),
      });
      if (inRangeAbroadSet.has(key)) {
        html.push(`<div class="day abroad" data-date="${key}">${d}</div>`);
        highlightedDays += 1;
      } else {
        html.push(`<div class="day" data-date="${key}">${d}</div>`);
      }
    }
    html.push("</div>");
    card.innerHTML = html.join("");
    frag.appendChild(card);
    if (cursorMonth === 12) {
      cursorMonth = 1;
      cursorYear += 1;
    } else {
      cursorMonth += 1;
    }
  }

  el.calendar.replaceChildren(frag);
  el.totalDays.innerHTML = `<span class="total-label">累计总出境天数</span><span class="total-value">${highlightedDays} 天</span>`;
  state.highlightedDays = highlightedDays;
  updateQualificationPanels(highlightedDays);
}

function ensureCalendarTooltip() {
  if (calendarTooltipEl && calendarTooltipEl.isConnected) return calendarTooltipEl;
  calendarTooltipEl = document.createElement("div");
  calendarTooltipEl.className = "calendar-tooltip hidden";
  document.body.appendChild(calendarTooltipEl);
  return calendarTooltipEl;
}

function hideCalendarTooltip() {
  if (!calendarTooltipEl) return;
  calendarTooltipEl.classList.add("hidden");
}

function showCalendarTooltip(dayCell) {
  if (!dayCell) return;
  const dateKey = normalizeText(dayCell.dataset.date);
  if (!dateKey) return;
  const dayInfo = state.calendarDayInfoMap.get(dateKey);
  if (!dayInfo) return;

  const tip = ensureCalendarTooltip();
  if (!dayInfo.inRange) {
    tip.innerHTML = "<p class=\"calendar-tooltip-row\">本日不在计算范围内</p>";
  } else if (!dayInfo.inAbroad) {
    tip.innerHTML = [
      `<p class="calendar-tooltip-row"><strong>${escapeHtml(formatDateZh(dateKey))}</strong></p>`,
      "<p class=\"calendar-tooltip-row\">本日不在境外</p>",
      `<p class="calendar-tooltip-row">入境：${escapeHtml(dayInfo.domesticEntryText)}</p>`,
      `<p class="calendar-tooltip-row">出境：${escapeHtml(dayInfo.domesticExitText)}</p>`,
    ].join("");
  } else {
    tip.innerHTML = [
      `<p class="calendar-tooltip-row"><strong>${escapeHtml(formatDateZh(dateKey))}</strong></p>`,
      "<p class=\"calendar-tooltip-row\">本日在境外</p>",
      `<p class="calendar-tooltip-row">出境：${escapeHtml(dayInfo.abroadExitText)}</p>`,
      `<p class="calendar-tooltip-row">入境：${escapeHtml(dayInfo.abroadEntryText)}</p>`,
    ].join("");
  }
  tip.classList.remove("hidden");

  const rect = dayCell.getBoundingClientRect();
  const margin = 10;
  const tipRect = tip.getBoundingClientRect();
  let left = rect.left + rect.width / 2 - tipRect.width / 2;
  left = Math.max(margin, Math.min(window.innerWidth - tipRect.width - margin, left));
  let top = rect.top - tipRect.height - 8;
  if (top < margin) {
    top = rect.bottom + 8;
  }
  tip.style.left = `${left}px`;
  tip.style.top = `${top}px`;
}

function getCalendarDayCell(event) {
  const target = event.target;
  if (!target || typeof target.closest !== "function") return null;
  return target.closest(".day[data-date]");
}

function handleCalendarInteraction(event) {
  const dayCell = getCalendarDayCell(event);
  if (!dayCell) return;
  state.activeCalendarDay = normalizeText(dayCell.dataset.date);
  showCalendarTooltip(dayCell);
}

function runValidation(dataset, adjustmentSummary = "") {
  state.workingData = dataset;
  state.adjustmentSummary = adjustmentSummary || "";
  const records = [...dataset["数据"]].sort((a, b) => toSerial(a) - toSerial(b));
  const latest = records.length ? records[0] : null;
  const earliest = records.length ? getEarliestRecord(records) : null;

  if (latest && normalizeText(latest["出境/入境"]) === "出境") {
    state.validated = false;
    configureAdjustBoxForLatestExit(latest);
    el.status.classList.add("hidden");
    el.adjustBox.classList.remove("hidden");
    el.adjustBox.classList.add("warn-tone");
    el.calcSection.classList.add("hidden");
    return;
  }

  if (earliest && normalizeText(earliest["出境/入境"]) === "入境") {
    state.validated = false;
    configureAdjustBoxForEarliestEntry(earliest);
    el.status.classList.add("hidden");
    el.adjustBox.classList.remove("hidden");
    el.adjustBox.classList.add("warn-tone");
    el.calcSection.classList.add("hidden");
    return;
  }

  const result = validateRecords(records);
  if (!result.valid) {
    state.validated = false;
    el.status.classList.remove("hidden");
    el.adjustBox.classList.add("hidden");
    el.adjustBox.classList.remove("warn-tone");
    el.calcSection.classList.add("hidden");
    setStatus(`核验失败：\n${result.errors.slice(0, 10).join("\n")}`, "warn");
    return;
  }

  state.validated = true;
  el.status.classList.remove("hidden");
  el.adjustBox.classList.add("hidden");
  el.adjustBox.classList.remove("warn-tone");
  setupCalc(state.workingData);
  const totalText = state.adjustmentSummary ? state.adjustmentSummary : String(records.length);
  setStatus(`核验通过：PDF 数据有效。\n总记录数：${totalText}`, "ok");
}

async function handlePdf(file) {
  if (!file || !file.name.toLowerCase().endsWith(".pdf")) {
    setStatus("请选择 PDF 文件。", "warn");
    return;
  }

  try {
    setStatus("正在提取 PDF 数据，请稍候...");
    const data = await extractRecordsFromPdf(file);
    state.originalData = JSON.parse(JSON.stringify(data));
    state.baseTotalRows = Number(data["总行数"]) || 0;
    state.adjustmentSummary = "";
    // 上传并解析成功后即可导出原始数据（与后续处理分支无关）
    el.btnExport.classList.remove("hidden");
    el.btnExport.disabled = false;
    renderStats(data);
    runValidation(JSON.parse(JSON.stringify(data)), "");
  } catch (err) {
    console.error(err);
    el.btnExport.classList.add("hidden");
    setStatus("PDF 解析失败，请检查文件格式。", "warn");
  }
}

function openPdfPicker() {
  // Reset value so selecting the same file still emits a change event.
  el.pdfInput.value = "";
  el.pdfInput.click();
}

el.dropzone.addEventListener("click", openPdfPicker);
el.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    openPdfPicker();
  }
});
el.pdfInput.addEventListener("change", (e) => {
  const file = e.target.files && e.target.files[0];
  handlePdf(file);
});

el.dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  el.dropzone.classList.add("drag");
});
el.dropzone.addEventListener("dragleave", () => el.dropzone.classList.remove("drag"));
el.dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  el.dropzone.classList.remove("drag");
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  handlePdf(file);
});

el.btnAddVirtual.addEventListener("click", () => {
  if (!state.workingData) return;

  if (state.adjustMode === "earliest-entry") {
    const copy = JSON.parse(JSON.stringify(state.workingData));
    const idx = getEarliestEntryIndex(copy["数据"]);
    if (idx >= 0) copy["数据"].splice(idx, 1);
    copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
    copy["总行数"] = copy["数据"].length;
    const base = state.baseTotalRows || copy["总行数"] + 1;
    runValidation(copy, `${base}-1（忽略最早入境记录）`);
    return;
  }

  const copy = JSON.parse(JSON.stringify(state.workingData));
  const latestExitIndex = getLatestExitIndex(copy["数据"]);
  if (latestExitIndex < 0) return;
  const latestExit = copy["数据"][latestExitIndex];
  copy["数据"].push(buildVirtualEntry(latestExit));
  copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
  copy["总行数"] = copy["数据"].length;
  const base = state.baseTotalRows || copy["总行数"] - 1;
  runValidation(copy, `${base}+1（计算截止到今日数据）`);
});

if (el.manualExitDate) {
  el.manualExitDate.addEventListener("change", () => setManualExitWarn(""));
}

if (el.btnAddManualExit) {
  el.btnAddManualExit.addEventListener("click", () => {
    if (!state.workingData || state.adjustMode !== "earliest-entry") return;

    const manualDate = normalizeText(el.manualExitDate?.value);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(manualDate)) {
      setManualExitWarn("请先填写有效日期。\n日期不得晚于最早入境记录。");
      return;
    }

    const earliestDate = normalizeText(state.earliestEntryDate);
    if (earliestDate && manualDate > earliestDate) {
      setManualExitWarn(`该日期晚于最早入境记录（${formatDateZh(earliestDate)}），请重新填写。`);
      return;
    }

    const copy = JSON.parse(JSON.stringify(state.workingData));
    const earliest = getEarliestRecord(copy["数据"]);
    if (!earliest) {
      setManualExitWarn("未找到可补录的最早入境记录，请重新上传 PDF。");
      return;
    }

    copy["数据"].push(buildManualExitFromEarliestEntry(earliest, manualDate));
    copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
    copy["总行数"] = copy["数据"].length;
    const base = state.baseTotalRows || copy["总行数"] - 1;
    runValidation(copy, `${base}+1（手动补录最早出境记录）`);
  });
}

el.btnUseLastEntry.addEventListener("click", () => {
  if (!state.workingData) return;
  if (state.adjustMode !== "latest-exit") return;
  const copy = JSON.parse(JSON.stringify(state.workingData));
  const idx = getLatestExitIndex(copy["数据"]);
  if (idx >= 0) copy["数据"].splice(idx, 1);
  copy["数据"].sort((a, b) => toSerial(a) - toSerial(b));
  copy["总行数"] = copy["数据"].length;
  const base = state.baseTotalRows || copy["总行数"] + 1;
  runValidation(copy, `${base}-1（移除最后一次出境记录）`);
});

el.btnExport.addEventListener("click", () => {
  if (!state.originalData) return;
  const blob = new Blob([JSON.stringify(state.originalData, null, 2)], {
    type: "application/json;charset=utf-8",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "data.json";
  document.body.appendChild(a);
  a.click();
  URL.revokeObjectURL(a.href);
  a.remove();
});

if (el.btnRandomDemo) {
  el.btnRandomDemo.addEventListener("click", () => {
    const demoData = buildRandomDemoData();
    state.originalData = JSON.parse(JSON.stringify(demoData));
    state.baseTotalRows = Number(demoData["总行数"]) || 0;
    state.adjustmentSummary = "";
    el.btnExport.classList.remove("hidden");
    el.btnExport.disabled = false;
    renderStats(demoData);
    const prepared = prepareDemoWorkingData(demoData);
    runValidation(prepared.dataset, prepared.summary);
  });
}

el.btnReupload.addEventListener("click", () => {
  el.pdfInput.value = "";
  el.pdfInput.click();
});

function performCalculation() {
  if (!state.validated || !state.workingData) {
    setCalcOrderWarn("");
    el.resultBox.classList.add("hidden");
    return;
  }

  const start = el.startDate.value;
  const end = el.endDate.value;
  if (!start || !end) {
    setCalcOrderWarn("请先选择开始/结束日期。");
    el.resultBox.classList.add("hidden");
    return;
  }

  const selectedCredValues = getCheckedValuesStrict(el.credentialList);
  const selectedPortValues = getCheckedValuesStrict(el.portList);
  if (!selectedCredValues.length || !selectedPortValues.length) {
    setCalcOrderWarn("请至少勾选一个有效证件和一个有效通行口岸。");
    el.resultBox.classList.add("hidden");
    return;
  }

  const selectedCred = new Set(selectedCredValues);
  const selectedPort = new Set(selectedPortValues);
  const rawPdfFilteredBySelection = getFilteredRawPdfRecords(selectedCred, selectedPort);

  const filteredBySelection = state.workingData["数据"].filter((r) => {
    const d = normalizeText(r["出入境日期"]);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return false;

    const c = `${normalizeText(r["证件名称"])}|${normalizeText(r["证件号码"])}`;
    if (!selectedCred.has(c)) return false;

    const p = normalizePortName(r["出入境口岸"]);
    if (!p || !selectedPort.has(p)) return false;

    return true;
  });

  const orderWarn = validateSelectedOrder(filteredBySelection, start, end);
  if (orderWarn) {
    setCalcOrderWarn(orderWarn);
    el.resultBox.classList.add("hidden");
    return;
  }

  setCalcOrderWarn("");

  const abroadMap = calculateAbroad(filteredBySelection);
  renderCalendar(start, end, new Set(abroadMap.keys()), rawPdfFilteredBySelection);
  el.resultBox.classList.remove("hidden");
}

function requestCalculation() {
  if (state.calcRafId) {
    cancelAnimationFrame(state.calcRafId);
  }
  state.calcRafId = requestAnimationFrame(() => {
    state.calcRafId = 0;
    performCalculation();
  });
}

el.startDate.addEventListener("change", requestCalculation);
el.endDate.addEventListener("change", requestCalculation);
el.calendar.addEventListener("mouseover", handleCalendarInteraction);
el.calendar.addEventListener("mousemove", handleCalendarInteraction);
el.calendar.addEventListener("click", handleCalendarInteraction);
el.calendar.addEventListener("mouseleave", hideCalendarTooltip);

updateTimeDisplays();
setInterval(updateTimeDisplays, 60 * 1000);
initQualificationControls();
