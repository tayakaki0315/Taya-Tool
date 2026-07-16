"use client";

import { useMemo, useState } from "react";

type Language = "ja" | "en" | "zh";
type SheetMode = "single" | "perExpert" | "custom";

type ExpertRecord = {
  id: string;
  number: string;
  name: string;
  company: string;
  relevantExperience: string;
  employmentHistory: string;
  introduction: string;
  screening: string;
  fee: string;
  availability: string;
  sheetName: string;
  warnings: string[];
};

const MONTHS =
  "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December";
const WEEKDAYS =
  "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday";

const translations = {
  ja: {
    title: "Taya Expert List Builder",
    version: "v1.2",
    subtitle: "エキスパート情報を貼り付け、Excel リストをすぐに作成できます。",
    privacy: "入力内容はブラウザ内だけで処理され、サーバーへ送信・保存されません。",
    inputTitle: "1. エキスパート情報を貼り付け",
    inputHelp:
      "#1.1 - Name - … から始まる形式に対応しています。複数名をまとめて貼り付けても自動で分割します。",
    rawLabel: "エキスパート情報",
    rawPlaceholder: "ここにエキスパート情報を貼り付けてください…",
    parse: "解析する",
    clear: "すべてクリア",
    resultTitle: "2. 抽出結果を確認・修正",
    resultHelp:
      "複数名を個別に確認・修正できます。紹介文とQ&Aも抽出結果から直接編集できます。",
    experts: "名を抽出",
    noResults: "解析後、ここに抽出結果が表示されます。",
    remove: "このエキスパートを削除",
    warning: "要確認",
    sheetPlanner: "Sheet の分け方",
    sheetModeSingle: "全員を1つのSheetにまとめる",
    sheetModePerExpert: "1名につき1つのSheetを作る",
    sheetModeCustom: "Sheetを自由に分ける（ドラッグ）",
    singleSheetName: "Sheet名",
    customSheetHelp: "Sheet管理画面を開き、エキスパートをドラッグして自由に分けられます。",
    perExpertSheetHelp: "番号と名前を使って、エキスパートごとにSheetを作成します。",
    openSheetOrganizer: "Sheet管理画面を開く",
    organizerTitle: "Sheetを自由に分ける",
    organizerHelp: "Sheetを追加・変更し、エキスパートカードを希望のSheetへドラッグしてください。",
    newSheetPlaceholder: "新しいSheet名",
    addSheet: "Sheetを追加",
    emptySheet: "ここにエキスパートをドロップ",
    dragHint: "ドラッグして移動",
    deleteSheet: "空のSheetを削除",
    done: "完了",
    exportTitle: "3. Excel を作成",
    exportHelp:
      "アップロードされたテンプレートの C–K 列構成を保ち、見やすい書式を加えて出力します。",
    fileName: "ファイル名",
    export: "Excel をダウンロード",
    exporting: "Excel を作成中…",
    exported: "Excel を作成し、ダウンロードを開始しました。",
    parsed: "名のエキスパートを抽出しました。内容をご確認ください。",
    parseError:
      "「#1.1 - Name - …」で始まるエキスパート情報を見つけられませんでした。",
    exportError: "Excel の作成に失敗しました。もう一度お試しください。",
    exportEmpty: "先にエキスパート情報を解析してください。",
    fields: {
      number: "番号",
      name: "名前",
      company: "企業",
      relevantExperience: "関連経歴",
      employmentHistory: "過去の経歴",
      introduction: "紹介",
      screening: "スクリーニング質問に対する回答",
      fee: "金額",
      availability: "インタビュー可能な日付・候補",
      sheetName: "出力先Sheet名",
    },
  },
  en: {
    title: "Taya Expert List Builder",
    version: "v1.2",
    subtitle: "Paste expert profiles and turn them into a client-ready Excel list.",
    privacy: "Everything is processed in your browser. Nothing is uploaded or stored.",
    inputTitle: "1. Paste expert information",
    inputHelp:
      "Supports the standard #1.1 - Name - … format. Paste multiple experts at once and they will be separated automatically.",
    rawLabel: "Expert information",
    rawPlaceholder: "Paste expert information here…",
    parse: "Parse experts",
    clear: "Clear all",
    resultTitle: "2. Review and edit",
    resultHelp:
      "Review and edit multiple experts individually, including the introduction and screening Q&A.",
    experts: "experts found",
    noResults: "Parsed experts will appear here.",
    remove: "Remove this expert",
    warning: "Check",
    sheetPlanner: "Sheet layout",
    sheetModeSingle: "Put all experts in one sheet",
    sheetModePerExpert: "Create one sheet per expert",
    sheetModeCustom: "Arrange sheets freely (drag and drop)",
    singleSheetName: "Sheet name",
    customSheetHelp: "Open the sheet organizer and drag experts into any sheet.",
    perExpertSheetHelp: "A separate sheet will be created for each expert using their number and name.",
    openSheetOrganizer: "Open sheet organizer",
    organizerTitle: "Arrange sheets freely",
    organizerHelp: "Add or rename sheets, then drag each expert card into the destination sheet.",
    newSheetPlaceholder: "New sheet name",
    addSheet: "Add sheet",
    emptySheet: "Drop experts here",
    dragHint: "Drag to move",
    deleteSheet: "Delete empty sheet",
    done: "Done",
    exportTitle: "3. Create Excel",
    exportHelp:
      "Uses the uploaded template’s C–K column structure with improved formatting and usability.",
    fileName: "File name",
    export: "Download Excel",
    exporting: "Creating Excel…",
    exported: "The Excel file was created and the download has started.",
    parsed: "experts extracted. Please review the fields below.",
    parseError: "No expert block beginning with “#1.1 - Name - …” was found.",
    exportError: "The Excel file could not be created. Please try again.",
    exportEmpty: "Parse at least one expert first.",
    fields: {
      number: "Number",
      name: "Name",
      company: "Company",
      relevantExperience: "Relevant experience",
      employmentHistory: "Employment history",
      introduction: "Introduction",
      screening: "Screening Q&A",
      fee: "Fee",
      availability: "Availability",
      sheetName: "Destination sheet",
    },
  },
  zh: {
    title: "Taya Expert List Builder",
    version: "v1.2",
    subtitle: "粘贴专家资料，一键整理并生成客户用 Excel 名单。",
    privacy: "所有内容仅在浏览器内处理，不会上传或保存到服务器。",
    inputTitle: "1. 粘贴专家信息",
    inputHelp:
      "支持以 #1.1 - Name - … 开头的标准格式，也可以一次粘贴多位专家。",
    rawLabel: "专家信息",
    rawPlaceholder: "请在这里粘贴专家信息……",
    parse: "解析专家",
    clear: "全部清空",
    resultTitle: "2. 确认并修改结果",
    resultHelp: "可以逐一确认和修改多位专家，包括专家介绍和筛选问答。",
    experts: "位专家",
    noResults: "解析后的专家资料会显示在这里。",
    remove: "删除这位专家",
    warning: "需确认",
    sheetPlanner: "Sheet 分组方式",
    sheetModeSingle: "全部专家放在同一个Sheet",
    sheetModePerExpert: "每位专家单独一个Sheet",
    sheetModeCustom: "自由安排Sheet（拖拽）",
    singleSheetName: "Sheet名称",
    customSheetHelp: "打开Sheet管理窗口，把专家拖到想要的Sheet中。",
    perExpertSheetHelp: "系统会使用编号和姓名，为每位专家建立单独的Sheet。",
    openSheetOrganizer: "打开Sheet管理窗口",
    organizerTitle: "自由安排Sheet",
    organizerHelp: "新增或修改Sheet，然后把专家卡片拖到目标Sheet中。",
    newSheetPlaceholder: "新的Sheet名称",
    addSheet: "新增Sheet",
    emptySheet: "把专家拖到这里",
    dragHint: "拖拽移动",
    deleteSheet: "删除空Sheet",
    done: "完成",
    exportTitle: "3. 生成 Excel",
    exportHelp: "保留上传模板的 C–K 列结构，并加入更易读的格式。",
    fileName: "文件名",
    export: "下载 Excel",
    exporting: "正在生成 Excel……",
    exported: "Excel 已生成并开始下载。",
    parsed: "位专家已成功解析，请确认以下内容。",
    parseError: "没有找到以“#1.1 - Name - …”开头的专家资料。",
    exportError: "Excel 生成失败，请重试。",
    exportEmpty: "请先解析至少一位专家。",
    fields: {
      number: "编号",
      name: "姓名",
      company: "企业",
      relevantExperience: "相关经历",
      employmentHistory: "全部工作经历",
      introduction: "专家介绍",
      screening: "筛选问题与回答",
      fee: "金额",
      availability: "可访谈日期与时间",
      sheetName: "输出Sheet名称",
    },
  },
} as const;

function cleanText(value: string) {
  return value
    .replace(/\r\n?/g, "\n")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function splitExpertBlocks(raw: string) {
  const text = cleanText(raw);
  const regex = /#\d+(?:\.\d+)*\s*-\s*/g;
  const starts: number[] = [];
  let match: RegExpExecArray | null;

  while ((match = regex.exec(text)) !== null) starts.push(match.index);
  return starts.map((start, index) =>
    text.slice(start, starts[index + 1] ?? text.length).trim(),
  );
}

type ScreeningBoundary = {
  markerIndex: number;
  contentStart: number;
};

function findRoleDateEnd(source: string) {
  const roleDate = new RegExp(
    `\\((?:\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\s*-\\s*(?:Present|\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\)`,
    "i",
  ).exec(source);
  return roleDate && roleDate.index !== undefined
    ? roleDate.index + roleDate[0].length
    : -1;
}

function findScreeningBoundary(source: string): ScreeningBoundary | null {
  const candidates: ScreeningBoundary[] = [];
  const labeledPatterns = [
    /Screening update(?:\s*\([^)]*\)|\s+\d{1,2}\/\d{1,2}\/\d{2,4})?\s*:?/i,
    /[\[【]\s*Screen(?:ed|ing)\b[^\]】]*[\]】]\s*:?/i,
  ];

  labeledPatterns.forEach((pattern) => {
    const match = pattern.exec(source);
    if (match && match.index !== undefined) {
      candidates.push({
        markerIndex: match.index,
        contentStart: match.index + match[0].length,
      });
    }
  });

  const lineQuestion = /(?:^|\n)\s*(Q(?:\d+)?(?:\s*[).:\-]|\s+))/im.exec(source);
  if (lineQuestion && lineQuestion.index !== undefined) {
    const questionOffset = lineQuestion[0].indexOf(lineQuestion[1]);
    candidates.push({
      markerIndex: lineQuestion.index + questionOffset,
      contentStart: lineQuestion.index + questionOffset,
    });
  }

  const inlineQuestion = /\s(Q(?:\d+)?\s*[).:\-])/i.exec(source);
  if (inlineQuestion && inlineQuestion.index !== undefined) {
    const questionOffset = inlineQuestion[0].indexOf(inlineQuestion[1]);
    candidates.push({
      markerIndex: inlineQuestion.index + questionOffset,
      contentStart: inlineQuestion.index + questionOffset,
    });
  }

  return candidates.sort((a, b) => a.markerIndex - b.markerIndex)[0] ?? null;
}

function matchIndex(source: string, regex: RegExp) {
  const match = regex.exec(source);
  return match ? match.index : -1;
}

function firstPositive(...values: number[]) {
  const available = values.filter((value) => value >= 0);
  return available.length ? Math.min(...available) : -1;
}

function sectionAfter(
  source: string,
  marker: RegExp,
  endingMarkers: RegExp[],
) {
  const match = marker.exec(source);
  if (!match) return "";
  const start = match.index + match[0].length;
  const tail = source.slice(start);
  const ends = endingMarkers.map((regex) => matchIndex(tail, regex));
  const end = firstPositive(...ends);
  return cleanText(end >= 0 ? tail.slice(0, end) : tail);
}

function formatEmploymentHistory(source: string) {
  return cleanText(source).replace(
    new RegExp(`\\s+(?=(?:${MONTHS})\\s+\\d{4}\\s*-\\s*)`, "gi"),
    "\n",
  );
}

function formatAvailability(source: string) {
  return cleanText(source)
    .replace(/\s+(?=Time Zone\s*:)/gi, "\n")
    .replace(new RegExp(`\\s+(?=(?:${WEEKDAYS})\\s+)`, "gi"), "\n")
    .trim();
}

function formatScreening(source: string) {
  return cleanText(source)
    .replace(/\s+(?=(?:Q|A)(?:\d+)?\s*[).:\-])/gi, "\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

function companyKey(value: string) {
  return value
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(
      /\b(corporation|corp|incorporated|inc|company|co|limited|ltd|llc|plc|gmbh|ag|sa)\b/g,
      "",
    )
    .replace(/[^a-z0-9\p{L}]+/gu, "")
    .trim();
}

function parseCompany(headline: string) {
  const dated = new RegExp(
    `\\bat\\s+(.+?)\\s*\\((?:\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\s*-\\s*(?:Present|\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\)`,
    "gi",
  );
  const matches = [...headline.matchAll(dated)];
  if (matches.length) return cleanText(matches[matches.length - 1][1]);

  const fallback = headline.match(/\bat\s+(.+?)(?:\s*\(|$)/i);
  return fallback ? cleanText(fallback[1]) : "";
}

function relatedHistory(employmentHistory: string, company: string) {
  if (!employmentHistory || !company) return "";
  const key = companyKey(company);
  if (!key) return "";
  return employmentHistory
    .split(/\n(?=(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December)\s+\d{4}\s*-)/i)
    .filter((entry) => companyKey(entry).includes(key))
    .join("\n")
    .trim();
}

function calculateWarnings(record: {
  name: string;
  company: string;
  relevantExperience: string;
  employmentHistory: string;
  introduction: string;
  screening: string;
  fee: string;
  availability: string;
}) {
  const warnings: string[] = [];
  if (!record.name) warnings.push("Name");
  if (!record.company) warnings.push("Company");
  if (!record.introduction) warnings.push("Introduction");
  if (!record.screening) warnings.push("Screening Q&A");
  if (!record.employmentHistory) warnings.push("Employment History");
  if (
    record.company &&
    record.employmentHistory &&
    !record.relevantExperience
  )
    warnings.push("Relevant Experience");
  if (!record.fee) warnings.push("Fee");
  if (!record.availability) warnings.push("Availability");
  return warnings;
}

function parseExpert(block: string, index: number): ExpertRecord | null {
  const header = block.match(/^#([\d.]+)\s*-\s*(.*?)\s*-\s*/s);
  if (!header) return null;

  const number = `#${header[1]}`;
  const name = cleanText(header[2]);
  const body = block.slice(header[0].length).trim();

  const firstLineEnd = body.indexOf("\n");
  const roleDateEnd = findRoleDateEnd(body);
  const fallbackIntroduction = matchIndex(
    body,
    /\b(?:This expert|The expert|This specialist|He has|She has)\b/i,
  );
  const headlineEnd =
    firstLineEnd >= 0
      ? firstLineEnd
      : roleDateEnd >= 0
        ? roleDateEnd
        : fallbackIntroduction;
  const headline = cleanText(
    headlineEnd >= 0 ? body.slice(0, headlineEnd) : body,
  );
  const content = cleanText(
    headlineEnd >= 0 ? body.slice(headlineEnd) : "",
  );
  const screeningBoundary = findScreeningBoundary(content);
  const employmentIndex = matchIndex(content, /Employment History\s*:/i);
  const availabilityIndex = matchIndex(content, /Availability\s*:/i);
  const feeIndex = matchIndex(content, /Hourly Fee\s*:/i);
  const profileEnd = firstPositive(
    screeningBoundary?.markerIndex ?? -1,
    employmentIndex,
    availabilityIndex,
    feeIndex,
  );
  const introduction = cleanText(
    profileEnd >= 0 ? content.slice(0, profileEnd) : content,
  );
  const company = parseCompany(headline);

  let screening = "";
  if (screeningBoundary) {
    const screeningTail = content.slice(screeningBoundary.contentStart);
    const screeningEnd = firstPositive(
      matchIndex(screeningTail, /Employment History\s*:/i),
      matchIndex(screeningTail, /Availability\s*:/i),
      matchIndex(screeningTail, /Hourly Fee\s*:/i),
    );
    screening = formatScreening(
      screeningEnd >= 0
        ? screeningTail.slice(0, screeningEnd)
        : screeningTail,
    );
  }
  const employmentHistory = formatEmploymentHistory(
    sectionAfter(body, /Employment History\s*:/i, [
      /Availability\s*:/i,
      /Book Now/i,
      /This specialist is based/i,
      /Hourly Fee\s*:/i,
    ]),
  );
  const availability = formatAvailability(
    sectionAfter(body, /Availability\s*:/i, [
      /Book Now/i,
      /This specialist is based/i,
      /Hourly Fee\s*:/i,
    ]),
  );
  const feeMatch = block.match(/Hourly Fee\s*:\s*([A-Z]{3}\s*[\d,.]+)/i);
  const fee = feeMatch ? `Hourly Fee: ${feeMatch[1].replace(/\s+/g, "")}` : "";
  const relevantExperience = relatedHistory(employmentHistory, company);

  const warnings = calculateWarnings({
    name,
    company,
    relevantExperience,
    employmentHistory,
    introduction,
    screening,
    fee,
    availability,
  });

  return {
    id: `${number}-${index}-${Date.now()}`,
    number,
    name,
    company,
    relevantExperience,
    employmentHistory,
    introduction,
    screening,
    fee,
    availability,
    sheetName: "Expert List",
    warnings,
  };
}

function parseExperts(raw: string) {
  return splitExpertBlocks(raw)
    .map((block, index) => parseExpert(block, index))
    .filter((record): record is ExpertRecord => Boolean(record));
}

function safeFileName(value: string) {
  const clean = value.trim().replace(/[\\/:*?"<>|]+/g, "-");
  if (!clean) return "Expert_List.xlsx";
  return clean.toLowerCase().endsWith(".xlsx") ? clean : `${clean}.xlsx`;
}

function safeSheetName(value: string, fallback = "Expert List") {
  const clean = value
    .trim()
    .replace(/[\\/:*?\[\]]+/g, "-")
    .replace(/^'+|'+$/g, "")
    .trim();
  return (clean || fallback).slice(0, 31);
}

function uniqueSheetName(value: string, used: Set<string>) {
  const base = safeSheetName(value);
  let candidate = base;
  let count = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${count})`;
    candidate = `${base.slice(0, 31 - suffix.length)}${suffix}`;
    count += 1;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function groupRecordsForSheets(
  records: ExpertRecord[],
  sheetMode: SheetMode,
  singleSheetName: string,
) {
  if (sheetMode === "single") {
    return [{ name: singleSheetName || "Expert List", records }];
  }

  if (sheetMode === "perExpert") {
    return records.map((record, index) => ({
      name: `${record.number || `#${index + 1}`} ${record.name || "Expert"}`,
      records: [record],
    }));
  }

  const grouped = new Map<string, ExpertRecord[]>();
  records.forEach((record) => {
    const name = record.sheetName.trim() || "Expert List";
    grouped.set(name, [...(grouped.get(name) ?? []), record]);
  });
  return [...grouped.entries()].map(([name, groupedRecords]) => ({
    name,
    records: groupedRecords,
  }));
}

function estimatedRowHeight(values: string[]) {
  const longest = Math.max(...values.map((value) => value.length), 0);
  return Math.min(300, Math.max(72, 60 + Math.ceil(longest / 150) * 15));
}

export default function Home() {
  const [language, setLanguage] = useState<Language>("ja");
  const [theme, setTheme] = useState<"light" | "dark">("light");
  const [raw, setRaw] = useState("");
  const [records, setRecords] = useState<ExpertRecord[]>([]);
  const [fileName, setFileName] = useState("Expert_List.xlsx");
  const [sheetMode, setSheetMode] = useState<SheetMode>("single");
  const [singleSheetName, setSingleSheetName] = useState("Expert List");
  const [customSheets, setCustomSheets] = useState(["Expert List"]);
  const [newSheetName, setNewSheetName] = useState("");
  const [sheetOrganizerOpen, setSheetOrganizerOpen] = useState(false);
  const [draggedExpertId, setDraggedExpertId] = useState("");
  const [dragOverSheet, setDragOverSheet] = useState("");
  const [message, setMessage] = useState("");
  const [messageType, setMessageType] = useState<"success" | "error" | "">("");
  const [exporting, setExporting] = useState(false);
  const t = translations[language];

  const warningCount = useMemo(
    () => records.reduce((total, record) => total + record.warnings.length, 0),
    [records],
  );

  function changeTheme() {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    document.documentElement.dataset.theme = next;
  }

  function runParser() {
    const parsed = parseExperts(raw);
    setRecords(parsed);
    setCustomSheets(["Expert List"]);
    if (!parsed.length) {
      setMessage(t.parseError);
      setMessageType("error");
      return;
    }
    setMessage(
      language === "en"
        ? `${parsed.length} ${t.parsed}`
        : `${parsed.length}${t.parsed}`,
    );
    setMessageType("success");
    window.setTimeout(
      () => document.getElementById("results")?.scrollIntoView({ behavior: "smooth" }),
      50,
    );
  }

  function clearAll() {
    setRaw("");
    setRecords([]);
    setSheetMode("single");
    setSingleSheetName("Expert List");
    setCustomSheets(["Expert List"]);
    setNewSheetName("");
    setSheetOrganizerOpen(false);
    setMessage("");
    setMessageType("");
  }

  function updateRecord(
    id: string,
    field: keyof Omit<ExpertRecord, "id" | "warnings">,
    value: string,
  ) {
    setRecords((current) =>
      current.map((record) => {
        if (record.id !== id) return record;
        const next = { ...record, [field]: value };
        return { ...next, warnings: calculateWarnings(next) };
      }),
    );
  }

  function changeSheetMode(nextMode: SheetMode) {
    setSheetMode(nextMode);
    if (nextMode === "custom") {
      const assignedSheets = records
        .map((record) => record.sheetName.trim())
        .filter(Boolean);
      setCustomSheets((current) => [
        ...new Set([...current, ...assignedSheets]),
      ]);
      setSheetOrganizerOpen(true);
    }
  }

  function addCustomSheet() {
    const used = new Set(customSheets.map((name) => name.toLowerCase()));
    const candidate = uniqueSheetName(
      newSheetName.trim() || (language === "ja" ? "新しいSheet" : language === "zh" ? "新Sheet" : "New Sheet"),
      used,
    );
    setCustomSheets((current) => [...current, candidate]);
    setNewSheetName("");
  }

  function renameCustomSheet(currentName: string, requestedName: string) {
    const used = new Set(
      customSheets
        .filter((name) => name !== currentName)
        .map((name) => name.toLowerCase()),
    );
    const nextName = uniqueSheetName(requestedName || currentName, used);
    if (nextName === currentName) return;
    setCustomSheets((current) =>
      current.map((name) => (name === currentName ? nextName : name)),
    );
    setRecords((current) =>
      current.map((record) =>
        record.sheetName === currentName
          ? { ...record, sheetName: nextName }
          : record,
      ),
    );
  }

  function deleteCustomSheet(sheetName: string) {
    const hasExperts = records.some((record) => record.sheetName === sheetName);
    if (hasExperts || customSheets.length <= 1) return;
    setCustomSheets((current) => current.filter((name) => name !== sheetName));
  }

  function moveExpertToSheet(expertId: string, sheetName: string) {
    setRecords((current) =>
      current.map((record) =>
        record.id === expertId ? { ...record, sheetName } : record,
      ),
    );
    setDraggedExpertId("");
    setDragOverSheet("");
  }

  async function exportExcel() {
    if (!records.length) {
      setMessage(t.exportEmpty);
      setMessageType("error");
      return;
    }

    setExporting(true);
    setMessage("");
    try {
      const { Workbook } = await import("exceljs");
      const workbook = new Workbook();
      workbook.creator = "Taya Expert List Builder";
      workbook.created = new Date();

      const sheetGroups = groupRecordsForSheets(
        records,
        sheetMode,
        singleSheetName,
      );
      const usedSheetNames = new Set<string>();

      sheetGroups.forEach((group) => {
      const sheetName = uniqueSheetName(group.name, usedSheetNames);
      const sheet = workbook.addWorksheet(sheetName, {
        views: [
          {
            state: "frozen",
            xSplit: 2,
            ySplit: 2,
            showGridLines: false,
          },
        ],
        pageSetup: {
          orientation: "landscape",
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          margins: {
            left: 0.25,
            right: 0.25,
            top: 0.5,
            bottom: 0.5,
            header: 0.2,
            footer: 0.2,
          },
        },
      });

      const widths = [8.63, 8.63, 15.5, 23.13, 21.5, 40.75, 43.38, 43.38, 75.38, 27, 45.13];
      widths.forEach((width, index) => {
        sheet.getColumn(index + 1).width = width;
      });

      sheet.mergeCells("C1:K1");
      const titleCell = sheet.getCell("C1");
      titleCell.value = `${sheetName}  |  ${new Date().toLocaleDateString("ja-JP")}`;
      titleCell.font = {
        name: "Yu Gothic",
        size: 14,
        bold: true,
        color: { argb: "FFF1F5F9" },
      };
      titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0B1E2D" } };
      titleCell.alignment = { vertical: "middle", horizontal: "left" };
      sheet.getRow(1).height = 30;

      const headers = [
        "番号",
        "名前",
        "企業",
        "関連経歴",
        "過去の経歴",
        "紹介",
        "スクリーニング質問に対する回答",
        "金額",
        "インタビュー可能な日付・候補",
      ];
      headers.forEach((header, index) => {
        const cell = sheet.getCell(2, index + 3);
        cell.value = header;
        cell.font = { name: "Yu Gothic", size: 11, bold: true, color: { argb: "FF102A3A" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDBE9F7" } };
        cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
        cell.border = {
          bottom: { style: "medium", color: { argb: "FF1F3A4D" } },
        };
      });
      sheet.getRow(2).height = 34;

      group.records.forEach((record, index) => {
        const rowNumber = index + 3;
        const values = [
          record.number,
          record.name,
          record.company,
          record.relevantExperience,
          record.employmentHistory,
          record.introduction,
          record.screening,
          record.fee,
          record.availability,
        ];
        values.forEach((value, valueIndex) => {
          const cell = sheet.getCell(rowNumber, valueIndex + 3);
          cell.value = value;
          cell.font = { name: "Yu Gothic", size: 10, color: { argb: "FF17212B" } };
          cell.alignment = { vertical: "top", horizontal: "left", wrapText: true };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: index % 2 === 0 ? "FFFFFFFF" : "FFF5F8FB" },
          };
          cell.border = {
            bottom: { style: "thin", color: { argb: "FFD0D9E2" } },
          };
        });
        sheet.getRow(rowNumber).height = estimatedRowHeight(values);
      });

      sheet.autoFilter = {
        from: "C2",
        to: `K${group.records.length + 2}`,
      };
      sheet.headerFooter.oddFooter = "&LGenerated by Taya Expert List Builder&CPage &P / &N";
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = safeFileName(fileName);
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 1500);
      setMessage(t.exported);
      setMessageType("success");
    } catch (error) {
      console.error(error);
      setMessage(t.exportError);
      setMessageType("error");
    } finally {
      setExporting(false);
    }
  }

  const longFields: Array<keyof Pick<
    ExpertRecord,
    | "relevantExperience"
    | "employmentHistory"
    | "introduction"
    | "screening"
    | "availability"
  >> = [
    "relevantExperience",
    "employmentHistory",
    "introduction",
    "screening",
    "availability",
  ];

  return (
    <main className="app-shell">
      <div className="container">
        <header className="topbar">
          <div>
            <div className="eyebrow">TAYA TOOL</div>
            <h1>
              {t.title} <span>{t.version}</span>
            </h1>
            <p className="subtitle">{t.subtitle}</p>
          </div>
          <div className="controls">
            <label className="sr-only" htmlFor="language">
              Language
            </label>
            <select
              id="language"
              value={language}
              onChange={(event) => setLanguage(event.target.value as Language)}
            >
              <option value="ja">日本語</option>
              <option value="en">English</option>
              <option value="zh">中文</option>
            </select>
            <button className="theme-toggle" type="button" onClick={changeTheme} aria-label="Toggle theme">
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </header>

        <div className="privacy-note">
          <span aria-hidden="true">🔒</span>
          {t.privacy}
        </div>

        <section className="card">
          <div className="section-heading">
            <div>
              <h2>{t.inputTitle}</h2>
              <p>{t.inputHelp}</p>
            </div>
          </div>

          <label className="field-label" htmlFor="raw-experts">
            {t.rawLabel}
          </label>
          <textarea
            id="raw-experts"
            className="raw-input"
            value={raw}
            onChange={(event) => setRaw(event.target.value)}
            placeholder={t.rawPlaceholder}
            spellCheck={false}
          />

          <div className="button-row">
            <button className="button button-primary" type="button" onClick={runParser}>
              {t.parse}
            </button>
            <button className="button button-danger" type="button" onClick={clearAll}>
              {t.clear}
            </button>
          </div>

          {message && (
            <div className={`message ${messageType}`} role="status">
              {message}
            </div>
          )}
        </section>

        <section className="card" id="results">
          <div className="section-heading">
            <div>
              <h2>{t.resultTitle}</h2>
              <p>{t.resultHelp}</p>
            </div>
            {records.length > 0 && (
              <div className="stat-group">
                <span className="stat-pill success">
                  {records.length} {t.experts}
                </span>
                {warningCount > 0 && (
                  <span className="stat-pill warning">
                    {warningCount} {t.warning}
                  </span>
                )}
              </div>
            )}
          </div>

          {records.length > 0 && (
            <div className="sheet-planner">
              <label>
                <span>{t.sheetPlanner}</span>
                <select
                  value={sheetMode}
                  onChange={(event) => changeSheetMode(event.target.value as SheetMode)}
                >
                  <option value="single">{t.sheetModeSingle}</option>
                  <option value="perExpert">{t.sheetModePerExpert}</option>
                  <option value="custom">{t.sheetModeCustom}</option>
                </select>
              </label>
              {sheetMode === "single" && (
                <label>
                  <span>{t.singleSheetName}</span>
                  <input
                    value={singleSheetName}
                    onChange={(event) => setSingleSheetName(event.target.value)}
                  />
                </label>
              )}
              {sheetMode === "perExpert" && (
                <p>{t.perExpertSheetHelp}</p>
              )}
              {sheetMode === "custom" && (
                <div className="sheet-organizer-launch">
                  <p>{t.customSheetHelp}</p>
                  <button
                    className="button button-secondary"
                    type="button"
                    onClick={() => setSheetOrganizerOpen(true)}
                  >
                    {t.openSheetOrganizer}
                  </button>
                </div>
              )}
            </div>
          )}

          {!records.length ? (
            <div className="empty-state">
              <div className="empty-icon">XL</div>
              <p>{t.noResults}</p>
            </div>
          ) : (
            <div className="expert-list">
              {records.map((record, index) => (
                <details
                  className="expert-panel"
                  key={record.id}
                  open={records.length <= 3 || index === 0}
                >
                  <summary>
                    <span className="expert-index">{record.number || `#${index + 1}`}</span>
                    <span className="expert-summary">
                      <strong>{record.name || "Unnamed expert"}</strong>
                      <small>
                        {record.company || "Company not detected"}
                        {sheetMode === "custom" && ` · ${record.sheetName || "Expert List"}`}
                      </small>
                    </span>
                    {record.warnings.length > 0 && (
                      <span className="warning-count">{record.warnings.length}</span>
                    )}
                    <span className="chevron" aria-hidden="true">⌄</span>
                  </summary>

                  <div className="expert-content">
                    {record.warnings.length > 0 && (
                      <div className="warning-box">
                        <strong>{t.warning}:</strong> {record.warnings.join(" / ")}
                      </div>
                    )}

                    <div className="short-fields">
                      {(["number", "name", "company", "fee"] as const).map((field) => (
                        <label key={field}>
                          <span>{t.fields[field]}</span>
                          <input
                            value={record[field]}
                            onChange={(event) => updateRecord(record.id, field, event.target.value)}
                          />
                        </label>
                      ))}
                    </div>

                    <div className="long-fields">
                      {longFields.map((field) => (
                        <label key={field}>
                          <span>{t.fields[field]}</span>
                          <textarea
                            className={field === "screening" ? "tall" : ""}
                            value={record[field]}
                            onChange={(event) => updateRecord(record.id, field, event.target.value)}
                          />
                        </label>
                      ))}
                    </div>

                    <button
                      className="text-danger"
                      type="button"
                      onClick={() => setRecords((current) => current.filter((item) => item.id !== record.id))}
                    >
                      {t.remove}
                    </button>
                  </div>
                </details>
              ))}
            </div>
          )}
        </section>

        <section className="card export-card">
          <div className="section-heading">
            <div>
              <h2>{t.exportTitle}</h2>
              <p>{t.exportHelp}</p>
            </div>
            <div className="excel-mark" aria-hidden="true">X</div>
          </div>
          <div className="export-controls">
            <label>
              <span>{t.fileName}</span>
              <input value={fileName} onChange={(event) => setFileName(event.target.value)} />
            </label>
            <button
              className="button button-primary export-button"
              type="button"
              onClick={exportExcel}
              disabled={exporting || !records.length}
            >
              {exporting ? t.exporting : t.export}
            </button>
          </div>
        </section>

        <footer>Taya Tool · Expert information to Excel</footer>
      </div>

      {sheetOrganizerOpen && sheetMode === "custom" && (
        <div
          className="sheet-modal-backdrop"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) setSheetOrganizerOpen(false);
          }}
        >
          <section
            className="sheet-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="sheet-organizer-title"
          >
            <div className="sheet-modal-header">
              <div>
                <span className="eyebrow">TAYA TOOL</span>
                <h2 id="sheet-organizer-title">{t.organizerTitle}</h2>
                <p>{t.organizerHelp}</p>
              </div>
              <button
                className="modal-close"
                type="button"
                onClick={() => setSheetOrganizerOpen(false)}
                aria-label={t.done}
              >
                ×
              </button>
            </div>

            <div className="sheet-add-row">
              <input
                value={newSheetName}
                onChange={(event) => setNewSheetName(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") addCustomSheet();
                }}
                placeholder={t.newSheetPlaceholder}
              />
              <button className="button button-secondary" type="button" onClick={addCustomSheet}>
                ＋ {t.addSheet}
              </button>
            </div>

            <div className="sheet-board">
              {customSheets.map((sheetName) => {
                const sheetExperts = records.filter(
                  (record) => record.sheetName === sheetName,
                );
                return (
                  <div
                    className={`sheet-column ${dragOverSheet === sheetName ? "is-over" : ""}`}
                    key={sheetName}
                    onDragOver={(event) => {
                      event.preventDefault();
                      setDragOverSheet(sheetName);
                    }}
                    onDragLeave={(event) => {
                      if (!event.currentTarget.contains(event.relatedTarget as Node)) {
                        setDragOverSheet("");
                      }
                    }}
                    onDrop={(event) => {
                      event.preventDefault();
                      if (draggedExpertId) moveExpertToSheet(draggedExpertId, sheetName);
                    }}
                  >
                    <div className="sheet-column-header">
                      <input
                        defaultValue={sheetName}
                        aria-label={t.singleSheetName}
                        onBlur={(event) => renameCustomSheet(sheetName, event.target.value)}
                      />
                      <span>{sheetExperts.length}</span>
                    </div>

                    <div className="sheet-drop-zone">
                      {sheetExperts.length ? (
                        sheetExperts.map((record) => (
                          <article
                            className={`drag-expert ${draggedExpertId === record.id ? "is-dragging" : ""}`}
                            draggable
                            key={record.id}
                            onDragStart={() => setDraggedExpertId(record.id)}
                            onDragEnd={() => {
                              setDraggedExpertId("");
                              setDragOverSheet("");
                            }}
                          >
                            <span className="drag-handle" aria-hidden="true">⠿</span>
                            <div>
                              <strong>{record.name || "Unnamed expert"}</strong>
                              <small>{record.number} · {record.company}</small>
                            </div>
                            <em>{t.dragHint}</em>
                          </article>
                        ))
                      ) : (
                        <div className="sheet-empty">{t.emptySheet}</div>
                      )}
                    </div>

                    {customSheets.length > 1 && !sheetExperts.length && (
                      <button
                        className="delete-sheet"
                        type="button"
                        onClick={() => deleteCustomSheet(sheetName)}
                      >
                        {t.deleteSheet}
                      </button>
                    )}
                  </div>
                );
              })}
            </div>

            <div className="sheet-modal-footer">
              <button
                className="button button-primary"
                type="button"
                onClick={() => setSheetOrganizerOpen(false)}
              >
                {t.done}
              </button>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
