"use client";

import { useEffect, useMemo, useState } from "react";

type Language = "ja" | "en" | "zh";
type SheetMode = "single" | "perExpert" | "custom";
type WorkflowMode = "create" | "update";
type UpdateSummaryLanguage = "ja" | "en";
type ToolView = "excel" | "slack" | "navi";
type NaviLanguage = "en" | "ja" | "zh_cn" | "zh_tw" | "mn";
type NaviMode = "PAST" | "CURRENT" | "BOTH";

type NaviHistoryItem = {
  companies: string;
  mode: NaviMode;
  keywords: string[];
  url: string;
  time: string;
};

type ExpertRecord = {
  id: string;
  stableId: string;
  number: string;
  name: string;
  company: string;
  title: string;
  relevantExperience: string;
  employmentHistory: string;
  introduction: string;
  screening: string;
  fee: string;
  availability: string;
  sheetName: string;
  warnings: string[];
};

type SlackQaItem = {
  id: string;
  question: string;
  answer: string;
};

type SlackExpertRecord = ExpertRecord & {
  location: string;
  screeningLabel: string;
  qaItems: SlackQaItem[];
};

const DATA_FIELDS = [
  "number",
  "name",
  "company",
  "title",
  "relevantExperience",
  "employmentHistory",
  "introduction",
  "screening",
  "fee",
  "availability",
] as const;

type DataField = (typeof DATA_FIELDS)[number];
type ComparisonStatus = "new" | "changed" | "unchanged";

type ComparisonChange = {
  field: DataField;
  oldValue: string;
  newValue: string;
  useLatest: boolean;
};

type ComparisonItem = {
  id: string;
  status: ComparisonStatus;
  latest: ExpertRecord;
  existingId: string;
  changes: ComparisonChange[];
  matchedBy: "stable" | "name-company" | "name" | "number" | "new";
};

type ImportedWorkbookSummary = {
  fileName: string;
  sheetNames: string[];
};

type UpdateSummaryStatus =
  | "new"
  | "updated"
  | "unchanged"
  | "retained"
  | "removed";

type UpdateSummaryChange = {
  field: DataField | "sheetName";
  oldValue: string;
  newValue: string;
};

type UpdateSummaryEntry = {
  status: UpdateSummaryStatus;
  number: string;
  name: string;
  company: string;
  changes: UpdateSummaryChange[];
};

const EXCEL_HEADERS = [
  "番号",
  "名前",
  "企業",
  "関連経歴",
  "過去の経歴",
  "紹介",
  "スクリーニング質問に対する回答",
  "金額",
  "インタビュー可能な日付・候補",
] as const;

const TAYA_META_SHEET = "_TAYA_META";
const TAYA_META_MARKER = "TAYA_TOOL_EXPERT_LIST";
const NAVI_SPLIT_REGEX = /[\r\n,，、。:：\/\\;；Ø|◊]+/g;
const NAVI_MAX_GROUPS = 4;
const NAVI_MAX_HISTORY = 10;
const NAVI_HISTORY_KEY = "tayaHistory_v34";
const GLOBAL_LANGUAGE_KEY = "tayaGlobalLanguageV1";
const GLOBAL_LANGUAGES: NaviLanguage[] = ["en", "ja", "zh_cn", "zh_tw", "mn"];
const EXPERT_ACCESS_KEY = "tayaExpertUnlockedV1";
const EXPERT_PASSWORD_HASH =
  "cfe0042d5ff7f0bba1453855ef82ab6074985b68f0a86e2ffbb4313ea52d33ef";

const MONTHS =
  "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December";
const WEEKDAYS =
  "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday";

const translations = {
  ja: {
    title: "Taya Expert List Builder",
    version: "v1.7",
    subtitle: "エキスパート情報を貼り付け、Excel リストをすぐに作成できます。",
    privacy: "入力内容はブラウザ内だけで処理され、サーバーへ送信・保存されません。",
    modeCreate: "新しいExcelを作成",
    modeUpdate: "既存Excelを更新",
    updateUploadTitle: "更新するExcelをアップロード",
    updateUploadHelp: "Taya Toolで作成したExcel（C–K列の専用テンプレート）のみ対応しています。",
    chooseExcel: "Excelを選択",
    readingExcel: "Excelを読み込み中…",
    uploadedExcel: "Excelを読み込みました",
    invalidExcel: "Taya Toolで作成された対応Excelではありません。",
    uploadRequired: "先に更新するExcelをアップロードしてください。",
    latestInfoTitle: "最新のエキスパート情報",
    latestInfoHelp: "最新情報を貼り付けると、旧Excelと比較して変更点を表示します。",
    compareTitle: "変更内容を確認",
    compareHelp: "変更された項目は最新内容が選択されています。必要な項目だけ旧内容に戻せます。",
    statusNew: "新規",
    statusChanged: "変更あり",
    statusUnchanged: "変更なし",
    statusRetained: "旧表のまま保持",
    oldValue: "旧内容",
    newValue: "最新内容",
    useLatest: "最新を使用",
    useOld: "旧内容を使用",
    acceptAll: "すべて最新を使用",
    keepAll: "すべて旧内容を使用",
    newExpertHelp: "新しいエキスパートとして追加されます。",
    unchangedHelp: "旧Excelから変更は検出されませんでした。",
    updateParsed: "最新情報を比較しました。変更内容をご確認ください。",
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
      "必要に応じてExpert Summaryを追加し、テンプレートの C–K 列構成を保って出力します。",
    generateExpertSummary: "Expert Summaryを追加",
    expertSummaryHelp: "全エキスパートの番号、名前、企業、Current Title、所属Sheetを一覧にします。",
    expertSummaryExample: "出力イメージ",
    generateUpdateSummary: "Update Summaryを追加",
    updateSummaryHelp: "最終的に採用した変更だけをまとめたSheetをExcelに追加します。",
    updateSummaryLanguage: "Summary language",
    updateSummaryReady: "Update SummaryをExcelに追加します",
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
      title: "Current Title",
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
    version: "v1.7",
    subtitle: "Paste expert profiles and turn them into a client-ready Excel list.",
    privacy: "Everything is processed in your browser. Nothing is uploaded or stored.",
    modeCreate: "Create a new Excel",
    modeUpdate: "Update an existing Excel",
    updateUploadTitle: "Upload the Excel to update",
    updateUploadHelp: "Only Excel files created by Taya Tool with the dedicated C–K template are supported.",
    chooseExcel: "Choose Excel",
    readingExcel: "Reading Excel…",
    uploadedExcel: "Excel loaded",
    invalidExcel: "This is not a supported Excel created by Taya Tool.",
    uploadRequired: "Upload the Excel you want to update first.",
    latestInfoTitle: "Latest expert information",
    latestInfoHelp: "Paste the latest profiles to compare them with the uploaded Excel.",
    compareTitle: "Review changes",
    compareHelp: "Latest values are selected by default. You can keep any old value field by field.",
    statusNew: "New",
    statusChanged: "Changed",
    statusUnchanged: "Unchanged",
    statusRetained: "Retained from old file",
    oldValue: "Old value",
    newValue: "Latest value",
    useLatest: "Use latest",
    useOld: "Keep old",
    acceptAll: "Use all latest values",
    keepAll: "Keep all old values",
    newExpertHelp: "This expert will be added as a new record.",
    unchangedHelp: "No changes were detected from the uploaded Excel.",
    updateParsed: "Latest information compared. Please review the changes.",
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
      "Optionally adds an Expert Summary while keeping the template’s C–K column structure.",
    generateExpertSummary: "Add Expert Summary",
    expertSummaryHelp: "Lists every expert’s number, name, company, current title, and destination sheet.",
    expertSummaryExample: "Example",
    generateUpdateSummary: "Add Update Summary",
    updateSummaryHelp: "Adds a sheet summarizing only the final changes you accepted.",
    updateSummaryLanguage: "Summary language",
    updateSummaryReady: "The Update Summary will be added to the Excel file",
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
      title: "Current Title",
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
    version: "v1.7",
    subtitle: "粘贴专家资料，一键整理并生成客户用 Excel 名单。",
    privacy: "所有内容仅在浏览器内处理，不会上传或保存到服务器。",
    modeCreate: "创建新的Excel",
    modeUpdate: "更新现有Excel",
    updateUploadTitle: "上传需要更新的Excel",
    updateUploadHelp: "仅支持由Taya Tool生成、使用C–K列专用模板的Excel。",
    chooseExcel: "选择Excel",
    readingExcel: "正在读取Excel……",
    uploadedExcel: "Excel读取成功",
    invalidExcel: "这不是由Taya Tool生成的受支持Excel。",
    uploadRequired: "请先上传需要更新的Excel。",
    latestInfoTitle: "最新专家信息",
    latestInfoHelp: "粘贴最新资料后，系统会与旧Excel比较并显示变化。",
    compareTitle: "确认变化内容",
    compareHelp: "有变化的字段默认使用最新内容，也可以逐项保留旧内容。",
    statusNew: "新增",
    statusChanged: "有变化",
    statusUnchanged: "无变化",
    statusRetained: "保留旧表内容",
    oldValue: "旧内容",
    newValue: "最新内容",
    useLatest: "使用最新内容",
    useOld: "保留旧内容",
    acceptAll: "全部使用最新内容",
    keepAll: "全部保留旧内容",
    newExpertHelp: "将作为新专家添加到Excel。",
    unchangedHelp: "与上传的旧Excel相比没有检测到变化。",
    updateParsed: "最新资料比较完成，请确认变化内容。",
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
    exportHelp: "可按需添加 Expert Summary，并保留模板的 C–K 列结构。",
    generateExpertSummary: "添加 Expert Summary",
    expertSummaryHelp: "汇总所有专家的编号、姓名、企业、Current Title及所属Sheet。",
    expertSummaryExample: "示例",
    generateUpdateSummary: "添加 Update Summary",
    updateSummaryHelp: "在Excel中添加一个Sheet，仅汇总最终采用的变更。",
    updateSummaryLanguage: "Summary语言",
    updateSummaryReady: "Update Summary将加入Excel",
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
      title: "Current Title",
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

const updateSummaryText = {
  ja: {
    sheetName: "更新サマリー",
    title: "Update Summary",
    generatedOn: "作成日",
    sourceFile: "更新元ファイル",
    status: "ステータス",
    number: "番号",
    name: "名前",
    company: "企業",
    updatedItems: "更新項目",
    details: "変更内容",
    statuses: {
      new: "新規追加",
      updated: "更新あり",
      unchanged: "変更なし",
      retained: "旧表より保持",
      removed: "削除",
    },
    fields: {
      number: "番号",
      name: "名前",
      company: "企業",
      title: "Current Title",
      relevantExperience: "関連経歴",
      employmentHistory: "過去の経歴",
      introduction: "紹介",
      screening: "スクリーニング回答",
      fee: "金額",
      availability: "Availability",
      sheetName: "Sheet",
    },
    addedDetail: "エキスパートリストに新規追加",
    unchangedDetail: "最終出力に変更なし",
    retainedDetail: "最新情報には含まれず、旧表の内容を保持",
    removedDetail: "最終出力から削除",
    updatedDetail: "更新",
  },
  en: {
    sheetName: "Update Summary",
    title: "Update Summary",
    generatedOn: "Generated on",
    sourceFile: "Source file",
    status: "Status",
    number: "Number",
    name: "Name",
    company: "Company",
    updatedItems: "Updated items",
    details: "Change details",
    statuses: {
      new: "New expert",
      updated: "Updated",
      unchanged: "Unchanged",
      retained: "Retained from old file",
      removed: "Removed",
    },
    fields: {
      number: "Number",
      name: "Name",
      company: "Company",
      title: "Current Title",
      relevantExperience: "Relevant experience",
      employmentHistory: "Employment history",
      introduction: "Introduction",
      screening: "Screening Q&A",
      fee: "Fee",
      availability: "Availability",
      sheetName: "Sheet",
    },
    addedDetail: "Added to the expert list",
    unchangedDetail: "No change in the final output",
    retainedDetail: "Not included in the latest input; retained from the old file",
    removedDetail: "Removed from the final output",
    updatedDetail: "Updated",
  },
} as const;

const naviTranslations = {
  en: {
    title: "LinkedIn Search Builder",
    version: "Taya Navi v1.2",
    subtitle: "Build a LinkedIn Sales Navigator search URL from company and keyword groups.",
    privacy: "Everything is processed in your browser. Search history stays on this device.",
    company: "Company names",
    companyPlaceholder: "Enter one company per line, or separate with commas…",
    mode: "Company filter",
    past: "Former only (exclude current)",
    current: "Current only",
    both: "Current + former",
    keyword: "Keyword group",
    keywordPlaceholder: "Enter alternative keywords for this OR group…",
    remove: "Remove",
    add: "Add keyword group (AND)",
    generate: "Generate search",
    open: "Open LinkedIn",
    copy: "Copy URL",
    copied: "Copied",
    clear: "Clear all",
    expandedCompanies: "Expanded companies",
    expandedKeywords: "Expanded keywords",
    generatedUrl: "Generated URL",
    noUrl: "Generate a search to see the URL here.",
    history: "Recent searches",
    noHistory: "Your recent searches will appear here.",
    load: "Load",
    delete: "Delete",
    clearHistory: "Clear history",
    confirmClear: "Clear all inputs?",
    confirmHistory: "Clear all search history?",
    companyRequired: "Enter at least one company name.",
  },
  ja: {
    title: "LinkedIn 検索ビルダー",
    version: "Taya Navi v1.2",
    subtitle: "会社名とキーワードグループから、LinkedIn Sales Navigator の検索URLを作成します。",
    privacy: "すべてブラウザ内で処理され、検索履歴はこの端末だけに保存されます。",
    company: "会社名",
    companyPlaceholder: "1行に1社、またはカンマ区切りで入力してください…",
    mode: "会社フィルター",
    past: "元職のみ（現職を除外）",
    current: "現職のみ",
    both: "現職＋元職",
    keyword: "キーワードグループ",
    keywordPlaceholder: "同じORグループにするキーワードを入力してください…",
    remove: "削除",
    add: "キーワードグループ追加（AND）",
    generate: "検索URLを生成",
    open: "LinkedInを開く",
    copy: "URLをコピー",
    copied: "コピーしました",
    clear: "すべてクリア",
    expandedCompanies: "展開された会社名",
    expandedKeywords: "展開されたキーワード",
    generatedUrl: "生成されたURL",
    noUrl: "検索URLを生成すると、ここに表示されます。",
    history: "最近の検索",
    noHistory: "最近の検索がここに表示されます。",
    load: "読込",
    delete: "削除",
    clearHistory: "履歴を削除",
    confirmClear: "入力内容をすべてクリアしますか？",
    confirmHistory: "検索履歴をすべて削除しますか？",
    companyRequired: "会社名を1社以上入力してください。",
  },
  zh_cn: {
    title: "LinkedIn 搜索生成器",
    version: "Taya Navi v1.2",
    subtitle: "根据公司名称和关键词组，生成 LinkedIn Sales Navigator 搜索网址。",
    privacy: "所有内容仅在浏览器内处理，搜索记录只保存在当前设备。",
    company: "公司名称",
    companyPlaceholder: "每行输入一家公司，也可以用逗号分隔……",
    mode: "公司筛选方式",
    past: "仅过去任职（排除现职）",
    current: "仅现职",
    both: "现职＋过去任职",
    keyword: "关键词组",
    keywordPlaceholder: "输入属于同一个 OR 组的关键词……",
    remove: "删除",
    add: "添加关键词组（AND）",
    generate: "生成搜索网址",
    open: "打开 LinkedIn",
    copy: "复制网址",
    copied: "已复制",
    clear: "全部清空",
    expandedCompanies: "展开后的公司",
    expandedKeywords: "展开后的关键词",
    generatedUrl: "生成的网址",
    noUrl: "生成搜索后，网址会显示在这里。",
    history: "最近搜索",
    noHistory: "最近的搜索记录会显示在这里。",
    load: "载入",
    delete: "删除",
    clearHistory: "清空记录",
    confirmClear: "确定清空全部输入内容吗？",
    confirmHistory: "确定清空全部搜索记录吗？",
    companyRequired: "请至少输入一家公司。",
  },
  zh_tw: {
    title: "LinkedIn 搜尋產生器",
    version: "Taya Navi v1.2",
    subtitle: "根據公司名稱和關鍵字群組，產生 LinkedIn Sales Navigator 搜尋網址。",
    privacy: "所有內容僅在瀏覽器內處理，搜尋紀錄只保存在目前裝置。",
    company: "公司名稱",
    companyPlaceholder: "每行輸入一家公司，也可以用逗號分隔……",
    mode: "公司篩選方式",
    past: "僅過去任職（排除現職）",
    current: "僅現職",
    both: "現職＋過去任職",
    keyword: "關鍵字群組",
    keywordPlaceholder: "輸入屬於同一個 OR 群組的關鍵字……",
    remove: "刪除",
    add: "新增關鍵字群組（AND）",
    generate: "產生搜尋網址",
    open: "開啟 LinkedIn",
    copy: "複製網址",
    copied: "已複製",
    clear: "全部清除",
    expandedCompanies: "展開後的公司",
    expandedKeywords: "展開後的關鍵字",
    generatedUrl: "產生的網址",
    noUrl: "產生搜尋後，網址會顯示在這裡。",
    history: "最近搜尋",
    noHistory: "最近的搜尋紀錄會顯示在這裡。",
    load: "載入",
    delete: "刪除",
    clearHistory: "清除紀錄",
    confirmClear: "確定清除全部輸入內容嗎？",
    confirmHistory: "確定清除全部搜尋紀錄嗎？",
    companyRequired: "請至少輸入一家公司。",
  },
  mn: {
    title: "LinkedIn хайлтын үүсгэгч",
    version: "Taya Navi v1.2",
    subtitle: "Компанийн нэр болон түлхүүр үгийн бүлгээс LinkedIn Sales Navigator хайлтын холбоос үүсгэнэ.",
    privacy: "Бүх боловсруулалт хөтөч дээр хийгдэж, хайлтын түүх зөвхөн энэ төхөөрөмжид хадгалагдана.",
    company: "Компанийн нэр",
    companyPlaceholder: "Нэг мөрөнд нэг компани эсвэл таслалаар тусгаарлана уу…",
    mode: "Компанийн шүүлтүүр",
    past: "Зөвхөн өмнөх (одоогийн ажлыг хасна)",
    current: "Зөвхөн одоогийн",
    both: "Одоогийн + өмнөх",
    keyword: "Түлхүүр үгийн бүлэг",
    keywordPlaceholder: "Нэг OR бүлгийн түлхүүр үгсийг оруулна уу…",
    remove: "Устгах",
    add: "Түлхүүр үгийн бүлэг нэмэх (AND)",
    generate: "Хайлт үүсгэх",
    open: "LinkedIn нээх",
    copy: "Холбоос хуулах",
    copied: "Хуулсан",
    clear: "Бүгдийг цэвэрлэх",
    expandedCompanies: "Задалсан компаниуд",
    expandedKeywords: "Задалсан түлхүүр үгс",
    generatedUrl: "Үүсгэсэн холбоос",
    noUrl: "Хайлт үүсгэсний дараа холбоос энд харагдана.",
    history: "Сүүлийн хайлтууд",
    noHistory: "Сүүлийн хайлтууд энд харагдана.",
    load: "Ачаалах",
    delete: "Устгах",
    clearHistory: "Түүх цэвэрлэх",
    confirmClear: "Бүх оруулгыг цэвэрлэх үү?",
    confirmHistory: "Бүх хайлтын түүхийг цэвэрлэх үү?",
    companyRequired: "Дор хаяж нэг компанийн нэр оруулна уу.",
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

function removeAvailabilityPlaceholder(value: string) {
  return cleanText(
    value.replace(
      /(?:This\s+(?:specialist|expert)\s+has\s+not(?:\s+yet)?\s+provided\s+(?:any\s+)?availability\.?|Request\s+Availability)/gi,
      "",
    ),
  );
}

function createStableId() {
  if (typeof crypto !== "undefined" && "randomUUID" in crypto) {
    return crypto.randomUUID();
  }
  return `taya-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function comparableText(value: string) {
  return cleanText(value).toLocaleLowerCase();
}

function personKey(value: string) {
  return value
    .toLocaleLowerCase()
    .normalize("NFKC")
    .replace(/[^a-z0-9\p{L}]+/gu, "")
    .trim();
}

function excelCellText(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return cleanText(String(value));
  }
  if (typeof value === "object") {
    const objectValue = value as {
      richText?: Array<{ text?: string }>;
      result?: unknown;
      text?: unknown;
      hyperlink?: unknown;
    };
    if (Array.isArray(objectValue.richText)) {
      return cleanText(objectValue.richText.map((part) => part.text ?? "").join(""));
    }
    if (objectValue.result !== undefined) return excelCellText(objectValue.result);
    if (objectValue.text !== undefined) return excelCellText(objectValue.text);
  }
  return cleanText(String(value));
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

type ScreeningMarker = {
  type: "question" | "answer";
  markerIndex: number;
  contentStart: number;
};

const SCREENING_MARKER_SOURCE = String.raw`(?:[\[【]\s*(?:Q(?:\d+)?\s*[.。:：)\-]?|A(?:\d+)?\s*[.。:：)\-]?|回答\s*[.。:：\-]?|→)\s*[\]】]|Q(?:\d+)?\s*[.。:：)\-]|A(?:\d+)?\s*[.。:：)\-]|回答\s*[:：]|→)`;

function screeningMarkerRegex(global = false) {
  return new RegExp(`(^|\\s)(${SCREENING_MARKER_SOURCE})`, global ? "gimu" : "imu");
}

function screeningMarkers(source: string) {
  const markers: ScreeningMarker[] = [];
  const regex = screeningMarkerRegex(true);
  let match: RegExpExecArray | null;

  while ((match = regex.exec(source)) !== null) {
    const prefixLength = match[1]?.length ?? 0;
    const marker = match[2];
    const normalized = marker.replace(/[\[\]【】\s.。:：)\-]/g, "");
    markers.push({
      type: /^Q/i.test(normalized) ? "question" : "answer",
      markerIndex: match.index + prefixLength,
      contentStart: match.index + match[0].length,
    });
  }
  return markers;
}

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

  const firstQuestion = screeningMarkers(source).find(
    (marker) => marker.type === "question",
  );
  if (firstQuestion) {
    candidates.push({
      markerIndex: firstQuestion.markerIndex,
      contentStart: firstQuestion.markerIndex,
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
  return removeAvailabilityPlaceholder(source).replace(
    new RegExp(`\\s+(?=(?:${MONTHS})\\s+\\d{4}\\s*-\\s*)`, "gi"),
    "\n",
  );
}

function formatAvailability(source: string) {
  return removeAvailabilityPlaceholder(source)
    .replace(/\s+(?=Time Zone\s*:)/gi, "\n")
    .replace(new RegExp(`\\s+(?=(?:${WEEKDAYS})\\s+)`, "gi"), "\n")
    .trim();
}

function formatScreening(source: string) {
  return removeAvailabilityPlaceholder(source)
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

function parseTitle(headline: string) {
  const datedSuffix = new RegExp(
    `\\s*\\((?:\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\s*-\\s*(?:Present|\\d{2}\\/\\d{4}|(?:${MONTHS})\\s+\\d{4})\\)\\s*$`,
    "i",
  );
  return cleanText(
    headline
      .replace(/^[^\p{L}\p{N}]+/u, "")
      .replace(datedSuffix, ""),
  );
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
  const introduction = removeAvailabilityPlaceholder(
    profileEnd >= 0 ? content.slice(0, profileEnd) : content,
  );
  const company = parseCompany(headline);
  const title = parseTitle(headline);

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
      /This\s+(?:specialist|expert)\s+has\s+not(?:\s+yet)?\s+provided\s+(?:any\s+)?availability/i,
      /Request\s+Availability/i,
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
    stableId: createStableId(),
    number,
    name,
    company,
    title,
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

function extractScreeningLabel(block: string) {
  const labeled = block.match(
    /Screening update(?:\s*\([^)]*\)|\s+\d{1,2}\/\d{1,2}\/\d{2,4})?\s*:?/i,
  );
  if (labeled) return cleanText(labeled[0]).replace(/\s*:?$/, ":");

  const screened = block.match(/[\[【]\s*Screen(?:ed|ing)\b[^\]】]*[\]】]\s*:?/i);
  if (screened) return cleanText(screened[0]);
  return "Screening Questions:";
}

function extractExpertLocation(block: string) {
  const location = block.match(
    /This\s+(?:specialist|expert)\s+is\s+based\s+in\s+(.+?)\.(?=\s|$)/i,
  );
  return location ? cleanText(location[1]) : "";
}

function parseSlackQaItems(source: string) {
  const text = removeAvailabilityPlaceholder(source);
  const markers = screeningMarkers(text);
  const items: SlackQaItem[] = [];

  if (!markers.length && text) {
    return [
      {
        id: `qa-fallback-${Math.random().toString(36).slice(2, 8)}`,
        question: "",
        answer: text,
      },
    ];
  }

  markers.forEach((marker, index) => {
    const nextMarker = markers[index + 1];
    const content = cleanText(
      text.slice(marker.contentStart, nextMarker?.markerIndex ?? text.length),
    );
    if (marker.type === "question") {
      items.push({
        id: `qa-${index}-${Math.random().toString(36).slice(2, 8)}`,
        question: content,
        answer: "",
      });
      return;
    }

    if (!items.length) {
      items.push({
        id: `qa-${index}-${Math.random().toString(36).slice(2, 8)}`,
        question: "",
        answer: content,
      });
      return;
    }
    const latest = items[items.length - 1];
    latest.answer = cleanText([latest.answer, content].filter(Boolean).join(" "));
  });

  return items;
}

function formatSlackScreening(items: SlackQaItem[]) {
  return items
    .map((item, index) => {
      const lines = [];
      if (item.question) lines.push(`Q${index + 1}) ${cleanText(item.question)}`);
      if (item.answer) lines.push(`A${index + 1}) ${cleanText(item.answer)}`);
      return lines.join("\n");
    })
    .filter(Boolean)
    .join("\n\n")
    .trim();
}

function slackHistoryLines(source: string) {
  return formatEmploymentHistory(source)
    .split("\n")
    .map((line) => cleanText(line))
    .filter(Boolean);
}

function slackAvailability(record: SlackExpertRecord) {
  const lines = formatAvailability(record.availability)
    .split("\n")
    .map((line) => cleanText(line))
    .filter(Boolean);
  const timeZone = lines.find((line) => /^Time Zone\s*:/i.test(line)) ?? "";
  return {
    heading: timeZone ? `Availability: ${timeZone}` : "Availability",
    slots: lines.filter((line) => line !== timeZone),
  };
}

function parseSlackExperts(raw: string) {
  return splitExpertBlocks(raw)
    .map((block, index) => {
      const record = parseExpert(block, index);
      if (!record) return null;
      return {
        ...record,
        location: extractExpertLocation(block),
        screeningLabel: extractScreeningLabel(block),
        qaItems: parseSlackQaItems(record.screening),
      };
    })
    .filter((record): record is SlackExpertRecord => Boolean(record));
}

function formatSlackExpert(record: SlackExpertRecord) {
  const parts: string[] = [];
  parts.push(`*${record.number} - ${record.name} - ✅${record.title}*`);

  if (record.introduction) parts.push(record.introduction);

  const screening = formatSlackScreening(record.qaItems);
  if (screening) {
    parts.push(`\`\`\`\n${record.screeningLabel}\n\n${screening}\n\`\`\``);
  }

  const history = slackHistoryLines(record.employmentHistory);
  if (history.length) {
    parts.push(`*Employment History*\n\n${history.map((line) => `• ${line}`).join("\n")}`);
  }

  const availability = slackAvailability(record);
  if (availability.slots.length) {
    parts.push(
      `*${availability.heading}*\n\n${availability.slots
        .map((slot) => `• \`${slot}\``)
        .join("\n")}`,
    );
  }

  if (record.location) {
    parts.push(`This specialist is based in ${record.location}.`);
  }
  if (record.fee) parts.push(`*${record.fee}*`);
  return parts.join("\n\n");
}

function escapeHtml(value: string) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function slackHtmlText(value: string) {
  return escapeHtml(value).replace(/\n/g, "<br>");
}

function formatSlackExpertHtml(record: SlackExpertRecord) {
  const blocks: string[] = [];
  blocks.push(
    `<p><strong>${escapeHtml(`${record.number} - ${record.name} - ✅${record.title}`)}</strong></p>`,
  );
  if (record.introduction) {
    blocks.push(`<p>${slackHtmlText(record.introduction)}</p>`);
  }

  const screening = formatSlackScreening(record.qaItems);
  if (screening) {
    blocks.push(
      `<pre>${escapeHtml(`${record.screeningLabel}\n\n${screening}`)}</pre>`,
    );
  }

  const history = slackHistoryLines(record.employmentHistory);
  if (history.length) {
    blocks.push(
      `<p><strong>Employment History</strong></p><ul>${history
        .map((line) => `<li>${escapeHtml(line)}</li>`)
        .join("")}</ul>`,
    );
  }

  const availability = slackAvailability(record);
  if (availability.slots.length) {
    blocks.push(
      `<p><strong>${escapeHtml(availability.heading)}</strong></p><ul>${availability.slots
        .map((slot) => `<li><code>${escapeHtml(slot)}</code></li>`)
        .join("")}</ul>`,
    );
  }

  if (record.location) {
    blocks.push(`<p>${escapeHtml(`This specialist is based in ${record.location}.`)}</p>`);
  }
  if (record.fee) blocks.push(`<p><strong>${escapeHtml(record.fee)}</strong></p>`);
  return `<div>${blocks.join("")}</div>`;
}

async function writeSlackClipboard(plainText: string, html: string) {
  try {
    if (navigator.clipboard.write && typeof ClipboardItem !== "undefined") {
      await navigator.clipboard.write([
        new ClipboardItem({
          "text/plain": new Blob([plainText], { type: "text/plain" }),
          "text/html": new Blob([html], { type: "text/html" }),
        }),
      ]);
      return;
    }
  } catch {
    // Fall back to plain Slack markup when rich clipboard data is unavailable.
  }
  await navigator.clipboard.writeText(plainText);
}

const slackTranslations = {
  en: {
    title: "Slack Expert Formatter",
    version: "v1.1",
    subtitle: "Turn multiple expert profiles into clean, copy-ready Slack posts.",
    privacy: "Everything is processed in your browser. Nothing is uploaded or stored.",
    inputTitle: "1. Paste expert profiles",
    inputHelp: "Paste one or many profiles beginning with #1.1 - Name - …",
    label: "Expert information",
    placeholder: "Paste expert profiles here…",
    generate: "Create Slack posts",
    clear: "Clear all",
    results: "2. Copy to Slack",
    resultsHelp: "Each expert is formatted separately. Copy one post or all posts at once.",
    empty: "Your Slack-ready expert posts will appear here.",
    copy: "Copy for Slack",
    copied: "Copied",
    copyAll: "Copy all experts",
    copiedAll: "All experts copied",
    found: "experts formatted",
    parseError: "No expert profile beginning with #1.1 - Name - … was found.",
    employment: "Employment History",
    edit: "Edit",
    doneEditing: "Done editing",
    editHelp: "Review the detected fields and correct any question or answer before copying.",
    number: "Number",
    name: "Name",
    titleField: "Title",
    introduction: "Expert introduction",
    screeningLabel: "Screening heading",
    question: "Question",
    answer: "Answer",
    addQa: "Add Q&A",
    removeQa: "Remove",
    history: "Employment History",
    availability: "Availability",
    location: "Location",
    fee: "Fee",
  },
  ja: {
    title: "Slack Expert Formatter",
    version: "v1.1",
    subtitle: "複数のエキスパート情報を、Slackに貼り付けやすい形式へ整えます。",
    privacy: "入力内容はブラウザ内だけで処理され、アップロードや保存はされません。",
    inputTitle: "1. エキスパート情報を貼り付け",
    inputHelp: "#1.1 - Name - … で始まる情報を、複数名まとめて貼り付けられます。",
    label: "エキスパート情報",
    placeholder: "ここにエキスパート情報を貼り付けてください…",
    generate: "Slack用に整形",
    clear: "すべてクリア",
    results: "2. Slackへコピー",
    resultsHelp: "エキスパートごとに個別コピー、または全員をまとめてコピーできます。",
    empty: "整形したSlack投稿がここに表示されます。",
    copy: "Slack用にコピー",
    copied: "コピーしました",
    copyAll: "全員をコピー",
    copiedAll: "全員をコピーしました",
    found: "名を整形",
    parseError: "#1.1 - Name - … で始まるエキスパート情報が見つかりませんでした。",
    employment: "Employment History",
    edit: "編集",
    doneEditing: "編集を完了",
    editHelp: "認識結果を確認し、コピー前に質問・回答や各項目を修正できます。",
    number: "番号",
    name: "名前",
    titleField: "タイトル",
    introduction: "エキスパート紹介",
    screeningLabel: "Screening見出し",
    question: "質問",
    answer: "回答",
    addQa: "Q&Aを追加",
    removeQa: "削除",
    history: "Employment History",
    availability: "Availability",
    location: "所在地",
    fee: "Fee",
  },
  zh: {
    title: "Slack Expert Formatter",
    version: "v1.1",
    subtitle: "将多位专家信息整理成可直接复制到 Slack 的格式。",
    privacy: "所有内容只在浏览器中处理，不会上传或保存。",
    inputTitle: "1. 粘贴专家信息",
    inputHelp: "可以一次粘贴多位以 #1.1 - Name - … 开头的专家信息。",
    label: "专家信息",
    placeholder: "在这里粘贴专家信息…",
    generate: "生成 Slack 内容",
    clear: "全部清除",
    results: "2. 复制到 Slack",
    resultsHelp: "每位专家独立生成，可以单独复制，也可以一次复制全部。",
    empty: "生成后的 Slack 内容会显示在这里。",
    copy: "复制到 Slack",
    copied: "已复制",
    copyAll: "复制全部专家",
    copiedAll: "已复制全部专家",
    found: "位专家已生成",
    parseError: "没有找到以 #1.1 - Name - … 开头的专家信息。",
    employment: "Employment History",
    edit: "编辑",
    doneEditing: "完成编辑",
    editHelp: "请确认识别结果，复制前可以修改问题、答案及其他字段。",
    number: "编号",
    name: "姓名",
    titleField: "Title",
    introduction: "专家介绍",
    screeningLabel: "Screening 标题",
    question: "问题",
    answer: "答案",
    addQa: "添加 Q&A",
    removeQa: "删除",
    history: "Employment History",
    availability: "Availability",
    location: "所在地",
    fee: "Fee",
  },
} as const;

function findExistingMatch(
  latest: ExpertRecord,
  existing: ExpertRecord[],
  usedIds: Set<string>,
) {
  const available = existing.filter((record) => !usedIds.has(record.id));
  const latestName = personKey(latest.name);
  const latestCompany = companyKey(latest.company);
  const latestNumber = comparableText(latest.number);

  const exact = available.find(
    (record) =>
      latestName &&
      latestCompany &&
      personKey(record.name) === latestName &&
      companyKey(record.company) === latestCompany,
  );
  if (exact) return { record: exact, matchedBy: "name-company" as const };

  const sameName = available.filter(
    (record) => latestName && personKey(record.name) === latestName,
  );
  if (sameName.length === 1) {
    return { record: sameName[0], matchedBy: "name" as const };
  }

  const sameNumber = available.filter(
    (record) => latestNumber && comparableText(record.number) === latestNumber,
  );
  if (sameNumber.length === 1) {
    return { record: sameNumber[0], matchedBy: "number" as const };
  }

  return null;
}

function compareWithExisting(
  latestRecords: ExpertRecord[],
  existingRecords: ExpertRecord[],
  defaultSheetName: string,
) {
  const usedExistingIds = new Set<string>();
  const mergedRecords = existingRecords.map((record) => ({ ...record }));
  const comparisons: ComparisonItem[] = [];

  latestRecords.forEach((latest) => {
    const match = findExistingMatch(latest, existingRecords, usedExistingIds);
    if (!match) {
      const newRecord = {
        ...latest,
        sheetName: defaultSheetName || "Expert List",
      };
      newRecord.warnings = calculateWarnings(newRecord);
      mergedRecords.push(newRecord);
      comparisons.push({
        id: `comparison-${latest.id}`,
        status: "new",
        latest: newRecord,
        existingId: "",
        changes: [],
        matchedBy: "new",
      });
      return;
    }

    usedExistingIds.add(match.record.id);
    const changes: ComparisonChange[] = DATA_FIELDS.flatMap((field) => {
      const oldValue = match.record[field];
      const newValue = latest[field];
      if (!newValue || comparableText(oldValue) === comparableText(newValue)) return [];
      return [{ field, oldValue, newValue, useLatest: true }];
    });

    const mergedIndex = mergedRecords.findIndex(
      (record) => record.id === match.record.id,
    );
    if (mergedIndex >= 0 && changes.length) {
      const merged = { ...mergedRecords[mergedIndex] };
      changes.forEach((change) => {
        merged[change.field] = change.newValue;
      });
      merged.warnings = calculateWarnings(merged);
      mergedRecords[mergedIndex] = merged;
    }

    comparisons.push({
      id: `comparison-${latest.id}`,
      status: changes.length ? "changed" : "unchanged",
      latest: { ...latest, sheetName: match.record.sheetName },
      existingId: match.record.id,
      changes,
      matchedBy: match.matchedBy,
    });
  });

  return {
    comparisons,
    mergedRecords,
    retainedCount: existingRecords.length - usedExistingIds.size,
  };
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

function finalRecordChanges(
  oldRecord: ExpertRecord,
  finalRecord: ExpertRecord,
): UpdateSummaryChange[] {
  const fields = [...DATA_FIELDS, "sheetName"] as const;
  return fields.flatMap((field) => {
    const oldValue = oldRecord[field];
    const newValue = finalRecord[field];
    if (comparableText(oldValue) === comparableText(newValue)) return [];
    return [{ field, oldValue, newValue }];
  });
}

function buildUpdateSummaryEntries(
  records: ExpertRecord[],
  existingRecords: ExpertRecord[],
  comparisonItems: ComparisonItem[],
) {
  const entries: UpdateSummaryEntry[] = [];
  const finalById = new Map(records.map((record) => [record.id, record]));
  const finalByStableId = new Map(
    records.map((record) => [record.stableId, record]),
  );
  const existingById = new Map(
    existingRecords.map((record) => [record.id, record]),
  );
  const matchedExistingIds = new Set<string>();

  comparisonItems.forEach((item) => {
    if (!item.existingId) {
      const finalRecord =
        finalById.get(item.latest.id) ??
        finalByStableId.get(item.latest.stableId);
      if (!finalRecord) return;
      entries.push({
        status: "new",
        number: finalRecord.number,
        name: finalRecord.name,
        company: finalRecord.company,
        changes: [],
      });
      return;
    }

    matchedExistingIds.add(item.existingId);
    const oldRecord = existingById.get(item.existingId);
    if (!oldRecord) return;
    const finalRecord = finalById.get(item.existingId);
    if (!finalRecord) {
      entries.push({
        status: "removed",
        number: oldRecord.number,
        name: oldRecord.name,
        company: oldRecord.company,
        changes: [],
      });
      return;
    }

    const changes = finalRecordChanges(oldRecord, finalRecord);
    entries.push({
      status: changes.length ? "updated" : "unchanged",
      number: finalRecord.number,
      name: finalRecord.name,
      company: finalRecord.company,
      changes,
    });
  });

  existingRecords.forEach((oldRecord) => {
    if (matchedExistingIds.has(oldRecord.id)) return;
    const finalRecord = finalById.get(oldRecord.id);
    if (!finalRecord) {
      entries.push({
        status: "removed",
        number: oldRecord.number,
        name: oldRecord.name,
        company: oldRecord.company,
        changes: [],
      });
      return;
    }
    const changes = finalRecordChanges(oldRecord, finalRecord);
    entries.push({
      status: changes.length ? "updated" : "retained",
      number: finalRecord.number,
      name: finalRecord.name,
      company: finalRecord.company,
      changes,
    });
  });

  return entries;
}

function updateSummaryDetails(
  entry: UpdateSummaryEntry,
  language: UpdateSummaryLanguage,
) {
  const text = updateSummaryText[language];
  if (entry.status === "new") return text.addedDetail;
  if (entry.status === "unchanged") return text.unchangedDetail;
  if (entry.status === "retained") return text.retainedDetail;
  if (entry.status === "removed") return text.removedDetail;

  const showValues = new Set<DataField | "sheetName">([
    "number",
    "name",
    "company",
    "title",
    "fee",
    "sheetName",
  ]);
  return entry.changes
    .map((change) => {
      const label = text.fields[change.field];
      if (!showValues.has(change.field)) return `${label}: ${text.updatedDetail}`;
      return `${label}: ${change.oldValue || "—"} → ${change.newValue || "—"}`;
    })
    .join("\n");
}

function naviUniq(values: string[]) {
  return [...new Set(values.map((value) => value.trim()).filter(Boolean))];
}

function encodeNaviKeyword(value: string) {
  const once = encodeURIComponent(value)
    .replace(/\(/g, "%28")
    .replace(/\)/g, "%29");
  return once.replace(/%/g, "%25");
}

function encodeNaviCompany(value: string) {
  return encodeURIComponent(value).replace(/%/g, "%25");
}

function formatNaviKeywordGroup(value: string) {
  const keywords = naviUniq((value || "").split(NAVI_SPLIT_REGEX));
  return keywords
    .map((keyword) => (/\p{Script=Han}/u.test(keyword) ? `“${keyword}”` : `"${keyword}"`))
    .join(" OR ");
}

function buildNaviFilters(names: string[], mode: NaviMode) {
  const values = (selectionType: "INCLUDED" | "EXCLUDED") =>
    names
      .map(
        (name) =>
          `(text%3A${encodeNaviCompany(name)}%2CselectionType%3A${selectionType})`,
      )
      .join("%2C");

  if (mode === "PAST") {
    return `(type%3APAST_COMPANY%2Cvalues%3AList(${values("INCLUDED")}))%2C(type%3ACURRENT_COMPANY%2Cvalues%3AList(${values("EXCLUDED")}))`;
  }
  if (mode === "CURRENT") {
    return `(type%3ACURRENT_COMPANY%2Cvalues%3AList(${values("INCLUDED")}))`;
  }
  return `(type%3APAST_COMPANY%2Cvalues%3AList(${values("INCLUDED")}))%2C(type%3ACURRENT_COMPANY%2Cvalues%3AList(${values("INCLUDED")}))`;
}

function createNaviGroup(value = "") {
  return {
    id: `navi-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    value,
  };
}

async function hashPassword(value: string) {
  const data = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest("SHA-256", data);
  return Array.from(new Uint8Array(hash))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function ToolSwitcher({
  active,
  onSelect,
}: {
  active: ToolView;
  onSelect: (tool: ToolView) => void;
}) {
  return (
    <nav className="tool-switcher" aria-label="Taya Tool applications">
      <button
        className={active === "excel" ? "is-active" : ""}
        type="button"
        onClick={() => onSelect("excel")}
      >
        <span className="tool-switcher-icon excel">XL</span>
        <span>
          <strong>Expert Excel</strong>
          <small>Create & update lists</small>
        </span>
      </button>
      <button
        className={active === "slack" ? "is-active" : ""}
        type="button"
        onClick={() => onSelect("slack")}
      >
        <span className="tool-switcher-icon slack">SL</span>
        <span>
          <strong>Slack Formatter</strong>
          <small>Copy-ready expert posts</small>
        </span>
      </button>
      <button
        className={active === "navi" ? "is-active" : ""}
        type="button"
        onClick={() => onSelect("navi")}
      >
        <span className="tool-switcher-icon search">⌕</span>
        <span>
          <strong>LinkedIn Search</strong>
          <small>Taya Navi</small>
        </span>
      </button>
    </nav>
  );
}

function ExpertPasswordGate({
  active,
  theme,
  password,
  error,
  checking,
  onPasswordChange,
  onSubmit,
  onToggleTheme,
  onSelectTool,
}: {
  active: Exclude<ToolView, "navi">;
  theme: "light" | "dark";
  password: string;
  error: string;
  checking: boolean;
  onPasswordChange: (value: string) => void;
  onSubmit: (event: React.FormEvent<HTMLFormElement>) => void;
  onToggleTheme: () => void;
  onSelectTool: (tool: ToolView) => void;
}) {
  const [showPassword, setShowPassword] = useState(false);

  return (
    <main className="app-shell access-shell">
      <div className="container access-container">
        <header className="topbar">
          <div>
            <div className="eyebrow">TAYA TOOL</div>
            <h1>
              Expert Tools <span>Testing access</span>
            </h1>
            <p className="subtitle">
              Expert Excel and Slack Formatter are currently limited to approved testers.
            </p>
          </div>
          <div className="controls">
            <button
              className="theme-toggle"
              type="button"
              onClick={onToggleTheme}
              aria-label="Toggle theme"
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </header>

        <ToolSwitcher active={active} onSelect={onSelectTool} />

        <section className="access-card">
          <div className="access-lock" aria-hidden="true">
            <span />
          </div>
          <div className="testing-badge">TESTING</div>
          <h2>Testing in Progress</h2>
          <p className="access-lead">Enter the tester password to continue.</p>
          <form className="access-form" onSubmit={onSubmit}>
            <label htmlFor="tester-password">Tester password</label>
            <div className="password-field">
              <input
                id="tester-password"
                type={showPassword ? "text" : "password"}
                value={password}
                onChange={(event) => onPasswordChange(event.target.value)}
                placeholder="Enter password"
                autoComplete="current-password"
                autoFocus
              />
              <button
                type="button"
                onClick={() => setShowPassword((current) => !current)}
                aria-label={showPassword ? "Hide password" : "Show password"}
              >
                {showPassword ? "Hide" : "Show"}
              </button>
            </div>
            {error && <div className="access-error" role="alert">{error}</div>}
            <button
              className="button button-primary access-submit"
              type="submit"
              disabled={checking || !password}
            >
              {checking ? "Checking…" : "Enter Expert Tools"}
            </button>
          </form>
          <p className="access-note">
            Expert data is processed only in this browser. Nothing is uploaded or stored on a server.
          </p>
        </section>

        <footer>Taya Tool · Tester access</footer>
      </div>
    </main>
  );
}

function TayaNaviPanel({
  theme,
  language,
  onToggleTheme,
  onLanguageChange,
  onSelectTool,
}: {
  theme: "light" | "dark";
  language: NaviLanguage;
  onToggleTheme: () => void;
  onLanguageChange: (language: NaviLanguage) => void;
  onSelectTool: (tool: ToolView) => void;
}) {
  const [companies, setCompanies] = useState("");
  const [mode, setMode] = useState<NaviMode>("PAST");
  const [keywordGroups, setKeywordGroups] = useState(() => [createNaviGroup()]);
  const [generatedURL, setGeneratedURL] = useState("");
  const [history, setHistory] = useState<NaviHistoryItem[]>([]);
  const [notice, setNotice] = useState("");
  const [noticeType, setNoticeType] = useState<"success" | "error" | "">("");
  const t = naviTranslations[language];

  useEffect(() => {
    let saved: NaviHistoryItem[] = [];
    try {
      const parsed = JSON.parse(
        window.localStorage.getItem(NAVI_HISTORY_KEY) || "[]",
      ) as NaviHistoryItem[];
      saved = Array.isArray(parsed) ? parsed.slice(0, NAVI_MAX_HISTORY) : [];
    } catch {
      saved = [];
    }
    const timer = window.setTimeout(() => setHistory(saved), 0);
    return () => window.clearTimeout(timer);
  }, []);

  const expandedCompanies = useMemo(
    () => naviUniq(companies.split(NAVI_SPLIT_REGEX)),
    [companies],
  );
  const expandedKeywords = useMemo(
    () =>
      keywordGroups
        .map((group) => formatNaviKeywordGroup(group.value))
        .filter(Boolean)
        .map((group) => `(${group})`)
        .join(" AND "),
    [keywordGroups],
  );

  function persistHistory(nextHistory: NaviHistoryItem[]) {
    setHistory(nextHistory);
    window.localStorage.setItem(NAVI_HISTORY_KEY, JSON.stringify(nextHistory));
  }

  function generateSearch() {
    if (!expandedCompanies.length) {
      setNotice(t.companyRequired);
      setNoticeType("error");
      return;
    }
    const filters = buildNaviFilters(expandedCompanies, mode);
    const keywordBlock = expandedKeywords
      ? `%2Ckeywords%3A${encodeNaviKeyword(expandedKeywords)}`
      : "";
    const query = `(spellCorrectionEnabled%3Atrue%2CrecentSearchParam%3A(doLogHistory%3Atrue)%2Cfilters%3AList(${filters})${keywordBlock})`;
    const url = `https://www.linkedin.com/sales/search/people?query=${query}&viewAllFilters=true`;
    const item: NaviHistoryItem = {
      companies,
      mode,
      keywords: keywordGroups.map((group) => group.value),
      url,
      time: new Date().toLocaleString(),
    };
    setGeneratedURL(url);
    persistHistory([item, ...history].slice(0, NAVI_MAX_HISTORY));
    setNotice("");
    setNoticeType("");
  }

  async function copySearchURL() {
    if (!generatedURL) return;
    await navigator.clipboard.writeText(generatedURL);
    setNotice(t.copied);
    setNoticeType("success");
  }

  function clearInputs() {
    if (!window.confirm(t.confirmClear)) return;
    setCompanies("");
    setMode("PAST");
    setKeywordGroups([createNaviGroup()]);
    setGeneratedURL("");
    setNotice("");
    setNoticeType("");
  }

  function loadHistoryItem(item: NaviHistoryItem) {
    setCompanies(item.companies || "");
    setMode(item.mode || "PAST");
    setKeywordGroups(
      item.keywords?.length
        ? item.keywords.slice(0, NAVI_MAX_GROUPS).map((value) => createNaviGroup(value))
        : [createNaviGroup()],
    );
    setGeneratedURL(item.url || "");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function clearSearchHistory() {
    if (!window.confirm(t.confirmHistory)) return;
    window.localStorage.removeItem(NAVI_HISTORY_KEY);
    setHistory([]);
  }

  return (
    <main className="app-shell navi-shell">
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
            <label className="sr-only" htmlFor="navi-language">
              Language
            </label>
            <select
              id="navi-language"
              value={language}
              onChange={(event) =>
                onLanguageChange(event.target.value as NaviLanguage)
              }
            >
              <option value="en">English</option>
              <option value="ja">日本語</option>
              <option value="zh_cn">中文（简体）</option>
              <option value="zh_tw">中文（繁體）</option>
              <option value="mn">Монгол</option>
            </select>
            <button
              className="theme-toggle"
              type="button"
              onClick={onToggleTheme}
              aria-label="Toggle theme"
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </header>

        <ToolSwitcher active="navi" onSelect={onSelectTool} />

        <div className="privacy-note">
          <span aria-hidden="true">🔒</span>
          {t.privacy}
        </div>

        <section className="card navi-builder-card">
          <div className="navi-grid">
            <label className="navi-field navi-companies">
              <span>{t.company}</span>
              <textarea
                value={companies}
                onChange={(event) => setCompanies(event.target.value)}
                placeholder={t.companyPlaceholder}
                spellCheck={false}
              />
            </label>

            <label className="navi-field navi-mode">
              <span>{t.mode}</span>
              <select
                value={mode}
                onChange={(event) => setMode(event.target.value as NaviMode)}
              >
                <option value="PAST">{t.past}</option>
                <option value="CURRENT">{t.current}</option>
                <option value="BOTH">{t.both}</option>
              </select>
            </label>
          </div>

          <div className="keyword-groups">
            {keywordGroups.map((group, index) => (
              <label className="navi-field keyword-group" key={group.id}>
                <span>{t.keyword} {index + 1}</span>
                {keywordGroups.length > 1 && (
                  <button
                    className="keyword-remove"
                    type="button"
                    onClick={() =>
                      setKeywordGroups((current) =>
                        current.filter((candidate) => candidate.id !== group.id),
                      )
                    }
                  >
                    {t.remove}
                  </button>
                )}
                <textarea
                  value={group.value}
                  onChange={(event) =>
                    setKeywordGroups((current) =>
                      current.map((candidate) =>
                        candidate.id === group.id
                          ? { ...candidate, value: event.target.value }
                          : candidate,
                      ),
                    )
                  }
                  placeholder={t.keywordPlaceholder}
                  spellCheck={false}
                />
              </label>
            ))}
          </div>

          <div className="navi-add-row">
            <button
              className="button button-muted"
              type="button"
              disabled={keywordGroups.length >= NAVI_MAX_GROUPS}
              onClick={() =>
                setKeywordGroups((current) => [...current, createNaviGroup()])
              }
            >
              ＋ {t.add}
            </button>
            <span>{keywordGroups.length} / {NAVI_MAX_GROUPS}</span>
          </div>

          <div className="button-row navi-actions">
            <button className="button button-primary" type="button" onClick={generateSearch}>
              {t.generate}
            </button>
            <button
              className="button button-secondary"
              type="button"
              disabled={!generatedURL}
              onClick={() => window.open(generatedURL, "_blank", "noopener,noreferrer")}
            >
              {t.open}
            </button>
            <button
              className="button button-muted"
              type="button"
              disabled={!generatedURL}
              onClick={copySearchURL}
            >
              {t.copy}
            </button>
            <button className="button button-danger" type="button" onClick={clearInputs}>
              {t.clear}
            </button>
          </div>

          {notice && (
            <div className={`message ${noticeType}`} role="status">
              {notice}
            </div>
          )}
        </section>

        <section className="card navi-output-card">
          <div className="navi-output-grid">
            <div>
              <h2>{t.expandedCompanies}</h2>
              <pre>{expandedCompanies.join("\n") || "—"}</pre>
            </div>
            <div>
              <h2>{t.expandedKeywords}</h2>
              <pre>{expandedKeywords || "—"}</pre>
            </div>
          </div>
          <div className="generated-url-block">
            <h2>{t.generatedUrl}</h2>
            <code>{generatedURL || t.noUrl}</code>
          </div>
        </section>

        <section className="card navi-history-card">
          <div className="section-heading">
            <div>
              <h2>{t.history}</h2>
              {!history.length && <p>{t.noHistory}</p>}
            </div>
            {history.length > 0 && (
              <button className="text-danger" type="button" onClick={clearSearchHistory}>
                {t.clearHistory}
              </button>
            )}
          </div>
          <div className="navi-history-list">
            {history.map((item, index) => (
              <article className="navi-history-item" key={`${item.time}-${index}`}>
                <div>
                  <strong>{item.time}</strong>
                  <p>{naviUniq((item.companies || "").split(NAVI_SPLIT_REGEX)).join(" · ")}</p>
                </div>
                <div className="history-actions">
                  <button className="button button-muted" type="button" onClick={() => loadHistoryItem(item)}>
                    {t.load}
                  </button>
                  <button
                    className="text-danger"
                    type="button"
                    onClick={() =>
                      persistHistory(history.filter((_, itemIndex) => itemIndex !== index))
                    }
                  >
                    {t.delete}
                  </button>
                </div>
              </article>
            ))}
          </div>
        </section>

        <footer>Taya Tool · LinkedIn search & Expert Excel</footer>
      </div>
    </main>
  );
}

function SlackFormatterPanel({
  theme,
  language,
  onToggleTheme,
  onLanguageChange,
  onSelectTool,
  onLock,
}: {
  theme: "light" | "dark";
  language: NaviLanguage;
  onToggleTheme: () => void;
  onLanguageChange: (language: NaviLanguage) => void;
  onSelectTool: (tool: ToolView) => void;
  onLock: () => void;
}) {
  const uiLanguage: Language =
    language === "ja"
      ? "ja"
      : language === "zh_cn" || language === "zh_tw"
        ? "zh"
        : "en";
  const t = slackTranslations[uiLanguage];
  const [rawProfiles, setRawProfiles] = useState("");
  const [experts, setExperts] = useState<SlackExpertRecord[]>([]);
  const [notice, setNotice] = useState("");
  const [noticeType, setNoticeType] = useState<"success" | "error" | "">("");
  const [copiedId, setCopiedId] = useState("");
  const [editingId, setEditingId] = useState("");

  function generatePosts() {
    const parsed = parseSlackExperts(rawProfiles);
    if (!parsed.length) {
      setExperts([]);
      setNotice(t.parseError);
      setNoticeType("error");
      return;
    }
    setExperts(parsed);
    setNotice(`${parsed.length} ${t.found}`);
    setNoticeType("success");
    window.setTimeout(() => {
      document.getElementById("slack-results")?.scrollIntoView({ behavior: "smooth" });
    }, 0);
  }

  function clearPosts() {
    setRawProfiles("");
    setExperts([]);
    setNotice("");
    setNoticeType("");
    setCopiedId("");
    setEditingId("");
  }

  async function copyExpert(expert: SlackExpertRecord) {
    await writeSlackClipboard(
      formatSlackExpert(expert),
      formatSlackExpertHtml(expert),
    );
    setCopiedId(expert.id);
    window.setTimeout(() => setCopiedId(""), 1800);
  }

  async function copyAllExperts() {
    const combined = experts
      .map((expert) => formatSlackExpert(expert))
      .join("\n\n────────────────────\n\n");
    const combinedHtml = experts
      .map((expert) => formatSlackExpertHtml(expert))
      .join('<hr style="margin:20px 0;border:0;border-top:1px solid #cccccc">');
    await writeSlackClipboard(combined, `<div>${combinedHtml}</div>`);
    setCopiedId("all");
    window.setTimeout(() => setCopiedId(""), 1800);
  }

  function updateExpert(
    expertId: string,
    field: keyof Omit<SlackExpertRecord, "qaItems" | "warnings">,
    value: string,
  ) {
    setExperts((current) =>
      current.map((expert) =>
        expert.id === expertId ? { ...expert, [field]: value } : expert,
      ),
    );
  }

  function updateQa(
    expertId: string,
    qaId: string,
    field: "question" | "answer",
    value: string,
  ) {
    setExperts((current) =>
      current.map((expert) =>
        expert.id === expertId
          ? {
              ...expert,
              qaItems: expert.qaItems.map((item) =>
                item.id === qaId ? { ...item, [field]: value } : item,
              ),
            }
          : expert,
      ),
    );
  }

  function addQa(expertId: string) {
    setExperts((current) =>
      current.map((expert) =>
        expert.id === expertId
          ? {
              ...expert,
              qaItems: [
                ...expert.qaItems,
                { id: `qa-${Date.now()}`, question: "", answer: "" },
              ],
            }
          : expert,
      ),
    );
  }

  function removeQa(expertId: string, qaId: string) {
    setExperts((current) =>
      current.map((expert) =>
        expert.id === expertId
          ? {
              ...expert,
              qaItems: expert.qaItems.filter((item) => item.id !== qaId),
            }
          : expert,
      ),
    );
  }

  return (
    <main className="app-shell slack-shell">
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
            <label className="sr-only" htmlFor="slack-language">
              Language
            </label>
            <select
              id="slack-language"
              value={language}
              onChange={(event) =>
                onLanguageChange(event.target.value as NaviLanguage)
              }
            >
              <option value="en">English</option>
              <option value="ja">日本語</option>
              <option value="zh_cn">中文（简体）</option>
              <option value="zh_tw">中文（繁體）</option>
              <option value="mn">Монгол</option>
            </select>
            <button
              className="lock-toggle"
              type="button"
              onClick={onLock}
              aria-label="Lock Expert Tools"
              title="Lock Expert Tools"
            >
              🔒
            </button>
            <button
              className="theme-toggle"
              type="button"
              onClick={onToggleTheme}
              aria-label="Toggle theme"
            >
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </header>

        <ToolSwitcher active="slack" onSelect={onSelectTool} />

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
          <label className="field-label" htmlFor="slack-raw-experts">
            {t.label}
          </label>
          <textarea
            id="slack-raw-experts"
            className="raw-input slack-raw-input"
            value={rawProfiles}
            onChange={(event) => setRawProfiles(event.target.value)}
            placeholder={t.placeholder}
            spellCheck={false}
          />
          <div className="button-row">
            <button className="button button-primary" type="button" onClick={generatePosts}>
              {t.generate}
            </button>
            <button className="button button-danger" type="button" onClick={clearPosts}>
              {t.clear}
            </button>
          </div>
          {notice && (
            <div className={`message ${noticeType}`} role="status">
              {notice}
            </div>
          )}
        </section>

        <section className="card slack-results" id="slack-results">
          <div className="section-heading">
            <div>
              <h2>{t.results}</h2>
              <p>{t.resultsHelp}</p>
            </div>
            {experts.length > 0 && (
              <button className="button button-secondary" type="button" onClick={copyAllExperts}>
                {copiedId === "all" ? t.copiedAll : t.copyAll}
              </button>
            )}
          </div>

          {!experts.length ? (
            <div className="empty-state">{t.empty}</div>
          ) : (
            <div className="slack-expert-list">
              {experts.map((expert) => {
                const screening = formatSlackScreening(expert.qaItems);
                const history = slackHistoryLines(expert.employmentHistory);
                const availability = slackAvailability(expert);
                const isEditing = editingId === expert.id;
                return (
                  <article className="slack-expert-card" key={expert.id}>
                    <div className="slack-expert-toolbar">
                      <div>
                        <strong>{expert.number} · {expert.name}</strong>
                        <small>{expert.company}</small>
                      </div>
                      <div className="slack-toolbar-actions">
                        <button
                          className="button button-muted"
                          type="button"
                          onClick={() => setEditingId(isEditing ? "" : expert.id)}
                        >
                          {isEditing ? t.doneEditing : t.edit}
                        </button>
                        <button
                          className="button button-primary slack-copy-button"
                          type="button"
                          onClick={() => copyExpert(expert)}
                        >
                          {copiedId === expert.id ? t.copied : t.copy}
                        </button>
                      </div>
                    </div>
                    {isEditing && (
                      <div className="slack-editor">
                        <p className="slack-editor-help">{t.editHelp}</p>
                        <div className="slack-editor-grid slack-editor-header-grid">
                          <label>
                            <span>{t.number}</span>
                            <input
                              value={expert.number}
                              onChange={(event) => updateExpert(expert.id, "number", event.target.value)}
                            />
                          </label>
                          <label>
                            <span>{t.name}</span>
                            <input
                              value={expert.name}
                              onChange={(event) => updateExpert(expert.id, "name", event.target.value)}
                            />
                          </label>
                          <label className="slack-editor-wide">
                            <span>{t.titleField}</span>
                            <input
                              value={expert.title}
                              onChange={(event) => updateExpert(expert.id, "title", event.target.value)}
                            />
                          </label>
                        </div>

                        <label className="slack-editor-field">
                          <span>{t.introduction}</span>
                          <textarea
                            value={expert.introduction}
                            onChange={(event) => updateExpert(expert.id, "introduction", event.target.value)}
                          />
                        </label>

                        <label className="slack-editor-field slack-screening-heading-field">
                          <span>{t.screeningLabel}</span>
                          <input
                            value={expert.screeningLabel}
                            onChange={(event) => updateExpert(expert.id, "screeningLabel", event.target.value)}
                          />
                        </label>

                        <div className="slack-qa-editor">
                          {expert.qaItems.map((item, index) => (
                            <div className="slack-qa-row" key={item.id}>
                              <div className="slack-qa-row-heading">
                                <strong>Q&amp;A {index + 1}</strong>
                                <button
                                  className="text-danger"
                                  type="button"
                                  onClick={() => removeQa(expert.id, item.id)}
                                >
                                  {t.removeQa}
                                </button>
                              </div>
                              <label>
                                <span>{t.question}</span>
                                <textarea
                                  value={item.question}
                                  onChange={(event) => updateQa(expert.id, item.id, "question", event.target.value)}
                                />
                              </label>
                              <label>
                                <span>{t.answer}</span>
                                <textarea
                                  value={item.answer}
                                  onChange={(event) => updateQa(expert.id, item.id, "answer", event.target.value)}
                                />
                              </label>
                            </div>
                          ))}
                          <button
                            className="button button-muted slack-add-qa"
                            type="button"
                            onClick={() => addQa(expert.id)}
                          >
                            ＋ {t.addQa}
                          </button>
                        </div>

                        <div className="slack-editor-grid">
                          <label>
                            <span>{t.history}</span>
                            <textarea
                              value={expert.employmentHistory}
                              onChange={(event) => updateExpert(expert.id, "employmentHistory", event.target.value)}
                            />
                          </label>
                          <label>
                            <span>{t.availability}</span>
                            <textarea
                              value={expert.availability}
                              onChange={(event) => updateExpert(expert.id, "availability", event.target.value)}
                            />
                          </label>
                          <label>
                            <span>{t.location}</span>
                            <input
                              value={expert.location}
                              onChange={(event) => updateExpert(expert.id, "location", event.target.value)}
                            />
                          </label>
                          <label>
                            <span>{t.fee}</span>
                            <input
                              value={expert.fee}
                              onChange={(event) => updateExpert(expert.id, "fee", event.target.value)}
                            />
                          </label>
                        </div>
                      </div>
                    )}
                    <div className="slack-preview">
                      <p className="slack-headline">
                        <strong>{expert.number} - {expert.name} - ✅{expert.title}</strong>
                      </p>
                      {expert.introduction && <p>{expert.introduction}</p>}
                      {screening && (
                        <pre className="slack-code-block">{expert.screeningLabel}{"\n\n"}{screening}</pre>
                      )}
                      {history.length > 0 && (
                        <div className="slack-history">
                          <strong>{t.employment}</strong>
                          <ul>
                            {history.map((line, index) => <li key={`${expert.id}-history-${index}`}>{line}</li>)}
                          </ul>
                        </div>
                      )}
                      {availability.slots.length > 0 && (
                        <div className="slack-availability">
                          <strong>{availability.heading}</strong>
                          <ul>
                            {availability.slots.map((slot, index) => (
                              <li key={`${expert.id}-slot-${index}`}><code>{slot}</code></li>
                            ))}
                          </ul>
                        </div>
                      )}
                      {expert.location && <p>This specialist is based in {expert.location}.</p>}
                      {expert.fee && <p className="slack-fee"><strong>{expert.fee}</strong></p>}
                    </div>
                  </article>
                );
              })}
            </div>
          )}
        </section>

        <footer>Taya Tool · Slack Formatter</footer>
      </div>
    </main>
  );
}

export default function Home() {
  const [activeTool, setActiveTool] = useState<ToolView>("excel");
  const [expertUnlocked, setExpertUnlocked] = useState(false);
  const [testerPassword, setTesterPassword] = useState("");
  const [passwordError, setPasswordError] = useState("");
  const [checkingPassword, setCheckingPassword] = useState(false);
  const [language, setLanguage] = useState<NaviLanguage>("en");
  const [theme, setTheme] = useState<"light" | "dark">("light");
  const [workflowMode, setWorkflowMode] = useState<WorkflowMode>("create");
  const [raw, setRaw] = useState("");
  const [records, setRecords] = useState<ExpertRecord[]>([]);
  const [existingRecords, setExistingRecords] = useState<ExpertRecord[]>([]);
  const [comparisonItems, setComparisonItems] = useState<ComparisonItem[]>([]);
  const [retainedCount, setRetainedCount] = useState(0);
  const [importedWorkbook, setImportedWorkbook] =
    useState<ImportedWorkbookSummary | null>(null);
  const [readingExcel, setReadingExcel] = useState(false);
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
  const [includeExpertSummary, setIncludeExpertSummary] = useState(false);
  const [includeUpdateSummary, setIncludeUpdateSummary] = useState(false);
  const [updateSummaryLanguage, setUpdateSummaryLanguage] =
    useState<UpdateSummaryLanguage>("ja");
  const expertLanguage: Language =
    language === "ja"
      ? "ja"
      : language === "zh_cn" || language === "zh_tw"
        ? "zh"
        : "en";
  const t = translations[expertLanguage];

  useEffect(() => {
    const unlocked = window.sessionStorage.getItem(EXPERT_ACCESS_KEY) === "unlocked";
    const savedLanguage = window.localStorage.getItem(GLOBAL_LANGUAGE_KEY);
    const timer = window.setTimeout(() => {
      setExpertUnlocked(unlocked);
      if (GLOBAL_LANGUAGES.includes(savedLanguage as NaviLanguage)) {
        setLanguage(savedLanguage as NaviLanguage);
      }
    }, 0);
    return () => window.clearTimeout(timer);
  }, []);

  function changeLanguage(nextLanguage: NaviLanguage) {
    setLanguage(nextLanguage);
    window.localStorage.setItem(GLOBAL_LANGUAGE_KEY, nextLanguage);
  }

  const warningCount = useMemo(
    () => records.reduce((total, record) => total + record.warnings.length, 0),
    [records],
  );

  const comparisonStats = useMemo(
    () =>
      comparisonItems.reduce(
        (stats, item) => ({ ...stats, [item.status]: stats[item.status] + 1 }),
        { new: 0, changed: 0, unchanged: 0 },
      ),
    [comparisonItems],
  );

  const updateSummaryEntries = useMemo(
    () =>
      buildUpdateSummaryEntries(records, existingRecords, comparisonItems),
    [records, existingRecords, comparisonItems],
  );

  const updateSummaryStats = useMemo(
    () =>
      updateSummaryEntries.reduce(
        (stats, entry) => ({
          ...stats,
          [entry.status]: stats[entry.status] + 1,
        }),
        { new: 0, updated: 0, unchanged: 0, retained: 0, removed: 0 },
      ),
    [updateSummaryEntries],
  );

  function changeTheme() {
    const next = theme === "light" ? "dark" : "light";
    setTheme(next);
    document.documentElement.dataset.theme = next;
  }

  async function unlockExpert(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setCheckingPassword(true);
    setPasswordError("");
    try {
      const passwordHash = await hashPassword(testerPassword);
      if (passwordHash !== EXPERT_PASSWORD_HASH) {
        setPasswordError("Incorrect password. Please try again.");
        return;
      }
      window.sessionStorage.setItem(EXPERT_ACCESS_KEY, "unlocked");
      setExpertUnlocked(true);
      setTesterPassword("");
    } catch {
      setPasswordError("Unable to verify the password. Refresh the page and try again.");
    } finally {
      setCheckingPassword(false);
    }
  }

  function lockExpert() {
    window.sessionStorage.removeItem(EXPERT_ACCESS_KEY);
    setExpertUnlocked(false);
    setTesterPassword("");
    setPasswordError("");
  }

  function switchWorkflowMode(nextMode: WorkflowMode) {
    setWorkflowMode(nextMode);
    setRaw("");
    setRecords([]);
    setExistingRecords([]);
    setComparisonItems([]);
    setRetainedCount(0);
    setImportedWorkbook(null);
    setSheetMode("single");
    setSingleSheetName("Expert List");
    setCustomSheets(["Expert List"]);
    setIncludeExpertSummary(false);
    setIncludeUpdateSummary(false);
    setUpdateSummaryLanguage("ja");
    setFileName(nextMode === "update" ? "Expert_List_updated.xlsx" : "Expert_List.xlsx");
    setMessage("");
    setMessageType("");
  }

  async function loadExistingExcel(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    event.target.value = "";
    if (!file) return;

    setReadingExcel(true);
    setMessage("");
    try {
      const { Workbook } = await import("exceljs");
      const workbook = new Workbook();
      const buffer = await file.arrayBuffer();
      await workbook.xlsx.load(buffer);

      const metaSheet = workbook.getWorksheet(TAYA_META_SHEET);
      const metaMarked =
        metaSheet && excelCellText(metaSheet.getCell("A1").value) === TAYA_META_MARKER;
      const metadataIds = new Map<string, string>();
      const metadataTitles = new Map<string, string>();
      if (metaMarked && metaSheet) {
        for (let rowNumber = 3; rowNumber <= metaSheet.actualRowCount; rowNumber += 1) {
          const row = metaSheet.getRow(rowNumber);
          const stableId = excelCellText(row.getCell(1).value);
          const sheetName = excelCellText(row.getCell(5).value);
          const excelRow = excelCellText(row.getCell(6).value);
          if (stableId && sheetName && excelRow) {
            const metadataKey = `${sheetName}::${excelRow}`;
            metadataIds.set(metadataKey, stableId);
            metadataTitles.set(
              metadataKey,
              excelCellText(row.getCell(7).value),
            );
          }
        }
      }

      const importedRecords: ExpertRecord[] = [];
      const supportedSheetNames: string[] = [];
      workbook.worksheets.forEach((sheet) => {
        if (sheet.name === TAYA_META_SHEET) return;
        const headers = EXCEL_HEADERS.map((_, index) =>
          excelCellText(sheet.getCell(2, index + 3).value),
        );
        const supported = EXCEL_HEADERS.every(
          (header, index) => headers[index] === header,
        );
        if (!supported) return;

        supportedSheetNames.push(sheet.name);
        for (let rowNumber = 3; rowNumber <= sheet.actualRowCount; rowNumber += 1) {
          const values = EXCEL_HEADERS.map((_, index) =>
            excelCellText(sheet.getCell(rowNumber, index + 3).value),
          );
          if (!values.some(Boolean)) continue;
          const recordBase = {
            number: values[0],
            name: values[1],
            company: values[2],
            title: metadataTitles.get(`${sheet.name}::${rowNumber}`) ?? "",
            relevantExperience: removeAvailabilityPlaceholder(values[3]),
            employmentHistory: removeAvailabilityPlaceholder(values[4]),
            introduction: removeAvailabilityPlaceholder(values[5]),
            screening: removeAvailabilityPlaceholder(values[6]),
            fee: values[7],
            availability: removeAvailabilityPlaceholder(values[8]),
          };
          importedRecords.push({
            id: `import-${sheet.name}-${rowNumber}-${Date.now()}`,
            stableId:
              metadataIds.get(`${sheet.name}::${rowNumber}`) ?? createStableId(),
            ...recordBase,
            sheetName: sheet.name,
            warnings: calculateWarnings(recordBase),
          });
        }
      });

      const creatorMatches = comparableText(workbook.creator ?? "").includes(
        "taya expert list builder",
      );
      if (!supportedSheetNames.length || (!metaMarked && !creatorMatches)) {
        throw new Error("Unsupported Taya workbook");
      }

      setExistingRecords(importedRecords);
      setRecords(importedRecords);
      setImportedWorkbook({ fileName: file.name, sheetNames: supportedSheetNames });
      setCustomSheets(supportedSheetNames);
      setSheetMode("custom");
      setComparisonItems([]);
      setRetainedCount(importedRecords.length);
      setFileName(
        `${file.name.replace(/\.xlsx$/i, "") || "Expert_List"}_updated.xlsx`,
      );
      setMessage(
        `${t.uploadedExcel}: ${file.name} · ${importedRecords.length} ${t.experts}`,
      );
      setMessageType("success");
    } catch (error) {
      console.error(error);
      setExistingRecords([]);
      setRecords([]);
      setImportedWorkbook(null);
      setComparisonItems([]);
      setRetainedCount(0);
      setMessage(t.invalidExcel);
      setMessageType("error");
    } finally {
      setReadingExcel(false);
    }
  }

  function runParser() {
    const parsed = parseExperts(raw);
    if (!parsed.length) {
      setMessage(t.parseError);
      setMessageType("error");
      return;
    }

    if (workflowMode === "update") {
      if (!importedWorkbook || !existingRecords.length) {
        setMessage(t.uploadRequired);
        setMessageType("error");
        return;
      }
      const comparison = compareWithExisting(
        parsed,
        existingRecords,
        customSheets[0] || "Expert List",
      );
      setRecords(comparison.mergedRecords);
      setComparisonItems(comparison.comparisons);
      setRetainedCount(comparison.retainedCount);
      setSheetMode("custom");
      setMessage(t.updateParsed);
    } else {
      setRecords(parsed);
      setCustomSheets(["Expert List"]);
      setComparisonItems([]);
      setRetainedCount(0);
      setMessage(
        expertLanguage === "en"
          ? `${parsed.length} ${t.parsed}`
          : `${parsed.length}${t.parsed}`,
      );
    }
    setMessageType("success");
    window.setTimeout(
      () => document.getElementById("results")?.scrollIntoView({ behavior: "smooth" }),
      50,
    );
  }

  function clearAll() {
    setRaw("");
    setRecords([]);
    setExistingRecords([]);
    setComparisonItems([]);
    setRetainedCount(0);
    setImportedWorkbook(null);
    setSheetMode("single");
    setSingleSheetName("Expert List");
    setCustomSheets(["Expert List"]);
    setNewSheetName("");
    setSheetOrganizerOpen(false);
    setIncludeUpdateSummary(false);
    setUpdateSummaryLanguage("ja");
    setMessage("");
    setMessageType("");
  }

  function toggleComparisonChange(itemId: string, field: DataField) {
    const item = comparisonItems.find((candidate) => candidate.id === itemId);
    const change = item?.changes.find((candidate) => candidate.field === field);
    if (!item || !change || !item.existingId) return;
    const nextUseLatest = !change.useLatest;

    setComparisonItems((current) =>
      current.map((candidate) =>
        candidate.id === itemId
          ? {
              ...candidate,
              changes: candidate.changes.map((candidateChange) =>
                candidateChange.field === field
                  ? { ...candidateChange, useLatest: nextUseLatest }
                  : candidateChange,
              ),
            }
          : candidate,
      ),
    );
    setRecords((current) =>
      current.map((record) => {
        if (record.id !== item.existingId) return record;
        const next = {
          ...record,
          [field]: nextUseLatest ? change.newValue : change.oldValue,
        };
        return { ...next, warnings: calculateWarnings(next) };
      }),
    );
  }

  function setAllComparisonChanges(useLatest: boolean) {
    setComparisonItems((current) =>
      current.map((item) => ({
        ...item,
        changes: item.changes.map((change) => ({ ...change, useLatest })),
      })),
    );
    setRecords((current) =>
      current.map((record) => {
        const item = comparisonItems.find(
          (candidate) => candidate.existingId === record.id,
        );
        if (!item?.changes.length) return record;
        const next = { ...record };
        item.changes.forEach((change) => {
          next[change.field] = useLatest ? change.newValue : change.oldValue;
        });
        return { ...next, warnings: calculateWarnings(next) };
      }),
    );
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
      newSheetName.trim() || (expertLanguage === "ja" ? "新しいSheet" : expertLanguage === "zh" ? "新Sheet" : "New Sheet"),
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
      workbook.modified = new Date();
      workbook.title = "Taya Tool Expert List";
      workbook.subject = "Taya Tool supported expert list template";

      const sheetGroups = groupRecordsForSheets(
        records,
        sheetMode,
        singleSheetName,
      );
      const usedSheetNames = new Set<string>();
      const metaRows: string[][] = [];
      const resolvedSheetNames = new Map<string, string>();

      let expertSummarySheet: ReturnType<typeof workbook.addWorksheet> | null =
        null;
      if (includeExpertSummary) {
        const expertSummarySheetName = uniqueSheetName(
          "Expert Summary",
          usedSheetNames,
        );
        expertSummarySheet = workbook.addWorksheet(expertSummarySheetName, {
          views: [
            {
              state: "frozen",
              ySplit: 4,
              showGridLines: false,
            },
          ],
        });
        [14, 27, 36, 72, 31].forEach((width, index) => {
          const column = expertSummarySheet?.getColumn(index + 1);
          if (column) column.width = width;
        });

        expertSummarySheet.mergeCells("A1:E1");
        const expertSummaryTitle = expertSummarySheet.getCell("A1");
        expertSummaryTitle.value = "Expert Summary";
        expertSummaryTitle.font = {
          name: "Yu Gothic",
          size: 16,
          bold: true,
          color: { argb: "FFF1F5F9" },
        };
        expertSummaryTitle.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0B1E2D" },
        };
        expertSummaryTitle.alignment = {
          vertical: "middle",
          horizontal: "left",
        };
        expertSummarySheet.getRow(1).height = 34;

        expertSummarySheet.mergeCells("A2:B2");
        expertSummarySheet.mergeCells("C2:E2");
        expertSummarySheet.getCell("A2").value = `Generated on: ${new Date().toLocaleDateString("en-GB")}`;
        expertSummarySheet.getCell("C2").value = `Experts: ${records.length}`;
        [
          expertSummarySheet.getCell("A2"),
          expertSummarySheet.getCell("C2"),
        ].forEach((cell) => {
          cell.font = {
            name: "Yu Gothic",
            size: 10,
            color: { argb: "FF536271" },
          };
          cell.alignment = { vertical: "middle", horizontal: "left" };
        });
        expertSummarySheet.getRow(2).height = 24;

        ["番号", "Expert Name", "Company", "Current Title", "Sheet"].forEach(
          (header, index) => {
            const cell = expertSummarySheet?.getCell(4, index + 1);
            if (!cell) return;
            cell.value = header;
            cell.font = {
              name: "Yu Gothic",
              size: 10,
              bold: true,
              color: { argb: "FF102A3A" },
            };
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: { argb: "FFDBE9F7" },
            };
            cell.alignment = {
              vertical: "middle",
              horizontal: "left",
              wrapText: true,
            };
            cell.border = {
              bottom: { style: "medium", color: { argb: "FF1F3A4D" } },
            };
          },
        );
        expertSummarySheet.getRow(4).height = 30;
      }

      if (
        workflowMode === "update" &&
        includeUpdateSummary &&
        comparisonItems.length > 0
      ) {
        const summaryText = updateSummaryText[updateSummaryLanguage];
        const summarySheetName = uniqueSheetName(
          summaryText.sheetName,
          usedSheetNames,
        );
        const summarySheet = workbook.addWorksheet(summarySheetName, {
          views: [
            {
              state: "frozen",
              ySplit: 7,
              showGridLines: false,
            },
          ],
        });

        [16, 13, 25, 30, 42, 72].forEach((width, index) => {
          summarySheet.getColumn(index + 1).width = width;
        });

        summarySheet.mergeCells("A1:F1");
        const summaryTitle = summarySheet.getCell("A1");
        summaryTitle.value = summaryText.title;
        summaryTitle.font = {
          name: "Yu Gothic",
          size: 16,
          bold: true,
          color: { argb: "FFF1F5F9" },
        };
        summaryTitle.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FF0B1E2D" },
        };
        summaryTitle.alignment = { vertical: "middle", horizontal: "left" };
        summarySheet.getRow(1).height = 34;

        summarySheet.mergeCells("A2:C2");
        summarySheet.mergeCells("D2:F2");
        summarySheet.getCell("A2").value = `${summaryText.generatedOn}: ${new Date().toLocaleDateString(updateSummaryLanguage === "ja" ? "ja-JP" : "en-GB")}`;
        summarySheet.getCell("D2").value = `${summaryText.sourceFile}: ${importedWorkbook?.fileName ?? "—"}`;
        [summarySheet.getCell("A2"), summarySheet.getCell("D2")].forEach(
          (cell) => {
            cell.font = { name: "Yu Gothic", size: 10, color: { argb: "FF536271" } };
            cell.alignment = { vertical: "middle", horizontal: "left" };
          },
        );
        summarySheet.getRow(2).height = 24;

        const countCells = [
          ["A4", summaryText.statuses.new, "B4", updateSummaryStats.new],
          ["C4", summaryText.statuses.updated, "D4", updateSummaryStats.updated],
          ["E4", summaryText.statuses.unchanged, "F4", updateSummaryStats.unchanged],
          ["A5", summaryText.statuses.retained, "B5", updateSummaryStats.retained],
          ["C5", summaryText.statuses.removed, "D5", updateSummaryStats.removed],
        ] as const;
        countCells.forEach(([labelCell, label, valueCell, value]) => {
          summarySheet.getCell(labelCell).value = label;
          summarySheet.getCell(valueCell).value = value;
          summarySheet.getCell(labelCell).font = {
            name: "Yu Gothic",
            size: 10,
            bold: true,
            color: { argb: "FF536271" },
          };
          summarySheet.getCell(valueCell).font = {
            name: "Yu Gothic",
            size: 12,
            bold: true,
            color: { argb: "FF102A3A" },
          };
          summarySheet.getCell(labelCell).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF1F5F9" },
          };
          summarySheet.getCell(valueCell).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFF1F5F9" },
          };
        });

        const summaryHeaders = [
          summaryText.status,
          summaryText.number,
          summaryText.name,
          summaryText.company,
          summaryText.updatedItems,
          summaryText.details,
        ];
        summaryHeaders.forEach((header, index) => {
          const cell = summarySheet.getCell(7, index + 1);
          cell.value = header;
          cell.font = {
            name: "Yu Gothic",
            size: 10,
            bold: true,
            color: { argb: "FF102A3A" },
          };
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFDBE9F7" },
          };
          cell.alignment = { vertical: "middle", horizontal: "left" };
          cell.border = {
            bottom: { style: "medium", color: { argb: "FF1F3A4D" } },
          };
        });
        summarySheet.getRow(7).height = 30;

        const statusFills: Record<UpdateSummaryStatus, string> = {
          new: "FFDDF4E8",
          updated: "FFFFE9C7",
          unchanged: "FFEAF0F5",
          retained: "FFEAF0F5",
          removed: "FFFFE1E1",
        };
        updateSummaryEntries.forEach((entry, index) => {
          const rowNumber = index + 8;
          const values = [
            summaryText.statuses[entry.status],
            entry.number,
            entry.name,
            entry.company,
            entry.changes
              .map((change) => summaryText.fields[change.field])
              .join(" / ") || "—",
            updateSummaryDetails(entry, updateSummaryLanguage),
          ];
          values.forEach((value, valueIndex) => {
            const cell = summarySheet.getCell(rowNumber, valueIndex + 1);
            cell.value = value;
            cell.font = {
              name: "Yu Gothic",
              size: 10,
              bold: valueIndex === 0,
              color: { argb: "FF17212B" },
            };
            cell.alignment = {
              vertical: "top",
              horizontal: "left",
              wrapText: true,
            };
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: {
                argb:
                  valueIndex === 0
                    ? statusFills[entry.status]
                    : index % 2 === 0
                      ? "FFFFFFFF"
                      : "FFF5F8FB",
              },
            };
            cell.border = {
              bottom: { style: "thin", color: { argb: "FFD0D9E2" } },
            };
          });
          summarySheet.getRow(rowNumber).height = Math.min(
            110,
            Math.max(34, 24 + entry.changes.length * 14),
          );
        });
        summarySheet.autoFilter = {
          from: "A7",
          to: `F${updateSummaryEntries.length + 7}`,
        };
        summarySheet.headerFooter.oddFooter =
          "&LGenerated by Taya Tool&CPage &P / &N";
      }

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

      EXCEL_HEADERS.forEach((header, index) => {
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
        resolvedSheetNames.set(record.id, sheetName);
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
        metaRows.push([
          record.stableId || createStableId(),
          record.number,
          record.name,
          record.company,
          sheetName,
          String(rowNumber),
          record.title,
        ]);
      });

      sheet.autoFilter = {
        from: "C2",
        to: `K${group.records.length + 2}`,
      };
      sheet.headerFooter.oddFooter = "&LGenerated by Taya Expert List Builder&CPage &P / &N";
      });

      if (expertSummarySheet) {
        records.forEach((record, index) => {
          const rowNumber = index + 5;
          const values = [
            record.number,
            record.name,
            record.company,
            record.title,
            resolvedSheetNames.get(record.id) ?? record.sheetName,
          ];
          values.forEach((value, valueIndex) => {
            const cell = expertSummarySheet?.getCell(
              rowNumber,
              valueIndex + 1,
            );
            if (!cell) return;
            cell.value = value;
            cell.font = {
              name: "Yu Gothic",
              size: 10,
              color: { argb: "FF17212B" },
            };
            cell.alignment = {
              vertical: "top",
              horizontal: "left",
              wrapText: true,
            };
            cell.fill = {
              type: "pattern",
              pattern: "solid",
              fgColor: {
                argb: index % 2 === 0 ? "FFFFFFFF" : "FFF5F8FB",
              },
            };
            cell.border = {
              bottom: { style: "thin", color: { argb: "FFD0D9E2" } },
            };
          });
          const summaryRow = expertSummarySheet?.getRow(rowNumber);
          if (summaryRow) {
            summaryRow.height = Math.min(
              100,
              Math.max(32, 24 + Math.ceil(record.title.length / 70) * 14),
            );
          }
        });
        expertSummarySheet.autoFilter = {
          from: "A4",
          to: `E${records.length + 4}`,
        };
        expertSummarySheet.headerFooter.oddFooter =
          "&LGenerated by Taya Tool&CPage &P / &N";
      }

      const metaSheet = workbook.addWorksheet(TAYA_META_SHEET);
      metaSheet.state = "veryHidden";
      metaSheet.addRow([TAYA_META_MARKER, "1.4"]);
      metaSheet.addRow([
        "Stable ID",
        "Number",
        "Name",
        "Company",
        "Sheet",
        "Row",
        "Title",
      ]);
      metaRows.forEach((row) => metaSheet.addRow(row));

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

  if (activeTool === "navi") {
    return (
      <TayaNaviPanel
        theme={theme}
        language={language}
        onToggleTheme={changeTheme}
        onLanguageChange={changeLanguage}
        onSelectTool={setActiveTool}
      />
    );
  }

  if (!expertUnlocked) {
    return (
      <ExpertPasswordGate
        active={activeTool}
        theme={theme}
        password={testerPassword}
        error={passwordError}
        checking={checkingPassword}
        onPasswordChange={(value) => {
          setTesterPassword(value);
          if (passwordError) setPasswordError("");
        }}
        onSubmit={unlockExpert}
        onToggleTheme={changeTheme}
        onSelectTool={setActiveTool}
      />
    );
  }

  if (activeTool === "slack") {
    return (
      <SlackFormatterPanel
        theme={theme}
        language={language}
        onToggleTheme={changeTheme}
        onLanguageChange={changeLanguage}
        onSelectTool={setActiveTool}
        onLock={lockExpert}
      />
    );
  }

  const longFields: Array<keyof Pick<
    ExpertRecord,
    | "title"
    | "relevantExperience"
    | "employmentHistory"
    | "introduction"
    | "screening"
    | "availability"
  >> = [
    "title",
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
              onChange={(event) =>
                changeLanguage(event.target.value as NaviLanguage)
              }
            >
              <option value="en">English</option>
              <option value="ja">日本語</option>
              <option value="zh_cn">中文（简体）</option>
              <option value="zh_tw">中文（繁體）</option>
              <option value="mn">Монгол</option>
            </select>
            <button
              className="lock-toggle"
              type="button"
              onClick={lockExpert}
              aria-label="Lock Expert Excel"
              title="Lock Expert Excel"
            >
              🔒
            </button>
            <button className="theme-toggle" type="button" onClick={changeTheme} aria-label="Toggle theme">
              {theme === "light" ? "🌙" : "☀️"}
            </button>
          </div>
        </header>

        <ToolSwitcher active="excel" onSelect={setActiveTool} />

        <div className="privacy-note">
          <span aria-hidden="true">🔒</span>
          {t.privacy}
        </div>

        <div className="workflow-tabs" role="tablist" aria-label="Excel workflow">
          <button
            className={workflowMode === "create" ? "is-active" : ""}
            type="button"
            role="tab"
            aria-selected={workflowMode === "create"}
            onClick={() => switchWorkflowMode("create")}
          >
            <span>＋</span>
            {t.modeCreate}
          </button>
          <button
            className={workflowMode === "update" ? "is-active" : ""}
            type="button"
            role="tab"
            aria-selected={workflowMode === "update"}
            onClick={() => switchWorkflowMode("update")}
          >
            <span>↻</span>
            {t.modeUpdate}
          </button>
        </div>

        <section className="card">
          <div className="section-heading">
            <div>
              <h2>{workflowMode === "update" ? t.latestInfoTitle : t.inputTitle}</h2>
              <p>{workflowMode === "update" ? t.latestInfoHelp : t.inputHelp}</p>
            </div>
          </div>

          {workflowMode === "update" && (
            <div className={`upload-panel ${importedWorkbook ? "is-loaded" : ""}`}>
              <div className="upload-icon" aria-hidden="true">XL</div>
              <div className="upload-copy">
                <strong>{t.updateUploadTitle}</strong>
                <p>{t.updateUploadHelp}</p>
                {importedWorkbook && (
                  <div className="upload-summary">
                    <span>{importedWorkbook.fileName}</span>
                    <span>{existingRecords.length} {t.experts}</span>
                    <span>{importedWorkbook.sheetNames.length} Sheet</span>
                  </div>
                )}
              </div>
              <label className="file-picker">
                <input
                  type="file"
                  accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  onChange={loadExistingExcel}
                  disabled={readingExcel}
                />
                {readingExcel ? t.readingExcel : t.chooseExcel}
              </label>
            </div>
          )}

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

          {workflowMode === "update" && comparisonItems.length > 0 && (
            <div className="comparison-review">
              <div className="comparison-heading">
                <div>
                  <h3>{t.compareTitle}</h3>
                  <p>{t.compareHelp}</p>
                </div>
                <div className="comparison-actions">
                  <button
                    className="button button-secondary"
                    type="button"
                    onClick={() => setAllComparisonChanges(true)}
                  >
                    {t.acceptAll}
                  </button>
                  <button
                    className="button button-muted"
                    type="button"
                    onClick={() => setAllComparisonChanges(false)}
                  >
                    {t.keepAll}
                  </button>
                </div>
              </div>

              <div className="comparison-stats">
                <span className="comparison-stat new">{comparisonStats.new} {t.statusNew}</span>
                <span className="comparison-stat changed">{comparisonStats.changed} {t.statusChanged}</span>
                <span className="comparison-stat unchanged">{comparisonStats.unchanged} {t.statusUnchanged}</span>
                <span className="comparison-stat retained">{retainedCount} {t.statusRetained}</span>
              </div>

              <div className="comparison-list">
                {comparisonItems.map((item, index) => {
                  const statusLabel =
                    item.status === "new"
                      ? t.statusNew
                      : item.status === "changed"
                        ? t.statusChanged
                        : t.statusUnchanged;
                  return (
                    <details
                      className={`comparison-item ${item.status}`}
                      key={item.id}
                      open={item.status === "changed" && (comparisonItems.length <= 5 || index === 0)}
                    >
                      <summary>
                        <span className={`comparison-badge ${item.status}`}>{statusLabel}</span>
                        <strong>{item.latest.name || "Unnamed expert"}</strong>
                        <small>{item.latest.number} · {item.latest.company}</small>
                        <span className="chevron" aria-hidden="true">⌄</span>
                      </summary>

                      <div className="comparison-content">
                        {item.status === "new" && <p>{t.newExpertHelp}</p>}
                        {item.status === "unchanged" && <p>{t.unchangedHelp}</p>}
                        {item.changes.map((change) => (
                          <div className="change-row" key={change.field}>
                            <div className="change-field">{t.fields[change.field]}</div>
                            <div className="change-value old">
                              <span>{t.oldValue}</span>
                              <pre>{change.oldValue || "—"}</pre>
                            </div>
                            <div className="change-arrow" aria-hidden="true">→</div>
                            <div className="change-value latest">
                              <span>{t.newValue}</span>
                              <pre>{change.newValue}</pre>
                            </div>
                            <button
                              className={`change-choice ${change.useLatest ? "use-latest" : "use-old"}`}
                              type="button"
                              aria-pressed={change.useLatest}
                              onClick={() => toggleComparisonChange(item.id, change.field)}
                            >
                              {change.useLatest ? t.useLatest : t.useOld}
                            </button>
                          </div>
                        ))}
                      </div>
                    </details>
                  );
                })}
              </div>
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
          <div
            className={`update-summary-option expert-summary-option ${includeExpertSummary ? "is-selected" : ""}`}
          >
            <label className="update-summary-toggle">
              <input
                type="checkbox"
                checked={includeExpertSummary}
                onChange={(event) =>
                  setIncludeExpertSummary(event.target.checked)
                }
              />
              <span>
                <strong>{t.generateExpertSummary}</strong>
                <small>{t.expertSummaryHelp}</small>
              </span>
            </label>

            <div className="expert-summary-example">
              <span>{t.expertSummaryExample}</span>
              <div className="expert-summary-example-scroll">
                <table>
                  <thead>
                    <tr>
                      <th>番号</th>
                      <th>Expert Name</th>
                      <th>Company</th>
                      <th>Current Title</th>
                      <th>Sheet</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>#1.1</td>
                      <td>Sample Expert</td>
                      <td>Example Company</td>
                      <td>Current Director of Operations</td>
                      <td>Expert List</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          {workflowMode === "update" && comparisonItems.length > 0 && (
            <div
              className={`update-summary-option ${includeUpdateSummary ? "is-selected" : ""}`}
            >
              <label className="update-summary-toggle">
                <input
                  type="checkbox"
                  checked={includeUpdateSummary}
                  onChange={(event) =>
                    setIncludeUpdateSummary(event.target.checked)
                  }
                />
                <span>
                  <strong>{t.generateUpdateSummary}</strong>
                  <small>{t.updateSummaryHelp}</small>
                </span>
              </label>

              {includeUpdateSummary && (
                <div className="update-summary-settings">
                  <fieldset>
                    <legend>{t.updateSummaryLanguage}</legend>
                    <label>
                      <input
                        type="radio"
                        name="update-summary-language"
                        value="ja"
                        checked={updateSummaryLanguage === "ja"}
                        onChange={() => setUpdateSummaryLanguage("ja")}
                      />
                      日本語
                    </label>
                    <label>
                      <input
                        type="radio"
                        name="update-summary-language"
                        value="en"
                        checked={updateSummaryLanguage === "en"}
                        onChange={() => setUpdateSummaryLanguage("en")}
                      />
                      English
                    </label>
                  </fieldset>
                  <div className="update-summary-confirmation" role="status">
                    <strong>✓ {t.updateSummaryReady}</strong>
                    <span>
                      {updateSummaryText[updateSummaryLanguage].statuses.new}: {updateSummaryStats.new}
                      {" · "}
                      {updateSummaryText[updateSummaryLanguage].statuses.updated}: {updateSummaryStats.updated}
                    </span>
                  </div>
                </div>
              )}
            </div>
          )}
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

        <footer>Taya Tool · LinkedIn search & Expert Excel</footer>
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
