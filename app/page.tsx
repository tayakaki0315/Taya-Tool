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
  sheetNames?: string[];
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
  screeningText: string;
};

type SlackHistoryItem = {
  date: string;
  detail: string;
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

const MONTHS =
  "Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|June|July|August|September|October|November|December";
const WEEKDAYS =
  "Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday";

const translations = {
  ja: {
    title: "Taya Expert List Builder",
    version: "v1.8",
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
      "#1.1 / #A-2 - Name - … のような形式に対応しています。複数名も自動で分割します。",
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
    customSheetHelp: "各Sheetと所属エキスパートを下に直接表示します。同じエキスパートを複数のSheetへコピーできます。",
    perExpertSheetHelp: "番号と名前を使って、エキスパートごとにSheetを作成します。",
    openSheetOrganizer: "Sheet管理画面を開く",
    organizerTitle: "Sheetを自由に分ける",
    organizerHelp: "エキスパートをドラッグすると別のSheetへ移動します。複数のangleで使う場合は、コピー欄から追加してください。",
    newSheetPlaceholder: "新しいSheet名",
    addSheet: "Sheetを追加",
    emptySheet: "まだエキスパートはいません",
    dragHint: "ドラッグして移動",
    copyAction: "コピー",
    copyToSheet: "コピー先Sheet",
    copyExpert: "このSheetにエキスパートをコピー",
    selectExpert: "エキスパートを選択…",
    removeFromSheet: "このSheetから外す",
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
      "「#番号 - Name - …」で始まるエキスパート情報を見つけられませんでした。",
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
    version: "v1.8",
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
      "Supports numeric or letter-based IDs such as #1.1 and #A-2. Multiple experts are separated automatically.",
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
    customSheetHelp: "See every sheet and its experts below. An expert can be copied to more than one sheet.",
    perExpertSheetHelp: "A separate sheet will be created for each expert using their number and name.",
    openSheetOrganizer: "Open sheet organizer",
    organizerTitle: "Arrange sheets freely",
    organizerHelp: "Drag an expert to move it to another sheet. Use the copy field when the same expert is needed in multiple angles.",
    newSheetPlaceholder: "New sheet name",
    addSheet: "Add sheet",
    emptySheet: "No experts in this sheet yet",
    dragHint: "Drag to move",
    copyAction: "Copy",
    copyToSheet: "Copy to sheet",
    copyExpert: "Copy an expert to this sheet",
    selectExpert: "Select an expert…",
    removeFromSheet: "Remove from this sheet",
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
    parseError: "No expert block beginning with “#ID - Name - …” was found.",
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
    version: "v1.8",
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
      "支持 #1.1、#A-2 等数字或字母编号，也可以一次粘贴多位专家。",
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
    customSheetHelp: "下方会直接显示每个Sheet及其中的专家。同一位专家可以复制到多个Sheet。",
    perExpertSheetHelp: "系统会使用编号和姓名，为每位专家建立单独的Sheet。",
    openSheetOrganizer: "打开Sheet管理窗口",
    organizerTitle: "自由安排Sheet",
    organizerHelp: "拖拽专家会将其移动到另一个Sheet。同一位专家需要用于多个angle时，请使用复制选项。",
    newSheetPlaceholder: "新的Sheet名称",
    addSheet: "新增Sheet",
    emptySheet: "这个Sheet中还没有专家",
    dragHint: "拖拽移动",
    copyAction: "复制",
    copyToSheet: "复制到Sheet",
    copyExpert: "复制专家到这个Sheet",
    selectExpert: "选择专家……",
    removeFromSheet: "从这个Sheet移除",
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
    parseError: "没有找到以“#编号 - Name - …”开头的专家资料。",
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
  const regex = /#\S+\s+-\s+/g;
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

const SCREENING_MARKER_SOURCE = String.raw`(?:[\[【]\s*(?:Q(?:\d+)?\s*[.。:：)\-]?|A(?:\d+)?\s*[.。:：),，;；\-]?|回答\s*[.。:：\-]?|→)\s*[\]】]|[①-⑳❶-❿⓫-⓴]|Q(?:\d+)?\s*[.。:：)\-]|A(?:\d+)?\s*[.。:：),，;；\-]|Q\d+\b|A\d+\b|回答\s*[:：]|→)`;

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
    const normalized = marker.replace(/[\[\]【】\s.。:：),，;；\-]/g, "");
    markers.push({
      type: /^(?:Q|[①-⑳❶-❿⓫-⓴])/i.test(normalized) ? "question" : "answer",
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
    /\bScreen(?:ed|ing)\s+\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\s*:?/i,
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
      .replace(/^[\s✅✔☑\uFE0F]+/u, "")
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
  const header = block.match(/^#(\S+)\s+-\s+(.+?)\s+-\s+/s);
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
    sheetNames: ["Expert List"],
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

  const screenedPlain = block.match(
    /\bScreen(?:ed|ing)\s+\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\s*:?/i,
  );
  if (screenedPlain) {
    return cleanText(screenedPlain[0]).replace(/\s*:?$/, ":");
  }
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

function slackHistoryItems(source: string): SlackHistoryItem[] {
  const lines = formatEmploymentHistory(source)
    .split("\n")
    .map((line) => cleanText(line))
    .filter(Boolean);

  const datePattern = new RegExp(
    `^((?:${MONTHS})\\s+\\d{4}\\s*-\\s*(?:Present|(?:${MONTHS})\\s+\\d{4}))(?:\\s+(.+))?$`,
    "i",
  );
  const items: SlackHistoryItem[] = [];

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const match = line.match(datePattern);
    if (!match) {
      items.push({ date: "", detail: line });
      continue;
    }

    let detail = cleanText(match[2] ?? "");
    const nextLine = lines[index + 1];
    if (!detail && nextLine && !datePattern.test(nextLine)) {
      detail = nextLine;
      index += 1;
    }
    items.push({ date: cleanText(match[1]), detail });
  }

  return items;
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
        screeningText:
          formatSlackScreening(parseSlackQaItems(record.screening)) ||
          removeAvailabilityPlaceholder(record.screening),
      };
    })
    .filter((record): record is SlackExpertRecord => Boolean(record));
}

function formatSlackExpertBody(
  record: SlackExpertRecord,
  screeningStyle: "code" | "canvas" = "code",
) {
  const parts: string[] = [];
  if (record.introduction) parts.push(record.introduction);

  const screening = record.screeningText.trim();
  if (screening) {
    parts.push(
      screeningStyle === "code"
        ? `\`\`\`\n${record.screeningLabel}\n\n${screening}\n\`\`\``
        : `*${record.screeningLabel}*\n\n${screening}`,
    );
  }

  const history = slackHistoryItems(record.employmentHistory);
  if (history.length) {
    parts.push(
      `*Employment History*\n\n${history
        .map(({ date, detail }) => `• ${[date, detail].filter(Boolean).join(" | ")}`)
        .join("\n")}`,
    );
  }

  const availability = slackAvailability(record);
  if (availability.slots.length) {
    parts.push(
      `*${availability.heading}*\n\n${availability.slots
        .map((slot) =>
          screeningStyle === "canvas" ? `• ${slot}` : `• \`${slot}\``,
        )
        .join("\n")}`,
    );
  }

  if (record.location) {
    parts.push(`This specialist is based in ${record.location}.`);
  }
  if (record.fee) parts.push(`*${record.fee}*`);
  return parts.join("\n\n");
}

function formatSlackExpert(record: SlackExpertRecord) {
  return [
    `*${record.number} - ${record.name} - ✅${record.title}*`,
    formatSlackExpertBody(record),
  ]
    .filter(Boolean)
    .join("\n\n");
}

function formatSlackExpertForCombinedCopy(record: SlackExpertRecord) {
  const body = formatSlackExpertBody(record);
  const quotedBody = body
    ? body
        .split("\n")
        .map((line) => (line ? `> ${line}` : ">"))
        .join("\n")
    : "";

  return [
    `*${record.number} - ${record.name} - ✅${record.title}*`,
    quotedBody,
  ]
    .filter(Boolean)
    .join("\n\n");
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

function canvasParagraphHtml(value: string) {
  return `<p style="white-space:normal;overflow-wrap:anywhere;word-break:break-word;">${slackHtmlText(value)}</p>`;
}

function formatSlackExpertBodyHtml(
  record: SlackExpertRecord,
  screeningStyle: "code" | "canvas" = "code",
) {
  const blocks: string[] = [];
  if (record.introduction) {
    blocks.push(`<p>${slackHtmlText(record.introduction)}</p>`);
  }

  const screening = record.screeningText.trim();
  if (screening) {
    if (screeningStyle === "code") {
      blocks.push(
        `<pre>${escapeHtml(`${record.screeningLabel}\n\n${screening}`)}</pre>`,
      );
    } else {
      const screeningParagraphs = screening
        .split(/\n{2,}/)
        .map((paragraph) => paragraph.trim())
        .filter(Boolean)
        .flatMap((paragraph) =>
          paragraph
            .split("\n")
            .map((line) => line.trim())
            .filter(Boolean),
        )
        .map((line) => canvasParagraphHtml(line))
        .join("");
      blocks.push(
        `<p style="white-space:normal;overflow-wrap:anywhere;word-break:break-word;"><strong>${escapeHtml(record.screeningLabel)}</strong></p>${screeningParagraphs}`,
      );
    }
  }

  const history = slackHistoryItems(record.employmentHistory);
  if (history.length) {
    blocks.push(
      `<p><strong>Employment History</strong></p><ul>${history
        .map(({ date, detail }) => {
          const dateHtml = date ? `<strong>${escapeHtml(date)}</strong>` : "";
          const separator = date && detail ? " | " : "";
          return `<li>${dateHtml}${separator}${escapeHtml(detail)}</li>`;
        })
        .join("")}</ul>`,
    );
  }

  const availability = slackAvailability(record);
  if (availability.slots.length) {
    blocks.push(
      `<p><strong>${escapeHtml(availability.heading)}</strong></p><ul>${availability.slots
        .map((slot) =>
          screeningStyle === "canvas"
            ? `<li>${escapeHtml(slot)}</li>`
            : `<li><code>${escapeHtml(slot)}</code></li>`,
        )
        .join("")}</ul>`,
    );
  }

  if (record.location) {
    blocks.push(`<p>${escapeHtml(`This specialist is based in ${record.location}.`)}</p>`);
  }
  if (record.fee) blocks.push(`<p><strong>${escapeHtml(record.fee)}</strong></p>`);
  return blocks.join("");
}

function formatSlackExpertHtml(record: SlackExpertRecord) {
  const header = `<p><strong>${escapeHtml(`${record.number} - ${record.name} - ✅${record.title}`)}</strong></p>`;
  return `<div>${header}${formatSlackExpertBodyHtml(record)}</div>`;
}

function formatSlackExpertHtmlForCombinedCopy(record: SlackExpertRecord) {
  const header = `<p><strong>${escapeHtml(`${record.number} - ${record.name} - ✅${record.title}`)}</strong></p>`;
  const body = formatSlackExpertBodyHtml(record);
  return `<div>${header}${body ? `<blockquote>${body}</blockquote>` : ""}</div>`;
}

function formatSlackExpertForCanvas(record: SlackExpertRecord) {
  const parts = [`**${record.number} - ${record.name} - ✅${record.title}**`];

  if (record.introduction) parts.push(record.introduction);

  const screening = record.screeningText.trim();
  if (screening) {
    const compactScreening = screening
      .replace(/\r\n?/g, "\n")
      .replace(/[ \t]+\n/g, "\n")
      .replace(/\n{3,}/g, "\n\n")
      .replace(/\n+(?=Q(?:\d+)?[\s).:：])/gi, "\n\n");
    const callout = [
      "::: {.callout}",
      `**${record.screeningLabel}**`,
      compactScreening,
      ":::",
    ].join("\n");
    parts.push(callout);
  }

  const history = slackHistoryItems(record.employmentHistory);
  if (history.length) {
    parts.push(
      `**Employment History**\n${history
        .map(({ date, detail }) => `- ${[date, detail].filter(Boolean).join(" | ")}`)
        .join("\n")}`,
    );
  }

  const availability = slackAvailability(record);
  if (availability.slots.length) {
    parts.push(
      `**${availability.heading}**\n${availability.slots
        .map((slot) => `- \`${slot}\``)
        .join("\n")}`,
    );
  }

  if (record.location) {
    parts.push(`This specialist is based in ${record.location}.`);
  }
  if (record.fee) parts.push(`**${record.fee}**`);

  return parts.join("\n\n");
}
function formatSlackExpertHtmlForCanvas(record: SlackExpertRecord) {
  const header = `<p><strong>${escapeHtml(`${record.number} - ${record.name} - ✅${record.title}`)}</strong></p>`;
  const body = formatSlackExpertBodyHtml(record, "canvas");
  return `<div>${header}${body ? `<blockquote>${body}</blockquote>` : ""}</div>`;
}

function formatSlackExpertList(records: SlackExpertRecord[]) {
  return [
    "*Expert List*",
    "",
    ...records.map(
      (record) =>
        `• *${record.number} - ${record.name}* - ✅${record.title}`,
    ),
  ].join("\n");
}

function formatSlackExpertListForCanvas(records: SlackExpertRecord[]) {
  return [
    "**Expert List**",
    ...records.map(
      (record) =>
        `- **${record.number} - ${record.name}** - ✅${record.title}`,
    ),
  ].join("\n");
}
function formatSlackExpertListHtml(records: SlackExpertRecord[]) {
  return `<div><p><strong>Expert List</strong></p><ul>${records
    .map(
      (record) =>
        `<li><strong>${escapeHtml(`${record.number} - ${record.name}`)}</strong> - ${escapeHtml(`✅${record.title}`)}</li>`,
    )
    .join("")}</ul></div>`;
}

function formatSlackCanvas(records: SlackExpertRecord[]) {
  const expertDetails = records
    .map((record) => formatSlackExpertForCanvas(record))
    .join("\n\n---\n\n");
  return [formatSlackExpertListForCanvas(records), expertDetails]
    .filter(Boolean)
    .join("\n\n");
}
function formatSlackCanvasHtml(records: SlackExpertRecord[]) {
  const expertDetails = records
    .map((record) => formatSlackExpertHtmlForCanvas(record))
    .join("<br>");
  return `<div>${formatSlackExpertListHtml(records)}<br>${expertDetails}</div>`;
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

async function writeCanvasClipboard(plainText: string) {
  // Slack Canvas prioritises rich clipboard HTML and can turn long screening
  // text into a horizontally scrolling code block. Plain text is intentional:
  // it guarantees normal Canvas line wrapping.
  await navigator.clipboard.writeText(plainText);
}

const slackTranslations = {
  en: {
    title: "Slack Expert Formatter",
    version: "v1.11",
    subtitle: "Turn multiple expert profiles into clean, copy-ready Slack posts.",
    privacy: "Everything is processed in your browser. Nothing is uploaded or stored.",
    inputTitle: "1. Paste expert profiles",
    inputHelp: "Use numeric or letter-based IDs such as #1.1 and #A-2. Paste one or many profiles at once.",
    label: "Expert information",
    placeholder: "Paste expert profiles here…",
    generate: "Create Slack posts",
    clear: "Clear all",
    results: "2. Copy to Slack",
    resultsHelp: "Canvas copy puts only Screening Questions in a callout and highlights availability as code.",
    empty: "Your Slack-ready expert posts will appear here.",
    copy: "Copy for Slack",
    copied: "Copied",
    copyAll: "Copy all experts",
    copiedAll: "All experts copied",
    copyCanvas: "Copy for Canvas",
    copiedCanvas: "Canvas copied",
    copyExpertList: "Copy expert list",
    copiedExpertList: "Expert list copied",
    found: "experts formatted",
    parseError: "No expert profile beginning with #ID - Name - … was found.",
    employment: "Employment History",
    edit: "Edit",
    doneEditing: "Done editing",
    editHelp: "Screening Questions is open by default. Expand another section only when you need to correct it.",
    screeningEditorLabel: "Screening Questions",
    basicFieldsGroup: "Number, name & title",
    introductionGroup: "Expert introduction",
    historyGroup: "Employment, availability & fee",
    number: "Number",
    name: "Name",
    company: "Company",
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
    version: "v1.11",
    subtitle: "複数のエキスパート情報を、Slackに貼り付けやすい形式へ整えます。",
    privacy: "入力内容はブラウザ内だけで処理され、アップロードや保存はされません。",
    inputTitle: "1. エキスパート情報を貼り付け",
    inputHelp: "#1.1、#A-2など数字・英字の番号に対応し、複数名をまとめて貼り付けられます。",
    label: "エキスパート情報",
    placeholder: "ここにエキスパート情報を貼り付けてください…",
    generate: "Slack用に整形",
    clear: "すべてクリア",
    results: "2. Slackへコピー",
    resultsHelp: "Canvas用コピーでは、Screening Questionsのみをcalloutにし、Availabilityをcode表示します。",
    empty: "整形したSlack投稿がここに表示されます。",
    copy: "Slack用にコピー",
    copied: "コピーしました",
    copyAll: "全員をコピー",
    copiedAll: "全員をコピーしました",
    copyCanvas: "Canvas用にコピー",
    copiedCanvas: "Canvas用にコピーしました",
    copyExpertList: "Expert Listをコピー",
    copiedExpertList: "Expert Listをコピーしました",
    found: "名を整形",
    parseError: "#番号 - Name - … で始まるエキスパート情報が見つかりませんでした。",
    employment: "Employment History",
    edit: "編集",
    doneEditing: "編集を完了",
    editHelp: "Screening Questionsは最初から開いています。修正が必要な項目だけ展開してください。",
    screeningEditorLabel: "Screening Questions",
    basicFieldsGroup: "番号・名前・タイトル",
    introductionGroup: "エキスパート紹介",
    historyGroup: "経歴・Availability・Fee",
    number: "番号",
    name: "名前",
    company: "Company",
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
    version: "v1.11",
    subtitle: "将多位专家信息整理成可直接复制到 Slack 的格式。",
    privacy: "所有内容只在浏览器中处理，不会上传或保存。",
    inputTitle: "1. 粘贴专家信息",
    inputHelp: "支持 #1.1、#A-2 等数字或字母编号，并可一次粘贴多位专家。",
    label: "专家信息",
    placeholder: "在这里粘贴专家信息…",
    generate: "生成 Slack 内容",
    clear: "全部清除",
    results: "2. 复制到 Slack",
    resultsHelp: "Canvas 复制仅将 Screening Questions 放入 callout，并用 code 格式突出 Availability。",
    empty: "生成后的 Slack 内容会显示在这里。",
    copy: "复制到 Slack",
    copied: "已复制",
    copyAll: "复制全部专家",
    copiedAll: "已复制全部专家",
    copyCanvas: "复制 Canvas",
    copiedCanvas: "Canvas 内容已复制",
    copyExpertList: "复制专家名单",
    copiedExpertList: "专家名单已复制",
    found: "位专家已生成",
    parseError: "没有找到以 #编号 - Name - … 开头的专家信息。",
    employment: "Employment History",
    edit: "编辑",
    doneEditing: "完成编辑",
    editHelp: "Screening Questions 默认展开，其他内容只在需要修改时展开即可。",
    screeningEditorLabel: "Screening Questions",
    basicFieldsGroup: "编号、姓名与 Title",
    introductionGroup: "专家介绍",
    historyGroup: "经历、Availability 与 Fee",
    number: "编号",
    name: "姓名",
    company: "Company",
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
  const mergedRecords = existingRecords.map((record) => ({
    ...record,
    sheetNames: getRecordSheetNames(record),
  }));
  const comparisons: ComparisonItem[] = [];

  latestRecords.forEach((latest) => {
    const match = findExistingMatch(latest, existingRecords, usedExistingIds);
    if (!match) {
      const newRecord = {
        ...latest,
        sheetName: defaultSheetName || "Expert List",
        sheetNames: [defaultSheetName || "Expert List"],
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
      latest: {
        ...latest,
        sheetName: match.record.sheetName,
        sheetNames: getRecordSheetNames(match.record),
      },
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

function getRecordSheetNames(record: ExpertRecord) {
  const names = (record.sheetNames?.length
    ? record.sheetNames
    : [record.sheetName || "Expert List"]
  )
    .map((name) => name.trim())
    .filter(Boolean);
  return [...new Set(names.length ? names : ["Expert List"])];
}

function consolidateImportedRecords(records: ExpertRecord[]) {
  const consolidated: ExpertRecord[] = [];

  records.forEach((record) => {
    const recordNameKey = personKey(record.name);
    const recordCompanyKey = companyKey(record.company);
    const match = consolidated.find(
      (candidate) =>
        (record.stableId && candidate.stableId === record.stableId) ||
        (recordNameKey &&
          recordCompanyKey &&
          personKey(candidate.name) === recordNameKey &&
          companyKey(candidate.company) === recordCompanyKey),
    );
    if (!match) {
      consolidated.push({
        ...record,
        sheetNames: getRecordSheetNames(record),
      });
      return;
    }

    match.sheetNames = [
      ...new Set([...getRecordSheetNames(match), ...getRecordSheetNames(record)]),
    ];
    match.sheetName = match.sheetNames[0] || match.sheetName;
  });

  return consolidated;
}

function groupRecordsForSheets(
  records: ExpertRecord[],
  sheetMode: SheetMode,
  singleSheetName: string,
  customSheets: string[] = [],
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
  customSheets.forEach((name) => grouped.set(name, []));
  records.forEach((record) => {
    getRecordSheetNames(record).forEach((name) => {
      grouped.set(name, [...(grouped.get(name) ?? []), record]);
    });
  });
  return [...grouped.entries()]
    .filter(([, groupedRecords]) => groupedRecords.length > 0)
    .map(([name, groupedRecords]) => ({
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
    const oldValue =
      field === "sheetName"
        ? getRecordSheetNames(oldRecord).join(" / ")
        : oldRecord[field];
    const newValue =
      field === "sheetName"
        ? getRecordSheetNames(finalRecord).join(" / ")
        : finalRecord[field];
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
        title="Expert Excel"
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
        title="Slack Formatter"
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
        title="LinkedIn Search"
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

type BreakBubble = {
  id: number;
  x: number;
  y: number;
  size: number;
  tone: "mint" | "blue" | "violet" | "peach";
  delay: number;
};

type BreakGameKind =
  | "bubbles"
  | "odd"
  | "numbers"
  | "memory"
  | "runner"
  | "snake"
  | "flappy"
  | "stack"
  | "2048";

type MemoryCard = {
  id: number;
  symbol: string;
};

const BREAK_BUBBLES: BreakBubble[] = [
  { id: 1, x: 8, y: 14, size: 46, tone: "mint", delay: 0 },
  { id: 2, x: 35, y: 8, size: 58, tone: "blue", delay: 0.35 },
  { id: 3, x: 70, y: 13, size: 40, tone: "peach", delay: 0.7 },
  { id: 4, x: 19, y: 43, size: 36, tone: "violet", delay: 0.5 },
  { id: 5, x: 51, y: 38, size: 48, tone: "mint", delay: 0.15 },
  { id: 6, x: 79, y: 48, size: 52, tone: "blue", delay: 0.9 },
  { id: 7, x: 5, y: 72, size: 54, tone: "peach", delay: 1.05 },
  { id: 8, x: 39, y: 72, size: 38, tone: "violet", delay: 0.25 },
  { id: 9, x: 66, y: 76, size: 45, tone: "mint", delay: 0.65 },
];

const MEMORY_SYMBOLS = ["✦", "☁", "❋"];
const RUNNER_BEST_KEY = "tayaRunnerBestMsV1";

function shuffled<T>(items: T[]) {
  const next = [...items];
  for (let index = next.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [next[index], next[swapIndex]] = [next[swapIndex], next[index]];
  }
  return next;
}

function createMemoryCards(): MemoryCard[] {
  return shuffled([...MEMORY_SYMBOLS, ...MEMORY_SYMBOLS]).map((symbol, id) => ({ id, symbol }));
}

function randomBreakIndex(length: number) {
  return Math.floor(Math.random() * length);
}

type BreakQuiz = {
  prompt: string;
  answer: number;
};

function createBreakQuiz(): BreakQuiz {
  const quizKind = randomBreakIndex(3);

  if (quizKind === 0) {
    const first = randomBreakIndex(15) + 3;
    const second = randomBreakIndex(12) + 2;
    return { prompt: `${first} + ${second}`, answer: first + second };
  }

  if (quizKind === 1) {
    const second = randomBreakIndex(10) + 2;
    const answer = randomBreakIndex(13) + 2;
    return { prompt: `${answer + second} − ${second}`, answer };
  }

  const first = randomBreakIndex(8) + 2;
  const second = randomBreakIndex(8) + 2;
  return { prompt: `${first} × ${second}`, answer: first * second };
}

type RunnerState = {
  status: "ready" | "running" | "over";
  elapsed: number;
  playerY: number;
  velocity: number;
  obstacleX: number;
  speed: number;
  passed: number;
};

type RunnerCopy = {
  start: string;
  jump: string;
  startButton: string;
  jumpButton: string;
  gameOver: string;
  time: string;
  best: string;
  restart: string;
  level: string;
};

const INITIAL_RUNNER: RunnerState = {
  status: "ready",
  elapsed: 0,
  playerY: 0,
  velocity: 0,
  obstacleX: 112,
  speed: 37,
  passed: 0,
};

function formatRunnerTime(milliseconds: number) {
  return `${(milliseconds / 1000).toFixed(1)}s`;
}

type ArcadeLevel = 1 | 2 | 3;

function getArcadeLevel(score: number, level2At: number, level3At: number): ArcadeLevel {
  if (score >= level3At) return 3;
  if (score >= level2At) return 2;
  return 1;
}

function ArcadeLevelBadge({
  score,
  level2At,
  level3At,
  label,
}: {
  score: number;
  level2At: number;
  level3At: number;
  label: string;
}) {
  const level = getArcadeLevel(score, level2At, level3At);
  const nextLevelAt = level === 1 ? level2At : level === 2 ? level3At : null;

  return (
    <span key={level} className={`taya-arcade-level is-level-${level}`}>
      <strong>{label} {level}</strong>
      {nextLevelAt !== null && <small>{score}/{nextLevelAt}</small>}
      {nextLevelAt === null && <small>MAX</small>}
    </span>
  );
}

function jumpRunner(current: RunnerState): RunnerState {
  if (current.status !== "running" || current.playerY > 1) return current;
  return { ...current, velocity: 230 };
}

function TayaRunner({ copy }: { copy: RunnerCopy }) {
  const [runner, setRunner] = useState<RunnerState>(INITIAL_RUNNER);
  const [bestTime, setBestTime] = useState(0);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(RUNNER_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestTime(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (runner.status !== "running") return;
    let animationFrame = 0;
    let previousTime = window.performance.now();

    function tick(now: number) {
      const delta = Math.min((now - previousTime) / 1000, 0.034);
      previousTime = now;
      setRunner((current) => {
        if (current.status !== "running") return current;

        const playerY = Math.max(0, current.playerY + current.velocity * delta);
        let velocity = current.velocity - 420 * delta;
        if (playerY === 0 && velocity < 0) velocity = 0;

        let obstacleX = current.obstacleX - current.speed * delta;
        let passed = current.passed;
        const level = getArcadeLevel(current.passed, 5, 12);
        const speedCap = level === 1 ? 45 : level === 2 ? 61 : 80;
        const acceleration = level === 1 ? 0.75 : level === 2 ? 1.05 : 1.4;
        let speed = Math.min(speedCap, current.speed + delta * acceleration);
        if (obstacleX < -10) {
          obstacleX = 104 + randomBreakIndex(18);
          passed += 1;
          const nextLevel = getArcadeLevel(passed, 5, 12);
          const nextCap = nextLevel === 1 ? 45 : nextLevel === 2 ? 61 : 80;
          const speedStep = nextLevel === 1 ? 1.25 : nextLevel === 2 ? 1.7 : 2.15;
          speed = Math.min(nextCap, speed + speedStep);
        }

        const collided = obstacleX < 23 && obstacleX > 8 && playerY < 20;
        return {
          status: collided ? "over" : "running",
          elapsed: current.elapsed + delta * 1000,
          playerY,
          velocity,
          obstacleX,
          speed,
          passed,
        };
      });
      animationFrame = window.requestAnimationFrame(tick);
    }

    animationFrame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(animationFrame);
  }, [runner.status]);

  useEffect(() => {
    if (runner.status !== "over" || runner.elapsed <= bestTime) return;
    const nextBest = runner.elapsed;
    window.localStorage.setItem(RUNNER_BEST_KEY, String(nextBest));
    const timer = window.setTimeout(() => setBestTime(nextBest), 0);
    return () => window.clearTimeout(timer);
  }, [bestTime, runner.elapsed, runner.status]);

  useEffect(() => {
    if (runner.status !== "running") return;
    function jumpFromKeyboard(event: KeyboardEvent) {
      if (event.code !== "Space" && event.code !== "ArrowUp") return;
      event.preventDefault();
      setRunner((current) => jumpRunner(current));
    }
    window.addEventListener("keydown", jumpFromKeyboard);
    return () => window.removeEventListener("keydown", jumpFromKeyboard);
  }, [runner.status]);

  function startRunner() {
    setRunner({ ...INITIAL_RUNNER, status: "running" });
  }

  function handleRunnerAction() {
    if (runner.status === "ready" || runner.status === "over") {
      startRunner();
      return;
    }
    setRunner((current) => jumpRunner(current));
  }

  return (
    <div className="taya-runner-shell">
      <button
        className={`taya-runner-stage is-${runner.status}`}
        type="button"
        aria-label={runner.status === "running" ? copy.jump : copy.start}
        onPointerDown={(event) => {
          event.preventDefault();
          handleRunnerAction();
        }}
        onClick={(event) => {
          if (event.detail === 0) handleRunnerAction();
        }}
      >
        <ArcadeLevelBadge score={runner.passed} level2At={5} level3At={12} label={copy.level} />
        <span className="taya-runner-hud">
          <span>{copy.time} <strong>{formatRunnerTime(runner.elapsed)}</strong></span>
          <span>{copy.best} <strong>{formatRunnerTime(bestTime)}</strong></span>
        </span>
        <span className="taya-runner-cloud cloud-one" aria-hidden="true">●●</span>
        <span className="taya-runner-cloud cloud-two" aria-hidden="true">●●●</span>
        <span className="taya-runner-ground" aria-hidden="true" />
        <span
          className="taya-runner-player"
          style={{ transform: `translate3d(0, -${runner.playerY}px, 0)` }}
          aria-hidden="true"
        >
          T
        </span>
        <span className="taya-runner-obstacle" style={{ left: `${runner.obstacleX}%` }} aria-hidden="true">
          <span />
        </span>
        {runner.status !== "running" && (
          <span className="taya-runner-overlay">
            <strong>{runner.status === "over" ? copy.gameOver : "TAYA RUNNER"}</strong>
            {runner.status === "over" && <em>{copy.time}: {formatRunnerTime(runner.elapsed)}</em>}
            <small>{runner.status === "over" ? copy.restart : copy.start}</small>
          </span>
        )}
        {runner.status === "running" && <span className="taya-runner-tip">{copy.jump}</span>}
      </button>
      <button className="taya-runner-jump" type="button" onClick={handleRunnerAction}>
        {runner.status === "running" ? copy.jumpButton : copy.startButton}
      </button>
    </div>
  );
}

type SnakeDirection = "up" | "down" | "left" | "right";

type SnakePoint = {
  x: number;
  y: number;
};

type SnakeState = {
  status: "ready" | "running" | "over";
  body: SnakePoint[];
  food: SnakePoint;
  direction: SnakeDirection;
  lastDirection: SnakeDirection;
  score: number;
};

type SnakeCopy = {
  start: string;
  gameOver: string;
  score: string;
  best: string;
  restart: string;
  up: string;
  down: string;
  left: string;
  right: string;
  level: string;
};

const SNAKE_WIDTH = 14;
const SNAKE_HEIGHT = 9;
const SNAKE_BEST_KEY = "tayaSnakeBestV1";

function createInitialSnake(status: SnakeState["status"] = "ready"): SnakeState {
  return {
    status,
    body: [
      { x: 5, y: 4 },
      { x: 4, y: 4 },
      { x: 3, y: 4 },
    ],
    food: { x: 10, y: 4 },
    direction: "right",
    lastDirection: "right",
    score: 0,
  };
}

function isSamePoint(first: SnakePoint, second: SnakePoint) {
  return first.x === second.x && first.y === second.y;
}

function isOppositeDirection(first: SnakeDirection, second: SnakeDirection) {
  return (
    (first === "up" && second === "down") ||
    (first === "down" && second === "up") ||
    (first === "left" && second === "right") ||
    (first === "right" && second === "left")
  );
}

function turnSnake(current: SnakeState, direction: SnakeDirection): SnakeState {
  if (current.status !== "running" || isOppositeDirection(current.lastDirection, direction)) return current;
  return { ...current, direction };
}

function createSnakeFood(body: SnakePoint[]): SnakePoint {
  for (let attempt = 0; attempt < 60; attempt += 1) {
    const point = {
      x: randomBreakIndex(SNAKE_WIDTH),
      y: randomBreakIndex(SNAKE_HEIGHT),
    };
    if (!body.some((segment) => isSamePoint(segment, point))) return point;
  }
  return { x: 0, y: 0 };
}

function advanceSnake(current: SnakeState): SnakeState {
  if (current.status !== "running") return current;
  const head = current.body[0];
  const movement = {
    up: { x: 0, y: -1 },
    down: { x: 0, y: 1 },
    left: { x: -1, y: 0 },
    right: { x: 1, y: 0 },
  }[current.direction];
  const nextHead = { x: head.x + movement.x, y: head.y + movement.y };
  const hitWall =
    nextHead.x < 0 ||
    nextHead.x >= SNAKE_WIDTH ||
    nextHead.y < 0 ||
    nextHead.y >= SNAKE_HEIGHT;
  const ateFood = isSamePoint(nextHead, current.food);
  const collisionBody = ateFood ? current.body : current.body.slice(0, -1);
  const hitSelf = collisionBody.some((segment) => isSamePoint(segment, nextHead));
  if (hitWall || hitSelf) return { ...current, status: "over" };

  const nextBody = [nextHead, ...current.body];
  if (!ateFood) nextBody.pop();
  return {
    ...current,
    body: nextBody,
    food: ateFood ? createSnakeFood(nextBody) : current.food,
    lastDirection: current.direction,
    score: current.score + (ateFood ? 1 : 0),
  };
}

function TayaSnake({ copy }: { copy: SnakeCopy }) {
  const [game, setGame] = useState<SnakeState>(() => createInitialSnake());
  const [bestScore, setBestScore] = useState(0);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(SNAKE_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestScore(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (game.status !== "running") return;
    const level = getArcadeLevel(game.score, 4, 8);
    const delay =
      level === 1
        ? Math.max(145, 170 - game.score * 6)
        : level === 2
          ? Math.max(105, 128 - (game.score - 4) * 5)
          : Math.max(68, 92 - (game.score - 8) * 3);
    const timer = window.setInterval(() => setGame((current) => advanceSnake(current)), delay);
    return () => window.clearInterval(timer);
  }, [game.score, game.status]);

  useEffect(() => {
    if (game.status !== "over" || game.score <= bestScore) return;
    window.localStorage.setItem(SNAKE_BEST_KEY, String(game.score));
    const timer = window.setTimeout(() => setBestScore(game.score), 0);
    return () => window.clearTimeout(timer);
  }, [bestScore, game.score, game.status]);

  useEffect(() => {
    function controlSnake(event: KeyboardEvent) {
      const directionByKey: Record<string, SnakeDirection | undefined> = {
        ArrowUp: "up",
        KeyW: "up",
        ArrowDown: "down",
        KeyS: "down",
        ArrowLeft: "left",
        KeyA: "left",
        ArrowRight: "right",
        KeyD: "right",
      };
      const direction = directionByKey[event.code];
      if (!direction) return;
      event.preventDefault();
      setGame((current) => turnSnake(current, direction));
    }
    window.addEventListener("keydown", controlSnake);
    return () => window.removeEventListener("keydown", controlSnake);
  }, []);

  function startSnake() {
    setGame(createInitialSnake("running"));
  }

  return (
    <div className="taya-snake-shell">
      <div className="taya-snake-hud">
        <ArcadeLevelBadge score={game.score} level2At={4} level3At={8} label={copy.level} />
        <span>{copy.score} <strong>{game.score}</strong></span>
        <span>{copy.best} <strong>{bestScore}</strong></span>
      </div>
      <div className="taya-snake-board" role="img" aria-label={`${copy.score}: ${game.score}`}>
        {Array.from({ length: SNAKE_WIDTH * SNAKE_HEIGHT }, (_, index) => {
          const point = { x: index % SNAKE_WIDTH, y: Math.floor(index / SNAKE_WIDTH) };
          const bodyIndex = game.body.findIndex((segment) => isSamePoint(segment, point));
          const hasFood = isSamePoint(game.food, point);
          return (
            <span
              key={index}
              className={`${bodyIndex === 0 ? "is-head" : bodyIndex > 0 ? "is-body" : ""} ${hasFood ? "is-food" : ""}`}
              aria-hidden="true"
            />
          );
        })}
        {game.status !== "running" && (
          <div className="taya-snake-overlay">
            <strong>{game.status === "over" ? copy.gameOver : "TAYA SNAKE"}</strong>
            {game.status === "over" && <em>{copy.score}: {game.score}</em>}
            <button type="button" onClick={startSnake}>
              {game.status === "over" ? copy.restart : copy.start}
            </button>
          </div>
        )}
      </div>
      <div className="taya-snake-controls" aria-label="Snake controls">
        <button className="is-up" type="button" aria-label={copy.up} onClick={() => setGame((current) => turnSnake(current, "up"))}>↑</button>
        <button className="is-left" type="button" aria-label={copy.left} onClick={() => setGame((current) => turnSnake(current, "left"))}>←</button>
        <button className="is-down" type="button" aria-label={copy.down} onClick={() => setGame((current) => turnSnake(current, "down"))}>↓</button>
        <button className="is-right" type="button" aria-label={copy.right} onClick={() => setGame((current) => turnSnake(current, "right"))}>→</button>
      </div>
    </div>
  );
}

type FlappyState = {
  status: "ready" | "running" | "over";
  birdY: number;
  velocity: number;
  pipeX: number;
  gapY: number;
  score: number;
};

type FlappyCopy = {
  start: string;
  startButton: string;
  gameOver: string;
  score: string;
  best: string;
  restart: string;
  flap: string;
  level: string;
};

const FLAPPY_BEST_KEY = "tayaFlappyBestV1";
const INITIAL_FLAPPY: FlappyState = {
  status: "ready",
  birdY: 45,
  velocity: 0,
  pipeX: 108,
  gapY: 48,
  score: 0,
};

function flapBird(current: FlappyState): FlappyState {
  if (current.status !== "running") return current;
  return { ...current, velocity: -29 };
}

function TayaFlappy({ copy }: { copy: FlappyCopy }) {
  const [game, setGame] = useState<FlappyState>(INITIAL_FLAPPY);
  const [bestScore, setBestScore] = useState(0);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(FLAPPY_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestScore(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (game.status !== "running") return;
    let animationFrame = 0;
    let previousTime = window.performance.now();

    function tick(now: number) {
      const delta = Math.min((now - previousTime) / 1000, 0.034);
      previousTime = now;
      setGame((current) => {
        if (current.status !== "running") return current;

        const nextVelocity = current.velocity + 73 * delta;
        const birdY = current.birdY + nextVelocity * delta;
        const level = getArcadeLevel(current.score, 4, 8);
        const pipeSpeed =
          level === 1
            ? 24 + current.score * 0.75
            : level === 2
              ? 31 + (current.score - 4) * 0.85
              : Math.min(46, 39 + (current.score - 8) * 0.7);
        let pipeX = current.pipeX - pipeSpeed * delta;
        let gapY = current.gapY;
        let score = current.score;

        if (pipeX < -16) {
          pipeX = 108;
          gapY = 31 + randomBreakIndex(38);
          score += 1;
        }

        const gapSize = level === 1 ? 35 : level === 2 ? 29 : 24;
        const pipeNearBird = pipeX < 31 && pipeX > 17;
        const gapTop = gapY - gapSize / 2;
        const gapBottom = gapY + gapSize / 2;
        const hitPipe = pipeNearBird && (birdY < gapTop || birdY + 7 > gapBottom);
        const hitEdge = birdY < 0 || birdY > 91;

        return {
          status: hitPipe || hitEdge ? "over" : "running",
          birdY,
          velocity: nextVelocity,
          pipeX,
          gapY,
          score,
        };
      });
      animationFrame = window.requestAnimationFrame(tick);
    }

    animationFrame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(animationFrame);
  }, [game.status]);

  useEffect(() => {
    if (game.status !== "over" || game.score <= bestScore) return;
    window.localStorage.setItem(FLAPPY_BEST_KEY, String(game.score));
    const timer = window.setTimeout(() => setBestScore(game.score), 0);
    return () => window.clearTimeout(timer);
  }, [bestScore, game.score, game.status]);

  useEffect(() => {
    if (game.status !== "running") return;
    function flapFromKeyboard(event: KeyboardEvent) {
      if (event.code !== "Space" && event.code !== "ArrowUp") return;
      event.preventDefault();
      setGame((current) => flapBird(current));
    }
    window.addEventListener("keydown", flapFromKeyboard);
    return () => window.removeEventListener("keydown", flapFromKeyboard);
  }, [game.status]);

  function startFlappy() {
    setGame({ ...INITIAL_FLAPPY, status: "running" });
  }

  function handleFlappyAction() {
    if (game.status !== "running") {
      startFlappy();
      return;
    }
    setGame((current) => flapBird(current));
  }

  const flappyLevel = getArcadeLevel(game.score, 4, 8);
  const visibleGapSize = flappyLevel === 1 ? 35 : flappyLevel === 2 ? 29 : 24;
  const gapTop = Math.max(8, game.gapY - visibleGapSize / 2);
  const gapBottom = Math.min(92, game.gapY + visibleGapSize / 2);

  return (
    <div className="taya-flappy-shell">
      <button
        className={`taya-flappy-stage is-${game.status}`}
        type="button"
        aria-label={game.status === "running" ? copy.flap : copy.start}
        onPointerDown={(event) => {
          event.preventDefault();
          handleFlappyAction();
        }}
        onClick={(event) => {
          if (event.detail === 0) handleFlappyAction();
        }}
      >
        <ArcadeLevelBadge score={game.score} level2At={4} level3At={8} label={copy.level} />
        <span className="taya-arcade-hud">
          <span>{copy.score} <strong>{game.score}</strong></span>
          <span>{copy.best} <strong>{bestScore}</strong></span>
        </span>
        <span className="taya-flappy-cloud cloud-one" aria-hidden="true">●●</span>
        <span className="taya-flappy-cloud cloud-two" aria-hidden="true">●●●</span>
        <span
          className="taya-flappy-bird"
          style={{ top: `${game.birdY}%`, transform: `rotate(${Math.max(-22, Math.min(34, game.velocity * 0.75))}deg)` }}
          aria-hidden="true"
        >
          T
        </span>
        <span className="taya-flappy-pipe is-top" style={{ left: `${game.pipeX}%`, height: `${gapTop}%` }} aria-hidden="true" />
        <span className="taya-flappy-pipe is-bottom" style={{ left: `${game.pipeX}%`, top: `${gapBottom}%` }} aria-hidden="true" />
        <span className="taya-flappy-ground" aria-hidden="true" />
        {game.status !== "running" && (
          <span className="taya-arcade-overlay">
            <strong>{game.status === "over" ? copy.gameOver : "TAYA FLAPPY"}</strong>
            {game.status === "over" && <em>{copy.score}: {game.score}</em>}
            <small>{game.status === "over" ? copy.restart : copy.start}</small>
          </span>
        )}
      </button>
      <button className="taya-arcade-action" type="button" onClick={handleFlappyAction}>
        {game.status === "running" ? copy.flap : copy.startButton}
      </button>
    </div>
  );
}

type MetroAction = "run" | "jump" | "slide";
type MetroObstacleKind = "train" | "barrier" | "gate";

type MetroObstacle = {
  id: number;
  lane: 0 | 1 | 2;
  y: number;
  kind: MetroObstacleKind;
};

type MetroState = {
  status: "ready" | "running" | "over";
  lane: 0 | 1 | 2;
  action: MetroAction;
  actionUntil: number;
  obstacles: MetroObstacle[];
  spawnIn: number;
  nextObstacleId: number;
  score: number;
};

type MetroCopy = {
  start: string;
  gameOver: string;
  score: string;
  best: string;
  restart: string;
  left: string;
  right: string;
  jump: string;
  slide: string;
  jumpLabel: string;
  slideLabel: string;
  switchLabel: string;
  nextLabel: string;
  level: string;
};

const METRO_BEST_KEY = "tayaMetroBestV1";
const METRO_LEVEL_2_SCORE = 5;
const METRO_LEVEL_3_SCORE = 12;

function metroObstacleKinds(level: ArcadeLevel): MetroObstacleKind[] {
  if (level === 1) return ["barrier", "gate", "train"];
  if (level === 2) return ["barrier", "gate", "train", "train"];
  return ["train", "barrier", "gate", "train", "gate"];
}

function createMetroWave(score: number, nextObstacleId: number) {
  const level = getArcadeLevel(score, METRO_LEVEL_2_SCORE, METRO_LEVEL_3_SCORE);
  const doubleWaveChance = level === 1 ? 0 : level === 2 ? 0.26 : 0.54;
  const obstacleCount = Math.random() < doubleWaveChance ? 2 : 1;
  const lanes = shuffled([0, 1, 2] as Array<0 | 1 | 2>).slice(0, obstacleCount);
  const kinds = metroObstacleKinds(level);
  const obstacles = lanes.map((lane, index) => ({
    id: nextObstacleId + index,
    lane,
    y: -16,
    kind: kinds[randomBreakIndex(kinds.length)],
  }));
  return { obstacles, nextObstacleId: nextObstacleId + obstacles.length };
}

function metroSpawnDelay(level: ArcadeLevel) {
  const randomDelay = Math.random() * (level === 1 ? 0.22 : level === 2 ? 0.16 : 0.1);
  if (level === 1) return 1.46 + randomDelay;
  if (level === 2) return 0.98 + randomDelay;
  return 0.68 + randomDelay;
}

function createMetroState(status: MetroState["status"] = "ready"): MetroState {
  const firstWave = createMetroWave(0, 1);
  return {
    status,
    lane: 1,
    action: "run",
    actionUntil: 0,
    obstacles: firstWave.obstacles,
    spawnIn: 1.35,
    nextObstacleId: firstWave.nextObstacleId,
    score: 0,
  };
}

function moveMetroLane(current: MetroState, amount: -1 | 1): MetroState {
  if (current.status !== "running") return current;
  return { ...current, lane: Math.max(0, Math.min(2, current.lane + amount)) as 0 | 1 | 2 };
}

function setMetroAction(current: MetroState, action: Exclude<MetroAction, "run">): MetroState {
  if (current.status !== "running") return current;
  const duration = action === "jump" ? 720 : 590;
  return { ...current, action, actionUntil: window.performance.now() + duration };
}

function metroLanePosition(lane: number, y: number) {
  const spread = 15 + Math.max(0, Math.min(100, y)) * 0.18;
  return 50 + (lane - 1) * spread;
}

function TayaMetro({ copy }: { copy: MetroCopy }) {
  const [game, setGame] = useState<MetroState>(() => createMetroState());
  const [bestScore, setBestScore] = useState(0);
  const [swipeStart, setSwipeStart] = useState<{ x: number; y: number } | null>(null);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(METRO_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestScore(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (game.status !== "running") return;
    let animationFrame = 0;
    let previousTime = window.performance.now();

    function tick(now: number) {
      const delta = Math.min((now - previousTime) / 1000, 0.034);
      previousTime = now;
      setGame((current) => {
        if (current.status !== "running") return current;

        const level = getArcadeLevel(current.score, METRO_LEVEL_2_SCORE, METRO_LEVEL_3_SCORE);
        const speed = level === 1 ? 39 : level === 2 ? 51 : 66;
        const action = current.actionUntil > now ? current.action : "run";
        const movedObstacles = current.obstacles.map((obstacle) => ({
          ...obstacle,
          y: obstacle.y + speed * delta,
        }));
        const collided = movedObstacles.some((obstacle) => {
          const sameLane = obstacle.lane === current.lane;
          const inCollisionZone = obstacle.y >= 75 && obstacle.y <= 98;
          const avoidedByAction =
            (obstacle.kind === "barrier" && action === "jump") ||
            (obstacle.kind === "gate" && action === "slide");
          return sameLane && inCollisionZone && !avoidedByAction;
        });

        if (collided) {
          return { ...current, status: "over", action, obstacles: movedObstacles };
        }

        const passed = movedObstacles.filter((obstacle) => obstacle.y > 112).length;
        const score = current.score + passed;
        let obstacles = movedObstacles.filter((obstacle) => obstacle.y <= 112);
        let spawnIn = current.spawnIn - delta;
        let nextObstacleId = current.nextObstacleId;
        if (spawnIn <= 0) {
          const nextLevel = getArcadeLevel(score, METRO_LEVEL_2_SCORE, METRO_LEVEL_3_SCORE);
          const wave = createMetroWave(score, nextObstacleId);
          obstacles = [...obstacles, ...wave.obstacles];
          nextObstacleId = wave.nextObstacleId;
          spawnIn = metroSpawnDelay(nextLevel);
        }

        return { ...current, action, obstacles, spawnIn, nextObstacleId, score };
      });
      animationFrame = window.requestAnimationFrame(tick);
    }

    animationFrame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(animationFrame);
  }, [game.status]);

  useEffect(() => {
    if (game.status !== "over" || game.score <= bestScore) return;
    window.localStorage.setItem(METRO_BEST_KEY, String(game.score));
    const timer = window.setTimeout(() => setBestScore(game.score), 0);
    return () => window.clearTimeout(timer);
  }, [bestScore, game.score, game.status]);

  useEffect(() => {
    if (game.status !== "running") return;
    function controlMetro(event: KeyboardEvent) {
      const handled = ["ArrowLeft", "KeyA", "ArrowRight", "KeyD", "ArrowUp", "KeyW", "Space", "ArrowDown", "KeyS"];
      if (!handled.includes(event.code)) return;
      event.preventDefault();
      if (event.code === "ArrowLeft" || event.code === "KeyA") {
        setGame((current) => moveMetroLane(current, -1));
      } else if (event.code === "ArrowRight" || event.code === "KeyD") {
        setGame((current) => moveMetroLane(current, 1));
      } else if (event.code === "ArrowDown" || event.code === "KeyS") {
        setGame((current) => setMetroAction(current, "slide"));
      } else {
        setGame((current) => setMetroAction(current, "jump"));
      }
    }
    window.addEventListener("keydown", controlMetro);
    return () => window.removeEventListener("keydown", controlMetro);
  }, [game.status]);

  function startMetro() {
    setGame(createMetroState("running"));
  }

  const playerLeft = metroLanePosition(game.lane, 100);
  const nextObstacle = [...game.obstacles]
    .filter((obstacle) => obstacle.y < 96)
    .sort((first, second) => second.y - first.y)[0];

  function obstacleActionLabel(kind: MetroObstacleKind) {
    if (kind === "barrier") return `↑ ${copy.jumpLabel}`;
    if (kind === "gate") return `↓ ${copy.slideLabel}`;
    return `↔ ${copy.switchLabel}`;
  }

  function handleMetroSwipe(endX: number, endY: number) {
    if (!swipeStart || game.status !== "running") {
      setSwipeStart(null);
      return;
    }
    const deltaX = endX - swipeStart.x;
    const deltaY = endY - swipeStart.y;
    setSwipeStart(null);
    if (Math.max(Math.abs(deltaX), Math.abs(deltaY)) < 24) return;
    if (Math.abs(deltaX) > Math.abs(deltaY)) {
      setGame((current) => moveMetroLane(current, deltaX < 0 ? -1 : 1));
    } else {
      setGame((current) => setMetroAction(current, deltaY < 0 ? "jump" : "slide"));
    }
  }

  return (
    <div className="taya-metro-shell">
      <ArcadeLevelBadge
        score={game.score}
        level2At={METRO_LEVEL_2_SCORE}
        level3At={METRO_LEVEL_3_SCORE}
        label={copy.level}
      />
      <div className="taya-arcade-hud">
        <span>{copy.score} <strong>{game.score}</strong></span>
        <span>{copy.best} <strong>{bestScore}</strong></span>
      </div>
      <div
        className={`taya-metro-stage is-${game.status} level-${getArcadeLevel(game.score, METRO_LEVEL_2_SCORE, METRO_LEVEL_3_SCORE)}`}
        role="application"
        aria-label="Taya Metro Rush"
        onPointerDown={(event) => setSwipeStart({ x: event.clientX, y: event.clientY })}
        onPointerUp={(event) => handleMetroSwipe(event.clientX, event.clientY)}
        onPointerCancel={() => setSwipeStart(null)}
      >
        <span className="taya-metro-skyline" aria-hidden="true" />
        <span className="taya-metro-track" aria-hidden="true">
          <i />
          <i />
          <i />
        </span>
        {nextObstacle && game.status === "running" && (
          <span className={`taya-metro-next is-${nextObstacle.kind}`}>
            <small>{copy.nextLabel}</small>
            <strong>{obstacleActionLabel(nextObstacle.kind)}</strong>
          </span>
        )}
        {game.obstacles.map((obstacle) => {
          const obstacleScale = 0.56 + Math.max(0, Math.min(116, obstacle.y + 16)) / 116 * 0.94;
          const obstacleLeft = metroLanePosition(obstacle.lane, obstacle.y);
          return (
            <span
              key={obstacle.id}
              className={`taya-metro-obstacle is-${obstacle.kind}`}
              style={{
                left: `${obstacleLeft}%`,
                top: `${obstacle.y}%`,
                transform: `translate(-50%, -50%) scale(${obstacleScale})`,
              }}
              aria-hidden="true"
            >
              {obstacle.kind === "train" ? <><b /><b /></> : obstacle.kind === "gate" ? <b /> : null}
              <em>{obstacleActionLabel(obstacle.kind)}</em>
            </span>
          );
        })}
        <span
          className={`taya-metro-player is-${game.action}`}
          style={{ left: `${playerLeft}%` }}
          aria-hidden="true"
        >
          <b>T</b>
        </span>
        {game.status !== "running" && (
          <div className="taya-arcade-overlay">
            <strong>{game.status === "over" ? copy.gameOver : "TAYA METRO RUSH"}</strong>
            {game.status === "over" && <em>{copy.score}: {game.score}</em>}
            <div className="taya-metro-instructions" aria-label="How to play">
              <span className="is-jump"><b>↑</b><small>{copy.jumpLabel}</small></span>
              <span className="is-slide"><b>↓</b><small>{copy.slideLabel}</small></span>
              <span className="is-switch"><b>↔</b><small>{copy.switchLabel}</small></span>
            </div>
            <button type="button" onClick={startMetro}>
              {game.status === "ready" ? copy.start : copy.restart}
            </button>
          </div>
        )}
      </div>
      <div className="taya-metro-controls" aria-label="Metro controls">
        <button type="button" aria-label={copy.left} onClick={() => setGame((current) => moveMetroLane(current, -1))}><b>←</b><small>{copy.switchLabel}</small></button>
        <button className="is-jump" type="button" aria-label={copy.jump} onClick={() => setGame((current) => setMetroAction(current, "jump"))}><b>↑</b><small>{copy.jumpLabel}</small></button>
        <button className="is-slide" type="button" aria-label={copy.slide} onClick={() => setGame((current) => setMetroAction(current, "slide"))}><b>↓</b><small>{copy.slideLabel}</small></button>
        <button type="button" aria-label={copy.right} onClick={() => setGame((current) => moveMetroLane(current, 1))}><b>→</b><small>{copy.switchLabel}</small></button>
      </div>
    </div>
  );
}

type StackBlock = {
  id: number;
  x: number;
  width: number;
};

type StackState = {
  status: "ready" | "running" | "over";
  blocks: StackBlock[];
  movingX: number;
  direction: 1 | -1;
  score: number;
};

type StackCopy = {
  start: string;
  gameOver: string;
  score: string;
  best: string;
  restart: string;
  drop: string;
  level: string;
};

const STACK_BEST_KEY = "tayaStackBestV1";

function createStackState(status: StackState["status"] = "ready"): StackState {
  return {
    status,
    blocks: [{ id: 0, x: 20, width: 60 }],
    movingX: 0,
    direction: 1,
    score: 0,
  };
}

function placeStackBlockState(current: StackState): StackState {
  if (current.status !== "running") return current;
  const previous = current.blocks[current.blocks.length - 1];
  const left = Math.max(previous.x, current.movingX);
  const right = Math.min(previous.x + previous.width, current.movingX + previous.width);
  const width = right - left;
  if (width <= 0.8) return { ...current, status: "over" };

  const nextScore = current.score + 1;
  const nextBlock = { id: current.blocks.length, x: left, width };
  const direction: 1 | -1 = nextScore % 2 === 0 ? 1 : -1;
  return {
    ...current,
    blocks: [...current.blocks, nextBlock],
    movingX: direction === 1 ? 0 : 100 - width,
    direction,
    score: nextScore,
  };
}

function TayaStack({ copy }: { copy: StackCopy }) {
  const [game, setGame] = useState<StackState>(() => createStackState());
  const [bestScore, setBestScore] = useState(0);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(STACK_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestScore(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (game.status !== "running") return;
    let animationFrame = 0;
    let previousTime = window.performance.now();

    function tick(now: number) {
      const delta = Math.min((now - previousTime) / 1000, 0.034);
      previousTime = now;
      setGame((current) => {
        if (current.status !== "running") return current;
        const width = current.blocks[current.blocks.length - 1].width;
        const level = getArcadeLevel(current.score, 5, 10);
        const speed =
          level === 1
            ? 27 + current.score * 1.4
            : level === 2
              ? 43 + (current.score - 5) * 1.8
              : Math.min(82, 61 + (current.score - 10) * 1.5);
        let movingX = current.movingX + current.direction * speed * delta;
        let direction = current.direction;
        if (movingX <= 0) {
          movingX = 0;
          direction = 1;
        } else if (movingX + width >= 100) {
          movingX = 100 - width;
          direction = -1;
        }
        return { ...current, movingX, direction };
      });
      animationFrame = window.requestAnimationFrame(tick);
    }

    animationFrame = window.requestAnimationFrame(tick);
    return () => window.cancelAnimationFrame(animationFrame);
  }, [game.status]);

  useEffect(() => {
    if (game.status !== "over" || game.score <= bestScore) return;
    window.localStorage.setItem(STACK_BEST_KEY, String(game.score));
    const timer = window.setTimeout(() => setBestScore(game.score), 0);
    return () => window.clearTimeout(timer);
  }, [bestScore, game.score, game.status]);

  useEffect(() => {
    if (game.status !== "running") return;
    function dropFromKeyboard(event: KeyboardEvent) {
      if (event.code !== "Space" && event.code !== "ArrowDown") return;
      event.preventDefault();
      setGame((current) => placeStackBlockState(current));
    }
    window.addEventListener("keydown", dropFromKeyboard);
    return () => window.removeEventListener("keydown", dropFromKeyboard);
  }, [game.status]);

  function startStack() {
    setGame(createStackState("running"));
  }

  function placeStackBlock() {
    setGame((current) => placeStackBlockState(current));
  }

  function handleStackAction() {
    if (game.status !== "running") {
      startStack();
      return;
    }
    placeStackBlock();
  }

  const visibleBlocks = game.blocks.slice(-8);
  const movingWidth = game.blocks[game.blocks.length - 1].width;

  return (
    <div className="taya-stack-shell">
      <button
        className={`taya-stack-stage is-${game.status}`}
        type="button"
        aria-label={game.status === "running" ? copy.drop : copy.start}
        onPointerDown={(event) => {
          event.preventDefault();
          handleStackAction();
        }}
        onClick={(event) => {
          if (event.detail === 0) handleStackAction();
        }}
      >
        <ArcadeLevelBadge score={game.score} level2At={5} level3At={10} label={copy.level} />
        <span className="taya-arcade-hud">
          <span>{copy.score} <strong>{game.score}</strong></span>
          <span>{copy.best} <strong>{bestScore}</strong></span>
        </span>
        <span className="taya-stack-ground" aria-hidden="true" />
        <span className="taya-stack-block is-base" style={{ left: "8%", bottom: "5%", width: "84%" }} aria-hidden="true" />
        {visibleBlocks.map((block, index) => (
          <span
            key={block.id}
            className={`taya-stack-block tone-${(block.id % 4) + 1}`}
            style={{ left: `${block.x}%`, bottom: `${9 + index * 9}%`, width: `${block.width}%` }}
            aria-hidden="true"
          />
        ))}
        {game.status === "running" && (
          <span
            className={`taya-stack-block is-moving tone-${((game.blocks.length + 1) % 4) + 1}`}
            style={{ left: `${game.movingX}%`, bottom: `${9 + visibleBlocks.length * 9}%`, width: `${movingWidth}%` }}
            aria-hidden="true"
          />
        )}
        {game.status !== "running" && (
          <span className="taya-arcade-overlay">
            <strong>{game.status === "over" ? copy.gameOver : "TAYA STACK"}</strong>
            {game.status === "over" && <em>{copy.score}: {game.score}</em>}
            <small>{game.status === "over" ? copy.restart : copy.start}</small>
          </span>
        )}
      </button>
      <button className="taya-arcade-action" type="button" onClick={handleStackAction}>
        {game.status === "running" ? copy.drop : copy.start}
      </button>
    </div>
  );
}

type Game2048State = {
  status: "ready" | "running" | "over";
  board: number[];
  score: number;
};

type Game2048Copy = {
  start: string;
  gameOver: string;
  score: string;
  best: string;
  restart: string;
  up: string;
  down: string;
  left: string;
  right: string;
  level: string;
};

type Game2048Direction = "up" | "down" | "left" | "right";

const GAME_2048_SIZE = 4;
const GAME_2048_BEST_KEY = "taya2048BestV1";

function get2048SpawnValue(board: number[]) {
  const highestTile = Math.max(0, ...board);
  const level = getArcadeLevel(highestTile, 128, 512);
  const roll = Math.random();

  if (level === 3) {
    if (roll < 0.08) return 8;
    if (roll < 0.36) return 4;
    return 2;
  }
  if (level === 2) return roll < 0.24 ? 4 : 2;
  return roll < 0.1 ? 4 : 2;
}

function add2048Tile(board: number[]) {
  const emptyIndexes = board
    .map((value, index) => (value === 0 ? index : -1))
    .filter((index) => index >= 0);
  if (emptyIndexes.length === 0) return board;

  const next = [...board];
  const targetIndex = emptyIndexes[randomBreakIndex(emptyIndexes.length)];
  next[targetIndex] = get2048SpawnValue(board);
  return next;
}

function create2048State(status: Game2048State["status"] = "ready"): Game2048State {
  const emptyBoard = Array.from({ length: GAME_2048_SIZE * GAME_2048_SIZE }, () => 0);
  return {
    status,
    board: add2048Tile(add2048Tile(emptyBoard)),
    score: 0,
  };
}

function merge2048Line(line: number[]) {
  const compact = line.filter((value) => value > 0);
  const merged: number[] = [];
  let gained = 0;

  for (let index = 0; index < compact.length; index += 1) {
    if (compact[index] === compact[index + 1]) {
      const value = compact[index] * 2;
      merged.push(value);
      gained += value;
      index += 1;
    } else {
      merged.push(compact[index]);
    }
  }

  return {
    line: [...merged, ...Array.from({ length: GAME_2048_SIZE - merged.length }, () => 0)],
    gained,
  };
}

function can2048Move(board: number[]) {
  if (board.some((value) => value === 0)) return true;

  for (let row = 0; row < GAME_2048_SIZE; row += 1) {
    for (let column = 0; column < GAME_2048_SIZE; column += 1) {
      const index = row * GAME_2048_SIZE + column;
      if (column < GAME_2048_SIZE - 1 && board[index] === board[index + 1]) return true;
      if (row < GAME_2048_SIZE - 1 && board[index] === board[index + GAME_2048_SIZE]) return true;
    }
  }

  return false;
}

function move2048(current: Game2048State, direction: Game2048Direction): Game2048State {
  if (current.status !== "running") return current;
  const nextBoard = Array.from({ length: GAME_2048_SIZE * GAME_2048_SIZE }, () => 0);
  let gained = 0;

  for (let lineIndex = 0; lineIndex < GAME_2048_SIZE; lineIndex += 1) {
    const original = Array.from({ length: GAME_2048_SIZE }, (_, offset) => {
      if (direction === "left" || direction === "right") {
        return current.board[lineIndex * GAME_2048_SIZE + offset];
      }
      return current.board[offset * GAME_2048_SIZE + lineIndex];
    });
    const workingLine = direction === "right" || direction === "down" ? [...original].reverse() : original;
    const merged = merge2048Line(workingLine);
    const resultLine = direction === "right" || direction === "down" ? [...merged.line].reverse() : merged.line;
    gained += merged.gained;

    resultLine.forEach((value, offset) => {
      if (direction === "left" || direction === "right") {
        nextBoard[lineIndex * GAME_2048_SIZE + offset] = value;
      } else {
        nextBoard[offset * GAME_2048_SIZE + lineIndex] = value;
      }
    });
  }

  const moved = nextBoard.some((value, index) => value !== current.board[index]);
  if (!moved) return can2048Move(current.board) ? current : { ...current, status: "over" };

  const boardWithNewTile = add2048Tile(nextBoard);
  return {
    status: can2048Move(boardWithNewTile) ? "running" : "over",
    board: boardWithNewTile,
    score: current.score + gained,
  };
}

function Taya2048({ copy }: { copy: Game2048Copy }) {
  const [game, setGame] = useState<Game2048State>(() => create2048State());
  const [bestScore, setBestScore] = useState(0);

  useEffect(() => {
    const saved = Number(window.localStorage.getItem(GAME_2048_BEST_KEY) || 0);
    const timer = window.setTimeout(() => setBestScore(Number.isFinite(saved) ? saved : 0), 0);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (game.score <= bestScore) return;
    window.localStorage.setItem(GAME_2048_BEST_KEY, String(game.score));
    const timer = window.setTimeout(() => setBestScore(game.score), 0);
    return () => window.clearTimeout(timer);
  }, [bestScore, game.score]);

  useEffect(() => {
    if (game.status !== "running") return;
    function control2048(event: KeyboardEvent) {
      const directionByKey: Record<string, Game2048Direction | undefined> = {
        ArrowUp: "up",
        KeyW: "up",
        ArrowDown: "down",
        KeyS: "down",
        ArrowLeft: "left",
        KeyA: "left",
        ArrowRight: "right",
        KeyD: "right",
      };
      const direction = directionByKey[event.code];
      if (!direction) return;
      event.preventDefault();
      setGame((current) => move2048(current, direction));
    }
    window.addEventListener("keydown", control2048);
    return () => window.removeEventListener("keydown", control2048);
  }, [game.status]);

  function start2048() {
    setGame(create2048State("running"));
  }

  const highestTile = Math.max(0, ...game.board);

  return (
    <div className="taya-2048-shell">
      <ArcadeLevelBadge score={highestTile} level2At={128} level3At={512} label={copy.level} />
      <div className="taya-2048-hud">
        <span>{copy.score} <strong>{game.score}</strong></span>
        <span>{copy.best} <strong>{bestScore}</strong></span>
      </div>
      <div className="taya-2048-board" role="img" aria-label={`${copy.score}: ${game.score}`}>
        {game.board.map((value, index) => (
          <span
            key={index}
            className={`taya-2048-tile ${value > 0 ? `tile-${Math.min(11, Math.log2(value))}` : "is-empty"}`}
            aria-hidden="true"
          >
            {value || ""}
          </span>
        ))}
        {game.status !== "running" && (
          <div className="taya-2048-overlay">
            <strong>{game.status === "over" ? copy.gameOver : "TAYA 2048"}</strong>
            {game.status === "over" && <em>{copy.score}: {game.score}</em>}
            <button type="button" onClick={start2048}>
              {game.status === "over" ? copy.restart : copy.start}
            </button>
          </div>
        )}
      </div>
      <div className="taya-2048-controls" aria-label="2048 controls">
        <button className="is-up" type="button" aria-label={copy.up} onClick={() => setGame((current) => move2048(current, "up"))}>↑</button>
        <button className="is-left" type="button" aria-label={copy.left} onClick={() => setGame((current) => move2048(current, "left"))}>←</button>
        <button className="is-down" type="button" aria-label={copy.down} onClick={() => setGame((current) => move2048(current, "down"))}>↓</button>
        <button className="is-right" type="button" aria-label={copy.right} onClick={() => setGame((current) => move2048(current, "right"))}>→</button>
      </div>
    </div>
  );
}

const breakGameCopy = {
  en: {
    open: "Take a tiny break",
    another: "Another game",
    random: "Random pause",
    done: "Nice. Take one breath, then return whenever you’re ready.",
    again: "Again",
    close: "Close game",
    bubble: "Pop bubble",
    odd: "Choose the different tile",
    number: "Choose number",
    card: "Turn over card",
    start: "Tap or press Space to start",
    jump: "Tap · Space · ↑ to jump",
    startButton: "START",
    jumpButton: "JUMP",
    gameOver: "Run complete",
    time: "Time",
    best: "Best",
    restart: "Tap to run again",
    snakeGameOver: "Snake stopped",
    snakeRestart: "Play again",
    score: "Score",
    up: "Move up",
    down: "Move down",
    left: "Move left",
    right: "Move right",
    quizTitle: "Quick check",
    quizHelp: "Solve one small question to unlock your break.",
    quizAnswer: "Your answer",
    quizContinue: "Unlock game",
    quizWrong: "Not quite — try again.",
    flapButton: "FLAP",
    dropButton: "DROP",
    level: "LEVEL",
    games: {
      bubbles: { title: "Bubble pause", help: "Pop the bubbles at your own pace.", progress: "bubbles popped" },
      odd: { title: "Soft focus", help: "Find the tile that is just a little different.", progress: "rounds found" },
      numbers: { title: "Number flow", help: "Tap the numbers from 1 to 9.", progress: "numbers cleared" },
      memory: { title: "Memory clouds", help: "Turn over the cards and find three pairs.", progress: "pairs found" },
      runner: { title: "Taya Runner", help: "Jump over obstacles. The longer you run, the faster it gets.", progress: "running time" },
      snake: { title: "Taya Snake", help: "Collect the fruit without hitting the wall or yourself.", progress: "score" },
      flappy: { title: "Taya Flappy", help: "Flap through the gaps. Each pipe makes the next one faster.", progress: "score" },
      stack: { title: "Taya Stack", help: "Drop each moving block as neatly as you can.", progress: "blocks" },
      "2048": { title: "Taya 2048", help: "Combine matching tiles. Reach 128 and 512 to unlock harder levels.", progress: "score" },
    },
  },
  ja: {
    open: "ひと休みする",
    another: "別のゲーム",
    random: "ランダム休憩",
    done: "おつかれさまです。ひと呼吸したら、いつでも戻れます。",
    again: "もう一度",
    close: "ゲームを閉じる",
    bubble: "泡をポップ",
    odd: "違うタイルを選ぶ",
    number: "数字を選ぶ",
    card: "カードをめくる",
    start: "タップまたはSpaceでスタート",
    jump: "タップ・Space・↑でジャンプ",
    startButton: "START",
    jumpButton: "JUMP",
    gameOver: "ラン終了",
    time: "タイム",
    best: "ベスト",
    restart: "タップでもう一度",
    snakeGameOver: "ゲーム終了",
    snakeRestart: "もう一度",
    score: "スコア",
    up: "上へ移動",
    down: "下へ移動",
    left: "左へ移動",
    right: "右へ移動",
    quizTitle: "クイッククイズ",
    quizHelp: "簡単な問題に答えるとゲームが開きます。",
    quizAnswer: "答え",
    quizContinue: "ゲームを開く",
    quizWrong: "もう一度お試しください。",
    flapButton: "FLAP",
    dropButton: "DROP",
    level: "LEVEL",
    games: {
      bubbles: { title: "バブル休憩", help: "気の向くままに、泡をタップしてください。", progress: "個の泡をポップ" },
      odd: { title: "やさしい集中", help: "少しだけ色の違うタイルを見つけてください。", progress: "ラウンド完了" },
      numbers: { title: "ナンバーフロー", help: "1から9まで順番にタップしてください。", progress: "個の数字をクリア" },
      memory: { title: "雲の記憶", help: "カードをめくって3組のペアを見つけてください。", progress: "組を発見" },
      runner: { title: "Taya Runner", help: "障害物をジャンプ。走り続けるほどスピードが上がります。", progress: "走行タイム" },
      snake: { title: "Taya Snake", help: "壁や自分にぶつからないようにフルーツを集めます。", progress: "スコア" },
      flappy: { title: "Taya Flappy", help: "タップして隙間を通過。進むほど速くなります。", progress: "スコア" },
      stack: { title: "Taya Stack", help: "動くブロックをできるだけきれいに積み上げます。", progress: "ブロック" },
      "2048": { title: "Taya 2048", help: "同じ数字を合体。128と512で難易度が上がります。", progress: "スコア" },
    },
  },
  zh_cn: {
    open: "休息一下",
    another: "换一个游戏",
    random: "随机小休息",
    done: "很好。深呼吸一下，准备好再继续。",
    again: "再来一次",
    close: "关闭小游戏",
    bubble: "戳破泡泡",
    odd: "选择不同的方块",
    number: "选择数字",
    card: "翻开卡片",
    start: "点击或按空格开始",
    jump: "点击 · 空格 · ↑ 跳跃",
    startButton: "开始",
    jumpButton: "跳跃",
    gameOver: "跑酷结束",
    time: "本次",
    best: "最佳",
    restart: "点击再跑一次",
    snakeGameOver: "游戏结束",
    snakeRestart: "再玩一次",
    score: "分数",
    up: "向上移动",
    down: "向下移动",
    left: "向左移动",
    right: "向右移动",
    quizTitle: "快速小测验",
    quizHelp: "答对一道简单题即可进入游戏。",
    quizAnswer: "你的答案",
    quizContinue: "进入游戏",
    quizWrong: "还差一点，请再试一次。",
    flapButton: "拍动",
    dropButton: "放下",
    level: "等级",
    games: {
      bubbles: { title: "泡泡休息", help: "慢慢戳破泡泡，放松一下。", progress: "个泡泡已戳破" },
      odd: { title: "轻松找不同", help: "找出颜色有一点点不同的方块。", progress: "轮已找到" },
      numbers: { title: "数字流", help: "按照1到9的顺序点击数字。", progress: "个数字已完成" },
      memory: { title: "云朵记忆", help: "翻开卡片，找出三组相同图案。", progress: "组已找到" },
      runner: { title: "Taya Runner", help: "跳过障碍物，坚持越久速度越快。", progress: "跑酷时间" },
      snake: { title: "Taya Snake", help: "吃到水果，同时不要撞墙或撞到自己。", progress: "分数" },
      flappy: { title: "Taya Flappy", help: "点击穿过缝隙，每通过一个障碍速度都会加快。", progress: "分数" },
      stack: { title: "Taya Stack", help: "看准时机，把移动的方块整齐叠起来。", progress: "方块" },
      "2048": { title: "Taya 2048", help: "合并相同数字，达到128和512后提高难度。", progress: "分数" },
    },
  },
  zh_tw: {
    open: "休息一下",
    another: "換一個遊戲",
    random: "隨機小休息",
    done: "很好。深呼吸一下，準備好再繼續。",
    again: "再來一次",
    close: "關閉小遊戲",
    bubble: "戳破泡泡",
    odd: "選擇不同的方塊",
    number: "選擇數字",
    card: "翻開卡片",
    start: "點擊或按空格開始",
    jump: "點擊 · 空格 · ↑ 跳躍",
    startButton: "開始",
    jumpButton: "跳躍",
    gameOver: "跑酷結束",
    time: "本次",
    best: "最佳",
    restart: "點擊再跑一次",
    snakeGameOver: "遊戲結束",
    snakeRestart: "再玩一次",
    score: "分數",
    up: "向上移動",
    down: "向下移動",
    left: "向左移動",
    right: "向右移動",
    quizTitle: "快速小測驗",
    quizHelp: "答對一道簡單題即可進入遊戲。",
    quizAnswer: "你的答案",
    quizContinue: "進入遊戲",
    quizWrong: "還差一點，請再試一次。",
    flapButton: "拍動",
    dropButton: "放下",
    level: "等級",
    games: {
      bubbles: { title: "泡泡休息", help: "慢慢戳破泡泡，放鬆一下。", progress: "個泡泡已戳破" },
      odd: { title: "輕鬆找不同", help: "找出顏色有一點點不同的方塊。", progress: "輪已找到" },
      numbers: { title: "數字流", help: "按照1到9的順序點擊數字。", progress: "個數字已完成" },
      memory: { title: "雲朵記憶", help: "翻開卡片，找出三組相同圖案。", progress: "組已找到" },
      runner: { title: "Taya Runner", help: "跳過障礙物，堅持越久速度越快。", progress: "跑酷時間" },
      snake: { title: "Taya Snake", help: "吃到水果，同時不要撞牆或撞到自己。", progress: "分數" },
      flappy: { title: "Taya Flappy", help: "點擊穿過縫隙，每通過一個障礙速度都會加快。", progress: "分數" },
      stack: { title: "Taya Stack", help: "看準時機，把移動的方塊整齊疊起來。", progress: "方塊" },
      "2048": { title: "Taya 2048", help: "合併相同數字，達到128和512後提高難度。", progress: "分數" },
    },
  },
  mn: {
    open: "Take a tiny break",
    another: "Another game",
    random: "Random pause",
    done: "Nice. Take one breath, then return whenever you’re ready.",
    again: "Again",
    close: "Close game",
    bubble: "Pop bubble",
    odd: "Choose the different tile",
    number: "Choose number",
    card: "Turn over card",
    start: "Tap or press Space to start",
    jump: "Tap · Space · ↑ to jump",
    startButton: "START",
    jumpButton: "JUMP",
    gameOver: "Run complete",
    time: "Time",
    best: "Best",
    restart: "Tap to run again",
    snakeGameOver: "Snake stopped",
    snakeRestart: "Play again",
    score: "Score",
    up: "Move up",
    down: "Move down",
    left: "Move left",
    right: "Move right",
    quizTitle: "Quick check",
    quizHelp: "Solve one small question to unlock your break.",
    quizAnswer: "Your answer",
    quizContinue: "Unlock game",
    quizWrong: "Not quite — try again.",
    flapButton: "FLAP",
    dropButton: "DROP",
    level: "LEVEL",
    games: {
      bubbles: { title: "Bubble pause", help: "Pop the bubbles at your own pace.", progress: "bubbles popped" },
      odd: { title: "Soft focus", help: "Find the tile that is just a little different.", progress: "rounds found" },
      numbers: { title: "Number flow", help: "Tap the numbers from 1 to 9.", progress: "numbers cleared" },
      memory: { title: "Memory clouds", help: "Turn over the cards and find three pairs.", progress: "pairs found" },
      runner: { title: "Taya Runner", help: "Jump over obstacles. The longer you run, the faster it gets.", progress: "running time" },
      snake: { title: "Taya Snake", help: "Collect the fruit without hitting the wall or yourself.", progress: "score" },
      flappy: { title: "Taya Flappy", help: "Flap through the gaps. Each pipe makes the next one faster.", progress: "score" },
      stack: { title: "Taya Stack", help: "Drop each moving block as neatly as you can.", progress: "blocks" },
      "2048": { title: "Taya 2048", help: "Combine matching tiles. Reach 128 and 512 to unlock harder levels.", progress: "score" },
    },
  },
} as const;

function BreakGame({ language }: { language: NaviLanguage }) {
  const [revealed, setRevealed] = useState(false);
  const [open, setOpen] = useState(false);
  const [gameKind, setGameKind] = useState<BreakGameKind>("runner");
  const [quiz, setQuiz] = useState<BreakQuiz>({ prompt: "2 + 3", answer: 5 });
  const [quizAnswer, setQuizAnswer] = useState("");
  const [quizPassed, setQuizPassed] = useState(false);
  const [quizError, setQuizError] = useState(false);
  const [bubbles, setBubbles] = useState<BreakBubble[]>(BREAK_BUBBLES);
  const [oddRound, setOddRound] = useState(0);
  const [oddIndex, setOddIndex] = useState(4);
  const [numberTiles, setNumberTiles] = useState([1, 2, 3, 4, 5, 6, 7, 8, 9]);
  const [nextNumber, setNextNumber] = useState(1);
  const [memoryCards, setMemoryCards] = useState<MemoryCard[]>(
    MEMORY_SYMBOLS.flatMap((symbol, index) => [
      { id: index * 2, symbol },
      { id: index * 2 + 1, symbol },
    ]),
  );
  const [flippedCards, setFlippedCards] = useState<number[]>([]);
  const [matchedSymbols, setMatchedSymbols] = useState<string[]>([]);
  const [memoryLocked, setMemoryLocked] = useState(false);
  const copy = breakGameCopy[language];
  const popped = BREAK_BUBBLES.length - bubbles.length;

  const gameProgress = {
    bubbles: { value: popped, max: BREAK_BUBBLES.length },
    odd: { value: oddRound, max: 5 },
    numbers: { value: Math.min(nextNumber - 1, 9), max: 9 },
    memory: { value: matchedSymbols.length, max: MEMORY_SYMBOLS.length },
    runner: { value: 0, max: 1 },
    snake: { value: 0, max: 1 },
    flappy: { value: 0, max: 1 },
    stack: { value: 0, max: 1 },
    "2048": { value: 0, max: 1 },
  }[gameKind];
  const isArcadeGame = ["runner", "snake", "flappy", "stack", "2048"].includes(gameKind);
  const isComplete = !isArcadeGame && gameProgress.value >= gameProgress.max;

  useEffect(() => {
    const timer = window.setTimeout(() => setRevealed(true), 9000);
    return () => window.clearTimeout(timer);
  }, []);

  useEffect(() => {
    if (!open) return;
    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") setOpen(false);
    }
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [open]);

  function resetGame(kind = gameKind) {
    setBubbles([...BREAK_BUBBLES]);
    setOddRound(0);
    setOddIndex(randomBreakIndex(9));
    setNumberTiles(shuffled([1, 2, 3, 4, 5, 6, 7, 8, 9]));
    setNextNumber(1);
    setMemoryCards(createMemoryCards());
    setFlippedCards([]);
    setMatchedSymbols([]);
    setMemoryLocked(false);
    setGameKind(kind);
  }

  function openRunner() {
    resetGame("runner");
    setQuiz(createBreakQuiz());
    setQuizAnswer("");
    setQuizPassed(false);
    setQuizError(false);
    setOpen(true);
  }

  function checkQuizAnswer() {
    if (Number(quizAnswer.trim()) !== quiz.answer) {
      setQuizAnswer("");
      setQuizError(true);
      return;
    }

    setQuizPassed(true);
    setQuizError(false);
    resetGame("runner");
  }

  function selectOddTile(index: number) {
    if (index !== oddIndex || oddRound >= 5) return;
    const nextRound = oddRound + 1;
    setOddRound(nextRound);
    if (nextRound < 5) setOddIndex(randomBreakIndex(9));
  }

  function selectNumber(value: number) {
    if (value === nextNumber) setNextNumber((current) => current + 1);
  }

  function flipMemoryCard(index: number) {
    const card = memoryCards[index];
    if (
      memoryLocked ||
      flippedCards.includes(index) ||
      matchedSymbols.includes(card.symbol)
    ) return;

    const nextFlipped = [...flippedCards, index];
    setFlippedCards(nextFlipped);
    if (nextFlipped.length !== 2) return;

    setMemoryLocked(true);
    const [firstIndex, secondIndex] = nextFlipped;
    const isMatch = memoryCards[firstIndex].symbol === memoryCards[secondIndex].symbol;
    window.setTimeout(() => {
      if (isMatch) {
        setMatchedSymbols((current) => [...current, memoryCards[firstIndex].symbol]);
      }
      setFlippedCards([]);
      setMemoryLocked(false);
    }, isMatch ? 380 : 700);
  }

  function renderCurrentGame() {
    if (gameKind === "2048") {
      return (
        <Taya2048
          copy={{
            start: copy.startButton,
            gameOver: copy.snakeGameOver,
            score: copy.score,
            best: copy.best,
            restart: copy.snakeRestart,
            up: copy.up,
            down: copy.down,
            left: copy.left,
            right: copy.right,
            level: copy.level,
          }}
        />
      );
    }

    if (gameKind === "stack") {
      return (
        <TayaStack
          copy={{
            start: copy.startButton,
            gameOver: copy.snakeGameOver,
            score: copy.score,
            best: copy.best,
            restart: copy.snakeRestart,
            drop: copy.dropButton,
            level: copy.level,
          }}
        />
      );
    }

    if (gameKind === "flappy") {
      return (
        <TayaFlappy
          copy={{
            start: copy.start,
            startButton: copy.startButton,
            gameOver: copy.snakeGameOver,
            score: copy.score,
            best: copy.best,
            restart: copy.snakeRestart,
            flap: copy.flapButton,
            level: copy.level,
          }}
        />
      );
    }

    if (gameKind === "snake") {
      return (
        <TayaSnake
          copy={{
            start: copy.startButton,
            gameOver: copy.snakeGameOver,
            score: copy.score,
            best: copy.best,
            restart: copy.snakeRestart,
            up: copy.up,
            down: copy.down,
            left: copy.left,
            right: copy.right,
            level: copy.level,
          }}
        />
      );
    }

    if (gameKind === "runner") {
      return (
        <TayaRunner
          copy={{
            start: copy.start,
            jump: copy.jump,
            startButton: copy.startButton,
            jumpButton: copy.jumpButton,
            gameOver: copy.gameOver,
            time: copy.time,
            best: copy.best,
            restart: copy.restart,
            level: copy.level,
          }}
        />
      );
    }

    if (gameKind === "bubbles") {
      return bubbles.map((bubble, index) => (
        <button
          key={bubble.id}
          className={`break-bubble tone-${bubble.tone}`}
          type="button"
          aria-label={`${copy.bubble} ${index + 1}`}
          style={{
            left: `${bubble.x}%`,
            top: `${bubble.y}%`,
            width: bubble.size,
            height: bubble.size,
            animationDelay: `${bubble.delay}s`,
          }}
          onClick={() => setBubbles((current) => current.filter((item) => item.id !== bubble.id))}
        />
      ));
    }

    if (gameKind === "odd") {
      return (
        <div className={`break-odd-grid palette-${oddRound % 4}`}>
          {Array.from({ length: 9 }, (_, index) => (
            <button
              key={`${oddRound}-${index}`}
              className={index === oddIndex ? "is-odd" : ""}
              type="button"
              aria-label={`${copy.odd} ${index + 1}`}
              onClick={() => selectOddTile(index)}
            />
          ))}
        </div>
      );
    }

    if (gameKind === "numbers") {
      return (
        <div className="break-number-grid">
          {numberTiles.map((number) => (
            <button
              key={number}
              className={number < nextNumber ? "is-cleared" : ""}
              type="button"
              aria-label={`${copy.number} ${number}`}
              onClick={() => selectNumber(number)}
            >
              {number < nextNumber ? "✓" : number}
            </button>
          ))}
        </div>
      );
    }

    return (
      <div className="break-memory-grid">
        {memoryCards.map((card, index) => {
          const isFaceUp = flippedCards.includes(index) || matchedSymbols.includes(card.symbol);
          return (
            <button
              key={card.id}
              className={`${isFaceUp ? "is-face-up" : ""} ${matchedSymbols.includes(card.symbol) ? "is-matched" : ""}`}
              type="button"
              aria-label={`${copy.card} ${index + 1}`}
              aria-pressed={isFaceUp}
              onClick={() => flipMemoryCard(index)}
            >
              <span>{isFaceUp ? card.symbol : "?"}</span>
            </button>
          );
        })}
      </div>
    );
  }

  return (
    <>
      <button
        className={`break-game-trigger ${revealed ? "is-visible" : ""}`}
        type="button"
        aria-label={copy.open}
        title={copy.open}
        onClick={openRunner}
      >
        <span aria-hidden="true">✦</span>
      </button>

      {open && (
        <div className="break-game-backdrop" role="presentation" onMouseDown={() => setOpen(false)}>
          <section
            className={`break-game-card ${quizPassed && isArcadeGame ? "is-runner" : "is-quiz"}`}
            role="dialog"
            aria-modal="true"
            aria-labelledby="break-game-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            {!quizPassed ? (
              <div className="break-quiz-view">
                <div className="break-game-heading">
                  <div>
                    <span className="break-game-kicker">TAYA TOOL · QUICK CHECK</span>
                    <h2 id="break-game-title">{copy.quizTitle}</h2>
                    <p>{copy.quizHelp}</p>
                  </div>
                  <div className="break-game-heading-actions">
                    <button className="break-game-close" type="button" aria-label={copy.close} onClick={() => setOpen(false)}>×</button>
                  </div>
                </div>

                <form
                  className="break-quiz-form"
                  onSubmit={(event) => {
                    event.preventDefault();
                    checkQuizAnswer();
                  }}
                >
                  <div className="break-quiz-question" aria-live="polite">
                    <span>{quiz.prompt}</span>
                    <strong>= ?</strong>
                  </div>
                  <label htmlFor="break-quiz-answer">{copy.quizAnswer}</label>
                  <div className="break-quiz-answer-row">
                    <input
                      id="break-quiz-answer"
                      type="text"
                      inputMode="numeric"
                      pattern="-?[0-9]*"
                      autoComplete="off"
                      autoFocus
                      value={quizAnswer}
                      onChange={(event) => {
                        setQuizAnswer(event.target.value);
                        setQuizError(false);
                      }}
                    />
                    <button type="submit" disabled={!quizAnswer.trim()}>{copy.quizContinue}</button>
                  </div>
                  {quizError && <p className="break-quiz-error" role="alert">{copy.quizWrong}</p>}
                </form>
              </div>
            ) : (
              <>
            <div className="break-game-heading">
              <div>
                <span className="break-game-kicker">TAYA TOOL · GAME</span>
                <h2 id="break-game-title">{copy.games[gameKind].title}</h2>
                <p>{copy.games[gameKind].help}</p>
              </div>
              <div className="break-game-heading-actions">
                <button className="break-game-close" type="button" aria-label={copy.close} onClick={() => setOpen(false)}>×</button>
              </div>
            </div>

            <div className="break-game-tabs" role="tablist" aria-label="Choose game">
              <button
                className={gameKind === "runner" ? "is-active" : ""}
                type="button"
                role="tab"
                aria-selected={gameKind === "runner"}
                onClick={() => resetGame("runner")}
              >
                Runner
              </button>
              <button
                className={gameKind === "snake" ? "is-active" : ""}
                type="button"
                role="tab"
                aria-selected={gameKind === "snake"}
                onClick={() => resetGame("snake")}
              >
                Snake
              </button>
              <button
                className={gameKind === "flappy" ? "is-active" : ""}
                type="button"
                role="tab"
                aria-selected={gameKind === "flappy"}
                onClick={() => resetGame("flappy")}
              >
                Flappy
              </button>
              <button
                className={gameKind === "stack" ? "is-active" : ""}
                type="button"
                role="tab"
                aria-selected={gameKind === "stack"}
                onClick={() => resetGame("stack")}
              >
                Stack
              </button>
              <button
                className={gameKind === "2048" ? "is-active" : ""}
                type="button"
                role="tab"
                aria-selected={gameKind === "2048"}
                onClick={() => resetGame("2048")}
              >
                2048
              </button>
            </div>

            {!isArcadeGame && (
              <div className="break-game-progress" aria-live="polite">
                <span style={{ width: `${(gameProgress.value / gameProgress.max) * 100}%` }} />
                <small>{gameProgress.value} / {gameProgress.max} {copy.games[gameKind].progress}</small>
              </div>
            )}

            <div className={`break-game-board game-${gameKind} ${isComplete ? "is-complete" : ""}`}>
              {!isComplete && renderCurrentGame()}
              {isComplete && (
                <div className="break-game-done">
                  <span aria-hidden="true">☁</span>
                  <p>{copy.done}</p>
                  <div className="break-game-done-actions">
                    <button type="button" onClick={() => resetGame()}>{copy.again}</button>
                  </div>
                </div>
              )}
            </div>
              </>
            )}
          </section>
        </div>
      )}
    </>
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
}: {
  theme: "light" | "dark";
  language: NaviLanguage;
  onToggleTheme: () => void;
  onLanguageChange: (language: NaviLanguage) => void;
  onSelectTool: (tool: ToolView) => void;
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
      .map((expert) => formatSlackExpertForCombinedCopy(expert))
      .join("\n\n\n");
    const combinedHtml = experts
      .map((expert) => formatSlackExpertHtmlForCombinedCopy(expert))
      .join("<br>");
    await writeSlackClipboard(combined, `<div>${combinedHtml}</div>`);
    setCopiedId("all");
    window.setTimeout(() => setCopiedId(""), 1800);
  }

  async function copyExpertList() {
    await writeSlackClipboard(
      formatSlackExpertList(experts),
      formatSlackExpertListHtml(experts),
    );
    setCopiedId("expert-list");
    window.setTimeout(() => setCopiedId(""), 1800);
  }

  async function copyForCanvas() {
    await writeCanvasClipboard(formatSlackCanvas(experts));
    setCopiedId("canvas");
    window.setTimeout(() => setCopiedId(""), 1800);
  }

  function updateExpert(
    expertId: string,
    field: keyof Omit<SlackExpertRecord, "warnings">,
    value: string,
  ) {
    setExperts((current) =>
      current.map((expert) =>
        expert.id === expertId ? { ...expert, [field]: value } : expert,
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
              <div className="slack-results-actions">
                <button
                  className="button button-muted"
                  type="button"
                  onClick={copyExpertList}
                >
                  {copiedId === "expert-list"
                    ? t.copiedExpertList
                    : t.copyExpertList}
                </button>
                <button
                  className="button button-secondary"
                  type="button"
                  onClick={copyAllExperts}
                >
                  {copiedId === "all" ? t.copiedAll : t.copyAll}
                </button>
                <button
                  className="button button-primary slack-canvas-button"
                  type="button"
                  onClick={copyForCanvas}
                >
                  {copiedId === "canvas" ? t.copiedCanvas : t.copyCanvas}
                </button>
              </div>
            )}
          </div>

          {!experts.length ? (
            <div className="empty-state">{t.empty}</div>
          ) : (
            <div className="slack-expert-list">
              {experts.map((expert) => {
                const screening = expert.screeningText.trim();
                const history = slackHistoryItems(expert.employmentHistory);
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
                      <div className="slack-editor slack-screening-editor">
                        <p className="slack-editor-help">{t.editHelp}</p>
                        <details className="slack-edit-section" open>
                          <summary>{t.screeningEditorLabel}</summary>
                          <div className="slack-edit-section-body">
                            <label>
                              <span>{t.screeningLabel}</span>
                              <input
                                value={expert.screeningLabel}
                                onChange={(event) =>
                                  updateExpert(expert.id, "screeningLabel", event.target.value)
                                }
                              />
                            </label>
                            <label>
                              <span>{t.screeningEditorLabel}</span>
                              <textarea
                                value={expert.screeningText}
                                onChange={(event) =>
                                  updateExpert(expert.id, "screeningText", event.target.value)
                                }
                                spellCheck={false}
                              />
                            </label>
                          </div>
                        </details>

                        <details className="slack-edit-section">
                          <summary>{t.basicFieldsGroup}</summary>
                          <div className="slack-edit-section-body slack-edit-grid">
                            <label>
                              <span>{t.number}</span>
                              <input
                                value={expert.number}
                                onChange={(event) =>
                                  updateExpert(expert.id, "number", event.target.value)
                                }
                              />
                            </label>
                            <label>
                              <span>{t.name}</span>
                              <input
                                value={expert.name}
                                onChange={(event) =>
                                  updateExpert(expert.id, "name", event.target.value)
                                }
                              />
                            </label>
                            <label>
                              <span>{t.company}</span>
                              <input
                                value={expert.company}
                                onChange={(event) =>
                                  updateExpert(expert.id, "company", event.target.value)
                                }
                              />
                            </label>
                            <label className="slack-edit-wide">
                              <span>{t.titleField}</span>
                              <input
                                value={expert.title}
                                onChange={(event) =>
                                  updateExpert(expert.id, "title", event.target.value)
                                }
                              />
                            </label>
                          </div>
                        </details>

                        <details className="slack-edit-section">
                          <summary>{t.introductionGroup}</summary>
                          <div className="slack-edit-section-body">
                            <label>
                              <span>{t.introduction}</span>
                              <textarea
                                value={expert.introduction}
                                onChange={(event) =>
                                  updateExpert(expert.id, "introduction", event.target.value)
                                }
                              />
                            </label>
                          </div>
                        </details>

                        <details className="slack-edit-section">
                          <summary>{t.historyGroup}</summary>
                          <div className="slack-edit-section-body slack-edit-grid">
                            <label className="slack-edit-wide">
                              <span>{t.history}</span>
                              <textarea
                                value={expert.employmentHistory}
                                onChange={(event) =>
                                  updateExpert(expert.id, "employmentHistory", event.target.value)
                                }
                              />
                            </label>
                            <label className="slack-edit-wide">
                              <span>{t.availability}</span>
                              <textarea
                                value={expert.availability}
                                onChange={(event) =>
                                  updateExpert(expert.id, "availability", event.target.value)
                                }
                              />
                            </label>
                            <label>
                              <span>{t.location}</span>
                              <input
                                value={expert.location}
                                onChange={(event) =>
                                  updateExpert(expert.id, "location", event.target.value)
                                }
                              />
                            </label>
                            <label>
                              <span>{t.fee}</span>
                              <input
                                value={expert.fee}
                                onChange={(event) =>
                                  updateExpert(expert.id, "fee", event.target.value)
                                }
                              />
                            </label>
                          </div>
                        </details>
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
                          <div className="slack-history-grid">
                            {history.map(({ date, detail }, index) => (
                              <div className="slack-history-box" key={`${expert.id}-history-${index}`}>
                                {date && <span className="slack-history-date">{date}</span>}
                                {detail && <span className="slack-history-detail">{detail}</span>}
                              </div>
                            ))}
                          </div>
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
  const [sheetMode, setSheetMode] = useState<SheetMode>("custom");
  const [singleSheetName, setSingleSheetName] = useState("Expert List");
  const [customSheets, setCustomSheets] = useState(["Expert List"]);
  const [newSheetName, setNewSheetName] = useState("");
  const [draggedExpertId, setDraggedExpertId] = useState("");
  const [draggedFromSheet, setDraggedFromSheet] = useState("");
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
    const savedLanguage = window.localStorage.getItem(GLOBAL_LANGUAGE_KEY);
    const timer = window.setTimeout(() => {
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

  function switchWorkflowMode(nextMode: WorkflowMode) {
    setWorkflowMode(nextMode);
    setRaw("");
    setRecords([]);
    setExistingRecords([]);
    setComparisonItems([]);
    setRetainedCount(0);
    setImportedWorkbook(null);
    setSheetMode("custom");
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
            sheetNames: [sheet.name],
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

      const consolidatedRecords = consolidateImportedRecords(importedRecords);
      setExistingRecords(consolidatedRecords);
      setRecords(consolidatedRecords);
      setImportedWorkbook({ fileName: file.name, sheetNames: supportedSheetNames });
      setCustomSheets(supportedSheetNames);
      setSheetMode("custom");
      setComparisonItems([]);
      setRetainedCount(consolidatedRecords.length);
      setFileName(
        `${file.name.replace(/\.xlsx$/i, "") || "Expert_List"}_updated.xlsx`,
      );
      setMessage(
        `${t.uploadedExcel}: ${file.name} · ${consolidatedRecords.length} ${t.experts}`,
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
    setSheetMode("custom");
    setSingleSheetName("Expert List");
    setCustomSheets(["Expert List"]);
    setNewSheetName("");
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
    field: DataField,
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
      current.map((record) => {
        const sheetNames = getRecordSheetNames(record).map((name) =>
          name === currentName ? nextName : name,
        );
        return {
          ...record,
          sheetName:
            record.sheetName === currentName ? nextName : record.sheetName,
          sheetNames: [...new Set(sheetNames)],
        };
      }),
    );
  }

  function deleteCustomSheet(sheetName: string) {
    const hasExperts = records.some((record) =>
      getRecordSheetNames(record).includes(sheetName),
    );
    if (hasExperts || customSheets.length <= 1) return;
    setCustomSheets((current) => current.filter((name) => name !== sheetName));
  }

  function copyExpertToSheet(expertId: string, sheetName: string) {
    setRecords((current) =>
      current.map((record) => {
        if (record.id !== expertId) return record;
        const sheetNames = [
          ...new Set([...getRecordSheetNames(record), sheetName]),
        ];
        return { ...record, sheetName: sheetNames[0], sheetNames };
      }),
    );
    setDraggedExpertId("");
    setDraggedFromSheet("");
    setDragOverSheet("");
  }

  function moveExpertToSheet(
    expertId: string,
    sourceSheet: string,
    destinationSheet: string,
  ) {
    if (!sourceSheet || sourceSheet === destinationSheet) {
      setDraggedExpertId("");
      setDraggedFromSheet("");
      setDragOverSheet("");
      return;
    }
    setRecords((current) =>
      current.map((record) => {
        if (record.id !== expertId) return record;
        const withoutSource = getRecordSheetNames(record).filter(
          (name) => name !== sourceSheet,
        );
        const sheetNames = [...new Set([...withoutSource, destinationSheet])];
        return { ...record, sheetName: sheetNames[0], sheetNames };
      }),
    );
    setDraggedExpertId("");
    setDraggedFromSheet("");
    setDragOverSheet("");
  }

  function removeExpertFromSheet(expertId: string, sheetName: string) {
    setRecords((current) =>
      current.map((record) => {
        if (record.id !== expertId) return record;
        const currentSheets = getRecordSheetNames(record);
        if (currentSheets.length <= 1) return record;
        const sheetNames = currentSheets.filter((name) => name !== sheetName);
        return { ...record, sheetName: sheetNames[0], sheetNames };
      }),
    );
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
        customSheets,
      );
      const usedSheetNames = new Set<string>();
      const metaRows: string[][] = [];
      const resolvedSheetNames = new Map<string, string[]>();

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
        resolvedSheetNames.set(record.id, [
          ...new Set([...(resolvedSheetNames.get(record.id) ?? []), sheetName]),
        ]);
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
            (resolvedSheetNames.get(record.id) ?? getRecordSheetNames(record)).join(" / "),
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
      <>
        <TayaNaviPanel
          theme={theme}
          language={language}
          onToggleTheme={changeTheme}
          onLanguageChange={changeLanguage}
          onSelectTool={setActiveTool}
        />
        <BreakGame language={language} />
      </>
    );
  }

  if (activeTool === "slack") {
    return (
      <>
        <SlackFormatterPanel
          theme={theme}
          language={language}
          onToggleTheme={changeTheme}
          onLanguageChange={changeLanguage}
          onSelectTool={setActiveTool}
        />
        <BreakGame language={language} />
      </>
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
                        {sheetMode === "custom" &&
                          ` · ${getRecordSheetNames(record).join(" / ")}`}
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

          {records.length > 0 && sheetMode === "custom" && (
            <section className="sheet-organizer-inline" aria-labelledby="sheet-organizer-title">
              <div className="sheet-inline-header">
                <div>
                  <span className="eyebrow">TAYA TOOL</span>
                  <h3 id="sheet-organizer-title">{t.organizerTitle}</h3>
                  <p>{t.organizerHelp}</p>
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
              </div>

              <div className="sheet-board">
                {customSheets.map((sheetName) => {
                  const sheetExperts = records.filter((record) =>
                    getRecordSheetNames(record).includes(sheetName),
                  );
                  const availableExperts = records.filter(
                    (record) => !getRecordSheetNames(record).includes(sheetName),
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
                        if (draggedExpertId) {
                          moveExpertToSheet(
                            draggedExpertId,
                            draggedFromSheet,
                            sheetName,
                          );
                        }
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

                      <label className="sheet-copy-control">
                        <span>{t.copyExpert}</span>
                        <select
                          value=""
                          disabled={!availableExperts.length}
                          onChange={(event) => {
                            if (event.target.value) copyExpertToSheet(event.target.value, sheetName);
                          }}
                        >
                          <option value="">{t.selectExpert}</option>
                          {availableExperts.map((record) => (
                            <option key={record.id} value={record.id}>
                              {record.number} · {record.name}
                            </option>
                          ))}
                        </select>
                      </label>

                      <div className="sheet-drop-zone">
                        {sheetExperts.length ? (
                          sheetExperts.map((record) => {
                            const recordSheets = getRecordSheetNames(record);
                            const copyTargets = customSheets.filter(
                              (name) => !recordSheets.includes(name),
                            );
                            return (
                              <article
                                className={`drag-expert ${draggedExpertId === record.id ? "is-dragging" : ""}`}
                                draggable
                                key={record.id}
                                onDragStart={() => {
                                  setDraggedExpertId(record.id);
                                  setDraggedFromSheet(sheetName);
                                }}
                                onDragEnd={() => {
                                  setDraggedExpertId("");
                                  setDraggedFromSheet("");
                                  setDragOverSheet("");
                                }}
                              >
                                <span className="drag-handle" aria-hidden="true">⠿</span>
                                <div className="drag-expert-summary">
                                  <strong>{record.name || "Unnamed expert"}</strong>
                                  <small>{record.number} · {record.company}</small>
                                </div>
                                <div className="drag-expert-actions">
                                  <em>{t.dragHint}</em>
                                  <select
                                    className="copy-expert-action"
                                    value=""
                                    disabled={!copyTargets.length}
                                    aria-label={`${t.copyToSheet}: ${record.name}`}
                                    title={t.copyToSheet}
                                    onChange={(event) => {
                                      if (event.target.value) {
                                        copyExpertToSheet(record.id, event.target.value);
                                      }
                                    }}
                                  >
                                    <option value="">{t.copyAction}</option>
                                    {copyTargets.map((targetSheet) => (
                                      <option key={targetSheet} value={targetSheet}>
                                        {targetSheet}
                                      </option>
                                    ))}
                                  </select>
                                  {recordSheets.length > 1 && (
                                    <button
                                      className="remove-sheet-expert"
                                      type="button"
                                      title={t.removeFromSheet}
                                      aria-label={`${t.removeFromSheet}: ${record.name}`}
                                      onClick={() => removeExpertFromSheet(record.id, sheetName)}
                                    >
                                      ×
                                    </button>
                                  )}
                                </div>
                              </article>
                            );
                          })
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
            </section>
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

      <BreakGame language={language} />

    </main>
  );
}
