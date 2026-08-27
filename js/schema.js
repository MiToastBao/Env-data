// schema.js
// Defines the exact field structure for the 5 monitoring categories,
// matching the government template files field-for-field.

const BASIC_INFO_FIELDS = [
  { key: '計畫代碼', label: '計畫代碼', type: 'text', required: true,
    help: '格式為「英文代碼-四碼流水號」，例如 E-0001；如有延伸編碼請填 E-0001.1（請勿使用舊版三碼格式）' },
  { key: '書件案號', label: '書件案號', type: 'text',
    help: '共8碼：前7碼為數字，第8碼為英文大寫（若無可不填）' },
  { key: '書件名稱', label: '書件名稱', type: 'text', help: '請填寫完整書件名稱，勿使用簡易寫法' },
  // A leading blank, like every other select in this schema. Without it the browser
  // auto-selected 施工前 for a project that had never set the field, so the screen
  // showed 施工前 while the export wrote an empty cell.
  { key: '執行現況', label: '執行現況', type: 'select',
    options: ['', '施工前', '施工中', '施工兼營運', '營運中'] },
  { key: '施工日期', label: '施工日期', type: 'date', help: '執行現況為施工中/施工兼營運時請務必填寫' },
  { key: '竣工日期', label: '竣工日期', type: 'date', help: '執行現況為施工兼營運/營運中時請填寫' },
  { key: '營運日期', label: '營運日期', type: 'date', help: '執行現況為施工兼營運/營運中時請務必填寫' },
  { key: '備註', label: '備註', type: 'textarea', help: '100字以內' },
];

// 檢測類別 option lists per category (from 環境監測填寫範本 資料辭典)
const CATEGORY_TYPE_OPTIONS = {
  air: ['大氣環境', '固定污染源', '移動污染源', '周界空氣品質', '室內空氣品質', '光污染'],
  water: ['飲用水', '自來水', '河川', '水庫', '海域海灘', '地下水', '水下噪音-打樁', '水下噪音-風場範圍', '污廢水', '回收水', '雨水'],
  geo: ['土壤品質', '土砂觀測', '土壤', '底泥品質', '毒化物質', '廢棄物'],
  noise: ['環境噪音', '道路交通噪音', '營建工程噪音', '公私場所噪音', '航空噪音', '低頻噪音', '振動'],
};

// 比較關係 is documented as ">, <, ND" only, but real completed filings also use
// free-text markers like "未檢測" (equipment failure) etc., so this is a text field
// with suggestions rather than a strict enum.
const COMPARE_SUGGESTIONS = ['>', '<', 'ND', '未檢測'];

const CATEGORIES = {
  air: {
    key: 'air',
    locationField: '採樣地點',
    itemField: '檢測項目',
    methodField: '檢測方法',
    unitField: '檢測濃度/質量單位',
    label: '空氣品質',
    sourceFile: '空氣品質監測填寫.xlsx',
    dataSheetName: '空品檢測項目',
    basicSheetName: '監測點基本資料',
    fields: [
      { key: '日期(起)', label: '日期(起)', type: 'date', required: true },
      { key: '時間(起)', required: true, label: '時間(起)', type: 'time' },
      { key: '日期(迄)', required: true, label: '日期(迄)', type: 'date' },
      { key: '時間(迄)', required: true, label: '時間(迄)', type: 'time' },
      { key: '採樣地點', required: true, label: '採樣地點', type: 'text' },
      { key: '座標系統', required: true, label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' },
        help: '2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）' },
      { key: '採樣座標-經度 X', required: true, label: '採樣座標-經度 X', type: 'number' },
      { key: '採樣座標-緯度 Y', required: true, label: '採樣座標-緯度 Y', type: 'number' },
      { key: '場所編號', label: '場所編號', type: 'text' },
      { key: '採樣地點高度(公尺)', label: '採樣地點高度(公尺)', type: 'number' },
      { key: '污染物採樣高度(公尺)', label: '污染物採樣高度(公尺)', type: 'number' },
      { key: '管制編號', label: '管制編號', type: 'text' },
      { key: '煙道編號', label: '煙道編號', type: 'text' },
      { key: '檢測類別', label: '檢測類別', type: 'select', required: true, options: ['', ...CATEGORY_TYPE_OPTIONS.air] },
      { key: '檢測項目', required: true, label: '檢測項目', type: 'text' },
      { key: '檢測濃度/質量單位', requiredUnlessNote: true, label: '檢測濃度/質量單位', type: 'unitcode' },
      { key: '其他檢測濃度/質量單位', label: '其他檢測濃度/質量單位', type: 'text' },
      { key: '比較關係', label: '比較關係', type: 'suggest', options: COMPARE_SUGGESTIONS },
      { key: '檢測數值', requiredUnlessNote: true, label: '檢測數值', type: 'number' },
      { key: '檢測極限', label: '檢測極限', type: 'number' },
      { key: '檢測方法', requiredUnlessNote: true, label: '檢測方法', type: 'text' },
      { key: '檢測機構許可證號', label: '檢測機構許可證號', type: 'agencycode', required: true },
      { key: '其他檢測機構名稱', label: '其他檢測機構名稱', type: 'text' },
      { key: '備註', label: '備註', type: 'textarea' },
    ],
  },
  water: {
    key: 'water',
    locationField: '採樣地點',
    itemField: '檢測項目',
    methodField: '檢測方法',
    unitField: '檢測濃度/質量單位',
    label: '水質',
    sourceFile: '水質檢測資料填寫.xlsx',
    dataSheetName: '水質檢測項目',
    basicSheetName: '監測點基本資料',
    fields: [
      { key: '日期(起)', label: '日期(起)', type: 'date', required: true },
      { key: '時間(起)', required: true, label: '時間(起)', type: 'time' },
      { key: '日期(迄)', required: true, label: '日期(迄)', type: 'date' },
      { key: '時間(迄)', required: true, label: '時間(迄)', type: 'time' },
      { key: '採樣地點', required: true, label: '採樣地點', type: 'text' },
      { key: '座標系統', required: true, label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' },
        help: '2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）' },
      { key: '採樣座標-經度 X', required: true, label: '採樣座標-經度 X', type: 'number' },
      { key: '採樣座標-緯度 Y', required: true, label: '採樣座標-緯度 Y', type: 'number' },
      { key: '採樣深度(公尺)', label: '採樣深度(公尺)', type: 'number' },
      { key: '採樣水深(公尺)', label: '採樣水深(公尺)', type: 'number' },
      { key: '管制編號', label: '管制編號', type: 'text' },
      { key: '檢測類別', label: '檢測類別', type: 'select', required: true, options: ['', ...CATEGORY_TYPE_OPTIONS.water] },
      { key: '檢測項目', required: true, label: '檢測項目', type: 'text' },
      { key: '檢測濃度/質量單位', requiredUnlessNote: true, label: '檢測濃度/質量單位', type: 'unitcode' },
      { key: '其他檢測濃度/質量單位', label: '其他檢測濃度/質量單位', type: 'text' },
      { key: '比較關係', label: '比較關係', type: 'suggest', options: COMPARE_SUGGESTIONS },
      { key: '檢測數值', requiredUnlessNote: true, label: '檢測數值', type: 'number' },
      { key: '檢測極限', label: '檢測極限', type: 'number' },
      { key: '檢測方法', requiredUnlessNote: true, label: '檢測方法', type: 'text' },
      { key: '檢測機構許可證號', label: '檢測機構許可證號', type: 'agencycode', required: true },
      { key: '其他檢測機構名稱', label: '其他檢測機構名稱', type: 'text' },
      { key: '備註', label: '備註', type: 'textarea' },
    ],
  },
  geo: {
    key: 'geo',
    locationField: '採樣地點',
    itemField: '檢測項目',
    methodField: '檢測方法',
    unitField: '檢測濃度/質量單位',
    label: '地質',
    sourceFile: '地質檢測資料填寫.xlsx',
    dataSheetName: '地質檢測項目',
    basicSheetName: '監測點基本資料',
    fields: [
      { key: '日期(起)', label: '日期(起)', type: 'date', required: true },
      { key: '時間(起)', required: true, label: '時間(起)', type: 'time' },
      { key: '日期(迄)', required: true, label: '日期(迄)', type: 'date' },
      { key: '時間(迄)', required: true, label: '時間(迄)', type: 'time' },
      { key: '採樣地點', required: true, label: '採樣地點', type: 'text' },
      { key: '座標系統', required: true, label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' },
        help: '2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）' },
      { key: '採樣座標-經度 X', required: true, label: '採樣座標-經度 X', type: 'number' },
      { key: '採樣座標-緯度 Y', required: true, label: '採樣座標-緯度 Y', type: 'number' },
      { key: '採樣深度(公尺)', label: '採樣深度(公尺)', type: 'number' },
      { key: '管制編號', label: '管制編號', type: 'text' },
      { key: '檢測類別', label: '檢測類別', type: 'select', required: true, options: ['', ...CATEGORY_TYPE_OPTIONS.geo] },
      { key: '檢測項目', required: true, label: '檢測項目', type: 'text' },
      { key: '檢測濃度/質量單位', requiredUnlessNote: true, label: '檢測濃度/質量單位', type: 'unitcode' },
      { key: '其他檢測濃度/質量單位', label: '其他檢測濃度/質量單位', type: 'text' },
      { key: '比較關係', label: '比較關係', type: 'suggest', options: COMPARE_SUGGESTIONS },
      { key: '檢測數值', requiredUnlessNote: true, label: '檢測數值', type: 'number' },
      { key: '檢測極限', label: '檢測極限', type: 'number' },
      { key: '檢測方法', requiredUnlessNote: true, label: '檢測方法', type: 'text' },
      { key: '檢測機構許可證號', label: '檢測機構許可證號', type: 'agencycode', required: true },
      { key: '其他檢測機構名稱', label: '其他檢測機構名稱', type: 'text' },
      { key: '備註', label: '備註', type: 'textarea' },
    ],
  },
  noise: {
    key: 'noise',
    locationField: '監測地點',
    itemField: '音源發聲特性',
    methodField: '監測方法',
    unitField: '監測單位',
    label: '噪音',
    sourceFile: '噪音監測資料填寫.xlsx',
    dataSheetName: '噪音檢測項目',
    basicSheetName: '監測點基本資料',
    fields: [
      { key: '日期(起)', label: '日期(起)', type: 'date', required: true },
      { key: '時間(起)', required: true, label: '時間(起)', type: 'time' },
      { key: '日期(迄)', required: true, label: '日期(迄)', type: 'date' },
      { key: '時間(迄)', required: true, label: '時間(迄)', type: 'time' },
      { key: '監測地點', required: true, label: '監測地點', type: 'text' },
      { key: '座標系統', required: true, label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' },
        help: '2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）' },
      { key: '採樣座標-經度 X', required: true, label: '採樣座標-經度 X', type: 'number' },
      { key: '採樣座標-緯度 Y', required: true, label: '採樣座標-緯度 Y', type: 'number' },
      { key: '管制標準', required: true, label: '管制標準', type: 'select', options: ['', '工廠（場）噪音', '營建工程', '娛樂場所/營業場所', '擴音設施', '其他-1', '其他-2', '噪音管制法第7條第1項', '噪音管制法第15條第3項', '噪音管制法第14條第2項', '無'] },
      { key: '管制區', required: true, label: '管制區', type: 'select', options: ['', '第1類', '第2類', '第3類', '第4類', '無'] },
      { key: '環境音量標準', label: '環境音量標準', type: 'select', options: ['', '0', '1', '2'],
        optionLabels: { '0': '0-不適用', '1': '1-緊鄰未滿八公尺之道路', '2': '2-緊鄰八公尺以上之道路' } },
      { key: '頻率範圍', label: '頻率範圍', type: 'select', options: ['', '20 Hz 至 200 Hz', '20 Hz 至 20kHz'] },
      { key: '音源發聲特性', required: true, label: '音源發聲特性', type: 'select', options: ['', '最大音量(Lmax)', '均能音量(Leq)', '均能音量(Leq,LF)', '最大振動位準(Lvmax)', '事件振動位準(Lveq)', 'Lvd(10)', 'Lvn(10)'] },
      { key: '檢測類別', label: '檢測類別', type: 'select', required: true, options: ['', ...CATEGORY_TYPE_OPTIONS.noise] },
      { key: '監測時段', required: true, label: '監測時段', type: 'select', options: ['', '日間', '晚間', '夜間', '全天'] },
      { key: '監測數值', requiredUnlessNote: true, label: '監測數值', type: 'number' },
      { key: '監測單位', requiredUnlessNote: true, label: '監測單位', type: 'unitcode' },
      { key: '其他監測單位', label: '其他監測單位', type: 'text' },
      { key: '監測方法', requiredUnlessNote: true, label: '監測方法', type: 'text' },
      { key: '檢測機構許可證號', label: '檢測機構許可證號', type: 'agencycode', required: true },
      { key: '其他檢測機構名稱', label: '其他檢測機構名稱', type: 'text' },
      { key: '備註', label: '備註', type: 'textarea' },
    ],
  },
  eco: {
    key: 'eco',
    locationField: '調查地點',
    itemField: '調查項目',
    methodField: null,
    unitField: null,
    label: '生態',
    sourceFile: '生態調查資料填寫.xlsx',
    dataSheetName: '生態檢測項目',
    basicSheetName: '監測點基本資料',
    fields: [
      { key: '日期(起)', label: '日期(起)', type: 'date', required: true },
      { key: '時間(起)', required: true, label: '時間(起)', type: 'time' },
      { key: '日期(迄)', required: true, label: '日期(迄)', type: 'date' },
      { key: '時間(迄)', required: true, label: '時間(迄)', type: 'time' },
      { key: '調查地點', required: true, label: '調查地點', type: 'text' },
      { key: '座標系統', required: true, label: '座標系統', type: 'select', options: ['', '2', '3'],
        optionLabels: { '2': '2：WGS84（全球座標）', '3': '3：TWD97-TM2（投影座標系）' },
        help: '2：WGS84（全球座標，例如經度 120.681，緯度 24.147）／3：TWD97-TM2（投影座標系，例如 X=193150, Y=2670900）' },
      { key: '採樣座標-經度 X', required: true, label: '採樣座標-經度 X', type: 'number' },
      { key: '採樣座標-緯度 Y', required: true, label: '採樣座標-緯度 Y', type: 'number' },
      { key: '調查項目', required: true, label: '調查項目', type: 'text' },
      { key: '調查頻率', required: true, label: '調查頻率', type: 'text' },
      { key: '環境現況描述', required: true, label: '環境現況描述', type: 'textarea' },
      { key: '學名', required: true, label: '學名', type: 'text' },
      { key: '中文名', required: true, label: '中文名', type: 'text' },
      { key: '數量', required: true, label: '數量', type: 'text' },
      { key: '特有性', label: '特有性', type: 'select', options: ['', '特有種', '特有亞種'] },
      { key: '保育分類', label: '保育分類', type: 'select', options: ['', 'I', 'II', 'III'] },
      { key: '調查方法描述', label: '調查方法描述', type: 'textarea', required: true },
      { key: '檢測機構許可證號', label: '檢測機構許可證號', type: 'agencycode', required: true },
      { key: '其他檢測機構名稱', label: '其他檢測機構名稱', type: 'text' },
      { key: '備註', label: '備註', type: 'textarea' },
    ],
  },
};

const CATEGORY_ORDER = ['air', 'water', 'geo', 'noise', 'eco'];

// ── 振動的音源發聲特性：日間 Lvd(10)、夜間 Lvn(10) ──────────────────────────
//
// 官方「噪音資料辭典」第 21 列把振動的音源發聲特性列為四種：
// 最大振動位準(Lvmax)、事件振動位準(Lveq)、Lvd(10)、Lvn(10)。
// 其中 Lvd 的 d 是 day、Lvn 的 n 是 night——環境振動的 24 小時報告會分別算出
// 「Lv日(Lv10)」與「Lv夜(Lv10)」兩個值（報告上通常還會註明日間 6~20 時、
// 夜間 20~翌日 6 時），日間那一筆要填 Lvd(10)，夜間那一筆要填 Lvn(10)。
//
// v4.29 以前解析器對日、夜兩筆都填 Lvd(10)，夜間那一筆是錯的。
//
// ⚠️ 只適用於「環境振動 24 小時報告的 Lv10 彙整值」。
//    ・營建工程振動（BV）用的是事件振動位準(Lveq)／最大振動位準(Lvmax)，
//      那兩個沒有日夜之分，不可套用本規則。
//    ・晚間、全天沒有對應的官方代碼（振動只分日、夜兩個時段），一律不改，
//      維持使用者自己選的值。
const VIB_LV10_DAY = 'Lvd(10)';
const VIB_LV10_NIGHT = 'Lvn(10)';

/** 依監測時段決定 Lv10 彙整值的音源發聲特性；只認得日間與夜間，其餘回傳 null。 */
function vibLv10ItemFor(timeSegment) {
  if (timeSegment === '日間') return VIB_LV10_DAY;
  if (timeSegment === '夜間') return VIB_LV10_NIGHT;
  return null;
}

/**
 * 把「夜間 + Lvd(10)」這個**錯誤組合**正規化成 Lvn(10)，其餘一律原樣回傳。
 *
 * 這支專門給「比對用的識別碼」使用（跨季測站記憶、缺少測項比對、重複列偵測），
 * 讓 v4.29 以前存下來的舊資料和本版新匯入的資料被視為同一個測項。少了它會發生：
 *   ・舊記憶裡的「Lvd(10)::夜間」在新匯入時看起來像「這個測項不見了」，
 *     而缺少測項的建議是**預設打勾**的，使用者不取消就會多出一筆空白列；
 *   ・重新匯入同一份報告時，新的 Lvn(10) 找不到對應的舊列，會被當成全新資料附加，
 *     同一天同測站就同時存在 Lvd(10) 與 Lvn(10) 兩筆夜間振動。
 * 它只動「夜間的 Lvd(10)」這一個組合——依定義那就是錯的，所以不會誤傷任何正確資料。
 */
function canonicalVibItemName(itemName, timeSegment) {
  return itemName === VIB_LV10_DAY && timeSegment === '夜間' ? VIB_LV10_NIGHT : itemName;
}

// ── 噪音的「監測數值」固定顯示兩位小數 ──────────────────────────────────────
//
// 官方資料辭典對「監測數值」的說明是「請填數值，小數點2位數」。
// 這裡做的是**補零**，不是提高精度：報告上印的是 39.2（檢測報告的儲存格格式
// 就是 0.0），所以填的仍然是 39.2，只是寫成 39.20，和報告對得起來。
// 非數值（空白、ND、文字說明）一律原樣保留。
const NOISE_VALUE_FIELD = '監測數值';
const NOISE_VALUE_DECIMALS = 2;

/**
 * ⚠️ 這支只會**補零**，絕對不會減少位數。
 *
 * 一開始寫成 Number(v).toFixed(2)，那是錯的——它會把 67.235 變成 67.23、
 * 把 0.0006 變成 0.00，等於把實驗室印出來的位數砍掉。smartparse 的 formatNumber
 * 為了避免這件事特地寫了一整段註解說明「只用來削掉浮點雜訊，絕不低於原始位數」，
 * 這裡不能反過來把它抵銷掉。
 *
 * 規則：小數位數不足兩位就補到兩位，已經有兩位以上就原樣保留。
 * 非數值（空白、ND、<0.5、未檢測…）一律不動。
 */
function formatNoiseValue(v) {
  const s = String(v ?? '').trim();
  if (s === '') return s;
  const m = s.match(/^([+-]?)(\d*)(?:\.(\d*))?$/);
  if (!m || (m[2] === '' && (m[3] || '') === '')) return s; // ND、<0.5、未檢測、科學記號…原樣保留
  const sign = m[1] === '-' ? '-' : '';
  const intPart = m[2] === '' ? '0' : m[2];
  const decPart = m[3] || '';
  if (decPart.length >= NOISE_VALUE_DECIMALS) return `${sign}${intPart}.${decPart}`;
  return `${sign}${intPart}.${decPart.padEnd(NOISE_VALUE_DECIMALS, '0')}`;
}

/** 兩個欄位值是不是「同一個數字、只是寫法不同」（39.2 與 39.20、0.3 與 0.30）。 */
function sameNumericValue(a, b) {
  const na = String(a ?? '').trim(), nb = String(b ?? '').trim();
  if (na === nb) return true;
  if (na === '' || nb === '') return false;
  const plain = /^[+-]?(\d+(\.\d+)?|\.\d+)$/;
  if (!plain.test(na) || !plain.test(nb)) return false;
  return Number(na) === Number(nb);
}

/**
 * 「未知格式自動偵測」這條備援路徑抓到的是報告上的**原字**，例如 Lv日(Lv10)。
 * 那不是合法的音源發聲特性（官方只認四種代碼），而且該路徑不會填監測時段，
 * 交出去的兩個欄位都是錯的。這裡把原字翻成官方代碼並補上時段。
 *
 * 只認這四種明確寫著日／夜的振動標示，其餘（Lv5、Lv50、Leq…）一律不動，
 * 讓使用者在畫面上自己判斷——猜錯比留白更難發現。
 */
const VIB_RAW_LABEL_MAP = {
  'Lv日(Lv10)': { item: VIB_LV10_DAY, timeSegment: '日間' },
  'Lv夜(Lv10)': { item: VIB_LV10_NIGHT, timeSegment: '夜間' },
  'Lv日(Lveq)': { item: '事件振動位準(Lveq)', timeSegment: '日間' },
  'Lv夜(Lveq)': { item: '事件振動位準(Lveq)', timeSegment: '夜間' },
};

function normalizeVibRawLabel(rawItemName) {
  // 報告上半形、全形括號都有人用，順便把結尾的等號吃掉（「Lv日(Lv10)=」）。
  const key = String(rawItemName ?? '')
    .replace(/\s+/g, '')
    .replace(/[（]/g, '(').replace(/[）]/g, ')')
    .replace(/[=＝]+$/, '');
  return VIB_RAW_LABEL_MAP[key] || null;
}

// ── 必填欄位檢查（依 115 年版官方資料辭典）────────────────────────────────
//
// 上面每個欄位的 required / requiredUnlessNote 是照「環境監測填寫範本」各類別
// 資料辭典的「是否必填」欄逐項對出來的：
//   ・required            → 辭典寫「必填」
//   ・requiredUnlessNote  → 辭典寫「必/選填」，說明是
//                           「若因特殊情況無法檢測，請填寫備註，且可以不用填此欄」
//
// 另外兩條條件式規則，辭典也寫得很明確：
//   ・單位代碼填 160（其他）時，「其他…單位」變成必填
//   � 檢測機構許可證號填 AA 時，「其他檢測機構名稱」變成必填
//
// ⚠️ 114 年以前的舊格式有些欄位當時不是必填。這裡一律以**今年的辭典為準**，
//    但只是「提醒」，不會擋住匯入，也不會擋住匯出——使用者可以自己判斷。

/** 這一列少了哪些必填欄位？回傳 [{ key, label, why }]，沒有就是空陣列。 */
function missingRequiredFields(row, cat) {
  const blank = (k) => String(row[k] ?? '').trim() === '';
  const out = [];
  const noteFilled = !blank('備註');
  cat.fields.forEach((f) => {
    if (!blank(f.key)) return;
    if (f.required) out.push({ key: f.key, label: f.label, why: '必填' });
    else if (f.requiredUnlessNote && !noteFilled)
      out.push({ key: f.key, label: f.label, why: '必填（若確實無法檢測，請在備註說明原因）' });
  });
  // 條件式必填
  const unitField = cat.unitField;
  if (unitField && String(row[unitField] ?? '').trim() === '160') {
    const other = cat.fields.find((f) => /^其他.*單位$/.test(f.key));
    if (other && blank(other.key))
      out.push({ key: other.key, label: other.label, why: '單位代碼填 160 時必填' });
  }
  if (String(row['檢測機構許可證號'] ?? '').trim().split(';').some((x) => x.trim() === 'AA')
      && blank('其他檢測機構名稱')) {
    out.push({ key: '其他檢測機構名稱', label: '其他檢測機構名稱', why: '許可證號填 AA 時必填' });
  }
  return out;
}

/** 這一格是不是「該填而沒填」——給表格畫紅框用。 */
function isMissingRequiredCell(row, cat, fieldKey) {
  return missingRequiredFields(row, cat).some((m) => m.key === fieldKey);
}
