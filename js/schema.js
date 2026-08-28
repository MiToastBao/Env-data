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

/*
 * 一次採樣就是一個位置，所以這三欄永遠是一組動的——手動改一格會連帶同步、
 * 匯入時把噪音的座標帶給振動也是帶這三欄。
 *
 * ⚠️ 放在這裡是刻意的：v4.37 以前它只寫在 app.js 某個函式裡面，
 * 匯入那一側要用就得再抄一份。同一份清單抄兩遍，遲早有一邊漏掉「座標系統」，
 * 而漏掉之後畫面上完全看不出來——座標有了、座標系統空白，上傳才被退件。
 */
const COORD_FIELDS = ['座標系統', '採樣座標-經度 X', '採樣座標-緯度 Y'];

/*
 * 噪音 ↔ 振動：哪些欄位「必須各自不同」（v4.37）
 * ─────────────────────────────────────────────
 * 預設是**可以同步、會先問**——這樣以後多一個欄位不用再改一次程式。
 * 下面這些是例外：它們在噪音與振動之間本來就不一樣，同步過去就是填錯。
 *
 * 不是憑印象列的，是把使用者的 5 份實際報告讀出來逐欄比對的結果：
 *
 *   管制標準      噪音：營建工程／噪音管制法第7條第1項   振動：無
 *   管制區        噪音：第2類／第3類                   振動：無
 *     ↑ 這兩個是官方明文規定：振動必須填「無」、噪音不能填「無」。
 *       115年7月起加強檢核，同步過去等於在退件上面加一個「確定」按鈕。
 *   監測單位      噪音：16（dB(A)）                   振動：159（dB）
 *   監測方法      噪音：NIEA P201／P205                振動：NIEA P204
 *   音源發聲特性   噪音：Leq／Lmax／Leq,LF              振動：Lveq／Lvmax／Lvd(10)／Lvn(10)
 *   頻率範圍      噪音：20 Hz 至 20kHz／20 Hz 至 200 Hz  振動：空白
 *   環境音量標準   道路交通噪音專用，振動不適用
 *   監測時段      噪音有日／晚／夜三段，振動只有日／夜；而且它是每一列自己的屬性
 *   檢測類別      這就是用來分辨兩者的欄位，另有專屬的同步流程
 *   監測數值      每一列的測值本來就該不一樣
 *
 * 沒列在這裡的（例如備註、以後新增的欄位）一律走「先問再同步」。
 */
const NOISE_VIB_DISTINCT_FIELDS = [
  '管制標準', '管制區', '環境音量標準', '頻率範圍', '音源發聲特性',
  '檢測類別', '監測時段', '監測數值', '監測單位', '其他監測單位', '監測方法',
];

/*
 * 匯入時要從噪音帶給振動的欄位：報告的振動那一半通常整欄留白。
 * 座標三欄 ＋ 檢測機構（同一次到場、同一個位置、同一家檢測機構做的）。
 * 日期與時間不在這裡——它們是配對用的鍵，本來就已經一致。
 */
const NOISE_VIB_IMPORT_FILL_FIELDS = [...COORD_FIELDS, '檢測機構許可證號', '其他檢測機構名稱'];

/*
 * 生態的「調查項目」——建議清單，不是嚴格下拉。
 *
 * 來源：官方「環境監測填寫範本 115.8.18」的「總分類」工作表，生態污染別共 11 項，
 * 逐字照抄（v4.36 依此份範本更新；舊版簡報只列 8 項，且第 8 項在同一份簡報裡
 * 「其他動植物」與「維管束植物」兩種寫法並存——總分類寫的是維管束植物，
 * 而且根本沒有「其他動植物」這一項）。
 *
 * ⚠️ 這裡刻意用 'suggest'（可選也可打）而不是 'select'（只能選）：
 * 官方生態資料辭典第 18 列寫的是「請輸入文字 30 字以內。項目分類**可以請參考**
 * 工作表-"總分類"或紅皮書名錄」，而且它自己的填寫示範用了「裸子植物」——
 * 那不在總分類這 11 項裡面。鎖成下拉會讓合法的值打不進去。
 */
const ECO_SURVEY_ITEMS = [
  '魚類', '鳥類', '爬蟲類', '兩棲類', '蝶類', '哺乳類', '昆蟲類',
  '淡水魚類', '陸域爬行類', '陸域哺乳類', '維管束植物',
];

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
      { key: '調查項目', required: true, label: '調查項目', type: 'suggest', options: ECO_SURVEY_ITEMS, maxLength: 30,
        help: '可從清單選，也可以自己輸入（官方寫「可以參考」總分類或紅皮書名錄，範本自己的示範就用了清單外的「裸子植物」）。30 字以內。' },
      { key: '調查頻率', required: true, label: '調查頻率', type: 'text' },
      { key: '環境現況描述', required: true, label: '環境現況描述', type: 'textarea' },
      { key: '學名', required: true, label: '學名', type: 'text' },
      { key: '中文名', required: true, label: '中文名', type: 'text' },
      { key: '數量', required: true, label: '數量', type: 'text' },
      /*
       * 「單位」——v4.36 補上。官方 115.8.18 的空白表單「生態檢測項目」共 21 欄，
       * 這一欄夾在「數量」與「特有性」之間；舊版少了它，匯出只有 20 欄，
       * 欄位對不上就是整份退件。**位置不可以移動**，環境部是按欄序讀的。
       *
       * ⚠️ 這是自由文字，不是單位代碼。官方辭典第 23 列：「選填。若不可數，
       * 可填寫單位，例如：批、群；或覆蓋度：%」——填的是「群」「%」這種字，
       * 不是 1~161 的代碼。所以 eco 的 unitField 仍然維持 null。
       */
      { key: '單位', label: '單位', type: 'text',
        help: '選填。若不可數可填單位（例如：批、群），或覆蓋度（%）。這一欄填文字，不是單位代碼。' },
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
  /*
   * 「晚間」是**噪音**的時段分法（日／晚／夜三段）。振動參考日本振動規制法，
   * 只分日間與夜間兩段——官方的音源發聲特性也只有 Lvd(10)、Lvn(10) 兩個代碼，
   * 沒有晚間版本。
   *
   * 而報告上的「測定時間」是用噪音那套三段規則判的，所以營建工程若在
   * 18:00~21:59 之間量測，振動列會拿到「晚間」這個對振動不存在的時段。
   * 舊版在這裡回傳 null，呼叫端就整列跳過——Lv10 **靜靜不見**，畫面上沒有提示。
   *
   * 依使用者確認：**振動的晚間併入夜間**（日本振動規制法多數地區日間是
   * 08:00~19:00，19:00 之後轉夜間，比 22:00 才轉更接近實務）。
   */
  if (timeSegment === '晚間') return VIB_LV10_NIGHT;
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
  // 走 vibLv10ItemFor 而不是自己比對「夜間」，這樣「晚間也算夜間」那條規則
  // 只需要維護在一個地方——否則跨季比對會把同一個測項當成兩個。
  if (itemName !== VIB_LV10_DAY) return itemName;
  return vibLv10ItemFor(timeSegment) === VIB_LV10_NIGHT ? VIB_LV10_NIGHT : itemName;
}

// ── 小數位數設定（使用者可調） ──────────────────────────────────────────────
//
// 起因：官方辭典的規定會改。v4.33 以前是把「哪些欄位要兩位小數」寫死在程式裡，
// 規定一改就得改程式、重新發布。現在改成**使用者自己在畫面上設定**，
// 程式只提供「出廠預設值 ＝ 目前官方辭典的規定」。
//
// **出廠預設只有一項：噪音的「監測數值」補零到 2 位。其餘一律「不處理」（原樣）。**
//
// 為什麼不把辭典寫到的六個欄位全部設成預設：使用者只親自確認過噪音那一條。
// 程式替使用者決定申報數值要寫幾位，等於替他做了一個他沒同意的決定——
// 而數值的位數本身就是資訊（補成兩位＝宣稱實驗室量到小數第二位）。
// 辭典寫了什麼仍然完整保留在 OFFICIAL_DECIMAL_RULES，面板上逐欄顯示，
// 並提供一鍵「套用官方辭典建議」，要不要採用由使用者自己按。
//
// 我把 115 年版五份辭典的說明欄整份讀出來比對過，寫了小數位數的只有這六個：
//   噪音   監測數值              「請填數值，小數點2位數」  ← 出廠預設就是它
//   空氣   採樣地點高度(公尺)     「請填小數點後兩位。」
//   空氣   污染物採樣高度(公尺)   「請填小數點後兩位。」
//   水質   採樣深度(公尺)         「請填小數點後兩位。」
//   水質   採樣水深(公尺)         「請填小數點後兩位。」
//   地質   採樣深度(公尺)         「請填小數點後兩位。」
//
// 空氣／水質／地質的「檢測數值」辭典**沒有**規定，所以預設不處理。
// 硬套會製造假精度：空品報告 274 個值裡有 138 個原本只有 0～1 位小數。
//
// ⚠️ 兩種模式的差別很重要：
//   補零（pad）    ：只加尾零，**絕不減少位數**。0.125 在兩位設定下仍是 0.125。
//                    改的只是寫法，數字完全沒動。
//   四捨五入（round）：真的會減少位數。0.125 在兩位設定下變成 0.13。
//                    **但只影響畫面顯示與匯出的檔案，存起來的永遠是原始值**，
//                    所以設定改回去、或改成別的位數，隨時都能還原。
const DECIMAL_MODES = { off: '不處理', pad: '補零', round: '四捨五入' };
const DECIMAL_SETTINGS_KEY = 'envapp_decimal_settings_v1';
const DECIMAL_MAX_DIGITS = 6;

/** 115 年版官方辭典寫了小數位數的欄位。**只用來在面板上顯示建議**，不自動套用。 */
const OFFICIAL_DECIMAL_RULES = {
  'noise|監測數值': { mode: 'pad', digits: 2 },
  'air|採樣地點高度(公尺)': { mode: 'pad', digits: 2 },
  'air|污染物採樣高度(公尺)': { mode: 'pad', digits: 2 },
  'water|採樣深度(公尺)': { mode: 'pad', digits: 2 },
  'water|採樣水深(公尺)': { mode: 'pad', digits: 2 },
  'geo|採樣深度(公尺)': { mode: 'pad', digits: 2 },
};

/** 出廠預設。只有噪音的監測數值，其餘欄位一律不處理（原樣）。 */
const DEFAULT_DECIMAL_SETTINGS = {
  'noise|監測數值': { mode: 'pad', digits: 2 },
};

/*
 * 設定面板要列出哪些欄位——**只列會填數字的欄位**。
 * 全部欄位都列的話一個類別就二十幾列，五個類別加起來根本找不到要改的那一個。
 * 日期、時間、地點、代碼那些欄位設小數位數沒有意義。
 */
const DECIMAL_CONFIGURABLE_FIELDS = {
  air: ['檢測數值', '檢測極限', '採樣地點高度(公尺)', '污染物採樣高度(公尺)'],
  water: ['檢測數值', '檢測極限', '採樣深度(公尺)', '採樣水深(公尺)'],
  geo: ['檢測數值', '檢測極限', '採樣深度(公尺)'],
  noise: ['監測數值'],
  // 生態不列：它的「數量」是隻數／株數，本來就是整數，沒有小數位數的問題。
  eco: [],
};

const DECIMAL_OFF = { mode: 'off', digits: 2 };

function decimalSettingKey(catKey, fieldKey) { return `${catKey}|${fieldKey}`; }

/** 讀出整份設定（含出廠預設）。壞掉的內容一律忽略，不讓它把畫面弄爆。 */
function readDecimalSettings() {
  let stored = {};
  try {
    const raw = localStorage.getItem(DECIMAL_SETTINGS_KEY);
    if (raw) stored = JSON.parse(raw) || {};
  } catch { stored = {}; }
  const out = {};
  for (const [key, rule] of Object.entries({ ...DEFAULT_DECIMAL_SETTINGS, ...stored })) {
    if (!rule || typeof rule !== 'object') continue;
    const mode = Object.prototype.hasOwnProperty.call(DECIMAL_MODES, rule.mode) ? rule.mode : 'off';
    let digits = Number(rule.digits);
    if (!Number.isInteger(digits) || digits < 0 || digits > DECIMAL_MAX_DIGITS) digits = 2;
    out[key] = { mode, digits };
  }
  return out;
}

function writeDecimalSettings(settings) {
  localStorage.setItem(DECIMAL_SETTINGS_KEY, JSON.stringify(settings));
}

/** 這個類別的這個欄位，目前要怎麼呈現小數。沒設定過就是「不處理」。 */
function decimalRuleFor(catKey, fieldKey, settings) {
  const all = settings || readDecimalSettings();
  return all[decimalSettingKey(catKey, fieldKey)] || DECIMAL_OFF;
}

/**
 * 依規則把一個值格式化。
 *
 * ⚠️ 非數值（空白、ND、<0.5、未檢測、科學記號…）一律原樣保留，兩種模式都一樣。
 * ⚠️ pad 只會**補零**，絕對不會減少位數——這一點是刻意的。
 *    早期版本寫成 Number(v).toFixed(2)，會把 67.235 變成 67.23、0.0006 變成 0.00，
 *    等於把實驗室印出來的位數砍掉。要減少位數請明確選「四捨五入」。
 */
function formatDecimal(v, rule) {
  const s = String(v ?? '').trim();
  if (s === '' || !rule || rule.mode === 'off') return s;
  const m = s.match(/^([+-]?)(\d*)(?:\.(\d*))?$/);
  if (!m || (m[2] === '' && (m[3] || '') === '')) return s;
  const digits = rule.digits;
  if (rule.mode === 'round') {
    const n = Number(s);
    return Number.isFinite(n) ? n.toFixed(digits) : s;
  }
  const sign = m[1] === '-' ? '-' : '';
  const intPart = m[2] === '' ? '0' : m[2];
  const decPart = m[3] || '';
  if (decPart.length >= digits) return digits === 0 && decPart === '' ? `${sign}${intPart}` : `${sign}${intPart}${decPart ? '.' + decPart : ''}`;
  return `${sign}${intPart}.${decPart.padEnd(digits, '0')}`;
}

/** 給畫面／匯出用的一行式：查規則 ＋ 套用。 */
function formatFieldValue(catKey, fieldKey, value, settings) {
  return formatDecimal(value, decimalRuleFor(catKey, fieldKey, settings));
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
/*
 * 噪音／振動的連動規則檢查（v4.42）
 * ──────────────────────────────────
 * 官方規定（115 年操作說明 p.19，並自 115 年 7 月起加強檢核）：
 *   ・檢測類別＝振動 → 管制標準、管制區**必須**填「無」
 *   ・檢測類別＝噪音 → 管制標準、管制區**不能**填「無」
 *
 * 到 v4.41 為止，程式只做到「改的時候不讓你弄錯」——表格編輯、同步、批次修改
 * 三條路徑都擋住了。但**已經填錯的資料仍然沒有人告訴你**：
 * 別家公司交來的檔案、上一季匯入的舊資料、報告本身就填錯的，
 * 全部安安靜靜地躺在表格裡，一路送到環境部才被退件。
 * 這個函式補的就是「事後也看得出來」。
 *
 * ⚠️ 空白不在這裡處理。「管制標準」「管制區」本來就是必填欄位，
 * 空白由 missingRequiredFields() 負責——同一件事不要兩個地方各講一次，
 * 那只會讓真正該修的被淹沒。
 * ⚠️ 檢測類別本身空白時不判斷：連是噪音還是振動都不知道，猜就是亂報。
 */
const NOISE_VIB_LINKED_FIELDS = ['管制標準', '管制區'];
function noiseVibRuleViolations(row, cat) {
  if (!cat || cat.key !== 'noise') return [];
  const category = String(row['檢測類別'] ?? '').trim();
  if (category === '') return [];
  const isVib = category === '振動';
  const out = [];
  NOISE_VIB_LINKED_FIELDS.forEach((key) => {
    const v = String(row[key] ?? '').trim();
    if (v === '') return;
    if (isVib && v !== '無') {
      out.push({ key, label: key, value: v,
        why: `檢測類別是「振動」時，「${key}」必須填「無」（目前是「${v}」）` });
    } else if (!isVib && v === '無') {
      out.push({ key, label: key, value: v,
        why: `檢測類別是「${category}」時，「${key}」不能填「無」` });
    }
  });
  return out;
}

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
