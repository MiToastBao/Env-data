// export.js
// Builds an .xlsx that matches the government blank-template structure
// exactly: sheet 1 "監測點基本資料" (one header row + one data row),
// sheet 2 "<category>檢測項目" (header row + all monitoring data rows).

const ExportEngine = {
  buildWorkbook(project, basicInfo, categoryKey) {
    const cat = CATEGORIES[categoryKey];
    const rows = DataStore.getData(project.id, categoryKey);

    const wb = XLSX.utils.book_new();

    // Sheet 1: 監測點基本資料
    const basicHeaders = BASIC_INFO_FIELDS.map(f => f.key);
    const basicRow = BASIC_INFO_FIELDS.map(f => basicInfo[f.key] || '');
    const wsBasic = XLSX.utils.aoa_to_sheet([basicHeaders, basicRow]);
    XLSX.utils.book_append_sheet(wb, wsBasic, cat.basicSheetName);

    // Sheet 2: <category>檢測項目
    const dataHeaders = cat.fields.map(f => f.key);
    const dataRows = rows.map(r => cat.fields.map(f => r[f.key] || ''));
    const wsData = XLSX.utils.aoa_to_sheet([dataHeaders, ...dataRows]);
    XLSX.utils.book_append_sheet(wb, wsData, cat.dataSheetName);

    return wb;
  },

  downloadCategory(project, basicInfo, categoryKey) {
    const cat = CATEGORIES[categoryKey];
    const wb = this.buildWorkbook(project, basicInfo, categoryKey);
    const fname = `${project.code}_${cat.sourceFile}`;
    XLSX.writeFile(wb, fname);
  },

  downloadAll(project) {
    const basicInfo = DataStore.getBasicInfo(project.id);
    CATEGORY_ORDER.forEach(catKey => {
      const rows = DataStore.getData(project.id, catKey);
      if (rows.length > 0) {
        this.downloadCategory(project, basicInfo, catKey);
      }
    });
  },
};
