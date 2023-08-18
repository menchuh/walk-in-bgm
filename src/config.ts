type Configs = {
  sheet: string;
  duration: number;
};

/**
 * シートから値を取得する関数
 * @param spreadSheet
 * @param sheetname
 * @return Configs
 */
export function getConfigs(sheet: GoogleAppsScript.Spreadsheet.Sheet): Configs {
  const SOURCE_SHEET_NAME = sheet.getRange(3, 3).getValue();
  const BGM_DURATION = sheet.getRange(4, 3).getValue();

  if (!SOURCE_SHEET_NAME || !BGM_DURATION) {
    throw new Error('Config Values are do not be set.');
  }

  return {
    sheet: SOURCE_SHEET_NAME,
    duration: BGM_DURATION,
  };
}
