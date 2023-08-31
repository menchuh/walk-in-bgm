//===============================================
// BGMリストを適当に作成するスクリプト
// ・指定したシートの楽曲一覧からBGMリストをよしなに作成する
// ・BGMの長さなどはシート上で指定する
//===============================================
type Configs = {
  sheet: string;
  duration: number;
};

type Track = {
  origin_number: number;
  artist: string;
  title: string;
  album: string;
  duration_s: number;
};

function main() {
  try {
    // 定数
    const SHEET_NAME = 'BGM';

    // スプレッドシートの取得
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    // シートの取得（設定情報）
    const outputAndConfigSheet = spreadSheet.getSheetByName(SHEET_NAME);
    if (!outputAndConfigSheet) {
      throw Error(`The sheet ${SHEET_NAME} doses not exist. Error!`);
    }

    //======================
    // シートから設定を取得する
    //======================
    const configs = getConfigs(outputAndConfigSheet);

    //======================
    // シートから楽曲データを取得する
    //======================
    // シートの取得（楽曲リスト）
    const sourceSheet = spreadSheet.getSheetByName(configs.sheet);
    if (!sourceSheet) {
      throw new Error(`The sheet ${configs.sheet} does not exist.`);
    }

    // 楽曲データを取得
    const tracks = getTracks(sourceSheet);
    if (tracks.length === 0) {
      return null;
    }

    //======================
    // プレイリストを作成する
    //======================
    const playlist = getPlaylist(tracks, configs.duration);

    //======================
    // プレイリストをシートに出力する
    //======================
    saveTracks(outputAndConfigSheet, playlist);
  } catch (e) {
    Logger.log(e);
  }
}

/**
 * シートから値を取得する関数
 * @param spreadSheet
 * @param sheetname
 * @return Configs
 */
function getConfigs(sheet: GoogleAppsScript.Spreadsheet.Sheet): Configs {
  // 定数
  const SHEET_NAME_ROW = 3;
  const DURATION_ROW = 4;
  const DATA_COLUMN = 3;

  // データの取得
  const SOURCE_SHEET_NAME = sheet
    .getRange(SHEET_NAME_ROW, DATA_COLUMN)
    .getValue();
  const BGM_DURATION = sheet.getRange(DURATION_ROW, DATA_COLUMN).getValue();
  if (!SOURCE_SHEET_NAME || !BGM_DURATION) {
    throw new Error('Config Values are do not be set.');
  }

  return {
    sheet: SOURCE_SHEET_NAME,
    duration: BGM_DURATION,
  };
}

/**
 * シートから楽曲リストを取得する関数
 * @param sourceSheet GoogleAppsScript.Spreadsheet.Sheet
 * @return Track[]
 */
function getTracks(sourceSheet: GoogleAppsScript.Spreadsheet.Sheet): Track[] {
  // 定数
  const NO_TRACK_ROWS_NUMBER = 7;
  const START_COLUMN = 1;
  const END_COLUMN = 5;

  // データの最終行を取得
  const lastRow = getLastRow(sourceSheet, END_COLUMN);

  // 楽曲リストのデータを取得
  const tracks = sourceSheet
    .getRange(
      NO_TRACK_ROWS_NUMBER + 1,
      START_COLUMN,
      lastRow - NO_TRACK_ROWS_NUMBER,
      END_COLUMN
    )
    .getValues()
    .map((t) => {
      return {
        origin_number: t[0],
        artist: t[1],
        title: t[2],
        album: t[3],
        duration_s: t[4].getMinutes() * 60 + t[4].getSeconds(), // トラックの秒数
      };
    });

  return tracks;
}

/**
 * 楽曲リストから抜粋しプレイリストを作る関数
 * @param tracks Track[]
 * @param duration number
 * @returns Track[]
 */
function getPlaylist(tracks: Track[], duration: number): Track[] {
  let indexArray = [...Array(tracks.length)].map((_, i) => i);
  const playlistTracks = [];

  let isContinue = true;

  while (isContinue) {
    // ランダムでトラックを追加
    const randomIndex = Math.floor(Math.random() * indexArray.length);
    playlistTracks.push(tracks[indexArray[randomIndex]]);

    // 取得対象からそのトラックを除外
    indexArray = indexArray.filter((i) => {
      return i !== randomIndex;
    });

    // 秒数を加算し、秒数が上限を超えるかチェック
    const present_duration = playlistTracks.reduce(
      (sum, i) => sum + i.duration_s,
      0
    );
    if (present_duration > duration * 60) {
      isContinue = false;
    }
  }

  return playlistTracks;
}

function saveTracks(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  playlistTracks: Track[]
) {
  // 定数
  const START_ROW = 8;
  const START_COLUMN = 2;
  const END_COLUMN = 7;
  const TOTAL_DURATION_ROW = 5;
  const TOTAL_DURATION_COLUMN = 7;

  // シートをクリーン
  sheet
    .getRange(
      START_ROW,
      START_COLUMN,
      getLastRow(sheet, END_COLUMN),
      END_COLUMN
    )
    .clearContent();

  // 合計時間
  const totalDuration_s = playlistTracks.reduce(
    (sum, i) => sum + i.duration_s,
    0
  );
  const formattedTotalDuration = getFormattedDuration(totalDuration_s);

  // プレイリスト
  const data = playlistTracks.map((t, index) => {
    return [
      index + 1,
      t.origin_number,
      t.artist,
      t.title,
      t.album,
      getFormattedDuration(t.duration_s),
    ];
  });

  // 書き込み
  // 合計時間
  sheet
    .getRange(TOTAL_DURATION_ROW, TOTAL_DURATION_COLUMN)
    .setValue(formattedTotalDuration);

  // プレイリスト
  const range = sheet.getRange(
    START_ROW,
    START_COLUMN,
    data.length,
    data[0].length
  );
  range.setValues(data);
}

/**
 * フォーマッティングされた時間データを取得する関数
 * @param duration number
 * @return string
 */
function getFormattedDuration(duration: number): string {
  return `${zeroPadding(Math.floor(duration / 3600), 2)}:${zeroPadding(
    Math.floor(duration / 60),
    2
  )}:${zeroPadding(duration % 60, 2)}`;
}

function getLastRow(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  column: number
): number {
  const lastRow = sheet
    .getRange(sheet.getMaxRows(), column)
    .getNextDataCell(SpreadsheetApp.Direction.UP)
    .getRow();
  return lastRow;
}

/**
 * 与えられた数値をゼロパディングする関数
 * @param num 数値
 * @param len 桁数
 * @returns string
 */
function zeroPadding(num: number, len: number): string {
  return (Array(len).join('0') + num).slice(-len);
}
