//===============================================
// BGMリストを適当に作成するスクリプト
// ・指定したシートの楽曲一覧からBGMリストをよしなに作成する
// ・BGMの長さなどはシート上で指定する
//===============================================
import { getConfigs } from './config';

function main() {
  try {
    // 定数
    const SHEET_NAME = 'BGM';

    // スプレッドシートの取得
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    // シートの取得
    const outputAndConfigSheet = spreadSheet.getSheetByName(SHEET_NAME);
    if (!outputAndConfigSheet) {
      throw Error(`The sheet ${SHEET_NAME} doses not exist. Error!`);
    }

    //======================
    // 1. シートから設定を取得する
    //======================
    const configs = getConfigs(outputAndConfigSheet);

    //======================
    // シートから楽曲データを取得する
    //======================
    Logger.log('aaa');

    //======================
    // プレイリストを作成する
    //======================
    Logger.log('bbb');

    //======================
    // プレイリストをシートに出力する
    //======================
    Logger.log('ccc');
  } catch (e) {
    Logger.log(e);
  }
}
