/**
 * このスクリプトの任意の場所から読んでよい、汎用性が高く、かつ仕様変更されにくい関数を定義するクラス。
 */
class CommonUtils{
  /**
   * 専用シートにログを残す関数。
   * トリガー実行でもログを「リアルタイムに」見るための関数。
   * リアルタイム性と、確実に一つずつ送信することを最優先にしているので、バッチ処理にはしていないです。
   */
  static sendLog(message){
    // 現在日時付きでログを作る
    const timestamp = new Date();
    const logEntry = [timestamp, message];
    // シートの一番下に追加
    logSheet.appendRow(logEntry);

    // コンソールにも出力
    Logger.log(`${timestamp.toISOString()} - ${message}`);
  }

}
