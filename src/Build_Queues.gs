/**
 * Build_Queues.gs
 * 更新キューを組み立てて、専用のシートにセットするクラス。
 * キューがすでにセットされている場合はセットを拒む。
 */
class BuildUpdateQueues {

  //更新キューの構築。これは実行しても何のブックも更新されない。生成されたキューはシート更新キューに登録される。
  static build() {
    const queuesSheet = updateQueueSheet;

    if (queuesSheet.getLastRow() > 0) {
      throw new Error('キューがすでに存在するので、新しくキューを構築できません。\nキューの処理中であれば、絶対にキューをいま更新してはいけませんが、なんらかの理由によってキューが放置されており、現在処理中でないのならば手動でキューシートをクリアしてください。\nキューが処理中でキューを再構築することを妨害する理由は、シートへの二重の更新の実行を防止するためです。シートのデータ移行が、新旧のシート構成に依存しているため、二重更新は危険になります。\nもし、エラーが起きて更新をかけなおそうという場合には、既存のキューをそのまま利用してください。そうすると、途中から処理が再開されます。ただし、エラーが起こったシートそれ自体については更新されない場合があります。その1シートについては、手動で修正してください。');
    }

    const bookUrlsArray = this._getStudentsBooksByUrl();
    if (bookUrlsArray.length === 0) {
      throw new Error("生徒のブックのurlがありません。誤った列をフィールド変数で登録しているかもしれません。");
    }
    if (sheetNames.length === 0) {
      throw new Error("更新対象のシート名の登録がありません。");
    }

    const valuesToSet = [];
    //キューデータを整形。　対象ブックurl × シート名　の直積
    for (let i = 0; i < bookUrlsArray.length; i++) {
      for (let j = 0; j < sheetNames.length; j++) {
        const row = [bookUrlsArray[i], sheetNames[j]];
        valuesToSet.push(row);
      }
    }
    CommonUtils.sendLog(`キューを作成しました。数:${valuesToSet.length}`);

    //キュー処理時に自動行削除をするので、このとき余分に行を追加して行数が0にならないようにする
    const emptyQueue = Array(2).fill(null);
    for (let j = 0; j < 10; j++) valuesToSet.push(emptyQueue);

    queuesSheet.getRange(1, 1, valuesToSet.length, valuesToSet[0].length).setValues(valuesToSet);
    CommonUtils.sendLog(`キューをセットしました。`);
  }

  // 「生徒マスター」から取得
  static _getStudentsBooksByUrl() {
    // 対象になりうるurlを収集
    const studentMaster = StudentMasterLib.getStudentMaster_V2();
    const allStudentData = studentMaster.getAllActiveStudentsDataRecordsArray();
    const allDefaultTargetUrls = allStudentData.map(rec => rec.speedPlannerUrl || '');

    // 除外URL → 除外IDのSetに変換
    const exclusionIdsSet = new Set(
      willNotUpdateBooksUrl
        .map(url => this._extractSpreadsheetId(url))
        .filter(id => id)
    );

    // URLごとにIDを見て除外
    const filteredTargetUrls = allDefaultTargetUrls.filter(url => {
      const id = this._extractSpreadsheetId(url);
      return id && !exclusionIdsSet.has(id);
    });

    // 重複排除
    return [...new Set(filteredTargetUrls)];
  }

  static _extractSpreadsheetId(url) {
    if (!url) return '';

    const match = String(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

}





