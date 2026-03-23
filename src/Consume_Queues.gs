/**
 * Consume_Queues.gs
 * 
 * 専用シートに存在する更新キューを消費して、更新を実行する関数。
 * キューを手動で構築すれば、任意のテスト用ブックに対して更新のテストが可能。
 * 詳しい使用方法はgoogle docsのドキュメントを参照のこと。
 * 
 * 新構成の導入は、枠線情報など、一部のシート構成に関する情報はシートごと差し替えなければ更新できないので、
 * テンプレートからシートをコピーしてきて、そのシートに古いバックアップから情報をコピーすることによって
 * 構成の移行を実現しています。
 * 
 * 更新途中にエラーが発生した場合にも、データはどこかのシートには残るようになっています。
 * 可能性は以下:
 * 
 * 1. 更新成功している（新しい構成のシートと、バックアップの両方にデータあり）（いまのところこれ以外経験していない）
 * 
 * 2. バックアップと古いシートだけあって新しいシートなし(両方にデータあり)（更新していないのとほぼ同じ。バックアップだけ最新のものになっているが、手動でシートの更新が必要。）
 * 
 * 3. バックアップだけあり（バックアップにデータあり。手動でシートの更新が必要。）
 * 
 * 4. バックアップと空の新しいシートだけあり（バックアップにデータあり。手動でデータ移行して、エラーの原因が通信障害などの不可避のものでないなら修正してもう一度API.gsから実行開始すればよい）
 * 
 * 5. 新しいシートにデータが移行されており、バックアップもあるが、数式が複数個所で#REF!になっている（移行は完了しているので、フォーマット修正マクロで数式をリフレッシュすれば完了します。）
 * 
 * 可能性は以上です。
 *  
 */
class ConsumeQueues{

  //更新キューを処理する。特定シートからキューを取得し、上の行から処理する。処理されたキューの行は削除。
  static process() {
    const lock = LockService.getScriptLock();
    let locked = false;
    try {
      lock.waitLock(1); //タイムロスを防ぐために最小の時間しか待たない
      locked = true;
      const started = Date.now();
      const msTimeLimit = started + macroProcessMsTimeLimit;
      const safetyBuffer = 15 * 1000;
      const safeMsTimeLimit = msTimeLimit - safetyBuffer;//安全バッファ付き時間制限の見込み

      const sheetUpdateQueSheet = updateQueueSheet;
      if (!sheetUpdateQueSheet) throw new Error("更新キューのシートが存在しません。");

      let msTimePerQue = 100;//これが最小値になる
      let processed = 0;
      let now = Date.now();
      while (sheetUpdateQueSheet.getLastRow() > 0) { // 値の入っている行を全部こなす
        const safeMsTimeLeft = safeMsTimeLimit - now;//残り時間安全バッファ適用済み
        if (safeMsTimeLeft < msTimePerQue) {
          CommonUtils.sendLog(`時間切れなので、続きのトリガーを発動します。: 残り ${safeMsTimeLeft}ms, 推定必要 ${msTimePerQue}ms\n今回の実行では${processed}件処理しました。`);
          this.makeTriggerToContinueUpdateSheet();
          return;
        }
        /*
        ここで一行ずつ（キューを一つずつ）取得しています。
        性能的にはキューをシート全体について一気に取得し、最後に処理した数だけ削除するほうが良いです。
        しかし、ここでは処理されたキューは確実に削除し、二重で更新をかけることを徹底的に防止することを優先しています。
        もしも、キューを複数まとめて削除するようにしてしまうと、途中でエラーが起こった際に、どこでエラーが発生したのかわかりにくくなり、さらに任意のタイミングで進捗を確認できなくなります。
        */

        //FIXME: キューの取得自体は一気にやって、deleteRowは今まで通りにすれば性能だけ改善できるのでそのほうが良いかもしれない。

        // キューをシートから取得
        const queue = sheetUpdateQueSheet.getRange(1, 1, 1, 2).getValues()[0];
        const studentBookUrl  = queue[0];
        const targetSheetName = queue[1];

        // 空行は削除して続行
        if (!studentBookUrl || !targetSheetName) {
          sheetUpdateQueSheet.deleteRow(1);
          continue;
        }

        if (!this.isUrl(studentBookUrl)) {
          throw new Error(`不正なURLがキューにあります: ${studentBookUrl}`);
        }

        //キュー実行
        const targetBook = SpreadsheetApp.openByUrl(studentBookUrl);
        const spManager = SheetIO.getSpeedPlannerIOManager(targetBook);
        this.updateStudentBook(spManager, targetSheetName);

        //実行出来たらキューを削除
        sheetUpdateQueSheet.deleteRow(1);

        //時間計測
        const newNow = Date.now();
        msTimePerQue = Math.max(msTimePerQue, newNow - now); // 最も時間のかかった処理時間
        now = newNow;
        processed++;

        CommonUtils.sendLog(`処理 ${processed} 件目: ${msTimePerQue}ms`);
      }
      CommonUtils.sendLog(`processUpdateQueues 完了: 合計 ${processed} 件`);

    } catch (e) {
      CommonUtils.sendLog("processUpdateQueues エラー: \n" + e);
      throw e;
    } finally {
      if (locked) lock.releaseLock();// ロックを持っている場合のみ解放
    }
  }

  //与えられたテキストが http または https の URL なら true を返すだけの関数
  static isUrl(text) {
    if (typeof text !== "string") return false;
    return /^https?:\/\/[^\s]+$/i.test(text.trim());
  }


  // 時間切れ時、再開用トリガーをセット
  static makeTriggerToContinueUpdateSheet() {
    const functionName = "processUpdateQueues";
    const triggers = ScriptApp.getProjectTriggers();
    // 既存の同関数ハンドラのトリガーを削除
    triggers.forEach(t => {
      if (t.getHandlerFunction && t.getHandlerFunction() === functionName) {
        ScriptApp.deleteTrigger(t); 
      }
    });
    // 1分後に発火（最短目安）
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .after(60 * 1000)
      .create();
  }



  /**
   * それぞれのキューを実行する単位処理の関数。一回呼ばれたら一つのキューを実行。
   */
  static updateStudentBook(spManager,sheetName) {
      try {
        const newBackUpSheet = this.replaceSheetWithTemplate(spManager, sheetName);//シートのテンプレートからのコピー
        this.copyBackupValuesExecutor(spManager, newBackUpSheet, sheetName);//旧シートのデータの移動
        this.refreshFormulas(spManager,sheetName);//数式リフレッシュ
        CommonUtils.sendLog(`${spManager.book.getName()}の「${sheetName}」を更新しました。`);
      } catch (e) {
        CommonUtils.sendLog(`${spManager.book.getName()}の更新中にエラー:\n${e}`);
      }
  }


  /**
   * テンプレートのシートで置き換えをする関数。
   * まず既存のシートのバックアップを取り、その後既存のシートをコピーしてデータ移行元のシートを作成する。
   * データ移行元のシートができれば、元のシートは削除し、新しいシートを元のシートと同じ名前で追加する。
   * データの移行は別の関数が行う。
   * 第1引数: コピー先のスプレッドシートID
   * 第2引数: 置き換えるシート名
   */
  static replaceSheetWithTemplate(spManager, sheetName) {
    
    const targetBook = spManager.book;

    const sourceSpreadsheet = srcBook;
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sourceSheet) throw new Error(`テンプレート側にシート「${sheetName}」がありません。`);

    // 置換対象シート
    const targetSheet = targetBook.getSheetByName(sheetName);
    if(!targetSheet){
      throw new Error(`エラー:\n更新対象のブック「${targetBook.getName()}」にシート「${sheetName}」がありません。\n古いシートのバックアップを作成できないので、このブックのこのシートの更新はスキップします。`);
    }

    //もともとの更新対象シートの位置を取得する。FIXME: なぜか常に1になる
    const insertPosition = (targetSheet)? (targetBook.getSheets().indexOf(targetSheet) + 1): 1;//デフォルト値1

    //コピー作って、元を消す。これをおえた段階でバックアップシートだけ存在している。
    //(ここでリネームしてしまうとこのシートを参照する数式も勝手にこちらを参照するように追従されてしまうのでこのようなやり方をしている)
    const newBackUpSheet = spManager.createBackUp(targetSheet);
    targetBook.deleteSheet(targetSheet);

    // テンプレートを対象のブックにコピーしてリネームする。
    //これをおえた時点で空の新構成シートとバックアップシートだけ存在している。
    const copiedSheet = sourceSheet.copyTo(targetBook);
    copiedSheet.setName(sheetName);

    // コピーしたシートを元の位置へ移動させる
    targetBook.setActiveSheet(copiedSheet);
    targetBook.moveActiveSheet(insertPosition);

    CommonUtils.sendLog(`置換完了：新しい「${sheetName}」を位置 ${insertPosition} に配置。いまからデータ移行を行います。`);
    return newBackUpSheet;
  }

  /**
   * _copyBackupValuesをよび、バックアップシートから新しいシートへのデータの移行を行う関数。
   * 実際の移行内容はConfigs.gsに定義されている。
   */
  static copyBackupValuesExecutor(spManager, backUpSheet, sheetName) {
    const ranges = dataMoveConfig[sheetName];
    if (!ranges) return; // 設定なしなら何もしない

    const targetSheet = spManager.book.getSheetByName(sheetName);
    if (!targetSheet) throw new Error(`コピー先シートが見つかりません: ${sheetName}`);

    ranges.forEach(r => {
      this._copyBackupValues(backUpSheet, targetSheet, r.src, r.dest);
    });
  }

  /**
   * 上のwrappperで呼び出している関数。
   * 旧シートのバックアップから新シートへのデータの移行を実行する関数である。
   * {sheetName}_src から新しい {sheetName} へ
   * 指定範囲の「値のみ」をコピーする（コピー先範囲も指定）
   */
  static _copyBackupValues(backupSheet, targetSheet, srcRangeText, destRangeText) {

    // コピー元範囲
    const srcRange  = backupSheet.getRange(srcRangeText);
    const srcRows   = srcRange.getNumRows();
    const srcCols   = srcRange.getNumColumns();

    // コピー先範囲
    const destRange = targetSheet.getRange(destRangeText);
    const destRows  = destRange.getNumRows();
    const destCols  = destRange.getNumColumns();

    // サイズ一致チェック
    if (!(srcRows === destRows && srcCols === destCols)) {
      throw new Error(`設定ミス：コピー元(${srcRows}x${srcCols}) と コピー先(${destRows}x${destCols}) のサイズが一致しません。`);
    }

    // 値をコピー
    const values = srcRange.getValues();
    destRange.setValues(values);

    CommonUtils.sendLog(`${targetSheet.getName()}の古いデータを${srcRangeText}から${destRangeText}に移行しました。`);
  }

  /**
   * 新しいシートに差し替え、データを移行しただけでは、もともとそのシートを参照していた別のシートの数式が#Ref!状態になるので、キャッシュをクリアするために数式を再セットする。
   */
  static refreshFormulas(spManager, sheetName) {
    const refreshTargetSheets = spManager.book.getSheets().filter(sh => {
      const name = sh.getName();
      // バックアップ用シートっぽいやつなどを除外。過剰に省いたとしても手動で数式の参照は直せる（リフレッシュするだけ）ので致命的な問題はない
      if (name.includes('_')) return false;
      if (name.includes('のコピー')) return false;
      if (name.includes('copy of')) return false;
      if (name.includes('Copy of')) return false;
      return true;
    });

    refreshTargetSheets.forEach(sh => {
      const r = sh.getDataRange();
      const formulas = r.getFormulas();
      let refreshed = false;

      for (let j = 0; j < formulas[0].length; j++) { // 数式が一列に並ぶ場合が多いので、列方向に最適化している。
        let blockStart = NaN;
        for (let i = 0; i <= formulas.length; i++) { // 最後にブロックをflushするため <=
          const f = i < formulas.length ? formulas[i][j] : null;

          // 対象の数式かどうか
          const isTarget = f && f.indexOf(sheetName) !== -1;

          if (isTarget) {
            if (Number.isNaN(blockStart)) blockStart = i; // ブロック開始
          } else {
            if (!Number.isNaN(blockStart)) {// ブロック終了 → まとめて setFormulas
              const blockLength = i - blockStart;
              const blockRange = sh.getRange(blockStart + 1, j + 1, blockLength, 1);
              const blockFormulas = formulas.slice(blockStart, i).map(row => [row[j]]);
              blockRange.setFormulas(blockFormulas);

              refreshed = true;
              blockStart = NaN; // 次のブロックへ
            }
          }
        }
      }
      if(refreshed) CommonUtils.sendLog(`refreshed formulas in: ${sh.getName()}`);
    });
  }


}