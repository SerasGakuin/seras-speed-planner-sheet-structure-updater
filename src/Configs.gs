//Configs.gs

const srcBook = SpreadsheetApp.openById("stub!"); // template 

const updateProjectBook = SpreadsheetApp.getActiveSpreadsheet();

const updateQueueSheet = updateProjectBook.getSheetByName('更新キュー');
const logSheet = updateProjectBook.getSheetByName('ログ');

const spreadsheetFirstRowNumber = 3;  // H3 からリンクの格納開始行
const spreadsheetColNumber = 8;       // H列　スピードプランナーのリンクの格納列

const macroProcessMsTimeLimit = 6*60*1000;//マクロ実行時間制限。ミリ秒単位。

/**
 * 更新対象のシートの名前。キューの構築に使用する。更新はキューを読んで行うので、本番実行時にはどのシートを更新するかを定義する役割を果たす。
 * 例: const sheetNames = ['今月プラン','月初']; 
 */
const sheetNames = ['月間実績'
]; 

/**
* 更新しない生徒のブックのurl一覧。更新するたびにデータ構造が変わる可能性があるので、二度同じブックの同じシートに対して同じアップデート（およびデータの移行）をすることをさけるために使う。
* 理由があって関数実行前に手動で更新した場合などにこの配列にurlを登録すればよい。
* 更新する生徒のブックのurlは生徒一覧シートの特定列から取得しているのだが、この配列に格納されたurlのブックは更新されなくなる。完全一致で判定するので必ず生徒一覧シートのurlをいれること。
* willNotUpdateBooksUrl =['url1'.'url2a']
*/
const willNotUpdateBooksUrl =[
]

/**　dataMoveConfig 詳細
 * 各種シートに対して、どのように古いシートから新しいシートへデータを移行するか定義するオブジェクト。
 * ここにシート名が登録されていたとしても、上の更新対象シート名配列に入っていなければ、更新は実行されない。
 * あくまで「更新する際に」どのようにデータを移行すればよいかを定義するためのものである。
 *
 * 構成
 * "シート名":[
 *  { src: '古いシート内の移行したいデータの範囲', dest: '新しいシートのデータ移行先の範囲'},
 *  ...
 * ],
 * ...
 * 
 * 例; 
 * "今月プラン": [
 *  { src: 'C1:C1', dest: 'C1:C1' },//何月度
 *  { src: 'A1:A30', dest: 'D1:D30'},//教材データ
 *  { src: 'D4:H30', dest: 'R4:R30'}//勉強時間
 * ]
 * この場合は「今月プラン」の古いシートのC1:C1のデータを新しいシートのC1:C1に、移行することになる。
 * 同様に、A1:A30をD1:D30に、D4:H30をR4:R30に。
 */

const dataMoveConfig = {
  "今月プラン": [
    { src: 'C1:C1', dest: 'C1:C1' }, // n月度テキスト
    { src: 'A1:A30', dest: 'A1:A30' }, // 教材データ
    { src: 'D4:H30', dest: 'D4:H30' }  // 勉強時間
  ],
  "週間管理": [
    { src: 'D1:D1', dest: 'D1:D1' },   // 最初の週の開始日
    { src: 'H4:H31', dest: 'H4:H31' }, // 週実績1
    { src: 'P4:P31', dest: 'P4:P31' }, // 週実績2
    { src: 'X4:X31', dest: 'X4:X31' }, // 週実績3
    { src: 'AF4:AF31', dest: 'AF4:AF31' }, // 週実績4
    { src: 'AN4:AN31', dest: 'AN4:AN31' }  // 週実績5
  ],
  "月初": [
    { src: 'A4:A31', dest: 'A4:A31' },   // 教材データ
    { src: 'C1:C1', dest: 'C1:C1' }, // n月度テキスト
    { src: 'BJ4:BN31', dest: 'BJ4:BN31' }, // 勉強時間
    { src: 'AA6:AA6', dest: 'AA6:AA6' }   // カレンダー基準日
  ],
  "今月実績": [
    { src: 'A4:A31', dest: 'A4:A31' }, // 教材データ
    { src: 'B1:B1', dest: 'B1:B1' }, // n月度テキスト
    { src: 'E4:I31', dest: 'D4:H31' }  // 勉強時間
  ]
};
