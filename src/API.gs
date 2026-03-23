//API.gs
function safer_updateSpeedPlanner(){}//誤操作防止


//更新実行開始する関数。キューシートに書き込まれた内容をキューとして読み込み、更新を実行する。
function processUpdateQueues(){
  try{
    ConsumeQueues.process();
  }catch(e){
    CommonUtils.sendLog('エラー: '+(e.message || e));
    throw e;
  }
}

//キューを構築する関数。キューの処理中には呼ばないでください。
//生徒マスターのurl情報を参照しています。
//キューの処理中にタブを閉じてしまうと実行が中止されるようです。別のタブを開く分には問題ないようですが、、、

function buildSheetUpdateQueues(){
  try{
    BuildUpdateQueues.build();
  }catch(e){
    CommonUtils.sendLog('エラー: '+(e.message || e));
    throw e;
  }
}
