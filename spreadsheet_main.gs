/**
* 作成：2020.01.09　ライブラリ 
*
* 時刻起動トリガーが実行されると
* 残数をチェックして内容に合わせたメールを送信する。
* 紐付けシートの名称はデフォルトのままとする("フォームの回答 1")
* 受付メール送信記録シートの名前は"受付メール送信記録"とする。
* メール本文は"本文"シートとする、ユーザー手動作成
* 列0:タイムスタンプ 列1:アドレス 列2:名前　前提条件
* 残2以上でもフォームを開いている人が何にもいて送信するとインスタンス扱いのようなので全てその値で送信記録される。
* そのため全数という最大数設定が判断には必要
* 全数は途中で変更可能
*/

function fnConfirmMail(){

  var ss =SpreadsheetApp.getActiveSpreadsheet();
  
  //------------------------------------------------------
  // 受付メール送信記録
  //--------------------------------------------------------
  var MAILADD_SHEET = '受付メール送信記録';
  var subject = '';
  var sheets = ss.getSheets();
  var sheetexists = sheets.map(function(sh){
    return sh.getName();
  });
  
  if(sheetexists.indexOf(MAILADD_SHEET)<0){//指定名のシートがない場合
    //受付メール送信記録シート追加 独自関数呼び出し(common_sub)
    InsertSheet(MAILADD_SHEET,2);
    var shsent = ss.getSheetByName(MAILADD_SHEET);
    var mailrecordtitle = ['送信時間','メールアドレス','宛先名','可否'];
    var judgecol = 3;//'可否列
    mailrecordtitle.map(function(title,index){
      shsent.getRange(1,index+1).setValue(title);
    });
    var m = Moment.moment('2019/01/01 00:00:00'); //比較初期値
    var lasttime = m.format('YYYY/MM/DD HH:mm:ss');
    var mailrecordinitialdata = [m.format('YYYY/MM/DD HH:mm:ss'),'サンプル','サンプル','OK'];
    mailrecordinitialdata.map(function(title,index){
      shsent.getRange(2,index+1).setValue(title);
    });    
  }else{
    var shsent = ss.getSheetByName(MAILADD_SHEET);
  }
  var sentdatas = shsent.getDataRange().getValues();
  sentdatas =  sentdatas.filter(function(e){return e[0] !== "";});//空の要素を削除する
  var lastrow=sentdatas.length-1;
  var lasttime = sentdatas[lastrow][0];//前回送信実行した時刻 
  //--------------------------------------------------------
  // 本文
  //--------------------------------------------------------
  var MAILCONTENTS_SHEET = '本文';
  try{
    var sscontents= ss.getSheetByName(MAILCONTENTS_SHEET).getDataRange().getValues(); 
    //2行目:OK 3行目:NGの場合のメール内容
    var ok_contents = {judge:sscontents[1][0],subject:sscontents[1][1],content:sscontents[1][2]};
    var ng_contents = {judge:sscontents[2][0],subject:sscontents[2][1],content:sscontents[2][2]};
  }catch(e){
    var msg='「本文」シートがありません、作成願います';
    Browser.msgBox(msg);
    return
  }  
  //---------------------------------------------------------
  // フォームの回答 1
  //-----------------------------------------------------------
  var FORMRET_SHEET = 'フォームの回答 1';
  var ssdatas = ss.getSheetByName(FORMRET_SHEET).getDataRange().getValues(); 
  var ssform=ss.getSheetByName(FORMRET_SHEET);
  var addcol=1;//メールアドレスはB列、ただし配列データで指定するときは-1となる
  var namecol=2;//宛名はC列  
  var datatitles = ssdatas.splice(0, 1)[0];//タイトル項目行
  datatitles = datatitles.filter(Boolean);
  var colmax = datatitles.length+1;//可否列
  ssdatas = ssdatas.filter(function(e){return e[0] !== "";});//空の要素を削除する
  var ssdatas_length = ssdatas.length-1;   
  
  //回答データがないときは処理を終了
  if(ssdatas.length<1){
    Logger.log('回答データはありません、処理終了');
    return;
  }
  ////未送信のデータがない場合は朱里を終了
  if(ssdatas[ssdatas_length][0] < lasttime ){
    Logger.log('未送信データはありません、処理終了');
    return;
  }
  
  //回答データがある時は以下を実行
  var maxnum = {start:'［全 ',last:'］'}; //全角括弧開始、全の後半角空白 + 数値n+全角括弧閉じ
  var cntkey = {start:'（残 ',last:'）'};//全角括弧開始、残の後半角空白+数値n+全角括弧閉じ
  var cntoutnum = 4;//上記の文字列数 
  var confirmrows=[];// NG行  

  var indexlists=[];//ターゲット項目(日にち時間)列群 
  for(var i=0;i<ssdatas.length;i++){//
    for(var j=1;j<ssdatas[i].length;j++){
      var value = ssdatas[i][j] + '';//数値があったら文字化(数値はindexOfでエラーになる)
      var inumstart = value.indexOf(maxnum.start);//1文字は目は0番目 
      if(-1<inumstart){
        indexlists.push(j);//配列のindex番号リストをセット 
      }
    }
  }
  indexlists=uniq(indexlists);//重複削除 独自関数呼び出し(common_sub)
  Logger.log('全を含むリスト'+indexlists);//列番号リスト
  //---------------------------------------------
  //申込数が全num以上あるかどうかをチェック
  //----------------------------------------------\
  //項目(日にち)列抽出 ---------------------------- for A S
  for(var i=0;i<indexlists.length;i++){//列数(日にち)
  
    var col = parseInt(indexlists[i]);//項目列
    var targetcollist=[];//列群
    var targetobj={};
    var targetobjarr=[];
    var targetitems = {};
    var itemlists=[];

    //必要データを連想配列化
    for(var j=0;j<ssdatas.length;j++){//1列
      var targetcol={};//項目(セル)内容
      var data = ssdatas[j][col];//セルデータ
      var pos = data.indexOf(maxnum.start);
      if(0 < pos){
        targetcol.item=data.substr(0,pos);        
        Logger.log('位置は %s 内容は %s',pos,targetcol.item);
        targetcol.max=fnMaxNumber(data,maxnum);
        targetcol.rest=fnMaxNumber(data,cntkey);     
        targetcol.row=j;        
        targetcollist.push(targetcol); 
      }
    } 
  //---------------------------------
  // 最大数と申込数チェック
  //---------------------------------
    targetcollist.map(function (value,index) {
      var cat=value["item"];        
      if(typeof targetitems[cat]=="undefined"){//なければ作る
        targetitems[cat]=[];
        itemlists.push(cat);
      }
      targetitems[cat].push(value);
    });
    
    Logger.log('アイテムリスト %s',itemlists);
    for(var n=0;n<itemlists.length;n++){//'10時'などごと
      
      var cat=itemlists[n];
      //降順に並び替え　項目中最新の全数を採用する
      targetitems[cat].sort(function(a,b){
        if (a.row < b.row) {
          return 1;
        } else {
          return -1;
        }
      })
      var diffnum = targetitems[cat][0].max - targetitems[cat].length
      //Logger.log('最大 %s 申込数 %s 違い %s',targetitems[cat][0].max, targetitems[cat].length,diffnum);
      
      if(diffnum < 0){//全数より申込数が多い
        var nummax =Math.abs(diffnum);
        for(var m=0;m<nummax;m++){
          confirmrows.push(targetitems[cat][m].row); 
        }
      }      
    }

  }//グルーピング　項目ごと(10時　など) 日時毎 ---- for A E
  //----------------------------------------
  // リストが削除処理後に送信されたケース　時間内容が送信されていない=NG
  //----------------------------------------
  for(var j=0;j<ssdatas.length;j++){//1列(日にち時間)
    var cnt_null = 0;
    for(var i=0;i<indexlists.length;i++){//日にち列数 - for A S
      var col = indexlists[i];
      if(!ssdatas[j][col]){
        cnt_null++;    
      }
    }
    if(cnt_null === indexlists.length){
      //Logger.log('回数　%s 行 %s',cnt_null,j);
      confirmrows.push(j);
    }
  }// - for A E
  Logger.log('NGリスト %s ',confirmrows);
  //----------------------------------------
  //前回実行時間後のタイムスタンプ記録行のみ実行
  //----------------------------------------
  var maildatas=[];

  var irowr = sentdatas.length //「受付メール送信記録」シートへの書き込み行
  for(var i=0;i<ssdatas.length;i++){//データの数だけ ---------- for A S
  
    var irow = i+2;//レンジはタイトル行があり、1スタート
    if(lasttime < ssdatas[i][0]){//未送信のデータ---------if B S
      //初期値OK
      var toadd =  ssdatas[i][addcol];
      var strjudge = 'OK';
      ssform.getRange(irow,colmax).setValue(strjudge);  
      var resline = ok_contents;
      //該当の宛先にメール送信　＋ 書き込み
      //NGの場合のみ上書き
      if(0<confirmrows.length){
        for(var j=0;j<confirmrows.length;j++){
         if(i===confirmrows[j]){
            var strjudge = 'NG';
            var irowf = confirmrows[j]+2;//レンジはタイトル行があり、1スタート
            ssform.getRange(irowf,colmax).setValue(strjudge);//form回答シート          
            var resline = ng_contents;
          }
          
        }
      }
      var datatitle = reTitle(datatitles,ssdatas[i]);  //独自関数呼び出し(common_sub)
      var content_datas = reData(datatitle,ssdatas[i],maxnum); //独自関数呼び出し(common_sub)
      var str_name = ssdatas[i][namecol];//宛名
      var content = createTextMail(toadd,str_name,content_datas,resline); 
      irowr++;
      var m = Moment.moment(); //作成日時
      var senddate = m.format('YYYY/MM/DD HH:mm:ss');
      //['送信時間','メールアドレス','宛先名','可否']
      var values = [[senddate,toadd,str_name,strjudge]];
      shsent.getRange(irowr,1,values.length,values[0].length).setValues(values);//送信記録シート
      
    } //未送信のデータ-----------------------------------------------------if B E
  }//データの数だけ------------------------------------------------ for A E

}
/**
 * 
 * TEXT mail 作成
 * @param array ssdata フォーム回答シートの1レコード
 * @param array content 本文シートのOK／NG別の1レコード
 * ssdata=>列0:タイムスタンプ 列1:アドレス 列2:名前　前提条件
 * https://www.sejuku.net/blog/21812
 */
function createTextMail(toadd,str_name,ssdata,contents){

  var subject = contents.subject;
  var datas = ssdata.join('\n');//配列を改行で結合して文字列で返す
  var body = str_name + '　様' + '\n\n' +  
    '[ 申込内容 ]' + '\n' + datas  + '\n\n' + contents.content;    
  sendMail(toadd,subject,body) ;
  
}