/**
* 作成：2019.12.25　ライブラリ Ver002
*
* 時刻起動トリガーが実行されると
* 残数をチェックして内容に合わせたメールを送信する。
* 紐付けシートの名称はデフォルトのままとする("フォームの回答 1")
* メール送信記録シートの名前は"メール送信記録"とする。
* メール本文は"本文"シートとする、ユーザー手動作成
* 列0:タイムスタンプ 列1:アドレス 列2:名前　前提条件
* 残2以上でもフォームを開いている人が何にもいて送信するとインスタンス扱いのようなので全てその値で送信記録される。
* そのため全数という最大数設定が判断には必要
* 全数は途中で変更可能
*/
function onFormSubmit(e) {

  GASTooL01formApplicationCounter.fncounterListItem(e)
  
}

function testRowCol(){
  var ss =SpreadsheetApp.getActiveSpreadsheet();
  var shsent = ss.getSheetByName('テスト');
  var sentdatas = shsent.getDataRange().getValues();
  var values = [
  ['Time','メアド','宛名','可否']
  ];
  shsent.getRange(1,1,1,4).setValues(values);//レンジは1から
  
}
function test(){
var maxnum={start:'pp'};
var datas=[
['かば','Apple','Orange'],
['かば','Orange','Apple']
];

  var indexlists=[];//ターゲット項目(時間)列群 
  for(var i=0;i<datas.length;i++){//
    for(var j=0;j<datas[i].length;j++){
      var inumstart = datas[i][j].indexOf(maxnum.start);//1文字は目は0番目 
      Logger.log(inumstart + ' ' + datas[i][j]);
       if(-1<inumstart){
        indexlists.push(j);//配列のindex番号リストをセット
      }
    }
  }
  Logger.log(indexlists);
}
function testMap(){
  var datatitle = ['タイムスタンプ','氏名	内線番号(数字のみ)','社員番号','希望日','12月10日','12月11日','12月12日','連絡事項','メールアドレス'];
  var maxnum = {start:'［全 ',last:'］'};
  var ssdatas=[];
  ssdatas.push(['高野1',111,'y1111','12月10日','10時［全 2］（残 2）','','','10時全2残2','sakiko.kono@y-yokohama.com']);
  datatitle = reTitle(datatitle,ssdatas);
  var content_datas  =reData(ssdatas[0],maxnum );
  Logger.log(content_datas);
}
/**
 * 配列データの形を整形する
 * @param Array datatitlee 1次配列 タイトル
 * @param Array ssdatas 1次配列　データ
 * return 
 */
function reTitle(datatitle,ssdatas){
//!2020/01/06 ここから　
  var content_titles = datatitle.map(function(value,index){
    if(ssdatas[index] == ''){
      return '';//こっちにこない
    }else{
      return value;
    }
  });
  //空の要素は削除する
  var result = content_titles.filter(function( item ) {    
    return !(item == '');    
  });
   Logger.log('残すタトル' + content_titles);
  return result;    
    
}
/**
 * 配列データの形を整形する
 * @param Array datatitlee 1次配列 タイトル
 * @param Array ssdatas 1次配列　データ
 * @param Array maxnum 連想配列　キーワード
 *　return Array 連想配列
 *
*/
function reData(ssdatas,maxnum){

  var content_datas = ssdatas.map(function(value,index){

    if(isString(value)){//文字列かどうか
      var item = value.substr(0,value.indexOf(maxnum.start));//キーが文字列に存在すれば取り除く
      if(item == ''){
        return value;//キーがない場合は受けた値のまま返す
      }else{
        return item;//整形後の文字列
      }
    }else{
      return value;
    }

  });
  Logger.log(content_datas);

  //空の要素は削除する
  var result = content_datas.filter(function( item ) {    
    return !(item == '');    
  });
  return result;
}

/**
* javascriptで変数が文字列かどうか判定する
* https://hacknote.jp/archives/5674/
*
*/
function isString(obj) {
    return typeof (obj) == "string" || obj instanceof String;
};
function testMax(){
  var data='10時［全 2］（残 2）';
  var maxnum = {start:'［全 ',last:'］'};
  var max = fnMaxNumber(data,maxnum);
  Logger.log(max);
}
/**
 * 文字列から引数キーワードに囲まれた桁数を抽出する
 * @param String data
 * @param Array maxnum 連想配列
 *　return Number max
 * 
 */
function fnMaxNumber(data,maxnum){

  var inumstart =data.indexOf(maxnum.start);//1文字は目は0番目 全開始キーワードの最初の文字が見つかった位置
  var inumlast =data.indexOf(maxnum.last);//1文字は目は0番目　全終了キーワードの最初の文字が見つかった位置
  var temp_start = maxnum.start;
  var maxnum_start_len = temp_start.length
  var temp_last = maxnum.last;
  var maxnum_last_len = temp_last.length
  var pos_start= inumstart + maxnum_start_len;//桁数抽出 開始位置
  var maxlen = inumlast-pos_start;//桁数
  var max=Number(data.substr(pos_start,maxlen));//該当の全数
  
  return max;
  
}


function fnConfirmMail2(){

  var ss =SpreadsheetApp.getActiveSpreadsheet();
  
  var MAILADD_SHEET = 'メール送信記録';
  var subject = '';
  var sheets = ss.getSheets();
  var sheetexists = sheets.map(function(sh){
    return sh.getName();
  });
  if(sheetexists.indexOf(MAILADD_SHEET)<0){//指定名のシートがない場合
    //メール送信記録シート追加
    InsertSheet(MAILADD_SHEET,2);    
  }
  var shsent = ss.getSheetByName(MAILADD_SHEET);
  var sentdatas = shsent.getDataRange().getValues();
  sentdatas =  sentdatas.filter(function(e){return e[0] !== "";});//空の要素を削除する  
  var mailrecordtitle = ['送信時間','メールアドレス','宛先名','可否'];
  var judgecol = 3;//'可否列
  mailrecordtitle.map(function(title,index){
    shsent.getRange(1,index+1).setValue(title);
  });
  if(sentdatas.length<2){//シート作成時のみ実行　初回答時に送信させるため。
    var m = Moment.moment(); //作成日時    
    var mailrecordinitialdata = [m.format('YYYY/MM/DD HH:mm:ss'),'サンプル','サンプル','OK'];
    mailrecordinitialdata.map(function(title,index){
      shsent.getRange(2,index+1).setValue(title);
    });
  }
  
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
  
  var FORMRET_SHEET = 'フォームの回答 1';
  var ssdatas = ss.getSheetByName(FORMRET_SHEET).getDataRange().getValues(); 
  var ssform=ss.getSheetByName(FORMRET_SHEET);
  var addcol=9;//メールアドレスはB列、ただし配列データで指定するときは-1となる
  var namecol=1;//宛名はC列  
  var datatitle = ssdatas.splice(0, 1)[0];
  ssdatas = ssdatas.filter(function(e){return e[0] !== "";});//空の要素を削除する 
  
  //回答データがないときは処理を終了
  if(ssdatas.length<1){
    return;
  }
  
  //回答データがある時は以下を実行
  var maxnum = {start:'［全 ',last:'］'}; //全角括弧開始、全の後半角空白 + 数値n+全角括弧閉じ
  var cntkey = {start:'（残 ',last:'）'};//全角括弧開始、残の後半角空白+数値n+全角括弧閉じ
  var cntoutnum = 4;//上記の文字列数 
  var confirmrows=[];// 
  
  //申込数が全num以上あるかどうかをチェック
  var indexlists=[];//ターゲット項目(時間)列群 
  for(var i=0;i<ssdatas.length;i++){//
    for(var j=1;j<ssdatas[i].length;j++){
      var value = ssdatas[i][j] + '';//数値があったら文字化
      var inumstart = value.indexOf(maxnum.start);//1文字は目は0番目 
      if(-1<inumstart){
        indexlists.push(j);//配列のindex番号リストをセット
      }
    }
  }
  Logger.log(indexlists);
  //項目(時間)列抽出 --A
  for(var i=0;i<indexlists.length;i++){//列数
    var col = indexlists[i];//項目列
    var targetcollist=[];//列群
    var targetobj={};
    var targetobjarr=[];
    var targetcol={};//項目(セル)内容
    //必要データを連想配列化
    for(var j=0;j<ssdatas.length;j++){//1列
      var data = ssdatas[j][col];//セルデータ
      targetcol.item=data.substr(0,data.indexOf(maxnum.start));
      targetcol.max=fnMaxNumber(data,maxnum);
      targetcol.rest=fnMaxNumber(data,cntkey);     
      targetcol.row=j;
      
      targetcollist.push(targetcol); 
    } 
    //グルーピング　項目ごと(10時　など)---B

    var itemlists=[];
    for(var k=0;k<targetcollist.length;k++){
        
      var targetitems = {};
      targetcollist.map(function (i,j) {
        var cat=i["item"];        
        if(typeof targetitems[cat]=="undefined"){
          targetitems[cat]=[];
          itemlists.push(cat);
        }
        targetitems[cat].push(i);
      });

      for(var n=0;n<itemlists.length;n++){
        var cat=itemlists[n];
        //降順に並び替え　項目中最新の全数を採用する
        targetitems[cat].sort(function(a,b){
          if (a.row < b.row) {
            return 1;
          } else {
            return -1;
          }
        })
        Logger.log('最大'+targetitems[cat][0].max);
        Logger.log('申込数' + targetitems[cat].length);
        var diffnum = targetitems[cat][0].max - targetitems[cat].length
        Logger.log(diffnum);

        if(diffnum < 0){//全数より申込数が多い
          for(var m=0;m<diffnum;m++){
            confirmrows.push(targetitems[cat][m].row); 
          }
        }      
      }

    }//グルーピング　項目ごと(10時　など) --B

  }//項目(時間)列抽出 --A
  
  //前回実行時間後のタイムスタンプ記録行のみ実行
  var maildatas=[];
  var lastrow=sentdatas.length-1;
  var lasttime = sentdatas[lastrow][0];
  //var colmax = ssdatas[0].length+1;
  var colmax = datatitle.length+1;
  var irowr = sentdatas.length //「メール送信記録」シートへの書き込み行
  for(var i=0;i<ssdatas.length;i++){//データの数だけ
  
    var irow = i+2;//レンジはタイトル行があり、1スタート
    if(lasttime < ssdatas[i][0]){//未送信のデータ
      var toadd =  ssdatas[i][addcol];
      //該当の宛先にメール送信　＋ 書き込み
      if(0<confirmrows.length){
        for(var j=0;j<confirmrows;j++){
          var strjudge = 'NG';
          var irowf = confirmrows[j];
          ssform.getRange(irowf,colmax).setValue(strjudge);//form回答シート          
          var resline = ng_contents;
        }
      }else{
        var strjudge = 'OK';
        ssform.getRange(irow,colmax).setValue(strjudge);  
        var resline = ok_contents;
      }
      Logger.log('行データの中身' + ssdatas[i]);

      var content_datas = reData(ssdatas[i],maxnum);
      var content = createTextMail(toadd,content_datas,resline); 
      irowr++;
      var m = Moment.moment(); //作成日時
      var senddate = m.format('YYYY/MM/DD HH:mm:ss');
      //['送信時間','メールアドレス','宛先名','可否']
      var values = [[senddate,toadd,ssdatas[i][namecol],strjudge]];
      shsent.getRange(irowr,1,values.length,values[0].length).setValues(values);//送信記録シート
      
    }
  }//データの数だけ

}
/**
 * シート追加
 * @param String sheetname シート名指定
 * @param Number num 何番目に追加か
 */
function InsertSheet(sheetname,num) {
  var objSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  objSpreadsheet.insertSheet(sheetname,num);
}
/**
 * 
 * TEXT mail 作成
 * @param array ssdata フォーム回答シートの1レコード
 * @param array content 本文シートのOK／NG別の1レコード
 * ssdata=>列0:タイムスタンプ 列1:アドレス 列2:名前　前提条件
 * https://www.sejuku.net/blog/21812
 */
function createTextMail(toadd,ssdata,contents){

  var subject = contents.subject;
  var datas = ssdata.join('\n');//配列を改行で結合して文字列で返す
  var body = ssdata[2] + '様' + '\n' +  
    '[ 申込内容 ]' + '\n' + datas  + '\n' + contents.content;    
  sendMail(toadd,subject,body) ;
  
}
/**
 * 参考
 * https://www.terakoya.work/google-apps-script-json-mail/
 * https://tonari-it.com/gas-coding-guide-line/
 */

/**
 * TEXT mail send.
 * 
 * @param String content
 * @param String subject
 * @param String content
 */
function sendMail(toadd,subject,content) {
  MailApp.sendEmail({
    to: toadd,
    subject: subject,
    body: content
  });
}

