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

/**
 * 配列データの形を整形する
 * @param Array datatitlee 1次配列 タイトル
 * @param Array ssdatas 1次配列　データ
 * return 
 */
function reTitle(datatitles,ssdatas){
    Logger.log(ssdatas)
  var content_titles = datatitles.map(function(value,index){

    if(ssdatas[index] == ''){
      return '';
    }else{
      return value;
    }
  },ssdatas);
  //空の要素は削除する
  var result = content_titles.filter(function( item ) {    
    return !(item == '');    
  });
   //Logger.log('残すタイトル' + content_titles);
  return result;    
    
}
/**
 * 配列データの形を整形する
 * @param Array datatitles 1次配列 タイトル
 * @param Array ssdatas 1次配列　データ
 * @param Array maxnum 連想配列　キーワード
 *　return Array 連想配列
 *
*/
function reData(datatitles,ssdatas,maxnum){

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
  var re_content_datas = content_datas.filter(function( item ) {    
    return !(item == '');    
  });
  //タイトルとデータをセットにする
  var result = re_content_datas.map(function(value,index){
    return datatitles[index] + '：' + value;
  });
  return result;
}

/**
 * 元の配列データから要素数のユニークなものだけを配列として返す
 * @param Array array  1次配列 タイトル
 *　return Array array ユニークな要素のみの配列
 *
*/
function uniq(array) {
  return array.filter(function(elem, index, self) {
    return self.indexOf(elem) === index;
  });
}
/**
* javascriptで変数が文字列かどうか判定する
* https://hacknote.jp/archives/5674/
*
*/
function isString(obj) {
    return typeof (obj) == "string" || obj instanceof String;
};

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