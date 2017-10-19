/*
* Project           : 趣味のプログラミング
*
* Program name      : ニコニコ動画ページ取得
*
* Purpose           : ニコニコ動画にアップロードされる実況プレイ動画について
*                   : いちいちサイトにサクセスして新着を調べるのは面倒なのでスクリプで自動取得する。
*
* Memo              : ・ニコニコ動画はJavascriptでページ生成をしているため、通常のUrlFetchAppではHTMLを取得できない
*                   : 　「PhantomJS Cloud API」を利用する。
*                   : ・アクセスするページはスプレッドシートのシート「キーワード」にて設定する。
*                   : ・キーワード、タグに対する除外条件はスプレッドシートのシート「除外」にて設定する。
*/


/**
* グローバル
*/
var START_ROW      = 2;
var START_COL      = 1;
var HEADER_ROW     = 1;
var LIMIT_VIEW_CNT = 20;
var TIME_ZONE      = "Asia/Tokyo";
var executionTime  = new Date();
var urlThisTime    = []; // 重複除去に使用する配列

var ss    = SpreadsheetApp.getActiveSpreadsheet();

/**
* メイン処理
*/
function main() {
  
  // アクセス対象が記載されているシート内容を取得し、連想配列に格納する
  const sheetHeaderAndValues = getSheetHeaderAndValues();

  // 連想配列を加工しメール送付内容を作成する
  const mailBody = generateMailBody(sheetHeaderAndValues);

  // メールを送付する
  if(!(mailBody)) return;
  mailSend(mailBody);
}


/**
* シート内容取得
*/
function getSheetHeaderAndValues(sheetHeaderAndValues) {
  const SHEET_NAME = 'キーワード';
  
  // キーワードシートの内容を連想配列化して、配列にする
  const sheet  = ss.getSheetByName(SHEET_NAME);
  const header = sheet.getSheetValues(HEADER_ROW, START_COL, 1, sheet.getLastColumn() - START_COL + 1)[0];
  const values = sheet.getSheetValues(START_ROW, START_COL, sheet.getLastRow() - START_ROW + 1, sheet.getLastColumn() - START_COL + 1);
  
  return {'Header':header,'Values':values};
}


/**
* メール本文作成
*/
function generateMailBody(sheetHeaderAndValues) {
  const SHEET_NAME = '除外';
  
  // 除外シートの内容を取得し、2次元配列からただの配列に変換する
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const sheet     = ss.getSheetByName(SHEET_NAME);
  const valeus    = sheet.getSheetValues(START_ROW, START_COL, sheet.getLastRow() - START_ROW + 1, 1);
  const skipArray = valeus.map(function(element){return element[0]});
  
  
  // PhantomJS Cloud APIを使用してニコニコ動画にアクセスし、ページ内容を取得する関数リテラル
  const fetchUrl = function(url){
    const userProperties = PropertiesService.getUserProperties();
    const API_KEY        = userProperties.getProperty('API_KEY');
    const urlPhantom     = "https://phantomjscloud.com/api/browser/v2/"+API_KEY+"/?request=%7Burl:%22"+ url +"%22,renderType:%22html%22,outputAsJson:true%7D";
    const response       = UrlFetchApp.fetch(urlPhantom);

    // 帰ってきたページ内容からテキストを抽出し、JSONで扱えるようにする
    const json       = JSON.parse(response.getContentText());
    // data情報以下を返却する
    return json.content.data;
  }
  
  
  // ページ内容から各動画の紹介部分HTMLを抜き出し、配列に格納する関数リテラル
  const extractItems = function(response){
    return response.match(/class="videoList01Wrap"([\s\S]*?)class="count comment"/gim);
  }
  

  // 動画のHTMLより動画のタイトルを抽出する関数リテラル
  const extractTitle = function(item){
    return item.replace(/^[\s\S]*?\<a title="([\s\S]*?)"[\s\S]*?$/i, "$1");
  }
  

  // 動画のHTMLより動画のURLを抽出する関数リテラル
  const extractUrl = function(item){
    return "http://www.nicovideo.jp/watch/" + item.replace(/^[\s\S]*?data-id="([\s\S]*?)"[\s\S]*?$/i, "$1");
  }
  

  // 動画のHTMLより動画のサムネイルを抽出する関数リテラル
  const extractImg = function(item){
    return item.replace(/^[\s\S]*?"(http:\/\/tn[\s\S]*?\.M)"[\s\S]*?$/gim, "$1");
  }
  

  // 動画のHTMLより視聴数を抽出する関数リテラル
  const extractView = function(item){
    return item.replace(/^[\s\S]*?class="count view"[\s\S]*?class="value"\>([\s\S]*?)\<[\s\S]*?$/gim, "$1").replace(/,/gim, "");
  }
  

  // 動画のHTMLより経過時間を抽出する関数リテラル
  const extractTime = function(item){
    return item.replace(/^[\s\S]*?"time new hour"\>(\d*?) hour ago[\s\S]*?$/gim, "$1");
  }
  
  
  // 動画のHTMLより欲しい情報を連想配列で取得する関数リテラル
  const extractInfo = function(item){
    return {
      'Title' : extractTitle(item)
      ,'Url'  : extractUrl(item)
      ,'Img'  : extractImg(item)
      ,'View' : extractView(item)
      ,'Time' : extractTime(item)
      ,'Tags' : ""
    };
  }
  
  
  // 改行を除去する関数リテラル
  const lfRemovedStr = function (str) {
    return str.replace(/\r?\n/gim,"");
  }
  
  
  // 動画のURLより詳細ページにアクセスし、動画のタグ情報を取得し動画情報の連想配列に追加する関数リテラル
  const addTags = function(dict){
    const url       = dict['Url'];
    const response  = lfRemovedStr(UrlFetchApp.fetch(url).getContentText());
    const tags      = response//
    .replace(/^[\s\S]*?\<!\-\- google_ad_section_start \-\-\>([\s\S]*?)\<!\-\- google_ad_section_end \-\-\>[\s\S]*?$/, "$1")//
    .replace(/\<.+?\>+/g, "")//
    .replace(/\&nbsp;\&nbsp;/g, " / ")//
    ;
    dict["Tags"]    = tags;
    return dict; 
  }
  
  
  // 視聴数が少ない動画を除去する関数リテラル
  const lowView = function(dict){
    return ( (dict['View']*1 / dict['Time']*1) >= LIMIT_VIEW_CNT);
  }
  
  
  // 除外キーワードがタイトルもしくはタグに含まれている場合は除去する関数リテラル
  const removeNgWord = function(dict){
    const title        = dict['Title'];
    const tags         = dict['Tags'];
    
    // 除外キーワード配列の内容がを一つでも含まれていればtrueを返却
    const shouldBeSkip = skipArray.some(function(element){
      const re = new RegExp(element, 'i');
      return re.test(title + tags)
    });

    return (!(shouldBeSkip));
  }
  
  
  // 重複を取り除く関数リテラル
  const removeDupli = function (dict) {
    const url = dict['Url'];
    if (urlThisTime.indexOf(url) == -1) {
      urlThisTime.push(url);
      return true;
    };
  };
  
  
  // 連想配列からメール本文を生成
  const convImgHtml = function(dict){
    
    const url       = dict['Url'];
    const title     = dict['Title'];
    const img       = dict['Img'];
    const tags      = dict['Tags'];
    
    var html = "";
    html += '<br><br><br><strong>' + title + '</strong><br>';
    html += '<a href="' + url + '"><img src="' + img + '"></a>';
    html += '<br><FONT color=#666666>' + tags + '</font>';
    return html;
  }
  
  
  // スプレッドシートの内容を連想配列にする連想配列
  const convArrToDict = function(element){
    var dict = {};
    
    // ヘッダーをキー、データを値
    element.forEach(function(element, index){
      // キー情報はヘッダーから取得
      const headerName = sheetHeaderAndValues['Header'][index];

      // URLの場合は、加工して最終的にはメール本文にする
      if (headerName == 'url') {
        Logger.log(element);
        element = extractItems(fetchUrl(element))//
        .map(lfRemovedStr)//
        .map(extractInfo)//
        .filter(lowView)//
        .filter(removeNgWord)//
        .map(addTags)//
        .filter(removeNgWord)//2回めを実行するのは、まずはキーワードで除去して、詳細ページへのアクセスを減らすため
        .filter(removeDupli)//
        .map(convImgHtml)//
        ;
      }
      dict[headerName] = element;
    });
    return dict
  };
  
  
  // キーワードの区切りを追加する関数リテラル
  const convHtmlBody = function(dict){
    const contents = dict['url'];
    if (!(contents.length)) return '';
    return '<br><br><hr color="GhostWhite"><h2>' + dict['keyword'] + '</h2>' + contents.join("");
  };
  
  
  // スプレッドシートの中身からメール分を生成する
  return sheetHeaderAndValues['Values']//
  .map(convArrToDict)//
  .map(convHtmlBody)//
  .join("")//
  ;
}


/**
* 送付
*/
function mailSend(mailBody){
  const sendAddress = Session.getActiveUser().getEmail().replace(/^([\s\S].+?)@([\s\S].+?)$/gim, "$1+star@$2");
  const ssLink      = '<a href="' + ss.getUrl() + '">台帳</a>';
  const cnt         = mailBody.match(/nicovideo/gim).length;
  const title       = "Nico " + Utilities.formatDate(executionTime, TIME_ZONE, "MM/dd") + " (" + cnt + ")";
  GmailApp.sendEmail(sendAddress, title, '', {htmlBody: ssLink + mailBody})
}