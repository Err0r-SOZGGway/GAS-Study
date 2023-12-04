function doGet() {
  myFunction();
  let html = HtmlService.createTemplateFromFile('index');
  return html.evaluate();
}

function myFunction() {
   // 現在開いているシートを参照する
  const sheet = SpreadsheetApp.getActiveSheet();

   // 取得するデータ範囲
  const firstCol = 1; // A列
  // const lastCol  = 1; // A列のみ
  // const firstRow = 2; // 2行目から
  // const lastRow  = sheet.getLastRow(); // 最終行まで

  // 入力するデータ範囲
  const urlCol       = 2; // B列
  const titleCol     = 3; // C列

  let response = UrlFetchApp.fetch("https://news.yahoo.co.jp/");
  let text = response.getContentText("utf-8");
  // console.log(text);

  //トップニュースのブロックを抽出
  let topic_block = Parser.data(text).from('class="sc-gjdhqi dgdDOH"').to('</div>').build();
  // console.log(topic_block);

  //ulタグで囲まれている記述（トップニュース）を抽出
  let content_block = Parser.data(topic_block).from('<ul>').to('</ul>').build();
  // console.log(content_block);

  // content_blockの要素のうち、aタグに囲まれている記述を抽出
  topics = Parser.data(content_block).from('<a').to('</a>').iterate();

  // ニュースリスト用の配列変数を宣言
  let newsList = new Array();

  // aタグに囲まれた記述の回数分、順位／タイトル／URLを抽出する
    for(news of topics){
       //配列内のインデックス番号+1を取得（ニュース掲載順位として利用）
       let newsRank = topics.indexOf(news) + 1;

        //URL取得
      let newsUrl = news.replace(/.*href="/,"").replace(/".*/,"");
      //タイトル取得
      let newsTitle = news.replace(/.*class="sc-dtLLSn dpehyt">/,"").replace(/<.*>/,"");

       // ニュース順位、URL、タイトルの組を作成
      let newsInfo = [newsRank, newsUrl, newsTitle];

      newsList.push(newsInfo);

      const Low = newsRank + 1;

    // シートに書き込み
    sheet.getRange(Low, firstCol).setValue(newsRank);
    sheet.getRange(Low, urlCol).setValue(newsUrl);
    sheet.getRange(Low, titleCol).setValue(newsTitle);
    
   }

  console.log(newsList);
}