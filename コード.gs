function setValueTest() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let url = ss.getUrl();

  // スプレッドシートの選択
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  // シートの選択
  let sheet = spreadsheet.getSheetByName(`シート1`);
  // 書き込み
  sheet.getRange('D2').setValue('未完了');
  sheet.getRange('D3').setValue('未完了');
}

function sendResult() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let url = ss.getUrl();

  // スプレッドシートの選択
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  // シートの選択
  let sheet = spreadsheet.getSheetByName(`シート1`);
  // 名前の取得
  let persons = sheet.getRange('A2:C3').getValues();
  let name = persons[0][0];
  // メールアドレスの取得
  let mailAddress = persons[0][1];
  // 獲得賞の取得
  let award = persons[0][2];

  // メールの本文作成
  let message = `${name} 様
  N予備校コンテスト運営局です。
  この度は「Webページコンテスト」にご応募いただき、誠にありがとうございます。
  
  コンテストの結果、 ${name} 様は見事 ${award} に選ばれました`;

  // オプションの設定
  // let options = {
  //   attachments: [DriveApp.getFileById('1QXzxzdZHR9z5z5S014NLSvPFXfbcLq9y').getBlob()]
  // };
  // メール送信
  MailApp.sendEmail(mailAddress, 'コンテストの結果について', message);
}

function sendResults() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let url = ss.getUrl();

  // スプレッドシートの選択
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  // シートの選択
  let sheet = spreadsheet.getSheetByName(`シート1`);
  let numOfPersons = sheet.getLastRow() - 1;
  let persons = sheet.getRange(2, 1, numOfPersons, 4).getValues();
  
  persons.forEach(function(person,i) {
    let name = person[0];
    let mailAddress = person[1];
    let award = person[2];
    let sendStatus = person[3];

    if (sendStatus === '未完了') {

      // メールの本文作成
      let message = `${name} 様
  N予備校コンテスト運営局です。
  この度は「Webページコンテスト」にご応募いただき、誠にありがとうございます。
  
  コンテストの結果、 ${name} 様は見事 ${award} に選ばれました`;

      // オプションの設定
      // let options = {
      //   attachments: [DriveApp.getFileById('1QXzxzdZHR9z5z5S014NLSvPFXfbcLq9y').getBlob()]
      // };
      // メール送信
      MailApp.sendEmail(mailAddress, 'コンテストの結果について', message);

      // メール送信状況の更新
      let range2 = sheet.getRange(`D${i+2}`);
      range2.setValue('完了');
    }
  });
}

function changeHeader() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let url = ss.getUrl();

  // スプレッドシートの選択
  let spreadsheet = SpreadsheetApp.openByUrl(url);
  // シートの選択
  let sheet = spreadsheet.getSheetByName(`シート1`);
  // 書き込み
  let range = sheet.getRange('A1:D1');
  range.setBackground('black');
  range.setFontColor('white');
}

function omikuji() {
  let unsei = ['大吉', '中吉', '小吉', '吉', '凶'];
  let rand = Math.floor(Math.random() * unsei.length);
  console.log(unsei[rand]);
}

function loopTest() {
  let persons = [
    [ '太郎', 'example1@gmail.com', '最優秀賞', '未送信' ],
    [ '二郎', 'example2@gmail.com', '優秀賞', '未送信' ],
    [ '三郎', 'example3@gmail.com', '健闘賞', '未送信' ]
  ];
  
  persons.forEach(function(person, i) {
    console.log(i, person);
  });

  for (let person of persons) {
    console.log(person);
  }
}

function ifTest() {
  let sendStatus = '未完了';
  if (sendStatus === '未完了') {
    console.log('メール送信します。');
  }
}
