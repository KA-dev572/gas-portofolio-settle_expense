function expenseApplication(e) {
  //  設計:１　まずここで社員からの経費精算申請を受け付け、経費精算書ファイルに転記して自動的に仕訳：借方）申請された科目/貸方）未払金→経理担当にも通知？
  //       2  経理担当が申請内容（ローデータ）を確認し、必要であれば修正後に経費精算部分を集計（これは本来の設計外？：やるなら経費精算書にコンテナバインド）
  //          精算したらその仕訳も行う。
  
  //前処理: まずは出力先  実運用では実際の設定に合わせる
  let outputSheet = SpreadsheetApp.openById("YOUR_SPREADSHEET_ID_HERE").getSheetByName("経費精算書");
  let oSHeaderRow = 1;
  let oSLastRow = outputSheet.getLastRow();
  let oSLastColumn = outputSheet.getLastColumn();
  let oSHeader = outputSheet.getRange(oSHeaderRow, 1, oSHeaderRow, oSLastColumn).getValues()[0];
  let oSDateColumn = oSHeader.indexOf("日付");
  let oSDetorItem = oSHeader.indexOf("借方科目");
  let oSDetorAmount = oSHeader.indexOf("借方金額");
  let oSCreditorItem = oSHeader.indexOf("貸方科目");
  let oSCreditorAmount = oSHeader.indexOf("貸方金額");
  let oSAbstract = oSHeader.indexOf("摘要");
  let oSName = oSHeader.indexOf("社員名");

  //eventの内容から取得→実務では配列作って一気に書き込みが主流。（一番下参照）
  let content = e.namedValues;  //回答順に辞書的配列になるはず.回答は文字列っぽいはず
  try {
    outputSheet.getRange(oSLastRow+1, oSDateColumn+1).setValue(new Date(content["立替日"][0]));
  } catch (errorDate) {
    Logger.log("日付エラー" + errorDate.message);
    if (errorDate.message === TypeError) {  //←成立しないらしい。単にLogger.log(errorDate)でよいらしい。
      Logger.log("Date型変換エラー？")
    }
  }
  outputSheet.getRange(oSLastRow+1, oSDetorItem+1).setValue(content["勘定科目"][0]);
  try {
    outputSheet.getRange(oSLastRow+1, oSDetorAmount+1).setValue((content["立替金額（円）"][0]));
    outputSheet.getRange(oSLastRow+1, oSCreditorAmount+1).setValue((content["立替金額（円）"][0]));
  } catch (errorNum) {
    Logger.log("金額エラー" + errorNum.message);
    if (errorNum.message === TypeError) { //←成立しないらしい。単にLogger.log(errorDate)でよいらしい。
      Logger.log("数値型変換エラー？")
      
    }
  }
  outputSheet.getRange(oSLastRow+1, oSCreditorItem+1).setValue("未払金");
  outputSheet.getRange(oSLastRow+1, oSAbstract+1).setValue(content["摘要（目的、相手方、場所等）"][0]);
  outputSheet.getRange(oSLastRow+1, oSName+1).setValue(content["氏名"][0]);

  //メールを送るときは以下を有効化
  // GmailApp.sendEmail("経理担当アドレス", "【自動送信】経費精算申請がありました。", `経費精算申請書および経費精算書を確認してください。\n\n申請者: ${content["氏名"][0]}, 勘定科目: ${content["勘定科目"][0]}, 金額: ${content["立替金額（円）"][0]}円, 摘要: ${content["摘要（目的、相手方、場所等）"][0]}`)
}

//考慮すべき点？
//フォーム側で半角数字の入力など指定ができないので金額が正しく反映されない可能性がある。回答を突き返すことはできないか？→よく見たら整数のみなど指定可能だった。
//さすがに煩雑すぎるか？メール送信時に金額をカンマで区切る方法があればそのほうがよい？（やろうと思えば無理にでもできるが…）
//→実務ではこう
// let row = new Array(oSLastColumn).fill("");
// row[oSDateColumn] = new Date(content["立替日"][0]);
// row[oSDetorItem] = content["勘定科目"][0];
// row[oSDetorAmount] = Number(content["立替金額（円）"][0]);
// row[oSCreditorItem] = "未払金";
// row[oSCreditorAmount] = Number(content["立替金額（円）"][0]);
// row[oSAbstract] = content["摘要（目的、相手方、場所等）"][0];
// row[oSName] = content["氏名"][0];

// outputSheet.getRange(oSLastRow+1, 1, 1, oSLastColumn).setValues([row]);

