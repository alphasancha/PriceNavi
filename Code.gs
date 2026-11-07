function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('端末価格ナビ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}


function getDatabaseData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. "DB" シートを取得（見つからなければ一番左のシートを使用）
    var dbSheet = ss.getSheetByName("DB") || ss.getSheets()[0];
    
    var lastRow = dbSheet.getLastRow();
    var lastCol = dbSheet.getLastColumn();
    
    if (lastRow < 2) {
      return { items: [], lastUpdate: "データがありません" };
    }
    
    // シートの全データを二次元配列で取得
    var dataValues = dbSheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = dataValues[0];
    var result = [];
    
    // 2行目以降（実データ）をオブジェクトに変換
    for (var i = 1; i < dataValues.length; i++) {
      var row = dataValues[i];
      var obj = {};
      var hasData = false;
      
      for (var j = 0; j < headers.length; j++) {
        var headerName = headers[j];
        if (headerName !== "" && headerName !== null && headerName !== undefined) {
          var key = String(headerName).trim();
          var val = row[j];
          
          // 日付や特殊オブジェクトなら文字列に変換
          if (val instanceof Date) {
            val = val.toLocaleDateString();
          }
          
          obj[key] = val;
          if (val !== "" && val !== null && val !== undefined) {
            hasData = true;
          }
        }
      }
      
      if (hasData) {
        result.push(obj);
      }
    }
    
    // 2. "設定" シートから最終更新日を取得
    var lastUpdateVal = "";
    var settingSheet = ss.getSheetByName("設定");
    if (settingSheet) {
      lastUpdateVal = settingSheet.getRange("B2").getValue();
      if (lastUpdateVal instanceof Date) {
        lastUpdateVal = lastUpdateVal.toLocaleDateString();
      }
    }
    
    return {
      items: result,
      lastUpdate: String(lastUpdateVal)
    };
    
  } catch (error) {
    return { items: [], lastUpdate: "エラー: " + error.message };
  }
}
