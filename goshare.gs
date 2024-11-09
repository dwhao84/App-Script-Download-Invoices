function saveGoShareInvoicesToPDF() {
  // 指定的資料夾名稱
  var folderName = "GoShare-Inv-2";
  var folder = DriveApp.getFoldersByName(folderName);
  
  if (!folder.hasNext()) {
    folder = DriveApp.createFolder(folderName);
  } else {
    folder = folder.next();
  }
  
  // 搜尋 Goshare 標籤且主旨含「電子發票開立通知」的郵件
  var threads = GmailApp.search("label:Goshare subject:電子發票開立通知", 0, 50);
  Logger.log("找到 " + threads.length + " 封符合條件的郵件");
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var date = message.getDate();
      
      // 將日期格式化為 YYMMDD
      var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyMMdd");
      
      // 新的檔名格式：YYMMDD - GoShare 電子發票開立通知.PDF
      var fileName = formattedDate + " - GoShare 電子發票開立通知.PDF";
      
      // 檢查是否已經存在同名檔案
      var existingFiles = folder.getFilesByName(fileName);
      if (existingFiles.hasNext()) {
        // 如果同一天有多張發票，加上序號
        var counter = 1;
        while (existingFiles.hasNext()) {
          fileName = formattedDate + " - GoShare 電子發票開立通知(" + counter + ").PDF";
          existingFiles = folder.getFilesByName(fileName);
          counter++;
        }
      }
      
      try {
        // 創建 HTML 內容
        var htmlContent = "<div style='font-family: Arial, sans-serif;'>" +
                         "<div>" + message.getBody() + "</div>" +
                         "</div>";
        
        // 轉換成 PDF 並儲存
        var blob = Utilities.newBlob(htmlContent, "text/html", fileName);
        var pdf = folder.createFile(blob.getAs("application/pdf"));
        
        Logger.log("成功儲存檔案: " + fileName);
        
        // 標記郵件為已讀（可選）
        message.markRead();
      } catch (error) {
        Logger.log("處理郵件時發生錯誤: " + error.toString());
      }
    }
  }
}

// 新增一個手動執行的函式，可以指定日期範圍
function saveGoShareInvoicesWithDateRange() {
  var startDate = "2024/03/05"; // 請修改成你要的開始日期
  var endDate = "2024/12/31";   // 請修改成你要的結束日期
  
  var searchQuery = "label:Goshare subject:電子發票開立通知 after:" + startDate + " before:" + endDate;
  var threads = GmailApp.search(searchQuery, 0, 500);
  
  Logger.log("開始處理 " + startDate + " 到 " + endDate + " 的發票");
  saveGoShareInvoicesToPDF();
}

// 建立自動執行觸發器
function createTrigger() {
  // 刪除現有的觸發器
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // 建立新的每日觸發器
  ScriptApp.newTrigger('saveGoShareInvoicesToPDF')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}
