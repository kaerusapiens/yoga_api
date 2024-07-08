function main(){
    const sheetName = 'Sheet1'; // Name of the sheet where you want to append data
    let sheetA = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    const calendarA = CalendarApp.getCalendarById(calendarId);
    
    // ヨガ予約完了メールアカウントで直近の10メールを取得
    let sender = "from:" + yoga_sender;
    var threads = GmailApp.search(sender, 0, 5);
    threads.forEach(thread => {
        const messages = thread.getMessages();
        messages.forEach(message => {
            let subject = message.getSubject();
            var judge1 = subject.includes("レッスンのご予約ありがとうございます。");
            if (judge1) {
                let body = message.getPlainBody();
                let receivedDate = dateformating(message.getDate());
                var startString = "ご予約いただいたレッスンの日時は以下となりますので、ご確認をお願いいたします。";
                var endString = "追加料金（※）：なし";
                var regex = new RegExp(startString + '[\\s\\S]*?' + endString);
                var match = body.match(regex);
                if (match) {
                    var result = match[0].replace(startString, '').replace(endString, '').trim();
                    var dataParts = result.split('\n').map(part => part.trim());
                    var location = dataParts[0];
                    var dateTime = dateformt(dataParts[1]);
                    var date = dateTime.date;
                    var time = dateTime.time;
                    var studio = dataParts[2];
                    var className = dataParts[3];
                    var datevalue = dateTime.values;
                    var endDateValue = new Date(datevalue.getTime() + 60 * 60 * 1000); //
                    var instructor = dataParts.length > 4 ? dataParts[4].replace("担当インストラクター：", "") : "";
                    if (!rowExists(sheetA, location, date, time, studio, className)) {
                      sheetA.appendRow([location, date, time, studio, className, instructor, receivedDate]);
                    //Logger.log('Appended to sheet: ' + location + ", " + date + ", " + time + ", " + studio + ", " + className);

          // Add event to calendar
          console.log(datevalue)
          calendarA.createEvent(`${location}_${className}`, datevalue, endDateValue, {description: `Instructor: ${instructor}`
          });
          Logger.log('Added calendar event: ' + location + ", " + date + ", " + time + ", " + studio + ", " + className);
        } 
      } else {
        Logger.log("No match found in reservation email");
      }
    }
  });
})

    var threads = GmailApp.search(sender, 0, 5);
    threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
        let subject = message.getSubject();
        var judge2 = subject.includes("キャンセル完了");
        if (judge2) {
        let body = message.getPlainBody();
        let canceledDate = message.getDate();
        var startString = "以下のご予約についてキャンセルを承りましたので、ご確認ください。";
        var endString = "万が一、誤りのある場合は、下記よりお手続きをお願い申し上げます。";
        var regex = new RegExp(startString + '[\\s\\S]*?' + endString);
        var match = body.match(regex);
        if (match) {
            var result = match[0].replace(startString, '').replace(endString, '').trim();
            var dataParts = result.split('\n').map(part => part.trim());
            var location = dataParts[0];
            var dateTime = dateformt(dataParts[1]);
            var date = dateTime.date;
            var time = dateTime.time;
            var studio = dataParts[2];
            var className = dataParts[3];
            var datevalue = dateTime.values;
            if (rowExists(sheetA,location, date, time, studio, className)) {
            removeRow(calendarA,sheetA,location, date, time, studio, className, canceledDate,datevalue);
            }
        } else {
            Logger.log("No match found in cancellation email");
      }
    }
  });
});

// 日付順番でメールが並ぶようにする
const range = sheetA.getRange('A1:G');
range.sort({ column: 2, ascending: true });
}
