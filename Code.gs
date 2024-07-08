function mail_2_calendar() {
  // 先ずはシート取得
  const sheetId = '1_bz_c5yzk6j9-ssStXsRW6xPom0a7dp297ymbD4ovOo'; // Your Google Sheet ID
  const sheetName = 'Sheet1'; // Name of the sheet where you want to append data
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);

  // Your calendar ID
  const calendarId = 'kmy63krr@gmail.com';
  const calendar = CalendarApp.getCalendarById(calendarId);

  // 既に予約済みのメールが複数くる場合があったので、現在のシートに予約情報があるかないか確認
  function rowExists(location, date, time, studio, className) {
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === location && data[i][1] === date && data[i][2] === time && data[i][3] === studio && data[i][4] === className) {
        return true;
      }
    }
    return false;
  }

  //地域、予約日、スタジオ、クラス名が同じで、キャンセルメールが予約メール受信日より後に来た場合メール削除
  //--キャンセルメールが予約メール受信日より後に来た場合メール削除:キャンセルしてから再予約した場合の削除防止
  function removeRow(location, date, time, studio, className, canceledDate,startDate) {
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === location && data[i][1] === date && data[i][2] === time && data[i][3] === studio && data[i][4] === className && data[i][6] <= canceledDate) {
        sheet.deleteRow(i + 1); // deleteRow uses 1-based index
        Logger.log("Removed row: " + location + ", " + date + ", " + time + ", " + studio + ", " + className);

        // Find and delete the corresponding calendar event
        console.log("removeing calander date")
        console.log(startDate)
        const events = calendar.getEventsForDay(startDate);
        events.forEach(event => {
          if (event.getTitle() === `${location}_${className}`) {
            event.deleteEvent();
            Logger.log("Removed calendar event: " + location + ", " + date + ", " + time + ", " + studio + ", " + className);
          }
        });

        return;
      }
    }
  }

  function dateformt(dates) {
    var dataParts = dates.split(" ");
    var datePart = dataParts[0];
    var timePart = dataParts[1].trim();
    var startTime = timePart.split('～');
    var startTime = startTime[0].trim();

    var dateParts = datePart.split("月");
    var month = dateParts[0];
    var day = dateParts[1].split("日(")[0];
    var weekday = dateParts[1].split("日(")[1];
    var month = month.padStart(2, '0');
    var day = day.padStart(2, '0');
    //2024
    var valueDate = new Date(`2024-${month}-${day}T${startTime}:00`);
    // Check if the date is valid
    if (isNaN(valueDate.getTime())) {
      throw new Error(`Invalid date: ${dates}`);
    }
    var formattedDate = `${month}月${day}日(${weekday}`;
    return { date: formattedDate, time: timePart ,values: valueDate};
  }

  function dateformating(receivedDate) {
    let yymmdd = receivedDate.getFullYear() + '-' + ('0' + (receivedDate.getMonth() + 1)).slice(-2) + '-' + ('0' + receivedDate.getDate()).slice(-2) + ' ' + ('0' + receivedDate.getHours()).slice(-2) + ':' + ('0' + receivedDate.getMinutes()).slice(-2) + ':' + ('0' + receivedDate.getSeconds()).slice(-2);
    return yymmdd.trim();
  }

  // ヨガ予約完了メールアカウントで直近の10メールを取得
  var threads = GmailApp.search('from:reserve@yoga-lava.com', 0, 5);
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
          if (!rowExists(location, date, time, studio, className)) {
            sheet.appendRow([location, date, time, studio, className, instructor, receivedDate]);
            //Logger.log('Appended to sheet: ' + location + ", " + date + ", " + time + ", " + studio + ", " + className);

            // Add event to calendar
            console.log(datevalue)
            calendar.createEvent(`${location}_${className}`, datevalue, endDateValue, {description: `Instructor: ${instructor}`
            });
            Logger.log('Added calendar event: ' + location + ", " + date + ", " + time + ", " + studio + ", " + className);
          } 
        } else {
          Logger.log("No match found in reservation email");
        }
      }
    });
  });

  var threads = GmailApp.search('from:reserve@yoga-lava.com', 0, 5);
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
          if (rowExists(location, date, time, studio, className)) {
            removeRow(location, date, time, studio, className, canceledDate,datevalue);
          }
        } else {
          Logger.log("No match found in cancellation email");
        }
      }
    });
  });

  // 日付順番でメールが並ぶようにする
  const range = sheet.getRange('A1:G');
  range.sort({ column: 2, ascending: true });
}