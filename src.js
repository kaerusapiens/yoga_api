  
function rowExists(sheet,location, date, time, studio, className) {
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
function removeRow(calendar,sheet,location, date, time, studio, className, canceledDate,startDate) {
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