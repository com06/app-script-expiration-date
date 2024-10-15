function initSpreadSheet() {
  ss = SpreadsheetApp.getActive();// เลือกสเปรดชีตที่กำลังเปิดใช้งานอยู่ในปัจจุบัน
  sheet = ss.getSheetByName(sheetName);// เลือกชีตที่ตรงกับชื่อที่กำหนด
  data = sheet.getDataRange().getValues();// ได้ข้อมูลทั้งหมดในชีตนั้นมาเก็บไว้ในตัวแปร data
  console.log('initSpreadSheet completed');
}

function checkExpiry() {
  // เรียกใช้ฟังก์ชัน initSpreadSheet() เพื่อเตรียมข้อมูลจากสเปรดชีต
  initSpreadSheet()

  // สร้าง Object Date เพื่อเก็บวันที่ปัจจุบัน
  var today = new Date();

  // สร้าง Array แล้วทำการวนลูปเพิ่มจำนวนวันที่ 90 30 15 ตาม Array ที่กำหนด และจัดรูปแบบวันที่ให้เป็น "yyyy-MM-dd"
  var formattedDates = [90, 30, 15].map(function(days) {
    var date = new Date();
    date.setDate(today.getDate() + days);
    return Utilities.formatDate(date, "Asia/Bangkok", "yyyy-MM-dd");
  });

  // แยกค่าจากอาร์เรย์ formattedDates ออกมาเก็บในตัวแปรต่างๆ
  var [formattedNinetyDays, formattedThirtyDays, formattedFifteenDays] = formattedDates;

  var emailBody = "";
  var lineMessage = "";

  // วนลูปอ่านข้อมูล วันหมดอายุ ทั้งหมดในชีต
  for (var i = 1; i < data.length; i++) {
    var expiryDate = new Date(data[i][1]); // คอลัมน์ที่ 2 คือ วันหมดอายุ
    expiryDate = Utilities.formatDate(expiryDate, "Asia/Bangkok", "yyyy-MM-dd");

    // เปรียบเทียบวันที่
    if (expiryDate == formattedFifteenDays) {
      emailBody += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 15 วัน (" + expiryDate + ")\n";
      lineMessage += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 15 วัน (" + expiryDate + ")\n";
    }else if (expiryDate == formattedThirtyDays){
      emailBody += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 30 วัน (" + expiryDate + ")\n";
      lineMessage += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 30 วัน (" + expiryDate + ")\n";
    }else if (expiryDate == formattedNinetyDays){
      emailBody += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 90 วัน (" + expiryDate + ")\n";
      lineMessage += "สินค้า " + data[i][0] + " จะหมดอายุในอีก 90 วัน (" + expiryDate + ")\n";
    }
  }

  // ส่ง Email
  if (emailBody !== "") {
    MailApp.sendEmail(email, 'แจ้งเตือนสินค้าใกล้หมดอายุ', emailBody);
    console.log("Send email : Success")
  }

  // ส่ง Line Notify
  if (lineMessage !== "") {
    sendLineMessageToUser(emailBody)
    console.log("Send LINE : Success")
  }
}

function sendLineMessageToUser(txt_message) {
  var url = 'https://api.line.me/v2/bot/message/push';
  var headers = {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer '+CHANNEL_ACCESS_TOKEN
  };
  var payload = {
    to: user_id, // เปลี่ยนเป็น User ID ของผู้รับ
    messages: [{
      type: 'text',
      text: txt_message
    }]
  };

  var options = {
    'method': 'post',
    'headers': headers,
    'payload': JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}

function doPost(e) {

}
