/**
 * GOOGLE APPS SCRIPT: PHIÊN BẢN CHO GMAIL CHÍNH CHỦ (OWNER)
 * Civilis - Chiến dịch "Xóa mù tiếng Hàn"
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // Dùng trực tiếp file đang mở
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName("data");
    var codeSheet = ss.getSheetByName("code");
    
    var name = e.parameter.name || "Học viên";
    var email = e.parameter.email || "";
    var zalo = e.parameter.zalo || "";
    var plan = e.parameter.plan || "";
    var url = e.parameter.url || "";
    var date = new Date();

    // 1. Xác định dòng mới để ghi dữ liệu (Dùng cách này để đảm bảo cột G luôn khớp)
    var newRow = dataSheet.getLastRow() + 1;
    
    // 2. Ghi dữ liệu vào 6 cột đầu: Thời gian | Họ và Tên | Số điện thoại | Email | Sản phẩm quan tâm | URL
    dataSheet.getRange(newRow, 1, 1, 6).setValues([[date, name, zalo, email, plan, url]]);

    var studyCode = "Liên hệ Zalo để nhận mã";
    var codesData = codeSheet.getDataRange().getValues();
    var codeRowIndex = -1;

    for (var i = 1; i < codesData.length; i++) {
        if (!codesData[i][1]) {
            studyCode = codesData[i][0];
            codeRowIndex = i + 1;
            break;
        }
    }

    if (codeRowIndex !== -1 && email !== "") {
        codeSheet.getRange(codeRowIndex, 2).setValue("SENT");
        codeSheet.getRange(codeRowIndex, 3).setValue(email);
        
        var timeStamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
        codeSheet.getRange(codeRowIndex, 4).setValue(timeStamp);

        sendStudyCodeEmail(name, email, studyCode);
        
        // 3. Cập nhật trạng thái "Đã gửi" vào đúng cột G của dòng vừa thêm
        dataSheet.getRange(newRow, 7).setValue("Đã gửi (Auto)");
    } else {
        // Nếu không có mã hoặc không có email, có thể ghi chú lại
        dataSheet.getRange(newRow, 7).setValue("Chưa gửi (Không có mã/email)");
    }

    return ContentService
          .createTextOutput(JSON.stringify({ 'result': 'success', 'code': studyCode }))
          .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
          .createTextOutput(JSON.stringify({ 'result': 'error', 'error': error.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 Civilis Tools')
      .addItem('Gửi mã cho các dòng đang chọn', 'sendEmailManually')
      .addToUi();
}

function sendEmailManually() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var codeSheet = ss.getSheetByName("code");
  
  if (activeSheet.getName() !== "data") {
    SpreadsheetApp.getUi().alert("LỖI: Hãy mở sheet 'data' và chọn các dòng cần gửi mail.");
    return;
  }

  var activeRange = activeSheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  var values = activeSheet.getRange(startRow, 1, numRows, activeSheet.getLastColumn()).getValues();
  
  var headers = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
  var emailColIdx = -1;
  var nameColIdx = -1;
  
  for (var h = 0; h < headers.length; h++) {
    var head = headers[h].toString().toLowerCase();
    if (head.indexOf("email") !== -1) emailColIdx = h;
    if (head.indexOf("tên") !== -1 || head.indexOf("name") !== -1) nameColIdx = h;
  }
  
  if (emailColIdx === -1) emailColIdx = 3; 
  if (nameColIdx === -1) nameColIdx = 1;

  var sentCount = 0;
  for (var i = 0; i < values.length; i++) {
    if (startRow + i === 1) continue;
    var name = values[i][nameColIdx] || "Học viên";
    var email = values[i][emailColIdx];
    
    if (email && email.toString().indexOf("@") !== -1) {
      var studyCode = "";
      var codesData = codeSheet.getDataRange().getValues();
      var codeRowIndex = -1;

      for (var j = 1; j < codesData.length; j++) {
          if (!codesData[j][1]) {
              studyCode = codesData[j][0];
              codeRowIndex = j + 1;
              break;
          }
      }

      if (codeRowIndex !== -1) {
        var timeStamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
        codeSheet.getRange(codeRowIndex, 2).setValue("SENT (MANUAL)");
        codeSheet.getRange(codeRowIndex, 3).setValue(email);
        codeSheet.getRange(codeRowIndex, 4).setValue(timeStamp);
        sendStudyCodeEmail(name, email, studyCode);
        
        // Cập nhật trạng thái "Đã gửi" vào cột G của dòng tương ứng trong sheet "data"
        activeSheet.getRange(startRow + i, 7).setValue("Đã gửi (Manual)");
        
        sentCount++;
      }
    }
  }
  
  if (sentCount > 0) {
    SpreadsheetApp.getUi().alert("✅ Đã gửi thành công " + sentCount + " email!");
  } else {
    SpreadsheetApp.getUi().alert("❌ Không gửi được mail nào. Hãy bôi đen dòng chứa Email nhé.");
  }
}

function sendStudyCodeEmail(name, email, code) {
  var subject = "TẶNG BẠN: Mã kích hoạt bộ 20 video học tiếng Hàn từ Civilis";
  
  var htmlBody = `
    <div style="font-family: 'Be Vietnam Pro', Arial, sans-serif; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-top: 10px solid #003366; border-radius: 15px; overflow: hidden; background-color: #ffffff;">
      <div style="padding: 30px; text-align: center; background-color: #f8fbff;">
        <h1 style="color: #003366; margin: 0; font-size: 24px;">Chúc mừng ${name}!</h1>
        <p style="color: #666; font-size: 16px; margin-top: 10px;">Bạn đã chính thức bắt đầu hành trình "Xóa mù tiếng Hàn"</p>
      </div>
      
      <div style="padding: 40px 30px; text-align: center;">
        <p style="color: #333; font-size: 16px; line-height: 1.6;">Chào mừng bạn đến với Civilis! Đây là mã kích hoạt học tập dành riêng cho bạn:</p>
        
        <div style="margin: 30px 0 10px 0; padding: 20px; background: #fff4e5; border: 2px dashed #ffb800; border-radius: 10px; display: inline-block;">
          <span style="font-size: 28px; font-weight: 900; color: #cc0000; letter-spacing: 2px;">${code}</span>
        </div>
        
        <p style="color: #cc0000; font-size: 14px; font-style: italic; margin-top: 0; margin-bottom: 30px;">
            * Bạn vui lòng truy cập và nhập mã kích hoạt trong vòng 24h để tránh bị thu hồi mã.
        </p>
        
        <div style="text-align: left; background: #f0f7ff; padding: 25px; border-radius: 12px; margin-top: 20px;">
          <h3 style="color: #003366; margin-top: 0;">4 Bước để bắt đầu ngay:</h3>
          <ul style="color: #444; list-style-type: none; padding-left: 0;">
            <li style="margin-bottom: 12px;"><strong>Bước 1:</strong> Truy cập website <a href="https://topik4u.vn" style="color: #cc0000; text-decoration: none; font-weight: bold;">Topik4u.vn</a></li>
            <li style="margin-bottom: 12px;"><strong>Bước 2:</strong> Đăng ký tài khoản (hoặc đăng nhập bằng tài khoản Google) và chọn phần <strong>Nhập mã kích hoạt</strong>.</li>
            <li style="margin-bottom: 12px;"><strong>Bước 3:</strong> Dán mã bên trên và chọn <strong>Kích hoạt khóa học</strong>.</li>
            <li style="margin-bottom: 12px;"><strong>Bước 4:</strong> Tìm khóa <strong>"Nhập môn tiếng Hàn chuẩn Edutech"</strong> và tiến hành học tập ngay.</li>
          </ul>
        </div>

        <div style="text-align: left; margin-top: 30px; border-top: 1px solid #eee; padding-top: 20px; font-size: 15px; color: #555;">
          <p style="margin-bottom: 10px;">- Mời bạn tham gia nhóm giải đáp thắc mắc khi học tập: <a href="https://zalo.me/g/l105v00zva5okswffcoo" style="color: #003366; text-decoration: underline;">Nhóm Zalo Hỗ Trợ</a></p>
        </div>
      </div>
      
      <div style="padding: 20px; background: #003366; color: #ffffff; text-align: center; font-size: 14px;">
        <p style="margin: 0;">© 2026 Civilis - Du học & Đào tạo quốc tế</p>
        <p style="margin: 5px 0 0 0;">Hotline/Zalo: 077.444.0000</p>
      </div>
    </div>
  `;

  GmailApp.sendEmail(email, subject, "", {
    htmlBody: htmlBody,
    name: "Civilis - Xóa mù tiếng Hàn"
  });
}
