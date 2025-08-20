// ==========================================
// Google Apps Script สำหรับรับข้อมูลจากแบบประเมินและบันทึก Certificate
// ==========================================

// !!! สำคัญ: ต้องเปลี่ยน ID นี้เป็น ID ของ Google Sheets ของคุณ !!!
const SPREADSHEET_ID = 'SPREADSHEET_ID';

// !!! สำคัญ: ต้องเปลี่ยน ID นี้เป็น ID ของ Folder ที่จะเก็บ Certificate !!!
// วิธีหา Folder ID: เปิด Google Drive ไปที่โฟลเดอร์ที่ต้องการ 
// URL จะเป็น: https://drive.google.com/drive/folders/FOLDER_ID_HERE
const CERTIFICATE_FOLDER_ID = 'CERTIFICATE_FOLDER_ID'; // ใส่ ID ของโฟลเดอร์ certificates

// ==========================================
// 1. Function หลักสำหรับรับข้อมูลจาก Web Form
// ==========================================
function doPost(e) {
  try {
    console.log('=== Start doPost ===');
    console.log('Request received at:', new Date());
    
    // ตรวจสอบว่ามีข้อมูลส่งมาหรือไม่
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }
    
    console.log('Raw data received:', e.postData.contents.substring(0, 500) + '...'); // Log first 500 chars
    
    // แปลง JSON string เป็น object
    const data = JSON.parse(e.postData.contents);
    console.log('Parsed data fields:', Object.keys(data));
    console.log('Email:', data.email);
    console.log('Has certificate image:', !!data.certificateImage);
    
    // เปิด Google Sheets
    console.log('Opening spreadsheet with ID:', SPREADSHEET_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    console.log('Sheet name:', sheet.getName());
    
    // หาแถวล่าสุดที่มีข้อมูล
    const lastRow = sheet.getLastRow();
    console.log('Current last row with data:', lastRow);
    
    // =============================================
    // บันทึก Certificate ไปยัง Google Drive
    // =============================================
    let certificateUrl = '';
    let certificateFileId = '';
    let certificateDownloadUrl = '';
    
    if (data.certificateImage) {
      try {
        console.log('=== Saving Certificate to Google Drive ===');
        
        // สร้างหรือเข้าถึงโฟลเดอร์
        const certificateFolder = getCertificateFolder();
        
        // สร้างชื่อไฟล์
        const fileName = `${data.email}.png`;
        console.log('Certificate filename:', fileName);
        
        // ตรวจสอบว่ามีไฟล์ชื่อนี้อยู่แล้วหรือไม่
        const existingFiles = certificateFolder.getFilesByName(fileName);
        if (existingFiles.hasNext()) {
          // ถ้ามีไฟล์อยู่แล้ว ให้ลบไฟล์เก่าก่อน
          const oldFile = existingFiles.next();
          console.log('Removing existing certificate file...');
          certificateFolder.removeFile(oldFile);
          // ลบไฟล์ถาวร
          DriveApp.getFileById(oldFile.getId()).setTrashed(true);
        }
        
        // แปลง base64 เป็น Blob
        const blob = Utilities.newBlob(
          Utilities.base64Decode(data.certificateImage),
          'image/png',
          fileName
        );
        
        // สร้างไฟล์ใน Google Drive
        const file = certificateFolder.createFile(blob);
        
        // ตั้งค่าการแชร์ (optional - ถ้าต้องการให้เข้าถึงได้ด้วย link)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // เก็บ URL และ File ID
        certificateUrl = file.getUrl();
        certificateFileId = file.getId();
        certificateDownloadUrl = `https://drive.google.com/uc?export=download&id=${certificateFileId}`;
        
        console.log('Certificate saved successfully!');
        console.log('File ID:', certificateFileId);
        console.log('File URL:', certificateUrl);
        console.log('Download URL:', certificateDownloadUrl);
        
        // =============================================
        // ส่งอีเมลพร้อม Certificate
        // =============================================
        if (data.email) {
          try {
            sendCertificateEmail(
              data.email,
              data.fullnameWithTitle || data.fullname,
              data.courseTitle || 'รายงานการประชุมยุคใหม่ สั่งได้ด้วย AI & MS Teams',
              data.courseDate || '6 สิงหาคม 2568',
              certificateUrl,
              certificateDownloadUrl,
              file
            );
            console.log('✅ Certificate email sent successfully to:', data.email);
          } catch (emailError) {
            console.error('⚠️ Error sending email:', emailError);
            // ไม่ให้ error email หยุดการทำงานหลัก
          }
        }
        
      } catch (certError) {
        console.error('Error saving certificate:', certError);
        console.error('Certificate error details:', certError.toString());
        // ไม่ให้ error นี้หยุดการบันทึกข้อมูลแบบประเมิน
        certificateUrl = 'Error: ' + certError.message;
      }
    }
    
    // =============================================
    // บันทึกข้อมูลแบบประเมินลง Sheet
    // =============================================
    
    // เตรียมข้อมูลสำหรับบันทึก
    const timestamp = new Date();
    const rowData = [
      timestamp,                          // Column A: Timestamp
      data.title || '',                   // Column B: คำนำหน้า
      data.fullname || '',                // Column C: ชื่อ-นามสกุล
      data.email || '',                   // Column D: อีเมล
      data.organization || '',            // Column E: หน่วยงาน
      data.q1_1 || '',                    // Column F: 1.1 หลักสูตรตรงตามวัตถุประสงค์
      data.q1_2 || '',                    // Column G: 1.2 นำไปใช้ประโยชน์ได้
      data.q1_3 || '',                    // Column H: 1.3 ระยะเวลาเหมาะสม
      data.q2_1 || '',                    // Column I: 2.1 วิทยากรมีความรู้
      data.q2_2 || '',                    // Column J: 2.2 ถ่ายทอดชัดเจน
      data.q2_3 || '',                    // Column K: 2.3 ตอบคำถามตรงประเด็น
      data.q3_1 || '',                    // Column L: 3.1 ความพร้อมสถานที่
      data.q3_2 || '',                    // Column M: 3.2 เจ้าหน้าที่ให้บริการดี
      data.q4_1 || '',                    // Column N: 4.1 ความรู้ก่อนอบรม
      data.q4_2 || '',                    // Column O: 4.2 ความรู้หลังอบรม
      data.suggestions || '',             // Column P: ข้อเสนอแนะ
      data.courseTitle || 'รายงานการประชุมยุคใหม่ สั่งได้ด้วย AI & MS Teams',  // Column Q: หลักสูตร
      data.courseDate || '6 สิงหาคม 2568',     // Column R: วันที่อบรม
      data.fullnameWithTitle || '',             // Column S: ชื่อเต็ม (พร้อมคำนำหน้า)
      certificateFileId,                         // Column T: Certificate File ID
      certificateUrl                             // Column U: Certificate URL
    ];
    
    console.log('Prepared row data with', rowData.length, 'columns');
    
    // บันทึกข้อมูลลง sheet
    const targetRow = lastRow + 1;
    console.log('Writing to row:', targetRow);
    
    // ตรวจสอบว่า sheet มี column พอหรือไม่
    const requiredColumns = rowData.length;
    const currentColumns = sheet.getMaxColumns();
    if (currentColumns < requiredColumns) {
      console.log(`Inserting ${requiredColumns - currentColumns} columns`);
      sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
    }
    
    // เขียนข้อมูล
    sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    console.log('Data written successfully to row', targetRow);
    
    // บันทึกและ flush
    SpreadsheetApp.flush();
    
    // ส่งผลลัพธ์กลับ
    const response = {
      'result': 'success',
      'message': 'บันทึกข้อมูลและ Certificate เรียบร้อยแล้ว',
      'timestamp': timestamp.toISOString(),
      'row': targetRow,
      'data_received': {
        'fullname': data.fullnameWithTitle,
        'email': data.email,
        'organization': data.organization,
        'certificate_saved': !!certificateFileId,
        'certificate_url': certificateUrl,
        'email_sent': true
      }
    };
    
    console.log('Sending response:', response);
    console.log('=== End doPost (Success) ===');
    
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('=== Error in doPost ===');
    console.error('Error message:', error.message);
    console.error('Error stack:', error.stack);
    console.error('Error details:', error.toString());
    
    const errorResponse = {
      'result': 'error',
      'error': error.toString(),
      'message': error.message,
      'stack': error.stack,
      'timestamp': new Date().toISOString()
    };
    
    return ContentService
      .createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 2. Function สำหรับส่งอีเมลพร้อม Certificate
// ==========================================
function sendCertificateEmail(recipientEmail, fullName, courseTitle, courseDate, certificateUrl, downloadUrl, certificateFile) {
  try {
    console.log('=== Sending Certificate Email ===');
    console.log('To:', recipientEmail);
    console.log('Name:', fullName);
    
    // หัวข้ออีเมล
    const subject = `🎓 Certificate - ${courseTitle}`;
    
    // เนื้อหาอีเมล HTML
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f5f5f5;
      margin: 0;
      padding: 0;
    }
    .container {
      max-width: 600px;
      margin: 20px auto;
      background-color: white;
      border-radius: 10px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 30px;
      text-align: center;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .content {
      padding: 30px;
    }
    .greeting {
      font-size: 18px;
      color: #333;
      margin-bottom: 20px;
    }
    .message {
      color: #666;
      line-height: 1.6;
      margin-bottom: 20px;
    }
    .course-info {
      background-color: #f8f9fa;
      border-left: 4px solid #667eea;
      padding: 15px;
      margin: 20px 0;
    }
    .course-info h3 {
      margin: 0 0 10px 0;
      color: #333;
      font-size: 16px;
    }
    .course-info p {
      margin: 5px 0;
      color: #666;
    }
    .download-section {
      text-align: center;
      margin: 30px 0;
    }
    .download-btn {
      display: inline-block;
      padding: 12px 30px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      text-decoration: none;
      border-radius: 25px;
      font-weight: 600;
      margin: 10px;
      transition: transform 0.2s;
    }
    .download-btn:hover {
      transform: translateY(-2px);
    }
    .secondary-btn {
      display: inline-block;
      padding: 10px 25px;
      background-color: #f0f0f0;
      color: #333;
      text-decoration: none;
      border-radius: 20px;
      font-weight: 500;
      margin: 10px;
    }
    .footer {
      background-color: #f8f9fa;
      padding: 20px;
      text-align: center;
      color: #999;
      font-size: 12px;
    }
    .logo {
      margin-bottom: 15px;
    }
    .success-icon {
      font-size: 48px;
      margin-bottom: 15px;
    }
    .divider {
      height: 1px;
      background-color: #e0e0e0;
      margin: 20px 0;
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="success-icon">🎉</div>
      <h1>ขอบคุณที่เข้าร่วมการอบรม</h1>
    </div>
    
    <div class="content">
      <div class="greeting">
        สวัสดีคุณ${fullName} 👋
      </div>
      
      <div class="message">
        ขอขอบคุณที่ท่านได้เข้าร่วมการอบรมและทำแบบประเมินเรียบร้อยแล้ว 
        ทางเราได้แนบเกียรติบัตรการเข้าร่วมอบรมมาพร้อมกับอีเมลฉบับนี้
      </div>
      
      <div class="course-info">
        <h3>📚 รายละเอียดหลักสูตร</h3>
        <p><strong>หลักสูตร:</strong> ${courseTitle}</p>
        <p><strong>วันที่อบรม:</strong> ${courseDate}</p>
        <p><strong>สถานที่:</strong> สถาบันบัณฑิตพัฒนบริหารศาสตร์ (NIDA)</p>
      </div>
      
      <div class="download-section">
        <h3 style="color: #333;">📄 ดาวน์โหลดเกียรติบัตร</h3>
        <a href="${downloadUrl}" class="download-btn">
          ⬇️ ดาวน์โหลดเกียรติบัตร (PNG)
        </a>
        <br>
        <a href="${certificateUrl}" class="secondary-btn">
          👁️ ดูเกียรติบัตรออนไลน์
        </a>
      </div>
      
      <div class="divider"></div>
      
      <div class="message" style="font-size: 14px;">
        <p>📌 <strong>หมายเหตุ:</strong></p>
        <ul style="color: #666; padding-left: 20px;">
          <li>เกียรติบัตรได้ถูกบันทึกไว้ใน Google Drive แล้ว</li>
          <li>สามารถดาวน์โหลดได้ตลอดเวลาจากลิงก์ด้านบน</li>
          <li>หากพบปัญหาในการดาวน์โหลด กรุณาติดต่อเจ้าหน้าที่</li>
        </ul>
      </div>
    </div>
    
    <div class="footer">
      <p><strong>สถาบันบัณฑิตพัฒนบริหารศาสตร์</strong></p>
      <p>National Institute of Development Administration (NIDA)</p>
      <p>📧 ติดต่อสอบถาม: idt@nida.ac.th</p>
      <div class="divider" style="margin: 15px auto; width: 50%;"></div>
      <p style="font-size: 11px; color: #aaa;">
        อีเมลฉบับนี้ถูกส่งอัตโนมัติ กรุณาอย่าตอบกลับ<br>
        © 2568 NIDA - Institute of Digital Technology
      </p>
    </div>
  </div>
</body>
</html>
    `;
    
    // เนื้อหาอีเมลแบบ Plain Text (สำหรับกรณีที่ไม่รองรับ HTML)
    const plainBody = `
สวัสดีคุณ${fullName}

ขอขอบคุณที่ท่านได้เข้าร่วมการอบรมและทำแบบประเมินเรียบร้อยแล้ว

รายละเอียดหลักสูตร:
- หลักสูตร: ${courseTitle}
- วันที่อบรม: ${courseDate}
- สถานที่: สถาบันบัณฑิตพัฒนบริหารศาสตร์ (NIDA)

ดาวน์โหลดเกียรติบัตร:
${downloadUrl}

ดูเกียรติบัตรออนไลน์:
${certificateUrl}

หมายเหตุ:
- เกียรติบัตรได้ถูกบันทึกไว้ใน Google Drive แล้ว
- สามารถดาวน์โหลดได้ตลอดเวลาจากลิงก์ด้านบน
- หากพบปัญหาในการดาวน์โหลด กรุณาติดต่อเจ้าหน้าที่

ติดต่อสอบถาม: idt@nida.ac.th

สถาบันบัณฑิตพัฒนบริหารศาสตร์
National Institute of Development Administration (NIDA)
    `;
    
    // ส่งอีเมลพร้อมแนบไฟล์
    MailApp.sendEmail({
      to: recipientEmail,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      attachments: [certificateFile],
      name: 'NIDA - Institute of Digital Technology'
    });
    
    console.log('Email sent successfully with attachment');
    
  } catch (error) {
    console.error('Error in sendCertificateEmail:', error);
    throw error;
  }
}

// ==========================================
// 3. Function สำหรับเข้าถึงโฟลเดอร์ Certificate
// ==========================================
function getCertificateFolder() {
  try {
    // ลองเข้าถึงโฟลเดอร์ด้วย ID ที่กำหนด
    const folder = DriveApp.getFolderById(CERTIFICATE_FOLDER_ID);
    console.log('Certificate folder found:', folder.getName());
    return folder;
    
  } catch (error) {
    console.error('Error accessing certificate folder:', error);
    
    // ถ้าไม่พบโฟลเดอร์ ให้สร้างโฟลเดอร์ใหม่ที่ root
    console.log('Creating new certificate folder at root...');
    const newFolder = DriveApp.createFolder('NIDA-IDT-680808-certificates');
    console.log('New folder created with ID:', newFolder.getId());
    console.log('!!! Please update CERTIFICATE_FOLDER_ID in the script !!!');
    
    return newFolder;
  }
}

// ==========================================
// 4. Function สำหรับตั้งค่าหัวตาราง (รันครั้งเดียว)
// ==========================================
function setupHeaders() {
  try {
    console.log('=== Setting up headers ===');
    console.log('Spreadsheet ID:', SPREADSHEET_ID);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    
    const headers = [
      'Timestamp',
      'คำนำหน้า',
      'ชื่อ-นามสกุล',
      'อีเมล',
      'หน่วยงาน',
      '1.1 หลักสูตรตรงตามวัตถุประสงค์',
      '1.2 นำไปใช้ประโยชน์ได้',
      '1.3 ระยะเวลาเหมาะสม',
      '2.1 วิทยากรมีความรู้',
      '2.2 ถ่ายทอดชัดเจน',
      '2.3 ตอบคำถามตรงประเด็น',
      '3.1 ความพร้อมสถานที่',
      '3.2 เจ้าหน้าที่ให้บริการดี',
      '4.1 ความรู้ก่อนอบรม',
      '4.2 ความรู้หลังอบรม',
      'ข้อเสนอแนะ',
      'หลักสูตร',
      'วันที่อบรม',
      'ชื่อเต็ม (พร้อมคำนำหน้า)',
      'Certificate File ID',
      'Certificate URL'
    ];
    
    // ตรวจสอบว่า sheet มี column พอหรือไม่
    const requiredColumns = headers.length;
    const currentColumns = sheet.getMaxColumns();
    if (currentColumns < requiredColumns) {
      sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
    }
    
    // เขียนหัวตาราง
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // จัดรูปแบบหัวตาราง
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);
    
    // ปรับความกว้างคอลัมน์
    sheet.setColumnWidth(1, 150);  // Timestamp
    sheet.setColumnWidth(2, 80);   // คำนำหน้า
    sheet.setColumnWidth(3, 150);  // ชื่อ-นามสกุล
    sheet.setColumnWidth(4, 200);  // อีเมล
    sheet.setColumnWidth(5, 200);  // หน่วยงาน
    
    // คอลัมน์คำถาม (6-15)
    for (let i = 6; i <= 15; i++) {
      sheet.setColumnWidth(i, 100);
    }
    
    sheet.setColumnWidth(16, 300); // ข้อเสนอแนะ
    sheet.setColumnWidth(17, 250); // หลักสูตร
    sheet.setColumnWidth(18, 120); // วันที่อบรม
    sheet.setColumnWidth(19, 200); // ชื่อเต็ม
    sheet.setColumnWidth(20, 150); // Certificate File ID
    sheet.setColumnWidth(21, 300); // Certificate URL
    
    // ตั้งค่าแถวหัวตาราง
    sheet.setRowHeight(1, 60);
    
    // Freeze แถวแรก
    sheet.setFrozenRows(1);
    
    console.log('Headers setup completed');
    console.log('Total columns:', headers.length);
    
  } catch (error) {
    console.error('Error in setupHeaders:', error);
    throw error;
  }
}

// ==========================================
// 5. Function ทดสอบการส่งอีเมล
// ==========================================
function testEmailSending() {
  console.log('=== Testing Email Sending ===');
  
  try {
    // ใส่อีเมลทดสอบที่นี่
    const testEmail = 'natthasath.sak@nida.ac.th'; // เปลี่ยนเป็นอีเมลของคุณเพื่อทดสอบ
    const testName = 'นายทดสอบ ระบบ';
    const testCourseTitle = 'หลักสูตรทดสอบการส่งอีเมล';
    const testCourseDate = '1 มกราคม 2568';
    const testCertUrl = 'https://drive.google.com/file/d/1VwqnqQCDAPFk5Ciyh7rNS7u0RKRG7jaf/view?usp=drive_link';
    const testDownloadUrl = 'https://drive.google.com/uc?export=download&id=test123';
    
    // สร้างไฟล์ทดสอบ
    const testBlob = Utilities.newBlob('Test Certificate Content', 'text/plain', 'test_certificate.txt');
    
    // ส่งอีเมลทดสอบ
    sendCertificateEmail(
      testEmail,
      testName,
      testCourseTitle,
      testCourseDate,
      testCertUrl,
      testDownloadUrl,
      testBlob
    );
    
    console.log('✅ Test email sent successfully to:', testEmail);
    return true;
    
  } catch (error) {
    console.error('❌ Email test failed:', error);
    return false;
  }
}

// ==========================================
// 6. Function ทดสอบการสร้าง Certificate
// ==========================================
function testCertificateSave() {
  console.log('=== Testing Certificate Save ===');
  
  try {
    // ทดสอบเข้าถึงโฟลเดอร์
    const folder = getCertificateFolder();
    console.log('✅ Folder access successful:', folder.getName());
    console.log('Folder ID:', folder.getId());
    
    // สร้างไฟล์ทดสอบ
    const testContent = 'This is a test certificate file';
    const blob = Utilities.newBlob(testContent, 'text/plain', 'test_certificate.txt');
    const file = folder.createFile(blob);
    
    console.log('✅ Test file created successfully');
    console.log('File ID:', file.getId());
    console.log('File URL:', file.getUrl());
    
    // ลบไฟล์ทดสอบ
    DriveApp.getFileById(file.getId()).setTrashed(true);
    console.log('✅ Test file removed');
    
    console.log('=== Certificate Save Test Passed ===');
    return true;
    
  } catch (error) {
    console.error('❌ Certificate Save Test Failed:', error);
    return false;
  }
}

// ==========================================
// 7. Function ทดสอบการส่งข้อมูลพร้อม Certificate และ Email
// ==========================================
function testDoPostWithCertificateAndEmail() {
  console.log('=== Starting Test with Certificate and Email ===');
  
  // สร้าง base64 string ทดสอบ (1x1 pixel PNG)
  const testImageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
  
  const testData = {
    postData: {
      contents: JSON.stringify({
        title: 'นาย',
        fullname: 'ทดสอบ ระบบ',
        email: 'natthasath.sak@nida.ac.th', // เปลี่ยนเป็นอีเมลของคุณเพื่อทดสอบ
        organization: 'หน่วยงานทดสอบ',
        q1_1: '5',
        q1_2: '4',
        q1_3: '5',
        q2_1: '5',
        q2_2: '5',
        q2_3: '4',
        q3_1: '4',
        q3_2: '5',
        q4_1: '3',
        q4_2: '5',
        suggestions: 'ทดสอบข้อเสนอแนะ - ระบบทำงานได้ดี',        
        courseTitle: 'หลักสูตรทดสอบ',
        courseDate: '1 มกราคม 2568',
        fullnameWithTitle: 'นายทดสอบ ระบบ',
        certificateImage: testImageBase64
      })
    }
  };
  
  console.log('Test data prepared with certificate image');
  
  try {
    const result = doPost(testData);
    const content = result.getContent();
    console.log('Test result:', content);
    
    const parsed = JSON.parse(content);
    if (parsed.result === 'success') {
      console.log('✅ Test PASSED - Data saved to row:', parsed.row);
      console.log('Certificate saved:', parsed.data_received.certificate_saved);
      console.log('Certificate URL:', parsed.data_received.certificate_url);
      console.log('Email sent:', parsed.data_received.email_sent);
    } else {
      console.log('❌ Test FAILED:', parsed.error);
    }
  } catch (error) {
    console.error('❌ Test ERROR:', error);
  }
  
  console.log('=== End Test ===');
}

// ==========================================
// 8. Function ตรวจสอบการเชื่อมต่อ
// ==========================================
function testConnection() {
  try {
    console.log('Testing connection to spreadsheet...');
    console.log('Spreadsheet ID:', SPREADSHEET_ID);
    
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ Successfully connected to:', spreadsheet.getName());
    
    const sheet = spreadsheet.getActiveSheet();
    console.log('Active sheet:', sheet.getName());
    console.log('Sheet dimensions:', sheet.getMaxRows(), 'rows x', sheet.getMaxColumns(), 'columns');
    console.log('Last row with data:', sheet.getLastRow());
    
    // ทดสอบเข้าถึงโฟลเดอร์ Certificate
    console.log('\nTesting certificate folder access...');
    console.log('Certificate Folder ID:', CERTIFICATE_FOLDER_ID);
    
    try {
      const folder = getCertificateFolder();
      console.log('✅ Certificate folder accessible:', folder.getName());
    } catch (folderError) {
      console.error('⚠️ Certificate folder not accessible:', folderError.message);
      console.log('A new folder will be created when needed');
    }
    
    // ทดสอบ Email quota
    console.log('\nChecking email quota...');
    const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    console.log('📧 Remaining email quota for today:', emailQuotaRemaining);
    
    return true;
  } catch (error) {
    console.error('❌ Connection failed:', error.message);
    return false;
  }
}

// ==========================================
// 9. Function ดูข้อมูลล่าสุด
// ==========================================
function viewLatestData() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 2) {
      console.log('No data rows found (only headers)');
      return;
    }
    
    console.log('=== Latest Data ===');
    console.log('Total data rows:', lastRow - 1);
    
    // ดูข้อมูล 5 แถวล่าสุด
    const rowsToShow = Math.min(5, lastRow - 1);
    const startRow = lastRow - rowsToShow + 1;
    
    console.log(`Showing last ${rowsToShow} rows:`);
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(startRow, 1, rowsToShow, sheet.getLastColumn()).getValues();
    
    data.forEach((row, index) => {
      console.log(`\n--- Row ${startRow + index} ---`);
      headers.forEach((header, colIndex) => {
        if (row[colIndex]) {
          console.log(`${header}: ${row[colIndex]}`);
        }
      });
    });
    
  } catch (error) {
    console.error('Error viewing data:', error);
  }
}

// ==========================================
// 10. Function ดู Certificate ที่บันทึกไว้
// ==========================================
function listCertificates() {
  try {
    console.log('=== Listing Certificates in Drive ===');
    
    const folder = getCertificateFolder();
    const files = folder.getFiles();
    
    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      count++;
      console.log(`${count}. ${file.getName()}`);
      console.log(`   Size: ${file.getSize()} bytes`);
      console.log(`   Created: ${file.getDateCreated()}`);
      console.log(`   URL: ${file.getUrl()}`);
      console.log('');
    }
    
    if (count === 0) {
      console.log('No certificates found in folder');
    } else {
      console.log(`Total: ${count} certificate(s)`);
    }
    
  } catch (error) {
    console.error('Error listing certificates:', error);
  }
}

// ==========================================
// 11. Function ส่งอีเมล Certificate ซ้ำ (Manual)
// ==========================================
function resendCertificateEmail(email) {
  try {
    console.log('=== Resending Certificate Email ===');
    console.log('To:', email);
    
    // หา Certificate ในโฟลเดอร์
    const folder = getCertificateFolder();
    const fileName = `${email}.png`;
    const files = folder.getFilesByName(fileName);
    
    if (!files.hasNext()) {
      console.error('Certificate not found for:', email);
      return false;
    }
    
    const file = files.next();
    const certificateUrl = file.getUrl();
    const certificateDownloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    
    // หาข้อมูลจาก Sheet
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    // หาแถวที่มี email ตรงกัน (คอลัมน์ D = index 3)
    let userData = null;
    for (let i = 1; i < data.length; i++) { // เริ่มจาก 1 เพื่อข้าม header
      if (data[i][3] === email) { // คอลัมน์ D = อีเมล
        userData = {
          fullName: data[i][18] || data[i][2], // คอลัมน์ S หรือ C
          courseTitle: data[i][16] || 'รายงานการประชุมยุคใหม่ สั่งได้ด้วย AI & MS Teams', // คอลัมน์ Q
          courseDate: data[i][17] || '6 สิงหาคม 2568' // คอลัมน์ R
        };
        break;
      }
    }
    
    if (!userData) {
      console.error('User data not found for:', email);
      return false;
    }
    
    // ส่งอีเมล
    sendCertificateEmail(
      email,
      userData.fullName,
      userData.courseTitle,
      userData.courseDate,
      certificateUrl,
      certificateDownloadUrl,
      file
    );
    
    console.log('✅ Certificate email resent successfully');
    return true;
    
  } catch (error) {
    console.error('Error resending certificate email:', error);
    return false;
  }
}

// ==========================================
// 12. Function ล้างข้อมูลทดสอบ (ระวังใช้!)
// ==========================================
function clearTestData() {
  const response = Browser.msgBox(
    'คำเตือน',
    'ต้องการล้างข้อมูลทั้งหมดยกเว้นหัวตาราง?',
    Browser.Buttons.YES_NO
  );
  
  if (response === Browser.Buttons.YES) {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
      console.log('Cleared', lastRow - 1, 'rows of data');
    } else {
      console.log('No data to clear');
    }
  }
}