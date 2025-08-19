// ==========================================
// Google Apps Script สำหรับรับข้อมูลจากแบบประเมิน
// ==========================================

// !!! สำคัญ: ต้องเปลี่ยน ID นี้เป็น ID ของ Google Sheets ของคุณ !!!
const SPREADSHEET_ID = 'SPREADSHEET_ID';

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
    
    console.log('Raw data received:', e.postData.contents);
    
    // แปลง JSON string เป็น object
    const data = JSON.parse(e.postData.contents);
    console.log('Parsed data:', JSON.stringify(data, null, 2));
    
    // เปิด Google Sheets
    console.log('Opening spreadsheet with ID:', SPREADSHEET_ID);
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getActiveSheet();
    console.log('Sheet name:', sheet.getName());
    console.log('Sheet dimensions:', sheet.getMaxRows(), 'x', sheet.getMaxColumns());
    
    // หาแถวล่าสุดที่มีข้อมูล
    const lastRow = sheet.getLastRow();
    console.log('Current last row with data:', lastRow);
    
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
      data.courseTitle || 'รายงานการประชุมยุคใหม่ สั่งได้ด้วย AI & MS Teams',  // Column V
      data.courseDate || '6 สิงหาคม 2568',                 // Column W
      data.fullnameWithTitle || ''                          // Column X
    ];
    
    console.log('Prepared row data:', rowData);
    console.log('Number of columns:', rowData.length);
    
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
      'message': 'บันทึกข้อมูลเรียบร้อยแล้ว',
      'timestamp': timestamp.toISOString(),
      'row': targetRow,
      'data_received': {
        'fullname': data.fullnameWithTitle,
        'email': data.email,
        'organization': data.organization
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
// 2. Function สำหรับตั้งค่าหัวตาราง (รันครั้งเดียว)
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
      'ชื่อเต็ม (พร้อมคำนำหน้า)'
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
    
    // คอลัมน์คะแนนเฉลี่ย (17-21)
    for (let i = 17; i <= 21; i++) {
      sheet.setColumnWidth(i, 120);
    }
    
    sheet.setColumnWidth(22, 250); // หลักสูตร
    sheet.setColumnWidth(23, 120); // วันที่อบรม
    sheet.setColumnWidth(24, 200); // ชื่อเต็ม
    
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
// 3. Function ทดสอบการส่งข้อมูล
// ==========================================
function testDoPost() {
  console.log('=== Starting Test ===');
  
  const testData = {
    postData: {
      contents: JSON.stringify({
        title: 'นาย',
        fullname: 'ทดสอบ ระบบ',
        email: 'test@example.com',
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
        fullnameWithTitle: 'นายทดสอบ ระบบ'
      })
    }
  };
  
  console.log('Test data prepared:', JSON.stringify(testData, null, 2));
  
  try {
    const result = doPost(testData);
    const content = result.getContent();
    console.log('Test result:', content);
    
    const parsed = JSON.parse(content);
    if (parsed.result === 'success') {
      console.log('✅ Test PASSED - Data saved to row:', parsed.row);
    } else {
      console.log('❌ Test FAILED:', parsed.error);
    }
  } catch (error) {
    console.error('❌ Test ERROR:', error);
  }
  
  console.log('=== End Test ===');
}

// ==========================================
// 4. Function ตรวจสอบการเชื่อมต่อ
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
    
    return true;
  } catch (error) {
    console.error('❌ Connection failed:', error.message);
    return false;
  }
}

// ==========================================
// 5. Helper Functions
// ==========================================
function calculateAverage(scores) {
  const validScores = scores
    .filter(score => score && !isNaN(parseInt(score)))
    .map(score => parseInt(score));
  
  if (validScores.length === 0) return 0;
  
  const sum = validScores.reduce((a, b) => a + b, 0);
  return Math.round((sum / validScores.length) * 100) / 100; // Round to 2 decimal places
}

// ==========================================
// 6. Function ดูข้อมูลล่าสุด
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
// 7. Function ล้างข้อมูลทดสอบ (ระวังใช้!)
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