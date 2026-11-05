/**
 * ระบบขอสอบแก้ตัวสำหรับนักเรียน
 * Google Apps Script Web App
 * 
 * ต้องตั้งค่า Script Properties ดังนี้:
 * - SPREADSHEET_ID
 * - TELEGRAM_TOKEN
 * - TELEGRAM_CHAT_ID
 */

const SS_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const TELEGRAM_TOKEN = PropertiesService.getScriptProperties().getProperty('TELEGRAM_TOKEN');
const TELEGRAM_CHAT_ID = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบขอสอบแก้ตัว (นักเรียน)')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * ดึงรายวิชาที่นักเรียนยังไม่ได้สอบแก้ตัว
 */
function getPendingSubjects(studentId) {
  if (!SS_ID) throw new Error('ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties');

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('RetakeResults');
  const data = sheet.getDataRange().getValues();
  
  const inputId = studentId.toString().trim().replace(/^0+/, '');
  const results = [];
  let studentName = '';

  const teacherMap = getTeacherMap();

  for (let i = 1; i < data.length; i++) {
    const sheetId = (data[i][2] || '').toString().trim().replace(/^0+/, '');
    const gradeRetake = (data[i][13] || '').toString().trim(); // คอลัมน์ N

    if (sheetId === inputId && gradeRetake === '') {
      studentName = data[i][3] || '';
      const shortName = (data[i][15] || '').toString().trim(); // คอลัมน์ P
      const teacher = teacherMap[shortName] || {};

      // ตรวจสอบว่าเคยขอไปแล้วหรือยัง (คอลัมน์ T = index 19)
      const alreadyRequested = !!(data[i][19] || '');

      results.push({
        rowIndex: i + 1,
        studentId: data[i][2],
        studentName: studentName,
        subjectCode: data[i][6] || '',
        subjectName: data[i][7] || '',
        teacherShort: shortName,
        teacherFull: teacher.fullName || shortName,
        date1: teacher.date1 || '',
        date2: teacher.date2 || '',
        alreadyRequested: alreadyRequested,
        requestedDate: data[i][20] || '' // คอลัมน์ U
      });
    }
  }

  return {
    success: results.length > 0,
    studentName: studentName,
    subjects: results
  };
}

/**
 * สร้าง Map ของครูจากชีต Teachers
 */
function getTeacherMap() {
  if (!SS_ID) throw new Error('ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties');

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('Teachers');
  const data = sheet.getDataRange().getValues();
  const map = {};

  for (let i = 1; i < data.length; i++) {
    const fullName = (data[i][1] || '').toString().trim();
    if (!fullName) continue;

    // สร้าง shortName เหมือนระบบครู
    let nameWithoutPrefix = fullName.replace(/^(นาย|นางสาว|นาง|ว่าที่ร้อยตรี|ร\.ต\.|ดร\.)\s*/, '');
    let firstName = nameWithoutPrefix.split(' ')[0];
    let shortName = 'ครู' + firstName;

    map[shortName] = {
      fullName: fullName,
      date1: (data[i][2] || '').toString().trim(), // คอลัมน์ C
      date2: (data[i][3] || '').toString().trim()  // คอลัมน์ D
    };
  }
  return map;
}

/**
 * นักเรียนกดขอสอบแก้ตัว
 * - บันทึกประวัติลงคอลัมน์ T, U, V
 * - ส่งแจ้งเตือนไป Telegram
 */
function requestRetake(payload) {
  if (!SS_ID) throw new Error('ยังไม่ได้ตั้งค่า SPREADSHEET_ID ใน Script Properties');

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('RetakeResults');
  const row = parseInt(payload.rowIndex);

  // บันทึกประวัติ
  const now = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  sheet.getRange(row, 20).setValue(now);                    // T: วันเวลาที่ขอ
  sheet.getRange(row, 21).setValue(payload.dateText);       // U: วันที่เลือก
  sheet.getRange(row, 22).setValue('นักเรียนขอสอบแก้ตัวแล้ว'); // V: สถานะ

  // ส่ง Telegram
  const message = `เรียนคุณครู${payload.teacherFull}\n` +
                  `${payload.studentName} (${payload.studentId}) ขอสอบแก้ตัว\n` +
                  `วิชา ${payload.subjectName} รหัสวิชา ${payload.subjectCode}\n` +
                  `ใน${payload.dateText}`;

  sendToTelegram(message);
  return { success: true };
}

/**
 * ส่งข้อความไป Telegram
 */
function sendToTelegram(message) {
  if (!TELEGRAM_TOKEN || !TELEGRAM_CHAT_ID) {
    throw new Error('ยังไม่ได้ตั้งค่า TELEGRAM_TOKEN หรือ TELEGRAM_CHAT_ID');
  }

  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({
      chat_id: TELEGRAM_CHAT_ID,
      text: message,
      parse_mode: 'Markdown'
    })
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    throw new Error(`Telegram Error: ${result.description}`);
  }
}
