const DRIVE_FOLDER_ID = '1ha0GBUdY2V_8uk_lpckXcu8_DeovmMZR';
const SHEET_ID        = '1YPvbtBzGgiKx3_4P7egMzYsqqLt9C0sVTwtyDcIVsRg';
const SHEET_NAME      = '報名資料';

const HEADERS = [
  '時間戳記',
  '中文姓名',
  '稱呼/暱稱',
  '出生年月日',
  '聯絡電話',
  'Email',
  'Instagram',
  '所屬店家/學校',
  '美髮年資',
  '創作理念',
  '染膏配方',
  '照片資料夾連結',
  '照片數量'
];

function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    const sheet = getOrCreateSheet();

    let folderUrl = '', photoCount = 0;
    try {
      const driveResult = savePhotosToDrive(data);
      folderUrl  = driveResult.folderUrl;
      photoCount = driveResult.photoCount;
    } catch (driveErr) {
      console.error('Drive 上傳失敗：', driveErr);
      folderUrl = '照片上傳失敗：' + driveErr.message;
    }

    sheet.appendRow([
      data.timestamp || new Date().toISOString(),
      data.name_zh   || '',
      data.name_en   || '',
      data.birthday  || '',
      data.phone     || '',
      data.email     || '',
      data.ig        || '',
      data.salon     || '',
      data.years     || '',
      data.intro     || '',
      data.formula   || '',
      folderUrl,
      photoCount
    ]);

    return jsonResponse({ result: 'success' });

  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ result: 'error', error: err.message });
  }
}

function getOrCreateSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1a2438')
      .setFontColor('#ffffff');
  }
  return sheet;
}

function savePhotosToDrive(data) {
  const photos = data.photos || [];
  if (photos.length === 0) {
    return { folderUrl: '', photoCount: 0 };
  }

  const parentFolder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const safeTimestamp = (data.timestamp || new Date().toISOString())
    .replace(/[:.]/g, '-')
    .slice(0, 19);
  const folderName   = `${data.name_zh || 'unknown'}_${safeTimestamp}`;
  const applicantDir = parentFolder.createFolder(folderName);

  // setSharing 已移除：Workspace 政策在 web app context 中禁止設定 ANYONE_WITH_LINK

  for (const photo of photos) {
    try {
      const decoded = Utilities.base64Decode(photo.data);
      const blob    = Utilities.newBlob(decoded, photo.mimeType, photo.name);
      applicantDir.createFile(blob);
    } catch (photoErr) {
      console.warn('照片上傳失敗：' + photo.name, photoErr);
    }
  }

  return {
    folderUrl:  applicantDir.getUrl(),
    photoCount: photos.length
  };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testDoPost() {
  DriveApp.getFolderById(DRIVE_FOLDER_ID);

  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        timestamp: new Date().toISOString(),
        name_zh:  '測試姓名',
        name_en:  'Test',
        birthday: '2000/01/01',
        phone:    '0912000000',
        email:    'test@test.com',
        ig:       '@test_ig',
        salon:    'FIOLE 測試店',
        years:    '1～2年',
        intro:    '這是一則測試投稿。',
        formula:  'Fiole 測試色號',
        photos:   []
      })
    }
  };
  const result = doPost(fakeEvent);
  Logger.log(result.getContent());
}
