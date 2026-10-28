/** WORK HUB — Google Apps Script backend */
const HUB = Object.freeze({ sheetName: '사이트목록', propertyKey: 'WORK_HUB_SPREADSHEET_ID', refreshSeconds: 60 });

function setupDashboard() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(HUB.propertyKey);
  let ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('CMED WORK HUB - 사이트 관리');
    props.setProperty(HUB.propertyKey, ss.getId());
  }
  let sheet = ss.getSheetByName(HUB.sheetName);
  if (!sheet) sheet = ss.insertSheet(HUB.sheetName);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 5).setValues([['카테고리', '사이트명', 'URL', '설명', '사용여부']]);
    sheet.getRange(2, 1, 9, 5).setValues([
      ['물류 · 배송','로젠택배','https://www.ilogen.com','국내 택배 배송 조회','Y'],
      ['물류 · 배송','UPS','https://www.ups.com/kr/ko','국제 특송 배송 조회','Y'],
      ['물류 · 배송','EMS','https://ems.epost.go.kr','국제우편 조회','Y'],
      ['무역 · 수출','관세청','https://www.customs.go.kr','관세 행정','Y'],
      ['무역 · 수출','UNI-PASS','https://unipass.customs.go.kr','전자통관 시스템','Y'],
      ['무역 · 수출','FTA 포털','https://www.customs.go.kr/ftaportalkor/main.do','FTA 원산지 정보','Y'],
      ['업무 관리','Google Sheets','https://sheets.google.com','스프레드시트','Y'],
      ['업무 관리','Google Drive','https://drive.google.com','파일 보관 및 공유','Y'],
      ['업무 관리','Gmail','https://mail.google.com','회사 이메일','Y']
    ]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,5).setBackground('#17324c').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setColumnWidths(1,2,180); sheet.setColumnWidth(3,360); sheet.setColumnWidth(4,220); sheet.setColumnWidth(5,110);
  sheet.getRange(2,5,Math.max(sheet.getMaxRows()-1,1),1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Y','N'],true).build());
  Logger.log('관리용 스프레드시트: ' + ss.getUrl());
  return ss.getUrl();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('CMED WORK HUB | 회사 업무 바로가기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSites() {
  const id = PropertiesService.getScriptProperties().getProperty(HUB.propertyKey);
  if (!id) throw new Error('관리자가 Apps Script에서 setupDashboard 함수를 먼저 실행해야 합니다.');
  const sheet = SpreadsheetApp.openById(id).getSheetByName(HUB.sheetName);
  if (!sheet) throw new Error('사이트목록 시트를 찾지 못했습니다.');
  const values = sheet.getDataRange().getDisplayValues();
  const categories = new Map();
  values.slice(1).forEach((row, i) => {
    const [category, name, url, description, active] = row.map(v => String(v || '').trim());
    if (!name || active.toUpperCase() !== 'Y') return;
    if (!/^https?:\/\//i.test(url)) return;
    const group = category || '기타';
    if (!categories.has(group)) categories.set(group, []);
    categories.get(group).push({ name, url, description, rowNumber: i + 2 });
  });
  return { groups: Array.from(categories, ([name, sites]) => ({name,sites})), updatedAt: new Date().toISOString(), refreshSeconds: HUB.refreshSeconds };
}
