function updateMeasurementsToNewColumn() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheetName = "DB";
  const dbSheet = spreadsheet.getSheetByName(dbSheetName);

  if (!dbSheet) {
    Logger.log("경고: 'DB'라는 이름의 시트를 찾을 수 없습니다. 시트 이름을 확인해 주세요.");
    return;
  }

  // 현재 날짜를 제목으로 사용하여 새 열 삽입
  const currentDate = new Date();
  const formattedDate = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yy-MM-dd");
  const newColumnIndex = 12; // L열은 12번째 열

  dbSheet.insertColumnAfter(newColumnIndex - 1); // L열 위치에 새 열 삽입 (insertColumnAfter는 지정된 열 다음 열에 삽입하므로 -1)
  dbSheet.getRange(1, newColumnIndex).setValue(formattedDate); // 새로 삽입된 열의 1행에 날짜 제목 설정

  // 'DB' 시트의 A열 데이터 가져오기 (측정지점)
  const dbSheetLastRow = dbSheet.getLastRow();
  if (dbSheetLastRow < 1) {
    Logger.log("'DB' 시트에 데이터가 없습니다.");
    return;
  }
  const measurementPoints = dbSheet.getRange("A2:A" + dbSheetLastRow).getValues();

  // FormList 시트에서 오늘 날짜와 일치하는 시트 목록 가져오기
  const formListSheet = spreadsheet.getSheetByName(FORM_LIST_SHEET_NAME);
  if (!formListSheet) {
    Logger.log("경고: 'FormList' 시트를 찾을 수 없습니다.");
    return;
  }

  const formListLastRow = formListSheet.getLastRow();
  if (formListLastRow < 2) {
    Logger.log("FormList 시트에 데이터가 없습니다.");
    return;
  }

  // FormList 시트에서 A열(시트이름)과 C열(LastModified) 데이터 가져오기
  const formListRange = formListSheet.getRange(2, 1, formListLastRow - 1, 3); // A2:C마지막행
  const formListValues = formListRange.getValues();

  // 오늘 날짜와 일치하는 시트 목록 생성
  const todaySheets = [];
  for (let i = 0; i < formListValues.length; i++) {
    const sheetName = formListValues[i][0]; // A열 (시트이름)
    const lastModified = formListValues[i][2]; // C열 (LastModified)
    
    // LastModified가 Date 객체인지 확인하고 오늘 날짜와 비교
    if (lastModified instanceof Date) {
      const lastModifiedFormatted = Utilities.formatDate(lastModified, Session.getScriptTimeZone(), "yy-MM-dd");
      if (lastModifiedFormatted === formattedDate) {
        todaySheets.push(sheetName);
      }
    }
  }

  Logger.log(`오늘 날짜(${formattedDate})와 일치하는 시트 목록: ${todaySheets.join(', ')}`);

  // 오늘 날짜와 일치하는 시트만 순회
  for (let i = 0; i < todaySheets.length; i++) {
    const sheetName = todaySheets[i];
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`경고: '${sheetName}' 시트를 찾을 수 없습니다.`);
      continue;
    }

    const sheetLastRow = sheet.getLastRow();
    if (sheetLastRow < 1) {
      Logger.log(`'${sheetName}' 시트에 데이터가 없습니다.`);
      continue; // 시트에 데이터가 없으면 다음 시트로
    }

    const dataRange = sheet.getRange("C2:F" + sheetLastRow); // C열부터 F열까지 데이터 가져오기
    const sheetValues = dataRange.getValues();

    Logger.log(`'${sheetName}' 시트에서 데이터를 처리 중...`);

    // 'DB' 시트의 각 측정지점에 대해 다른 시트에서 일치하는 값 찾기
    for (let row = 0; row < measurementPoints.length; row++) {
      const currentMeasurementPoint = measurementPoints[row][0]; // A열 값 (측정지점)

      for (let otherRow = 0; otherRow < sheetValues.length; otherRow++) {
        const otherSheetCValue = sheetValues[otherRow][0]; // C열 값 (배열 인덱스 0)
        const otherSheetFValue = sheetValues[otherRow][3]; // F열 값 (배열 인덱스 3)

        // 측정지점이 동일한 경우
        if (currentMeasurementPoint === otherSheetCValue && currentMeasurementPoint !== "") {
          // 'DB' 시트의 새로 삽입된 열에 L열 값 작성 (행 인덱스는 0부터 시작하므로 +1)
          dbSheet.getRange(row+2, newColumnIndex).setValue(otherSheetFValue);
          break; // 첫 번째로 찾은 값만 복사하고 다음 측정지점으로 이동
        }
      }
    }
    Logger.log(`'${sheetName}' 시트를 DB 시트에 기록했습니다.`);
  }
}