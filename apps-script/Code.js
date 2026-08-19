// code.js
// [설정] 배포된 웹 앱 URL (CORS 헤더용 - 필요 시 수정)
// const ALLOWED_ORIGIN = "https://somyun.github.io"; // 보안 강화 시 사용
const ALLOWED_ORIGIN = "*";

// [설정] 부산교통공사 로그인 정보는 '프로젝트 설정 > 스크립트 속성'에서 관리합니다.

// [설정] 대상 스프레드시트 ID (기본값)
// 이 ID는 코드 내에서 기본값으로 사용됩니다.
const GLOBAL_TARGET_SPREADSHEET_ID = '19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s';
const FOLDER_ID = '1p0iXXAWa6iwmj2nUuHm_PntlaV9CgQ7w'; // 업로드할 폴더
const TARGET_SPREADSHEET_ID = '19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s'; // ERP점검웹앱 파일 ID
const FORM_LIST_SHEET_NAME = 'FormList'; // 양식 목록을 저장할 시트 이름

/**
 * GET 요청을 처리합니다 (데이터 조회용)
 */
function doGet(e) {
    console.log("doGet 함수 시작. 매개변수:", JSON.stringify(e.parameter));

    const action = e.parameter.action;

    // 1. 단순 웹 앱 접속 (HTML 반환, 실제 프론트는 깃허브)
    if (!action && !e.parameter.fileId) {
        return ContentService.createTextOutput("이 웹 앱은 API 서버로 동작합니다. 깃허브 프론트엔드를 이용해주세요.");
    }

    // 2. XLSX 다운로드
    if (e.parameter.fileId && e.parameter.sheetName) {
        return handleXlsxDownload(e);
    }

    // 3. API 요청 처리
    try {
        let result;
        switch (action) {
            case 'getFormList':
                result = getFormList(); // 기존 getFormList (웹앱용) 재사용
                break;

            case 'getFormDataForWeb':
                result = getFormDataForWeb(e.parameter.sheetName);
                break;

            case 'getValidationDataFromDB':
                // JSON 문자열로 넘어온 uniqueIds 파싱
                const uniqueIds = JSON.parse(e.parameter.uniqueIds || '[]');
                result = getValidationDataFromDB(uniqueIds);
                break;

            case 'getBoardData':
                result = handleGetBoardData();
                break;

            case 'getUserSettings':
                result = getUserSettings(e.parameter.token);
                break;

            case 'getPushLogs':
                result = getPushLogs(e.parameter.token);
                break;

            // 오토핫키용 레거시 지원 (필요하다면 유지)
            case 'getMeasurements':
                const data = getMeasurementData(e.parameter.sheetName);
                // Blob 리턴은 JSON 래핑 불가하므로 직접 리턴 (CORS 문제 가능성 있음, 깃허브에선 안씀)
                return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);

            default:
                result = { error: `알 수 없는 action: ${action} ` };
        }

        return createCorsResponse(result);

    } catch (err) {
        return createCorsResponse({ error: err.message, stack: err.stack });
    }
}

/**
 * POST 요청을 처리합니다 (데이터 변경/업로드용)
 */
function doPost(e) {
    console.log("doPost 함수 시작.");

    try {
        // POST 데이터 파싱 (JSON 바디)
        const requestData = JSON.parse(e.postData.contents);
        const action = requestData.action;
        let result;

        switch (action) {
            case 'uploadFileBase64':
                result = uploadFileBase64(requestData.fileData, requestData.userChoice);
                break;

            case 'saveMeasurementsToSheet':
                result = saveMeasurementsToSheet(
                    TARGET_SPREADSHEET_ID,
                    requestData.sheetName,
                    requestData.measurements
                );
                break;

            case 'registerToken':
                result = registerToken(
                    requestData.token,
                    requestData.userAgent,
                    requestData.keywords,
                    requestData.isActive
                );
                break;

            case 'requestMapAuthCode':
                result = requestMapAuthCode(requestData.email);
                break;

            case 'verifyMapAuthCode':
                result = verifyMapAuthCode(requestData.email, requestData.code);
                break;

            default:
                result = { error: `알 수 없는 POST action: ${action} ` };
        }

        return createCorsResponse(result);

    } catch (err) {
        return createCorsResponse({ error: err.message, stack: err.stack });
    }
}

/**
 * CORS 헤더를 포함한 JSON 응답을 생성하는 헬퍼 함수
 */
function createCorsResponse(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
    // .setHeader('Access-Control-Allow-Origin', '*'); // Apps Script `ContentService`에는 setHeader 없음
    // Apps Script 웹앱은 기본적으로 리다이렉트를 통해 CORS를 어느정도 허용하지만,
    // fetch 모드시 redirect: 'follow'가 필요함.
}

// -----------------------------------------------------------------------
// 아래부터는 기존 비즈니스 로직 함수들 (기존 code.js 내용 그대로 유지 + 일부 수정)
// -----------------------------------------------------------------------

function handleXlsxDownload(e) {
    // XLSX 다운로드 처리는 GET으로 URL을 열어서 처리하므로 JSON 반환이 아님.
    // 기존 로직 복사 (생략 없이 구현 필요 시 복사)
    // 여기서는 기존 code.js의 37~100라인 내용을 함수화했다고 가정

    // (기존 코드의 해당 부분 로직)
    const fileId = e.parameter.fileId;
    const filename = decodeURIComponent(e.parameter.filename || 'downloaded_file.xlsx');
    const sheetName = decodeURIComponent(e.parameter.sheetName);
    const expectedRevision = e.parameter.expectedRevision
        ? decodeURIComponent(e.parameter.expectedRevision)
        : '';
    let tempSpreadsheetId = null;

    try {
        const originalSpreadsheet = SpreadsheetApp.openById(fileId);
        if (expectedRevision) {
            assertXlsxSourceRevision(originalSpreadsheet, sheetName, expectedRevision);
        }
        const sourceSheet = originalSpreadsheet.getSheetByName(sheetName);
        if (!sourceSheet) throw new Error(`시트 '${sheetName}'을(를) 찾을 수 없습니다.`);

        const tempSpreadsheet = SpreadsheetApp.create('ExportTemp_' + new Date().getTime());
        tempSpreadsheetId = tempSpreadsheet.getId();

        const copiedSheet = sourceSheet.copyTo(tempSpreadsheet);
        copiedSheet.setName(sheetName);
        tempSpreadsheet.setActiveSheet(copiedSheet);
        tempSpreadsheet.moveActiveSheet(1);

        const defaultSheet = tempSpreadsheet.getSheetByName('Sheet1');
        if (defaultSheet) {
            tempSpreadsheet.deleteSheet(defaultSheet);
        }

        SpreadsheetApp.flush();

        const exportUrl = `https://docs.google.com/spreadsheets/d/${tempSpreadsheetId}/export?format=xlsx`;
        const response = UrlFetchApp.fetch(exportUrl, {
            headers: {
                Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
            },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            throw new Error(`엑셀 다운로드 실패: 상태코드 ${response.getResponseCode()}`);
        }

        const blob = response.getBlob().setName(filename);
        const base64Data = Utilities.base64Encode(blob.getBytes());

        const jsonOutput = {
            filename: filename,
            base64: base64Data
        };

        return createCorsResponse(jsonOutput);

    } catch (error) {
        console.error('엑셀 다운로드 처리 중 오류:', error);
        return createCorsResponse({ error: error.message });
    } finally {
        if (tempSpreadsheetId) {
            try {
                DriveApp.getFileById(tempSpreadsheetId).setTrashed(true);
            } catch (e) { }
        }
    }
}

function assertXlsxSourceRevision(spreadsheet, sheetName, expectedRevision) {
    const expectedTime = new Date(expectedRevision).getTime();
    if (!Number.isFinite(expectedTime)) throw new Error('INVALID_EXPECTED_REVISION');
    const formListSheet = spreadsheet.getSheetByName(FORM_LIST_SHEET_NAME);
    if (!formListSheet || formListSheet.getLastRow() < 2) {
        throw new Error('FORM_LIST_ENTRY_NOT_FOUND');
    }
    const rows = formListSheet.getRange(2, 1, formListSheet.getLastRow() - 1, 3).getValues();
    const normalizedSheetName = String(sheetName).normalize('NFC');
    const match = rows.find(row => String(row[0]).normalize('NFC') === normalizedSheetName);
    if (!match) throw new Error('FORM_LIST_ENTRY_NOT_FOUND');
    const actualTime = match[2] instanceof Date
        ? match[2].getTime()
        : new Date(match[2]).getTime();
    if (!Number.isFinite(actualTime)) throw new Error('INVALID_FORM_LIST_REVISION');
    if (actualTime !== expectedTime) throw new Error('SOURCE_REVISION_MISMATCH');
}

// ... (나머지 기존 함수들: initializeFormListSheet, uploadFileBase64, recordUploadedForm, getFormList, etc.)
// ... (기존 code.js의 나머지 함수들을 모두 여기에 포함시켜야 함. 사용자에게는 기존 함수들을 복사하라고 안내)

function initializeFormListSheet() {
    const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    let formListSheet = spreadsheet.getSheetByName(FORM_LIST_SHEET_NAME);

    if (!formListSheet) {
        formListSheet = spreadsheet.insertSheet(FORM_LIST_SHEET_NAME);
        formListSheet.appendRow(['Sheet Name', 'Spreadsheet ID', 'Last Modified']);
    }
}

function uploadFileBase64(fileData, userChoice) {
    // ... (기존 uploadFileBase64 코드 복사)
    console.log(`uploadFileBase64 함수 시작. 사용자 선택: ${userChoice}`);
    try {
        if (!fileData || !fileData.name || !fileData.mimeType || !fileData.data) {
            throw new Error("파일 데이터가 불완전합니다. 이름, MIME 타입, 데이터가 모두 필요합니다.");
        }

        const standardXlsxMimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), standardXlsxMimeType, fileData.name);
        const convertedFileMetadata = {
            name: fileData.name.replace(/\.[^/.]+$/, "") + '_temp_upload_' + new Date().getTime(),
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [FOLDER_ID],
            convert: true
        };
        const convertedDriveFile = Drive.Files.create(convertedFileMetadata, blob);
        const tempFileId = convertedDriveFile.id;

        const sourceSpreadsheet = SpreadsheetApp.openById(tempFileId);
        const sourceSheet = sourceSpreadsheet.getSheets()[0];
        const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const newSheetName = fileData.name.replace(/\.[^/.]+$/, "");

        const existingSheet = targetSpreadsheet.getSheetByName(newSheetName);

        if (existingSheet) {
            if (userChoice === 'overwrite') {
                targetSpreadsheet.deleteSheet(existingSheet);
            } else if (userChoice === 'preserve') {
                DriveApp.getFileById(tempFileId).setTrashed(true);
                return { success: true, preserved: true, message: `기존 '${newSheetName}' 시트가 보존되었습니다.` };
            } else {
                DriveApp.getFileById(tempFileId).setTrashed(true);
                return { success: false, requiresChoice: true, message: `'${newSheetName}' 시트가 이미 존재합니다. 덮어쓰시겠습니까?` };
            }
        }

        const newSheet = sourceSheet.copyTo(targetSpreadsheet);
        newSheet.setName(newSheetName);

        DriveApp.getFileById(tempFileId).setTrashed(true);

        recordUploadedForm(newSheetName, TARGET_SPREADSHEET_ID);
        const formData = getFormDataForWeb(newSheetName);

        return {
            success: true,
            message: `파일이 성공적으로 업로드 및 변환되었습니다: ${fileData.name}`,
            formData: formData,
            spreadsheetId: TARGET_SPREADSHEET_ID,
            sheetName: newSheetName
        };

    } catch (error) {
        console.error("파일 업로드 중 오류 발생:", error);
        try { if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e) { }
        return { success: false, message: `오류 발생: ${error.message}` };
    }
}

function recordUploadedForm(sheetName, spreadsheetId) {
    // ... (기존과 동일)
    try {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const formListSheet = spreadsheet.getSheetByName(FORM_LIST_SHEET_NAME);

        if (!formListSheet) return;

        const data = formListSheet.getDataRange().getValues();
        let foundRowIndex = -1;
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === sheetName) {
                foundRowIndex = i + 1;
                break;
            }
        }

        const now = new Date();
        const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

        if (foundRowIndex !== -1) {
            formListSheet.getRange(foundRowIndex, 3).setValue(formattedDate);
        } else {
            formListSheet.appendRow([sheetName, spreadsheetId, formattedDate]);
        }
    } catch (error) {
        console.error(error);
    }
}

function getFormList() {
    // ... (기존과 동일)
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const formListSheet = spreadsheet.getSheetByName(FORM_LIST_SHEET_NAME);

        if (!formListSheet || formListSheet.getLastRow() < 2) return [];

        const range = formListSheet.getRange(2, 1, formListSheet.getLastRow() - 1, 3);
        const values = range.getValues();

        if (!Array.isArray(values) || values.length === 0) return [];

        return values.map(row => {
            let lastModifiedDateString = '';
            if (row[2] instanceof Date) {
                lastModifiedDateString = row[2].toISOString();
            } else {
                lastModifiedDateString = String(row[2]);
            }
            return {
                sheetName: row[0],
                spreadsheetId: row[1],
                lastModifiedDate: lastModifiedDateString
            };
        });
    } catch (error) {
        console.error(error);
        return [];
    }
}

function getFormDataForWeb(sheetName) {
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);

        if (!sheet) return [];

        const startRow = 2;
        const lastRow = sheet.getLastRow();
        if (lastRow < startRow) return [];

        const dataRange = sheet.getRange(startRow, 3, lastRow - startRow + 1, 5);
        const values = dataRange.getDisplayValues();

        // 1. Prepare DB Data for lookup (Batch read)
        const dbSheet = spreadsheet.getSheetByName("DB");
        let dbDataMap = new Map();
        let dbHeaderDates = [];

        if (dbSheet && dbSheet.getLastRow() >= 2) {
            const dbLastRow = dbSheet.getLastRow();
            const dbLastCol = dbSheet.getLastColumn();

            // Fetch all data largely
            const dbValues = dbSheet.getRange(2, 1, dbLastRow - 1, dbLastCol).getDisplayValues();

            // Fetch header for dates (Row 1) if needed for 'RecentDate' formatting
            if (dbLastCol > 0) {
                dbHeaderDates = dbSheet.getRange(1, 1, 1, dbLastCol).getValues()[0];
            }

            dbValues.forEach(row => {
                const uId = row[0]; // Unique ID (Column A)
                if (uId) {
                    dbDataMap.set(uId, row);
                }
            });
        }

        const formData = [];
        values.forEach(row => {
            if (row[1]) {
                const itemObj = {
                    uniqueId: row[0],
                    location: row[1],
                    item: row[2],
                    value: row[3],
                    unit: row[4],
                    validation: null,
                    recentInfo: null
                };

                // 2. Lookup Validation Data
                if (itemObj.uniqueId && dbDataMap.has(itemObj.uniqueId)) {
                    const dbRow = dbDataMap.get(itemObj.uniqueId);

                    // Min/Max values (Index 5: F, Index 4: E based on previous logic observation)
                    // Note: User suspected column mismatch but asked to just optimize flow first.
                    // Keeping existing column index logic: E(4)=Max, F(5)=Min from previous getValidationDataFromDB
                    const maxValue = dbRow[4];
                    const minValue = dbRow[5];

                    let recentValue = null;
                    let recentDate = null;

                    // Recent value logic: Find first non-empty cell starting from Index 11 (L Column)
                    for (let j = 11; j < dbRow.length; j++) {
                        if (dbRow[j] && dbRow[j].toString().trim() !== '') {
                            recentValue = dbRow[j];
                            // Format date from header
                            if (dbHeaderDates && dbHeaderDates[j]) {
                                try {
                                    // Use Utilities.formatDate for server-side formatting
                                    recentDate = Utilities.formatDate(new Date(dbHeaderDates[j]), Session.getScriptTimeZone(), "M월d일");
                                } catch (e) { recentDate = ""; }
                            }
                            break;
                        }
                    }

                    itemObj.validation = {
                        minValue: minValue,
                        maxValue: maxValue
                    };
                    itemObj.recentInfo = {
                        value: recentValue,
                        date: recentDate
                    };
                }

                formData.push(itemObj);
            }
        });
        return formData;
    } catch (error) {
        console.error(error);
        return [];
    }
}

function saveMeasurementsToSheet(spreadsheetId, sheetName, measurements) {
    // ... (기존과 동일)
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);

        if (!sheet) throw new Error(`시트 '${sheetName}'을(를) 찾을 수 없습니다.`);

        const startRow = 2;
        const lastRow = sheet.getLastRow();
        if (lastRow < startRow) throw new Error("시트에 데이터가 없습니다.");

        const lookupDataRange = sheet.getRange(startRow, 4, lastRow - startRow + 1, 7);
        const lookupDataValues = lookupDataRange.getDisplayValues();

        const existingValueColumnValues = sheet.getRange(startRow, 6, lastRow - startRow + 1, 1).getValues();
        const valuesToUpdate = existingValueColumnValues.map(row => [row[0]]);

        measurements.forEach(measurement => {
            const targetLocation = measurement.location;
            const targetItem = measurement.item;
            const inputValue = measurement.value;

            for (let i = 0; i < lookupDataValues.length; i++) {
                const sheetLocation = lookupDataValues[i][0];
                const sheetItem = lookupDataValues[i][1];

                if (sheetLocation.trim().toLowerCase() === targetLocation.trim().toLowerCase() &&
                    sheetItem.trim().toLowerCase() === targetItem.trim().toLowerCase()) {

                    if (inputValue !== null && inputValue !== undefined && inputValue !== "") {
                        const decimalPlacesRaw = lookupDataValues[i][6];
                        const decimalPlaces = parseInt(decimalPlacesRaw, 10);
                        let finalValue;

                        if (Number.isInteger(decimalPlaces) && decimalPlaces >= 0) {
                            finalValue = Number(inputValue).toFixed(decimalPlaces);
                        } else {
                            finalValue = String(inputValue);
                        }
                        valuesToUpdate[i][0] = finalValue;
                    } else {
                        valuesToUpdate[i][0] = ""
                    }
                    break;
                }
            }
        });

        sheet.getRange(startRow, 6, valuesToUpdate.length, 1)
            .setNumberFormat('@')
            .setValues(valuesToUpdate);

        recordUploadedForm(sheetName, TARGET_SPREADSHEET_ID);

        return { success: true, message: `측정값이 시트에 성공적으로 저장되었습니다.` };

    } catch (error) {
        return { success: false, message: `오류 발생: ${error.message}` };
    }
}

function getValidationDataFromDB(uniqueIds) {
    // ... (기존과 동일)
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const dbSheet = spreadsheet.getSheetByName("DB");
        if (!dbSheet) return {};
        const dbLastRow = dbSheet.getLastRow();
        if (dbLastRow < 2) return {};

        const lastColumn = dbSheet.getLastColumn();
        const dbRange = dbSheet.getRange(2, 1, dbLastRow - 1, lastColumn);
        const dbValues = dbRange.getDisplayValues();

        const validationData = {};

        uniqueIds.forEach(uniqueId => {
            if (!uniqueId) return;

            for (let i = 0; i < dbValues.length; i++) {
                const dbUniqueId = dbValues[i][0];

                if (dbUniqueId === uniqueId) {
                    const minValue = dbValues[i][5];
                    const maxValue = dbValues[i][4];

                    let recentValue = null;
                    let recentDate = null;
                    for (let j = 11; j < dbValues[i].length; j++) {
                        if (dbValues[i][j] && dbValues[i][j].toString().trim() !== '') {
                            recentValue = dbValues[i][j];
                            recentDate = Utilities.formatDate(dbSheet.getRange(1, j + 1).getValue(), Session.getScriptTimeZone(), "M월d일");
                            break;
                        }
                    }

                    validationData[uniqueId] = {
                        minValue: minValue,
                        maxValue: maxValue,
                        recentValue: recentValue,
                        recentDate: recentDate
                    };
                    break;
                }
            }
        });
        return validationData;

    } catch (error) {
        console.error(error);
        return {};
    }
}

// --- FCM Token Management ---
function registerToken(token, userAgent, keywords = "", isActive = true) {
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheetName = 'FCM_Tokens';
        let sheet = spreadsheet.getSheetByName(sheetName);

        if (!sheet) {
            sheet = spreadsheet.insertSheet(sheetName);
            sheet.appendRow(['Token', 'UserAgent', 'LastUpdated', 'Keywords', 'IsActive']);
        } else {
            // 헤더 확인 및 보강
            const headers = sheet.getRange(1, 1, 1, 5).getValues()[0];
            if (headers.length < 5 || headers[4] === "") {
                sheet.getRange(1, 4).setValue('Keywords');
                sheet.getRange(1, 5).setValue('IsActive');
            }
        }

        const data = sheet.getDataRange().getValues();
        let foundRow = -1;

        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === token) {
                foundRow = i + 1;
                break;
            }
        }

        const now = new Date();
        if (foundRow > 0) {
            sheet.getRange(foundRow, 2).setValue(userAgent);
            sheet.getRange(foundRow, 3).setValue(now);
            sheet.getRange(foundRow, 4).setValue(keywords);
            sheet.getRange(foundRow, 5).setValue(isActive);
        } else {
            sheet.appendRow([token, userAgent, now, keywords, isActive]);
        }

        return { success: true, message: 'Settings saved successfully' };
    } catch (e) {
        console.error('Registration failed', e);
        return { success: false, message: e.message };
    }
}

function getUserSettings(token) {
    if (!token) return { success: false, message: 'Token missing' };
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName('FCM_Tokens');
        if (!sheet) return { success: true, keywords: "", isActive: false };

        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === token) {
                const isActiveRaw = data[i][4];
                let isActive = false;
                if (isActiveRaw === true) {
                    isActive = true;
                } else if (typeof isActiveRaw === 'string') {
                    isActive = isActiveRaw.trim().toLowerCase() === 'true';
                }

                return {
                    success: true,
                    keywords: data[i][3] || "",
                    isActive: isActive
                };
            }
        }
        return { success: true, keywords: "", isActive: false };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

/**
 * 키워드 매칭 로직 (쉼표: OR, 공백: AND)
 * 예: "전기사업소, 경전철 전기"
 * => ("전기사업소" 포함) OR ("경전철" AND "전기" 포함)
 */
function checkKeywordMatch(postTitle, keywordString) {
    if (!keywordString || keywordString.trim() === "") return true; // 키워드 없으면 전체 알림

    const orGroups = keywordString.split(',').map(s => s.trim()).filter(s => s !== "");

    return orGroups.some(group => {
        const andTerms = group.split(/\s+/).filter(s => s !== "");
        return andTerms.every(term => postTitle.toLowerCase().includes(term.toLowerCase()));
    });
}

// --- Real Data Notification Dispatch ---
// --- Real Data Notification Dispatch ---
function checkNewPosts() {
    console.log("새 게시글 확인 시작...");
    const scriptProperties = PropertiesService.getScriptProperties();
    let lastPostId = parseInt(scriptProperties.getProperty('LAST_POST_ID') || '0', 10);

    const boardData = fetchBoardList(loginToHumetro()); // hugether.js 함수 호출

    if (!boardData || boardData.length === 0) {
        console.log("게시글을 가져오지 못했습니다.");
        return;
    }

    // 새 글 필터링 (저장된 ID보다 큰 ID)
    const newPosts = boardData.filter(post => post.id > lastPostId);

    if (newPosts.length === 0) {
        console.log("새로운 게시글이 없습니다.");
        return;
    }

    // ID 오름차순 정렬 (옛날 글 -> 최신 글 순으로 알림 발송)
    newPosts.sort((a, b) => a.id - b.id);

    console.log(`새 글 발견: ${newPosts.length}개`);

    // 알림 발송
    sendNotificationsForNewPosts(newPosts);

    // 마지막 ID 갱신 (가장 큰 ID)
    const maxId = newPosts[newPosts.length - 1].id;
    scriptProperties.setProperty('LAST_POST_ID', maxId.toString());
    console.log(`LAST_POST_ID 갱신: ${maxId}`);
}

function sendNotificationsForNewPosts(posts) {
    if (!posts || posts.length === 0) return;

    const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName('FCM_Tokens');
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    posts.forEach(post => {
        const postTitle = post.title;
        const postLink = post.link;

        for (let i = 1; i < data.length; i++) {
            const token = data[i][0];
            const keywords = data[i][3] || "";
            const isActive = data[i][4];

            // 알람 꺼둔 사용자 제외
            if (isActive === false || isActive === "FALSE" || isActive === "false") continue;

            if (checkKeywordMatch(postTitle, keywords)) {
                console.log(`알림 매칭 성공: [${postTitle}] -> ${token.substring(0, 10)}...`);

                const sendResult = sendFCMNotification(
                    token,
                    "해피휴게더",
                    postTitle, // 본문에 게시글 제목 넣기
                    "https://www.humetro.busan.kr/homepage/default/img/common/logo.png", // 아이콘
                    postLink // 클릭 시 이동 링크
                );

                logFCMNotification(token, postTitle, postLink, keywords, sendResult.success, sendResult.message);
            }
        }
    });
}

/**
 * 10분마다 실행될 트리거 설정 (한 번만 실행 필요)
 */
function setupTrigger() {
    // 기존 트리거 삭제 (중복 방지)
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === 'checkNewPosts') {
            ScriptApp.deleteTrigger(trigger);
        }
    }

    // 새 트리거 생성 (10분 간격)
    ScriptApp.newTrigger('checkNewPosts')
        .timeBased()
        .everyMinutes(10)
        .create();

    console.log("checkNewPosts 트리거가 설정되었습니다 (10분 간격).");
}

// 테스트용 함수 업데이트
function testHugetherPush() {
    // 테스트용: 최근 글을 가져와서 강제로 알림 전송 (LAST_POST_ID 무시)
    const boardData = fetchBoardList(loginToHumetro());
    if (boardData && boardData.length > 0) {
        console.log("[테스트] 가장 최근 글 1개로 알림 전송 테스트");
        sendNotificationsForNewPosts([boardData[0]]);
    }
}

// --- 푸시 알림 로그 관리 ---
function logFCMNotification(token, title, link, keywords, success, errorMessage) {
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheetName = 'FCM_Logs';
        let sheet = spreadsheet.getSheetByName(sheetName);

        if (!sheet) {
            sheet = spreadsheet.insertSheet(sheetName);
            sheet.appendRow(['Timestamp', 'Token', 'Title', 'Link', 'MatchedKeywords', 'Success', 'ErrorMessage']);
            // 첫번째 행 강조 및 고정
            sheet.getRange(1, 1, 1, 7).setBackground('#e0e0e0').setFontWeight('bold');
            sheet.setFrozenRows(1);
        }

        const now = new Date();
        sheet.appendRow([now, token, title, link, keywords, success ? "성공" : "실패", errorMessage || ""]);

    } catch (e) {
        console.error('로그 저장 실패', e);
    }
}

function getPushLogs(token) {
    if (!token) return { success: false, message: 'Token missing' };
    try {
        const spreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName('FCM_Logs');
        if (!sheet) return { success: true, logs: [] };

        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return { success: true, logs: [] };

        const logs = [];
        // 역순 탐색 (최근 기록부터 상위 50개)
        for (let i = data.length - 1; i >= 1; i--) {
            const rowToken = data[i][1];
            if (rowToken === token) {
                let timestampStr = "";
                if (data[i][0] instanceof Date) {
                    timestampStr = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "MM/dd HH:mm:ss");
                } else {
                    timestampStr = String(data[i][0]);
                }

                logs.push({
                    timestamp: timestampStr,
                    title: data[i][2],
                    link: data[i][3],
                    keywords: data[i][4],
                    status: data[i][5],
                    error: data[i][6]
                });

                if (logs.length >= 50) break; // 최대 50개까지만 호출
            }
        }
        return { success: true, logs: logs };
    } catch (e) {
        return { success: false, message: e.message };
    }
}
