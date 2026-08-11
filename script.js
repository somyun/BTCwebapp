console.log("Script.js 로드됨.");

// =================================================================================
// [설정] 구글 앱스 스크립트(GAS) 웹 앱 URL을 여기에 입력하세요.
// 'backend_gas_v2.js'를 웹 앱으로 배포한 후 주소를 복사해 넣으세요.
// =================================================================================
const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbzuWS4Q5kTzDRH4IBpeXBa69KngElRdArtTCzTV0NDQsB3y4oABBIzrTLuPOZH5KOPP/exec';
const SUBMIT_MEASUREMENTS_URL = 'https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net/submitMeasurements';
const GET_SUBMISSION_URL = 'https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net/getMeasurementSubmission';
const SUBMISSION_POLL_INTERVAL_MS = 1000;
const SUBMISSION_POLL_TIMEOUT_MS = 90000;
const FUNCTION_REQUEST_TIMEOUT_MS = 20000;

// --- Firebase Config ---
const firebaseConfig = {
    apiKey: "AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM",
    authDomain: "btcwebapp-551bd.firebaseapp.com",
    projectId: "btcwebapp-551bd",
    storageBucket: "btcwebapp-551bd.firebasestorage.app",
    messagingSenderId: "237989935469",
    appId: "1:237989935469:web:07fc002a5c2ab2f5858264",
    measurementId: "G-SFSSEHRPMN"
};

// VAPID Key (Public)
const VAPID_KEY = "BCIeuJhwW92Usr-QS3BFOUWnP2pZ4rqulcmZBlxXdv8Ayms7zllnqLy-jNj9NtmOrkJfE9ywMkkj0IegbKxDDmE";

// Initialize Firebase (Compat)
let messaging = null;
try {
    if (!firebase.apps.length) firebase.initializeApp(firebaseConfig);
    messaging = firebase.messaging();
    console.log("Firebase initialized.");
} catch (e) {
    console.error("Firebase initialization failed:", e);
}

// 전역 변수
let currentSheetInfo = null;
let favorites = {};
let isMeasurementDirty = false;
let preparedDownload = null;
let pendingSubmission = null;
let validationData = {};
let sortableInstance = null;
let isSortMode = false;
let isMapViewActive = false;
let selectedFormRequestId = 0;

// 양식 선택 전 홈 화면에서만 보이는 요소를 한 곳에서 관리합니다.
// 테스트 앱에도 같은 목록과 전환 함수를 두어 화면 상태가 서로 어긋나지 않게 합니다.
const HOME_ONLY_ELEMENT_IDS = [
    'favoritesSection',
    'formMessage',
    'mainToggleContainer',
    'openMapBtn'
];

function setHomeOnlyElementsVisible(isVisible) {
    HOME_ONLY_ELEMENT_IDS.forEach((id) => {
        document.getElementById(id)?.classList.toggle('home-only-hidden', !isVisible);
    });
}

// --- API 통신 헬퍼 함수 ---
async function callApi(action, method = 'GET', data = null) {
    let url = GAS_API_URL;
    const options = {
        method: method,
    };

    if (method === 'GET') {
        url += `?action=${action}`;
        if (data) {
            for (const key in data) {
                url += `&${key}=${encodeURIComponent(data[key])}`;
            }
        }
    } else if (method === 'POST') {
        options.body = JSON.stringify({ action, ...data });
        // Google Apps Script 웹 앱은 보통 text/plain으로 보내도 잘 처리하지만, 
        // fetch 특성상 리다이렉트를 따르도록 설정이 필요할 수 있음.
        options.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
    }

    try {
        const response = await fetch(url, options);
        if (!response.ok) {
            throw new Error(`서버 통신 오류: ${response.status}`);
        }
        const result = await response.json();
        return result;
    }
    catch (error) {
        console.error(`API Error (${action}):`, error);
        throw error;
    }
}

// --- Client-side Storage Helper Functions ---
function saveToStorage(key, value) {
    try {
        localStorage.setItem(key, JSON.stringify(value));
    } catch (e) {
        console.error("Error saving to localStorage", e);
        showStatus('즐겨찾기를 저장하는 데 실패했습니다.', 'error', 3000);
    }
}

function getFromStorage(key) {
    try {
        const item = localStorage.getItem(key);
        return item ? JSON.parse(item) : null;
    } catch (e) {
        console.error("Error reading from localStorage", e);
        return null;
    }
}

// --- Sidenav / Hamburger Menu Logic ---
const hamburger = document.getElementById('hamburger');
const sidenav = document.getElementById('sidenav');
const overlay = document.getElementById('overlay');

function closeMenu() {
    sidenav.classList.remove('open');
    overlay.classList.remove('visible');
}

function openMenu() {
    sidenav.classList.add('open');
    overlay.classList.add('visible');
}

if (hamburger) hamburger.addEventListener('click', openMenu);
if (overlay) overlay.addEventListener('click', closeMenu);

const openMapBtn = document.getElementById('openMapBtn');
if (openMapBtn) openMapBtn.addEventListener('click', openMapView);

async function openMapView() {
    if (isMeasurementDirty && !confirm('변경사항이 저장되지 않았습니다. 지도 화면으로 이동하시겠습니까?')) {
        return;
    }

    if (window.BWAAuth && !(await window.BWAAuth.requireMapAccess())) {
        closeMenu();
        return;
    }

    const homeView = document.getElementById('homeView');
    const mapView = document.getElementById('mapView');
    const headerTitle = document.querySelector('.header-title');

    isMapViewActive = true;
    homeView?.classList.add('view-hidden');
    mapView?.classList.remove('view-hidden');
    document.body.classList.add('map-mode');
    if (headerTitle) headerTitle.textContent = '차량기지 도면 지도';

    closeMenu();
    updateHomeButtonVisibility();
    window.BWAMap?.initialize();
}

function closeMapView() {
    const homeView = document.getElementById('homeView');
    const mapView = document.getElementById('mapView');
    const headerTitle = document.querySelector('.header-title');

    isMapViewActive = false;
    mapView?.classList.add('view-hidden');
    homeView?.classList.remove('view-hidden');
    document.body.classList.remove('map-mode');
    if (headerTitle) headerTitle.textContent = 'ERP 점검 웹앱';
    updateHomeButtonVisibility();
}

const resetBtn = document.getElementById('resetFavoritesBtn');
if (resetBtn) {
    resetBtn.addEventListener('click', function () {
        if (confirm('정말로 모든 즐겨찾기를 초기화하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
            localStorage.removeItem('favorites');
            favorites = {};
            updateFavoriteButtons();
            showStatus('즐겨찾기가 초기화되었습니다.', 'success', 3000);
            closeMenu();
        }
    });
}

// 웹페이지 로드 시 초기화
window.onload = function () {
    document.getElementById('uploadBtn')?.addEventListener('click', handleUpload);
    document.getElementById('formSelect')?.addEventListener('change', loadSelectedForm);
    loadFormList();
    initializeFavorites();
    updateHomeButtonVisibility();
    addHomeStateToHistory();

    window.addEventListener('popstate', () => {
        // 뒤로가기 시 홈 화면으로 복귀
        const formSelect = document.getElementById('formSelect');
        if (formSelect) formSelect.value = '';
        loadSelectedForm();
    });

    // [Issue 3 Fix] 새로고침/닫기 시 변경사항 경고 (beforeunload 복원)
    window.addEventListener('beforeunload', function (e) {
        if (isMeasurementDirty) {
            e.preventDefault();
            e.returnValue = '입력한 측정값이 저장되지 않았습니다. 정말로 페이지를 이동하시겠습니까?';
            return e.returnValue;
        }
    });

    // 앱 시작 시 운영 scope에 서비스 워커를 등록합니다.
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./firebase-messaging-sw.js', { scope: './' })
            .then((registration) => {
                console.log('Service Worker registered with scope:', registration.scope);
                void bootstrapFirebaseNotificationMigration(registration);
            }).catch((err) => {
                console.log('Service Worker registration failed:', err);
            });
    }

    initializeNotificationHeaderState();
};

async function bootstrapFirebaseNotificationMigration(registration) {
    if (!messaging || !window.BWANotificationStore ||
        getFromStorage('isNotificationActive') !== true ||
        typeof Notification === 'undefined' || Notification.permission !== 'granted') return;
    try {
        const [identity, token] = await Promise.all([
            window.BWANotificationStore.getOrCreateIdentity(),
            messaging.getToken({ vapidKey: VAPID_KEY, serviceWorkerRegistration: registration })
        ]);
        if (!token) return;
        const response = await fetch('https://asia-northeast3-btcwebapp-551bd.cloudfunctions.net/registerNotificationDevice', {
            method: 'POST',
            cache: 'no-store',
            headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
            body: JSON.stringify({
                deviceId: identity.deviceId,
                deviceSecret: identity.deviceSecret,
                token,
                userAgent: navigator.userAgent,
                keywords: String(getFromStorage('userKeywords', '') || ''),
                active: true
            })
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok || payload.ok !== true) return;
        if (payload.result?.migratedLegacy) {
            saveToStorage('userKeywords', String(payload.result.keywords || ''));
            saveToStorage('isNotificationActive', payload.result.active === true);
        }
    } catch (_) {
        // Existing legacy delivery remains active and migration retries on the next visit.
    }
}

function initializeNotificationHeaderState() {
    const historyButton = document.getElementById('notificationHistoryBtn');
    if (!historyButton) return;
    const active = getFromStorage('isNotificationActive') === true &&
        typeof Notification !== 'undefined' && Notification.permission === 'granted';
    historyButton.classList.toggle('notification-off', !active);
    historyButton.title = active ? '알림 내역' : '알림 꺼짐 · 설정 및 내역';
}

// --- Firebase Notification Logic ---
async function requestNotificationPermission() {
    if (!messaging) {
        showStatus('Firebase가 초기화되지 않았습니다.', 'error');
        return null;
    }

    try {
        const permission = await Notification.requestPermission();
        if (permission === 'granted') {
            console.log('Notification permission granted.');

            let registration = await navigator.serviceWorker.getRegistration();

            if (!registration) {
                console.log('No active registration found. Registering new one...');
                try {
                    registration = await navigator.serviceWorker.register('./firebase-messaging-sw.js', { scope: './' });
                } catch (regErr) {
                    console.error('Explicit registration failed:', regErr);
                    throw new Error('서비스 워커 등록 실패');
                }
            }

            // 등록 대기
            if (!registration.active && registration.installing) {
                await new Promise(resolve => {
                    const worker = registration.installing;
                    worker.addEventListener('statechange', () => {
                        if (worker.state === 'activated') resolve();
                    });
                });
            }

            const token = await messaging.getToken({
                vapidKey: VAPID_KEY,
                serviceWorkerRegistration: registration
            });

            if (token) {
                console.log('FCM Token:', token);
                return token;
            } else {
                console.log('No registration token available.');
                showStatus('토큰을 가져올 수 없습니다.', 'error');
                return null;
            }
        } else {
            console.log('Unable to get permission to notify.');
            showStatus('알림 권한이 거부되었습니다.', 'error');
            return null;
        }
    } catch (err) {
        console.log('An error occurred while retrieving token. ', err);
        showStatus(`알림 설정 실패: ${err.message}`, 'error');
        return null;
    }
}

async function sendTokenToServer(token, keywords = "", isActive = true) {
    // showStatus('서버에 설정을 저장 중입니다...', 'loading'); // 사용자 요청으로 제거
    try {
        const response = await callApi('registerToken', 'POST', {
            token: token,
            userAgent: navigator.userAgent,
            keywords: keywords,
            isActive: isActive
        });
        if (response.success) {
            // showStatus('알림 설정이 저장되었습니다!', 'success', 3000); // 사용자 요청으로 제거
            // 성공 시 로컬 스토리지도 확실히 갱신
            saveToStorage('isNotificationActive', isActive);
            saveToStorage('userKeywords', keywords);
            console.log("서버 저장 성공:", response);
        } else {
            showStatus(`서버 저장 실패: ${response.message}`, 'error');
            console.error("서버 저장 실패:", response);
        }
    } catch (e) {
        console.error(e);
        showStatus('서버 통신 오류', 'error');
    }
}

async function syncNotificationSettingsWithServer() {
    if (!messaging) return;

    // 초기 로딩 시 토글 잠금 (사이드 이펙트 방지)
    // 주의: window.onload 안에 있는 지역 변수 sideToggle 등에 접근 불가하므로
    // DOM에서 직접 가져와야 함.
    const sideToggle = document.getElementById('notificationToggle');
    const mainToggle = document.getElementById('notificationToggleMain');

    if (sideToggle) sideToggle.disabled = true;
    if (mainToggle) mainToggle.disabled = true;

    try {
        // 권한이 없으면 동기화할 토큰도 없음.
        if (Notification.permission !== 'granted') {
            console.log("No notification permission, skipping sync.");
            if (sideToggle) sideToggle.disabled = false;
            if (mainToggle) mainToggle.disabled = false;
            return;
        }

        const token = await messaging.getToken({ vapidKey: VAPID_KEY });
        if (!token) {
            if (sideToggle) sideToggle.disabled = false;
            if (mainToggle) mainToggle.disabled = false;
            return;
        }

        console.log("Fetching settings from server...");
        const response = await callApi('getUserSettings', 'GET', { token: token });
        if (response.success) {
            console.log("Server settings synced:", response);
            // 로컬 스토리지 및 UI 갱신
            saveToStorage('userKeywords', response.keywords || "");
            saveToStorage('isNotificationActive', response.isActive);

            if (sideToggle) {
                sideToggle.checked = response.isActive;
                sideToggle.style.opacity = '1';
            }
            if (mainToggle) {
                mainToggle.checked = response.isActive;
                mainToggle.style.opacity = '1';
            }
        } else {
            console.warn("Failed to fetch settings:", response);
        }
    } catch (e) {
        console.log("Sync failed (not usually an error if first time):", e);
    } finally {
        if (sideToggle) sideToggle.disabled = false;
        if (mainToggle) mainToggle.disabled = false;
    }
}

// --- Keyword Modal Logic ---
function openKeywordModal() {
    document.getElementById('keywordModalOverlay').classList.add('visible');
    document.getElementById('keywordModal').classList.add('visible');
    // TODO: 기존 키워드 불러오기 (서버 연동 전엔 로컬스토리지 or 빈값)
    const storedKeywords = getFromStorage('userKeywords') || '';
    document.getElementById('keywordInput').value = storedKeywords;
    // 참고: 모달이 열려있는 동안은 토글이 잠겨있음 (handleToggleChange에서 설정)
}

function closeKeywordModal() {
    document.getElementById('keywordModalOverlay').classList.remove('visible');
    document.getElementById('keywordModal').classList.remove('visible');

    // 취소 시 토글이 켜져있었다면 끄기 (저장되지 않았으므로)
    // 단, 이미 로컬 state가 active라면(원래 켜져있던 상태에서 수정하려다 취소) 유지해야 하지만,
    // 현재 로직상 ON -> Modal Open 흐름이므로, 
    // 저장을 안했으면 '취소'로 간주하고 토글을 다시 OFF로 돌리는게 맞음.
    // (만약 '수정' 기능이 있다면 로직이 달라져야 함. 현재는 Toggle ON -> Modal임)

    // 하지만 "이미 켜져있는 상태"에서 모달을 열 수 있는 경로가 마땅히 없음 (토글을 껐다 켜야 함).
    // 따라서 취소 = 토글 OFF 원복이 타당함.

    const sideToggle = document.getElementById('notificationToggle');
    const mainToggle = document.getElementById('notificationToggleMain');

    if (sideToggle) sideToggle.checked = false;
    if (mainToggle) mainToggle.checked = false;
}

async function handleKeywordSave() {
    const keywordInput = document.getElementById('keywordInput');
    const keywords = keywordInput.value.trim();

    // 1. Optimistic UI: 즉시 저장 및 UI 반영
    saveToStorage('userKeywords', keywords);
    saveToStorage('isNotificationActive', true);

    // 토글 UI 켜기 (저장 버튼 누른 시점에 켜짐)
    const sideToggle = document.getElementById('notificationToggle');
    const mainToggle = document.getElementById('notificationToggleMain');
    if (sideToggle) sideToggle.checked = true;
    if (mainToggle) mainToggle.checked = true;

    // 모달 닫기
    document.getElementById('keywordModalOverlay').classList.remove('visible');
    document.getElementById('keywordModal').classList.remove('visible');

    // 메뉴 닫기
    closeMenu();

    // showStatus('알림 설정을 저장하고 있습니다...', 'loading'); // 사용자 요청으로 제거

    // 2. 백그라운드: 권한 요청 및 서버 전송
    try {
        const token = await requestNotificationPermission();

        if (token) {
            // 권한 성공: 서버 전송 (await 하되 UI는 이미 완료됨)
            // sendTokenToServer 내부에서 showStatus('success')를 호출하여 완료를 알림
            await sendTokenToServer(token, keywords, true);
        } else {
            // 권한 실패/거부: 롤백(Rollback) 수행
            throw new Error("Token retrieval failed");
        }
    } catch (e) {
        console.error("알림 설정 실패 (롤백):", e);

        // 롤백: 로컬 상태 및 UI 원복
        saveToStorage('isNotificationActive', false);

        const sideToggle = document.getElementById('notificationToggle');
        const mainToggle = document.getElementById('notificationToggleMain');

        if (sideToggle) sideToggle.checked = false;
        if (mainToggle) mainToggle.checked = false;

        showStatus('권한을 얻지 못해 알림 설정을 취소했습니다.', 'error');
    }
}

async function disableNotification() {
    // 키워드 유지: 빈 값("") 대신 기존에 저장된 키워드를 보냄
    const storedKeywords = getFromStorage('userKeywords') || "";

    // 토큰이 있나?
    if (!messaging) return;

    // 현재 토큰 가져오기 (권한이 이미 있으므로 바로 나올 것임)
    try {
        const token = await messaging.getToken({ vapidKey: VAPID_KEY });
        if (token) {
            // isActive만 false로 보냄
            await sendTokenToServer(token, storedKeywords, false);
        }
    } catch (e) {
        console.error("Disable error", e);
    }
}

// 포그라운드 메시지 수신 (페이지가 열려있을 때)
if (messaging) {
    messaging.onMessage((payload) => {
        console.log('Message received. ', payload);
        // Data-only 메시지 처리
        const data = payload.data;
        const title = data.title;
        const options = {
            body: data.body,
            icon: data.icon
        };
        // 브라우저 기본 알림 띄우기 (페이지가 포커스 되어 있어도 알림을 띄우고 싶다면)
        // 또는 커스텀 토스트 메시지 사용 가능
        showStatus(`[알림] ${title}: ${options.body}`, 'success', 5000);
        // 필요 시 new Notification(title, options) 호출 가능 (사용자 제스처 필요할 수 있음)
    });
}

function showStatus(message, type, duration = 0) {
    const statusDiv = document.getElementById('status');
    if (!message) {
        statusDiv.style.opacity = '0';
        setTimeout(() => { statusDiv.style.display = 'none'; }, 300);
        return;
    }
    if (statusDiv.hideTimer) clearTimeout(statusDiv.hideTimer);
    statusDiv.textContent = message;
    statusDiv.className = type;
    statusDiv.style.display = 'block';
    statusDiv.offsetHeight; // force reflow
    statusDiv.style.opacity = '1';

    if (duration > 0) {
        statusDiv.hideTimer = setTimeout(() => {
            statusDiv.style.opacity = '0';
            setTimeout(() => { statusDiv.style.display = 'none'; }, 300);
        }, duration);
    }
}

function renderFormMessage(message) {
    const container = document.getElementById('dynamicFormContainer');
    if (!container) return;
    container.replaceChildren();
    const title = document.createElement('h3');
    title.textContent = '측정값 입력 폼';
    const paragraph = document.createElement('p');
    paragraph.id = 'formMessage';
    paragraph.textContent = message;
    container.append(title, paragraph);
}

function renderFormError(message, retryHandler) {
    const container = document.getElementById('dynamicFormContainer');
    if (!container) return;
    container.replaceChildren();
    const title = document.createElement('h3');
    title.textContent = '측정값 입력 폼';
    const paragraph = document.createElement('p');
    paragraph.className = 'read-error-message';
    paragraph.setAttribute('role', 'alert');
    paragraph.textContent = message;
    const retryButton = document.createElement('button');
    retryButton.type = 'button';
    retryButton.className = 'read-retry-button';
    retryButton.textContent = '다시 시도';
    retryButton.addEventListener('click', retryHandler);
    container.append(title, paragraph, retryButton);
}

function renderListFailure(error) {
    const formSelect = document.getElementById('formSelect');
    const formListStatus = document.getElementById('formListStatus');
    if (formSelect) formSelect.disabled = true;
    if (!formListStatus) return;
    formListStatus.replaceChildren();
    const message = document.createElement('div');
    message.className = 'read-error-message';
    message.setAttribute('role', 'alert');
    message.textContent = `로드 오류: ${error.message}`;
    const retryButton = document.createElement('button');
    retryButton.type = 'button';
    retryButton.className = 'read-retry-button compact';
    retryButton.textContent = '목록 다시 시도';
    retryButton.addEventListener('click', loadFormList);
    formListStatus.append(message, retryButton);
}

// --- iOS PWA Install Guide Logic ---
// 변경: 로드시 체크가 아니라, 토글 동작 시 호출되어 가이드 표시 여부를 결정
function checkIosPwaStatusAndShowGuide() {
    const isIos = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
    const isStandalone = window.navigator.standalone === true || window.matchMedia('(display-mode: standalone)').matches;

    // 아이폰이면서 브라우저(사파리 등)인 경우
    if (isIos && !isStandalone) {
        const guide = document.getElementById('iosInstallGuide');
        if (guide) {
            guide.classList.add('visible');
        }
        return true; // 가이드를 띄웠음 (차단 필요)
    }
    return false; // 통과
}


function closeIosSettingsGuide() {
    const guide = document.getElementById('iosInstallGuide');
    if (guide) {
        guide.classList.remove('visible');
    }
}

// --- 파일 업로드 로직 ---
function handleUpload() {
    const fileInput = document.getElementById('excelFile');
    const file = fileInput.files[0];

    if (!file) {
        showStatus('파일을 선택해주세요.', 'error', 3000);
        return;
    }

    closeMenu();
    showStatus('파일을 읽는 중입니다...', 'loading');

    const reader = new FileReader();
    reader.onload = function (e) {
        const base64Data = e.target.result.split(',')[1];
        const fileData = {
            name: file.name,
            mimeType: file.type || 'application/octet-stream',
            data: base64Data
        };
        processFileUpload(fileData, undefined);
    };
    reader.onerror = function (error) {
        showStatus(`파일 읽기 오류: ${error.message}`, 'error');
    };
    reader.readAsDataURL(file);
}

// --- 서버 통신: 파일 업로드 ---
async function processFileUpload(fileData, userChoice) {
    const statusMessage = userChoice ? '사용자 선택을 반영하여 처리 중...' : '파일을 업로드 및 처리 중입니다...';
    showStatus(statusMessage, 'loading');

    const formContainer = document.getElementById('dynamicFormContainer');
    if (!userChoice) {
        formContainer.innerHTML = '<h3>측정값 입력 폼</h3><p id="formMessage" class="loading">폼 생성 중...</p>';
    }

    try {
        const response = await callApi('uploadFileBase64', 'POST', {
            fileData: fileData,
            userChoice: userChoice
        });

        if (response.success) {
            showStatus(response.message, 'success', 3000);
            document.getElementById('excelFile').value = '';

            if (response.preserved) {
                formContainer.innerHTML = '<h3>측정값 입력 폼</h3><p id="formMessage">업로드가 취소되었습니다. 다른 파일을 업로드하거나 기존 양식을 선택하세요.</p>';
            } else if (response.formData && Array.isArray(response.formData)) {
                // 업로드 성공 후 시트 정보 추출
                // (Backend 응답에 lastModifiedDate가 포함되어 있지 않을 수 있으므로 현재 시간 사용 가능)
                // 단 backend_gas_v2.js에서는 recordUploadedForm만 하고 데이터엔 안담아줄 수도 있음.
                // 편의상 여기서 처리.
                currentSheetInfo = {
                    spreadsheetId: response.spreadsheetId,
                    sheetName: response.sheetName,
                    displayName: response.sheetName,
                    lastModifiedDate: new Date().toISOString()
                };
                createDynamicForm(response.formData, response.sheetName);
                loadFormList();
            }

        } else {
            if (response.requiresChoice) {
                if (confirm(response.message)) {
                    processFileUpload(fileData, 'overwrite');
                } else {
                    processFileUpload(fileData, 'preserve');
                }
            } else {
                const msg = response.message || "알 수 없는 오류";
                showStatus(msg, 'error');
                renderFormError(
                    `폼 생성 실패: ${msg}`,
                    () => processFileUpload(fileData, userChoice)
                );
            }
        }
    } catch (error) {
        const msg = `서버 통신 오류: ${error.message}`;
        showStatus(msg, 'error');
        renderFormError(msg, () => processFileUpload(fileData, userChoice));
    }
}

// --- 동적 폼 생성 ---
function createDynamicForm(formData, formTitle) {
    const formContainer = document.getElementById('dynamicFormContainer');
    const sheetName = currentSheetInfo.sheetName;

    // [1] 저장된 순서가 있으면 데이터 정렬
    formData = sortFormData(formData, sheetName);

    // 날짜 포맷
    let lastDateStr = '';
    let fileDateStr = '';
    if (currentSheetInfo?.lastModifiedDate) {
        try {
            const d = new Date(currentSheetInfo.lastModifiedDate);
            lastDateStr = `(${d.getFullYear().toString().slice(2, 4)}.${('0' + (d.getMonth() + 1)).slice(-2)}.${('0' + d.getDate()).slice(-2)})`;
            fileDateStr = `${d.getFullYear().toString().slice(2, 4)}${('0' + (d.getMonth() + 1)).slice(-2)}${('0' + d.getDate()).slice(-2)}`;
        } catch (e) { }
    }

    // 엑셀 다운로드 버튼
    let downloadBtnHtml = '';
    if (formTitle) {
        downloadBtnHtml = `<button id="xlsxDownloadBtn"
        onclick="triggerPreparedDownload('xlsxDownloadBtn')"
        disabled
        style="width: auto; margin-left: 5px; margin-right: 5px; font-size:0.9em; padding: 6px 10px; background-color: #ccc; color: #666;
              border: none; border-radius: 4px; font-weight: bold; cursor: not-allowed;">
        저장 후 XLSX
      </button>`
    }

    // [1] 간격 조절 콤보박스 (평소에는 숨김 display: none)
    // 4px(좁게), 8px(보통), 12px(넓게) 옵션 제공
    const spacingSelectHtml = `
        <select id="spacingSelect" onchange="handleSpacingChange(this.value)" 
            style="display: none; width: auto; padding: 6px 10px; margin: 10px 0 10px 0 ; font-size: 0.9em; align-items: center; justify-content: center; vertical-align: bottom; border: 1px solid rgb(204, 204, 204); border-radius: 4px; background: rgb(255, 255, 255);">
            <option value="4">좁게</option>
            <option value="8">보통</option>
            <option value="12" selected>넓게</option>
        </select>
    `;

    // [2]정렬 초기화 버튼 (평소에는 숨겨져 있음 display: none)
    const resetBtnHtml = `
        <button id="resetOrderBtn" type="button" onclick="handleResetOrder()" title="초기 순서로 복원"
            style="display: none; width: auto; padding: 6px 10px; margin-left: 5px; margin-right: 5px; font-size:0.9em; align-items: center; justify-content: center;
                   background:transparent; color: #666; border: 1px solid #ffcdd2; border-radius: 4px;">
            <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                <polyline points="1 4 1 10 7 10"></polyline>
                <path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"></path>
            </svg>
            기본값
        </button>`;

    // 초기화
    setHomeOnlyElementsVisible(false);

    const oldForm = document.getElementById('measurementForm');
    if (oldForm) oldForm.remove();

    // [2] 헤더 구성: 제목 + 순서변경 버튼 + 다운로드 버튼
    let h3 = formContainer.querySelector('h3');
    if (!h3) {
        h3 = document.createElement('h3');
        formContainer.prepend(h3);
    }

    // 순서 변경 버튼 HTML
    const sortBtnHtml = `
        <button id="toggleSortBtn" type="button" onclick="toggleSortMode()" 
            style="margin-left:auto; width:auto; padding: 6px 10px; background:transparent; color:#333; border:1px solid #ccc; font-size:0.9em;">
            ⇅ 정렬
        </button>`;

    h3.style.display = 'flex';
    h3.style.alignItems = 'center';
    h3.style.justifyContent = 'space-between';
    h3.style.flexWrap = 'nowrap';
    h3.style.marginBottom = '16px';
    h3.style.gap = '5px';

    // 제목 영역
    h3.innerHTML = `
        <span style="flex: 1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-size: 1.1em;">
            ${formTitle || '측정값 입력 폼'}
        </span>
        <div style="display: flex; flex-shrink: 0;">
            ${spacingSelectHtml} ${resetBtnHtml} ${downloadBtnHtml} ${sortBtnHtml}
        </div>
    `;

    const formElement = document.createElement('form');
    formElement.id = 'measurementForm';

    if (!formData || !Array.isArray(formData) || formData.length === 0) {
        formContainer.innerHTML += '<p class="error">데이터가 없습니다.</p>';
        return;
    }

    const uniqueIds = formData.map(d => d.uniqueId).filter(id => id);
    if (uniqueIds.length > 0) loadValidationData(uniqueIds);

    // [3] 폼 필드 생성 Loop
    let prevLocPrefix = null;

    formData.forEach((data, index) => {
        const formGroup = document.createElement('div');
        formGroup.className = 'form-group';
        formGroup.dataset.uniqueId = data.uniqueId; // 정렬 저장을 위해 ID 심어둠

        // 위치(Location) 변경 시 구분선 처리 (별도 div 대신 클래스 추가)
        const currLocPrefix = (data.location || '').substring(0, 3);
        if (index > 0 && prevLocPrefix !== null && prevLocPrefix !== currLocPrefix) {
            formGroup.classList.add('group-start'); // CSS에서 border-top 처리
        }
        prevLocPrefix = currLocPrefix;

        // 드래그 핸들 추가 (햄버거 아이콘)
        const handle = document.createElement('div');
        handle.className = 'drag-handle';
        handle.innerHTML = '☰'; // 또는 SVG 아이콘 사용
        formGroup.appendChild(handle);

        const locationSpan = document.createElement('span');
        locationSpan.className = 'item-location';
        locationSpan.textContent = data.location;

        const itemSpan = document.createElement('span');
        itemSpan.className = 'item-detail';

        let itemText = '';
        let placeholderText = '측정값';
        const words = (data.item || '').trim().split(/\s+/).filter(w => w);
        if (words.length > 1) {
            placeholderText = words.pop();
            itemText = words.join(' ');
        } else if (words.length === 1) {
            itemText = words[0];
            placeholderText = words[0];
        }
        itemSpan.textContent = itemText;

        const input = document.createElement('input');
        input.type = 'number';
        input.inputMode = 'decimal';
        input.step = 'any';
        input.placeholder = placeholderText;
        input.value = '';
        input.dataset.location = data.location;
        input.dataset.item = data.item;
        input.dataset.unit = data.unit;
        input.dataset.uniqueId = data.uniqueId;

        input.addEventListener('blur', function () { validateInputValue(this); });

        const unitSpan = document.createElement('span');
        unitSpan.className = 'measurement-unit';
        unitSpan.textContent = data.unit || '';

        formGroup.appendChild(locationSpan);
        formGroup.appendChild(itemSpan);
        formGroup.appendChild(input);
        formGroup.appendChild(unitSpan);

        formElement.appendChild(formGroup);
    });

    // 폼 생성 직후, 저장된 간격 설정이 있다면 적용
    const savedSpacing = localStorage.getItem('userFormSpacing');
    if (savedSpacing) {
        handleSpacingChange(savedSpacing);
        // 콤보박스 값도 동기화
        const select = document.getElementById('spacingSelect');
        if (select) select.value = savedSpacing;
    }

    const submitButton = document.createElement('button');
    submitButton.type = 'button';
    submitButton.textContent = '측정값 저장';
    submitButton.id = 'saveMeasurements';
    submitButton.onclick = saveMeasurements;
    submitButton.disabled = !currentSheetInfo?.formKey;
    if (submitButton.disabled) submitButton.textContent = '목록 반영 후 저장';

    const submissionStatus = document.createElement('div');
    submissionStatus.id = 'submissionStatus';
    submissionStatus.className = 'submission-status status-idle';
    submissionStatus.setAttribute('role', 'status');
    submissionStatus.textContent = currentSheetInfo?.formKey
        ? '저장 대기'
        : '업로드 양식을 목록에서 다시 선택한 뒤 저장할 수 있습니다.';
    formElement.appendChild(submissionStatus);

    // 저장 버튼은 드래그 영역 밖이어야 안전하므로 formElement 밖이나 마지막에 배치
    formElement.appendChild(submitButton);
    formContainer.appendChild(formElement);

    formElement.addEventListener('input', () => {
        isMeasurementDirty = true;
        pendingSubmission = null;
        preparedDownload = null;
        const downloadButton = document.getElementById('xlsxDownloadBtn');
        if (downloadButton) {
            downloadButton.disabled = true;
            downloadButton.style.backgroundColor = '#ccc';
            downloadButton.style.color = '#666';
            downloadButton.style.cursor = 'not-allowed';
            downloadButton.textContent = '저장 후 XLSX';
        }
        updateSubmissionStatus('입력값이 변경되었습니다.', 'idle');
    });
    addHomeStateToHistory();

    // [4] Sortable 초기화 (비활성화 상태로 시작)
    initSortable();
}

// --- XLSX 다운로드 준비 ---
async function prepareXlsxInAdvance(fileId, sheetName, fileName, fileDateStr) {
    // GAS API는 fileId, sheetName, filename을 파라미터로 받아서 Base64를 리턴하도록 되어있음
    try {
        let url = `${GAS_API_URL}?fileId=${encodeURIComponent('ignored')}&sheetName=${encodeURIComponent(sheetName)}&filename=${encodeURIComponent(fileName)}`;
        // 백엔드가 fileId를 필수라고 생각한다면 더미값 전달. backend_gas_v2.js에서는 openById(fileId)를 하므로
        // IMPORTANT: backend_gas_v2.js의 handleXlsxDownload는 fileId를 받는다.
        // 하지만 우리는 TARGET_SPREADSHEET_ID를 백엔드가 알고있다.
        // 만약 백엔드가 fileId를 필수로 받는다면 여기서 TARGET_SPREADSHEET_ID를 알아야한다.
        // 일단 사용자가 backend_gas_v2.js에 상수로 ID를 박았으므로, fileId파라미터가 없어도 동작하도록 백엔드를 수정하거나
        // 아니면 여기서 상수로 ID를 가지고 있어야 한다.
        // 프론트에 ID를 노출하고 싶지 않다면 백엔드 수정 필요.
        // 지금은 backend_gas_v2.js가 fileId를 받아서 openById 한다고 가정되어 있음.
        // 따라서 기존 로직 호환을 위해 더미 ID 또는 실제 ID가 필요함.
        // 편의상 아래 상수를 정의해서 사용.
    } catch (err) { }

    // Note: Since we removed the ID injection, download feature might break if backend strictly requires ID param.
    // For now, let's assume backend defaults to global ID if param is missing, OR we fetch it first.
    // We will pass sheetName.

    const options = { method: 'GET' };
    // URL Construct again
    // We need to pass TARGET_SPREADSHEET_ID... but we removed it from Index.html.
    // Let's assume we pass 'default' and backend handles it, OR fetch 'getFormList' returned spreadsheetId.
    const targetId = currentSheetInfo?.spreadsheetId || '19rgzRnTQtOwwW7Ts5NbBuItNey94dAZsEnO7Tk0cm6s'; // Fallback to hardcoded ID if needed

    let fetchUrl = `${GAS_API_URL}?fileId=${encodeURIComponent(targetId)}&sheetName=${encodeURIComponent(sheetName)}&filename=${encodeURIComponent(fileName)}`;

    try {
        const res = await fetch(fetchUrl);
        const json = await res.json();
        if (!res.ok || json.error || !json.base64 || !json.filename) {
            throw new Error(json.error || `XLSX_HTTP_${res.status}`);
        }

        preparedDownload = json;
        const btn = document.getElementById('xlsxDownloadBtn');
        if (btn) {
            btn.disabled = false;
            btn.style.backgroundColor = '#4CAF50';
            btn.style.color = 'white';
            btn.style.cursor = 'pointer';
            btn.innerText = `⬇ ${fileDateStr} 엑셀`;
        }
        return true;
    } catch (err) {
        console.error(err);
        preparedDownload = null;
        const btn = document.getElementById('xlsxDownloadBtn');
        if (btn) {
            btn.disabled = false;
            btn.innerText = 'XLSX 준비 다시 시도';
        }
        return false;
    }
}

function xlsxFilename() {
    const now = new Date();
    const date = `${String(now.getFullYear()).slice(-2)}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    return {
        date,
        filename: `${currentSheetInfo.displayName || currentSheetInfo.sheetName}_${date}.xlsx`
    };
}

async function prepareXlsxAfterSync() {
    if (!currentSheetInfo) return false;
    const button = document.getElementById('xlsxDownloadBtn');
    if (button) {
        button.disabled = true;
        button.textContent = 'XLSX 준비 중…';
        delete button.dataset.prepareRetry;
    }
    const { date, filename } = xlsxFilename();
    const prepared = await prepareXlsxInAdvance(
        currentSheetInfo.spreadsheetId,
        currentSheetInfo.sheetName,
        filename,
        date
    );
    if (prepared) {
        updateSubmissionStatus('Google Sheet 동기화 및 XLSX 준비 완료', 'synced');
        return true;
    }
    if (button) button.dataset.prepareRetry = 'true';
    updateSubmissionStatus('Google Sheet 동기화 완료 · XLSX 준비 실패', 'xlsx-error');
    return false;
}

async function triggerPreparedDownload(buttonId) {
    const button = document.getElementById(buttonId);
    if (!preparedDownload && button?.dataset.prepareRetry === 'true') {
        await prepareXlsxAfterSync();
    }
    if (!preparedDownload) {
        alert('파일이 아직 준비되지 않았습니다.');
        return;
    }
    const a = document.createElement('a');
    a.href = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${preparedDownload.base64}`;
    a.download = preparedDownload.filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
}

// --- 측정값 저장 ---
function updateSubmissionStatus(message, status = 'idle') {
    const element = document.getElementById('submissionStatus');
    if (!element) return;
    element.className = `submission-status status-${status}`;
    element.textContent = message;
}

function createIdempotencyKey() {
    if (!window.crypto || typeof window.crypto.randomUUID !== 'function') {
        throw new Error('SECURE_RANDOM_UUID_UNAVAILABLE');
    }
    return window.crypto.randomUUID();
}

function collectMeasurements() {
    const inputs = Array.from(document.querySelectorAll('#measurementForm input[type="number"]'));
    if (!currentSheetInfo || inputs.length !== currentSheetInfo.rowCount) {
        throw new Error('FORM_STATE_MISMATCH');
    }
    return inputs.map((input) => ({
        uniqueId: input.dataset.uniqueId || '',
        location: input.dataset.location || '',
        item: input.dataset.item || '',
        value: input.value.trim(),
        unit: input.dataset.unit || ''
    }));
}

function submissionPayload() {
    if (!currentSheetInfo?.formKey || !currentSheetInfo?.sourceRevision) {
        throw new Error('FIRESTORE_FORM_STATE_REQUIRED');
    }
    const payload = {
        schemaVersion: 1,
        idempotencyKey: pendingSubmission?.idempotencyKey || createIdempotencyKey(),
        formKey: currentSheetInfo.formKey,
        sheetName: currentSheetInfo.sheetName,
        formRevision: currentSheetInfo.sourceRevision,
        measurements: collectMeasurements()
    };
    const semantic = window.BWAReadAdapter.stableStringify({ ...payload, idempotencyKey: null });
    if (pendingSubmission && pendingSubmission.semantic !== semantic) {
        payload.idempotencyKey = createIdempotencyKey();
    }
    pendingSubmission = { idempotencyKey: payload.idempotencyKey, semantic };
    return payload;
}

async function fetchFunctionJson(url, body) {
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), FUNCTION_REQUEST_TIMEOUT_MS);
    try {
        const response = await fetch(url, {
            method: 'POST',
            cache: 'no-store',
            headers: { Accept: 'application/json', 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
            signal: controller.signal
        });
        const payload = await response.json().catch(() => ({}));
        if (!response.ok || payload.ok !== true) {
            throw new Error(payload.error || `HTTP_${response.status}`);
        }
        return payload.result;
    } finally {
        window.clearTimeout(timeout);
    }
}

function sleep(ms) {
    return new Promise((resolve) => window.setTimeout(resolve, ms));
}

async function getSubmissionStatus(idempotencyKey) {
    return fetchFunctionJson(GET_SUBMISSION_URL, { idempotencyKey });
}

async function pollSubmission(idempotencyKey) {
    const deadline = Date.now() + SUBMISSION_POLL_TIMEOUT_MS;
    let lastStatus = null;
    while (Date.now() < deadline) {
        const status = await getSubmissionStatus(idempotencyKey);
        lastStatus = status;
        if (status.status === 'synced') return status;
        if (status.status === 'failed' && !status.retryable) {
            throw new Error(status.errorCode || 'SUBMISSION_SYNC_FAILED');
        }
        updateSubmissionStatus(
            status.status === 'failed'
                ? `일시 오류 후 자동 재시도 대기 · ${status.errorCode || '원인 확인 중'}`
                : `접수 ${idempotencyKey.slice(0, 8)}… · ${status.status} · 시트 반영 확인 중`,
            status.status
        );
        await sleep(SUBMISSION_POLL_INTERVAL_MS);
    }
    throw new Error(`SUBMISSION_STATUS_TIMEOUT:${lastStatus?.status || 'unknown'}`);
}

async function submitOrRecover(payload) {
    try {
        return await fetchFunctionJson(SUBMIT_MEASUREMENTS_URL, payload);
    } catch (submitError) {
        updateSubmissionStatus('접수 응답이 불확실하여 같은 요청의 상태를 확인합니다.', 'recovering');
        try {
            const status = await getSubmissionStatus(payload.idempotencyKey);
            return { created: false, status };
        } catch (statusError) {
            if (statusError.message !== 'SUBMISSION_NOT_FOUND') throw submitError;
            return fetchFunctionJson(SUBMIT_MEASUREMENTS_URL, payload);
        }
    }
}

function setSaveButtonBusy(isBusy, label) {
    const button = document.getElementById('saveMeasurements');
    if (!button) return;
    button.disabled = isBusy;
    if (label) button.textContent = label;
}

async function saveMeasurements() {
    let payload;
    try {
        payload = submissionPayload();
        setSaveButtonBusy(true, '저장 중...');
        updateSubmissionStatus('Firestore에 측정값을 접수하는 중입니다.', 'submitting');
        showStatus('측정값을 저장 중입니다...', 'loading');
        const receipt = await submitOrRecover(payload);
        const initialStatus = receipt.status;
        const synced = initialStatus.status === 'synced'
            ? initialStatus
            : await pollSubmission(payload.idempotencyKey);
        isMeasurementDirty = false;
        currentSheetInfo.sourceRevision = synced.sourceRevisionAfterSync || currentSheetInfo.sourceRevision;
        currentSheetInfo.lastModifiedDate = currentSheetInfo.sourceRevision;
        pendingSubmission = null;
        updateSubmissionStatus(
            `Google Sheet 동기화 완료 · ${synced.updatedCellCount}개 셀 · XLSX 준비 중`,
            'synced'
        );
        setSaveButtonBusy(true, '저장 완료');
        await loadFormList();
        await prepareXlsxAfterSync();
        showStatus('측정값이 Google Sheet에 저장되었습니다.', 'success', 3000);
    } catch (error) {
        updateSubmissionStatus(`저장 실패 · GAS 재전송 없음 · ${error.message}`, 'failed');
        setSaveButtonBusy(false, '다시 저장');
        showStatus(`저장 오류: ${error.message}`, 'error');
    }
}

// --- 양식 목록 로드 ---
async function loadFormList() {
    const formSelect = document.getElementById('formSelect');
    const formListStatus = document.getElementById('formListStatus');
    const originalValue = formSelect.value;

    formSelect.disabled = true;
    formListStatus.textContent = '양식 목록 불러오는 중...';
    formListStatus.className = 'loading';

    try {
        const { items: formList } = await window.BWAProductionRead.loadFormList(
            () => callApi('getFormList', 'GET')
        );
        formSelect.disabled = false;
        formSelect.innerHTML = '<option value="">-- 양식을 선택해주세요 --</option>';

        if (formList && formList.length > 0) {
            formList.sort((a, b) => new Date(b.lastModifiedDate) - new Date(a.lastModifiedDate));
            formList.forEach(form => {
                const option = document.createElement('option');
                option.value = form.sheetName;
                const cleanName = (form.displayName || form.sheetName).split('_')[0];
                option.textContent = `${cleanName} (수정: ${formatDateForDisplay(form.lastModifiedDate)})`;
                option.dataset.displayName = cleanName;
                option.dataset.lastModifiedDate = form.lastModifiedDate;
                option.dataset.spreadsheetId = form.spreadsheetId;
                if (form.formKey) option.dataset.formKey = form.formKey;
                formSelect.appendChild(option);
            });
            formSelect.value = originalValue;
            formListStatus.textContent = '양식 목록 로드 완료.';
            formListStatus.className = 'success';
        } else {
            formListStatus.textContent = '저장된 양식이 없습니다.';
            formListStatus.className = '';
        }
        updateFavoriteButtons();

    } catch (error) {
        renderListFailure(error);
        updateFavoriteButtons();
    }
}

async function loadSelectedForm() {
    if (isMeasurementDirty && !confirm('변경사항이 저장되지 않았습니다. 이동하시겠습니까?')) {
        document.getElementById('formSelect').value = currentSheetInfo ? currentSheetInfo.sheetName : '';
        return;
    }

    const formSelect = document.getElementById('formSelect');
    const selectedOption = formSelect.options[formSelect.selectedIndex];
    const sheetName = selectedOption.value;
    const formContainer = document.getElementById('dynamicFormContainer');

    // 양식 선택이 없을 경우 (또는 홈 버튼 클릭 시) 초기 화면으로 복귀
    if (!sheetName) {
        selectedFormRequestId += 1;
        renderFormMessage('메뉴(☰)를 열어 새 양식을 업로드하거나 기존 양식을 선택해주세요.');
        setHomeOnlyElementsVisible(true);

        currentSheetInfo = null;
        updateHomeButtonVisibility();

        // [Issue 2 Fix] 홈 복귀 시 dirty flag 초기화
        isMeasurementDirty = false;

        // [Issue 4 Fix] 홈 상태 히스토리 추가 (이동 확정 시점)
        addHomeStateToHistory();

        closeMenu();
        return;
    }

    const requestId = ++selectedFormRequestId;
    currentSheetInfo = null;
    pendingSubmission = null;
    preparedDownload = null;
    setHomeOnlyElementsVisible(false);
    updateHomeButtonVisibility();
    showStatus(`${sheetName} 로드 중...`, 'loading');
    isMeasurementDirty = false;
    closeMenu();

    try {
        const { rows: formData } = await window.BWAProductionRead.loadForm(
            sheetName,
            {
                formKey: selectedOption.dataset.formKey || null,
                sheetName,
                displayName: selectedOption.dataset.displayName,
                lastModifiedDate: selectedOption.dataset.lastModifiedDate,
                spreadsheetId: selectedOption.dataset.spreadsheetId
            },
            () => callApi('getFormDataForWeb', 'GET', { sheetName })
        );
        if (requestId !== selectedFormRequestId) return;

        currentSheetInfo = {
            spreadsheetId: selectedOption.dataset.spreadsheetId,
            formKey: selectedOption.dataset.formKey || null,
            sheetName: sheetName,
            displayName: selectedOption.dataset.displayName,
            lastModifiedDate: selectedOption.dataset.lastModifiedDate,
            sourceRevision: selectedOption.dataset.lastModifiedDate,
            rowCount: formData.length
        };

        if (formData && formData.length > 0) {
            createDynamicForm(formData, currentSheetInfo.displayName);
            showStatus('로드 완료', 'success', 3000);
            updateHomeButtonVisibility();
        } else {
            renderFormError('데이터가 없습니다.', loadSelectedForm);
            setHomeOnlyElementsVisible(true);
        }

    } catch (error) {
        if (requestId !== selectedFormRequestId) return;
        const message = `폼 로딩 오류: ${error.message}`;
        renderFormError(message, loadSelectedForm);
        showStatus(message, 'error');
        // 에러 발생 시에도 안전하게 초기화
        currentSheetInfo = null;
        setHomeOnlyElementsVisible(true);
        updateHomeButtonVisibility();
    }
}

// --- 즐겨찾기 로직 ---
function initializeFavorites() {
    favorites = getFromStorage('favorites') || {};
    document.getElementById('favoritesSection').addEventListener('click', handleFavoriteClick);

    // [복원] 홈버튼 이벤트 리스너
    const homeBtn = document.getElementById('homeBtn');
    if (homeBtn) {
        homeBtn.addEventListener('click', function () {
            if (isMapViewActive) {
                closeMapView();
                return;
            }
            document.getElementById('formSelect').value = '';
            // [Issue 4 Fix] 여기서 상태를 초기화하지 않고 loadSelectedForm에 위임
            // currentSheetInfo = null; 
            loadSelectedForm();
            // updateHomeButtonVisibility();
            // addHomeStateToHistory();
        });
    }
}

function updateFavoriteButtons() {
    for (let i = 1; i <= 3; i++) {
        const btn = document.getElementById(`favBtn${i}`);
        if (favorites[i]) {
            btn.textContent = favorites[i].displayName;
            btn.classList.add('registered');
            btn.disabled = false;
        } else {
            btn.textContent = '비어있음';
            btn.classList.remove('registered');
            btn.disabled = false;
        }
    }
}

function handleFavoriteClick(e) {
    if (!e.target.matches('.fav-button')) return;
    const favId = e.target.dataset.favId;

    if (favorites[favId]) {
        // 로드
        const fav = favorites[favId];
        const formSelect = document.getElementById('formSelect');
        // Select option logic simliar to original
        let opt = [...formSelect.options].find(o => o.value === fav.sheetName);
        if (!opt) opt = [...formSelect.options].find(o => (o.dataset.displayName) === fav.displayName);

        if (opt) {
            formSelect.value = opt.value;
            loadSelectedForm();
        } else {
            showStatus('즐겨찾기 된 양식을 찾을 수 없어 초기화합니다.', 'error');
            delete favorites[favId];
            saveToStorage('favorites', favorites);
            updateFavoriteButtons();
        }
    } else {
        // 등록
        if (currentSheetInfo && currentSheetInfo.sheetName) {
            if (confirm(`현재 양식 '${currentSheetInfo.displayName}'를 이 즐겨찾기에 등록하시겠습니까?`)) {
                favorites[favId] = { sheetName: currentSheetInfo.sheetName, displayName: currentSheetInfo.displayName };
                saveToStorage('favorites', favorites);
                updateFavoriteButtons();
                showStatus(`'${currentSheetInfo.displayName}'가 즐겨찾기에 등록되었습니다.`, 'success', 3000);
            }
        } else {
            // [복원] 양식이 로드되지 않았을 경우, 선택 팝업을 띄움
            promptForFavoriteSelection(favId);
        }
    }
}

function promptForFavoriteSelection(favId) {
    const formSelect = document.getElementById('formSelect');
    if (formSelect.options.length <= 1) {
        showStatus('등록할 양식이 없습니다. 먼저 새 양식을 업로드해주세요.', 'error', 3000);
        return;
    }

    const overlay = document.createElement('div');
    overlay.className = 'fav-modal-overlay';

    const modal = document.createElement('div');
    modal.className = 'fav-modal-content';

    modal.innerHTML = `
            <h4>즐겨찾기 등록</h4>
            <p>이 슬롯에 등록할 양식을 선택해주세요.</p>
            <select id="favModalSelect"></select>            
            <div class="fav-modal-buttons">
                <button id="favModalCancel">취소</button>
                <button id="favModalRegister" class="primary">등록</button>
            </div>
        `;

    document.body.appendChild(overlay);
    document.body.appendChild(modal);

    const modalSelect = document.getElementById('favModalSelect');
    for (let i = 1; i < formSelect.options.length; i++) {
        const clonedOption = formSelect.options[i].cloneNode(true);
        const cleanName = clonedOption.dataset.displayName || clonedOption.textContent.split(' (')[0];
        clonedOption.textContent = cleanName; // 팝업에서는 수정 날짜 없이 깔끔한 이름만 보여줌
        modalSelect.appendChild(clonedOption);
    }

    function closeModal() {
        document.body.removeChild(overlay);
        document.body.removeChild(modal);
    }

    document.getElementById('favModalRegister').onclick = function () {
        const selectedOption = modalSelect.options[modalSelect.selectedIndex];
        const sheetName = selectedOption.value;
        const displayName = selectedOption.dataset.displayName || sheetName;

        favorites[favId] = { sheetName, displayName };
        saveToStorage('favorites', favorites);
        updateFavoriteButtons();
        showStatus(`'${displayName}'가 즐겨찾기에 등록되었습니다.`, 'success', 3000);
        closeModal();
    };

    document.getElementById('favModalCancel').onclick = closeModal;
    overlay.onclick = closeModal;
}

// --- 유효성 검사 로직 (Validation) ---
async function loadValidationData(uniqueIds) {
    try {
        // uniqueIds array to JSON
        const data = await callApi('getValidationDataFromDB', 'GET', { uniqueIds: JSON.stringify(uniqueIds) });
        validationData = data;
    } catch (e) { console.error(e); }
}

function validateInputValue(input) {
    const val = parseFloat(input.value);
    const uid = input.dataset.uniqueId;
    if (!uid || isNaN(val)) return;

    const info = validationData[uid];
    if (!info) return;

    if ((info.minValue && val < info.minValue) || (info.maxValue && val > info.maxValue)) {
        showValidationWarning(input, val, info.minValue, info.maxValue, info.recentValue, info.recentDate);
    }
}

function showValidationWarning(input, value, min, max, recentVal, recentDate) {
    // [Issue 1 Fix] DOM 요소 직접 생성 및 연결 방식으로 데드락 방지
    const overlay = document.createElement('div');
    overlay.className = 'validation-modal-overlay';
    // 오버레이 클릭 시 닫기
    overlay.onclick = closeModal;

    const modal = document.createElement('div');
    modal.className = 'validation-modal-content';

    const recentValueText = recentVal ? recentVal : '없음';

    // 내용 구성 (버튼 제외)
    modal.innerHTML = `<h4>⚠️ 범위 경고</h4>
          <p>입력값이 유효범위를 벗어납니다.</p>
          <p>${recentDate || ''} 값: ${recentValueText}<br>현재 값: ${value}</p>`;

    // 버튼 컨테이너 생성
    const btnContainer = document.createElement('div');
    btnContainer.className = 'validation-modal-buttons';

    // "수정" 버튼 (값 지우고 포커스)
    const btnYes = document.createElement('button');
    btnYes.className = 'primary';
    btnYes.textContent = '수정';
    btnYes.onclick = function () {
        input.value = '';
        input.focus();
        closeModal();
    };

    // "무시하기" 버튼 (값 유지)
    const btnNo = document.createElement('button');
    btnNo.textContent = '무시하기';
    btnNo.onclick = function () {
        closeModal();
    };

    btnContainer.appendChild(btnYes);
    btnContainer.appendChild(btnNo);
    modal.appendChild(btnContainer);

    document.body.appendChild(overlay);
    document.body.appendChild(modal);

    function closeModal() {
        if (overlay.parentNode) document.body.removeChild(overlay);
        if (modal.parentNode) document.body.removeChild(modal);
    }
}

// --- 유틸리티 ---
function formatDateForDisplay(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return `${d.getMonth() + 1}/${d.getDate()}`;
}
function updateHomeButtonVisibility() {
    const homeBtn = document.getElementById('homeBtn');
    if (homeBtn && (currentSheetInfo || isMapViewActive)) homeBtn.classList.add('visible');
    else if (homeBtn) homeBtn.classList.remove('visible');
}
function addHomeStateToHistory() {
    const url = new URL(window.location.href);
    url.searchParams.set('page', 'home');
    window.history.pushState({ page: 'home' }, 'Home', url);
}

// --- 로컬 스토리지 관련 헬퍼함수 추가 ---
// 순서 저장 키 생성
function getOrderStorageKey(sheetName) {
    return `itemOrder_${sheetName}`;
}

// 데이터 정렬 함수
function sortFormData(formData, sheetName) {
    const savedOrder = getFromStorage(getOrderStorageKey(sheetName));
    if (!savedOrder || !Array.isArray(savedOrder)) return formData;

    // 저장된 ID 순서대로 정렬
    // uniqueId가 있는 경우를 가정 (없으면 item+location 조합 등을 써야 함)
    // 여기서는 간단히 uniqueId가 있다고 가정합니다.

    const orderMap = {};
    savedOrder.forEach((id, index) => { orderMap[id] = index; });

    return formData.sort((a, b) => {
        const indexA = orderMap[a.uniqueId] !== undefined ? orderMap[a.uniqueId] : 9999;
        const indexB = orderMap[b.uniqueId] !== undefined ? orderMap[b.uniqueId] : 9999;
        return indexA - indexB;
    });
}

//Sortable 제어함수 추가
function initSortable() {
    const el = document.getElementById('measurementForm');
    if (!el) return;

    // 저장 버튼은 드래그 대상에서 제외하기 위해 필터링 필요
    // draggable 옵션으로 .form-group만 드래그 가능하게 설정
    sortableInstance = Sortable.create(el, {
        animation: 150,
        handle: '.drag-handle', // 이 핸들을 잡아야만 드래그 가능
        draggable: '.form-group', // 드래그 가능한 요소
        disabled: true, // 초기엔 비활성화
        ghostClass: 'sortable-ghost',
        onEnd: function (evt) {
            // 드래그가 끝날 때마다 순서 저장
            saveCurrentOrder();

            // 구분선(group-start) 재계산 (순서가 바뀌었으니)
            recalculateDividers();
        }
    });
}

function toggleSortMode() {
    const btn = document.getElementById('toggleSortBtn');
    const downloadBtn = document.getElementById('xlsxDownloadBtn'); // 다운로드 버튼 ID
    const resetBtn = document.getElementById('resetOrderBtn');      // 초기화 버튼 ID
    const spacingSelect = document.getElementById('spacingSelect'); // 간격 콤보박스
    const form = document.getElementById('measurementForm');

    isSortMode = !isSortMode;

    if (isSortMode) {
        // 편집 모드 켜기

        // 1. 버튼 스타일 변경 (파란색 활성화)
        btn.style.background = '#e7f3ff';
        btn.style.borderColor = '#2196f3';
        btn.style.color = '#0b69d3';

        // 2. 버튼 스위칭 (다운로드 숨김, 초기화, 간격 콤보박스 보임)
        if (downloadBtn) downloadBtn.style.display = 'none';
        if (resetBtn) resetBtn.style.display = 'flex';
        if (spacingSelect) spacingSelect.style.display = 'block';

        // 3. Sortable 활성화
        form.classList.add('sort-mode');
        if (sortableInstance) sortableInstance.option('disabled', false);

        showStatus('핸들(☰)을 드래그하여 순서를 변경하세요.', 'success', 2000);
    } else {
        // 편집 모드 끄기

        // 1. 버튼 스타일 복구
        btn.style.background = '#fff';
        btn.style.borderColor = '#ccc';
        btn.style.color = '#333';

        // 2. 버튼 스위칭 (다운로드 보임, 초기화 숨김, 간격 콤보박스 숨김)
        if (downloadBtn) downloadBtn.style.display = 'flex';
        if (resetBtn) resetBtn.style.display = 'none';
        if (spacingSelect) spacingSelect.style.display = 'none';

        // 3. Sortable 비활성화
        form.classList.remove('sort-mode');
        if (sortableInstance) sortableInstance.option('disabled', true);

        showStatus('순서가 저장되었습니다.', 'success', 2000);
    }
}

// 정렬 초기화 핸들러
function handleResetOrder() {
    if (!currentSheetInfo) return;

    if (confirm('정렬을 기본 순서로 되돌리시겠습니까?')) {
        // 1. 로컬 스토리지에서 순서 데이터 삭제
        const key = getOrderStorageKey(currentSheetInfo.sheetName);
        localStorage.removeItem(key);

        // 2. 알림 표시
        showStatus('정렬이 초기화되었습니다.', 'success', 1000);

        // 3. 폼 새로고침 (기본 순서로 다시 렌더링)
        // 주의: toggleSortMode 상태는 초기화되므로 다시 폼이 로드되면 일반 모드가 됩니다.
        loadSelectedForm();
    }
}

// 간격 변경 핸들러
function handleSpacingChange(pxValue) {
    // 1. CSS 변수 값을 변경하여 즉시 반영
    document.documentElement.style.setProperty('--input-margin', pxValue + 'px');

    // 2. 사용자가 선택한 값을 로컬스토리지에 저장 (다음에 왔을 때도 유지)
    localStorage.setItem('userFormSpacing', pxValue);

    // 3. 콤보박스 상태 동기화 (함수가 코드로 호출될 경우를 대비)
    const select = document.getElementById('spacingSelect');
    if (select && select.value !== pxValue) {
        select.value = pxValue;
    }
}

function saveCurrentOrder() {
    if (!currentSheetInfo) return;

    const formGroups = document.querySelectorAll('#measurementForm .form-group');
    const newOrderIds = Array.from(formGroups).map(el => el.dataset.uniqueId);

    saveToStorage(getOrderStorageKey(currentSheetInfo.sheetName), newOrderIds);
}

function recalculateDividers() {
    // 순서 변경 후, "위치"가 달라지는 지점에 다시 줄을 그어줌
    const formGroups = document.querySelectorAll('#measurementForm .form-group');
    let prevLocPrefix = null;

    formGroups.forEach((group, index) => {
        // 내부 input에서 location 정보 가져오기
        const input = group.querySelector('input');
        const loc = input ? input.dataset.location : '';
        const currLocPrefix = loc.substring(0, 3);

        // 기존 클래스 제거
        group.classList.remove('group-start');

        if (index > 0 && prevLocPrefix !== null && prevLocPrefix !== currLocPrefix) {
            group.classList.add('group-start');
        }
        prevLocPrefix = currLocPrefix;
    });
}
