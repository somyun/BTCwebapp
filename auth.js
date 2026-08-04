(function () {
    'use strict';

    const ALLOWED_DOMAIN = '@humetro.busan.kr';
    const HUMETRO_UID_PREFIX = 'humetro:';
    const GAS_API_URL = 'https://script.google.com/macros/s/AKfycbzuWS4Q5kTzDRH4IBpeXBa69KngElRdArtTCzTV0NDQsB3y4oABBIzrTLuPOZH5KOPP/exec';
    const firebaseConfig = {
        apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
        authDomain: 'btcwebapp-551bd.firebaseapp.com',
        projectId: 'btcwebapp-551bd',
        storageBucket: 'btcwebapp-551bd.firebasestorage.app',
        messagingSenderId: '237989935469',
        appId: '1:237989935469:web:07fc002a5c2ab2f5858264',
        measurementId: 'G-SFSSEHRPMN'
    };

    let auth = null;
    let currentUser = null;
    let pendingEmail = '';
    let authReadyResolve;
    const authReady = new Promise((resolve) => { authReadyResolve = resolve; });

    function normalizedEmail(value) {
        return String(value || '').trim().toLowerCase();
    }

    function emailIdFromAddress(value) {
        const email = normalizedEmail(value);
        return email.endsWith(ALLOWED_DOMAIN)
            ? email.slice(0, -ALLOWED_DOMAIN.length)
            : email;
    }

    function addressFromEmailId(value) {
        const emailId = normalizedEmail(value);
        if (!/^[a-z0-9._%+-]+$/.test(emailId)) return '';
        return `${emailId}${ALLOWED_DOMAIN}`;
    }

    function emailFromUser(user) {
        if (user?.email) return normalizedEmail(user.email);
        if (user?.uid?.startsWith(HUMETRO_UID_PREFIX)) {
            return normalizedEmail(user.uid.slice(HUMETRO_UID_PREFIX.length));
        }
        return '';
    }

    function allowed(user) {
        return emailFromUser(user).endsWith(ALLOWED_DOMAIN);
    }

    function elements() {
        return {
            overlay: document.getElementById('mapAuthOverlay'),
            message: document.getElementById('mapAuthMessage'),
            email: document.getElementById('mapAuthEmail'),
            send: document.getElementById('mapAuthSendCodeBtn'),
            codeSection: document.getElementById('mapAuthCodeSection'),
            code: document.getElementById('mapAuthCode'),
            verify: document.getElementById('mapAuthVerifyBtn'),
            cancel: document.getElementById('mapAuthCancelBtn')
        };
    }

    function setMessage(text, isError = false) {
        const message = elements().message;
        if (!message) return;
        message.textContent = text;
        message.classList.toggle('error', isError);
    }

    function setBusy(button, busy, busyText, normalText) {
        if (!button) return;
        button.disabled = busy;
        button.textContent = busy ? busyText : normalText;
    }

    async function callAuthApi(action, data) {
        const response = await fetch(GAS_API_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify({ action, ...data }),
            redirect: 'follow'
        });
        if (!response.ok) throw new Error(`AUTH_SERVER_${response.status}`);
        return response.json();
    }

    function hideDialog(result = false) {
        const { overlay } = elements();
        if (overlay) overlay.hidden = true;
        if (hideDialog.resolve) {
            hideDialog.resolve(result);
            hideDialog.resolve = null;
        }
    }

    function showDialog() {
        const { overlay, email, code, codeSection } = elements();
        if (!overlay || !email) return Promise.resolve(false);
        overlay.hidden = false;
        pendingEmail = '';
        email.value = emailIdFromAddress(emailFromUser(currentUser));
        if (code) code.value = '';
        if (codeSection) codeSection.hidden = true;
        setMessage('');
        email.focus();
        return new Promise((resolve) => { hideDialog.resolve = resolve; });
    }

    async function requestCode() {
        const { email, send, codeSection, code } = elements();
        const value = addressFromEmailId(email?.value);
        if (!value) {
            setMessage('부산교통공사 이메일 아이디를 입력해 주세요.', true);
            email?.focus();
            return;
        }

        setBusy(send, true, '발송 중…', '인증코드 다시 받기');
        setMessage('인증코드를 발송하고 있습니다.');
        try {
            const result = await callAuthApi('requestMapAuthCode', { email: value });
            if (!result.success) throw new Error(result.code || 'REQUEST_FAILED');
            pendingEmail = value;
            if (email) email.value = emailIdFromAddress(value);
            if (codeSection) codeSection.hidden = false;
            setMessage('인증코드를 발송했습니다. 메일의 6자리 숫자를 입력하세요.');
            code?.focus();
        } catch (error) {
            if (error.message === 'RATE_LIMITED') {
                setMessage('잠시 후 다시 요청해 주세요.', true);
            } else {
                console.error('Auth code request failed:', error);
                setMessage('인증코드를 발송하지 못했습니다. 잠시 후 다시 시도해 주세요.', true);
            }
        } finally {
            setBusy(send, false, '', pendingEmail ? '인증코드 다시 받기' : '인증코드 받기');
        }
    }

    async function verifyCode() {
        const { email, code, verify } = elements();
        const value = addressFromEmailId(email?.value);
        const codeValue = String(code?.value || '').replace(/\D/g, '').slice(0, 6);
        if (!pendingEmail || value !== pendingEmail) {
            setMessage('이메일이 변경되었습니다. 인증코드를 다시 받아 주세요.', true);
            return;
        }
        if (!/^\d{6}$/.test(codeValue)) {
            setMessage('메일로 받은 6자리 숫자를 입력해 주세요.', true);
            code?.focus();
            return;
        }

        setBusy(verify, true, '확인 중…', '인증하고 도면 열기');
        setMessage('인증코드를 확인하고 있습니다.');
        try {
            const result = await callAuthApi('verifyMapAuthCode', {
                email: pendingEmail,
                code: codeValue
            });
            if (!result.success || !result.token) throw new Error(result.code || 'VERIFY_FAILED');
            const credential = await auth.signInWithCustomToken(result.token);
            if (!allowed(credential.user)) {
                await auth.signOut();
                throw new Error('DOMAIN_REJECTED');
            }
            currentUser = credential.user;
            hideDialog(true);
        } catch (error) {
            if (error.message === 'INVALID_CODE') {
                setMessage('인증코드가 올바르지 않습니다.', true);
            } else if (error.message === 'CODE_EXPIRED') {
                setMessage('인증코드가 만료되었습니다. 새 코드를 받아 주세요.', true);
            } else if (error.message === 'TOO_MANY_ATTEMPTS') {
                setMessage('입력 횟수를 초과했습니다. 새 코드를 받아 주세요.', true);
            } else {
                console.error('Auth code verification failed:', error);
                setMessage('인증하지 못했습니다. 잠시 후 다시 시도해 주세요.', true);
            }
        } finally {
            setBusy(verify, false, '', '인증하고 도면 열기');
        }
    }

    async function requireMapAccess() {
        await authReady;
        if (allowed(currentUser)) return true;
        return showDialog();
    }

    function initialize() {
        try {
            if (!firebase.apps.length) firebase.initializeApp(firebaseConfig);
            auth = firebase.auth();
            auth.setPersistence(firebase.auth.Auth.Persistence.LOCAL).catch((error) => {
                console.error('Firebase auth persistence failed:', error);
            });
            auth.onAuthStateChanged(async (user) => {
                if (user && !allowed(user)) {
                    await auth.signOut();
                    currentUser = null;
                } else {
                    currentUser = user;
                }
                authReadyResolve();
            });
        } catch (error) {
            console.error('Firebase auth initialization failed:', error);
            authReadyResolve();
        }

        const { send, verify, cancel, email, code } = elements();
        send?.addEventListener('click', requestCode);
        verify?.addEventListener('click', verifyCode);
        cancel?.addEventListener('click', () => hideDialog(false));
        email?.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') requestCode();
        });
        code?.addEventListener('input', () => {
            code.value = code.value.replace(/\D/g, '').slice(0, 6);
        });
        code?.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') verifyCode();
        });
    }

    initialize();

    window.BWAAuth = {
        requireMapAccess,
        isAllowed: () => allowed(currentUser),
        currentEmail: () => emailFromUser(currentUser)
    };
}());
