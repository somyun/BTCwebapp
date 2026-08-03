(function () {
    'use strict';

    const ALLOWED_DOMAIN = '@humetro.busan.kr';
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
    let authReadyResolve;
    const authReady = new Promise((resolve) => { authReadyResolve = resolve; });

    function allowed(user) {
        return Boolean(user?.email && user.email.toLowerCase().endsWith(ALLOWED_DOMAIN));
    }

    function elements() {
        return {
            overlay: document.getElementById('mapAuthOverlay'),
            message: document.getElementById('mapAuthMessage'),
            login: document.getElementById('mapAuthLoginBtn'),
            cancel: document.getElementById('mapAuthCancelBtn')
        };
    }

    function setMessage(text, isError = false) {
        const message = elements().message;
        if (!message) return;
        message.textContent = text;
        message.classList.toggle('error', isError);
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
        const { overlay, login, cancel } = elements();
        if (!overlay || !login || !cancel) return Promise.resolve(false);
        overlay.hidden = false;
        setMessage('');
        login.focus();
        return new Promise((resolve) => { hideDialog.resolve = resolve; });
    }

    async function signIn() {
        const { login } = elements();
        if (!auth || !login) return;
        login.disabled = true;
        setMessage('로그인 창을 여는 중입니다.');

        try {
            const provider = new firebase.auth.GoogleAuthProvider();
            provider.setCustomParameters({
                hd: 'humetro.busan.kr',
                prompt: 'select_account'
            });
            const result = await auth.signInWithPopup(provider);
            if (!allowed(result.user)) {
                await auth.signOut();
                setMessage('부산교통공사(@humetro.busan.kr) 계정만 이용할 수 있습니다.', true);
                return;
            }
            currentUser = result.user;
            hideDialog(true);
        } catch (error) {
            if (error?.code === 'auth/popup-closed-by-user') {
                setMessage('로그인이 취소되었습니다.', true);
            } else if (error?.code === 'auth/popup-blocked'
                || error?.code === 'auth/operation-not-supported-in-this-environment') {
                setMessage('로그인 페이지로 이동합니다.');
                await auth.signInWithRedirect(provider);
            } else if (error?.code === 'auth/operation-not-allowed') {
                setMessage('Firebase에서 Google 로그인을 먼저 활성화해야 합니다.', true);
            } else if (error?.code === 'auth/configuration-not-found') {
                setMessage('Firebase Authentication 시작 설정이 필요합니다.', true);
            } else {
                console.error('Map sign-in failed:', error);
                setMessage('로그인하지 못했습니다. 잠시 후 다시 시도해 주세요.', true);
            }
        } finally {
            login.disabled = false;
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

        elements().login?.addEventListener('click', signIn);
        elements().cancel?.addEventListener('click', () => hideDialog(false));
    }

    initialize();

    window.BWAAuth = {
        requireMapAccess,
        isAllowed: () => allowed(currentUser),
        currentEmail: () => currentUser?.email || ''
    };
}());
