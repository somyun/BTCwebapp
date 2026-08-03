(function () {
    'use strict';

    const KAKAO_JAVASCRIPT_KEY = '708065ee6e872ac3f158928a61d3252e';
    const HOPO_DEPOT_CENTER = { latitude: 35.286239431, longitude: 129.01636858 };

    let sdkPromise = null;
    let map = null;

    function loadKakaoMapSdk() {
        if (window.kakao && window.kakao.maps) {
            return new Promise((resolve) => window.kakao.maps.load(resolve));
        }

        if (sdkPromise) return sdkPromise;

        sdkPromise = new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${KAKAO_JAVASCRIPT_KEY}&autoload=false`;
            script.async = true;
            script.addEventListener('load', () => {
                if (!window.kakao || !window.kakao.maps) {
                    reject(new Error('카카오 지도 SDK를 시작하지 못했습니다.'));
                    return;
                }
                window.kakao.maps.load(resolve);
            });
            script.addEventListener('error', () => {
                reject(new Error('카카오 지도 SDK를 내려받지 못했습니다.'));
            });
            document.head.appendChild(script);
        });

        return sdkPromise;
    }

    function setMapType(type) {
        if (!map) return;

        const isSkyview = type === 'skyview';
        map.setMapTypeId(isSkyview
            ? window.kakao.maps.MapTypeId.HYBRID
            : window.kakao.maps.MapTypeId.ROADMAP);

        document.getElementById('roadmapBtn')?.classList.toggle('active', !isSkyview);
        document.getElementById('skyviewBtn')?.classList.toggle('active', isSkyview);
    }

    function moveToHopoDepot() {
        if (!map) return;
        map.panTo(new window.kakao.maps.LatLng(
            HOPO_DEPOT_CENTER.latitude,
            HOPO_DEPOT_CENTER.longitude
        ));
        map.setLevel(3);
    }

    function bindControls() {
        document.getElementById('roadmapBtn')?.addEventListener('click', () => setMapType('roadmap'));
        document.getElementById('skyviewBtn')?.addEventListener('click', () => setMapType('skyview'));
        document.getElementById('recenterMapBtn')?.addEventListener('click', moveToHopoDepot);
    }

    async function initialize() {
        const loading = document.getElementById('mapLoading');

        try {
            await loadKakaoMapSdk();

            if (!map) {
                map = new window.kakao.maps.Map(document.getElementById('kakaoMap'), {
                    center: new window.kakao.maps.LatLng(
                        HOPO_DEPOT_CENTER.latitude,
                        HOPO_DEPOT_CENTER.longitude
                    ),
                    level: 3
                });

                const zoomControl = new window.kakao.maps.ZoomControl();
                map.addControl(zoomControl, window.kakao.maps.ControlPosition.RIGHT);
                bindControls();
            }

            map.relayout();
            moveToHopoDepot();
            if (loading) loading.classList.add('hidden');
        } catch (error) {
            console.error('Kakao map initialization failed:', error);
            if (loading) {
                loading.textContent = '지도를 불러오지 못했습니다. 카카오 개발자 도메인 설정을 확인해 주세요.';
                loading.classList.add('error');
            }
        }
    }

    window.BWAMap = {
        initialize,
        relayout() {
            if (map) map.relayout();
        }
    };
}());
