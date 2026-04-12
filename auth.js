/**
 * auth.js — 수소·암모니아 브리핑 포털 인증 모듈
 * 허용 사번 목록 기반 단순 검증
 */

const AUTH = (() => {

  // ─────────────────────────────────────────────
  //  허용 사번 목록 — 관리자가 직접 수정
  // ─────────────────────────────────────────────
  const ALLOWED_IDS = [
    '14161422',
    '96112661',
    '95114458',
    '07160402',
    '07160771',
    '12162855',
    '14161252',
    '93151593',
    '17261053',
    '12261616',
    '20162163',
    '23161028',

    // 사번 추가 시 여기에 한 줄씩 추가
  ];

  // 세션 유지 시간 (기본 8시간)
  const SESSION_DURATION = 8 * 60 * 60 * 1000;
  const SESSION_KEY      = 'argus_auth_id';
  const SESSION_TIME_KEY = 'argus_auth_time';

  // ─────────────────────────────────────────────
  //  공개 API
  // ─────────────────────────────────────────────

  function login(employeeId) {
    const id = employeeId.trim().toUpperCase();
    if (!id) return { success: false, message: '사번을 입력해주세요.' };

    if (ALLOWED_IDS.map(x => x.toUpperCase()).includes(id)) {
      sessionStorage.setItem(SESSION_KEY, id);
      sessionStorage.setItem(SESSION_TIME_KEY, Date.now().toString());
      return { success: true };
    }

    return { success: false, message: '등록되지 않은 사번입니다.\n관리자에게 문의하세요.' };
  }

  function isLoggedIn() {
    const id   = sessionStorage.getItem(SESSION_KEY);
    const time = sessionStorage.getItem(SESSION_TIME_KEY);
    if (!id || !time) return false;

    if (Date.now() - parseInt(time, 10) > SESSION_DURATION) {
      logout();
      return false;
    }
    return ALLOWED_IDS.map(x => x.toUpperCase()).includes(id);
  }

  function logout() {
    sessionStorage.removeItem(SESSION_KEY);
    sessionStorage.removeItem(SESSION_TIME_KEY);
  }

  function requireAuth() {
    if (!isLoggedIn()) {
      const path  = window.location.pathname;
      const match = path.match(/^(\/[^/]+)/);
      const root  = match ? match[1] : '';
      window.location.replace(root + '/index.html');
    }
  }

  return { login, isLoggedIn, logout, requireAuth };

})();
