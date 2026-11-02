'use strict';

/* ───────────── 공통 유틸 ───────────── */
const $ = (s, el = document) => el.querySelector(s);
const fmt = n => Math.round(n || 0).toLocaleString('ko-KR');
const won = n => fmt(n) + '원';
const pad = n => String(n).padStart(2, '0');
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const distText = m => m == null ? '-' : m < 1000 ? `${Math.round(m)}m` : `${(m / 1000).toFixed(m < 10000 ? 2 : 1).replace(/\.?0+$/, '')}km`;
const DOW = ['일', '월', '화', '수', '목', '금', '토'];

const store = {
  get(k, d) { try { const v = localStorage.getItem('cha.' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('cha.' + k, JSON.stringify(v)); } catch { /* 저장 불가 환경은 무시 */ } },
};

function toast(msg, isErr = false) {
  const t = $('#toast');
  t.textContent = msg;
  t.className = 'show' + (isErr ? ' err' : '');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => { t.className = ''; }, isErr ? 5000 : 2600);
}

/* ───────────── 상수 ───────────── */
// 회사 법인카드 양식 컬럼 → 내부 키 (헤더의 공백은 무시하고 비교)
const COLUMNS = [
  ['순번', 'seq'], ['카드번호', 'card'], ['이용자명', 'user'], ['이용일', 'date'], ['가맹점명', 'merchant'],
  ['이용금액', 'amount'], ['가맹점주소', 'addr'], ['사업자등록번호', 'bizNo'], ['승인번호', 'approvalNo'],
  ['과세유형', 'taxType'], ['전표구분', 'slipType'], ['미화금액', 'usd'], ['미화환율', 'usdRate'],
  ['승인시간', 'time'], ['가맹점업종코드', 'catCode'], ['가맹점업종명', 'catName'], ['유종', 'fuel'], ['리터당단가', 'literPrice'],
];
const REQUIRED = ['user', 'date', 'merchant', 'amount', 'addr'];
const ONLINE_RE = /PG|온라인|통신판매|전자상거래|인터넷/i;
const DOMESTIC_RE = /^(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충청|충북|충남|전라|전북|전남|경상|경북|경남|제주)/;

const DEFAULT_HOLIDAYS = `2025-01-01 신정
2025-01-27 임시공휴일
2025-01-28 설날 연휴
2025-01-29 설날
2025-01-30 설날 연휴
2025-03-03 삼일절 대체공휴일
2025-05-05 어린이날·부처님오신날
2025-05-06 대체공휴일
2025-06-03 대통령선거일
2025-06-06 현충일
2025-08-15 광복절
2025-10-03 개천절
2025-10-05 추석 연휴
2025-10-06 추석
2025-10-07 추석 연휴
2025-10-08 대체공휴일
2025-10-09 한글날
2025-12-25 성탄절
2026-01-01 신정
2026-02-16 설날 연휴
2026-02-17 설날
2026-02-18 설날 연휴
2026-03-02 삼일절 대체공휴일
2026-05-05 어린이날
2026-05-25 부처님오신날 대체공휴일
2026-06-03 지방선거일
2026-08-17 광복절 대체공휴일
2026-09-24 추석 연휴
2026-09-25 추석
2026-09-26 추석 연휴
2026-10-05 개천절 대체공휴일
2026-10-09 한글날
2026-12-25 성탄절`;

/* ───────────── 상태 ───────────── */
const state = {
  txns: [],
  users: [],
  homes: store.get('homes', {}),     // { 이용자명: { addr, lat, lng, src: 'geo'|'manual', approx } }
  analyzed: false,
  view: { mapUser: '', fUser: '', fView: 'all', fDay: 'all', q: '', sortKey: 'date', sortDir: -1, selected: null },
};

/* ───────────── 파싱 ───────────── */
function num(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v ?? '').replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function ymd(y, m, d) {
  const dt = new Date(y, m - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== m - 1 || dt.getDate() !== d) return null; // 2/29 같은 잘못된 날짜
  return `${y}-${pad(m)}-${pad(d)}`;
}

function parseDate(v) {
  if (v instanceof Date && !isNaN(v)) return ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === 'number' && v > 20000 && v < 80000) { const d = XLSX.SSF.parse_date_code(v); return ymd(d.y, d.m, d.d); }
  const s = String(v ?? '').trim();
  const m = s.match(/^(\d{4})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})/) || s.match(/^(\d{4})(\d{2})(\d{2})$/);
  return m ? ymd(+m[1], +m[2], +m[3]) : null;
}

function parseTime(v) {
  const hms = (h, m, s) => (h < 24 && m < 60 && s < 60) ? `${pad(h)}:${pad(m)}:${pad(s)}` : null;
  if (v instanceof Date && !isNaN(v)) return hms(v.getHours(), v.getMinutes(), v.getSeconds());
  if (typeof v === 'number') {
    if (v < 1) { const t = Math.round(v * 86400); return hms(Math.floor(t / 3600), Math.floor(t % 3600 / 60), t % 60); }
    const s = String(Math.round(v)).padStart(6, '0');
    return hms(+s.slice(0, 2), +s.slice(2, 4), +s.slice(4, 6));
  }
  const s = String(v ?? '').trim();
  const m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/) || s.match(/^(\d{2})(\d{2})(\d{2})?$/);
  return m ? hms(+m[1], +m[2], +(m[3] || 0)) : null;
}

const normAddr = a => String(a ?? '').replace(/\s+/g, ' ').trim();
const normHeader = h => String(h ?? '').replace(/\s+/g, '');

async function readWorkbook(file) {
  if (!window.XLSX) throw new Error('엑셀 라이브러리(vendor/xlsx.full.min.js)를 불러오지 못했습니다. vendor 폴더가 함께 올라갔는지 확인해 주세요');
  const buf = await file.arrayBuffer();
  if (/\.csv$/i.test(file.name)) {
    // 카드사 CSV는 CP949(EUC-KR)인 경우가 많아 UTF-8 실패 시 EUC-KR로 재시도
    let text;
    try { text = new TextDecoder('utf-8', { fatal: true }).decode(buf); } catch { text = new TextDecoder('euc-kr').decode(buf); }
    return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
  }
  return XLSX.read(buf, { type: 'array', cellDates: true });
}

function parseTransactions(wb) {
  for (const name of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null });
    // 카드사 엑셀은 상단에 제목 줄이 있는 경우가 있어 헤더 행을 찾는다
    const hIdx = rows.slice(0, 30).findIndex(r => {
      const hs = (r || []).map(normHeader);
      return hs.includes('이용자명') && hs.includes('가맹점주소');
    });
    if (hIdx < 0) continue;
    const header = rows[hIdx].map(normHeader);
    const colIdx = {};
    for (const [label, key] of COLUMNS) { const i = header.indexOf(label); if (i >= 0) colIdx[key] = i; }
    const missing = REQUIRED.filter(k => colIdx[k] == null);
    if (missing.length) throw new Error('필수 컬럼 없음: ' + missing.map(k => COLUMNS.find(c => c[1] === k)[0]).join(', '));

    const txns = [];
    rows.slice(hIdx + 1).forEach(r => {
      if (!r) return;
      const get = k => colIdx[k] == null ? null : r[colIdx[k]];
      const merchant = String(get('merchant') ?? '').trim();
      const user = String(get('user') ?? '').trim();
      if (!merchant || !user) return; // 합계·빈 줄
      const rawDate = get('date'), rawTime = get('time');
      const date = parseDate(rawDate), time = parseTime(rawTime);
      const raw = {};
      header.forEach((h, i) => { if (h) raw[rows[hIdx][i]] = r[i] instanceof Date ? (h === '승인시간' ? time : date) : r[i]; });
      txns.push({
        id: txns.length,
        seq: get('seq'), card: get('card'), user, merchant,
        dateRaw: rawDate instanceof Date ? date : String(rawDate ?? ''),
        date, dow: date ? new Date(date + 'T00:00:00').getDay() : null,
        time: time || '', hour: time ? +time.slice(0, 2) : null,
        amount: num(get('amount')),
        addr: normAddr(get('addr')),
        usd: num(get('usd')),
        catCode: get('catCode'), catName: String(get('catName') ?? '').trim(),
        raw,
      });
    });
    return { sheet: name, txns };
  }
  throw new Error('헤더 행(이용자명, 가맹점 주소 …)을 찾지 못했습니다. 회사 법인카드 양식인지 확인해 주세요.');
}

/* ───────────── 카카오 지도 / 지오코딩 ───────────── */
let kakaoReady = null;

function loadKakao() {
  if (window.kakao?.maps?.services) return Promise.resolve();
  const key = store.get('kakaoKey', '');
  if (!key) return Promise.reject(new Error('카카오 JavaScript 키를 먼저 설정해 주세요.'));
  if (kakaoReady) return kakaoReady;
  kakaoReady = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${encodeURIComponent(key)}&libraries=services&autoload=false`;
    s.onload = () => window.kakao?.maps ? kakao.maps.load(resolve) : reject(new Error('카카오 SDK 초기화 실패'));
    s.onerror = () => { kakaoReady = null; s.remove(); reject(new Error('카카오 SDK 로드 실패 — 키 또는 등록 도메인을 확인해 주세요.')); };
    document.head.appendChild(s);
  });
  return kakaoReady;
}

const Geo = {
  cache: store.get('geo', {}),
  errors: 0,
  answered: 0,
  save() { store.set('geo', this.cache); },
  clear() { this.cache = {}; this.save(); },

  call(fn) {
    // 콜백이 오지 않는 경우(도메인 미등록 등)를 대비해 타임아웃
    return new Promise(resolve => {
      const timer = setTimeout(() => { this.errors++; resolve(null); }, 8000);
      fn((result, status) => {
        clearTimeout(timer);
        const S = kakao.maps.services.Status;
        if (status === S.ERROR) this.errors++;
        else this.answered++;
        resolve(status === S.OK && result.length ? { lat: +result[0].y, lng: +result[0].x } : null);
      });
    });
  },
  addressSearch(q) { this.geocoder ||= new kakao.maps.services.Geocoder(); return this.call(cb => this.geocoder.addressSearch(q, cb)); },
  keywordSearch(q) { this.places ||= new kakao.maps.services.Places(); return this.call(cb => this.places.keywordSearch(q, cb)); },

  // "서울 강남구 테헤란로 281 효성빌딩 10층" → 도로명+건물번호까지만 잘라 정확도를 높인다
  candidates(a) {
    const precise = [];
    const road = a.match(/^(.*?(?:로|길)\s?\d+(?:-\d+)?)(?=\s|$)/);
    if (road) precise.push(road[1]);
    const jibun = a.match(/^(.*?(?:동|리|가)\d*\s\d+(?:-\d+)?)(?=\s|$)/);
    if (jibun) {
      precise.push(jibun[1]);
      const sub = a.slice(jibun[1].length).match(/^\s(\d+)(?=\s|$)/); // "서초동 1670 15" → 1670-15
      if (sub) precise.unshift(`${jibun[1]}-${sub[1]}`);
    }
    precise.push(a);
    const toks = a.split(' ');
    const approx = [];
    for (let n = toks.length - 1; n >= 3; n--) {
      if (/(로|길|동|리|가)$|\d/.test(toks[n - 1])) approx.push(toks.slice(0, n).join(' '));
    }
    return { precise: [...new Set(precise)], approx: [...new Set(approx)].filter(q => !precise.includes(q)) };
  },

  async resolve(addr) {
    const a = normAddr(addr);
    if (!a) return null;
    if (this.cache[a]) return this.cache[a];
    const { precise, approx } = this.candidates(a);
    let hit = null, approxFlag = false;
    for (const q of precise) { if ((hit = await this.addressSearch(q))) break; }
    if (!hit) { hit = await this.keywordSearch(a); approxFlag = !!hit; }
    if (!hit) for (const q of approx) { if ((hit = await this.addressSearch(q))) { approxFlag = true; break; } }
    if (!hit) return null;
    return (this.cache[a] = { ...hit, approx: approxFlag });
  },
};

async function pool(items, size, worker, onProgress) {
  let next = 0, done = 0;
  const run = async () => {
    while (next < items.length) {
      const i = next++;
      await worker(items[i], i);
      onProgress?.(++done, items.length);
    }
  };
  await Promise.all(Array.from({ length: Math.min(size, items.length) }, run));
}

/* ───────────── 분석 ───────────── */
function haversine(a, b) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (b.lat - a.lat) * rad, dLng = (b.lng - a.lng) * rad;
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(a.lat * rad) * Math.cos(b.lat * rad) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(h));
}

function settings() {
  const holidays = new Map();
  $('#holidayText').value.split('\n').forEach(line => {
    const m = line.trim().match(/^(\d{4}-\d{2}-\d{2})\s*(.*)$/);
    if (m) holidays.set(m[1], m[2] || '공휴일');
  });
  return {
    radius: +$('#radius').value,
    excludeOnline: $('#excludeOnline').checked,
    nightStart: +$('#nightStart').value,
    nightEnd: +$('#nightEnd').value,
    holidays,
  };
}

function isNight(hour, s, e) {
  if (hour == null || s === e) return false;
  return s < e ? hour >= s && hour < e : hour >= s || hour < e;
}

function compute() {
  const cfg = settings();
  for (const t of state.txns) {
    t.excluded = ''; t.dist = null; t.loc = null; t.near = false; t.risk = '';
    t.offDay = t.date ? (cfg.holidays.get(t.date) || (t.dow === 0 || t.dow === 6 ? DOW[t.dow] + '요일' : '')) : '';
    t.dayType = !t.date ? '날짜오류' : cfg.holidays.has(t.date) ? '공휴일' : t.offDay ? '주말' : '평일';
    t.night = isNight(t.hour, cfg.nightStart, cfg.nightEnd);

    if (t.usd > 0 || !DOMESTIC_RE.test(t.addr)) { t.excluded = '해외·주소없음'; continue; }
    if (cfg.excludeOnline && ONLINE_RE.test(t.catName)) { t.excluded = '온라인·PG'; continue; }
    const home = state.homes[t.user];
    t.loc = Geo.cache[t.addr] || null;
    if (!home?.lat) { t.excluded = '집주소 미등록'; continue; }
    if (!t.loc) { t.excluded = '가맹점 위치 미확인'; continue; }
    t.dist = haversine(home, t.loc);
    t.near = t.dist <= cfg.radius;
    t.risk = t.near ? (t.offDay || t.night ? 'high' : 'mid') : '';
  }
  return cfg;
}

async function runAnalysis() {
  if (!state.txns.length) return toast('먼저 카드 사용내역 파일을 불러와 주세요.', true);
  const btn = $('#btnRun');
  btn.disabled = true;
  const prog = $('#progress');
  const setProg = (done, total, label) => {
    prog.hidden = false;
    prog.querySelector('i').style.width = (total ? done / total * 100 : 100) + '%';
    prog.querySelector('span').textContent = `${label} ${done} / ${total}`;
  };
  try {
    await loadKakao();
    updateKeyStatus(true);
    Geo.errors = Geo.answered = 0;

    // 1) 집 주소 좌표
    const homeUsers = state.users.filter(u => state.homes[u]?.addr && !state.homes[u].lat);
    await pool(homeUsers, 3, async u => {
      const h = state.homes[u];
      const hit = await Geo.resolve(h.addr);
      if (hit) Object.assign(h, { lat: hit.lat, lng: hit.lng, approx: hit.approx, src: 'geo' });
      else h.failed = true;
    }, (d, n) => setProg(d, n, '집 주소 좌표 변환'));
    store.set('homes', state.homes);
    renderHomes();

    // 2) 가맹점 주소 좌표 (해외·온라인 결제는 좌표 불필요)
    const cfg = settings();
    const addrs = [...new Set(state.txns
      .filter(t => DOMESTIC_RE.test(t.addr) && !(t.usd > 0) && !(cfg.excludeOnline && ONLINE_RE.test(t.catName)))
      .map(t => t.addr))].filter(a => !Geo.cache[a]);
    await pool(addrs, 4, a => Geo.resolve(a), (d, n) => setProg(d, n, '가맹점 주소 좌표 변환'));
    Geo.save();

    if (Geo.errors > 0 && Geo.answered === 0) {
      throw new Error('카카오 주소 검색이 모두 실패했습니다. 카카오 디벨로퍼스에 현재 도메인(' + location.origin + ')이 등록됐는지, 카카오맵 사용 설정이 켜져 있는지 확인해 주세요.');
    }
    setProg(1, 1, '완료');
    state.analyzed = true;
    $('#results').hidden = false;
    refresh();
    $('#results').scrollIntoView({ behavior: 'smooth' });
  } catch (e) {
    toast(e.message, true);
    if (/키|SDK/.test(e.message)) updateKeyStatus(false);
  } finally {
    btn.disabled = false;
    setTimeout(() => { prog.hidden = true; }, 1500);
  }
}

/* ───────────── 렌더링: 입력 단계 ───────────── */
function renderHomes() {
  const box = $('#homeList');
  if (!state.users.length) return;
  box.innerHTML = state.users.map(u => {
    const h = state.homes[u] || {};
    const badge = h.lat
      ? (h.src === 'manual' ? '<span class="badge manual">지도 지정</span>' : `<span class="badge ok">${h.approx ? '근사 위치' : '좌표 확인'}</span>`)
      : h.failed ? '<span class="badge fail">주소 못 찾음</span>' : '<span class="badge none">미확인</span>';
    return `<div class="home-row" data-user="${esc(u)}">
      <span class="name" title="${esc(u)}">${esc(u)}</span>
      <input type="text" value="${esc(h.addr || '')}" placeholder="집 주소 (예: 서울 서초구 서초중앙로 200)">
      ${badge}
      <button type="button" class="icon-btn" data-pick title="지도에서 직접 지정">📍</button>
    </div>`;
  }).join('');
}

function onHomeInput(e) {
  const row = e.target.closest('.home-row');
  if (!row || e.target.tagName !== 'INPUT') return;
  const u = row.dataset.user;
  const addr = normAddr(e.target.value);
  if ((state.homes[u]?.addr || '') === addr) return;
  state.homes[u] = { addr };
  store.set('homes', state.homes);
  renderHomes();
}

function updateKeyStatus(ok) {
  const el = $('#keyStatus');
  const hasKey = !!store.get('kakaoKey', '');
  if (ok === true) { el.textContent = '카카오 지도 연결됨'; el.className = 'pill ok'; }
  else if (!hasKey) { el.textContent = '카카오 키 미설정'; el.className = 'pill warn'; }
  else if (ok === false) { el.textContent = '카카오 연결 실패'; el.className = 'pill warn'; }
  else { el.textContent = '카카오 키 설정됨'; el.className = 'pill'; }
}

/* ───────────── 렌더링: 결과 ───────────── */
function refresh() {
  if (!state.analyzed) return;
  const cfg = compute();
  renderKpis(cfg);
  renderUserTable();
  renderDistChart(cfg);
  renderDayTable();
  renderAnomalies(cfg);
  renderMap(cfg);
  renderTxTable();
}

function renderKpis(cfg) {
  const all = state.txns;
  const analyzed = all.filter(t => !t.excluded);
  const near = analyzed.filter(t => t.near);
  const high = near.filter(t => t.risk === 'high');
  const unknown = all.filter(t => t.excluded === '집주소 미등록' || t.excluded === '가맹점 위치 미확인');
  const sum = arr => arr.reduce((s, t) => s + t.amount, 0);
  const pct = (a, b) => b ? (a / b * 100).toFixed(1) + '%' : '-';
  const sa = sum(analyzed);
  $('#kpis').innerHTML = `
    <div class="kpi"><div class="k-label">분석 대상 사용</div>
      <div class="k-value">${fmt(analyzed.length)}<small>건</small></div>
      <div class="k-sub">${won(sa)} · 전체 ${fmt(all.length)}건 중 제외 ${fmt(all.length - analyzed.length)}건</div></div>
    <div class="kpi"><div class="k-label"><i class="dot-i" style="background:var(--near)"></i>집 근처 ${distText(cfg.radius)} 이내</div>
      <div class="k-value">${fmt(near.length)}<small>건</small></div>
      <div class="k-sub">${won(sum(near))} · 금액 비중 ${pct(sum(near), sa)}</div></div>
    <div class="kpi alert"><div class="k-label">⚠ 집 근처 + 휴일·심야</div>
      <div class="k-value">${fmt(high.length)}<small>건</small></div>
      <div class="k-sub">${won(sum(high))} · 심야 ${pad(cfg.nightStart)}시~${pad(cfg.nightEnd)}시</div></div>
    <div class="kpi"><div class="k-label">위치 확인 불가</div>
      <div class="k-value">${fmt(unknown.length)}<small>건</small></div>
      <div class="k-sub">집주소 미등록 ${fmt(unknown.filter(t => t.excluded === '집주소 미등록').length)} · 가맹점 미확인 ${fmt(unknown.filter(t => t.excluded === '가맹점 위치 미확인').length)}</div></div>`;
}

function userStats() {
  return state.users.map(u => {
    const ts = state.txns.filter(t => t.user === u);
    const an = ts.filter(t => !t.excluded);
    const near = an.filter(t => t.near);
    const amt = an.reduce((s, t) => s + t.amount, 0);
    const nearAmt = near.reduce((s, t) => s + t.amount, 0);
    return {
      user: u, home: !!state.homes[u]?.lat, total: ts.length, analyzed: an.length,
      near: near.length, nearAmt, amt, ratio: amt ? nearAmt / amt : 0,
      high: near.filter(t => t.risk === 'high').length,
    };
  });
}

function renderUserTable() {
  const rows = userStats().sort((a, b) => b.high - a.high || b.nearAmt - a.nearAmt);
  $('#userTable').innerHTML = `<thead><tr><th>이용자</th><th class="r">분석건수</th><th class="r">집근처</th><th class="r">집근처 금액</th><th class="r">금액비중</th><th class="r">휴일·심야</th></tr></thead>
    <tbody>${rows.map(r => `<tr data-user="${esc(r.user)}" class="${state.view.fUser === r.user ? 'sel' : ''}">
      <td>${esc(r.user)}${r.home ? '' : ' <span class="tag ex">집주소 없음</span>'}</td>
      <td class="r">${fmt(r.analyzed)}</td><td class="r">${fmt(r.near)}</td><td class="r">${won(r.nearAmt)}</td>
      <td class="r">${r.analyzed ? (r.ratio * 100).toFixed(1) + '%' : '-'}</td>
      <td class="r">${r.high ? `<span class="tag high">${r.high}</span>` : '0'}</td></tr>`).join('')}</tbody>`;
}

function renderDistChart(cfg) {
  const ts = state.txns.filter(t => !t.excluded && (!state.view.fUser || t.user === state.view.fUser));
  const edges = [300, 500, 1000, 2000, 3000, 5000, 10000, 30000, Infinity];
  const buckets = edges.map((e, i) => ({ lo: i ? edges[i - 1] : 0, hi: e, n: 0, amt: 0 }));
  ts.forEach(t => { const b = buckets.find(b => t.dist <= b.hi); b.n++; b.amt += t.amount; });
  const max = Math.max(1, ...buckets.map(b => b.n));
  const label = b => b.hi === Infinity ? `${distText(b.lo)} 초과` : `~${distText(b.hi)}`;
  $('#distChart').innerHTML = buckets.map((b, i) => {
    const inR = b.hi <= cfg.radius;
    return `<div class="hbar-row" data-i="${i}">
      <span class="lbl">${inR ? '<b>반경 내</b> ' : ''}${label(b)}</span>
      <div class="track"><div class="fill${inR ? ' in' : ''}" style="width:${b.n / max * 100}%"></div></div>
      <span class="val">${fmt(b.n)}건</span></div>`;
  }).join('') + `<div class="hbar-note">${state.view.fUser ? esc(state.view.fUser) + ' · ' : ''}분석 대상 ${fmt(ts.length)}건 기준. 빨간 막대는 설정 반경(${distText(cfg.radius)}) 이내 구간입니다.</div>`;
  $('#distChart').dataset.buckets = JSON.stringify(buckets.map(b => ({ l: label(b), n: b.n, amt: b.amt })));
}

// 휴일 여부는 위치와 무관하게 확인할 수 있도록 전체 건수 기준, 집근처는 분석 대상 기준
function renderDayTable() {
  const ts = state.txns.filter(t => !state.view.fUser || t.user === state.view.fUser);
  const rows = [['weekday', '평일'], ['weekend', '주말'], ['publicHoliday', '공휴일'], ['holiday', '휴일 합계']].map(([key, label]) => {
    const all = ts.filter(DAY_FILTERS[key]);
    const near = all.filter(t => t.near);
    const sum = arr => arr.reduce((s, t) => s + t.amount, 0);
    return { key, label, n: all.length, amt: sum(all), nearN: near.length, nearAmt: sum(near) };
  });
  const bad = ts.filter(t => t.dayType === '날짜오류').length;
  $('#dayTable').innerHTML = `<thead><tr><th>구분</th><th class="r">전체 건수</th><th class="r">전체 금액</th><th class="r">집근처 건수</th><th class="r">집근처 금액</th></tr></thead>
    <tbody>${rows.map(r => `<tr data-day="${r.key}" class="${state.view.fDay === r.key ? 'sel' : ''}${r.key === 'holiday' ? ' total' : ''}">
      <td>${r.key === 'weekday' ? '<span class="tag weekday">평일</span>' : r.key === 'holiday' ? '<b>휴일 합계</b>' : `<span class="tag off">${r.label}</span>`}</td>
      <td class="r">${fmt(r.n)}</td><td class="r">${won(r.amt)}</td>
      <td class="r">${fmt(r.nearN)}</td><td class="r">${won(r.nearAmt)}</td></tr>`).join('')}</tbody>`;
  $('#dayNote').textContent = (state.view.fUser ? state.view.fUser + ' · ' : '')
    + '행을 클릭하면 상세 내역이 그 구분으로 걸러집니다.' + (bad ? ` 날짜를 인식하지 못한 ${bad}건은 제외됩니다.` : '');
}

function setDayFilter(key) {
  state.view.fDay = key;
  $('#fDay').value = key;
  renderDayTable();
  renderTxTable();
}

/* ── 지도 ── */
const M = { map: null, overlays: [], popup: null };

function ensureMap() {
  if (M.map) return M.map;
  M.map = new kakao.maps.Map($('#map'), { center: new kakao.maps.LatLng(37.5665, 126.978), level: 7 });
  M.map.addControl(new kakao.maps.ZoomControl(), kakao.maps.ControlPosition.RIGHT);
  M.popup = new kakao.maps.CustomOverlay({ position: M.map.getCenter(), content: document.createElement('div'), yAnchor: 1, zIndex: 10, clickable: true });
  kakao.maps.event.addListener(M.map, 'click', () => M.popup.setMap(null));
  return M.map;
}

function renderMap(cfg) {
  if (!window.kakao?.maps) return;
  const map = ensureMap();
  map.relayout();
  M.overlays.forEach(o => o.setMap(null));
  M.overlays = [];
  M.popup.setMap(null);

  const sel = state.view.mapUser;
  const users = sel ? [sel] : state.users;
  // 기본 화면은 집 반경과 반경 내 가맹점에 맞추고, 집이 없을 때만 전체 가맹점에 맞춘다
  const bounds = new kakao.maps.LatLngBounds();
  const allBounds = new kakao.maps.LatLngBounds();
  let hasBounds = false, hasAll = false;
  const add = o => { o.setMap(map); M.overlays.push(o); };

  for (const u of users) {
    const h = state.homes[u];
    if (!h?.lat) continue;
    const pos = new kakao.maps.LatLng(h.lat, h.lng);
    add(new kakao.maps.Circle({ center: pos, radius: cfg.radius, strokeWeight: 2, strokeColor: '#d03b3b', strokeOpacity: .8, strokeStyle: 'dash', fillColor: '#d03b3b', fillOpacity: .07 }));
    const el = document.createElement('div');
    el.className = 'mk-home';
    el.textContent = '🏠 ' + u;
    add(new kakao.maps.CustomOverlay({ position: pos, content: el, yAnchor: 1.4, zIndex: 5 }));
    bounds.extend(pos); hasBounds = true;
    // 반경 원 전체가 보이도록 동서남북 끝점도 포함
    const dLat = cfg.radius / 111320, dLng = cfg.radius / (111320 * Math.cos(h.lat * Math.PI / 180));
    bounds.extend(new kakao.maps.LatLng(h.lat + dLat, h.lng + dLng));
    bounds.extend(new kakao.maps.LatLng(h.lat - dLat, h.lng - dLng));
  }

  // 같은 좌표의 사용건은 하나의 점으로 묶는다
  const groups = new Map();
  state.txns.filter(t => !t.excluded && users.includes(t.user)).forEach(t => {
    const k = t.loc.lat.toFixed(6) + ',' + t.loc.lng.toFixed(6);
    if (!groups.has(k)) groups.set(k, { loc: t.loc, list: [] });
    groups.get(k).list.push(t);
  });
  for (const g of groups.values()) {
    const near = g.list.some(t => t.near);
    const el = document.createElement('div');
    el.className = `mk-dot ${near ? 'near' : 'far'}${g.list.length > 1 ? ' multi' : ''}`;
    el.textContent = g.list.length > 1 ? g.list.length : '';
    el.title = g.list[0].merchant;
    const pos = new kakao.maps.LatLng(g.loc.lat, g.loc.lng);
    el.addEventListener('click', e => { e.stopPropagation(); openPopup(pos, g.list); });
    add(new kakao.maps.CustomOverlay({ position: pos, content: el, zIndex: near ? 4 : 3, clickable: true }));
    if (near) { bounds.extend(pos); hasBounds = true; }
    allBounds.extend(pos); hasAll = true;
  }
  if (hasBounds) map.setBounds(bounds, 40, 40, 40, 40);
  else if (hasAll) map.setBounds(allBounds, 40, 40, 40, 40);
}

function openPopup(pos, list) {
  const sorted = [...list].sort((a, b) => (b.date || '').localeCompare(a.date || '') || b.time.localeCompare(a.time));
  const el = document.createElement('div');
  el.className = 'mk-pop';
  const total = list.reduce((s, t) => s + t.amount, 0);
  el.innerHTML = `<h4><span>${esc(sorted[0].merchant)}</span><button class="close" type="button" aria-label="닫기">×</button></h4>
    <div class="addr">${esc(sorted[0].addr)} · 집에서 ${distText(sorted[0].dist)}</div>
    <div class="addr">${list.length}건 · ${won(total)}</div>
    <table>${sorted.slice(0, 30).map(t => `<tr><td>${esc(t.date || t.dateRaw)} (${t.dow != null ? DOW[t.dow] : '-'})</td><td>${esc(t.time.slice(0, 5))}</td><td>${esc(t.user)}</td><td class="r">${won(t.amount)}</td><td>${popupTag(t)}</td></tr>`).join('')}</table>
    ${sorted.length > 30 ? `<div class="addr">외 ${sorted.length - 30}건</div>` : ''}`;
  el.querySelector('.close').addEventListener('click', () => M.popup.setMap(null));
  el.addEventListener('mousedown', e => e.stopPropagation());
  M.popup.setContent(el);
  M.popup.setPosition(pos);
  M.popup.setMap(M.map);
}

function focusTxn(t) {
  if (!t.loc || !M.map) return;
  const pos = new kakao.maps.LatLng(t.loc.lat, t.loc.lng);
  const same = state.txns.filter(x => x.loc && !x.excluded && x.loc.lat === t.loc.lat && x.loc.lng === t.loc.lng && (!state.view.mapUser || x.user === state.view.mapUser));
  M.map.setLevel(4);
  M.map.panTo(pos);
  openPopup(pos, same.length ? same : [t]);
  $('#map').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// 팝업은 폭이 좁아 가장 중요한 표시 하나만 보여준다
function popupTag(t) {
  if (t.risk === 'high') return `<span class="tag high">${t.offDay ? '휴일' : '심야'}</span>`;
  if (t.offDay) return `<span class="tag off">휴일</span>`;
  if (t.night) return '<span class="tag night">심야</span>';
  return '';
}

// 평일은 담담하게, 주말·공휴일은 이름까지 태그로 보여준다
function dayTag(t) {
  if (t.dayType === '평일') return '<span class="tag weekday">평일</span>';
  if (t.dayType === '날짜오류') return '<span class="tag ex">날짜오류</span>';
  return `<span class="tag off">${esc(t.offDay)}</span>`;
}

const DAY_FILTERS = {
  weekday: t => t.dayType === '평일',
  holiday: t => t.dayType === '주말' || t.dayType === '공휴일',
  weekend: t => t.dayType === '주말',
  publicHoliday: t => t.dayType === '공휴일',
};

/* ── 상세 테이블 ── */
function riskTags(t) {
  const tags = [];
  if (t.risk === 'high') tags.push('<span class="tag high">집근처·휴일/심야</span>');
  else if (t.risk === 'mid') tags.push('<span class="tag mid">집근처</span>');
  if (t.night) tags.push('<span class="tag night">심야</span>');
  if (t.excluded) tags.push(`<span class="tag ex">${esc(t.excluded)}</span>`);
  return tags.join('');
}

const TX_COLS = [
  { key: 'risk', label: '판정', val: t => ({ high: 3, mid: 2 }[t.risk] || (t.excluded ? 0 : 1)) },
  { key: 'date', label: '이용일', val: t => (t.date || t.dateRaw.replace(/[./]/g, '-')) + ' ' + t.time },
  { key: 'dayType', label: '평일/휴일', val: t => ({ 평일: 1, 주말: 2, 공휴일: 3 }[t.dayType] || 0) },
  { key: 'time', label: '승인시간', val: t => t.time },
  { key: 'user', label: '이용자', val: t => t.user },
  { key: 'merchant', label: '가맹점명', val: t => t.merchant },
  { key: 'catName', label: '업종', val: t => t.catName },
  { key: 'amount', label: '이용금액', val: t => t.amount, r: true },
  { key: 'dist', label: '집까지 거리', val: t => t.dist ?? Infinity, r: true },
  { key: 'addr', label: '가맹점 주소', val: t => t.addr },
];

function filteredTxns() {
  const { fUser, fView, fDay, q } = state.view;
  const qq = q.trim().toLowerCase();
  return state.txns.filter(t => {
    if (fUser && t.user !== fUser) return false;
    if (state.view.anomalyIds && !state.view.anomalyIds.has(t.id)) return false;
    if (DAY_FILTERS[fDay] && !DAY_FILTERS[fDay](t)) return false;
    if (fView === 'near' && !t.near) return false;
    if (fView === 'high' && t.risk !== 'high') return false;
    if (fView === 'off' && !(t.offDay || t.night)) return false;
    if (fView === 'excluded' && !t.excluded) return false;
    if (qq && !`${t.merchant} ${t.addr} ${t.catName}`.toLowerCase().includes(qq)) return false;
    return true;
  });
}

function renderTxTable() {
  const { sortKey, sortDir } = state.view;
  const col = TX_COLS.find(c => c.key === sortKey);
  const rows = filteredTxns().sort((a, b) => {
    const x = col.val(a), y = col.val(b);
    return (typeof x === 'number' ? x - y : String(x).localeCompare(String(y), 'ko')) * sortDir;
  });
  const anomName = anomalyFilterName();
  $('#fCount').innerHTML = `${fmt(rows.length)}건 · ${won(rows.reduce((s, t) => s + t.amount, 0))}`
    + (anomName ? ` <button type="button" class="chip on" data-clear-anom title="이상징후 필터 해제">이상징후: ${esc(anomName)} ✕</button>` : '');
  const LIMIT = 2000;
  $('#txTable').innerHTML = `<thead><tr>${TX_COLS.map(c => `<th data-key="${c.key}" class="${c.r ? 'r ' : ''}${c.key === sortKey ? 'sorted' + (sortDir > 0 ? ' asc' : '') : ''}">${c.label}</th>`).join('')}</tr></thead>
    <tbody>${rows.length ? rows.slice(0, LIMIT).map(t => `<tr data-id="${t.id}" class="${t.risk ? 'risk-' + t.risk : ''}">
      <td>${riskTags(t) || '<span class="muted small">-</span>'}</td>
      <td>${esc(t.date || t.dateRaw)}${t.dow != null ? ` (${DOW[t.dow]})` : ''}</td>
      <td>${dayTag(t)}</td>
      <td>${esc(t.time)}</td>
      <td>${esc(t.user)}</td>
      <td class="merchant" title="${esc(t.merchant)}">${esc(t.merchant)}</td>
      <td>${esc(t.catName)}</td>
      <td class="r">${won(t.amount)}</td>
      <td class="r">${distText(t.dist)}${t.loc?.approx ? ' <span class="tag approx" title="상세주소를 못 찾아 도로/동 단위로 근사한 위치">근사</span>' : ''}</td>
      <td class="addr" title="${esc(t.addr)}">${esc(t.addr)}</td></tr>`).join('')
    : `<tr><td colspan="${TX_COLS.length}" class="empty">조건에 맞는 내역이 없습니다.</td></tr>`}
    ${rows.length > LIMIT ? `<tr><td colspan="${TX_COLS.length}" class="empty">상위 ${fmt(LIMIT)}건만 표시합니다. 전체는 엑셀로 저장해 확인하세요.</td></tr>` : ''}</tbody>`;
}

function fillUserSelects() {
  const opts = `<option value="">전체 이용자</option>` + state.users.map(u => `<option>${esc(u)}</option>`).join('');
  $('#mapUser').innerHTML = opts;
  $('#fUser').innerHTML = opts;
  if (state.users.length === 1) state.view.mapUser = state.view.fUser = state.users[0];
  $('#mapUser').value = state.view.mapUser;
  $('#fUser').value = state.view.fUser;
}

function selectUser(u) {
  state.view.mapUser = state.view.fUser = u;
  $('#mapUser').value = u;
  $('#fUser').value = u;
  refresh();
}

/* ───────────── 엑셀 내보내기 ───────────── */
function exportExcel() {
  const cfg = settings();
  const rows = filteredTxns().map(t => ({
    ...t.raw,
    '집까지 거리(m)': t.dist == null ? '' : Math.round(t.dist),
    [`집근처(${distText(cfg.radius)} 이내)`]: t.excluded ? '' : t.near ? 'Y' : 'N',
    '평일/휴일': t.dayType,
    '휴일명': t.dayType === '평일' ? '' : t.offDay,
    '심야': t.night ? 'Y' : '',
    '판정': t.risk === 'high' ? '집근처·휴일/심야' : t.risk === 'mid' ? '집근처' : '',
    '분석 제외 사유': t.excluded,
    '위치 정확도': t.loc ? (t.loc.approx ? '근사' : '정확') : '',
    '이상징후': anomalyNames(t),
  }));
  const summary = userStats().map(r => ({
    '이용자': r.user, '전체건수': r.total, '분석건수': r.analyzed, '집근처 건수': r.near,
    '집근처 금액': r.nearAmt, '집근처 금액비중(%)': +(r.ratio * 100).toFixed(1), '집근처+휴일·심야 건수': r.high,
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '상세내역');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), '이용자별 요약');
  const d = new Date();
  XLSX.writeFile(wb, `법인카드_집근처분석_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}.xlsx`);
}

/* ───────────── 집 주소 가져오기/내보내기 ───────────── */
async function importHomes(file) {
  const wb = await readWorkbook(file);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
  let n = 0;
  rows.forEach(r => {
    const keys = Object.keys(r);
    const nameKey = keys.find(k => /이용자|이름|성명/.test(k));
    const addrKey = keys.find(k => /주소/.test(k));
    if (!nameKey || !addrKey) return;
    const u = String(r[nameKey]).trim(), addr = normAddr(r[addrKey]);
    if (!u || !addr) return;
    if (state.homes[u]?.addr !== addr) state.homes[u] = { addr };
    if (!state.users.includes(u) && !state.txns.length) state.users.push(u);
    n++;
  });
  store.set('homes', state.homes);
  renderHomes();
  toast(n ? `집 주소 ${n}건을 불러왔습니다.` : '이용자명 / 집주소 컬럼을 찾지 못했습니다.', !n);
}

function exportHomes() {
  const names = state.users.length ? state.users : Object.keys(state.homes);
  const csv = '﻿이용자명,집주소\n' + names.map(u => `"${u.replace(/"/g, '""')}","${(state.homes[u]?.addr || '').replace(/"/g, '""')}"`).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  a.download = '이용자_집주소.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ───────────── 지도에서 집 위치 지정 ───────────── */
const Pick = { map: null, marker: null, user: null, pos: null };

async function openPick(user) {
  try { await loadKakao(); updateKeyStatus(true); } catch (e) { return toast(e.message, true); }
  Pick.user = user; Pick.pos = null;
  $('#pickTitle').textContent = `${user} — 지도에서 집 위치 지정`;
  $('#btnPickSave').disabled = true;
  $('#pickQuery').value = state.homes[user]?.addr || '';
  $('#pickDlg').showModal();
  const h = state.homes[user];
  const center = h?.lat ? new kakao.maps.LatLng(h.lat, h.lng) : new kakao.maps.LatLng(37.5665, 126.978);
  if (!Pick.map) {
    Pick.map = new kakao.maps.Map($('#pickMap'), { center, level: 4 });
    Pick.marker = new kakao.maps.Marker();
    kakao.maps.event.addListener(Pick.map, 'click', e => setPickPos(e.latLng));
  }
  Pick.map.relayout();
  Pick.map.setCenter(center);
  Pick.marker.setMap(null);
  if (h?.lat) setPickPos(center);
}

function setPickPos(latLng) {
  Pick.pos = { lat: latLng.getLat(), lng: latLng.getLng() };
  Pick.marker.setPosition(latLng);
  Pick.marker.setMap(Pick.map);
  $('#btnPickSave').disabled = false;
}

async function pickSearch() {
  const q = normAddr($('#pickQuery').value);
  if (!q) return;
  const hit = await Geo.addressSearch(q) || await Geo.keywordSearch(q);
  if (!hit) return toast('위치를 찾지 못했습니다.', true);
  const ll = new kakao.maps.LatLng(hit.lat, hit.lng);
  Pick.map.setCenter(ll);
  Pick.map.setLevel(3);
  setPickPos(ll);
}

/* ───────────── 이벤트 연결 ───────────── */
async function loadFile(file) {
  const info = $('#fileInfo');
  info.textContent = `읽는 중… (${file.name})`;
  try {
    let wb;
    try { wb = await readWorkbook(file); } catch (e) {
      throw new Error(`파일을 읽지 못했습니다 (${e.message}). 사내 보안(DRM) 프로그램이 걸린 파일이면 엑셀에서 '다른 이름으로 저장' 후 다시 시도해 주세요.`);
    }
    const { sheet, txns } = parseTransactions(wb);
    state.txns = txns;
    state.users = [...new Set(txns.map(t => t.user))];
    state.analyzed = false;
    state.view.mapUser = state.view.fUser = '';
    resetAi();
    $('#results').hidden = true;
    const badDates = txns.filter(t => !t.date).length;
    const total = txns.reduce((s, t) => s + t.amount, 0);
    info.innerHTML = `<b>${esc(file.name)}</b> (${esc(sheet)})<br>${fmt(txns.length)}건 · ${won(total)} · 이용자 ${state.users.length}명`
      + (badDates ? `<br><span class="err">날짜를 인식하지 못한 행 ${badDates}건 (휴일 판정 제외)</span>` : '');
    renderHomes();
    fillUserSelects();
  } catch (e) {
    info.innerHTML = `<span class="err">${esc(e.message)}</span>`;
  }
}

function init() {
  // 심야 시간 선택지
  const hours = Array.from({ length: 24 }, (_, h) => `<option value="${h}">${pad(h)}</option>`).join('');
  $('#nightStart').innerHTML = hours;
  $('#nightEnd').innerHTML = hours;
  const saved = store.get('settings', {});
  $('#radius').value = saved.radius ?? 1000;
  $('#excludeOnline').checked = saved.excludeOnline ?? true;
  $('#nightStart').value = saved.nightStart ?? 22;
  $('#nightEnd').value = saved.nightEnd ?? 6;
  $('#holidayText').value = saved.holidayText ?? DEFAULT_HOLIDAYS;

  const syncRadius = () => {
    const r = +$('#radius').value;
    $('#radiusLabel').textContent = distText(r);
    document.querySelectorAll('.chip').forEach(c => c.classList.toggle('on', +c.dataset.r === r));
  };
  const saveSettings = () => store.set('settings', {
    radius: +$('#radius').value, excludeOnline: $('#excludeOnline').checked,
    nightStart: +$('#nightStart').value, nightEnd: +$('#nightEnd').value, holidayText: $('#holidayText').value,
  });
  syncRadius();

  // 파일 업로드
  // 드롭 영역 위에서는 파일 입력이 직접 받고, 그 밖에 떨어뜨려도 페이지가 파일을 열어버리지 않게 받아 처리한다
  const drop = $('#dropZone'), fileInput = $('#fileInput');
  fileInput.addEventListener('change', e => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; });
  fileInput.addEventListener('dragenter', () => drop.classList.add('over'));
  fileInput.addEventListener('dragleave', () => drop.classList.remove('over'));
  fileInput.addEventListener('drop', () => drop.classList.remove('over'));
  window.addEventListener('dragover', e => { if (e.target !== fileInput) e.preventDefault(); });
  window.addEventListener('drop', e => {
    if (e.target === fileInput) return;
    e.preventDefault();
    const f = e.dataTransfer?.files[0];
    if (!f) return;
    if (/\.(xlsx|xls|csv)$/i.test(f.name) && !$('#homeList').contains(e.target)) loadFile(f);
  });
  window.addEventListener('error', e => toast('오류: ' + (e.message || '알 수 없는 오류'), true));
  window.addEventListener('unhandledrejection', e => toast('오류: ' + (e.reason?.message || e.reason), true));

  // 집 주소
  $('#homeList').addEventListener('change', onHomeInput);
  $('#homeList').addEventListener('click', e => { const b = e.target.closest('[data-pick]'); if (b) openPick(b.closest('.home-row').dataset.user); });
  $('#btnHomeImport').addEventListener('click', () => $('#homeFile').click());
  $('#homeFile').addEventListener('change', e => { if (e.target.files[0]) importHomes(e.target.files[0]).catch(err => toast(err.message, true)); e.target.value = ''; });
  $('#btnHomeExport').addEventListener('click', exportHomes);

  // 분석 기준 (분석 후에는 좌표 재조회 없이 즉시 재계산)
  $('#radius').addEventListener('input', () => { syncRadius(); refresh(); });
  $('#radius').addEventListener('change', saveSettings);
  document.querySelectorAll('.chip').forEach(c => c.addEventListener('click', () => { $('#radius').value = c.dataset.r; syncRadius(); saveSettings(); refresh(); }));
  ['excludeOnline', 'nightStart', 'nightEnd', 'holidayText'].forEach(id => $('#' + id).addEventListener('change', () => { saveSettings(); if (id === 'excludeOnline' && state.analyzed) runAnalysis(); else refresh(); }));
  $('#btnRun').addEventListener('click', runAnalysis);

  // 결과 영역
  $('#mapUser').addEventListener('change', e => { state.view.mapUser = e.target.value; refresh(); });
  $('#fUser').addEventListener('change', e => selectUser(e.target.value));
  $('#fView').addEventListener('change', e => { state.view.fView = e.target.value; renderTxTable(); });
  $('#fDay').addEventListener('change', e => setDayFilter(e.target.value));
  $('#dayTable').addEventListener('click', e => {
    const tr = e.target.closest('tr[data-day]');
    if (tr) setDayFilter(state.view.fDay === tr.dataset.day ? 'all' : tr.dataset.day);
  });
  $('#fQuery').addEventListener('input', e => { state.view.q = e.target.value; renderTxTable(); });
  $('#btnExport').addEventListener('click', exportExcel);
  $('#userTable').addEventListener('click', e => { const tr = e.target.closest('tr[data-user]'); if (tr) selectUser(state.view.fUser === tr.dataset.user ? '' : tr.dataset.user); });
  $('#txTable').addEventListener('click', e => {
    const th = e.target.closest('th[data-key]');
    if (th) {
      const k = th.dataset.key;
      state.view.sortDir = state.view.sortKey === k ? -state.view.sortDir : (k === 'dist' ? 1 : -1);
      state.view.sortKey = k;
      return renderTxTable();
    }
    const tr = e.target.closest('tr[data-id]');
    if (tr) focusTxn(state.txns[+tr.dataset.id]);
  });

  // 거리 차트 툴팁
  const tip = document.createElement('div');
  tip.className = 'tip'; tip.hidden = true; document.body.appendChild(tip);
  $('#distChart').addEventListener('mousemove', e => {
    const row = e.target.closest('.hbar-row');
    if (!row) { tip.hidden = true; return; }
    const b = JSON.parse($('#distChart').dataset.buckets)[+row.dataset.i];
    tip.innerHTML = `${esc(b.l)}<br><b>${fmt(b.n)}건</b> · ${won(b.amt)}`;
    tip.hidden = false;
    tip.style.left = e.clientX + 12 + 'px';
    tip.style.top = e.clientY + 12 + 'px';
  });
  $('#distChart').addEventListener('mouseleave', () => { tip.hidden = true; });

  // 설정 대화상자
  $('#originText').textContent = location.origin;
  $('#btnSettings').addEventListener('click', () => {
    $('#kakaoKey').value = store.get('kakaoKey', '');
    $('#settingsDlg').showModal();
  });
  $('#settingsDlg').addEventListener('close', () => {
    if ($('#settingsDlg').returnValue !== 'save') return;
    const key = $('#kakaoKey').value.trim();
    if (key === store.get('kakaoKey', '')) return;
    store.set('kakaoKey', key);
    if (window.kakao?.maps) { toast('키가 바뀌었습니다. 새로고침 후 적용됩니다.'); setTimeout(() => location.reload(), 1200); return; }
    updateKeyStatus();
    if (key) loadKakao().then(() => { updateKeyStatus(true); toast('카카오 지도 연결 완료'); }).catch(e => { updateKeyStatus(false); toast(e.message, true); });
  });
  $('#btnClearCache').addEventListener('click', () => { Geo.clear(); toast('주소 좌표 캐시를 비웠습니다.'); });

  // 지도 지정 대화상자
  $('#btnPickSearch').addEventListener('click', pickSearch);
  $('#pickQuery').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); pickSearch(); } });
  $('#btnPickCancel').addEventListener('click', () => $('#pickDlg').close());
  $('#btnPickSave').addEventListener('click', () => {
    const u = Pick.user;
    state.homes[u] = { addr: state.homes[u]?.addr || '(지도에서 지정)', ...Pick.pos, src: 'manual' };
    store.set('homes', state.homes);
    $('#pickDlg').close();
    renderHomes();
    refresh();
  });

  if (location.protocol === 'file:') $('#fileWarn').hidden = false;
  updateKeyStatus();
  if (store.get('kakaoKey', '')) loadKakao().then(() => updateKeyStatus(true)).catch(() => updateKeyStatus(false));
}

init();
