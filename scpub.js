/* 강CRM 공시 수집기 (V305)
 * 스마트초이스 '단말기 지원금 조회' 화면에서 북마크(CRM공시)를 누르면 실행돼요.
 * - 사이트 화면과 똑같은 조회를 사람 속도로 하나씩 해요 (대기열(넷퍼넬)도 화면처럼 그대로 거침)
 * - 기종마다 한 번 조회 → 요금 구간 전체(~3.9만 ~ ~13만)의 SKT · KT · LG U+ 요금제 · 출고가 · 선택약정 · 공통지원금(기변 · 번이) · 공시일
 * - 다 모으면 강CRM을 열어서 넘겨줘요 (CRM에서 확인을 눌러야 저장)
 */
(function () {
  'use strict';
  if (window.__crmPubRun) { alert('공시 수집이 이미 돌고 있어요.'); return; }
  const CRM = String(window.__crm || 'https://kanggun1983.github.io/my-scanner/db.html');
  const MONTHS = 24, GAP = 1200, TIERS = ['all'];   // 한 번 조회하면 요금 구간 전체가 나옴
  const CARS = ['SKT', 'KT', 'LGU'];
  if (!/smartchoice\.or\.kr$/.test(location.hostname)) { alert('스마트초이스 화면에서 눌러 주세요.\n(단말기 지원금 조회 → 단말기 기준으로 조회)'); return; }
  window.__crmPubRun = true;
  const isM = /type=m/.test(location.search) || /^m\./.test(location.hostname);
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const num = (s) => Number(String(s || '').replace(/[^\d]/g, '')) || 0;
  const clean = (s) => String(s || '').replace(/\s+/g, ' ').trim();

  // ---------- 화면 (위에 뜨는 작은 상자)
  const box = document.createElement('div');
  box.style.cssText = 'position:fixed;left:8px;right:8px;top:8px;z-index:2147483647;background:#111827;color:#fff;border-radius:16px;padding:14px 16px;font:14px/1.5 sans-serif;box-shadow:0 10px 30px rgba(0,0,0,.4)';
  document.body.appendChild(box);
  let stop = false;
  const ui = (html) => { box.innerHTML = `<div style="font-weight:800;font-size:15px;margin-bottom:6px">📊 강CRM 공시 가져오기</div>${html}`; };
  const btn = (id, label, bg) => `<button data-x="${id}" style="margin:8px 6px 0 0;padding:10px 14px;border:0;border-radius:12px;font-weight:800;font-size:14px;background:${bg};color:#fff">${label}</button>`;
  box.addEventListener('click', (e) => { const b = e.target.closest('[data-x]'); if (!b) return; const x = b.dataset.x; (handlers[x] || (() => {}))(); });
  const handlers = { close: () => { stop = true; box.remove(); window.__crmPubRun = false; }, stop: () => { stop = true; } };

  // ---------- 사이트 요청 (사이트가 쓰는 주소 그대로)
  async function post(path, data) {
    const body = new URLSearchParams(); Object.entries(data).forEach(([k, v]) => body.append(k, v == null ? '' : v));
    const r = await fetch(location.origin + path, { method: 'POST', credentials: 'include', headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }, body });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    return r.text();
  }
  function nf(fn) {   // 넷퍼넬: 화면의 검색 버튼과 같은 순서 (시작 → 조회 → 종료). 없으면 그냥 조회
    const key = { projectKey: 'service_1', segmentKey: 'segKey_6704' };
    if (typeof window.nfStart !== 'function') return fn();
    return new Promise((ok, no) => {
      let done = false;
      const t = setTimeout(() => { if (!done) { done = true; no(new Error('대기열 응답 없음')); } }, 120000);
      window.nfStart(key, (res) => {
        if (done) return; done = true; clearTimeout(t);
        if (res && (res.status === 'Block')) return no(new Error('사이트가 잠시 막았어요 (대기열)'));
        Promise.resolve().then(fn).then((v) => { try { window.nfStop && window.nfStop(key, () => {}); } catch (e) {} ok(v); }, (e) => { try { window.nfStop && window.nfStop(key, () => {}); } catch (x) {} no(e); });
      });
    });
  }

  // ---------- 결과 화면 읽기 (글자 기준: '출고가' · '선택약정할인' · '기기변경' · '번호이동' · '공시일' 줄)
  function parse(html) {   // 구간마다: [구간] [요금제 월 39,000원 ×3] → 선택약정 · 기기변경 · 번호이동 · 공시일 (웹: 전 구간 한 화면 · 모바일: 고른 구간만)
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const rows = [...doc.querySelectorAll('tr')].map(tr => [...tr.children].map(td => clean(td.textContent)));
    const last3 = (cs) => cs.slice(-3), isWon = (c) => /월\s*[\d,]+\s*원/.test(c);
    let price = ['', '', ''];
    const pi = rows.findIndex(cs => cs.some(c => /^출고가$/.test(c)));
    if (pi >= 0) { let v = last3(rows[pi].filter(c => !/^출고가$/.test(c))); if (!v.some(c => /\d/.test(c)) && rows[pi + 1]) v = last3(rows[pi + 1]); price = v; }
    const out = {}; let cur = null;
    const flush = () => { if (!cur || !cur.chg) return; CARS.forEach((ck, j) => { const c = num(cur.chg[j]), m = num((cur.mnp || [])[j]); if (!c && !m) return;
      (out[ck] = out[ck] || []).push([cur.name[j].slice(0, 30), cur.pp[j], num(price[j]), num((cur.sel || [])[j]), c, m, (String((cur.date || [])[j] || '').match(/\d{4}-\d{2}-\d{2}/) || [''])[0]]); }); };
    rows.forEach((cs, i) => {
      if (cs.filter(isWon).length >= 2) { flush(); const v = last3(cs.filter(isWon)), prev = rows[i - 1] ? last3(rows[i - 1]) : ['', '', ''];
        cur = { pp: v.map(c => num((c.match(/월\s*([\d,]+)\s*원/) || [])[1])), name: v.map((c, j) => { const n = clean(c.replace(/월\s*[\d,]+\s*원.*/, '')); return n || clean(prev[j] || ''); }) }; return; }
      if (!cur) return;
      const lab = cs.find(c => /^(선택약정할인|기기\s*변경|번호\s*이동|공시일)/.test(c)); if (!lab) return;
      const v = last3(cs.filter(c => c !== lab && !/^공통\s*지원금$/.test(c))); if (v.length < 3) return;
      if (/^선택약정/.test(lab)) cur.sel = v; else if (/^기기/.test(lab)) cur.chg = v; else if (/^번호/.test(lab)) cur.mnp = v; else cur.date = v;
    });
    flush();
    return Object.keys(out).length ? out : null;
  }

  // ---------- 기종 목록
  function recent(rel) { const m = String(rel || '').match(/(\d{4})\D+(\d{1,2})/); if (!m) return false; const d = new Date(+m[1], +m[2] - 1, 1), lim = new Date(); lim.setMonth(lim.getMonth() - MONTHS); return d >= lim; }
  async function models() {
    const all = [];
    for (const mau of ['삼성전자', '애플']) {
      const t = await post('/smc/mobile/getProductList.do', { dan_Mau: mau, p_Group: 'PHONE', dan_Service: '5G' });
      let list = []; try { list = JSON.parse(t); } catch (e) { list = []; }
      list.forEach(x => { if (!x || /\(New\)/i.test(x.dan_Pc_Name || '') || !recent(x.used_Phone_Release)) return;
        const pick = [['SK', x.skt_Pn], ['KT', x.kt_Pn], ['LG', x.lg_Pn]].filter(([, p]) => p && String(p).trim());
        const best = pick.find(([, p]) => !/^-/.test(p)) || pick[0]; if (!best) return;
        all.push({ mau, name: clean(x.dan_Pc_Name), pn: best[0] + best[1], code: String(x.skt_Pn || x.kt_Pn || x.lg_Pn || '').replace(/^-/, ''), rel: x.used_Phone_Release || '' });
      });
    }
    return all;
  }

  function formBase() {   // 화면의 검색 폼 칸 그대로 (없으면 기본값)
    const f = document.forms.dantongForm, o = {};
    if (f) new FormData(f).forEach((v, k) => { o[k] = v; });
    return Object.assign({ datType: '', p_Group: 'PHONE', planLteChoice: 'all', plan3GChoice: 'all' }, o);
  }

  let lastFail = '';
  async function run() {
    ui('기종 목록을 불러오는 중…' + btn('close', '닫기', '#4b5563'));
    let list;
    try { list = await models(); } catch (e) { ui(`기종 목록을 못 불러왔어요 (${e.message}).<br>단말기 지원금 조회 화면에서 다시 눌러 주세요.` + btn('close', '닫기', '#4b5563')); return; }
    if (!list.length) { ui('최근 기종을 찾지 못했어요.' + btn('close', '닫기', '#4b5563')); return; }
    const min = Math.ceil(list.length * TIERS.length * (GAP + 700) / 60000);
    ui(`최근 ${MONTHS}개월 기종 <b>${list.length}개</b> (삼성 · 애플)<br>기종마다 한 번씩 조회해요 (요금 구간 전체). 약 ${min}분 걸려요.<br><span style="color:#9ca3af;font-size:12px">조회하는 동안 이 화면을 켜 두세요.</span><br>` + btn('go', '▶ 시작', '#2563eb') + btn('close', '취소', '#4b5563'));
    await new Promise(r => { handlers.go = r; });
    const base = formBase(), items = [];
    let n = 0, fail = 0;
    for (const md of list) {
      if (stop) break;
      const rec = { n: md.name, c: md.code, r: md.rel, d: {} };
      for (const tier of TIERS) {
        if (stop) break;
        n++;
        ui(`조회 중 ${n} / ${list.length * TIERS.length}<br><b>${md.name}</b><div style="height:6px;background:#374151;border-radius:3px;margin-top:8px"><div style="height:6px;width:${Math.round(n / (list.length * TIERS.length) * 100)}%;background:#60a5fa;border-radius:3px"></div></div>` + btn('stop', '■ 여기까지만', '#b91c1c'));
        try {
          const html = await nf(() => post('/smc/mobile/dantongList.do', Object.assign({}, base, { dan_Service: '5G', dan_Mau: md.mau, productName: md.pn, dan_Pc_Name: md.name, plan5GChoice: tier })));
          const p = parse(html);
          if (p) CARS.forEach(ck => { if (p[ck]) rec.d[ck] = (rec.d[ck] || []).concat(p[ck]); });
          else { fail++; lastFail = html; }
        } catch (e) {
          fail++; if (/막았/.test(e.message)) { stop = true; ui(e.message + btn('close', '닫기', '#4b5563')); }
        }
        if (fail >= 6 && !items.length) { stop = true; break; }
        await sleep(GAP);
      }
      if (Object.keys(rec.d).length) items.push(rec);
    }
    if (!items.length) {
      ui('결과 화면을 읽지 못했어요. 아래 [결과 화면 저장]으로 파일을 받아 보내 주시면 고칠게요.' + btn('dl', '결과 화면 저장', '#2563eb') + btn('close', '닫기', '#4b5563'));
      handlers.dl = () => { const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([lastFail || document.documentElement.outerHTML], { type: 'text/html' })); a.download = 'smartchoice-result.html'; document.body.appendChild(a); a.click(); a.remove(); };
      return;
    }
    const pay = { v: 1, at: new Date().toISOString(), src: location.hostname, items };
    const url = CRM + '#scpub=' + encodeURIComponent(JSON.stringify(pay));
    ui(`✅ ${items.length}개 기종 공시를 모았어요${fail ? ` <span style="color:#fca5a5">(못 읽은 조회 ${fail}개)</span>` : ''}.<br>강CRM으로 넘겨서 확인 후 저장해요.` + btn('send', '강CRM으로 보내기 ▶', '#16a34a') + btn('close', '닫기', '#4b5563'));
    handlers.send = () => { location.href = url; };
  }
  run().catch(e => { ui('오류: ' + (e && e.message || e) + btn('close', '닫기', '#4b5563')); });
})();
