/* matchflag.js v1.1
   - 상단 Nb2 이동 버튼은 HTML에서 처리
   - A검사 결과를 로그에 상세 출력
   - 엑셀 3종 출력: 전체 / 불일치 / 일치
*/
(function () {
  'use strict';

  // ===== 로그 유틸 =====
  const LOG_MISMATCH_LIMIT = 200; // 로그에 찍을 최대 불일치 행수 (UI 보호)
  const log = (m) => {
    const el = document.getElementById('mfLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(m);
  };
  window.onerror = function (msg, url, line, col, error) {
    log([
      '=== 전역 에러 감지 ===',
      '메시지: ' + msg,
      '파일: ' + url,
      '줄/컬럼: ' + line + ':' + col,
      '오브젝트: ' + (error && error.stack ? error.stack : error),
    ].join('\n'));
    return false;
  };

  // ===== XLSX 준비 =====
  function waitForXLSX(timeoutMs = 5000) {
    return new Promise((resolve, reject) => {
      const t0 = Date.now();
      (function loop() {
        if (window.XLSX) return resolve();
        if (Date.now() - t0 > timeoutMs) return reject(new Error('XLSX not loaded'));
        setTimeout(loop, 100);
      })();
    });
  }
  async function readWorkbook(file) {
    const buf = await file.arrayBuffer();
    return XLSX.read(new Uint8Array(buf), { type: 'array', cellFormula: true, cellNF: true, cellText: true });
  }
  function sheetToAOA(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error(`시트 없음: ${name}`);
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  // ===== 문자열/헤더 =====
  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, '').replace(/[()（）\[\]【】]/g, '').trim();
  const LABELS = {
    품명: ['품명', '품 명'],
    규격: ['규격', '규 격', '사양', '사 양'],
    단위: ['단위'],
  };
  function findHeaderRowAndCols(arr, scanRows = 8) {
    const labelSet = {};
    for (const k in LABELS) labelSet[k] = LABELS[k].map((x) => norm(x));
    const rows = Math.min(scanRows, arr.length);
    const found = {}; Object.keys(labelSet).forEach(k => found[k] = []);
    for (let r = 0; r < rows; r++) {
      const row = arr[r] || [];
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (typeof v !== 'string') continue;
        const key = norm(v);
        for (const lab in labelSet) if (labelSet[lab].includes(key)) found[lab].push([r, c]);
      }
    }
    const rc = new Map();
    for (const hits of Object.values(found)) for (const [r] of hits) rc.set(r, (rc.get(r) || 0) + 1);
    if (!rc.size) throw new Error('머리글을 찾지 못함');
    const headerRow = [...rc.entries()].sort((a,b)=>b[1]-a[1])[0][0];
    function pickCol(lab){ const hits = found[lab].filter(([r])=>r===headerRow); return hits.length?hits[0][1]:null; }
    const colMap = { 품명: pickCol('품명'), 규격: pickCol('규격'), 단위: pickCol('단위') };
    if (colMap.품명 == null || colMap.규격 == null) throw new Error('필수 컬럼(품명/규격) 미탐색');
    return { headerRow: headerRow, colMap };
  }
  const buildKey = (name, spec) => norm(name) + '|' + norm(spec);
  function isSummaryRow(rowStr) {
    if (!rowStr) return false;
    const s = String(rowStr);
    return /합계|TOTAL|소계|%/.test(s);
  }

  // ===== 수식 참조 추출 =====
  const REF_RE = /(?:'([^']+)'|([^'!]+))!([$]?[A-Z]+[$]?\d+)/g;
  function extractRefs(formula) {
    const out = [];
    if (typeof formula !== 'string') return out;
    let m;
    while ((m = REF_RE.exec(formula)) !== null) {
      const sheet = m[1] || m[2];
      const a1 = m[3];
      out.push({ sheet, a1 });
    }
    return out;
  }
  function pickSheetByName(names, target) {
    if (names.includes(target)) return target;
    const t = norm(target);
    for (const n of names) if (norm(n).includes(t)) return n;
    return target;
  }

  // ===== A검사 =====
  function runCheckA(wb) {
    const details = [];
    const cache = new Map();
    const getAOA = (name) => { if (cache.has(name)) return cache.get(name); const arr = sheetToAOA(wb, name); cache.set(name, arr); return arr; };

    const srcName = pickSheetByName(wb.SheetNames, '일위대가');
    const srcArr = sheetToAOA(wb, srcName);
    const { headerRow: srcHdr, colMap: srcCols } = findHeaderRowAndCols(srcArr);
    const ws = wb.Sheets[srcName];
    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let r = srcHdr + 1; r <= range.e.r; r++) {
      const name = srcArr[r]?.[srcCols.품명];
      if (!name || isSummaryRow(name)) continue;
      const spec = srcArr[r]?.[srcCols.규격] ?? '';
      const myKey = buildKey(name, spec);

      // 행 r의 수식 수집
      const refs = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell || !cell.f) continue;
        const ex = extractRefs(cell.f);
        for (const ref of ex) refs.push({ ...ref, addr, formula: cell.f });
      }
      if (refs.length === 0) continue;

      // 대표 시트(최빈)
      const bySheet = new Map();
      for (const rf of refs) bySheet.set(rf.sheet, (bySheet.get(rf.sheet) || 0) + 1);
      const rep = [...bySheet.entries()].sort((a,b)=>b[1]-a[1])[0];
      const repSheet = rep ? rep[0] : refs[0].sheet;

      let refName = '', refSpec = '', ok = false, reason = '';
      try {
        const refArr = getAOA(repSheet);
        const { headerRow: refHdr, colMap: refCols } = findHeaderRowAndCols(refArr);
        const refNameVal = refArr[r]?.[refCols.품명];
        const refSpecVal = refArr[r]?.[refCols.규격] ?? '';
        refName = refNameVal ?? '';
        refSpec = refSpecVal ?? '';
        if (norm(refNameVal) === '' && norm(refSpecVal) === '') { ok = false; reason = '대표시트 동일행에 품명/규격 빈값'; }
        else { ok = (myKey === buildKey(refName, refSpec)); reason = ok ? '키일치' : '키불일치'; }
      } catch (e) {
        ok = false; reason = '대표시트 해석 실패: ' + (e.message || e);
      }

      details.push({
        검사: 'A',
        시트: srcName,
        행번호: r + 1,
        내_품명: name,
        내_규격: spec,
        대표시트: repSheet,
        대표_품명: refName,
        대표_규격: refSpec,
        키_일치: ok ? 'TRUE' : 'FALSE',
        사유: reason,
      });
    }

    const total = details.length;
    const pass = details.filter(d => d.키_일치 === 'TRUE').length;
    const fail = total - pass;
    return { details, summary: { 항목: 'A', 총건수: total, 일치: pass, 불일치: fail } };
  }

  // ===== 저장 유틸 =====
  function objectsToAOA(objs) {
    if (!objs.length) return [['결과 없음']];
    const headers = Object.keys(objs[0]);
    const aoa = [headers];
    for (const o of objs) aoa.push(headers.map(h => o[h]));
    return aoa;
  }

  function saveThreeFiles(baseName, allRows) {
    const passRows = allRows.filter(r => r.키_일치 === 'TRUE');
    const failRows = allRows.filter(r => r.키_일치 !== 'TRUE');

    // 전체
    {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(allRows)), 'A_전체');
      XLSX.writeFile(wb, baseName + '_A_전체.xlsx');
    }
    // 불일치
    {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(failRows)), 'A_불일치');
      XLSX.writeFile(wb, baseName + '_A_불일치.xlsx');
    }
    // 일치
    {
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(passRows)), 'A_일치');
      XLSX.writeFile(wb, baseName + '_A_일치.xlsx');
    }
  }

  // ===== 오케스트레이션 =====
  function runChecksAndSave(wb, srcFileName) {
    const A = runCheckA(wb);

    // 로그 출력 (요약 + 불일치 상세)
    log(`[A] 총건수=${A.summary.총건수}, 일치=${A.summary.일치}, 불일치=${A.summary.불일치}`);
    if (A.details.length === 0) {
      log('[A] 상세 없음');
    } else {
      const fails = A.details.filter(d => d.키_일치 !== 'TRUE');
      if (fails.length) {
        log(`[A] 불일치 상세 (${Math.min(fails.length, LOG_MISMATCH_LIMIT)}개 표시 / 총 ${fails.length})`);
        const sample = fails.slice(0, LOG_MISMATCH_LIMIT);
        for (const d of sample) {
          log(`- 행${d.행번호} [${d.사유}] 내=(${d.내_품명} | ${d.내_규격}) vs 대표(${d.대표시트})=(${d.대표_품명} | ${d.대표_규격})`);
        }
        if (fails.length > LOG_MISMATCH_LIMIT) log(`… 생략 ${fails.length - LOG_MISMATCH_LIMIT}건`);
      } else {
        log('[A] 불일치 없음 (전부 일치)');
      }
    }

    // 파일 3종 저장
    const base = srcFileName.replace(/\.[^.]+$/, '');
    saveThreeFiles(base, A.details);
    log('엑셀 저장 완료: _A_전체 / _A_불일치 / _A_일치');
  }

  // ===== 실행 =====
  async function run() {
    try {
      document.getElementById('mfLog').textContent = '';
      log('초기화…'); await waitForXLSX();

      const f = document.getElementById('mfFile').files[0];
      if (!f) throw new Error('엑셀 파일을 선택하세요.');

      log('파일 로딩 중…');
      const wb = await readWorkbook(f);
      log('시트: ' + wb.SheetNames.join(', '));

      runChecksAndSave(wb, f.name);
      log('완료.');
    } catch (e) {
      console.error(e); log(`ERROR: ${e.message || e}\n${e.stack || ''}`);
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    const btn = document.getElementById('mfRun');
    if (btn) btn.addEventListener('click', run);
  });
})();

