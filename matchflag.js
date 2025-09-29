/* matchflag.js — GitHub Pages / 브라우저 내 처리 (SheetJS 필요)
   구현범위 v1: A검사 완전 동작
   - '일위대가' 시트 각 행의 수식에서 외부시트 참조 추출
   - 현재 행 (품명|규격) 키 vs 참조된 시트의 같은 행 키 비교
   - 일치/불일치 상세와 요약 생성 → XLSX.writeFile 로 다운로드
   확장 훅: runChecks() 내부에 B~E를 추가 구현 가능
*/
(function () {
  'use strict';

  // ====== UI 로그 ======
  const log = (m) => {
    const el = document.getElementById('mfLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(m);
  };

  // 전역 에러 표시
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

  // ====== SheetJS 로딩 보장 ======
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

  // ====== 파일 읽기 / 시트 변환 ======
  async function readWorkbook(file) {
    const buf = await file.arrayBuffer();
    return XLSX.read(new Uint8Array(buf), {
      type: 'array',
      cellFormula: true,
      cellNF: true,
      cellText: true,
    });
  }
  function sheetToAOA(wb, name) {
    const ws = wb.Sheets[name];
    if (!ws) throw new Error(`시트 없음: ${name}`);
    return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  }

  // ====== 문자열 정규화 & 숫자 ======
  const norm = (s) =>
    String(s == null ? '' : s)
      .replace(/\s+/g, '')
      .replace(/[()（）\[\]【】]/g, '')
      .trim();
  function toNum(x) {
    if (x == null || x === '') return null;
    const n = Number(String(x).replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  // ====== 헤더 탐색 (상단 몇 행 스캔해 품명/규격 등 위치 잡기) ======
  // 라벨 후보는 너의 NB2와 동일 개념으로 구성
  const LABELS = {
    품명: ['품명', '품 명'],
    규격: ['규격', '규 격', '사양', '사 양'],
    단위: ['단위'],
  };
  function findHeaderRowAndCols(arr, scanRows = 8) {
    const labelSet = {};
    for (const k in LABELS) labelSet[k] = LABELS[k].map((x) => norm(x));
    const rows = Math.min(scanRows, arr.length);
    const found = {}; // lab -> positions
    Object.keys(labelSet).forEach((k) => (found[k] = []));
    for (let r = 0; r < rows; r++) {
      const row = arr[r] || [];
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (typeof v !== 'string') continue;
        const key = norm(v);
        for (const lab in labelSet) {
          if (labelSet[lab].includes(key)) found[lab].push([r, c]);
        }
      }
    }
    // 헤더행: 가장 많이 히트한 r
    const rc = new Map();
    for (const hits of Object.values(found)) {
      for (const [r] of hits) rc.set(r, (rc.get(r) || 0) + 1);
    }
    if (!rc.size) throw new Error('머리글을 찾지 못함');
    const headerRow = [...rc.entries()].sort((a, b) => b[1] - a[1])[0][0];

    function pickCol(lab) {
      const hits = found[lab].filter(([r]) => r === headerRow);
      return hits.length ? hits[0][1] : null;
    }
    const colMap = {
      품명: pickCol('품명'),
      규격: pickCol('규격'),
      단위: pickCol('단위'),
    };
    if (colMap.품명 == null || colMap.규격 == null)
      throw new Error('필수 컬럼(품명/규격) 미탐색');
    return { headerRow, colMap };
  }

  // ====== (품명|규격) 키 생성 ======
  function buildKey(name, spec) {
    return norm(name) + '|' + norm(spec);
  }

  // ====== 행 → 키 맵 생성 ======
  function buildKeyMap(arr, headerRow, colMap) {
    const map = new Map(); // key -> {r, name, spec}
    for (let r = headerRow + 1; r < arr.length; r++) {
      const row = arr[r] || [];
      const name = row[colMap.품명];
      if (name == null || String(name).trim() === '') continue;
      const spec = row[colMap.규격] == null ? '' : row[colMap.규격];
      const key = buildKey(name, spec);
      if (!map.has(key)) map.set(key, { r, name, spec });
    }
    return map;
  }

  // ====== A1 표기 → 좌표 (A→0, 1-based row→0-based) ======
  const A1_RE = /^(\$?)([A-Z]+)(\$?)(\d+)$/;
  function a1ToRC(a1) {
    const m = String(a1).match(A1_RE);
    if (!m) return null;
    const colLetters = m[2];
    const row1 = parseInt(m[4], 10);
    let col = 0;
    for (let i = 0; i < colLetters.length; i++) {
      col = col * 26 + (colLetters.charCodeAt(i) - 64);
    }
    return { r: row1 - 1, c: col - 1 };
  }

  // ====== 수식에서 외부시트 참조 추출 ======
  // '시트 명'!A1  또는  시트명!$B$12  등
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

  // ====== 워크북 시트 → AOA 캐시 ======
  function getAOA(wb, name, cache) {
    if (cache.has(name)) return cache.get(name);
    const arr = sheetToAOA(wb, name);
    cache.set(name, arr);
    return arr;
  }

  // ====== A검사 구현 ======
  // 대상 시트: '일위대가' (없으면 가장 유사한 후보를 자동 탐색)
  function pickSheetByName(names, target) {
    if (names.includes(target)) return target;
    // 유사 후보: 공백 제거 후 includes
    const t = norm(target);
    let best = null;
    for (const n of names) {
      if (norm(n).includes(t)) {
        best = n; break;
      }
    }
    return best || target; // 없으면 그냥 target (나중에 시트없음 에러)
  }

  function isSummaryRow(rowStr) {
    if (!rowStr) return false;
    const s = String(rowStr);
    return /합계|TOTAL|소계|%/.test(s);
  }

  function runCheckA(wb) {
    const details = []; // 결과 상세 행들
    const cache = new Map(); // 시트 AOA 캐시

    const srcName = pickSheetByName(wb.SheetNames, '일위대가');
    const srcArr = sheetToAOA(wb, srcName);
    const { headerRow: srcHdr, colMap: srcCols } = findHeaderRowAndCols(srcArr);
    const srcKeyMap = buildKeyMap(srcArr, srcHdr, srcCols);

    // 모든 셀 순회: 같은 행(r)에서 수식이 있는 셀을 스캔
    const ws = wb.Sheets[srcName];
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let r = srcHdr + 1; r <= range.e.r; r++) {
      const name = srcArr[r]?.[srcCols.품명];
      if (!name || isSummaryRow(name)) continue;
      const spec = srcArr[r]?.[srcCols.규격] ?? '';
      const myKey = buildKey(name, spec);

      // 행 r의 모든 열에 대해 수식 셀 찾기
      const refs = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell || !cell.f) continue;
        const ex = extractRefs(cell.f);
        for (const ref of ex) refs.push({ ...ref, addr, formula: cell.f });
      }
      if (refs.length === 0) {
        // 외부참조가 전혀 없으면 패스(논쟁지점: 불일치로 볼지 보류)
        continue;
      }

      // 대표 참조(최빈 시트) 계산
      const bySheet = new Map();
      for (const rf of refs) bySheet.set(rf.sheet, (bySheet.get(rf.sheet) || 0) + 1);
      const rep = [...bySheet.entries()].sort((a, b) => b[1] - a[1])[0];
      const repSheet = rep ? rep[0] : refs[0].sheet;

      // 대표 시트에서 같은 행 키 가져와 비교
      let refKey = '(N/A)';
      let refName = '', refSpec = '';
      let ok = false, reason = '';

      try {
        const refArr = getAOA(wb, repSheet, cache);
        const { headerRow: refHdr, colMap: refCols } = findHeaderRowAndCols(refArr);
        const refNameVal = refArr[r]?.[refCols.품명];
        const refSpecVal = refArr[r]?.[refCols.규격] ?? '';
        refName = refNameVal ?? '';
        refSpec = refSpecVal ?? '';
        refKey = buildKey(refName, refSpec);

        if (norm(refNameVal) === '' && norm(refSpecVal) === '') {
          ok = false; reason = '대표시트 동일행에 품명/규격 빈값';
        } else {
          ok = myKey === refKey;
          reason = ok ? '키일치' : '키불일치';
        }
      } catch (e) {
        ok = false; reason = '대표시트 해석 실패: ' + (e.message || e);
      }

      details.push({
        검사: 'A',
        시트: srcName,
        행번호: r + 1, // 1-based
        내_품명: name,
        내_규격: spec,
        대표시트: repSheet,
        대표_품명: refName,
        대표_규격: refSpec,
        키_일치: ok ? 'TRUE' : 'FALSE',
        사유: reason,
      });
    }

    // 요약
    const total = details.length;
    const pass = details.filter((d) => d.키_일치 === 'TRUE').length;
    const fail = total - pass;

    return { details, summary: { 항목: 'A', 총건수: total, 일치: pass, 불일치: fail } };
  }

  // ====== 결과 통합 → 엑셀 저장 ======
  function objectsToAOA(objs) {
    if (!objs.length) return [['결과 없음']];
    const headers = Object.keys(objs[0]);
    const aoa = [headers];
    for (const o of objs) aoa.push(headers.map((h) => o[h]));
    return aoa;
  }
  function saveResultToExcel(srcFileName, results) {
    const wbOut = XLSX.utils.book_new();

    // Summary
    const sumAoa = objectsToAOA(results.summaries);
    XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet(sumAoa), '요약');

    // Each detail sheet
    for (const [name, rows] of Object.entries(results.detailByName)) {
      const aoa = objectsToAOA(rows);
      XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet(aoa), name);
    }

    const outName = srcFileName.replace(/\.[^.]+$/, '') + '_Matchflag_A.xlsx';
    XLSX.writeFile(wbOut, outName);
    log(`저장 완료: ${outName}`);
  }

  // ====== 실행 오케스트레이션 ======
  function runChecks(wb) {
    // A 검사 실행
    const A = runCheckA(wb);

    // 나중 확장(B~E) 자리
    const summaries = [A.summary];
    const detailByName = {
      'A_검사': A.details,
      // 'B_검사': B.details, ...
    };
    return { summaries, detailByName };
  }

  // ====== 버튼 핸들러 ======
  async function run() {
    try {
      document.getElementById('mfLog').textContent = '';
      log('초기화…');
      await waitForXLSX();

      const f = document.getElementById('mfFile').files[0];
      if (!f) throw new Error('엑셀 파일을 선택하세요.');

      log('파일 로딩 중…');
      const wb = await readWorkbook(f);
      log('시트: ' + wb.SheetNames.join(', '));

      const results = runChecks(wb);
      log('검사 완료. 엑셀로 저장합니다…');
      saveResultToExcel(f.name, results);
      log('완료.');
    } catch (e) {
      console.error(e);
      log(`ERROR: ${e.message || e}\n${e.stack || ''}`);
    }
  }

  // ====== DOM 바인딩 ======
  window.addEventListener('DOMContentLoaded', () => {
    const btn = document.getElementById('mfRun');
    if (btn) btn.addEventListener('click', run);
  });
})();
