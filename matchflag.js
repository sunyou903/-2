/* matchflag.js v1.2 — A검사: 파이썬 로직과 동등화
   - 로그: 요약만 출력
   - 출력: 전체 / 불일치 / 일치 3종 XLSX
   - 핵심 로직:
     * '일위대가' 각 행의 수식에서 '단가대비표' 또는 '일위대가목록' 참조 추출
     * 참조된 셀의 '행번호(rr)'를 사용하여 해당 시트의 (품명|규격) 키를 가져와 현재 행 키와 비교
     * 참조가 전혀 없으면 '수량' 값 직입 여부 판단 → 불일치(단, % 포함 시 ‘제외’)
*/

(function () {
  'use strict';

  // ===== 로그 =====
  const log = (m) => {
    const el = document.getElementById('mfLog');
    if (el) el.textContent += (el.textContent ? '\n' : '') + String(m);
  };
  window.onerror = function (msg, url, line, col, error) {
    log(`ERROR: ${msg} @ ${url}:${line}:${col}`);
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

  // ===== 정규화 =====
  const normSimple = (s) => (s == null ? null : String(s).trim());
  function normCommasWs(s) {
    if (s == null) return '';
    let t = String(s).replace(/,/g, ' ');
    t = t.replace(/\s+/g, ' ').trim();
    return t;
  }
  function normKey(key) {
    if (key == null) return '';
    const s = String(key);
    if (s.includes('|')) {
      const [a, b] = s.split('|', 1 + 1);
      return `${normCommasWs(a)}|${normCommasWs(b)}`;
    }
    return normCommasWs(s);
  }
   // 0) 워크시트에서 "수식 있는 셀만" 행별로 모으기
   function buildFormulaMapByRow(ws, rowStart /* 0-based */) {
     const byRow = new Map();
     for (const addr in ws) {
       if (addr[0] === '!') continue;
       const cell = ws[addr];
       if (!cell || !cell.f || typeof cell.f !== 'string') continue;
       const rc = XLSX.utils.decode_cell(addr);
       if (rc.r < rowStart) continue;
       if (!byRow.has(rc.r)) byRow.set(rc.r, []);
       byRow.get(rc.r).push({ f: cell.f, addr });
     }
     return byRow;
   }
   
   // 1) 원본과 동일한 시트참조 정규식 (따옴표/비따옴표, 시트명, 열, 행)
   const SHEET_REF_RE = /'(.*?)'!\$?([A-Z]{1,3})\$?(\d+)|([^'!]+)!\$?([A-Z]{1,3})\$?(\d+)/g;
   
   // 2) 가장 가까운 헤더행(최근접 위쪽 r<=행)을 찾는 헬퍼
   function nearestHeaderRow(targetRow1Based, headerRowCandidates /* number[] 1-based */) {
     let best = null, bestDist = Infinity;
     for (const hr of headerRowCandidates) {
       if (hr <= targetRow1Based) {
         const d = targetRow1Based - hr;
         if (d < bestDist) { bestDist = d; best = hr; }
       }
     }
     return best; // 없으면 null
   }

  // ===== 헤더 탐색 (필수 라벨 모두 존재하는 행 찾기; 라벨은 공백 제거 비교) =====
  function normLabel(s) {
    if (s == null) return null;
    return String(s).replace(/[\u3000 ]/g, '').trim();
  }
  function findHeaderRowAndColsRequired(arr, required, scanRows = 40, scanCols = 200) {
    const req = new Set(required.map(normLabel));
    for (let r = 0; r < Math.min(scanRows, arr.length); r++) {
      const found = new Map();
      const row = arr[r] || [];
      for (let c = 0; c < Math.min(scanCols, row.length); c++) {
        const nv = normLabel(row[c]);
        for (const want of req) {
          if (nv === want && !found.has(want)) found.set(want, c);
        }
      }
      if (found.size === req.size) {
        const pos = {};
        // map back to original labels
        for (const lab of required) {
          const k = normLabel(lab);
          pos[lab] = found.get(k);
        }
        return { headerRow: r, pos };
      }
    }
    throw new Error(`헤더 라벨 탐색 실패: ${required.join(',')}`);
  }

  // ===== 키맵 (1-based 행번호 -> "품명|규격") =====
  function buildKeyMap(arr, headerRow, colName, colSpec) {
    const map = {};
    for (let r = headerRow + 1; r < arr.length; r++) {
      const name = normSimple(arr[r]?.[colName]);
      const spec = normSimple(arr[r]?.[colSpec]);
      if ((name == null || name === '') && (spec == null || spec === '')) continue;
      const rr = r + 1; // 1-based
      map[rr] = `${name ?? ''}|${spec ?? ''}`;
    }
    return map;
  }

  // ===== 합계/소계 등 제외 규칙 (원본 A와 동일) =====
  function isSumRow(nameCell) {
    if (!nameCell) return false;
    const t = String(nameCell);
    return t.includes('합') && t.includes('계') && t.includes('[') && t.includes(']');
  }
  // UI에 제어권 잠깐 넘겨서 '멈춘 것처럼' 보이지 않게 함
  function uiYield() { return new Promise(r => setTimeout(r, 0)); }

  // N번째마다 진행 로그 찍기
  function progressLog(prefix, i, total, step=200) {
    if (i % step === 0) log(`${prefix}... ${i}/${total}`);
  }


  // ===== 수식 참조 추출:  '단가대비표' / '일위대가목록' ! $A$123 =====
  const REF_RE = /(?:'?)((?:단가대비표)|(?:일위대가목록))(?:'?)!\$?([A-Z]{1,3})\$?(\d+)/g;


  // ===== 객체 배열→AOA =====
  function objectsToAOA(objs) {
    if (!objs.length) return [['결과 없음']];
    const headers = Object.keys(objs[0]);
    const aoa = [headers];
    for (const o of objs) aoa.push(headers.map(h => o[h]));
    return aoa;
  }


  function appendThreeSheets(wb, prefix, rows) {
     const passRows = rows.filter(r => r.일치여부 === '일치');
     const failRows = rows.filter(r => r.일치여부 === '불일치');
   
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(rows)), `${prefix}_전체`);
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(failRows)), `${prefix}_불일치`);
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(passRows)), `${prefix}_일치`);
   }

  function saveAllSectionsAsOneWorkbook(baseName, sections /* [{name:'A', rows:[...]}, ...] */) {
     const wb = XLSX.utils.book_new();
     for (const sec of sections) {
       appendThreeSheets(wb, `${sec.name}`, sec.rows || []);
     }
     const outName = `${baseName}_종합검사.xlsx`;
     XLSX.writeFile(wb, outName);
  }
 
  // ===== A검사 =====
  function runCheckA(wb) {
    const srcName = pickSheetByName(wb.SheetNames, '일위대가');
    const upName  = pickSheetByName(wb.SheetNames, '단가대비표');
    const lsName  = pickSheetByName(wb.SheetNames, '일위대가목록');

    const ulArr = sheetToAOA(wb, srcName);
    const upArr = sheetToAOA(wb, upName);
    const lsArr = sheetToAOA(wb, lsName);

    const ulHdr = findHeaderRowAndColsRequired(ulArr, ['품명','규격','단위','수량']);
    const upHdr = findHeaderRowAndColsRequired(upArr, ['품명','규격','단위']);
    const lsHdr = findHeaderRowAndColsRequired(lsArr, ['품명','규격']);

    const ulPos = ulHdr.pos, upPos = upHdr.pos, lsPos = lsHdr.pos;

    const ulKey = buildKeyMap(ulArr, ulHdr.headerRow, ulPos['품명'], ulPos['규격']);
    const upKey = buildKeyMap(upArr, upHdr.headerRow, upPos['품명'], upPos['규격']);
    const lsKey = buildKeyMap(lsArr, lsHdr.headerRow, lsPos['품명'], lsPos['규격']);

    const ws = wb.Sheets[srcName];
    const rng = XLSX.utils.decode_range(ws['!ref']);

    const records = [];
    let checked = 0;

    for (let r = ulHdr.headerRow + 1; r <= rng.e.r; r++) {
      const r1 = r + 1; // 1-based
      const cur_key = ulKey[r1];
      const nameCell = ulArr[r]?.[ulPos['품명']];
      if (!cur_key || isSumRow(nameCell)) continue;

      const specCell = ulArr[r]?.[ulPos['규격']];
      const pname_cur = normSimple(nameCell);
      const gname_cur = normSimple(specCell);

      let rowHasRef = false;

      // --- 이 행의 모든 셀에서 수식 검사 ---
      for (let c = rng.s.c; c <= rng.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        const f = cell && typeof cell.f === 'string' ? cell.f : null;
        if (!f) continue;
        if (!f.includes('단가대비표') && !f.includes('일위대가목록')) continue;

        let m;
        REF_RE.lastIndex = 0;
        while ((m = REF_RE.exec(f)) !== null) {
          const sheet_name = m[1];
          const colLetters = m[2];
          const rr = parseInt(m[3], 10); // 참조된 '행번호'

          const ref_key = sheet_name.startsWith('단가대비표') ? upKey[rr] : lsKey[rr];
          if (!ref_key) continue;

          checked += 1;
          let status = (normKey(ref_key) === normKey(cur_key)) ? '일치' : '불일치';

          // % 포함 시 불일치 → 제외
          try {
            if (status === '불일치' && ((pname_cur && String(pname_cur).includes('%')) || (gname_cur && String(gname_cur).includes('%')))) {
              status = '제외';
            }
          } catch (_) {}

          const shortF = f.length > 140 ? f.slice(0, 140) + '...' : f;

          records.push({
            "일위대가_행": r1,
            "일위대가_품명|규격": cur_key,
            "참조시트": sheet_name,
            "참조셀": `${sheet_name}!${colLetters}${rr}`,
            "참조_품명|규격": ref_key,
            "수식_셀": addr,
            "수식_일부": shortF,
            "일치여부": status
          });
          rowHasRef = true;
        }
      }

      // --- 참조가 전혀 없으면: '수량' 값 직접입력 여부 검사 ---
      if (!rowHasRef) {
        const qtyCol = ulPos['수량'];
        const val = ulArr[r]?.[qtyCol];
        if (!(val == null || val === '' || val === 0)) {
          let status_di = '불일치';
          try {
            if ((pname_cur && String(pname_cur).includes('%')) || (gname_cur && String(gname_cur).includes('%'))) {
              status_di = '제외';
            }
          } catch (_) {}

          records.push({
            "일위대가_행": r1,
            "일위대가_품명|규격": cur_key,
            "참조시트": "",
            "참조셀": "",
            "참조_품명|규격": "",
            "수식_셀": XLSX.utils.encode_cell({ r, c: qtyCol }),
            "수식_일부": String(val),
            "일치여부": status_di
          });
        }
      }
    }

    const total = records.length;
    const ok = records.filter(x => x.일치여부 === '일치').length;
    const bad = records.filter(x => x.일치여부 === '불일치').length;

    return {
      summary: { "A_검사한_참조": checked, "A_일치": ok, "A_불일치": bad },
      details: records
    };
  }

  // ===== 시트명 선택 (정확/유사) =====
  function pickSheetByName(names, target) {
    if (names.includes(target)) return target;
    const t = String(target).replace(/[\u3000 ]/g, '');
    for (const n of names) {
      if (String(n).replace(/[\u3000 ]/g, '').includes(t)) return n;
    }
    return target; // 못 찾으면 그대로(나중에 시트 없음 에러)
  }

  // ===== 저장: 전체/불일치/일치 =====
  function saveAsOneWorkbook(baseName, allRows) {
     const passRows = allRows.filter(r => r.일치여부 === '일치');
     const failRows = allRows.filter(r => r.일치여부 === '불일치');

     const wb = XLSX.utils.book_new();
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(allRows)), 'A_전체');
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(failRows)), 'A_불일치');
     XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(objectsToAOA(passRows)), 'A_일치');
   
     const outName = `${baseName}_일위대가 검사.xlsx`;
     XLSX.writeFile(wb, outName);
} 
  // 공통: 시트 AOA 캐시
   function makeAOACache(wb) {
     const cache = new Map();
     const get = (name) => {
       if (cache.has(name)) return cache.get(name);
       const arr = sheetToAOA(wb, name);
       cache.set(name, arr);
       return arr;
     };
     return get;
   }
   
   // B) 일위대가목록 검사: 목록 각 행의 수식이 '단가대비표' (또는 일위대가) 특정 행을 참조 → 키 비교
   function runCheckB(wb) {
     const lsName = pickSheetByName(wb.SheetNames, '일위대가목록');
     const ulName = pickSheetByName(wb.SheetNames, '일위대가');
   
     const lsArr = sheetToAOA(wb, lsName);
     const ulArr = sheetToAOA(wb, ulName);
   
     // 목록/일위대가 헤더 포지션
     const { headerRow: lsHR, pos: lsPos } = findHeaderRowAndColsRequired(lsArr, ['품명','규격']);
     const { headerRow: ulHR, pos: ulPosFull } = findHeaderRowAndColsRequired(
       ulArr, ['품명','규격','단위','수량','합계 단가','합계금액','합계단가','합계단가(원)'].map(s=>s) // 라벨 변형 흡수
     );
   
     // 일위대가의 헤더 후보(원본은 “헤더행 후보 모음”에서 최근접 선택)
     const ulHeaderCandidates = [ulHR]; // 필요 시 추가 후보를 여기에 넣어 확장
   
     // 목록 키
     const keyOfList = (r) => `${normSimple(lsArr[r]?.[lsPos['품명']])}|${normSimple(lsArr[r]?.[lsPos['규격']])}`;
   
     // 일위대가: 특정 헤더행 기준으로 (품명|규격) 키 만들기
     function keyFromULRowByHeader(rr1, hdr1) {
       const r0 = rr1 - 1;
       const name = ulArr[r0]?.[ulPosFull['품명']];
       const spec = ulArr[r0]?.[ulPosFull['규격']];
       return `${normSimple(name)}|${normSimple(spec)}`;
     }
   
     const wsList = wb.Sheets[lsName];
     const rng = XLSX.utils.decode_range(wsList['!ref']);
     const formulasByRow = buildFormulaMapByRow(wsList, lsHR + 1);
   
     const records = [];
     let checked = 0;
   
     for (let r = lsHR + 1; r <= rng.e.r; r++) {
       const r1 = r + 1;
       const curKey = keyOfList(r);
       if (curKey === "null|null" || curKey === "undefined|undefined") continue;
   
       // 이 행의 수식들만 스캔
       const forms = formulasByRow.get(r) || [];
       let any = false;
       let mismatched = false;
   
       for (const { f } of forms) {
         let m; SHEET_REF_RE.lastIndex = 0;
         while ((m = SHEET_REF_RE.exec(f)) !== null) {
           const sheet = (m[1] || m[4] || '').trim().replace(/\s+/g,'');
           const rr = parseInt(m[3] || m[6], 10);
           if (!sheet.includes('일위대가')) continue; // B는 일위대가 참조만 사용
           any = true;
   
           // 최근접 헤더 선택
           const hdr1 = nearestHeaderRow(rr, ulHeaderCandidates);
           if (!hdr1) continue;
           const refKey = keyFromULRowByHeader(rr, hdr1) || '';
   
           checked++;
           const status = (normKey(refKey) === normKey(curKey)) ? '일치' : '불일치';
           if (status === '불일치') mismatched = true;
   
           records.push({
             "검사":"B",
             "시트": lsName,
             "행번호": r1,
             "내_품명|규격": curKey,
             "참조시트": ulName,
             "참조셀": `${ulName}!${(m[2]||m[5])}${rr}`,
             "참조_품명|규격": refKey,
             "일치여부": status
           });
         }
       }
   
       // 원본: “참조 전혀 없음”은 B에서 불일치로 치지 않음 (그대로 둠)
       // 필요 시 ‘금액 열 직접입력’ 추가 판정 가능(원본 주석과 동일한 처리)
     }
   
     const ok = records.filter(x=>x.일치여부==='일치').length;
     const bad = records.filter(x=>x.일치여부==='불일치').length;
     return { summary: { "B_참조셀": checked, "B_일치": ok, "B_불일치": bad }, details: records };
   }

   // C) 공종별내역서: 대표참조(최빈 시트) 기준 키 비교, 참조없고 '합계 단가' 직접입력되면 불일치
   function runCheckC(wb) {
     const wsName = pickSheetByName(wb.SheetNames, '공종별내역서');
     const arr = sheetToAOA(wb, wsName);
     const { headerRow: hr, pos } = findHeaderRowAndColsRequired(arr, ['품명','규격']);
     // 합계 단가 후보 라벨
     let sumCol = null;
     for (let c = 0; c < (arr[hr]||[]).length; c++) {
       const k = String(arr[hr]?.[c] ?? '').replace(/[\s\u3000]/g,'');
       if (/합계단가|합계단가\(원\)|총단가|합계금액/.test(k)) { sumCol = c; break; }
     }
     const ws = wb.Sheets[wsName];
     const rng = XLSX.utils.decode_range(ws['!ref']);
     const formulasByRow = buildFormulaMapByRow(ws, hr + 1);
   
     const records = [];
     let withDirectRefs = 0, directValMismatch = 0;
   
     const keyOf = (r) => `${normSimple(arr[r]?.[pos['품명']])}|${normSimple(arr[r]?.[pos['규격']])}`;
   
     for (let r = hr + 1; r <= rng.e.r; r++) {
       const r1 = r + 1;
       const curKey = keyOf(r);
       if (curKey === 'null|null' || curKey === 'undefined|undefined') continue;
   
       const forms = formulasByRow.get(r) || [];
       const refPairs = [];
       for (const { f } of forms) {
         let m; SHEET_REF_RE.lastIndex = 0;
         while ((m = SHEET_REF_RE.exec(f)) !== null) {
           const sheet = (m[1] || m[4] || '').trim().replace(/\s+/g,'');
           const rr = parseInt(m[3] || m[6], 10);
           if (!sheet || sheet === wsName) continue; // 자기 시트 제외
           refPairs.push([sheet, rr]);
         }
       }
   
       if (refPairs.length) {
         withDirectRefs++;
         // 대표 (시트,행) = 최빈
         const counter = new Map();
         for (const k of refPairs) counter.set(k.toString(), (counter.get(k.toString())||0)+1);
         const [repKeyStr] = [...counter.entries()].sort((a,b)=>b[1]-a[1])[0];
         const [repSheet, repRowStr] = repKeyStr.split(',');
         const repRow = parseInt(repRowStr, 10);
   
         // 대표 참조 시트의 (품명|규격)
         let refKey = '';
         try {
           const tarArr = sheetToAOA(wb, repSheet);
           const { headerRow: th, pos: tpos } = findHeaderRowAndColsRequired(tarArr, ['품명','규격']);
           refKey = `${normSimple(tarArr[repRow-1]?.[tpos['품명']])}|${normSimple(tarArr[repRow-1]?.[tpos['규격']])}`;
         } catch (_) {}
   
         const status = (normKey(refKey) === normKey(curKey)) ? '일치' : '불일치';
         records.push({
           "검사":"C","시트":wsName,"행번호":r1,
           "내_품명|규격":curKey,"대표참조_시트":repSheet,"대표참조_행":repRow,
           "대표참조_품명|규격":refKey,"일치여부":status
         });
       } else {
         // 참조 전혀 없고 합계단가 셀에 값이 있으면 불일치(값 직접입력)
         if (sumCol != null) {
           const val = arr[r]?.[sumCol];
           if (!(val == null || val === '' || val === 0)) {
             directValMismatch++;
             records.push({
               "검사":"C","시트":wsName,"행번호":r1,
               "내_품명|규격":curKey,"대표참조_시트":"","대표참조_행":"",
               "대표참조_품명|규격":"","일치여부":"불일치"
             });
           }
         }
       }
     }
   
     const ok = records.filter(x=>x.일치여부==='일치').length;
     const bad = records.filter(x=>x.일치여부==='불일치').length;
     return { summary: {
         "C_검사대상_행수(직접참조 보유)": withDirectRefs,
         "C_일치": ok,
         "C_불일치": bad,
         "C_값직접입력_불일치": directValMismatch
       }, details: records };
   }

   // D) 공종별집계표: 재/노/경 단가 수식 → 참조행의 '품명'만 비교(완화)
   function runCheckD(wb) {
     const sSum = pickSheetByName(wb.SheetNames, '공종별집계표');
     const arr = sheetToAOA(wb, sSum);
     const { headerRow: hr, pos } = findHeaderRowAndColsRequired(arr, ['품명']); // 품명만 필수
   
     // 재/노/경 단가 열 감지
     const targets = [];
     for (let c=0; c<(arr[hr]||[]).length; c++) {
       const k = String(arr[hr]?.[c] ?? '').replace(/[\s\u3000]/g,'');
       if (/재료비적용단가|노무비|경비적용단가|경비단가/.test(k)) targets.push([k,c]);
     }
     const ws = wb.Sheets[sSum];
     const rng = XLSX.utils.decode_range(ws['!ref']);
     const formulasByRow = buildFormulaMapByRow(ws, hr + 1);
   
     const records = [];
     let checked = 0;
   
     const nameOf = (r) => String(arr[r]?.[pos['품명']] ?? '').trim();
   
     for (let r = hr + 1; r <= rng.e.r; r++) {
       const r1 = r + 1;
       const curName = nameOf(r);
       if (!curName) continue;
   
       // 이 행의 대상 열들 중 수식에서 외부 참조 수집
       const refPairs = [];
       const forms = formulasByRow.get(r) || [];
       for (const { f } of forms) {
         let m; SHEET_REF_RE.lastIndex = 0;
         while ((m = SHEET_REF_RE.exec(f)) !== null) {
           const sheet = (m[1] || m[4] || '').trim().replace(/\s+/g,'');
           const rr = parseInt(m[3] || m[6], 10);
           if (!sheet || sheet === sSum) continue;
           refPairs.push([sheet, rr]); checked++;
         }
       }
   
       let repSheet=null, repRow=null, refName='';
       if (refPairs.length) {
         const counter = new Map();
         for (const k of refPairs) counter.set(k.toString(), (counter.get(k.toString())||0)+1);
         const [repKeyStr] = [...counter.entries()].sort((a,b)=>b[1]-a[1])[0];
         const [s, rrStr] = repKeyStr.split(','); repSheet = s; repRow = parseInt(rrStr,10);
   
         try {
           const tarArr = sheetToAOA(wb, repSheet);
           const { headerRow: th, pos: tpos } = findHeaderRowAndColsRequired(tarArr, ['품명']);
           refName = String(tarArr[repRow-1]?.[tpos['품명']] ?? '').trim();
         } catch (_) {}
       }
   
       const status = (normCommasWs(refName) === normCommasWs(curName)) ? '일치' : '불일치';
       records.push({
         "검사":"D","시트":sSum,"행번호":r1,
         "내_품명":curName,"대표참조_시트":repSheet,"대표참조_행":repRow,
         "참조_품명":refName,"일치여부":status
       });
     }
   
     const ok = records.filter(x=>x.일치여부==='일치').length;
     const bad = records.filter(x=>x.일치여부==='불일치').length;
     return { summary: { "D_참조셀": checked, "D_일치": ok, "D_불일치": bad }, details: records };
   }

   // E) 단가대비표: 재료비 적용단가/노무비 수식의 참조 행매칭 (장비 단가산출서 등을 참조해도 rr로 비교)
   function runCheckE(wb) {
     const sDv = pickSheetByName(wb.SheetNames, '단가대비표');
     const dvArr = sheetToAOA(wb, sDv);
     const { headerRow: hr, pos } = findHeaderRowAndColsRequired(dvArr, ['품명','규격']);
   
     const nameCol = pos['품명'], specCol = pos['규격'];
   
     // 대상 열 찾기
     const cCols = [];
     for (let c=0; c<(dvArr[hr]||[]).length; c++) {
       const k = String(dvArr[hr]?.[c] ?? '').replace(/[\s\u3000]/g,'');
       if (/재료비적용단가|노무비|경비적용단가/.test(k)) cCols.push([c,k]);
     }
   
     // 특정 시트면 (품명|사양) 우선
     function keyFromTargetSheet(arr, rr1) {
       // 시트별 헤더 조합 후보
       const candidates = [
         ['품명','사양'],
         ['품명','규격'],
         ['품평','사양'],
         ['품평','규격']
       ];
       for (const [h1,h2] of candidates) {
         try {
           const { headerRow: th, pos: p } = findHeaderRowAndColsRequired(arr, [h1, h2]);
           const name = normSimple(arr[rr1-1]?.[p[h1]]);
           const spec = normSimple(arr[rr1-1]?.[p[h2]]);
           if (name || spec) return { key: `${name}|${spec}`, used: `${h1}|${h2}` };
         } catch(_) {}
       }
       return { key: '', used: null };
     }
   
     const ws = wb.Sheets[sDv];
     const rng = XLSX.utils.decode_range(ws['!ref']);
     const formulasByRow = buildFormulaMapByRow(ws, hr + 1);
   
     const records = [];
   
     for (let r = hr + 1; r <= rng.e.r; r++) {
       const r1 = r + 1;
       const baseKey = `${normSimple(dvArr[r]?.[nameCol])}|${normSimple(dvArr[r]?.[specCol])}`;
       if (baseKey === 'null|null' || baseKey === 'undefined|undefined') continue;
   
       // 대상 열 중 수식이 있는 것만
       const forms = formulasByRow.get(r) || [];
       for (const { f } of forms) {
         // 외부 참조 수집(자기 시트 제외)
         let rr = null, sheetName = '', colA = null;
         let m; SHEET_REF_RE.lastIndex = 0;
         while ((m = SHEET_REF_RE.exec(f)) !== null) {
           const sheet = (m[1] || m[4] || '').trim();
           if (!sheet || sheet === sDv) continue;
           sheetName = sheet; rr = parseInt(m[3] || m[6], 10); colA = (m[2] || m[5]);
           break; // 대표 1개만
         }
         if (!rr) continue;
   
         // 대상 시트 로드
         let refKey = '', usedHdr = null;
         try {
           const tarArr = sheetToAOA(wb, sheetName);
           const { key, used } = keyFromTargetSheet(tarArr, rr);
           refKey = key; usedHdr = used;
         } catch(_) {}
   
         const status = refKey ? ((normKey(refKey) === normKey(baseKey)) ? '일치' : '불일치') : '불일치';
         records.push({
           "검사":"E","시트":sDv,"행번호":r1,
           "내_품명|규격": baseKey,
           "참조시트": sheetName,
           "참조셀": `${sheetName}!${colA||'A'}${rr}`,
           "참조_품명|규격": refKey,
           "참조키_사용헤더": usedHdr,
           "일치여부": status
         });
       }
     }
   
     const ok = records.filter(x=>x.일치여부==='일치').length;
     const bad = records.filter(x=>x.일치여부==='불일치').length;
     return { summary: { "E_일치": ok, "E_불일치": bad }, details: records };
   }
   // ===== 실행 엔트리 =====
   async function run() {
     try {
       // UI 상태
       const btn = document.getElementById('mfRun');
       if (btn) { btn.disabled = true; btn.textContent = '실행 중...'; }
       const logEl = document.getElementById('mfLog');
       if (logEl) logEl.textContent = '';
   
       // SheetJS 로드 확인 (cdn 또는 로컬 lib/xlsx.full.min.js)
       await waitForXLSX();
   
       // 파일 가져오기
       const fileInput = document.getElementById('mfFile');
       const f = fileInput?.files?.[0];
       if (!f) {
         log('엑셀 파일을 선택해 주세요.');
         return;
       }
   
       // 워크북 읽기
       const wb = await readWorkbook(f);
       const base = f.name.replace(/\.[^.]+$/, '');
   
       // 검사 실행(A~E)
       log('검사 시작 (A~E)…');
       const A = runCheckA(wb); await uiYield();
       const B = runCheckB(wb); await uiYield();
       const C = runCheckC(wb); await uiYield();
       const D = runCheckD(wb); await uiYield();
       const E = runCheckE(wb);
   
       // 요약 로그
       log(`[A] 일치 ${A.summary.A_일치}, 불일치 ${A.summary.A_불일치}`);
       log(`[B] 일치 ${B.summary.B_일치}, 불일치 ${B.summary.B_불일치}`);
       log(`[C] 일치 ${C.summary.C_일치}, 불일치 ${C.summary.C_불일치}`);
       log(`[D] 일치 ${D.summary.D_일치}, 불일치 ${D.summary.D_불일치}`);
       log(`[E] 일치 ${E.summary.E_일치}, 불일치 ${E.summary.E_불일치}`);
   
       // 결과 저장 (A~E 각 3시트: 전체/불일치/일치)
       saveAllSectionsAsOneWorkbook(base, [
         { name: 'A', rows: A.details },
         { name: 'B', rows: B.details },
         { name: 'C', rows: C.details },
         { name: 'D', rows: D.details },
         { name: 'E', rows: E.details },
       ]);
   
       log('엑셀 저장 완료: 종합검사.xlsx (A~E 각 3시트)');
   
     } catch (e) {
       console.error(e);
       log(`ERROR: ${e && e.message ? e.message : e}`);
     } finally {
       const btn = document.getElementById('mfRun');
       if (btn) { btn.disabled = false; btn.textContent = '검사 실행 (A~E)'; }
     }
   }

  // ===== 실행 =====
   const base = f.name.replace(/\.[^.]+$/, '');

   const A = runCheckA(wb); await new Promise(r=>setTimeout(r,0));
   const B = runCheckB(wb); await new Promise(r=>setTimeout(r,0));
   const C = runCheckC(wb); await new Promise(r=>setTimeout(r,0));
   const D = runCheckD(wb); await new Promise(r=>setTimeout(r,0));
   const E = runCheckE(wb);
   // 로그는 요약만
   log(`[A] 일치 ${A.summary.A_일치}, 불일치 ${A.summary.A_불일치}`);
   log(`[B] 일치 ${B.summary.B_일치}, 불일치 ${B.summary.B_불일치}`);
   log(`[C] 일치 ${C.summary.C_일치}, 불일치 ${C.summary.C_불일치}`);
   log(`[D] 일치 ${D.summary.D_일치}, 불일치 ${D.summary.D_불일치}`);
   log(`[E] 일치 ${E.summary.E_일치}, 불일치 ${E.summary.E_불일치}`);
   
   // 한 파일에 시트로 모두 저장
   saveAllSectionsAsOneWorkbook(base, [
     { name: 'A', rows: A.details },
     { name: 'B', rows: B.details },
     { name: 'C', rows: C.details },
     { name: 'D', rows: D.details },
     { name: 'E', rows: E.details },
   ]);
   
   log('엑셀 저장 완료: 종합검사.xlsx (A~E 각 3시트)');
      
    } catch (e) {
      console.error(e);
      log(`ERROR: ${e.message || e}`);
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    const btn = document.getElementById('mfRun');
    if (btn) btn.addEventListener('click', run);
  });
})();
