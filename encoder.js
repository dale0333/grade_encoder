/* ================================================================
   ADSJ Grade encoder — encoder.js  v5.0
   NEW:
   - Smart column matching: maps uploaded columns to live sheet
     by keyword (PRELIM/MIDTERM/FINAL) regardless of exact label
   - Unlock panel: reveals hidden Save/Finalize/Print buttons,
     force-unlocks locked grade columns (MIDTERM, FINAL)
   - AJAX queue interceptor: serializes all savegrade calls
   ================================================================ */
(function () {
  if (window.__aimsEncoderLoaded) return;
  window.__aimsEncoderLoaded = true;

  var POLL_MS       = 30;
  var INPUT_TIMEOUT = 4000;
  var MAX_RETRIES   = 3;
  var PERIOD_KEYS   = ['PRELIM', 'MIDTERM', 'FINAL'];
  var unlockKeepAlive = null;

  /* ══════════════════════════════════════════════════════════════
     AJAX QUEUE — serialise all savegrade requests
  ══════════════════════════════════════════════════════════════ */
  function installSaveQueue() {
    var $ = window.jQuery;
    if ($.fn.__aimsQueued) return;
    $.fn.__aimsQueued = true;
    var queue = [], running = false, orig = $.ajax.bind($);
    function run() {
      if (running || !queue.length) return;
      running = true;
      var item = queue.shift();
      var uc   = item.s.complete;
      item.s.complete = function (x, st) {
        running = false;
        if (typeof uc === 'function') uc(x, st);
        item.r(st); run();
      };
      orig(item.s);
    }
    $.ajax = function (settings) {
      var url = (typeof settings === 'string') ? settings : (settings && settings.url) || '';
      if (String(url).indexOf('savegrade') > -1) {
        if (typeof settings === 'string') settings = { url: settings };
        var d = $.Deferred();
        queue.push({ s: settings, r: d.resolve });
        run(); return d.promise();
      }
      return orig.apply(this, arguments);
    };
  }

  /* ══════════════════════════════════════════════════════════════
     COLUMN KEY  — normalises a label to a keyword
     "PRELIM PCTG GRADE" → "PRELIM"
     "MIDTERM PCTG GRADE" → "MIDTERM"
     "FINAL PCTG GRADE"  → "FINAL"
  ══════════════════════════════════════════════════════════════ */
  function colKey(label) {
    var u = (label || '').toUpperCase();
    if (u.indexOf('PRELIM')  > -1) return 'PRELIM';
    if (u.indexOf('MIDTERM') > -1) return 'MIDTERM';
    if (u.indexOf('FINAL')   > -1) return 'FINAL';
    return u.replace(/[^A-Z0-9]/g,'').slice(0,12);
  }

  /* ══════════════════════════════════════════════════════════════
     SCAN PAGE — reads ALL grade cols (open + locked)
  ══════════════════════════════════════════════════════════════ */
  function scanPage() {
    var $ = window.jQuery;
    var info = {
      sy:      ($('#display_sy').text()      || '').trim(),
      sem:     ($('#display_sem').text()     || '').trim(),
      subject: ($('#display_subject').text() || '').trim(),
      section: ($('#display_section').text() || '').trim(),
    };
    var open = [], all = [];
    $('#gradingsheet_datable thead tr:first th').each(function (idx) {
      var $th = $(this), id = ($th.attr('data-id') || '').trim();
      if (!id || id === 'name') return;
      if (parseInt($th.attr('data-istotal') || 0)) return;
      var notEd = parseInt($th.attr('data-noteditable') || 1);
      var col   = { id: id, label: $th.text().trim(), key: colKey($th.text()), thIndex: idx, open: !notEd };
      all.push(col);
      if (!notEd) open.push(col);
    });
    var students = [];
    $('#gradingsheet_datable tbody tr').each(function () {
      var sno  = $(this).find('.col2 span:first-child b').text().trim();
      var name = $(this).find('.col2 span:last-child').text().trim();
      if (sno) students.push({ sno: sno, name: name });
    });
    return { info: info, open: open, all: all, students: students };
  }

  function getPeriodColumns(cols) {
    var byKey = {};
    cols.forEach(function (c) { byKey[c.key] = c; });
    return PERIOD_KEYS.map(function (k) {
      return byKey[k] || {
        id: '',
        label: k + ' PCTG GRADE',
        key: k,
        thIndex: null,
        open: false
      };
    });
  }

  /* ══════════════════════════════════════════════════════════════
     UNLOCK  — force-open locked cols + reveal hidden buttons
  ══════════════════════════════════════════════════════════════ */
  function forceShowButtons() {
    var $ = window.jQuery;
    var btnSel = ['#button_save', '#button_finalize', '#button_print'];
    btnSel.forEach(function (sel) {
      $(sel).each(function () {
        this.style.setProperty('display', 'inline-block', 'important');
        this.style.setProperty('visibility', 'visible', 'important');
        this.style.removeProperty('opacity');
        this.removeAttribute('hidden');
      });
    });
    ['#btnsave','#btnfinalize','#btnprint',
     '.btn-save','[id*="save"]','[id*="finalize"]','[id*="print"]'].forEach(function(sel){
      $(sel).each(function(){
        var el = this;
        el.style.removeProperty('display');
        el.style.removeProperty('visibility');
        el.removeAttribute('hidden');
        $(el).removeClass('hidden d-none ng-hide').show();
      });
    });
    $('.widget-footer, tfoot, .grading-footer, .grade-actions, .actions').find('*').addBack().each(function(){
      if ($(this).css('display') === 'none') $(this).show();
    });
  }

  function forceUnlockColumnsByKeys(keys) {
    var $ = window.jQuery;
    var unlockedByIndex = {};
    $('#gradingsheet_datable thead tr:first th').each(function () {
      var $th  = $(this);
      var id   = ($th.attr('data-id') || '').trim();
      if (!id || id === 'name') return;
      if (parseInt($th.attr('data-istotal') || 0)) return;
      var key = colKey($th.text());
      if (!keys || !keys.length || keys.indexOf(key) > -1) {
        $th.attr('data-noteditable', '0');
        unlockedByIndex[$th.index()] = true;
      }
    });
    $('#gradingsheet_datable tbody tr').each(function () {
      var $tr = $(this);
      Object.keys(unlockedByIndex).forEach(function (idx) {
        var $td = $tr.find('td').eq(Number(idx));
        if (!$td.hasClass('istotal') && !$td.hasClass('studentname')) {
          $td.removeClass('noteditable');
          $td.removeAttr('data-noteditable');
        }
      });
    });
  }

  function enableUnlockPersistence() {
    if (unlockKeepAlive) return;
    unlockKeepAlive = setInterval(function () {
      if (!document.body.classList.contains('_aims-unlocked')) return;
      forceUnlockColumnsByKeys(PERIOD_KEYS);
      forceShowButtons();
    }, 500);
  }

  function unlockAll() {
    var $ = window.jQuery;

    /* 1. Reveal hidden page buttons (Save / Finalize / Print) */
    document.body.classList.add('_aims-unlocked');
    forceShowButtons();
    forceUnlockColumnsByKeys(PERIOD_KEYS);
    enableUnlockPersistence();

    return scanPage();   /* re-scan so panel reflects newly unlocked cols */
  }

  /* ══════════════════════════════════════════════════════════════
     BUILD TEMPLATE  — always includes ALL grade cols
     (open + locked), marks locked ones clearly
  ══════════════════════════════════════════════════════════════ */
  function buildTemplate(data, includeAll) {
    var info  = data.info;
    var cols  = getPeriodColumns(data.all);
    var students = data.students;
    var wb    = XLSX.utils.book_new();
    var hdr   = ['Student Number', 'Student Name', 'PRELIM', 'MIDTERM', 'FINAL', 'REMARKS'];
    var rows  = [
      ['SY: '+info.sy, 'Sem: '+info.sem, 'Subject: '+info.subject, 'Section: '+info.section],
      [],
      hdr
    ];
    students.forEach(function(s){
      rows.push([s.sno, s.name].concat(cols.map(function(){return '';})).concat(['']));
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    for (var r = 4; r <= 3 + students.length; r++) {
      var a = 'A'+r; if (ws[a]) { ws[a].t='s'; ws[a].z='@'; }
    }
    ws['!cols'] = [{wch:18},{wch:42},{wch:14},{wch:14},{wch:14},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws, 'Grade Input');

    /* META — store key (PRELIM/MIDTERM/FINAL) for smart matching on upload */
    var meta = [['__META__','v5'],['col_count', cols.length]];
    cols.forEach(function(c,i){
      meta.push(['col_'+i+'_id',    c.id]);
      meta.push(['col_'+i+'_label', c.label]);
      meta.push(['col_'+i+'_key',   c.key]);
      meta.push(['col_'+i+'_thIdx', c.thIndex]);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(meta), '__META__');
    var sect = (info.section||'grade').replace(/\s+/g,'_').replace(/[^A-Za-z0-9_-]/g,'');
    var fn   = 'AIMS_'+(info.subject||'SUBJ')+'_'+sect+'_template.xlsx';
    XLSX.writeFile(wb, fn);
    return fn;
  }

  /* ══════════════════════════════════════════════════════════════
     PARSE UPLOAD — smart column matching by keyword
  ══════════════════════════════════════════════════════════════ */
  function parseUpload(file, selectedPeriod, onDone) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb     = XLSX.read(e.target.result, { type:'binary', raw:false });

        /* Read META */
        var templateCols = [];
        var wsMeta = wb.Sheets['__META__'];
        if (wsMeta) {
          var mr  = XLSX.utils.sheet_to_json(wsMeta,{header:1,defval:''});
          var cnt = 0;
          mr.forEach(function(r){ if(r[0]==='col_count') cnt=parseInt(r[1]); });
          for (var i=0;i<cnt;i++) {
            var o={};
            mr.forEach(function(r){
              if(r[0]==='col_'+i+'_id')    o.id=r[1];
              if(r[0]==='col_'+i+'_label') o.label=r[1];
              if(r[0]==='col_'+i+'_key')   o.key=r[1];
              if(r[0]==='col_'+i+'_thIdx') o.thIndex=parseInt(r[1]);
            });
            if(o.id) templateCols.push(o);
          }
        }

        /* Smart-match template cols → live page cols by key */
        var liveCols = scanPage().all;
        var liveByKey = {};
        liveCols.forEach(function(c){ liveByKey[c.key] = c; });

        /* Build matched cols: prefer live thIndex (page may have shifted) */
        var matchedCols = templateCols.map(function(tc){
          var live = liveByKey[tc.key] || liveByKey[colKey(tc.label)];
          return {
            id:      live ? live.id      : tc.id,
            label:   live ? live.label   : tc.label,
            key:     tc.key,
            thIndex: live ? live.thIndex : tc.thIndex,
            open:    live ? live.open    : false
          };
        });

        /* Fall back if no META */
        if (!matchedCols.length) matchedCols = liveCols.filter(function(c){ return c.open; });

        /* Read grade data */
        var wsG  = wb.Sheets['Grade Input'] || wb.Sheets[wb.SheetNames[0]];
        var rows = XLSX.utils.sheet_to_json(wsG,{header:1,defval:'',raw:false});
        var hi   = -1;
        rows.forEach(function(r,i){
          if(String(r[0]).toLowerCase().indexOf('student number')>-1) hi=i;
        });

        /* Also detect column positions from header row */
        var headerRow = hi >= 0 ? rows[hi] : null;
        var colPositions = [];   /* index in Excel row → matched live col */
        if (headerRow) {
          for (var ci = 2; ci < headerRow.length - 1; ci++) {
            var hkey = colKey(String(headerRow[ci]));
            var live2 = liveByKey[hkey];
            if (live2) colPositions.push({ excelIdx: ci, liveCol: live2 });
          }
        }
        /* Fall back to sequential matchedCols */
        if (!colPositions.length) {
          matchedCols.forEach(function(mc, mi){ colPositions.push({ excelIdx: mi+2, liveCol: mc }); });
        }

        /* Keep only selected upload period */
        if (selectedPeriod && selectedPeriod !== 'ALL') {
          colPositions = colPositions.filter(function(cp){
            return cp.liveCol && cp.liveCol.key === selectedPeriod;
          });
          if (!colPositions.length && liveByKey[selectedPeriod]) {
            var manualIdx = -1;
            if (headerRow) {
              for (var hci = 0; hci < headerRow.length; hci++) {
                if (colKey(String(headerRow[hci])) === selectedPeriod) {
                  manualIdx = hci;
                  break;
                }
              }
            }
            if (manualIdx > -1) {
              colPositions.push({ excelIdx: manualIdx, liveCol: liveByKey[selectedPeriod] });
            }
          }
        }

        var map   = {};
        var start = (hi>=0) ? hi+1 : 0;
        for (var ri=start; ri<rows.length; ri++) {
          var row = rows[ri];
          var raw = String(row[0]||'').trim();
          if (!raw || raw.toLowerCase().indexOf('student')>-1) continue;
          var sno = norm(raw);
          map[sno] = {};
          colPositions.forEach(function(cp){
            var g = String(row[cp.excelIdx]||'').trim().toUpperCase();
            if(g) map[sno][cp.liveCol.id] = { grade:g, thIndex:cp.liveCol.thIndex, key:cp.liveCol.key };
          });
        }

        onDone({ map:map, cols:matchedCols, colPositions:colPositions, total:Object.keys(map).length });
      } catch(err) { alert('ADSJ Grade encoder — Error reading file: '+err.message); }
    };
    reader.readAsBinaryString(file);
  }

  /* ══════════════════════════════════════════════════════════════
     WRITE ONE CELL
  ══════════════════════════════════════════════════════════════ */
  function pollForInput($cell, onFound, onTimeout) {
    var e=0, t=setInterval(function(){
      var $i=$cell.find('input');
      if($i.length){clearInterval(t);onFound($i);return;}
      if((e+=POLL_MS)>=INPUT_TIMEOUT){clearInterval(t);onTimeout();}
    },POLL_MS);
  }

  function waitForCellText($cell, expected, onDone, onTimeout) {
    var e=0, ne=parseFloat(expected);
    var t=setInterval(function(){
      var shown=$cell.clone().find('input,select').remove().end().text().trim();
      var ns=parseFloat(shown);
      var ok=(shown===expected)||(!isNaN(ne)&&!isNaN(ns)&&ne===ns)||(shown.length>0&&!$cell.find('input').length);
      if(ok){clearInterval(t);onDone(shown);return;}
      if((e+=POLL_MS)>6000){clearInterval(t);onTimeout(shown);}
    },POLL_MS);
  }

  function writeOneCell($, $cell, grade, attempt, onSuccess, onFail) {
    function forceSetCellValue() {
      $cell.data('grade', grade);
      $cell.attr('data-grade', grade);
      $cell.html(grade);
    }
    if (attempt > MAX_RETRIES) { onFail('max_retries'); return; }
    if (document.activeElement && document.activeElement !== document.body) {
      document.activeElement.blur();
    }
    setTimeout(function(){
      $cell.trigger($.Event('click', { trueClick: true }));
      pollForInput($cell, function($inp){
        var ns = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype,'value');
        if (ns && ns.set) {
          ns.set.call($inp[0], grade);
          $inp[0].dispatchEvent(new Event('input',  {bubbles:true}));
          $inp[0].dispatchEvent(new Event('change', {bubbles:true}));
        } else {
          $inp.val(grade).trigger('input').trigger('change');
        }
        $cell.data('click', true);
        $inp.trigger('blur');
        waitForCellText($cell, grade,
          function(){ onSuccess(); },
          function(){
            if (attempt >= 2) {
              forceSetCellValue();
              waitForCellText($cell, grade,
                function(){ onSuccess(); },
                function(){ writeOneCell($, $cell, grade, attempt+1, onSuccess, onFail); }
              );
            } else {
              writeOneCell($, $cell, grade, attempt+1, onSuccess, onFail);
            }
          }
        );
      }, function(){
        if (attempt >= 2) {
          forceSetCellValue();
          waitForCellText($cell, grade,
            function(){ onSuccess(); },
            function(){ writeOneCell($, $cell, grade, attempt+1, onSuccess, onFail); }
          );
        } else {
          writeOneCell($, $cell, grade, attempt+1, onSuccess, onFail);
        }
      });
    }, attempt===1 ? 80 : 400);
  }

  /* ══════════════════════════════════════════════════════════════
     APPLY ALL GRADES
  ══════════════════════════════════════════════════════════════ */
  function applyGrades(gradeMap, onProgress, onDone) {
    var $ = window.jQuery;
    installSaveQueue();
    var keys=Object.keys(gradeMap), total=keys.length;
    var applied=0, notFound=[], locked=[], failed=[];
    var retryQueue = [];

    function getModuleLink() {
      var p = (window.location.pathname || '').split('/').filter(Boolean);
      return {
        module: p[0] || '',
        link: p[1] || ''
      };
    }

    function collectCurrentSheetPayload() {
      var data_type = $("#gtabs li.active").attr("data-type") || "gradingsheet";
      var gt = $((data_type==="attendance"?"#attendancesheet_datable":"#gradingsheet_datable"));
      var collected = {};
      gt.find("tbody tr").each(function(){
        var sid = $(this).data("studclasslistid");
        if (typeof sid === "undefined") return;
        var tmp = {};
        $(this).find("td:not(.studentname)").each(function(){
          tmp[$(this).data("id")] = (typeof $(this).data("grade") !== "undefined")
            ? $(this).data("grade").toString().replace("&nbsp;","")
            : "";
        });
        collected[sid] = tmp;
      });
      return {
        gt: gt,
        data_type: data_type,
        data_sheet: gt.attr("data-sheet") || "",
        collected: collected
      };
    }

    function persistCellToBackend(cellid, attempts, cb) {
      var ml = getModuleLink();
      var payload = collectCurrentSheetPayload();
      $.ajax({
        url: window.location.origin + "/new_savegrade",
        type: "POST",
        timeout: 20000,
        data: {
          collected: JSON.stringify(payload.collected),
          data_type: payload.data_type,
          data_sheet: payload.data_sheet,
          data_process: "save",
          link: ml.link,
          module: ml.module,
          selectedCellId: cellid
        },
        success: function(){ cb(true); },
        error: function(){
          if (attempts > 1) return setTimeout(function(){
            persistCellToBackend(cellid, attempts-1, cb);
          }, 250);
          cb(false);
        }
      });
    }

    function nextStudent(i) {
      if (i >= total) { return; }
      onProgress(i, total);
      var sno=keys[i], cols=gradeMap[sno], $row=null;
      $('#gradingsheet_datable tbody tr').each(function(){
        if(norm($(this).find('.col2 span:first-child b').text().trim())===sno){$row=$(this);return false;}
      });
      if(!$row){notFound.push(sno);return setTimeout(function(){nextStudent(i+1);},60);}
      var cids=Object.keys(cols);
      function nextCol(ci){
        if(ci>=cids.length){applied++;return setTimeout(function(){nextStudent(i+1);},60);}
        var info=$cell=null;
        info=cols[cids[ci]];
        var $cell=$row.find('td').eq(info.thIndex);
        if(!$cell.length||$cell.hasClass('noteditable')||$cell.hasClass('istotal')){
          locked.push(sno+' ['+(info.key||cids[ci])+']');return nextCol(ci+1);
        }
        var cellid = $cell.data("cellid");
        function persistThenContinue() {
          if (!cellid) return nextCol(ci+1);
          persistCellToBackend(cellid, 3, function(ok){
            if (ok) return nextCol(ci+1);
            /* failed backend save: rewrite once and persist again */
            writeOneCell($,$cell,info.grade,2,function(){
              persistCellToBackend(cellid, 2, function(ok2){
                if (!ok2) failed.push(sno+' ['+(info.key||cids[ci])+'] backend_save_failed');
                nextCol(ci+1);
              });
            },function(){
              failed.push(sno+' ['+(info.key||cids[ci])+'] rewrite_failed');
              nextCol(ci+1);
            });
          });
        }
        writeOneCell($,$cell,info.grade,1,
          function(){persistThenContinue();},
          function(r){
            retryQueue.push({
              sno: sno,
              key: (info.key||cids[ci]),
              thIndex: info.thIndex,
              grade: info.grade
            });
            failed.push(sno+' ['+(info.key||cids[ci])+'] '+r);
            nextCol(ci+1);
          }
        );
      }
      nextCol(0);
    }
    function runRetryPass(done){
      if(!retryQueue.length){ done(); return; }
      var q = retryQueue.slice();
      retryQueue = [];
      var ri = 0;
      function nextRetry(){
        if(ri >= q.length){ done(); return; }
        var item = q[ri++];
        var $row = null;
        $('#gradingsheet_datable tbody tr').each(function(){
          if(norm($(this).find('.col2 span:first-child b').text().trim())===item.sno){$row=$(this);return false;}
        });
        if(!$row){ return nextRetry(); }
        var $cell = $row.find('td').eq(item.thIndex);
        if(!$cell.length){ return nextRetry(); }
        writeOneCell($,$cell,item.grade,2,function(){ nextRetry(); },function(){ nextRetry(); });
      }
      nextRetry();
    }
    nextStudent(0);
    var watcher = setInterval(function(){
      if (applied + notFound.length >= total) {
        clearInterval(watcher);
        runRetryPass(function(){
          onDone({applied:applied,notFound:notFound,locked:locked,failed:failed});
        });
      }
    }, 120);
  }

  /* ══════════════════════════════════════════════════════════════
     CSS
  ══════════════════════════════════════════════════════════════ */
  var CSS = `
    #_aFab{position:fixed;bottom:24px;right:24px;z-index:2147483647;
      width:54px;height:54px;border-radius:50%;
      background:linear-gradient(135deg,#005C2A,#00A84F);
      color:#fff;font-size:26px;border:none;cursor:pointer;
      box-shadow:0 4px 20px rgba(0,120,60,.55);
      display:flex;align-items:center;justify-content:center;
      transition:transform .15s,box-shadow .15s;font-family:sans-serif;}
    #_aFab:hover{transform:scale(1.12);box-shadow:0 6px 28px rgba(0,120,60,.65);}
    #_aFab .tip{position:absolute;right:62px;bottom:50%;transform:translateY(50%);
      background:#111;color:#fff;font-size:11px;white-space:nowrap;padding:4px 10px;
      border-radius:5px;opacity:0;pointer-events:none;transition:opacity .2s;
      font-family:'Segoe UI',sans-serif;}
    #_aFab:hover .tip{opacity:1;}

    #_aPanel{position:fixed;bottom:90px;right:20px;width:390px;z-index:2147483646;
      background:#fff;border-radius:14px;
      box-shadow:0 12px 50px rgba(0,0,0,.32);border:1.5px solid #00833F;
      overflow:hidden;animation:_aPop .22s cubic-bezier(.34,1.56,.64,1);}
    @keyframes _aPop{from{opacity:0;transform:scale(.88) translateY(16px)}to{opacity:1;transform:scale(1)translateY(0)}}
    #_aPanel *{box-sizing:border-box;font-family:'Segoe UI',Tahoma,sans-serif;margin:0;padding:0;}
    #_aPanel .ah{background:linear-gradient(135deg,#004d22,#00A84F);color:#fff;
      padding:12px 15px;display:flex;justify-content:space-between;align-items:center;
      cursor:move;user-select:none;}
    #_aPanel .ah-title{font-weight:700;font-size:13px;display:flex;align-items:center;gap:8px;}
    #_aPanel .ah-badge{background:rgba(255,255,255,.2);border-radius:4px;font-size:9px;padding:2px 7px;font-weight:700;}
    #_aPanel .ah-x{background:none;border:none;color:#fff;font-size:18px;cursor:pointer;opacity:.7;}
    #_aPanel .ah-x:hover{opacity:1;}
    #_aPanel .ab{padding:13px;max-height:calc(100vh - 170px);overflow-y:auto;}
    #_aPanel .sched{background:#EAF7EF;border:1px solid #A8D5BB;border-radius:7px;
      padding:8px 11px;font-size:11px;color:#1B5E34;margin-bottom:10px;line-height:1.8;}
    #_aPanel .sched strong{display:block;font-size:12px;margin-bottom:1px;}
    #_aPanel .pills{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:11px;}
    #_aPanel .pill{padding:3px 10px;border-radius:99px;font-size:10.5px;font-weight:600;border:1px solid;}
    #_aPanel .pill.o{background:#ECFDF5;border-color:#34C079;color:#0A6B36;}
    #_aPanel .pill.c{background:#F5F5F5;border-color:#CCC;color:#999;}
    #_aPanel .pill.u{background:#FEF3C7;border-color:#F59E0B;color:#92400E;}

    /* Unlock panel */
    #_aPanel .unlock-box{border:1.5px solid #F59E0B;border-radius:9px;
      background:#FFFBEB;padding:11px 13px;margin-bottom:11px;}
    #_aPanel .unlock-box h4{font-size:12px;font-weight:700;color:#92400E;margin-bottom:8px;
      display:flex;align-items:center;gap:6px;}
    #_aPanel .unlock-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px;}
    #_aPanel .ub{padding:7px 8px;border:1.5px solid #D97706;border-radius:7px;
      background:#fff;font-size:11px;font-weight:600;color:#92400E;cursor:pointer;
      text-align:center;transition:background .15s;}
    #_aPanel .ub:hover{background:#FEF3C7;}
    #_aPanel .ub.active{background:#D97706;color:#fff;border-color:#B45309;}
    #_aPanel .ub-all{grid-column:1/-1;}

    #_aPanel .steps{display:flex;flex-direction:column;gap:7px;}
    #_aPanel .step{border-radius:9px;border:1px solid #E0E0E0;overflow:hidden;}
    #_aPanel .step-h{padding:9px 12px;display:flex;align-items:center;gap:9px;
      background:#FAFAFA;font-size:12px;font-weight:600;color:#333;}
    #_aPanel .sn{width:22px;height:22px;border-radius:50%;display:flex;
      align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;}
    #_aPanel .sn.a{background:#00833F;color:#fff;}
    #_aPanel .sn.d{background:#34C079;color:#fff;}
    #_aPanel .sn.i{background:#E0E0E0;color:#999;}
    #_aPanel .step-b{padding:10px 12px;display:none;}
    #_aPanel .step.active .step-b{display:block;}
    #_aPanel .btn{width:100%;padding:9px;border:none;border-radius:7px;
      font-weight:700;font-size:12.5px;cursor:pointer;
      transition:background .15s,transform .1s;
      display:flex;align-items:center;justify-content:center;gap:7px;}
    #_aPanel .btn:active{transform:scale(.97);}
    #_aPanel .btn.g{background:#00833F;color:#fff;}
    #_aPanel .btn.g:hover:not(:disabled){background:#006530;}
    #_aPanel .btn.sk{background:none;border:1px dashed #CCC;color:#888;font-size:11.5px;margin-top:6px;}
    #_aPanel .btn.sk:hover{border-color:#00833F;color:#00833F;background:#F5FDF8;}
    #_aPanel .btn:disabled{background:#CCC!important;color:#999!important;cursor:not-allowed;}
    #_aPanel .drop{border:2px dashed #CCC;border-radius:8px;padding:16px;
      text-align:center;cursor:pointer;font-size:12px;color:#888;
      transition:border-color .2s,background .2s;margin-bottom:9px;}
    #_aPanel .drop:hover,#_aPanel .drop.drag{border-color:#00833F;background:#F0FAF4;color:#00833F;}
    #_aPanel .drop input{display:none;}
    #_aPanel .prev{max-height:175px;overflow-y:auto;
      border:1px solid #E8E8E8;border-radius:6px;margin-bottom:9px;font-size:11px;}
    #_aPanel .prev table{width:100%;border-collapse:collapse;}
    #_aPanel .prev th{background:#00833F;color:#fff;padding:5px 7px;
      position:sticky;top:0;text-align:left;font-size:10.5px;white-space:nowrap;}
    #_aPanel .prev td{padding:4px 7px;border-bottom:1px solid #F2F2F2;}
    #_aPanel .prev tr:nth-child(even) td{background:#F9FAF9;}
    #_aPanel .prev .nf td{color:#C0392B!important;background:#FFF5F5!important;}
    #_aPanel .pw{background:#EEE;border-radius:4px;height:8px;
      margin-bottom:8px;overflow:hidden;display:none;}
    #_aPanel .pb{height:100%;background:#00833F;border-radius:4px;
      width:0%;transition:width .3s ease,background .4s ease;}
    #_aPanel .ctr{font-size:12px;font-weight:700;color:#00833F;margin-bottom:5px;display:none;}
    #_aPanel .status{font-size:11.5px;min-height:14px;line-height:1.6;color:#333;}
    #_aPanel .warn{background:#FFF8E1;border:1px solid #FFD54F;border-radius:7px;
      padding:9px 11px;font-size:11px;color:#7A5800;margin-bottom:10px;}
    #_aPanel .col-map{background:#F8F8F8;border:1px solid #E0E0E0;border-radius:6px;
      padding:8px 10px;margin-bottom:8px;font-size:11px;color:#555;line-height:1.8;}
    #_aPanel .col-map strong{color:#00833F;}
    #_aPanel .col-map .miss{color:#C0392B;}
    body._aims-unlocked #button_save,
    body._aims-unlocked #button_finalize,
    body._aims-unlocked #button_print{
      display:inline-block !important;
      visibility:visible !important;
    }
  `;

  /* ══════════════════════════════════════════════════════════════
     PANEL
  ══════════════════════════════════════════════════════════════ */
  function openPanel() {
    if (document.getElementById('_aPanel')) return;
    var pd  = unlockAll();
    var inf = pd.info, open = pd.open, all = pd.all, students = pd.students;

    var openPills   = open.map(function(c){ return '<span class="pill o">✓ '+c.label+'</span>'; }).join('');
    var closedPills = all.filter(function(c){ return !c.open; })
                        .map(function(c){ return '<span class="pill c">🔒 '+c.label+'</span>'; }).join('');
    var noSched = !inf.subject
      ? '<div class="warn">⚠ No schedule selected — select one then click 📊 again.</div>' : '';

    var el = document.createElement('div');
    el.id  = '_aPanel';
    el.innerHTML =
      '<div class="ah">' +
        '<div class="ah-title">📊 ADSJ Grade encoder <span class="ah-badge">1.0</span></div>' +
        '<button class="ah-x" id="_aX">✕</button>' +
      '</div>' +
      '<div class="ab">' +
        noSched +
        '<div class="sched"><strong>📋 '+(inf.subject||'—')+' &nbsp;·&nbsp; '+(inf.section||'—')+'</strong>' +
          'SY '+(inf.sy||'—')+' &nbsp;|&nbsp; '+(inf.sem||'—')+' &nbsp;|&nbsp; '+students.length+' students</div>' +
        '<div class="pills" id="_aPills">'+openPills+closedPills+'</div>' +

        /* ── UNLOCK PANEL ── */
        '<div class="unlock-box">' +
          '<h4>🔓 Unlock Controls</h4>' +
          '<div class="unlock-grid">' +
            '<button class="ub" id="_uPrelim">Unlock PRELIM</button>' +
            '<button class="ub" id="_uMidterm">Unlock MIDTERM</button>' +
            '<button class="ub" id="_uFinal">Unlock FINAL</button>' +
            '<button class="ub" id="_uButtons">Show Save/Finalize</button>' +
            '<button class="ub ub-all" id="_uAll">🔓 Unlock Everything</button>' +
          '</div>' +
          '<div class="status" id="_uStatus" style="font-size:11px;color:#92400E;min-height:12px;"></div>' +
        '</div>' +

        '<div class="steps">' +
          '<div class="step '+(open.length||all.length?'active':'')+'" id="_aS1">' +
            '<div class="step-h"><div class="sn '+(open.length?'a':'i')+'" id="_aN1">1</div>Download grade template</div>' +
            '<div class="step-b">' +
              '<p style="font-size:11.5px;color:#555;line-height:1.6;margin-bottom:8px">' +
                'Template includes dedicated <strong>PRELIM / MIDTERM / FINAL</strong> columns.' +
              '</p>' +
              '<button class="btn g" id="_aDl" '+(all.length?'':' disabled')+'>⬇ Download Template</button>' +
              '<button class="btn sk" id="_aDlAll" style="margin-top:6px">⬇ Re-download Template</button>' +
              '<button class="btn sk" id="_aSkip" style="margin-top:4px">↩ Skip — I already have the template</button>' +
            '</div>' +
          '</div>' +

          '<div class="step" id="_aS2">' +
            '<div class="step-h"><div class="sn i" id="_aN2">2</div>Upload filled template</div>' +
            '<div class="step-b">' +
              '<div style="margin-bottom:8px">' +
                '<label for="_aPeriod" style="display:block;font-size:11px;color:#555;margin-bottom:4px">Upload target period</label>' +
                '<select id="_aPeriod" style="width:100%;padding:8px;border:1px solid #ccc;border-radius:7px;font-size:12px">' +
                  '<option value="">- Select period before upload -</option>' +
                  '<option value="PRELIM">PRELIM</option>' +
                  '<option value="MIDTERM">MIDTERM</option>' +
                  '<option value="FINAL">FINAL</option>' +
                '</select>' +
              '</div>' +
              '<div class="drop" id="_aDrop">' +
                '<input type="file" id="_aFile" accept=".xlsx,.xls,.csv">' +
                '📂 Click or drag your filled Excel file here' +
                '<div style="font-size:10.5px;color:#aaa;margin-top:3px">.xlsx / .xls / .csv</div>' +
              '</div>' +
              '<div id="_aColMap" style="display:none" class="col-map"></div>' +
              '<div class="prev" id="_aPrev" style="display:none"></div>' +
              '<div class="status" id="_aSt2"></div>' +
            '</div>' +
          '</div>' +

          '<div class="step" id="_aS3">' +
            '<div class="step-h"><div class="sn i" id="_aN3">3</div>Apply grades to the sheet</div>' +
            '<div class="step-b">' +
              '<div class="ctr" id="_aCtr"></div>' +
              '<div class="pw" id="_aPW"><div class="pb" id="_aPB"></div></div>' +
              '<div class="status" id="_aSt3" style="margin-bottom:8px"></div>' +
              '<button class="btn g" id="_aApply" disabled>▶ Apply Grades Now</button>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>';
    document.body.appendChild(el);

    /* Draggable */
    var hd=el.querySelector('.ah'), drag=false, dx=0, dy=0;
    hd.addEventListener('mousedown',function(e){
      if(e.target.id==='_aX') return;
      drag=true;var r=el.getBoundingClientRect();dx=e.clientX-r.left;dy=e.clientY-r.top;
    });
    document.addEventListener('mousemove',function(e){
      if(!drag)return;
      el.style.right='auto';el.style.bottom='auto';
      el.style.left=(e.clientX-dx)+'px';el.style.top=(e.clientY-dy)+'px';
    });
    document.addEventListener('mouseup',function(){drag=false;});
    document.getElementById('_aX').onclick=function(){el.remove();};

    /* Step manager */
    function go(n){
      [1,2,3].forEach(function(i){
        var s=document.getElementById('_aS'+i),nm=document.getElementById('_aN'+i);
        if(i<n){s.classList.remove('active');nm.className='sn d';nm.textContent='✓';}
        else if(i===n){s.classList.add('active');nm.className='sn a';nm.textContent=i;}
        else{s.classList.remove('active');nm.className='sn i';nm.textContent=i;}
      });
    }

    /* Refresh pills after unlock */
    function refreshPills(newPd) {
      var pp = document.getElementById('_aPills');
      if (!pp) return;
      var op = newPd.open.map(function(c){ return '<span class="pill o">✓ '+c.label+'</span>'; }).join('');
      var cp = newPd.all.filter(function(c){ return !c.open; }).map(function(c){ return '<span class="pill u">🔓 '+c.label+'</span>'; }).join('');
      pp.innerHTML = op + cp;
    }

    /* ── UNLOCK BUTTONS ── */
    function unlockKey(key, label) {
      forceUnlockColumnsByKeys([key]);
      forceShowButtons();
      enableUnlockPersistence();
      var uSt = document.getElementById('_uStatus');
      if(uSt) uSt.textContent = '✅ '+label+' unlocked!';
      var btnId = '_u' + label.charAt(0) + label.slice(1).toLowerCase();
      var b = document.getElementById(btnId);
      if (b) b.classList.add('active');
      refreshPills(scanPage());
    }

    function showButtons() {
      document.body.classList.add('_aims-unlocked');
      forceShowButtons();
      enableUnlockPersistence();
      var uSt=document.getElementById('_uStatus');
      if(uSt) uSt.textContent='✅ Save/Finalize/Print buttons should now be visible.';
      document.getElementById('_uButtons').classList.add('active');
    }

    document.getElementById('_uPrelim').onclick  = function(){ unlockKey('PRELIM','PRELIM'); };
    document.getElementById('_uMidterm').onclick = function(){ unlockKey('MIDTERM','MIDTERM'); };
    document.getElementById('_uFinal').onclick   = function(){ unlockKey('FINAL','FINAL'); };
    document.getElementById('_uButtons').onclick  = showButtons;
    document.getElementById('_uAll').onclick = function(){
      var newPd = unlockAll();
      showButtons();
      refreshPills(newPd);
      ['_uPrelim','_uMidterm','_uFinal','_uButtons'].forEach(function(id){
        var b=document.getElementById(id); if(b) b.classList.add('active');
      });
      document.getElementById('_uStatus').textContent='✅ Everything unlocked!';
    };

    /* ── STEP 1: DOWNLOAD ── */
    document.getElementById('_aDl').onclick = function(){
      var fn=buildTemplate(pd, false); go(2);
      document.getElementById('_aSt2').innerHTML='<span style="color:#00833F">✅ '+fn+' — fill grades &amp; upload below.</span>';
    };
    document.getElementById('_aDlAll').onclick = function(){
      /* Re-scan after any unlocks */
      var fn=buildTemplate(scanPage(), true); go(2);
      document.getElementById('_aSt2').innerHTML='<span style="color:#00833F">✅ '+fn+' (all columns) — fill grades &amp; upload below.</span>';
    };
    document.getElementById('_aSkip').onclick = function(){
      go(2);
      document.getElementById('_aSt2').innerHTML='<span style="color:#888">ℹ Using existing template.</span>';
    };

    /* ── STEP 2: UPLOAD ── */
    var drop=document.getElementById('_aDrop'),fi=document.getElementById('_aFile'),parsed=null;
    drop.addEventListener('dragover',  function(e){e.preventDefault();drop.classList.add('drag');});
    drop.addEventListener('dragleave', function(){drop.classList.remove('drag');});
    drop.addEventListener('drop', function(e){e.preventDefault();drop.classList.remove('drag');doFile(e.dataTransfer.files[0]);});
    drop.addEventListener('click', function(){fi.click();});
    fi.addEventListener('change', function(e){doFile(e.target.files[0]);});

    function doFile(f){
      if(!f)return;
      var periodSel = document.getElementById('_aPeriod');
      var targetPeriod = periodSel ? periodSel.value : '';
      if(!targetPeriod){
        document.getElementById('_aSt2').innerHTML='⚠ Select PRELIM, MIDTERM, or FINAL before uploading.';
        return;
      }
      document.getElementById('_aSt2').innerHTML='⏳ Reading file…';
      parseUpload(f,targetPeriod,function(r){
        parsed=r;
        showColMap(r);
        renderPreview(r);
        if (!r.colPositions.length) {
          document.getElementById('_aSt2').innerHTML='⚠ No '+targetPeriod+' column found/mapped in the uploaded file.';
          document.getElementById('_aApply').disabled=true;
          return;
        }
        document.getElementById('_aSt2').innerHTML=
          '✅ <strong>'+r.total+'</strong> students · '+r.colPositions.length+' column(s) mapped for <strong>'+targetPeriod+'</strong>';
        document.getElementById('_aApply').disabled=false;
        go(3);
      });
    }

    /* Show which Excel columns mapped to which live sheet columns */
    function showColMap(r){
      var cm=document.getElementById('_aColMap'); cm.style.display='block';
      var liveCols=scanPage().all, liveByKey={};
      liveCols.forEach(function(c){liveByKey[c.key]=c;});
      var html='<strong>Column mapping detected:</strong><br>';
      r.colPositions.forEach(function(cp){
        var lc=cp.liveCol;
        var inSheet=!!liveByKey[lc.key];
        html+=(inSheet
          ? '<strong>'+lc.key+'</strong> → col '+lc.thIndex+' ✅'
          : '<span class="miss">'+lc.key+' → not found in sheet ⚠</span>')+'<br>';
      });
      cm.innerHTML=html;
    }

    function renderPreview(r){
      var pv=document.getElementById('_aPrev'); pv.style.display='block';
      var known={};
      window.jQuery('#gradingsheet_datable tbody tr').each(function(){
        var sn=norm(window.jQuery(this).find('.col2 span:first-child b').text().trim());
        if(sn)known[sn]=true;
      });
      var ths=r.colPositions.map(function(cp){
        return '<th>'+cp.liveCol.key+'</th>';
      }).join('');
      var html='<table><thead><tr><th>Student No.</th>'+ths+'</tr></thead><tbody>';
      var keys=Object.keys(r.map),shown=Math.min(keys.length,22);
      for(var i=0;i<shown;i++){
        var sno=keys[i],ok=known[sno];
        var tds=r.colPositions.map(function(cp){
          var g=(r.map[sno][cp.liveCol.id]||{}).grade||'';
          return '<td style="text-align:center">'+(g||'—')+'</td>';
        }).join('');
        html+='<tr'+(ok?'':' class="nf"')+'><td>'+sno+(ok?'':' ⚠')+'</td>'+tds+'</tr>';
      }
      if(keys.length>shown)html+='<tr><td colspan="'+(r.colPositions.length+1)+'" style="text-align:center;color:#999;padding:5px">… '+(keys.length-shown)+' more</td></tr>';
      pv.innerHTML=html+'</tbody></table>';
    }

    /* ── STEP 3: APPLY ── */
    document.getElementById('_aApply').onclick=function(){
      if(!parsed)return;
      var btn=this;btn.disabled=true;
      var ctr=document.getElementById('_aCtr');
      var pw=document.getElementById('_aPW'),pb=document.getElementById('_aPB');
      var st=document.getElementById('_aSt3');
      ctr.style.display='block';pw.style.display='block';
      btn.textContent='⏳ Writing grades…';
      applyGrades(parsed.map,
        function(i,total){
          ctr.textContent='Writing student '+(i+1)+' / '+total+'…';
          pb.style.width=Math.round(((i+1)/total)*100)+'%';
        },
        function(s){
          pb.style.width='100%';
          pb.style.background=s.failed.length?'#E67E22':'#34C079';
          ctr.textContent=s.failed.length?'⚠ Done with errors':'✅ All done!';
          var allOk=s.failed.length===0;
          var msg='';
          if(allOk){
            msg='<span style="color:#00833F;font-size:13px;font-weight:700">✅ All '+s.applied+' grades written!</span>';
          } else {
            msg='<strong style="color:#C0392B">⚠ '+s.failed.length+' failed:</strong><br>';
            msg+='<span style="font-size:10.5px;color:#888">'+s.failed.join('<br>')+'</span><br>';
          }
          if(s.notFound.length) msg+='<br>⚠ Not found: <strong>'+s.notFound.length+'</strong>';
          if(s.locked.length)   msg+='<br>🔒 Still locked: <strong>'+s.locked.length+'</strong> — use Unlock above';
          msg+='<br><br><span style="font-weight:700;font-size:12.5px;color:'+(allOk?'#00833F':'#E67E22')+'">'
              +(allOk?'→ Now click Save in AIMS!':'→ Fix above, then click Save.')+'</span>';
          st.innerHTML=msg;
          btn.innerHTML=allOk?'✔ Click Save in AIMS':'⚠ Some failed';
          btn.style.background=allOk?'#34C079':'#E67E22';
          btn.disabled=false;
        }
      );
    };
  }

  /* ══════════════════════════════════════════════════════════════
     FAB
  ══════════════════════════════════════════════════════════════ */
  function injectFAB(){
    if(document.getElementById('_aFab'))return;
    var style=document.createElement('style');
    style.textContent=CSS;document.head.appendChild(style);
    var fab=document.createElement('button');
    fab.id='_aFab';
    fab.innerHTML='📊<span class="tip">ADSJ Grade encoder</span>';
    document.body.appendChild(fab);
    fab.addEventListener('click',function(){
      var p=document.getElementById('_aPanel');
      if(p){p.remove();return;}
      if(!document.getElementById('gradingsheet_datable')){
        alert('ADSJ Grade encoder: Please go to Grading Sheet and select a schedule first.');return;
      }
      openPanel();
    });
  }

  function norm(s){return s.replace(/[^0-9]/g,'').padStart(11,'0');}

  if(document.body){injectFAB();}
  else{document.addEventListener('DOMContentLoaded',injectFAB);}
})();
