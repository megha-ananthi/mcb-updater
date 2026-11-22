/*
 * mcb-updater: the panel the bookmark opens inside a logged-in MyClassBoard (MCB) page.
 * It talks only to the MCB site it is opened on, with the teacher's own session.
 * Nothing is stored, and nothing is sent anywhere else.
 */
(function () {
  'use strict';
  if (window.__mcbUpdater) { window.__mcbUpdater.toggle(); return; }

  const script = document.currentScript;
  const BASE = script && script.src ? script.src.replace(/[^/]*(\?.*)?$/, '') : '';
  const fakeApi = window.MCB_UPDATER_FAKE_API || null; // set only by test/harness.html
  const onLocalTest = /^(localhost|127\.0\.0\.1)$/.test(location.hostname); // no MCB login exists there
  if (!fakeApi && !onLocalTest && !/(^|\.)myclassboard\.com$/i.test(location.hostname)) {
    alert('Open this bookmark on your MyClassBoard page, after you log in.');
    return;
  }

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const sleep = (ms) => new Promise((done) => setTimeout(done, ms));

  // ---------- MCB calls ----------
  //
  // Every MCB request goes through this one object, so the rest of the panel does not care how MCB works:
  //   listBranches()                            -> [{id, name}]
  //   listClasses(branchId)                     -> [{id, name}]
  //   listExams(branchId, classId)              -> [{id, name, label}]   newest first
  //   listSubjects({branchId, classId, examId}) -> [{id, name, max}]     max = null when MCB does not say
  //   listSections(branchId, classId)           -> [{id, name}]
  //   getSection(ctx)                           -> {students, subjects, saved, gradeOnly}  (core.parseSectionInfo)
  //   uploadMarks(ctx, subject, students, marks) -> resolves when MCB took the file; throws MCB's message if not
  // ctx = {branchId, classId, examId, examName, sectionId, subjects: [{id, name}]}
  //
  // These are the requests MCB's own "Bulk Upload Marks/Grades" page makes (see README, "MCB endpoints").
  function createMcbApi() {
    const lost = 'MCB did not answer as expected. Your login may have ended: reload the MCB page, log in, and click the bookmark again.';
    const enc = encodeURIComponent;
    let ids = null; // {orgId, branchId, academicYearId}, read once from MCB's Other Exams dashboard
    const sectionsOf = new Map(); // 'branchId|classId' -> sections, so the list is asked for once

    async function request(path, options) {
      const init = { method: 'POST', credentials: 'same-origin', headers: { 'X-Requested-With': 'XMLHttpRequest' } };
      if (options && options.method) init.method = options.method;
      if (options && options.form) init.body = new URLSearchParams(options.form);
      if (options && options.body) init.body = options.body;
      let response;
      try { response = await fetch(path, init); } catch (e) { throw new Error(lost); }
      if (!response.ok) throw new Error('MCB answered with an error (' + response.status + ') for ' + path.split('?')[0] + '.');
      return response.text();
    }
    async function json(path, options) {
      const text = await request(path, options);
      try { return JSON.parse(text); } catch (e) { throw new Error(lost); }
    }
    async function start() {
      if (ids) return ids;
      const page = await request('/Examination/OtherExaminationsDashboard/?X-Requested-With=XMLHttpRequest&_=' + Date.now(), { method: 'GET' });
      const found = window.McbCore.parseDashboardIds(page);
      if (!found.orgId || !found.academicYearId) throw new Error(lost);
      ids = found;
      return ids;
    }

    const api = {
      connected: true,
      async listBranches() {
        const { orgId } = await start();
        return window.McbCore.parseBranches(await json('/MasterData/GetBranchLocationsForSearchForstudent', { form: { OrganisationID: orgId } }));
      },
      async listClasses(branchId) {
        return window.McbCore.parseClasses(await json('/AzureClass/GetClassesbyAzure', { form: { BranchID: branchId } }));
      },
      async listSections(branchId, classId) {
        const key = branchId + '|' + classId;
        if (!sectionsOf.has(key)) {
          const { academicYearId } = await start();
          sectionsOf.set(key, window.McbCore.parseSections(await json('/AzureClass/GetSectionsbyAzure?ClassID=' + enc(classId) +
            '&BranchID=' + enc(branchId) + '&AcademicYearID=' + enc(academicYearId))));
        }
        return sectionsOf.get(key);
      },
      async listExams(branchId, classId) {
        // MCB lists examinations per section; the first section stands for the class.
        const { academicYearId } = await start();
        const sections = await api.listSections(branchId, classId);
        if (!sections.length) return [];
        return window.McbCore.parseExams(await json('/Examination/GetExamsByBranchSectionByAcademicYearID', {
          form: { classId, BranchSectionID: sections[0].id, AcademicYearID: academicYearId, BranchID: branchId },
        }));
      },
      async listSubjects(pick) {
        const list = window.McbCore.parseSubjectList(await json('/Examination/GetSubjectsByExaminationByBranchID?ExaminationId=' +
          enc(pick.examId) + '&ClassID=' + enc(pick.classId) + '&BranchID=' + enc(pick.branchId)));
        if (!list.length) return [];
        // The maximum marks only come with a section's student list.
        const sections = await api.listSections(pick.branchId, pick.classId);
        const info = sections.length
          ? await api.getSection({ branchId: pick.branchId, examId: pick.examId, sectionId: sections[0].id, subjects: list })
          : { subjects: [] };
        return list.map((s) => {
          const live = info.subjects.find((x) => x.id === s.id);
          return { id: s.id, name: s.name, max: live && live.max > 0 ? live.max : null };
        });
      },
      async getSection(ctx) {
        const { academicYearId } = await start();
        return window.McbCore.parseSectionInfo(await json('/Examination/GetStrudents', {
          form: {
            chksub: window.McbCore.studentsQuerySubjects(ctx.subjects), ExaminationId: ctx.examId,
            BranchSectionID: ctx.sectionId, AcademicYearID: academicYearId, BranchID: ctx.branchId,
          },
        }));
      },
      async uploadMarks(ctx, subject, students, marks) {
        const { academicYearId } = await start();
        const data = new FormData();
        const csv = window.McbCore.buildMcbCsv(subject, students, marks);
        data.append('FileUpload2', new Blob([csv], { type: 'text/csv' }), 'BulkUploadExamMarks.csv');
        const fields = window.McbCore.uploadFields(Object.assign({}, ctx, { academicYearId }), subject);
        Object.keys(fields).forEach((name) => data.append(name, fields[name]));
        const outcome = window.McbCore.uploadOutcome(await json('/Examination/SaveBulkExaminationMarksTosheet', { body: data }));
        if (!outcome.ok) throw new Error('MCB did not take the file. ' + (outcome.message || 'It gave no reason.'));
      },
    };
    return api;
  }
  const api = fakeApi || createMcbApi();

  // ---------- loading the two libraries ----------

  function loadScript(url) {
    return new Promise((resolve, reject) => {
      const tag = document.createElement('script');
      tag.src = url;
      tag.onload = resolve;
      tag.onerror = () => reject(new Error('Could not load ' + url));
      document.head.appendChild(tag);
    });
  }

  async function loadLibraries() {
    // A page with an AMD loader would take the libraries for its own modules, so hide it while they load.
    const amd = window.define;
    if (typeof amd === 'function' && amd.amd) window.define = undefined;
    try {
      if (!window.ExcelJS) await loadScript(BASE + 'vendor/exceljs.min.js');
      if (!window.McbCore) await loadScript(BASE + 'core.js?v=' + Date.now());
    } finally {
      if (amd) window.define = amd;
    }
  }

  // ---------- tiny DOM helpers (all text goes in as text, never as HTML) ----------

  function el(tag, props, ...kids) {
    const node = document.createElement(tag);
    Object.entries(props || {}).forEach(([key, value]) => {
      if (key === 'class') node.className = value;
      else if (key.startsWith('on')) node.addEventListener(key.slice(2), value);
      else if (key in node) node[key] = value;
      else node.setAttribute(key, value);
    });
    kids.flat().forEach((kid) => { if (kid != null && kid !== false) node.append(kid); });
    return node;
  }

  function say(box, text, kind) {
    box.className = 'status ' + (kind || '');
    box.textContent = text || '';
  }

  function saveFile(name, blob) {
    const url = URL.createObjectURL(blob);
    const link = el('a', { href: url, download: name });
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
  }

  const CSS = `
    :host { all: initial; }
    * { box-sizing: border-box; }
    .panel { position: fixed; top: 0; right: 0; width: 440px; max-width: 100vw; height: 100vh; z-index: 2147483647;
      display: flex; flex-direction: column; background: #fff; color: #1c2430; border-left: 1px solid #c9d1dc;
      box-shadow: -6px 0 24px rgba(20, 30, 50, .18); font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
    header { display: flex; align-items: center; gap: 8px; padding: 12px 14px; background: #12233f; color: #fff; }
    header b { font-size: 15px; } header small { opacity: .7; flex: 1; }
    header button { background: none; border: 0; color: #fff; font-size: 20px; cursor: pointer; line-height: 1; }
    nav { display: flex; border-bottom: 1px solid #dde3ea; }
    nav button { flex: 1; padding: 10px 6px; border: 0; background: #f3f6fa; cursor: pointer; font: inherit; color: #45536a;
      border-bottom: 3px solid transparent; }
    nav button.on { background: #fff; color: #12233f; font-weight: 600; border-bottom-color: #1565d8; }
    main { flex: 1; overflow: auto; padding: 14px; }
    .note { background: #fff6dd; border: 1px solid #ecd28a; padding: 8px 10px; border-radius: 6px; margin-bottom: 12px; }
    label.field { display: block; margin-bottom: 10px; font-weight: 600; font-size: 12px; color: #45536a; }
    select, input[type=text], input[type=url], input[type=number] { display: block; width: 100%; margin-top: 3px; padding: 7px 8px;
      border: 1px solid #b9c3d0; border-radius: 6px; font: inherit; font-weight: 400; color: inherit; background: #fff; }
    .subject { border: 1px solid #dde3ea; border-radius: 6px; padding: 8px 10px; margin-bottom: 8px; }
    .subject label { font-weight: 600; } .subject input[type=number] { margin-top: 6px; }
    .subject label.splitline { display: block; font-weight: 400; font-size: 13px; margin-top: 6px; }
    .parts { margin-top: 6px; padding: 8px 8px 6px; background: #f3f6fa; border-radius: 6px; }
    .partrow { display: grid; grid-template-columns: 1fr 84px 58px; gap: 6px; align-items: center; margin-bottom: 6px; }
    .partrow.head { font-size: 12px; font-weight: 600; color: #5c6b82; margin-bottom: 3px; }
    .subject .partrow input { margin-top: 0; }
    .hint { font-size: 12px; color: #5c6b82; margin-top: 4px; } .hint.bad { color: #b3261e; } .hint.good { color: #1a7f37; }
    button.go { padding: 9px 14px; border: 0; border-radius: 6px; background: #1565d8; color: #fff; font: inherit; font-weight: 600; cursor: pointer; }
    button.go:disabled { background: #9db3d3; cursor: default; }
    button.plain { background: none; border: 0; color: #1565d8; cursor: pointer; font: inherit; padding: 0; text-decoration: underline; }
    .status { margin: 10px 0; min-height: 1.4em; white-space: pre-wrap; } .status.bad { color: #b3261e; } .status.good { color: #1a7f37; }
    .or { text-align: center; color: #7a8799; font-size: 12px; margin: 8px 0; }
    h3 { margin: 14px 0 6px; font-size: 14px; }
    details { border: 1px solid #dde3ea; border-radius: 6px; margin-bottom: 6px; }
    summary { padding: 8px 10px; cursor: pointer; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
    summary b { flex: 0 0 auto; } summary .state { flex: 1 1 100%; font-size: 12px; }
    .tag { font-size: 12px; padding: 1px 7px; border-radius: 10px; background: #e8eef7; color: #2a3b55; }
    .tag.bad { background: #fde7e5; color: #b3261e; } .tag.warn { background: #fff1cc; color: #7a5200; } .tag.good { background: #dff3e4; color: #1a7f37; }
    .body { padding: 0 10px 8px; font-size: 13px; } .body p { margin: 6px 0 2px; font-weight: 600; }
    .body ul { margin: 0; padding-left: 18px; } .body li.bad { color: #b3261e; }
    .row { display: flex; gap: 12px; align-items: center; margin-top: 10px; flex-wrap: wrap; }
  `;

  // ---------- the panel ----------

  function start() {
    const core = window.McbCore;
    const host = el('div', { id: 'mcb-updater' });
    const shadow = host.attachShadow({ mode: 'open' });
    shadow.append(el('style', {}, CSS));

    const log = [];
    const note = (text) => log.push(new Date().toISOString() + '  ' + text);

    // ----- tab 1: make template -----
    const lists = { branches: [], classes: [], exams: [] };
    let subjects = [];
    const branchSel = el('select', { onchange: () => guard(tStatus, onBranch) });
    const classSel = el('select', { onchange: () => guard(tStatus, onClass) });
    const examSel = el('select', { onchange: () => guard(tStatus, onExam) });
    const subjectBox = el('div');
    const tStatus = el('div', { class: 'status' });
    const makeBtn = el('button', { class: 'go', onclick: () => guard(tStatus, makeTemplate, makeBtn) }, 'Download template');

    function fillSelect(select, items, placeholder) {
      select.replaceChildren(el('option', { value: '' }, placeholder), ...items.map((i) => el('option', { value: String(i.id) }, i.label || i.name)));
    }
    const chosen = (select, items) => items.find((i) => String(i.id) === select.value) || null;

    async function loadBranches() {
      say(tStatus, 'Reading your branches from MCB…');
      lists.branches = await api.listBranches();
      fillSelect(branchSel, lists.branches, 'Choose a branch');
      fillSelect(classSel, [], 'Choose a branch first');
      fillSelect(examSel, [], 'Choose a class first');
      say(tStatus, '');
      if (lists.branches.length === 1) { branchSel.value = String(lists.branches[0].id); await onBranch(); }
    }
    async function onBranch() {
      subjects = []; subjectBox.replaceChildren();
      fillSelect(examSel, [], 'Choose a class first');
      lists.classes = branchSel.value ? await api.listClasses(branchSel.value) : [];
      fillSelect(classSel, lists.classes, 'Choose a class');
    }
    async function onClass() {
      subjects = []; subjectBox.replaceChildren();
      lists.exams = classSel.value ? await api.listExams(branchSel.value, classSel.value) : [];
      fillSelect(examSel, lists.exams, lists.exams.length ? 'Choose an examination' : 'No examinations for this class');
    }
    async function onExam() {
      subjects = []; subjectBox.replaceChildren();
      if (!examSel.value) return;
      const found = await api.listSubjects({ branchId: branchSel.value, classId: classSel.value, examId: examSel.value });
      subjects = found.map((s) => {
        const item = { id: s.id, name: s.name, max: s.max, rows: [] };
        item.tick = el('input', { type: 'checkbox', onchange: () => checkSubject(item) });
        item.maxBox = s.max == null
          ? el('input', { type: 'number', min: '1', step: 'any', placeholder: 'Maximum marks', oninput: () => checkSubject(item) })
          : null;
        item.split = el('input', { type: 'checkbox', onchange: () => {
          if (item.split.checked) {
            item.tick.checked = true;
            while (item.rows.length < 3) addPartRow(item);
          }
          item.partsWrap.hidden = !item.split.checked;
          checkSubject(item);
        } });
        item.rowsBox = el('div');
        item.partsWrap = el('div', { class: 'parts', hidden: true },
          el('div', { class: 'partrow head' }, el('span', {}, 'Part'), el('span', {}, 'Maximum'), el('span', {})),
          item.rowsBox,
          el('button', { class: 'plain', onclick: () => { addPartRow(item); checkSubject(item); } }, '+ Add a part'));
        item.hint = el('div', { class: 'hint' });
        subjectBox.append(el('div', { class: 'subject' },
          el('label', {}, item.tick, ' ', s.name, s.max == null ? '' : ' (max ' + s.max + ')'),
          item.maxBox,
          el('label', { class: 'splitline' }, item.split, ' Split into parts, for example Physics, Chemistry, Biology'),
          item.partsWrap, item.hint));
        return item;
      });
      if (!subjects.length) say(tStatus, 'MCB lists no subjects for this examination.', 'bad');
    }

    /** One part of a subject: a name box, a maximum box, and a way to take it out again. */
    function addPartRow(item) {
      const row = {
        name: el('input', { type: 'text', placeholder: 'Part name', oninput: () => checkSubject(item) }),
        max: el('input', { type: 'number', min: '0', step: 'any', placeholder: 'Max', oninput: () => checkSubject(item) }),
      };
      row.node = el('div', { class: 'partrow' }, row.name, row.max,
        el('button', { class: 'plain', title: 'Remove this part', onclick: () => {
          item.rows.splice(item.rows.indexOf(row), 1);
          row.node.remove();
          checkSubject(item);
        } }, 'Remove'));
      item.rows.push(row);
      item.rowsBox.append(row.node);
    }

    /** The subject as it goes into the template, or the reason it cannot. */
    function checkSubject(item) {
      const parts = [];
      let rowProblem = null;
      if (item.split.checked) {
        for (const row of item.rows) {
          const name = row.name.value.trim();
          const max = row.max.value.trim();
          if (!name && !max) continue; // a row left empty is not a part
          if (!name) rowProblem = rowProblem || 'Give a name to the part with maximum ' + max + '.';
          else if (!(Number(max) > 0)) rowProblem = rowProblem || 'Give ' + name + ' a maximum above 0.';
          else parts.push({ name, max: Number(max) });
        }
        if (!rowProblem && !parts.length) rowProblem = 'Fill in the parts, or untick "Split into parts".';
      }
      let max = item.max;
      if (max == null && item.maxBox.value) max = Number(item.maxBox.value);
      if (max == null && parts.length) max = parts.reduce((sum, p) => sum + Math.round(p.max * 100), 0) / 100;
      const problem = rowProblem || core.checkParts(parts, max) || (max > 0 ? null : 'Give the maximum marks.');
      item.hint.className = 'hint ' + (problem ? 'bad' : 'good');
      item.hint.textContent = problem || (parts.length
        ? 'Adds up to ' + max + ' of ' + max + '. The sheet gets a column for ' + parts.map((p) => p.name + ' (' + p.max + ')').join(', ') +
          ', and a TOTAL column that adds them up. Only the total goes to MCB.'
        : 'One column, marks out of ' + max + '.');
      return { problem, subject: { id: item.id, name: item.name, max, parts } };
    }

    async function makeTemplate() {
      const branch = chosen(branchSel, lists.branches);
      const cls = chosen(classSel, lists.classes);
      const exam = chosen(examSel, lists.exams);
      if (!branch || !cls || !exam) throw new Error('Choose the branch, class and examination first.');
      const picked = subjects.filter((s) => s.tick.checked).map(checkSubject);
      if (!picked.length) throw new Error('Tick at least one subject.');
      const broken = picked.find((p) => p.problem);
      if (broken) throw new Error(broken.subject.name + ': ' + broken.problem);
      const wanted = picked.map((p) => p.subject);

      say(tStatus, 'Reading the sections of ' + cls.name + '…');
      const sections = await api.listSections(branch.id, cls.id);
      const filled = [];
      const empty = [];
      for (let i = 0; i < sections.length; i += 1) {
        const section = sections[i];
        say(tStatus, 'Reading the students of ' + section.name + '… (' + (i + 1) + ' of ' + sections.length + ')');
        const info = await api.getSection({
          branchId: branch.id, classId: cls.id, examId: exam.id, examName: exam.name, sectionId: section.id,
          subjects: wanted.map((s) => ({ id: s.id, name: s.name })),
        });
        if (info.gradeOnly) throw new Error('MCB takes grades, not marks, for this examination. That is not supported.');
        const students = info.students;
        if (students.length) filled.push({ id: section.id, name: section.name, students });
        else empty.push(section.name);
      }
      if (!filled.length) throw new Error('MCB returned no students for ' + cls.name + '.');

      const workbook = await core.buildWorkbook(window.ExcelJS, {
        host: location.hostname, branch, cls, exam: { id: exam.id, name: exam.name }, subjects: wanted, sections: filled,
      });
      const bytes = await workbook.xlsx.writeBuffer();
      const fileName = [cls.name, exam.name, wanted.map((s) => s.name).join(' + ')].join(' - ').replace(/[\\/:*?"<>|]+/g, ' ') + '.xlsx';
      saveFile(fileName, new Blob([bytes], { type: XLSX_MIME }));
      const count = filled.reduce((sum, s) => sum + s.students.length, 0);
      note('Template made: ' + fileName + ', ' + filled.length + ' sections, ' + count + ' students.');
      say(tStatus, 'Downloaded "' + fileName + '": ' + filled.length + ' sections, ' + count + ' students.' +
        (empty.length ? '\nNo students in: ' + empty.join(', ') + '.' : '') +
        '\nUpload it to Google Drive, open it with Google Sheets, and fill in the marks.', 'good');
    }

    // ----- tab 2: upload marks -----
    let plan = null; // { meta, items: [{ section, judged, summary, ctx, info, marks, diff, box, state }] }
    const fileBox = el('input', { type: 'file', accept: '.xlsx', onchange: () => { linkBox.value = ''; } });
    const linkBox = el('input', { type: 'url', placeholder: 'https://docs.google.com/spreadsheets/d/…' });
    const readBtn = el('button', { class: 'go', onclick: () => guard(uStatus, readChosenSheet, readBtn) }, 'Check the sheet');
    const uStatus = el('div', { class: 'status' });
    const previewBox = el('div');
    const goBtn = el('button', { class: 'go', disabled: true, onclick: () => guard(uStatus, upload, goBtn) }, 'Upload');
    const firstOnly = el('button', { class: 'plain', hidden: true, onclick: tickFirstOnly }, 'Tick only the first section');
    const logBtn = el('button', { class: 'plain', hidden: true, onclick: () =>
      saveFile('mcb-updater-log.txt', new Blob([log.join('\n') + '\n'], { type: 'text/plain' })) }, 'Download the log');

    async function readChosenSheet() {
      let data;
      if (fileBox.files && fileBox.files[0]) {
        data = await fileBox.files[0].arrayBuffer();
        note('Sheet chosen from a file.');
      } else if (linkBox.value.trim()) {
        const url = core.googleSheetExportUrl(linkBox.value);
        if (!url) throw new Error('That is not a Google Sheets link.');
        say(uStatus, 'Downloading the sheet from Google…');
        const help = 'Could not open that link. In Google Sheets use Share → General access → "Anyone with the link", ' +
          'or use File → Download → Microsoft Excel (.xlsx) and choose that file here.';
        let response;
        try { response = await fetch(url, { credentials: 'omit' }); } catch (e) { throw new Error(help); }
        if (!response.ok || !/spreadsheetml/.test(response.headers.get('content-type') || '')) throw new Error(help);
        data = await response.arrayBuffer();
        note('Sheet downloaded from a Google Sheets link.');
      } else {
        throw new Error('Choose the filled .xlsx file, or paste the Google Sheets link.');
      }
      await readBuffer(data);
    }

    async function readBuffer(data) {
      plan = null;
      previewBox.replaceChildren();
      goBtn.disabled = true; firstOnly.hidden = true;
      const { meta, sections } = await core.readWorkbook(window.ExcelJS, data);
      if (!fakeApi && !onLocalTest && meta.host && meta.host !== location.hostname) {
        throw new Error('This sheet was made on ' + meta.host + ', but this page is ' + location.hostname + '.');
      }
      const items = [];
      for (let i = 0; i < sections.length; i += 1) {
        const section = sections[i];
        say(uStatus, 'Checking ' + section.name + ' against MCB… (' + (i + 1) + ' of ' + sections.length + ')');
        const judged = core.evaluateSection(section, meta);
        const item = { section, judged, ctx: null, info: null, marks: {}, diff: {} };
        if (!judged.errors.length) {
          item.ctx = {
            branchId: meta.branch.id, classId: meta.cls.id, examId: meta.exam.id, examName: meta.exam.name, sectionId: section.id,
            subjects: meta.subjects.map((s) => ({ id: s.id, name: s.name })),
          };
          item.info = await api.getSection(item.ctx);
          core.reconcileSection(judged, item.info.students);
          meta.subjects.forEach((s) => {
            item.marks[s.id] = core.marksFor(judged, s.id);
            item.diff[s.id] = core.diffMarks(item.marks[s.id], item.info.saved[s.id]);
            if (!item.marks[s.id].size) return;
            const live = item.info.subjects.find((x) => x.id === s.id);
            if (item.info.gradeOnly) judged.errors.push('MCB takes grades, not marks, for this examination.');
            else if (!live) judged.errors.push('MCB does not list ' + s.name + ' for this section and examination.');
            else if (core.entryClosed(live.lastDate)) judged.errors.push('MCB stopped taking ' + s.name + ' marks on ' + live.lastDate + '.');
            else if (live.max !== s.max) judged.errors.push('MCB now has ' + live.max + ' as the maximum for ' + s.name + ', but this sheet was made for ' + s.max + '.');
          });
        }
        item.summary = core.summarizeSection(judged, meta.subjects);
        items.push(item);
      }
      plan = { meta, items };
      note('Sheet checked: ' + meta.cls.name + ', ' + meta.exam.name + ', ' + items.length + ' sections.');
      renderPreview();
    }

    function renderPreview() {
      const { meta, items } = plan;
      previewBox.replaceChildren(el('h3', {}, meta.cls.name + ' · ' + meta.exam.name + ' · ' + meta.subjects.map((s) => s.name).join(', ')));
      for (const item of items) {
        const sum = item.summary;
        const usable = !sum.blocked && sum.toWrite > 0;
        item.box = el('input', { type: 'checkbox', checked: usable, disabled: !usable, onclick: (e) => e.stopPropagation(), onchange: refreshGo });
        item.state = el('span', { class: 'state' });
        const nameOf = new Map(item.judged.rows.map((r) => [r.code, r.name]));

        const errors = item.judged.errors.map((text) => el('li', { class: 'bad' }, text));
        const warnings = item.judged.warnings.map((text) => el('li', {}, text));
        for (const row of item.judged.rows) {
          const who = 'Row ' + row.row + (row.name ? ', ' + row.name : '') + ': ';
          row.errors.forEach((text) => errors.push(el('li', { class: 'bad' }, who + text)));
          row.warnings.forEach((text) => warnings.push(el('li', {}, who + text)));
        }
        const changes = [];
        let same = 0;
        let known = false;
        meta.subjects.forEach((s) => {
          const diff = item.diff[s.id];
          if (!diff) return;
          known = true;
          same += diff.same;
          diff.changed.forEach((c) => changes.push(el('li', {},
            (nameOf.get(c.code) || c.code) + (meta.subjects.length > 1 ? ', ' + s.name : '') + ': ' + c.from + ' → ' + c.to)));
        });

        previewBox.append(el('details', { open: sum.blocked },
          el('summary', {}, item.box, el('b', {}, item.section.name),
            sum.blocked
              ? el('span', { class: 'tag bad' }, sum.errors + ' error' + (sum.errors > 1 ? 's' : '') + ', not uploaded')
              : el('span', { class: 'tag good' }, sum.toWrite + ' to upload'),
            sum.absent ? el('span', { class: 'tag' }, sum.absent + ' absent') : null,
            sum.blank ? el('span', { class: 'tag' }, sum.blank + ' not entered') : null,
            changes.length ? el('span', { class: 'tag warn' }, changes.length + ' replace a mark in MCB') : null,
            sum.warnings ? el('span', { class: 'tag warn' }, sum.warnings + ' to look at') : null,
            item.state),
          el('div', { class: 'body' },
            errors.length ? [el('p', {}, 'Fix these in the sheet, then check it again'), el('ul', {}, errors)] : null,
            changes.length ? [el('p', {}, 'MCB already has a different mark for'), el('ul', {}, changes)] : null,
            warnings.length ? [el('p', {}, 'Worth a look'), el('ul', {}, warnings)] : null,
            known && same ? el('p', {}, same + ' mark(s) are already the same in MCB.') : null,
            !errors.length && !changes.length && !warnings.length ? el('p', {}, 'Nothing to flag.') : null)));
      }
      refreshGo();
      const blocked = items.filter((i) => i.summary.blocked).length;
      say(uStatus, blocked
        ? blocked + ' section(s) have errors and will not be uploaded until the sheet is fixed.'
        : 'The sheet is fine. Review the sections, then upload.', blocked ? 'bad' : 'good');
    }

    const ticked = () => (plan ? plan.items.filter((i) => i.box && i.box.checked) : []);
    function refreshGo() {
      const items = ticked();
      const marks = items.reduce((sum, i) => sum + i.summary.toWrite, 0);
      goBtn.disabled = !items.length;
      goBtn.textContent = items.length ? 'Upload ' + marks + ' marks in ' + items.length + ' section' + (items.length > 1 ? 's' : '') : 'Upload';
      firstOnly.hidden = !plan || plan.items.filter((i) => i.box && !i.box.disabled).length < 2;
    }
    function tickFirstOnly() {
      let done = false;
      plan.items.forEach((i) => { if (!i.box.disabled) { i.box.checked = !done; done = true; } });
      refreshGo();
    }

    async function upload() {
      const { meta } = plan;
      const items = ticked();
      const marks = items.reduce((sum, i) => sum + i.summary.toWrite, 0);
      const question = 'Write ' + marks + ' marks into MCB?\n\n' + meta.cls.name + ' · ' + meta.exam.name + '\n' +
        items.map((i) => i.section.name).join(', ') + '\n\nMarks already in MCB for these students are replaced.';
      if (!window.confirm(question)) return;
      logBtn.hidden = false;
      let ok = 0;
      for (const item of items) {
        const setState = (text, kind) => { item.state.textContent = text; item.state.style.color = kind === 'bad' ? '#b3261e' : kind === 'good' ? '#1a7f37' : ''; };
        item.box.disabled = true;
        try {
          setState('Uploading…');
          let sent = 0;
          for (const s of meta.subjects) {
            const marks = item.marks[s.id];
            if (!marks.size) continue;
            await api.uploadMarks(item.ctx, item.info.subjects.find((x) => x.id === s.id), item.info.students, marks);
            sent += marks.size;
            await sleep(300);
          }
          setState('Checking what MCB saved…');
          const after = await api.getSection(item.ctx);
          const wrong = [];
          meta.subjects.forEach((s) => wrong.push(...core.verifyMarks(item.marks[s.id], after.saved[s.id])));
          if (wrong.length) {
            throw new Error(wrong.length + ' of ' + sent + ' mark(s) in MCB differ from the sheet: ' +
              wrong.map((w) => w.code + ' (sent ' + w.sent + ', MCB has ' + (w.found || 'nothing') + ')').join(', ') + '.');
          }
          setState('Saved and checked: ' + sent + ' marks.', 'good');
          note(item.section.name + ': saved and checked, ' + sent + ' marks.');
          item.box.checked = false;
          ok += 1;
        } catch (error) {
          setState('Stopped: ' + error.message, 'bad');
          note(item.section.name + ': FAILED. ' + error.message);
          refreshGo();
          throw new Error('Stopped at ' + item.section.name + '. ' + ok + ' section(s) before it were uploaded and checked; the sections after it were not sent.');
        }
        await sleep(400); // go easy on MCB
      }
      refreshGo();
      say(uStatus, 'Done: ' + ok + ' section(s) uploaded.', 'good');
    }

    // ----- shell -----
    async function guard(statusBox, work, button) {
      if (button) button.disabled = true;
      try {
        await work();
      } catch (error) {
        say(statusBox, error.message, 'bad');
      } finally {
        if (button && button !== goBtn) button.disabled = false;
        if (button === goBtn) refreshGo();
      }
    }

    const templateTab = el('main', {},
      el('label', { class: 'field' }, 'Branch', branchSel),
      el('label', { class: 'field' }, 'Class', classSel),
      el('label', { class: 'field' }, 'Examination', examSel),
      subjectBox, tStatus, makeBtn,
      el('div', { class: 'hint' }, 'You get one Excel file with a tab for every section of the class.'));
    const uploadTab = el('main', { hidden: true },
      el('label', { class: 'field' }, 'The filled sheet, as an .xlsx file', fileBox),
      el('div', { class: 'or' }, 'or'),
      el('label', { class: 'field' }, 'A Google Sheets link (shared as "Anyone with the link")', linkBox),
      readBtn, uStatus, previewBox,
      el('div', { class: 'row' }, goBtn, firstOnly, logBtn));

    const tabButtons = [
      el('button', { class: 'on', onclick: () => show(0) }, '1 · Make template'),
      el('button', { onclick: () => show(1) }, '2 · Upload marks'),
    ];
    function show(index) {
      tabButtons.forEach((b, i) => b.classList.toggle('on', i === index));
      templateTab.hidden = index !== 0;
      uploadTab.hidden = index !== 1;
    }
    const panel = el('div', { class: 'panel' },
      el('header', {}, el('b', {}, 'MCB marks updater'), el('small', {}, 'v' + core.VERSION),
        el('button', { title: 'Close', onclick: () => toggle() }, '×')),
      el('nav', {}, tabButtons), templateTab, uploadTab);
    shadow.append(panel);
    document.body.appendChild(host);

    function toggle() { host.style.display = host.style.display === 'none' ? '' : 'none'; }
    window.__mcbUpdater = {
      toggle,
      // test/harness.html hands a sheet straight in, without the file picker
      readBuffer: (data) => { show(1); return guard(uStatus, () => readBuffer(data)); },
    };
    guard(tStatus, loadBranches);
  }

  loadLibraries().then(start).catch((error) => alert('MCB marks updater could not start: ' + error.message));
})();
