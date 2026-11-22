/*
 * mcb-updater core: everything that needs no page and no network.
 * Marks arithmetic, sheet checks, MCB's request and response formats, and the .xlsx template (ExcelJS is passed in).
 * Runs in the browser (window.McbCore) and in Node (require) so it can be tested.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.McbCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const VERSION = '0.1.0';
  const ABSENT_CODES = ['A', 'AB', 'L', 'AM']; // what MCB's marks box accepts for an absent student
  const META_SHEET = '_meta';
  const META_TAG = 'mcb-updater';
  const HEAD_SNO = 'S.No';
  const HEAD_CODE = 'Enrollment code';
  const HEAD_NAME = 'Name';

  // ---------- small helpers ----------

  const norm = (s) => String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toUpperCase();
  const normCode = (s) => String(s == null ? '' : s).replace(/\s+/g, '').toUpperCase();
  // Marks are added in hundredths so 3.5 + 6.1 never becomes 9.600000000000001.
  const toHundredths = (n) => Math.round(Number(n) * 100);
  const formatMark = (hundredths) => String(hundredths / 100);

  /** What one sheet cell holds: blank, a number, an absent code, or something that is not a mark. */
  function parseCell(raw) {
    if (raw == null) return { kind: 'blank' };
    if (typeof raw === 'number') {
      return Number.isFinite(raw) ? { kind: 'num', value: raw } : { kind: 'bad', text: String(raw) };
    }
    const text = String(raw).trim();
    if (text === '') return { kind: 'blank' };
    const upper = text.toUpperCase();
    if (ABSENT_CODES.includes(upper)) return { kind: 'absent', code: upper };
    if (/^\d+(\.\d+)?$/.test(text)) return { kind: 'num', value: Number(text) };
    return { kind: 'bad', text };
  }

  /** Two marks are the same when they are the same number, or the same absent code. */
  function sameMark(a, b) {
    const x = parseCell(a);
    const y = parseCell(b);
    if (x.kind !== y.kind) return false;
    if (x.kind === 'num') return toHundredths(x.value) === toHundredths(y.value);
    if (x.kind === 'absent') return x.code === y.code;
    if (x.kind === 'blank') return true;
    return norm(x.text) === norm(y.text);
  }

  // ---------- subject parts ("Physics 7, Chemistry 6, Biology 7") ----------

  /** Reads the parts text box. Returns { parts: [{name, max}], errors: [text] }. Empty text = no parts. */
  function parseParts(text) {
    const parts = [];
    const errors = [];
    const pieces = String(text || '').split(/[,;\n]+/).map((p) => p.trim()).filter(Boolean);
    for (const piece of pieces) {
      const m = piece.match(/^(.*?)[\s(=:-]+(\d+(?:\.\d+)?)\s*\)?$/);
      if (!m || !m[1].trim()) {
        errors.push('"' + piece + '" should look like "Physics 7".');
        continue;
      }
      parts.push({ name: m[1].trim(), max: Number(m[2]) });
    }
    return { parts, errors };
  }

  /** Why these parts cannot be used for this subject, or null when they are fine. */
  function checkParts(parts, subjectMax) {
    if (!parts.length) return null;
    if (parts.length === 1) return 'One part is the same as no parts. Give two or more, or leave it empty.';
    const seen = new Set();
    let sum = 0;
    for (const p of parts) {
      if (!(p.max > 0)) return p.name + ' needs a maximum above 0.';
      if (seen.has(norm(p.name))) return p.name + ' is listed twice.';
      seen.add(norm(p.name));
      sum += toHundredths(p.max);
    }
    if (subjectMax != null && sum !== toHundredths(subjectMax)) {
      return 'The parts add up to ' + formatMark(sum) + ', but MCB expects ' + subjectMax + '.';
    }
    return null;
  }

  /** The columns one subject takes in a section tab. */
  function subjectColumns(subject, manySubjects) {
    const parts = subject.parts || [];
    if (!parts.length) {
      return { columns: [subject.name + ' (' + subject.max + ')'], totalHeader: null };
    }
    const prefix = manySubjects ? subject.name + ': ' : '';
    return {
      columns: parts.map((p) => prefix + p.name + ' (' + p.max + ')'),
      totalHeader: subject.name + ' TOTAL (' + subject.max + ')',
    };
  }

  // ---------- one student's mark for one subject ----------

  /**
   * rawCells: the student's cells for this subject, in column order.
   * Returns { status: 'ok' | 'absent' | 'blank' | 'error', mark, errors, warnings }.
   * mark is the text MCB gets: "14.5", or an absent code.
   */
  function computeSubject(rawCells, subject) {
    const cols = subject.parts && subject.parts.length
      ? subject.parts
      : [{ name: subject.name, max: subject.max }];
    const cells = cols.map((_, i) => parseCell(rawCells ? rawCells[i] : null));
    const errors = [];
    const warnings = [];
    if (cells.every((c) => c.kind === 'blank')) return { status: 'blank', mark: null, errors, warnings };

    const absent = [];
    let numbers = 0;
    let total = 0;
    cells.forEach((cell, i) => {
      const col = cols[i];
      if (cell.kind === 'bad') {
        errors.push(col.name + ': "' + cell.text + '" is not a mark.');
      } else if (cell.kind === 'absent') {
        absent.push(cell.code);
      } else if (cell.kind === 'num') {
        numbers += 1;
        if (cell.value < 0) errors.push(col.name + ': ' + cell.value + ' is below 0.');
        else if (col.max != null && toHundredths(cell.value) > toHundredths(col.max)) {
          errors.push(col.name + ': ' + cell.value + ' is more than the maximum ' + col.max + '.');
        }
        total += toHundredths(cell.value);
      }
    });

    if (absent.length) {
      if (numbers) errors.push('An absent code and marks are in the same row.');
      else if (new Set(absent).size > 1) errors.push('Different absent codes in the same row.');
      if (errors.length) return { status: 'error', mark: null, errors, warnings };
      return { status: 'absent', mark: absent[0], errors, warnings };
    }
    if (errors.length) return { status: 'error', mark: null, errors, warnings };

    cells.forEach((cell, i) => {
      if (cell.kind === 'blank') warnings.push(cols[i].name + ' is blank, counted as 0.');
    });
    if (subject.max != null && total > toHundredths(subject.max)) {
      return {
        status: 'error', mark: null, warnings,
        errors: ['Total ' + formatMark(total) + ' is more than the maximum ' + subject.max + '.'],
      };
    }
    return { status: 'ok', mark: formatMark(total), total: total / 100, errors, warnings };
  }

  // ---------- checking a sheet that was read back ----------

  /**
   * section: one entry of readWorkbook().sections. Returns the section with every row judged:
   * { id, name, tab, errors, warnings, rows: [{ row, code, name, results: {subjectId: computeSubject()}, errors, warnings }] }
   */
  function evaluateSection(section, meta) {
    const out = { id: section.id, name: section.name, tab: section.tab, errors: [], warnings: [], rows: [] };
    if (!section.found) {
      out.errors.push('The tab "' + section.tab + '" is missing. Was it renamed or deleted?');
      return out;
    }
    out.errors.push(...section.problems);
    if (out.errors.length) return out;

    const firstRowOf = new Map();
    for (const r of section.rows) {
      const entry = { row: r.row, code: normCode(r.code), name: String(r.name || '').trim(), results: {}, errors: [], warnings: [] };
      let hasMarks = false;
      for (const subject of meta.subjects) {
        const res = computeSubject(r.cells[subject.id], subject);
        entry.results[subject.id] = res;
        if (res.status !== 'blank') hasMarks = true;
        entry.errors.push(...res.errors);
        entry.warnings.push(...res.warnings);
        const shown = parseCell(r.totals ? r.totals[subject.id] : null);
        if (res.status === 'ok' && shown.kind === 'num' && toHundredths(shown.value) !== toHundredths(res.total)) {
          entry.warnings.push('The TOTAL cell shows ' + shown.value + ' but the parts add up to ' + res.mark + '. ' + res.mark + ' is used.');
        }
      }
      if (!entry.code) {
        if (!hasMarks) continue; // an empty row, or a name with nothing to upload
        entry.errors.push('This row has marks but no enrollment code.');
      } else if (firstRowOf.has(entry.code)) {
        entry.errors.push('Enrollment code ' + entry.code + ' is also in row ' + firstRowOf.get(entry.code) + '.');
      } else {
        firstRowOf.set(entry.code, r.row);
      }
      out.rows.push(entry);
    }
    return out;
  }

  /** Checks a judged section against MCB's current student list for that section: [{code, name}]. */
  function reconcileSection(evaluated, roster) {
    const inMcb = new Map(roster.map((s) => [normCode(s.code), s]));
    const inSheet = new Set();
    for (const entry of evaluated.rows) {
      if (!entry.code) continue;
      inSheet.add(entry.code);
      if (!inMcb.has(entry.code)) {
        entry.errors.push(entry.code + ' is not in MCB\'s list for ' + evaluated.name + '.');
      }
    }
    const missing = roster.filter((s) => !inSheet.has(normCode(s.code)));
    if (missing.length) {
      evaluated.warnings.push(
        missing.length + ' student(s) in MCB are not in this tab and stay untouched: ' +
        missing.map((s) => s.name + ' (' + s.code + ')').join(', ') + '.'
      );
    }
    return evaluated;
  }

  /** Counts for the preview. A section with any error is blocked: nothing from it is uploaded. */
  function summarizeSection(evaluated, subjects) {
    let toWrite = 0;
    let absent = 0;
    let blank = 0;
    let rowErrors = 0;
    let warnings = evaluated.warnings.length;
    for (const entry of evaluated.rows) {
      if (entry.errors.length) rowErrors += 1;
      warnings += entry.warnings.length;
      for (const subject of subjects) {
        const res = entry.results[subject.id];
        if (!res) continue;
        if (res.status === 'ok') toWrite += 1;
        else if (res.status === 'absent') { toWrite += 1; absent += 1; }
        else if (res.status === 'blank') blank += 1;
      }
    }
    const errors = evaluated.errors.length + rowErrors;
    return { toWrite, absent, blank, errors, warnings, blocked: errors > 0 };
  }

  /** code -> mark text for one subject, only rows that have a mark. */
  function marksFor(evaluated, subjectId) {
    const marks = new Map();
    for (const entry of evaluated.rows) {
      const res = entry.results[subjectId];
      if (entry.code && res && res.mark != null) marks.set(entry.code, res.mark);
    }
    return marks;
  }

  /** Compares what will be written with what MCB already holds (Map code -> mark, may be empty). */
  function diffMarks(next, existing) {
    const out = { fresh: 0, same: 0, changed: [] };
    for (const [code, mark] of next) {
      const old = existing ? existing.get(code) : null;
      if (parseCell(old).kind === 'blank') out.fresh += 1;
      else if (sameMark(old, mark)) out.same += 1;
      else out.changed.push({ code, from: String(old), to: mark });
    }
    return out;
  }

  /** After an upload: which marks in MCB differ from what was sent. Empty list = all good. */
  function verifyMarks(sent, saved) {
    const wrong = [];
    for (const [code, mark] of sent) {
      const now = saved ? saved.get(code) : null;
      if (!sameMark(now, mark)) wrong.push({ code, sent: mark, found: now == null ? '' : String(now) });
    }
    return wrong;
  }

  // ---------- MCB's wire formats ----------
  // Taken from a capture of MCB's own "Bulk Upload Marks/Grades" page. Field names (and their typos) are MCB's.

  /** "/Date(1794162600000)/" -> Date, or null. */
  function mcbDate(value) {
    const m = /\/Date\((-?\d+)\)\//.exec(String(value || ''));
    return m ? new Date(Number(m[1])) : null;
  }

  /** The ids MCB writes into its Other Exams dashboard page: organisation, home branch, academic year. */
  function parseDashboardIds(html) {
    const pick = (name) => {
      const m = String(html).match(new RegExp('var\\s+' + name + "\\s*=\\s*'(\\d+)'"));
      return m ? m[1] : null;
    };
    return { orgId: pick('OrgID'), branchId: pick('BranchID1'), academicYearId: pick('AID') };
  }

  function parseBranches(json) {
    const seen = new Set();
    return (json && json.totRes || []).filter((b) => !seen.has(b.BranchID) && seen.add(b.BranchID))
      .map((b) => ({ id: String(b.BranchID), name: String(b.BranchName).trim() }));
  }

  function parseClasses(json) {
    return (Array.isArray(json) ? json : []).map((c) => ({ id: String(c.ClassID), name: String(c.ClassName).trim() }));
  }

  /** MCB works with BranchSectionID everywhere on this page, so that is the section's id. */
  function parseSections(json) {
    return (Array.isArray(json) ? json : []).filter((s) => s.Active !== 0)
      .map((s) => ({ id: String(s.BranchSectionID), name: String(s.Section).trim() }));
  }

  /** Newest examination first. label carries the date because names repeat. */
  function parseExams(json) {
    const list = (json && Array.isArray(json.dt2) ? json.dt2 : []).map((e) => {
      const date = mcbDate(e.Date);
      const name = String(e.ExamName).trim();
      const day = date ? date.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
      return { id: String(e.ExaminationID), name, label: day ? name + ' (' + day + ')' : name, time: date ? date.getTime() : 0 };
    });
    return list.sort((a, b) => b.time - a.time);
  }

  /** [{Text: "SCIENCE", Value: "36513,-1"}] -> [{id: "36513,-1", name: "SCIENCE"}]. The id is SubjectID,SubjectSkillID. */
  function parseSubjectList(json) {
    return (Array.isArray(json) ? json : []).map((s) => ({ id: String(s.Value), name: String(s.Text).trim() }));
  }

  /** The "chksub" value of GetStrudents: "36513,-1, SCIENCE$36514,-1, SOCIAL SCIENCE". */
  function studentsQuerySubjects(subjects) {
    return subjects.map((s) => s.id + ', ' + s.name).join('$');
  }

  /** True when MCB no longer takes marks for a subject ("11/24/2026 12:00:00 AM" has passed). */
  function entryClosed(lastDate, now) {
    if (!lastDate) return false;
    const end = new Date(lastDate);
    return !Number.isNaN(end.getTime()) && end < (now || new Date());
  }

  /**
   * The answer of /Examination/GetStrudents for one section:
   *   ds1 students, ds2 subjects (a JSON string), ds4 marks already saved, ds5 "grades only" flag.
   * Returns { students: [{code, name}], subjects: [{id, name, max, lastDate, raw}], saved: {subjectId: Map(code -> mark)}, gradeOnly }.
   */
  function parseSectionInfo(json) {
    const list = (value) => {
      if (Array.isArray(value)) return value;
      if (typeof value === 'string' && value.trim().startsWith('[')) {
        try { return JSON.parse(value); } catch (e) { return []; }
      }
      return [];
    };
    const ds1 = list(json && json.ds1);
    const codeOf = new Map(ds1.map((s) => [String(s.StudentEnrollmentID), normCode(s.StudentEnrollmentCode)]));
    const subjects = list(json && json.ds2).map((s) => ({
      id: s.SubjectID + ',' + s.SubjectSkillID,
      name: String(s.SubjectName).trim(),
      max: Number(s.MaxMarks),
      lastDate: s.MarksEntryLastDate || '',
      raw: s,
    }));
    const saved = {};
    subjects.forEach((s) => { saved[s.id] = new Map(); });
    for (const row of list(json && json.ds4)) {
      const marks = saved[row.SubjectID + ',' + row.SubjectSkillID];
      const code = codeOf.get(String(row.StudentEnrollmentID));
      if (!marks || !code) continue;
      const absent = Number(row.MarksAcquired) === -1; // MCB stores an absent student as -1.00
      marks.set(code, absent ? String(row.AbsentType || 'A').trim().toUpperCase() : String(row.MarksAcquired));
    }
    return {
      students: ds1.map((s) => ({
        code: normCode(s.StudentEnrollmentCode),
        name: String(s.StudentName || s.FullName || '').replace(/\s+/g, ' ').trim(),
      })),
      subjects, saved,
      gradeOnly: json != null && (json.ds5 === true || json.ds5 === 'true'),
    };
  }

  /**
   * The file MCB's "Upload marks" tab takes, for one subject, keyed by enrollment code only:
   *   Student EnrollmentCode, SCIENCE
   *   16XX010052,14.5
   * Only students with a mark are listed: MCB rejects a line with an empty mark.
   */
  function buildMcbCsv(subject, students, marks) {
    const lines = ['Student EnrollmentCode,' + subject.raw.SubjectName];
    students.forEach((s) => { if (marks.has(s.code)) lines.push(s.code + ',' + marks.get(s.code)); });
    return lines.join('\r\n') + '\r\n';
  }

  /** The form fields that go with that file to /Examination/SaveBulkExaminationMarksTosheet. */
  function uploadFields(ctx, subject) {
    const raw = subject.raw;
    return {
      slectsubjectCount: '2', // one identity column + one subject
      dlUploadingSubjects: [raw.SubjectID, raw.SubjectSkillID, raw.SubjectName, raw.MinMarks + ' ', raw.MaxMarks, raw.MarksEntryLastDate].join('#'),
      chktext: 'Student EnrollmentCode',
      CheckBoxList1values: 'StudentEnrollmentCode',
      ddlExamnameID: String(ctx.examId),
      subjects: subject.id + '#' + raw.SubjectName,
      ClassID: String(ctx.classId),
      BranchSectionID: String(ctx.sectionId),
      ddlExamname: ctx.examName || '',
      AcademicYearID: String(ctx.academicYearId),
      BranchID: String(ctx.branchId),
    };
  }

  /** MCB's answer to an upload: dt4 is 0 when nothing was saved; dt7 / dt2 carry its message as HTML. */
  function uploadOutcome(json) {
    const text = String((json && (json.dt7 || json.dt2)) || '')
      .replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '')
      .replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
    const lines = text.split('\n').map((l) => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const message = lines.slice(0, 4).join(' | ') + (lines.length > 4 ? ' | and ' + (lines.length - 4) + ' more' : '');
    return { ok: json != null && json.dt4 != null && Number(json.dt4) !== 0, message };
  }

  // ---------- the .xlsx template ----------

  function uniqueTabName(name, used) {
    const base = String(name || 'Section').replace(/[\[\]:*?\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'Section';
    let tab = base;
    for (let n = 2; used.has(tab.toUpperCase()); n += 1) {
      const suffix = ' (' + n + ')';
      tab = base.slice(0, 31 - suffix.length) + suffix;
    }
    used.add(tab.toUpperCase());
    return tab;
  }

  function validationFor(cellRef, max) {
    const codes = ABSENT_CODES.map((c) => cellRef + '="' + c + '"').join(',');
    return {
      type: 'custom',
      allowBlank: true,
      formulae: ['OR(AND(ISNUMBER(' + cellRef + '),' + cellRef + '>=0,' + cellRef + '<=' + max + '),' + codes + ')'],
      showErrorMessage: true,
      errorStyle: 'stop',
      errorTitle: 'Not a valid mark',
      error: 'Enter a number from 0 to ' + max + ', or ' + ABSENT_CODES.join(' / ') + ' for an absent student.',
    };
  }

  /**
   * spec: { host, branch:{id,name}, cls:{id,name}, exam:{id,name},
   *         subjects:[{id,name,max,parts:[{name,max}]}],
   *         sections:[{id,name,students:[{code,name}]}] }
   * Returns an ExcelJS workbook: one tab per section and a hidden _meta tab that says what the file is for.
   */
  async function buildWorkbook(ExcelJS, spec) {
    const wb = new ExcelJS.Workbook();
    wb.creator = META_TAG;
    wb.created = new Date();
    const many = spec.subjects.length > 1;
    const subjects = spec.subjects.map((s) => Object.assign(
      { id: s.id, name: s.name, max: s.max, parts: s.parts || [] },
      subjectColumns(s, many)
    ));
    const usedTabs = new Set([META_SHEET.toUpperCase()]);
    const metaSections = [];
    const grey = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };
    const headFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E8FB' } };
    const thin = { style: 'thin', color: { argb: 'FFBBBBBB' } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    for (const section of spec.sections) {
      const tab = uniqueTabName(section.name, usedTabs);
      metaSections.push({ id: section.id, name: section.name, tab });
      const ws = wb.addWorksheet(tab, { views: [{ state: 'frozen', xSplit: 3, ySplit: 1 }] });

      const headers = [HEAD_SNO, HEAD_CODE, HEAD_NAME];
      const layout = []; // per subject: the column numbers of its cells and of its total
      for (const subject of subjects) {
        const first = headers.length + 1;
        headers.push(...subject.columns);
        const cols = subject.columns.map((_, i) => first + i);
        let totalCol = null;
        if (subject.totalHeader) { headers.push(subject.totalHeader); totalCol = headers.length; }
        layout.push({ subject, cols, totalCol });
      }
      ws.addRow(headers);
      const headRow = ws.getRow(1);
      headRow.height = 32;
      headRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = headFill;
        cell.border = border;
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      });
      ws.getColumn(1).width = 6;
      ws.getColumn(2).width = 18;
      ws.getColumn(3).width = 30;
      for (let c = 4; c <= headers.length; c += 1) ws.getColumn(c).width = 16;

      const students = section.students.slice().sort((a, b) =>
        String(a.name).localeCompare(String(b.name), 'en', { sensitivity: 'base' }));
      students.forEach((student, i) => {
        const r = i + 2;
        const fixed = [i + 1, String(student.code), student.name];
        fixed.forEach((value, c) => {
          const cell = ws.getCell(r, c + 1);
          cell.value = value;
          cell.fill = grey;
          cell.border = border;
          if (c === 1) cell.numFmt = '@';
          if (c !== 2) cell.alignment = { horizontal: 'center' };
        });
        for (const { subject, cols, totalCol } of layout) {
          const maxes = subject.parts.length ? subject.parts.map((p) => p.max) : [subject.max];
          cols.forEach((c, k) => {
            const cell = ws.getCell(r, c);
            cell.border = border;
            cell.alignment = { horizontal: 'center' };
            cell.protection = { locked: false };
            cell.dataValidation = validationFor(cell.address, maxes[k]);
          });
          if (totalCol) {
            const range = ws.getCell(r, cols[0]).address + ':' + ws.getCell(r, cols[cols.length - 1]).address;
            const cell = ws.getCell(r, totalCol);
            cell.value = { formula: 'IF(COUNT(' + range + ')=0,"",SUM(' + range + '))' };
            cell.font = { bold: true };
            cell.fill = grey;
            cell.border = border;
            cell.alignment = { horizontal: 'center' };
          }
        }
      });
      await ws.protect('', {
        selectLockedCells: true, selectUnlockedCells: true,
        formatCells: true, formatColumns: true, formatRows: true, sort: true, autoFilter: true,
      });
    }

    const meta = {
      tag: META_TAG, version: VERSION, createdAt: new Date().toISOString(),
      host: spec.host || '', branch: spec.branch, cls: spec.cls, exam: spec.exam,
      subjects, sections: metaSections,
    };
    const metaWs = wb.addWorksheet(META_SHEET, { state: 'hidden' });
    metaWs.getCell('A1').value = 'mcb-updater needs this tab to upload the marks. Do not edit or delete it.';
    metaWs.getCell('A2').value = JSON.stringify(meta);
    return wb;
  }

  /** A cell's plain value: formulas give their result, rich text gives its text. */
  function plain(value) {
    if (value == null) return null;
    if (typeof value !== 'object') return value;
    if (value instanceof Date) return value.toISOString();
    if ('result' in value) return plain(value.result);
    if ('formula' in value || 'sharedFormula' in value) return null; // a formula nobody has calculated yet
    if (Array.isArray(value.richText)) return value.richText.map((t) => t.text).join('');
    if ('text' in value) return plain(value.text);
    if ('error' in value) return String(value.error);
    return String(value);
  }

  /**
   * Reads a filled template (ArrayBuffer / Buffer). Returns
   * { meta, sections: [{ id, name, tab, found, problems, rows: [{ row, code, name, cells: {subjectId: [raw]}, totals: {subjectId: raw} }] }] }
   * Columns are found by their header text, so moving or adding columns is fine.
   */
  async function readWorkbook(ExcelJS, data) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    const metaWs = wb.getWorksheet(META_SHEET);
    if (!metaWs) {
      throw new Error('This file was not made by "Make template": its hidden ' + META_SHEET + ' tab is missing.');
    }
    let meta;
    try { meta = JSON.parse(String(plain(metaWs.getCell('A2').value))); } catch (e) { meta = null; }
    if (!meta || meta.tag !== META_TAG || !Array.isArray(meta.sections) || !Array.isArray(meta.subjects)) {
      throw new Error('The hidden ' + META_SHEET + ' tab was changed, so the file cannot be matched to an exam. Make a fresh template.');
    }

    const sections = meta.sections.map((sec) => {
      const out = { id: sec.id, name: sec.name, tab: sec.tab, found: false, problems: [], rows: [] };
      const ws = wb.getWorksheet(sec.tab);
      if (!ws) return out;
      out.found = true;

      const colOf = new Map();
      ws.getRow(1).eachCell((cell, colNumber) => {
        const key = norm(plain(cell.value));
        if (key && !colOf.has(key)) colOf.set(key, colNumber);
      });
      const need = (header) => {
        const col = colOf.get(norm(header));
        if (!col) out.problems.push('The column "' + header + '" is missing from the tab "' + sec.tab + '".');
        return col;
      };
      const codeCol = need(HEAD_CODE);
      const nameCol = colOf.get(norm(HEAD_NAME));
      const layout = meta.subjects.map((subject) => ({
        id: subject.id,
        cols: subject.columns.map(need),
        totalCol: subject.totalHeader ? colOf.get(norm(subject.totalHeader)) : null,
      }));
      if (out.problems.length) return out;

      for (let r = 2; r <= ws.rowCount; r += 1) {
        const row = ws.getRow(r);
        const entry = {
          row: r,
          code: plain(row.getCell(codeCol).value),
          name: nameCol ? plain(row.getCell(nameCol).value) : '',
          cells: {},
          totals: {},
        };
        for (const item of layout) {
          entry.cells[item.id] = item.cols.map((c) => plain(row.getCell(c).value));
          entry.totals[item.id] = item.totalCol ? plain(row.getCell(item.totalCol).value) : null;
        }
        out.rows.push(entry);
      }
      return out;
    });
    return { meta, sections };
  }

  /** "https://docs.google.com/spreadsheets/d/<id>/edit..." -> the link that downloads it as .xlsx, or null. */
  function googleSheetExportUrl(link) {
    const m = String(link || '').match(/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})/);
    return m ? 'https://docs.google.com/spreadsheets/d/' + m[1] + '/export?format=xlsx' : null;
  }

  return {
    VERSION, ABSENT_CODES, META_SHEET,
    norm, normCode, parseCell, sameMark, formatMark,
    parseParts, checkParts, subjectColumns, computeSubject,
    evaluateSection, reconcileSection, summarizeSection, marksFor, diffMarks, verifyMarks,
    mcbDate, parseDashboardIds, parseBranches, parseClasses, parseSections, parseExams, parseSubjectList,
    studentsQuerySubjects, entryClosed, parseSectionInfo, buildMcbCsv, uploadFields, uploadOutcome,
    buildWorkbook, readWorkbook, googleSheetExportUrl,
  };
});
