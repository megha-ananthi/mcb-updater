'use strict';
// Run with: node --test test/
// Every student here is made up.
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs.min.js');
const core = require('../core.js');

const SCIENCE = { id: 's1', name: 'SCIENCE', max: 20, parts: [{ name: 'Physics', max: 7 }, { name: 'Chemistry', max: 6 }, { name: 'Biology', max: 7 }] };
const MATHS = { id: 's2', name: 'MATHEMATICS', max: 20, parts: [] };

const SPEC = {
  host: 'demo.myclassboard.com',
  branch: { id: '1', name: 'Demo School' },
  cls: { id: '99', name: 'Grade 8' },
  exam: { id: '1595', name: 'CHAPTER TEST 1' },
  subjects: [SCIENCE],
  sections: [
    { id: '11', name: 'Section A', students: [{ code: '16XX000003', name: 'CHITRA K' }, { code: '16XX000001', name: 'ARUN M' }, { code: '16XX000002', name: 'BALA S' }] },
    { id: '12', name: 'Section B', students: [{ code: '17XX000010', name: 'DEVI R' }, { code: '17XX000011', name: 'ESWAR P' }] },
  ],
};

test('parseCell tells numbers, absent codes, blanks and junk apart', () => {
  assert.deepEqual(core.parseCell(4.5), { kind: 'num', value: 4.5 });
  assert.deepEqual(core.parseCell(' 6 '), { kind: 'num', value: 6 });
  assert.deepEqual(core.parseCell('ab'), { kind: 'absent', code: 'AB' });
  assert.equal(core.parseCell('').kind, 'blank');
  assert.equal(core.parseCell(null).kind, 'blank');
  assert.equal(core.parseCell('six').kind, 'bad');
  assert.equal(core.parseCell('-2').kind, 'bad');
});

test('parts must be two or more, differently named, and add up to the subject maximum', () => {
  assert.equal(core.checkParts(SCIENCE.parts, 20), null);
  assert.equal(core.checkParts([], 20), null); // no split
  assert.match(core.checkParts(SCIENCE.parts, 25), /add up to 20, but MCB expects 25/);
  assert.match(core.checkParts([{ name: 'A', max: 5 }, { name: 'a', max: 5 }], 10), /twice/);
  assert.match(core.checkParts([{ name: 'Theory', max: 20 }], 20), /two or more/);
});

test('a student total is the sum of the parts', () => {
  assert.equal(core.computeSubject([5, 3.5, 5.5], SCIENCE).mark, '14');
  assert.equal(core.computeSubject([6.5, 4.5, 3.5], SCIENCE).mark, '14.5');
  assert.equal(core.computeSubject([0.1, 0.2, 0], SCIENCE).mark, '0.3'); // no 0.30000000000000004
  assert.equal(core.computeSubject(['7', '6', '7'], SCIENCE).mark, '20');
});

test('a blank part counts as 0 with a warning; a fully blank row is not entered', () => {
  const res = core.computeSubject([6, 3.5, null], SCIENCE);
  assert.equal(res.status, 'ok');
  assert.equal(res.mark, '9.5');
  assert.match(res.warnings[0], /Biology is blank/);
  assert.equal(core.computeSubject([null, '', undefined], SCIENCE).status, 'blank');
});

test('absent codes pass through; mixing them with marks is an error', () => {
  assert.deepEqual(core.computeSubject(['a', null, null], SCIENCE).mark, 'A');
  assert.equal(core.computeSubject(['AB', 'AB', 'AB'], SCIENCE).mark, 'AB');
  assert.equal(core.computeSubject(['A', 3, null], SCIENCE).status, 'error');
  assert.equal(core.computeSubject(['A', 'L', null], SCIENCE).status, 'error');
});

test('marks above a maximum and junk are errors', () => {
  assert.match(core.computeSubject([7.5, 3, 3], SCIENCE).errors[0], /Physics: 7.5 is more than the maximum 7/);
  assert.match(core.computeSubject([5, 'x', 3], SCIENCE).errors[0], /Chemistry: "x" is not a mark/);
  assert.match(core.computeSubject([21], MATHS).errors[0], /more than the maximum 20/);
  assert.equal(core.computeSubject([18.5], MATHS).mark, '18.5');
});

// The shapes below are MCB's (taken from a capture of its Bulk Upload Marks/Grades page); the data is made up.
const STUDENTS_ANSWER = {
  ds1: [
    { StudentEnrollmentID: 31218, StudentEnrollmentCode: '16XX000001', StudentName: 'ARUN  M  ', FullName: 'ARUN  M  ' },
    { StudentEnrollmentID: 31253, StudentEnrollmentCode: '16XX000002', StudentName: 'BALA.S  ', FullName: 'BALA.S  ' },
    { StudentEnrollmentID: 31289, StudentEnrollmentCode: '16XX000003', StudentName: 'CHITRA K', FullName: 'CHITRA K' },
  ],
  ds2: '[{"SubjectName":" SCIENCE","SubjectID":"36513","SubjectSkillID":"-1","MaxMarks":"20.00","MinMarks":"0.00","MarksEntryLastDate":"11/24/2026 12:00:00 AM","Priority":"1","Status":"1"}]',
  ds3: '[]',
  ds4: [
    { StudentEnrollmentID: 31218, SubjectID: '36513', SubjectSkillID: '-1', MarksAcquired: '14.50', AbsentType: '', ExaminationMarksID: 9001 },
    { StudentEnrollmentID: 31253, SubjectID: '36513', SubjectSkillID: '-1', MarksAcquired: '-1.00', AbsentType: 'AB', ExaminationMarksID: 9002 },
  ],
  ds5: false,
};

test("MCB's lists are read into id and name", () => {
  assert.deepEqual(core.parseDashboardIds("<script> var BranchID1 = '35'; var OrgID = '7'; ... var AID = '18'; </script>"),
    { orgId: '7', branchId: '35', academicYearId: '18' });
  assert.deepEqual(core.parseBranches({ totRes: [{ BranchID: 35, BranchName: 'Demo School ' }, { BranchID: 35, BranchName: 'Demo School ' }] }),
    [{ id: '35', name: 'Demo School' }]);
  assert.deepEqual(core.parseClasses([{ ClassID: 99, ClassName: 'Grade 8' }]), [{ id: '99', name: 'Grade 8' }]);
  assert.deepEqual(core.parseSections([{ SectionID: 424, Section: 'Section A', BranchSectionID: 1779, Active: 1 }, { SectionID: 9, Section: 'Old', BranchSectionID: 5, Active: 0 }]),
    [{ id: '1779', name: 'Section A' }]);
  assert.deepEqual(core.parseSubjectList([{ Selected: false, Text: 'SCIENCE', Value: '36513,-1' }]), [{ id: '36513,-1', name: 'SCIENCE' }]);
  const exams = core.parseExams({ dt1: 2, dt2: [
    { ExamName: 'APRIL TEST ', ExaminationID: 14232, Date: '/Date(1776191400000)/' },
    { ExamName: 'NOVEMBER TEST', ExaminationID: 15955, Date: '/Date(1794162600000)/' },
  ] });
  assert.deepEqual(exams.map((e) => e.id), ['15955', '14232']); // newest first
  assert.equal(exams[1].name, 'APRIL TEST');
  assert.match(exams[0].label, /^NOVEMBER TEST \(\d+ Nov 2026\)$/);
});

test("a section's students, subjects and saved marks are read from MCB's answer", () => {
  const info = core.parseSectionInfo(STUDENTS_ANSWER);
  assert.deepEqual(info.students, [{ code: '16XX000001', name: 'ARUN M' }, { code: '16XX000002', name: 'BALA.S' }, { code: '16XX000003', name: 'CHITRA K' }]);
  assert.equal(info.subjects.length, 1);
  assert.deepEqual([info.subjects[0].id, info.subjects[0].name, info.subjects[0].max, info.subjects[0].lastDate],
    ['36513,-1', 'SCIENCE', 20, '11/24/2026 12:00:00 AM']);
  assert.deepEqual([...info.saved['36513,-1']], [['16XX000001', '14.50'], ['16XX000002', 'AB']]); // -1.00 means absent
  assert.equal(info.gradeOnly, false);
  assert.equal(core.studentsQuerySubjects([{ id: '36513,-1', name: 'SCIENCE' }, { id: '36514,-1', name: 'SOCIAL SCIENCE' }]),
    '36513,-1, SCIENCE$36514,-1, SOCIAL SCIENCE');
  // MCB answers 0 or nothing when a section has no students
  assert.deepEqual(core.parseSectionInfo({ ds1: 0, ds2: '', ds4: 0 }).students, []);
});

test("the upload file and its form fields match what MCB's page sends", () => {
  const info = core.parseSectionInfo(STUDENTS_ANSWER);
  const marks = new Map([['16XX000003', 'A'], ['16XX000001', '14.5']]);
  // roster order, only students with a mark, CRLF line ends
  assert.equal(core.buildMcbCsv(info.subjects[0], info.students, marks),
    'Student EnrollmentCode, SCIENCE\r\n16XX000001,14.5\r\n16XX000003,A\r\n');
  const ctx = { examId: '15955', examName: 'NOVEMBER TEST', classId: '99', sectionId: '1779', academicYearId: '18', branchId: '35' };
  assert.deepEqual(core.uploadFields(ctx, info.subjects[0]), {
    slectsubjectCount: '2',
    dlUploadingSubjects: '36513#-1# SCIENCE#0.00 #20.00#11/24/2026 12:00:00 AM',
    chktext: 'Student EnrollmentCode',
    CheckBoxList1values: 'StudentEnrollmentCode',
    ddlExamnameID: '15955',
    subjects: '36513,-1# SCIENCE',
    ClassID: '99',
    BranchSectionID: '1779',
    ddlExamname: 'NOVEMBER TEST',
    AcademicYearID: '18',
    BranchID: '35',
  });
});

test("MCB's answer to an upload is read as accepted or refused", () => {
  const refused = core.uploadOutcome({ dt1: '', dt2: '<b>Line number: 2  Code: 16XX000001</b>  >>These students marks are not in correct format<br/>', dt4: 0, dt7: '' });
  assert.equal(refused.ok, false);
  assert.equal(refused.message, 'Line number: 2 Code: 16XX000001 >>These students marks are not in correct format');
  assert.equal(core.uploadOutcome({ dt1: '[]', dt2: '', dt4: 1, dt7: '' }).ok, true);
  assert.equal(core.uploadOutcome({}).ok, false);
});

test('a subject is closed once its marks entry last date has passed', () => {
  assert.equal(core.entryClosed('11/24/2026 12:00:00 AM', new Date(2026, 10, 22)), false);
  assert.equal(core.entryClosed('11/24/2026 12:00:00 AM', new Date(2026, 10, 25)), true);
  assert.equal(core.entryClosed('', new Date()), false);
});

async function filledTemplate(spec, fill) {
  const wb = await core.buildWorkbook(ExcelJS, spec);
  if (fill) fill(wb);
  const bytes = await wb.xlsx.writeBuffer();
  return core.readWorkbook(ExcelJS, bytes);
}

test('the template has one tab per section, sorted by name, and reads back', async () => {
  const wb = await core.buildWorkbook(ExcelJS, SPEC);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ['Section A', 'Section B', '_meta']);
  const a = wb.getWorksheet('Section A');
  assert.deepEqual(a.getRow(1).values.slice(1), ['S.No', 'Enrollment code', 'Name', 'Physics (7)', 'Chemistry (6)', 'Biology (7)', 'SCIENCE TOTAL (20)']);
  assert.deepEqual([2, 3, 4].map((r) => a.getCell(r, 3).value), ['ARUN M', 'BALA S', 'CHITRA K']);
  assert.equal(a.getCell('G2').value.formula, 'IF(COUNT(D2:F2)=0,"",SUM(D2:F2))');
  assert.equal(wb.getWorksheet('_meta').state, 'hidden');

  const read = await filledTemplate(SPEC);
  assert.equal(read.meta.exam.id, '1595');
  assert.equal(read.sections.length, 2);
  assert.equal(read.sections[0].rows.length, 3);
  const judged = core.evaluateSection(read.sections[0], read.meta);
  // nothing filled yet: every student is "not entered" and nothing would be uploaded
  assert.deepEqual(core.summarizeSection(judged, read.meta.subjects), { toWrite: 0, absent: 0, blank: 3, errors: 0, warnings: 0, blocked: false });
  assert.equal(core.marksFor(judged, 's1').size, 0);
});

test('a filled template gives the marks to upload and flags the bad rows', async () => {
  const read = await filledTemplate(SPEC, (wb) => {
    const a = wb.getWorksheet('Section A');
    a.getCell('D2').value = 5; a.getCell('E2').value = 3.5; a.getCell('F2').value = 5.5; // ARUN 14
    a.getCell('D3').value = 'A';                                                         // BALA absent
    a.getCell('D4').value = 6; a.getCell('E4').value = 3.5;                              // CHITRA 9.5, Biology blank
    const b = wb.getWorksheet('Section B');
    b.getCell('D2').value = 9; b.getCell('E2').value = 1; b.getCell('F2').value = 1;     // DEVI: Physics over 7
  });
  const [a, b] = read.sections.map((s) => core.evaluateSection(s, read.meta));

  core.reconcileSection(a, SPEC.sections[0].students);
  assert.deepEqual([...core.marksFor(a, 's1')], [['16XX000001', '14'], ['16XX000002', 'A'], ['16XX000003', '9.5']]);
  assert.deepEqual(core.summarizeSection(a, read.meta.subjects), { toWrite: 3, absent: 1, blank: 0, errors: 0, warnings: 1, blocked: false });

  // MCB now has a student who joined after the template was made
  core.reconcileSection(b, SPEC.sections[1].students.concat([{ code: '17XX000012', name: 'FATHIMA N' }]));
  const sum = core.summarizeSection(b, read.meta.subjects);
  assert.equal(sum.blocked, true);
  assert.equal(sum.errors, 1);
  assert.equal(sum.blank, 1); // ESWAR P has no marks yet
  assert.match(b.warnings[0], /1 student\(s\) in MCB are not in this tab.*FATHIMA N \(17XX000012\)/);
});

test('students unknown to MCB, duplicate codes and renamed tabs are caught', async () => {
  const read = await filledTemplate(SPEC, (wb) => {
    const a = wb.getWorksheet('Section A');
    a.getCell('D2').value = 5;
    a.getCell('B3').value = '16XX000001'; a.getCell('D3').value = 4; // same code twice
    a.getCell('B4').value = '99ZZ000000'; a.getCell('D4').value = 4; // not in MCB
    wb.getWorksheet('Section B').name = 'B renamed';
  });
  const [a, b] = read.sections.map((s) => core.evaluateSection(s, read.meta));
  core.reconcileSection(a, SPEC.sections[0].students);
  assert.match(a.rows[1].errors[0], /also in row 2/);
  assert.match(a.rows[2].errors[0], /not in MCB's list/);
  assert.match(b.errors[0], /tab "Section B" is missing/);
});

test('two subjects in one sheet keep their columns apart', async () => {
  const spec = Object.assign({}, SPEC, { subjects: [SCIENCE, MATHS] });
  const read = await filledTemplate(spec, (wb) => {
    const a = wb.getWorksheet('Section A');
    assert.deepEqual(a.getRow(1).values.slice(4), ['SCIENCE: Physics (7)', 'SCIENCE: Chemistry (6)', 'SCIENCE: Biology (7)', 'SCIENCE TOTAL (20)', 'MATHEMATICS (20)']);
    a.getCell('D2').value = 7; a.getCell('E2').value = 6; a.getCell('F2').value = 7;
    a.getCell('H2').value = 18.5;
  });
  const a = core.evaluateSection(read.sections[0], read.meta);
  assert.deepEqual([...core.marksFor(a, 's1')], [['16XX000001', '20']]);
  assert.deepEqual([...core.marksFor(a, 's2')], [['16XX000001', '18.5']]);
});

test('a file that is not a template is refused', async () => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('A').getCell('A1').value = 'hello';
  await assert.rejects(core.readWorkbook(ExcelJS, await wb.xlsx.writeBuffer()), /not made by "Make template"/);
});

test('diff and verify compare marks by value', () => {
  const next = new Map([['1', '14'], ['2', '9.5'], ['3', 'A'], ['4', '7']]);
  const old = new Map([['1', '14.0'], ['2', '8'], ['3', 'a'], ['4', '']]);
  assert.deepEqual(core.diffMarks(next, old), { fresh: 1, same: 2, changed: [{ code: '2', from: '8', to: '9.5' }] });
  assert.deepEqual(core.verifyMarks(next, new Map([['1', 14], ['2', '9.50'], ['3', 'A']])), [{ code: '4', sent: '7', found: '' }]);
});

test('a Google Sheets link becomes its .xlsx download link', () => {
  assert.equal(
    core.googleSheetExportUrl('https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz012345/edit?gid=5#gid=5'),
    'https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz012345/export?format=xlsx'
  );
  assert.equal(core.googleSheetExportUrl('https://example.com/x'), null);
});
