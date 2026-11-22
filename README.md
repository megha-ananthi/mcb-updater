# MCB marks updater

A teacher fills the marks for a whole class in one sheet (one tab per section), and this tool writes them into
MyClassBoard (MCB) section by section. A subject can be split into parts, for example Science as Physics 7 +
Chemistry 6 + Biology 7. MCB only gets the total.

**Status: wired to MCB's real requests and checked against a capture of them. Not yet run against the live
MCB**, so the first real upload should be one section only (see [Not checked yet](#not-checked-yet)).

## How it works

It is a bookmarklet. The teacher logs in to MCB as usual and clicks a bookmark. That loads `mcb-updater.js` from
this site into the MCB page, where it runs with her own session. No password is handled, and marks go from her
sheet to MCB and nowhere else. A normal website cannot do this: the browser does not let another site use her
MCB login.

1. **Make template.** Pick branch, class, examination and subject(s); optionally tick "Split into parts" and give each part a name and a maximum. The
   panel reads every section's student list from MCB and downloads one `.xlsx`: a tab per section with
   `S.No | Enrollment code | Name | <parts> | TOTAL`, and a hidden `_meta` tab that records which exam the file
   is for.
2. **Fill.** In Google Sheets or Excel. Absent = `A`, `AB`, `L` or `AM`. A blank part counts as 0. A fully blank
   student is left "not entered".
3. **Upload marks.** Give the panel the `.xlsx` file or a Google Sheets link (shared as "Anyone with the link").
   It checks every row against MCB's current student list and shows a preview per section. A section with any
   error is skipped whole. On Upload it sends MCB one small CSV per section and subject (enrollment code and
   mark, only students who have a mark), then reads the section back and compares. It stops at the first section
   that does not match.

## Files

| File | What it is |
|---|---|
| `index.html` | The page with the bookmark to drag and the steps for the teacher |
| `mcb-updater.js` | The panel and the MCB calls (`createMcbApi`) |
| `core.js` | Marks arithmetic, sheet checks, MCB's request and response formats, the `.xlsx` template. No page, no network |
| `vendor/exceljs.min.js` | ExcelJS 4.4.0 (MIT), reads and writes `.xlsx`. sha256 `7e49da68…f742c9` |
| `test/core.test.js` | Tests for `core.js` |
| `test/harness.html` | A fake MCB page to try the panel without touching real marks |
| `discovery/` | Captures from the real portal. Git-ignored, never commit it |

No build step. Plain JavaScript.

## Try it

```bash
node --test test/core.test.js
```

```bash
python3 -m http.server 8765 --bind 127.0.0.1
```

Then open `http://localhost:8765/test/harness.html`, press "Open the panel", make a template, and press "Hand the
panel a filled sample sheet". The two tick boxes on that page add a mistake to the sheet and make the fake MCB
drop a mark, to show the checks catching both.

## MCB endpoints

MCB has no public API docs. The panel makes the same requests as MCB's own "Bulk Upload Marks/Grades" page
(CBSE and State → Other Exams → Marks Entry Dashboard → Proceed to Enter Marks → Bulk Upload Marks/Grades). They
are same-origin, use the login cookie, and carry `X-Requested-With: XMLHttpRequest`. The calls are in
`createMcbApi()` in `mcb-updater.js`; the formats are in the "MCB's wire formats" part of `core.js`.

| Need | Request | Notes |
|---|---|---|
| Organisation and academic year ids | `GET /Examination/OtherExaminationsDashboard/` | read from `var OrgID`, `var AID` in the page |
| Branches | `POST /MasterData/GetBranchLocationsForSearchForstudent` | form: `OrganisationID` |
| Classes | `POST /AzureClass/GetClassesbyAzure` | form: `BranchID` |
| Sections | `POST /AzureClass/GetSectionsbyAzure?ClassID=&BranchID=&AcademicYearID=` | the section id used everywhere is `BranchSectionID` |
| Examinations | `POST /Examination/GetExamsByBranchSectionByAcademicYearID` | form: `classId`, `BranchSectionID`, `AcademicYearID`, `BranchID`; asked for the first section |
| Subjects | `POST /Examination/GetSubjectsByExaminationByBranchID?ExaminationId=&ClassID=&BranchID=` | `Value` is `SubjectID,SubjectSkillID` |
| Students, maximum marks, saved marks | `POST /Examination/GetStrudents` | form: `chksub`, `ExaminationId`, `BranchSectionID`, `AcademicYearID`, `BranchID`. `ds1` students, `ds2` subjects with `MaxMarks` and `MarksEntryLastDate`, `ds4` saved marks (`-1.00` = absent, with `AbsentType`), `ds5` grades-only flag |
| Upload | `POST /Examination/SaveBulkExaminationMarksTosheet` | multipart: the CSV as `FileUpload2` plus the fields in `core.uploadFields()`. Answer: `dt4` is 0 when nothing was saved, `dt2` / `dt7` hold the message |

The uploaded file is what MCB's "Upload marks" tab takes with only "Student EnrollmentCode" ticked:

```
Student EnrollmentCode, SCIENCE
16XX010052,14.5
16XX010087,A
```

MCB refuses a line with an empty mark, so students without a mark are left out of the file.

Examinations that take grades instead of marks use other MCB requests and are refused by the panel.

## Not checked yet

Everything above was checked against a capture of MCB's requests and against MCB's page script, not against the
live site. The capture holds one upload, of an empty file, which MCB refused. So these are still assumptions,
and the read-back after each upload is what catches them:

- A successful upload answers with a non-zero `dt4` (this is how MCB's own page decides).
- MCB takes decimal marks such as `14.5` and the absent codes in an uploaded file.
- An upload replaces a mark that is already saved. MCB's page uses a different request to edit a saved mark, so
  it may keep the old one. The preview lists these as "replace a mark in MCB".
- Saved marks come back in `ds4` the way MCB's page script reads them.
- The template opened in Google Sheets and downloaded again as `.xlsx` keeps its cell validation, locked
  columns and hidden `_meta` tab (only the ExcelJS round trip is tested).

For the first real run, use "Tick only the first section", then look at that section in MCB.
