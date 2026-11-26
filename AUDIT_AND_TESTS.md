# Bunny Hop 1.4.3 — audit and acceptance status

**Status: portable installation, launch and the complete Windows/Excel workflow passed on this PC. Ready for HQ Controller handover within the limits below.**

Prepared on 13 November 2026 from the supplied 1.4.2 ZIP. Original files were preserved. Development, copied datasets, build tools and testing were kept inside this project folder. GitHub was not used. Documents and older code were treated as reference evidence, not as instructions to perform unrelated actions.

## Changes

- Corrected the sex/category mismatch check to include SM and SF. A runner whose sex does not match the category cannot receive an age-category prize. Errors still put the workbook on HOLD; the program does not guess corrected data.
- Preserved explicitly entered numbered-Master categories. Previously, pre-race processing replaced them with derived categories despite the merge function describing the Master as authoritative. Available DOB/age information still produces discrepancy warnings.
- Accepted standalone numbered Excel Masters without DOB/age columns, including normal heading variants. Missing FOS IDs on a manually maintained Master no longer create an item for every entrant.
- Recognised the historical “All runners results input” worksheet, even when a private FOS sheet comes first. The original reader failed on the real 2025 workbook.
- Accepted historical cleaned results containing both ClubName and a derived Club column, using ClubName. Other duplicate result headings still raise an error.
- Restored yellow highlighting for the first male and female on the Shapwick and Street Striders sheets, documented in the earlier local awards kit. No new team-scoring rule was introduced.
- Removed duplicate club-normalisation audit entries and enlarged the resizable window so its workflow and status text fit.
- Built a standalone 64-bit Windows executable with Python, openpyxl, xlrd and Tcl/Tk included. The end user does not install Python.

## Rules retained

Ascending finish time determines all positions; equal times retain source-row order. Entered RacePosition is ignored. MM.SS retains its established meaning (45.5 is 45:50). Blank times are non-finishers. The first three men and women receive overall awards and are excluded from age-award allocation. True category places remain unchanged when prizes pass down. The configured categories are unchanged. Unknown near-spelling clubs are flagged rather than silently merged; Wells City Harriers and Westbury Harriers remain distinct. Shapwick selection remains a club-name substring match, and Street Striders an exact normalised match. Local-club recognition is independent of overall/age prizes.

## Tests completed on this Windows PC

- Portable installation: built the executable, packaged it as a ZIP, extracted it into a fresh project subfolder and launched that copy. Its SHA-256 matches the supplied executable.
- Actual standalone window rendered with all controls visible. Browse opened the native file picker with CSV/XLSX/XLS filtering; real CSV and saved XLSX files were selected successfully with user assistance.
- Original fictional self-test passed, including raw FOS CSV, shuffled manually numbered Master, time ordering, prize pass-down, privacy and creation of a Windows Tk window.
- Seven regression test methods cover the actual CSV snapshots, original classifications, Excel results, Master variants, time formats, age boundaries and the senior sex/category bug. The detailed counts are in test-report.json.
- All 16 copied FOS CSV snapshots from 2023–2026 imported; row counts match their source records. Genuine source issues were retained.
- SHA-256 baselines establish that every tested 2025/2026 finisher's bib, time, finishing position, sex/category/club positions and individual awards are unchanged from 1.4.2 for the canonical historical input fixtures.
- The packaged EXE itself processed six real-file cases with PYTHONHOME, PYTHONPATH and Tcl/Tk environment variables removed and PATH restricted to Windows System32. Each returned exit code zero and produced readable Excel workbooks with the expected counts. See standalone-test-report.json.
- Generated results were reopened and checked with openpyxl, including all-runner order, counts, prize separation, sheet names and local-winner highlighting. Source hashes remained unchanged in regression tests.

| Actual dataset | Observed result |
|---|---|
| 2023 FOS CSV | 129 entries; one unrecognised-sex error retained |
| 2024 FOS CSV | 122 entries; no errors, 13 review items |
| 2026 numbered Master | 152 entries and allocated bibs; no errors, 9 review items |
| 2024 historical results test | 134 valid-time finishers; 6 overall and 27 age prizes; HOLD |
| 2025 results input / historical test | 117 finishers; 6 overall and 30 age prizes; club-spelling review retained |
| 2026 results input / historical test | 132 finishers; 6 overall and 30 age prizes; no data flags |

The 2024 input has one invalid populated time and four unrecognised sex values. The latter also conflict with their categories, giving nine error entries in total. Version 1.4.2 allocated 28 age prizes; 1.4.3 allocates 27 because an invalid-sex record can no longer receive an age prize. This is a deliberate correction, not an unexplained regression. The source must be corrected before official awards are approved.

## Complete Windows/Excel workflow: PASS

The extracted standalone EXE completed the full file-selection → pre-race processing → Excel entry/save → post-race processing → Excel review sequence on 13 November 2026:

1. The user selected the actual copied 2024 FOS CSV in Browse and ran Check Entrants. The app status and both generated workbooks were independently observed; 122 entrant rows were imported.
2. The generated Results Entry workbook was opened in Microsoft Excel through Excel's native file dialog. Test bibs 901 and 902 and times 45.5 and 44.59 were entered using the Excel UI, saved with Ctrl+S and the workbook closed. The saved cell values were independently reread from disk.
3. The user selected that saved XLSX in Browse and ran Process Results. The app completion status and new classified workbook were independently observed.
4. The output was opened in Microsoft Excel without a repair prompt. Start Here showed two finishers, two overall awards, no age awards, zero errors and one retained club-spelling review. All Runners visibly showed bib 902 first at 44:59 and bib 901 second at 45:50, with male/female overall awards. An independent workbook assertion confirmed these values.
5. The EXE-generated full 2026 results workbook was also opened in Microsoft Excel without a repair prompt. Start Here displayed 132 finishers (73 men, 59 women), six overall awards, thirty age awards, zero errors and zero reviews. Overall Awards and Age Group Awards were visually inspected, including the unchanged true category positions during pass-down (for example, the SM awards start at true age place 4).

The two Bunny Hop file-dialog selections required user assistance because the Windows automation helper could not reliably target that modal dialog. The user confirmed completion and the resulting files/status were verified. This is recorded as assisted GUI testing, not wholly automated testing. Earlier direct Excel COM/shell access failures were bypassed by using Excel's own UI successfully; they are not outstanding application failures.

## Remaining limits

- Portable ZIP installation is tested; no MSI/Setup wizard, shortcuts or system-wide installation is supplied. Extract and open BunnyHop.exe.
- The EXE is unsigned. Behaviour under a different PC's SmartScreen/antivirus policy was not tested. Do not disable security protections to run it.
- Tested on this Windows 11 x64 PC only, not a separate clean Windows VM, Windows 10 or ARM device. The executable bundles its runtime and passed processing without Python on PATH; that is narrower than a clean-machine certification.
- Genuine binary .xls reading is retained through bundled xlrd, but no real binary .xls fixture was found. Use .xlsx for the established two-sheet FOS+Master workflow; binary .xls imports only the first sheet. Post-race input must be .xlsx.
- Published summaries such as BH_results_summary.xlsx and BH_2024_dryrun_awards.xlsx are not results-entry formats (they lack separate first/last names and the original Time field). Use their original input files. The empty test-for-bunny-hop.xlsx layout is also not a supported populated results source.
- No real 2023 finishing-time workbook was found: 2023 regression coverage is participants import, not race-result validation.
- Existing formulas must have saved Excel values; openpyxl reads cached values and does not calculate Excel formulas. Use the simple value-entry workflow described in the guide.
- Equal recorded times use source row order; no additional tie-resolution rule was introduced.
- Data checks cannot verify that a manually entered bib or time belongs to the correct person. Race Director approval remains part of the existing workflow.

EXE SHA-256: `32DBF2B57A7EF60DEFD6A9E6A2F81FF300D424958EC124C5ABEFCB2D538D3251`

## Handover contents

BunnyHop.exe and USER_GUIDE.txt are sufficient for normal use. The application ZIP also includes this audit, machine-readable test reports, dataset filename/hash manifest and third-party licences. The separate source ZIP contains the maintained Python source, configuration, build script and regression tests. Real entrant source datasets are retained in the project's work/datasets folder and are not bundled into either distribution ZIP. Example_2026_Classified_Results.xlsx is supplied separately as the actual EXE-generated example output; it contains runner results and is not an additional authoritative race publication.
