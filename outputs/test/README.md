# Test Suite -- visualize.html

Automated tests for the Whiteshirt Data Visualizer.

## Tests
1. Files exist (`visualize.html`, `data.json`)
2. Extract JS blocks from HTML
3. Build P6_IDS from `data.json` (Python)
4. Run predicates via Node.js
5. `ISSUE_SUMMARY` has 12 items P1-P12
6. `ISSUE_MARKS` covers all P1-P12
7. `isFlaggedRow` counts match expected
8. JS syntax check (`node --check`)

## Requirements
- Python 3.10+
- Node.js (for JS execution)

## Run
From repo root (`video-content-generator/`):

    python3 outputs/test/test_visualize.py

Or use the helper:

    outputs/test/run.sh

## Exit codes
- `0` -- all tests passed
- `1` -- one or more tests failed

## Expected counts
| Issue | Sheet | Expected |
|-------|-------|----------|
| P1  | DAILY_ACCOUNT       | 1   |
| P2  | COMMENT_EXECUTION   | 1   |
| P3  | DAILY_ACCOUNT       | 40  |
| P4  | COMMENT_EXECUTION   | 3   |
| P5  | COMMENT_EXECUTION   | 83  |
| P6  | COMMENT_RECHECK     | 28  |
| P7  | DAILY_ACCOUNT       | 8   |
| P8  | COMMENT_EXECUTION   | 1   |
| P9  | COMMENT_EXECUTION   | 100 |
| P10 | COMPLETENESS_CHECK  | 4   |
| P11 | EVIDENCE_INDEX      | 1   |
| P12 | DAILY_BRAND         | 1   |
