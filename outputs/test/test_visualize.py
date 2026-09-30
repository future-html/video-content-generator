#!/usr/bin/env python3
"""
Automated test suite for outputs/visualize.html

Tests:
  1. Files exist
  2. Extract JS blocks from HTML
  3. Build P6_IDS from data.json (Python)
  4. Run predicates in Node VM
  5. ISSUE_SUMMARY has 12 items P1-P12
  6. ISSUE_MARKS covers all P1-P12
  7. isFlaggedRow counts match expected
  8. JS syntax check (node --check)

Usage:
  python3 outputs/test/test_visualize.py
  outputs/test/run.sh
"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
HTML = ROOT / 'outputs' / 'visualize.html'
DATA = ROOT / 'outputs' / 'data.json'

PASS = 0
FAIL = 0
FAILURES = []


def ok(name, cond, detail=''):
    global PASS, FAIL
    if cond:
        print(f'  [PASS] {name}')
        PASS += 1
    else:
        print(f'  [FAIL] {name}' + (f' -- {detail}' if detail else ''))
        FAIL += 1
        FAILURES.append(name)


def section(title):
    print(f'\n=== {title} ===')


def main():
    section('1. Files')
    ok('visualize.html exists', HTML.exists(), str(HTML))
    ok('data.json exists', DATA.exists(), str(DATA))
    if FAIL:
        return 1

    html = HTML.read_text(encoding='utf-8')
    data = json.loads(DATA.read_text(encoding='utf-8'))

    section('2. Extract JS')
    sm = re.search(r'<script>([\s\S]*?)</script>', html)
    ok('script block found', bool(sm))
    if not sm:
        return 1
    script = sm.group(1)

    def extract(pattern, name):
        m = re.search(pattern, script)
        ok(f'extract {name}', bool(m))
        return m.group(0) if m else None

    summary_src = extract(r'const ISSUE_SUMMARY = \[[\s\S]*?\n\];', 'ISSUE_SUMMARY')
    marks_src = extract(r'const ISSUE_MARKS = \{[\s\S]*?\n\};', 'ISSUE_MARKS')
    flag_src = extract(r'function isFlaggedRow\(sheet, rec\) \{[\s\S]*?\n\}', 'isFlaggedRow')
    if not all([summary_src, marks_src, flag_src]):
        return 1

    section('3. Build P6_IDS (Python)')
    by_id = {}
    for r in data['sheets']['COMMENT_RECHECK']['records']:
        cid = str(r.get('Comment ID') or '')
        tp = str(r.get('Verification Timepoint') or '')
        if not cid or not tp:
            continue
        by_id.setdefault(cid, {})[tp] = str(r.get('Recheck Result') or '')
    p6_ids = sorted({cid for cid, p in by_id.items() if p.get('30m') != p.get('24h')})
    ok('P6_IDS = 28', len(p6_ids) == 28, f'actual: {len(p6_ids)}')

    section('4. Run predicates in Node')
    wrapper_js = """
const fs = require('fs');
const data = JSON.parse(fs.readFileSync('__DATA__', 'utf-8'));
const P6_IDS = new Set(__P6__);
const window = { P6_IDS };

__SUMMARY__

__MARKS__

__FLAG__

function cnt(sheet, pid) {
  return data.sheets[sheet].records.filter(r => isFlaggedRow(sheet, r).includes(pid)).length;
}
function uniq(sheet, pid) {
  const s = new Set();
  for (const r of data.sheets[sheet].records) {
    if (isFlaggedRow(sheet, r).includes(pid)) s.add(r['Comment ID']);
  }
  return s.size;
}

const result = {
  summary_ids: ISSUE_SUMMARY.map(x => x.id),
  summary_len: ISSUE_SUMMARY.length,
  marked: (() => {
    const s = new Set();
    for (const marks of Object.values(ISSUE_MARKS))
      for (const pid of Object.keys(marks)) s.add(pid);
    return [...s];
  })(),
  counts: {
    P1_DAILY_ACCOUNT: cnt('DAILY_ACCOUNT', 'P1'),
    P2_COMMENT_EXECUTION: cnt('COMMENT_EXECUTION', 'P2'),
    P3_DAILY_ACCOUNT: cnt('DAILY_ACCOUNT', 'P3'),
    P4_COMMENT_EXECUTION: cnt('COMMENT_EXECUTION', 'P4'),
    P5_COMMENT_EXECUTION: cnt('COMMENT_EXECUTION', 'P5'),
    P7_DAILY_ACCOUNT: cnt('DAILY_ACCOUNT', 'P7'),
    P8_COMMENT_EXECUTION: cnt('COMMENT_EXECUTION', 'P8'),
    P9_COMMENT_EXECUTION: cnt('COMMENT_EXECUTION', 'P9'),
    P10_COMPLETENESS_CHECK: cnt('COMPLETENESS_CHECK', 'P10'),
    P11_EVIDENCE_INDEX: cnt('EVIDENCE_INDEX', 'P11'),
    P12_DAILY_BRAND: cnt('DAILY_BRAND', 'P12'),
    P6_uniq: uniq('COMMENT_RECHECK', 'P6'),
  }
};
console.log(JSON.stringify(result));
"""
    wrapper_js = (
        wrapper_js
        .replace('__DATA__', str(DATA))
        .replace('__P6__', json.dumps(p6_ids))
        .replace('__SUMMARY__', summary_src)
        .replace('__MARKS__', marks_src)
        .replace('__FLAG__', flag_src)
    )

    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as f:
        f.write(wrapper_js)
        tmp = f.name
    try:
        try:
            r = subprocess.run(['node', tmp], capture_output=True, text=True, timeout=30)
        except FileNotFoundError:
            ok('node available', False, 'node not found in PATH')
            return 1
        if r.returncode != 0:
            ok('node wrapper run', False, r.stderr.strip()[:200])
            return 1
        ok('node wrapper run', True)
        result = json.loads(r.stdout.strip())
    finally:
        Path(tmp).unlink(missing_ok=True)

    section('5. ISSUE_SUMMARY')
    ok('has 12 items', result['summary_len'] == 12, f"actual: {result['summary_len']}")
    for i in range(1, 13):
        ok(f'contains P{i}', f'P{i}' in result['summary_ids'])

    section('6. ISSUE_MARKS coverage')
    marked = set(result['marked'])
    for pid in result['summary_ids']:
        ok(f'mapping exists: {pid}', pid in marked)

    section('7. isFlaggedRow counts')
    expected = [
        ('P1_DAILY_ACCOUNT', 1),
        ('P2_COMMENT_EXECUTION', 1),
        ('P3_DAILY_ACCOUNT', 40),
        ('P4_COMMENT_EXECUTION', 3),
        ('P5_COMMENT_EXECUTION', 83),
        ('P7_DAILY_ACCOUNT', 8),
        ('P8_COMMENT_EXECUTION', 1),
        ('P9_COMMENT_EXECUTION', 100),
        ('P10_COMPLETENESS_CHECK', 4),
        ('P11_EVIDENCE_INDEX', 1),
        ('P12_DAILY_BRAND', 1),
        ('P6_uniq', 28),
    ]
    for key, exp in expected:
        actual = result['counts'][key]
        ok(f'{key} = {exp}', actual == exp, f'actual: {actual}')

    section('8. JS syntax check')
    tmp_js = Path('/tmp/viz_check.js')
    tmp_js.write_text(script, encoding='utf-8')
    try:
        r = subprocess.run(['node', '--check', str(tmp_js)], capture_output=True, text=True)
        ok('node --check', r.returncode == 0, r.stderr.strip()[:200])
    except FileNotFoundError:
        ok('node --check', False, 'node not found')

    section('9. Result')
    print(f'\n  PASS: {PASS}   FAIL: {FAIL}')
    if FAILURES:
        print('\n  Failures:')
        for name in FAILURES:
            print(f'    - {name}')
        return 1
    print('\n  ALL TESTS PASSED')
    return 0


if __name__ == '__main__':
    sys.exit(main())
