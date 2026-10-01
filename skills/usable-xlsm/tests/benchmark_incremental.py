"""Paired benchmark of real incremental edits; uses disposable books only.

uv run --project skills/usable-xlsm python skills/usable-xlsm/tests/benchmark_incremental.py
Add --excel on a clean, provisioned Windows Excel host for the full workflow.
"""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
import types

import usable_xlsm.core as current_core
import usable_xlsm.security as current_security
from usable_xlsm.creation import create_workbook
from usable_xlsm.planning import plan_changes
from usable_xlsm.runner import excel_pids
from usable_xlsm.syntax import check_directory

BASELINE = 'eabda8e33ffaf1616d10fb6de1bad01ac879b5ee'
REPO = Path(__file__).resolve().parents[3]


def load_baseline(name, revision):
    filename = f'skills/usable-xlsm/src/usable_xlsm/{name}.py'
    source = subprocess.check_output(['git', '-C', str(REPO), '-c',
        'safe.directory=' + REPO.as_posix(), 'show', revision + ':' + filename], encoding='utf-8')
    if name == 'core':
        source = source.replace('from .security import (', 'from ._benchmark_security import (')
    module = types.ModuleType('usable_xlsm._benchmark_' + name)
    module.__package__ = 'usable_xlsm'
    module.__file__ = str(REPO / filename)
    sys.modules[module.__name__] = module
    exec(compile(source, module.__file__, 'exec'), module.__dict__)
    return module


def emit(value):
    print(json.dumps(value), flush=True)


def calc(value):
    return f'''Attribute VB_Name = "Calc"
Option Explicit
Public Function Amount() As Long
    Amount = {value}
End Function
Public Sub Test_Calc()
    AssertEqual Amount(), {value}
End Sub
'''


def ui(caption):
    return f'''Attribute VB_Name = "UI"
Option Explicit
Public Function Caption() As String
    Caption = "{caption}"
End Function
Public Sub Test_UI()
    AssertEqual Caption(), "{caption}"
End Sub
'''


EXTRA = '''Attribute VB_Name = "Extra"
Option Explicit
Public Function AddedAmount() As Long
    AddedAmount = Calc.Amount() + 4
End Function
Public Sub Test_Extra()
    AssertEqual AddedAmount(), 16
End Sub
'''


def write_source(folder, name, text):
    folder.mkdir(exist_ok=True)
    path = folder / name
    path.write_text(text, encoding='utf-8', newline='')
    return path


def score(operation, samples):
    old = statistics.median([sample[0] for sample in samples])
    new = statistics.median([sample[1] for sample in samples])
    return {'operation': operation, 'pairs': len(samples), 'before_seconds': old,
            'after_seconds': new, 'time_reduction_pct': (1 - new / old) * 100,
            'score_before': 100, 'score_after': old / new * 100, 'samples_seconds': samples}


def full_workflow(core, template, folder, label):
    folder.mkdir()
    book = folder / 'book.xlsm'
    shutil.copy2(template, book)
    records = []
    start = time.perf_counter()
    revisions = [('number', 'Calc.bas', calc(12)), ('caption', 'UI.bas', ui('Reviewed')),
                 ('add_module', 'Extra.bas', EXTRA)]
    for index, (action, name, text) in enumerate(revisions):
        source = folder / f'patch-{index}'
        path = write_source(source, name, text)
        add_only = action == 'add_module'
        before = core.extract_vba(book)
        step = time.perf_counter()
        assert not check_directory(source)
        plan = plan_changes(path if add_only else source, book, add_only=add_only, partial=not add_only, include_diff=True)
        assert plan['ok'] and not plan['removed']
        result = core.update_vba(path if add_only else source, book, add_only=add_only,
            partial=not add_only, trust_workbook=True, timeout=30)
        changed_seconds = time.perf_counter() - step
        assert result['status'] == 'updated' and result['tests'] == (3 if add_only else 2)
        after = core.extract_vba(book)
        assert set(after) == set(before) | ({name} if add_only else set())
        assert all(core._normalized_module_source(after[key]) == core._normalized_module_source(old)
                   for key, old in before.items() if key != name)
        assert core._normalized_module_source(after[name]) == core._normalized_module_source(text)
        # Repeating the exact same user request must not save/run Excel again.
        repeated = folder / f'repeated-{index}'
        write_source(repeated, name, after[name])
        digest = hashlib.sha256(book.read_bytes()).hexdigest()
        step = time.perf_counter()
        noop = core.update_vba(repeated, book, partial=True, skip_unchanged=True, trust_workbook=True)
        noop_seconds = time.perf_counter() - step
        assert noop['status'] == 'unchanged' and noop['excel_jobs'] == 0 and noop['tests'] is None
        assert hashlib.sha256(book.read_bytes()).hexdigest() == digest
        assert not excel_pids()
        record = {'version': label, 'step': action, 'changed_seconds': changed_seconds,
                  'unchanged_seconds': noop_seconds, 'vba_tests': result['tests'], 'office': result['office_version']}
        records.append(record)
        emit(record)
    return time.perf_counter() - start, records


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--baseline', default=BASELINE)
    parser.add_argument('--repeats', type=int, default=2)
    parser.add_argument('--helpers', type=int, default=18)
    parser.add_argument('--excel', action='store_true')
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error('--repeats must be positive')
    if args.helpers < 0:
        parser.error('--helpers must be nonnegative')
    if args.excel and (sys.platform != 'win32' or excel_pids()):
        parser.error('--excel requires a clean Windows Excel host')
    old_security = load_baseline('security', args.baseline)
    old_core = load_baseline('core', args.baseline)
    rows = []
    with tempfile.TemporaryDirectory(prefix='usable-xlsm-incremental-') as directory:
        root = Path(directory)
        source = root / 'source'
        write_source(source, 'Calc.bas', calc(10))
        write_source(source, 'UI.bas', ui('Draft'))
        for i in range(args.helpers):
            code = f'Attribute VB_Name = "Helper{i:02d}"\nOption Explicit\n'
            code += '\n'.join(f'Public Function Value{k:02d}() As Long\nValue{k:02d} = {k}\nEnd Function' for k in range(10))
            write_source(source, f'Helper{i:02d}.bas', code)
        spec = root / 'spec.json'
        spec.write_text(json.dumps({'sheets': [{'name': 'Report', 'rows': [['Draft report'], [10]]}]}), encoding='utf-8')
        template = root / 'template.xlsm'
        create_workbook(spec, template, source=source, experimental=True, target_os='windows')
        original_hash = hashlib.sha256(template.read_bytes()).hexdigest()
        modules = current_core.extract_vba(template)
        emit({'phase': 'fixture', 'platform': sys.platform, 'baseline': args.baseline,
              'modules': len(modules), 'source_bytes': sum(len(s.encode('utf-8')) for s in modules.values()),
              'python': sys.version.split()[0], 'repeats': args.repeats})
        scans, noops = [], []
        partial = root / 'unchanged'
        write_source(partial, 'Calc.bas', modules['Calc.bas'])
        for i in range(args.repeats):
            elapsed = [0.0, 0.0]
            reports = [None, None]
            for j in ([0, 1] if i % 2 == 0 else [1, 0]):
                begin = time.perf_counter()
                reports[j] = (old_security, current_security)[j].preflight_workbook(template, operation='edit', trust_workbook=True)
                elapsed[j] = time.perf_counter() - begin
            if reports[0].to_dict() != reports[1].to_dict():
                emit({'preflight_difference': {key: [reports[0].to_dict()[key], reports[1].to_dict()[key]]
                      for key in reports[0].to_dict() if reports[0].to_dict()[key] != reports[1].to_dict()[key]}})
                raise AssertionError('Scanner reports changed')
            assert reports[0].allowed
            scans.append(elapsed)
            elapsed = [0.0, 0.0]
            for j in ([0, 1] if i % 2 == 0 else [1, 0]):
                begin = time.perf_counter()
                result = (old_core, current_core)[j].update_vba(partial, template, partial=True,
                    skip_unchanged=True, trust_workbook=True)
                elapsed[j] = time.perf_counter() - begin
                assert result['status'] == 'unchanged' and result['excel_jobs'] == 0
                assert hashlib.sha256(template.read_bytes()).hexdigest() == original_hash
            noops.append(elapsed)
        rows.extend([score('full_security_preflight', scans), score('unchanged_update', noops)])
        for row in rows:
            emit(row)
        if args.excel:
            samples, all_records = [], []
            for i in range(args.repeats):
                elapsed = [0.0, 0.0]
                for j in ([0, 1] if i % 2 == 0 else [1, 0]):
                    elapsed[j], records = full_workflow((old_core, current_core)[j], template,
                        root / f'workflow-{i}-{j}', 'before' if j == 0 else 'after')
                    all_records.extend(records)
                samples.append(elapsed)
            rows.append(score('3_small_edits_plus_3_repeated_requests', samples))
            rows[-1]['records'] = all_records
            emit(rows[-1])
        assert hashlib.sha256(template.read_bytes()).hexdigest() == original_hash
    emit({'benchmark_result': rows})


if __name__ == '__main__':
    main()
