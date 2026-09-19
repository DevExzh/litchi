#!/usr/bin/env python3
"""Retain raw native phase samples from four independent process legs."""
import hashlib
import json
import subprocess
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
binary = ROOT.parent / 'litchi-0691-bin/native'
cases = {
    'real': str(ROOT / 'test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx'),
    'control': str(P / 'marker-control.pptx'),
    'generated': 'generated:12x8',
    'notes-poi': str(ROOT / 'test-data/poi/test-data/slideshow/prProps.pptx'),
    'notes-lo': str(ROOT / 'test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx'),
}
def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
records = []
out = P / 'native'
out.mkdir(exist_ok=True)
for leg in range(4):
    ordered = list(cases.items())
    if leg % 2:
        ordered.reverse()
    for name, source in ordered:
        command = ['taskset', '-c', '12', str(binary), 'phases', source, '100', '5']
        output = out / f'{name}-{leg}.tsv'
        error = out / f'{name}-{leg}.stderr'
        started = time.monotonic()
        with output.open('w') as stdout, error.open('w') as stderr:
            result = subprocess.run(command, cwd=ROOT, stdout=stdout, stderr=stderr)
        records.append(dict(case=name, leg=leg, command=command, exit_code=result.returncode,
                            seconds=time.monotonic()-started, output=str(output.relative_to(P)),
                            output_sha256=sha(output), stderr_sha256=sha(error),
                            binary_sha256=sha(binary),
                            source_sha256=None if source.startswith('generated:') else sha(Path(source))))
        (P / 'native-runs.json').write_text(json.dumps(records, indent=2)+'\n')
        print(name, leg, result.returncode, flush=True)
        assert result.returncode == 0, error.read_text()
