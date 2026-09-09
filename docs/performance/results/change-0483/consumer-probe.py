#!/usr/bin/env python3
"""Run the public file-source example and separate LibreOffice text readback.

Run under gate.py after building the example. This is an independent consumer
probe on synthetic fixtures, not a Microsoft Office compatibility result.
"""
import argparse
import os
import signal
import shutil
import subprocess
from pathlib import Path

from common import ENV, REPO, ROOT, TEMP, meta, now, write


def run_command(argv, label, cwd, directory):
    start = now()
    out = directory / f'{label}.stdout'
    err = directory / f'{label}.stderr'
    timed_out = False
    with out.open('xb') as stdout, err.open('xb') as stderr:
        process = subprocess.Popen(argv, cwd=cwd, env=ENV, stdout=stdout, stderr=stderr,
                                   start_new_session=True)
        try:
            exit_code = process.wait(timeout=120)
        except subprocess.TimeoutExpired:
            timed_out = True
            os.killpg(process.pid, signal.SIGKILL)
            exit_code = process.wait()
    write(directory / f'{label}.json', dict(argv=argv, cwd=str(cwd), started_utc=start,
          finished_utc=now(), exit_code=exit_code, timed_out=timed_out, stdout=meta(out), stderr=meta(err)))
    if exit_code:
        raise RuntimeError(f'{label} failed with exit {exit_code}')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--example', required=True, type=Path)
    parser.add_argument('--attempt', required=True)
    args = parser.parse_args()
    if not args.attempt or not all(c.isalnum() or c in '-_' for c in args.attempt):
        raise ValueError('attempt must be a path-safe token')
    example = args.example.resolve(strict=True)
    retained = ROOT / 'consumer' / args.attempt
    retained.mkdir(parents=True, exist_ok=False)
    work = TEMP / f'consumer-{args.attempt}'
    work.mkdir(parents=True, exist_ok=False)
    text = 'Independent consumer append <&> café'
    rows = []
    for case in ('plain', 'section'):
        folder = retained / case
        folder.mkdir()
        source = folder / 'source.docx'
        candidate = folder / 'candidate.docx'
        shutil.copyfile(ROOT / 'fuzz/seeds' / f'{case}-deflate.docx', source)
        run_command([str(example), str(source), str(candidate), text], 'append', REPO, folder)
        text_dir = folder / 'text'
        text_dir.mkdir()
        profile = work / f'profile-{case}'
        run_command(['/usr/bin/libreoffice', '--headless', '-env:UserInstallation=' + profile.as_uri(),
                     '--convert-to', 'txt:Text (encoded):UTF8', '--outdir', str(text_dir), str(source), str(candidate)],
                    'libreoffice', REPO, folder)
        before_raw = (text_dir / 'source.txt').read_bytes()
        after_raw = (text_dir / 'candidate.txt').read_bytes()
        before = before_raw.decode('utf-8-sig').replace('\r\n', '\n')
        after = after_raw.decode('utf-8-sig').replace('\r\n', '\n')
        expected_before = ' seed & text \n'
        expected_after = expected_before + text + '\n'
        row = dict(case=case, source=meta(source), candidate=meta(candidate),
                   source_text=meta(text_dir / 'source.txt'), candidate_text=meta(text_dir / 'candidate.txt'),
                   decoded_source_text=before, decoded_candidate_text=after,
                   expected_source_text=expected_before, expected_candidate_text=expected_after,
                   verified=before == expected_before and after == expected_after)
        write(folder / 'readback.json', row)
        rows.append(row)
        if not row['verified']:
            raise RuntimeError(f'{case}: independent consumer text differs; retained raw output for review')
    write(retained / 'result.json', dict(status='pass', finished_utc=now(),
          example=dict(path=str(example), **meta(example)), rows=rows,
          scope='LibreOffice synthetic DOCX open and TXT export; no native Microsoft Office claim'))
    print('LibreOffice independently read the source and appended candidate for both fixtures.')


if __name__ == '__main__':
    main()
