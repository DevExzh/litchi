"""Observe file-ID reuse in the owned test filesystem; no production execution."""
import json
from pathlib import Path
import tempfile
from run import HERE, TARGET, write


def identity(path):
    value=path.stat();return [value.st_dev,value.st_ino]


if __name__=='__main__':
    observations=[]
    with tempfile.TemporaryDirectory(prefix='identity-probe-',dir=TARGET/'test-tmp') as directory:
        root=Path(directory)
        for mode in ('unlink','rename'):
            for iteration in range(100):
                path=root/'stage';displaced=root/'displaced'
                path.write_bytes(b'x'*4096);before=identity(path)
                if mode=='unlink':path.unlink()
                else:path.rename(displaced)
                path.write_bytes(b'attacker replacement');after=identity(path)
                observations.append(dict(mode=mode,iteration=iteration,before=before,after=after,reused=before==after))
                path.unlink()
                if displaced.exists():displaced.unlink()
    report=dict(scope='Filesystem-only identity experiment on owned disk-backed TMPDIR; supports failure mechanism but does not retrospectively observe the failed Rust test inode.',observations=observations,reused={mode:sum(r['reused'] for r in observations if r['mode']==mode) for mode in ('unlink','rename')},owned_probe_directory_absent=not Path(directory).exists())
    assert report['reused']['rename']==0
    write(HERE/'temp-identity-probe.json',report);print(json.dumps(report['reused']))
