"""Add the separately frozen ignored workspace lock to later source checks.

The baseline pipeline began before this wrapper was needed/discovered. Its
workspace lock was separately frozen before its first workspace build and is
checked again at pipeline completion. Later pipelines check it per child.
The original frozen run.py/capture.py remain byte-identical.
"""

import json
import runpy

import run as R

original_check = R.check_source
binding = json.loads((R.HERE / 'workspace-lock.json').read_text())


def checked_source():
    original_check()
    assert R.sha(R.REPO / binding['path']) == binding['sha256'], 'workspace lock changed'
    assert R.sha(R.HERE / binding['retained']) == binding['sha256'], 'retained lock changed'


if __name__ == '__main__':
    R.check_source = checked_source
    runpy.run_path(str(R.HERE / 'capture.py'), run_name='__main__')
