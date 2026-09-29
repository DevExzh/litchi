"""Add corrected-reader admission before the unchanged profile schedule."""
import driver as d
import runpy,sys
assert sys.argv[1:]==['profiles']
value=d.read(d.P/'admission-v2.json')
assert all(d.desc(x['path'])==x for x in value['files'])
assert all(d.desc(x['path'])==x for x in d.read(d.P/'admission.json')['files'])
runpy.run_path(str(d.P/'capture.py'),run_name='__main__')
