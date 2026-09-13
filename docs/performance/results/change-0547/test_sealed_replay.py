"""Regression check: removed executable requires matching cleanup custody."""
from pathlib import Path
import copy,importlib.util,json
B=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('profiles0547',B/'analyze_profiles.py')
m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m)
plan=m.plan_data();identity=m._validate_build(plan)
assert identity['sha256']==json.loads((B/'cleanup.json').read_text())['binary_sha256']
original=m.load_json
for change in ['missing','binary','references','removed']:
    def altered(path):
        value=original(path)
        if Path(path).name!='cleanup.json':return value
        if change=='missing':raise m.EvidenceError('cleanup unavailable')
        value=copy.deepcopy(value)
        if change=='binary':value['binary_sha256']='0'*64
        if change=='references':value['accessible_process_references']=['live process']
        if change=='removed':value['removed']=[]
        return value
    m.load_json=altered
    try:m._validate_build(plan)
    except m.EvidenceError:pass
    else:raise AssertionError('invalid cleanup admitted: '+change)
    finally:m.load_json=original
print('PASS: valid sealed cleanup accepted; missing, wrong-binary, live-reference and incomplete-removal custody refused')
