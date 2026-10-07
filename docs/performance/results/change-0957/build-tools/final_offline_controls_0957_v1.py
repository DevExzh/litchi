from pathlib import Path
import sys,json,tempfile,copy,hashlib
r=Path('/home/zhuhe/litchi-performance-0957');sys.path.insert(0,str(r/'tools-v2'))
import controls_0957_v2 as reports
import custody_controls_0957_v2 as custody
from capture_0957_v2 import load_and_check_plan
from common_0957_v2 import ContractError
plan=r/'plan-final-v2.json'
reports.PLAN=plan;reports.OUT=r/'review/root-report-controls-v1.json';assert reports.main()==0
custody.PLAN_PATH=plan;custody.OUT=r/'review/root-custody-controls-v1.json';assert custody.main()==0
assert len(load_and_check_plan(plan)[1])==27
p=json.loads(plan.read_text())
with tempfile.TemporaryDirectory() as tmp:
 wrong=Path(tmp)/'wrong-runner.py';wrong.write_text('# wrong runner\n')
 p['runner']={'path':str(wrong),'bytes':wrong.stat().st_size,'sha256':hashlib.sha256(wrong.read_bytes()).hexdigest()}
 trial=Path(tmp)/'plan.json';trial.write_text(json.dumps(p))
 try:load_and_check_plan(trial)
 except ContractError as e:assert 'executing' in str(e)
 else:raise AssertionError('wrong executing runner admitted')
(r/'review/root-runner-identity-control-v1.json').write_text(json.dumps({'status':'passed','wrong_executing_runner':'rejected','formal_children':27})+'\n')
