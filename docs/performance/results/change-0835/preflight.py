"""Root-owned independent custody and statistics checks before formal capture."""
import ast
import driver as d
import audit,analyze

def main():
 for path in d.P.glob('*.py'):ast.parse(path.read_text(),filename=str(path))
 audit.load_origin();audit.validate_freeze();audit.validate_host();audit.validate_quality_reuse()
 build,command=audit.validate_build();audit.validate_qualification(build['binary']);audit.validate_plan()
 analyze.validate_plan(d.read(d.P/'measurement-plan.json'))
 values=[30,10,10,20]*7+[5,40]
 assert len(values)==30
 expected=dict(min=5,p50=15,p95=30,p99=40,max=40,mean=sum(values)/30)
 assert analyze.summary(values)==expected
 assert audit.sample_summary(values)==expected
 d.write(d.P/'preflight.json',dict(status='pass',plan=d.desc(d.P/'measurement-plan.json'),reader=d.desc(d.P/'reader.py'),analyzer=d.desc(d.P/'analyze.py'),auditor=d.desc(d.P/'audit.py'),checks=['Python syntax','source/host/quality/build/qualification custody','both statistics plan validators','30-value midpoint/nearest-rank oracle']))
 print('independent preflight PASS')
if __name__=='__main__':main()
