"""Verify that independent replay rejects altered summaries and commands."""
import copy, json
import audit
from run import P, read
if __name__=='__main__':
    controls=[]
    def reject(name,path,mutate):
        original=(P/path).read_bytes(); value=json.loads(original); mutate(value)
        try:
            (P/path).write_text(json.dumps(value,indent=2)+'\n')
            try:audit.main()
            except (audit.AuditError,AssertionError,KeyError,TypeError,ValueError):
                controls.append({'name':name,'rejected':True})
            else:raise AssertionError('accepted corruption: '+name)
        finally:(P/path).write_bytes(original)
    reject('derived process median','analysis.json',lambda d:d['processes'][0]['stats']['lifecycle_ns'].__setitem__('p50',1))
    reject('group bootstrap interval','analysis.json',lambda d:d['groups'][0]['metrics_ns']['plan_ns'].__setitem__('bootstrap_median_95',[1,2]))
    reject('omitted process','analysis.json',lambda d:d['processes'].pop())
    reject('changed command','captures/manifest.json',lambda d:d[0]['command'].__setitem__(2,'11'))
    audit.main()
    (P/'audit-negative.json').write_text(json.dumps({'status':'passed','controls':controls},indent=2)+'\n')
    print('PASS four independent replay corruption controls; originals restored')
