"""Deliberate qualification-report corruption controls for the real validator."""
import copy, json
from run import P, read, projection, schedule

def run():
    rows=schedule(True); reports=[read(f'qualification/{i:03d}.json') for i in range(4)]
    controls=[]
    def check(name, index, mutate):
        d=copy.deepcopy(reports[index]);mutate(d)
        try:
            got=projection(d,rows[index]); assert got==read('oracle.json')[rows[index]['case']]
        except (AssertionError,KeyError,TypeError,ValueError):
            controls.append({'name':name,'rejected':True});return
        raise AssertionError('accepted corruption: '+name)
    def x(d):return d['results'][0]['source']['pptx_cross_copy']
    check('output digest',0,lambda d:x(d)['output_sha256'].__setitem__(0,'0'*64))
    check('semantic gate',0,lambda d:x(d)['gates'].__setitem__('semantic_output_verified',False))
    check('missing gate',0,lambda d:x(d)['gates'].pop('dependency_closure_verified'))
    check('corpus digest',0,lambda d:d['results'][0]['corpus'].__setitem__('archive_sha256','0'*64))
    check('sample order',0,lambda d:d['results'][0]['elapsed_ns'].__setitem__('sample_order',[1]))
    check('phase nesting',0,lambda d:x(d)['plan_ns'].__setitem__(0,10**18))
    check('missing sample',0,lambda d:x(d)['reopen_ns'].clear())
    check('wrong executable',0,lambda d:d['binary_identity'].__setitem__('binary_sha256','0'*64))
    check('wrong affinity',0,lambda d:d['environment'].__setitem__('cpu_affinity','11'))
    check('wrong warmups',0,lambda d:d['configuration'].__setitem__('warmup_iterations_per_case',9))
    check('missing allocation',2,lambda d:d['results'][0]['operation_metrics']['allocation'].__setitem__('status','unavailable'))
    check('negative allocation',2,lambda d:d['results'][0]['operation_metrics']['allocation']['allocated_bytes']['values'].__setitem__(0,-1))
    result={'status':'passed','controls':controls}
    (P/'negative.json').write_text(json.dumps(result,indent=2)+'\n')
    print(f'PASS {len(controls)} qualification corruption controls')

if __name__=='__main__':run()
