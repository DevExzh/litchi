"""Attribute planning-only Ir; retain lifecycle dumps and verify exact raw edges."""
import argparse
import hashlib
import importlib.util
import json
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / 'change-0521/analyze_profiles.py'
spec = importlib.util.spec_from_file_location('planning_profile_helpers', HELPER)
h = importlib.util.module_from_spec(spec)
spec.loader.exec_module(h)
h.HERE = HERE

def display_name(text):
    # Only the final bracketed field is the executable. Rust symbols can
    # contain earlier brackets, e.g. core::slice::<impl [T]>::sort_unstable_by.
    text = text.rsplit(' [', 1)[0].strip()
    text = re.sub(r'\s+\([\d,]+x\)$', '', text)
    if text.startswith('???:'):
        return text[4:]
    return text.rsplit(':', 1)[-1] if text.startswith(('./', '/')) else text

h.display_name = display_name
PARENT = 'litchi_perf_baseline::run_xlsx_cell_values_edit_save'
LIFECYCLE = 'litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates'

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def analyze(create_annotations=False):
    plan = json.loads((HERE/'plan.json').read_text())
    observation=json.loads((HERE/'symbol-observation.json').read_text())
    assert observation['plan_sha256']==sha(HERE/'plan.json')
    owner=observation['owner']
    assert owner in plan['profile']['owner_candidates']
    rows=[]
    for repeat in range(1,plan['profile']['repeats']+1):
        for shape in plan['profile']['shapes']:
            stem=HERE/'baseline'/f'profile-r{repeat}-{shape}'
            raw=Path(str(stem)+'.callgrind')
            numbered=sorted(raw.parent.glob(raw.name+'.[0-9]*'),key=lambda p:int(p.suffix[1:]))
            assert numbered, 'No owner dumps captured'
            assert [int(p.suffix[1:]) for p in numbered] == list(range(1,len(numbered)+1))
            parts=[]; selected=[]
            for path in numbered:
                txt=path.read_text();total=h.summary_ir(txt,str(path))
                edge=h.target_edge_summary(path,owner,PARENT)
                life=h.target_edge_summary(path,owner,LIFECYCLE)
                assert h.part_number(txt,str(path))==int(path.suffix[1:])
                assert owner in h.trigger(txt,str(path)), 'Unexpected dump trigger'
                parts.append(dict(path=str(path.relative_to(HERE)),sha256=sha(path),summary_ir=total,measured_parent=edge,lifecycle_parent=life))
                if edge['inclusive_ir']>0:
                    assert edge['positive_edge_count']==1 and edge['calls']==1
                    assert edge['inclusive_ir']==total
                    selected.append(path)
                else:
                    assert life['inclusive_ir']==total and life['positive_edge_count']==1
            assert len(selected)==1 and selected[0]==numbered[-1], 'Measured call must be unique and final'
            assert h.summary_ir(raw.read_text(),str(raw))==0
            assert h.trigger(raw.read_text(),str(raw))=='Program termination'
            paths=[];annotations=[]
            for inclusive,suffix in [(True,'.inclusive.txt'),(False,'.self.txt')]:
                path=Path(str(stem)+suffix)
                output,command=h.run_annotation(selected[0],inclusive)
                if create_annotations:
                    if path.exists():
                        assert path.read_text()==output, 'Existing annotation differs'
                    else:
                        path.write_text(output)
                else:
                    assert path.read_text()==output, 'Annotation replay differs'
                paths.append(path);annotations.append(output)
            inc,own=(h.parse_annotation(txt,owner,str(path)) for txt,path in zip(annotations,paths))
            direct=h.direct_map(inc['direct'])
            assert direct==h.direct_map(own['direct'])
            assert inc['selected_ir']==own['selected_ir']+sum(direct.values())==parts[-1]['summary_ir']
            owners={}
            # Follow material edges through six levels. A function's annotation
            # includes all of its callers within this selected profile; retain
            # parent-specific raw edge costs and never sum nested owner totals.
            pending=[(owner,0)]
            while pending:
                name,depth=pending.pop(0)
                if name in owners:
                    continue
                i=h.parse_annotation(annotations[0],name,str(paths[0]));s=h.parse_annotation(annotations[1],name,str(paths[1]))
                children=h.direct_map(i['direct']);assert children==h.direct_map(s['direct'])
                assert i['selected_ir']==s['selected_ir']+sum(children.values())
                for child,cost in children.items():
                    edge=h.target_edge_summary(selected[0],child,name)
                    assert edge['inclusive_ir']==cost,(name,child)
                owners[name]=dict(inclusive_ir=i['selected_ir'],self_ir=s['selected_ir'],direct=dict(sorted(children.items(),key=lambda x:(-x[1],x[0]))))
                if depth<6:
                    pending.extend((n,depth+1) for n,c in children.items() if c>=inc['selected_ir']*.05)
            rows.append(dict(repeat=repeat,shape=shape,parts=parts,termination=dict(path=str(raw.relative_to(HERE)),sha256=sha(raw),summary_ir=0),selected=str(selected[0].relative_to(HERE)),annotations={str(p.relative_to(HERE)):sha(p) for p in paths},owners=owners))
    return dict(status='pass',plan_sha256=sha(HERE/'plan.json'),helper_sha256=sha(HELPER),owner=owner,rows=rows,scope='Selected planning method Ir only. edit_sheets is the preferred native planning boundary; snapshot-loader fallback would exclude wrapper execution check and empty staging initialization. Selector construction is outside native planning. No latency improvement or production change is claimed.',limitations=['Nested owners overlap; only immediate children of one owner are disjoint.','Call metadata outside positive selected edges may contain collection-off work and is not an allocation count.','Generated two-shape corpus, one measured owner call per child; no cold/range/scaling claim.'])

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--create-annotations',action='store_true');parser.add_argument('--output',type=Path)
    args=parser.parse_args();result=analyze(args.create_annotations)
    output=args.output or HERE/'planning-analysis.json'
    output.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
    print('Planning profile edges and annotations verified:',output)
