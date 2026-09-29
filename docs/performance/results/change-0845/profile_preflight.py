"""Exercise profile bindings/qualification without executing an observer."""
import driver as d
import profile as ip
value=ip.plan()
index={str(path.resolve()):d.desc(path) for path in d.P.rglob('*') if path.is_file()}
bindings=ip.validate_bindings(value,index,require_current_after=True)
checked=[]
for leg in ['before','after']:
    selected=ip.qualification_index(leg,index)
    assert len(selected)==4
    for shape,parts in selected:
        row=next(x for x in value['rows'] if x['leg']==leg and x['shape']==shape and x['parts']==parts)
        result=ip.expected_qualification(row,selected,bindings,index)
        checked.append(dict(leg=leg,shape=shape,parts=parts,report=d.desc(result['report_path'])))
assert len(checked)==8
d.write(d.P/'profile-preflight.json',dict(status='pass',plan=d.desc(d.P/'profile-plan.json'),runner=d.desc(d.P/'profile.py'),checker=d.desc(d.P/'profile_preflight.py'),qualification_reports=checked,note='Pure existing-artifact/binary binding preflight; no workload or timing sample executed.'))
print('Profile binding preflight PASS',len(checked),'qualification reports')
