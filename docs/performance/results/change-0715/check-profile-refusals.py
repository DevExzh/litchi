#!/usr/bin/env python3
"""Exercise profile selection refusal without modifying retained captures."""
import importlib.util,json,re,shutil,tempfile
from pathlib import Path
P=Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('profile_negative0715',P/'analyze.py');A=importlib.util.module_from_spec(spec);spec.loader.exec_module(A)
T=A.module(P/'analyze_profiles.py','profile_parser_negative0715')
def main():
    target=P/'profile-negative-checks.json';assert not target.exists()
    assert (json.dumps(A.analyze(),indent=2,sort_keys=True)+'\n').encode()==(P/'profile-analysis.json').read_bytes()
    stem=P/'profile-r1-generated-counting_publish.callgrind'; paths=T.paths_for_stem(stem)
    protected={str(p):A.C.C.sha(p) for p in paths};checks=[]
    for name in ['missing-setup-part','wrong-measured-parent','changed-summary']:
        with tempfile.TemporaryDirectory(prefix='litchi-0715-profile-negative-') as directory:
            folder=Path(directory)
            for p in paths:shutil.copy2(p,folder/p.name)
            selected=folder/(stem.name+'.6')
            if name=='missing-setup-part':(folder/(stem.name+'.2')).unlink()
            elif name=='wrong-measured-parent':
                data=selected.read_text();assert T.MEASURED_PARENT in data
                selected.write_text(data.replace(T.MEASURED_PARENT,'litchi_perf_baseline::ordinary_save::wrong_parent'))
            else:
                data=selected.read_text();data,n=re.subn(r'^summary: (\d+)$',lambda m:'summary: '+str(int(m[1])+1),data,flags=re.M);assert n==1;selected.write_text(data)
            try:T.analyze_profile(folder/stem.name,'generated',A.C.read(P/'plan.json')['profile'])
            except ValueError as error:checks.append(dict(name=name,rejected=True,error=str(error)))
            else:raise AssertionError(name)
    assert all(A.C.C.sha(Path(n))==h for n,h in protected.items())
    A.C.write(target,dict(status='pass',checks=checks,retained_inputs_unchanged=True,temporary_inputs_removed=True,positive_exact_replay=True,analysis_sha256=A.C.C.sha(P/'profile-analysis.json'),parser_sha256=A.C.C.sha(P/'analyze_profiles.py')))
    print('PASS: three profile corruptions rejected, exact positive replay')
if __name__=='__main__':main()
