"""Regression for inline-file return tokens, plus every inherited reader check."""
import contextlib
import io
from pathlib import Path
import tempfile
import driver as d
import reader as old
import reader_v2 as r
import reader_tests as inherited

def main():
    inherited.r=r
    passed=[];rejected=[]
    inherited._zip_checks(rejected,passed)
    inherited._profile_checks(rejected,passed)
    qualification_count=inherited._retained_qualification_check(passed)
    assert qualification_count==64 and len(rejected)==34
    regressions=[]
    with tempfile.TemporaryDirectory(prefix='litchi-0845-reader-v2-') as temporary:
        root=Path(temporary)
        for label,return_token in [('compressed','fe=(10)'),('declared','fe=(10) probe/src/main.rs'),('plain','fe=probe/src/main.rs')]:
            numbered,terminal,command=inherited._profile_fixture(root/label)
            text=numbered.read_text().replace('fl=probe/src/main.rs','fl=(10) probe/src/main.rs')
            text=text.replace('1 3\n','fi=(90) probe/src/inline.rs\n1 1\n'+return_token+'\nfn=(2)\n1 2\n')
            numbered.write_text(text)
            try:old.validate_profile(numbered,terminal,command=command)
            except old.ReaderError as error:assert 'function source file changed' in str(error)
            else:raise AssertionError('legacy reader did not reproduce the inline return failure')
            value=r.validate_profile(numbered,terminal,command=command)
            assert value['summary_ir']==10 and value['owner']['self_ir']==3 and value['owner']['direct_children_ir']==7
            assert value['owner']['file']=='probe/src/main.rs'
            assert all(value['validation'].values())
            regressions.append(label+'-inline-return')
            broken=text.replace(return_token+'\n','')
            numbered.write_text(broken)
            try:r.validate_profile(numbered,terminal,command=command)
            except r.ReaderError as error:assert 'function source file changed' in str(error)
            else:raise AssertionError('missing return incorrectly accepted')
            regressions.append(label+'-missing-return-rejected')
        numbered,terminal,command=inherited._profile_fixture(root/'callee-alias')
        text=numbered.read_text().replace('cfn=(3) child','cfl=(11) probe/src/archive.rs\ncfn=(3) child')
        numbered.write_text(text)
        value=r.validate_profile(numbered,terminal,command=command)
        assert value['summary_ir']==10 and value['owner']['file']=='probe/src/main.rs'
        regressions.append('callee-file-alias-does-not-change-caller')
        numbered,terminal,command=inherited._profile_fixture(root/'conflicting-file-id')
        text=numbered.read_text().replace('fl=probe/src/main.rs','fl=(10) probe/src/main.rs').replace('1 3\n','fe=(10) other.rs\n1 3\n')
        numbered.write_text(text)
        try:r.validate_profile(numbered,terminal,command=command)
        except r.ReaderError as error:assert 'conflicting names' in str(error)
        else:raise AssertionError('conflicting source file ID accepted')
        regressions.append('conflicting-file-id-rejected')
    captures=[]
    folders=sorted((d.P/'profile-runs').iterdir());assert len(folders)==4
    for folder in folders:
        command=d.read(d.P/'commands'/('profile-'+folder.name)/'receipt.json');assert command['exit_code']==0
        numbered=folder/'profile.callgrind.1';terminal=folder/'profile.callgrind'
        value=r.validate_profile(numbered,terminal,command=command['argv'])
        error_text=None
        try:
            prior=old.validate_profile(numbered,terminal,command=command['argv'])
            assert prior['summary_ir']==value['summary_ir']
            assert prior['owner']['self_ir']==value['owner']['self_ir'] and prior['owner']['direct_children_ir']==value['owner']['direct_children_ir']
        except old.ReaderError as error:error_text=str(error)
        captures.append(dict(numbered=d.desc(numbered),terminal=d.desc(terminal),command=d.desc(d.P/'commands'/('profile-'+folder.name)/'receipt.json'),summary_ir=value['summary_ir'],legacy_error=error_text))
    assert sum(x['legacy_error'] is not None for x in captures)==1
    d.write(d.P/'reader-v2-tests.json',dict(status='pass',reader=d.desc(d.P/'reader_v2.py'),previous=d.desc(d.P/'reader.py'),test_source=d.desc(d.P/'reader_v2_tests.py'),inherited_test_source=d.desc(d.P/'reader_tests.py'),inherited_checks=passed,inherited_mutations_rejected=rejected,qualification_reports=qualification_count,regressions=regressions,captures=captures))
    print('Reader v2 PASS:',qualification_count,'qualifications;',len(rejected),'inherited mutations;',len(regressions),'grammar checks;',len(captures),'retained profiles')
if __name__=='__main__':main()
