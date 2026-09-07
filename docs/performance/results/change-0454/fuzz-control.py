#!/usr/bin/env python3
"""Prepare/capture/clean only the exclusive OPC fuzz task; never run CPU jobs."""
import hashlib,json,shutil,sys,zipfile,io
from pathlib import Path
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[3]
TASK=Path('/tmp/litchi-goal-0454-opc-fuzz')
def sha(raw):return hashlib.sha256(raw).hexdigest()
def row(path,base):
    assert path.is_file() and not path.is_symlink()
    raw=path.read_bytes();return {'path':str(path.relative_to(base)),'bytes':len(raw),'sha256':sha(raw)}
def write(name,value):
    p=ROOT/name;assert not p.exists();p.write_text(json.dumps(value,indent=2,sort_keys=True)+'\n')
def verify_inputs():
    proof=json.loads((ROOT/'checks/fuzz-prepared.json').read_text())
    for r in proof['inputs']:
        assert row(TASK/r['path'],TASK)==r
    assert proof['driver_sha256']==sha(Path(__file__).read_bytes())
    return proof
def main():
    assert TASK==Path('/tmp/litchi-goal-0454-opc-fuzz') and not TASK.is_symlink()
    mode=sys.argv[1]
    if mode=='prepare':
        assert not TASK.exists();TASK.mkdir(mode=0o700)
        (TASK/'fuzz_targets').mkdir();(TASK/'corpus').mkdir();(ROOT/'seeds').mkdir()
        src=REPO/'crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs'
        (TASK/'fuzz_targets/parse_opc.rs').write_bytes(src.read_bytes())
        manifest=(REPO/'crates/litchi-opc/fuzz/Cargo.toml').read_text().replace('path = ".."', 'path = '+json.dumps(str(REPO/'crates/litchi-opc')))
        (TASK/'Cargo.toml').write_text(manifest)
        (ROOT/'checks/fuzz-Cargo.toml.txt').write_text(manifest)
        for p in sorted((ROOT.parent/'change-0450/seeds').glob('*.zip')):
            shutil.copyfile(p,ROOT/'seeds'/p.name)
        shutil.copyfile(REPO/'test-data/poi/test-data/slideshow/EmbeddedVideo.pptx',ROOT/'seeds/native-EmbeddedVideo.pptx')
        entries={'[Content_Types].xml':b'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>', '_rels/.rels':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>', 'word/document.xml':b'<document/>', 'custom/source.bin':bytes(range(256))*32}
        for method,label in [(zipfile.ZIP_STORED,'store'),(zipfile.ZIP_DEFLATED,'deflate')]:
            out=io.BytesIO()
            with zipfile.ZipFile(out,'w') as z:
                for name,data in entries.items():
                    info=zipfile.ZipInfo(name,date_time=(1980,1,1,0,0,0));info.compress_type=method;z.writestr(info,data)
            (ROOT/'seeds'/('combined-'+label+'.zip')).write_bytes(out.getvalue())
        variants={
            'prefixed':b'<r:Relationships xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships"><!--kept--><r:Relationship Target=\'word/document.xml\' Type=\'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\' Id=\'rId1\'></r:Relationship><?kept data?></r:Relationships>',
            'selfclosing':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" />',
            'unknown-attribute':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships" strange="value"/>',
            'dtd':b'<!DOCTYPE Relationships [<!ENTITY target "word/document.xml">]><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
            'nested':b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="urn:type" Target="word/document.xml"><Relationship/></Relationship></Relationships>',
        }
        for label,rels in variants.items():
            out=io.BytesIO()
            with zipfile.ZipFile(out,'w') as z:
                for name,data in entries.items():
                    info=zipfile.ZipInfo(name,date_time=(1980,1,1,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;z.writestr(info,rels if name=='_rels/.rels' else data)
            (ROOT/'seeds'/('relationships-'+label+'.zip')).write_bytes(out.getvalue())
        seeds=[row(p,ROOT) for p in sorted((ROOT/'seeds').iterdir())];assert len(seeds)==21
        write('seed-manifest.json',seeds)
        for r in seeds:shutil.copyfile(ROOT/r['path'],TASK/'corpus'/Path(r['path']).name)
        write('checks/fuzz-prepared.json',{'status':'pass','task':str(TASK),'driver_sha256':sha(Path(__file__).read_bytes()),'inputs':[row(p,TASK) for p in sorted(TASK.rglob('*')) if p.is_file()]})
    elif mode=='capture':
        verify_inputs()
        for name in ['Cargo.lock']:
            (ROOT/'checks'/('fuzz-'+name+'.txt')).write_bytes((TASK/name).read_bytes())
        artifacts=[row(p,TASK) for p in sorted(TASK.rglob('*')) if p.is_file()]
        write('checks/fuzz-artifacts.json',{'status':'pass','task':str(TASK),'artifacts':artifacts,'files':len(artifacts),'bytes':sum(r['bytes'] for r in artifacts)})
    elif mode=='cleanup':
        verify_inputs();assert json.loads((ROOT/'checks/precleanup.json').read_text())['status']=='pass'
        proof=json.loads((ROOT/'checks/fuzz-artifacts.json').read_text())
        assert [row(p,TASK) for p in sorted(TASK.rglob('*')) if p.is_file()]==proof['artifacts']
        shutil.rmtree(TASK)
        write('checks/fuzz-cleanup.json',{'status':'pass','task':str(TASK),'temporary_directory_absent':not TASK.exists(),'files_removed':proof['files'],'bytes_removed':proof['bytes'],'artifact_manifest_sha256':sha((ROOT/'checks/fuzz-artifacts.json').read_bytes())})
    else:raise ValueError(mode)
    print(json.dumps({'status':'pass','action':mode,'task':str(TASK)}))
if __name__=='__main__':main()
