"""Static generated-archive span eligibility, not dynamic read or allocation counts."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import zipfile
P=Path(__file__).resolve().parent
ROOT=P.parents[3]
PARENT='66b944b07fcf4c6fa37f10e7ed83a02417455f57'
def desc(f):return {'path':str(f.relative_to(ROOT)),'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()}
def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument("--output",type=Path,default=P/"span-eligibility.json");args=parser.parse_args()
    plan=json.loads((P/'capture-plan.json').read_text());leg=next(x for x in plan['legs'] if x['route']=='owned' and x['variant']=='baseline');source=P/'capture'/leg['label']/'artifacts/source.pptx'
    code={name:subprocess.check_output(['git','show',PARENT+':'+name],cwd=ROOT) for name in ['crates/soapberry-zip/src/office.rs','crates/soapberry-zip/src/archive.rs']}
    assert b'const MAX_MEMBER_SPAN_READ_BYTES: u64 = 64 * 1024;' in code['crates/soapberry-zip/src/office.rs']
    assert b'const STRICT_LOCAL_HEADER_WINDOW: usize = 640;' in code['crates/soapberry-zip/src/archive.rs']
    assert b'const MAX_DATA_DESCRIPTOR_SIZE: u64 = 24;' in code['crates/soapberry-zip/src/archive.rs']
    rows=[]
    with zipfile.ZipFile(source) as archive:
        infos=sorted(archive.infolist(),key=lambda i:i.header_offset);directory=archive.start_dir
        for index,entry in enumerate(infos):
            base=entry.header_offset;next_offset=infos[index+1].header_offset if index+1<len(infos) else directory;descriptor=24 if entry.flag_bits&8 else 0
            length=min(min(next_offset,directory)+descriptor,base+640+entry.compress_size+descriptor,directory)-base
            eligible=base<directory and 30<=length<=65536
            rows.append({'name':entry.filename,'directory':entry.is_dir(),'compressed_bytes':entry.compress_size,'decoded_bytes':entry.file_size,'proposed_span_bytes':length,'eligible':eligible})
    files=[r for r in rows if not r['directory']];media=[r for r in files if r['name'].startswith('ppt/media/')];assert len(media)==8 and all(not r['eligible'] for r in media)
    result={'status':'pass','scope':'Static per-physical-member admission bound; not measured calls, allocation counts, or dynamic copied bytes.','source':desc(source),'code_commit':PARENT,'code':[{'path':n,'bytes':len(b),'sha256':hashlib.sha256(b).hexdigest()} for n,b in code.items()],'members':len(rows),'file_members':len(files),'eligible_file_members':sum(r['eligible'] for r in files),'one_read_per_eligible_file_span_bytes':sum(r['proposed_span_bytes'] for r in files if r['eligible']),'media_compressed_bytes':sum(r['compressed_bytes'] for r in media),'media_decoded_bytes':sum(r['decoded_bytes'] for r in media),'media_all_already_passthrough':True,'rows':rows}
    with args.output.open('x') as out:json.dump(result,out,indent=2);out.write('\n')
    print({k:v for k,v in result.items() if k not in ('rows','code','source')})
if __name__=='__main__':main()
