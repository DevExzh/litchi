"""Run from repository root; writes deterministic diagnostic fixtures to argv[1]."""
from pathlib import Path
import hashlib, json, sys, zipfile
out=Path(sys.argv[1]);out.mkdir(parents=True,exist_ok=True)
manifest={}
for size in [96,128,160,164,256]:
    if size >=160:
        archive=Path(f'docs/performance/results/change-0544/cap-boundary/baseline/cap-native-r1-{size}.fixture.bin')
        with zipfile.ZipFile(archive) as z:data=z.read('xl/worksheets/sheet1.xml')
        provenance={'archive':str(archive),'archive_sha256':hashlib.sha256(archive.read_bytes()).hexdigest(),'member':'xl/worksheets/sheet1.xml'}
    else:
        def column(n):
            s=''
            while n:n,r=divmod(n-1,26);s=chr(65+r)+s
            return s
        data=('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'+''.join('<row r="'+str(r)+'">'+''.join(f'<c r="{column(c)}{r}"><v>{r*size+c}</v></c>' for c in range(1,size+1))+'</row>' for r in range(1,size+1))+'</sheetData></worksheet>').encode()
        provenance={'generator':'generate.py: synthetic numeric XML, distinct from primary workflow fixture'}
    name=f'grid-{size}.xml';(out/name).write_bytes(data)
    manifest[str(size)]={'file':name,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),**provenance}
for name,data in [('sparse',b'<x>'+b'a'*(1024*1024)+b'</x>'),('early-reject',b'<!--x-->'*131073+b'<x>'+b'a'*(1024*1024)+b'</x>')]:
    file=name+'.xml';(out/file).write_bytes(data)
    manifest[name]={'file':file,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest(),'generator':'generate.py: deterministic lexical stress, not Office workflow'}
(out/'fixtures.json').write_text(json.dumps(manifest,indent=2)+'\n')
