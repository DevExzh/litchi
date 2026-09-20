#!/usr/bin/env python3
"""Resolve strace file offsets through ELF executable segments and symbol ranges."""
import hashlib,json,re,struct,subprocess
from pathlib import Path
P=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def main():
    dest=P/'symbolization.json';assert not dest.exists()
    binary=json.loads((P/'build.json').read_text())['binary'];exe=Path(binary['path']);data=exe.read_bytes();assert sha(exe)==binary['sha256'] and len(data)==binary['bytes']
    header=struct.unpack_from('<16sHHIQQQIHHHHHH',data);assert header[0][:6]==b'\x7fELF\x02\x01' and header[2]==62
    segments=[]
    for i in range(header[10]):
        kind,flags,offset,vaddr,physical,filesz,memsz,align=struct.unpack_from('<IIQQQQQQ',data,header[5]+i*header[9])
        if kind==1 and flags&1:segments.append(dict(file_offset=offset,virtual_address=vaddr,file_size=filesz,memory_size=memsz,flags=flags,alignment=align))
    assert segments
    traces=sorted(P.glob('mapping-r*.strace'));assert len(traces)==4
    pattern=re.compile(re.escape(str(exe))+r'\([^\n]*\) \[(0x[0-9a-fA-F]+)\]')
    offsets=sorted({int(m,16) for f in traces for m in pattern.findall(f.read_text())});assert offsets
    frames={}
    for offset in offsets:
        found=[s for s in segments if s['file_offset']<=offset<s['file_offset']+s['file_size']];assert len(found)==1,(offset,found)
        s=found[0];frames[hex(offset)]=dict(file_offset=offset,virtual_address=offset-s['file_offset']+s['virtual_address'])
    commands=[('program-headers',['/usr/bin/readelf','-lW',str(exe)],None),('symbols',['/usr/bin/nm','-n','-S','-C','--defined-only',str(exe)],None),('resolved',['/usr/bin/addr2line','-f','-C','-e',str(exe)],''.join(hex(f['virtual_address'])+'\n' for f in frames.values()))]
    records=[]
    for name,cmd,input_text in commands:
        r=subprocess.run(cmd,input=input_text,text=True,capture_output=True);assert r.returncode==0 and not r.stderr,(cmd,r.stderr)
        f=P/('symbolization-'+name+'.txt');f.write_text(r.stdout)
        if input_text is not None:(P/'symbolization-addresses.txt').write_text(input_text)
        records.append(dict(command=cmd,exit_code=r.returncode,stdout=f.name,sha256=sha(f)))
    symbols=[]
    for line in (P/'symbolization-symbols.txt').read_text().splitlines():
        m=re.fullmatch(r'([0-9a-f]+) ([0-9a-f]+) ([TtWw]) (.+)',line)
        if m and int(m[2],16)>0:symbols.append((int(m[1],16),int(m[2],16),m[4]))
    resolved=(P/'symbolization-resolved.txt').read_text().splitlines();assert len(resolved)==2*len(frames)
    for i,frame in enumerate(frames.values()):
        function,location=resolved[2*i:2*i+2];addr=frame['virtual_address'];matches=[dict(address=a,bytes=z,name=n) for a,z,n in symbols if a<=addr<a+z]
        assert function!='??' and matches,(frame,function)
        assert any(function==m['name'] for m in matches),(function,matches)
        frame.update(function=function,location=location,covering_symbols=matches)
    source=P/'strace-unwind-libunwind.c.txt'
    result=dict(binary=binary,method='Strace libunwind true_offset is a mapped-file offset; translate through the unique executable PT_LOAD segment. Resolve ELF virtual addresses with addr2line and require exact name agreement with a covering nm symbol range.',segments=segments,frames=frames,traces={f.name:sha(f) for f in traces},commands=records,tools={n:dict(sha256=sha(Path(n)),bytes=Path(n).stat().st_size) for n in ['/usr/bin/readelf','/usr/bin/nm','/usr/bin/addr2line']},source_reference=dict(url='https://raw.githubusercontent.com/strace/strace/v6.19/src/unwind-libunwind.c',file=source.name,sha256=sha(source)),addresses_sha256=sha(P/'symbolization-addresses.txt'),script_sha256=sha(P/'symbolize.py'))
    dest.write_text(json.dumps(result,indent=2)+'\n');print('PASS',len(frames),'exact file-offset/function bindings')
if __name__=='__main__':main()
