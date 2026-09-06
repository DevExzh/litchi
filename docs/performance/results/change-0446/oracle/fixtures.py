"""Independent bounded ZIP32/XML and deterministic payload oracle. No extraction."""
import hashlib, io, json, struct, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET
ROOT = Path(__file__).resolve().parent.parent
SHAPES = {'tiny':64,'medium':1024,'large':4096}
ADDED='benchmark/added/leaf.bin'
ADDED_TYPE='application/vnd.litchi.perf.added-part'
REL_TYPE='urn:litchi:perf:relationships:added-part'
CT='[Content_Types].xml'
RELS='_rels/.rels'
GATES={'payload_and_topology_verified','raw_untouched_members_verified','content_types_lexical_verified','exact_noop_verified','duplicate_part_refusal_verified','missing_target_refusal_verified','stale_source_refusal_verified','short_read_verified','partial_sink_verified','output_limit_verified','part_limit_verified'}
def require(condition, message):
    if not condition: raise ValueError(message)
def sha(raw): return hashlib.sha256(raw).hexdigest()
def payload(index, size, compressible):
    if compressible:
        block=b'litchi-perf-baseline-compressible-payload-v1\n'
        return bytes(block[(offset+index)%len(block)] for offset in range(size))
    mask=(1<<64)-1
    state=(index*0x9e3779b97f4a7c15+0xd1b54a32d192ed03)&mask
    output=bytearray()
    for _ in range(size):
        state^=(state<<13)&mask
        state^=state>>7
        state^=(state<<17)&mask
        output.append((state>>24)&255)
    return bytes(output)
def inspect(raw, expected):
    require(0<len(raw)<16*1024*1024,'archive byte bound')
    z=zipfile.ZipFile(io.BytesIO(raw)); infos=z.infolist()
    require(len(infos)==len(expected) and set(z.namelist())==expected,'exact unique member names')
    decoded={}
    for info in infos:
        require(info.file_size<=1024*1024 and info.compress_type==8 and not info.flag_bits&1,'member resource/codec bound')
        decoded[info.filename]=z.read(info)
    physical=sorted(infos,key=lambda v:v.header_offset)
    locals={v.filename:raw[v.header_offset:(physical[i+1].header_offset if i+1<len(physical) else z.start_dir)] for i,v in enumerate(physical)}
    central={}; cursor=z.start_dir
    for info in infos:
        require(raw[cursor:cursor+4]==b'PK\x01\x02','central header')
        n,e,c=struct.unpack_from('<HHH',raw,cursor+28)
        record=bytearray(raw[cursor:cursor+46+n+e+c]);record[42:46]=b'\0'*4
        central[info.filename]=bytes(record);cursor+=len(record)
    require(raw[cursor:cursor+4]==b'PK\x05\x06','ZIP32 end record')
    return decoded,locals,central,[v.filename for v in infos],[v.filename for v in physical],z.comment

def verify_fixture(shape, directory=None, bind=True):
    directory=Path(directory) if directory else ROOT/'fixtures'
    count=SHAPES[shape];names={f'benchmark/parts/{i:05}.bin' for i in range(count)}|{CT,RELS}
    source=(directory/f'{shape}-source.zip').read_bytes();output=(directory/f'{shape}-output.zip').read_bytes()
    before=inspect(source,names);after=inspect(output,names|{ADDED})
    for i in range(count):
        name=f'benchmark/parts/{i:05}.bin';expected=payload(i,1024,i%2==0)
        require(before[0][name]==expected and after[0][name]==expected,'ordinary payload formula/preservation')
        require(before[1][name]==after[1][name] and before[2][name]==after[2][name],'raw untouched local/central records')
    added=payload(444,65536,False)
    require(after[0][ADDED]==added,'added payload formula')
    for index in (3,4):require(before[index]==[n for n in after[index] if n!=ADDED],'original member order')
    require(before[5]==after[5],'ZIP comment preservation')
    override=f'<Override PartName="/{ADDED}" ContentType="{ADDED_TYPE}"/>'.encode()
    require(before[0][CT].count(b'</Types>')==1,'source CT close')
    require(after[0][CT]==before[0][CT].replace(b'</Types>',override+b'</Types>'),'exact CT lexical insertion')
    ns='{http://schemas.openxmlformats.org/package/2006/content-types}'
    for content,is_output in ((before[0][CT],False),(after[0][CT],True)):
        root=ET.fromstring(content);require(root.tag==ns+'Types','CT root')
        overrides=[v.attrib for v in root if v.tag==ns+'Override']
        found={v['PartName']:v['ContentType'] for v in overrides}
        expected={'/'+n:'application/octet-stream' for n in names-{CT,RELS}}
        if is_output:expected['/'+ADDED]=ADDED_TYPE
        require(len(found)==len(overrides) and found==expected,'content type declarations')
    rns='{http://schemas.openxmlformats.org/package/2006/relationships}'
    old={'Id':'rIdBenchmarkMain','Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument','Target':f'benchmark/parts/{count//2:05}.bin'}
    new={'Id':'rIdPartAddition','Type':REL_TYPE,'Target':ADDED}
    for content,expected in ((before[0][RELS],[old]),(after[0][RELS],[old,new])):
        root=ET.fromstring(content); require(root.tag==rns+'Relationships','relationship root')
        actual=[]
        for v in root:
            require(v.tag==rns+'Relationship','relationship child')
            attrs=dict(v.attrib)
            require(attrs.pop('TargetMode','Internal')=='Internal','relationship mode')
            actual.append(attrs)
        require(sorted(actual,key=lambda v:v['Id'])==sorted(expected,key=lambda v:v['Id']),'relationship topology')
    summary=json.loads((directory/f'{shape}.json').read_text())
    expected={'shape':shape,'source_part_count':count,'output_part_count':count+1,'source_member_count':count+2,'output_member_count':count+3,'source_archive_sha256':sha(source),'output_archive_sha256':sha(output),'source_archive_bytes':len(source),'output_archive_bytes':len(output),'added_member':ADDED,'added_content_type':ADDED_TYPE,'added_payload_bytes':len(added),'added_payload_sha256':sha(added),'relationship_id':new['Id'],'relationship_type':REL_TYPE,'role':'part_addition','source_mode':'instrumented_read_at_v1','lifecycle_ns':[],'output_sha256':[]}
    for key,value in expected.items():require(summary.get(key)==value,'fixture metadata '+key)
    require(set(summary['gates'])==GATES and all(v is True for v in summary['gates'].values()),'producer refusal/preservation gate declarations')
    if bind:
        contract=json.loads((ROOT/'oracle/protocol.json').read_text())
        for suffix in ('-source.zip','-output.zip','.json'):
            name=shape+suffix;raw=(directory/name).read_bytes()
            require(contract['fixtures'][name]=={'bytes':len(raw),'sha256':sha(raw)},'frozen fixture identity '+name)
    return {'source_sha':sha(source),'source_bytes':len(source),'output_sha':sha(output),'output_bytes':len(output)}
if __name__=='__main__':
    for shape in SHAPES:print(shape,verify_fixture(shape,bind=False))
