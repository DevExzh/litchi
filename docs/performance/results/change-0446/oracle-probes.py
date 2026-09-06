#!/usr/bin/env python3
"""Reject independently corrupted exported archives and serialized pilot reports."""
import copy, importlib.util, io, json, shutil, sys, tempfile, zipfile
from pathlib import Path
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'oracle'))
from fixtures import verify_fixture
spec=importlib.util.spec_from_file_location('report_oracle',ROOT/'oracle/verify-report.py')
oracle=importlib.util.module_from_spec(spec);spec.loader.exec_module(oracle)
def rejected(call, name):
    try:call()
    except (ValueError,KeyError,zipfile.BadZipFile):return name
    raise AssertionError('accepted mutation: '+name)
def main():
    rows=[]
    with tempfile.TemporaryDirectory(prefix='litchi-goal-0446-oracle-probes-') as temporary:
        directory=Path(temporary)
        for source_mode in ('plain',):
            pilot=ROOT/'pilots'/'before'/'initial'
            for mode in ('normal','allocator'):
                name=mode+'-tiny';report=json.loads((pilot/(name+'.json')).read_text())
                path=directory/(name+'.json');shutil.copyfile(pilot/(name+'-catalog.json'),directory/(name+'-catalog.json'))
                path.write_text(json.dumps(report));oracle.validate_report(path,mode,'tiny',source_mode=source_mode,samples=3,warmups=1)
                def case(label, mutate):
                    value=copy.deepcopy(report);mutate(value['results'][0],value)
                    path.write_text(json.dumps(value))
                    rows.append(rejected(lambda:oracle.validate_report(path,mode,'tiny',source_mode=source_mode,samples=3,warmups=1),source_mode+':'+mode+':'+label))
                case('sample_count',lambda r,v:v['configuration'].__setitem__('samples_per_case',2))
                case('output_hash',lambda r,v:r.__setitem__('output_sha256','0'*64))
                case('fixture_part_count',lambda r,v:r['source']['opc_part_add'].__setitem__('output_part_count',66))
                case('typed_gate',lambda r,v:r['source']['opc_part_add']['gates'].__setitem__('stale_source_refusal_verified',False))
                case('chronological_vector',lambda r,v:r['source']['opc_part_add']['lifecycle_ns'].__setitem__(0,1))
                case('sample_alignment',lambda r,v:r['operation_metrics']['sample_indices'].__setitem__(0,99))
                case('source_vector',lambda r,v:r['source'].__setitem__('read_bytes',[1,1,1]))
                case('source_unavailable',lambda r,v:r['operation_metrics']['source'].__setitem__('status','unavailable' if source_mode=='observed' else 'measured'))
                case('codec_invention',lambda r,v:r['operation_metrics']['source']['decompressed_bytes'].__setitem__('status','measured'))
                case('sink_bytes',lambda r,v:r['sink'].__setitem__('accepted_bytes',1))
                case('catalog_digest',lambda r,v:v['corpus_catalog'].__setitem__('catalog_sha256','0'*64))
                if mode=='allocator':
                    case('allocator_balance',lambda r,v:r['operation_metrics']['allocation']['allocated_bytes']['values'].__setitem__(0,0))
                    case('allocator_peak',lambda r,v:r['operation_metrics']['allocation']['region_peak_live_bytes']['values'].__setitem__(0,0))
        for path in (ROOT/'fixtures').glob('tiny*'):shutil.copyfile(path,directory/path.name)
        verify_fixture('tiny',directory,bind=False)
        output=directory/'tiny-output.zip';original=output.read_bytes()
        for label,target,change in [
            ('added_payload','benchmark/added/leaf.bin',lambda raw:b'X'+raw[1:]),
            ('untouched_payload','benchmark/parts/00000.bin',lambda raw:b'X'+raw[1:]),
            ('content_type','[Content_Types].xml',lambda raw:raw.replace(b'application/vnd.litchi.perf.added-part',b'application/octet-stream')),
            ('relationship','_rels/.rels',lambda raw:raw.replace(b'rIdPartAddition',b'rIdWrongAddition')),
        ]:
            rebuilt=io.BytesIO()
            with zipfile.ZipFile(io.BytesIO(original)) as source, zipfile.ZipFile(rebuilt,'w') as sink:
                for info in source.infolist():
                    raw=source.read(info);sink.writestr(info,change(raw) if info.filename==target else raw)
            output.write_bytes(rebuilt.getvalue())
            rows.append(rejected(lambda:verify_fixture('tiny',directory,bind=False),'archive:'+label))
        changed=bytearray(original)
        with zipfile.ZipFile(io.BytesIO(original)) as source:offset=source.getinfo('benchmark/parts/00000.bin').header_offset
        changed[offset+10]^=1;output.write_bytes(changed)
        rows.append(rejected(lambda:verify_fixture('tiny',directory,bind=False),'archive:raw_local_timestamp'))
    value={'change':446,'status':'pass','rejected_mutations':rows,'count':len(rows),'scope':'actual exported ZIP mutations plus report corruption; fixture rebuild mutations can fail any preservation gate, without claiming isolated attribution'}
    with (ROOT/'oracle-probes.json').open('x') as stream:stream.write(json.dumps(value,indent=2)+'\n')
    print(json.dumps({'status':'pass','rejected':len(rows)}))
if __name__=='__main__':main()
