"""Disjoint source-reviewed instruction attribution; no latency projection."""
from pathlib import Path
import hashlib,json
B=Path(__file__).resolve().parent

def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()

def category(offset):
    if 0x1c0 <= offset < 0x1f8 or 0x1fd <= offset < 0x201:
        return 'visited_bounds_address_test_update'
    if 0x201 <= offset < 0x23c:
        return 'sector_vector_append_capacity'
    if offset == 0x1f8 or 0x23c <= offset < 0x26f or 0x40b <= offset < 0x414:
        return 'fat_lookup_marker_checks_loop_state'
    if 0x6a <= offset < 0x76 or 0x2c5 <= offset < 0x2da:
        return 'success_status_and_shared_return'
    if any(a <= offset < b for a,b in [(0x3f,0x64),(0x7e,0xc8),(0x26f,0x2c5),(0x2da,0x40b),(0x414,0x499)]):
        return 'cold_error_construction_reset'
    return 'entry_reservation_visited_preparation'

def compute():
    p=B/'baseline/instruction-analysis.json';d=json.loads(p.read_text())
    index=json.loads((B/'baseline/assembly-index.json').read_text())
    collector=[r for r in index['rows'] if 'SectorChainScratch13collect_exact' in r['symbol']]
    assert len(collector)==1 and collector[0]['address_hex']=='0000000002f223c0' and collector[0]['size_bytes']==1436
    # This mapping is deliberately bound to reviewed machine code, never reused
    # for another binary merely because its Rust function has the same name.
    assert index['binary_sha256']=='2d1e12c0d1e156ba8dde848d72f4f3c684a73887d5c5fc87ae0f4c57538b9875'
    rows=[]
    for job in d['jobs']:
        timed=[x for x in job['dumps'] if x['role']=='timed'];assert len(timed)==5
        categories={n:0 for n in ['visited_bounds_address_test_update','sector_vector_append_capacity','fat_lookup_marker_checks_loop_state','success_status_and_shared_return','cold_error_construction_reset','entry_reservation_visited_preparation']}
        instructions={};children={};self_ir=0;owner_ir=0
        for dump in timed:
            owner_ir+=dump['owner']['incoming']['inclusive_ir']
            for f in dump['collector']['functions']:
                self_ir+=f['self_ir']
                for i in f['instruction_ir']:
                    c=category(int(i['offset_hex'],16));categories[c]+=i['ir']
                    key=i['address_hex'];r=instructions.setdefault(key,dict(address_hex=key,offset_hex=i['offset_hex'],assembly=i['assembly'],category=c,ir=0));r['ir']+=i['ir']
                for e in f['direct_callees']:children[e['callee']]=children.get(e['callee'],0)+e['inclusive_ir']
        assert sum(categories.values())==self_ir
        rows.append(dict(job=job['name'],constructor_inclusive_ir=owner_ir,collector_self_ir=self_ir,collector_direct_ir=sum(children.values()),collector_inclusive_ir=self_ir+sum(children.values()),categories=categories,direct_children=children,category_percent_of_collector={k:v/self_ir*100 for k,v in categories.items()},category_percent_of_constructor={k:v/owner_ir*100 for k,v in categories.items()},instructions=sorted(instructions.values(),key=lambda x:int(x['address_hex'],16))))
    return dict(status='pass',scope='Four separate five-timed-dump totals; disjoint self-Ir categories. Direct children are separate. Source attribution is not removable-cost or latency prediction.',instruction_analysis_sha256=sha(p),assembly_index_sha256=sha(B/'baseline/assembly-index.json'),mapping_script_sha256=sha(Path(__file__)),rows=rows)
if __name__=='__main__':print(json.dumps(compute(),indent=2,sort_keys=True))
