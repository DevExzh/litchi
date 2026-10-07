#!/usr/bin/env python3
"""Cross-check author diagnostic cell spans against frozen independent v4."""
from __future__ import annotations
import argparse, hashlib, json
from pathlib import Path

V4_SHA = "94f350341ae4bc922c234a0fb7b31d7dfd587ce6d72499e49ecc0d7116d4dbeb"

def sha256(p: Path) -> str:
    h=hashlib.sha256()
    with p.open('rb') as f:
        for b in iter(lambda:f.read(1<<20),b''): h.update(b)
    return h.hexdigest()

def require(c,m):
    if not c: raise AssertionError(m)

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--independent',type=Path,required=True); ap.add_argument('--author',type=Path,required=True); ap.add_argument('--output',type=Path,required=True); a=ap.parse_args()
    independent=json.loads(a.independent.read_text()); author=json.loads(a.author.read_text())
    require(sha256(a.independent)==V4_SHA,'independent v4 receipt hash differs')
    records={(r['case'],r['member']):r for r in author['records']}
    require(len(records)==len(author['records'])==12,'author record keys are not unique 12')
    checked=[]; changed_cells=0; source_bytes=0; replacement_bytes=0; complete=0; windows=0; workbook_metadata=0
    for case in independent['cases']:
        for part in case['parts']:
            key=(case['id'],part['member']); r=records.get(key); require(r is not None,f'missing author record {key}')
            require(r['original']['bytes']==part['original_bytes'] and r['replacement']['bytes']==part['replacement_bytes'],f'size {key}')
            require(r['original']['sha256']==part['original_sha256'] and r['replacement']['sha256']==part['replacement_sha256'],f'hash {key}')
            eligibility=r['cell_window_eligibility']; candidates=eligibility['candidates']
            if part['changed_cell_count']:
                expected={(x['original_start'],x['original_end'],x['replacement_start'],x['replacement_end'],x['original_bytes'],x['replacement_bytes']) for x in part['changed_cells']}
                observed={(x['source']['start'],x['source']['end'],x['replacement']['start'],x['replacement']['end'],x['source']['bytes'],x['replacement']['bytes']) for x in candidates}
                require(expected==observed,f'candidate span set {key}')
                require(part['changed_cell_count']==eligibility['candidate_count']==len(candidates),f'candidate count {key}')
                require(part['changed_cell_original_bytes']==eligibility['candidate_source_bytes_union'],f'source bytes {key}')
                require(part['changed_cell_replacement_bytes']==eligibility['candidate_replacement_bytes_union'],f'replacement bytes {key}')
                require(all(x['lexical_balanced'] and not x['boundary_flags'] for x in candidates),f'candidate safety flags {key}')
                changed_cells += part['changed_cell_count']; source_bytes += part['changed_cell_original_bytes']; replacement_bytes += part['changed_cell_replacement_bytes']
            else:
                require(eligibility['candidate_count']==0,f'non-cell candidate count {key}')
                spans=r['changed_spans']; require(len(spans)==1 and spans[0]['kind']=='insert',f'non-cell author span {key}')
                s=part['non_cell_difference']; span=spans[0]
                require(span['source']['start']==s['common_prefix_bytes'] and span['source']['end']==s['common_prefix_bytes'],f'non-cell source gap {key}')
                require(span['replacement']['start']==s['common_prefix_bytes'] and span['replacement']['end']==s['common_prefix_bytes']+s['replacement_gap_bytes'],f'non-cell replacement gap {key}')
                require(span['source']['bytes']==s['original_gap_bytes'] and span['replacement']['bytes']==s['replacement_gap_bytes'],f'non-cell gap bytes {key}')
                require(span['source_sha256']==s['original_gap_sha256'] and span['replacement_sha256']==s['replacement_gap_sha256'],f'non-cell gap hashes {key}')
                workbook_metadata += 1
            proof=r['audit_proof']['observed']['kind']; require(proof==r['audit_proof']['expected']['kind'],f'proof mismatch {key}')
            if proof=='Complete': complete += 1
            elif proof=='Window': windows += 1
            else: raise AssertionError(f'unexpected proof {proof} {key}')
            checked.append({'case':case['id'],'member':part['member'],'changed_cell_count':part['changed_cell_count'],'proof':proof})
    agg=author['aggregate']
    require(agg['changed_groups']==changed_cells+workbook_metadata,'aggregate changed groups')
    require(agg['cell_candidate_source_bytes']==source_bytes and agg['cell_candidate_replacement_bytes']==replacement_bytes,'aggregate cell bytes')
    require(agg['proof_kinds']=={'Complete':complete,'Error':0,'Identical':0,'Window':windows},'aggregate proof kinds')
    require(agg['parts']==12,'aggregate parts')
    result={'schema':'litchi.xlsx-author-span-crosscheck-0947.v1','status':'pass','scope':'Independent v4 cell spans and per-gap bindings cross-checked against diagnostic-final.json; no workload/build/native measurement.','inputs':{'independent_v4':{'path':str(a.independent),'bytes':a.independent.stat().st_size,'sha256':sha256(a.independent)},'author_diagnostic':{'path':str(a.author),'bytes':a.author.stat().st_size,'sha256':sha256(a.author)}},'checked_records':checked,'totals':{'changed_cell_pairs':changed_cells,'changed_cell_source_bytes':source_bytes,'changed_cell_replacement_bytes':replacement_bytes,'workbook_metadata_differences':workbook_metadata,'complete_parts':complete,'window_parts':windows,'author_aggregate_changed_groups':agg['changed_groups']}}
    require(not a.output.exists(),f'output exists {a.output}'); a.output.write_text(json.dumps(result,indent=2,sort_keys=True)+'\n'); print(json.dumps({'status':'pass','output':str(a.output),'bytes':a.output.stat().st_size,'sha256':sha256(a.output),'totals':result['totals']},sort_keys=True))
if __name__=='__main__': main()
