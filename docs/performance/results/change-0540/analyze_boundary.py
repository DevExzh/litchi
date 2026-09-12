"""Derive disjoint direct traversal edges without adding overlapping owners."""
import hashlib
import json
from pathlib import Path
H = Path(__file__).resolve().parent
SNAPSHOT = 'litchi_xlsx::cell_values::snapshot::Snapshot::from_source_selected'
VALIDATOR = 'litchi_xlsx::cell_values::validation::validate_xml'
RAW = 'litchi_xlsx::raw::worksheet::parse'
PARSER = 'litchi_xlsx::raw::worksheet::codec::<impl litchi_xlsx::raw::worksheet::model::Parser>::parse'
TRAVERSAL = ['quick_xml::reader::Reader<R>::read_event_impl', 'quick_xml::reader::ns_reader::NsReader<R>::process_event', 'quick_xml::name::NamespaceResolver::resolve_event']

def analyze():
    path = H/'planning-analysis.json'
    data = json.loads(path.read_text())
    rows = []
    for row in data['rows']:
        owners = row['owners']
        total = owners[data['owner']]['inclusive_ir']
        direct = owners[SNAPSHOT]['direct']
        loops = {}
        for name in [VALIDATOR, PARSER]:
            costs = {child:owners[name]['direct'][child] for child in TRAVERSAL}
            count = sum(costs.values())
            loops[name] = dict(direct_edges=costs, sum_ir=count, percent_planning=100*count/total)
        allocated = {name:cost for name,cost in owners[VALIDATOR]['direct'].items() if name in ['__rustc::__rust_alloc','__rustc::__rust_dealloc']}
        rows.append(dict(repeat=row['repeat'],shape=row['shape'],planning_ir=total,
            worksheet_validation_ir=direct[VALIDATOR],worksheet_validation_percent=100*direct[VALIDATOR]/total,
            raw_worksheet_ir=direct[RAW],raw_worksheet_percent=100*direct[RAW]/total,
            traversal=loops,validator_direct_allocator_edges=allocated,
            validator_direct_allocator_percent=100*sum(allocated.values())/total))
    return dict(status='pass',planning_analysis_sha256=hashlib.sha256(path.read_bytes()).hexdigest(),rows=rows,
        scope='Each traversal sum adds three disjoint immediate children of one loop. Validator-wide totals include workbook calls as well as selected worksheets; snapshot direct validation cost isolates selected worksheets. The two loop sums describe measured work, not removable work or a latency estimate. Allocator child Ir is neither an allocation count nor a bound on all ownership cost.')

if __name__ == '__main__':
    result = analyze()
    (H/'boundary-analysis.json').write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
    for row in result['rows']:
        print(row['shape'],row['repeat'],'planning Ir',row['planning_ir'],'validation %',round(row['worksheet_validation_percent'],3),'raw %',round(row['raw_worksheet_percent'],3))
