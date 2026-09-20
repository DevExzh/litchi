#!/usr/bin/env python3
"""Derive program-break transitions from retained traces; no new execution."""
import hashlib,json,re,sys
from pathlib import Path
P=Path(__file__).resolve().parent
CALL=re.compile(r'^([a-z][a-z0-9_]*)\((.*)\)\s+=\s+(.*)$')
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def read(p):return json.loads(p.read_text())
def analyze():
    primary=read(P/'analysis.json');rows=[]
    for row in primary['rows']:
        if row['job']['lane']!='mapping':continue
        name=row['job']['name'];trace=P/(name+'.strace');assert sha(trace)==row['artifacts'][trace.name]
        attrib=row['attribution'];events={e['event_index']:(group,e) for group,data in attrib['groups'].items() for e in data['events']}
        windows=[]
        for kind in ['control','warmup','measured']:
            windows += [(kind,w) for w in attrib['marker_windows'][kind+'_windows']]
        index=0;previous=None;items=[]
        for line in trace.read_text().splitlines():
            match=CALL.fullmatch(line)
            if not match:continue
            syscall,args,result=match.groups();group,event=events[index];assert event['syscall']==syscall and event['result']==result
            if syscall=='brk':
                assert re.fullmatch(r'(NULL|0x[0-9a-f]+)',args) and re.fullmatch(r'0x[0-9a-f]+',result)
                requested=None if args=='NULL' else int(args,16);returned=int(result,16)
                success=requested is None or requested==returned
                delta=None if previous is None else returned-previous
                matches=[(kind,w) for kind,w in windows if w['start_event_index']<=index<=w['end_event_index']];assert len(matches)<=1
                kind,window=matches[0] if matches else ('outside',None)
                functions=[f['resolved'] for f in event['stack_evidence'] if f.get('resolved')]
                items.append(dict(event_index=index,classification=group,window_kind=kind,pair_index=None if window is None else window['pair_index'],minor_faults=None if window is None else window.get('minor_faults'),requested_break=requested,returned_break=returned,previous_returned_break=previous,returned_break_delta_bytes=delta,request_succeeded=success,query=requested is None,deflate_init_stack=any('zlib_rs::deflate::init' in f for f in functions),functions=functions))
                previous=returned
            index+=1
        assert index==len(events)==attrib['totals']['request_count']
        assert all(x['request_succeeded'] for x in items),'retained trace contains failed brk request; do not count as successful growth'
        groups=[]
        for scope in ['control','warmup','measured','outside']:
            for owner in ['phase','setup_publication','other_unresolved']:
                subset=[x for x in items if x['window_kind']==scope and x['classification']==owner]
                if not subset:continue
                deltas=[x['returned_break_delta_bytes'] for x in subset if x['returned_break_delta_bytes'] is not None]
                groups.append(dict(window_kind=scope,classification=owner,count=len(subset),growth_count=sum(x>0 for x in deltas),shrink_count=sum(x<0 for x in deltas),unchanged_count=sum(x==0 for x in deltas),unknown_delta_count=len(subset)-len(deltas),positive_delta_bytes=sum(x for x in deltas if x>0),negative_delta_bytes=sum(x for x in deltas if x<0),event_indices=[x['event_index'] for x in subset]))
        rows.append(dict(name=name,trace_sha256=sha(trace),events=items,groups=groups))
    return dict(scope='Within-trace program-break returns and deltas; not resident pages, faults, operation allocation totals, or cross-process address comparisons.',rows=rows,analysis_sha256=sha(P/'analysis.json'),script_sha256=sha(P/'brk-analysis.py'))
def main():
    value=analyze();dest=P/'brk-analysis.json'
    if sys.argv[1:]==['--check']:assert value==read(dest);print('PASS exact program-break replay')
    else:assert not dest.exists();dest.write_text(json.dumps(value,indent=2)+'\n');print('PASS program-break attribution')
if __name__=='__main__':main()
