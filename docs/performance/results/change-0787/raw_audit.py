"""Independent corpus reconstruction and raw paired-latency audit (no native code)."""
from pathlib import Path
import hashlib,json,random,statistics,struct,sys,math
P=Path(__file__).resolve().parent
h=lambda b:hashlib.sha256(b).hexdigest()
expected={}
for shape in ('small','large','mixed'):
    manifests=[];seq=hashlib.sha256();total=0
    for i in range(32):
        size=4096 if shape=='small' or (shape=='mixed' and i==31) else 262144
        label=f'litchi-0786-member-{i:02}-'.encode();offset=i*11 if i*11<=255 else 0
        period=bytes((label[k%len(label)]+(k%97)*3+offset)%256 for k in range(len(label)*97))
        payload=(period*((size+len(period)-1)//len(period)))[:size]
        manifests.append(h(payload));seq.update(struct.pack('<Q',i));seq.update(payload);total+=size
    expected[shape]={'members':manifests,'sequence':seq.hexdigest(),'logical_bytes':total}
rows=[];identities={};count=0;measurements={};report_count=0
for lane in ('qualification-before','qualification-after','native','observer'):
    receipts=json.loads((P/lane/'receipts.json').read_text())
    for r in receipts:
        report_count+=1
        path=P/lane/Path(r['report']['path']).name;raw=path.read_bytes()
        assert h(raw)==r['report']['sha256'] and len(raw)==r['report']['bytes'] and r['exit_code']==0
        d=json.loads(raw);shape=r['shape'];e=expected[shape]
        assert [m['sha256'] for m in d['corpus']['members']]==e['members']
        ids=(d['corpus']['opc_sha256'],d['corpus']['cfb_sha256']);assert identities.setdefault(shape,ids)==ids
        for s in d['samples']:
            v=s['verification'];assert v['ordered'] and v['all_member_sha256_match'] and v['members']==32
            assert v['sequence_sha256']==e['sequence'] and v['logical_bytes']==e['logical_bytes']
            a=s['resources'];before=32 if r['state']=='primed' else 0
            assert a['before_operation']['cpu_tasks']==before
            assert a['after_operation']['cpu_tasks']==before+32==a['after_drop']['cpu_tasks']
            assert a['after_drop']['workers']==a['after_drop']['io_concurrency']==0
            for name in ('before_operation','after_operation','after_drop'):
                assert all(0<=a[name][k]<=a['limits'][k] for k in ('workers','io_concurrency','cpu_tasks'))
            m=s['source_metrics']
            if lane!='native' and r['route']!='opc':
                assert m['short_reads']==m['active_reads_after_operation']==0
                assert m['requested_bytes']==m['returned_bytes'] and sum(m['request_size_histogram'])==m['logical_calls']
                assert m['max_simultaneous_reads']<=r['workers']
                if r['route']=='cfb':assert m['logical_calls']==32 and m['returned_bytes']==e['logical_bytes']
                elif r['state']=='primed':assert m['logical_calls']==m['returned_bytes']==0
                else:assert m['logical_calls']==64
            count+=1
        if lane=='native':
            key=(r['route'],shape,r['state'],r['task_floor'],r['workers'],r['block'],r['leg'])
            assert key not in measurements
            times=sorted(s['wall_ns'] for s in d['samples'])
            rss_path=P/lane/Path(r['rss']['path']).name
            assert h(rss_path.read_bytes())==r['rss']['sha256']
            measurements[key]={'p50':times[14],'p95':times[28],'p99':times[29], 'rss':int(rss_path.read_text())}
for case in sorted({k[:5] for k in measurements}):
    metrics={}
    for metric in ('p50','p95','p99','rss'):
        ratios=[measurements[(*case,b,'after')][metric]/measurements[(*case,b,'before')][metric] for b in range(6)]
        rng=random.Random(787078)
        boot=sorted(statistics.median(rng.choice(ratios) for _ in ratios) for _ in range(10000))
        metrics[metric]={'paired_ratios':ratios,'median':statistics.median(ratios),'ci95':[boot[249],boot[9749]]}
    rows.append({'case':list(case),'metrics':metrics})
assert count==22200 and report_count==1080 and len(rows)==60
rejections=[];benefits=[]
for row in rows:
    for metric in ('p50','rss'):
        m=row['metrics'][metric]
        if m['median']>1.05 and m['ci95'][0]>1:rejections.append({'case':row['case'],'metric':metric})
    m=row['metrics']['p50']
    if row['case'][2]=='primed' and row['case'][4]>1 and m['median']<=.97 and m['ci95'][1]<1:benefits.append(row['case'])
result={'reports':report_count,'samples':count,'independently_reconstructed_payloads':96,'corpora':expected,'container_identities':identities,'rows':rows,'rejections':rejections,'benefits':benefits,'candidate_eligible_for_retention':not rejections and bool(benefits)}
encoded=json.dumps(result,indent=2,sort_keys=True)+'\n';out=P/'raw-audit.json'
if '--check' in sys.argv:assert out.read_text()==encoded
else:
    assert not out.exists()
    out.write_text(encoded)
print('independent audit PASS: 1080 reports, 22200 samples, 96 reconstructed payloads, 60 paired cases')
