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
rows=[];identities={};count=0;p50={}
for lane in ('qualification','native','observer'):
    receipts=json.loads((P/lane/'receipts.json').read_text())
    for r in receipts:
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
            key=(r['route'],shape,r['state'],r['task_floor'],r['workers'],r['block'])
            p50[key]=sorted(s['wall_ns'] for s in d['samples'])[14]
for family in sorted({k[:4] for k in p50}):
    for w in (1,2,4,8,32):
        ratios=[p50[(*family,1,b)]/p50[(*family,w,b)] for b in range(6)]
        rng=random.Random(786078);boot=sorted(statistics.median(rng.choice(ratios) for _ in ratios) for _ in range(10000))
        rows.append({'case':list(family)+[w],'paired_speedups':ratios,'median':statistics.median(ratios),'ci95':[boot[249],boot[9749]]})
assert count==22200 and len(rows)==120
result={'reports':1080,'samples':count,'independently_reconstructed_payloads':96,'corpora':expected,'container_identities':identities,'rows':rows}
encoded=json.dumps(result,indent=2,sort_keys=True)+'\n';out=P/'raw-audit.json'
if '--check' in sys.argv:assert out.read_text()==encoded
else:out.write_text(encoded)
print('independent audit PASS: 1080 reports, 22200 samples, 96 reconstructed payloads, 120 paired curves')
