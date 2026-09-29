"""Replay the exact 0832 after-source quality witness; run no tests or workloads."""
import sys
import driver as d

P=d.P
PREV=P.parent/'change-0832'

def value():
    current=d.read(P/'prepare.json')['source']
    old=d.read(PREV/'inputs.json')
    quality=d.read(PREV/'quality-after.json')
    assert quality['status']=='pass' and quality['gates']==8 and not quality['reused']
    after=quality['source_map']
    assert len(after)==4
    assert all(v==after.get(k,old.get(k)) for k,v in current.items())
    seal=d.read(PREV/'seal.json')['files']
    bindings={}
    paths=[PREV/'inputs.json',PREV/'quality-after.json']
    for name in ['fmt','check','test','clippy','doc','boundaries','oracle','harness-test']:
        receipt=PREV/f'commands/quality-after-{name}.json'
        started=PREV/f'commands/quality-after-{name}.started.json'
        log=PREV/f'commands/quality-after-{name}.log'
        row=d.read(receipt)
        assert row['exit_code']==0 and row['source_leg']=='after'
        assert row['source_map']==after
        assert row['input_inventory_sha256']==d.sha(PREV/'inputs.json')
        assert row['log_sha256']==d.sha(log)
        assert row['finished_unix']>=row['started_unix']
        paths.extend([receipt,started,log])
    for path in paths:
        key=str(path.relative_to(d.ROOT))
        assert d.sha(path)==seal[key],key
        bindings[key]=d.sha(path)
    return dict(status='pass',reused=True,reused_from=str(PREV.relative_to(d.ROOT)),
      gates=8,normalized_current_inputs=len(current),normalization=after,
      source_identity='All 9388 current build/normative inputs equal sealed 0832 after source',
      evidence=bindings,previous_seal=d.desc(PREV/'seal.json'),
      prepare_sha256=d.sha(P/'prepare.json'),reader_sha256=d.sha(P/'reuse_quality.py'),
      scope='Reused historical quality gates; no fresh test execution in 0833')

if __name__=='__main__':
    v=value()
    if sys.argv[1:]==['--write']: d.write(P/'quality.json',v)
    else:
        assert not sys.argv[1:]
        assert d.read(P/'quality.json')==v
    print('Reused quality witness PASS: 9388 unchanged inputs, 8 sealed gates')
