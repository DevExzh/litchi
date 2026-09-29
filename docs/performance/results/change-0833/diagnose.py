"""Untimed-admission diagnostics; never a retry of the formal matrix."""
import driver as d

P=d.P
ROWS=[('pptx-warm','pptx_file_source_open_selected_slide_lifecycle','warm'),
      ('pptx-cold','pptx_file_source_open_selected_slide_lifecycle','cold-verified'),
      ('opc-save-pair-cold','opc_file_eager_one_part_atomic_save,opc_file_source_one_part_atomic_save','cold-verified')]
if __name__=='__main__':
    assert d.read(P/'qualification.json')['status']=='failed'
    assert not (P/'capture-freeze.json').exists()
    assert d.desc(d.BINARY)==d.read(P/'build.json')['binary']
    d.write(P/'diagnostic-plan.json',dict(rows=ROWS,scope='Failure isolation only; no performance claims',
        script=d.desc(P/'diagnose.py'),qualification=d.desc(P/'qualification.json')))
    results=[]
    for label,case,state in ROWS:
        path=P/f'diagnostic-{label}.json'
        assert not path.exists()
        argv=['taskset','-c','12',str(d.BINARY),'--case',case,'--warmup','0','--samples','1',
              '--filesystem-cache',state,'--filesystem-root',str(d.SCRATCH),'--json',str(path)]
        r=d.run(f'diagnostic-{label}',argv)
        results.append(dict(label=label,exit_code=r['exit_code'],
            receipt=d.desc(P/f'commands/diagnostic-{label}/receipt.json'),
            report=d.desc(path) if path.exists() else None))
    d.write(P/'diagnostics.json',dict(results=results,performance_claim='none'))
