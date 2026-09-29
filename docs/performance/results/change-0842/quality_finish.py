import sys
sys.path.insert(0,'/home/zhuhe/code/litchi/docs/performance/results/change-0842')
import driver as d
d.check()
source=d.source()
paths=['Cargo.toml','crates/soapberry-zip/Cargo.toml']
saved={p:(d.ROOT/p).read_bytes() for p in paths}
try:
 for p in paths:(d.ROOT/p).write_bytes((d.P/'source-before'/p).read_bytes())
 code=d.run('baseline-iwork-doc-control',['cargo','doc','--offline','--locked','-p','litchi-numbers-wire','--all-features','--no-deps'])
 assert code==101
finally:
 for p,b in saved.items():(d.ROOT/p).write_bytes(b)
assert d.source()==source
log=(d.P/'commands/baseline-iwork-doc-control/output.log').read_text()
assert 'unresolved link to `read_table_cells`' in log and 'unresolved link to `SidecarReference`' in log
d.write(d.P/'workspace-doc-limitation.json',dict(status='confirmed-preexisting-iwork-doc-failure',error='Unresolved read_table_cells and SidecarReference rustdoc links in litchi-numbers-wire',baseline_receipt=d.desc(d.P/'commands/baseline-iwork-doc-control/receipt.json'),candidate_receipt=d.desc(d.P/'commands/after-final-workspace-doc/receipt.json'),candidate_restored=True,action='Keep excluded iWork unchanged; require warning-denied docs for the five affected crates.'))
receipts=[d.desc(d.P/'commands'/('after-final-'+n)/'receipt.json') for n in ['fmt','workspace-check','owner-check','workspace-clippy']]
owners=['-p','soapberry-zip','-p','litchi-opc','-p','litchi-docx','-p','litchi-xlsx','-p','litchi-pptx']
for name,args in [('owner-doc',['cargo','doc','--offline','--locked',*owners,'--all-features','--no-deps']),('owner-test',['cargo','test','--offline','--locked',*owners,'--all-features']),('boundaries',['python3','-B','tools/check_crate_boundaries.py'])]:
 label='after-final-'+name
 code=d.run(label,args)
 assert code==0
 receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
assert d.source()==source
d.write(d.P/'quality-after.json',dict(status='pass',scope='Workspace library/test checks and library Clippy; five affected ZIP/OPC/OOXML crates all-target checks, docs, and tests; format and boundaries. Full workspace all-target and documentation gates blocked by baseline-confirmed excluded iWork failures.',limitations=[d.desc(d.P/'workspace-gate-limitation.json'),d.desc(d.P/'workspace-doc-limitation.json')],receipts=receipts,source=source))
