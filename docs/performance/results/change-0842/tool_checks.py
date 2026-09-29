"""Compile live standalone consumers with their repaired locked dependency graphs."""
import driver as d
d.check();source=d.source();receipts=[]
for name,extra in [('perf-baseline',['--lib']),('perf-execution',[]),('native-resave',['--features','docx,xlsx,pptx'])]:
 label=name+'-locked-check'
 assert d.run(label,['cargo','check','--offline','--locked','--manifest-path',str(d.ROOT/'tools'/name/'Cargo.toml'),*extra])==0
 receipts.append(d.desc(d.P/'commands'/label/'receipt.json'))
assert d.source()==source
d.write(d.P/'tool-checks.json',dict(status='pass',receipts=receipts,source=d.desc(d.P/'source-final.json')))
