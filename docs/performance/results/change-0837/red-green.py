"""Prove the retained ZIP regression fails on the old adapter, then restore it."""
import driver as d
path=d.ROOT/'crates/soapberry-zip/src/office.rs'
final=path.read_bytes()
argv=['cargo','test','--offline','--locked','-p','soapberry-zip','--test','streaming_interrupted_write']
try:
 path.write_bytes((d.P/'sources/baseline/crates/soapberry-zip/src/office.rs').read_bytes())
 code=d.run('baseline-interruption-regression',argv)
 log=(d.P/'commands/baseline-interruption-regression/output.log').read_text()
 assert code==101 and 'assertion failed: !entry.is_poisoned()' in log
 assert 'test result: FAILED. 0 passed; 1 failed' in log
finally:path.write_bytes(final)
assert d.run('final-interruption-regression',argv)==0
d.write(d.P/'red-green.json',dict(status='pass',baseline='expected runtime assertion failure: retryable Interrupted poisoned bounded entry',final='Store and Deflate payload retries, byte accounting and readback pass',receipts=[d.desc(d.P/f'commands/{n}/receipt.json') for n in ['baseline-interruption-regression','final-interruption-regression']]))
