"""Add derived-symbol attribution replay to the frozen closure protocol."""
import driver as d
import attribution
import ancestor_allocation
import runpy
assert attribution.verify()==d.read(d.P/'attribution.json')
assert ancestor_allocation.verify()==d.read(d.P/'ancestor-allocation.json')
runpy.run_path(str(d.P/'close.py'),run_name='__main__')
