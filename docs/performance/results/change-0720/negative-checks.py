#!/usr/bin/env python3
"""Ensure body trace invariants reject plausible corrupted observations."""
import copy,json
from unittest.mock import patch
import analyze
from custody import P

original=analyze.traces

def missing_mce(rows):rows['generated-medium'][0]['active_calls'].pop(0)
def wrong_offset(rows):rows['generated-medium'][0]['active_calls'][1]['input_offsets'][0]+=1
def unequal_events(rows):rows['numbered-list'][0]['range_passes'][0]['event_count']-=1
def wrong_source(rows):rows['numbered-list'][0]['xml_sha256']='0'*64

results=[]
for mutate in [missing_mce,wrong_offset,unequal_events,wrong_source]:
 def corrupted(name):
  rows=copy.deepcopy(original(name));mutate(rows);return rows
 with patch.object(analyze,'traces',corrupted):
  try:analyze.analyze()
  except AssertionError:results.append({'mutation':mutate.__name__,'rejected':True})
  else:raise AssertionError('corruption accepted: '+mutate.__name__)
assert analyze.analyze()==analyze.read('analysis.json')
(P/'negative-checks.json').write_text(json.dumps(results,indent=2)+'\n')
print('PASS: four corrupted observations rejected; original evidence still passes')
