"""Replay all four independent workload populations after root cleanup."""
import driver as d
import reader

total=0
for manifest,output,expected in [
 ('qualification','qualification-validation',4),
 ('native','native-validation',288),
 ('profiles','profiles-validation',120),
 ('traces','traces-validation',32),
]:
 result=reader.validate_manifest(d.P/(manifest+'.json'))
 assert result==d.read(d.P/(output+'.json')),manifest
 assert result['validated_sample_count']==expected
 total+=expected
assert total==444
print('post-cleanup report replay PASS: 34 reports / 444 measured outputs')
