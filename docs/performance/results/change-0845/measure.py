"""Execute separate serial populations only after immutable admission."""
import driver as d
assert d.read(d.P/'admission.json')['status']=='pass'
assert d.source()==d.read(d.P/'source-after.json')['source']
for label,argv in [
 ('native-capture',['python3','-B',str(d.P/'capture.py'),'native']),
 ('heap-capture',['python3','-B',str(d.P/'capture.py'),'heap']),
 ('profile-capture',['python3','-B',str(d.P/'profile.py'),'capture']),
 ('main-analysis',['python3','-B',str(d.P/'analyze.py')]),
 ('independent-crosscheck',['python3','-B',str(d.P/'crosscheck.py')]),
 ('render-tables',['python3','-B',str(d.P/'tables.py')]),
 ('profile-analysis',['python3','-B',str(d.P/'profile.py'),'analyze']),
]:
 assert d.run(label,argv)==0,label
print('Measurements and reductions PASS')
