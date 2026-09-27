"""Apply only the historically archived candidate, or restore its exact baseline."""
import sys
import custody as c
mode=sys.argv[1];assert mode in ('apply','restore')
binding=c.read(c.P/'candidate-binding.json')
archive=c.P.parent/'change-0787/candidate'
for key in ['candidate_receipt','historical_quality','patch']:
 row=binding[key];assert c.sha(c.P/row['path'])==row['sha256']
oldleg,newleg=('before','after') if mode=='apply' else ('after','before')
old=c.read(c.P.parent/f'change-0787/build-{oldleg}/source.json')['production']['files']
new=c.read(c.P.parent/f'change-0787/build-{newleg}/source.json')['production']['files']
assert c.source()['files']==old
for row in binding['files']:
 name=row['path'];src=archive/('files' if mode=='apply' else 'before')/name
 assert c.sha(src)==new[name]
for row in binding['files']:
 name=row['path'];src=archive/('files' if mode=='apply' else 'before')/name
 (c.ROOT/name).write_bytes(src.read_bytes())
assert c.source()['files']==new
out=c.P/('applied-source.json' if mode=='apply' else 'restored-source.json')
assert not out.exists();c.write(out,c.source())
print(mode,len(new),'exact source files',flush=True)
