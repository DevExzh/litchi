"""Raw no-inline stack census independent of profile graph attribution."""
import gzip,re,sys,collections,hashlib
import custody as c
owner='litchi_pptx::package::model::Package::opened_presentation_with_limits'
rows=[]
for r in c.read(c.P/'perf-fp/frame-receipts.json'):
 z=r['compressed'];path=c.P/z['path'].split('/change-0807/',1)[1];assert path.stat().st_size==z['bytes'] and c.sha(path)==z['sha256']
 raw=gzip.decompress(path.read_bytes());assert len(raw)==r['frames']['bytes'] and hashlib.sha256(raw).hexdigest()==r['frames']['sha256']
 blocks=raw.decode().strip().split('\n\n');qualified=0;unknown=0;leaf=collections.Counter();nested=collections.Counter()
 for block in blocks:
  lines=block.splitlines();assert re.fullmatch(r'\S+\s+\d+\s+\d+\.\d+:\s+\d+ cycles:u:\s*',lines[0])
  names=[]
  for line in lines[1:]:
   m=re.fullmatch(r'\s*[0-9a-f]+ (.+?)(?:\+0x[0-9a-f]+)? \(.+\)',line);assert m,line
   names.append(m[1])
  if owner not in names:continue
  assert names.count(owner)==1;qualified+=1;inside=names[:names.index(owner)]
  if inside:leaf[inside[0]]+=1
  unknown+=int(any('[unknown]' in n for n in inside))
  for name in set(inside):nested[name]+=1
 rows.append({'repeat':r['repeat'],'all_samples':len(blocks),'qualified_samples':qualified,'unknown_interior':unknown,'top_leaf':[list(item) for item in leaf.most_common(20)],'selected_nested':{n:v for n,v in sorted(nested.items()) if any(t in n for t in ['notes::codec::scan_processed_xml','notes::codec::inspect_element','notes::resolved','attributes::IterState::next','CheckedAttributes','from_utf8','load_index_with_slide_root_proofs','package_fingerprint_with_memo'])}})
result={'scope':'Exact-owner no-inline native frame counts. Nested counts overlap, warmup included, no population phase fraction or speedup claim.','rows':rows};dest=c.P/'root-frame-counts.json'
if '--check' in sys.argv:assert c.read(dest)==result;print('Independent native frame census PASS')
else:assert not dest.exists();c.write(dest,result);print(__import__('json').dumps(result,indent=2))
