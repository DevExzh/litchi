"""Before candidate application, bind exact current artifacts to old goldens."""
import driver as d
import qualify as q
rows=[]
old=d.P.parent/'change-0782';seal=d.read(old/'seal.json')
for item in q.folders('qualification','before'):
 row=item['row'];folder=item['folder'];report=d.read(folder/'report.json')
 name=f"qualification/0-{row['shape']}-{row['mode']}-before.json"
 historical=old/name;assert seal['files'][name]==d.sha(historical)
 expected=d.read(historical)
 for k in ['shape','mode','slides','boxes_per_slide','text_boxes','source','expected_semantic_text_bytes','expected_raw_text_bytes','timing_scope']:
  assert report[k]==expected[k],(name,k)
 assert report['samples_requested']==1 and report['warmup']==0 and len(report['samples'])==1
 sample=report['samples'][0];hs=expected['samples'][0]
 assert sample['verification']==hs['verification'] and sample['output']==hs['output'] and sample['source_sha256']==report['source']['sha256']
 artifact=d.desc(folder/'artifact.ppt');assert sample['output']==report['artifact']=={k:artifact[k] for k in ['bytes','sha256']}
 rows.append(dict(shape=row['shape'],mode=row['mode'],historical=d.desc(historical),current=d.desc(folder/'report.json'),artifact=artifact))
d.write(d.P/'historical-preflight.json',dict(status='pass',rows=rows,scope='Current baseline artifacts match sealed0782 output goldens and every semantic/raw verification field before candidate application. Independent current parser gate remains required before formal measurements.'))
print('Historical fixture preflight PASS',len(rows))
