"""Bounded offline admission for the first planned publication capture."""
import importlib.util,json,pathlib,hashlib,sys
P=pathlib.Path(__file__).resolve().parent
spec=importlib.util.spec_from_file_location('capture0901',P/'capture.py');c=importlib.util.module_from_spec(spec);sys.modules[spec.name]=c;spec.loader.exec_module(c)
plan=c.read(P/'plan.json');c.check(plan);folder=P/'capture/00-medium';cmd=c.read(folder/'command.json');c.require(cmd['exit_code']==0,'exit');c.require(cmd['argv']==c.command(plan,plan['schedule'][0],folder),'command')
for d in cmd['artifacts']:c.require(c.desc(d['path'])==d,'artifact '+d['path'])
sem=c.validate_semantics(plan,folder,'medium');c.require(sem==c.read(folder/'semantics.json'),'semantic receipt')
reader=P.parent/'change-0845/reader_v2.py';r=c.load('reader0845',reader)
files=sorted(folder.glob('callgrind.out.*'),key=lambda f:int(f.suffix[1:]));rows=[]
for index,f in enumerate(files,1):
 parsed=r.parse_raw(f);h=parsed['header'];c.require(h['part']==index and h['events']==['Ir'],'part/event');c.require(h['trigger']=='--dump-after='+plan['owner'],'trigger');c.require(h['summary_ir']==h['totals_ir'] and h['summary_ir']>0,'totals');c.require(h['command'].split()==cmd['argv'][cmd['argv'].index(plan['binary']['path']):],'header command')
 edges=[e for fid,name in parsed['names'].items() if name==plan['owner'] for e in r.incoming_edges(parsed,fid)]
 c.require(len(edges)==1,'unique positive owner incoming');edge=edges[0];c.require(edge['calls']==1 and edge['inclusive_ir']==h['summary_ir'],'owner call/cost');c.require(edge['caller'] in [plan['measured_parent'],plan['lifecycle_parent']],'parent')
 owner=r.function_view(parsed['functions'][edge['callee_id']],parsed['names']);c.require(owner['self_ir']+owner['direct_children_ir']==h['summary_ir'],'owner partition');c.require(sum(f['self_ir'] for f in parsed['functions'].values())==h['summary_ir'],'exclusive conservation')
 rows.append(dict(artifact=c.desc(f),header=h,edge=edge))
c.require(sum(row['edge']['caller']==plan['measured_parent'] for row in rows)==1,'single measured call')
terminal=folder/'callgrind.out';h=r._header_0844(terminal.read_text().splitlines(),str(terminal));c.require(h['part']==len(rows)+1 and h['events']==['Ir'] and h['summary_ir']==h['totals_ir']==0 and h['trigger']=='Program termination','terminal')
c.write(P/'first-child-admission.json',dict(status='pass',helper=c.desc(__file__),reader=c.desc(reader),plan=c.desc(P/'plan.json'),command=c.desc(folder/'command.json'),semantics=c.desc(folder/'semantics.json'),rows=rows,terminal=c.desc(terminal),measured_calls=1,lifecycle_calls=len(rows)-1,scope=plan['scope']))
print('first child admission pass')
