#!/usr/bin/env python3
import copy, datetime as dt, hashlib, importlib.util, json, math, random, statistics
from pathlib import Path

ROOT = Path('/home/zhuhe/litchi-performance-0943')
PACKET = ROOT / 'formal-v1'
MANIFEST_PATH = ROOT / 'formal-authorized-v1.json'
REDUCTION_PATH = ROOT / 'reductions/formal-v1.json'
RUNS = ROOT / 'runs/formal-v1'
OUT = ROOT / 'audit-formal-v1'
MAX = 1 << 20
ENV = {'PATH':'/usr/bin:/bin','LC_ALL':'C','LANG':'C','TZ':'UTC','RUST_BACKTRACE':'0'}
VERSIONS = ('baseline','candidate')
FIXTURES = ('docnohf','docfloat')
EDITS = (1,2,8,32)
ROUTES = ('zero','default','release')
NATIVE_CPU = 14
ALLOC_CPU = 14
DEADLINE = 900
METRICS = {'allocated_bytes','deallocated_bytes','allocation_calls','peak_live_bytes','retained_bytes'}


def fail(msg):
    if isinstance(msg, bool):
        if not msg: return
        raise AssertionError('condition failed')
    raise AssertionError(msg)
def sha(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1<<20),b''): h.update(b)
    return h.hexdigest()
def desc(path, allow_large=True):
    path=Path(path).resolve(); fail(f'missing {path}') if not path.is_file() or path.is_symlink() else None
    size=path.stat().st_size
    if not allow_large: fail(f'large artifact {path}') if size>=MAX else None
    return {'path':str(path),'bytes':size,'sha256':sha(path)}
def strict(path):
    def hook(pairs):
        d={}
        for k,v in pairs:
            fail(f'duplicate key {k} in {path}') if k in d else None
            d[k]=v
        return d
    return json.loads(Path(path).read_text(encoding='utf8'), object_pairs_hook=hook)
def rel_desc(value, packet, allow_large=True):
    fail('descriptor not object') if not isinstance(value,dict) else None
    raw=value.get('path'); fail(f'bad packet path {raw!r}') if not isinstance(raw,str) or not raw or Path(raw).is_absolute() or '..' in Path(raw).parts else None
    path=(packet/Path(raw)).resolve()
    fail(f'packet escape {raw}') if not path.is_relative_to(packet.resolve()) else None
    actual=desc(path,allow_large)
    fail(f'descriptor mismatch {raw}') if actual['bytes']!=value.get('bytes') or actual['sha256']!=value.get('sha256') else None
    actual['path']=Path(raw).as_posix()
    fail(f'descriptor fields mismatch {raw}') if actual != value else None
    return path,actual

def iso(v):
    x=dt.datetime.fromisoformat(v.replace('Z','+00:00'))
    return x.astimezone(dt.timezone.utc)

def fixture_path(source, case):
    row=next(x for x in protocol['fixtures'] if x['case']==case)
    p=(Path(source)/row['relative_path']).resolve()
    fail(f'fixture outside source {p}') if not p.is_relative_to(Path(source).resolve()) else None
    fail(f'fixture identity {case}') if p.stat().st_size!=row['bytes'] or sha(p)!=row['sha256'] else None
    return p,row

def native_rows():
    out=[]
    for rnd in range(1,13):
        cases=list(FIXTURES) if rnd%2 else list(reversed(FIXTURES))
        edits=list(EDITS) if rnd%2 else list(reversed(EDITS))
        versions=list(VERSIONS) if rnd%2 else list(reversed(VERSIONS))
        for case in cases:
            for e in edits:
                pair=f'native-r{rnd:02d}-{case}-e{e:02d}'
                for ver in versions:
                    out.append({'lane':'native','round':rnd,'case':case,'edits':e,'retention':'default','version':ver,'pair_id':pair,'label':f'{pair}-{ver}','warmups':5,'samples':50,'cpu':NATIVE_CPU,'deadline_seconds':DEADLINE})
    return out

def alloc_rows():
    cells=[(c,e,r) for c in FIXTURES for e in EDITS for r in ROUTES]
    out=[]
    for rep in (1,2):
        ordered=cells if rep==1 else list(reversed(cells)); vers=list(VERSIONS) if rep==1 else list(reversed(VERSIONS))
        for c,e,r in ordered:
            pair=f'allocator-r{rep:02d}-{c}-e{e:02d}-{r}'
            for ver in vers:
                out.append({'lane':'allocator','repetition':rep,'case':c,'edits':e,'retention':r,'version':ver,'pair_id':pair,'label':f'{pair}-{ver}','warmups':0,'samples':1,'cpu':ALLOC_CPU,'deadline_seconds':DEADLINE})
    return out

def schedule_hash(rows): return hashlib.sha256(json.dumps(rows,sort_keys=True,separators=(',',':')).encode()).hexdigest()

def expected_command(spec, manifest, origin):
    ver=spec['version']; lane=spec['lane']; case=spec['case']
    binary=manifest['builds'][ver][lane]['binary']['path']
    source=manifest['sources'][ver]['source_root']
    fp, _ = fixture_path(source,case)
    rss=(Path(origin)/'children'/spec['label']/'max-rss-kib.txt').resolve()
    cpu=NATIVE_CPU if lane=='native' else ALLOC_CPU
    return ['/usr/bin/time','-f','%M','-o',str(rss),'/usr/bin/taskset','--cpu-list',str(cpu),binary,'--case',case,'--input',str(fp),'--edits',str(spec['edits']),'--retention',spec['retention'],'--warmup',str(spec['warmups']),'--samples',str(spec['samples'])]

def nearest(vals, frac):
    ordered=sorted(vals); return ordered[min(max(math.ceil(frac*len(ordered))-1,0),len(ordered)-1)]
def stats(vals): return {k:nearest(vals,f) for k,f in [('p50',.5),('p95',.95),('p99',.99)]}|{'mean':statistics.fmean(vals)}
def inc(before,after,limit):
    value=float('inf') if before==0 and after>0 else ((after/before)-1 if before else 0)
    return {'before':before,'after':after,'increase_fraction':value,'flag':value>limit}
def boot(ratios):
    rng=random.Random(943); meds=[statistics.median(ratios[rng.randrange(len(ratios))] for _ in ratios) for _ in range(10000)]
    return {'draws':10000,'seed':943,'lower':nearest(meds,.025),'upper':nearest(meds,.975),'median':statistics.median(ratios),'ratios':ratios}
def strip_metrics(node):
    if isinstance(node,dict):
        for k in list(node):
            if k in METRICS: node.pop(k,None)
            else: strip_metrics(node[k])
    elif isinstance(node,list):
        for x in node: strip_metrics(x)
def capacity(report):
    routes={}
    for route in report.get('route_outputs',[]):
        routes[route['retention']]={'output_capacity':route['output_capacity'],'final_retained_before_commit':route['final_retained_before_commit'],'stages':[{k:s[k] for k in ('before','after','after_release')} for s in route['stages']]}
    samples=[]
    for sample in report.get('samples',[]):
        a=sample.get('allocation')
        samples.append({'output_capacity':sample['output_capacity'],'retained_bytes':a.get('retained_bytes') if isinstance(a,dict) else None,'stages':[{k:s[k] for k in ('before','after','after_release')} for s in sample.get('stages',sample.get('stage_observations',[]))]})
    a=report.get('allocation')
    return {'output_capacity':report['output_capacity'],'retained_bytes':a.get('retained_bytes') if isinstance(a,dict) else None,'routes':routes,'samples':samples}
def semantic(report):
    v=copy.deepcopy(report); v.pop('input',None); v.pop('elapsed_ns',None)
    for s in v.get('samples',[]):
        if isinstance(s,dict): s.pop('elapsed_ns',None)
    strip_metrics(v)
    cap=capacity(report)
    def strip_cap(n):
        if isinstance(n,dict):
            n.pop('output_capacity',None); n.pop('final_retained_before_commit',None)
            for x in n.values(): strip_cap(x)
        elif isinstance(n,list):
            for x in n: strip_cap(x)
    strip_cap(v); v['__capacity_shape']={'route_count':len(cap['routes']),'sample_count':len(cap['samples'])}
    return v

def report_pair_equal(a,b): return semantic(a)==semantic(b) and capacity(a)==capacity(b)

manifest=strict(MANIFEST_PATH); protocol=strict(ROOT/'tooling/comparison-protocol.json')
source_rev={v:(lambda q:q.get('source_revision',q.get('revision',q.get('head'))))(strict(Path(manifest['sources'][v]['source_receipt']['path']))) for v in VERSIONS}
rows=native_rows()+alloc_rows(); fail(len(rows)!=288)
expected_labels={r['label']:r for r in rows}
fail(schedule_hash(native_rows()) != manifest['schedule_sha256']['native'])
fail(schedule_hash(alloc_rows()) != manifest['schedule_sha256']['allocator'])
fail(schedule_hash(rows) != manifest['schedule_sha256']['combined'])
# Outer terminal custody.
outer_start=strict(RUNS/'start.json'); outer_term=strict(RUNS/'terminal.json')
fail(outer_term.get('exit_code')!=0 or outer_term.get('argv')!=outer_start.get('argv') or outer_term.get('terminal_written_before_formal_reduction') is not True)
# Child receipts and terminals.
receipt_dir=PACKET/'receipts'; receipt_files=sorted(receipt_dir.glob('*.json')); fail(len(receipt_files)!=288)
loaded={}; validator_path=Path('/home/zhuhe/code/litchi/docs/performance/results/change-0917/capture.py')
vs=importlib.util.spec_from_file_location('v0917',validator_path); vm=importlib.util.module_from_spec(vs); vs.loader.exec_module(vm)
validator_pass=0; all_clean=True
previous_finish=None
for index,spec in enumerate(rows):
    rp=receipt_dir/f"{spec['label']}.json"; fail(f'missing receipt {rp}') if not rp.is_file() else None
    rec=strict(rp); fail(f'receipt status/spec {spec["label"]}') if rec.get('status')!='passed' or rec.get('spec')!=spec else None
    fail(rec.get('schedule_sha256')!=schedule_hash(rows))
    fail(rec.get('protocol') != {'path':str((ROOT/'tooling/comparison-protocol.json').resolve()),'bytes':(ROOT/'tooling/comparison-protocol.json').stat().st_size,'sha256':sha(ROOT/'tooling/comparison-protocol.json')})
    # Manifest descriptor is exact and current.
    fail(rec.get('manifest') != desc(MANIFEST_PATH,False))
    st=iso(rec['started_utc']); fi=iso(rec['finished_utc']); fail(st>fi or previous_finish is not None and fi < previous_finish); previous_finish=fi
    term_path,_=rel_desc(rec['terminal'],PACKET,False); term=strict(term_path)
    fail(term['schema']!='litchi.performance.0943.child-terminal.v1' or term['status']!='terminal' or term['spec']!=spec or term['label']!=spec['label'])
    fail(term.get('exit_code') != 0 or rec.get('exit_code') != 0 or term.get('deadline_seconds')!=DEADLINE or term.get('deadline_exceeded') is not False or term.get('runner_error') is not None or term.get('process_group_closed') is not True or term.get('process_group_forced_cleanup') is not False)
    fail(term.get('started_utc')!=rec.get('started_utc') or term.get('ended_utc')!=rec.get('finished_utc') or term.get('environment')!=ENV or rec.get('environment')!=ENV)
    ver=spec['version']; lane=spec['lane']; source=manifest['sources'][ver]; build=manifest['builds'][ver][lane]
    expected_source={'root':source['source_root'],'revision':source_rev[ver]}; fail(rec.get('source')!=expected_source or term.get('source_revision')!=source_rev[ver])
    fail(rec.get('binary')!=build['binary'] or term.get('binary')!=build['binary'])
    origin=Path(rec['capture_origin']).resolve(); command=rec['command']; fail(command!=expected_command(spec,manifest,origin))
    packet_command=[]
    origin_s=str(origin)
    for token in command:
        packet_command.append('<packet>' if token==origin_s else '<packet>/'+token[len(origin_s)+1:] if token.startswith(origin_s+'/') else token)
    fail(rec.get('packet_command')!=packet_command or term.get('argv')!=command or term.get('packet_command')!=packet_command or term.get('capture_origin')!=str(origin) or term.get('cwd')!=rec.get('cwd') or term.get('cwd')!=source['source_root'])
    for field in ('terminal','validation','report','stderr','max_rss_kib'):
        path,d=rel_desc(rec[field],PACKET, field not in ('terminal','validation','max_rss_kib'))
        if field in ('report','stderr','max_rss_kib'):
            td=term.get(field); fail(td is None or td!=rec[field])
    rss_path,_=rel_desc(rec['max_rss_kib'],PACKET,False); rss_text=rss_path.read_text().strip(); fail(not rss_text.isdigit() or int(rss_text)!=rec.get('rss_kib'))
    val_path,_=rel_desc(rec['validation'],PACKET,False); val=strict(val_path); fail(val.get('status')!='passed')
    report_path,report_d=rel_desc(rec['report'],PACKET,True); report=strict(report_path)
    fixture_row=next(x for x in protocol['fixtures'] if x['case']==spec['case']); fixture={'case':spec['case'],'relative_path':fixture_row['relative_path'],'bytes':fixture_row['bytes'],'sha256':fixture_row['sha256'],'path':fixture_row['relative_path'],'path_abs':str((Path(source['source_root'])/fixture_row['relative_path']).resolve())}
    vm.validate_report(report,dict(spec),fixture,spec['lane'],expected_warmups=spec['warmups'],expected_samples=spec['samples']); validator_pass+=1
    fail(val.get('report')!=report_d or val.get('report_summary') is None)
    loaded[spec['label']]={'spec':spec,'receipt':rec,'term':term,'report':report,'rss':int(rss_text)}
# Pair and metric checks.
native_cells=[]; alloc_cells=[]; sem_equal=cap_equal=True; native_flags={'p50':[],'p95':[],'p99':[],'rss':[]}; alloc_flags={'allocation':[],'peak_live':[],'rss':[]}; failed_cells=[]
for c in FIXTURES:
  for e in EDITS:
    rounds=[]; ratios=[]
    for rnd in range(1,13):
      pfx=f'native-r{rnd:02d}-{c}-e{e:02d}'
      a=loaded[pfx+'-baseline']; b=loaded[pfx+'-candidate']; sem_equal &= semantic(a['report'])==semantic(b['report']); cap_equal &= capacity(a['report'])==capacity(b['report'])
      av=[float(x) for x in a['report']['elapsed_ns']]; bv=[float(x) for x in b['report']['elapsed_ns']]; sa=stats(av); sb=stats(bv); ratio=sa['p50']/sb['p50']; ratios.append(ratio)
      flags={k:inc(sa[k],sb[k],.05) for k in ('p50','p95','p99')}; rf=inc(float(a['rss']),float(b['rss']),.05)
      for k,v in flags.items():
        if v['flag']: native_flags[k].append({'label':pfx,'version_pair':'baseline/candidate','stats':v})
      if rf['flag']: native_flags['rss'].append({'label':pfx,'version_pair':'baseline/candidate','stats':rf})
      rounds.append({'label':pfx,'p50_ratio':ratio,'latency':flags,'rss':rf})
    bt=boot(ratios); gate=bt['lower'] >= 1/1.05
    if not gate: failed_cells.append({'case':c,'edits':e,'lower':bt['lower'],'upper':bt['upper'],'median':bt['median']})
    native_cells.append({'case':c,'edits':e,'bootstrap':bt,'latency_gate':gate,'rounds':rounds})
for rep in (1,2):
 for c in FIXTURES:
  for e in EDITS:
   for route in ROUTES:
    pfx=f'allocator-r{rep:02d}-{c}-e{e:02d}-{route}'; a=loaded[pfx+'-baseline']; b=loaded[pfx+'-candidate']; sem_equal &= semantic(a['report'])==semantic(b['report']); cap_equal &= capacity(a['report'])==capacity(b['report'])
    aa=a['report']['allocation']; bb=b['report']['allocation']; af=inc(float(aa['allocated_bytes']),float(bb['allocated_bytes']),.01); pf=inc(float(aa['peak_live_bytes']),float(bb['peak_live_bytes']),.05); rf=inc(float(a['rss']),float(b['rss']),.05)
    row={'label':pfx,'allocation':af,'peak_live':pf,'rss':rf}
    for key,flag in [('allocation',af),('peak_live',pf),('rss',rf)]:
      if flag['flag']: alloc_flags[key].append(row)
    alloc_cells.append(row)
target=[]
for c in FIXTURES:
  vals=[loaded[f'allocator-r{rep:02d}-{c}-e32-default-{v}']['report']['allocation']['allocated_bytes'] for rep in (1,2) for v in ('baseline','candidate')]
  base=statistics.median([loaded[f'allocator-r{rep:02d}-{c}-e32-default-baseline']['report']['allocation']['allocated_bytes'] for rep in (1,2)])
  cand=statistics.median([loaded[f'allocator-r{rep:02d}-{c}-e32-default-candidate']['report']['allocation']['allocated_bytes'] for rep in (1,2)])
  target.append({'case':c,'baseline_median':base,'candidate_median':cand,'ratio':cand/base,'saves_at_least_5_percent':cand<=base*.95})
result={'schema':'litchi.performance.0943.independent-formal-audit.v1','status':'passed','manifest':desc(MANIFEST_PATH,False),'runner':desc(ROOT/'tooling/comparison-runner.py',False),'protocol':desc(ROOT/'tooling/comparison-protocol.json',False),'outer_terminal':desc(RUNS/'terminal.json',False),'packet_root':str(PACKET),'receipt_count':len(receipt_files),'schedule_counts':{'native':len(native_rows()),'allocator':len(alloc_rows()),'combined':len(rows)},'validator_passes':validator_pass,'terminal_clean_count':len(loaded),'pair_counts':{'native':96,'allocator':48},'semantic_equal_pairs':sem_equal,'capacity_equal_pairs':cap_equal,'native':native_cells,'allocator':alloc_cells,'native_failed_bootstrap_cells':failed_cells,'native_flag_counts':{k:len(v) for k,v in native_flags.items()},'native_flags':native_flags,'allocator_flag_counts':{k:len(v) for k,v in alloc_flags.items()},'allocator_flags':alloc_flags,'admission_targets':target,'gates':{'all_8_native_bootstrap_lower_bounds':not failed_cells,'candidate_saves_at_least_5_percent_in_one_32_edit_default_fixture':any(x['saves_at_least_5_percent'] for x in target),'every_allocator_pair_gross_requested_within_1_percent':not alloc_flags['allocation'],'every_allocator_pair_peak_live_within_5_percent':not alloc_flags['peak_live'],'all_explicit_capacity_fields_equal':sem_equal and cap_equal},'adoption':False,'reduction_sha256':sha(REDUCTION_PATH),'independent_reduction_sha256':sha(ROOT/'independent-reduction-audit-20261007.json'),'regeneration_command':'python3 -B tooling/comparison-runner.py --reduce --manifest /home/zhuhe/litchi-performance-0943/formal-authorized-v1.json --lane all --output /home/zhuhe/litchi-performance-0943/independent-reduction-audit-20261007.json'}
(OUT/'independent-audit-v1.json').write_text(json.dumps(result,indent=2,sort_keys=True)+'\n')
print(json.dumps({'status':'passed','output':str(OUT/'independent-audit-v1.json'),'bytes':(OUT/'independent-audit-v1.json').stat().st_size,'sha256':sha(OUT/'independent-audit-v1.json'),'receipt_count':len(receipt_files),'validator_passes':validator_pass,'failed_cells':failed_cells,'native_flag_counts':{k:len(v) for k,v in native_flags.items()},'allocator_flag_counts':{k:len(v) for k,v in alloc_flags.items()},'targets':target,'gates':result['gates']},sort_keys=True))
