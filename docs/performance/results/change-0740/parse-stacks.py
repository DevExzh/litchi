"""Exclusive sampled-period partition; no wall-time conversion or child double count."""
import collections,json,pathlib,re,sys
ROOT='litchi_perf_baseline::run_pptx_cross_copy_lifecycle'
BASE='litchi_pptx::opened::cross_copy_plan::'
PLAN=BASE+'plan_cross_slide_copy_for_slides'
PREP=BASE+'prepare_cross_slide_copy_for_slides'
APPLY=[BASE+'apply_plan','litchi_pptx::package::model::Package::apply_cross_slide_copy_plan']
HEADER=re.compile(r'^\S+\s+\d+/\d+\s+\d+\.\d+:\s+(\d+) cycles:u:\s*$')
FRAME=re.compile(r'^\s+[0-9a-f]+ (.+) \((.+)\)$')
def parse(text):
    assert 'LOST' not in text and 'PERF_RECORD_LOST' not in text
    for block in re.split(r'\n\s*\n',text.strip()):
        lines=block.splitlines(); m=HEADER.fullmatch(lines[0]);assert m,lines[0]
        frames=[]
        for line in lines[1:]:
            f=FRAME.fullmatch(line);assert f,line
            frames.append(f.group(1))
        yield int(m.group(1)),frames

def category(frames):
    if not frames:return 'ambiguous'
    root=ROOT in frames;plan=PLAN in frames;prep=PREP in frames;apply=any(a in frames for a in APPLY)
    if any('[unknown]' in f for f in frames) or len(frames)>=127 or (plan and apply):return 'ambiguous'
    if not root:return 'unrooted_setup_oracle'
    if apply:return 'rooted_apply_excluded'
    if plan and prep:
        if not frames.index(PREP)<frames.index(PLAN)<frames.index(ROOT):return 'ambiguous'
        return 'rooted_plan_prepare'
    if prep:return 'ambiguous'
    return 'rooted_other'

def analyze(text):
    buckets=collections.defaultdict(lambda:{'samples':0,'period':0});inclusive=collections.Counter();leaves=collections.Counter();total=0;count=0;stages=collections.defaultdict(lambda:{'samples':0,'period':0})
    unknown=0;depth=0
    for period,frames in parse(text):
        count+=1;total+=period;c=category(frames);buckets[c]['samples']+=1;buckets[c]['period']+=period
        unknown+=any('[unknown]' in f for f in frames);depth+=len(frames)>=127
        if c=='rooted_plan_prepare':
            candidate=BASE+'build_candidate' in frames
            writer='litchi_opc::pkgwriter::PackageWriter::write_to_stream' in frames
            generated='soapberry_zip::preserve::generated_entry' in frames
            deflate='zlib_rs::deflate::deflate' in frames
            stage=('candidate_writer_generated_deflate' if candidate and writer and generated and deflate else
                   'candidate_writer_other' if candidate and writer else 'candidate_other' if candidate else 'plan_other')
            stages[stage]['samples']+=1;stages[stage]['period']+=period
            leaves[frames[0]]+=period
            for f in set(frames):inclusive[f]+=period
    assert sum(x['period'] for x in buckets.values())==total
    return {'total_samples':count,'total_period':total,'unknown_chain_samples':unknown,'depth_limit_samples':depth,'buckets':dict(buckets),'strict_plan_partition':dict(stages),'strict_plan_leaf_period':dict(leaves.most_common()),'strict_plan_inclusive_symbol_period_overlapping':dict(inclusive.most_common())}
if __name__=='__main__':
    print(json.dumps(analyze(pathlib.Path(sys.argv[1]).read_text()),indent=2))
