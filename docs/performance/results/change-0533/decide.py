"""Bind reviewed 0533 native, allocation, profile and quality admission evidence."""
import argparse
import json
from pathlib import Path
from run import HERE, sha, write


def read(path):return json.loads(path.read_text())


def evaluate():
    comparison=read(HERE/'comparison.json')['comparison']
    admission=comparison['admission']
    primary=admission['primary_workflow_p50']['all_four_cases_both_repeats_pass']
    memory=admission['allocation_guard']['passes']
    profiles={stage:read(HERE/stage/'profile-analysis.json') for stage in ('baseline','candidate')}
    owner=[]
    for repeat in (1,2):
        values={stage:sum(x['constructor']['inclusive_ir'] for x in next(r for r in profiles[stage]['profiles'] if r['group']=='xls-owned' and r['repeat']==repeat)['constructor_attribution']) for stage in profiles}
        owner.append(dict(repeat=repeat,**values,change_percent=100*(values['candidate']/values['baseline']-1)))
    profile=all(r['candidate']<r['baseline'] for r in owner)
    quality=read(HERE/'quality-summary.json')
    assert quality['status']=='pass' and isinstance(quality['checks'], list) \
        and len(quality['checks'])==14
    review=read(HERE/'adverse-review.json')
    assert review['comparison_sha256']==sha(HERE/'comparison.json')
    assert review['complete']
    for label,key in [('matched','matched_adverse_flags_over_five_percent'),('same_build','same_build_variations_over_five_percent')]:
        assert len(review[label])==len(comparison[key])
        for raw,checked in zip(comparison[key],review[label]):
            assert all(checked[k]==v for k,v in raw.items()) and checked['review']
    disposition='accepted' if primary and memory and profile and review['adoption_allowed'] else 'rejected'
    return dict(disposition=disposition,native_primary_gate=primary,memory_gate=memory,profile_gate=profile,
        constructor_ir=owner,adverse_review_complete=True,quality_gates=14,
        other_rejection_reason=review.get('rejection_reason'),
        comparison_sha256=sha(HERE/'comparison.json'),profiles_sha256={stage:sha(HERE/stage/'profile-analysis.json') for stage in profiles},
        quality_summary_sha256=sha(HERE/'quality-summary.json'),adverse_review_sha256=sha(HERE/'adverse-review.json'),
        scope='Fixed synthetic in-memory CFB/XLS comparison; OLE2/OOXML active, ODF deferred, iWork excluded')


if __name__=='__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', nargs='?', type=Path)
    parser.add_argument('--output', dest='output_option', type=Path)
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error('provide output either positionally or with --output')
    result = evaluate()
    destination = args.output_option or args.output or HERE/'decision.json'
    destination.parent.mkdir(parents=True, exist_ok=True)
    write(destination, result)
    print(json.dumps(result))
