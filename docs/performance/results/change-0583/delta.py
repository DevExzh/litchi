"""Isolate exactly which verdicts change 0583 moved relative to change 0580 alone.

Pairs the three reports on (input, profile, API, member) and reports, for every
member verdict that differs between the 0580-only build and the 0580+0583
build, what the pre-change (93a610ded) build said.
"""
import collections, json, re, sys
sys.path.insert(0, sys.argv[4])
import classify

before, only0580, fixed = sys.argv[1], sys.argv[2], sys.argv[3]
moved = collections.Counter()
moved_inputs = collections.Counter()
moved_family = collections.Counter()
transitions = collections.Counter()
regressions = []
for (n0, r0, _p0), (n1, r1, _p1), (n2, r2, _p2) in zip(
        classify.parse(before), classify.parse(only0580), classify.parse(fixed)):
    assert n0 == n1 == n2
    name = n0.split(" ")[0]
    family = name.split("/")[0]
    for key in sorted(set(r1) | set(r2)):
        profile, api, member = key
        a, b = r1.get(key, "<absent>"), r2.get(key, "<absent>")
        if a == b:
            continue
        base = r0.get(key, "<absent>")
        ka, kb, kbase = classify.kind_of(a), classify.kind_of(b), classify.kind_of(base)
        transitions[(kbase, ka, kb)] += 1
        moved[api] += 1
        moved_inputs[family] += 1
        moved_family[(family, classify.normalise(classify.error_text(base)) if kbase == "err" else "ACCEPTED")] += 1
        if kb == "ok" and ka == "err":
            regressions.append((name, profile, api, member, a, b))
print(json.dumps({
    "moved_verdicts_total": sum(moved.values()),
    "moved_per_api": {k: v for k, v in sorted(moved.items())},
    "moved_per_family": {k: v for k, v in sorted(moved_inputs.items())},
    "transitions_base_0580_fixed": {f"{k[0]}->{k[1]}->{k[2]}": v for k, v in sorted(transitions.items())},
    "pre_change_reason_of_moved": {f"{k[0]} | {k[1]}": v for k, v in sorted(moved_family.items())},
    "fix_made_something_readable_that_0580_refused": regressions[:20],
    "fix_made_something_readable_count": len(regressions),
}, indent=1))
