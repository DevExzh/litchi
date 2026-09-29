# Independent 0845 results review

## Native and allocation evidence

The retained native evidence is internally consistent. `analysis.json` is
`pass` with 512 reports and 211,328 samples: 384 native reports, 64
qualification reports and 64 one-sample Heaptrack reports. Its 224 summary
rows match every row rendered in `tables.md` (before/after medians, ratios and
10,000-draw intervals). `crosscheck.json` is `pass`, binds the current
`analysis.json` SHA-256, independently recomputes all 224 metrics from the 384
raw native reports and reports no arithmetic mismatch.

The headline values in the draft are correct. At 4,096 Parts, publication p50
ratios are 0.888293 (shared), 0.911912 (root), 0.904003 (deep-shared) and
0.927642 (deep-unique), corresponding to 7.24–11.17% improvement. Lifecycle
p50 ratios span 0.958164–0.976926, corresponding to 2.31–4.18% improvement.
The shared/4,096 publication medians are 3.549550 ms before and 3.156874 ms
after, with ratio 0.888293 and interval 0.884338–0.891285.

The maximum p99 point ratio is 1.007768 (deep-shared/4 publication), and the
maximum RSS point ratio is 1.018926 (root/1,024 publication), so both remain
below the 1.05 review threshold. `regression_flags`, `tail_review`,
`spread_flags` and `allocation_flags` are empty. All 32 allocation rows match
their two before and two after repeats exactly for both allocation events and
requested bytes; every allocation ratio is 1.000000.

The 64 qualification rows are present with both legs for all 32 shape/part/mode
cases. The qualification and reader-v2 receipts retain the independent ZIP
identity, logical digest, CRC, raw-span and target-metadata checks described by
the draft. Main capture receipts, source/build bindings and affected-owner
gates are passing in the packet. The native result remains a warm synthetic
OPC substrate measurement; it does not support semantic Office, mixed-catalog,
cold-cache, durability, scaling or peak-live-memory claims.

## Source scope

`applied.patch` and the formatted worktree diff contain exactly the selected
archive iterator plus the two OPC scan substitutions and one ZIP iterator test.
The direct-ID path keeps physical order, progress checks, checked totals,
relationship classification, reservations, folded-name ambiguity accounting
and first-representative behavior. It adds no retained cache or memo and does
not alter preservation policy. The source scope matches the draft.

## Preliminary Callgrind closure review (before resume)

The separate profile population is not closed in the current packet. The
frozen `profile-plan.json` requires 16 reports, but only four profile run
directories currently exist. `commands/profile-capture/receipt.json` has
`exit_code: 1`; capture stopped on
`profile-runs/01-shared-4096-before/profile.callgrind.1` because the inherited
`reader.py` rejected a valid Callgrind source-file switch. There is no
`capture-profile.json` or `profile-analysis.json` yet. The four retained
profiles are useful diagnostics, and `reader-v2-tests.json` passes its replay
coverage, but they cannot support the draft's final 16-report instruction
evidence or a green profile admission claim.

Before closing the profile section, the resumed run must retain all 16
before/after rows in the frozen order, bind the same binaries and
qualifications, validate all numbered and terminal dumps with the repaired
reader, produce `capture-profile.json` and `profile-analysis.json`, and enforce
the plan limits (one owner, full self-cost conservation, zero terminal-dump
cost, and no native/Heaptrack overlap). Native and profile instruction counts
must remain separate populations.

No build, workload, or production edit was performed for this review. This
file is the only packet file written by the reviewer.

## Root closure addendum after independent protocol review

The preliminary blocker above is resolved. `profile-resume.json` binds all 16
scheduled reports with `reader_v2.py`, and `profile-analysis-v2.json` contains
the fully validated instruction comparison. The preliminary request to create
`capture-profile.json`/`profile-analysis.json` was superseded: those original
success artifacts remain absent to preserve the failed original capture's
meaning. The additive manifest and analyzer carry their actual identities.

The original four raw captures and first three completed receipts are unchanged.
Only the missing fourth receipt was completed from existing successful output;
the remaining twelve commands executed once in the frozen order. Full
`close_v2.py verify` and `clean` passed, including exact reconstruction of the
resume manifest, all raw costs, prefix custody, main metrics and all 587 command
receipts. The independent protocol reviewer separately reported green; see
`protocol-final-review.md`. The native/source findings above remain unchanged.
