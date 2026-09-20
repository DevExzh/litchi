# Paired-statistics batch disposition

Accepted for the eight-function scope in `contract.md`. All seven isolated
gates pass: 1,518 tests, zero failures or ignored tests, strict all-target
Clippy, rustdoc with warnings denied, both format checks, crate boundaries,
and source diff checks. Frozen source remained unchanged during the gates.
Independent semantic/numeric and resource/cache reviews pass.

The corpus includes 464 independent observations over 39 fixtures and 42
pinned native observations. Of the native rows, 41 compare against the host
cache; one STEYX row retains a 2,005-ULP host deviation and instead asserts
the independently proved exact mathematical reference. The documented
INTERCEPT profile uses an included constant despite the contradictory
normative cross-reference.

## Performance decision

The first complete capture exposed ten latency flags in existing AVEDEV and
DEVSQ cases, reaching 69.57%. That source, all seven gate receipts, profile
inputs, and all 4,320 samples remain under `attempts/initial-divider`. The
only subsequent production change replaced a per-bit shifted comparison
with an allocation-free limb-wise view; three boundary tests were added.
The final source was frozen and fully gated before the second capture.

Both captures use the same harness, fixtures, lock, baseline commit, three
warmups, and fifteen fresh children per phase. Root independently verified
both sets of raw output/time receipts, source closures, preflight read counts,
binary identities, sample matrices, medians, and cleanup receipts. The final
capture contains 1,170 baseline and 3,150 candidate samples: 78 matched
control groups and 210 candidate groups.

Final matched latency median changes range from -41.70% to +8.41%. Allocation
calls, requested/released bytes, peak live bytes, retained result budgets,
work, and resolver reads are unchanged across all matched groups. RSS
median changes range from -5.55% to +4.69%, with no positive 5% trigger.
No large-reference control has a latency trigger. Three small inline cases
remain explicitly accepted:

| Case | Phase | Baseline ns | Candidate ns | Delta | Bootstrap 95% median-delta interval |
| --- | --- | ---: | ---: | ---: | --- |
| sensitive DEVSQ | evaluate | 3,210 | 3,480 | +8.41% | +5.81% to +11.32% |
| sensitive DEVSQ | parse-evaluate | 3,740 | 4,050 | +8.29% | +5.90% to +9.38% |
| sensitive SKEWP | parse-evaluate | 4,830 | 5,080 | +5.18% | +2.67% to +6.50% |

These are measured slowdowns, not cleared flags or dismissed noise. The
batch accepts the 250–310 ns cost on these tiny fixtures while retaining
fixed memory, corrected direct subnormal rounding, and the broader measured
gains. This is an implementation tradeoff, not proof that the overhead is
necessary for correctness. Further divider optimization remains possible.
The bootstrap uses seed 20260919 and 10,000 independent median resamples;
it describes this capture's uncertainty and does not establish causality or
cross-platform performance. No timing-only retry was used to clear a flag.

`verification.json` and `verification-receipt.json` hold the independent
result and verifier identity. Run `python3 verify.py` from this directory to
check both captures and the current frozen source. The frozen gate lock is
used instead of the different ambient workspace lock. `scratch-cleanup.json`
records removal of owned temporary worktrees, targets, and caches.
