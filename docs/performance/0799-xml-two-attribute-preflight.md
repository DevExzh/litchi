# 0799 — two-attribute duplicate-check preflight

**Do not advance this candidate to workflow trials.** Both target classes improve,
but six protected error-boundary regressions veto advancement. Production is
unchanged.

This batch tests an iterator candidate selected from the 0798 workload census.
It does not adopt production code. The earlier 0794 and 0797 candidates remain
rejected, and historical measurements are not pooled with this experiment.

The [packet](results/change-0799/README.md) starts at `049e242195`. In the observed
large PPTX capture, one- and two-attribute iterator instances together accounted
for 99.777% of instances, and all observed instances were fully consumed. The new
candidate aims to avoid duplicate-tracking allocation and first-attribute replay
on both of those classes, while preserving early duplicate refusal and the
existing bounded fallback. These counts motivate an experiment; they do not
predict speedup or establish coverage for other workloads.

## Protocol

The exact baseline and candidate helper modules compile into one standalone
binary. All five canonical copies and shared tests are also compiled in isolated
minimal workspaces. Production files are never changed. These checks do not
constitute a fresh full production-crate or public-workflow regression run.

Thirty-nine inputs retain the 0797 cases and add valid duplicate, long quoted and
unterminated duplicate, and three syntax-error tails after two successful items.
Construction exposes the iterator to `black_box` and drops it. Consumption includes
construction, first-error/end iteration, checksums, and destruction. All inputs
are hot and repeatedly reused; this is a direct-helper diagnostic.

Root runs all builds, tests, native captures, and profiles serially. Six native
block pairs alternate order, with 30 samples after three warmups and 4,096
iterations per sample. Process p50 uses sorted sample index 14; the median of
six paired ratios has 10,000 bootstrap resamples, seed 799079, and zero-based
endpoints 250 and 9749. CPU 12 is recorded. Two Callgrind repeats reverse order,
with one sample/iteration and no warmup. Named non-inlined owners delimit Ir,
Bc, Bcm, Bi, and Bim. Guest counters and simulated branch events are not native
cycle fractions or allocator API counts.

The frozen advancement rule requires semantic and custody gates, plus at least
3% improvement with bootstrap upper bound below one for both one- and
two-attribute consumption. Confirmed regressions (ratio above 1.05 and interval
lower bound above one) veto advancement on 18 explicit protected cases: every
input with at most two accepted prefix attributes, and all long duplicate cases.
This includes zero attributes, malformed transitions after two, and long
duplicates after 33. Every other regression remains a reported review trigger.

This differs from 0797's universal consume-regression veto and is frozen before
compilation or measurement. The 0798 census motivates examining both dominant
classes and treating uncommon multi-attribute micro costs as questions for fresh
public-workflow trials, without hiding them or estimating a weighted speedup.
The independent preflight review additionally requires protection of empty and
error boundaries. A direct pass can select a workflow experiment only; public
adoption still requires fresh public-workflow, resource and cross-format gates.

## Candidate and correctness

The iterator adds `First`, `Second(usize)` and `ReplayFirstTwo` phases without
adding a struct field. It reads the first item unchecked because no prior key
exists. On the second call it reconstructs the first and second raw key prefixes,
recognizes the equals sign using quick-xml's exact lexical order, and rejects a
matching key before asking quick-xml to scan the value. Missing equals signs
retain their lexical error. A leading `=` in an unusual key follows quick-xml's
first-byte consumption rule. The existing late-error `name_at` helper is unchanged.

After either successful prefix item, a whitespace-only tail proves exhaustion.
If a third item may follow, a fresh checked iterator consumes the two already
returned items internally to seed duplicate state, then resumes at the third.
The existing 32-name transition and ordered map remain unchanged. This adds
prefix scans and replay costs to longer tags; all measured costs remain visible.

The isolated final helper gates pass 65 baseline and 110 candidate test executions
(13 and 22 tests across five copies), zero ignored or failed, plus Clippy with
warnings denied. The shared tests include malformed key/equals/value precedence,
long duplicate values, empty values, differential parsing, local clone advances
zero through four, fused behavior, and bounded comparison checks. The direct probe
additionally verifies clones at 0, 1, 2, 3, 4, 5, 32 and 33 advances, exact
borrowed keys/values, first errors and positions, checksums, and iterator layout.
Its formatting, locked release build/check/Clippy, catalog, and self-check pass.

Setup corrections are retained: archival patch headers initially named archive
paths; an OLE-common helper copy omitted its shared test-path attribute; Clippy
requested a collapsed duplicate-check condition; and the first probe catalog
accidentally included an extra unquoted-after-two case. The frozen 39-case oracle
rejected that catalog before any capture. The corrected generator matches the
unchanged plan. Failed source mirrors, logs, catalog, and binary identity remain
available; no native case is retried or omitted.

## Results and disposition

All 1,248 capture processes complete successfully, with 28,392 measured samples.
Independent readers verify input bytes, accepted sequences, first errors,
checksums, native ratios and intervals, and all five counters across 624 raw
Callgrind dumps. All 312 owners qualify. Both iterator sizes are 120 bytes in
every report. All 78 independently computed native rows agree with the full
analyzer. No captures are retried or omitted.

Ratios below are candidate/baseline consumption p50. Guest Ir uses one iteration;
both counter repeats are identical for these rows.

| Case | Native ratio | Bootstrap interval | Before Ir | Candidate Ir | Protected veto |
|---|---:|---:|---:|---:|---|

| distinct-0 | 0.628882 | 0.624174–1.050778 | 170 | 183 | no |
| distinct-1 | 0.651018 | 0.627111–0.655196 | 574 | 432 | no |
| distinct-2 | 0.832837 | 0.832013–0.838709 | 910 | 890 | no |
| distinct-3 | 1.552761 | 1.535245–1.584619 | 1,286 | 2,108 | no |
| distinct-4 | 1.431872 | 1.411620–1.487191 | 1,702 | 2,549 | no |
| distinct-5 | 1.246714 | 1.227802–1.258834 | 2,818 | 3,690 | no |
| distinct-32 | 1.060278 | 1.055744–1.071060 | 25,780 | 27,325 | no |
| distinct-33 | 1.027661 | 1.010677–1.033208 | 50,263 | 51,830 | no |
| distinct-64 | 1.022225 | 1.017452–1.031382 | 80,839 | 83,088 | no |
| duplicate-valid-after-1 | 0.636832 | 0.633056–0.642358 | 754 | 594 | no |
| duplicate-valid-after-2 | 1.518736 | 1.494876–1.566236 | 1,090 | 1,887 | yes |
| duplicate-long-quoted-after-1 | 0.641187 | 0.630486–0.652401 | 754 | 594 | no |
| duplicate-long-unterminated-after-1 | 0.649098 | 0.600610–0.656932 | 754 | 594 | no |
| duplicate-long-quoted-after-2 | 1.508976 | 1.488662–1.518779 | 1,090 | 1,887 | yes |
| duplicate-long-unterminated-after-2 | 1.535865 | 1.503638–1.589069 | 1,090 | 1,887 | yes |
| syntax-flag-after-2 | 1.557298 | 1.518724–1.573107 | 994 | 1,791 | yes |
| syntax-unique-tail-after-2 | 1.503305 | 1.498072–1.525634 | 1,104 | 1,901 | yes |
| syntax-equals-value-after-2 | 1.559910 | 1.529931–1.571839 | 1,006 | 1,803 | yes |

One-attribute consumption improves 34.898%; two-attribute consumption improves
16.716%, both with intervals entirely below one. The proposed fast path therefore
clears both dominant-class requirements. Empty-tag results have a wide interval
that crosses one, so its lower median is not a qualified benefit.

However, all six protected failures occur immediately after the two-item prefix:
valid duplicate, long quoted and unterminated duplicate, and all three syntax
tails. Their ratios range from 1.503305 to 1.559910, with lower interval bounds
above one. Nineteen consumption rows trigger the general diagnostic regression
flag, including three distinct attributes at 1.552761 and four at 1.431872. The
frozen rule rejects advancement even though both target classes improve.

The counter evidence is consistent with a shifted replay boundary. Second-item
duplicates return before value scanning; quoted and unterminated 4,096-byte
values after one item have identical candidate Ir. Third-item duplicates re-enter
the checked parser after replaying the two successful items, also preserving
early duplicate refusal but adding work. This explains the code path, not a
native cycle fraction or an allocation API count.

Between-process p50 spread exceeds 5% in 23 of 156 case/mode/leg groups. Every
spread, raw tail and individual interval remains in the packet. All construction
rows avoid the diagnostic regression flag, but opaque construction exposes
materialization that inlined callers may eliminate; it cannot establish a
workflow benefit. Timer, checksum, dispatch, repeated hot-input, compiler layout,
and host effects remain limitations of the direct measurements.

The candidate is archived and production remains unchanged. The next design
needs to avoid rebuilding duplicate state at the first item beyond the short
prefix. Merely moving replay to a later item transfers the cost rather than
removing it. Any replacement still needs exact lexical/error behavior and the
hostile-input bound, then fresh direct and public-workflow/resource/cross-format
evidence. No production optimization, baseline, or CRUD/producer coverage is
promoted by this packet.

After capture and audit, the owned target was removed (180,338,787 logical bytes).
The cleanup witness retains both the final binary identity and the failed catalog
build's identity. Post-cleanup replay passes. All 9,196 production files and 35
architecture inputs remain exact; unrelated files and worktrees are preserved.
Failed setup attempts, final source/gates, every raw capture, independent audits,
and reviews are retained and sealed with the packet.
