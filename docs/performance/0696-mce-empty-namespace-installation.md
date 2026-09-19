# 0696 — skip empty MCE namespace installation

Status: retained with scoped declaration-heavy and tail costs.
`performance_claim: none`. Baseline: `442f2dd7f`.

A three-line private guard avoids cloning and immediately releasing the current
namespace scope when an element declares no new namespaces. The measured
real-deck one-edit phase total improves **6.50–6.72%**, from 11.736–11.768 ms
to 10.973–10.976 ms. Real no-op and two-slide edits improve similarly.
Allocation counts, requested bytes and measured live-byte metrics match exactly
across all 78 case/phase comparisons. There is no retained-state or public API
change.

The [evidence packet](results/change-0696/README.md) retains source/probe/build
bindings, the unchanged goal/ADR constraints, raw measurements, focused tests,
assembly review, differential oracles, final validation and cleanup receipts.
No CRUD coverage or registered performance claim is promoted.

## Mechanism and semantic contract

0695 established that presentation-only reuse was a small part of remaining
MCE cost. This experiment targets shared `start` work instead. The old
`Namespaces::with_local` empty branch returns `Ok(self.clone())`; assigning
that result immediately drops the equivalent old namespace owner. A guard
around that call preserves the same head, binding count and inherited pointer
identity. The empty branch cannot return a recoverable error.

All attribute decoding and namespace syntax checks remain before the guard.
Nonempty declarations still follow the existing duplicate, limit and QName
checks. `xmlns=""` remains a real declaration and takes that path. No streaming
codec, namespace representation, cache, global state or parallel policy changes.
The accepted ADR constraints remain unchanged and are reviewed in
[design.md](results/change-0696/design.md).

Independent assembly review confirms a baseline atomic increment/decrement
pair for empty declarations with an inherited namespace head. The candidate
empty branch skips that pair. Root scopes with no head never paid the atomics;
other frame/context clones remain. The compiler now outlines `with_local` on
the nonempty path. The `start` stack reservation remains `0x598` bytes, but the
nonempty path adds the helper frame. Native `.text` grows **260 bytes**; the
relevant `start` plus helper symbol-size sum grows roughly 288 bytes. Neither
static assembly nor these sizes prove a whole-program stack/RSS bound.

## Native end-to-end phase evidence

The matrix has 13 workflows, two baseline A/A legs plus four A/B/B/A legs,
100 samples and five warmups per process: 7,800 measured workflow samples.
The fixed CPU 12 host is warm and shared; no quiescence claim is made.

Timers cover capture, working clone, text edit, commit and apply. Source reading,
initial opening, target selection, output save/reopen and preservation checks
are outside those timers. Thus these are native edit-phase totals, not full
file-open/edit/save latency. The same-length marker control is a mechanism
counterfactual, not a generally valid or equivalent document edit.

| Workflow / input | Baseline medians (ms) | Candidate medians (ms) | Paired change |
| --- | ---: | ---: | ---: |
| one-real | 11.7675 / 11.7361 | 10.9762 / 10.9731 | -6.72% / -6.50% |
| one-control | 5.9984 / 5.9929 | 5.9646 / 5.9878 | -0.56% / -0.09% |
| one-generated | 1.5647 / 1.5720 | 1.5599 / 1.5516 | -0.31% / -1.30% |
| one-notes-poi | 1.0493 / 1.0565 | 1.0502 / 1.0553 | +0.09% / -0.12% |
| one-notes-lo | 1.6647 / 1.6695 | 1.6594 / 1.6597 | -0.32% / -0.59% |
| noop-real | 5.5562 / 5.5249 | 5.1540 / 5.1625 | -7.24% / -6.56% |
| noop-control | 2.9035 / 2.8955 | 2.8832 / 2.8712 | -0.70% / -0.84% |
| noop-generated | 0.7645 / 0.7642 | 0.7525 / 0.7580 | -1.57% / -0.81% |
| noop-notes-poi | 0.5083 / 0.4833 | 0.4761 / 0.4773 | -6.34% / -1.23% |
| noop-notes-lo | 0.6145 / 0.6116 | 0.6094 / 0.6106 | -0.84% / -0.16% |
| two-real | 12.0784 / 12.0813 | 11.2283 / 11.3864 | -7.04% / -5.75% |
| two-control | 6.2238 / 6.1876 | 6.1554 / 6.1525 | -1.10% / -0.57% |
| two-generated | 1.8082 / 1.8169 | 1.7864 / 1.7793 | -1.20% / -2.07% |

| Real one-edit phase | Baseline medians (ms) | Candidate medians (ms) |
| --- | ---: | ---: |
| capture | 5.0779 / 5.0507 | 4.7218 / 4.7221 |
| clone | 0.0167 / 0.0166 | 0.0163 / 0.0166 |
| settext | 1.0797 / 1.0761 | 1.0259 / 1.0179 |
| commit | 5.4891 / 5.4877 | 5.1155 / 5.1223 |
| apply | 0.0928 / 0.0935 | 0.0913 / 0.0915 |

| One-edit allocation diagnostic | Baseline | Candidate |
| --- | ---: | ---: |
| alloc_calls | 143,848 | 143,848 |
| realloc_calls | 4,870 | 4,870 |
| requested_bytes | 12,430,874 | 12,430,874 |
| peak_above_start | 459,170 | 459,170 |
| net_live_change | 195,730 | 195,730 |

The first A/A total medians range from −0.91% to +2.36%. The final baseline
pair is within about 0.70% except the notes-POI no-op at −4.92%; its first
candidate apparent gain must not be treated as a reliable optimization effect.
The real one-edit within-leg bootstrap median intervals are 11.751–11.780 /
11.726–11.751 ms for the paired baselines, versus 10.965–10.983 /
10.965–10.983 ms for the candidates. These intervals characterize the retained
samples, not future host variance or a population tail guarantee.

No total p50/mean/p95/p99 comparison exceeds the +5% review threshold.
Eleven **phase-tail** flags remain in `native-review-triggers.json`: the largest
relative flag is generated no-op clone p99 +61.41% (+7.56 µs), and the largest
absolute flag is notes-POI capture p99 +8.61% (+35.88 µs). Real one-edit clone
p99 rises +28.24% (+5.08 µs). All flags are retained; no universal tail gain is
claimed. Small marker-free shifts are not attributed solely to the guard.

## Allocation, hardware counters and refusal controls

Separate counting-allocator binaries retain three samples per workflow. Every
allocation/reallocation/requested/live metric is identical across the 78
case/phase comparisons; peak-above-start and net-live changes also match.
Requested totals already include full replacement realloc sizes. There is no
allocation-reduction claim for this optimization.

The native open-plus-capture profile has a different denominator from the
edit-phase timers. Startup-subtracted 10/210-iteration slopes show:

| Metric per open/capture | Baseline | Candidate | Change |
| --- | ---: | ---: | ---: |
| Cycles | 28,177,943 | 26,976,295 | −4.26% |
| Instructions | 105,311,206 | 105,241,432 | −0.07% |
| Branches | 21,104,134 | 21,089,819 | −0.07% |
| Branch misses | 108,174 | 106,241 | −1.79% |
| Cache misses | 27,389 | 26,752 | −2.33% |
| Page faults | 194.89 | 179.21 | −8.05% |
| Task clock (ms) | 6.2795 | 6.0122 | −4.26% |

MCE `start` self samples move from 12.11% to 9.28%. Nearly flat instructions
with lower cycles are consistent with removing costly atomics, rather than a
large reduction in instruction count. Single native-child peak RSS is
5,716 → 5,700 KiB; that small diagnostic difference is not a memory saving
or a hard bound. Raw profile reports and binary sizes remain retained.

The ten-case refusal/control probe keeps exact error and package-graph parity
across 6,000 samples. No candidate p50/mean/p95/p99 comparison triggers +5%.
The over-limit control's first baseline is 26.30 ms versus 22.55 ms in its last
baseline; its apparent −14.20% first-pair change is not attributed to this
namespace optimization. All individual values and paired changes remain in
[refusal-tables.md](results/change-0696/refusal-tables.md).

## Shared MCE output and capability controls

The exact oracle compares 192 cases under five profiles: 36 real XML members
from 12 DOCX/XLSX/PPTX archives, 144 deterministic mutations and 12 synthetic
cases. All 1,920 baseline/candidate invocations agree on output SHA-256, length,
borrowed/owned result and full Report, or exact Debug error representation.
This compares existing behavior, including the marker-free borrowed fast path;
it is not an independent complete XML validator. No fresh native Office GUI
interoperability claim follows.

Secondary native controls time MCE processing and output destruction, with
input/capability construction and identity hashing outside. Each matrix uses
AA/ABBA, ten warmups and 300 samples. The two synthetic inputs cover ordinary
and opaque subtrees; the real controls select marker-bearing DOCX numbering,
XLSX worksheet and PPTX slide members. All exact output identities repeat.
These are individual XML kernels, not complete DOCX/XLSX workflows.

| Input | Capability profile | Paired median change |
| --- | --- | ---: |
| synthetic: ordinary | baseline | -3.29% / -7.52% |
| synthetic: ordinary | opaque | -4.31% / -6.53% |
| synthetic: ordinary | opaque-many | -4.63% / -6.04% |
| synthetic: opaque | baseline | -6.08% / -3.25% |
| synthetic: opaque | opaque | -4.94% / -5.65% |
| synthetic: opaque | opaque-many | -5.27% / -6.54% |
| real part: docx | baseline | -5.42% / -4.46% |
| real part: docx | opaque | -2.68% / -3.28% |
| real part: docx | opaque-many | -4.21% / -2.96% |
| real part: xlsx | baseline | -5.64% / -6.91% |
| real part: xlsx | opaque | -3.25% / -4.16% |
| real part: xlsx | opaque-many | -2.40% / -2.35% |
| real part: pptx | baseline | -3.26% / -5.29% |
| real part: pptx | opaque | -5.18% / -3.39% |
| real part: pptx | opaque-many | -3.29% / -2.15% |

There are 10,800 synthetic and 16,200 real-part samples. Synthetic medians
improve 3.25–7.52%; real-part medians improve 2.15–6.91%. No p50/mean/p95/p99
candidate comparison in these two matrices triggers +5%. The actual finite
profile setup and every raw sample remain in the packet.

## Nonempty-path cost and disposition

Assembly showed that the compiler outlined the nonempty helper, so a final
control directly prices that tradeoff. `declared` has 1,000 siblings each
redeclaring its namespace; `mixed` adds an inherited child to each sibling.
Both use the same three capability profiles and 300-sample AA/ABBA design
(10,800 additional samples), after all Cargo/repository checks completed.
Exact output identities agree across every leg.

| Input / capability | Baseline medians (µs) | Candidate medians (µs) | Paired change |
| --- | ---: | ---: | ---: |
| declared / baseline | 312.67 / 313.62 | 324.37 / 324.82 | +3.74% / +3.57% |
| declared / opaque | 360.09 / 361.70 | 369.64 / 371.90 | +2.65% / +2.82% |
| declared / opaque-many | 363.71 / 358.77 | 372.98 / 372.63 | +2.55% / +3.86% |
| mixed / baseline | 566.14 / 569.77 | 555.42 / 559.09 | -1.89% / -1.87% |
| mixed / opaque | 673.24 / 671.76 | 651.81 / 650.40 | -3.18% / -3.18% |
| mixed / opaque-many | 664.11 / 666.56 | 647.60 / 647.61 | -2.49% / -2.84% |

The declaration-heavy medians cost **2.55–3.86%**. This is a real, repeatable
tradeoff, not dismissed as noise because it falls below the review threshold.
Mixed medians improve 1.87–3.18%. No p50/mean/p95/p99 comparison exceeds +5%.
The retained implementation favors the measured common inherited-scope path:
real-deck edit gains are 6.50–6.72%, and all prior real/synthetic capability
controls improve. The declaration-heavy cost, larger code/helper frame and
phase-tail flags remain explicit; no universal parser-speedup claim is made.
Further nonempty-path tuning would require a separate measured candidate.

## Correctness and verification scope

The focused MCE suite passes 88 tests, including three new tests for empty
inherited scopes/default reset, exact namespace limits and QName/directive
refusal order. One initial new-test expectation omitted the implicit `xml`
binding from the finite limit; the corrected ceiling is three for two explicit
declarations. Both the initial failure and final passing log are retained.

All seven integration gates pass. The default suites report 1,211 passes
and two existing ignores; all-feature suites report 4,135 passes and 33 existing
ignores; the facade has 45 passes. These 5,391 passing results include repeated
default/all-feature runs and are not a unique-test count.

The final integration receipts cover formatting, all-feature/all-target checks,
warning-denied library Clippy, default/all-feature consumer tests, the narrow
PPTX facade and warning-denied rustdoc. All six separate repository gates also pass: boundaries, strict/structural
claims, report classification, CRUD coverage and non-iWork verification. The independent audit checks source/binary/corpus and
raw-result bindings before cleanup. Fuzz/toolchain availability is recorded;
no new fuzz or Miri campaign is claimed.
