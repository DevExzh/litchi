# Ordinary managed DOCX edits (0495)

0495 enables finite-context ordinary DOCX snapshots, direct-body paragraph text
edits, commits, source-bound patches, and sequential publication. Source owners
and reservations survive through semantic views and candidates. An authenticated
complete-artifact publication inverse is also implemented. Managed durable
history, composition, other mutators, and dependency-bearing edits remain open.

This is a necessary measured enabler for the refusal recorded in 0494.
`claim_authorized` remains false: the old managed API supplies no successful
managed timing baseline. The paired comparison uses the same new harness in
unmanaged mode before and after.

## Measurements

[Formal analysis](analysis/formal1.json) and its [verification](verification/formal1.json)
retain 72 processes and 2,160 samples, with 30 samples and three warmups per
process. [Pilot2](verification/pilot2-pilot.json) retains 36 processes and 108
samples. Two reversed formal repeats cover six providers in normal and
allocator builds. The pinned corpus is 200 paragraphs and eight 2 MiB media
members in a 16,793,036-byte DOCX. Every accepted output is 16,793,048 bytes and
passes exact-output, semantic, untouched-media, and source checks.

Normal managed p50 values, in milliseconds:

| Provider | Repeat 1 | Repeat 2 |
| --- | ---: | ---: |
| owned | 5.592 | 5.556 |
| instrumented | 3.277 | 3.298 |
| file-warm | 5.682 | 5.853 |
| short | 4.066 | 4.060 |
| delayed | 572.748 | 576.147 |
| range-zero | 182.327 | 181.990 |

These are scoped baselines, not a provider ranking. The unmanaged owned
allocation sample changes from 22,859 to 9,696 calls and 5,721,334 to 1,623,696
allocated bytes. Eight paired whole-child RSS observations and three latency
comparison cells cross the five-percent review threshold. All flags and repeat
variation remain visible in [results-review.md](results-review.md). Operation
heap observations do not explain whole-child RSS by themselves.

All 720 managed formal rows conserve charged input and accepted sink output,
have zero reservation failures, and release memory, object, and depth gauges to
baseline after drop. `cache_before` is post-open; `cache_live` is
post-edit/pre-publication. Resource `live` is post-publication; `after_drop`
follows package consumption and returned-snapshot/commit release. There is no
cache after-drop observation. The harness patch inverse is XML-only; complete
artifact inverse behavior is covered by production tests, not timed here.

## Validation and profiles

Final gates pass 478 harness tests (one ignored), 937 DOCX unit tests,
119 DOCX integration tests, 388 OPC unit tests, 79 doctests (31 ignored), and
58 Python helper tests. Formatting, all-feature checks, warning-denied Clippy
and rustdoc, and crate boundaries pass. See [final-gates.json](final-gates.json),
[ownership-review.md](ownership-review.md), and
[publication-review.md](publication-review.md).

[Profiles](profiles/profiles1/profile-summary.json) cover both APIs with owned,
warm-file, and short-read inputs. Core counters, syscall traces, and two owned
stack captures are available; LLC events are unsupported. Whole-child profiles
include setup and output verification. They do not isolate operation CPU cost
or establish the cause of RSS flags. [Profile review](profile-review.md) records
the limits and required phase-specific follow-up.

## Reproduction and custody

[environment.json](environment.json) records the AMD EPYC 9R45 host, Rust 1.98.1,
release flags, storage, and tools. [protocol.json](protocol.json) binds all four
retained executables, source manifests, helpers, provider policies, and samples.
Build commands are in `build-{before,after}-{normal,allocator}.json`.
The disposable source checkout and build target were removed after process
and custody checks; [cleanup.json](cleanup.json) and
[owned-worktree-cleanup.json](owned-worktree-cleanup.json) record approximately
5.85 GiB of allocated intermediates removed. The four executables remain under
`/home/zhuhe/.cache/litchi-goal-0495/retained/final1/` for local replay.

[candidate-source-reproduction.json](candidate-source-reproduction.json) and
[candidate-source.patch](candidate-source.patch) retain the 31-file candidate
overlay against the recorded base. The patch passes a reverse-apply check in
the measured checkout. `baseline-tracked.patch` is the before-harness overlay;
`workspace-Cargo.lock` and fixture-input receipts pin dependencies and fixtures.
Source manifests distinguish the candidate from its base revision even though
the disposable checkout has the base HEAD.

From the main repository, verify retained measurements with:

```sh
export LITCHI_MANAGED_EDIT_SOURCE_ROOT=/home/zhuhe/.cache/litchi-goal-0495/source
python3 -B docs/performance/results/change-0495/measure.py verify --attempt pilot2 --pilot
python3 -B docs/performance/results/change-0495/measure.py verify --attempt formal1
python3 -B docs/performance/results/change-0495/verify.py verify
```

For new captures, restore the source checkout and use a fresh attempt name;
never replace an existing capture. Pilot1 is rejected: its child succeeded,
but its verifier equated counters from different phases. Original helpers and
protocol remain in [rejected-pilot1](rejected-pilot1/reason.json). The later
bundle-verifier correction is archived in `rejected-verifier1`. No original
terminal receipt was rewritten as a success.

The full non-iWork performance goal remains open. Genuine borrowed sources,
atomic filesystem save, cold and independent-producer intersections, bounded
parallel scaling, broad CRUD coverage, and several managed edit surfaces still
need implementation and evidence. The protected `~/code/litchi-spec-gaps`
worktree was untouched.
