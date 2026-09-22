# 0732 — native public PPT phase attribution

This batch adds opt-in PPT commit diagnostics and measures the existing public
slide-removal workflow on `45543.ppt`. It makes no runtime optimization or
speedup claim. The ordinary commit body remains byte-identical to 0731.

## Why native attribution

0731's Callgrind profile used software SHA while Valgrind masked native SHA
acceleration. Its 67.78% artifact-hash instruction fraction could not rank
native latency. The new `performance-diagnostics` feature emits synchronous,
content-free phase events from `Transaction::commit_profiled`, without library
clocks or retained events. The probe timestamps events externally.

The diagnostic copy preserves the ordinary sequence, local allocation
lifetimes, partial moves, inline digest evaluation, validations, typed errors,
and source-checked reversible patch. In particular, the document-commit branch
stays outside the observer closure and the live-record buffer remains in its
outer scope. `StructuralNoOp` names an empty document patch; formatting-only
transactions can still change the complete artifact on that branch.

## Frozen experiment

The fixed matrix uses Rust 1.95.0 and one release binary with diagnostics
enabled on CPU 12 of an AMD EPYC 9R45 host:
three cycles × three rounds × four routes, each with 50 samples and three
warmups. All 1,800 measured owners are retained; 108 additional warmups execute before
measurement. The routes are
ordinary opaque, ordinary split, profiled empty, and profiled clock. Each round
compares the same three adjacent route pairs despite rotated/reversed execution
order. Absolute p50 or mean differences above 5% are interpretation flags;
there are no selective reruns or trimmed tails.

Whole timing includes public open, edit, second-slide removal, commit, output
copy, and destruction of the local transaction/commit/snapshot owners. Returned
output remains alive for the untimed preservation oracle. Clock traces use a
fixed-capacity stack recorder. Its separate 20-event calibration is reported
without subtracting it from measured phases. Individual phase fractions use
the same owner's whole and commit durations. Unobserved commit work remains
an explicit residual.

Every measured output must match the sealed 0728 digest, raw-directory/stream inventory,
length-changing replacement proof, surviving slide payloads and logical text.
The same eight negative oracle controls must reject their corruptions. The
four-route qualification and a synthetic analysis preflight precede capture;
the synthetic data validate schema only and are never timing evidence.

## Results and observer controls

Ranges below span the nine process statistics per route, in microseconds.
They show repeat spread, not confidence intervals. With 50 samples per process,
the nearest-rank p99 is the maximum; every raw tail remains retained.

| Route | p50 | Mean | p95 | p99 / maximum |
| --- | ---: | ---: | ---: | ---: |
| Ordinary opaque | 978.53–997.45 | 1005.76–1014.01 | 1121.10–1141.56 | 1154.93–1177.53 |
| Ordinary split | 998.46–1011.51 | 1010.49–1017.14 | 1126.05–1144.92 | 1157.33–1198.26 |
| Profiled empty | 1000.00–1007.35 | 1014.29–1016.70 | 1129.65–1141.79 | 1155.34–1172.72 |
| Profiled clock | 1047.53–1054.42 | 1036.44–1039.35 | 1161.49–1169.31 | 1167.05–1185.72 |

| Fixed same-round comparison | p50 difference | Mean difference | >5% central flags |
| --- | ---: | ---: | ---: |
| Opaque → split | +0.57% to +3.37% | −0.02% to +0.84% | 0 |
| Split → profiled empty | −0.85% to +0.66% | −0.15% to +0.51% | 0 |
| Profiled empty → clock | +4.08% to +5.38% | +2.09% to +2.34% | 3 p50 |

The three flags are all three rounds of cycle 0. All 27 comparisons and all
54 central checks are retained; no process was rerun. The clock recorder's
separate median calibration is 490–520 ns, which does not explain the full
route difference and is not subtracted. Clock instrumentation and compiler
layout can affect more than the callback's direct clock cost. Consequently,
the following fractions describe the observed route, not precise ordinary
commit fractions or prospective savings.

| Observed commit phase | Process median duration (µs) | Median same-owner whole fraction |
| --- | ---: | ---: |
| Document commit | 28.70–29.01 | 2.73–2.79% |
| Before payload capture | 26.48–26.94 | 2.51–2.57% |
| Embedded editor open | 22.77–23.00 | 2.14–2.19% |
| Live document read/check | 0.32–0.35 | 0.031–0.034% |
| Embedded finish | 188.49–190.83 | 16.73–16.87% |
| Unrelated stream validation | 13.93–14.09 | 1.33–1.35% |
| Public reopen | 104.75–105.69 | 9.68–9.84% |
| After payload capture | 29.47–39.57 | 3.11–3.75% |
| Artifact hash before | 173.84–174.00 | 16.51–16.66% |
| Artifact hash after | 176.08–176.26 | 16.73–16.89% |

Combined hashing is 350.06–350.49 µs and 33.25–33.87% of the observed whole,
or 42.77–43.40% of observed commit. These combined ratios are calculated per
owner before taking medians, not by adding component medians. Commit itself
occupies 77.07–77.27% of the observed whole, with 54.83–56.04 µs median residual
inside commit outside the ten event spans. Medians of durations and ratios
need not divide into one another. All 450 clock owners carry complete balanced
20-event traces: 9,000 main-matrix events, each contained in its commit window.

## Validation and disposition

All twelve final quality commands pass: feature-off check, feature-on PPT
all-target tests (1,206 passed; three existing ignored), Clippy with warnings
denied, doctests (14 passed; eight existing ignored), documentation, probe
format/tests (nine passed)/Clippy/docs/build, and repository boundary audit.
Feature-on tests cover real-edit output/patch/inverse equality, exact no-op
source sharing, formatting-only publication, source mismatch and absent
persisted-record typed errors. Three failed build attempts and their exact
source/probe snapshots remain visible; the successful receipt is `quality-3`.

All four qualification processes and all 36 native processes match the sealed
oracle. The independent audit recomputes process statistics, phase arithmetic,
same-owner denominators, route comparisons, and custody. The
actual analyzer accepts the complete packet and rejects 23 isolated corruption
controls. The first synthetic preflight exposed a shallow temporary-path import
assumption; its script/freeze snapshot and failure note are retained. After the
isolation fix, preflight passed before the sole native matrix was collected.

Retain the opt-in diagnostic seam. Native hashing remains substantial in the
observed path, but its two required digests are different artifacts and cannot
simply be removed. Next investigate the work inside embedded finish, where
observed time is also substantial, preserving complete publication validation
and comparing any candidate on the ordinary public path. Refine the clock
control before claiming exact ordinary phase fractions. No validation bypass,
unbounded digest cache, save-policy change, or claimed speedup is justified.

## Scope

All routes use the same diagnostic-feature binary. Their controls measure
splitting, compiler and observer effects together; they do not compare default
feature code generation. This one warm in-memory fixture establishes no broad
producer coverage, cold I/O, allocation, RSS, concurrency or speedup result.
Required preservation validation, artifact hashes and Reuse policy remain.
The non-iWork performance goal remains active.

[Retained evidence and replay instructions](results/change-0732/README.md).
