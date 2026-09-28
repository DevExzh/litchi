# 0815 pre-build protocol review

The current production source is the retained 0813 direct event-result
dispatch at commit `55bb2ead3498043dd22b53555507b24402486248`; all 9,196
tracked source-file hashes are frozen before the first 0815 build. The new
candidate is limited to `notes/codec.rs` and binds the `Start` and `Empty`
event payloads by reference before passing them to the existing resolver and
inspection code. The four corresponding explicit borrows are removed. No
resolver, attribute merge, parser, error, test oracle, API, dependency, unsafe
code, or iWork change belongs to this trial.

The six-file public probe is copied byte-for-byte from sealed 0813. Root
`Cargo.lock`, `rustfmt.toml`, architecture inputs, unrelated workspace files,
toolchain observation, candidate archive, and driver identities are custody
inputs. The six production quality gates are reused only because the source
file manifest equals sealed 0813 after; no baseline Cargo run is implied.
Fresh probe formatting, all-feature release tests, and release Clippy remain
mandatory for both legs.

Before application, build fresh baseline native, allocation, and profile
executables and verify eighteen one-sample public-workflow qualification
reports against the sealed fixture, output, and semantic oracles. Apply only
the reviewed archived candidate after the qualification audit. Run all six
candidate production gates, including the buffered malformed/refusal
differential corpus. Build candidate executables and retain bounded ordinary
and profile scanner assembly for both legs.

The code-generation gate asks whether the borrowed `Start`/`Empty` bindings
remove the targeted arm-local copy sequences in the ordinary release scanner.
Static assembly evidence describes code shape only; it is not a runtime
frequency, phase fraction, causal cycle, or adoption claim. If the mechanism
gate fails, restore the exact baseline and stop the hypothesis.

After the mechanism gate, run six counterbalanced native blocks over all 18
capture/commit/lifecycle cases (216 reports and 6,480 samples), two paired
allocation blocks (72 reports and 216 samples), and four owner-scoped Callgrind
publications (4 reports and 4 samples). With qualification, the frozen total is
310 reports and 6,718 samples. Use nearest-rank quantiles and 10,000 median
bootstrap resamples with seed 815815 and sorted endpoints 250/9749. Retain all
tail, RSS, spread, unresolved-frame, and lost-event observations.

Adoption requires the stated 3 percent capture/lifecycle benefit with a
bootstrap upper endpoint below one, no latency veto, and non-increasing paired
allocation medians for calls, bytes, net live bytes, and peak-above-entry.
Semantic equality, all quality gates, independent replay, source custody, and
the final source disposition audit remain mandatory. Reject and restore if any
gate fails.

Root alone executes Cargo, formatting, workloads, profiling, and binary tools;
agents prepare and review source, drivers, and offline readers. Heavy offline
replay waits until all captures are terminal. Keep each live process handle
through completion and preserve actual failures in immutable receipts. After
an explicit retain/reject disposition, remove only the owned target, update
the indexes, audit exact staged paths, seal the packet, commit the batch, and
verify `HEAD`. The broader OLE2/OOXML performance goal remains active.
