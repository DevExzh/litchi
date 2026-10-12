# 0210 — Checked PPT source policies

PPT now exports `SourceLimits`, `SourceLimitsBuilder` and `SourceLimitKind`.
The public `RecordLimits` name and writable limit fields are removed. Consuming
`max_<dimension>` setters reject zero and values above each hard ceiling with
`Error::InvalidLimit { kind, value, maximum }`; ingress/parser refusals use
`Error::Limit { kind, observed, maximum }`. Other domain resource refusals
retain their existing classification. The facade maps invalid profiles to
`invalid_argument` and exceeded profiles to `limit_exceeded`.

`OpenOptions` has private fields and canonical `password`, `leniency` and
`source_limits` setters, matching `get_*` accessors, strict defaults and
credential-redacted debug output. `to_builder()` preserves a selected profile.
Package projection limits combine component-wise with the admitted profile;
unspecified options do not silently replace a deliberately larger profile
with defaults. The facade uses `ppt_source_limits` and retains that profile
for its editor. Its positional metadata projection follows source admission.

## Dimensions and admission

| Dimension | Default | Hard ceiling |
|---|---:|---:|
| Complete input bytes | 512 MiB | 1 GiB |
| Declared stream entry bytes | 512 MiB | 512 MiB |
| Aggregate declared stream bytes | 2 GiB | 2 GiB |
| Encoded record bytes | 256 MiB | 256 MiB |
| Record events | 1,000,000 | 1,000,000 |
| Record depth | 256 | 256 |
| Individual record payload bytes | 256 MiB − 8 | 256 MiB − 8 |
| Aggregate record payload bytes | 768 MiB | 2 GiB |

The old complete-input default was 1 GiB, aggregate-stream default 768 MiB
and depth default 128. The new defaults follow the shared policy vocabulary.
Record payload accounting retains the existing per-record charge whether
payload storage is borrowed or copied, including nested container payloads.
This is distinct from the declared-stream aggregate.

Eager, already-parsed CFB and positional constructors validate complete input
size and every declared stream before exposing payload-reading APIs. Metadata,
opaque and embedded streams count; the root mini-stream backing does not,
because its child stream lengths already count. Stricter projection policies
repeat admission before payload reads. Positional metadata scans fence the
source before and after, including failed limit checks, so stale-source errors
take precedence. The Pictures payload remains deferred on positional reads.
The performance harness now computes exact aggregate limits from all declared
streams rather than only the three presentation streams.

## Verification

All 1,359 native PPT and 1,548 all-feature facade tests/doctests pass.
A final focused native run passes 12 checks, including two additional
regressions for above-default policy retention and core error dimensions;
a final facade regression checks invalid-profile classification and retained
native error details. Existing ignored native cases remain ignored.

Coverage includes zero/over-ceiling profiles, exact input boundaries, record
payload/event/depth gates, retained stricter policies, source changes, metadata
and unrelated-stream entry/aggregate overflow, and rejection before stream
payload reads. Existing leniency, encryption, fonts, Pictures deferral,
source-backed editing and replay tests pass with typed source-limit errors.

Warning-denied all-target Clippy and rustdoc pass for PPT and the facade.
Minimal and PPT-only feature checks pass. The new PPT glossary profile passes
14 banned-name checks; it deliberately declares no canonical primary
transaction family yet. The crate-boundary audit passes for 66 packages,
268 internal dependency declarations and 11 existing debt items. The generated
OfficeCLI table is current and all 15 generator tests pass.

The standalone performance harness compiles along with its umbrella consumer,
and `ppt_picture_selectors_preserve_semantics_and_phase_read_evidence` passes.
It checks deterministic corpus bytes, exact picture semantics and positional
read behavior with the migrated source profiles. No benchmark speedup or
performance acceptance is claimed. Run it with `cargo test --manifest-path
tools/perf-baseline/Cargo.toml --lib
ppt_picture_selectors_preserve_semantics_and_phase_read_evidence`.

Workspace checks use an external offline lock selecting cached cc 1.6.0 and
an external build target; the standalone harness uses a separate copied lock. Reproduce with `cargo test -p litchi-ppt --all-features`,
`cargo test -p litchi-ops --all-features`,
`cargo clippy -p litchi-ppt -p litchi-ops --all-features --all-targets -- -D warnings`,
and `python3 tools/check_api_glossary.py --profile ppt-minimal`.

## Remaining scope

This is a PPT API/source-admission prerequisite. The primary eager owner is
still generic and mutable; canonical immutable Package, primary transactions,
Builder and the unified PPT read/write adapter remain required. It does not
establish performance acceptance, Python/CLI completion or all-format goal
completion. Existing XLS performance follow-ups remain open.

The base remains an ancestor at `32ad64e3207655b4ba558fc2793d9a7396655cdf`;
no additional rebase was needed. Repository lockfiles remain unchanged.
