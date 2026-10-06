# 0942 CFB operation profile harness

This standalone crate is a harness-only derivative of the qualified 0917 DOC
repeated-edit probe. It keeps the same public lifecycle, correctness gates,
JSON schema, fixture identities, and report fields. The two binaries have
0942-specific names: `doc_retention_probe_0942_cfb_profile` reports the
existing native samples, and `doc_retention_probe_0942_cfb_profile_alloc`
uses the 0730 counting `System` allocator in a separate lane. The allocator
lane reports allocation counters and makes no timing claim.

Only the timed and allocator measurement paths call the uniquely named,
`#[inline(never)]` `profiled_measured_case` wrapper. The wrapper calls the
unchanged `run_case`, applies a post-call `black_box` barrier, and returns the
same result. Preparation, correctness routes, and per-sample validation call
`run_case` directly. Build the profile binary with frame pointers and debug
information, then filter callchains containing `profiled_measured_case`; the
instrumented binary is for operation attribution and is not a timing
qualification binary.

The CLI contract is:

```text
--case <docnohf|docfloat> --input PATH --edits <1|2|8|32> \
--retention <zero|default|release> --warmup N --samples N
```

The native defaults are three warmups and thirty samples. The allocator lane
requires `--warmup 0 --samples 1`. Inputs are the real
`test-data/ole/doc/NoHeadFoot.doc` and `FloatingPictures.doc` fixtures; their
byte counts and SHA-256 values are checked before preparation.

Every stage replaces paragraph zero in the same public `Edit`, alternating the
two exact texts reported as `alternating_texts`. Edit counts therefore mean
repeated replacements of one paragraph, rather than a bulk edit of distinct
paragraphs. The three retention routes use zero retained-render bytes, an
8 MiB ceiling, and the same 8 MiB ceiling followed by
`release_retained_render()` after each successful stage. Stage observations use
a fixed 32-entry array and nullable capacities, so the measured lifecycle does
not allocate an observation vector.

The native clock and allocator region cover only
`Snapshot::open_bounded` from the source clone through one `Edit`, all staged
replacements and releases, `commit`, the returned output `Vec` copy, and local
snapshot/edit/commit destruction. Source loading, reference routes, hashes,
public-reader checks, the 0730 DOC oracle, and JSON serialization are outside
the measured region. Output is copied before the region ends and validated
afterward.

Preparation requires exact cross-route output equality, all stage text changes,
final paragraph and public reader readback, exact no-op and source allocation
reuse, forward and inverse patches, foreign and stale patch refusal, atomic
same-value and operation-limit controls, capacity ceilings, retention release,
semantic patch counts, and preservation of untouched streams and directory
metadata. The focused `oracle_0730` module retains the qualified DOC checks:
complete CFB stream/path bytes, raw directory normalization, CLSIDs, and DOC
paragraph, story, table, field, revision, and embedded-object survivor
projections. PPT code and the old 0730 executable/timer paths are not part of
this crate.

The JSON schema remains `litchi.performance.0917.doc_retention.v1`. It reports
source/output hashes and byte counts, selected paragraph, edit count, route
ceilings, nullable capacities per stage, raw native elapsed samples, and
allocator bytes/calls/peak/live retained values. The production crates are not
modified by this harness. Root owns the immutable export, final build, capture
schedule, and evidence custody; this directory contains source only and no
binaries or measured captures.
