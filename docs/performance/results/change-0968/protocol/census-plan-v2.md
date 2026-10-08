# 0968 staged CFB commit census plan v2

**Status:** diagnostic source and harness implementation is in the external
`source-diagnostic` tree. The implementation has only passed compile and unit
controls so far. Root review and formal capture GO are still required.

## Scope

There is one instrumented current baseline variant. The formal packet is the
existing 24-cell matrix (`NoHeadFoot` and `FloatingPictures`; 1, 2, 8, and 32
edits; zero, default, and release retention), with two fresh serial children
per cell: 48 children total. There is no optimization candidate and no paired
96-child comparison in this task.

Each child will use `--warmup 0 --samples 1`. Fixture preparation, the three
correctness routes, output hashing, JSON serialization, and all independent
oracles stay outside the census session. The measured lifecycle is the current
public DOC edit path. `timing_claim` is always `false`; the harness has no
clock or allocator instrumentation.

The ordinary 0967 executable may serve only as an output oracle if an
independent receipt proves byte identity for the runtime, harness, dependency
inputs, and configuration. A Git revision difference alone is not evidence of
staleness. No new ordinary build is needed when that identity holds.

## Diagnostic boundary

The feature-gated observer starts at
`ObjectEditor::commit_candidate_with_rendered` and remains active through
render, physical reopen, codec admission, package capture, stream-allocation
reuse, and object discovery. The commit scope is marked successful only after
object discovery returns successfully; any early return is a commit error.

Route observations are emitted only inside that scope:

- `overlay_not_attempted` records a non-`Reuse` policy;
- `overlay_attempted` partitions into `overlay_success`, `overlay_declined`, or
  `overlay_error`;
- each `render_with_layout` call partitions into
  `layout_reuse_success`, `layout_non_reuse`, or `layout_error` using the
  writer's successful `last_sector_layout()` report;
- `Package::capture` partitions into `capture_success` or `capture_error`;
- successful capture stream opens add one `capture_streams` count and the exact
  `capture_payload_bytes` length.

Capture counters are active only at the post-render `Package::capture` seam,
not at initial admission, object-target discovery, final output serialization,
or later serializers. No metadata-projection counters are included because
that route is not present in the current code. The counters are fixed scalar
fields and retain no input, stream data, path, or event vector.

## Child schema

Schema identifier: `litchi.performance.0968.cfb_commit_census.v2`.
Every packet contains the existing 0967 identity and correctness fields plus
one `sample` census object. Required census fields are:

```
commit_attempts, commit_successes, commit_errors,
overlay_attempted, overlay_success, overlay_declined, overlay_error,
overlay_not_attempted,
layout_attempted, layout_reuse_success, layout_non_reuse, layout_error,
capture_attempted, capture_success, capture_error,
capture_streams, capture_payload_bytes, saturated
```

The packet must report `warmups: 0`, `samples_requested: 1`,
`timing_claim: false`, and no elapsed or allocator fields. `route_outputs` and
the existing 20 correctness gates remain required; all gates must be true and
the measured output must equal the zero-retention reference output by bytes and
SHA-256.

The independent reducer must require exact key/type sets and reject missing or
extra census fields, `saturated: true`, negative/non-integer counters, duplicate
or missing cells, unequal repetitions, output mismatches, and any false
correctness gate. It may assert only observed stage conservation:

```
commit_attempts = commit_successes + commit_errors
overlay_attempted + overlay_not_attempted = commit_attempts
overlay_attempted = overlay_success + overlay_declined + overlay_error
layout_attempted = layout_reuse_success + layout_non_reuse + layout_error
capture_attempted = capture_success + capture_error
```

On a successful corpus sample (the required packet has zero commit, overlay,
layout, and capture errors), the reducer may apply these checks. Error packets
are rejected by the correctness gates before these equalities are used, so the
first two equalities do not claim to classify a pre-overlay `package.check`
failure:

```
commit_attempts = commit_successes
overlay_attempted + overlay_not_attempted = commit_successes
overlay_attempted = overlay_success + overlay_declined
layout_attempted = layout_reuse_success + layout_non_reuse
capture_attempted = capture_success
```

The reducer must not infer that every overlay decline reaches layout in an
erroring sample or that a capture stream total is allocator bytes.
`capture_payload_bytes` is the exact sum of bytes returned by successful stream
reads inside the captured package; it is a diagnostic recapture workload signal
only.

## Build and custody

Source baseline remains immutable. The diagnostic tree is an external copy of
current HEAD plus only the feature-gated observer and the separate census
harness. Build receipts will record the exact source tree hash, feature/config
flags, command, compiler identity, and binary SHA-256. No formal child launch
occurs until the source, harness, schema, and checker are sealed and root gives
GO.
