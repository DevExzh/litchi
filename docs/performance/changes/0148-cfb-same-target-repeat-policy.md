# CFB target-aware repeat-policy harness extension

Date: 2026-08-16

Status: production policy and correctness evidence accepted. The follow-up
[clean release ABBA](0149-cfb-same-target-repeat-release-abba.md) accepts only
the named configured-simulator aggregate repeat result; local and resource
claims remain withheld. This change does not modify iWork coverage.

Change 0148 extends the production `SharedOleFile::open_stream` evidence
runner from the one-shot/repeat workloads in [0146](0146-cfb-open-stream-evidence.md)
and the historical release result in [0147](0147-cfb-open-stream-release-abba.md).
It keeps the default 36 cases / 198 records unchanged, while adding six
opt-in selectors (291 selectable names total):

- `cfb_open_stream_mini_shared_different_sid`
- `cfb_open_stream_mini_shared_bulk`
- `cfb_open_stream_mini_shared_concurrent`
- `cfb_open_stream_mini_4095_shared_different_sid`
- `cfb_open_stream_mini_4095_shared_bulk`
- `cfb_open_stream_mini_4095_shared_concurrent`

The selectors use the existing `many-small` and `wide-root` selective CFB
corpora, with target lengths 36 and 4095 bytes. Each result records the
ordered logical stream names, invocation/batch count, output hashes and
lengths, exact source positional events, source-version stability, and the
typed `OleError::StreamNotFound` refusal check. No private cache state is
serialized as if it were public.

## Accepted source vectors

Let `L` be the selected target payload length, `R` the declared root Mini
Stream byte length, and `D`/`C` the direct target and root-cache source events:

```text
D = [target_start, L, L]
C = [512, R, R]
```

The runner accepts all three policy generations needed for a clean
control/current/candidate comparison:

```text
same-target repeat-3:  control [R, 0, 0]
                      current [L, R, 0]
                   target-aware [L, L, L]

same-target repeat-8:  control [R, 0, 0, 0, 0, 0, 0, 0]
                      current [L, R, 0, 0, 0, 0, 0, 0]
                   target-aware [L, L, L, L, L, L, L, L]

different-SID A-B-A:  control [R, 0, 0]
                      current/target-aware [L, R, 0]

public bulk A-B-A:    control/target-aware aggregate {C}
                      prior current aggregate {D, C}

same-target overlap:  control aggregate {C}
                      direct/cache aggregate {D, C}
                      direct-repeat aggregate {D, D}
```

The target-aware repeat formula is asserted exactly in the harness unit tests;
the production evidence tests remain baseline-compatible by accepting the
control and prior-current vectors as well. Concurrent events are compared as
an aggregate/multiset because source request completion order is not a public
contract. The concurrent selector uses a harness-only condition-variable gate:
workers announce entry before `open_stream`, and a target-sized direct source
read is released only after both workers have entered. Root-sized reads do not
wait on that gate, so the control path cannot deadlock on the harness.

Bulk coverage calls only the public `SharedOleFile::bulk_read` /
`read_streams` API. The workload is one deterministic A-B-A batch, with one
worker and a byte budget covering the three returned payloads. It does not
call candidate-only production helpers.

## Evidence boundary and release follow-up

These selectors are correctness and source-event evidence. They do not claim a
latency, allocation, RSS, physical-I/O, native-format, CRUD, or generic
performance improvement. Failure/retry, ineligible-root, FAT, native semantic,
resource, and performance acceptance for the extended bulk/concurrent selectors
remain open. The existing 0147 raw/results are historical and are left
untouched.

The required comparison is now retained in
[change 0149](0149-cfb-same-target-repeat-release-abba.md): identical harness,
clean release worktrees, and explicit `A1 control, B1 candidate, B2 candidate,
A2 control` order. Its control restores both `shared.rs` and `shared_bulk.rs`.
The configured simulator accepts aggregate repeat-3/repeat-8 total improvements
of roughly 56-64% in both directions while retaining the later-invocation
direct-read tradeoff. Local wall-clock, per-invocation, bulk, concurrent,
allocation/RSS, physical-I/O, native-format, and generic claims remain withheld.
