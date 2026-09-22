# 0735 PPT record-staging source review

This independent review covers the 0735 candidate for the private embedded
PowerPoint editor's complete-record staging functions. It compares the frozen
baseline with the candidate copy and makes no production, Cargo, probe, native,
formatter, or profiler change. The review is source-only; the frozen quality
and process measurements still decide whether the production hunk is retained.

Disposition: **approved for the frozen semantic and performance gates**. I
found no correctness blocker in removing the complete `Editor` clone from
`replace_persisted_record` and `insert_persisted_record`. Their typed checks,
check order, ownership of the caller's `Vec<u8>`, and `changed` transition are
unchanged. The candidate mutates only the staging map and the scalar changed
flag after all recoverable checks have succeeded.

The reviewed source identities are:

| item | path | SHA-256 |
| --- | --- | --- |
| archived baseline | [`source-archive/before/.../records.rs`](source-archive/before/crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs) | `06ec08f367a1d1f72544d91cac5f4764f7c8b5ef0f869c1aeb3f445a59f5940e` |
| candidate copy | [`candidate/.../records.rs`](candidate/crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs) | `0fe5168733248170e4d7bbedea39d348bae402177431649f430118c808f30264` |
| formatted live source | [`records.rs`](../../../../crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs) | `858bd401710d8aa780222b6edde3d0538f6d884252c4dd83300cc0bc723a0b69` |
| canonical after archive | [`source-archive/after/.../records.rs`](source-archive/after/crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs) | `858bd401710d8aa780222b6edde3d0538f6d884252c4dd83300cc0bc723a0b69` |
| candidate notes | [`implementation-notes.md`](candidate/implementation-notes.md) | `9d62167489ef938475dec98f4e8feb9d873114f6358d8d1d23a5553f0b9cf1aa` |
| source-equivalence receipt | [`source-equivalence.json`](source-equivalence.json) | `311f5bb42125529c243cddb1a079cb5f3454d4f3d5f4925e1392f717f6658761` |
| source guard | [`source-guard.py`](source-guard.py) | `112b24f41f3fc5140de6339cc516bdc984c175a3e5be0ecb37993c3e9aa5e2cf` |

The live production file and canonical after archive now match the formatted
candidate exactly. The coordinator applied and formatted the candidate only
after the baseline capture. The 0735 constraint snapshot is
[`constraints.json`](constraints.json), SHA-256
`13fbdce87065a406b83b58f85ca2f32668a49f4b8306f5012afd2a3092a20863`, and the
review is anchored to baseline commit `fdd63ee3867c4a4ab4c7b96c5374e68284f22319`.

The source guard passed after formatting. Its receipt records byte-identical
typed-validation prefixes of 602 bytes for replacement and 704 bytes for
insertion, with the only production transition difference after those
prefixes being the direct map insertion and `changed` update.

## Applicable constraints

I read the accepted ADR index and the relevant accepted records: [ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0002](../../../adr/0002-crate-topology.md), [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md), [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), [ADR 0008](../../../adr/0008-migration-and-verification.md), [ADR 0010](../../../adr/0010-facade-archive-ownership.md), [ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md), and [ADR 0024](../../../adr/0024-current-topology.md). The 0735 constraint manifest freezes all accepted ADR hashes.

The candidate stays inside the existing private `litchi-ppt` editor. It adds no
public type, dependency, unsafe code, cache, executor, ambient I/O, or new
publication policy. The existing OLE editor remains the owner of the staging
state, and the iWork exclusion is unaffected.

## Validation and refusal ordering

`replace_persisted_record` retains the exact baseline order:

1. It checks that the identifier is present in `mappings` and is not in
   `removed_persist_ids`.
2. It checks the eight-byte minimum, the 128 MiB maximum, and complete record
   framing through `rewrite::slice(&record, 0)`.
3. It inserts the owned record under the identifier and sets `changed`.

`insert_persisted_record` likewise keeps its exact baseline order:

1. It checks the nonzero, 20-bit identifier range and rejects identifiers in
   either `mappings` or `staged_storage`.
2. It applies the same record-size and complete-framing checks.
3. It inserts the owned record and sets `changed`.

The candidate did not move a check across another check, change an error class
or message, broaden an accepted identifier, or turn a typed refusal into a
partial edit. `rewrite::slice` is still the only framing authority. In every
recoverable error path the editor's fields are untouched; the caller-provided
record is consumed in the same way as before and is dropped on refusal.

After validation, `BTreeMap::insert` takes ownership of the already validated
`Vec<u8>`, and `changed = true` is a scalar store. There is no `Result`-returning
operation, source read, semantic parse, limit check, or publication step after
the checks. `BTreeMap::insert` can still encounter an ordinary process-level
allocation failure, just as the old clone and insertion path could. That event
is not a recoverable `Error::AllocationFailed` boundary and must not be
described as one. The key is `u32`, so ordering has no user code or fallible
comparison; replacing an existing staged vector only drops a `Vec<u8>`.

The old clone did not provide an additional recoverable failure boundary: its
`Editor::clone`, map insertion, and assignment exposed no `Result`, while all
typed validation already occurred before the clone. The direct path therefore
preserves the operation's failure-atomic contract for every returned error and
removes the full-editor allocation work. It does not claim panic recovery or
convert allocation aborts into typed errors.

## Public callsite scope

I traced every current callsite in `litchi-ppt`:

* The public `Editor::replace_persisted_record` method is an ordinary staging
  verb. A successful call is expected to change that editor's transaction-local
  state. Its refusal paths return before state mutation.
* The crate-private insertion method is used by the slide-order transfer
  publication path after identifier, master, picture, and dependency checks.
  The embedded editor is local to the publication attempt; any later typed
  error drops that local candidate and cannot mutate the source snapshot.
* Animation transactions clone their outer semantic editor before editing the
  embedded package. The stage helper completes record validation, reparsing,
  and round-trip checks before calling the staging verb. The remaining work in
  the successful path updates transaction-local vectors and flags; a process
  allocation failure remains a process failure rather than a returned typed
  error.
* Font transactions, document-comparison publication, slide-order commits,
  and text-edit publication all open or clone a local embedded editor, perform
  their source checks and record construction before staging, then either
  finish that local editor or discard it on a later error. Direct staging does
  not publish or mutate their immutable source snapshots.
* Existing embedded-editor tests and finish helpers use the same staged state
  and consume the editor at finish. No caller relies on the clone's temporary
  duplicate allocations or on an editor being unchanged after a successful
  staging call.

No callsite passes a borrowed alias to a private staged vector: the API takes a
`Vec<u8>` by value, and internal callers pass owned or cloned records. Replacing
an already staged record therefore has the same map-key and drop behavior as
the baseline.

## Focused evidence and residual notes

The candidate adds focused tests for unknown, removed, out-of-range, mapped,
already staged, too-short, malformed, trailing, and oversized records. Each
refusal compares every editor field with its pre-call state. Successful
replacement and insertion compare the direct transition with a test-only copy
of the former clone transition, retain unrelated staged records, and compare a
fixture-backed finished package for a replacement.

The focused tests do not attempt to manufacture allocator panics or a custom
`BTreeMap` comparator; neither is part of the recoverable `Result` contract,
and the production key type cannot invoke user comparison code. The oversized
case necessarily allocates a 128 MiB-plus test vector before exercising the
size check; this is an advisory test-resource cost, not a production semantic
issue. The required workspace gates and frozen public probes remain responsible
for compile, semantic, negative-control, allocation, latency, and tail evidence.

No independent review finding requires a source correction. Retain the
candidate only if the frozen output, semantic, refusal, allocation, latency,
peak-live-byte, and statistical gates pass. If those gates reject the change,
the archived baseline is sufficient to restore the one owned production file.
