# 0361f — IWA raw metadata constructors

- Date: 2026-10-21.
- Original base: `a9dad32ed19aaef0186bc221d5d640c7ef7a9a25`.
- Coordinator integration and measurement base: `ddf2ac8ae2`.
- Contract: 0351 §2.5 and 0355a §2.
- Status: source migration and bounded validation complete on the candidate
  branch; performance qualification remains required before acceptance.

The schema-neutral IWA metadata constructors now have one required limit
profile and one fallible shape:

```text
FieldPath::new(Vec<u32>, Limits) -> Result<FieldPath>
FieldInfo::new(FieldPath, Limits) -> Result<FieldInfo>
MessageInfo::new(u32, usize, Limits) -> Result<MessageInfo>
ArchiveInfo::new(u64, Vec<MessageInfo>, Limits) -> Result<ArchiveInfo>
```

`From<Vec<u32>>` is removed because it cannot carry a raw admission profile.
Empty paths, zero path components and identifier zero retain their existing
wire and later-validation semantics. The fixed `[1, 0, 5]` message-version
vector is reserved only after the message length, header and metadata checks
pass. Payload lengths are checked against the message-byte limit before the
checked `usize` to `u32` conversion, so a large caller length cannot be
silently truncated or change the configured-limit precedence.

Each constructor admits the neutral record it publishes against encoded
header bytes, wire-field count, nesting, neutral header memory and metadata
items. `ArchiveInfo` additionally checks message count and every nested
`MessageInfo`/`FieldInfo` shape. The neutral memory calculation counts the
logical bytes of the owned record model (collection/string lengths, not spare
caller capacity); this is the same stable logical accounting used by archive
ingress. Parser decoder scratch remains admitted by the existing wire
preflight before codec allocations. The hard nesting ceiling is
`Limits::MAX_HEADER_NESTING` (64); constructors use the caller's validated
profile and do not hardcode a replacement default. The canonical archive
schema reaches depth three for an archive containing a message, field info and
field path.

`archive_info_encoded_len` now computes the canonical neutral length without
cloning the compatibility codec. It uses checked varint, packed-vector and
length-delimited arithmetic and keeps optional-field presence and unknown
closed-enum values, including negative enum values. The existing codec
encoder remains the publication path; the focused parity test compares the
allocation-free length with the codec length for full metadata and unknown
enum values.

`ArchiveObject::new` preflights the complete conventional header directly
from borrowed `RawMessage` metadata before reserving its message-info
collection or fixed version vectors. It then constructs and publishes the
already-admitted `ArchiveInfo` without a second full metadata walk. Public raw
fields remain available for source-preserving record work;
`ArchiveInfo::validate` performs the same complete neutral admission at
publication time. Numbers comment creation/reply transitions and Pages
footnote reference creation pass their retained archive profile. All other
format, test, fuzz and tool consumers use the canonical result and explicit
profile; no alias or default-only constructor remains.

The focused core test covers exact and one-under limits for path, field,
message and archive header bytes, field counts, nesting, neutral memory,
metadata items, message count and payload bytes. It also covers zero/empty
values, checked `u32` overflow and length parity for optional and unknown
metadata, plus the aggregate object preflight rejecting a raw-message list
before metadata construction. A test-only `MessageInfo::try_new` entry counter
proves that each rejected fields/bytes/memory profile constructs no metadata,
while an admitted two-message object constructs exactly two records.

The coordinator integrated the complete source migration onto `ddf2ac8ae2`
without a production-code repair. The integrated core test suite and
all-target/all-feature checks of the archive, IWA, Pages, Numbers and Keynote
consumers pass. Receipts remain under the external
`constructor-0361f-coordinator` directory. The original worker's broader
validation claims are not counted until their receipts are located and
verified; integration does not imply native performance acceptance.

The coordinator replaced the prepared qualification protocol before building
it. The earlier protocol timed its counting-allocator binary, omitted host
guards and warmups, and calibrated from the slower endpoint. It has no
accepted measurements. The current builder creates protected native and
counting binaries from explicit, complete committed source snapshots. Shared
path dependencies, resolved registry versions/checksums, toolchain, lockfile,
probe and commands are recorded and checked. Native binaries have no counting
allocator. Every generated artifact and source tar remains outside Git.

```sh
python3 docs/unified-ops/results/0361f/constructor-qualify.py build /outside/0361f \
  --base ddf2ac8ae2 --candidate CANDIDATE_COMMIT
python3 docs/unified-ops/results/0361f/constructor-qualify.py allocation /outside/0361f \
  --attempt allocation-a1
```

The shared probe validates exact constructor fields and canonical archive
bytes before timing. It does not hash outputs inside the encode timing loop.
Separate allocator processes reset their counters immediately before each
route, count full reallocation requests, and report peak/retained/requested
bytes and calls. These receipts make no latency or process-RSS claim.

The native runner uses five complete ABBA blocks, two warmups, common
iteration counts calibrated from the faster endpoint to at least 500 ms,
and a minimum measured leg of 150 ms. It records each invocation's raw
stdout/stderr, command, exit and host processes before/during/after the child;
actual executable and script identities reject overlapping builds/profilers
and benchmark jobs. Exact outputs and semantic checksums must agree. All
attempts, including interrupted runs and earlier clean blocks, remain evidence.

```sh
python3 docs/unified-ops/results/0361f/constructor-abba.py \
  /outside/0361f/build-base/constructor-native \
  /outside/0361f/build-candidate/constructor-native \
  /outside/0361f/native-a1 --cpu 29
```

The nine routes are `FieldPath::new`, `FieldInfo::new`, `MessageInfo::new`,
`ArchiveInfo::new`, `ArchiveObject::new`, tight-header-byte object refusal, archive
parse, validate and encode. The tight-header route measures preflight; the
unit test remains the correctness authority for rejection before metadata
buffers are reserved.

This slice does not close the remaining shared raw-constructor audit or the
integrated W4/W4b/W8 gate. Those rows remain required engineering work.

The first coordinator build (`matched-v1`) stopped during base validation:
the prepared tight-header-field refusal assumed enforcement absent in the
base constructor. It produced no qualified timings. The raw traceback,
protected binary and original probe are retained. The corrected comparison
uses a one-byte header ceiling, enforced by both endpoints; candidate-only
header-field rejection remains covered by its correctness tests.

The corrected `matched-v2` fixture and repeated allocation checks pass. Its
standalone probe inherited the earlier preparation's debug-enabled, non-LTO
release profile, so those counts remain diagnostic and no native timing was
started. The final builder uses the workspace release settings (`lto = true`,
`panic = "abort"`, default debug level) for both native and counting binaries.
Both earlier build attempts and their protocol sources remain external.

The production-profile `matched-v3` build passes all four binary preflights:
the six constructor/archive outputs match across native/counting builds and
both endpoints. Four repeated allocation legs per route agree. Archive-object
construction reduces requested bytes from 566,400 to 176,000 per 100 calls;
parse, validation and encoding also allocate less. No route increases requested,
peak or retained bytes. These are allocation results, not native latency
acceptance. [The verified receipt](results/0361f/evidence.json) records the
complete build, source manifests, toolchain, dependencies and measurements.
Regenerate it with `constructor-collect.py /outside/0361f --out OUTPUT`.

Both `matched-v3/native-a1` and `native-a2` terminated with exit 1 when the
host guard observed an unrelated Python test process. The first attempt
stopped during warmup; the second stopped within its first ABBA block.
Neither produced a complete qualified block. Their original manifests retain
the last recorded `running` status because the runner raised on interference;
the process exits, logs, blocked host snapshots and every partial sample are
retained in the receipt. Neither process is still running. Native latency
remains unqualified, and these incomplete samples are not acceptance evidence.
