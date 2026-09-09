# Change 0484 one-shot replay-store measurement plan

Status: **design proposal, unmeasured**. This note closes the measurement
design gap identified in [`measurement-plan-review.md`](measurement-plan-review.md#scope-blockers-against-the-0484-design-claim).
It does not turn the current deterministic correctness checkpoint into a
performance result, and it does not revise the open full-document goal.

The current ten-case deterministic driver design remains a valid, separate
arm.
The no-argument harness invocation keeps that route and its existing report
shape. The store extension is opt-in and uses the already public
`tail_append_plain_paragraphs_from_producer` / `AuthoredReplayStore` seam. A
filesystem store in this plan is a harness-owned implementation of that seam;
production remains provider-based and does not gain an ambient path or a
temporary-file fallback.

## Routes and lifecycle

Every route uses the same source archive, authored event generator, text mode,
chunk partition, finite limits, package writer, 4 KiB non-retaining hashing
sink, and independent candidate/oracle artifacts. The route is an additional
axis. Source paragraph count and authored paragraph count remain independent;
varying one never derives or scales the other.

| Route | Authored input | Timed store work | Expected authored replay opens |
| --- | --- | --- | ---: |
| `deterministic` | Existing `GeneratedParagraphSource` and a fresh cursor per pass | None; each pass encodes through its bounded replay window | 5 |
| `memory_store` | A one-shot `GeneratedParagraphProducer` emitting the same events into the production `MemoryReplayStore` | `prepare_for_operation`, producer/encoder ingestion, every append, proof hashing, and `finish` | 4 |
| `file_store` | The same one-shot producer and a harness-only bounded positional file store | File creation, bounded append, flush/hash proof, and `finish` | 4 |

The deterministic five opens are the existing sealing, DOCX candidate,
OPC-candidate, and two publication passes. A one-shot route has one producer
invocation and one store finish instead of a sealing open, followed by the same
four replay-reader passes. The counters must make this distinction explicit;
the store arm must not report a producer invocation as a replay open.

The timed scope starts before source package admission and before creation of
the empty per-iteration replay store. It includes source reads, one-shot
producer execution, encoder and store append calls, store finalization, all
candidate and publication replay opens, package publication, sink hashing, and
operation-owned drops. It ends only after the sink is finished and the source,
edit, plan, provider, store handle, and sink have dropped. The deterministic
route keeps its current timing brackets.

Source archive construction, source-file population or source-cache warming,
authored proof/oracle construction, and the one untimed route-correctness
preflight are outside the sample timer. The replay store's authored ingestion
is part of the operation and is always timed. For the file route, only the
attempt directory is prepared outside the timer; creating the unique file,
writing authored bytes, flushing, hashing, and closing it are timed. A measured
sample must never reuse prefilled store bytes or a candidate output buffer.

The sink continues to hash and discard writes. It records accepted bytes,
write calls, largest write, histogram, and digest, but never collects the
candidate archive. A candidate archive `Vec` may exist in the untimed fixture
oracle, with its custody and size stated separately; no output-size buffer is
hidden in a timed route.

## Harness implementation

The small additive implementation belongs in
`tools/perf-baseline/src/docx_replayable_tail_append.rs`; the binary wrapper in
`tools/perf-baseline/src/bin/docx_replayable_tail_append.rs` remains thin.
No production crate or TOML change is required for these measurement arms.

The harness should add an internal `ProviderMode` with
`Deterministic`, `MemoryStore`, and `FileStore` values. The default is
`Deterministic`, preserving the current `Config`, lifecycle, and report
compatibility. `GeneratedParagraphProducer` must share the exact bounded
`fill_text`, event framing, chunk limits, and proof inputs with
`GeneratedParagraphSource`; it emits once to `ParagraphEventSink` and does not
build an event or XML `Vec`.

The memory route should use the production `MemoryReplayStore` directly,
wrapped by a harness counter adapter. The adapter counts `prepare`, append
calls and bytes, `finish`, handle opens, reader calls and bytes, and terminal
proof checks while delegating storage and bounds to the production store. A
generic wrapper around `AuthoredReplayStore` and its returned handle keeps the
production implementation under test rather than substituting a harness copy.

The file route should add a private `FileReplayStore` and authenticated reader
only to the performance harness. Its contract is:

* The caller supplies an attempt-scoped root and a maximum retained-byte
  ceiling. The store accepts only paths below that root, uses `create_new` for
  a unique per-iteration file, and rejects each append before the write if the
  next logical length would exceed the ceiling.
* Appends write encoder chunks directly to the file. `finish` flushes the
  selected sync policy, obtains the actual file length, hashes the exact file
  through a fixed buffer, and compares length and SHA-256 with the authored
  proof. It records actual logical bytes and, on the target filesystem, actual
  allocated blocks separately. The configured ceiling is never reported as
  retained or allocated bytes.
* Each reader opens the same caller-bound file from offset zero with a fixed
  buffer. It checks the stable length and provider identity before reading,
  hashes the bytes while returning them, and checks the proof at EOF. The
  handle retains the opened file identity, so a later path replacement cannot
  redirect that reader to a different inode. A replacement or length change
  observed before external publication is a preflight refusal with zero sink
  output. An in-place mutation discovered after publication has accepted a
  prefix must retain the typed `IncompleteOutput` result and exact accepted
  byte count; an EOF digest check cannot promise that every TOCTOU mutation is
  detected before output begins.
* The sealed handle carries a bounded provider reference containing a version,
  attempt-local relative file identity, stable length, and SHA-256. The report
  records the reference kind, bounded token length, and token hash, while the
  raw token and absolute path remain outside the report. A digest alone is not
  a replay reference: the handle must identify the caller-owned provider that
  can reopen the bytes.

The file store is therefore an explicit test provider, not evidence that the
production package may create or resolve arbitrary paths. Its resident-memory
claim is limited to the fixed reader window and package owners. Its retained
file bytes, filesystem blocks, page-cache state, and cleanup are separate
owners.

The route-specific observation should be additive to the current authored
record:

```text
route: deterministic | memory_store | file_store
producer_invocations
store_prepare_calls
store_append_calls
store_appended_bytes
store_finish_calls
replay_opens
replay_read_calls
replay_returned_bytes
replay_sha256_checks
retained_logical_bytes
retained_capacity_bytes       # memory only; actual capacity, not the limit
file_logical_bytes            # file only; actual length
file_allocated_bytes          # file only; stat result when available
file_write_calls
file_sync_calls
file_cleanup_verified
durable_reference_kind
durable_reference_bytes
durable_reference_sha256
```

The existing source request/returned counters and histograms, authored event
and text counters, sink histograms, allocator sample, and process delta remain
separate records. For deterministic samples the exact checks remain
`replay_opens == 5`, authored events/text equal the proof multiplied by five,
and no store counters are present. For either store route the checks are
`producer_invocations == 1`, `store_prepare_calls == 1`,
`store_finish_calls == 1`, `replay_opens == 4`,
`store_appended_bytes == proof.encoded_xml_bytes`, and
`replay_returned_bytes == proof.encoded_xml_bytes * 4`; every replay reader
must finish with the same encoded proof. Append-call count is reported and
must be positive, but is not fixed independently of encoder chunking.

Memory retention is reported as actual logical bytes and actual reserved
capacity. The production memory store may reserve its selected ceiling, so
allocator evidence must distinguish that reservation from the encoded logical
payload and from each live replay window. The file route reports its actual
file length and filesystem allocation when observable; it must use `null` or
an explicit unavailable state when block accounting is unavailable. Neither
route may replace an unknown physical quantity with its configured budget.

The route extension must preserve the other measurement axes even when the
first store protocol selects one value for each. Every case identity and
report carries `input_mode`, `sink_write_bytes`, and `compression_mode`:

* `input_mode=owned` keeps the current caller-owned positional source. The
  planned filesystem-backed, short-read/range, and latency adapters must use
  the same source length/hash and returned-byte proof, with source-file/cache
  population outside the timer and source reads inside it.
* `sink_write_bytes` remains the current 4 KiB default and continues to feed
  the non-retaining sink histogram. Future 512-byte and 64 KiB sink arms are
  separate protocol cases, never silently averaged into the 4 KiB result.
* `compression_mode` identifies the selected OPC member framing, initially the
  current package-writer profile and later explicit `store` and `deflate`
  profiles. Each compression profile gets its own expected candidate archive,
  raw-member oracle, output length/hash, and process inventory.

The initial store protocol fixes `owned`, 4 KiB, and the current compression
profile to keep the missing storage arm small. The fields and validators are
still required from the first implementation, so adding an input, sink, or
compression arm cannot accidentally compare different physical workloads or
reuse an oracle from another profile.

## CLI and driver protocol

The CLI additions should be small and explicit:

```text
--authored-provider deterministic|memory-store|file-store
--replay-dir DIR                 # required for file-store
--replay-max-bytes BYTES         # explicit store ceiling
--replay-sync none|data           # fixed and recorded per profile
--input-mode owned|file|short-read|latency
--compression-mode current|store|deflate
```

With no provider flag, the current deterministic invocation and v1 report are
unchanged. Store reports add a route/storage object and an authored provider
label; the driver accepts those additive fields only when the selected route
is a store arm. `--replay-dir` is never inferred from `TMPDIR`, a process
global, or a production configuration. The driver creates an attempt-specific
directory and passes its absolute path as an explicit caller capability.

No `protocol.json` is assumed to exist or to have been frozen by this
documentation task. Keep the current ten-case deterministic driver design
unchanged, then add a separate store-arm protocol at the root's later freeze
with the same ten deduplicated source/authored/chunk/text cases and the route
axis `{deterministic, memory_store, file_store}`. Its formal inventory is
exactly:

```text
3 routes × 10 cases × 2 roles (normal, allocator) × 2 repeats = 120 child processes
```

Each child process receives exactly one route/case/role/repeat tuple, runs the
same 30 measured samples and three warmups, and writes one non-replacing
receipt. The pilot inventory is 3 routes × 10 cases × 2 roles = 60 processes.
When frozen, the current deterministic design remains a separate 40-process
inventory. The repeated deterministic route in the new protocol provides
same-protocol comparison custody without changing that design; no historical
capture is implied here.

The driver must bind route, store ceiling, sync policy, replay-root token,
source manifest, copied binary hash, argv, environment, and profile to every
receipt. Its report validator should enforce the route-specific counter
arithmetic above, exact candidate sink bytes/hash, authored proof equality, raw
member preservation flags, semantic/order/text equality, inverse equality,
and finite storage fields. It should reject a report that has only a positive
store budget but no actual append, finish, replay, length, or hash evidence.

## Filesystem, page-cache, profiling, and cleanup custody

The formal file profile is named explicitly, for example
`filesystem_warm_write_then_replay`, and records filesystem type/device when
available, sync policy, page-cache policy, and whether physical block data was
observed. The process-level `/usr/bin/time -v` record and existing procfs
delta remain whole-child or best-effort observations; they are not relabeled
as operation-only file RSS. Allocator samples continue to run in a fresh
allocator process with the same route and case.

The default profile does not claim a cold-cache result. A file is written and
then replayed four times in the same operation, so those reads are labeled
`write_then_replay` and page-cache warmness is part of the tested shape. A
separate cold profile may be added only when the runner can perform and retain
an explicit cache-control receipt outside the timed interval. If cache control
is unavailable, the run is labeled unknown and cannot support a cold/warm
comparison. Repeats and case order are retained so cache drift is visible.

The driver uses a physical cleanup guard. Every file name includes a fresh
attempt/case/iteration token and is created exclusively. The child writes a
small ownership marker before the first append, closes the file before
cleanup, verifies the recorded length/hash, and unlinks only files whose marker
and parent are under the requested attempt root. Unlinking and directory
verification occur after the timer and after all operation-owned drops. The
parent receipt requires the replay directory to be empty at successful exit;
on a crash or failed cleanup it preserves the directory for diagnosis and
marks the process failed. No recursive deletion or cleanup outside the
caller-selected root is permitted.

## Correctness gate and remaining scope

Before formal samples, each route gets one untimed correctness preflight using
the same source and authored vector. It must prove identical candidate XML,
semantic paragraph order/text, source immutability, untouched raw local and
central records with only the local-offset normalization, physical member
order, opaque member identity, exact candidate archive hash, and immediate
inverse. The route's store proof and every replay reader proof must match the
independent authored event and encoded-byte proof. A replay-file metadata,
length, byte, or reference mutation detected before external publication must
leave the caller sink untouched. If a mutation is detected during actual
publication after a prefix was accepted, the result must be the typed
`IncompleteOutput` failure with the exact accepted count and the mutation
diagnostic preserved.

The first formal extension keeps the existing owned-byte positional source,
4 KiB sink, and default package writer so that storage-route effects remain
interpretable. Filesystem-backed source input, instrumented short reads,
high-latency ranges, multiple sink sizes, and Store versus Deflate selected
member cases remain separate required arms from the design. They need their
own source/sink/compression fields and process inventory rather than being
silently folded into this store protocol.

Implementation scope for the store arm is the harness module and thin CLI,
the store-arm driver/validator and its focused Python tests, plus retained
protocol/receipt schema documentation. No production Rust, TOML, or default
deterministic protocol change is part of this source-freeze task. Formal
captures remain pending until the correctness gate, source custody, route
counter checks, filesystem cleanup guard, and normal/allocator receipts all
pass. No route comparison or speedup claim follows from this plan alone.
