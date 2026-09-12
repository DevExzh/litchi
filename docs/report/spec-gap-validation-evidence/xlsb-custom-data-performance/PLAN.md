# XLSB Custom Data lifecycle performance profile

Status: candidate-only characterization approved after independent review and root replay; see `review.json` and `root-replay-verification.json`.

This profile measures the public `litchi_xlsb::custom_data` lifecycle added by
commit `16102fe751d7c5492042330f1bd1f49c304495f0`. It uses only deterministic,
authored OPC+BIFF12 fixtures derived from
`crates/litchi-xlsb/tests/custom_data_lifecycle.rs`. The fixtures are valid
under the pinned ExtConn14 grammar, but they are not native-producer files.
The profile makes no baseline, speedup, Office-acceptance, or native-producer
claim.

The harness is built with Cargo's optimized `--release` profile, with
`CARGO_INCREMENTAL=0`, no compiler override flags, and the retained lockfile.
The allocator observer remains enabled in that optimized binary, so elapsed
values include its atomic accounting overhead.

## Matrix

Each size class varies both storage and host-reference cardinality:

| class | storages | references | payload per storage | purpose |
| --- | ---: | ---: | ---: | --- |
| small | 2 | 16 (many-to-one) | 1 KiB | ordinary source closure |
| medium | 4 | 128 (many-to-one plus one edge per remaining UID) | 8 KiB | catalog and graph growth |
| large | 16 | 512 (many-to-one plus one edge per remaining UID) | 32 KiB | larger retained payload and binding inventory |

The timed lanes are `read`, `snapshotclone`, `noop`, `editpayload`,
`rename-many-to-one`, `mixedinsertremove`, `patchinverse`, and `publicapply`.
Fixture creation, source hashing, commit/patch preparation, and expected-output
construction are setup work. `read` times one public catalog read;
`snapshotclone` times cloning a prepared immutable snapshot;
`noop` times a public empty transaction, commit, and no-op apply;
`editpayload` times payload staging and commit;
`rename-many-to-one` times a UID rename that rewrites all admitted references
and commits;
`mixedinsertremove` times one insertion, one unreferenced removal, and commit;
`patchinverse` times applying a prepared patch and its prepared inverse; and
`publicapply` times applying a prepared public commit.

The transaction lanes construct their caller replacement data inside the timed
operation. The `patchinverse` and `publicapply` lanes use equal prepared patch
or commit inputs outside timing for every sample. These are separate public
cost scopes; the receipts record the input-preparation boundary so their
elapsed values are not treated as an apples-to-apples speed comparison.

Every lane/class is run in a fresh process with five measured samples after
one warm-up. The elapsed values include the process-local allocator observer's
atomic accounting overhead; they characterize this instrumented binary rather
than a zero-overhead production run. The retained medians are descriptive
five-sample summaries, not tail-latency or uncertainty certification. The raw
JSONL sample files are retained. The runner records the exact command, binary
hash, source hash manifest, toolchain, and fixture hashes. A replay command
repeats the same matrix from a fresh target/result directory and the verifier
checks that all semantic and preservation gates still pass.

## Measurement boundary

The harness installs a process-local `GlobalAlloc` observer. Each sample
records allocation calls, direct requested bytes, successful realloc-new and
realloc-old bytes, deallocated bytes, live bytes before/after the timed call,
and peak live bytes observed during that call. It checks the allocator balance
equation and reports failed allocations or live-byte underflow. The
`logical_peak_bytes` field is the observed peak live allocation value; it is
not RSS and is not a language-level leak proof. It is sampled from aggregate
allocator counters and does not claim to expose transient overlap hidden by a
reallocator's implementation. `/usr/bin/time -v` records
process maximum RSS for each fresh process.

Each lane uses the same prepared source fixture and, where applicable, the
same prebuilt commit/patch object for all five samples. Transaction-lane
caller buffers are intentionally built inside their timed public transaction
scope; the per-row input-preparation field records this boundary. Production
copy traffic is not instrumented by this candidate. Receipts
therefore expose `copies_bytes_observed` only for explicit harness validation
copies performed after timing (source/candidate byte comparisons); they do not
pretend to count internal parser, archive, or allocator copies. Those copies
are excluded from the timed operation. This scope is repeated in the report.

Correctness gates run outside the timer and are required for every successful
changed lane:

* an empty transaction and its public apply retain exact package bytes;
* every changed result retains the unrelated opaque member byte-for-byte;
* patch inverse returns the complete original package bytes;
* semantic storage count, IDs, payload lengths, and reference counts match the
  lane recipe; and
* each sample retains source, candidate, and opaque-member hashes.

The harness also checks exact expected payload bytes and lengths for every
untouched, edited, inserted, renamed, and restored storage, and emits a
candidate package SHA-256 for every sample. The public inverse and opaque
member checks remain required independently of the payload assertion.
It checks the complete per-ID reference-count map, including zero-reference
storages and inserted/removed IDs. The `patchinverse` recipe validates and
hashes its prepared forward candidate outside the timed apply/inverse scope,
then independently checks that the timed inverse result restores the source.

The profile is intentionally separate from the user-owned invalid fixture
`crates/litchi-xlsb/tests/custom_data.rs`, which is not read, copied, or built
by the harness.
