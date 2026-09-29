# 0839 cached ordered-Part scheduling candidate

This is an archive-only current-base candidate for a fresh reconsideration of
the rejected 0787 cache-hit scheduling experiment. It makes no production
claim and contains no build, test, benchmark, profiler, or adoption result.
The 0787 rejection remains the decision until root-owned current measurements
pass a newly frozen memory and latency policy.

## Current base and archive

The current production files touched by the candidate have the same bytes as
the 0787 baseline. The observed hashes are:

| Path | Bytes | SHA-256 |
| --- | ---: | --- |
| `crates/litchi-opc/src/source_backed.rs` | 919,523 | `029dac7bc5dbe173a333bd35f316ec79286e8c15432f0766bc3bbc48ce0c1df3` |
| `crates/litchi-opc/src/source_backed/batch.rs` | 42,039 | `57461fc249e7f9b6729affecd50ebf2685ddea658f0b2b2dfd67085a0eae0491` |
| `crates/litchi-opc/tests/source_backed_batch.rs` | 46,597 | `b6309038bcc105373b80c7dfeee4da3e3b62dc2d21f6239552f9e5fdf53d0137` |

Accordingly, `candidate.patch` is the 0787 model patch carried forward as a
current-base unified diff. It changes only those three paths and ports the
same seven focused tests: cache-state eligibility, stale-hint eviction,
multiwave private reads, explicit caller-worker dispatch, two cached panic
cases, and source change after the helper fence.

## Candidate behavior

`PartCache::all_entries_ready` is a private, read-only scheduling hint. Under
the existing cache mutex it requires every prepared entry to be present in
`entries` and absent from both `flights` and `pending`. It does not advance the
LRU clock, counters, reservations, pins, or payload references. It does not
make the cache a second authority.

The hint is checked only after the existing prepared request, output
reservation, scheduler reservation, context fence, and `CpuTasks` charge. It
is reached only in the private-worker branch. A true hint uses
`read_cached_serial`: it still calls the ordinary `read_part_prepared` path for
every request, keeps the existing `wave_end` boundaries and fence before each
multiwave segment, catches request panics, drains the current wave, selects the
lowest ordinal error, and stops before a later wave after an error. A stale
hint therefore becomes an ordinary authoritative cache read, including a
cold load when an entry was evicted.

The explicit `ScopedWorkers` branch remains before the hint and keeps its
caller callback and per-wave task shape. The old private one-wave and reused
worker paths remain unchanged when the hint is false.

## Compatibility and refusal boundaries

The candidate retains request-count and declared-size preflight, source
identity/version fences, cancellation checks, output `Memory` and `Objects`
reservations, scheduler `Workers`, `IoConcurrency`, `Memory`, and `Objects`
reservations, and the cumulative `CpuTasks` charge. Reservations stay held
through the operation and are released by their existing owners. Ordered
output, duplicate requests, payload pinning, cache eviction, and typed source
and execution errors remain on the current paths.

The only deliberate reachability difference is internal: after a true
all-ready private hint, construction of private worker threads, channels, and
their associated control allocations is avoided. A failure that could have
occurred while constructing that private machinery is consequently not
reachable on that path. Admission-time refusals and all public or caller-owned
worker behavior remain unchanged. A cache race cannot turn the hint into a
stale result because every request still performs the authoritative prepared
read and its source/context checks.

## ADR and owner alignment

The change stays inside the existing `litchi-opc` cache and batch owner. It
adds no public API, archive dependency, executor, global state, retention
policy, derived memo, or physical package authority. This is consistent with
ADR 0005's immutable positional source, authoritative bounded cache,
explicit execution context, preservation default, and measured-evidence
contract. It is consistent with ADR 0031 because admission and accounting are
retained, the caller-supplied worker facility remains authoritative for its
route, and the hint does not bypass finite worker or I/O permits. It also
preserves the owner decision in 0758 that a candidate must prove bounded
resource behavior and reproducible output before adoption.

## Why this is not adoption evidence

Change 0788 did not explain the historical small-case RSS increase and did not
justify relaxing the 0787 guard. Changes 0789 and 0790 calibrated the
relationship between observed residency, exit resource accounting, and
launcher topology; they did not identify the candidate's cause or authorize a
threshold change. They make a new measurement protocol more precise, not the
old failure disappear.

Root must therefore apply this archive in an isolated worktree, run all
quality gates and the seven focused tests, and capture a newly frozen paired
matrix. The reconsideration must retain native exit RSS alongside phase
residency and allocation observations, qualify its launcher topology, verify
exact bytes/order/source work and every budget release, and report tails and
uncertainty. Any memory or fresh-path regression stays visible. Until those
gates pass, this patch is only a reviewable candidate and production remains
unchanged.

