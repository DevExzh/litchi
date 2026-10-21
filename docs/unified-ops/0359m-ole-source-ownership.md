# 0359m — W5 source ownership continuation

- Date: 2026-10-21.
- Base: `1c7c6d8f7b` (the frozen W5 COW/state-reuse candidate).
- Candidate source: `daa2d35860`.
- Status: source repair and focused validation are complete; native
  qualification remains pending a coordinator quiet window.

The cell-value owner previously kept complete artifacts as `Arc<[u8]>`, then
copied each committed snapshot into a `Vec<u8>` before reopening a source-backed
package.  Snapshot construction also copied the common object package's
already-owned source into a second slice allocation.  That shape obscured the
same immutable bytes from the source-backed CFB owner and added one
artifact-sized copy on the source-backed publication path.

The internal cell-value `Snapshot` and exact `Patch` endpoints now retain
`Arc<Vec<u8>>`.  `Snapshot::open_package` adopts the common object's shared
vector directly, and source-backed `package_from_snapshot` passes a cloned
handle through `SharedOleFile::open_owned_vec_with_limits`, retaining the
snapshot's source version and sealed immutable-vector path.  The common
overlay owner has a corresponding `open_owned_vec` entry point that uses the sealed
`SharedOleFile::open_owned_vec_with_limits` path, so source-backed numeric
plans keep their immutable-provenance fast path.  Public byte projections
remain borrowed slices; no new format-facing owner API is introduced.

The eager package path intentionally retains its one copy into
`Cursor<Vec<u8>>`; this repair does not weaken or redesign that reader.  The
positional `from_shared_ole_file` ingress still materializes its source once,
as required by its existing reader contract.  Exact patch comparison,
forward/inverse replay, source freshness fencing, candidate reopen and
admission remain unchanged.  A regression asserts that a committed snapshot
and its patch share both endpoint allocations, and the common overlay test
covers the new sealed-vector entry point's protection classification.

Validation at this candidate passed with one capped, offline Cargo process:

```
cargo test --offline -p litchi-ole-common -p litchi-xls --all-targets --all-features
cargo clippy --offline -p litchi-ole-common -p litchi-xls --all-targets --all-features -- -D warnings
```

Both commands used the external target
`/home/zhuhe/code/litchi-worktrees/.coord/w5-ole-source-arc-final-target`.
The frozen 0359j–0359l receipts remain historical evidence; this candidate
has no accepted performance result until the source-matched owner, cell,
history and inverse routes are rebuilt and requalified in a quiet window.
