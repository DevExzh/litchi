# 0501 production candidate review

Independent read-only review of the applied `source_cross_copy.rs` candidate.
No build or test command was run in this review.

## Diff boundary

The production diff has exactly three semantic changes:

1. The image payload byte update in `digest_touched` was replaced by
   `check_payload_cancellation`.
2. The chart payload byte update in `digest_touched` was replaced by the same
   helper.
3. `check_payload_cancellation` was added. It visits the same 64 KiB chunk
   sequence as the old `digest_bytes` call and invokes `check_execution` for
   each chunk. It allocates nothing and does not consume `Resource::Work`.

The only other production change is a private test-module declaration. No
public API, manifest, source authority, topology, or publication code changed.

## Proof review

The candidate retains the exact payload identity chain:

- `PreparedImage` and `PreparedChart` equality still compares source URI,
  target URI, content type, declared size, and exact `PreparedPayload` bytes.
- `reuse_or_clone_payload` still calls checked exact byte comparison before
  reusing a prior payload.
- `Prepared::matches` still compares both complete image and chart vectors
  before checking `touched_digest`.
- `verify_candidate` still rereads selected image/chart Parts, checks decoded
  size and content type, rejects relationship-bearing leaves, compares every
  decoded byte exactly, and validates chart XML.
- `graph_digest` and its layout/master/theme content and relationship proof are
  unchanged. `digest_touched` still includes graph digest, source-layout URI,
  all XML byte inputs, relationship metadata/order, names, URIs, declared
  lengths, and payload lengths.
- Source lineage/revision checks, source freshness checks, source-monitored
  publication, compressed capture authorization, CRC/decoded equality, and
  partial-output handling are unchanged.

The two removed byte updates were therefore duplicate evidence for the exact
payload comparisons. A changed payload with an equal candidate digest remains
rejected by vector equality and candidate readback. The digest remains private
and cannot authorize publication by itself.

## Cancellation and budgets

The replacement helper preserves the old per-payload 64 KiB cancellation
observation cadence. Its behavior is intentionally limited to checking the
execution context: it does not hash, allocate, reserve memory or objects, or
charge work. This keeps managed cancellation and budget counters compatible
while removing the redundant SHA work. The helper also retains the old empty
payload behavior: the outer per-item check runs, and the chunk loop performs no
additional check for an empty slice.

The helper does iterate chunk descriptors for an unmanaged payload, where
`check_execution(None)` is a no-op. This has no correctness or resource effect;
the matched benchmark will determine whether a further no-context fast path is
warranted. Such a fast path would be a separate optimization and is not needed
for proof acceptance.

## Private tests reviewed

`payload_guard_tests.rs` adds two useful guards:

- It flips one byte at equal length in a private image or chart plan. The new
  metadata/graph digest remains equal, while `Prepared::matches` rejects the
  exact payload mismatch and publication emits no bytes.
- It flips one byte in a prepared candidate and verifies that candidate
  readback rejects it as stale.

These tests directly establish that the reduced digest is not an authorization
predicate for either payload family. Existing source-version, mid-publication
source-change, capture-reuse, graph, relationship, and raw-preservation tests
remain required gates.

## Verdict

No production proof blocker was found. The candidate is appropriately scoped
for the 0501 matched measurement. Acceptance still requires the normal focused
PPTX tests and the frozen before/after evidence gates, including cancellation,
source/cache/budget release, exact output, and every retained latency/RSS flag.
The result must remain scoped to the named source-backed PPTX workflow; the
historical broad 63.2--63.9% SHA hotness figure is not a causal speedup claim.
