# Validation and scope notes

The initial ZIP gate (`formal/checks/zip-initial.json`) failed to compile:
`PreparedSizedMember::write_local` moved its mutable writer reference before
using it again. The candidate reborrows that reference. The failed receipt,
log and source manifest remain, and `zip-initial-r1` passes 460 tests with
2 ignored. This is a compile correction before the candidate build/captures,
not a discarded workload measurement.

The changed fuzzer retains the original bounded archive admission and exercises
verified stored/deflated token publication, exact decoded readback, complete
accepted-byte accounting, short writes and partial payload failure. Six retained
deterministic seeds span empty, 17-byte and 128 KiB+17-byte payloads with UTF-8
names; the ASAN run passes 1,000 iterations with seed 456. The initial seed
manifest and the final mutated corpus are distinct artifacts; cleanup inventories
the latter rather than presenting it as the original corpus.

Independent production review found no material correctness issue in the direct
framing path. Shared stored payloads retain their Arc, and verified Store/Deflate
payloads retain the trusted token; ordinary owned Store/Deflate regeneration
keeps the existing buffered path. Local framing and central metadata share the
ordinary writer's grammar. Full layout preflight, ZIP64 offset promotion and
accepted-byte accounting precede or accompany output as before. Focused tests
cover pointer ownership, exact framing/output, partial failures and synthetic
ZIP64 without a multi-gigabyte allocation.

Managed OPC reservations remain conservative. The precompressed admission still
reserves capture plus a writer payload allowance; the existing source comment
at `source_backed.rs` describes that older complete-member allowance. This batch
does not lower it or claim that a lower physical allocation peak lowers the
admission budget. The 64 KiB stack copy buffer is unchanged.

The formal capture/derivation scripts were frozen before baseline capture.
Their inherited Markdown title says “transfer-chunk” and bootstrap seed uses
455 + candidate lane; neither selects or changes observations. The current
experiment is direct shared-payload framing, as declared in the frozen hypothesis.
All positive and negative absolute >5% flags remain in the derived report.

The first native check matched the 42,948-byte golden archive SHA, then failed
a relationship metadata digest assertion. A broad 0455-to-0456 text replacement
when copying evidence changed `045db62a04555453...` to `045db62a04565453...`
inside the expected digest. The immutable `native-expected.json`, original
driver/protocol, failed receipt and original output/report remain retained or
inventoried. The new `native-expected-r1.json` is an exact byte copy of the prior
expected file; `native-r1.py` uses fresh output paths and passes both providers.
This is an evidence transcription correction, not a production change or an
updated golden output.
