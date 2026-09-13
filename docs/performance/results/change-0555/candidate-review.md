# Change 0555 physical marker candidate review

**Disposition:** source review pass for the isolated candidate; eligible for
the coordinator's frozen measurement and correctness gates.

**Performance claim:** none.

This review is bound to the isolated witness against the restored 0554
baseline; it does not describe the coordinator's later candidate application.
I did not edit live Rust, run Cargo, run tests, or run captures. The historical
baseline input was revision
`d3c62f19a7632114f9cce5d1ad350e237215582d`, with
`crates/litchi-cfb/src/file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`.

The final candidate witness, formatted patch, and candidate source review
are:

| Artifact | SHA-256 |
| --- | --- |
| `candidate-preparation/file.rs` | `ae1383c3277d8695a5ff1495200079bd1a8f2156676d44058e1071a2de9fe173` |
| `candidate-preparation/candidate.patch` | `709654664c2a57c25d031278871370d6c5960572a1a12558dec3ada6c45432a0` |
| candidate's isolated source review | `4667a7de7ffa13917d7350144d6f9bb7de4b5487e0cd20bd654209fe2966bad1` |

The patch is confined to `crates/litchi-cfb/src/file.rs`; it introduces no
public API, dependency, unsafe code, persistent metadata, or archive owner
change.

## Qualification correction

The retained `qualification-01` witness and its review remain immutable. Its
targeted-01 quality result (`0ab1ad9c659d35282a24c99c6f7dc8f101e28c712f30eb9073c181e69351c377`)
rejected only the qualified `std::mem::size_of` spelling in the added
capacity test. The corrected source differs from that witness by exactly one
test line: it uses the prelude `size_of` spelling. The corrected patch carries
the corresponding test-only change; regenerated patch metadata and hunk
context do not change production code.

`candidate-correction.json` binds the original source and corrected source and
patch hashes. Its stated scope is test-only qualification; the production
candidate implementation is identical, and no candidate captures existed.

## Mechanism review

The candidate adds a private `PhysicalSectorRole::Pending` state. It records a
non-`FREESECT` marker in that existing role vector while decoding each entry of
the already exactly reserved FAT vector. It limits the state to physical
indices, so FAT padding is ignored. Header and DIFAT FAT-sector claims already
exist before this decode and remain concrete, so their existing `FATSECT` and
`DIFSECT` checks retain their original phase.

Later `claim_sector` calls accept `Unclaimed` and `Pending`, then replace
`Pending` with the concrete physical role after conversion, bounds, and
conflict checks. No claim-time marker error is emitted. The final validator
searches the role vector in ascending order and reads the FAT only for the
first pending marker. If none exists, it reports the missing FAT entry at
`fat.len()` when that is below the physical-role length. This is the
role-only-final-scan hypothesis selected for this campaign; it does not claim
an O(1) final check or use the unused cursor idea from the initial design
review.

The allocation boundary is compatible with the proof. `fat` retains its exact
fallible reservation and the replacement `push` calls do not require another
reservation. The pending state uses no bitset, map, marker vector, or hidden
heap allocation. The extra role comparison/write during FAT decoding and the
extra claim branch remain unmeasured work and must be included in the native,
allocator, RSS, and owner attribution gates.

## Semantic and phase findings

The role invariant is valid for production construction:

```text
Pending[i] iff i < min(P, F), fat[i] != FREESECT,
             and the physical role is still logically unclaimed.
```

FAT/DIFAT claims establish the pre-existing concrete exceptions. All later
physical claims route through `claim_sector`, and chain collection checks the
FAT table before a later claim, so no production path can create a pending
role at or beyond a missing-FAT position. Consequently the first pending role
is the first marker that the baseline paired loop would report. A marker below
`F` still wins over the virtual missing position; no marker in FAT padding is
considered; and claimed non-free sectors are accepted. Earlier header,
FAT/DIFAT, directory, MiniFAT, stream-chain, bounds, and conflict errors remain
before the final reconciliation call site.

The special handling for an incoming private `Unclaimed` role on a `Pending`
slot preserves the marker fact. Production never requests `Unclaimed` as a
concrete owner; the branch is relevant only to the existing crate-private role
tests. It must not become a `Pending -> Unclaimed` transition, which would lose
the marker invariant.

The `.position()` scan relies on the reachable production invariant that every
`Pending` entry was created while decoding a physical-prefix FAT entry, so it
is below `F`. The final source's fixtures establish that state explicitly. A
manually fabricated `Pending` beyond the FAT is outside that production state
machine and is not a required admission case.

The bound source includes the clone snapshot in
`pending_claim_is_consumed_without_shifting_conflict_errors` and the corrected
`actual.err()` comparison in the differential test. Those tests exercise
pending creation through `record_decoded_fat_entry`, short-FAT precedence,
padding tolerance, out-of-order claims, pending consumption, and failed-claim
nonmutation. They still require execution by the coordinator, along with the
existing scalar, classic-Mac root two-view, name/cache, and atomic-publication
tests.

## Decision and measurement boundary

The candidate is eligible for the frozen 0555 CFB/XLS campaign. The role-only
scan must be evaluated as a net-work hypothesis: the FAT decode now adds a
per-entry role branch while the final loop removes per-sector FAT lookups. The
campaign must retain all positive owner dumps and enforce every primary XLS,
CFB control, allocator, RSS, correctness, malformed-input, and quality gate. A
source review or a lower instruction count alone cannot justify adoption. Any
failed mandatory gate requires rejection and baseline restoration.

The 0553/0534 paired-prefix rewrite and 0548/0549 collector variants remain
excluded. ODF remains deferred until the OLE2/OOXML optimization goal
completes.
