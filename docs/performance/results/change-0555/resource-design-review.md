# Change 0555 physical marker accounting review

**Disposition:** conditional source-review pass for the allocation-free role
encoding; the prepared source is not yet ready to freeze.

**Performance claim:** none.

This is an independent, read-only review of the physical marker accounting
hypothesis in `change-0554/proof-note.md`. I did not edit live Rust, run Cargo,
run tests, or run captures. The live CFB source remains
`crates/litchi-cfb/src/file.rs` at SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`.
The reviewed proof note is SHA-256
`80fe4343a9ea99bdd3d3d1b287ebd1d0cb9befeee5f90a591a8481a4dd67baa6` and its
selected baseline is revision `53330ff6745dd7416f3bfb48ef891a020ee9d151`.

## ADR and ownership audit

The 0555 ADR manifest binds the ADR index and all 29 decision files. Their
recorded hashes match the current files. The records are accepted, with ADR
0014 explicitly amended by ADR 0015; ADR 0025 uses the repository's lowercase
`accepted` spelling. The physical candidate is confined to the private CFB
parser, which is the owner established by ADRs 0002, 0024, and 0026. The
proposal therefore does not add a public API, dependency, archive shortcut, or
format-layer ownership edge.

The constraints that govern this review are:

* ADR 0001 and ADR 0004 require the ordinary API to remain typed and free of
  physical implementation details. The proposed state is private and does not
  leak a sector identity.
* ADR 0003 requires immutable publication boundaries. Open-time accounting may
  mutate the private staged parser, but no partially validated object may be
  published.
* ADR 0005 requires finite resource use and fallible reservations, and permits
  a performance claim only after representative measurements. An additional
  pending bitset reservation would create a new allocation/failure boundary;
  this review evaluates the existing-role encoding specifically because it can
  avoid that reservation.
* ADR 0006 requires lossless, deterministic validation and established error
  precedence. A marker must not be reported from `claim_sector`, and a failed
  claim must not alter ownership or pending state.
* ADR 0008 requires an independently reproducible candidate and all applicable
  correctness, malformed-input, allocation, native, and quality gates before
  any support or optimization claim. ADRs 0009, 0011--0023, and 0025, 0027--
  0029 do not create a new owner or exception for this CFB-only experiment.

## Baseline state and proof boundary

At the reviewed baseline, `OleFile::open_with_limits` installs one
`PhysicalSectorRole` per physical sector before `load_fat`. `load_fat` claims
header and DIFAT-discovered FAT sectors before `self.fat` exists, materializes
the complete FAT including tolerated padding, checks `FATSECT`/`DIFSECT`, and
only then installs `self.fat`. Directory, MiniFAT, root, and regular-stream
claims all go through `claim_sector`. The final
`validate_physical_sector_layout` pass walks physical sectors in ascending
order and implements the following precedence:

1. a non-free unclaimed marker at the first physical prefix position wins;
2. otherwise a missing FAT entry at `F` wins when `F < P`;
3. padding at `i >= P` is ignored;
4. claimed sectors tolerate any marker.

Here `P = sector_roles.len()` and `F = fat.len()`. The exact marker value and
sector number in the existing corruption text are part of the contract. All
earlier header, FAT/DIFAT, directory, MiniFAT, stream-chain, bounds, and claim
conflict errors must continue to win before this final check.

## Feasibility of role encoding plus a monotonic cursor

Encoding a boolean pending condition as a private additional
`PhysicalSectorRole` variant is sound and allocation-free. The variant must
mean “logically unclaimed, with a non-`FREESECT` marker,” not a new ownership
role. The existing `fat` vector remains the sole source of the marker value.
Adding one enum variant does not increase the size of the current small enum,
and a single `Option<usize>` cursor is bounded scalar state. This avoids the
new `try_reserve` and allocation-error precedence problem of the proof note's
separate bitset.

The cursor is also sound, but only if it is maintained eagerly at every state
transition. Initialize it while the already-reserved FAT vector is being
materialized, in ascending entry order, after FAT/DIFAT claims have been made.
For each entry `i`, set `Pending` only when `i < min(P, F)`, the marker is not
`FREESECT`, and the current role is `Unclaimed`. A preclaimed FAT or DIFAT
sector must remain concrete so its existing marker validation retains its
place. FAT padding must never acquire `Pending`.

After the existing conversion, bounds, and conflict checks succeed,
`claim_sector(i, role)` may consume either `Unclaimed` or `Pending`. A
successful `Pending -> role` transition clears the pending condition. If and
only if `i` is the current cursor, advance the cursor forward through the role
vector to the next `Pending` entry. Each role slot is examined at most once by
these advances; claims at other positions do not scan. The cursor must never
move backward, and no path may change a concrete role back to `Unclaimed` or
`Pending`.

The final validator can then remain `&self` and read only the cursor and FAT:

```text
if first_pending exists at m:
    report the existing marker error using fat[m]
else if F < P:
    report the existing missing-entry error at F
else:
    succeed
```

The implementation should still compare `m` with `F` defensively, or retain a
debug invariant that proves `m < min(P, F)`. This protects the padding and
short-FAT rules if a future caller constructs private test state incorrectly.
The final validator must not lazily advance the cursor through interior
mutability: validation is read-only and must not become a mutation phase.

The proof is by induction. FAT materialization establishes

```text
role[i] == Pending
  iff 0 <= i < min(P, F)
      and fat[i] != FREESECT
      and role[i] was Unclaimed after FAT/DIFAT claims.
```

Each successful later claim changes exactly one pending role to a concrete
role and advances to the next pending slot when necessary. A failed claim
changes neither role nor cursor. Therefore the cursor identifies exactly the
first `m` that the original final paired loop would report, while `F` remains
the virtual missing-entry position. The proof does not depend on chain order;
claims may arrive in any order.

## Current prepared source finding

The current isolated preparation adds `PhysicalSectorRole::Pending` and folds
marker classification into the FAT materialization loop, which is the correct
place to account for pre-`self.fat` FAT/DIFAT claims without a second FAT scan.
It also makes `Pending` claimable while preserving the earlier error checks.
Those parts are source-compatible with the proof.

However, the prepared source currently declares `first_pending_sector` but
does not initialize or advance it. Its final validator still calls
`sector_roles.iter().position(...)`, which is a full physical role scan. That
is a different, role-only final-loop hypothesis; it is not the promised
monotonic-cursor implementation. The unused field also cannot survive the
warning-denied quality gate. The direct synthetic `OleFile` test fixtures must
likewise initialize a coherent cursor if the cursor implementation is kept.

The candidate owner should choose one of these explicit paths before the plan
is frozen:

1. complete the cursor implementation above, add focused invariant and
   precedence tests, and measure that candidate; or
2. remove the unused cursor field and document/measure the role-only scan as a
   separate candidate, with an explicit net-work gate for the extra
   per-FAT-entry role classification.

The role-only scan is semantically plausible when `Pending` is initialized
correctly, but it does not remove total role traversal: it adds a role access
and branch during each FAT entry and still scans roles at the final call site.
No source review can infer that it beats the baseline. It must not be described
as the cursor design or as eliminating the final scan.

## Required correctness and resource checks

Before native timing, the candidate must compare complete `OleError` variants,
text, and phase precedence for at least these cases:

* all-free and one or multiple unclaimed non-free markers, including markers
  before and after claimed sectors;
* `P > F` with a marker below `F`, which must report the marker before the
  missing-entry error, and `P > F` with no pending marker, which must report
  the missing entry at `F`;
* `F > P` with arbitrary non-free FAT padding, which must succeed when all
  physical roles are valid;
* FAT and DIFAT sectors claimed before FAT installation, including their
  existing `FATSECT`/`DIFSECT` validation errors;
* directory, MiniFAT, root, regular-stream, bounds, and duplicate-claim
  failures, proving that no marker error moved earlier;
* a claim of the current first pending sector, a claim of a later pending
  sector, a failed conversion/bounds/conflict claim, and an out-of-order chain;
* direct test fixtures with a coherent pending-role/cursor state, while
  preserving the existing public scalar, root/classic-Mac, and source
  publication tests.

The allocation-free design must retain the baseline `physical sector roles`
reservation and all existing FAT/chain/resource labels and ordering. It must
not add a bitset, map, ordered set, marker vector, or hidden lazy allocation.
The exact-reserved FAT push path must remain within its existing reservation;
the new role branch must not introduce a fallible error after the FAT
reservation succeeds. The cursor's additional branch and any forward scans
must be covered by the allocator/RSS lane and by the existing malformed-input
limits. No allocation, latency, RSS, or instruction result is established by
this review.

If the cursor version passes those focused checks, a full matched OLE2/XLS
campaign is justified under the frozen 0554 admission contract, including CFB
shape controls, native owner latency, allocation/RSS, malformed-input,
correctness, profile/instruction attribution, and final quality. If it fails
the resource/error contract or the bookkeeping cost does not clear the
measured gate, close the candidate without an adoption claim. The rejected
0553/0534 paired-prefix and 0548/0549 collector variants remain excluded.

ODF remains deferred until the OLE2/OOXML optimization goal completes.
