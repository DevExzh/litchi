# 0554 proof note: incremental physical marker accounting

Status: source and semantics proof only. This note does not implement the
candidate, report a new profile, claim a speedup, or approve an optimization.
It is the bounded follow-up to
`docs/performance/results/change-0554/next-attribution.md`, whose selected
independent baseline row is the leaf
`OleFile<R>::validate_physical_sector_layout`.

The proof is against baseline revision
`53330ff6745dd7416f3bfb48ef891a020ee9d151`, with the baseline
`crates/litchi-cfb/src/file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`.
The frozen 0554 plan is
`docs/performance/results/change-0554/plan.json`, SHA-256
`d901c87af4e8d2585af997dbc850832db21db0ce0e359d3dcf7d90c7b415e98c`.
The preceding attribution note is SHA-256
`53305c66bd5a847cfc9625fd69e64baae804bc247bf5fa83637e9bb598a9c747`.
The 0554 profile driver SHA-256 is
`af789e447d0093d1e2ad6f2f01a44109ea6701735517c5d1cf2ed9754705f315` and
the immutable 0536 raw parser SHA-256 is
`5ce0d3a0c9f9246f207b9be063791cf6ccf014f6582da47562de42c7e0d364a`.
The baseline receipt SHA-256 is
`66991507f940ac9e0cbf90247292a64dcad473ac2e161b89fbf7f900b08d05d1`, and
the baseline source-manifest SHA-256 is
`c6ec30f2e7c8116c6e9bd77d1622295d8245a36261daf9fabae657be4fe9a8e9`.

The owner-profile evidence that motivated this proof is retained here so a
future experiment cannot silently change the attribution boundary. The owner
is `litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits`
for case `xls_owned_source_open_one_cell`; every dump is a positive timed
owner dump and uses Callgrind `Ir` only. The five raw dump hashes are:

| Dump | Owner summary `Ir` | Raw SHA-256 |
| --- | ---: | --- |
| `baseline/profile-r1-xls-owned.callgrind.1` | 2,263,310 | `e1062ccec1059976592fd9dd39d457f690db3f960d31820ea45bee7fce82b527` |
| `baseline/profile-r1-xls-owned.callgrind.2` | 2,263,116 | `af2418bc07bf8e8ed29bc2f6754657b8434b564d2345d2adc0b1efc5cc525bb1` |
| `baseline/profile-r1-xls-owned.callgrind.3` | 2,265,417 | `2b5d91bcbc4df58180b7cab76257ec5b2b2556dbfc3d9bf4539dadda63eb8b82` |
| `baseline/profile-r1-xls-owned.callgrind.4` | 2,263,697 | `63235eb34d53cf7aa5a31a300d98fe9eb4c8e667cd5faad2a4b73ca5c42fdbf8` |
| `baseline/profile-r1-xls-owned.callgrind.5` | 2,263,711 | `00ca0ae63cf6ee5927a5600b4ba10421e51c2bd392a47445c1dbc700d5734a8e` |

The selected owner median is 2,263,697 `Ir`; the physical reconciliation
leaf is 398,342 self `Ir` in every dump. These values remain attribution
evidence only. Inclusive descendant rows are not added to this self row, and
no missing-function or zero-call result is inferred.

## Actual state and validation order

Let `P = sector_roles.len()` be the number of physical sectors, and let
`F = fat.len()` be the number of entries materialized from the complete FAT
sectors. The baseline allocates one `PhysicalSectorRole` for every physical
sector. The FAT can contain more entries than `P` because each FAT sector is
decoded at full width; those entries are padding for this reconciliation.

The baseline call path is

```text
OleFile::open_with_limits
  -> load_fat
  -> load_directory
  -> load_minifat              [when the header count is nonzero]
  -> validate_stream_allocations
  -> validate_physical_sector_layout
```

The relevant baseline source paths are:

| Path | Meaning |
| --- | --- |
| `crates/litchi-cfb/src/file.rs:571-712` | ingress, header, count, and physical-role setup |
| `crates/litchi-cfb/src/file.rs:769-922` | FAT/DIFAT claims, FAT materialization, and FAT/DIFAT marker checks |
| `crates/litchi-cfb/src/file.rs:966-1054` | directory chain claim, reads, and parse/tree validation |
| `crates/litchi-cfb/src/file.rs:924-963` | MiniFAT chain claim and materialization |
| `crates/litchi-cfb/src/file.rs:1057-1076` | checked role claims and chain claims |
| `crates/litchi-cfb/src/file.rs:1079-1164` | root, mini-stream, and regular-stream allocation validation |
| `crates/litchi-cfb/src/file.rs:1178-1191` | final physical reconciliation |

The observable error precedence is therefore:

1. Source size, header, version/shift, start/count, and physical-sector
   checks run before the role vector is installed.
2. `load_fat` validates declared counts, header FAT locations, DIFAT chain
   locations and continuation, FAT-sector count, reads, and allocations. It
   then checks every FAT sector for `FATSECT` and every DIFAT sector for
   `DIFSECT`. These errors precede all directory and final-layout errors.
3. `load_directory` collects and checks its chain, claims its physical
   sectors, reads the directory, and performs directory/name/tree validation.
4. `load_minifat`, when present, collects and checks its chain, claims its
   physical sectors, reads it, and materializes its table.
5. `validate_stream_allocations` checks the root mini-stream chain, claims it
   as `MiniStream`, checks each mini stream against the MiniFAT and root
   capacity, and checks/claims each regular stream as `RegularStream`.
6. Only after all preceding phases succeed does
   `validate_physical_sector_layout` inspect physical sectors in ascending
   order. For sector `i`, it first reports a missing FAT entry if `i >= F`.
   If the entry exists, it reports an unclaimed sector only when its marker is
   not `FREESECT`. Claimed sectors tolerate any marker.

`claim_sector` has its own precedence inside phases. Conversion and bounds
errors occur first; a role conflict occurs next; only a successful
`Unclaimed -> role` assignment changes ownership. Marker accounting must run
after that assignment, so a failed claim cannot retire a pending marker or
replace the existing bounds/conflict error.

There are two role-claim times that matter for a candidate. `load_fat` claims
header and DIFAT-discovered FAT sectors as `Fat`, and DIFAT sectors as
`Difat`, before `self.fat` is installed. Directory and MiniFAT chains claim
`Directory` and `MiniFat`; the root chain claims `MiniStream`; regular stream
chains claim `RegularStream`. Mini-sector ownership in
`claimed_mini_sectors` is indexed by the MiniFAT and does not claim a
physical-sector role. A candidate that only updates accounting after FAT
installation must explicitly account for the earlier FAT/DIFAT claims.

## Exact final semantics

The final loop is equivalent to considering the physical prefix
`0..min(P, F)`, plus a virtual missing-entry position at `F` when `F < P`.
Define

```text
U = { i | 0 <= i < min(P, F),
             sector_roles[i] == Unclaimed,
             fat[i] != FREESECT }
```

Let `m = min(U)` when `U` is nonempty. The original result is exactly:

```text
if F < P and (m is absent or F < m):
    CorruptedFile("FAT does not contain an entry for physical sector F")
else if m exists:
    CorruptedFile("unclaimed physical sector m has FAT marker 0x{fat[m]:08X}")
else:
    Ok
```

The marker branch wins whenever `m < F`; the missing-entry branch wins at the
first missing index when no earlier marker exists. The text must use the
actual sector and `fat[m]` value. Entries at `i >= P` are padding and are not
part of `U`, even when they contain `ENDOFCHAIN` or another non-`FREESECT`
marker. This is the source of the accepted behavior for short physical final
sectors and producer-filled FAT padding.

## Bounded proof candidate

A concrete candidate can carry a bounded pending-marker bitset `M` alongside
the existing role vector. Its bit length is
`min(P, fat_entry_count)`, where `fat_entry_count` is already computed in
`load_fat`; no padding index receives a bit.

While the existing FAT loop decodes each entry, it can set `M[i]` exactly
when `i < P`, the decoded marker is not `FREESECT`, and the role already
stored at `sector_roles[i]` is `Unclaimed`. At this point the only roles that
can already exist are the FAT/DIFAT claims made earlier in `load_fat`, so this
single materialization pass accounts for those pre-FAT claims without a
second physical scan. The existing FATSECT and DIFSECT checks must remain in
their current place and must still return before directory work.

After `self.fat` is installed, every successful `claim_sector(i, role)` clears
`M[i]` if the bit exists. It does so only after the existing conversion,
bounds, conflict, and role assignment steps succeed. Claims with `i >= F`
need no bit operation; normal chain validation already prevents later chain
walks from using a missing FAT entry, and the final missing position remains
`F`.

At the original final validation call site, the candidate returns the first
set bit in `M` if one exists, otherwise it returns the missing-entry error at
`F` when `F < P`, otherwise `Ok`. To preserve the original ascending order,
the implementation must compare the first set bit with `F` using the rule
above. A bitset can keep a monotonic first-set cursor: after clearing the
current first bit, search forward for the next set bit. The cursor never moves
backward because roles only transition out of `Unclaimed`. A word-level search
or an equivalent ordered pending set is acceptable only if it retains the
same first-sector and marker information.

The invariant after FAT materialization and after every successful later claim
is:

```text
M[i] == 1
  iff 0 <= i < min(P, F)
      and fat[i] != FREESECT
      and sector_roles[i] == Unclaimed
```

Initialization establishes the invariant because each physical-prefix FAT
entry is tested once while it is decoded, and pre-existing FAT/DIFAT roles
are read from `sector_roles` before setting the bit. For preservation, a
successful claim changes exactly one role from `Unclaimed` to a concrete role
and clears exactly that bit. A failed claim changes neither role nor bit.
No code in the inspected flow changes a concrete role back to `Unclaimed`.
Thus the first set bit is exactly `m` in the original final loop, and the
virtual position `F` reproduces the original missing-entry precedence. The
final logical check can remain at the same call site, so the candidate does
not emit a physical-layout error earlier than baseline.

The proof does not make the additional state free. A fallible bitset
reservation could itself change allocation-failure precedence by returning
before the existing FAT/DIFAT marker checks. A future implementation must
either preserve the established allocation/resource contract, fold the
bounded state into an already-covered reservation, or reject this design on
that basis. Marker bookkeeping must not be allowed to turn an allocation
failure into an earlier logical corruption error, and it must not move a
FATSECT/DIFSECT, directory, MiniFAT, stream-chain, bounds, or conflict error.
The bitset also adds a branch/write during FAT materialization and a clear on
each successful physical claim; the net cost is unknown until a fresh
candidate is measured.

## Counterexamples that constrain the implementation

These small states are enough to reject several tempting but incorrect
variants:

| State or action | Required baseline result | Variant disproved |
| --- | --- | --- |
| `P=4`, `F=4`, markers `[FREESECT, 0x10, FREESECT, 0x20]`, all roles initially unclaimed; claim sector 1 | marker error at sector 3 with marker `0x20` | A pending count alone cannot identify the remaining sector or marker. |
| `P=4`, `F=2`, marker `fat[0]=0x10`, roles 0 and 1 unclaimed | marker error at sector 0 | Checking `F < P` before the prefix can incorrectly return the missing error at sector 2. |
| `P=2`, `F=2`, sector 1 has a non-`FREESECT` marker; a directory claim succeeds before a later directory parse error | the later directory error wins; physical marker reconciliation is never reached | Emitting a marker error from `claim_sector` shifts the error earlier and hides the directory error. |
| FAT/DIFAT sectors are successfully claimed before `self.fat` is installed | existing FAT/DIFAT marker checks retain their current place and text | Reading `self.fat` from `claim_sector` is unavailable; a design that forgets these preclaims leaves false pending markers. |
| `P=2`, `F=4`, `fat[3]=ENDOFCHAIN` | padding at index 3 is ignored; physical reconciliation can succeed | Tracking all FAT entries or checking `i >= P` rejects tolerated padding. |
| `P=5`, `F=3`, non-free unclaimed marker at index 1 | marker error at 1, before missing entry 3 | A design that reports the missing suffix whenever `F < P` changes ascending precedence. |
| `P=5`, `F=3`, no non-free unclaimed marker below 3 | missing FAT entry error at 3 | A design that reports success from an empty marker set forgets the virtual missing position. |

The counterexamples support the bitset invariant and rule out a marker count
without identity, an early claim-time check, and a padding-wide map. They do
not show that the bitset candidate is faster. The candidate must remain
separate from the rejected 0534 paired-prefix loop rewrite and from the
rejected 0548/0549 collector variants; none of those changes is revived by
this proof.

## Bounded next experiment

If source review accepts the allocation and state-lifetime contract, the next
experiment should implement only the bitset/cursor accounting described here,
leaving chain collection, directory/name parsing, and public APIs unchanged.
Before native timing, it should exercise the listed short-FAT, padding,
claim-conflict, pre-FAT FAT/DIFAT, all-free, and unclaimed-marker fixtures and
compare the complete `OleError` text and phase precedence with baseline. It
then needs the existing 0554 owner-boundary, correctness, allocation/RSS,
malformed-input, and strict XLS gates. A failed allocation-contract review or
an accounting cost that erases the 398,342-`Ir` attribution closes this
candidate without an adoption claim.

This note contains no Cargo/build/capture activity and no live Rust change.
