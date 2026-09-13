# 0555 physical marker candidate source review

Status: isolated candidate witness only. This file and the adjacent
`file.rs` are not production source and have not been built, benchmarked, or
captured. The live checkout remains at the restored 0554 baseline.

## Frozen inputs

| Item | SHA-256 |
| --- | --- |
| baseline revision | `d3c62f19a7632114f9cce5d1ad350e237215582d` |
| baseline `crates/litchi-cfb/src/file.rs` | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| 0554 proof note | `80fe4343a9ea99bdd3d3d1b287ebd1d0cb9befeee5f90a591a8481a4dd67baa6` |
| 0554 attribution note | `53305c66bd5a847cfc9625fd69e64baae804bc247bf5fa83637e9bb598a9c747` |
| 0554 plan | `d901c87af4e8d2585af997dbc850832db21db0ce0e359d3dcf7d90c7b415e98c` |
| candidate `file.rs` | `9860e87ab34ce88c7e05d2914248a5d51acf1b4f34f1cc42406dacafd1836afa` |
| candidate patch | `5b94327fbecd1902ecf50ff15c6419885d6388d03cd5a69bc535e0dcbdbd9b25` |

The baseline revision is the restored 0554 production revision. The source
hash in the second row is the source copied before this candidate was edited.
The patch is a binary-capable unified diff from that source to this witness.

## Candidate mechanism

The existing `PhysicalSectorRole::Unclaimed` state is extended with the
private `Pending` state. During the existing FAT-sector decode, each decoded
entry is pushed into the already exactly reserved `fat` vector. For entries in
the physical prefix, a non-`FREESECT` marker changes an as-yet unclaimed role
to `Pending`. FAT padding is excluded, and roles already claimed as `Fat` or
`Difat` remain concrete before the decode starts.

Later validated ownership claims accept `Pending` as an unclaimed slot and
replace it with the concrete role. A request to assign the private
`Unclaimed` role does not erase `Pending`; this preserves the existing role
helper's semantics and keeps the marker fact intact. Failed conversion,
bounds, and conflict checks occur before any role change.

The final reconciliation uses one ascending `sector_roles` scan to find the
first `Pending` sector. It reads one corresponding FAT entry only when
forming the required marker error. If there is no pending sector, it reports
the first missing FAT entry when `fat.len() < sector_roles.len()`. Thus the
candidate removes the baseline's paired role/FAT read on every physical
sector while retaining an O(P) ordered role scan. It does not claim an O(1)
final check and does not add a heap allocation, cursor field, or second
physical index.

## Invariant and error order

After FAT decoding and after every successful concrete role claim, for every
physical index `i` below the shorter of the role and FAT vectors:

```text
sector_roles[i] == Pending
  iff sector_roles[i] was unclaimed after all claims made so far
      and fat[i] != FREESECT
```

Indices beyond the physical role vector are padding and never become
`Pending`. FAT/DIFAT claims happen before FAT entries are decoded, so their
markers do not become pending. Directory, MiniFAT, root, and stream claims
consume pending states only after their existing chain, conversion, bounds,
and conflict checks succeed. No production path changes a concrete role back
to `Unclaimed`.

The first pending index is therefore the baseline's first unclaimed
non-free-marker index. If it exists, its marker error precedes the virtual
missing entry at `fat.len()`, matching ascending baseline order. If no pending
index exists and the FAT is shorter than the physical role vector, the missing
entry error is emitted at exactly `fat.len()`. FAT/DIFAT marker checks remain
immediately after FAT materialization, before directory loading; all earlier
header, table, chain, directory, and claim errors retain their original
order. Reconciliation errors remain deferred until the original final call
site.

The direct tests retain exact marker text, short-FAT precedence, padding
tolerance, role ownership, conflict nonmutation, and pending consumption. The
candidate must still be reviewed and then measured by the root coordinator;
this source review supplies no correctness, allocation, latency, or native
admission result.
