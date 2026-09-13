# Change-0554 OLE2 directory-name handoff review

`Disposition: source review pass; capture and admission hold`

`performance_claim: none`

This is a read-only source review of the isolated OLE2 N candidate. I inspected
the live CFB parser, the 0553 attribution plan and source note, the original
candidate, the formatted candidate preparation, the accepted ADRs 0001, 0002,
0003, 0005, 0006, 0024, and 0026, and the MS-CFB directory-entry rules. I did
not edit live Rust, run Cargo, run tests, or change any earlier evidence bundle.

The current live revision is `53330ff6745dd7416f3bfb48ef891a020ee9d151` and
the live `crates/litchi-cfb/src/file.rs` hash is
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
original isolated candidate remains immutable at
`docs/performance/results/change-0553/ole2-name-candidate/`:

| Artifact | SHA-256 |
| --- | --- |
| original `file.rs` | `e7d6b72e2389b4a47d4628e04f1337a143a06abeec3dcd0f108d42494fc1b6c9` |
| original `candidate.patch` | `af6639d1ff99f9303031c4b51e26e3aa19a5154ad93314e6c949193ad067fb91` |
| prepared `candidate-preparation/file.rs` | `a96667162dbd8c3d7ebe88cad3bb0f07670296ebf2fd16279b7971cc375bbdd9` |
| prepared `candidate-preparation/candidate.patch` | `33f815a7fee8dac1aa13bd461d54235c5153a34b6e66044540eda4c6e47bc7ed` |

The prepared patch is a formatted copy of the same N mechanism plus the
focused tests described below. `git apply --check` passed against the current
live source. The prepared source is not a live source replacement.

## Correctness findings

The candidate preserves the validation authority. `validated_directory_entries`
still parses every 128-byte record, validates the name length, UTF-16,
terminator, forbidden-name rules, scalar structural fields, and the SID-owned
sibling trees before graph construction. The only representation change is
`ValidatedDirectoryEntry.name: String` to `Option<String>`.

For every non-root graph node, the builder computes the checked SID index,
requires the corresponding validated entry and matching SID, and passes its
name slot to `parse_directory_entry`. After the existing raw-record and
name-length checks, `Option::take()` moves the validated `String` into the
public `DirectoryEntry`. The visited set and validation traversal still reject
repeated SIDs, cycles, cross-storage ownership, invalid SIDs, and ordering
violations before another handoff can occur. No name clone or second normal
name allocation is introduced.

The final cache loop remains a useful proof boundary. A normal public entry is
accepted only when its SID equals the validated slot and that slot is `None`,
which proves that the exact validated name was consumed. Empty slots still
pair only with empty public slots. The root is the explicit exception: its
validated name may remain present because the public root parse deliberately
keeps the historical root view. No externally reachable path can take a
normal name twice; a second take would produce the new internal
`validated CFB directory SID {sid} name was already handed off` corruption
error, but the unchanged visited/SID checks make that state unreachable for a
validated source.

Scalar extraction is unchanged. The raw parser still supplies `entry_type`,
both sibling SIDs, child SID, CLSID text, start sector, version-3-masked size,
MiniFAT classification, and an empty child vector. The prepared focused test
compares every public scalar to the same raw directory record and checks the
SID-aligned comparison data, rather than checking only the transferred name.

Malformed-source precedence is preserved for the ordinary paths. Validation
still runs before public directory allocation and graph construction. The
public parser retains its name-length check before taking a name, so the
existing malformed fixture continues to report exactly:

```text
invalid CFB directory name length 3 at SID 1
```

Invalid object kinds, invalid UTF-16, embedded or missing terminators,
forbidden names, invalid colors, structural field errors, and tree errors are
still produced by the earlier validation pass with their existing variants and
ordering. The candidate changes the allocation path for a normal name:
`decode_utf16le` is no longer called for that entry, so its fallible
`"decoded CFB directory name"` reserve is no longer exercised there. This is
an intentional ownership-transfer consequence, not a malformed-input
diagnostic change, but the allocation-failure lane must explicitly record and
accept this resource-boundary difference before production admission.

## Classic-Mac root two-view behavior

The candidate keeps root parsing separate. Validation continues to accept only
the already-supported exact SID-0 `name_len == 2`, bytes `00 52`, zero-padded
classic-Mac spelling and canonicalizes it to `Root Entry` for ordering and
lookup. The public parser receives no validated name and decodes the raw
record's historical view, which is the empty string for this encoding. The
final cache stores the canonical `Root Entry` comparison data, so ordinary
child lookup remains valid. Standard UTF-16 `Root Entry` records continue to
produce the same public and comparison views as before. The prepared test
exercises both the public empty root view and successful lookup through the
canonical cache.

No root name is transferred into the public entry, so the root retains one
bounded duplicate `String` during the directory phase by design. This is the
required two-view compatibility exception and is independent of normal-entry
name handoff.

## Memory and publication review

The candidate adds no vector, map, scratch buffer, or unbounded retained
structure. `Option<String>` has the same ownership role as the previous
validated string slot; normal names move from the validated vector to the
public vector and are no longer dropped twice. `DirectoryNameData`, the
validated structural records, directory bytes, public graph, comparison cache,
and all existing finite reservations remain in the same phases. The root
exception is constant-sized. These are source-level allocation-shape findings,
not RSS or latency results; the planned allocation and native capture is still
required.

`load_directory` stages `dir_data`, validated entries, the public graph, and
comparison data in locals. `self.root`, `self.dir_entries`, and
`self.dir_name_data` are assigned only after the complete SID-aligned loop
succeeds. A validation, graph, cache-alignment, I/O, or fallible queue failure
drops the staged values without publishing them. The prepared candidate adds
`failed_directory_load_does_not_publish_a_partial_graph`, which reuses a
populated receiver, injects the exact malformed SID-1 name length, and checks
the old root, public entries, and comparison cache remain unchanged.

## Required follow-up before admission

There is no source-level correctness blocker to a controlled candidate build.
The candidate is still unbuilt, untested, unprofiled, and inadmissible. Before
any adoption decision, the owner should:

1. Run the prepared focused tests and the applicable CFB, DOC, XLS, and PPT
   integration and public-boundary tests against the exact prepared manifest.
2. Run differential malformed-input checks for invalid length, terminator,
   UTF-16, object type, color, SID/tree, and storage/stream fields, preserving
   error variant, exact text where already asserted, and precedence.
3. Run the dedicated classic-Mac root fixture, standard root fixture, deep and
   wide sibling trees, and repeated/cross-storage SID fixtures. Keep the
   no-partial-publication check in the candidate receipt.
4. Exercise the configured allocation-failure/resource-limit lane. In
   particular, document whether removing the duplicate decoder reserve is
   accepted under the repository's typed allocation-error contract; do not
   infer parity from valid-input tests.
5. Run the fresh CFB many-small, few-large, tiny, and actual XLS owner
   attribution and native/allocation/RSS campaign from the 0553 OLE2 plan.
   Any adoption or speed claim requires that matched evidence and the existing
   decision gates. This review supplies no performance result.

The 0553 N source note and attribution plan remain the authority for campaign
scope. ODF remains deferred until the OLE2/OOXML optimization goal completes.
