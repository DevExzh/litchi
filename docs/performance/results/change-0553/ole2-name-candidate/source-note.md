# 0553 OLE2 N candidate source note

`status: isolated, unmeasured, not admitted`

`scope: N name-only handoff; NF scalar-field handoff is excluded`

`performance_claim: none`

This candidate is a source-only experiment artifact for the directory handoff
attribution plan. It was prepared from repository revision
`8aa0c5baf0616d16c79eba0c6c28dc1716338ad6` with the live
`crates/litchi-cfb/src/file.rs` SHA-256
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The
isolated candidate file SHA-256 is
`e7d6b72e2389b4a47d4628e04f1337a143a06abeec3dcd0f108d42494fc1b6c9`, and
the patch SHA-256 is
`af6639d1ff99f9303031c4b51e26e3aa19a5154ad93314e6c949193ad067fb91`.
The patch headers use `a/crates/litchi-cfb/src/file.rs` and
`b/crates/litchi-cfb/src/file.rs`, so a future owner can apply it to the live
path after an explicit source review. The live crate and frozen change-0553
inputs were not edited.

## Candidate boundary

The current validation pass still creates one
`ValidatedDirectoryEntry` for every 128-byte record and retains its validated
`DirectoryNameData`. The N arm changes only the private validated name slot
from `String` to `Option<String>`:

* `Some(name)` owns the validated normal-entry string before graph construction;
* the iterative graph builder receives a mutable validated-entry slice and
  passes the exact SID's name slot to `parse_directory_entry`; and
* after raw parsing and the existing name-length check, the public parser uses
  `Option::take()` to move that string into `DirectoryEntry.name`.

The raw parser still performs `RawDirectoryEntry::read_from_bytes`, name-length
validation, CLSID formatting, version-3 size masking, MiniFAT classification,
and all existing scalar-field extraction. N does not add the NF arm's private
CLSID/start/size field seed. It therefore isolates normal-entry name reuse.
The root call deliberately passes no validated name, so the historical public
root parser remains in charge of the public root view.

The final SID-aligned cache loop now treats `validated_entry.name.is_none()`
as the normal-entry handoff proof. That replaces the old string equality check
because the candidate intentionally moves the only owned string; there is no
second string left to compare. The proof is exact: the graph builder looks up
the validated slot by the same checked SID, verifies its stored SID, and takes
the name only after the public raw record has passed the existing parse and
name-length checks. The final loop still checks the public and validated SID
alignment and rejects a normal entry whose name was not consumed. A root SID
is the explicit exception because validation canonicalizes the accepted
classic-Mac `00 52` spelling to `Root Entry` while the public parser retains
its historical decoded value.

This is an ownership-transfer invariant, not a relaxed integrity check. The
validation pass has already established that each nonempty SID is unique,
owned by one storage, and reachable under the validated ordering walk. The
builder's visited set retains its cycle/repeated-SID refusal. Consequently a
normal `None` at finalization can only mean that the exact SID's validated
name was moved into its public entry. An absent entry, wrong SID, or unconsumed
name still reaches the existing
`"directory entries and validated name cache disagree"` refusal. The root
exception retains the existing two-view comparison behavior.

## Ownership, drop, and refusal proof

The name allocation is moved, never cloned. While the graph is being built,
the private vector retains the `DirectoryNameData` and structural validation
fields, while the public vector owns each transferred `String`. Once the
final cache loop moves `name_data`, the private validated records contain no
normal-entry name allocation to drop. The public `DirectoryEntry` values,
comparison cache, and graph remain installed only by the existing final
assignments after the complete directory phase succeeds.

If raw parsing, name-length validation, or any later existing parse operation
returns an error, the staged `dir_entries`, validated records, and any moved
names are dropped locally; `self.root`, `self.dir_entries`, and
`self.dir_name_data` are not assigned. The new
`"validated CFB directory SID ... name was already handed off"` refusal is an
internal impossible-state guard after validation and SID visitation. It does
not replace an externally reachable malformed-input diagnostic: duplicate,
cyclic, cross-storage, invalid-SID, ordering, and ownership failures are still
returned by the unchanged validation or graph checks before a second handoff
can occur.

The validation error order remains first because
`validated_directory_entries` still runs before directory-entry allocation and
graph construction. Within graph construction, the candidate keeps raw entry
decoding and the name-length check before taking a name. The scalar parser and
all subsequent physical stream validation remain in their existing order.
The existing fallible resource labels and reservation sequence are unchanged:
directory data, directory entries, directory traversal queue/map, directory
name comparison data, and the prior FAT/MiniFAT/physical claims remain at the
same boundaries. The candidate adds no reservation and no public API.

## Focused checks included in the isolated source

The candidate file adds three focused unit tests next to the existing CFB
tests. They were included for the future differential/semantic gate but were
not built or run in this task:

* `normal_name_handoff_keeps_public_name_cache_and_scalar_fields` checks the
  transferred normal name, SID-aligned comparison data, stream kind, empty
  CLSID, size, MiniFAT classification, and child vector.
* `classic_mac_root_keeps_public_and_validated_name_views` rewrites the root
  directory record to the exact `name_len == 2`, `00 52`, zero-padded form and
  checks the public empty historical view, canonical `Root Entry` comparison
  data, and successful stream lookup.
* `name_handoff_does_not_bypass_validated_name_errors` changes SID 1's name
  length to `3` and checks the exact existing
  `invalid CFB directory name length 3 at SID 1` corruption refusal.

These focused checks supplement, rather than replace, the existing directory
validation, graph-cycle, cache-alignment, deep-tree, allocation-limit, source,
and public lookup tests. No test result is claimed here.

## Remaining measurement and review requirements

Before any adoption decision, apply the patch to a clean private candidate
source and run the fresh attribution campaign from
[`ole2-attribution-plan.md`](../ole2-attribution-plan.md):

1. Profile `cfb_open` over `tiny`, `many-small`, and `few-large` from
   `litchi-cfb-synthetic-v1`, plus
   `xls_owned_source_open_one_cell` over the fixed
   `litchi-xls-comments-opaque-heavy-v1` corpus. Attribute the positive timed
   owner edges rather than dump ordinals.
2. For standard roots, verify that the positive
   `parse_directory_entry -> decode_utf16le` edge falls to the root parse only
   (`1` call per open) while validation-side name work remains. Inspect
   `format_clsid`, scalar lines, inlining, and drop paths separately. This is a
   predicted mechanism check, not a result from this artifact.
3. Run the matched `litchi-perf-baseline-alloc` lane and compare allocation,
   deallocation, reallocation, failed-call, byte, live-before/after,
   peak-before/after, and region-peak fields. `Unavailable` and `Overflow`
   remain nonzero statuses, and allocator-region evidence does not establish
   RSS.
4. If attribution is clear, run the CPU-2 native ABBA latency matrix and the
   valid/error differential controls. Preserve the four primary XLS rows,
   CFB shape controls, source version/locality checks, exact public fields,
   error text/precedence, resource labels, and no-partial-publication proof.
5. Exercise the dedicated classic-Mac root fixture before retention. The
   synthetic harness has no selector for that two-view encoding, so a normal
   corpus pass cannot stand in for it.

The 0547 Callgrind evidence selected this mechanism but supplies no candidate
delta, allocation result, or native latency result. The 0548 Brent/checkpoint
and 0549 checked-test-and-mark candidates remain rejected and are not reused
as evidence. This N artifact is therefore only a concrete, reviewable source
candidate; it makes no speedup, regression, compatibility, or admission claim.
