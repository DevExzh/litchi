# XLSB unified worksheet-name writes

`File::set("/sheet[1]", props! { "name" => "Input" })` now renames an XLSB
worksheet through the native primary `Package::edit` transaction. Source-backed
and eager owners participate in queries, explicit transactions, process-local
patch replay/inverse, atomic/best-effort batches, streaming publication and
atomic path saves. Path opens retain their origin for `File::save`.

The local static registry marks only `sheet.name` as writable. Admission rejects
unsupported targets/properties, duplicate keys, ambiguous multi-sheet renames,
invalid names and Unicode-caseless duplicate names across the complete tab catalog before materializing a native
editor. Limits account for targets, operations, staged names and retained patch
images. Exact no-ops retain the original immutable facade owner. A successful
changed native package is projected into a new eager owner, retaining metadata,
origin and input policy.

Signed files publish their exact original bytes until the caller explicitly
uses `Signed::unsign`. Signed mutation attempts refuse even for a same-name
assignment. Source freshness fences bracket reads and edits; atomic publication
checks the source immediately before replacement. `Overwrite::Refuse` uses the
shared native no-clobber publisher. XLSB re-encryption remains an explicit
unsupported capability and cannot silently publish plaintext.

The budget regression exposed native structural quotas encoded as
`PackageError::InvalidLength`, which the facade correctly classified as corrupt
input. Structural operation, encoded-patch byte and history-byte quotas now emit
the existing typed `PackageError::LimitExceeded`, preserving the native source
chain and the facade's `limit_exceeded` code. Malformed envelope lengths remain
corruption errors. This does not complete the scoped native Error/LimitKind
migration.

Review also found that the native rename codec discarded tolerated trailing
bytes in `BrtBundleSh`. It now preserves the suffix after the original name;
regression coverage checks rename publication and durable replay/inverse.
A semantic same-name assignment also preserves the original record bytes,
including tolerated unpaired UTF-16 code units, rather than rewriting them
from the decoded replacement character.

## Scope

Cell writes, additional sheet properties, structural add/remove/move/swap,
creation and encrypted XLSB output remain unsupported. Source writes still
materialize an eager workbook; native structural staging retains complete
before/after images. Multiple facade commands currently use separate native
commits. Native scoped-family migration, bounded coalescing, wider corpus/fuzz
acceptance, remaining formats and Python remain part of the active goal.

## Verification

Before the last preservation/catalog review fixes, the full XLSB facade suite
passed 500 tests/doctests, the combined DOCX/XLSX/PPTX/XLSB/encryption facade suite
passed 1,168, and the native `sign,vba-inspection` suite passed 1,088. After those
fixes, all 13 native structural tests and all 13 combined-feature public XLSB
write tests passed. The latter includes the final no-plaintext re-encryption
regression. These focused reruns cover the final changes; the earlier full-suite
counts are not presented as final-code full-suite results.

Warning-denied native/facade all-target Clippy, warning-denied rustdoc, minimal
facade compilation, scoped formatting and diff checks passed.
The XLSB capability golden and OfficeCLI delta were regenerated; deterministic
parity checking and all 14 generator tests passed. Crate boundaries passed for
66 packages and 268 internal dependency declarations.

Reproduce with `cargo test -p litchi-ops --no-default-features --features xlsb`,
`cargo test -p litchi-ops --no-default-features --features
docx,xlsx,pptx,xlsb,encryption`, and `cargo test -p litchi-xlsb --features
sign,vba-inspection`. The final focused targets are `--test xlsb_writes` and
`--test workbook_structure_edit`. Full-feature runs used an external copy of
the root lockfile resolved offline to cached `cc` 1.6.0 through Cargo's unstable
lockfile-path option. Both inherited lockfiles remain unchanged.


## Paired rename measurement

[Results](0169-xlsb-sheet-name-results.json) retain all samples, corpus hashes,
source-file digests and binary provenance. The standalone harness is
`tools/perf-baseline/unified-xlsb-sheet-name`. It pairs prepared eager owners
with eager owners and source-backed owners with source-backed owners. The
source native leg includes materialization into a package within its clock,
as the facade must do. Both legs perform one rename and one native commit;
opening, prepared target/property construction, serialization and output
comparison are outside the clocks.

Three ABBA rounds provide six samples per leg, pinned to CPU 0 with no other
validation or measurement job running. Every case has exact ZIP-byte equality
and the failure matrix is empty.

| Stored cells | Eager unified/native | Source unified/native |
|---:|---:|---:|
| 1 | 1.082 | 0.997 |
| 10 | 1.031 | 1.021 |
| 100 | 1.034 | 0.998 |
| 1,000 | 1.031 | 1.030 |

One-cell eager native/unified medians were 0.311/0.337 milliseconds.
One-cell source native/unified medians were 0.448/0.446 milliseconds.
The measured cases fall within the 10% single-write review threshold; ratios
below one are timing variation. This is one rename per file, not a many-operation
batch or real-world corpus acceptance. Allocation/RSS and source materialization
cost reductions remain unmeasured follow-ups.

## Cleanup

Temporary targets, copied lockfiles, raw logs and measurement scratch are removed
after retaining the binary name, size, SHA-256 and regeneration command in the
result record. No new worktrees or borrowed fixtures were created. The rewritten
base remains an ancestor. The all-format goal remains active.
