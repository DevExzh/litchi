# 0190 — Immutable DOC package and document projection

## Change

The native root `Package` is now a concrete immutable owner. It retains one
owned CFB index, a shared parsed `Document`, metadata, and its limits/leniency.
`document()` returns an owned one-pointer projection whose clones share the
parsed tables. It stays usable after the package is dropped. Publication
streams the retained bytes; encrypted packages retain ciphertext while their
semantic projection contains the successfully decrypted content.

`OpenOptions` combines package/stream limits, leniency, and the zeroizing
password. `PackageOpenOptions` and root `open_with` are removed. Canonical
`open_with_options`, `from_reader_with_options`, and `from_bytes_with_options`
consume the policy at ingress. Passwords are not retained in the resulting
package. Borrowed byte input is size-checked before a fallible allocation.
Path and reader ingress seal owned bytes: source versions identify that sealed
snapshot, and subsequent changes to the original file do not alter it.

The cursor-based physical reader is now explicitly `raw::Package<R>`.
Physical stream, signature, VBA and property-set consumers use that boundary.
Raw ingress accepts physical limits; it does not expose an options entry that
would silently discard password or leniency. Semantic raw reads still accept
explicit document options. The immutable owner transfers the existing CFB
index into shared storage without re-indexing the container.

`Builder::build` returns the immutable owner and validates its semantic
projection. Encrypted builders supply their existing password to that final
validation. Unified creation and lazy edited-document reads consume the new
owner. Detection retains the raw ingress path so `PasswordRequired` can return
the original CFB container for the facade's locked state. Frozen physical-read
benchmark routes retain their original raw operations under the explicit
namespace; their route names and timing boundaries are unchanged.

## Verification

The combined DOC and DOC/DOCX/XLSX/XLSB/PPTX/encryption facade suite passes
2,696 tests and doctests, with 14 ignored. The earlier DOC-only run passed
1,892 tests/doctests. New tests cover detached cheap clones with shared borrowed
text, exact byte/stream publication, merged input/stream limits, path ingress
limits and owned-source identity, encrypted builder validation, missing/wrong/
correct passwords, and exact ciphertext publication. The public package
example is compiled and executed as a doctest.

All-target warning-denied Clippy passes for native DOC, the combined facade,
and the DOC-enabled public crate. Warning-denied rustdoc and minimal/full DOC
glossary checks pass. Glossary policy now rejects the removed options type,
old root ingress spelling, option-bearing projection access, and ordinary
mutable OLE access on the semantic owner. Performance-tool library and binary
consumers and the DOC parser fuzz target compile (no fuzz campaign is claimed). The generated OfficeCLI capability delta is unchanged.
The boundary audit passes for 66 packages and 268 internal declarations
(11 existing debt items). Repository lockfiles are unchanged; checks use an external copied lockfile
with cached cc 1.6.0. No iWork crate is modified.

## Remaining scope and limits

This is the immutable read/publication and ingress-options migration, not the
full native transaction facade. Primary `Package::edit/apply`, the canonical
Commit/Patch/Diagnostics family, `SourceLimits` naming/options builders, and
removing the remaining focused/raw vocabulary variants are still open. The
facade's detection adapter still directly retains a raw-derived document and
shared source. Generic source access on the new owner also remains an explicit
follow-up for the final raw-boundary API audit.

DOC limits bound the outer artifact and selected DOC payload streams; CFB
index allocations retain the CFB reader's separate finite defaults. This wave
does not add caller-configurable CFB allocation limits. Canonical package
metadata parsing returns typed errors; the facade's pre-existing tolerant
metadata projection is not migrated here.

Eager semantic validation in `Builder::build` adds work relative to the old
physical-reader result. Existing fresh-write and open performance-program
routes require before/after qualification before performance acceptance. No
claim of unchanged builder cost or improved edit speed is made. Changed-closure
ownership, the single-write budget miss from records 0188–0189, general DOC
CRUD, XLS/PPT loops, tier 3, Python/CLI and the full goal remain open.

## Paired paragraph measurements

The [after](results/0190-doc-paragraph-after.json) and
[repeat](results/0190-doc-paragraph-repeat.json) receipts use the unchanged
native Snapshot open/edit/commit versus freshly opened facade set boundaries.
Both run nine ABBA rounds on CPU 0, with no concurrent build, test, or boundary
audit. Both failure matrices are empty and all paired output bytes match.
Source manifests and binary/lockfile digests accompany the raw samples.

| Existing paragraphs | After native / facade µs | Repeat native / facade µs |
|---:|---:|---:|
| 1 | 47.820 / 49.461 | 47.360 / 48.730 |
| 10 | 60.610 / 65.261 | 60.441 / 64.700 |
| 100 | 187.161 / 205.491 | 183.871 / 208.791 |
| 1,000 | 2367.751 / 2428.351 | 2276.530 / 2388.991 |

The repeated 100-paragraph facade overhead is 13.6%, above the 10% budget.
The 1,000-paragraph ratio is lower than in the prior investigation, but these
receipts do not establish a causal fix or qualify every existing route. The
native leg also changes with the shared projection implementation; ratio
improvement alone is not evidence of an end-to-end speed improvement. Fresh
builder validation cost is outside these paired edit clocks and remains a
separate required qualification.
