# 0359f — Raw OLE common bounds and typed errors

- Date: 2026-10-21.
- Base: `45a5f9e9d1`.
- Status: raw shared OLE directory, OLEDS, link, and protection migrations are
  implemented on `uo2/w5-ole`; integration and the workspace gate remain
  coordinator work.

This slice completes the required-limit and error migration for the shared
raw OLE helpers. Every parser, constructor, serializer, and source-backed
replay route receives its limit policy as a required argument. `Default` still
describes a caller-selected policy; no raw helper creates one implicitly and
there are no `_with_limits` twins or aliases.

The directory catalog and transaction now use the common OLE error. Catalog
validation reports exact `LimitKind` dimensions for directory entries, name
bytes, total catalog bytes, and recursive directory depth, and directory
patches report `PatchConflict` through the common error. OLEDS presentation, native,
TOC, clipboard-format, target-device, and serialization routes use
`ole_streams::Limits` and return the common typed error. Oversized complete
streams report `LimitKind::StreamBytes`, including the snapshot constructors;
presentation and native snapshots retain the policy for edits and replay.

Object links use a required `object::link::Limits` policy throughout parsing,
serialization, snapshot edits, and patches. Link patches preserve that policy
when replaying or inverting and classify stale source bytes as
`PatchConflict`. Object, DOC, and test consumers were migrated in the same
wave. The OLEDS consumer correction also removes redundant error conversions
left after the raw helpers began returning the common error.

Protection inspection has a public `protection::Limits` policy with a required
nonzero directory-entry bound. Both mutable CFB and shared positional guards
charge directory traversal against it and return typed low-level limit errors
when the bound is crossed. Object admission, Property Set editing, and every
source-backed overlay constructor pass an explicit bound. Protected marker
refusals and their existing byte-preserving behavior remain unchanged.

The implementation commits are `d306553924` (directory), `3aa07c0496`
(OLEDS), `6066d2e86f` (OLEDS typed consumer correction), `7afe696ea7`
(links), and `f7d99def3f` (protection). The object physical-candidate test
also now asserts the typed `StreamBytes` result required by the common error
conversion introduced in the base object slice. Source-backed overlay guards
use `SharedOleFile::directory_entry_count()` so admission does not perform a
second O(n) directory count before the bounded protection walk.

The full `litchi-ole-common --all-targets` suite passes (222 unit tests and all
integration targets, with one pre-existing externally sourced test ignored), as
does common all-target clippy with `-D warnings`; DOC/PPT test-profile checks
also pass. No native performance measurement belongs to this raw-helper slice.
The shared object owner/options/source-sharing migration, legacy consumer
integration, and the full workspace gate remain required.

The common conversion boundary matches only the fixed labels owned by these
raw helpers. Physical CFB labels such as input, directory-image, and
allocation-table bytes remain structured `Error::Cfb` failures rather than
being guessed as stream bytes. The reverse `From<Error> for OleError` adapter
was removed; the directory codec keeps its one internal SID bridge explicit,
so owner refusals and patch conflicts cannot lose their category in a
low-level format string.
