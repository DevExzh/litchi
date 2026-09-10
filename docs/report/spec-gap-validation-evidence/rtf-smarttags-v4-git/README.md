# RTF SmartTags and move bookmarks

The clean feature candidate is based on
`f0097deb0e80aa4ca7d4347cbba1aebdeadee764`. Its 24 source paths are
bound by `root-source.json` and `candidate-manifest.json`.

The change adds typed SmartTag/XML namespace metadata, direct attribute PCDATA
parsing, source-preserving editing, and move-bookmark metadata. It validates
namespace references and indexed move identities, supports namespace zero,
and charges attribute/destination/aggregate bounds before materializing data.
No-op and inverse paths retain their exact source bytes. This batch does not
include the separately pending modern password-hash implementation.

Root Rust 1.95.0 validation on the real Git candidate passed 1,367 ordinary
tests, nine documentation tests, strict all-target/all-feature Clippy,
warning-denied rustdoc, and formatting for the owned Rust files. Compressed
logs retain exact raw hashes in `root-gates.json`.

Independent review checked all 24 file hashes against the byte-identical
preserved V4 copy, local RTF 1.9.1 grammar, direct and nested attribute forms,
malformed ordering and counts, pre-allocation charging, exact no-op/inverse
behavior, and the 65,536-entry indexed stress case. The clean V4 Git worktree
was reconstructed and verified separately; its root gates are authoritative.

The API retains the documented main-body scope and inert unmatched moves.
Hybrid direct-name/nested-value syntax lies outside the normative grammar.
