# Checked ODS data-style value model

This batch adds the pure `litchi_ods::data_style` value model and checked
builders. It is a prerequisite for the full
[data-style owner design](../ods-data-style-gap-design.md), not evidence that
the package scanner or ordinary graph transactions are complete.

The model covers decimal, scientific, fraction, embedded text (including
currency/percentage bodies), transliteration, common metadata patches, and
bounded graphs. Existing date/time/boolean values use valid ODF particles.
Typed catalog entries carry their explicit source owner. The legacy public
style types remain unchanged. Opaque catalog bodies cannot be regenerated
through a raw XML escape hatch.

Independent review approved the frozen model hash
`6036bb67049db0f281c5d599661d26d816693ed18d2ab8b70a3cde41e5b73216`:
NCName and metadata lexical domains, XML whitespace and numeric overflow
distinctions, complete Unicode decimal-digit value-one transliteration forms,
ignored transliteration settings when format is absent, explicit/inherited
decimal resolution, fixed-denominator semantics, valid date/time particles,
affix/body constraints, and aggregate preallocation checks.

Root tested only this model and its public module export in an isolated checkout
at `3abd56b93c212b0d344f109a09a9923b719d7e03`. The results were **269 library tests passed**, strict
library Clippy passed, and rustdoc passed with warnings denied. No lint
suppression was used. These counts exclude the concurrent source scanner,
package integration, and new public integration tests. They establish no
runtime speedup or native producer interoperability claim.

The source manifest, toolchain, lockfile, command/exit records, and compressed
logs are retained here. `tested-source.patch.gz` reconstructs the tested two-file
change from the manifest's base. To replay, create a checkout at that base,
apply the decompressed patch, copy the retained Cargo.lock to its root, and
provide the repository's vendored `3rdparty` fixtures:

```sh
TMPDIR=/var/tmp cargo test --locked -p litchi-ods --offline --lib
TMPDIR=/var/tmp cargo clippy --locked -p litchi-ods --offline --lib -- -D warnings
TMPDIR=/var/tmp RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-ods --offline --no-deps
```

Source-qualified lookup, common-style read-only enforcement, metadata source
splicing, package graph publication, and end-to-end inverse/stale behavior retain
their separate integration and review gates.
