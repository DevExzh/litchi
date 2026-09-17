# Change 0677 evidence

[Record](../../0677-xml-publication-bom-offsets.md).

- `before.txt`: the new admission regression on the prior implementation; it fails with the original byte-zero declaration refusal.
- `suite.txt`: complete `xml-minifier` test suite on the corrected implementation.
- `publication.txt`: public OPC tests, including marked replacement publication, no-output DTD refusal, and all 16 marked members from the real corpus.
- `clippy.txt`, `doc.txt`: warning-denied library lint and rustdoc.
- `corpus.tsv`: marked XML member census across the repository's OOXML fixture suffixes.

Reproduce with `cargo test -p xml-minifier --locked`,
`cargo test -p litchi-opc --test source_xml_publication`,
`cargo clippy -p xml-minifier --lib --no-deps -- -D warnings`, and
`RUSTDOCFLAGS='-D warnings' cargo doc -p xml-minifier --no-deps --locked`.
No benchmark or registered performance claim is added.
