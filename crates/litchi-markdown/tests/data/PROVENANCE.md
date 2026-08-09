# Markdown test-data provenance

`normative.rs` contains selected verbatim Markdown inputs from the generated
test suite distributed in the locally installed `pulldown-cmark 0.13.4` crate:

- `tests/suite/spec.rs`, generated from CommonMark 0.31.2 examples.
- `tests/suite/gfm_strikethrough.rs`, `gfm_table.rs`, and `gfm_tasklist.rs`.

The upstream crate identifies its license as MIT. The selected cases exercise
this crate's exact-source model; they do not constitute the complete normative
CommonMark or GitHub Flavored Markdown suites and do not independently verify
HTML rendering.

The `roundtrip-*.md` documents are project-authored integration fixtures under
the repository's Apache-2.0 license.
