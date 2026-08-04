# Deferred ODS integration coverage

The Rust sources in `legacy/` are preserved from the pre-split umbrella suite.
They exercise sheet mutation, conditional styles, shapes, images, sparklines,
DDE sources, and tracked changes, but currently require the old monolithic
`FlatSpreadsheet`/sheet-authoring surface. The dedicated `litchi-ods` facade
currently owns package validation, content/style access, RDF CRUD, and the
formula/model codecs; the legacy files stay here until those semantic owners
are wired into the facade. They are intentionally outside Cargo's direct
integration-test discovery so the active crate test build remains compile-safe.
