# Deferred ODP integration coverage

The Rust sources in `legacy/` are preserved from the pre-split umbrella suite.
They require the former mutable presentation, image, drawing-style, handout,
and layout/master-page APIs. The dedicated `litchi-odp` facade currently owns
package validation, slide/settings/declaration/layout readers, RDF CRUD, and
the presentation builder; these files remain scoped here until the missing
authoring owners are connected. They are outside Cargo's direct integration
test discovery so the active crate test build remains compile-safe.
