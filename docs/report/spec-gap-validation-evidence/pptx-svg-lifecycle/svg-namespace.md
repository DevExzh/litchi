# SVG default namespace clearing

The shared SVG fragment codec rejected a legal `xmlns=""` declaration on a
prefixed `asvg:svgBlip`. The retained LibreOffice fixture
`3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf169496_hidden_graphic.pptx`
uses this form, so source-backed SVG inventory and lifecycle operations refused
it before reaching the host edit.

`Namespace::new` now allows an empty default namespace URI and still rejects an
empty URI for a named prefix. Reserved XML/XMLNS binding checks remain intact.
The regression verifies exact unchanged replay, relationship mutation/readback,
retained namespace clearing, preserved unqualified opaque children, both public
default-prefix forms, and rejected empty named bindings.

The root focused command is `cargo test --locked -p litchi-drawingml --test
svg_blip`, with eight passing tests and no warning suppression. The raw log is
`svg-namespace-focused.log`. Native fixture parsing is a compatibility
observation, not evidence that a native Office application accepts generated
output. Full PPTX lifecycle validation and profiling are separate work.
