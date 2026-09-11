# PPTX SVG attachment lifecycle evidence

This directory is being prepared for a source-backed PPTX SVG attachment and
detachment implementation. It is not yet a final feature approval. The completed
OPC removal and shared SVG namespace prerequisites have their own committed,
source-bound review records (`opc-review.json` and `svg-namespace-review.json`).

The host operation selects an existing direct slide picture with an internal PNG
fallback. Attachment adds the native SVG extension, an owning slide relationship,
and an internal SVG media part. Detachment removes the selected owner and retains
resources referenced elsewhere. External resources remain inert. Selection and
rewrite ranges come from the original XML, with resolved inherited namespaces.
Unrelated opaque extension payloads and package members must remain unchanged.

`requirements.md` defines the evidence needed to close this batch. The validation
runner starts with compilation and records checks, tests, strict clippy, rustdoc,
formatting, and diff results against an unchanged source manifest. The manifest
includes local dependency source, tests, examples, benches, validation tools, and
retained fixture identities. `run_probe.py` runs the public downstream harness and
offline complete-slide and SVG-extension schema checks. `verify.py` requires all
production files in the changed ownership path to have source-bound review; an
empty or partial review cannot pass.

The scripts and generated outputs are preparation artifacts until current
`gates/receipt.json`, `review.json`, `probe-receipt.json`, and `verification.json`
exist and agree. Preliminary generated output alone is not final evidence.
Performance measurements are maintained separately in
`../pptx-svg-lifecycle-performance/`; final measurements must follow source freeze.

Exact patch inverse means applying the retained inverse to its authorized
in-memory target snapshot restores the stored source state. An independently
reopened saved package does not carry that patch provenance. A fresh detach after
save is a new semantic edit and need not reconstruct whether an original empty
blip used paired or self-closing syntax. The downstream fixture deliberately uses
a paired blip with an existing opaque extension list so it can additionally assert
exact source-slide restoration after fresh attach/save/detach/save.

The corpus scan records local producer fixtures and explicit refusals, not native
Office acceptance of newly generated files. Rendering, rasterization, external
fetching, picture creation/reordering, and durable cross-process patch transport
remain separate work. No performance improvement is claimed without measurements.
