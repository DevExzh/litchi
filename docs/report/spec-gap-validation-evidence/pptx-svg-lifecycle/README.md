# PPTX SVG attachment lifecycle evidence

This directory records evidence for source-backed PPTX SVG attachment and
detachment. Approval depends on the matching source-bound receipts described
below. The completed OPC removal and shared SVG namespace prerequisites have
their own committed,
source-bound review records (`opc-review.json` and `svg-namespace-review.json`).

The host operation selects an existing direct slide picture with an internal PNG
fallback. Attachment adds the native SVG extension, an owning slide relationship,
and an internal SVG media part. Detachment removes the selected owner and retains
resources referenced elsewhere. SVG media bytes are retained as opaque payloads;
the host validates package ownership and slide markup, not the SVG image grammar.
External resources remain inert. Selection and
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

`worktree-input.patch` records the pre-existing import-order-only change in OPC
`phys_pkg.rs` included in the build inputs. It is retained for reproducibility;
the production file is not part of this feature change. Gate receipts include
only logs executed by that invocation, and the verifier recomputes test totals
from the retained test log.

The standalone `harness/` exercises the ordinary fallible `commit()` API. It
checks exact no-op publication, attach/save/reopen, detach/save/reopen, retained
in-memory inverse patches, and preservation of raster and opaque package data.
It writes the three slide XML inputs used by the offline schema validator.
Running this example is a smoke check until the source-bound gate and probe
receipts described above are sealed.

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

The raw owner scanner has an explicit bounded profile: at most 32 MiB of slide
XML, depth 256, one million element nodes, 256 attributes and namespace
declarations per element, 16,384 active namespace bindings, and 4 KiB decoded
names/prefixes/namespace URIs. Namespace declaration lexical values have a 16 KiB
preflight bound. Namespace-complete parser fragments are admitted before their
output allocation and also respect the caller limit passed by the host. These
are support limits, not claims that larger documents are invalid XML. Namespace
lookup uses a scoped prefix index. Picture records share immutable ancestor
declaration layers, and source inventory uses one raw layout scan. Parser-only
namespace closure materialization remains explicit work; this does not imply
fully linear runtime for every workload.
The many-picture profile is required to assess host scan amplification separately
from payload size and inherited-namespace cost.
