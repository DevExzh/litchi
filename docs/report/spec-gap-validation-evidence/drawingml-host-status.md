# Committed DrawingML Ink/SVG host status

Review date: 2026-09-11.

This note records only committed capability. Working-tree DOCX/XLSX SVG edits and
Ink action-mutation experiments are excluded until they have their own commit,
source-bound review, and gates.

## SVG

The committed shared `svg_blip` codec and contextual namespace support are in
`1d29becce`. PPTX host coverage is split across two committed batches:

- `de165e1e6` replaces and retargets an existing admitted SVG picture with
  source-bound SVG payload, raster fallback, bounded media/relationship graph
  edits, stale-source checks, and exact no-op/inverse behavior.
- `511818cd8` attaches or detaches SVG on an existing direct slide picture with
  an internal PNG fallback. Its bounded commit validates slide XML,
  relationships, content types, media ownership, shared-resource retention,
  and save/reopen output; the retained inverse is an authorized in-memory
  inverse.

The lifecycle evidence is
[review.json](pptx-svg-lifecycle/review.json),
[verification.json](pptx-svg-lifecycle/verification.json), and
[schema-validation.json](pptx-svg-lifecycle/schema-validation.json).
Its approved scope excludes picture creation or
reordering, ambiguous/MCE-wrapped picture owners, external or linked SVG
targets, rendering, native Office acceptance, and durable cross-process patch
transport. DOCX and XLSX have no equivalent committed SVG relationship,
fallback, and resource lifecycle owner; current working-tree implementations
do not change this status.

## Ink and InkAction

The shared InkML codec has bounded metadata/source replay and detached finite
X/Y authoring. `2913e62c2` adds a structured `ink::actions` read/write profile;
`836398167` adds the XML character and decoded namespace validation follow-up.
These commits provide bounded action-container structure, exact source replay,
and opaque definitions/transforms/traces. They do not provide action mutation,
full InkML semantics, recognition, replay, rendering, or execution.

PPTX's committed host owner is the inert InkML `p:contentPart` inventory/storage
path (`customXml`, `application/inkml+xml`). No PPTX `inkAction` package owner is
claimed. [Packaging evidence](ink-action-profile/packaging-evidence.md) and
its [scan receipt](ink-action-profile/packaging-scan.json) record no raw action
marker matches in the successfully scanned inputs. Encoded namespace spellings
and unreadable inputs are outside that negative result; it does not prove
that the corpus contains no action-bearing package. No native or independently
validated action package has been established, so its action content type,
relationship URI, and target path still need evidence. Current working-tree
action CRUD or package-host experiments are outside this committed status.

These boundaries keep shared codec support separate from concrete package
ownership and avoid treating the InkML part contract as an InkAction contract.
