# 0362ul156 — remove public ODT and PPTX format aliases

- Date: 2026-10-25.
- Base: `f8cd398a770f7363f628be40edb3ce4219566c16`.
- Status: source-ready; the coordinator must run the focused Rust and
  rustdoc/profile validation before accepting this lane.

The ODT `transaction::Edit` no longer exposes the eleven deprecated unchecked
`usize` selector forwarders: RDF triple removal/move, form-control
removal/move, form removal/move, embedded-chart removal, embedded-object and
embedded-image removal, and embedded-object and embedded-image moves. The
checked `Position` routes (`*_at` and `*_to`) remain the sole public routes and
retain the existing no-op and operation-limit behavior. The source audit found
no ODT consumers needing migration; similarly named ODP, ODS, ODG and DOC
APIs are separate surfaces and remain unchanged.

PPTX `animations::Effect::to_preset` is deleted. Existing writer, parser,
example and test consumers already use `preset_class`; no behavior or encoded
animation data changes. The `preset_id` and `from_preset` capabilities remain
available.

Both PPTX glossary profiles and the ODT profile register exact bans for the
removed public names. The human glossary records the checked ODT selector
family and the canonical PPTX effect accessor. No deprecated alias or public
forwarder remains in either target surface.

## Source checks

The following source-only checks passed on the changed candidate:

```text
git diff --check
python3 -m json.tool tools/api_glossary.json
python3 -B - <<'PY'
import json
from tools import check_api_glossary
check_api_glossary.checked_policy(json.load(open("tools/api_glossary.json")))
PY
```

Targeted `rg` scans found no `to_preset` declaration or active ODT call to the
eleven removed names; the remaining same-spelled matches are private ODT
codec helpers or separate ODP/ODS/ODG/DOC APIs. The checked ODT routes and
PPTX `preset_class` consumers remain present. No Cargo, rustc, rustdoc, native,
corpus or performance command was run in this source-preparation phase.

The coordinator's validation proposal is the focused `litchi-odt` and
`litchi-pptx` test/build run, followed by both PPTX glossary profiles and the
ODT glossary profile against fresh rustdoc. The shared glossary JSON may need
conflict resolution when composed with the coordinator's concurrent batch
worker; the rows in this commit are complete and scoped to these exact APIs.
