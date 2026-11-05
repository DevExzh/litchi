# 0362zac — legacy raw Reader limits and owned errors

Status: preparation packet, based on `e35cede32ae730d351ce9d29d243c35edd321ddf`.
This record covers the initial legacy DOC, PPT and XLS cursor-backed reader
boundary. Its validation and required typed-error repair are recorded in
follow-up `0362zaca`.

## Boundary decision

The raw `Read + Seek` reader is the narrow codec exception in 0351 §2.1.6.
Its `&mut self` receiver represents cursor/container codec state, not mutation
of a package owner. The reader still returns the canonical shared models where
those models are the established substrate types: OLE property-set streams and
sections, the MS-OVBA project, signature reports, and the DataSpaces graph,
snapshot, and patch. No format-local wrapper was added solely to rename a
shared OLE or crypto model.

Every foreign fallible operation now crosses the owning format's error type:

- DOC adds `CustomXml`, `Vba`, `Signature`, and `Crypto` error variants. A
  DataSpaces patch conflict is classified as the local `PatchConflict`; other
  DataSpaces errors retain their typed source in `Crypto`.
- PPT adds `CustomXml` and feature-gated `Signature` variants.
- XLS adds `CustomXml` and `Signature` variants while retaining its existing
  `Vba` variant.

The new variants implement `Display`, `Error::source`, conversion into the
format's `litchi_core::Error`, and the corresponding `From` conversions. The
shared result/value types remain available through the typed source and return
values, but no `litchi_*::Result` escapes these raw format APIs.

## Required limits and inherited budgets

The former default twins were removed and all repository consumers were
migrated:

| Format | Raw API | Limit/policy boundary |
| --- | --- | --- |
| DOC | `custom_xml_data_store(Limits)` | caller-selected Custom XML limits |
| DOC | `vba(&litchi_vba::Limits)` | caller-selected MS-OVBA limits |
| DOC | `vba_signature(vba_signature::Limits)` | caller-selected blob/payload limits |
| DOC | `signatures(&Policy)` | caller-selected signature policy |
| PPT | `custom_xml_data_store(Limits)` | caller-selected Custom XML limits |
| PPT | `fonts(font::Limits)` | caller limits constrained by the Reader's source profile before parsing |
| PPT | `picture_bullets(PictureBulletLimits)` | caller limits constrained by the Reader's source profile before parsing |
| PPT | `signatures(&Policy)` | caller-selected signature policy |
| XLS | `custom_xml_data_store(Limits)` | caller-selected Custom XML limits |
| XLS | `vba(&litchi_vba::Limits)` | caller-selected MS-OVBA limits |
| XLS | `signatures(&Policy)` | caller-selected signature policy |
| XLS | `revision_log(&SourceLimits)` | the Reader retains its ingress `SourceLimits`; CFB declared-stream preflight runs before materialization and the lazy method checks the component-wise stricter requested profile before the revision parser |

XLS now stores the selected `SourceLimits` in the raw Reader and exposes a
read-only profile getter. Its validation-only constructor initializes the
same field explicitly. The open-options test proves that a custom profile
survives admission. The revision-log check is defensive after ingress
preflight, so an already-admitted source cannot silently widen the budget at
the lazy parser boundary.

## Raw output receiver ruling

DOC `to_bytes`, `write_to`, `save`, and `save_with_durability` retain
`&mut self`. Existing contract records 0351 §2.1.6 and 0359a §4 classify these
as raw cursor/codec transport methods; 0362rd records the same substrate scope.
This packet therefore does not invent a new owner exception: the raw reader is
the explicitly permitted mutable codec model, while the DOC package owner
outputs remain `&self`, as required by the owner contract.

## Consumer migration and checks

The DOC, PPT and XLS signature/VBA tests, the XLS source-limit test, and the
security-corpus consumers in `tools/perf-baseline` use the new explicit APIs.
`litchi-vba` is declared directly by that tool because the raw method's
required limit type is intentionally not hidden behind a format alias.

Completed during preparation without Cargo:

```text
git diff --check                         PASS
raw API consumer search                  PASS (no legacy raw no-argument calls)
staged-size/binary guard                 pending until commit staging
```

The full glossary/inventory gate remains coordinator work after this commit is
merged into the pinned staging tip.
