# 0362ul62 — PPTX raw-content limits and source propagation

Status: source-ready implementation packet, based on the coordinator replay
integration tip `a942320908`. Native Cargo validation, the workspace gate and
ABBA qualification remain coordinator work because this wave was run without a
build slot.

## Boundary

The remaining public raw XML boundaries in `litchi-pptx` now take an explicit
module policy. The old media aliases and the default `TryFrom<&str>` extension
parser are removed; all in-tree callers use the canonical names and pass the
selected policy.

| Area | Public policy and raw operations |
| --- | --- |
| `media_parts` | `Limits`, `parse`, `write_pictures`, and `ExtensionList::parse`; media payload, extension, node, string, and output budgets are owner-derived for package operations. |
| `tag` | `Limits`, `parse`, `write`, and `List::xml`; raw XML input/output budgets are explicit. |
| `notes` | `Limits`, `master_xml`, `write_text`, `write_text_with`, and `Slide::text`; MCE, structure, text, and aggregate graph budgets follow the source profile. |
| `animations` | `Limits`, semantic parse methods, and `to_xml`/`to_xml_for_slide`; MCE input and fallible bounded XML output use the selected policy. |
| classic `comments` | `Limits`, all author/comment parse and write forms, part projections, and collaboration snapshots; package, eager, source-backed, and edit consumers derive the source policy. |

## Safety and fidelity

MCE processing is bounded before semantic parsing. Media, animation, and
comment serializers use checked sinks that refuse a limit before reserving the
next chunk. Notes rewrites preflight the final byte count before the exact
reservation and preserve opaque source markup. Existing source-splice and
no-op paths retain their original bytes; the collaboration snapshots retain
the selected limits and reparse with the same profile after edits and replay.

Each codec has exact-wire/one-byte-under policy coverage, with media, notes,
comments, and animation round-trip coverage kept beside their raw tests. The
glossary records the required limits signatures for both `pptx-minimal` and
`pptx-full`.

## Source-only checks

```text
rustfmt --edition 2024 on all changed Rust files       PASS
git diff --check                                      PASS
legacy raw aliases/call-site search                  PASS
Cargo/native/full gate/ABBA                            pending coordinator slot
```

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
