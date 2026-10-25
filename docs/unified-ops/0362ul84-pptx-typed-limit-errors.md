# 0362ul84 — PPTX typed limit error boundary

- Date: 2026-10-25.
- Base: `0d38e80bd8fbe1d9a5f060097636524e08e2cb81`.
- Scope: the shared `litchi-pptx::Error` limit boundary and every migrated
  PPTX/ops consumer present on this source lane.
- Validation state: source-only preparation; Cargo, rustdoc, native fixtures
  and performance commands were intentionally not run before coordinator
  composition.

## Canonical payload

PPTX now owns one finite `LimitKind` vocabulary. The primary opened-edit
dimensions (`Parts`, `PatchBytes`, `TextBytes`, history and retained-byte
budgets) remain direct variants. Codec dimensions are named variants of
`CodecLimitKind`, and shared OOXML dimensions are carried by the common
typed enum. The public failures are exactly:

```text
Error::Limit { kind, observed, maximum }
Error::InvalidLimit { kind, value, maximum }
```

The old `resource: &'static str`/`limit` payload and the duplicate edit-limit
path are gone. Display text is fixed and content-free. `observed` is the
measured count or byte charge; `usize::MAX` is emitted only by a checked
arithmetic or checked conversion failure, never as a normal `maximum + 1`
substitute. The migration removes the former `Count`, `Candidates`,
`Entries`, `Items`, `Structure`, `Work` and `Events` catch-all dimensions.

## Consumer migration

The root error, opened re-exports, PPTX codecs, source-backed slide paths,
cross-copy/replay paths, streaming writer, embedded ink/actions, controls,
OLE, charts, tables, comments, tags, fonts, model/view properties, notes,
transitions, and the litchi-ops error/facade mappings use the typed payload.
Observed values were corrected at the high-risk boundaries: ink annotation,
trace/group/point, chart series/point, controls, OLE objects, content parts,
designer tags, table rows/cells/styles, comment entries/replies/authors,
source-copy conversion checks, and output versus input XML checks. Layout
names use `CodecLimitKind::NameBytes`; text-property accumulation uses
`ShapeProperties` or `TextBytes` according to the measured quantity; source
rebase/output checks use `OutputBytes`.

The focused error regression tests assert the finite dimensions, exact
observations, and content-free display strings. The glossary now records the
PPTX root/codec dimensions and canonical payload shape.

## Deliberately separate owners and follow-up refresh

Wrapped errors remain owned by their source crate and are not flattened into
the PPTX enum: OPC source limits, encryption/crypto limits, shared OOXML
limits, DrawingML limits, and validation-report limits retain their typed
foreign payloads. Private structural `Error::Invalid` checks in focused
validators are likewise not reclassified as the public PPTX limit boundary.

The versioned cross-copy digest transcript still contains historical
resource-string/refusal renderings and incomplete `InvalidLimit` renderings.
Its exact new `observed` values must be regenerated from the composed source
after the coordinator grants the Cargo slot; no value is fabricated in this
source-only lane. The transcript is listed explicitly so it cannot be treated
as a completed consumer during integration.

Static checks completed on this lane:

```text
stale CodecLimitKind catch-all variant search       PASS
root Limit { resource, limit } search               PASS (source consumers)
checked observed/MAX audit                          PASS (reviewed source sites)
git diff --check                                    PASS
Python glossary/no-3rdparty unit checks             PASS
```

The final composed PPTX wave must rerun the glossary, compile/test, and
transcript regeneration checks against the coordinator's integrated tip.

## 84a source audit

The follow-up audit repaired every two-argument local `limit` call in the
changed PresentationML helper families, including presentation properties,
readonlyRecommended, view properties, designer tags, comments, model3d, and
the source-backed primary topology/durable/insertion paths. Each call now
supplies its finite dimension and the measured attempted charge. Input,
output, XML, text, attribute, name, relationship, part, slide, and proof
charges are kept distinct.

The former `CandidateEntries` catch-all was removed. Font, ink, OLE, table,
tag, SVG, and other allocation loops now report their concrete part,
relationship, or family dimension. Tag owner references use the configured
node budget; mapped tag shapes use the configured shape budget. Slide event
bytes and decoded text bytes use byte/text dimensions rather than event
counts.

Admission checks now report prospective added parts and the two-copy
presentation XML patch charge. Source-backed budgets preserve a finite total
when it exceeds its bound; `usize::MAX` remains reserved for checked
arithmetic or conversion overflow. Depth and incoming-item observations are
reported as the actual prospective count. `opened::Limits::maximum` returns
`Option<usize>` for dimensions owned by that policy, so foreign codec/common
dimensions no longer use a sentinel maximum.

This audit remains source-only. Rustfmt and static repository checks were run;
Cargo, rustdoc, native fixture execution, and performance measurement remain
for the coordinator's composed-tip gate. The historical cross-copy transcript
still awaits exact regeneration on that composed tip.

## 84b call-shape audit

The whole-crate `limit("…")` search was checked against each callee rather
than treating every one-argument helper as the same API. The remaining calls
to typed three-argument helpers were repaired in modern-comment authors and
the inert OLE slide/package paths. Modern-comment processed XML now reports
`XmlBytes` with the measured byte length, and author strings report
`StringBytes` with their measured length. OLE object, relationship, part, and
payload failures use their concrete dimensions and prospective counts. The
source-backed insertion overflow path now retains the configured owner part
maximum while keeping `usize::MAX` as the genuine overflow observation.

The other one-argument matches belong to private legacy helpers that still
construct their own invalid-input text; they are not calls to a typed
three-argument helper and were left outside this repair. A final composed-tip
compile is required to validate all cross-lane callers.
