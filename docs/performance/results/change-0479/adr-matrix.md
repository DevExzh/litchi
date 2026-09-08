# 0479 ADR matrix: DOCX plain-paragraph append baseline

This matrix records whether the existing 0479 harness stays within the
accepted architecture while measuring the public source-backed DOCX
plain-paragraph copy path. It records contract evidence only. It contains no
performance result or support claim, and it does not authorize an explicit
source-window append API.

The 30-file accepted ADR/README manifest was revalidated without hash changes;
the machine-readable result is
[`adr-refresh.json`](adr-refresh.json), whose `all_prior_adr_hashes_unchanged`
field is true. The audit also leaves the user-owned `docs/GOAL.md` and
`docs/report/spec-gap-audit.md` unchanged. iWork is outside this workload.

## Actual baseline boundary

The harness in
[`docx_plain_paragraph_tail_append.rs`](../../../../tools/perf-baseline/src/docx_plain_paragraph_tail_append.rs)
measures one existing public lifecycle:

```text
caller-owned positional ReadAt source
  -> source-backed package open
  -> plain-paragraph snapshot
  -> one isolated copy edit
  -> commit
  -> sequential caller-owned Write publication
  -> drop retained operation owners
```

The measured operation copies source paragraph zero to the source-order tail
slot (`before == paragraph_count`). The harness does not fuse repeated
appends, introduce a section property, or substitute a source-window
representation. The total mode keeps the lifecycle in one region; the phase
mode measures separate non-overlapping `open`, `snapshot`, `stage`, `commit`,
`publish`, and `drop` regions. The phase executions use the same semantic
operation as total mode.

The source provider is an explicit in-memory `ReadAt` adapter owned by the
harness. It exposes stable source identity and records logical read calls,
requested bytes, returned bytes, and request-size buckets. The output provider
is an explicit sequential `Write` sink that retains only scalar counters and a
digest while deliberately handling short accepts. Neither provider performs
implicit filesystem, network, or plaintext-temporary-file work.

Corpus construction and correctness oracles occur before the measured
lifecycle. The retained corpus records an independently constructed expected
tail XML, semantic reopen and paragraph order, raw identities for untouched
members, physical member order, and exact inverse restoration. The measured
source and candidate archives remain owned by the corpus so the operation does
not rebuild its inputs between samples. This is a stated ownership boundary,
not an assertion about resident-memory performance.

## Matrix of accepted obligations

| ADR | Baseline obligation | Actual 0479 evidence and boundary | Status |
| --- | --- | --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md:18) | Correctness and safety win conflicts; specialized low-level paths stay explicit, typed, panic-free, and do not expose raw archive types through ordinary format signatures (`:18-52`). | The harness calls the named `source_backed::paragraph_copy` capability, records typed failures in its contract tests, and publishes through the DOCX/OPC owners. It does not add a raw ZIP type or a general mutable XML escape hatch. | Satisfied for this control scope. |
| [0002](../../../adr/0002-crate-topology.md:6) and [0024](../../../adr/0024-current-topology.md:15) | Concrete DOCX semantics depend downward on shared vocabulary and OPC; physical package ownership remains below the format owner (`0002:8-23`; `0024:17-42`). | `litchi-docx` owns snapshot/edit/patch semantics; `litchi-opc` owns source-backed ZIP indexing, preservation, and publication. The harness observes the public boundary and does not move package IDs or ZIP grammar into it. | Satisfied; ownership remains unchanged. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md:6) | Snapshots are immutable and shareable; edits are isolated; commit returns a new snapshot and reversible patch without mutating the source (`:8-25`). | The lifecycle calls `Snapshot::edit`, exactly one copy operation, `commit`, and publication. Existing tests cover exact source bytes, durable replay, inverse application, stale artifacts, no-op behavior, and failure atomicity (`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:114-328`). | Satisfied by the measured API path. |
| [0005](../../../adr/0005-io-memory-and-performance.md:6) | Use positional `ReadAt` with stable identity/version (`:8-12`), finite limits and explicit scratch (`:19-28`), sequential sinks and incomplete-output reporting (`:30-46`), and evidence before optimization (`:54-61`). | The harness supplies explicit `ReadAt` and `Write` providers, passes finite `Limits`, records logical read/sink observations, and reports the run as baseline/attribution evidence. It makes no constant-memory, speedup, or source-window claim. | Satisfied for measurement; future optimization remains gated. |
| [0006](../../../adr/0006-validation-security-and-compatibility.md:6) | Preserve untouched bytes and lexical details where possible; validation does not mutate; unsupported or unsafe edits remain typed refusals (`:8-23,87-111`). | Corpus checks compare untouched member identities and order, independent main-story bytes, semantic reopen, and exact inverse output. Existing tests cover external relationships, signatures, macros, protection, structured content, unknown namespaces, malformed positions, and sink failures (`source_backed_paragraph_copy.rs:330-473,547-720`). | Satisfied for the narrow plain-story corpus; no broader Office claim. |
| [0008](../../../adr/0008-migration-and-verification.md:1) | Migration slices remain buildable and no support claim is made before required evidence gates (`:8-28`). | This artifact and the neighboring contract/harness review identify a baseline only. Compile/test gates, final capture, and coordinator evidence remain separate requirements; no new production capability is declared. | Satisfied as a bounded evidence slice. |
| [0010](../../../adr/0010-facade-archive-ownership.md:14) and [0011](../../../adr/0011-ooxml-physical-package-ownership.md:15) | Archive mechanics stay below the facade; `litchi-opc` owns physical OPC and public format APIs do not expose concrete ZIP implementation types (`0010:14-27`; `0011:21-37,42-52`). | The harness uses `litchi_docx` and `litchi_opc` only through their existing APIs. Raw archive inspection is confined to the untimed corpus oracle, and no archive type is added to the DOCX public contract. | Satisfied; oracle use does not change production ownership. |

The remaining accepted records remain covered by the unchanged manifest and do
not expand this narrow DOCX workload. This matrix intentionally maps only the
obligations exercised by the source-backed append control rather than claiming
coverage for unrelated iWork, ODF, spreadsheet, presentation, or legacy
binary capabilities.

## Public contract invariants

The current API accepts one exact paragraph copy per `Edit`. The operation
checks `operation.is_some()` and rejects a second successful copy with
`Error::Limit { resource: "operations", max: 1, actual: 2 }`
(`crates/litchi-docx/src/source_backed/paragraph_copy.rs:291-357`). Source
positions are in `0..paragraph_count`; insertion slots are in
`0..=paragraph_count`; the final slot inserts before `</w:body>`
(`paragraph_copy.rs:318-324`). Repeated appends therefore require separate
reopen/snapshot/edit/commit/publication lifecycles and are outside this
single-append baseline.

The source-backed operation is bounded by caller-selected finite `Limits`.
The convenience snapshot uses the default paragraph ceiling, while the large
corpus path explicitly supplies limits that admit the source plus its one
copied paragraph (`paragraph_copy.rs:123-215,941-978`). The harness records
the selected limits in its untimed corpus record. It does not treat a larger
limit as permission to bypass parser, integer, output, or durable-patch
checks.

The accepted parser topology is deliberately narrow:

```text
document -> body -> direct p -> direct r -> direct t
```

`scan_document` and `next_scope` reject unsupported children, wrappers,
unknown markup, MCE, and sections
(`paragraph_copy.rs:1554-1718,1766-1831`). Both a body-level `w:sectPr` and a
paragraph-property `w:sectPr` are existing typed refusal cases; they are not a
final-section placement oracle. The focused test exercises both forms in
`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:429-473`, and the
harness has a corresponding refusal check at
`tools/perf-baseline/src/docx_plain_paragraph_tail_append.rs:1367-1380`.

The successful path retains exact materialized source and projected XML,
checks paragraph readback, applies patches only to the exact source artifact,
and publishes through the OPC preservation path. The relevant implementation
references are `paragraph_copy.rs:1319-1369,1432-1458,394-424,980-1067` and
`crates/litchi-opc/src/source_backed.rs:7705-7727,7764-7829,9251-9578`.

## Ownership and explicit-provider rules

The harness's source and sink are test-owned providers. The source adapter owns
the corpus bytes and supplies positional reads; it does not model physical
disk traffic. The sink accepts sequential writes and records progress without
retaining a second output archive. Source and candidate archives, semantic
records, member identities, and inverse oracles are retained by `Corpus`
outside each measured operation so correctness work cannot be confused with
the public transaction's ownership.

Production ownership is unchanged: DOCX owns the narrow semantic operation,
OPC owns source-backed archive reads and preservation, and the caller owns the
input source and output sink. No production file, public API, or parser rule
is changed by this matrix. No iWork source, fixture, or result is included.

The 0479 source-backed baseline is therefore an evidence control for the
existing materialized lifecycle. It does not establish that a bounded source
window can satisfy `Snapshot::xml_bytes`, exact reversible patch retention,
whole-artifact stale checks, or publication inverse retention. Those would
require a separately reviewed capability and ADR evidence.

