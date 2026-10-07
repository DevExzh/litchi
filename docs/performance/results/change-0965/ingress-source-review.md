# 0965 — 0964 ingress source-path review

**Disposition: the 0964 candidate loops do not execute inside `ingress_ns` in
the frozen managed benchmark call path.** This is a source-backed callgraph
result, not a claim that the measured ingress-tail flags are harmless. The
candidate can still change the release image and runtime state enough to affect
a short constructor phase indirectly; that needs a constructor-only diagnostic.

This review is read-only. I read `docs/GOAL.md`, the accepted ADR index and the
accepted constraints relevant to ownership, validation, finite resources,
measurement and verification (0001, 0002, 0003, 0005, 0006, 0008, 0024,
0030, 0031 and 0032) before inspecting source. No build, benchmark, workload
or production edit was run. The existing unrelated worktree changes remain
untouched.

## Reviewed candidate

The frozen 0964 candidate is at
`/home/zhuhe/litchi-performance-0964/source-candidate`; the baseline is at
`/home/zhuhe/litchi-performance-0964/source-baseline`, both based on the
source commit recorded in `docs/performance/0964-xlsx-direct-event-transport-trial.md`.
The only production changes are the two transport loops below. The candidate
file hashes are retained by the 0964 source review:

| file | candidate SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `f588cd2b1983c96df3d58a07c21ba5b4c14a57e776e35d5ac152159d788cb39e` |
| `crates/litchi-xlsx/src/raw/worksheet/codec.rs` | `0ed7754f2b6ebb9eb16fcdb9896d0383c471709af2579770138936147e6bebaa` |

`validate_xml` replaces `NsReader` event transport with `Reader` and a local
`NamespaceResolver` (`validation.rs:333-377`). The raw
`parse_source_with_observer` loop makes the same substitution
(`codec.rs:1633-1750`). It retains the deferred pop, declaration push,
namespace resolution, event spans, observer order, parser order, MCE proof and
fallback boundaries. These are the two loops being attributed here; this
review does not reinterpret the 0964 static equivalence result as a timing
result.

## Timed call path

The benchmark starts the constructor timer at
`tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs:416-425` and
stops it immediately after
`SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context`
returns. The relevant path is:

```text
[ingress_started]
SourceBackedEditor::from_read_at_with_limits_and_cache_limits_and_execution_context
  -> SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context
    -> SourceBackedPackage::from_read_at_inner
       - source.version(), source.len(), input limits and freshness checks
       - exact managed SourceReadPolicy (no read-ahead)
       - indexed_archive_with_limits(SourceReader { ... })
       - PackageReader::source_catalog(SessionedArchive, limits)
       - OPC content-types/relationship catalog projection, cache and package
  -> SourceBackedEditor::from_source_backed_package
       - package.check_execution()
[ingress_ns stops]

stage_timed_commit(&editor, updates)
  -> editor.edit_sheets(...)
    -> MultiSnapshot::load_source_backed
      -> load_source_catalog
         -> workbook.data(), validation::workbook_xml, raw::parse_catalog
      -> Snapshot::from_source_selected
         -> worksheet.data(), source_stream_admission
         -> Some(admission): worksheet_xml_and_parse_source
            -> raw::parse_source_with_observer (0964 candidate loop)
         -> None: validation::worksheet_xml, parse_refusing_repairs
```

The editor wrapper and constructor delegation are in
`crates/litchi-xlsx/src/cell_values/source.rs:425-445`; the managed OPC
constructor and its `from_read_at_inner` body are in
`crates/litchi-opc/src/source_backed.rs:6912-6924,6989-7221`. The post-timer
selector and worksheet path are in
`crates/litchi-xlsx/src/cell_values/snapshot.rs:205-255,467-535,544-570,2329-2368`.
The benchmark report itself defines `ingress_ns` as the direct constructor and
`edit_commit_ns` as selector staging, cell sets and commit
(`xlsx_managed_compression_benchmark.rs:702-705`).

`from_read_at_inner` receives an `Arc<dyn ReadAt>`, read limits, cache limits
and the managed execution context. It reads source metadata, indexes the ZIP
archive and parses the OPC structural catalog. It does not load or validate
the workbook or worksheet SpreadsheetML parts. In particular, the 0964
`litchi_xlsx::cell_values::validation::{validate_xml,...}` and
`raw::worksheet::codec::parse_source_with_observer` symbols have no source
call edge from this constructor path. The first `validation::workbook_xml`
call in this flow is in `load_source_catalog`, after the timer has closed.

## Candidate inputs, state and ordering

The two changed loops require decompressed XML bytes and XLSX semantic state,
which the constructor does not have:

| loop | inputs and state | ordering that matters |
| --- | --- | --- |
| `validate_xml` | `content: &[u8]`, `XmlOwner`, `Validator` depth/element stack/dialect/copied-subtree/error state, local `NamespaceResolver` | pop before the next read; push `Start`/`Empty` before resolution and observation; keep `End` in scope for its iteration; finish at EOF |
| `parse_source_with_observer` | worksheet `content`, `SourceAdmission`, strings closure, observer, `Repairs::Refuse`, `ReaderOrigin` BOM offsets, event count, parser stack/root state and optional MCE equivalence | compute source span, install namespace scope, count/cap, resolve namespace, run MCE equivalence, call validator/facts observer, then parser transition; provisional or reader failure falls back to the established authoritative pass |

The worksheet shared traversal is conditional on
`source_stream_admission(...)` returning `Some`; otherwise the candidate raw
loop is not entered. The candidate validation loop is still used by workbook
and worksheet validation after part data is loaded, including the fallback
path. None of these inputs, resolver states or parser callbacks exists during
the timed OPC constructor.

## Codegen check and causal limits

The read-only receipt
`/home/zhuhe/litchi-performance-0964/review/codegen-candidate-v1.json`
passed its narrow check: direct `NsReader::process_event` and
`NamespaceResolver::resolve_event` edges disappeared from the two changed
callers, with named local resolver calls observed. The receipt explicitly
does not infer dynamic execution or removal of namespace work. The ordinary
baseline binary is 43,259,328 bytes
(`5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3`), and
the candidate binary is 43,270,848 bytes
(`d21863a6613800391a01914865fc4eaecdf54ad3061447be8b6be7d47d0dc910`).
This supports a changed process image, not a proof about the constructor's
dynamic instruction stream. The identical-binary control was intentionally
rejected by the checker and is not evidence for either outcome.

The direct source hypothesis is therefore ruled out: the two candidate loops
cannot be the instructions that the frozen `ingress_ns` timer directly
measures. The observed ingress flags still admit these causal hypotheses:

1. **Release-image layout.** The candidate changes code size, inlining and
   link placement in a fat-LTO executable. Alignment, I-cache/BTB placement,
   register allocation in shared code and relocation layout can change the
   constructor even when its logical operations and source inputs are the
   same.
2. **Short-phase runtime variance.** Constructor samples are roughly a
   80–110 microsecond phase. Scheduler interruptions, frequency state, cache
   warmth, page faults, allocator state and fresh process address layout can
   move p95/p99 ratios materially. This is a hypothesis about the measurement
   environment, not a dismissal of a real regression.
3. **Fresh-state and wrapper effects.** Each timed iteration creates a fresh
   managed context and `SourceBackedPackage`; source version/length, ZIP index,
   object reservations and package drop/worker state are all inside or adjacent
   to the measured region. Candidate code elsewhere can alter heap and process
   image state that these operations observe.
4. **Unseen shared codegen dependencies.** A changed quick-xml use site can
   alter monomorphization, section placement or identical-code folding for
   shared dependencies without creating a logical call from OPC ingress to
   XLSX validation. Only dynamic counters and a controlled layout comparison
   can separate this from runtime variance.

## Historical ingress evidence

Earlier rejected candidates changed later phases, yet their ingress controls
also moved. This establishes recurring sensitivity, not that all ingress flags
are noise:

| trial and changed phase | historical ingress observation |
| --- | --- |
| 0959 numeric ownership / `Number(Box<str>)` | External comparison `analysis/comparison-v3.json` reports medium-one-edit ingress p50/p95/p99 of `1.0386/1.0517/1.0287`; dense-sparse one-percent is `0.9885/0.9785/0.9739`, and medium one-percent is `0.9861/0.9735/0.9338`. The one-edit p95 gate and multiple leg flags remained open even though the candidate was numeric storage. |
| 0960 worksheet capacity hint | The retained record reports one-edit ingress p95 `1.0657` and p99 `1.1703`; this candidate added a body recognizer and capacity hint after admission. |
| 0961 edit-staging capacity | The screen reports dense-sparse ingress p95 `1.0594` and medium-one-edit ingress p99 `1.0532`, while the candidate changed `MultiSourceEdit::apply_batch`. |
| 0964 direct event transport | Epoch 2 reports ingress-p99 medians `1.0247` (dense-sparse), `1.1075` (medium one-edit) and `1.1116` (medium one percent); the latter two triggered the existing safety rule. |

The repeated post-constructor candidates with ingress flags make phase noise,
layout and process-state effects plausible. They do not invalidate the 0964
safety rejection or prove that its two flagged tails lack a candidate-induced
component.

## Concrete next diagnostic

Build one temporary constructor-only benchmark from the same frozen baseline
and candidate sources. Keep the source corpus, release flags, host affinity,
fresh-process schedule and managed arguments identical; start and stop the
timer around only the editor constructor, and place editor/package drops
consistently outside the timer. Run serial pinned ABBA/BAAB blocks with the
existing warmup/sample policy. Add a direct `SourceBackedPackage` timing
split if the wrapper contribution is material.

Around that constructor-only region, collect `perf stat` counters (cycles,
instructions, branches, branch misses, cache misses) and a read-only
`perf record -g` callgraph. The callgraph should verify that no
`cell_values::validation` or `parse_source_with_observer` frame occurs during
ingress, while the counters show whether the candidate changes constructor
work. If the timing difference persists, compare a layout-only control with the
same release/link settings and measure allocator/page-fault and CPU-frequency
state. If it disappears when edit/publication are removed, inspect timer
boundaries and pair scheduling before attributing the 0964 flags to the
transport.

Until that diagnostic distinguishes these causes, retain the 0964
`reject_screen`, the `1.05` ingress-tail gate and the unchanged production
source.
