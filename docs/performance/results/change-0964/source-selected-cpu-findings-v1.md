# 0964 retained 0962 source-selected CPU findings

This note is a read-only attribution of the retained 0962 frame-pointer wave.
It uses the exact paired inline/no-inline exports and the exact executable
`xlsx_managed_compression_benchmark.fp-debug2.v1` (SHA-256
`843ae35a0f5131a8cba965f917b2a4c5fccb6926ae3d66862e5f78512820fb33`, GNU
build ID `7eb92ba2f0f8f3b248b77d8691e27b80b114fb22`). No workload, build, or
new perf export was run.

The frozen paired qualification selects 12,319 samples and 52,852,646,778
period weights. Restricting those rows to the exact physical binary frame
`litchi_xlsx::cell_values::snapshot::Snapshot::from_source_selected` gives
7,645 samples and 32,395,631,006 period weights. The six child partitions
match the prior source-selected boundary receipt exactly. Unknown frames,
source/rewrite ambiguity, and wrong-DSO exact-symbol matches are all zero in
this selected population. Physical source locations are retained as observed
evidence only; they do not establish source-file coverage or an owner claim.

The exclusive self rows below use only the deepest no-inline physical frame,
so each selected period appears once and callee work is excluded.

| Physical leaf | Samples | Period | Share of source-selected |
| --- | ---: | ---: | ---: |
| `quick_xml::reader::ns_reader::NsReader<R>::process_event` | 531 | 2,242,851,999 | 6.923% |
| `quick_xml::name::NamespaceResolver::resolve_event` | 341 | 1,442,015,420 | 4.451% |
| `litchi_xlsx::cell_values::validation::worksheet_xml_and_parse_source` | 820 | 3,475,748,072 | 10.729% |

The leading exact instruction offsets are:

| Function | ELF offset | Instruction | Samples | Period |
| --- | ---: | --- | ---: | ---: |
| `process_event` | `+0x92` | `movups %xmm0,0x10(%rdi)` | 278 | 1,176,807,546 |
| `process_event` | `+0x71` | `movups 0x10(%rdx),%xmm1` | 167 | 703,626,971 |
| `resolve_event` | `+0x13b` | `movups %xmm0,(%r14)` | 178 | 753,159,347 |
| `resolve_event` | `+0x175` | `movups %xmm0,(%r14)` | 150 | 634,732,830 |
| `worksheet_xml_and_parse_source` | `+0x140f` | `movups 0x10(%rcx),%xmm1` | 340 | 1,434,492,417 |
| `worksheet_xml_and_parse_source` | `+0x148e` | `movups 0x10(%rcx),%xmm1` | 185 | 782,608,992 |

`worksheet_xml_and_parse_source` disassembles as a clear event-transport
sequence: it calls `Reader::read_event_impl`, calls
`NsReader<R>::process_event`, copies the returned event payload at `+0x140f`,
then calls `NamespaceResolver::resolve_event` and copies the returned event
payload again at `+0x148e`. The two leading `resolve_event` offsets are the
tuple-return writes in the exact ELF, while the leading `process_event`
offsets are event payload copies around its `Start`/`Empty`/other branches.
These rows support a codegen/transport hypothesis; they do not convert sample
period into latency or prove an optimization benefit.

The namespace semantic scan remains a separate scope. `process_event` calls
`NamespaceResolver::push`, and `resolve_event` calls `resolve_prefix`; their
resolver traversal and binding-vector work is callee work and is excluded
from the exclusive self rows above. It must not be inferred from, or added to,
the payload-copy periods.

The narrowly scoped follow-up candidate is the 0810-style direct event
transport applied only to the XLSX source traversal: use a direct
`quick_xml::Reader<&[u8]>` plus a local public `NamespaceResolver` in
`raw::worksheet::parse_source_with_observer`, and apply the same transport to
the authoritative `validation::validate_xml` loop if the shared helper is
not factored. The candidate must preserve the exact `NsReader` state machine:

* consume a pending namespace pop before every read, including EOF and parser
  errors;
* push `Start` and `Empty` declarations before observer/parser checks;
* mark `Empty` and `End` for the next-read pop;
* preserve `NamespaceResolver::push` errors, declaration limits, resolver
  timing, parser/MCE equivalence checks, malformed-input error order, depth and
  event limits, and borrowed event lifetimes.

This is a source-preserving codegen experiment distinct from the historical
PPTX 0812 payload-map-error and later direct-event batches. It requires the
existing differential/oracle controls and an ordinary-build matched screen;
the retained diagnostic does not authorize a source edit or claim a speedup.

Evidence:

* full machine-readable attribution:
  `source-selected-cpu-receipt-v1.json` (SHA-256
  `0636efe9c9c4df2b05e14e2748831f5c91a015c7168b5efbfc658709bf23b5fd`);
* execution history and helper binding:
  `source-selected-cpu-execution-receipt-v1.json` (SHA-256
  `18a725354ba52fb9596039c8a73eab3c5c462e144c4f8f01d315045ec9925c27`);
* final helper SHA-256
  `e5c6a91f8c89e0f91db8eb34cd408d7c809d4f2ecffebeeb9a239e9fb9a7767d`.
