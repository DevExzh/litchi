# Heaptrack format and attribution boundary

`heap_analyze.py` consumes the decoded interpreted Heaptrack stream exported
from `heap-H1/heaptrack.zst` and `heap-H2/heaptrack.zst`. The capture driver
must write the decoded bytes as `decoded.stdout.gz`; the analyzer deliberately
rejects `.zst` so an unavailable decompressor cannot turn a binary trace into
silently accepted text.

The parser follows Heaptrack 1.5.0's own reader. Its pinned primary source is
[`src/analyze/accumulatedtracedata.cpp`](https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/analyze/accumulatedtracedata.cpp): version 3 enables sized `s` strings; `i` records an instruction pointer and its direct/inlined frames; `t` links an instruction pointer to a parent trace; `a` records a requested size and trace index; `+` starts an allocation lifetime; `-` ends the matching lifetime; `c` records a timestamp; and `R` records RSS. The hexadecimal field convention and sized-string parsing are pinned to [`src/util/linereader.h`](https://raw.githubusercontent.com/KDE/heaptrack/v1.5.0/src/util/linereader.h). The parser accepts the interpreted v3 record set (`v`, `X`, `s`, `i`, `t`, `a`, `+`, `-`, `c`, `R`, `I`, `S`, and `#`) and fails closed on malformed fields, unknown modes, out-of-range allocation descriptors, backwards timestamps, duplicate command/version records, and unbalanced lifetimes.

For each selected `+` event, the analyzer resolves the descriptor's requested size and complete trace chain before incrementing calls and requested bytes. The same descriptor is decremented on `-`; concurrent lifetimes using one descriptor are counted independently. `peak_live_bytes` is therefore the maximum of the actual Heaptrack plus/minus event timeline for descriptors with exact streaming `run` ancestry in the scalable path. The large materialization preflight has hundreds of millions of events, so its phase is retained as an explicit unavailable scope instead of being silently treated as zero. The default path skips unrelated events with a chunked C-regex scan and marks whole-process calls/requested bytes/peak as unavailable rather than treating the filtered projection as a process total. `--full` visits every plus/minus event and is intended only for bounded traces. A selected peak is not a modeled XML bound, allocator-internal peak, process RSS value, or the peak reported by a filtered `heaptrack_print` run. Leaked selected allocations remain visible in `live_bytes_at_end` and `outstanding_allocations_at_end` rather than being discarded.

The phase projection is deliberately conservative and visible in the output. A trace containing the exact `pptx_streaming_create::build_corpus` function (or its Rust v0 mangled `pptx_streaming_create12build_corpus` form) is labeled `writer-under-build_corpus`; otherwise one containing the exact `pptx_streaming_create::run` function (or `pptx_streaming_create3run`) is labeled `writer-under-run`; broad helper/module symbols remain `other`. Build-corpus precedence matters because the untimed materialized preflight calls the same writer helper. Category labels use the nearest visible matching frame, from leaf toward caller, so nested OPC, ZIP, or Deflate work is not swallowed by its writer caller: `writer-helper`, `OPC-name-validation`, `ZIP-metadata`, `Deflate`, or `other`. `Deflate` means a visible Deflate/compression frame; it does not claim initialization unless the symbol itself proves that. OPC/ZIP category overlap is resolved by nearest frame and is recorded as such. The raw mangled strings are retained in each stack row, and unresolved chains are counted and listed. These labels are attribution evidence for matching frames, not a claim that every process allocation belongs to the PPTX writer.

`heaptrack_print` exports are optional fallback context. The parser records its top-level calls and whole-process peak rows separately, with `requested_bytes: null` because that text does not expose each row's requested byte total and its filter does not establish a phase-local peak. Raw `+`/`-` data remains the only source for requested-byte and timeline-peak fields.

Example invocation after the capture export:

```text
PYTHONDONTWRITEBYTECODE=1 python3 heap_analyze.py \
  --trace H1=heap-H1/decoded.stdout.gz \
  --trace H2=heap-H2/decoded.stdout.gz \
  --heaptrack-print H1=heap-H1/print.stdout.gz \
  --heaptrack-print H2=heap-H2/print.stdout.gz \
  --output heap-attribution.json
```

The result records compressed input hashes, command identity, record counts,
whole-trace timeline state, phase/category totals, and aligned stack rows for
each lane. Its scope section separately counts primary-phase events,
unknown-scope events, and writer-helper records that lack exact run/build
ancestry (descriptor records in fast mode, allocation events in full mode). It
does not convert allocation samples into operation-only totals;
the trace command includes the complete process, including corpus
materialization. Any residual `other` or unresolved rows must remain in the
report and are not silently assigned to the writer.
