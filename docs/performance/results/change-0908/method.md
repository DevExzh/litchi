# 0908 PPTX multi-slide Callgrind reducer

This packet qualifies the instruction-cost attribution of the existing PPTX
multi-slide batch workflow. It does not measure or claim native elapsed time.
The frozen target is the 0907 native binary
`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84`, on CPU
12, with four serial children in this order:

`eager-a`, `source-a`, `source-b`, `eager-b`.

Each child uses warmup `0` and samples `1`. The Callgrind collection contract
is:

```text
--tool=callgrind
--collect-atstart=no
--toggle-collect=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit
--zero-before=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit
--dump-after=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit
```

The owner is
`litchi_perf_baseline::publish_pptx_multi_slide_batch_edit`; its measured
parent is
`litchi_perf_baseline::run_pptx_multi_slide_batch_edit_save`. Both zero/dump
options and the raw trigger must use that exact owner string; the shorthand
`same` is not a valid reproducer option.

For this binary, the scoped positive owner edge must resolve to source
position `51174`, the measured call in `run_pptx_multi_slide_batch_edit_save`.
The expected-output call at `51143` is outside the retained scoped edge. The
reducer records both positions so a future harness rearrangement cannot make
the caller-name check look sufficient by itself.

The analyzer imports the retained compressed-cost reader from
`change-0845/reader_v2.py`. It does not duplicate or approximate the
Callgrind grammar. Every scoped dump must have exactly the `Ir` event, a
positive summary equal to `totals`, one owner symbol, one positive incoming
edge from the measured parent, and an owner self-plus-immediate-callee
partition that reconstructs the summary. The sum of all function self costs
must reconstruct it as a second conservation check. The terminal
`callgrind.out` is retained and must be an empty `Program termination` dump.

Callgrind can produce one or two numbered scoped dumps for this owner,
depending on the emitted call structure and inlining. The analyzer therefore
does not hard-code that count. Run the packet's first-child admission helper
after capture of `eager-a`:

```bash
python3 docs/performance/results/change-0908/admit-first.py
```

The receipt records whether one or two numbered dumps were actually retained.
The original eager-a wrapper descriptor is accepted for that first receipt;
later children may use the v2 semantic wrapper because its command and
Callgrind options are unchanged. The reducer also supports `--first-child`
for a self-contained admission when that helper is not used. Only after that
pass, reduce all four children:

```bash
python3 docs/performance/results/change-0908/analyze.py
```

Full reduction requires the passing first-child receipt, binds it to the same
plan, and requires every remaining child to have the qualified count. It
revalidates each start/command receipt, exact argv and fixed CPU, helper/plan
descriptor, every command artifact hash, and the terminal dump. It also
revalidates each corrected semantic receipt against its retained 0907 native
reference: corpus, sink, output digest, source counters, and operation metrics
are checked after reducing constant native 200-sample vectors to one sample.
The original eager-a wrapper failure is accepted only because that child exited
zero and its corrected semantic receipt passes. The result contains per-part
owner edges and immediate direct callees, their `calls` and inclusive `Ir`,
source attribution, conservation checks, and category totals. It does not add
nested callees to the immediate partition. Result paths are create-only; an
existing analysis output causes a failure instead of being overwritten.
The retained stderr logs are custody artifacts only. Valgrind warnings and
repeated PID/runtime probe diagnostics are not added to `Ir` or interpreted as
workload cost.

Direct rows are classified conservatively. A `tools/perf-baseline` source path
or `litchi_perf_baseline::` name is a harness guard. A source path under
`crates/` is high-confidence production evidence. Known production namespaces
without a `crates/` path are retained as `production_candidate`. Runtime,
toolchain, and allocator rows are retained for conservation but excluded from
production totals. Anything else is `unclassified`; its source caveat is
carried in the JSON and it is never promoted to production by the reducer.

The report is intentionally an instruction-count artifact. `Ir` is not native
wall time, hardware cycles, allocation volume, memory use, or a latency ratio;
no such claim may be derived from this packet.
