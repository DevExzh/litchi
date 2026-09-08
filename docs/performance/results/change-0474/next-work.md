# Next measured streaming work

The fresh PPTX matrix disproves constant total operation memory for the tested
path: peak above entry grows from 435,541 to 8,875,092 bytes. It also observes
6,809,604,013 requested allocation bytes in the 8,192-slide operation. These
are two different targets: retained structural metadata and repeated temporary
allocation work. Zero retained output and a single semantic slide do not close
the full explicit-window requirement.

Next, capture stack-attributed allocation and CPU evidence over this same
public case, distinguishing the timed runner from materialized preflight and
measurement observers. Quantify Deflate initialization, OPC/ZIP name validation,
central-directory capacity and transient name preparation separately. Keep
normal latency, allocator callback metrics, profiling totals and RSS distinct.
Do not assign measured peak bytes to a map using only type sizes or source.

Use those profiles to decide the smallest coherent shared transport change.
Potential work includes eliminating repeated compressor/name preparation or
reusing bounded scratch through explicit ownership, and examining an explicit
caller-provided spool for directory metadata. Spooling the directory alone
cannot bound all memory while OPC and ZIP name-validation indexes still grow.
Preserve duplicate/equivalent/ancestor rejection, ZIP framing, output limits,
partial-output errors, determinism, and the existing no-ambient-I/O contracts.
A new provider/API needs an ownership review before production implementation.

A metadata reservation or lower proportional coefficient may be useful if
measured, but it does not establish output-independent total memory. Do not
rename a part-count ceiling or emitted-XML counter an authoring heap window.

Keep fresh streaming creation distinct from logical existing-structure append,
Part addition and arbitrary edits/repackaging. Other CRUD coverage, native
producers, source variants, real parallel scaling and the larger measured edit
bottlenecks remain open under the full non-iWork goal.
