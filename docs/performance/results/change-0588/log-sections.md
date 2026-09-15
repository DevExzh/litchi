# change 0588 — log paragraphs for the coordinator to merge

The four sections below are written for `HOTSPOTS.md`, `GOAL_AUDIT.md`,
`REPORT.md` and `ADR_COMPLIANCE.md`. This change did not edit those files.

---

## For `HOTSPOTS.md`

## 0588 — the MCE codec's per-attribute allocations, and why its namespace emission cannot move yet

Survey item XML-1 (rank 1 of change [0587](0587-remaining-opportunity-survey.md))
is priced, and its **byte-identical** subset is implemented. `start` allocated
two `String`s per attribute and an owned `xml_name::QualifiedName` plus an owned
`Name` per attribute on three paths; `BoundedOutput::reserve` grew by
`try_reserve_exact(additional)`, so a rewritten part reallocated once per written
run — 2,954,106 `__rust_realloc` calls and 30.46% of the codec's instructions on
the survey's worksheet. Attributes are now borrowed from the event with
quick-xml's `Cow`, a new `expand_parts` resolves a qualified name into borrowed
halves with the identical lexical check and the identical error identity, the
writer materializes an owned `Name` only in the rare preservation branch, and the
output buffer grows geometrically **clamped to `max_output_bytes`**. The MCE
codec on a real Excel worksheet falls **723.6 M to 305.1 M Ir per call
(-57.84%)**, the public eager open plus one cell **1,037.8 M to 497.5 M
(-52.08%)** and the source-backed read of the same cell **1,075.3 M to 654.1 M
(-39.17%)**; paired timing gives p50 **-43.88%** and **-34.89%** on those two
public reads against an A/A floor of p50 -0.92%/-0.09% and p99 +0.38%/+0.81% in
the same window. The codec's borrowed fast path and four existing XLSX selectors
on the marker-free harness corpora move less than the floor. Output bytes are
**identical**, proven over 6,964 real fixture parts and 30,000 mutants.
**XML-1's headline — emit only an element's own declarations — is implemented,
measured and withdrawn.** It is worth a further -79.9% of the codec's remaining
instructions (305.1 M to 61.4 M; the worksheet's output falls 3,540,261 to
194,508 bytes and the eager open to 156.0 M Ir), but the redundant declarations
are load-bearing: every element span of the processed buffer is namespace
self-contained, and `litchi-docx` slices inner `w:p` and `w:tbl` byte ranges out
of it and parses them standalone, so the rewrite breaks every
`Paragraph::extensions()` read on a real Word document. Two writers also publish
the processed stream (`litchi-pptx` guides `rewrite_source`, `litchi-xlsb`
`drawing_transfer`). A sweep of all **133** production call sites of the five
codec entry points outside iWork and ODF sorts the exposure into one hard case,
a partial tier behind a single-dominant-prefix fallback, and a publish tier of
public accessors that hand out processed slices verbatim; five sites already
re-inject declarations correctly and are the template for a fix. **And the
property is already conditional**: the codec borrows the input unchanged when a
part carries no MCE namespace, and a probe against the untouched base shows
`Paragraph::extensions()` **already refusing** with the identical error on the
one `test-data` `.docx` (of 62) whose `document.xml` has no MCE namespace. The
rewrite universalizes a pre-existing `litchi-docx` defect rather than creating
one, which is the strongest argument that the fix belongs in the slicing
consumers, not in the writer.
`performance_claim: none`.
[Change and limitations](0588-mce-codec-namespace-emission.md);
[evidence](results/change-0588/README.md). OLE2/OOXML remain active; ODF is
deferred until completion and iWork excluded.

---

## For `GOAL_AUDIT.md`

## 0588: the cheapest read-side transform in the OOXML path, and a contract nobody had written down

`docs/GOAL.md`'s optimization order puts "eliminate unnecessary allocation and
copying" ahead of layout and algorithms, and the MCE codec had both in the same
function: two heap allocations per attribute, an owned `QualifiedName` per
qualified name on three paths, and an exact reservation per written run.
[0588](0588-mce-codec-namespace-emission.md) removes them **without changing a
single output byte** — 57.84% of the codec's instructions on a real Excel
worksheet, 52.08% of a public eager open plus one cell, p50 -43.88% wall clock —
which is the cleanest form the decision rules ask for: a change that cannot move
a refusal, cannot move a limit and cannot change what any consumer sees.

The more durable result is what stopped the rest. 0587 ranked XML-1 first and set
the precondition "confirm the processed stream is never published". Two writers
publish it, but the binding constraint turned out to be different and unnamed:
because the writer re-declares every in-scope binding on every emitted start tag,
**any element span of the processed buffer can be sliced out and parsed
standalone**, and `litchi-docx`'s paragraph, table and row views do exactly that.
That is a contract the codebase relies on and no record stated. It is now a test
(`every_element_span_of_the_output_is_namespace_self_contained`). It is also
already broken where the codec's fast path applies: on the one `test-data` `.docx`
whose `document.xml` declares no MCE namespace, `Paragraph::extensions()` already
refuses at the base revision, so `docs/GOAL.md`'s correctness-first rule is
engaged independently of any optimization. The audit's
"Apply layout, cache, or SIMD tuning only from measured hot loops" row gains a
sibling requirement from this: an internal representation that several crates
slice needs its slicing contract written down before it is optimized, or the
optimization is discovered by a single consumer test. OLE2/OOXML remain the
active priority; ODF stays deferred and iWork excluded.

---

## For `REPORT.md`

## 0588 — a 52% read with no byte changed, and a 92% one that cannot land yet

Retained, partially implemented. The MCE codec's per-attribute `String` pairs,
its per-qualified-name `QualifiedName`/`Name` allocations and its exact
per-run output reservation are replaced by borrowed attributes, a borrowed name
resolver with identical lexical checks and error identity, and geometric growth
clamped to `max_output_bytes`. The codec falls 723.6 M to 305.1 M Ir per call on
a real Excel worksheet (-57.84%), a public eager open plus one cell 1,037.8 M to
497.5 M (-52.08%) and a source-backed read 1,075.3 M to 654.1 M (-39.17%);
paired p50 improves 43.88% and 34.89% against an A/A floor of p50 <= 0.92% and
p99 <= 0.81% in the same window, and four existing XLSX selectors on the
marker-free harness corpora move less than their own floor. **Output bytes,
`Report` counters, the borrow-versus-own decision and every refusal identity are
unchanged**, proven by a differential over 320 fixtures and 6,964 XML parts in
both a byte-exact and a namespace-resolving mode, and by 30,000 mutants of which
16,661 refused, all with zero mismatches. Survey item XML-1's headline change —
emit only an element's own declarations, hoisting a dropped wrapper's onto its
first emitted child — was implemented and passes the same oracles, and is
**withdrawn**: the redundant declarations make every element span of the
processed buffer namespace self-contained, and `litchi-docx` slices inner spans
out of it and parses them standalone, so the rewrite breaks
`Paragraph::extensions()` on real Word documents; two writers also publish the
stream; and a sweep of all 133 production call sites of the five codec entry
points outside iWork and ODF sorts the rest of the exposure into a partial tier
behind a single-dominant-prefix fallback and a publish tier of public accessors
that hand out processed slices verbatim. A probe against the untouched base shows
that same public call **already refusing**, with the identical error, on the one
`test-data` `.docx` of 62 whose `document.xml` carries no MCE namespace — because
the codec then borrows the input unchanged — so the rewrite universalizes a
pre-existing `litchi-docx` defect rather than creating one, and the fix belongs in
the slicing consumers. Its patch, its two differential reports, the probe and its
measurements (a further
-79.9% of the codec, output 16.9x to 0.93x of input) are retained. All gates
pass, including 246 test binaries and 5,402 tests across the eight OOXML
consumer crates; a pre-existing `litchi-iwa` example compile failure blocks
`cargo test --workspace` and is reproduced on the untouched base. No latency,
allocation, RSS or cold-cache claim is registered. OLE2/OOXML remain active; ODF
is deferred until completion and iWork excluded.
[Change and limitations](0588-mce-codec-namespace-emission.md);
[retained evidence](results/change-0588/README.md); `performance_claim: none`.

---

## For `ADR_COMPLIANCE.md`

## 0588: nothing published, nothing relocated, and two publication paths reported rather than changed

[0588](0588-mce-codec-namespace-emission.md) changes allocation strategy inside
the private MCE writer of the crate that owns markup-compatibility
preprocessing. No public type, signature, export or crate dependency changes —
`Name`, `Capabilities`, `Limits`, `Report` and `Output` are untouched — so ADRs
0002, 0003, 0010, 0011 and 0024 are not engaged, and the change is invisible
above `litchi-ooxml-common`.

ADR 0006 governs the substance, and the change is byte-identical against it:
every check keeps its position, its identity and its ordering; duplicate
attribute detection, value decoding and normalization, DTD/PI rejection, custom
entity rejection and `MustUnderstand` all run where they ran; and the new
`expand_parts` reproduces `xml_name::codec::parse` exactly, including its
`invalid QName: invalid XML QName '<value>'` message, so no input changes
category. The output bound is unchanged to the byte: `BoundedOutput::reserve`
still refuses on written length before reserving, and the geometric target is
`min(2 * capacity, max).max(len)`, which never asks the allocator for more than
the limit already admits. Every reservation stays fallible and maps to
`Error::Allocation` with the same `resource` strings; no `unsafe` is added.
ADR 0005 is untouched: no read, positional source or mandatory validation moves.

Two boundary facts are **recorded, not worked around**. First, the codec's
processed stream is published by two writers —
`litchi-pptx`'s `presentation_properties::metadata::guides::codec::rewrite_source`
(the whole `presentation.xml` on a changed guides edit; its no-op path
deliberately skips the helper, so an exact no-op stays exact) and `litchi-xlsb`'s
`cell_values::drawing_transfer` (an anchor fragment cut from the processed source
drawing part). Whether a read-side transform's output belongs in a published part
is a contract question this record raises and does not answer. Second, an element
unwrapped by `mc:ProcessContent` cannot declare a prefixed namespace
(`NonConformant("unbound prefix xmlns")`) because the unwrap check expands every
attribute name including `xmlns:z`. That refusal is pre-existing at the base
revision; it is now pinned by a test so the borrowed resolver cannot move it, and
it is deliberately not fixed, because fixing it would move when a refusal
happens.

A third boundary fact is recorded for correctness rather than compliance:
`litchi-docx`'s `Paragraph::extensions()` already refuses
(`InvalidFormat("Word extension XML must have one [112] root")`) on a
`word/document.xml` that carries no markup-compatibility namespace, because the
codec then borrows the input unchanged and the sliced `w:p` span was never
namespace self-contained. Measured on the one such fixture of 62 against the
untouched base checkout. It belongs to `litchi-docx`, not to this crate, and is
reported here rather than fixed.
