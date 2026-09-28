# 0805 conditional next-trials review

This note records the smallest fresh end-to-end evidence authorized by the
formal 0805 preflight decision. It is a plan only. The decision authorizes
fresh trials; it does not authorize applying the candidate to production or
interpreting the preflight as a public-workflow result.

The follow-up must compare the exact current production source recorded in
`origin.json` with the exact candidate archive in
`candidate/manifest.json`. It must not multiply the 0805 ratios by historical
ratios. The candidate archive remains immutable and retained throughout the
trials. A temporary, source-bound application after the before qualifications
is permitted to build the after binaries. Preflight alone cannot justify
committing that source change; adoption requires every fresh gate below.

## Reusable evidence harness

The latest qualified public PPTX harness is retained in
`../change-0794`. Its probe and fixture lineage is:

* `probe-src/Cargo.toml.template`, `probe-src/Cargo.lock`, and
  `probe-src/src/{main.rs,allocation_metrics.rs,counting_allocator.rs}`;
* `build.py`, `capture.py`, `analyze.py`, `profile.py`,
  `profile_analysis.py`, `quality.py`, `validate.py`, `seal_packet.py`, and
  `custody.py`;
* `plan.json`, `adoption-policy.json`, and `analysis-plan.json`;
* the sealed fixture qualification inherited from
  `../change-0792/seal.json`, as recorded in
  `../change-0794/inheritance.json`.

The probe's unchanged matrix is fifteen rows: `capture`, `commit`, and
`lifecycle` for each of `tiny`, `medium`, `large`, `vendor`, and
`unicode-vendor`. Dimensions are 3×4, 12×8, 100×100, 12×8, and 12×8 slides
and text boxes. Vendor fixtures inject six unknown namespace attributes on
each text tag; ordinary fixtures inject none. This matrix therefore supplies
ordinary and unknown-extension controls, but it does not contain a dedicated
four-attribute fixture. The four-attribute 0805 rows must remain visible in
the decision; the six-attribute vendor rows cannot be presented as proof that
the four-attribute cost is absent.

The next packet must add one public end-to-end variant alongside those fifteen
unchanged rows: `valid-4attr`. It uses the same deterministic 12×8 slide/text
box corpus as `vendor`, inserts exactly four distinct, quoted, well-formed
namespaced extension attributes on each text tag, and retains the complete
semantic text and unknown-extension preservation oracle. The four attributes
must have stable namespace declarations, valid UTF-8 names and values, and no
duplicates, malformed tails, or syntax-error cases. This is a public CRUD
fixture for the valid four-attribute path, not another hot direct-helper
microbenchmark. Freeze its generator source, dimensions, expected semantic
output and preservation contract before the before build. Record the generated
package bytes and hashes in the before-only qualification, independently
verify that new oracle, and freeze it before building or measuring the after leg.

The variant adds three rows (`capture`, `commit`, `lifecycle`) to the existing
matrix. The fresh PPTX campaign therefore has eighteen rows: 216 native
reports / 6,480 measured samples, 72 allocation reports / 216 allocation
samples, and eighteen before-only qualification reports. The original fifteen
rows remain mandatory and cannot be replaced or averaged away by the new
variant. The future packet must add the variant to its packet-local probe and
fixture manifest before baseline compilation; this 0805 packet makes no probe
implementation change.

The 0794 cross-format harness is independently retained in
`../change-0794/{cross-plan.json,cross_build.py,cross_capture.py,cross_analysis.py,cross_root_audit.py,cross-format-plan-review.md,cross-Cargo.lock}`.
It invokes the existing `tools/perf-baseline` binary and does not add a
production benchmark or dependency.

The probe source, cross-format selectors, fixture generators, semantic
readback oracles, and timer scopes are reusable. The drivers are not all
byte-for-byte reusable: `custody.py` hard-codes the owned target, and the
offline analyzers hard-code packet schemas, source lineage, cleanup paths,
bootstrap seeds, and historical packet names. Copy the qualified logic into a
new packet, change those custody and identity fields, and freeze every changed
hash before compilation. Do not reuse the 0794 target, output paths, or
historical timing values. The 0794 supplemental helper-count diagnostic is
also outside this minimum gate; 0805 already supplies the bounded helper
diagnostic and no new microbenchmark scope is needed.

## Conditional execution protocol

Run all work serially from a fresh owned target. Before applying the candidate:

```sh
python3 -B <packet>/build.py before
python3 -B <packet>/capture.py qualification
python3 -B <packet>/cross_build.py before
python3 -B <packet>/cross_capture.py qualification
```

The before qualification is source-bound. Its fifteen inherited PPTX rows must
match the sealed 0792 fixture identities, while the three new `valid-4attr`
rows must match their newly frozen fixture and public semantic oracles. The
cross-format qualification has a separate lineage: its eight result keys and
corpus identities are bound to the packet-local 0794-derived cross plan and
the current before binary; they are not 0792 fixture-seal rows and no 0794
timings are reused. None of these qualification samples is pooled with later
timings. Only after these receipts and the candidate source application have
been independently verified may the candidate quality and after builds start:

```sh
python3 -B <packet>/quality.py
python3 -B <packet>/build.py after
python3 -B <packet>/cross_build.py after
python3 -B <packet>/capture.py native
python3 -B <packet>/cross_capture.py native
python3 -B <packet>/capture.py allocation
python3 -B <packet>/profile.py
```

The production quality gate must use the six qualified 0794 commands: format,
offline locked all-target check, all-feature tests, warning-denied Clippy,
warning-denied rustdoc, and crate-boundary validation over the same fourteen
OOXML/OLE2 packages. The 0805 mirror-crate quality logs alone are not a
replacement for this candidate-source gate.

The PPTX native lane is six alternating process blocks, thirty samples and
three warmups on CPU 12, in the existing order
`AB, BA, AB, BA, BA, AB`. It yields 216 reports and 6,480 measured samples
over the eighteen rows. The allocation lane is two blocks, three samples and
no warmup over the same eighteen rows and yields 72 reports and 216 allocation
samples. Each report
must retain source, fixture, output, and semantic verification identities.

The resource contract compares each paired allocation block median for
`allocation_calls`, `allocated_bytes`, `net_live`, and `peak_above_entry`.
Any increase in any of those four fields is a resource violation; allocation
count alone is insufficient. All raw allocation fields, failed-allocation
checks, and spread flags remain in the report. The large-capture heaptrack
companion has four serial children and eight offline decodes. It qualifies
whole-process conservation, exact owner-subset conservation, and equality of
owner allocation cost to the operation `allocation_calls` counter. It supplies
resource attribution only; it supplies no latency or RSS claim.

## Cross-format regression gate

Use the exact 0794 selector matrix:

```text
--case docx_semantic_open,docx_semantic_full_text,xlsx_open_owned,xlsx_full_cell_scan
--semantic-shape tiny,large
--xlsx-shape tiny,dense-wide
--warmup 3
--samples 30
```

The eight rows are DOCX open and full-text over 24 and 10,000 paragraphs, and
XLSX open and Sheet1 full-cell scan over 3×8×8 and 2×256×256 cells. The
before-only qualification uses warmup 0 and one sample. Native capture uses
the same six `AB, BA, AB, BA, BA, AB` blocks, one fresh process per leg, and
produces twelve reports / 2,880 measured samples. Corpus names, generators,
archive hashes, compression metadata, counts, and semantic shapes must match
between legs; a nonzero harness exit fails the semantic gate.

The cross-format lane is a regression veto, not a benefit requirement. Freeze a
new packet-specific bootstrap seed before its baseline build; retain 10,000
resamples and zero-based ranks 250 and 9,749. Reject the candidate if any row
has paired p50 ratio above 1.05 and bootstrap lower endpoint above 1.0. Run
both `cross_analysis.py` and the independent raw-array
`cross_root_audit.py`; do not pool these timings with PPTX samples or with
0805/0794 history. Remove the two copied cross binaries under a separate exact
identity witness.

## Adoption policy and unresolved four-attribute cost

The fresh public decision retains the qualified 0794 policy: at least one
PPTX `capture` or `lifecycle` row must improve by at least 3% with bootstrap
upper endpoint below 1.0; no PPTX p50 row may exceed 5% with bootstrap lower
endpoint above 1.0; and no paired resource block may increase any guarded
metric. The cross-format veto and all semantic/quality/source-custody gates
must also pass. Every individual row, spread flag, RSS observation, and
profile limitation remains visible; no geometric average may hide a failing
row.

The formal 0805 decision satisfies both required benefit rows and records no
protected consume veto. It retains four non-protected consume flags at four
attributes: `distinct-4`, `duplicate-valid-after-4`,
`syntax-flag-after-4`, and `syntax-equals-value-after-4`. Those flags do not
block these fresh trials under the frozen preflight policy, but they remain
required review rows and do not establish public-workflow benefit. The new
`valid-4attr` fixture supplies the missing end-to-end evidence without
changing, replacing, or averaging away the original fifteen-row matrix. Any
significant regression in the fresh public four-attribute row is handled by
the same per-row 5% veto and remains visible even when another workload row
improves.

After all lanes finish, replay every machine-readable analysis, verify exact
binary identities before deleting the owned target and cross binaries, run the
packet validator with its final seal, and only then stage or adopt the
candidate. Any failed gate archives the candidate and leaves current
production unchanged.
