# DOCX structural-scan qualification probe

This standalone binary extends the 0712 public diagnostic matrix to 19 XML
controls and two package corpora. The 200-paragraph generator reproduces
`semantic_docx_bytes(SemanticShape::Medium)` (source text and publication route)
from the ordinary-save harness. NumberedList is opened from its admitted
repository fixture. Both package cases enter the real `document_mut()` facade.

The synthetic XML packages deliberately contain only a main part. Their anchor
cases diagnose parsing and selection, not relationship resolution or save
correctness. `known_body_block_starts` is the inherited lexical prefix probe;
it is not the actual range scanner's result and must not be used as that
oracle. Actual ranges come from the scoped trace. In particular, BOM-relative
positions, nested suppression and foreign namespaces differ from lexical hits.

The inherited `active-offset-count-overflow` is a separate public API check;
its artificial million-offset input is outside the writer boundary. It does
not prove the writer encounters that many ranges. Source-marker text is a
control, not exhaustive collision coverage. This finite matrix does not prove
all XML/namespace/limit/allocation behaviors of a future fused implementation.

From repository root, use the packet's `run.py` and reversible `instrument.py`
recipe. It records a baseline build/run, trace build, and two trace runs.
The instrumented build is an observation tool and must never be used for native
performance comparisons. `CASE0720 generated-setup` identifies generator setup
and is excluded from the representative scanner totals.
