# Next end-to-end investigation: repeated PPTX MCE processing

Read-only priority review by `program_priority`; no PPTX production change or
new PPTX speedup is claimed in 0690. The prior real-deck edit is 24.19 ms versus
an 8.3 ms marker-free control in [0653](../../0653-mce-namespace-emission-rewrite.md).
[0649 call-site evidence](../change-0649/summary.txt) identifies 44 MCE calls /
819,319 input bytes per capture and three whole-slide passes per slide. Commit
recaptures, giving six passes across two captures. These historical findings
must be revalidated at the next baseline before implementation.

Current path reviewed:

- `Package::opened_presentation_transaction` → `opened::capture_internal`.
- `SlidePart::from_part` → `root_name` → `parts::processed_xml` → `process_ooxml`.
- `SlidePart::name` → `c_sld_name` → the same processing path.
- `notes::load_snapshot` → `root_conformance` → `scan_xml`, including slides
  without notes in the measured deck.
- `Transaction::commit` → `capture_with_revision_and_digests` after staging.

Candidate design to price, not authority to bypass checks: a private bounded
capture/transaction context retaining only successful preprocessed bytes.
An identity key needs the raw blob allocation, length and processing profile;
retaining the original Arc prevents address reuse. A blob_arc that does not
alias blob must bypass reuse. Keep all root/name/relationship/notes-topology
validation and limits on every call; do not cache errors or reuse a distinct
semantic-text processing profile. Replaced/removed parts must miss or be
pruned; no public snapshot or global cache should inherit temporary retention.
An explicit aggregate byte ceiling and reservation/fallback policy must be
reviewed against ADR 0005 before code. Allocation identity and replacement
behavior must be proven on the actual package owners, not assumed.

Next measurement: trace stage, URI, input/output bytes, processing profile,
raw identity, borrowed/owned result and repeated identities through existing
0649 real/control phase probes. Price capture-local reuse first, then a
transaction handoff if the additional lifetime is worthwhile. Cover ordinary
save/edit lifecycle, semantic opened phases, no-op, one/two-slide edits and a
notes-bearing fixture. Keep byte digests, semantic reopen and typed refusals;
report retained/peak allocation costs and marker-free controls independently.
No cache implementation is justified until that evidence is captured.
