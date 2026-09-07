# Next work after 0465

0465 closes one checked default-baseline gap: the synthetic owned ODP logical
append lifecycle now participates in both full timing runs. The representative
index has 11 measured and 22 correctness-only mappings. The full non-iWork goal
remains open; adding a default row does not certify every checklist direction.

The new complete default capture also supplies a current profiling lead. In R2,
`xlsx_one_percent_commit_save` on `xlsx-dense-wide` has a 406.849 ms p50,
`xlsx_one_percent_commit` 308.386 ms, `xlsx_one_cell_commit_save` 196.845 ms,
and `xlsx_one_cell_commit` 153.289 ms. The large ODP append p50 is 136.844 ms.
These are different operations/corpora; absolute latency identifies candidates
for investigation, not a normalized ranking of user impact or a causal phase
breakdown. Profile the current dense XLSX ordinary one-percent commit/save path
before proposing another optimization. Keep the opened-document preservation,
patch, source identity, and save boundaries matched, and establish phase,
allocation and hardware-counter evidence before changing production code.

Continue promoting the remaining representative correctness-only workflows
only when deterministic checked identities and actual full-run report rows
pass their relevant semantic oracles. The mixed-category rule now supports
incremental progress without promoting unrelated fresh streaming cases. Native
producer coverage, delete/reorder, merge/split, patch composition, security and
malformed-input matrices remain independent requirements; a synthetic measured
row does not discharge them.

Caller-supplied nonzero-latency range input, physical cold/warm I/O, and bounded
worker scaling still need matched scenario evidence. The 0464 pair provides a
zero-delay derived control; 0465 uses owned bytes and one worker. Neither
supplies remote, cold-cache, or parallel-scaling proof.

The separately retained native-image follow-up identifies a possible independent
package oracle for the 0464 LibreOffice resave: direct picture payloads match,
while markup compatibility appears in a transition outside the picture tree.
Implement and negatively test that oracle before claiming its post-save image
facts. Preserve the production picture-inventory refusal and keep rendering,
Microsoft Office acceptance, and independently authored pair claims separate.
