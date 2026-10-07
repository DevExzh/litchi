# 0964 prospective mechanism and scope

The candidate changes only event transport in XLSX `validate_xml` and
`parse_source_with_observer`. A slice `Reader` and the same quick-xml
`NamespaceResolver` replace `NsReader`. Deferred pops, declaration pushes,
all Start/Empty/End namespace lookups, reader configuration, error conversion,
event spans, limits, MCE equivalence, validator/facts observation, semantic
transitions and authoritative fallback stay in their original order.

Matching a borrowed event for `resolve_prefix` avoids moving the event through
`resolve_event`. Explicit success/error matching retains the validator's
original error text while allowing the compiler to avoid Result payload
transport. No namespace/attribute scan, validation pass, final readback,
allocation policy, source guard, or publication work is removed.

The 0962 frame-pointer executable is an attribution build, not an ordinary
timing control. The 0964 offline replay conserves 7,645 source-admission
samples / 32,395,631,006 period weights inside 12,319 qualified stage samples.
Within that source-admission denominator, exclusive physical self contributes
6.923% for `NsReader::process_event`, 4.451% for `resolve_event`, and 10.729%
for the fused source function. The target self populations are disjoint.
Several leading exact offsets are event-payload load/store instructions.
Sampling skid, inlining and frame-pointer perturbation prevent treating their
sample weights as causal cost or a predicted speedup.

The precedent is mixed: 0810 adopted direct transport for PPTX notes capture;
0823 rejected the same pattern for Scene because its public-workflow benefit
was below the gate. Neither result transfers to XLSX. The first-two attribute
iterator and reduced-readback fusion ideas rejected/deferred in 0963 are not
part of this candidate.

Fresh matched ordinary native and allocator screening is required after exact
source/build binding and correctness gates. The frozen 0964 selector requires
a named useful signal and rejects safety regressions; screening alone cannot
authorize adoption. Any positive screen requires fresh confirmation and
ordinary/fallback control review. A negative screen leaves production unchanged.

Correctness controls compare existing NsReader behavior with the candidate,
including scope restoration, typed declaration/reader errors, BOM offsets,
caps, MCE and provisional fallback. New tests and the full XLSX suite are
quality evidence, never timing observations. No benchmark result was examined
when this hypothesis was recorded.
