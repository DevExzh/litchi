# Next work after change 0464

## Current evidence and limits

0464 now has the formal harness result: R1 and R2 each passed the four
serialized lanes (normal and operation-scoped allocator binaries, bytes and
bounded logical-range providers), with 30 retained samples per lane, for 240
samples total. Binary, source-revision, pair-input custody, output identity,
semantic checks, and the independent package oracle pass. The retained raw
reports are the evidence; the nearest-rank comparison summary is still
pending, so this change makes no speedup or regression claim.

The range control changed the observed logical request count while returning
the same logical payload: total calls rose from 550 to 812, with 93,113
returned bytes in both observations; source calls rose from 195 to 265 and
destination calls from 355 to 547. `delay_us` was zero. These counters show
bounded-provider short-read/request-shape behavior and its small adapter cost;
they do not measure remote latency, physical I/O, or cold-cache behavior. The
formal runner also remains serialized with `workers: 1`, so it supplies no
parallel-scaling result.

The pair is still explicitly `same-source-derived`: an unmodified LibreOffice
QA source and a Litchi-created destination. The final R2
`native_application_roundtrip` receipt passes LibreOffice application save and
semantic readback for the three-slide output: source-backed and eager slide
counts, slide size, and ordered text all match. This is scoped application
save/readback evidence for the derived control, not an independently authored
source/destination pair, Microsoft Office acceptance, rendering equivalence,
or image equivalence. Each post-save source-backed image inventory is typed
`unsafe_edit` and refuses markup-compatibility elements/attributes, so the
pre-save one-image-per-slide inventories (including equal payload identities)
cannot be treated as post-save image proof. No production optimization or
selector/default promotion follows from 0464. The coverage state remains 439
selectors, 36 defaults, 15 categories, 33 representative mappings, 10
measured mappings, and 23 correctness-only mappings; the full non-iWork goal
remains open.

The formal descriptive bundle is retained in `summary.json`; its claims
exclude comparison, optimization, causality, physical/network I/O, native
Office acceptance, and normal-lane allocator totals. The prior ODP results
remain scoped to their own ordinary ODP lifecycle. They cannot rank this PPTX
pair globally, and the pending nearest-rank summary must be available before
comparing the 0464 lanes.

## Priority 1: make the next measured input behavior meaningful

Use the passing 0464 bytes/range matrix as the control for a predeclared
nonzero-latency and cold-source experiment. Keep the same pair, binary/source
bindings, output oracle, and timing boundary so only provider behavior changes.
Vary the range cap and fixed service delay, including a delay representative
of the checked range policy, and run fresh-process/filesystem cold cases beside
warm cases. Report p50/p95/p99, source and destination logical calls,
requested versus returned bytes, short reads, transfer delay, API phase clocks,
allocator counters, process RSS, output identity, and oracle status. Keep input
loading, adapter setup, sink/artifact writes, oracle work, and teardown outside
the API sum as in the formal protocol. Do not describe the current zero-delay
result as high latency or physical cold I/O.

This is a measurement extension, not a production optimization. The derived
pair remains suitable as a deterministic control, while any user-facing claim
must wait for an independent native pair and application readback.

## Priority 2: close one actual checked opened-document CRUD baseline

The largest completion gap is scenario coverage. The checked catalog and
`CRUD_Scenario_Checklist.md` require identity-bound measurements for semantic
opened-document workflows; correctness-only selectors do not close that gate.
The four representative PPTX cross-document rows
(`pptx_source_backed_cross_copy_plain`, its lifecycle form, the media-rich
lifecycle form, and `pptx_cross_copy_media_rich`) and
`odp_existing_append_lifecycle` remain correctness-only. The default manifest's
PPT case is legacy PPT/CFB fresh writing, not this PPTX opened-document
cross-copy contract.

Make one correctness-only opened-document row the next measured catalog case.
The 0464 lifecycle is the most reusable candidate, but the current derived
pair cannot be promoted as independent native coverage. Bind a fixed catalog
identity, producer/version and licensing provenance, source/destination slide
selectors, dependency closure, limits, binary/revision receipt, output
artifact, semantic readback, raw-member preservation, and failure-boundary
checks. Then run the actual default measurement and promote the selector only
when the default runner validates report rows. A static manifest, a generated
per-run corpus, or the current pair's positive-control status is not timing
proof.

If the independent PPTX pair remains unavailable, choose another correctness-
only row only after its inputs can satisfy the same fixed checked-catalog and
full CRUD-oracle contract. Do not relabel existing correctness evidence as
measured.

## Priority 3: measure bounded concurrency after the serial control

Build a separate aggregate workload from independent publications under
explicit execution contexts of 1, 2, and 4 workers (or another predeclared
bounded set). The current pair runner's serialized lane order and
`workers: 1` are not scaling evidence. Keep each publication's source/output
ownership, limits, cancellation, and oracle independent, and report throughput,
p50/p95/p99 tail latency, serial fraction/Amdahl fit, lock or wait data, peak
RSS, and allocator growth. Compare only matched semantic contracts and
providers; do not infer scaling from lower single-lane elapsed time.

## Priority 4: close the native image boundary and obtain an independent pair

The scoped LibreOffice application roundtrip is complete for the derived
control, but its post-save source-backed picture inventory remains unknown.
The next compatibility action is to make markup-compatibility-bearing slides
safe to inspect through the typed image-inventory path, or preserve the
explicit refusal and add a separate lossless package-level image oracle. A
future image claim needs post-save image count, target kind, and payload
identity evidence; the current equal slide text/count/size is insufficient.

Separately, find two distinct original producer packages whose dependency
graphs pass an actual `plan_cross_slide_copy` probe. The static audit in
`native-pair-audit/graph-audit.json` is only a lead list with documented
heuristic false positives and negatives. Preserve producer/version, license,
archive hashes, graph closure, refusal results, and native application
readback. Keep that result separate from Microsoft Office acceptance,
rendering fidelity, and the Litchi-derived positive control.

These steps address the remaining `docs/GOAL.md` requirements for checked CRUD
coverage, caller-supplied range sources, cold/warm behavior, explicit bounded
parallelism, and independent producer evidence. Deletion, structural edits,
merge/split, reversible patch, concurrent composition, security, malformed
input, and broader producer-matrix rows remain open until their own applicable
checklist evidence is bound; 0464 does not complete the program.
