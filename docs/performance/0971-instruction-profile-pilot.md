# 0971 — instruction-profile pilot fails marker admission

**Status: pilot failed; attribution is unusable.** The corrected pre-capture
tooling review passed and root GO authorized the historical diagnostic, but
child 01 recovered zero samples for the exact native marker. The runner exited
with `pilot_gate_failed` after one child and did not retry. No 12-child cohort,
reducer association, source attribution, qualification, adoption, speedup, or
runtime change follows.

## Binding and pilot result

The run used the retained 0967 ordinary baseline executable in the historical
source environment (SHA-256
`0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb`) and the
`FloatingPictures/e1/default` fixture, pinned to CPU 31. The plan requested 12
serial children, `instructions:u` at period 5,000,000, DWARF callgraphs with a
16,384-byte stack limit, five warmups, and 80 measured samples. The existing
`profiled_measured_case` marker is called 85 times per child, combining the
warmups and measured calls; only the 80 `Instant` samples are the timing
target.

Child 01 was the required pilot. It recovered zero exact marker samples, a
zero marker-period share, and zero lost events. The declared gate required at
least 100 marker samples, zero lost events, and a positive marker-period share,
so the run stopped with exit code 2. The manifest contains one child, records
`no_retry: true`, and does not support a partial 12-child attribution result.

The profiler's whole-process event scope includes the launched child process
tree and excludes profiler bookkeeping. The harness `Instant` scope is the
operation-local region and excludes wrappers and sidecar I/O. The independent
audit confirms the native child returned rc0 with a valid 80-sample report,
matching PID/CPU and 1,387 `instructions:u` samples at 5,000,000 instructions
each (6,935,000,000 total period), with zero lost events. It also confirms the
exact 125-byte marker ELF symbol and FDE, but neither decode contains the
marker: 906 samples have no physical frames, 618 frame records are unresolved,
and only 11 samples contain 23 frames from the native binary, all other
symbols. Attribution is therefore unusable even though native execution and
raw custody pass. The audit does not establish the underlying unwind failure
or assert that the 16 KiB stack setting is its cause.

The sealed offline failure analysis independently re-decodes the profile and
matches the retained decode; native leaf-IP symbolization works, but no marker
return-word or caller ancestry is observed. It isolates no underlying cause.

## Completed retained-work analysis

The separate offline `Float/e1` analysis is complete for the retained 0970
packet: 24 children, 12 pinned and 12 unpinned, with 80 timed samples each.
It finds invariant report work and output shape across all children: 342,528
output bytes and capacity, three routes, 32 stage slots, one active stage per
sample, and stable semantic-change counts. The records expose no stage
duration or subphase counter.

Whole-process instruction counts have descriptive Pearson associations of
0.9553 pinned and 0.8918 unpinned with child p50. The analysis covers all 12
children in each arm and uses no fast/slow cutoff. These associations include
launcher/process-tree work, while the p50 is the operation-local `Instant`;
they do not identify a stage or a source owner.

## Unproven source hypothesis

The source review identifies seeded intermediate `HashSet`/`HashMap` work in
CFB writer and reopen paths as a plausible process-specific cost contributor:
iteration order can change hash probes, B-tree comparisons, sorting branches,
and allocation layout while preserving canonical output. The retained data
does not expose those orders, probes, addresses, or phase counts. Stable bytes,
route records, and correctness gates rule out an observed semantic/output
split, but do not prove this mechanism. The seeded-container explanation
remains unproven, as do scheduler, cache, allocator, frequency, and wrapper
alternatives.

## Disposition and audit boundary

The corrected pre-capture review passed the helper, parser, custody, timeout,
source-binding, and synthetic mutation controls, and root GO bound the exact
historical binary, fixture, CPU, and output roots. The failed pilot is not
reduced or aggregated and no replacement child was run. The next step is
offline callchain forensics to isolate the coverage failure; any future native
attribution must first demonstrate non-empty marker ancestry while reporting
all empty, unresolved, and truncated samples. No rerun, native rebuild, or
production change occurred.

## Evidence

| Evidence | SHA-256 |
| --- | --- |
| [root terminal](results/change-0971/receipts/root-terminal-0971-v1.json) | `0664f420f5d56654e3e61dda83b8d4e2ca06498b3e2a985343a8947d5c18e35a` |
| [failed pilot manifest](results/change-0971/capture/capture-manifest-v2.json) | `bb51f571c33cb6426d3adf75ac91f7ef3f1720a2cec17738832e90205c08673f` |
| [root GO](results/change-0971/receipts/root-go-0971-v2.json) | `ac84a543596842798d09a4727e7e32096bda580270b3e56e37dd55871768f13e` |
| [corrected pre-capture review](results/change-0971/review/independent-final-review-0971-v2.json) | `ab2d1c292e008ea2f398b3dbb9b3e923ab4581207eaa99c2efd659a3703ca92d` |
| [failed-pilot independent audit](results/change-0971/review/failed-pilot-independent-audit-0971-v1.json) | `99eefd86a035d8f6ac87dfa3bd768f1d1d85fee7e152526d2e26e0c40112a2a` |
| [offline failure-analysis receipt](results/change-0971/analysis/instruction-profile-v2-failure-v1/receipt-v1.json) | `5eb4a186159247e2369b595c53a8ce8d11dab9cd5ed56ff899950b65c7bfdae7` |
| [offline failure analysis](results/change-0971/analysis/instruction-profile-v2-failure-v1/failure-analysis-v1.md) | `6aecc34e84edd357401549f8591bd921cf01249b67a22163d0888c2f95658c52` |
| [retained Float/e1 analysis](results/change-0971/analysis/float-e1-work-v1.json) | `1e27a1e84b12d696133bf97fab17e59180e8ba99db64d70fb0ea85f9f61c19fc` |
| [analysis receipt](results/change-0971/receipts/float-e1-work-receipt-v1.json) | `cc9da6e1ef8357691df9b91ed361a9a81f672a92138bedddf85e5e23501ed4da` |
| [temporary cleanup](results/change-0971/receipts/temporary-cleanup-0971-v1.json) | `e25b7eb218960e440b187b4aa73f74acecd507acbe66f1f2b2ccfdad81801832` |
