# 0797 — first-attribute candidate preflight

This batch tests whether avoiding quick-xml's name allocation for the first
attribute is worth replaying that attribute when a second is encountered.
The candidate must preserve early duplicate rejection through the first 32
names and the existing bounded map fallback. Production is not modified by this
preflight, and no direct-helper result authorizes adoption.

The protocol inherits 0796's one-binary baseline/candidate comparison and adds
distinct-name counts 2 and 3 before freezing any build. Thirty-three cases in
construction and consumption modes use six alternating native block pairs,
30 samples after three warmups, 4,096 repeated hot-input iterations per sample,
and CPU 12. The nearest-rank process p50 feeds a median of six paired ratios;
bootstrap seed 797079 uses 10,000 resamples and ranks 250/9749. Ratios >1.05 with
interval lower bound >1 are diagnostic flags, not public-workflow vetoes.

Two counter repeats use opposite leg orders, one sample/iteration and no
warmup. Five Callgrind counters must conserve in one positive owner dump and
an empty termination dump. Exact named non-inlined owners require one incoming
call and a self-plus-direct-child partition. Guest counters do not measure native
cycles or establish allocation-call counts. Native micro timing and counter
profiles use different iteration counts; no cross-lane timing pooling is valid.

The matrix has 792 native reports/23,760 samples and 264 counter reports/samples:
1,056 reports and 24,024 measured samples. Literal inputs are frozen independently
of the binary listing. Full item/error sequences, error positions, borrowed
values, clone transitions, and terminal fusion are checked outside timing.
The first-error adapter models the checked helper's contract. Checksums are
independently reconstructed for every measured sample.

Separate minimal test workspaces compile the exact helper modules and their
shared unit tests, including canonical-copy parity. These are helper tests,
not full production-crate or public-workflow verification. The direct probe
excludes cfg(test). Opaque construction forces state materialization and is
not representative of every inlined production call; hot-input timing includes
function-pointer/timer and checksum work.

Root alone builds, tests, and captures, all serially. Offline analysis starts
only after captures terminate. Source and review agents do no native work.
Historical architecture/source evidence is hash-bound and no old timing is
pooled. The target is owned by this batch and removed after binary verification.
Full recapture requires a fresh packet/target because drivers refuse overwrite.

To replay retained evidence after sealing:

```sh
python3 -B docs/performance/results/change-0797/validate.py --require-final-seal
```

No baseline or CRUD coverage promotion follows. Public-workflow paired latency,
resource, and cross-format gates remain prerequisites for any later adoption.
OLE2/OOXML remain active, ODF deferred, and iWork excluded.

The frozen preflight advancement rule requires matching semantics, at least 3%
improvement for distinct-1 consumption with interval high below one, and no
consumption diagnostic regression. Construction flags remain diagnostic. Failure
archives this candidate without advancing it to public-workflow trials; success
only justifies those additional trials and cannot adopt production code.
