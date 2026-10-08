# 0973 design precedence clarification

Status: immutable pre-capture binding. This note records how the sealed
`design-v1.md` governs any wording left historical or underspecified in the
metric proposal. It does not change either source document and authorizes no
build, workload, or capture.

The sealed design takes precedence over older metric-proposal wording in these
specific cases:

* The physical leaf proxy is frame 0 of each marker-containing sample across
  all DSOs. A resolved frame is keyed by `<DSO>::<symbol>`; an absent or
  unresolved frame is an explicit `unknown_leaf` bucket. There is no target-DSO
  filter and no reassignment to a later ancestor.
* `D` is total sampled `instructions:u` period, `M` is marker-containing
  period, and `R = D - M` is labeled `marker ancestry not observed`. `R` is
  descriptive residual coverage and is not an outside-operation or causal
  phase bucket.
* Inclusive marker-chain owners are resolved `<DSO>::<symbol>` keys. They are
  counted once per sample, may overlap, are normalized against `M`, and are
  never summed into a phase fraction. The disjoint physical-leaf buckets plus
  `unknown_leaf` conserve `M` per child.

All other denominators, warmup/measured timing scopes, population moments, and
correlation rules remain those of the sealed design and metric proposal.
