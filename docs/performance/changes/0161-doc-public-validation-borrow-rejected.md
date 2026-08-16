# Native DOC borrowed public-validation input rejected

Date: 2026-08-17

Status: rejected experiment; no production change retained

## Hypothesis and mechanism

Change 0160 identified the initial and final complete public-reader validations
as the largest grouped named phase for the deterministic large and
payload-heavy native DOC corpora. The smallest candidate replaced two
`Vec<u8>` clones used only to construct `Cursor` inputs with borrowed slices.
It retained the strict revision owner, the independent complete public reader,
their order and error behavior, final reopen/readback, exact no-op, patch,
refusal, output-hash, and untouched-stream gates.

The candidate existed only as temporary commit
`00151acdd6ed7303700af0442d76a7160dcf5db0` over clean control
`c0314155d304322044b9e2ef00508b7f6f83bb05`. It changed two expressions in
`litchi-doc::body_text`; it introduced no API, dependency, runtime, cache,
lock, unsafe code, or validation shortcut.

The discarded two-occurrence patch is fully reproducible as:

```text
crate::Package::from_reader(Cursor::new(bytes.clone()))
    -> crate::Package::from_reader(Cursor::new(bytes.as_slice()))
```

## Balanced release result

Separate release binaries ran in `A1 control, B1 candidate, B2 candidate, A2
control` order on the named AMD EPYC 9575F host, pinned to CPU 2. Every leg
used 20 warmups and 500 retained samples for each of tiny, large, and
payload-heavy, for 6,000 retained samples. All 12 report/shape gate sets passed;
source and output digests were identical.

Positive percentages below mean the candidate was faster. Each cell is
`p50 / mean / p95` for the complete measured lifecycle.

| Shape | A1 control -> B1 candidate | A2 control -> B2 candidate | Decision |
|---|---:|---:|---|
| tiny | +3.20% / +2.10% / +0.80% | +3.24% / +4.87% / +13.87% | Improvement is shape-local; p99 directions disagree |
| large | -3.06% / -6.37% / -37.52% | -7.31% / -8.27% / -14.49% | Reject: both directions regress and one p50 crosses the 5% trigger |
| payload-heavy | -0.18% / -0.26% / -1.06% | +2.52% / +1.95% / +0.25% | Reject: lifecycle direction disagrees |

The grouped complete public-reader interval itself improved 4.09%/4.32% p50
for tiny and 3.42%/6.59% for payload-heavy, but regressed 8.18%/11.54% for
large. Same-implementation lifecycle p50 drift was +0.07%/-2.30%/+1.65% for
the tiny/large/payload-heavy controls and +0.03%/+1.72%/-1.08% for the
candidates. Tail drift was noisier and is retained in the summary rather than
hidden.

The narrow mechanism is therefore rejected. One plausible explanation is that
the removed clone also touched the complete input and changed the allocation or
cache state immediately before parsing; this is an inference, not a measured
cache or allocation claim. Whatever the micro-mechanism, the representative
large end-to-end regression is sufficient to discard the candidate under the
program gate. The temporary commit, branch, worktrees, and build targets were
removed; production remains at the control implementation.

## Artifacts and reproduction

The [machine-readable summary](../results/doc-public-borrow-0161-summary.json)
contains every per-leg statistic, paired direction, same-implementation drift,
binary/source identity, and decision gate. The four compressed raw reports are
identified by the [SHA-256 manifest](../results/doc-public-borrow-0161.sha256).

```bash
taskset -c 2 <control-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy \
  --warmup 20 --samples 500 --json /tmp/a1-control.json
taskset -c 2 <candidate-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy \
  --warmup 20 --samples 500 --json /tmp/b1-candidate.json
taskset -c 2 <candidate-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy \
  --warmup 20 --samples 500 --json /tmp/b2-candidate.json
taskset -c 2 <control-binary> \
  --case doc_owner_public_phases \
  --writer-shape tiny,large,payload-heavy \
  --warmup 20 --samples 500 --json /tmp/a2-control.json
```

The control/candidate release binary SHA-256 values are
`e75772030c4cc285cf2c28d00fa03c7eca842946acf44036a40d514ae613d5be`
and `e78bb779e2c83905d498c8df3a72027065a7c43f01fa8b536ca720ba5a850bb9`.

## Claim boundary and next work

This result is limited to the exact in-memory deterministic DOC lifecycle and
the named host/builds. It makes no physical-I/O, allocation, peak-heap, RSS,
cold-cache, filesystem, real-producer, or generic DOC claim. It does not show
that borrowing is intrinsically slow; it shows that this particular naked
clone removal is not a useful production optimization.

Do not retry the same substitution. A future DOC optimization should remove
material work through a private shared physical CFB/stream substrate, fused
fingerprint/copy proof, or parsed-state ownership handoff while retaining both
independent validation layers, and must again pass balanced end-to-end guards.
