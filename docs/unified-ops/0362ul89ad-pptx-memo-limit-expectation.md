# 0362ul89ad — PPTX slide-root memo limit expectation repair

Date: 2026-10-25. This is a source/test repair only. It does not claim that
the PPTX qualification, the versioned digest transcript, or the final gate
passes.

## Frozen failure and diagnosis

The scoped full-feature capture is frozen at source commit
`ae6f32d202a04e3d6675fa431b2579c0d5f27192` (tree
`361e2d24d8c9d0cb2c838f109646b3ede9150545`; the full tree is recorded by the
qualification packet). Its library-test log is outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/pptx-combined-0362ul89-f5-qualification-ae6f32d202/fullfeatures-tests/0000.log`,
378,736 bytes, SHA-256
`4924bcd19683868b0cc259024ceefdb93f3ba4ef0d3ec7d8653fcf564639fb6a`.

The memo test's resident-size assertions pass. The failing branch constructs
`Limits::new(6, ...)`; the first constructor argument is `max_parts`, not a
codec slide limit. Capture checks the seven strict slide references against
that primary dimension before constructing `SlideRootMemo`, and therefore
correctly returns:

```text
Limit { kind: Parts, observed: 7, maximum: 6 }
```

The test now matches that exact typed error. No memo allocation or resident
byte assertion was weakened, and the capture implementation is unchanged.

## Digest transcript remains open

The other frozen failure is intentionally not repaired here. At transcript
line 159, the generator returns
`Limit { kind: PatchBytes, observed: 137457, maximum: 82578 }`, while
`digest_reuse_lptx0003.golden.txt` still expects
`InvalidLimit { kind: PatchBytes, maximum: 82578 }`. This is a stale golden
expectation: `Limit` represents a measured runtime charge, whereas
`InvalidLimit` is for an invalid caller-selected limit. The expected row also
does not include the current `InvalidLimit` value field. The source code,
revision/publication transcript, and golden file are left unchanged until an
isolated, source-bound regeneration is reviewed.

After the coordinator composes this repair with the remaining PPTX returns,
the approved F5 lane should run this exact generator in a fresh isolated
target, without writing the tracked golden file:

```text
env PATH=/home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/bin:$PATH \
    CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/targets/pptx-lptx0003-golden-regeneration/target \
    /home/zhuhe/code/litchi-worktrees/uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/bin/cargo \
    test --offline --locked -p litchi-pptx --all-features --lib \
    digest_reuse_tests::print_the_golden_transcript -- --ignored --nocapture
```

The runner must bind the command to the composed source commit, tree, branch,
three role-lock hashes, approved runner/helper hashes, and an empty isolated
target. It must retain the transcript stdout and command log outside Git,
record their byte counts and SHA-256 values, and compare the generated rows
with the tracked versioned golden before any golden edit. Review must cover
all wire hashes, revisions, publications, refusal classifications, and row
length; no single refusal-row update is sufficient. Only a separately
reviewed producer-golden commit may then update the tracked golden.

No Cargo, build, rustdoc, test, or benchmark command was run for this repair.
Source inspection and `git diff --check` are the only checks claimed here.
