# DrawingML native timing and fidelity return

Independent review of `418cff01db` / assembled `29f072f19c` finds that the protocol cannot yet support final qualification. The actual current compile and ten untimed preflights in 0362dvid remain valid diagnostics, but those preflights generate fixtures and do not exercise the complete fidelity and owner proof paths.

## Required repairs

Timed owner commit paths currently reopen source, copy output bytes and clone patches; replay paths copy output and construct inverse patches. Raw refusal paths still select caps or construct proof/error details during samples. Per-operation details run after timing. These operations must instead use prepared native owners, values, policies and patches; correctness/details must pass before the first warmup. Timed rows measure the named native calls and their documented intrinsic output work.

Normal SVG routes must actually exercise contextual reading and writing, with standalone routes distinguished. Formatcode2 and ThemeFamily need noncanonical retained-source mutation fixtures and exact untouched lexical checks. Complete-family add/replace must verify the requested name/id/variant, and owner proofs must verify expected semantic changes and compare them across baseline/candidate. Presence and unequal hashes are not adequate semantic oracles. Ink's retained-source fidelity and the operation/detail timing descriptions must be corrected in the protocol record.

The coordinator preserves every existing six raw and five owner operation row. The previous plan's `outside_clock` wording conflated replay/refusal proof executions with the named native performance rows. The new plan explicitly separates them: the native replay or refusal call can be measured in its own declared row; the expected variant/reason, exact bytes and semantic proof stay outside the sample clock. This does not permit extra oracle copies, error formatting, refusal assertions or inverse construction in timing, and no threshold or sample requirement is weakened.

## Plan identity

[Immutable qualification plan](results/0362dvie/qualification-plan.json), 38,152 bytes, SHA-256 `e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505`. External identical artifact: `/home/zhuhe/code/litchi-worktrees/.coord/final-acceptance-2026-10-24/drawingml-additional-qualification-plan-20261024T150355Z.json`. It supersedes plan 662a9159; that plan and earlier captures remain unchanged. Future manifest/runner pins must refer to this tracked plan and its exact hash. The final candidate revision remains pending repairs and review.

Independent review confirms unchanged coverage, sample counts, duration thresholds and all four baseline fences. One stale sentence about paired fixture comparison was corrected before committing the plan: that comparison already runs before ABBA and remains mandatory. The first external draft (55bbf4ee) remains unchanged.

Owner/harness integration and raw fidelity/timing are assigned to separate leaves. Their consumers, validators, fixtures, records and hash bindings must migrate together. The fresh all-68 API capture is independent of this benchmark return; its documentation failure is recorded in 0362dvk. No native performance acceptance or push is claimed.

## Harness isolation integration

Reviewed source change `1ad640cc656bee4d563ff694585fc33ff0622a16` merges at `47ea67dc442b8d6ab3c26a5f371c543d7cc2a9ee`. Ordinary performance-harness defaults now select the current RTF and DrawingML adapters. The DrawingML module and binaries require their probe feature, and paired profiles explicitly select it with the appropriate historical/current gates. Generic manifest validation checks the exact feature arguments. The baseline fences remain unchanged.

Independent review finds no blocking source issue and passes sixteen tests from the exact committed source. The composed compile-driver, DrawingML qualification and generic runner suites then pass all forty tests; the [receipt](results/0362dvie/harness-isolation-tooling.json) binds the merged source and log. Actual broad historical/current Cargo validation, the timing/fidelity repairs above and final native qualification remain required.
