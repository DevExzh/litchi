# 0898 MCE emitted-name output-span qualification packet

This packet defines the bounded paired qualification for
`mce-codec-emitted-name-output-span-0898`. The frozen trial completed all 72 children / 3,320 samples and failed the primary
2% adoption threshold: paired p50 ratio 0.985065896346, 95% interval
0.977254702311–0.987741000127. Production source is restored; the candidate
overlay and measured evidence remain retained. Independent raw replay and post-cleanup verification pass;
`decision.json` rejects adoption and `owned-cleanup.json` records removal of
the marked source/build roots. No worktree was created.

The adapter reuses the retained 0894 runner, child validator, corpus oracle,
and command construction, plus the retained 0887 reducer. That keeps the
strict 0893 guard contract authoritative. The candidate overlay is exactly:

- `crates/litchi-ooxml-common/src/mce/codec.rs`
- `crates/litchi-ooxml-common/src/mce/tests.rs`

The baseline is the retained 0894 native and allocator binary pair, inherited
byte-for-byte with its release settings: Rust 1.95, opt-level 3, debug level
1, fat LTO, one codegen unit, and two build jobs. The baseline source commit
is `506aa2edae6666d9c76ae6827c2f5941a8e47896`; the evaluation head is the
full `8d517a7861163f04663dae6481e6123b6a314b3c` commit.

## Frozen protocol

The schedule remains fixed at 72 fresh sequential children and 3,320 measured
samples:

| Lane | Children | Warmup per child | Samples per child | Measured samples |
| --- | ---: | ---: | ---: | ---: |
| Native | 64 | 3 | 50 | 3,200 |
| Allocator | 8 | 3 | 15 | 120 |

The four cases are the retained producer XLSX primary case,
marker-free XLSX control, DOCX control, and PPTX control. Native children
alternate AB/BA by block and case; allocator children alternate by block.
All children use the warm filesystem cache and CPU 12. Capture is fail-fast
with no retries or replacements.

The primary gate is candidate/baseline native primary p50 at most `0.98`
with a bootstrap 95% upper bound below `1.0`. The bootstrap uses seed
`898` and 10,000 draws. Pair-ratio and p50-spread thresholds remain
descriptive review flags.

## Freeze input

The root coordinator supplies the future source, quality, and candidate-build
receipts before a plan can be frozen. Use:

```text
python3 -B freeze.py --bindings BINDINGS.json --output plan.json
```

The bindings contract follows 0896. Required names are
`source_manifest`, `source_identity`, `candidate_freeze`,
`baseline_build`, `candidate_build`, `baseline_quality`,
`candidate_quality`, and `guard_reference`; optional review receipts are
retained when supplied. Every binding is a regular non-symlink descriptor.
The candidate build must expose native and allocator binaries, matching
depfiles, build receipts, and ELF build IDs. The candidate source receipt
must bind the exact two-file overlay and keep `production_adoption: false`.

The adapter emits distinct 0898 plan, command, child-receipt, experiment,
reduction, and verification schemas. The final frozen plan is retained as `plan.json`, with passing source, quality,
corpus, build, and methodology receipts.

## Capture and offline replay

Once the plan and sealed binaries exist, the root coordinator may run:

```text
python3 -B experiment.py --plan plan.json --output-root capture-v1
python3 -B reduce.py --plan plan.json --experiment-root capture-v1 --output reduction.json
python3 -B verify.py --plan plan.json --experiment-root capture-v1 --output verification.json
```

The reducer and verifier only reopen retained receipts and capture outputs
after the workload has completed. These five adapters perform no source
build, Cargo invocation, workload launch, worktree creation, or production
edit while preparing the packet.


## Result custody

The primary scan allocator lane records 25,933→17,565 allocation calls,
37,618,900→37,609,931 allocated bytes, and 9,252,430→9,252,427 peak live bytes.
These observations do not override the failed native gate. RSS remains
descriptive because of the retained runner/executable residency limitation.
Initial Name-import compilation failure, ENOSPC PPTX attempt, and versioned
offline replay-helper failures remain visible. Locally retained binaries are
ignored by Git and hash-bound; compressed differential reports, source
overlays, lock, raw samples, logs, and receipts are tracked.
