# Change 0702 — borrow inherited MCE namespace scopes

`performance_claim: none`. Disposition: **retained with scoped evidence**. The packet
starts from the retained 0701 first-byte MCE namespace helper and measures one
independent ownership change in `codec.rs`: `Inherited` carries references to
the existing namespace `Arc`s for the duration of one XML start-event instead
of cloning those `Arc`s into a short-lived value. The helper, its early
predicate, parsing semantics, limits, reports, and tests are held constant.

The candidate source change is restricted to
`crates/litchi-ooxml-common/src/mce/codec.rs`. The focused `mce::` test file
is unchanged and both source phases must report 104 passing tests. The packet
records a fresh 602-file baseline source census at HEAD
`72d6f4500d906b9a5d6d4a0e571b0bb840ab1e68`; no Cargo manifest, lockfile,
dependency, public API, test, or probe change belongs to this experiment.

The source audit requires the 0701 helper source and generated helper size to
remain unchanged between phases, rejects `memmem` and the old `windows` scan,
and checks the candidate's borrowed `Option<&Arc<NamespaceLayer>>` shape. The
processor assembly receipt includes the helper and address-bounded marker
search for both phases. The refusal probe keeps its independent `0693-refusal`
transcript identity, and the native edit probe is `0702`.

## Reproduce

Use a disposable checkout with this packet at the same relative location.
Preserve the root `Cargo.lock`, copy `workspace-Cargo.lock` into the
disposable checkout root, and use the Rust toolchain recorded in
`environment.json`. CPU affinity and scratch paths are host-specific. Do not
refresh dependencies while reproducing the packet.

Run the control check and build the baseline while the checkout still has the
recorded HEAD source:

```sh
python3 docs/performance/results/change-0702/prepare-control.py
python3 docs/performance/results/change-0702/check-control.py
python3 docs/performance/results/change-0702/build.py baseline
python3 docs/performance/results/change-0702/build-refusal.py baseline
python3 docs/performance/results/change-0702/build-oracle.py baseline
python3 docs/performance/results/change-0702/assembly.py baseline
python3 docs/performance/results/change-0702/processor-assembly.py baseline
python3 docs/performance/results/change-0702/measure.py baseline
python3 docs/performance/results/change-0702/measure-allocations.py baseline
python3 docs/performance/results/change-0702/measure-refusal.py baseline
python3 docs/performance/results/change-0702/profile.py baseline
```

Apply only the codec path recorded by `source-diff.patch`, format it, freeze
the exact one-path diff, and run the candidate focused census:

```sh
git apply --include=crates/litchi-ooxml-common/src/mce/codec.rs \
  docs/performance/results/change-0702/source-diff.patch
cargo fmt --all
python3 docs/performance/results/change-0702/source-diff.py
python3 docs/performance/results/change-0702/run-focused.py candidate
```

The source diff must contain only the codec path. Both focused receipts must
bind the same unchanged 104-test source and report zero failures. The
candidate build and all timing binaries use the same lockfile and probe tree:

```sh
python3 docs/performance/results/change-0702/build.py candidate
python3 docs/performance/results/change-0702/build-refusal.py candidate
python3 docs/performance/results/change-0702/build-oracle.py candidate
python3 docs/performance/results/change-0702/assembly.py candidate
python3 docs/performance/results/change-0702/processor-assembly.py candidate
python3 docs/performance/results/change-0702/measure.py compare
python3 docs/performance/results/change-0702/measure-allocations.py candidate
python3 docs/performance/results/change-0702/measure-refusal.py compare
python3 docs/performance/results/change-0702/profile.py candidate
python3 docs/performance/results/change-0702/binary-sizes.py
```

The primary native matrix contains 13 workflows over the real, mechanism
control, generated, and notes inputs. It uses four baseline and two candidate
legs in AA/ABBA order, with 100 samples and five warmups per process. The
refusal matrix contains ten cases with the same schedule. Allocation runs are
separate diagnostics and are not folded into native timing claims.

Run the exact differential oracle before secondary timing controls. It covers
192 deterministic inputs, five capability profiles, and both binaries (1,920
invocations), comparing output bytes and lengths, `Cow` ownership, complete
reports, and typed/debug error identity:

```sh
python3 docs/performance/results/change-0702/oracle/corpus.py \
  /path/to/litchi-0702-bin/baseline-oracle \
  /path/to/litchi-0702-bin/candidate-oracle \
  /path/to/checkout/test-data \
  /path/to/packet/oracle-results
python3 docs/performance/results/change-0702/measure-oracle-controls.py
python3 docs/performance/results/change-0702/measure-oracle-real-controls.py
python3 docs/performance/results/change-0702/measure-oracle-declaration-controls.py
python3 docs/performance/results/change-0702/measure-marker-controls.py
```

The three XML control families and the five marker controls use independent
300-sample, ten-warmup AA/ABBA measurements. Keep Cargo builds, profiles, and
timing lanes disjoint. Derive summaries only after all raw measurements are
terminal:

```sh
python3 docs/performance/results/change-0702/summarize.py
python3 docs/performance/results/change-0702/summarize-allocations.py
python3 docs/performance/results/change-0702/summarize-refusal.py
python3 docs/performance/results/change-0702/report-metrics.py
```

The seven integration gates and six evidence gates are mandatory and run after
the initial measurement lane:

```sh
python3 docs/performance/results/change-0702/run-integration.py
python3 docs/performance/results/change-0702/quality-summary.py
python3 docs/performance/results/change-0702/run-evidence.py
```

Only after all seven integration and six evidence rows are terminal, run the
isolated follow-up. It writes a separate tree and leaves the initial matrices,
profiles, oracle results, and marker controls untouched:

```sh
python3 docs/performance/results/change-0702/measure-followup.py
python3 docs/performance/results/change-0702/audit_followup.py
```

The follow-up repeats seven native cases (`one-real`, `noop-real`, `two-real`,
`two-generated`, `one-control`, `one-generated`, and `one-notes-poi`) and
the complete ten-case refusal matrix with four ABBA legs, 300 samples, and ten
warmups. It produces 28 native rows (8,400 native samples) and 40 refusal
rows. Its independent audit recomputes statistics, pair deltas, metadata
identities, source and binary hashes, commands, output paths, and sample
indexes.

Run the complete audit before cleanup and again from the sealed packet:

```sh
python3 docs/performance/results/change-0702/audit.py
python3 docs/performance/results/change-0702/cleanup.py --apply
python3 docs/performance/results/change-0702/seal.py
```

The audit requires all 13 native cases, ten refusal cases, 78 allocation
comparisons, 1,920 oracle invocations, five marker controls, the isolated
follow-up, all seven plus six gates, exact source/probe/build bindings, helper
assembly for both phases, and deterministic summaries. Cleanup removes only
the four batch-owned scratch paths: `litchi-target-0702`, `litchi-0702-bin`,
`litchi-0702-profile`, and the generated `marker-control.pptx`; it preserves
the root Cargo lockfile and unrelated scratch.

## Interpretation

Native timers cover capture, working clone, text editing, commit, and apply on
fresh prepared packages. Loading, initial opening, target search, save,
serialization, and semantic oracle work are outside that timer. Profiles have
their own denominator. Allocation requested-byte totals include replacement
sizes for reallocations and must not double-count those bytes.

The host is shared and warm, so this packet makes no cold-cache, quiescence,
RSS bound, native Office, or general throughput claim. Every review trigger
remains in the raw receipts. The marker-control archive changes the namespace
URI and is a mechanism control, not a semantically equivalent document.
Exact oracle parity and focused tests are required before any result can
support retention.

## Retained result

The seven-case native follow-up and ten-case refusal follow-up pass their
independent audit. Real one-edit medians improve 3.72–4.77%, real no-op
4.29–4.43%, and real two-edit 1.15–3.74%. Exact oracle parity holds for
1,920 invocations, both focused suites pass 104 tests, and all seven plus six
integration/evidence gates pass. Allocation metrics are unchanged.

The start-handler reservation grows 128 bytes and tiny marked medians cost
20 ns. One early-name follow-up pair flags mean +5.14%, p95 +55.05% and
p99 +15.40%, despite near-neutral medians; every sample remains included.
Other clone and shared XML tails are preserved in the report. This scoped
retention is not a universal MCE, memory, refusal-path or full-save gain.
See the [performance record](../../0702-mce-borrowed-inherited-after-marker-search.md)
and [independent review](code-review.md) for the full tradeoff.
