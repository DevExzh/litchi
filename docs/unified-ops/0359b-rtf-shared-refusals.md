# 0359b — RTF shared refusal contract

- Date: 2026-10-21.
- Status: coordinator follow-up to 0359/0359a; targeted validation complete. The final integrated workspace gate remains pending.
- Scope: owned and logical-tail RTF edit errors, source-backed forwarding,
  unified classification and every matching consumer.

## Error contract

`edit::Error` and `tail_append::Error` now carry
`Refused(litchi_core::edit::Refusal)`. Unsupported source content uses
`UnsupportedContent`; incompatible edit combinations and a missing retained
source needed by the route use `UnsupportedOperation`; unsupported tail
selectors use `UnsupportedTarget`; active protection uses `Protected`.
Operation and reason are fixed static text. The crate retains its own error
types, limit dimensions and patch-conflict variants.

The old refusal variants are removed, including the unused compatibility
`OperationAlreadyStaged`. Private constructors centralize fixed operation
names; they do not add public aliases. Fifteen checked internal-invariant
failures use `edit::Error::Internal` instead of reporting unsupported content.
Examples include a missing staged operation, inconsistent selector cursor or
missing source precondition in an internally constructed operation.

The facade classifies `Refused` by `RefusalKind`. Unsupported content and
operation combinations retain `unsupported_capability`, protection retains
`document_protected`, and patch conflicts retain `invalid_input`. The internal
invariant failures now classify as `internal_error`. Existing protection
metadata remains readable from the document; the refusal itself is content-free.

Producer-paragraph fallback branches keep their precise content-kind guards,
including the fixed group-continuation reason. A protection or operation refusal
cannot accidentally enter those content-recovery branches.

Every native test, facade test and performance-harness matcher is migrated in
this commit. Specific-reason assertions still check the same reason in addition
to the typed kind. The redaction module's error constructors are migrated with
the other consumers; its behavior and the deferred item-22 audit are unchanged.
The glossary registers bans for all ten removed variant names across the two
error types.

## Validation and performance

The complete RTF suite passes: 1,562 tests, zero ignored. OPC-independent native
coverage includes protected no-ops, exact source/inverse output, corpus sources,
producer paragraph recovery, logical-tail publication and durable patches.
RTF and ops all-target/all-feature clippy passes with warnings denied.

The 1,020-file corpus run reports no unexpected fidelity failures. Its golden
report changes only fixed refusal text in 76 RTF rows (63 revert-detail cells
and 19 structural-detail cells, with overlap). File classifications, no-op,
revert and structural outcomes, and unified error codes are identical. The
candidate report was captured outside the checkout, reviewed, and then accepted;
`results/0359b/corpus-review.json` records the comparison and both TSV hashes.

The native probe is `crates/litchi-rtf/examples/refusal_contract_bench.rs`: it compares
1,000-paragraph open, no-op, commit, save, replay and tail append, plus unsupported
content and protected publication. Baseline and candidate use the same source
and separate release target directories; no compiled artifacts are committed.

The facade's eight layout tests and thirteen write tests pass alongside the
unblessed corpus rerun. The performance harness compiles with the migrated
matchers. The `rtf-minimal` glossary profile passes (103 bans, two families,
three immutable owners), and all 27 glossary checker tests pass.

The baseline is `86d708878f`; the RTF implementation is unchanged between that
commit and the candidate's base `dd00ae1611`. Copy the committed probe into a
baseline checkout before building. Both builds use the memory-capped cargo
shim, `CARGO_INCREMENTAL=0`, four build jobs, separate absolute
`CARGO_TARGET_DIR` directories and:

```sh
cargo build --locked --release -p litchi-rtf --example refusal_contract_bench
```

The release binaries are retained outside the repository; `binaries.sha256`
records their identities. The heap runs report identical 3.57 MB peak heap,
177,971 allocations and 1,258 temporary allocations. The raw captures remain
outside the checkout; `heap-artifacts.json` records their names, sizes, hashes
and regeneration commands. The committed heap summaries include RSS with
heaptrack overhead; that value is not used as a heap-allocation comparison.

Three isolated A-B-B-A trials on CPU 7 compare six runs of each binary. Each
run reports the median of 31 samples after three warmups; the table uses the
median of the six run medians. The per-case loop counts, identical probe hash,
source commits and build command are in `results/0359b/summary.json`. The
100 ms process monitor observed no compiler or heaptrack processes. Workers
were held for the measurement and released immediately afterward.

| Route | Before (µs) | After (µs) | After / before |
| --- | ---: | ---: | ---: |
| `open_1k` | 151.931000 | 150.551000 | 0.991 |
| `noop_1k` | 0.046000 | 0.044850 | 0.975 |
| `text_commit_1k` | 214.581000 | 214.006500 | 0.997 |
| `text_commit_save_1k` | 214.026000 | 213.376500 | 0.997 |
| `replay_1k` | 0.011300 | 0.011300 | 1.000 |
| `tail_commit_1k` | 163.441000 | 164.401000 | 1.006 |
| `unsupported_content` | 0.013550 | 0.013995 | 1.033 |
| `protected_commit` | 0.189255 | 0.182950 | 0.967 |

All eight routes are within 1.033× of baseline. The largest ratio is the
unsupported-content refusal, an increase below one nanosecond per call. Native
open/commit/save/replay routes remain within 1.006×; peak heap and allocation
counts are unchanged. These results cover this error-contract change; they do
not replace W8's source-backed replay measurements or the final workspace gate.
