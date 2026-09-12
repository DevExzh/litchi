# Change 0521 final source review

Scope was limited to the live source patch, its differential tests, and the
commit allocation instrumentation. No production or harness files were edited
for this review, and no build, test, profile, or capture command was run.

## Reviewed identities

| item | SHA-256 |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/validation.rs` (candidate) | `52e7d2d18f59e716c686f1c4c59b5835632fe6981dc6789ef7b06b555d136c0b` |
| `crates/litchi-xlsx/src/cell_values/validation_borrow_tests.rs` | `161a6ba19eaaeaff2e6e17a0bb77fde2b0e418ecf6784ac87a05afa1abccf61b` |
| `tools/perf-baseline/src/lib.rs` | `4a45ef0f89e18d9242c3b939706001bbf280bc9a1a0ef58aa57a1d4bad916479` |
| `tools/perf-baseline/src/xlsx_commit_metrics_tests.rs` | `6216c258d63a9ef9ccdf8a8ca035522cfce74b79c69476f5baf2ebd3f96c1d88` |
| `docs/performance/results/change-0521/borrow-candidate.patch` | `601df02c4d3e73ad327f6188b61acf29abc15bf6afa6b571aa844baeea3b17ae` |
| `docs/performance/results/change-0521/plan.json` | `711754fb64bcfa7a975509bd961befef6a5d2ddad3d36bf4267dbc04364cd171` |
| `baseline/source-manifest.json` | `e21a054732bf171e423c1b8c087ac2111dcb8f17cb9f653b8e9c72224ff90c37` |
| `candidate/source-manifest.json` | `765fe38ee9ec1ea9293081a3d439f85b8cf836c31b0068157a5f413f0fec6df8` |

The baseline manifest records the prior validation hash as
`b51f7753e71808d2ea0b6d7a4da670ed7c1d098a8b084455aee585922d7e3f5d`.
Comparing the two manifests found one source entry difference: validation.rs
changed to the candidate hash above. The differential test and performance
harness are present identically in both manifests, so the live `#[cfg(test)]`
module does not enlarge the baseline/candidate production comparison.

## Findings

The production change is safe at the reviewed call site. For
`NsReader<&[u8]>`, `read_event()` returns an event borrowing the input slice.
`resolve_event` returns the same event and borrows the resolver only for the
namespace result. The loop resolves before the next reader mutation, so the
resolver's delayed pop for `End` and `Empty` events remains in effect. Parser
error mapping, namespace push/error ordering, stack/name ownership, dialect
binding, and all validation branches remain unchanged. Removing
`Event::into_owned()` and the per-event resolver clone removes the candidate
copying while preserving the observed event and namespace values.

The added differential tests exercise valid aliases and scopes, temporary
prefix rebinding followed by an `Empty` element and a sibling after scope pop,
foreign and unknown namespaces, mismatched/truncated input, malformed
attributes/references, DTD refusal, and workbook errors. They compare the
borrowed implementation with the frozen owned-event loop and check selected
exact errors. This is adequate focused source coverage; the test run remains
unverified here.

The allocation sample brackets the existing `commit_ns` interval: it starts
before the staging loop, includes `edit.commit()`, and finishes before the
publication timer. Publication, reopen verification, and post-commit checks
are outside both measurements. Source-backed evidence records one sample per
retained iteration; the normal test binary emits an explicit unavailable
sample, while the allocator lane can report the existing operation-global
system-allocator counter. The alignment test checks vector lengths, sample
serialization, acquisition-order mapping, and
`open + plan + commit + publication == elapsed`.

The metric is intentionally operation-wide and can include concurrent process
allocations; it cannot isolate `validate_xml` allocations. It therefore
supports a relative commit-lane comparison, but not a validator-only allocation
claim.

No concrete source blocker was found. The following gates are unverified by
this review and must remain open until the existing check/capture workflow
confirms them: formatting, focused and workspace tests, clippy/rustdoc and
boundary/claims checks, sealed baseline/candidate captures, allocation/profile
evidence, semantic/output/resource oracles, and the plan's performance
admission criteria.
