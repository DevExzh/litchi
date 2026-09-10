# 0497 production review

This is a read-only architecture review of the DOCX filesystem publication
route. No Cargo command, capture, or source build was run for this review.
The protected `/home/zhuhe/code/litchi-spec-gaps` worktree was not accessed.

## Reviewed inputs

The current implementation and evidence inputs are bound by these SHA-256
digests:

| Input | SHA-256 | Role |
| --- | --- | --- |
| `crates/litchi-docx/src/source_backed/tail_append_stream.rs` | `963373a9d3b826b5f6651067f6f1e12cc4ad00371b29a58ce35d06d6835c8bf3` | 45 added lines for filesystem publication |
| `crates/litchi-docx/tests/source_backed_tail_append_stream.rs` | `06641feb08c539083a4fa7a3f5b9ab0fd50f6c98bc7f957b77d64706b5bd1e6b` | DOCX path, alias, failure, and resource cases |
| `crates/litchi-opc/src/atomic.rs` | `5b714bc59fb4a3f68fd9eef7084c33f3e7a5b3ea300e27e729093f13c90a5739` | Existing OPC sibling-temporary replacement helper |
| `tools/perf-baseline/src/docx_replayable_tail_append.rs` | `b6303b7b353415a4bad997951f98b2dc1c655f795233e10c908f17c684f9b5be` | Frozen benchmark harness input |
| `tools/perf-baseline/src/docx_replayable_tail_append/route_smoke_tests.rs` | `f6a1fb5915c3fdb6cd01acaf4389fe972c926f995a6662e8328cd981665e448a` | Frozen legacy-route smoke tests |
| `tools/perf-baseline/src/docx_replayable_tail_append/publication_route_tests.rs` | `3e2e515d1b4f04748e53bfb5e88461a525ed7ddf80abbea541cde21f2ab529d7` | Frozen after-only route test input |
| `crates/litchi-docx/tests/source_backed_sections.rs` | `a0018a5d3a2d42f9faae2d56e848a4ef5bf4ebce49bcbd570bb01b257f7a60ea` | Corrected section admission/event-limit fixture |
| `docs/performance/results/change-0497/candidate-source.json` | `f4b7438287bbd70524c8a3596648493c3d7238ce7c6e0f32152f5db049a7f49a` | Six-file candidate binding |
| `docs/performance/results/change-0497/plan.md` | `ddec69717ec829403e35a382a93077ac4a044f511808e9ca9a422866aa13ba66` | 0497 scope and validation contract |
| `docs/adr/0005-io-memory-and-performance.md` | `34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770` | Accepted I/O, budget, and publication policy |
| `docs/adr/README.md` | `8cd9b1ec6cded6f3c2a4fbe7f8cd0f4d42f8c259406c7968bbf49fd0ae51b54d` | ADR index binding |
| `docs/performance/results/change-0497/adr-refresh.json` | `5693d4d0227f416b0a0cbaa444102dc375e74884c1ccf0594938ed5d8f91036f` | Recorded ADR refresh manifest |

The current candidate receipts have exit code zero and
`source_unchanged: true` unless stated otherwise. The focused DOCX lane
records 33 passing tests. The DOCX default and all-features all-targets lanes
also pass, as do the six focused OPC atomic tests. The current OPC
preservation lane passes 408 library tests plus its other target groups, and
the DOCX/OPC documentation lane passes 74 tests with 31 ignored and five
doctests. Production workspace check, warning-denied Clippy and rustdoc, and
the allocator-metrics harness (483 passed, one ignored) all pass; harness
Clippy and rustdoc also pass.

The section fixture correction is test-only. Under the tight 16 MiB budget it
now asserts the typed memory admission refusal that occurs before the event
limit; under the wider budget it reaches the intended namespace/event-limit
assertion. The current DOCX default and all-features receipts validate that
ordering without changing the production route.

The earlier `focused1` development lane exited 101 after 31 passes because
its cancellation assertion expected
`OpcError::Execution(ExecutionError::Cancelled)`; the OPC boundary correctly
returns `OpcError::Cancelled`. The retained failed receipt is superseded by
the later 33/33 focused receipt. Earlier section and OPC-preservation failures
are likewise superseded by the corrected section fixture and candidate4
preservation receipt. These historical receipts remain useful audit evidence
and are not current blockers.

The current focused receipt is `candidate3-focused.json`
(`390f3b1bb49d8145aef1fed856a4e7ac59b2d73128fa082dd57e2f5189c308da`);
the default and all-features receipts are `candidate3-docx-default.json`
(`88362ff955c011a1d7262f59bc14d4858920d9b7e5dd220c3307f15a3205dfb9`) and
`candidate3-docx-features.json`
(`a5a6860300ffd1320a7dc8dad8573f880e8485a00eb156864f5222402e63c905`).
The focused OPC receipt is `candidate3-opc-atomic.json`
(`bab54840d80bbd34b595d8d2c49105d9eba8715dcead9c7e89ec3a2f6b78e4f0`),
and the current preservation and documentation receipts are
`candidate4-opc-preservation.json`
(`7992db787c50644535de1fcd00fefd693b11bc159268851e4bd765dfa45e2a11`) and
`candidate4-doc-tests.json`
(`efd0ca296e2f3c663ee47f41fc0a4e51a02e8596196dd89f01a74233bcc6ffb6`).
The allocator receipt is `candidate5-harness-allocator.json`
(`88ed43200b9ddac1445701b2a6609ba09db7fa2da19476a0c8b69d140684e13b`).

## Implementation assessment

`ParagraphStreamPlan::write_to_path` and
`ParagraphStreamCommit::write_to_path` consume the existing plan/commit and
delegate to the existing `publish_stream` path. The callback writes only to
the OPC sibling temporary. Source, candidate, replay, budget, and cooperative
cancellation checks therefore remain in the existing source-part splice path;
the OPC helper synchronizes and persists the temporary only after the callback
returns successfully.

The path tests now cover new and existing destinations, exact bytes and
reopening, durable forward/inverse authorization, the open source path as the
destination, and a Unix hardlink destination. The hardlink case verifies that
replacing the alias leaves the original source inode readable with the exact
source archive while the alias contains the candidate. Symlink and nonregular
destinations, stale sources, cancellation, late managed output failure,
permissions, temporary cleanup, and managed workspace release are also covered
in the current test source. No Windows same-path guarantee is inferred from
these Unix and source comments; a Windows result requires a Windows gate.

The `Committed` behavior is now documented on the commit method: after the
destination has been replaced, a parent-directory synchronization failure
returns `OpcError::Committed`, and this API returns no publication or inverse
product. The local publication value is dropped on that error. That is
consistent with the typed no-blind-retry contract; callers must inspect the
destination and retain any inverse authorization independently if they need it.
The existing OPC helper has a focused injected directory-sync test for this
state, although the DOCX route has no public way to inject that failure.

Source freshness is checked at the end of the existing splice callback, before
the atomic helper performs temporary-file synchronization, persistence, and
parent-directory synchronization. This is the cooperative callback/commit
boundary, not a compare-and-swap or continuous source watch. A source change
or cancellation observed by the callback leaves the destination untouched; a
change after the callback's final fence is outside that cooperative check.
That boundary is stated in the current API documentation and does not require
porting the later iWork-specific revalidation helper into OPC.

The existing OPC helper's trusted-parent assumptions remain inherited behavior:
it does not add adversarial destination identity revalidation at the final
persist point. That is outside this DOCX route's change and is not treated as a
new 0497 blocker. The route does satisfy the scoped callback-failure contract:
the destination remains unchanged and the private temporary is removed when a
source, replay, limit, cancellation, or sink failure occurs before replacement.

The sixth candidate file, `source_backed_sections.rs`, only repairs the test
fixture's admission ordering and is not part of the filesystem publication
implementation. Its passing coverage closes the evidence mismatch that had
previously been reported as a section event-limit failure.

## Root gates still pending

The focused DOCX, OPC, documentation, lint, workspace-check, and harness
receipts above are passing, but they do not complete 0497. Root should still
record and validate:

- final source/build/capture manifest custody after the source freeze;
- benchmark command/schema freeze, after-only route smoke, and the planned
  before/after capture matrix with child terminal, RSS, allocation, read,
  replay, output, and cleanup evidence;
- Python evidence-validator tests and final source/build/capture hash binding;
- the workspace formatting gate. `candidate6-fmt.json` still exits 1 because
  `crates/litchi-keynote/src/document.rs:592` differs from rustfmt; that is
  outside the six-file DOCX candidate and remains an explicit workspace gate;
- cleanup of owned temporary worktrees, Cargo targets, and intermediate output
  only after evidence custody is recorded, followed by the requested rebase
  onto `origin/feat/office-format-completeness`.

There is no new production blocker in the reviewed DOCX route. The scoped
implementation and validation approval covers the current source/candidate
fence, same-path replacement, Unix hardlink behavior, atomic failure contract,
and the receipt set above. Measurements and the final workspace formatting,
cleanup, and rebase gates remain pending, so this review makes no performance
or cross-platform publication claim.
