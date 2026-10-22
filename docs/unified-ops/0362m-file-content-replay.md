# 0362m — unified File patch content replay

- Date: 2026-10-22
- Status: implementation and capped validation complete on `uo2/w8-file-replay`;
  native timing remains unqualified and is intentionally left to the
  coordinator after the W4/W4b integration.
- Scope: `litchi_ops::transaction::Patch::apply` and every unified format
  consumer that previously treated a separately opened equal-content package
  as an unconditional patch conflict.

`Patch::apply` keeps the constant-work path for the original immutable owner
and its cheap clones. For a same-variant independently opened candidate it
now streams the candidate and the recorded directional source through
`File::write_to` into a length counter and SHA-256 digest. The candidate is
accepted only when both values match. No second publication buffer, cached
digest field or durable fingerprint is retained. A different format variant,
different security disposition, changed content, or a source that changes
during the proof remains refused as `PatchConflict` or the existing source
error. Source checks still run before the proof and after the retained target
is selected.

The proof is deliberately conservative around protected state. `Signed` and
`Locked` facade variants continue to fail through the existing protection
path, and the helper does not equate signed or encrypted dispositions merely
because their bytes happen to match. Exact unsigned no-op patches remain
byte-identical. The inverse uses the same directional content proof and keeps
the existing source freshness checks; it does not recreate a signed state.

The consumer migration covers DOCX, PPTX, PPT, XLSX, XLSB, XLS, RTF and the
ODF/iWork families exercised by the unified transaction matrix. Equal-content
reopen tests now assert successful forward replay and exact inverse restoration;
changed-content cases continue to assert typed `PatchConflict` behavior. The
DOCX protection matrix also applies no-op patches to signed, locked and
unlocked-encrypted recipients, confirming that no-op status cannot bypass the
existing protection policy; stale-source tests remain in the matrix.

The implementation intentionally has no native timing claim yet. The
coordinator must validate the foreign-content path separately from the
identity fast path and confirm that the streaming proof does not change the
ordinary edit/commit or same-owner replay cost. The capped validation below
records the source-independent correctness checks; native qualification remains
open.

## Facade protection boundary and PPTX source replay

The unified `File` facade cannot currently create a signed-base no-op
`Commit`. `crates/litchi-ops/src/transaction.rs` runs `admit` and
`check_source` before it creates an edit, and its existing protection path
rejects signed input (including a signed no-op) before `File::edit` can return
an editable owner. This is distinct from the leaf contract: a directly opened
PPTX source package admits an exact signed no-op because its package apply
guard refuses signed input only when `!patch.is_noop`; the focused PPTX owner
tests cover that legal no-op and still refuse signed changes. No blanket
signed-content exclusion was added to facade replay.

Commit `d100f18016` extends the PPTX source cross-slide-copy family to accept
an independently opened package when its published content matches the
recorded directional source. Pristine packages use the retained source
artifact fingerprint; edited packages stream the publication into a SHA-256
sink without retaining a second artifact buffer. The deferred result is
rebased onto the recipient's source, limits, execution context and freshness
fences, while the recorded artifact remains authenticated before and after
publication. The inverse follows the patch's directional before/after pair;
it is not selected by a separate `inverse` flag. The focused source-family
test now covers forward replay onto an equal reopened package, exact inverse
restoration from the streamed output, and a same-shaped changed recipient
returning `PatchConflict`.

The already-run validation for this slice was `cargo fmt --all -- --check`,
the focused PPTX source-family test, the 59-test cross-copy engine module,
and the seven-case `source_backed_cross_copy_adversarial` integration test,
all under the capped W8 environment. Native timing for this additional leaf
path remains unqualified.

Validation on the branch used the generated, untracked lockfile
`2c10b15dfc669fffc555913731171ccd507691659d3226c77a162ecafc3d2b72` and the
capped environment from `.coord/w8-file/env.sh`, with
`CARGO_TARGET_DIR=/home/zhuhe/code/litchi-worktrees/uo2-w8-file-replay/target`,
`CARGO_BUILD_JOBS=2` and `RUST_TEST_THREADS=2`:

- `cargo fmt --all -- --check` passed.
- The focused replay integrations passed: `docx_transactions` (22),
  `ppt_unified_write` (14), `pptx_unified_write` (16), `rtf_unified_write`
  (13), `xls_unified_remove` (15), `xls_unified_write` (21),
  `xlsb_cell_writes` (24), `xlsx_cell_writes` (71), and `xls_unified_batch`
  (14): 210 tests in total.
- `crud_matrix` passed its one active test after blessing the intended
  0362m disposition text; one maintainer discovery test stayed ignored. The
  updated rows report equal-content reopened copies as
  `accepts (content-checked)` while changed-content rows remain refused.
- `cargo test -p litchi-ops --all-features --lib --offline --locked` passed
  965 tests, and `cargo clippy -p litchi-ops --all-targets --all-features
  --offline --locked -- -D warnings` passed.
- `cargo check -p litchi-py --offline --locked` passed, and Python AST parsing
  passed for both modified binding test modules. The replay target remained
  below the 20 GiB cap after these checks. No benchmark or executable was
  added to the repository. The final source tip is the commit carrying this
  record and the changes above; the coordinator should record its hash when
  integrating it.

## Reproducible affected-suite receipt (2026-10-22)

The earlier focused receipts did not retain the complete affected-suite
output. This section records a new capped run with its output kept outside the
checkout. The aggregate run used production replay source `51120e13bf` (the
later commits in this branch add only focused PPTX coverage and the typed hash
sink guard):

```text
source /home/zhuhe/code/litchi-worktrees/.coord/w8-file/env.sh
CARGO_BUILD_JOBS=2 RUST_TEST_THREADS=2 \
  cargo test -p litchi-cfb -p litchi-xls -p litchi-rtf -p litchi-odp -p litchi-pptx \
  --offline --locked
```

The command exited `0`. Its 429 `test result:` lines contain 6,303 passed,
zero failed and seven ignored tests. The external raw log is
`/home/zhuhe/code/litchi-worktrees/.coord/w8-replay-validation-20261022-run2.log`
(623,025 bytes,
SHA-256 `469fe3f33051e7dd6b2bdce34dc8dffdf8ddd4b2e2ddc3c6f9db53612759a830`);
the explicit exit receipt is the adjacent
`w8-replay-validation-20261022-run2.exit` containing `0`.

The focused PPTX recipient-policy test was added at source `55d4795feb` and
passed one test with exit `0`; it uses an independently opened recipient whose
`max_input_bytes` is exactly the admitted source length. Its external log is
`/home/zhuhe/code/litchi-worktrees/.coord/w8-pptx-recipient-policy-20261022.log`
(530 bytes, SHA-256
`746d7712c4784571e4bcc539c405acf3941117215348dc2a4047f6ccb81a21e3`). The
full PPTX library suite at that source passed 1,278 tests with one ignored;
its log is
`/home/zhuhe/code/litchi-worktrees/.coord/w8-pptx-final-20261022.log`
(131,716 bytes, SHA-256
`9319241316460ec2a1d8c519cc9563787d6eb4082701c5c891ec76ddc13de26b`).

The focused cross-copy coverage includes
`cross_slide_copy_family_publishes_a_source_backed_package` (same-owner
identity, independent equal-content replay, strict recipient policy, exact
inverse, and changed-recipient conflict),
`source_backed_cross_copy_refuses_stale_foreign_and_same_source_plans`,
`source_backed_cross_copy_reports_source_change_during_output_with_exact_progress`,
and `source_backed_cross_copy_refuses_limits_and_cancellation_before_output`.
The final PPTX source tip is `a1b479f025`; its typed replay-length overflow
test passed separately with exit `0` in
`/home/zhuhe/code/litchi-worktrees/.coord/w8-pptx-overflow-20261022.log`
(517 bytes, SHA-256
`468ad6dd3ca7dbaaf376584eb6d0ac488411cc2a44422c12131e8168a2d78442`). The
test seeds the sink at `u64::MAX`, verifies the typed overflow source, and
verifies that neither the digest nor the byte counter advances on failure.
The sink's ordinary input-length conversion and checked addition therefore
cannot silently wrap; the content proof remains the `(u64 length, SHA-256)`
pair required by 0355a §1.

Commit `df86770cf6` maps that private sink source at the
`replay_content_fingerprint` write boundary to the public
`litchi_pptx::Error::CapacityOverflow` variant. It maps both the direct
`Error::Io` and nested `OpcError::IoError` wrappers, while preserving true
destination/source I/O errors unchanged. Focused boundary tests cover both
wrappers and unchanged sink state on overflow; the unified classifier maps
the structural capacity variant to `internal_error` without fabricating a
limit observation.
The boundary-focused command was `cargo test -p litchi-pptx replay_hash_ --lib
--offline --locked` (four tests, exit `0`), followed by
`cargo test -p litchi-ops row_allocation_commit_and_other_are_internal_error
--all-features --lib --offline --locked` (one test, exit `0`).

All runs used the capped environment above, whose target was
`/home/zhuhe/code/litchi-worktrees/uo2-w8-file/target`; it measured 6.4 GiB
after completion. The lockfile used by the runs is 139,521 bytes with
SHA-256 `2c10b15dfc669fffc555913731171ccd507691659d3226c77a162ecafc3d2b72`.
The source commits in this validation slice are `ebfc35fc31`, `d100f18016`,
`5fbed7708a`, `cdc2696e89`, `f965b0f9c1`, `51120e13bf`, `55d4795feb`,
`a1b479f025`, and `df86770cf6`; no native performance qualification is
claimed here.
