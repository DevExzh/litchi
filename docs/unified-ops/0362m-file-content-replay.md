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
