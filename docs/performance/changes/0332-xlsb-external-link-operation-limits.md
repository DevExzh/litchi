# Change 0332: XLSB external-link operation limits

## Scope

XLSB external-link parsing, writing, and mutation now have an explicit,
XLSB-owned `ExternalLinkLimits` policy. The policy uses private fields, a
fluent checked builder, finite `DEFAULT` values, and public getters. Builder
failures and runtime exhaustion are typed through `ExternalLinkResource` and
`LimitExceeded`. A zero custom budget is valid when the selected limits are
internally consistent. Mutable accounting counters remain private and are
scoped to the operation that owns them.

## Implementation

- Standalone `*_with_limits` parse, read, write, and apply operations create a
  fresh operation budget. Eager `Workbook` loading shares one budget across
  every external part reached through `SUP_BOOK_SRC`. `Package`, `Workbook`,
  mutation, and durable-patch validation carry the selected policy through
  eager reconstruction; source-backed paths are unchanged.
- Accounting is exact at the controlled boundaries for external parts and
  records, `BrtExtern` caches, entries, matrices, dense cells, UTF-16 units,
  exact UTF-8 string bytes, formula-token bytes, and logically retained model
  objects. Opaque records and bytes are charged when snapshots retain or
  writers re-emit them; standalone/eager semantic parsing skips them. UTF-16
  decoding, byte copies, vector reservations, and model validation scratch
  use fallible allocation paths.
- Writers preflight exact wire size and record headers before allocating
  records or final output. Limits do not activate external targets and never
  silently truncate content. The lossless `BrtSupNameBits` and opaque-frame
  provenance introduced by Change 0331 remains preserved.

## Validation evidence

Validation was serialized with `CARGO_BUILD_JOBS=1` and one isolated Cargo
target directory: `/dev/shm/litchi-0332-target`.

- Focused external-link tests: `42/42` passed.
- Focused workbook tests: `86/86` passed.
- Downstream public-surface tests: `5/5` passed.
- Full library tests: `581/581` passed.
- Integration tests: `126` passed, with exactly
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` skipped.
- Strict Clippy, `RUSTDOCFLAGS="-D warnings"` rustdoc, the minimal facade
  XLSB feature check, crate-scoped formatting, and the diff check passed.

These results make no performance, RSS, global/process, concurrency, or
absolute OOM claim.

## Residuals

- Source-backed external-link semantic parity, cache publication, and
  cancellation are not implemented by this change.
- Independent concurrent operations maintain independent budgets.
- Later caller-owned semantic clones, mutation candidate clones, `Patch`
  inverse images, and the final `Patch::apply_with_limits` output clone are
  outside the operation budget.
- Workbook metadata surrounding `SUP_BOOK_SRC` is outside the external-part
  policy. Its collection growth is fallible and each string is hard bounded,
  but shared raw string-decoder allocations are not charged to this policy.
- The high-level `WorkbookWriter` has no custom profile and uses the finite
  default policy. `From<OpcPackage>` likewise uses the default policy with
  lazy validation.
- Exact allocator overhead is not accounted for.
- The known drawing-corpus test remains skipped as stated above.
- Workspace-wide formatting is not proof because the unrelated
  `crates/litchi/src/presentation/prs.rs` has an existing formatting mismatch;
  the crate-scoped check passed.
