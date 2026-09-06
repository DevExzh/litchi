# Retained development failures

All command logs and receipts remain in this bundle, including failed attempts.

- `development-check`: five ZIP error-helper inference errors. The helper now
  accepts a concrete archive error and callers explicitly convert error kinds.
- `development-check-v2`: the OPC mapper still named the older verified-reader
  error type. It now maps the distinct immediate-abort precompressed error.
- `transfer-tests`: twelve cases passed; the native `bug62513.pptx` case was
  refused for trailing archive bytes. That unmodified input remains a typed
  refusal test. Unmodified `EmbeddedVideo.pptx` supplies the positive OPC image
  transfer case. This is OPC substrate evidence, not native PPTX slide-copy
  acceptance.
- `consumer-tests`: an existing shared-allocation assertion needed to inspect
  the new decoded enum variant before applying `Arc::ptr_eq`.
- `fmt`: batch-owned formatting was corrected. The unrelated DOCX
  `glossary_authoring.rs` finding remains; `fmt-debt.json` proves that file is
  byte-identical to the before-capture manifest. `affected-fmt` passes.
- `affected-strict`: the transfer variant enlarged every topology payload enum.
  Boxing that variant keeps ordinary decoded additions compact.
- `affected-strict-v2`: eight test calls attempted to drop a `Copy` Part view.
  Those ineffective calls were removed; explicit package-owner drops remain.
  `affected-strict-v3` passes with warnings denied.

The integrated suite passed 1,754 tests with five ignored. After the boxed-layout
and test cleanup corrections, the OPC/PPTX suite passed 1,305 tests with three
ignored. ZIP passed 449 tests with two ignored. The fuzz build and 1,000-run
AddressSanitizer smoke passed and its isolated directory was removed. The final
ZIP source additionally clarifies in documentation that progress counts are
cumulative; that comment does not change the tested implementation.

Memory reservations model retained compressed capture and generated-member
capacity/name terms. They are not a count of all heap allocations or allocator
peak bytes. The writer still buffers the generated member; RSS must be assessed
from the matched measurements.

## Refined capture granularity

The first matched candidate passed output gates but regressed simulated delayed
range API medians by 8.8–10%. Its complete reports remain in the sibling
`change-0431-first-attempt` bundle. Commit `0556401e2` changes compressed
capture requests from 16 KiB to 64 KiB while retaining bounded decoding,
short-read refill, source fencing and immediate cancellation. An instrumented
read test checks that behavior. The refined affected suite passed 1,755 tests
with five ignored; strict Clippy, affected formatting and a fresh 1,000-run
AddressSanitizer fuzz smoke also passed. Earlier rustdoc, workspace, default
feature and ownership-boundary checks remain retained.
