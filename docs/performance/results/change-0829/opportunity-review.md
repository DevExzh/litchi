# 0829 bounded opportunity review

Two PPTX candidates remain after current-source review:

- `opened/model.rs::capture_internal` reads `PresentationPart::slide_references`
  before constructing a fresh `Presentation`. `capture_slides_inner` then calls
  that view's `catalog`, whose empty memo causes another catalog parse. Passing
  the existing reference slice into the internal capture path could remove only
  the duplicate parse while retaining the initial limit/error order, per-slide
  relationship checks, one-to-one identities, names, notes and fingerprinting.
  Existing capture projection/error-precedence and catalog memo tests constrain
  this candidate. Historical 0822 leaf samples were only 17/15, so a practical
  end-to-end benefit is not established.
- `selected_raw_span_for_shape` still runs `raw_shape_span(..., true)` after a
  successful unrewritten Scene. Directly returning the Scene span is unsafe:
  the raw scanner counts local family/contentPart candidates even in unsupported
  namespaces and checks raw-map/tree consistency, while Scene intentionally
  ignores some such elements. The existing proof establishes XML/limit/MCE
  properties, not these raw refusal rules. Historical raw-span self leaves
  were 38/44. A combined scan needs an explicit raw-map proof and differential
  refusal tests before it can be considered equivalent.

The capture/publish profile phases can localize catalog work; set-text contains
raw-span work. Historical sample counts are context only, not current cost or
predicted savings. Both candidates must retain complete package revision,
signed-package handling, patch capture and publication source checks.

The current batch profiles remaining PPTX edit work. A separate read-only review
identified the next non-PPTX investigation: XLSX one-cell semantic commit
materialization/readback. This is a source hypothesis, not a measured attribution
or authorization to relax validation.

0827 records XLSX edit at 0.2644935 ms, with 2,510 allocation calls, 2,898,254
allocated bytes and 1,066,957 peak bytes above entry on an 8,435-byte input.
DOCX edit is about 0.056 ms with 121,273 allocated bytes. Default/full path
publication remains around 5 ms; 0821's explicit-policy evidence does not permit
weakening the durability default.

`workbook/model.rs::Worksheet::store` parses and validates the source worksheet.
`workbook/edit/semantic/transaction.rs::Edit::commit` performs provenance-based
rewriting and compaction, then complete post-write worksheet parsing when the
reduced-readback proof does not apply. Finally it constructs the new workbook
through `Workbook::from_package_with_styles` and adopts validated stores.
The real fixture's eight cells are below the 4,096-cell/1-MiB sparse-admission
threshold and its MCE/x14ac and shared strings require the full path.

A future measured investigation should distinguish source store/style validation,
rewrite/compaction, post-write parse/style/changed-cell verification, and final
snapshot/store adoption, using real, unmarked direct-string and larger-sheet
controls. Threshold relaxation alone is not justified: refusal precedence,
shared-string semantics, complete owner preservation and graph checks must stay
intact. The earlier 0471 rewrite-buffer lifetime experiment was rejected and is
not revived by these allocation totals.

This review used source and retained 0821/0827 evidence only. It ran no workload
and claims no new allocation or CPU phase share.
