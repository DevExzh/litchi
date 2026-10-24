# 0362ugt — PPTX refusal classification by kind

The PPTX arm of `litchi-ops::error::Classify` had two historical exceptions
for `slide_removal`: a signed-package refusal was recognized by its operation
and exact reason to attach the signed-edit suggestion, and `FinalSlide` was
reclassified from its `InvalidValue` kind to `invalid_argument`. That made the
facade's code depend on operation/reason text instead of the shared
`RefusalKind` contract in record 0351 §2.4.

The arm now maps every `Error::Refused(Refusal)` through `refusal.kind()`.
`FinalSlide` therefore reports `invalid_value`; `SignedPackage` and
`ProtectedPresentation` retain `Protected` and report `document_protected`.
The operation and reason remain the native, static, content-free diagnostics;
they do not affect the unified code. No native PPTX refusal kind or reason was
changed.

The same change updates every active consumer of the old exception in this
commit: the error-classification regression, the final-slide eager and source
integration assertions, the capability-gap assertion and PPTX capability
golden. The active W8 error-contract wording and the native `FinalSlide`
documentation now name `invalid_value`. The historical 0356/0361 exception
sentences are superseded by this correction record; they are retained as
historical wave reports. The capability disposition remains the same gap and
follow-up; only its code changes.

The focused classifier regression iterates every current `RefusalKind` with
both the old `slide_removal`/signed reason pair and an unrelated operation and
caller-authored reason marker. It asserts the kind table and that the facade
display remains content-free.

## Validation

Validation is to run from the fresh `uo2/pptx-refusal-kind` leaf using
`.coord/pptx-refusal-kind/env.sh`: one-job, offline, external temporary and
target directories, with the 20 GB memory cap. Required checks are the focused
PPTX classifier test, `pptx_slide_remove`, the PPTX capability golden, and
strict PPTX-enabled `litchi-ops` Clippy. Their exact source commit, lock
overlay provenance, logs, counts and target size will be appended after the
capped build completes.
