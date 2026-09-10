# DOCX section-boundary failure review

## Finding

The failure in `managed_document_bounds_namespace_scan_depth_and_events` is a
pre-existing test-fixture/admission-order mismatch. It is not caused by the
0497 atomic tail-path change, and production code should not be changed to
make this particular assertion pass.

The same assertion fails in the isolated before tree and in the candidate:

- [`baseline-sections1.stdout`](validation/baseline-sections1.stdout)
  and its stderr report the `event limit` assertion failure.
- [`candidate1-docx-default.stdout`](validation/candidate1-docx-default.stdout)
  and its stderr report the identical failure.
- The before and after trees are both at `66af25e2fc3c208e6819fffa5fa29e102942bc8e`.
  The 0497 production diff is limited to the atomic tail-append path; neither
  `source_backed.rs` nor `source_backed_sections.rs` changed.

The gate JSON records an after-tree cwd/source-manifest for the baseline
invocation because of the harness's fixed `after` path. The source file is
identical in the two trees, and the isolated before run reproduces the failure,
so that metadata issue does not affect the diagnosis.

## Cause

The fixture helper at
[`source_backed_sections.rs:115`](../../../../crates/litchi-docx/tests/source_backed_sections.rs#L115)
gives every case a 16 MiB `Budget`. The event-boundary case at
[`source_backed_sections.rs:472`](../../../../crates/litchi-docx/tests/source_backed_sections.rs#L472)
creates 1,000,001 `<w:p/>` elements, about 6 MB of XML.

Before `quick_xml` is entered, `ensure_source_document_xml` in
[`source_backed.rs:2132`](../../../../crates/litchi-docx/src/source_backed.rs#L2132)
admits a conservative memory envelope of
`xml.len() * 32 + 131,072`, plus object and depth reservations. For this
fixture the memory reservation is 192,134,880 bytes (about 183.23 MiB), so the 16 MiB budget returns
the typed `Error::Opc(OpcError::Execution(ExecutionError::ResourceLimit(...)))`
at workspace admission. The event loop therefore never reaches the
1,000,000-event check at `source_backed.rs:2179`.

That admission is intentional: the scanner charges its bounded namespace and
topology workspace before parser work. Removing or weakening it would violate
the finite-resource and fail-closed behavior described by
[ADR 0005](../../../adr/0005-io-memory-and-performance.md) and
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md).

## Safe test disposition

Keep the helper's existing 16 MiB default for all ordinary section tests. Add
one helper parameter so the event case can use a finite 256 MiB budget, which
is large enough to admit this deliberately 6 MiB fixture and still exercises
the production event ceiling. Before that assertion, call the same event
fixture through the default helper and explicitly assert the typed memory
admission refusal. This preserves coverage of the tight budget while making
the event-limit assertion reach the intended branch.

The following is a patch draft only. It has not been applied in this review.

```diff
diff --git a/crates/litchi-docx/tests/source_backed_sections.rs b/crates/litchi-docx/tests/source_backed_sections.rs
--- a/crates/litchi-docx/tests/source_backed_sections.rs
+++ b/crates/litchi-docx/tests/source_backed_sections.rs
@@ -4,10 +4,10 @@ use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
 use litchi_core::{
-    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
-    OwnedSource, ReadAt, Resource, SourceVersion,
+    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits,
+    Limits as CoreLimits, OwnedSource, ReadAt, Resource, SourceVersion,
 };
@@
-use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter};
+use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter};
@@
 fn managed_document_fixture(document: &[u8]) -> (Budget, source_backed::Package) {
+    managed_document_fixture_with_memory(document, 16 * 1024 * 1024)
+}
+
+fn managed_document_fixture_with_memory(
+    document: &[u8],
+    memory: u64,
+) -> (Budget, source_backed::Package) {
     let bytes = malformed_fixture(document);
-    let memory = 16 * 1024 * 1024;
     let budget = Budget::root(
@@
     for _ in 0..1_000_001 {
         events.push_str("<w:p/>");
     }
     events.push_str("</w:body></w:document>");
-    let (_budget, package) = managed_document_fixture(events.as_bytes());
+    let (_tight_budget, tight_package) = managed_document_fixture(events.as_bytes());
+    assert!(matches!(
+        tight_package.document(),
+        Err(Error::Opc(OpcError::Execution(
+            ExecutionError::ResourceLimit(limit),
+        ))) if limit.resource == Resource::Memory
+    ));
+    let (_budget, package) =
+        managed_document_fixture_with_memory(events.as_bytes(), 256 * 1024 * 1024);
     assert!(matches!(
         package.document(),
         Err(Error::InvalidFormat(reason)) if reason.contains("event limit")
```

The explicit tight-budget assertion is deliberately typed rather than a broad
`Error::Opc(_)` match. It documents that this fixture is refused at resource
admission, while the 256 MiB assertion documents the independent event-limit
boundary. No production, Cargo, or protected-worktree file was modified for
this review; no build or test command was run by this review.
