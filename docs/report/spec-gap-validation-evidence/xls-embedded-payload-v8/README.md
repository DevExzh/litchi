XLS ordinary embedded OLE payload editing
========================================

The 29 source files match reviewed V8. Root Rust 1.95 validation passed 1,967 tests across 87 targets including doctests, with 3 ignored; strict Clippy and warning-denied rustdoc also passed.

The feature covers ordinary embedded OLE object payload publication, conditional Obj record validation, bounded CFB/object traversal, and failure-atomic reversible edits. Generic form-control conditional validation remains a separate gap. The pending user_names module/export must remain outside the feature commit.

The clean validation overlay uses fd95; the following DOCX-only commit does not change this source closure. Security-corpus changes are not included in this commit; their later integration must retain this batch and pass combined validation. Raw logs are compressed without alteration.
