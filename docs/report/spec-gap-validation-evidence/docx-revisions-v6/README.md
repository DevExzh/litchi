DOCX inline and paired move revision authoring
=============================================

Source is byte-identical to independently reviewed V6, integrated over fd95ae50a6ee30b95c52dc2ded7440d6b59d5de7. Root Rust 1.95 checks passed: 1,599 tests including doctests, 32 ignored, strict all-target/all-feature Clippy, warning-denied rustdoc, and scoped rustfmt.

The API supports source-bound main-document inline insertion/deletion and paired move accept/reject with reversible patches. Property revisions, secondary stories and dependency-bearing unsupported owners remain explicitly refused; this batch does not claim general revision resolution. The feature matrix records that boundary.

The manifest includes a separately observed sibling checkout, which is not used for integration. Integration preimages are checked against /home/zhuhe/code/litchi-spec-gaps. Raw root logs are gzip-compressed without alteration.
