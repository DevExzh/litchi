//! Fresh-document builders must work in every build profile.
//!
//! Record 0297 found `NumbersDocumentBuilder::build()` failing only in
//! release builds; run with `cargo test --release -p litchi-iwa --test
//! release_builders` to cover the optimized profile.

use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};

#[test]
fn numbers_builder_builds() {
    let editor = NumbersDocumentBuilder::new().build().expect("build");
    let bytes = editor.to_bytes().expect("serialize");
    assert!(!bytes.is_empty());
}

#[test]
fn numbers_editor_creates() {
    NumbersEditor::create().expect("create");
}
