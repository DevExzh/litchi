//! Fresh-document builders must work in every build profile.
//!
//! Record 0297 found `numbers::Builder::build()` failing only in
//! release builds; run with `cargo test --release -p litchi-iwa --test
//! release_builders` to cover the optimized profile.

use litchi_iwa::numbers::{self};

#[test]
fn numbers_builder_builds() {
    let editor = numbers::Builder::new().build().expect("build");
    let bytes = editor.to_bytes().expect("serialize");
    assert!(!bytes.is_empty());
}

#[test]
fn numbers_editor_creates() {
    numbers::Draft::create().expect("create");
}
