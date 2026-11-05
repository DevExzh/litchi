//! Emits `cfg(litchi_unified)` when any format feature is enabled.
//!
//! The format features are derived from `Cargo.toml` (see
//! `build/format_features.rs`), so a new format feature gates the unified
//! entry point (`litchi::open`, `File`) without editing a cfg list.

use std::path::Path;

#[path = "build/format_features.rs"]
mod format_features;

fn main() {
    println!("cargo::rerun-if-changed=Cargo.toml");
    println!("cargo::rerun-if-changed=build/format_features.rs");
    println!("cargo::rustc-check-cfg=cfg(litchi_unified)");

    let directory = std::env::var_os("CARGO_MANIFEST_DIR").expect("cargo sets CARGO_MANIFEST_DIR");
    let manifest_path = Path::new(&directory).join("Cargo.toml");
    let manifest = std::fs::read_to_string(&manifest_path)
        .unwrap_or_else(|error| panic!("read {}: {error}", manifest_path.display()));
    let features = format_features::format_features(&manifest);
    assert!(
        !features.is_empty(),
        "no format feature (one enabling \"litchi-ops/<format>\") found in {}",
        manifest_path.display()
    );
    let enabled = features.iter().any(|feature| {
        let variable = format!(
            "CARGO_FEATURE_{}",
            feature.to_ascii_uppercase().replace('-', "_")
        );
        std::env::var_os(variable).is_some()
    });
    if enabled {
        println!("cargo::rustc-cfg=litchi_unified");
    }
}
