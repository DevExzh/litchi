//! Emits `cfg(odf_objects)` when an ODF object format feature is enabled.
//!
//! `odf-objects` is the internal umbrella of the shared object-tree adapter
//! (`format::odf_tree`); every ODF object format turns it on. The adapter's
//! code needs at least one object format, so it is gated on this cfg rather
//! than on the umbrella feature, and `--features odf-objects` alone builds
//! the umbrella's dependencies only (record 0358a).

const OBJECT_FORMATS: [&str; 5] = ["ODG", "ODC", "ODI", "ODB", "ODF_FORMULA"];

fn main() {
    println!("cargo::rerun-if-changed=build.rs");
    println!("cargo::rustc-check-cfg=cfg(odf_objects)");
    let enabled = OBJECT_FORMATS
        .iter()
        .any(|format| std::env::var_os(format!("CARGO_FEATURE_{format}")).is_some());
    if enabled {
        println!("cargo::rustc-cfg=odf_objects");
    }
}
