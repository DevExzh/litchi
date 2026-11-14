//! Emits `cfg(odf_objects)` when an ODF object format feature is enabled,
//! and `cfg(iwork_adapter)` when an iWork format feature is enabled.
//!
//! `odf-objects` is the internal umbrella of the shared object-tree adapter
//! (`format::odf_tree`); every ODF object format turns it on. The adapter's
//! code needs at least one object format, so it is gated on this cfg rather
//! than on the umbrella feature, and `--features odf-objects` alone builds
//! the umbrella's dependencies only.
//!
//! `iwork-adapter` is the internal umbrella of the iWork adapter
//! (`format::iwork`) in the same way: its code is gated on
//! `cfg(iwork_adapter)`, so `--features iwork-adapter` alone reports only the
//! crate's `compile_error!` that names the iWork format features.

const OBJECT_FORMATS: [&str; 5] = ["ODG", "ODC", "ODI", "ODB", "ODF_FORMULA"];
const IWORK_FORMATS: [&str; 3] = ["PAGES", "KEYNOTE", "NUMBERS"];

fn main() {
    println!("cargo::rerun-if-changed=build.rs");
    emit_if_any("odf_objects", &OBJECT_FORMATS);
    emit_if_any("iwork_adapter", &IWORK_FORMATS);
}

/// Declare `cfg(name)` and set it when one of `features` is enabled.
fn emit_if_any(name: &str, features: &[&str]) {
    println!("cargo::rustc-check-cfg=cfg({name})");
    let enabled = features
        .iter()
        .any(|feature| std::env::var_os(format!("CARGO_FEATURE_{feature}")).is_some());
    if enabled {
        println!("cargo::rustc-cfg={name}");
    }
}
