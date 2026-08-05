//! XLSX façade coverage for packages carrying DrawingML theme parts.
//!
//! Theme semantics are owned by `litchi-drawingml`; the XLSX façade only
//! needs to validate and retain the package graph while opening it.

use litchi_xlsx::{Package, Workbook};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsx")
        .join(name)
}

#[test]
fn standalone_package_opens_a_theme_bearing_workbook() {
    let package = Package::open(fixture("cell-borders.xlsx")).unwrap();
    let workbook = package.workbook().unwrap();

    assert!(!workbook.is_empty());
    assert!(!package.to_bytes().unwrap().is_empty());
}

#[test]
fn standalone_workbook_opens_xlsx_theme_variants() {
    for name in ["autofilter.xlsx", "column_style.xlsx"] {
        let workbook =
            Workbook::open(fixture(name)).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert!(!workbook.is_empty(), "{name}");
    }
}
