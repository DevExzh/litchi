use litchi_core::sheet::WorkbookTrait as _;
use litchi_xlsb::builder;
use litchi_xlsb::{Builder, Package};

#[path = "support/projection.rs"]
mod projection;
use projection::OwnedWorkbook as _;

#[test]
fn package_new_is_a_deterministic_minimal_workbook() {
    let first = Package::new().expect("minimal package");
    let second = Package::new().expect("minimal package");

    assert_eq!(
        first.to_bytes().expect("first package bytes"),
        second.to_bytes().expect("second package bytes")
    );
    let workbook = first.owned_workbook().expect("minimal workbook");
    assert_eq!(workbook.worksheet_names(), &["Sheet1".to_owned()]);
}

#[test]
fn builder_build_returns_an_editable_package() {
    let mut builder = Builder::new();
    builder.add_worksheet(builder::Worksheet::new("Sheet1"));

    let package = builder.build().expect("built package");
    package.edit().expect("built package is editable");
}

#[test]
fn empty_builder_refuses_then_can_build_after_a_sheet_is_added() {
    let mut builder = Builder::new();
    assert!(builder.build().is_err());
    builder.add_worksheet(builder::Worksheet::new("Data"));
    let package = builder.build().expect("retry with a worksheet");
    assert_eq!(
        package.workbook().expect("projection").worksheet_names(),
        &["Data".to_owned()]
    );
}

#[test]
fn repeated_builds_recompute_shared_strings_without_mutating_prior_packages() {
    let mut builder = Builder::new();
    let mut sheet = builder::Worksheet::new("Data");
    sheet.set_cell(0, 0, "hello");
    builder.add_worksheet(sheet);
    let first = builder.build().expect("first build");
    let bytes = first.to_bytes().expect("first bytes");
    assert_eq!(
        builder
            .build()
            .expect("repeat build")
            .to_bytes()
            .expect("repeat bytes"),
        bytes
    );

    builder
        .worksheet_mut(0)
        .expect("worksheet")
        .set_cell(0, 0, 42.0);
    let second = builder.build().expect("replace the text cell");
    let uri = litchi_opc::PackURI::new("/xl/sharedStrings.bin").expect("SST URI");
    assert!(second.opc_package().get_part(&uri).is_err());
    assert_eq!(first.to_bytes().expect("unchanged first package"), bytes);
}

#[test]
fn build_refuses_invalid_numeric_cells_before_package_publication() {
    let mut builder = Builder::new();
    let mut sheet = builder::Worksheet::new("Data");
    sheet.set_cell(0, 0, f64::NAN);
    builder.add_worksheet(sheet);
    assert!(builder.build().is_err());
    builder
        .worksheet_mut(0)
        .expect("worksheet")
        .set_cell(0, 0, 1.0);
    builder.build().expect("corrected cell is publishable");
}
