use std::path::PathBuf;

use litchi_numbers::{
    Error as PackageError, Package, cell::Value, compatibility_tables_from_bytes,
};

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn iwork_fixture(application: &str, filename: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork")
        .join(application)
        .join(filename)
}

fn assert_expected_cells(package: &Package) -> Result<(), Box<dyn std::error::Error>> {
    let sheet = package
        .workbook_ref()
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first sheet"))?;
    let table = sheet
        .at(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first table"))?;
    assert_eq!(table.row_count(), 22);
    assert_eq!(table.column_count(), 7);

    let marker = table.get_a1("B2")?;
    assert!(
        matches!(
            marker,
        Some(Value::Text(text)) if text == "Litchi native Numbers fixture"
        ),
        "unexpected B2 value {marker:?}; stored cells: {:?}",
        table.iter_cells().collect::<Vec<_>>()
    );
    let numeric_value = table.get_a1("B3")?;
    let Some(Value::Number(number)) = numeric_value else {
        panic!(
            "unexpected B3 value {numeric_value:?}; stored cells: {:?}",
            table.iter_cells().collect::<Vec<_>>()
        );
    };
    assert!((number.get() - 42.0).abs() <= f64::EPSILON);
    Ok(())
}

#[test]
#[allow(deprecated)]
fn native_numbers_table_info_projection_opens_rooted_and_compatibility_paths()
-> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let package = Package::open(&path)?;
    assert_expected_cells(&package)?;

    let bytes = std::fs::read(path)?;
    let from_bytes = Package::from_bytes(&bytes)?;
    // The rooted reader reaches its table model through the strict private
    // TableInfo Buffa projection. The separate global-compatibility entry
    // must retain the same native table result without first constructing the
    // rooted package snapshot.
    assert_expected_cells(&from_bytes)?;
    assert_eq!(
        compatibility_tables_from_bytes(&bytes)?,
        from_bytes.extract_structured_tables()?
    );
    assert_eq!(
        from_bytes.workbook_ref().sheet_count(),
        package.workbook_ref().sheet_count()
    );
    Ok(())
}

#[test]
fn native_foreign_iwork_fixtures_are_typed_not_numbers() -> Result<(), Box<dyn std::error::Error>> {
    for path in [
        iwork_fixture("pages", "basic.pages"),
        iwork_fixture("keynote", "basic.key"),
    ] {
        let bytes = std::fs::read(&path)?;
        assert!(matches!(
            Package::from_bytes(&bytes),
            Err(PackageError::UnexpectedFormat { .. })
        ));
        assert!(matches!(
            compatibility_tables_from_bytes(&bytes),
            Err(PackageError::UnexpectedFormat { .. })
        ));
    }
    Ok(())
}
