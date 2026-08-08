use std::path::PathBuf;

use litchi_numbers::{Package, PackageError, cell::Value, compatibility_tables_from_bytes};

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
        .document()
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first sheet"))?;
    let table = sheet
        .at(0)?
        .ok_or_else(|| std::io::Error::other("native Numbers fixture has no first table"))?;

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
fn native_numbers_fixture_opens_from_path_and_bytes() -> Result<(), Box<dyn std::error::Error>> {
    let path = fixture_path();
    let package = Package::open(&path)?;
    assert_expected_cells(&package)?;
    assert!(package.object_count() > 0);

    let bytes = std::fs::read(path)?;
    let from_bytes = Package::from_bytes(&bytes)?;
    assert_expected_cells(&from_bytes)?;
    assert_eq!(
        compatibility_tables_from_bytes(&bytes)?,
        from_bytes.extract_structured_tables()?
    );
    assert_eq!(
        from_bytes.document().sheet_count(),
        package.document().sheet_count()
    );
    assert_eq!(from_bytes.object_count(), package.object_count());
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
            Err(PackageError::NotNumbers)
        ));
        assert!(matches!(
            compatibility_tables_from_bytes(&bytes),
            Err(PackageError::NotNumbers)
        ));
    }
    Ok(())
}
