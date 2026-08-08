#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::Database;

#[test]
fn real_libreoffice_query_package_is_an_exact_inert_snapshot() {
    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb")
            .to_vec();
    let database = Database::from_bytes(bytes.clone()).unwrap();

    assert_eq!(database.as_bytes(), bytes.as_slice());
    assert!(database.content_xml().contains("<office:database"));
    assert!(database.content_xml().contains("<db:queries"));
    assert!(
        database
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "content.xml")
    );
}

#[test]
fn real_libreoffice_schema_package_is_an_exact_inert_snapshot() {
    let bytes =
        include_bytes!("../../../test-data/libreoffice-core/extras/source/database/biblio.odb")
            .to_vec();
    let database = Database::from_bytes(bytes.clone()).unwrap();

    assert_eq!(database.as_bytes(), bytes.as_slice());
    assert!(database.content_xml().contains("<office:database"));
    assert!(database.content_xml().contains("<db:table-representation"));
    assert!(
        database
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "settings.xml")
    );
}
