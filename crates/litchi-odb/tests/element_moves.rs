#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

//! Query, table and column moves relocate one element's bytes among its
//! siblings and leave every other byte in place (record 0292).

use litchi_odb::Builder;

const HEAD: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:href="sdbc:embedded:firebird" x:type="simple"/></d:connection-data></d:data-source>"#;
const TAIL: &str = "</o:database></o:body></o:document-content>";

fn source(body: &str) -> String {
    format!("{HEAD}{body}{TAIL}")
}

const QUERIES: &str = r#"<d:queries><d:query d:name="a" d:command="SELECT 1"/><d:query d:name="b" d:command="SELECT 2"/><d:query d:name="c" d:command="SELECT 3"/></d:queries>"#;

#[test]
fn query_moves_keep_bytes_and_invert() {
    let content = source(QUERIES);
    let database = Builder::new().content_xml(&content).build().unwrap();
    let mut edit = database.edit();
    edit.move_query("a", 2).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let names: Vec<String> = commit
        .package()
        .database()
        .catalog()
        .unwrap()
        .queries()
        .iter()
        .map(|query| query.name().to_owned())
        .collect();
    assert_eq!(names, ["b", "c", "a"]);
    let expected = source(
        r#"<d:queries><d:query d:name="b" d:command="SELECT 2"/><d:query d:name="c" d:command="SELECT 3"/><d:query d:name="a" d:command="SELECT 1"/></d:queries>"#,
    );
    assert_eq!(commit.package().database().content_xml(), expected);
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().database().content_xml(), content);

    // Back to the front, through the other branch.
    let mut edit = commit.package().edit();
    edit.move_query("a", 0).unwrap();
    let back = edit.commit().unwrap();
    let names: Vec<String> = back
        .package()
        .database()
        .catalog()
        .unwrap()
        .queries()
        .iter()
        .map(|query| query.name().to_owned())
        .collect();
    assert_eq!(names, ["a", "b", "c"]);
}

#[test]
fn self_moves_are_exact_noops_and_bounds_are_checked() {
    let database = Builder::new().content_xml(source(QUERIES)).build().unwrap();
    let mut edit = database.edit();
    edit.move_query("b", 1).unwrap();
    assert!(edit.changes().is_empty());
    assert!(!edit.commit().unwrap().diagnostics().changed());
    let mut edit = database.edit();
    assert!(edit.move_query("b", 3).is_err());
    assert!(edit.move_query("missing", 0).is_err());
}

#[test]
fn table_and_column_moves_stay_within_their_collection() {
    let body = r#"<d:table-representations><d:table-representation d:name="t1"><d:columns><d:column d:name="x"/><d:column d:name="y"/></d:columns></d:table-representation><d:table-representation d:name="t2"/></d:table-representations>"#;
    let database = Builder::new().content_xml(source(body)).build().unwrap();
    let mut edit = database.edit();
    edit.move_table("t2", 0).unwrap();
    edit.move_column("t1", "y", 0).unwrap();
    let commit = edit.commit().unwrap();
    let catalog = commit.package().database().catalog().unwrap();
    let tables: Vec<&str> = catalog.tables().iter().map(|table| table.name()).collect();
    assert_eq!(tables, ["t2", "t1"]);
    let columns: Vec<&str> = catalog.tables()[1]
        .columns()
        .iter()
        .map(|column| column.name())
        .collect();
    assert_eq!(columns, ["y", "x"]);
    assert_eq!(
        commit.package().database().content_xml(),
        source(
            r#"<d:table-representations><d:table-representation d:name="t2"/><d:table-representation d:name="t1"><d:columns><d:column d:name="y"/><d:column d:name="x"/></d:columns></d:table-representation></d:table-representations>"#
        )
    );
    let mut edit = database.edit();
    assert!(edit.move_column("t1", "x", 2).is_err());
    assert!(edit.move_table("t1", 2).is_err());
}

#[test]
fn producer_query_moves_keep_every_other_byte() {
    let bytes = Vec::from(include_bytes!(
        "../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb"
    ));
    let database = litchi_odb::Package::from_bytes(&bytes).unwrap();
    let catalog = database.database().catalog().unwrap();
    let names: Vec<String> = catalog
        .queries()
        .iter()
        .map(|query| query.name().to_owned())
        .collect();
    let source = database.database().content_xml().to_owned();
    let mut edit = database.edit();
    if names.len() < 2 {
        edit.move_query(&names[0], 0).unwrap();
        assert!(!edit.commit().unwrap().diagnostics().changed());
        return;
    }
    edit.move_query(&names[0], names.len() - 1).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().database().content_xml();
    assert_eq!(after.len(), source.len());
    assert!(after.starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"));
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().as_bytes(), bytes.as_slice());
}
