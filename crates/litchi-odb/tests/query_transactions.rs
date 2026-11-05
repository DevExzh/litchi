#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::{Builder, Package};

const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:href="sdbc:embedded:firebird" x:type="simple"/></d:connection-data></d:data-source><d:queries><d:query d:name="ledger" d:command="SELECT 1"/></d:queries></o:database></o:body></o:document-content>"#;

const GOLDEN: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:href="sdbc:embedded:firebird" x:type="simple"/></d:connection-data></d:data-source><d:queries><d:query d:name="ledger" d:command="SELECT &quot;amount&quot; &lt; 10" d:escape-processing="false"/></d:queries></o:database></o:body></o:document-content>"#;

#[test]
fn compact_query_edit_matches_the_exact_golden_xml() {
    let database = Builder::new().content_xml(SOURCE).build().unwrap();
    let mut edit = database.edit();
    edit.set_query_command("ledger", "SELECT \"amount\" < 10")
        .unwrap();
    edit.set_query_escape_processing("ledger", Some(false))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(commit.diagnostics().changed());
    assert_eq!(commit.package().database().content_xml(), GOLDEN);
    let change = commit.patch().change().unwrap();
    assert_eq!(change.name(), "ledger");
    assert_eq!(change.before_command(), "SELECT 1");
    assert_eq!(change.after_command(), "SELECT \"amount\" < 10");
    assert_eq!(change.before_escape_processing(), None);
    assert_eq!(change.after_escape_processing(), Some(false));
    assert!(!commit.patch().is_applicable_to(commit.package()));
    assert!(commit.package().apply(commit.patch()).is_err());
    let restored = commit.package().apply(&commit.patch().inverse()).unwrap();
    assert_eq!(restored.into_package().database().content_xml(), SOURCE);
}

#[test]
fn real_libreoffice_pretty_query_package_edits_losslessly_and_reopens() {
    let bytes = Vec::from(include_bytes!(
        "../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb"
    ));
    let database = Package::from_bytes(&bytes).unwrap();
    let mut edit = database.edit();
    edit.set_query_command("AliasTest", "SELECT 7").unwrap();
    edit.set_query_escape_processing("AliasTest", Some(true))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(database.as_bytes(), bytes.as_slice());
    assert!(
        commit
            .package()
            .database()
            .content_xml()
            .starts_with("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<office:document-content")
    );
    assert_eq!(
        commit
            .package()
            .database()
            .catalog()
            .unwrap()
            .query("AliasTest")
            .unwrap()
            .unwrap()
            .command(),
        "SELECT 7"
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .as_bytes(),
        bytes
    );
}

#[test]
fn transaction_rejects_ambiguous_and_absent_query_selectors() {
    let content = SOURCE.replace(
        "</d:queries>",
        r#"<d:query d:name="other" d:command="SELECT 2"/></d:queries>"#,
    );
    let database = Builder::new().content_xml(content).build().unwrap();
    let mut edit = database.edit();
    assert!(edit.set_query_command("missing", "SELECT 0").is_err());
    edit.set_query_command("ledger", "SELECT 3").unwrap();
    // A second query stages too (record 0288); the patch then reports no
    // single legacy query change.
    edit.set_query_command("other", "SELECT 4").unwrap();
    assert!(edit.set_query_command("missing", "SELECT 0").is_err());
    let commit = edit.commit().unwrap();
    assert!(commit.patch().change().is_none());
    let catalog = commit.package().database().catalog().unwrap();
    assert_eq!(
        catalog.query("ledger").unwrap().unwrap().command(),
        "SELECT 3"
    );
    assert_eq!(
        catalog.query("other").unwrap().unwrap().command(),
        "SELECT 4"
    );

    let duplicate = SOURCE.replace(
        "</d:queries>",
        r#"<d:query d:name="ledger" d:command="SELECT 2"/></d:queries>"#,
    );
    let duplicate_database = Builder::new().content_xml(duplicate).build().unwrap();
    assert!(
        duplicate_database
            .edit()
            .set_query_command("ledger", "SELECT 5")
            .is_err()
    );
}

#[test]
fn semantic_noop_preserves_the_exact_source_artifact() {
    let database = Builder::new().content_xml(SOURCE).build().unwrap();
    let mut edit = database.edit();
    edit.set_query_command("ledger", "SELECT 1").unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().change().is_none());
    assert_eq!(commit.package().as_bytes(), database.as_bytes());
}
