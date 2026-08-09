#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected fixture failures"
)]

use litchi_odb::{Builder, Database};

const SOURCE: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:href="sdbc:embedded:firebird" x:type="simple"/></d:connection-data></d:data-source><d:queries><d:query d:name="ledger" d:command="SELECT 1"/></d:queries></o:database></o:body></o:document-content>"#;

const GOLDEN: &str = r#"<?xml version="1.0" encoding="UTF-8"?><o:document-content xmlns:d="urn:oasis:names:tc:opendocument:xmlns:database:1.0" xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:database><d:data-source><d:connection-data><d:connection-resource xmlns:x="http://www.w3.org/1999/xlink" x:href="sdbc:embedded:firebird" x:type="simple"/></d:connection-data></d:data-source><d:queries><d:query d:name="ledger" d:command="SELECT &quot;amount&quot; &lt; 10" d:escape-processing="false"/></d:queries></o:database></o:body></o:document-content>"#;

#[test]
fn compact_query_edit_matches_the_exact_golden_xml() {
    let database =
        Database::from_bytes(Builder::new().content_xml(SOURCE).build().unwrap()).unwrap();
    let mut edit = database.edit();
    edit.set_query_command("ledger", "SELECT \"amount\" < 10")
        .unwrap();
    edit.set_query_escape_processing("ledger", Some(false))
        .unwrap();
    let commit = edit.commit().unwrap();

    assert!(commit.changed());
    assert_eq!(commit.database().content_xml(), GOLDEN);
    let change = commit.patch().change().unwrap();
    assert_eq!(change.name(), "ledger");
    assert_eq!(change.before_command(), "SELECT 1");
    assert_eq!(change.after_command(), "SELECT \"amount\" < 10");
    assert_eq!(change.before_escape_processing(), None);
    assert_eq!(change.after_escape_processing(), Some(false));
    assert!(!commit.patch().is_applicable_to(commit.database()));
    assert!(commit.patch().apply(commit.database()).is_err());
    let restored = commit.patch().inverse().apply(commit.database()).unwrap();
    assert_eq!(restored.content_xml(), SOURCE);
}

#[test]
fn real_libreoffice_noncompact_query_package_is_refused_without_mutation() {
    let bytes = Vec::from(include_bytes!(
        "../../../test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb"
    ));
    let database = Database::from_bytes(bytes.clone()).unwrap();
    let mut edit = database.edit();
    edit.set_query_command("AliasTest", "SELECT 7").unwrap();
    edit.set_query_escape_processing("AliasTest", Some(true))
        .unwrap();
    assert!(edit.commit().is_err());
    assert_eq!(database.as_bytes(), bytes.as_slice());
    assert_eq!(
        database
            .catalog()
            .unwrap()
            .query("AliasTest")
            .unwrap()
            .unwrap()
            .command(),
        "SELECT \"tid\" \"TestId\", \"tname\" \"TestName\" FROM \"test\""
    );
}

#[test]
fn transaction_rejects_ambiguous_absent_and_second_query_selectors() {
    let content = SOURCE.replace(
        "</d:queries>",
        r#"<d:query d:name="other" d:command="SELECT 2"/></d:queries>"#,
    );
    let database =
        Database::from_bytes(Builder::new().content_xml(content).build().unwrap()).unwrap();
    let mut edit = database.edit();
    assert!(edit.set_query_command("missing", "SELECT 0").is_err());
    edit.set_query_command("ledger", "SELECT 3").unwrap();
    assert!(edit.set_query_command("other", "SELECT 4").is_err());

    let duplicate = SOURCE.replace(
        "</d:queries>",
        r#"<d:query d:name="ledger" d:command="SELECT 2"/></d:queries>"#,
    );
    let duplicate_database =
        Database::from_bytes(Builder::new().content_xml(duplicate).build().unwrap()).unwrap();
    assert!(
        duplicate_database
            .edit()
            .set_query_command("ledger", "SELECT 5")
            .is_err()
    );
}

#[test]
fn semantic_noop_preserves_the_exact_source_artifact() {
    let database =
        Database::from_bytes(Builder::new().content_xml(SOURCE).build().unwrap()).unwrap();
    let mut edit = database.edit();
    edit.set_query_command("ledger", "SELECT 1").unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.changed());
    assert!(commit.patch().change().is_none());
    assert_eq!(commit.database().as_bytes(), database.as_bytes());
}
