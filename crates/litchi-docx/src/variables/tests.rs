use super::{MAX_DOCUMENT_VARIABLE_VALUE_CHARS, Snapshot, Variables};
use std::sync::Arc;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

#[test]
fn source_checked_edits_are_canonical_and_inverse_is_byte_exact() {
    let source = format!(
        r#"<?xml version="1.0"?><q:settings xmlns:q="{W}" xmlns:x="urn:opaque"><q:compat/><q:docVars><q:docVar q:name='old' q:val='old'/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
    );
    let snapshot = Snapshot::from_xml(source.as_bytes().to_vec()).unwrap();

    let mut no_op = snapshot.edit();
    no_op.set("old", "old").unwrap();
    let no_op = no_op.commit().unwrap();
    assert!(!no_op.changed());
    assert!(no_op.patch().is_empty());
    assert_eq!(no_op.snapshot().xml_bytes(), source.as_bytes());

    let mut edit = snapshot.edit();
    edit.set("Company & Team", "A < B").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().xml_bytes(),
        format!(
            r#"<?xml version="1.0"?><q:settings xmlns:q="{W}" xmlns:x="urn:opaque"><q:compat/><q:docVars><q:docVar q:name="old" q:val="old"/><q:docVar q:name="Company &amp; Team" q:val="A &lt; B"/></q:docVars><x:opaque><![CDATA[a < b]]></x:opaque><q:rsids/></q:settings>"#
        )
        .as_bytes()
    );
    assert_eq!(
        commit.snapshot().variables().get("Company & Team"),
        Some("A < B")
    );

    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), source.as_bytes());
    assert_eq!(restored.variables(), snapshot.variables());

    let stale = Snapshot::from_xml(source.replace("urn:opaque", "urn:other").into_bytes()).unwrap();
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn exact_noop_shares_the_retained_settings_allocation() {
    let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
    let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
    let source_xml = snapshot.shared_xml();
    let commit = snapshot.edit().commit().unwrap();

    assert!(!commit.changed());
    assert!(Arc::ptr_eq(&source_xml, &commit.snapshot().shared_xml()));
}

#[test]
fn transaction_failures_are_atomic_and_limits_are_enforced() {
    let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
    let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
    let mut edit = snapshot.edit();
    assert!(edit.set("", "invalid").is_err());
    assert!(edit.variables().is_empty());
    assert!(
        edit.set(
            "too-long",
            "x".repeat(MAX_DOCUMENT_VARIABLE_VALUE_CHARS + 1)
        )
        .is_err()
    );
    assert!(edit.variables().is_empty());

    assert!(
        edit.edit_variables(|variables| {
            variables.insert("staged", "value")?;
            Err(super::super::Error::Invalid("abort".into()))
        })
        .is_err()
    );
    assert!(edit.variables().is_empty());

    edit.set("first", "one").unwrap();
    edit.set("second", "two").unwrap();
    assert_eq!(edit.remove("first"), Some("one".into()));
    edit.clear();
    assert!(edit.variables().is_empty());
}

#[test]
fn strict_empty_settings_are_inserted_in_canonical_form_and_removed_atomically() {
    let source = format!(r#"<settings xmlns="{STRICT_W}"/>"#);
    let snapshot = Snapshot::from_xml(source.as_bytes().to_vec()).unwrap();
    let mut edit = snapshot.edit();
    edit.set("strict", "value").unwrap();
    let commit = edit.commit().unwrap();
    let expected = format!(
        r#"<settings xmlns="{STRICT_W}"><w:docVars xmlns:w="{STRICT_W}"><w:docVar w:name="strict" w:val="value"/></w:docVars></settings>"#
    );
    assert_eq!(commit.snapshot().xml_bytes(), expected.as_bytes());

    let mut remove = commit.snapshot().edit();
    assert_eq!(remove.remove("strict"), Some("value".into()));
    let removed = remove.commit().unwrap();
    assert_eq!(
        removed.snapshot().xml_bytes(),
        format!(r#"<settings xmlns="{STRICT_W}"></settings>"#).as_bytes()
    );
    assert_eq!(
        removed
            .patch()
            .inverse()
            .apply(removed.snapshot())
            .unwrap()
            .xml_bytes(),
        commit.snapshot().xml_bytes()
    );
}

#[test]
fn complete_collection_replacement_validates_before_publication() {
    let source = format!(
        r#"<w:settings xmlns:w="{W}"><w:docVars><w:docVar w:name="old" w:val="value"/></w:docVars></w:settings>"#
    );
    let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
    let mut edit = snapshot.edit();
    let mut replacement = Variables::new();
    replacement.insert("new", "value").unwrap();
    edit.replace(replacement).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.patch().before().get("old"), Some("value"));
    assert_eq!(commit.patch().after().get("new"), Some("value"));
}

#[test]
fn aggregate_encoded_output_is_rejected_before_publication() {
    let source = format!(r#"<w:settings xmlns:w="{W}"/>"#);
    let snapshot = Snapshot::from_xml(source.into_bytes()).unwrap();
    let mut replacement = Variables::new();
    for index in 0..129 {
        replacement
            .insert(
                format!("v{index}"),
                "&".repeat(MAX_DOCUMENT_VARIABLE_VALUE_CHARS),
            )
            .unwrap();
    }
    assert!(replacement.to_xml().is_err());
    let mut edit = snapshot.edit();
    edit.replace(replacement).unwrap();
    assert!(edit.commit().is_err());
}
