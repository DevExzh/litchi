use super::{EXTENSION_URI, NAMESPACE, Snapshot};

const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

fn source(payload: &str) -> Vec<u8> {
    format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}"><p:extLst><p:ext uri="{EXTENSION_URI}">{payload}</p:ext></p:extLst></p:presentationPr>"#
    )
    .into_bytes()
}

#[test]
fn reads_xsd_boolean_lexical_forms() {
    for (lexical, expected) in [("true", true), ("1", true), ("false", false), ("0", false)] {
        let snapshot = Snapshot::from_xml(source(&format!(
            r#"<p1710:readonlyRecommended val="{lexical}"/>"#
        )))
        .unwrap();
        assert_eq!(snapshot.value(), Some(expected));
    }
}

#[test]
fn xml_attribute_whitespace_preserves_scalar_edit_and_inverse() {
    for separator in ["\t", "\n", "\r", "\r\n", " \t\n"] {
        let original = source(&format!(
            "<p1710:readonlyRecommended{separator}val\t=\n'1'/>"
        ));
        let snapshot = Snapshot::from_xml(&original).unwrap();
        assert_eq!(snapshot.value(), Some(true));

        let mut noop = snapshot.edit();
        noop.set_readonly_recommended(true).unwrap();
        let noop = noop.commit().unwrap();
        assert!(!noop.is_changed());
        assert_eq!(noop.snapshot().source_xml(), original.as_slice());

        let mut edit = snapshot.edit();
        edit.set_readonly_recommended(false).unwrap();
        let commit = edit.commit().unwrap();
        let expected = source(&format!(
            "<p1710:readonlyRecommended{separator}val\t=\n'false'/>"
        ));
        assert_eq!(commit.snapshot().source_xml(), expected.as_slice());
        let reopened = Snapshot::from_xml(&expected).unwrap();
        assert_eq!(reopened.value(), Some(false));
        let mut restored = expected;
        commit.patch().inverse().apply(&mut restored).unwrap();
        assert_eq!(restored, original);
    }
}

#[test]
fn scalar_edit_and_inverse_preserve_unrelated_source() {
    let original = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}"><p:extLst><!-- before --><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended val="1"/></p:ext><!-- after --><p:ext uri="urn:example"><x:opaque xmlns:x="urn:example" a="b"/></p:ext></p:extLst></p:presentationPr>"#
    )
    .into_bytes();
    let snapshot = Snapshot::from_xml(&original).unwrap();
    let mut edit = snapshot.edit();
    edit.set_readonly_recommended(false).unwrap();
    let commit = edit.commit().unwrap();
    let changed = commit.snapshot().source_xml().to_vec();
    assert_eq!(commit.after_value(), Some(false));
    assert!(
        changed
            .windows(b"val=\"false\"".len())
            .any(|window| window == b"val=\"false\"")
    );
    assert!(
        changed
            .windows(b"<!-- before -->".len())
            .any(|window| window == b"<!-- before -->")
    );
    assert!(
        changed
            .windows(b"<!-- after -->".len())
            .any(|window| window == b"<!-- after -->")
    );
    assert!(
        changed
            .windows(b"<x:opaque".len())
            .any(|window| window == b"<x:opaque")
    );

    let mut target = changed.clone();
    commit.patch().inverse().apply(&mut target).unwrap();
    assert_eq!(target, original);
}

#[test]
fn exact_noop_keeps_source_bytes() {
    let original = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}"><p:extLst><!-- retained --><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended val="1"/></p:ext><p:ext uri="urn:example"><x:opaque xmlns:x="urn:example"/></p:ext></p:extLst></p:presentationPr>"#
    )
    .into_bytes();
    let snapshot = Snapshot::from_xml(&original).unwrap();
    let mut edit = snapshot.edit();
    edit.set_readonly_recommended(true).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.is_changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), original.as_slice());
}

#[test]
fn exact_noop_keeps_character_references_in_opaque_extensions() {
    let original = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}"><p:extLst><p:ext uri="urn:example"><x:opaque xmlns:x="urn:example">Tom &amp; Jerry</x:opaque></p:ext></p:extLst></p:presentationPr>"#
    )
    .into_bytes();
    let snapshot = Snapshot::from_xml(&original).unwrap();
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.is_changed());
    assert_eq!(commit.snapshot().source_xml(), original.as_slice());
}

#[test]
fn insertion_and_removal_keep_owner_structure() {
    let original = br#"<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:extLst><!-- retained --></p:extLst></p:presentationPr>"#.to_vec();
    let snapshot = Snapshot::from_xml(&original).unwrap();
    let mut edit = snapshot.edit();
    edit.set_readonly_recommended(true).unwrap();
    let commit = edit.commit().unwrap();
    let inserted = String::from_utf8(commit.snapshot().source_xml().to_vec()).unwrap();
    assert!(inserted.contains("readonlyRecommended"));
    assert!(inserted.contains("<!-- retained -->"));

    let mut remove = commit.snapshot().edit();
    remove.remove().unwrap();
    let removed = remove.commit().unwrap();
    assert_eq!(removed.snapshot().source_xml(), original.as_slice());
}

#[test]
fn insertion_expands_an_empty_root_without_losing_its_namespace_context() {
    let original = format!(r#"<q:presentationPr xmlns:q="{P_NS}"/>"#).into_bytes();
    let snapshot = Snapshot::from_xml(&original).unwrap();
    let mut edit = snapshot.edit();
    edit.set(Some(true)).unwrap();
    let commit = edit.commit().unwrap();
    let updated = String::from_utf8(commit.snapshot().source_xml().to_vec()).unwrap();
    assert!(updated.starts_with(&format!(r#"<q:presentationPr xmlns:q="{P_NS}">"#)));
    assert!(updated.contains(r#"<q:extLst><q:ext uri="{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}">"#));
    assert!(updated.ends_with("</q:presentationPr>"));
}

#[test]
fn mce_choice_selects_p1710_and_inverse_restores_fallback() {
    let source = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p1710="{NAMESPACE}" mc:Ignorable="p1710"><p:extLst><mc:AlternateContent><mc:Choice Requires="p1710"><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended val="1"/></p:ext></mc:Choice><mc:Fallback><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended val="0"/></p:ext></mc:Fallback></mc:AlternateContent></p:extLst></p:presentationPr>"#
    )
    .into_bytes();
    let snapshot = Snapshot::from_xml(&source).unwrap();
    assert_eq!(snapshot.value(), Some(true));
    let mut edit = snapshot.edit();
    edit.set_readonly_recommended(false).unwrap();
    let commit = edit.commit().unwrap();
    let updated = String::from_utf8(commit.snapshot().source_xml().to_vec()).unwrap();
    assert!(updated.contains(r#"<p1710:readonlyRecommended val="false"/>"#));
    assert!(updated.contains(r#"<p1710:readonlyRecommended val="0"/>"#));
    let mut target = commit.snapshot().source_xml().to_vec();
    commit.patch().inverse().apply(&mut target).unwrap();
    assert_eq!(target, source);
}

#[test]
fn namespace_and_required_attribute_validation_is_fail_closed() {
    let missing = source(r#"<p1710:readonlyRecommended/>"#);
    assert!(Snapshot::from_xml(missing).is_err());

    let qualified_value = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}" xmlns:q="urn:other"><p:extLst><p:ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended q:val="true"/></p:ext></p:extLst></p:presentationPr>"#
    );
    assert!(Snapshot::from_xml(qualified_value).is_err());

    let foreign_payload = format!(
        r#"<p:presentationPr xmlns:p="{P_NS}" xmlns:p1710="{NAMESPACE}" xmlns:q="urn:other"><p:extLst><p:ext uri="{EXTENSION_URI}"><q:readonlyRecommended val="true"/></p:ext></p:extLst></p:presentationPr>"#
    );
    assert!(Snapshot::from_xml(foreign_payload).is_err());
}

#[test]
fn default_namespace_and_strict_root_contexts_are_retained() {
    let strict = format!(
        r#"<presentationPr xmlns="http://purl.oclc.org/ooxml/presentationml/main" xmlns:p1710="{NAMESPACE}"><extLst><ext uri="{EXTENSION_URI}"><p1710:readonlyRecommended val="0"/></ext></extLst></presentationPr>"#
    );
    let snapshot = Snapshot::from_xml(&strict).unwrap();
    assert_eq!(snapshot.value(), Some(false));
    let mut edit = snapshot.edit();
    edit.set(Some(true)).unwrap();
    let commit = edit.commit().unwrap();
    let updated = String::from_utf8(commit.snapshot().source_xml().to_vec()).unwrap();
    assert!(updated.contains("xmlns=\"http://purl.oclc.org/ooxml/presentationml/main\""));
    assert!(updated.contains(r#"val="true""#));

    let defaulted = format!(
        r#"<presentationPr xmlns="{P_NS}"><extLst><ext uri="{EXTENSION_URI}"><readonlyRecommended xmlns="{NAMESPACE}" val="1"/></ext></extLst></presentationPr>"#
    );
    assert_eq!(Snapshot::from_xml(defaulted).unwrap().value(), Some(true));
}
