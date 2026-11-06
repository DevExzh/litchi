//! Native Word protection metadata and explicit publication policy.
use litchi_core::Position;
use litchi_doc::parts::protection::{ProtectionAuthorization, ProtectionPolicy};
use litchi_doc::{EditOptions, Package};
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word/word-document-protection.doc"),
    )
    .unwrap()
}
fn policy() -> ProtectionPolicy {
    ProtectionPolicy::allow_protected(
        ProtectionAuthorization::audited(
            "native-protection-regression",
            "Exercise an explicitly authorized local edit of the tracked test fixture",
        )
        .unwrap(),
    )
}
fn check(p: &Package, first: &str) {
    let d = p.document();
    assert_eq!(d.paragraph_count().unwrap(), 7);
    assert_eq!(
        d.body_paragraph_text(Position::new(0)).unwrap(),
        Some(first)
    );
    let properties = d.document_properties().unwrap().unwrap();
    assert_eq!(
        properties.version(),
        litchi_doc::parts::document_properties::DocumentPropertyVersion::Word2013
    );
    let protection = properties.base().protection();
    assert!(protection.comments_or_read_only);
    assert!(
        !protection.form_fields
            && !protection.tracked_revisions
            && !protection.track_revisions
            && !protection.vba_project_locked
    );
    assert_eq!(properties.base().protection_key(), 0);
    // This fixture has document-wide protection, not range-protection bookmarks.
    assert!(d.protected_ranges().is_none());
    let comments = d.comments().unwrap();
    assert_eq!(comments.len(), 1);
    let c = &comments[0];
    assert_eq!(c.author, "Ryker Zhu");
    assert_eq!(c.initials, "RZ");
    assert_eq!(
        (
            c.reference_position,
            c.range_start,
            c.range_end,
            c.bookmark_tag
        ),
        (136, Some(118), Some(136), Some(435933812))
    );
    assert_eq!(
        c.text().trim(),
        "PROTECTION_COMMENT_008: native comments-only restriction fixture."
    );
    let metadata = c.extended_metadata.unwrap();
    let at = metadata.modified_at.unwrap();
    assert_eq!(
        (at.year, at.month, at.day, at.hour, at.minute),
        (2026, 10, 15, 16, 20)
    );
    assert_eq!(
        (metadata.depth, metadata.parent_index, metadata.is_ink),
        (0, None, false)
    );
}
#[test]
fn native_comments_protection_retains_settings_comment_range_and_noop_bytes() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    check(&p, "WORD_PROTECTION_008");
    assert!(p.to_bytes().unwrap() == bytes);
    let commit = p.edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.package().to_bytes().unwrap() == bytes);
}
#[test]
fn native_protection_default_refusal_is_atomic() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit().unwrap();
    let result = edit.replace_paragraph(Position::new(0), "EDIT_PROTECTION_008");
    assert!(matches!(
        result,
        Err(litchi_doc::package::edit::Error::Body(
            litchi_doc::body_text::Error::Invalid(litchi_doc::Error::ProtectionDenied(
                litchi_doc::parts::protection::EditProtection::Document
            ))
        ))
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.package().to_bytes().unwrap() == bytes);
    check(commit.package(), "WORD_PROTECTION_008");
}
#[test]
fn authorized_native_body_edit_preserves_protection_and_requires_fresh_authority() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let properties = p
        .document()
        .document_properties()
        .unwrap()
        .unwrap()
        .to_bytes()
        .unwrap();
    let mut edit = p
        .edit_with_options(EditOptions::default().protection_policy(policy()))
        .unwrap();
    edit.replace_paragraph(Position::new(0), "EDIT_PROTECTION_008")
        .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    assert_other_streams(&bytes, &after);
    let reopened = Package::open(output.path()).unwrap();
    check(&reopened, "EDIT_PROTECTION_008");
    assert_eq!(
        reopened
            .document()
            .document_properties()
            .unwrap()
            .unwrap()
            .to_bytes()
            .unwrap(),
        properties
    );
    let error = reopened.apply(&commit.patch().inverse()).unwrap_err();
    assert!(matches!(
        error,
        litchi_doc::package::edit::Error::Body(litchi_doc::body_text::Error::Refused(
            litchi_doc::body_text::Refusal::Protected
        ))
    ));
    assert!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}

fn streams(bytes: &[u8]) -> std::collections::BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let refs: Vec<_> = path.iter().map(String::as_str).collect();
            let bytes = ole.open_stream(&refs).unwrap();
            (path, bytes)
        })
        .collect()
}

fn assert_other_streams(before: &[u8], after: &[u8]) {
    let before = streams(before);
    let after = streams(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in before {
        if path != ["WordDocument"] && path != ["1Table"] {
            assert!(after[&path] == data, "changed {path:?}");
        }
    }
}
#[test]
fn native_forms_protection_keeps_ffdata_when_an_unrelated_edit_is_authorized() {
    let bytes = std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word/word-legacy-form.doc"),
    )
    .unwrap();
    let p = Package::from_bytes(&bytes).unwrap();
    let d = p.document();
    let props = d.document_properties().unwrap().unwrap();
    let protection = props.base().protection();
    assert!(protection.form_fields);
    assert!(
        !protection.comments_or_read_only
            && !protection.tracked_revisions
            && !protection.track_revisions
            && !protection.vba_project_locked
    );
    assert_eq!(props.base().protection_key(), 0);
    let dop = props.to_bytes().unwrap();
    assert_eq!(dop.len(), 694);
    let fields = d.legacy_form_fields().unwrap();
    assert_eq!(fields.len(), 3);
    assert!(d.protected_ranges().is_none());
    let mut denied = p.edit().unwrap();
    assert!(matches!(
        denied.replace_paragraph(Position::new(0), "EDIT_LEGACY_FORM_010"),
        Err(litchi_doc::package::edit::Error::Body(
            litchi_doc::body_text::Error::Invalid(litchi_doc::Error::ProtectionDenied(
                litchi_doc::parts::protection::EditProtection::Document
            ))
        ))
    ));
    assert!(denied.commit().unwrap().package().to_bytes().unwrap() == bytes);
    let mut edit = p
        .edit_with_options(EditOptions::default().protection_policy(policy()))
        .unwrap();
    edit.replace_paragraph(Position::new(0), "EDIT_LEGACY_FORM_010")
        .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_other_streams(&bytes, &after);
    assert_eq!(
        reopened
            .document()
            .document_properties()
            .unwrap()
            .unwrap()
            .to_bytes()
            .unwrap(),
        dop
    );
    assert_eq!(reopened.document().legacy_form_fields().unwrap(), fields);
    assert_eq!(
        reopened
            .document()
            .body_paragraph_text(Position::new(0))
            .unwrap(),
        Some("EDIT_LEGACY_FORM_010")
    );
    assert!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
