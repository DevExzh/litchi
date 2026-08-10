use super::codec;
use super::{Key, Kind, Policy, Transaction};

const PACKAGE_SETTINGS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"
    xmlns:opaque="urn:example:opaque">
  <office:settings>
    <config:config-item-set config:name="ooo:configuration-settings">
      <config:config-item config:name="ProtectForm" config:type="boolean">true</config:config-item>
      <opaque:kept opaque:value="1">tail</opaque:kept>
    </config:config-item-set>
  </office:settings>
</office:document-settings>"#;

const FLAT_DOCUMENT: &[u8] = br#"<office:document
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">
  <office:settings/>
  <office:body/>
</office:document>"#;

#[test]
fn reads_typed_package_policy_with_unknown_content() {
    let policy = codec::parse_package(PACKAGE_SETTINGS);
    assert!(policy.is_ok());
    let policy = policy.unwrap_or_default();
    assert_eq!(policy.forms, Some(true));
    assert_eq!(policy.bookmarks, None);
    assert_eq!(policy.read_only, None);
    assert_eq!(policy.redline_key, None);
}

#[test]
fn transaction_changes_values_without_dropping_unknown_xml() {
    let mut transaction = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    transaction.set_bookmarks(Some(false));
    transaction.set_read_only(Some(true));
    transaction
        .set_redline_key(Some(Key::new(vec![1, 2, 3]).expect("bounded key")))
        .expect("valid key");

    let commit = transaction.commit().expect("transaction commits");
    let output = commit.as_bytes();
    assert!(
        output
            .windows(b"opaque:kept".len())
            .any(|window| window == b"opaque:kept")
    );
    assert!(
        output
            .windows(b"AQID".len())
            .any(|window| window == b"AQID")
    );

    let policy = codec::parse_package(output).expect("committed settings parse");
    assert_eq!(policy.forms, Some(true));
    assert_eq!(policy.bookmarks, Some(false));
    assert_eq!(policy.read_only, Some(true));
    assert_eq!(
        policy.redline_key.as_ref().map(Key::as_bytes),
        Some([1, 2, 3].as_slice())
    );
}

#[test]
fn no_op_transaction_keeps_source_bytes() {
    let transaction = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    let commit = transaction.commit().expect("no-op commits");
    assert!(commit.is_unchanged());
    assert_eq!(commit.as_bytes(), PACKAGE_SETTINGS);
}

#[test]
fn source_checked_rewrite_rejects_a_stale_policy() {
    let expected = Policy::default();
    let replacement = Policy::default().with_bookmarks(Some(true));
    let result = codec::rewrite(PACKAGE_SETTINGS, Kind::Package, &expected, &replacement);
    assert!(result.is_err());
}

#[test]
fn flat_settings_can_receive_new_policy_items() {
    let expected = codec::parse_flat(FLAT_DOCUMENT).expect("valid flat document");
    assert_eq!(expected, Policy::default());
    let replacement = Policy::default().with_forms(Some(false));
    let output = codec::rewrite(FLAT_DOCUMENT, Kind::Flat, &expected, &replacement)
        .expect("flat policy rewrite");
    let parsed = codec::parse_flat(&output).expect("rewritten flat document parses");
    assert_eq!(parsed, replacement);
}

#[test]
fn durable_patch_inverse_transfer_and_disjoint_merge_are_semantic() {
    let mut forms = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    forms.set_forms(Some(false));
    let forms = forms.commit().expect("form policy commits");

    let durable = forms.patch().durable().expect("durable patch");
    let wire = durable.to_deterministic_json().expect("deterministic JSON");
    let reopened =
        super::DurablePatch::from_deterministic_json(&wire).expect("durable patch reopens");
    let replayed = reopened.apply(PACKAGE_SETTINGS).expect("patch replays");
    let restored = reopened
        .inverse()
        .apply(replayed.as_bytes())
        .expect("inverse restores");
    assert_eq!(restored.as_bytes(), PACKAGE_SETTINGS);

    let transferred = forms
        .patch()
        .transfer()
        .apply_flat(FLAT_DOCUMENT)
        .expect("policy transfers");
    assert_eq!(
        codec::parse_flat(transferred.as_bytes())
            .expect("transferred policy parses")
            .forms,
        Some(false)
    );

    let mut bookmarks = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    bookmarks.set_bookmarks(Some(true));
    let bookmarks = bookmarks.commit().expect("bookmark policy commits");
    let merged = super::Patch::merge(forms.patch(), bookmarks.patch())
        .expect("merge plan")
        .finish()
        .expect("disjoint merge finishes");
    let merged = merged
        .apply(PACKAGE_SETTINGS)
        .expect("merged patch applies");
    assert_eq!(merged.after().forms, Some(false));
    assert_eq!(merged.after().bookmarks, Some(true));
}

#[test]
fn conflicting_protection_merge_requires_an_explicit_branch() {
    let mut left = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    left.set_forms(Some(false));
    let left = left.commit().expect("left commits");

    let mut right = Transaction::package(PACKAGE_SETTINGS).expect("valid settings");
    right.set_forms(None);
    let right = right.commit().expect("right commits");

    let mut merge = super::Patch::merge(left.patch(), right.patch()).expect("merge plan");
    assert_eq!(
        merge.conflicts().collect::<Vec<_>>(),
        vec![super::Field::Forms]
    );
    assert!(merge.clone().finish().is_err());
    merge
        .resolve(super::Field::Forms, super::Resolution::Right)
        .expect("conflict resolves");
    assert_eq!(
        merge
            .finish()
            .expect("resolved merge finishes")
            .apply(PACKAGE_SETTINGS)
            .expect("merged patch applies")
            .after()
            .forms,
        None
    );
}
