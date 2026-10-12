#![allow(
    clippy::unwrap_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    reason = "test assertions use bounded fixtures and rebind policy variants"
)]

use litchi_rtf::{OpenOptions, Package, SourceLimits};

const SOURCE: &[u8] = br"{\rtf1\ansi old}";

#[test]
fn canonical_parts_replay_and_inverse_preserve_source() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(0, "replacement").unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.package().text(), "replacement");
    let (package, patch) = commit.into_parts();
    assert_eq!(
        source.apply(&patch).unwrap().to_bytes().unwrap(),
        package.to_bytes().unwrap()
    );
    assert_eq!(
        package.apply(&patch.inverse()).unwrap().to_bytes().unwrap(),
        SOURCE
    );
    assert_eq!(source.to_bytes().unwrap(), SOURCE);
    assert!(
        Package::parse(r"{\rtf1 foreign}")
            .unwrap()
            .apply(&patch)
            .is_err()
    );
}

#[test]
fn exact_noop_replay_keeps_recipient_owner_and_policy() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let commit = source.edit().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    let limits = SourceLimits::builder()
        .max_input_bytes(SOURCE.len())
        .unwrap()
        .build()
        .unwrap();
    let recipient =
        Package::from_bytes_with_options(SOURCE, OpenOptions::new().source_limits(limits)).unwrap();
    let replay = recipient.apply(commit.patch()).unwrap();
    assert!(recipient.same_snapshot(&replay));
    assert_eq!(replay.source_limits(), limits);
    assert!(source.same_snapshot(&commit.into_package()));
}

#[test]
fn replay_admits_candidate_under_recipient_policy() {
    let source = Package::from_bytes(SOURCE).unwrap();
    let mut edit = source.edit();
    edit.replace_paragraph_text(0, "a much longer replacement")
        .unwrap();
    let commit = edit.commit().unwrap();
    let limits = SourceLimits::builder()
        .max_input_bytes(SOURCE.len())
        .unwrap()
        .build()
        .unwrap();
    let recipient =
        Package::from_bytes_with_options(SOURCE, OpenOptions::new().source_limits(limits)).unwrap();
    assert!(matches!(
        recipient.apply(commit.patch()),
        Err(litchi_rtf::edit::Error::Rtf(litchi_rtf::Error::Limit {
            kind: litchi_rtf::SourceLimitKind::InputBytes,
            ..
        }))
    ));
    assert_eq!(recipient.to_bytes().unwrap(), SOURCE);
    let limits = SourceLimits::builder()
        .max_input_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    let recipient =
        Package::from_bytes_with_options(SOURCE, OpenOptions::new().source_limits(limits)).unwrap();
    let replay = recipient.apply(commit.patch()).unwrap();
    assert_eq!(replay.source_limits(), limits);
    assert_eq!(replay.text(), "a much longer replacement");
}
