//! Primary commits carry their native package and its retained owner policy.

use crate::{Package, Result};
use litchi_opc::{BlobPart, PackURI};

fn opened() -> Result<Package> {
    let mut builder = crate::builder::Builder::new()?;
    for text in ["owner-secret-text", "second slide"] {
        builder
            .presentation()?
            .add_slide()?
            .add_text_box(text, 0, 0, 914_400, 914_400);
    }
    let authored = builder.build()?;
    Package::from_reader(std::io::Cursor::new(authored.to_bytes()?))
}

#[test]
fn direct_commit_and_immutable_replay_preserve_their_sources() -> Result<()> {
    let source = opened()?;
    let original = source.to_bytes()?;
    let mut edit = source.edit()?;
    edit.set_shape_text(0_usize, 0_usize, "changed")?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let (changed, patch) = commit.into_parts();
    assert_eq!(source.to_bytes()?, original);
    let output = changed.to_bytes()?;
    let replayed = source.apply(&patch)?.into_package();
    assert_eq!(replayed.to_bytes()?, output);
    assert_eq!(source.to_bytes()?, original);
    let restored = changed.apply(&patch.inverse())?.into_package();
    assert_eq!(restored.to_bytes()?, original);
    assert_eq!(changed.to_bytes()?, output);
    assert!(changed.apply(&patch).is_err());
    assert_eq!(changed.to_bytes()?, output);
    Ok(())
}

#[test]
fn cloned_commits_and_snapshot_derived_edits_keep_owner_evidence() -> Result<()> {
    let mut source = opened()?;
    source
        .tolerance
        .record(crate::leniency::TimingDefect::OrphanedTarget, 42);
    let report = source.tolerance_report().clone();
    let snapshot = source.opened_presentation()?;
    let mut edit = snapshot.edit();
    edit.set_shape_text(0_usize, 0_usize, "changed")?;
    let commit = edit.commit()?;
    assert_eq!(commit.package().tolerance_report(), &report);
    let first = commit.clone().into_package();
    let second = commit.into_package();
    assert_eq!(first.tolerance_report(), &report);
    assert_eq!(second.tolerance_report(), &report);
    assert_eq!(
        first.physical_source_provenance,
        source.physical_source_provenance
    );
    assert_eq!(first.to_bytes()?, second.to_bytes()?);
    let next = first.opened_presentation()?.edit().commit()?.into_package();
    assert_eq!(next.tolerance_report(), &report);
    let noop = source.edit()?.commit()?;
    assert!(!noop.diagnostics().changed());
    assert_eq!(
        source
            .apply(noop.patch())?
            .into_package()
            .tolerance_report(),
        &report
    );
    let plan = source.opened_presentation()?.plan_slide_removal(1_usize)?;
    let result = source.apply_slide_removal_plan(&plan)?;
    assert_eq!(
        result.edit().commit()?.package().tolerance_report(),
        &report
    );
    Ok(())
}

#[test]
fn direct_publication_keeps_the_candidate_independent_of_source_drift() -> Result<()> {
    let mut source = opened()?;
    let mut edit = source.edit()?;
    edit.set_shape_text(0_usize, 0_usize, "changed")?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let extra = PackURI::new("/private-drift.bin").map_err(crate::Error::Invalid)?;
    std::sync::Arc::make_mut(&mut source.opc).try_add_part(Box::new(BlobPart::new(
        extra.clone(),
        "application/octet-stream".into(),
        b"later".to_vec(),
    )))?;
    let direct = commit.into_package();
    assert!(!direct.opc.contains_part(&extra));
    assert!(source.opc.contains_part(&extra));
    // Durable patch replay deliberately uses the destination's unrelated state.
    let replayed = source.apply(&patch)?.into_package();
    assert!(replayed.opc.contains_part(&extra));
    Ok(())
}

#[test]
fn no_op_package_and_commit_debug_do_not_disclose_document_content() -> Result<()> {
    let source = opened()?;
    let original = source.to_bytes()?;
    let mut edit = source.edit()?;
    let edit_debug = format!("{edit:?}");
    let source_debug = format!("{:?}", edit.source()?);
    let commit = edit.commit()?;
    let debug = format!("{commit:?} {:?} {:?}", commit.package(), commit.patch());
    for value in [&edit_debug, &source_debug, &debug] {
        for secret in ["owner-secret-text", "/ppt/", "rId", "<p:", "<a:"] {
            assert!(!value.contains(secret), "debug leaked {secret}");
        }
    }
    let cloned = commit.clone().into_package();
    let moved = commit.into_package();
    assert_eq!(cloned.to_bytes()?, original);
    assert_eq!(moved.to_bytes()?, original);
    Ok(())
}

#[cfg(feature = "encryption")]
#[test]
fn encryption_provenance_survives_commits_clones_replay_and_reencryption() -> Result<()> {
    use litchi_crypto::ooxml::Mode;
    let source = opened()?;
    let encrypted = source.to_encrypted("test password", Mode::Standard)?;
    let source =
        Package::from_reader_with_password(std::io::Cursor::new(encrypted), "test password")?;
    let noop = source.edit()?.commit()?;
    let replayed_noop = source.apply(noop.patch())?.into_package();
    assert_eq!(replayed_noop.encryption(), Some(Mode::Standard));
    assert!(replayed_noop.to_bytes().is_err());
    let mut edit = source.opened_presentation()?.edit();
    edit.set_shape_text(0_usize, 0_usize, "encrypted change")?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let mut cloned = commit.clone().into_package();
    let mut changed = commit.into_package();
    for package in [&mut cloned, &mut changed] {
        assert_eq!(package.encryption(), Some(Mode::Standard));
        assert!(package.to_bytes().is_err());
    }
    let replayed = source.apply(&patch)?.into_package();
    assert_eq!(replayed.encryption(), Some(Mode::Standard));
    assert!(replayed.to_bytes().is_err());
    let bytes = changed.to_reencrypted("new password")?;
    let reopened = Package::from_reader_with_password(std::io::Cursor::new(bytes), "new password")?;
    assert_eq!(reopened.encryption(), Some(Mode::Standard));
    assert!(
        reopened.presentation()?.slides()?[0]
            .text()?
            .contains("encrypted change")
    );
    Ok(())
}

#[test]
fn signed_replay_refuses_changed_sources_and_signature_injection() -> Result<()> {
    let bytes = std::fs::read("../../test-data/poi/test-data/xmldsign/hello-world-signed.pptx")?;
    let signed = Package::from_reader(std::io::Cursor::new(bytes.clone()))?;
    let noop = signed.edit()?.commit()?;
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.clone().into_package().to_bytes()?, bytes);
    assert_eq!(
        signed.apply(noop.patch())?.into_package().to_bytes()?,
        bytes
    );
    let mut unsigned = Package::from_reader(std::io::Cursor::new(bytes.clone()))?;
    unsigned = unsigned.unsign()?;
    let mut edit = unsigned.edit()?;
    edit.add_text_box(0_usize, "changed", (0, 0, 914_400, 914_400))?;
    let commit = edit.commit()?;
    assert!(matches!(
        signed.apply(commit.patch()),
        Err(crate::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert_eq!(signed.to_bytes()?, bytes);
    let snapshot = unsigned.opened_presentation()?;
    let injection = super::Patch::capture(
        &unsigned.opc,
        &signed.opc,
        snapshot.presentation_name,
        super::Limits::default(),
    )?;
    let before = unsigned.to_bytes()?;
    assert!(matches!(
        unsigned.apply(&injection),
        Err(crate::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
    assert_eq!(unsigned.to_bytes()?, before);
    Ok(())
}
