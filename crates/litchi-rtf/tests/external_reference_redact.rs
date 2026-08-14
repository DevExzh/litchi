#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "redaction fixtures intentionally panic on malformed test setup"
)]

use litchi_rtf::Document;
use litchi_rtf::edit::Error as EditError;
use litchi_rtf::redact::{Mode, ReferenceKind, UnsupportedReference};
use serde_json::Value;

fn patch_limits(max_operations: usize) -> litchi_core::patch::PatchLimits {
    litchi_core::patch::PatchLimits::new(
        litchi_core::patch::BlobLimits::new(0, 0, 0),
        128 * 1024,
        max_operations,
        8,
        64 * 1024,
        256 * 1024,
    )
}

#[test]
fn inventories_exact_top_level_groups_and_redacts_forward_only() {
    let source = Document::parse(
        r#"{\rtf1\ansi{\*\template C:\\Templates\\base.dot}{\*\nextfile queue\\next.rtf}Body\par}"#,
    )
    .unwrap();
    let inventory = source.external_reference_redaction_snapshot().unwrap();
    assert_eq!(inventory.references().len(), 2);
    assert_eq!(inventory.references()[0].kind(), ReferenceKind::NextFile);
    assert_eq!(inventory.references()[1].kind(), ReferenceKind::Template);
    assert!(
        inventory
            .references()
            .iter()
            .all(|reference| reference.has_exact_source_range())
    );
    assert!(inventory.unsupported().is_empty());
    let commit = inventory
        .plan(Mode::Strict, patch_limits(2))
        .unwrap()
        .apply()
        .unwrap();
    assert_eq!(commit.diagnostics().removed_references(), 2);
    assert!(commit.diagnostics().is_complete());
    assert_eq!(commit.document().text(), source.text());
    assert!(commit.document().external_references().is_empty());
    let output = commit.document().to_bytes().unwrap();
    assert!(
        !output
            .windows(b"\\template".len())
            .any(|window| window == b"\\template")
    );
    assert!(
        !output
            .windows(b"\\nextfile".len())
            .any(|window| window == b"\\nextfile")
    );
    assert!(
        output
            .windows(b"\\ansi Body".len())
            .any(|window| window == b"\\ansi Body")
    );

    let replayed = source
        .apply_external_reference_redaction(commit.patch())
        .unwrap();
    assert_eq!(replayed.to_bytes().unwrap(), output);
    assert!(
        matches!(commit.patch().operations().first(), Some(operation) if operation.op == "external-reference.remove")
    );
}

#[test]
fn strict_refuses_external_fields_but_best_effort_is_explicitly_incomplete() {
    let source = Document::parse(
        r#"{\rtf1\ansi{\*\template template.dot}{\field{\*\fldinst HYPERLINK "https://example.invalid"}{\fldrslt Link}}Body}"#,
    )
    .unwrap();
    let inventory = source.external_reference_redaction_snapshot().unwrap();
    assert!(
        inventory
            .unsupported()
            .contains(&UnsupportedReference::Field)
    );
    assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());

    let commit = inventory
        .plan(Mode::BestEffort, patch_limits(2))
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.diagnostics().is_incomplete());
    assert!(
        commit
            .diagnostics()
            .unsupported()
            .contains(&UnsupportedReference::Field)
    );
    assert!(commit.document().external_references().is_empty());
    assert_eq!(commit.document().text(), source.text());
}

#[test]
fn strict_refuses_opaque_and_protected_surfaces() {
    let opaque =
        Document::parse(r#"{\rtf1\ansi{\*\unknown opaque}{\*\template t.dot}Body}"#).unwrap();
    let inventory = opaque.external_reference_redaction_snapshot().unwrap();
    assert!(
        inventory
            .unsupported()
            .contains(&UnsupportedReference::OpaqueSyntax)
    );
    assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());

    let protected =
        Document::parse(r#"{\rtf1\ansi\readprot\enforceprot1{\*\template t.dot}Body}"#).unwrap();
    let inventory = protected.external_reference_redaction_snapshot().unwrap();
    assert!(
        inventory
            .unsupported()
            .contains(&UnsupportedReference::Protection)
    );
    assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
}

#[test]
fn strict_refuses_every_object_kind_including_embedded_ole_and_linkself() {
    for source in [
        r#"{\rtf1\ansi{\object\objemb{\*\objdata 00}}{\*\template t.dot}Body}"#,
        r#"{\rtf1\ansi{\object\objocx\linkself{\*\objdata 00}}{\*\template t.dot}Body}"#,
    ] {
        let document = Document::parse(source).unwrap();
        let inventory = document.external_reference_redaction_snapshot().unwrap();
        assert!(
            inventory
                .unsupported()
                .contains(&UnsupportedReference::Object)
        );
        assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
    }
}

#[test]
fn strict_refuses_valid_legacy_formfield_destinations() {
    for source in [
        // These are valid legacy fields even without a separate formfield
        // property group; the active surface is represented by the generic
        // field store and must not depend on `form_fields()` being populated.
        r#"{\rtf1\ansi{\field{\*\fldinst FORMTEXT}{\fldrslt cached text}}}"#,
        r#"{\rtf1\ansi{\field{\*\fldinst FORMCHECKBOX}{\fldrslt cached checkbox}}}"#,
        r#"{\rtf1\ansi{\field{\*\fldinst FORMDROPDOWN}{\fldrslt cached selection}}}"#,
        // A fully modeled checkbox destination is covered as well.
        r#"{\rtf1\ansi{\field{\*\fldinst FORMCHECKBOX{\*\formfield{\fftype1\fftypetxt0\ffhps20\ffdefres0\ffres0}}}{\fldrslt }}}"#,
    ] {
        let source = Document::parse(source).unwrap();
        let inventory = source.external_reference_redaction_snapshot().unwrap();
        assert!(
            inventory
                .unsupported()
                .contains(&UnsupportedReference::Field)
        );
        assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
    }
}

#[test]
fn external_reference_unicode_fallback_rejects_active_control_words() {
    assert!(Document::parse(r#"{\rtf1\ansi{\*\template \u65\par}}"#).is_err());
}

#[test]
fn strict_shape_text_unicode_fallback_does_not_consume_active_controls() {
    let source = concat!(
        r#"{\rtf1\ansi{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
        r#"{\shptxt {\u65\par}}}}}"#,
    );
    let document = Document::parse(source).unwrap();
    // `\par` is a supported shape-text control, but it is not a Unicode
    // fallback character.  The owner must see it after the finite fallback
    // scanner stops, preserving the paragraph break.
    assert_eq!(document.shapes()[0].text, "A\n");
}

#[test]
fn strict_refuses_annotation_skipped_destinations() {
    for hidden in [
        r#"{\*\template hidden.dot}"#,
        // The unknown starred wrapper is skipped as one annotation-owned
        // group; active object/picture syntax inside it must not disappear
        // from strict diagnostics.
        r#"{\*\vendor{\object\objemb{\*\objdata 00}}}"#,
        r#"{\*\vendor{\pict\pngblip 89504e470d0a1a0a}}"#,
    ] {
        let source = [
            r#"{\rtf1\ansi{\*\atnid I}{\*\atnauthor A}\chatn{\*\annotation visible"#,
            hidden,
            r#"}}"#,
        ]
        .concat();
        let document = Document::parse(&source)
            .unwrap_or_else(|error| panic!("annotation hidden destination {hidden}: {error:?}"));
        let inventory = document.external_reference_redaction_snapshot().unwrap();
        assert!(
            inventory
                .unsupported()
                .contains(&UnsupportedReference::UnknownSyntax),
            "annotation hidden destination was not diagnosed: {hidden}"
        );
        assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
    }
}

#[test]
fn strict_refuses_unretained_upr_ansi_fallbacks() {
    for hidden in [
        r#"{\*\template hidden.dot}"#,
        r#"{\field{\*\fldinst HYPERLINK "https://example.invalid"}{\fldrslt hidden}}"#,
        r#"{\object\objemb{\*\objdata 00}}"#,
    ] {
        let source = [
            r#"{\rtf1\ansi{\upr{ansi "#,
            hidden,
            r#"}{\*\ud\uc0 Unicode}}Body}"#,
        ]
        .concat();
        let document = Document::parse(&source)
            .unwrap_or_else(|error| panic!("upr hidden destination {hidden}: {error:?}"));
        let inventory = document.external_reference_redaction_snapshot().unwrap();
        assert!(
            inventory
                .unsupported()
                .contains(&UnsupportedReference::UnknownSyntax),
            "upr hidden destination was not diagnosed: {hidden}"
        );
        assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
    }
}

#[test]
fn external_reference_decoder_caps_transport_and_unicode_intermediates() {
    let mut transport = String::from(r#"{\rtf1\ansi{\*\template "#);
    transport.push_str(&"x".repeat(65_537));
    transport.push_str("}}");
    assert!(Document::parse(&transport).is_err());

    // The encoded transport is below 64 KiB, but a non-ASCII code page byte
    // can expand to up to four UTF-8 bytes.  The parser must reject before
    // invoking the decoder rather than allocating an oversized Cow.
    let mut decoded_intermediate = String::from(r#"{\rtf1\ansi{\*\template "#);
    decoded_intermediate.push_str(&"é".repeat(16_385));
    decoded_intermediate.push_str("}}");
    assert!(Document::parse(&decoded_intermediate).is_err());

    let mut unicode_fallback = String::from(r#"{\rtf1\ansi{\*\template \u65"#);
    unicode_fallback.push_str(&"a".repeat(65_538));
    unicode_fallback.push_str("}}");
    assert!(Document::parse(&unicode_fallback).is_err());
}

fn assert_strict_refuses_shape_hyperlink(source: &str) {
    let document = Document::parse(source).unwrap();
    let inventory = document.external_reference_redaction_snapshot().unwrap();
    assert!(
        inventory
            .unsupported()
            .contains(&UnsupportedReference::Shape)
    );
    assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
}

#[test]
fn strict_shape_diagnostics_cover_header_table_and_field_result_stories() {
    let shape = concat!(
        r#"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}"#,
        r#"{\sp{\sn hyperlink}{\sv }{\hl {\hlloc http://example.test/x}}}"#,
        r#"{\shptxt x}}}"#,
    );
    assert_strict_refuses_shape_hyperlink(
        &[r#"{\rtf1\ansi{\header H"#, shape, r#"T}Body}"#].concat(),
    );
    assert_strict_refuses_shape_hyperlink(
        &[
            r#"{\rtf1\ansi\trowd\cellx5000\intbl "#,
            shape,
            r#"\cell\row}"#,
        ]
        .concat(),
    );
    assert_strict_refuses_shape_hyperlink(
        &[
            r#"{\rtf1\ansi{\field{\*\fldinst TEST}{\fldrslt "#,
            shape,
            r#"}}}"#,
        ]
        .concat(),
    );
}

#[test]
fn strict_shape_text_diagnostics_refuse_skipped_active_destinations() {
    let prefix = r#"{\rtf1\ansi{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt x"#;
    let suffix = r#"}}}}"#;
    for nested in [
        // Each destination is syntactically nested inside shptxt and is
        // intentionally not retained by the visible shape-text story model.
        r#"{\*\template hidden.dot}"#,
        r#"{\field{\*\fldinst HYPERLINK "https://example.invalid"}{\fldrslt hidden}}"#,
        r#"{\object\objemb{\*\objdata 00}}"#,
        r#"{\pict\pngblip 89504e470d0a1a0a}"#,
    ] {
        let source = [prefix, nested, suffix].concat();
        let document = Document::parse(&source)
            .unwrap_or_else(|error| panic!("nested destination {nested} failed: {error:?}"));
        let inventory = document.external_reference_redaction_snapshot().unwrap();
        assert!(
            inventory
                .unsupported()
                .contains(&UnsupportedReference::UnknownSyntax),
            "nested shape-text destination was not diagnosed: {nested}"
        );
        assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());
    }
}

#[test]
fn compressed_transport_has_inventory_but_no_source_span_and_best_effort_noop() {
    let source = br#"{\rtf1\ansi{\*\template t.dot}Body}"#;
    let compressed = litchi_rtf::compress(source, true).unwrap();
    let document = Document::from_bytes(&compressed).unwrap();
    let inventory = document.external_reference_redaction_snapshot().unwrap();
    assert!(
        inventory
            .unsupported()
            .contains(&UnsupportedReference::MissingSourceSpan(
                ReferenceKind::Template
            ))
    );
    assert!(inventory.plan(Mode::Strict, patch_limits(2)).is_err());

    let commit = inventory
        .plan(Mode::BestEffort, patch_limits(2))
        .unwrap()
        .apply()
        .unwrap();
    assert!(commit.diagnostics().is_incomplete());
    assert_eq!(commit.diagnostics().removed_references(), 0);
    assert_eq!(commit.document().to_bytes().unwrap(), compressed);
}

#[test]
fn foreign_and_stale_artifacts_are_rejected_before_candidate_publication() {
    let source = Document::parse(r#"{\rtf1\ansi{\*\template t.dot}Body}"#).unwrap();
    let commit = source
        .external_reference_redaction_snapshot()
        .unwrap()
        .plan(Mode::Strict, patch_limits(1))
        .unwrap()
        .apply()
        .unwrap();
    let foreign = Document::parse(r#"{\rtf1\ansi{\*\template other.dot}Body}"#).unwrap();
    assert!(matches!(
        foreign.apply_external_reference_redaction(commit.patch()),
        Err(EditError::PatchConflict)
    ));
    let stale = Document::parse(r#"{\rtf1\ansi{\*\template t.dot}Changed}"#).unwrap();
    assert!(matches!(
        stale.apply_external_reference_redaction(commit.patch()),
        Err(EditError::PatchConflict)
    ));
}

#[test]
fn exact_raw_span_is_authenticated_before_redaction() {
    let source = Document::parse(r#"{\rtf1\ansi{\*\template t.dot}Body}"#).unwrap();
    let inventory = source.external_reference_redaction_snapshot().unwrap();
    let occurrence = inventory
        .references()
        .iter()
        .find(|reference| reference.kind() == ReferenceKind::Template)
        .unwrap();
    let range = occurrence.source_range().unwrap();
    let bytes = source.to_bytes().unwrap();
    assert_eq!(&bytes[range], br#"{\*\template t.dot}"#);

    let commit = inventory
        .plan(Mode::Strict, patch_limits(1))
        .unwrap()
        .apply()
        .unwrap();
    let mut operations = commit.patch().operations().to_vec();
    operations[0]
        .preconditions
        .insert("source_span".to_string(), Value::String("0:1".to_string()));
    let forged = litchi_core::patch::Patch::<litchi_core::patch::ForwardOnly>::new(
        patch_limits(1),
        "litchi-rtf",
        operations,
        litchi_core::patch::BlobBundle::new(patch_limits(1).blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_external_reference_redaction(&forged),
        Err(EditError::PatchConflict)
    ));
}

#[test]
fn forged_oversized_operation_batch_is_rejected_before_selection_allocation() {
    let source = Document::parse(r#"{\rtf1\ansi{\*\template t.dot}Body}"#).unwrap();
    let commit = source
        .external_reference_redaction_snapshot()
        .unwrap()
        .plan(Mode::Strict, patch_limits(1))
        .unwrap()
        .apply()
        .unwrap();
    let operation = commit.patch().operations()[0].clone();
    let patch_limits = patch_limits(3);
    let forged = litchi_core::patch::Patch::<litchi_core::patch::ForwardOnly>::new(
        patch_limits,
        "litchi-rtf",
        vec![operation.clone(), operation.clone(), operation],
        litchi_core::patch::BlobBundle::new(patch_limits.blobs()),
    )
    .unwrap();
    assert!(matches!(
        source.apply_external_reference_redaction(&forged),
        Err(EditError::Rtf(litchi_rtf::RtfError::LimitExceeded {
            resource: "external-reference patch operations",
            observed: 3,
            limit: 2,
        }))
    ));
}

#[test]
fn zero_operation_patch_has_no_cross_document_authority_and_is_exact_noop() {
    let source = Document::parse(r#"{\rtf1\ansi Plain}"#).unwrap();
    let foreign = Document::parse(r#"{\rtf1\ansi{\*\template t.dot}Foreign}"#).unwrap();
    let limits = patch_limits(0);
    let patch = litchi_core::patch::Patch::<litchi_core::patch::ForwardOnly>::new(
        limits,
        "litchi-rtf",
        Vec::<litchi_core::patch::PatchOperation>::new(),
        litchi_core::patch::BlobBundle::new(limits.blobs()),
    )
    .unwrap();
    let applied = foreign.apply_external_reference_redaction(&patch).unwrap();
    assert_eq!(applied.to_bytes().unwrap(), foreign.to_bytes().unwrap());
    assert_eq!(applied.text(), foreign.text());
    assert!(source.apply_external_reference_redaction(&patch).is_ok());
}
