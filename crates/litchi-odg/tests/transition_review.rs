#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{
    BlobId, CompositionError, CompositionLimitKind, CompositionLimits, DiagnosticFingerprint,
    Error, SubEditJoinFailure,
};
use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Package, Transition, edit};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_AUTOMATIC: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual" presentation:transition-style="fade-from-left"/></style:style><style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:transition-style="dissolve"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/><draw:page draw:name="Page 2" draw:style-name="dp2"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/><draw:page draw:name="Page 2" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_SHAPE_ALIAS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1"><draw:rect draw:name="Box"/></draw:page></office:drawing></office:body></office:document-content>"#;

const CONTENT_INHERITED_SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="base"/><draw:page draw:name="Page 2" draw:style-name="parent"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_NAMED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const STYLES_NAMED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:styles></office:document-styles>"#;

const CONTENT_INHERITED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="child"/></office:drawing></office:body></office:document-content>"#;

const STYLES_INHERITED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="parent" style:family="drawing-page"><style:drawing-page-properties presentation:transition-style="fade-from-left"/></style:style><style:style style:name="child" style:family="drawing-page" style:parent-style-name="parent"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:styles></office:document-styles>"#;

const STYLES_INHERITED_SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="base" style:family="drawing-page"/><style:style style:name="parent" style:family="drawing-page"><style:drawing-page-properties presentation:transition-style="fade-from-left"/></style:style><style:style style:name="child" style:family="drawing-page" style:parent-style-name="parent"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:styles></office:document-styles>"#;

const CONTENT_FOREIGN_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:example:foreign-presentation" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:vendor="keep"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_LOCAL_FOREIGN_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:litchi_presentation="urn:example:occupied" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties xmlns:presentation="urn:example:foreign-presentation" presentation:vendor="keep"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

fn package(content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some(styles) = styles {
        writer.add_file("styles.xml", styles.as_bytes()).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn transition(transition_type: &'static str, style: &'static str) -> Transition {
    let mut value = Transition::new();
    value.set_transition_type(Some(transition_type)).unwrap();
    value.set_style(Some(style)).unwrap();
    value
}

fn transition_with_smil() -> Transition {
    let mut value = transition("automatic", "dissolve");
    value.set_smil_type(Some("fade")).unwrap();
    value
}

fn assert_exact_durable_round_trip(source: Package, value: Transition) {
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(value)).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let replayed = durable.apply(&source).unwrap();
    assert_eq!(replayed.as_bytes(), commit.package().as_bytes());
    let restored = durable.inverse().apply(&replayed).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

fn mixed_replacement_page_name(
    source: &Package,
    template_page: usize,
    target_page: usize,
    target_before: &str,
    target_after: &str,
    added_name: &str,
) -> (edit::DurablePatch, Package) {
    let mut structural_edit = source.edit();
    structural_edit
        .add_page(litchi_odg::page::Page::new(added_name))
        .unwrap();
    let structural = structural_edit.commit().unwrap();
    let mut mixed_wire: serde_json::Value = serde_json::from_slice(
        &structural
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();

    let mut name_edit = source.edit();
    name_edit
        .set_page_name(template_page, "Template rename")
        .unwrap();
    let name_wire: serde_json::Value = serde_json::from_slice(
        &name_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    mixed_wire["operations"]
        .as_array_mut()
        .unwrap()
        .push(name_wire["operations"][0].clone());
    let target = format!("page/{target_page}");
    mixed_wire["operations"][1]["forward"]["target"] = serde_json::Value::String(target.clone());
    mixed_wire["operations"][1]["forward"]["value"] =
        serde_json::Value::String(target_after.to_owned());
    mixed_wire["operations"][1]["inverse"]["target"] = serde_json::Value::String(target);
    mixed_wire["operations"][1]["inverse"]["value"] =
        serde_json::Value::String(target_before.to_owned());
    mixed_wire["operations"][1]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String("0".repeat(64));

    let provisional =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&mixed_wire).unwrap())
            .unwrap();
    let expected = provisional.apply(source).unwrap();
    let final_source = DiagnosticFingerprint::of(expected.as_bytes()).as_hex();
    mixed_wire["operations"][0]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String(final_source.clone());
    mixed_wire["operations"][1]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String(final_source);
    let mixed =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&mixed_wire).unwrap())
            .unwrap();
    (mixed, expected)
}

fn signed_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    let manifest = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="META-INF/documentsignatures.xml" manifest:media-type="application/vnd.oasis.opendocument.package"/></manifest:manifest>"#;
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", manifest)
        .unwrap();
    archive
        .write_deflated(
            "META-INF/documentsignatures.xml",
            br"<document-signatures/>",
        )
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

fn base64_encode(bytes: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut output = String::with_capacity(bytes.len().div_ceil(3) * 4);
    for chunk in bytes.chunks(3) {
        let first = chunk[0];
        let second = chunk.get(1).copied().unwrap_or(0);
        let third = chunk.get(2).copied().unwrap_or(0);
        output.push(TABLE[(first >> 2) as usize] as char);
        output.push(TABLE[((first & 0x03) << 4 | (second >> 4)) as usize] as char);
        output.push(if chunk.len() > 1 {
            TABLE[((second & 0x0f) << 2 | (third >> 6)) as usize] as char
        } else {
            '='
        });
        output.push(if chunk.len() > 2 {
            TABLE[(third & 0x3f) as usize] as char
        } else {
            '='
        });
    }
    output
}

#[test]
fn equal_transition_is_a_noop_even_when_the_style_is_shared() {
    let source = Package::from_bytes(&package(CONTENT_SHARED, None)).unwrap();
    let current = source.pages()[0].transition().cloned();
    let mut edit = source.edit();
    edit.set_page_transition(0, current).unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().changes().is_empty());
    assert_eq!(commit.package().as_bytes(), source.as_bytes());
    assert_eq!(commit.patch().durable().unwrap().operations().len(), 0);
}

#[test]
fn inherited_transition_can_be_replaced_but_cannot_be_cleared_locally() {
    let source = Package::from_bytes(&package(CONTENT_INHERITED, Some(STYLES_INHERITED))).unwrap();
    assert_eq!(
        source.pages()[0].transition().unwrap().style(),
        Some("fade-from-left")
    );

    let mut clear = source.edit();
    assert!(clear.set_page_transition(0, None).is_err());

    let mut replacement = source.edit();
    replacement
        .set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let output = replacement.commit().unwrap().into_package();
    assert_eq!(
        output.pages()[0].transition().unwrap().style(),
        Some("dissolve")
    );
    assert!(
        output
            .styles_xml()
            .unwrap()
            .contains("style:name=\"parent\"")
    );
    assert!(
        output
            .styles_xml()
            .unwrap()
            .contains("presentation:transition-style=\"fade-from-left\"")
    );
}

#[test]
fn named_style_transition_replays_and_inverts_durably() {
    let source = Package::from_bytes(&package(CONTENT_NAMED, Some(STYLES_NAMED))).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "page.transition.set");

    let replayed = durable.apply(&source).unwrap();
    assert_eq!(replayed.as_bytes(), commit.package().as_bytes());
    let restored = durable.inverse().apply(&replayed).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn transition_inverse_preserves_new_namespace_bindings_for_automatic_and_named_styles() {
    let automatic_content = CONTENT_AUTOMATIC.replace(
        r#" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0""#,
        "",
    );
    assert_exact_durable_round_trip(
        Package::from_bytes(&package(&automatic_content, None)).unwrap(),
        transition_with_smil(),
    );
    assert_exact_durable_round_trip(
        Package::from_bytes(&package(CONTENT_NAMED, Some(STYLES_NAMED))).unwrap(),
        transition_with_smil(),
    );
}

#[test]
fn transition_inverse_preserves_first_property_insertion_for_automatic_and_named_styles() {
    let automatic_content = CONTENT_AUTOMATIC.replace(
        r#"<style:drawing-page-properties presentation:transition-type="manual" presentation:transition-style="fade-from-left"/></style:style>"#,
        "</style:style>",
    );
    assert_exact_durable_round_trip(
        Package::from_bytes(&package(&automatic_content, None)).unwrap(),
        transition("automatic", "dissolve"),
    );

    let named_styles = STYLES_NAMED.replace(
        r#"<style:drawing-page-properties presentation:transition-type="manual"/></style:style>"#,
        "</style:style>",
    );
    assert_exact_durable_round_trip(
        Package::from_bytes(&package(CONTENT_NAMED, Some(&named_styles))).unwrap(),
        transition("automatic", "dissolve"),
    );
}

#[test]
fn transition_edit_does_not_rebind_a_foreign_presentation_prefix() {
    let source = Package::from_bytes(&package(CONTENT_FOREIGN_PRESENTATION, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit
            .package()
            .content_xml()
            .contains("presentation:vendor=\"keep\"")
    );
    assert!(commit.package().content_xml().contains(
        "xmlns:litchi_presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\""
    ));
    let durable = commit.patch().durable().unwrap();
    let restored = durable
        .inverse()
        .apply(&durable.apply(&source).unwrap())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn transition_edit_avoids_foreign_alias_and_local_shadowing() {
    let source = Package::from_bytes(&package(CONTENT_LOCAL_FOREIGN_PRESENTATION, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    let content = commit.package().content_xml();
    assert!(content.contains("xmlns:litchi_presentation=\"urn:example:occupied\""));
    assert!(content.contains("xmlns:presentation=\"urn:example:foreign-presentation\""));
    assert!(content.contains(
        "xmlns:litchi_presentation1=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\""
    ));
    assert!(content.contains("litchi_presentation1:transition-type=\"automatic\""));
    assert!(content.contains("presentation:vendor=\"keep\""));
    let durable = commit.patch().durable().unwrap();
    let restored = durable
        .inverse()
        .apply(&durable.apply(&source).unwrap())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn disjoint_transition_sub_edits_join_and_same_page_edits_conflict() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let limits = CompositionLimits::new(4, 4, 8, 8);

    let mut first_edit = source.edit();
    first_edit
        .set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let first = first_edit.commit().unwrap();

    let mut second_edit = source.edit();
    second_edit
        .set_page_transition(1, Some(transition("manual", "fade-from-left")))
        .unwrap();
    let second = second_edit.commit().unwrap();

    let mut joined = source.joined_edits(limits);
    joined
        .join(first.patch().prepare("first", limits).unwrap())
        .unwrap();
    joined
        .join(second.patch().prepare("second", limits).unwrap())
        .unwrap();
    let mut joined_edit = source.edit();
    joined_edit.join(joined).unwrap();
    let output = joined_edit.commit().unwrap().into_package();
    assert_eq!(
        output.pages()[0].transition().unwrap().style(),
        Some("dissolve")
    );
    assert_eq!(
        output.pages()[1].transition().unwrap().style(),
        Some("fade-from-left")
    );

    let mut competing_edit = source.edit();
    competing_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let competing = competing_edit.commit().unwrap();
    let mut conflicting = source.joined_edits(limits);
    conflicting
        .join(first.patch().prepare("left", limits).unwrap())
        .unwrap();
    let error = conflicting
        .join(competing.patch().prepare("right", limits).unwrap())
        .unwrap_err();
    assert!(matches!(error.failure(), SubEditJoinFailure::Overlap(_)));
}

#[test]
fn joined_transition_rejects_a_preceding_page_style_retarget() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);

    let mut transition_edit = source.edit();
    transition_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let transition_commit = transition_edit.commit().unwrap();
    let mut joined = source.joined_edits(limits);
    joined
        .join(
            transition_commit
                .patch()
                .prepare("transition", limits)
                .unwrap(),
        )
        .unwrap();

    let mut style_edit = source.edit();
    style_edit.set_page_style_name(0, "dp2").unwrap();
    let error = match style_edit.join(joined) {
        Ok(_) => panic!("style and transition effects were treated as disjoint"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    let committed = style_edit.commit().unwrap();
    assert_eq!(committed.package().pages()[0].style_name(), Some("dp2"));
}

#[test]
fn page_style_retarget_conflicts_with_shared_transition_in_any_order() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut style_edit = source.edit();
    style_edit.set_page_style_name(0, "dp2").unwrap();
    let style_commit = style_edit.commit().unwrap();
    let mut transition_edit = source.edit();
    transition_edit
        .set_page_transition(1, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let transition_commit = transition_edit.commit().unwrap();

    let make_plan = |style_id: &str, transition_id: &str| {
        let mut joined = source.joined_edits(limits);
        joined
            .join(style_commit.patch().prepare(style_id, limits).unwrap())
            .unwrap();
        joined
            .join(
                transition_commit
                    .patch()
                    .prepare(transition_id, limits)
                    .unwrap(),
            )
            .unwrap();
        joined
    };

    for (style_id, transition_id) in [("a-style", "z-transition"), ("z-style", "a-transition")] {
        let mut owner = source.edit();
        let error = match owner.join(make_plan(style_id, transition_id)) {
            Ok(_) => panic!("dependent style and transition effects were accepted"),
            Err(error) => error,
        };
        assert!(matches!(
            error.error(),
            Error::Refused(refusal)
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
        ));
        assert!(!owner.commit().unwrap().diagnostics().changed());
    }

    let mut ordinary_before = source.edit();
    ordinary_before.set_page_style_name(0, "dp2").unwrap();
    let error = match ordinary_before.join({
        let mut joined = source.joined_edits(limits);
        joined
            .join(
                transition_commit
                    .patch()
                    .prepare("ordinary-before", limits)
                    .unwrap(),
            )
            .unwrap();
        joined
    }) {
        Ok(_) => panic!("ordinary style retarget was accepted before join"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(ordinary_before.commit().is_ok());

    let mut ordinary_after = source.edit();
    ordinary_after
        .join({
            let mut joined = source.joined_edits(limits);
            joined
                .join(
                    transition_commit
                        .patch()
                        .prepare("ordinary-after", limits)
                        .unwrap(),
                )
                .unwrap();
            joined
        })
        .unwrap();
    ordinary_after.set_page_style_name(0, "dp2").unwrap();
    assert!(ordinary_after.commit().is_err());
}

#[test]
fn inherited_page_style_dependencies_are_checked_in_both_orders_and_staging_positions() {
    let source = Package::from_bytes(&package(
        CONTENT_INHERITED_SHARED,
        Some(STYLES_INHERITED_SHARED),
    ))
    .unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut style_edit = source.edit();
    style_edit.set_page_style_name(0, "child").unwrap();
    let style_commit = style_edit.commit().unwrap();
    let mut transition_edit = source.edit();
    transition_edit
        .set_page_transition(1, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let transition_commit = transition_edit.commit().unwrap();

    let make_plan = |style_id: &str, transition_id: &str| {
        let mut joined = source.joined_edits(limits);
        joined
            .join(style_commit.patch().prepare(style_id, limits).unwrap())
            .unwrap();
        joined
            .join(
                transition_commit
                    .patch()
                    .prepare(transition_id, limits)
                    .unwrap(),
            )
            .unwrap();
        joined
    };

    for (style_id, transition_id) in [
        ("a-inherited-style", "z-inherited-transition"),
        ("z-inherited-style", "a-inherited-transition"),
    ] {
        let mut owner = source.edit();
        let error = match owner.join(make_plan(style_id, transition_id)) {
            Ok(_) => panic!("inherited style and transition effects were accepted"),
            Err(error) => error,
        };
        assert!(matches!(
            error.error(),
            Error::Refused(refusal)
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
        ));
        assert!(!owner.commit().unwrap().diagnostics().changed());
    }

    let mut ordinary_before = source.edit();
    ordinary_before.set_page_style_name(0, "child").unwrap();
    assert!(
        ordinary_before
            .join({
                let mut joined = source.joined_edits(limits);
                joined
                    .join(
                        transition_commit
                            .patch()
                            .prepare("inherited-ordinary-before", limits)
                            .unwrap(),
                    )
                    .unwrap();
                joined
            })
            .is_err()
    );
    assert!(ordinary_before.commit().is_ok());

    let mut ordinary_after = source.edit();
    ordinary_after
        .join({
            let mut joined = source.joined_edits(limits);
            joined
                .join(
                    transition_commit
                        .patch()
                        .prepare("inherited-ordinary-after", limits)
                        .unwrap(),
                )
                .unwrap();
            joined
        })
        .unwrap();
    ordinary_after.set_page_style_name(0, "child").unwrap();
    assert!(ordinary_after.commit().is_err());
}

#[test]
fn projected_sub_edit_style_state_conflicts_with_a_later_shared_retarget() {
    let content = CONTENT_AUTOMATIC.replace(
        r#"<draw:page draw:name="Page 2" draw:style-name="dp2"/>"#,
        r#"<draw:page draw:name="Page 2" draw:style-name="dp1"/>"#,
    );
    let source = Package::from_bytes(&package(&content, None)).unwrap();
    let mut ordered_edit = source.edit();
    ordered_edit.set_page_style_name(0, "dp2").unwrap();
    ordered_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let ordered_commit = ordered_edit.commit().unwrap();
    assert_eq!(ordered_commit.patch().changes().len(), 2);

    let mut retarget_edit = source.edit();
    retarget_edit.set_page_style_name(1, "dp2").unwrap();
    let retarget_commit = retarget_edit.commit().unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    for (ordered_id, retarget_id) in [("a-ordered", "z-retarget"), ("z-ordered", "a-retarget")] {
        let mut joined = source.joined_edits(limits);
        joined
            .join(ordered_commit.patch().prepare(ordered_id, limits).unwrap())
            .unwrap();
        joined
            .join(
                retarget_commit
                    .patch()
                    .prepare(retarget_id, limits)
                    .unwrap(),
            )
            .unwrap();
        let mut owner = source.edit();
        let error = match owner.join(joined) {
            Ok(_) => panic!("projected shared style dependency was accepted"),
            Err(error) => error,
        };
        assert!(matches!(
            error.error(),
            Error::Refused(refusal)
                if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
        ));
        assert!(!owner.commit().unwrap().diagnostics().changed());
    }
}

#[test]
fn superseded_style_retargets_do_not_conflict_with_another_group() {
    let content = CONTENT_AUTOMATIC
        .replace(
            r#"<draw:page draw:name="Page 1" draw:style-name="dp1"/>"#,
            r#"<draw:page draw:name="Page 1" draw:style-name="dp2"/>"#,
        )
        .replace(
            r#"<draw:page draw:name="Page 2" draw:style-name="dp2"/>"#,
            r#"<draw:page draw:name="Page 2" draw:style-name="dp1"/>"#,
        );
    let source = Package::from_bytes(&package(&content, None)).unwrap();
    let mut style_edit = source.edit();
    style_edit.set_page_style_name(0, "dp1").unwrap();
    let first_style_wire: serde_json::Value = serde_json::from_slice(
        &style_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let first_style_operation = first_style_wire["operations"][0].clone();
    let mut second_style_operation = first_style_operation.clone();
    second_style_operation["forward"]["value"] = serde_json::Value::String("dp2".to_owned());
    second_style_operation["inverse"]["value"] = serde_json::Value::String("dp1".to_owned());
    let mut style_wire = first_style_wire;
    style_wire["operations"] =
        serde_json::Value::Array(vec![first_style_operation, second_style_operation]);
    let style_payload =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&style_wire).unwrap())
            .unwrap();
    assert_eq!(style_payload.operations().len(), 2);

    let mut transition_edit = source.edit();
    transition_edit
        .set_page_transition(1, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let transition_commit = transition_edit.commit().unwrap();
    let mut same_page_transition_edit = source.edit();
    same_page_transition_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let same_page_transition_commit = same_page_transition_edit.commit().unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let make_plan = |style_id: &str, transition_id: &str| {
        let mut joined = source.joined_edits(limits);
        let lineage = joined.lineage().clone();
        joined
            .join(
                edit::SubEdit::new(
                    lineage,
                    limits,
                    style_id,
                    Vec::<String>::new(),
                    vec!["page/0#page.style.set".to_owned()],
                    style_payload.clone(),
                )
                .unwrap(),
            )
            .unwrap();
        joined
            .join(
                transition_commit
                    .patch()
                    .prepare(transition_id, limits)
                    .unwrap(),
            )
            .unwrap();
        joined
    };

    for (style_id, transition_id) in [
        ("a-superseded-style", "z-superseded-transition"),
        ("z-superseded-style", "a-superseded-transition"),
    ] {
        let mut owner = source.edit();
        owner.join(make_plan(style_id, transition_id)).unwrap();
        let commit = owner.commit().unwrap();
        assert_eq!(commit.package().pages()[0].style_name(), Some("dp2"));
        assert_eq!(
            commit.package().pages()[1].transition().unwrap().style(),
            Some("fade-from-right")
        );
    }

    let mut same_page_joined = source.joined_edits(limits);
    let same_page_lineage = same_page_joined.lineage().clone();
    same_page_joined
        .join(
            edit::SubEdit::new(
                same_page_lineage,
                limits,
                "same-page-superseded-style",
                Vec::<String>::new(),
                vec!["page/0#page.style.set".to_owned()],
                style_payload.clone(),
            )
            .unwrap(),
        )
        .unwrap();
    same_page_joined
        .join(
            same_page_transition_commit
                .patch()
                .prepare("same-page-transition", limits)
                .unwrap(),
        )
        .unwrap();
    let mut same_page_owner = source.edit();
    same_page_owner.join(same_page_joined).unwrap();
    let same_page_commit = same_page_owner.commit().unwrap();
    assert_eq!(
        same_page_commit.package().pages()[0].style_name(),
        Some("dp2")
    );
    assert_eq!(
        same_page_commit.package().pages()[0]
            .transition()
            .unwrap()
            .style(),
        Some("fade-from-right")
    );

    let mut intermediate_wire: serde_json::Value =
        serde_json::from_slice(&style_payload.to_deterministic_json().unwrap()).unwrap();
    let transition_wire: serde_json::Value = serde_json::from_slice(
        &same_page_transition_commit
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let style_first_operation = intermediate_wire["operations"][0].clone();
    let style_last_operation = intermediate_wire["operations"][1].clone();
    intermediate_wire["operations"] = serde_json::Value::Array(vec![
        style_first_operation,
        transition_wire["operations"][0].clone(),
        style_last_operation,
    ]);
    let intermediate_payload = edit::DurablePatch::from_deterministic_json(
        &serde_json::to_vec(&intermediate_wire).unwrap(),
    )
    .unwrap();
    let mut intermediate_joined = source.joined_edits(limits);
    let intermediate_lineage = intermediate_joined.lineage().clone();
    intermediate_joined
        .join(
            edit::SubEdit::new(
                intermediate_lineage,
                limits,
                "intermediate-transition",
                Vec::<String>::new(),
                vec![
                    "page/0#page.style.set".to_owned(),
                    "page/0#page.transition.set".to_owned(),
                ],
                intermediate_payload,
            )
            .unwrap(),
        )
        .unwrap();
    intermediate_joined
        .join(
            transition_commit
                .patch()
                .prepare("intermediate-independent-transition", limits)
                .unwrap(),
        )
        .unwrap();
    let mut intermediate_owner = source.edit();
    let error = match intermediate_owner.join(intermediate_joined) {
        Ok(_) => panic!("intermediate style dependency was retired with its retarget"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(!intermediate_owner.commit().unwrap().diagnostics().changed());

    let mut ordinary_before = source.edit();
    ordinary_before.set_page_style_name(0, "dp1").unwrap();
    ordinary_before.set_page_style_name(0, "dp2").unwrap();
    ordinary_before
        .join({
            let mut joined = source.joined_edits(limits);
            joined
                .join(
                    transition_commit
                        .patch()
                        .prepare("superseded-ordinary-before", limits)
                        .unwrap(),
                )
                .unwrap();
            joined
        })
        .unwrap();
    let before_commit = ordinary_before.commit().unwrap();
    assert_eq!(before_commit.package().pages()[0].style_name(), Some("dp2"));

    let mut ordinary_after = source.edit();
    ordinary_after
        .join({
            let mut joined = source.joined_edits(limits);
            joined
                .join(
                    transition_commit
                        .patch()
                        .prepare("superseded-ordinary-after", limits)
                        .unwrap(),
                )
                .unwrap();
            joined
        })
        .unwrap();
    ordinary_after.set_page_style_name(0, "dp1").unwrap();
    ordinary_after.set_page_style_name(0, "dp2").unwrap();
    let after_commit = ordinary_after.commit().unwrap();
    assert_eq!(after_commit.package().pages()[0].style_name(), Some("dp2"));
}

#[test]
fn joined_page_name_renames_validate_final_names() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut first_edit = source.edit();
    first_edit.set_page_name(0, "First renamed").unwrap();
    let first = first_edit.commit().unwrap();
    let mut second_edit = source.edit();
    second_edit.set_page_name(1, "Second renamed").unwrap();
    let second = second_edit.commit().unwrap();

    let mut joined = source.joined_edits(limits);
    joined
        .join(first.patch().prepare("first", limits).unwrap())
        .unwrap();
    joined
        .join(second.patch().prepare("second", limits).unwrap())
        .unwrap();
    let mut owner = source.edit();
    owner.join(joined).unwrap();
    let committed = owner.commit().unwrap();
    assert_eq!(committed.diagnostics().change_count(), 2);
    assert_eq!(committed.package().pages()[0].name(), Some("First renamed"));
    assert_eq!(
        committed.package().pages()[1].name(),
        Some("Second renamed")
    );

    let mut duplicate_first = source.edit();
    duplicate_first.set_page_name(0, "Same final name").unwrap();
    let duplicate_first = duplicate_first.commit().unwrap();
    let mut duplicate_second = source.edit();
    duplicate_second
        .set_page_name(1, "Same final name")
        .unwrap();
    let duplicate_second = duplicate_second.commit().unwrap();
    let mut duplicate = source.joined_edits(limits);
    duplicate
        .join(
            duplicate_first
                .patch()
                .prepare("duplicate-first", limits)
                .unwrap(),
        )
        .unwrap();
    duplicate
        .join(
            duplicate_second
                .patch()
                .prepare("duplicate-second", limits)
                .unwrap(),
        )
        .unwrap();
    let mut unchanged = source.edit();
    let error = match unchanged.join(duplicate) {
        Ok(_) => panic!("duplicate final page names were accepted"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(!unchanged.commit().unwrap().diagnostics().changed());
}

#[test]
fn page_rename_preserves_unrelated_existing_duplicate_names() {
    let content = CONTENT_AUTOMATIC
        .replace("draw:name=\"Page 1\"", "draw:name=\"A\"")
        .replace("draw:name=\"Page 2\"", "draw:name=\"A\"")
        .replace(
            "</office:drawing>",
            r#"<draw:page draw:name="C" draw:style-name="dp1"/></office:drawing>"#,
        );
    let source = Package::from_bytes(&package(&content, None)).unwrap();
    assert_eq!(
        source
            .pages()
            .iter()
            .map(|page| page.name())
            .collect::<Vec<_>>(),
        [Some("A"), Some("A"), Some("C"),]
    );

    let mut direct = source.edit();
    direct.set_page_name(2, "D").unwrap();
    let direct_commit = direct.commit().unwrap();
    let durable = direct_commit.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        direct_commit.package().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(direct_commit.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut joined = source.joined_edits(limits);
    joined
        .join(direct_commit.patch().prepare("rename-c", limits).unwrap())
        .unwrap();
    let mut joined_owner = source.edit();
    joined_owner.join(joined).unwrap();
    let joined_commit = joined_owner.commit().unwrap();
    assert_eq!(
        joined_commit.package().as_bytes(),
        direct_commit.package().as_bytes()
    );
    assert_eq!(
        source
            .apply(joined_commit.patch())
            .unwrap()
            .package()
            .as_bytes(),
        direct_commit.package().as_bytes()
    );
    assert_eq!(
        joined_commit
            .package()
            .apply(&joined_commit.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );

    let mut direct_duplicate = source.edit();
    assert!(direct_duplicate.set_page_name(2, "A").is_err());
    let mut duplicate_wire: serde_json::Value =
        serde_json::from_slice(&durable.to_deterministic_json().unwrap()).unwrap();
    duplicate_wire["operations"][0]["forward"]["value"] = serde_json::Value::String("A".to_owned());
    let duplicate =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&duplicate_wire).unwrap())
            .unwrap();
    let mut duplicate_joined = source.joined_edits(limits);
    let duplicate_lineage = duplicate_joined.lineage().clone();
    let duplicate_sub_edit = edit::SubEdit::new(
        duplicate_lineage,
        limits,
        "duplicate-c",
        Vec::<String>::new(),
        vec!["page/2#page.name.set".to_owned()],
        duplicate,
    )
    .unwrap();
    duplicate_joined.join(duplicate_sub_edit).unwrap();
    let mut duplicate_owner = source.edit();
    let error = match duplicate_owner.join(duplicate_joined) {
        Ok(_) => panic!("duplicate touched page name was accepted"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(!duplicate_owner.commit().unwrap().diagnostics().changed());

    let (mixed, expected) = mixed_replacement_page_name(&source, 0, 2, "C", "D", "New");
    assert_eq!(
        mixed.apply(&source).unwrap().as_bytes(),
        expected.as_bytes()
    );
    assert_eq!(
        mixed.inverse().apply(&expected).unwrap().as_bytes(),
        source.as_bytes()
    );
    let mut replacement_joined = source.joined_edits(limits);
    let replacement_lineage = replacement_joined.lineage().clone();
    let replacement_sub_edit = edit::SubEdit::new(
        replacement_lineage,
        limits,
        "replacement-rename-c",
        Vec::<String>::new(),
        vec!["package".to_owned(), "page/2#page.name.set".to_owned()],
        mixed,
    )
    .unwrap();
    replacement_joined.join(replacement_sub_edit).unwrap();
    let mut replacement_owner = source.edit();
    replacement_owner.join(replacement_joined).unwrap();
    let replacement_commit = replacement_owner.commit().unwrap();
    assert_eq!(replacement_commit.package().as_bytes(), expected.as_bytes());
    assert_eq!(
        replacement_commit
            .package()
            .apply(&replacement_commit.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );

    let (replacement_duplicate, _) = mixed_replacement_page_name(&source, 0, 2, "C", "D", "New");
    let mut replacement_duplicate_wire: serde_json::Value =
        serde_json::from_slice(&replacement_duplicate.to_deterministic_json().unwrap()).unwrap();
    replacement_duplicate_wire["operations"][1]["forward"]["value"] =
        serde_json::Value::String("A".to_owned());
    let replacement_duplicate = edit::DurablePatch::from_deterministic_json(
        &serde_json::to_vec(&replacement_duplicate_wire).unwrap(),
    )
    .unwrap();
    let mut replacement_duplicate_joined = source.joined_edits(limits);
    let replacement_duplicate_lineage = replacement_duplicate_joined.lineage().clone();
    let replacement_duplicate_sub_edit = edit::SubEdit::new(
        replacement_duplicate_lineage,
        limits,
        "replacement-duplicate-c",
        Vec::<String>::new(),
        vec!["package".to_owned(), "page/2#page.name.set".to_owned()],
        replacement_duplicate,
    )
    .unwrap();
    replacement_duplicate_joined
        .join(replacement_duplicate_sub_edit)
        .unwrap();
    let mut replacement_duplicate_owner = source.edit();
    replacement_duplicate_owner
        .join(replacement_duplicate_joined)
        .unwrap();
    assert!(replacement_duplicate_owner.commit().is_err());
}

#[test]
fn joined_payload_effects_cover_forged_cross_sub_edit_writes() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let (forged_replacement, _) =
        mixed_replacement_page_name(&source, 0, 2, "Added", "Added renamed", "Added");
    let mut forged_plan = source.joined_edits(limits);
    let forged_lineage = forged_plan.lineage().clone();
    let forged_sub_edit = edit::SubEdit::new(
        forged_lineage,
        limits,
        "forged-replacement",
        Vec::<String>::new(),
        vec!["page/2#page.name.set".to_owned()],
        forged_replacement,
    )
    .unwrap();
    forged_plan.join(forged_sub_edit).unwrap();
    let mut ordinary = source.edit();
    ordinary.set_page_style_name(0, "dp2").unwrap();
    let error = match ordinary.join(forged_plan) {
        Ok(_) => panic!("underdeclared replacement payload was accepted"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    let recovered = error.into_rejected();
    let ordinary_commit = ordinary.commit().unwrap();
    assert_eq!(
        ordinary_commit.package().pages()[0].style_name(),
        Some("dp2")
    );
    let mut retry = source.edit();
    assert!(retry.join(recovered).is_err());

    let mut first_edit = source.edit();
    first_edit.set_page_name(0, "Duplicate once").unwrap();
    let first_wire: serde_json::Value = serde_json::from_slice(
        &first_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let mut second_edit = source.edit();
    second_edit.set_page_name(0, "Duplicate twice").unwrap();
    let second_wire: serde_json::Value = serde_json::from_slice(
        &second_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let first_operation = first_wire["operations"][0].clone();
    let second_operation = second_wire["operations"][0].clone();
    let mut duplicate_wire = first_wire;
    duplicate_wire["operations"] =
        serde_json::Value::Array(vec![first_operation, second_operation]);
    let duplicate =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&duplicate_wire).unwrap())
            .unwrap();
    let mut duplicate_plan = source.joined_edits(limits);
    let duplicate_lineage = duplicate_plan.lineage().clone();
    duplicate_plan
        .join(
            edit::SubEdit::new(
                duplicate_lineage,
                limits,
                "concealed-duplicate",
                Vec::<String>::new(),
                vec!["page/0#page.name.set".to_owned()],
                duplicate,
            )
            .unwrap(),
        )
        .unwrap();
    let mut duplicate_owner = source.edit();
    duplicate_owner.join(duplicate_plan).unwrap();
    let duplicate_commit = duplicate_owner.commit().unwrap();
    assert_eq!(
        duplicate_commit.package().pages()[0].name(),
        Some("Duplicate twice")
    );
    assert!(duplicate_commit.diagnostics().changed());

    let mut first_edit = source.edit();
    first_edit.set_page_name(0, "First cross-sub-edit").unwrap();
    let first = first_edit.commit().unwrap();
    let mut second_edit = source.edit();
    second_edit
        .set_page_name(0, "Second cross-sub-edit")
        .unwrap();
    let second = second_edit.commit().unwrap();
    let mut forged_cross_sub_edit = source.joined_edits(limits);
    let lineage = forged_cross_sub_edit.lineage().clone();
    forged_cross_sub_edit
        .join(
            edit::SubEdit::new(
                lineage.clone(),
                limits,
                "actual-page-zero",
                Vec::<String>::new(),
                vec!["page/0#page.name.set".to_owned()],
                first.patch().durable().unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    forged_cross_sub_edit
        .join(
            edit::SubEdit::new(
                lineage,
                limits,
                "forged-page-one",
                Vec::<String>::new(),
                vec!["page/1#page.name.set".to_owned()],
                second.patch().durable().unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    let mut cross_owner = source.edit();
    let error = match cross_owner.join(forged_cross_sub_edit) {
        Ok(_) => panic!("concealed cross-sub-edit semantic overlap was accepted"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(!cross_owner.commit().unwrap().diagnostics().changed());
}

#[test]
fn joined_effect_keys_canonicalize_page_and_shape_selector_aliases() {
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let page_source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let mut first_page_edit = page_source.edit();
    first_page_edit.set_page_name(0, "Canonical page").unwrap();
    let first_page = first_page_edit.commit().unwrap();
    let mut alias_page_edit = page_source.edit();
    alias_page_edit.set_page_name(0, "Aliased page").unwrap();
    let mut alias_page_wire: serde_json::Value = serde_json::from_slice(
        &alias_page_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    alias_page_wire["operations"][0]["forward"]["target"] =
        serde_json::Value::String("page/00".to_owned());
    alias_page_wire["operations"][0]["inverse"]["target"] =
        serde_json::Value::String("page/00".to_owned());
    let alias_page =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&alias_page_wire).unwrap())
            .unwrap();
    let mut page_joined = page_source.joined_edits(limits);
    let page_lineage = page_joined.lineage().clone();
    page_joined
        .join(
            edit::SubEdit::new(
                page_lineage.clone(),
                limits,
                "canonical-page",
                Vec::<String>::new(),
                vec!["page/0#page.name.set".to_owned()],
                first_page.patch().durable().unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    page_joined
        .join(
            edit::SubEdit::new(
                page_lineage,
                limits,
                "aliased-page",
                Vec::<String>::new(),
                vec!["page/00#page.name.set".to_owned()],
                alias_page,
            )
            .unwrap(),
        )
        .unwrap();
    let mut page_owner = page_source.edit();
    let error = match page_owner.join(page_joined) {
        Ok(_) => panic!("page selector aliases bypassed joined overlap detection"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    assert!(!page_owner.commit().unwrap().diagnostics().changed());

    let shape_source = Package::from_bytes(&package(CONTENT_SHAPE_ALIAS, None)).unwrap();
    let mut alias_shape_edit = shape_source.edit();
    alias_shape_edit
        .set_shape_name(0, 0, "Aliased shape")
        .unwrap();
    let mut alias_shape_wire: serde_json::Value = serde_json::from_slice(
        &alias_shape_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    alias_shape_wire["operations"][0]["forward"]["target"] =
        serde_json::Value::String("page/00/shape/00".to_owned());
    alias_shape_wire["operations"][0]["inverse"]["target"] =
        serde_json::Value::String("page/00/shape/00".to_owned());
    let alias_shape = edit::DurablePatch::from_deterministic_json(
        &serde_json::to_vec(&alias_shape_wire).unwrap(),
    )
    .unwrap();
    let mut shape_joined = shape_source.joined_edits(limits);
    let shape_lineage = shape_joined.lineage().clone();
    shape_joined
        .join(
            edit::SubEdit::new(
                shape_lineage,
                limits,
                "aliased-shape",
                Vec::<String>::new(),
                vec!["page/00/shape/00#shape.name.set".to_owned()],
                alias_shape,
            )
            .unwrap(),
        )
        .unwrap();
    let mut ordinary = shape_source.edit();
    ordinary.set_shape_name(0, 0, "Ordinary shape").unwrap();
    let error = match ordinary.join(shape_joined) {
        Ok(_) => panic!("shape selector aliases bypassed staged overlap detection"),
        Err(error) => error,
    };
    assert!(matches!(
        error.error(),
        Error::Refused(refusal)
            if refusal.kind() == litchi_core::edit::RefusalKind::UnsupportedOperation
    ));
    let ordinary_commit = ordinary.commit().unwrap();
    assert_eq!(
        ordinary_commit.package().pages()[0].shapes()[0].name(),
        Some("Ordinary shape")
    );
}

#[test]
fn ordered_replacements_use_the_current_security_policy_state() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let signed_bytes = signed_package(CONTENT_AUTOMATIC);
    let signed = Package::from_bytes(&signed_bytes).unwrap();
    assert!(signed.security().is_signed());
    let mut plain_edit = source.edit();
    plain_edit
        .add_page(litchi_odg::page::Page::new("Plain target"))
        .unwrap();
    let plain_target = plain_edit.commit().unwrap().into_package();

    let source_id = BlobId::of(source.as_bytes()).as_hex();
    let signed_id = BlobId::of(&signed_bytes).as_hex();
    let plain_id = BlobId::of(plain_target.as_bytes()).as_hex();
    let source_fingerprint = DiagnosticFingerprint::of(source.as_bytes()).as_hex();
    let plain_fingerprint = DiagnosticFingerprint::of(plain_target.as_bytes()).as_hex();
    let wire = serde_json::json!({
        "format": "litchi.odg",
        "forward_blobs": [
            {"bytes": base64_encode(&signed_bytes), "sha256": signed_id},
            {"bytes": base64_encode(plain_target.as_bytes()), "sha256": plain_id}
        ],
        "operations": [
            {
                "forward": {
                    "op": "package.replace",
                    "target": "package",
                    "preconditions": {"source": source_fingerprint},
                    "value": signed_id
                },
                "inverse": {
                    "op": "package.replace",
                    "target": "package",
                    "preconditions": {"source": plain_fingerprint},
                    "value": source_id
                }
            },
            {
                "forward": {
                    "op": "package.replace",
                    "target": "package",
                    "preconditions": {"source": source_fingerprint},
                    "value": plain_id
                },
                "inverse": {
                    "op": "package.replace",
                    "target": "package",
                    "preconditions": {"source": plain_fingerprint},
                    "value": signed_id
                }
            }
        ],
        "reverse_blobs": [
            {"bytes": base64_encode(source.as_bytes()), "sha256": source_id},
            {"bytes": base64_encode(&signed_bytes), "sha256": signed_id}
        ],
        "version": 1
    });
    let durable =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&wire).unwrap()).unwrap();
    assert!(durable.apply(&source).is_err());
    assert_eq!(
        durable
            .apply_with_security_policy(&source, edit::SecurityWritePolicy::RemoveSignatures)
            .unwrap()
            .as_bytes(),
        plain_target.as_bytes()
    );

    let limits = CompositionLimits::new(2, 4, 4, 4);
    let mut refused_joined = source.joined_edits(limits);
    let refused_lineage = refused_joined.lineage().clone();
    refused_joined
        .join(
            edit::SubEdit::new(
                refused_lineage,
                limits,
                "two-replacements",
                Vec::<String>::new(),
                vec!["package".to_owned()],
                durable.clone(),
            )
            .unwrap(),
        )
        .unwrap();
    let mut refused_owner = source.edit();
    refused_owner.join(refused_joined).unwrap();
    assert!(refused_owner.commit().is_err());

    let mut permitted_joined = source.joined_edits(limits);
    let permitted_lineage = permitted_joined.lineage().clone();
    permitted_joined
        .join(
            edit::SubEdit::new(
                permitted_lineage,
                limits,
                "two-replacements-permitted",
                Vec::<String>::new(),
                vec!["package".to_owned()],
                durable,
            )
            .unwrap(),
        )
        .unwrap();
    let mut permitted_owner =
        source.edit_with_security_policy(edit::SecurityWritePolicy::RemoveSignatures);
    permitted_owner.join(permitted_joined).unwrap();
    assert_eq!(
        permitted_owner.commit().unwrap().package().as_bytes(),
        plain_target.as_bytes()
    );
}

#[test]
fn one_prepared_style_and_transition_sequence_replays_exactly() {
    let content = CONTENT_AUTOMATIC.replace(
        "draw:name=\"Page 2\" draw:style-name=\"dp2\"",
        "draw:name=\"Page 2\" draw:style-name=\"dp1\"",
    );
    let source = Package::from_bytes(&package(&content, None)).unwrap();
    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut multi_edit = source.edit();
    multi_edit.set_page_style_name(0, "dp2").unwrap();
    multi_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let multi_commit = multi_edit.commit().unwrap();
    assert_eq!(multi_commit.patch().changes().len(), 2);

    let mut joined = source.joined_edits(limits);
    joined
        .join(multi_commit.patch().prepare("ordered", limits).unwrap())
        .unwrap();
    let mut owner = source.edit();
    owner.join(joined).unwrap();
    let committed = owner.commit().unwrap();
    assert_eq!(committed.package().pages()[0].style_name(), Some("dp2"));
    assert_eq!(
        committed.package().pages()[0].transition().unwrap().style(),
        Some("fade-from-right")
    );
    assert_eq!(
        source
            .apply(committed.patch())
            .unwrap()
            .package()
            .as_bytes(),
        committed.package().as_bytes()
    );
    assert_eq!(
        committed
            .package()
            .apply(&committed.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );
    let durable = committed.patch().durable().unwrap();
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        committed.package().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(committed.package())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn one_ordered_mixed_replacement_payload_keeps_standalone_capability() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let mut structural_edit = source.edit();
    structural_edit
        .add_page(litchi_odg::page::Page::new("Added"))
        .unwrap();
    let structural = structural_edit.commit().unwrap();
    let structural_wire: serde_json::Value = serde_json::from_slice(
        &structural
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();

    let mut semantic_edit = source.edit();
    semantic_edit
        .set_page_transition(0, Some(transition("manual", "fade-from-right")))
        .unwrap();
    let semantic_wire: serde_json::Value = serde_json::from_slice(
        &semantic_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let mut mixed_wire = structural_wire;
    mixed_wire["operations"]
        .as_array_mut()
        .unwrap()
        .push(semantic_wire["operations"][0].clone());
    let mixed =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&mixed_wire).unwrap())
            .unwrap();
    let expected = mixed.apply(&source).unwrap();

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut joined = source.joined_edits(limits);
    let sub_edit = edit::SubEdit::new(
        joined.lineage().clone(),
        limits,
        "mixed-ordered",
        Vec::<String>::new(),
        vec![
            "package".to_owned(),
            "page/0#page.transition.set".to_owned(),
        ],
        mixed.clone(),
    )
    .unwrap();
    joined.join(sub_edit).unwrap();
    let mut owner = source.edit();
    owner.join(joined).unwrap();
    let committed = owner.commit().unwrap();
    assert_eq!(committed.package().as_bytes(), expected.as_bytes());
    assert_eq!(
        committed
            .package()
            .apply(&committed.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn ordered_replacement_can_rename_a_page_added_by_the_same_payload() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let mut structural_edit = source.edit();
    structural_edit
        .add_page(litchi_odg::page::Page::new("Added"))
        .unwrap();
    let structural = structural_edit.commit().unwrap();
    let structural_wire: serde_json::Value = serde_json::from_slice(
        &structural
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();

    let mut name_edit = source.edit();
    name_edit.set_page_name(0, "Unused").unwrap();
    let name_wire: serde_json::Value = serde_json::from_slice(
        &name_edit
            .commit()
            .unwrap()
            .patch()
            .durable()
            .unwrap()
            .to_deterministic_json()
            .unwrap(),
    )
    .unwrap();
    let name_operation = name_wire["operations"][0].clone();

    let mut mixed_wire = structural_wire;
    mixed_wire["operations"]
        .as_array_mut()
        .unwrap()
        .push(name_operation);
    mixed_wire["operations"][1]["forward"]["target"] =
        serde_json::Value::String("page/2".to_owned());
    mixed_wire["operations"][1]["forward"]["value"] =
        serde_json::Value::String("Added renamed".to_owned());
    mixed_wire["operations"][1]["inverse"]["target"] =
        serde_json::Value::String("page/2".to_owned());
    mixed_wire["operations"][1]["inverse"]["value"] = serde_json::Value::String("Added".to_owned());
    mixed_wire["operations"][1]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String("0".repeat(64));

    let provisional =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&mixed_wire).unwrap())
            .unwrap();
    let expected = provisional.apply(&source).unwrap();
    let final_source = DiagnosticFingerprint::of(expected.as_bytes()).as_hex();
    mixed_wire["operations"][0]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String(final_source.clone());
    mixed_wire["operations"][1]["inverse"]["preconditions"]["source"] =
        serde_json::Value::String(final_source);
    let mixed =
        edit::DurablePatch::from_deterministic_json(&serde_json::to_vec(&mixed_wire).unwrap())
            .unwrap();
    assert_eq!(
        mixed.apply(&source).unwrap().as_bytes(),
        expected.as_bytes()
    );
    assert_eq!(
        mixed.inverse().apply(&expected).unwrap().as_bytes(),
        source.as_bytes()
    );

    let limits = CompositionLimits::new(4, 8, 16, 8);
    let mut joined = source.joined_edits(limits);
    let sub_edit = edit::SubEdit::new(
        joined.lineage().clone(),
        limits,
        "replacement-then-rename",
        Vec::<String>::new(),
        vec!["package".to_owned(), "page/2#page.name.set".to_owned()],
        mixed,
    )
    .unwrap();
    joined.join(sub_edit).unwrap();
    let mut owner = source.edit();
    owner.join(joined).unwrap();
    let committed = owner.commit().unwrap();
    assert_eq!(committed.package().as_bytes(), expected.as_bytes());
    assert_eq!(committed.package().pages()[2].name(), Some("Added renamed"));
    assert_eq!(
        source
            .apply(committed.patch())
            .unwrap()
            .package()
            .as_bytes(),
        expected.as_bytes()
    );
    assert_eq!(
        committed
            .package()
            .apply(&committed.patch().inverse())
            .unwrap()
            .package()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn transition_sub_edit_admission_honors_each_composition_bound() {
    let source = Package::from_bytes(&package(CONTENT_AUTOMATIC, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();

    let per_edit = CompositionLimits::new(2, 0, 4, 2);
    assert!(matches!(
        commit.patch().prepare("too-many-effects", per_edit),
        Err(Error::InvalidFormat(message)) if message.contains("EffectsPerSubEdit")
    ));

    let one_sub_edit = CompositionLimits::new(1, 2, 4, 2);
    let mut limited_sub_edits = source.joined_edits(one_sub_edit);
    limited_sub_edits
        .join(commit.patch().prepare("first", one_sub_edit).unwrap())
        .unwrap();
    let error = limited_sub_edits
        .join(commit.patch().prepare("second", one_sub_edit).unwrap())
        .unwrap_err();
    assert!(matches!(
        error.failure(),
        SubEditJoinFailure::Limit(CompositionError::Limit {
            kind: CompositionLimitKind::SubEdits,
            ..
        })
    ));

    let one_total_effect = CompositionLimits::new(2, 2, 1, 2);
    let mut limited_total = source.joined_edits(one_total_effect);
    limited_total
        .join(commit.patch().prepare("first", one_total_effect).unwrap())
        .unwrap();
    let error = limited_total
        .join(commit.patch().prepare("second", one_total_effect).unwrap())
        .unwrap_err();
    assert!(matches!(
        error.failure(),
        SubEditJoinFailure::Limit(CompositionError::Limit {
            kind: CompositionLimitKind::TotalEffects,
            ..
        })
    ));
}
