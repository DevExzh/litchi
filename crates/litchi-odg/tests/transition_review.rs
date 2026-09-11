#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{CompositionError, CompositionLimitKind, CompositionLimits, SubEditJoinFailure};
use litchi_odf_common::core::PackageWriter;
use litchi_odg::{Drawing, Transition};

const CONTENT_AUTOMATIC: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual" presentation:transition-style="fade-from-left"/></style:style><style:style style:name="dp2" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" presentation:transition-style="dissolve"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/><draw:page draw:name="Page 2" draw:style-name="dp2"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_SHARED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/><draw:page draw:name="Page 2" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_NAMED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const STYLES_NAMED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:styles></office:document-styles>"#;

const CONTENT_INHERITED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="child"/></office:drawing></office:body></office:document-content>"#;

const STYLES_INHERITED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.4"><office:styles><style:style style:name="parent" style:family="drawing-page"><style:drawing-page-properties presentation:transition-style="fade-from-left"/></style:style><style:style style:name="child" style:family="drawing-page" style:parent-style-name="parent"><style:drawing-page-properties presentation:transition-type="manual"/></style:style></office:styles></office:document-styles>"#;

const CONTENT_FOREIGN_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:example:foreign-presentation" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:vendor="keep"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

const CONTENT_LOCAL_FOREIGN_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:litchi_presentation="urn:example:occupied" office:version="1.4"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties xmlns:presentation="urn:example:foreign-presentation" presentation:vendor="keep"/></style:style></office:automatic-styles><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1"/></office:drawing></office:body></office:document-content>"#;

fn package(content: &str, styles: Option<&str>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
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

fn assert_exact_durable_round_trip(source: Drawing, value: Transition) {
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(value)).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let replayed = durable.apply(source.snapshot()).unwrap();
    assert_eq!(replayed.as_bytes(), commit.snapshot().as_bytes());
    let restored = durable.inverse().apply(&replayed).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn equal_transition_is_a_noop_even_when_the_style_is_shared() {
    let source = Drawing::from_bytes(package(CONTENT_SHARED, None)).unwrap();
    let current = source.pages()[0].transition().cloned();
    let mut edit = source.edit();
    edit.set_page_transition(0, current).unwrap();
    let commit = edit.commit().unwrap();

    assert!(!commit.changed());
    assert!(commit.patch().changes().is_empty());
    assert_eq!(commit.snapshot().as_bytes(), source.as_bytes());
    assert_eq!(commit.patch().durable().unwrap().operations().len(), 0);
}

#[test]
fn inherited_transition_can_be_replaced_but_cannot_be_cleared_locally() {
    let source = Drawing::from_bytes(package(CONTENT_INHERITED, Some(STYLES_INHERITED))).unwrap();
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
    let output = replacement.commit().unwrap().into_snapshot();
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
    let source = Drawing::from_bytes(package(CONTENT_NAMED, Some(STYLES_NAMED))).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "page.transition.set");

    let replayed = durable.apply(source.snapshot()).unwrap();
    assert_eq!(replayed.as_bytes(), commit.snapshot().as_bytes());
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
        Drawing::from_bytes(package(&automatic_content, None)).unwrap(),
        transition_with_smil(),
    );
    assert_exact_durable_round_trip(
        Drawing::from_bytes(package(CONTENT_NAMED, Some(STYLES_NAMED))).unwrap(),
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
        Drawing::from_bytes(package(&automatic_content, None)).unwrap(),
        transition("automatic", "dissolve"),
    );

    let named_styles = STYLES_NAMED.replace(
        r#"<style:drawing-page-properties presentation:transition-type="manual"/></style:style>"#,
        "</style:style>",
    );
    assert_exact_durable_round_trip(
        Drawing::from_bytes(package(CONTENT_NAMED, Some(&named_styles))).unwrap(),
        transition("automatic", "dissolve"),
    );
}

#[test]
fn transition_edit_does_not_rebind_a_foreign_presentation_prefix() {
    let source = Drawing::from_bytes(package(CONTENT_FOREIGN_PRESENTATION, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("presentation:vendor=\"keep\"")
    );
    assert!(commit.snapshot().content_xml().contains(
        "xmlns:litchi_presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\""
    ));
    let durable = commit.patch().durable().unwrap();
    let restored = durable
        .inverse()
        .apply(&durable.apply(source.snapshot()).unwrap())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn transition_edit_avoids_foreign_alias_and_local_shadowing() {
    let source = Drawing::from_bytes(package(CONTENT_LOCAL_FOREIGN_PRESENTATION, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();
    let content = commit.snapshot().content_xml();
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
        .apply(&durable.apply(source.snapshot()).unwrap())
        .unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn disjoint_transition_sub_edits_join_and_same_page_edits_conflict() {
    let source = Drawing::from_bytes(package(CONTENT_AUTOMATIC, None)).unwrap();
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
    let output = source.snapshot().apply_joined(joined).unwrap();
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
fn transition_sub_edit_admission_honors_each_composition_bound() {
    let source = Drawing::from_bytes(package(CONTENT_AUTOMATIC, None)).unwrap();
    let mut edit = source.edit();
    edit.set_page_transition(0, Some(transition("automatic", "dissolve")))
        .unwrap();
    let commit = edit.commit().unwrap();

    let per_edit = CompositionLimits::new(2, 0, 4, 2);
    assert!(matches!(
        commit.patch().prepare("too-many-effects", per_edit),
        Err(litchi_core::Error::InvalidFormat(message)) if message.contains("EffectsPerSubEdit")
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
