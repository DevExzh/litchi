#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::{CompositionLimits, HistoryLimits, MergeChoice, SubEditJoinFailure};
use litchi_odf_common::{
    compact_xml,
    core::{OwnedPackage, PackageWriter},
};
use litchi_odg::{Drawing, PackageDurablePatch, PackageMergePlan, page::Page, shape::ShapeKind};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:xml="http://www.w3.org/XML/1998/namespace" office:version="1.3"><office:body><office:drawing><draw:page draw:name="Page 1" draw:style-name="dp1" draw:master-page-name="Default" xml:id="page1"><draw:layer-set><draw:layer draw:name="Foreground" draw:display="always" draw:protected="false"/><draw:layer draw:name="Background"/></draw:layer-set><draw:rect draw:name="Label" draw:layer="Foreground" draw:style-name="gr1" draw:text-style-name="P1" draw:z-index="7" svg:x="1cm" svg:y="2cm" svg:width="3cm" svg:height="4cm"><svg:title>Label title</svg:title><svg:desc>Label description</svg:desc><text:p>Old label</text:p></draw:rect><draw:frame draw:name="Photo" draw:layer="Background" svg:width="2cm" svg:height="1cm"><svg:title>Photo title</svg:title><svg:desc>Photo description</svg:desc></draw:frame></draw:page></office:drawing></office:body></office:document-content>"#;

fn package(content: &str) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.graphics")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn raw_negative_fixture_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.graphics";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.graphics"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

#[test]
fn semantic_package_views_use_layers_shapes_and_shared_frame_context() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    assert_eq!(
        drawing.page(0usize).unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert_eq!(
        drawing.page("Page 1").unwrap().unwrap().name(),
        Some("Page 1")
    );
    assert!(drawing.page(9usize).unwrap().is_none());
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.pages()[0].name(), Some("Page 1"));
    assert_eq!(drawing.pages()[0].xml_id(), Some("page1"));
    assert_eq!(drawing.pages()[0].style_name(), Some("dp1"));
    assert_eq!(drawing.pages()[0].master_page_name(), Some("Default"));
    assert!(drawing.layers().is_empty());
    assert_eq!(drawing.pages()[0].layers().len(), 2);
    assert_eq!(drawing.pages()[0].layers()[0].name(), "Foreground");
    assert_eq!(drawing.pages()[0].layers()[0].display(), Some("always"));
    assert_eq!(drawing.pages()[0].layers()[0].protected(), Some(false));
    let shape = &drawing.pages()[0].shapes()[0];
    assert_eq!(drawing.pages()[0].shape("Label").unwrap(), Some(shape));
    assert_eq!(drawing.pages()[0].shape(0usize).unwrap(), Some(shape));
    assert!(drawing.pages()[0].shape("Missing").unwrap().is_none());
    assert_eq!(shape.name(), Some("Label"));
    assert_eq!(shape.layer(), Some("Foreground"));
    assert_eq!(shape.kind(), ShapeKind::Rectangle);
    assert_eq!(shape.style_name(), Some("gr1"));
    assert_eq!(shape.text_style_name(), Some("P1"));
    assert_eq!(shape.z_index(), Some(7));
    assert_eq!(shape.x(), Some("1cm"));
    assert_eq!(shape.y(), Some("2cm"));
    assert_eq!(shape.width(), Some("3cm"));
    assert_eq!(shape.height(), Some("4cm"));
    assert_eq!(shape.title(), Some("Label title"));
    assert_eq!(shape.description(), Some("Label description"));
    assert_eq!(shape.text(), "Old label");
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .width
            .as_deref(),
        Some("2cm")
    );
    assert_eq!(
        drawing.pages()[0].shapes()[1]
            .frame()
            .unwrap()
            .title
            .as_deref(),
        Some("Photo title")
    );
}

#[test]
fn exact_shape_name_selector_rejects_ambiguity() {
    let duplicate = CONTENT.replace("draw:name=\"Photo\"", "draw:name=\"Label\"");
    let drawing = Drawing::from_bytes(package(&duplicate)).unwrap();
    assert!(drawing.pages()[0].shape("Label").is_err());
}

#[test]
fn package_edit_is_source_checked_reversible_and_compact() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_text(0, 0, "New <label>").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].text(),
        "New <label>"
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("<text:p>New &lt;label&gt;</text:p>")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    let output = OwnedPackage::from_bytes(commit.snapshot().as_bytes().to_vec()).unwrap();
    let archive = output.package().unwrap();
    for path in ["content.xml", "META-INF/manifest.xml"] {
        compact_xml::validate(&archive.get_file(path).unwrap()).unwrap();
    }
    assert!(!commit.snapshot().content_xml().contains(">\r\n<"));
    assert!(commit.patch().is_applicable_to(drawing.snapshot()));
    let reapplied = commit.patch().apply(drawing.snapshot()).unwrap();
    assert_eq!(reapplied.as_bytes(), commit.snapshot().as_bytes());
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());
}

#[test]
fn package_shape_name_edit_is_source_checked_reversible_and_compact() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_name(0, 0, "Renamed & exact").unwrap();
    let commit = transaction.commit().unwrap();

    let shape = &commit.snapshot().pages()[0].shapes()[0];
    assert_eq!(shape.name(), Some("Renamed & exact"));
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:name=\"Renamed &amp; exact\"")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:layer=\"Foreground\"")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:name=\"Photo\"")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    assert!(!commit.snapshot().content_xml().contains(">\r\n<"));

    let change = commit.patch().name_change().unwrap();
    assert_eq!(change.before(), "Label");
    assert_eq!(change.after(), "Renamed & exact");
    assert!(commit.patch().is_applicable_to(drawing.snapshot()));
    let different =
        Drawing::from_bytes(package(&CONTENT.replace("Photo", "Different photo"))).unwrap();
    assert!(!commit.patch().is_applicable_to(different.snapshot()));
    assert!(commit.patch().apply(different.snapshot()).is_err());
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());
}

#[test]
fn package_layer_edit_is_declared_source_checked_and_reversible() {
    let drawing = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_layer(0, 0, "Background").unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].layer(),
        Some("Background")
    );
    assert!(
        commit
            .snapshot()
            .content_xml()
            .contains("draw:layer=\"Background\"")
    );
    assert!(!commit.snapshot().content_xml().contains(">\n<"));
    let change = commit.patch().layer_change().unwrap();
    assert_eq!(change.before(), "Foreground");
    assert_eq!(change.after(), "Background");
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), drawing.as_bytes());

    let mut invalid = drawing.edit();
    assert!(invalid.set_shape_layer(0, 0, "Missing").is_err());
}

#[test]
fn ended_page_scope_does_not_capture_shapes_outside_a_page() {
    let malformed = CONTENT.replace(
        "</draw:page></office:drawing>",
        "</draw:page><draw:g><draw:rect draw:name=\"outside\"/></draw:g></office:drawing>",
    );
    assert!(Drawing::from_bytes(package(&malformed)).is_err());
}

#[test]
fn dtd_and_noncompact_rewrite_are_refused_without_execution() {
    let dtd = CONTENT.replacen("<office:body>", "<!DOCTYPE drawing><office:body>", 1);
    assert!(Drawing::from_bytes(raw_negative_fixture_package(&dtd)).is_err());

    let noncompact = CONTENT.replacen("<office:body>", "\n<office:body>", 1);
    let drawing = Drawing::from_bytes(raw_negative_fixture_package(&noncompact)).unwrap();
    let mut transaction = drawing.edit();
    transaction.set_shape_text(0, 0, "New label").unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn durable_semantic_patch_roundtrips_inverts_and_refuses_stale_sources() {
    let source = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.set_shape_text(0, 0, "durable text").unwrap();
    edit.set_shape_geometry(0, 0, "2cm", "3cm", "5cm", "6cm")
        .unwrap();
    edit.set_shape_style_name(0, 0, "gr2").unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    let wire = durable.to_deterministic_json().unwrap();
    assert_eq!(wire, durable.to_deterministic_json().unwrap());
    let reopened = PackageDurablePatch::from_deterministic_json(&wire).unwrap();
    assert_eq!(reopened.operations().len(), 3);
    assert_eq!(
        reopened.apply(source.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        reopened
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let stale = Drawing::from_bytes(package(&CONTENT.replace("Photo", "Changed"))).unwrap();
    assert!(reopened.apply(stale.snapshot()).is_err());
}

#[test]
fn durable_structural_patch_reopens_full_package_and_inverts_exactly() {
    let source = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut edit = source.edit();
    edit.add_page(Page::new("Added")).unwrap();
    let commit = edit.commit().unwrap();
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "package.replace");
    let reopened =
        PackageDurablePatch::from_deterministic_json(&durable.to_deterministic_json().unwrap())
            .unwrap();
    let applied = reopened.apply(source.snapshot()).unwrap();
    assert_eq!(applied.as_bytes(), commit.snapshot().as_bytes());
    assert_eq!(applied.pages().len(), 2);
    assert_eq!(
        reopened.inverse().apply(&applied).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn path_data_is_typed_losslessly_editable_and_durable() {
    let content = CONTENT.replace(
        "</draw:page>",
        r#"<draw:path draw:name="Route" svg:d="M 0 0 L 10 10"/></draw:page>"#,
    );
    let source = Drawing::from_bytes(package(&content)).unwrap();
    assert_eq!(
        source.pages()[0].shapes()[2].path_data(),
        Some("M 0 0 L 10 10")
    );
    let mut edit = source.edit();
    edit.set_shape_path_data(0, 2, "M 1 2 C 3 4 5 6 7 8")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[2].path_data(),
        Some("M 1 2 C 3 4 5 6 7 8")
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "shape.path.set");
    assert_eq!(
        durable.apply(source.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
    assert_eq!(
        durable
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn form_controls_remain_inert_but_their_references_are_editable() {
    let content = CONTENT.replace(
        "</draw:page>",
        r#"<draw:control draw:name="Button" draw:control="button1"/></draw:page>"#,
    );
    let source = Drawing::from_bytes(package(&content)).unwrap();
    let control = &source.pages()[0].shapes()[2];
    assert_eq!(control.kind(), ShapeKind::Control);
    assert_eq!(control.control_reference(), Some("button1"));
    let mut edit = source.edit();
    edit.set_shape_control_reference(0, 2, "button2").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[2].control_reference(),
        Some("button2")
    );
    let durable = commit.patch().durable().unwrap();
    assert_eq!(durable.operations()[0].op, "shape.control.set");
    assert_eq!(
        durable.apply(source.snapshot()).unwrap().as_bytes(),
        commit.snapshot().as_bytes()
    );
}

#[test]
fn joined_edits_are_deterministic_disjoint_and_conflict_aware() {
    let source = Drawing::from_bytes(package(CONTENT)).unwrap();
    let limits = CompositionLimits::new(8, 8, 32, 16);

    let mut text_edit = source.edit();
    text_edit.set_shape_text(0, 0, "joined").unwrap();
    let text_commit = text_edit.commit().unwrap();
    let mut name_edit = source.edit();
    name_edit.set_shape_name(0, 1, "Joined photo").unwrap();
    let name_commit = name_edit.commit().unwrap();

    let mut joined = source.snapshot().joined_edits(limits);
    joined
        .join(text_commit.patch().prepare("z-text", limits).unwrap())
        .unwrap();
    joined
        .join(name_commit.patch().prepare("a-name", limits).unwrap())
        .unwrap();
    assert_eq!(
        joined
            .sub_edits()
            .map(litchi_odg::PackagePreparedEdit::id)
            .collect::<Vec<_>>(),
        ["a-name", "z-text"]
    );
    let applied = source.snapshot().apply_joined(joined).unwrap();
    assert_eq!(applied.pages()[0].shapes()[0].text(), "joined");
    assert_eq!(applied.pages()[0].shapes()[1].name(), Some("Joined photo"));

    let mut conflicting_edit = source.edit();
    conflicting_edit.set_shape_text(0, 0, "other").unwrap();
    let conflicting_commit = conflicting_edit.commit().unwrap();
    let mut conflicts = source.snapshot().joined_edits(limits);
    conflicts
        .join(text_commit.patch().prepare("left", limits).unwrap())
        .unwrap();
    let error = conflicts
        .join(conflicting_commit.patch().prepare("right", limits).unwrap())
        .unwrap_err();
    assert!(matches!(error.failure(), SubEditJoinFailure::Overlap(_)));
}

#[test]
fn three_way_plan_does_not_mutate_until_conflicts_are_resolved() {
    let source = Drawing::from_bytes(package(CONTENT)).unwrap();
    let original = source.as_bytes().to_vec();
    let limits = CompositionLimits::new(8, 8, 32, 16);
    let mut left_edit = source.edit();
    left_edit.set_shape_name(0, 0, "Left").unwrap();
    let left_commit = left_edit.commit().unwrap();
    let mut right_edit = source.edit();
    right_edit.set_shape_name(0, 0, "Right").unwrap();
    let right_commit = right_edit.commit().unwrap();
    let mut left = source.snapshot().joined_edits(limits);
    left.join(left_commit.patch().prepare("left", limits).unwrap())
        .unwrap();
    let mut right = source.snapshot().joined_edits(limits);
    right
        .join(right_commit.patch().prepare("right", limits).unwrap())
        .unwrap();

    let plan = PackageMergePlan::new(left, right).unwrap();
    assert_eq!(source.as_bytes(), original);
    assert_eq!(plan.conflicts().len(), 1);
    let mut unresolved = plan.finish().unwrap_err();
    unresolved.resolve(MergeChoice::Right);
    let staged = unresolved.finish().unwrap();
    let applied = source.snapshot().apply_joined(staged).unwrap();
    assert_eq!(applied.pages()[0].shapes()[0].name(), Some("Right"));
}

#[test]
fn history_enforces_step_and_weight_bounds_without_mutating_on_refusal() {
    let source = Drawing::from_bytes(package(CONTENT)).unwrap();
    let mut first_edit = source.edit();
    first_edit.set_shape_name(0, 0, "First").unwrap();
    let first = first_edit.commit().unwrap().into_snapshot();
    let mut second_edit = first.edit();
    second_edit.set_shape_name(0, 0, "Second").unwrap();
    let second = second_edit.commit().unwrap().into_snapshot();
    let mut history = source.snapshot().history(HistoryLimits::new(2, 10));
    history.record(first.clone(), 6).unwrap();
    history.record(second.clone(), 6).unwrap();
    assert_eq!(history.retained_weight(), 6);
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), first.as_bytes());
    assert!(!history.undo());
    assert!(history.record(second, 11).is_err());
    assert_eq!(history.current().as_bytes(), first.as_bytes());
}
