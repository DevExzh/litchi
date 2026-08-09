#![allow(
    clippy::unwrap_used,
    reason = "integration tests panic on unexpected errors to keep assertions concise"
)]

use litchi_odc::{AxisSpec, AxisUpdate, Chart, ChartClass, Definition, chart::Dimension};
use litchi_odf_common::core::PackageWriter;
use soapberry_zip::office::StreamingArchiveWriter;

fn source_definition() -> Definition {
    let mut definition = Definition::new(ChartClass::bar());
    let mut horizontal = AxisSpec::new(Dimension::X);
    horizontal.name = Some("primary-x".to_string());
    definition.plot_area.axes.push(horizontal);
    let mut vertical = AxisSpec::new(Dimension::Y);
    vertical.name = Some("primary-y".to_string());
    definition.plot_area.axes.push(vertical);
    definition
}

fn package(content: &str, auxiliary: Option<(&str, &[u8])>) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.chart")
        .unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    if let Some((path, bytes)) = auxiliary {
        writer.add_file(path, bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn raw_negative_fixture_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.chart";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
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
fn package_axis_transaction_is_atomic_exact_source_checked_and_reversible() {
    let source = Chart::from_definition(source_definition()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert!(!no_op.changed());
    assert_eq!(no_op.chart().as_bytes(), source.as_bytes());

    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("revenue&amp<2026>"))
        .unwrap();
    edit.update_axis(1, AxisUpdate::unnamed()).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.patch().changes().len(), 2);
    assert!(commit.patch().is_applicable_to(&source));
    assert!(
        commit
            .chart()
            .content_xml()
            .contains("chart:name=\"revenue&amp;amp&lt;2026&gt;\"")
    );
    assert!(!commit.chart().content_xml().contains("primary-y"));
    assert!(!commit.chart().content_xml().contains("><\n<"));
    assert!(!commit.chart().content_xml().contains("><\r\n<"));

    let stale = Chart::from_definition(Definition::new(ChartClass::line())).unwrap();
    assert!(!commit.patch().is_applicable_to(&stale));
    assert!(commit.patch().apply(&stale).is_err());

    let restored = commit.patch().inverse().apply(commit.chart()).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn package_edit_preserves_auxiliary_payloads_and_refuses_signed_or_pretty_xml() {
    let content = litchi_odc::serialize_content(&source_definition()).unwrap();
    let bytes = package(&content, Some(("Pictures/keep.bin", b"keep-exact")));
    let source = Chart::from_bytes(bytes).unwrap();
    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("changed")).unwrap();
    let committed = edit.commit().unwrap().into_chart();
    let archive =
        litchi_odf_common::core::OwnedPackage::from_bytes(committed.into_bytes()).unwrap();
    assert_eq!(
        archive.get_file("Pictures/keep.bin").unwrap(),
        b"keep-exact"
    );

    let signed = Chart::from_bytes(package(
        &content,
        Some(("META-INF/documentsignatures.xml", b"<signatures/>")),
    ))
    .unwrap();
    let mut signed_edit = signed.edit();
    signed_edit
        .update_axis(0, AxisUpdate::named("blocked"))
        .unwrap();
    assert!(signed_edit.commit().is_err());

    let pretty = content.replacen("><", ">\n<", 1);
    let pretty_chart = Chart::from_bytes(raw_negative_fixture_package(&pretty)).unwrap();
    let mut pretty_edit = pretty_chart.edit();
    pretty_edit
        .update_axis(0, AxisUpdate::named("blocked"))
        .unwrap();
    assert!(pretty_edit.commit().is_err());
}

#[test]
fn explicit_chart_replacement_publishes_full_typed_definition_and_is_reversible() {
    let original = source_definition();
    let source = Chart::from_definition(original.clone()).unwrap();
    let mut no_op_edit = source.edit();
    no_op_edit.replace_chart(&original).unwrap();
    let no_op = no_op_edit.commit().unwrap();
    assert!(!no_op.changed());
    assert_eq!(no_op.chart().as_bytes(), source.as_bytes());

    let mut replacement = Definition::new(ChartClass::ring());
    replacement.title = Some(litchi_odc::Text::new("Quarterly totals"));
    let mut category_axis = AxisSpec::new(Dimension::X);
    category_axis.name = Some("categories".to_string());
    replacement.plot_area.axes.push(category_axis);

    let mut edit = source.edit();
    edit.replace_chart(&replacement).unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(commit.patch().replaces_chart());
    assert!(commit.patch().changes().is_empty());
    assert_eq!(
        commit.chart().class().unwrap().kind(),
        replacement.class.kind()
    );
    assert!(
        commit
            .chart()
            .content_xml()
            .contains("<text:p>Quarterly totals</text:p>")
    );
    assert!(!commit.chart().content_xml().contains('\n'));

    let restored = commit.patch().inverse().apply(commit.chart()).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}
