#![allow(
    clippy::unwrap_used,
    reason = "Tests use unwrap to keep successful chart-construction assertions concise."
)]

use super::anchor::{validate, write_all};
use super::relationship::{
    ExternalDataPart, ExternalDataTarget, is_external_data_type, is_user_shapes_type,
    user_shapes_ids,
};
use super::{Anchor, Chart, Series, read, write};

#[test]
fn anchor_validation_preserves_worksheet_bounds_and_offsets() {
    assert!(validate(&Anchor::new(0, 0, 16_383, 1_048_575)).is_ok());
    assert!(validate(&Anchor::new(0, 0, 16_384, 1)).is_err());
    assert!(validate(&Anchor::new(2, 0, 1, 1)).is_err());
    assert!(validate(&Anchor::with_offsets(0, -1, 0, 0, 1, 0, 1, 0)).is_err());
}

#[test]
fn anchor_writer_emits_cell_anchored_chart_frame() {
    let chart = Chart::bar_chart(
        "Sales",
        "Sheet1!$A$1:$A$2",
        "Sheet1!$B$1:$B$2",
        Anchor::new(1, 2, 7, 14),
    )
    .unwrap();
    let mut xml = String::new();
    write_all(&mut xml, &[chart], 4, 8).unwrap();
    assert!(xml.contains("<xdr:from><xdr:col>1</xdr:col>"));
    assert!(xml.contains("<xdr:to><xdr:col>7</xdr:col>"));
    assert!(xml.contains("id=\"5\""));
    assert!(xml.contains("r:id=\"rId9\""));
}

#[test]
fn chart_model_round_trips_through_shared_drawingml_codec() {
    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Sheet1!$A$2:$A$4",
        &["Jan", "Feb", "Mar"],
        "Sheet1!$B$2:$B$4",
        &[10.0, 20.0, 30.0],
        Anchor::default(),
    )
    .unwrap();
    let xml = write(&chart.chart).unwrap();
    let parsed = read(&xml, chart.anchor.clone()).unwrap();
    assert_eq!(parsed.series_count(), 1);
    assert_eq!(parsed.anchor.from_col, chart.anchor.from_col);
    assert_eq!(parsed.chart_type(), chart.chart_type());
}

#[test]
fn strict_and_transitional_chart_namespaces_are_accepted() {
    for namespace in [
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "http://purl.oclc.org/ooxml/drawingml/chart",
    ] {
        let xml = format!(
            r#"<c:chartSpace xmlns:c="{namespace}"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#
        );
        let chart = read(xml.as_bytes(), Anchor::default()).unwrap();
        assert_eq!(chart.series_count(), 0);
    }
}

#[test]
fn relationship_vocabulary_keeps_external_resources_inert() {
    let embedded = ExternalDataPart::embedded_workbook(vec![1, 2, 3]);
    assert!(matches!(
        embedded.target,
        ExternalDataTarget::Embedded { ref extension, .. } if extension == "xlsx"
    ));
    let linked = ExternalDataPart::linked_package("https://example.test/book.xlsx");
    assert!(matches!(
        linked.target,
        ExternalDataTarget::Linked { ref target } if target == "https://example.test/book.xlsx"
    ));
    assert!(is_external_data_type(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package"
    ));
    assert!(is_user_shapes_type(
        "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes"
    ));
    let _series = Series::new(0);
}

#[test]
fn strict_user_shapes_scan_retains_relationship_references() {
    let xml = br#"<c:userShapes xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><c:sp r:id="rIdShape"/></c:userShapes>"#;
    let ids = user_shapes_ids(xml).unwrap();
    assert!(ids.contains("rIdShape"));
}
