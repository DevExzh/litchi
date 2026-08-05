use litchi_odc::{
    Builder, Chart,
    chart::{Dimension, Kind, Position},
};

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Dimension::Z, Dimension::Z);
    assert_eq!(Position::BottomEnd, Position::BottomEnd);

    let bytes = Builder::new().build().unwrap();
    let chart = Chart::from_bytes(bytes).unwrap();
    assert!(chart.content_xml().contains("<office:chart"));
    assert_eq!(chart.chart().kind(), Kind::Chart);
    assert!(chart.plot_area().is_some());
}

#[test]
fn validation_is_namespace_aware_and_structural() {
    let content = r#"<?xml version="1.0"?><o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:c="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><o:body><o:chart><c:chart><c:plot-area/></c:chart></o:chart></o:body></o:document-content>"#;
    let bytes = Builder::new().content_xml(content).build().unwrap();
    assert!(
        Chart::from_bytes(bytes)
            .unwrap()
            .chart()
            .plot_area()
            .is_some()
    );

    let invalid = Builder::new().content_xml(
        r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body></o:document-content>"#,
    );
    assert!(invalid.build().is_err());
}
