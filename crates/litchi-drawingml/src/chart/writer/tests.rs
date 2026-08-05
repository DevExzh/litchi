//! Focused public-facade checks for chart XML emission.

use super::{write, write_with_rels};
use crate::chart::{Chart, ExternalData};

#[test]
fn writer_facade_emits_a_chart_space() {
    let mut xml = Vec::new();
    write(&mut xml, &Chart::new()).unwrap();
    let xml = std::str::from_utf8(&xml).unwrap();
    assert!(xml.starts_with("<?xml version=\"1.0\""));
    assert!(xml.contains("<c:chartSpace"));
    assert!(xml.ends_with("</c:chartSpace>"));
}

#[test]
fn relationship_overrides_are_forwarded_to_external_data() {
    let mut chart = Chart::new();
    chart.external_data = Some(ExternalData::pending());

    let mut xml = Vec::new();
    write_with_rels(&mut xml, &chart, Some("rId7"), None).unwrap();
    assert!(
        std::str::from_utf8(&xml)
            .unwrap()
            .contains(r#"<c:externalData r:id="rId7">"#)
    );
}
