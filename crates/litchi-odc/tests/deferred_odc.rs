#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected errors"
)]

use litchi_odc::{AxisUpdate, FlatChart};

const ODC: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3">"#,
    r#"<office:styles><style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="keep-me"/></office:styles>"#,
    r#"<office:body><office:chart><chart:chart chart:class="chart:bar"><chart:title><text:p>Revenue</text:p></chart:title><chart:legend chart:legend-position="end"/><chart:plot-area>"#,
    r#"<chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"/>"#,
    r#"<chart:series chart:values-cell-range-address="Sheet1.$B$1:.$B$3"><chart:data-point chart:repeated="2"/></chart:series>"#,
    r#"<table:table table:name="local-table"><table:table-row><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell></table:table-row></table:table>"#,
    r#"</chart:plot-area></chart:chart></office:chart></office:body></office:document>"#,
);

fn exercise(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    match FlatChart::from_bytes(bytes) {
        Ok(flat) => {
            if let Some(plot) = flat.plot_area() {
                for axis in plot.axes() {
                    if let Err(error) = axis.dimension() {
                        return Some(error);
                    }
                }
                for series in plot.series() {
                    let _ = series.values_range();
                    let _ = series.attached_axis();
                }
            }
            None
        },
        Err(error) => Some(error),
    }
}

fn assert_invalid(case: &str, bytes: Vec<u8>) {
    assert!(
        matches!(exercise(bytes), Some(litchi_core::Error::InvalidFormat(_))),
        "{case}"
    );
}

#[test]
fn flat_chart_reads_tree_preserves_bytes_and_rejects_wrong_family() {
    let flat = FlatChart::from_bytes(ODC.as_bytes().to_vec()).unwrap();
    assert_eq!(flat.as_bytes(), ODC.as_bytes());
    assert!(flat.chart().all_text().contains("Revenue"));
    assert!(flat.find_axis("primary-x").is_some());
    assert_invalid(
        "wrong family",
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body><office:drawing/></office:body></office:document>"#
            .as_bytes()
            .to_vec(),
    );
    assert_invalid("packaged input", b"PK\x03\x04mimetype".to_vec());
}

#[test]
fn flat_axis_transaction_is_lossless_source_checked_and_reversible() {
    let source = FlatChart::from_bytes(ODC.as_bytes().to_vec()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert_eq!(no_op.snapshot().as_bytes(), source.as_bytes());
    assert!(no_op.patch().changes().is_empty());

    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("renamed-x")).unwrap();
    let commit = edit.commit().unwrap();
    let output = std::str::from_utf8(commit.snapshot().as_bytes()).unwrap();
    assert!(output.contains("style:name=\"keep-me\""));
    assert!(output.contains("<text:p>Revenue</text:p>"));
    assert!(commit.snapshot().find_axis("renamed-x").is_some());
    assert!(source.find_axis("primary-x").is_some());

    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(restored.snapshot().as_bytes(), source.as_bytes());
    assert!(commit.patch().apply(commit.snapshot()).is_err());
}

#[test]
fn malformed_and_misplaced_chart_inputs_return_typed_errors() {
    let content = |inner: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3"><office:body>{inner}</office:body></office:document>"#,
        )
        .into_bytes()
    };
    assert_invalid(
        "wrong root",
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.chart"><office:body/></office:document-content>"#
            .as_bytes()
            .to_vec(),
    );
    assert_invalid(
        "duplicate wrapper",
        content(
            r#"<office:chart><chart:chart chart:class="chart:bar"/></office:chart><office:chart><chart:chart chart:class="chart:line"/></office:chart>"#,
        ),
    );
    assert_invalid(
        "invalid dimension",
        content(
            r#"<office:chart><chart:chart chart:class="chart:bar"><chart:plot-area><chart:axis chart:dimension="invalid" chart:name="bad"/></chart:plot-area></chart:chart></office:chart>"#,
        ),
    );
    assert_invalid(
        "DOCTYPE",
        r#"<?xml version="1.0"?><!DOCTYPE office><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.chart"><office:body/></office:document>"#
            .as_bytes()
            .to_vec(),
    );
    let foreign = content(
        r#"<office:chart><chart:chart chart:class="chart:bar"><chart:plot-area/><ext:thing xmlns:ext="urn:example" ext:attr="1"/></chart:chart></office:chart>"#,
    );
    assert!(FlatChart::from_bytes(foreign).is_ok());
    let mut invalid_utf8 = ODC.as_bytes().to_vec();
    invalid_utf8.insert(invalid_utf8.len() / 2, 0x80);
    assert_invalid("invalid UTF-8", invalid_utf8);
}

#[test]
fn every_truncation_and_single_byte_mutation_is_panic_free() {
    let bytes = ODC.as_bytes();
    for end in 0..bytes.len() {
        let _ = exercise(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        let _ = exercise(mutated);
    }
    assert!(exercise(bytes.to_vec()).is_none());
}
