#![allow(
    clippy::unwrap_used,
    reason = "tests are expected to panic on unexpected errors"
)]

use litchi_odc::flat::Chart as FlatChart;
use litchi_odc::{AxisUpdate, ChartClassKind};

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
    match FlatChart::from_bytes(&bytes) {
        Ok(flat) => {
            if let Some(plot) = flat.plot_area() {
                for axis in plot.axes() {
                    if let Err(error) = axis.dimension() {
                        return Some(error);
                    }
                }
                for series in plot.series() {
                    std::hint::black_box(series.values_range());
                    std::hint::black_box(series.attached_axis());
                }
            }
            None
        },
        Err(error) => Some(error.into()),
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
    let flat = FlatChart::from_bytes(ODC.as_bytes()).unwrap();
    assert_eq!(flat.as_bytes(), ODC.as_bytes());
    assert!(flat.chart().all_text().contains("Revenue"));
    assert!(flat.find_axis("primary-x").is_some());
    assert_eq!(
        flat.chart().chart_class().unwrap().kind(),
        ChartClassKind::Bar
    );
    assert_invalid(
        "wrong family",
        r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body><office:drawing/></office:body></office:document>"#
            .as_bytes()
            .to_vec(),
    );
    assert_invalid("packaged input", b"PK\x03\x04mimetype".to_vec());
}

#[test]
fn flat_chart_preserves_a_chart_class_alias_losslessly() {
    let input = ODC
        .replace("chart:class=\"chart:bar\"", "chart:class=\"c:bar\"")
        .replace(
            "xmlns:chart=",
            "xmlns:c=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" xmlns:chart=",
        );
    let flat = FlatChart::from_bytes(&input.clone().into_bytes()).unwrap();
    let class = flat.chart().chart_class().unwrap();
    assert_eq!(class.kind(), ChartClassKind::Bar);
    assert_eq!(class.lexical(), "c:bar");
    assert_eq!(
        flat.edit().commit().unwrap().package().as_bytes(),
        input.as_bytes()
    );
}

#[test]
fn flat_axis_transaction_is_lossless_source_checked_and_reversible() {
    let source = FlatChart::from_bytes(ODC.as_bytes()).unwrap();
    let no_op = source.edit().commit().unwrap();
    assert_eq!(no_op.package().as_bytes(), source.as_bytes());
    assert!(no_op.patch().changes().is_empty());

    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("renamed-x")).unwrap();
    let commit = edit.commit().unwrap();
    let output = std::str::from_utf8(commit.package().as_bytes()).unwrap();
    assert!(output.contains("style:name=\"keep-me\""));
    assert!(output.contains("<text:p>Revenue</text:p>"));
    assert!(!output.contains(">\n<"));
    assert!(!output.contains(">\r\n<"));
    assert!(commit.package().find_axis("renamed-x").is_some());
    assert!(source.find_axis("primary-x").is_some());
    assert!(commit.patch().is_applicable_to(&source));

    let unrelated =
        FlatChart::from_bytes(&ODC.replace("keep-me", "also-kept").into_bytes()).unwrap();
    assert!(!commit.patch().is_applicable_to(&unrelated));
    assert!(matches!(
        unrelated.apply(commit.patch()),
        Err(litchi_odc::flat::Error::PatchConflict { .. })
    ));

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply(&inverse).unwrap();
    assert!(restored.diagnostics().changed());
    assert_eq!(restored.package().as_bytes(), source.as_bytes());
    assert!(commit.package().apply(commit.patch()).is_err());
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
    assert!(FlatChart::from_bytes(&foreign).is_ok());
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

#[test]
fn flat_chart_takes_caller_limits_through_open_options() {
    use litchi_odc::{LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits};

    let one_axis = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_axes(1)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(FlatChart::from_bytes_with_options(ODC.as_bytes(), one_axis).is_err());
    let shallow = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_depth(2)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(FlatChart::from_bytes_with_options(ODC.as_bytes(), shallow).is_err());

    let source = FlatChart::from_bytes(ODC.as_bytes()).unwrap();
    assert_eq!(source.options(), OpenOptions::new());
    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("renamed-x")).unwrap();
    edit.reset();
    let noop = edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().as_bytes(), source.as_bytes());
    let reread = FlatChart::from_reader(std::io::Cursor::new(ODC.as_bytes())).unwrap();
    assert_eq!(reread.as_bytes(), source.as_bytes());
    let tiny = OpenOptions::new().source_limits(
        litchi_odc::SourceLimits::builder()
            .max_input_bytes(8)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(matches!(
        FlatChart::from_bytes_with_options(ODC.as_bytes(), tiny),
        Err(litchi_odc::Error::Limit {
            kind: LimitKind::Source(_),
            ..
        })
    ));
    let _ = SemanticLimitKind::Axes;
}
