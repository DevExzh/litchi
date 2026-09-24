use litchi_odc::{
    AxisSpec, AxisUpdate, Builder, Chart, ChartErrorCategory, ChartPackageKind, ChartStyleProperty,
    ChartStyleValue, ChartSymbolName, ChartSymbolType, Definition, Limits, chart::Dimension,
};
use litchi_odf_common::core::OwnedPackage;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::error::Error;

type TestResult<T = ()> = Result<T, Box<dyn Error>>;

const STYLES: &str = concat!(
    "<?xml version=\"1.0\"?><office:document-styles ",
    "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" ",
    "xmlns:chart=\"urn:oasis:names:tc:opendocument:xmlns:chart:1.0\" ",
    "xmlns:ext=\"urn:example:unknown\"><office:styles>",
    "<style:style style:name=\"axisStyle\" style:family=\"chart\"><style:chart-properties ",
    "chart:logarithmic=\"false\" chart:tick-marks-major-inner=\"true\" ",
    "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"circle\" ",
    "chart:error-category=\"standard-error\" chart:regression-type=\"linear\" ",
    "ext:opaque=\" preserved &amp; lexical \"/><!--keep-comment--></style:style>",
    "</office:styles></office:document-styles>"
);

fn definition() -> Definition {
    let mut definition = Definition::new(litchi_odc::ChartClass::line());
    let mut axis = AxisSpec::new(Dimension::X);
    axis.style_name = Some("axisStyle".into());
    definition.plot_area.axes.push(axis);
    definition
}

fn with_pretty_styles(bytes: &[u8]) -> TestResult<Vec<u8>> {
    let reader = ArchiveReader::new(bytes)?;
    let mut writer = StreamingArchiveWriter::new();
    for name in reader.file_names() {
        let payload = if name == "styles.xml" {
            STYLES.replace("><", ">\n<").into_bytes()
        } else {
            reader.read(name)?
        };
        if name == "mimetype" {
            writer.write_stored(name, &payload)?;
        } else {
            writer.write_deflated(name, &payload)?;
        }
    }
    Ok(writer.finish_to_bytes()?)
}

#[test]
fn template_mime_and_manifest_kind_survive_a_public_edit() -> TestResult<()> {
    let source_bytes = Builder::new()
        .with_definition(definition())
        .with_package_kind(ChartPackageKind::Template)
        .with_styles_xml(STYLES)
        .build()?;
    let source = Chart::from_bytes(source_bytes.clone())?;
    assert_eq!(source.package_kind(), ChartPackageKind::Template);
    assert_eq!(source.mime_type(), ChartPackageKind::Template.mime_type());
    assert!(source.is_template());

    let mut edit = source.edit();
    edit.update_axis(0, AxisUpdate::named("renamed"))?;
    let committed = edit.commit()?.into_chart();
    assert_eq!(committed.package_kind(), ChartPackageKind::Template);
    assert_eq!(
        committed.mime_type(),
        ChartPackageKind::Template.mime_type()
    );

    let archive = OwnedPackage::from_bytes(committed.as_bytes().to_vec())?;
    assert_eq!(archive.mimetype()?, ChartPackageKind::Template.mime_type());
    let package = archive.package()?;
    assert_eq!(
        package.manifest().get_media_type("/"),
        Some(ChartPackageKind::Template.mime_type())
    );
    assert_eq!(
        package.manifest().get_media_type("content.xml"),
        Some("text/xml")
    );
    assert_ne!(committed.as_bytes(), source_bytes.as_slice());
    Ok(())
}

#[test]
fn chart_style_properties_are_typed_atomic_and_exact() -> TestResult<()> {
    let compact = Builder::new()
        .with_definition(definition())
        .with_styles_xml(STYLES)
        .build()?;
    let source = Chart::from_bytes(with_pretty_styles(&compact)?)?;
    assert_eq!(
        source.chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(false))
    );
    assert_eq!(
        source.chart_style_property("axisStyle", ChartStyleProperty::SymbolName)?,
        Some(ChartStyleValue::SymbolName(ChartSymbolName::Circle))
    );

    let mut edit = source.edit();
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    edit.update_chart_style_property("axisStyle", ChartStyleProperty::SymbolName, None)?;
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::SymbolType,
        Some(ChartStyleValue::SymbolType(ChartSymbolType::Automatic)),
    )?;
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::ErrorCategory,
        Some(ChartStyleValue::ErrorCategory(
            ChartErrorCategory::Percentage,
        )),
    )?;
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMinorOuter,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMinorInner,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    let commit = edit.commit()?;
    assert_eq!(commit.patch().style_property_changes().len(), 6);
    let decoded = litchi_odc::Patch::from_bytes(&commit.patch().to_bytes(), source.limits())?;
    let mut decoded_changes = decoded.style_property_changes().to_vec();
    let mut committed_changes = commit.patch().style_property_changes().to_vec();
    decoded_changes.sort_by_key(|change| (change.style_name().to_owned(), change.property()));
    committed_changes.sort_by_key(|change| (change.style_name().to_owned(), change.property()));
    assert_eq!(decoded_changes, committed_changes);
    let changed = commit.chart();
    assert_eq!(
        changed.chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(true))
    );
    assert!(changed.styles_xml().is_some_and(|styles| {
        styles.contains(">\n<")
            && styles.contains("ext:opaque=\" preserved &amp; lexical \"")
            && styles.contains("<!--keep-comment-->")
            && styles.contains("chart:logarithmic=\"true\"")
            && styles.contains("chart:tick-marks-minor-outer=\"true\"")
            && styles.contains("chart:tick-marks-minor-inner=\"true\"")
    }));
    assert_eq!(
        commit.patch().inverse().apply(changed)?.as_bytes(),
        source.as_bytes()
    );

    let mut noop = source.edit();
    noop.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(false)),
    )?;
    let noop_commit = noop.commit()?;
    assert!(!noop_commit.changed());
    assert_eq!(noop_commit.chart().as_bytes(), source.as_bytes());

    let mut wrong_type = source.edit();
    assert!(
        wrong_type
            .update_chart_style_property(
                "axisStyle",
                ChartStyleProperty::Logarithmic,
                Some(ChartStyleValue::Text("true".into()))
            )
            .is_err()
    );
    let mut missing = source.edit();
    assert!(
        missing
            .update_chart_style_property(
                "missing",
                ChartStyleProperty::Logarithmic,
                Some(ChartStyleValue::Boolean(true))
            )
            .is_err()
    );
    Ok(())
}

#[test]
fn native_odfdom_chart_fixture_opens_and_keeps_chart_kind() -> TestResult<()> {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/odf/odc-producer-evidence/odfdom-created.odc"
    );
    let chart = Chart::open(path)?;
    assert_eq!(chart.package_kind(), ChartPackageKind::Chart);
    assert_eq!(chart.mime_type(), ChartPackageKind::Chart.mime_type());
    assert!(chart.content_xml().contains("chart:chart"));
    Ok(())
}

#[test]
fn chart_style_types_follow_odf_lexicals_and_retain_prefixes() -> TestResult<()> {
    let special = STYLES.replace(
        "chart:regression-type=\"linear\"",
        "chart:regression-type=\"linear\" chart:regression-name=\"\" chart:error-margin=\"INF\"",
    );
    let chart = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(special)
            .build()?,
    )?;
    assert_eq!(
        chart.chart_style_property("axisStyle", ChartStyleProperty::RegressionName)?,
        Some(ChartStyleValue::Text(String::new()))
    );
    assert_eq!(
        chart.chart_style_property("axisStyle", ChartStyleProperty::ErrorMargin)?,
        Some(ChartStyleValue::Decimal(f64::INFINITY))
    );

    let bad_boolean = STYLES.replace("chart:logarithmic=\"false\"", "chart:logarithmic=\"1\"");
    let bad_boolean = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(bad_boolean)
            .build()?,
    )?;
    assert!(
        bad_boolean
            .chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)
            .is_err()
    );

    let bad_degree = STYLES.replace(
        "chart:regression-type=\"linear\"",
        "chart:regression-type=\"linear\" chart:regression-max-degree=\"1\"",
    );
    let bad_degree = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(bad_degree)
            .build()?,
    )?;
    assert!(
        bad_degree
            .chart_style_property("axisStyle", ChartStyleProperty::RegressionMaxDegree)
            .is_err()
    );

    let prefixed = STYLES
        .replace("xmlns:chart=", "xmlns:c=")
        .replace(" chart:", " c:");
    let prefixed = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(prefixed)
            .build()?,
    )?;
    let mut edit = prefixed.edit();
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMinorInner,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    let changed = edit.commit()?.into_chart();
    assert!(
        changed
            .styles_xml()
            .is_some_and(|styles| styles.contains("c:tick-marks-minor-inner=\"true\""))
    );

    let automatic_with_name = STYLES.replace(
        "chart:symbol-type=\"named-symbol\"",
        "chart:symbol-type=\"automatic\"",
    );
    let automatic_with_name = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(automatic_with_name)
            .build()?,
    )?;
    assert!(
        automatic_with_name
            .chart_style_property("axisStyle", ChartStyleProperty::SymbolType)
            .is_err()
    );

    let image_without_element = STYLES.replace(
        "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"circle\"",
        "chart:symbol-type=\"image\"",
    );
    let image_without_element = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(image_without_element)
            .build()?,
    )?;
    assert!(
        image_without_element
            .chart_style_property("axisStyle", ChartStyleProperty::SymbolType)
            .is_err()
    );

    let image_styles = STYLES
        .replace(
            "xmlns:ext=\"urn:example:unknown\"",
            "xmlns:ext=\"urn:example:unknown\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        )
        .replace(
            "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"circle\"",
            "chart:symbol-type=\"image\"",
        )
        .replace(
            "ext:opaque=\" preserved &amp; lexical \"/><!--keep-comment--></style:style>",
            "ext:opaque=\" preserved &amp; lexical \"><chart:symbol-image xlink:href=\"Pictures/s.png\"/></style:chart-properties><!--keep-comment--></style:style>",
        );
    let image_chart = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(image_styles.clone())
            .build()?,
    )?;
    assert_eq!(
        image_chart.chart_style_property("axisStyle", ChartStyleProperty::SymbolType)?,
        Some(ChartStyleValue::SymbolType(ChartSymbolType::Image))
    );

    let image_without_href = image_styles.replace(" xlink:href=\"Pictures/s.png\"", "");
    let image_without_href = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(image_without_href)
            .build()?,
    )?;
    assert!(
        image_without_href
            .chart_style_property("axisStyle", ChartStyleProperty::SymbolType)
            .is_err()
    );
    Ok(())
}

#[test]
fn chart_style_names_accept_xml_ncname_combining_marks() -> TestResult<()> {
    let combining_name = "a\u{301}";
    let replacement = format!("style:name=\"{combining_name}\"");
    let styles = STYLES.replace("style:name=\"axisStyle\"", &replacement);
    let chart = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(styles)
            .build()?,
    )?;
    assert_eq!(
        chart.chart_style_property(combining_name, ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(false))
    );
    let mut edit = chart.edit();
    edit.update_chart_style_property(
        combining_name,
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    assert!(edit.commit()?.changed());
    Ok(())
}

#[test]
fn disjoint_chart_style_patches_join_and_transfer_by_property() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;

    let mut left_edit = source.edit();
    left_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    let left_commit = left_edit.commit()?;
    let left_patch = left_commit.patch().clone();

    let mut right_edit = source.edit();
    right_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMajorInner,
        Some(ChartStyleValue::Boolean(false)),
    )?;
    let right_commit = right_edit.commit()?;
    let right_patch = right_commit.patch().clone();
    let right_chart = right_commit.into_chart();
    let merged = left_patch.join(&right_patch)?;
    assert!(merged.is_merged(), "disjoint typed style writes must join");
    let merged_chart = merged.patch().expect("merged patch").apply(&source)?;
    assert_eq!(
        merged_chart.chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(true))
    );
    assert_eq!(
        merged_chart.chart_style_property("axisStyle", ChartStyleProperty::TickMarksMajorInner)?,
        Some(ChartStyleValue::Boolean(false))
    );

    let transferred = left_patch.transfer_to(&right_chart)?;
    assert!(
        transferred.is_merged(),
        "disjoint typed style writes must transfer"
    );
    let transferred_chart = transferred
        .patch()
        .expect("transferred patch")
        .apply(&right_chart)?;
    assert_eq!(
        transferred_chart.chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(true))
    );
    assert_eq!(
        transferred_chart
            .chart_style_property("axisStyle", ChartStyleProperty::TickMarksMajorInner)?,
        Some(ChartStyleValue::Boolean(false))
    );
    Ok(())
}

#[test]
fn typed_style_publication_rejects_output_above_content_limit_before_rebuild() -> TestResult<()> {
    let padded_styles = STYLES.replace(
        " ext:opaque=",
        &format!(" ext:padding=\"{}\" ext:opaque=", "x".repeat(2_048)),
    );
    let bytes = Builder::new()
        .with_definition(definition())
        .with_styles_xml(padded_styles)
        .build()?;
    let unconstrained = Chart::from_bytes(bytes.clone())?;
    let content_limit = unconstrained.styles_xml().map_or(0, str::len);
    let limits = Limits::new().with_content_bytes(content_limit)?;
    let bounded = Chart::from_bytes_with_limits(bytes, limits)?;
    let mut edit = bounded.edit();
    edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMinorInner,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    assert!(edit.commit().is_err());
    Ok(())
}

#[test]
fn composed_style_changes_collapse_replays_and_survive_wire_roundtrip() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;

    let mut first_edit = source.edit();
    first_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    first_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::RegressionName,
        Some(ChartStyleValue::Text("curve".into())),
    )?;
    let first = first_edit.commit()?;
    let first_patch = first.patch().clone();
    let first_chart = first.into_chart();

    let mut second_edit = first_chart.edit();
    second_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(false)),
    )?;
    second_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::RegressionName,
        None,
    )?;
    second_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::TickMarksMajorInner,
        Some(ChartStyleValue::Boolean(false)),
    )?;
    let second = second_edit.commit()?;
    let second_patch = second.patch().clone();

    let composed = first_patch.compose(&second_patch)?;
    assert_eq!(composed.style_property_changes().len(), 1);
    assert_eq!(
        composed.style_property_changes()[0].property(),
        ChartStyleProperty::TickMarksMajorInner
    );
    assert_eq!(
        composed.style_property_changes()[0].before(),
        Some(&ChartStyleValue::Boolean(true))
    );
    assert_eq!(
        composed.style_property_changes()[0].after(),
        Some(&ChartStyleValue::Boolean(false))
    );

    let transferred = composed.transfer_to(&source)?;
    assert!(transferred.is_merged());
    let transferred_chart = transferred
        .patch()
        .expect("transferred patch")
        .apply(&source)?;
    assert_eq!(
        transferred_chart.chart_style_property("axisStyle", ChartStyleProperty::Logarithmic)?,
        Some(ChartStyleValue::Boolean(false))
    );
    assert_eq!(
        transferred_chart
            .chart_style_property("axisStyle", ChartStyleProperty::TickMarksMajorInner)?,
        Some(ChartStyleValue::Boolean(false))
    );

    let durable = litchi_odc::Patch::from_bytes(&composed.to_bytes(), source.limits())?;
    assert_eq!(
        durable.style_property_changes(),
        composed.style_property_changes()
    );
    let durable_transfer = durable.transfer_to(&source)?;
    assert!(durable_transfer.is_merged());
    Ok(())
}

#[test]
fn durable_added_style_falls_back_to_lossless_whole_part_join_and_transfer() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;
    let added_styles = STYLES.replace(
        "</office:styles>",
        "<style:style style:name=\"addedStyle\" style:family=\"chart\"><style:chart-properties chart:logarithmic=\"true\"/></style:style></office:styles>",
    );
    let mut style_edit = source.edit();
    style_edit.set_styles_xml(added_styles.clone());
    let style_commit = style_edit.commit()?;
    let durable = litchi_odc::Patch::from_bytes(&style_commit.patch().to_bytes(), source.limits())?;
    assert!(
        durable
            .style_property_changes()
            .iter()
            .any(|change| change.style_name() == "addedStyle")
    );

    let transferred = durable.transfer_to(&source)?;
    assert!(transferred.is_merged());
    let transferred_chart = transferred
        .patch()
        .expect("transferred patch")
        .apply(&source)?;
    assert_eq!(transferred_chart.styles_xml(), Some(added_styles.as_str()));

    let mut right_edit = source.edit();
    right_edit.update_axis(0, AxisUpdate::named("joined"))?;
    let right = right_edit.commit()?;
    let joined = durable.join(right.patch())?;
    assert!(joined.is_merged());
    let joined_chart = joined.patch().expect("joined patch").apply(&source)?;
    assert_eq!(joined_chart.styles_xml(), Some(added_styles.as_str()));
    assert_eq!(
        joined_chart
            .plot_area()
            .and_then(|plot| plot.axes().next())
            .and_then(|axis| axis.name()),
        Some("joined")
    );

    let invalid_styles = STYLES.replace("chart:logarithmic=\"false\"", "chart:logarithmic=\"1\"");
    let mut invalid_edit = source.edit();
    invalid_edit.set_styles_xml(invalid_styles);
    let invalid_commit = invalid_edit.commit()?;
    assert!(
        litchi_odc::Patch::from_bytes(&invalid_commit.patch().to_bytes(), source.limits()).is_err()
    );
    Ok(())
}

#[test]
fn durable_symbol_image_child_uses_whole_part_join_and_transfer() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;
    let image_styles = STYLES
        .replace(
            "xmlns:ext=\"urn:example:unknown\"",
            "xmlns:ext=\"urn:example:unknown\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        )
        .replace(
            "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"circle\"",
            "chart:symbol-type=\"image\"",
        )
        .replace(
            "ext:opaque=\" preserved &amp; lexical \"/><!--keep-comment--></style:style>",
            "ext:opaque=\" preserved &amp; lexical \"><chart:symbol-image xlink:href=\"Pictures/s.png\"/></style:chart-properties><!--keep-comment--></style:style>",
        );
    let mut edit = source.edit();
    edit.set_styles_xml(image_styles.clone());
    let committed = edit.commit()?;
    let durable = litchi_odc::Patch::from_bytes(&committed.patch().to_bytes(), source.limits())?;
    assert!(
        durable
            .style_property_changes()
            .iter()
            .any(|change| change.property() == ChartStyleProperty::SymbolType)
    );

    let transferred = durable.transfer_to(&source)?;
    assert!(transferred.is_merged());
    let transferred_chart = transferred
        .patch()
        .expect("transferred patch")
        .apply(&source)?;
    assert_eq!(transferred_chart.styles_xml(), Some(image_styles.as_str()));

    let mut right_edit = source.edit();
    right_edit.update_axis(0, AxisUpdate::named("joined"))?;
    let right = right_edit.commit()?;
    let joined = durable.join(right.patch())?;
    assert!(joined.is_merged());
    let joined_chart = joined.patch().expect("joined patch").apply(&source)?;
    assert_eq!(joined_chart.styles_xml(), Some(image_styles.as_str()));
    assert_eq!(
        joined_chart
            .plot_area()
            .and_then(|plot| plot.axes().next())
            .and_then(|axis| axis.name()),
        Some("joined")
    );
    Ok(())
}

#[test]
fn typed_transfer_missing_chart_properties_reports_conflict_before_staging() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;
    let mut left_edit = source.edit();
    left_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(true)),
    )?;
    let left = left_edit.commit()?;
    let properties = concat!(
        "<style:chart-properties chart:logarithmic=\"false\" ",
        "chart:tick-marks-major-inner=\"true\" ",
        "chart:symbol-type=\"named-symbol\" chart:symbol-name=\"circle\" ",
        "chart:error-category=\"standard-error\" chart:regression-type=\"linear\" ",
        "ext:opaque=\" preserved &amp; lexical \"/>"
    );
    let destination_styles = STYLES.replace(properties, "");
    let destination = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(destination_styles)
            .build()?,
    )?;
    let transferred = left.patch().transfer_to(&destination)?;
    assert!(!transferred.is_merged());
    assert_eq!(
        transferred
            .conflicts()
            .iter()
            .map(|conflict| conflict.path())
            .collect::<Vec<_>>(),
        vec!["package.styles[axisStyle].Logarithmic"]
    );
    Ok(())
}

#[test]
fn composed_whole_style_then_typed_style_uses_final_whole_part() -> TestResult<()> {
    let source = Chart::from_bytes(
        Builder::new()
            .with_definition(definition())
            .with_styles_xml(STYLES)
            .build()?,
    )?;
    let structural_styles = STYLES
        .replace("chart:logarithmic=\"false\"", "chart:logarithmic=\"true\"")
        .replace(
            "</office:styles>",
            "<style:style style:name=\"addedStyle\" style:family=\"chart\"><style:chart-properties chart:logarithmic=\"true\"/></style:style></office:styles>",
        );
    let mut structural_edit = source.edit();
    structural_edit.set_styles_xml(structural_styles);
    let structural = structural_edit.commit()?;
    let structural_patch = structural.patch().clone();
    let structural_chart = structural.into_chart();
    let mut typed_edit = structural_chart.edit();
    typed_edit.update_chart_style_property(
        "axisStyle",
        ChartStyleProperty::Logarithmic,
        Some(ChartStyleValue::Boolean(false)),
    )?;
    let typed = typed_edit.commit()?;
    let composed = structural_patch.compose(typed.patch())?;
    let transferred = composed.transfer_to(&source)?;
    assert!(transferred.is_merged());
    let chart = transferred
        .patch()
        .expect("transferred patch")
        .apply(&source)?;
    assert_eq!(chart.styles_xml(), typed.chart().styles_xml());
    Ok(())
}
