use litchi_odc::{
    AxisSpec, Builder, CachedCell, CachedRow, CachedTable, CachedValue, ChartClass, DataPointSpec,
    Definition, ExactAttribute, ExactTarget, GridSpec, History, LegendSpec, OpenOptions, Package,
    Patch, SemanticLimits, SeriesSpec, StyleElement, StyleTarget, Text,
    chart::{Class, Dimension, Labels},
    validate_range_list,
};
use litchi_odf_common::core::{PackageWriter, Profile};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::{error::Error, fmt::Write as _};

const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.4"><office:styles/></office:document-styles>"#;
const RAW_MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;

type TestResult<T = ()> = Result<T, Box<dyn Error>>;

fn definition() -> Definition {
    let mut value = Definition::new(ChartClass::bar());
    let mut horizontal = AxisSpec::new(Dimension::X);
    horizontal.name = Some("x".into());
    value.plot_area.axes.push(horizontal);
    let mut vertical = AxisSpec::new(Dimension::Y);
    vertical.name = Some("y".into());
    value.plot_area.axes.push(vertical);
    value.plot_area.series.push(SeriesSpec {
        attached_axis: Some("x".into()),
        data_points: vec![DataPointSpec::default()],
        ..SeriesSpec::default()
    });
    value
}

fn package(content: &str, auxiliary: Option<(&str, &[u8])>) -> TestResult<Vec<u8>> {
    // Intentionally noncanonical parser inputs bypass the fail-closed
    // production writer through this test-only raw archive constructor.
    let mut manifest = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.4\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.chart\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>",
    );
    if let Some((path, bytes)) = auxiliary {
        write!(
            &mut manifest,
            "<manifest:file-entry manifest:full-path=\"{path}\" manifest:media-type=\"application/octet-stream\"/>"
        )?;
        let mut archive =
            StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
        archive.write_stored("mimetype", b"application/vnd.oasis.opendocument.chart")?;
        archive.write_deflated("content.xml", content.as_bytes())?;
        archive.write_deflated(path, bytes)?;
        manifest.push_str("</manifest:manifest>");
        archive.write_deflated("META-INF/manifest.xml", manifest.as_bytes())?;
        return Ok(archive.finish_to_bytes()?);
    }
    manifest.push_str("</manifest:manifest>");
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", b"application/vnd.oasis.opendocument.chart")?;
    archive.write_deflated("content.xml", content.as_bytes())?;
    archive.write_deflated("META-INF/manifest.xml", manifest.as_bytes())?;
    Ok(archive.finish_to_bytes()?)
}

fn encrypted_envelope(content: &str) -> TestResult<Vec<u8>> {
    let mut encrypted_writer =
        PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
    encrypted_writer.set_mimetype("application/vnd.oasis.opendocument.chart")?;
    encrypted_writer.set_encryption("password", Profile::compatible())?;
    encrypted_writer.add_file("content.xml", content.as_bytes())?;
    let encrypted = encrypted_writer.finish_to_bytes()?;
    let encrypted_archive =
        ArchiveReader::new(&encrypted, soapberry_zip::office::ArchiveLimits::default())?;
    let ciphertext = encrypted_archive.read("content.xml")?;
    let manifest = String::from_utf8(encrypted_archive.read("META-INF/manifest.xml")?)?.replace(
        "full-path=\"content.xml\"",
        "full-path=\"META-INF/secret.bin\"",
    );

    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", b"application/vnd.oasis.opendocument.chart")?;
    archive.write_deflated("content.xml", content.as_bytes())?;
    archive.write_deflated("META-INF/secret.bin", &ciphertext)?;
    archive.write_deflated("META-INF/manifest.xml", manifest.as_bytes())?;
    Ok(archive.finish_to_bytes()?)
}

fn raw_negative_package(content: &str) -> TestResult<Vec<u8>> {
    let mut archive =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    archive.write_stored("mimetype", b"application/vnd.oasis.opendocument.chart")?;
    archive.write_deflated("content.xml", content.as_bytes())?;
    archive.write_deflated("META-INF/manifest.xml", RAW_MANIFEST)?;
    Ok(archive.finish_to_bytes()?)
}

#[test]
fn opened_canonical_definition_edits_are_atomic_and_reversible() -> TestResult<()> {
    let source_definition = definition();
    let source = Builder::new()
        .with_definition(source_definition.clone())
        .build()?;
    assert_eq!(source.chart().definition()?, source_definition);

    let mut edit = source.edit();
    edit.insert_axis(2, AxisSpec::new(Dimension::Z))?;
    edit.insert_series(1, SeriesSpec::default())?;
    edit.insert_data_point(1, 0, DataPointSpec::default())?;
    edit.set_style(StyleTarget::Series(1), Some("added".into()))?;
    let commit = edit.commit()?;
    assert_eq!(commit.patch().definition_changes().len(), 4);
    let reopened = Package::from_bytes(commit.package().as_bytes())?;
    let projected = reopened.chart().definition()?;
    assert_eq!(projected.plot_area.axes.len(), 3);
    assert_eq!(projected.plot_area.series.len(), 2);
    assert_eq!(
        projected.plot_area.series[1].style_name.as_deref(),
        Some("added")
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())?
            .into_package()
            .as_bytes(),
        source.as_bytes()
    );
    Ok(())
}

#[test]
fn opened_projection_refuses_noncanonical_xml_without_mutation() -> TestResult<()> {
    let canonical = litchi_odc::serialize_content(&definition())?;
    let noncanonical = canonical.replacen("><", ">\n<", 1);
    let chart = Package::from_bytes(&raw_negative_package(&noncanonical)?)?;
    let original = chart.as_bytes().to_vec();
    assert!(chart.chart().definition().is_err());
    let mut edit = chart.edit();
    assert!(edit.insert_axis(0, AxisSpec::new(Dimension::X)).is_err());
    edit.update_axis(0, litchi_odc::AxisUpdate::styled("exact-style"))?;
    let committed = edit.commit()?;
    assert!(committed.package().chart().content_xml().contains(">\n<"));
    assert_eq!(
        committed
            .package()
            .chart()
            .plot_area()
            .ok_or("missing plot area")?
            .axes()
            .next()
            .ok_or("missing axis")?
            .style_name(),
        Some("exact-style")
    );
    assert_eq!(chart.as_bytes(), original);
    Ok(())
}

#[test]
fn noncanonical_chart_plot_and_series_attributes_use_checked_exact_spans() -> TestResult<()> {
    let mut authored = definition();
    authored.width = Some("10cm".into());
    authored.height = Some("8cm".into());
    authored.plot_area.cell_range_address = Some("Data.A1:.B4".into());
    authored.plot_area.x = Some("1cm".into());
    authored.plot_area.y = Some("2cm".into());
    authored.plot_area.width = Some("9cm".into());
    authored.plot_area.height = Some("7cm".into());
    authored.plot_area.series[0].values_cell_range_address = Some("Data.B2:.B4".into());
    authored.plot_area.series[0].label_cell_address = Some("Data.B1".into());
    let canonical = litchi_odc::serialize_content(&authored)?;
    let noncanonical = canonical.replacen(
        "><office:body>",
        ">\n<!-- retained producer extension boundary -->\n<office:body>",
        1,
    );
    let source = Package::from_bytes(&raw_negative_package(&noncanonical)?)?;
    assert!(source.chart().definition().is_err());

    let mut edit = source.edit();
    edit.update_exact(
        ExactTarget::Chart,
        ExactAttribute::Class,
        Some("chart:line".into()),
    )?;
    edit.update_exact(
        ExactTarget::Chart,
        ExactAttribute::Width,
        Some("11cm".into()),
    )?;
    edit.update_exact(ExactTarget::Chart, ExactAttribute::Height, None)?;
    edit.update_exact(
        ExactTarget::PlotArea,
        ExactAttribute::CellRangeAddress,
        Some("Data.A1:.C4".into()),
    )?;
    edit.update_exact(ExactTarget::PlotArea, ExactAttribute::X, Some("3cm".into()))?;
    edit.update_exact(
        ExactTarget::PlotArea,
        ExactAttribute::StyleName,
        Some("plot-style".into()),
    )?;
    edit.update_exact(
        ExactTarget::Series(0),
        ExactAttribute::Class,
        Some("chart:line".into()),
    )?;
    edit.update_exact(
        ExactTarget::Series(0),
        ExactAttribute::ValuesCellRangeAddress,
        Some("Data.C2:.C4".into()),
    )?;
    edit.update_exact(
        ExactTarget::Series(0),
        ExactAttribute::AttachedAxis,
        Some("y".into()),
    )?;
    edit.update_exact(
        ExactTarget::Series(0),
        ExactAttribute::StyleName,
        Some("series-style".into()),
    )?;
    let commit = edit.commit()?;
    assert_eq!(commit.patch().exact_changes().len(), 10);
    assert!(
        commit
            .package()
            .chart()
            .content_xml()
            .contains("\n<!-- retained producer extension boundary -->\n")
    );
    assert!(
        commit
            .package()
            .chart()
            .content_xml()
            .contains("svg:width=\"11cm\"")
    );
    assert!(
        !commit
            .package()
            .chart()
            .content_xml()
            .contains("svg:height=\"8cm\"")
    );
    assert!(
        commit
            .package()
            .chart()
            .content_xml()
            .contains("chart:style-name=\"plot-style\"")
    );
    assert!(
        commit
            .package()
            .chart()
            .content_xml()
            .contains("chart:attached-axis=\"y\"")
    );
    assert_eq!(commit.package().chart().class()?, ChartClass::line());
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())?
            .into_package()
            .as_bytes(),
        source.as_bytes()
    );

    let wire = commit.patch().to_bytes();
    let decoded = Patch::from_bytes(&wire, source.chart().options())?;
    assert_eq!(decoded.exact_changes().len(), 10);
    assert_eq!(
        source.apply(&decoded)?.into_package().as_bytes(),
        commit.package().as_bytes()
    );
    Ok(())
}

#[test]
fn odf_1_4_coordinate_region_exact_edit_transfers_without_normalizing() -> TestResult<()> {
    // Derived from the OASIS ODF 1.4 OS Relax NG
    // `chart-coordinate-region` definition. This is standards-derived test
    // XML, not a producer fixture or repackaged embedded chart.
    let mut authored = definition();
    authored.width = Some("10cm".into());
    let canonical = litchi_odc::serialize_content(&authored)?;
    let standards_derived = canonical.replacen(
        "<chart:axis",
        "<chart:coordinate-region svg:x=\"1cm\" svg:y=\"2cm\" svg:width=\"8cm\" svg:height=\"6cm\"/><chart:axis",
        1,
    );
    let source = Package::from_bytes(&raw_negative_package(&standards_derived)?)?;
    assert!(source.chart().definition().is_err());

    let mut edit = source.edit();
    edit.update_exact(
        ExactTarget::CoordinateRegion,
        ExactAttribute::X,
        Some("3cm".into()),
    )?;
    edit.update_exact(
        ExactTarget::CoordinateRegion,
        ExactAttribute::Width,
        Some("7cm".into()),
    )?;
    let changed = edit.commit()?;
    assert_eq!(changed.patch().exact_changes().len(), 2);
    assert!(changed
        .package()
        .chart().content_xml()
        .contains("<chart:coordinate-region svg:x=\"3cm\" svg:y=\"2cm\" svg:width=\"7cm\" svg:height=\"6cm\"/>"));

    let destination_xml = standards_derived.replacen(
        "<office:body>",
        "<!-- independently retained -->\n<office:body>",
        1,
    );
    let destination = Package::from_bytes(&raw_negative_package(&destination_xml)?)?;
    let transfer = changed.patch().transfer_to(&destination)?;
    assert!(transfer.is_merged());
    let transferred = destination
        .apply(
            transfer
                .patch()
                .ok_or("missing coordinate-region transfer patch")?,
        )?
        .into_package();
    assert!(
        transferred
            .chart()
            .content_xml()
            .contains("<!-- independently retained -->\n")
    );
    assert!(transferred.chart().content_xml()
        .contains("<chart:coordinate-region svg:x=\"3cm\" svg:y=\"2cm\" svg:width=\"7cm\" svg:height=\"6cm\"/>"));
    Ok(())
}

#[test]
fn noncanonical_titles_footers_and_legends_have_controlled_exact_edits() -> TestResult<()> {
    let mut authored = definition();
    authored.title = Some(Text {
        text: "Producer title".into(),
        cell_range: Some("Data.A1".into()),
        style_name: Some("title-old".into()),
        x: Some("1cm".into()),
        y: Some("2cm".into()),
        ..Text::default()
    });
    authored.footer = Some(Text::new("Producer footer"));
    authored.legend = Some(LegendSpec {
        style_name: Some("legend-old".into()),
        x: Some("12cm".into()),
        y: Some("1cm".into()),
        ..LegendSpec::default()
    });
    let canonical = litchi_odc::serialize_content(&authored)?;
    let noncanonical = canonical
        .replacen(
            "<chart:legend",
            "<chart:legend chart:legend-align=\"center\"",
            1,
        )
        .replacen("><office:body>", ">\n<office:body>", 1);
    let source = Package::from_bytes(&raw_negative_package(&noncanonical)?)?;
    assert!(source.chart().definition().is_err());

    let mut edit = source.edit();
    edit.update_exact(
        ExactTarget::Title,
        ExactAttribute::StyleName,
        Some("title-new".into()),
    )?;
    edit.update_exact(
        ExactTarget::Title,
        ExactAttribute::CellRange,
        Some("Data.B1".into()),
    )?;
    edit.update_exact(ExactTarget::Title, ExactAttribute::X, Some("3cm".into()))?;
    edit.update_exact(
        ExactTarget::Legend,
        ExactAttribute::LegendPosition,
        Some("start".into()),
    )?;
    edit.update_exact(
        ExactTarget::Legend,
        ExactAttribute::LegendAlign,
        Some("end".into()),
    )?;
    edit.update_exact(ExactTarget::Legend, ExactAttribute::Y, Some("2cm".into()))?;
    edit.update_exact(
        ExactTarget::Footer,
        ExactAttribute::StyleName,
        Some("footer-new".into()),
    )?;
    let changed = edit.commit()?;
    assert_eq!(changed.patch().exact_changes().len(), 7);
    let xml = changed.package().chart().content_xml();
    assert!(xml.contains("\n<office:body>"));
    assert!(xml.contains("table:cell-range=\"Data.B1\""));
    assert!(xml.contains("chart:legend-position=\"start\""));
    assert!(xml.contains("chart:legend-align=\"end\""));
    assert!(xml.contains("chart:style-name=\"footer-new\""));
    assert_eq!(
        changed
            .package()
            .apply(&changed.patch().inverse())?
            .into_package()
            .as_bytes(),
        source.as_bytes()
    );
    Ok(())
}

#[test]
fn embedded_specimen_plot_axis_grid_and_surface_semantics_use_exact_spans() -> TestResult<()> {
    // This vocabulary is exercised by checked-in LibreOffice FODS/FODT chart
    // fragments. The generated standalone package remains test input and is
    // not represented as producer-created ODC evidence.
    let mut authored = definition();
    authored.plot_area.data_source_labels = Some(Labels::Both);
    authored.plot_area.axes[0].categories_cell_range_address = Some("Data.A2:.A6".into());
    authored.plot_area.axes[0].grids.push(GridSpec {
        class: Class::Major,
        style_name: Some("grid-old".into()),
        extensions: Default::default(),
    });
    authored.plot_area.wall = Some(StyleElement {
        style_name: Some("wall-old".into()),
        ..StyleElement::default()
    });
    authored.plot_area.floor = Some(StyleElement {
        style_name: Some("floor-old".into()),
        ..StyleElement::default()
    });
    authored.legend = Some(LegendSpec {
        expansion: Some("high".into()),
        ..LegendSpec::default()
    });
    let canonical = litchi_odc::serialize_content(&authored)?;
    let noncanonical = canonical.replacen(
        "><office:body>",
        ">\n<!-- retained embedded-chart layout vocabulary -->\n<office:body>",
        1,
    );
    let source = Package::from_bytes(&raw_negative_package(&noncanonical)?)?;
    assert!(source.chart().definition().is_err());

    let mut edit = source.edit();
    edit.update_axis(0, litchi_odc::AxisUpdate::styled("axis-new"))?;
    edit.update_exact(
        ExactTarget::PlotArea,
        ExactAttribute::DataSourceHasLabels,
        Some("row".into()),
    )?;
    edit.update_exact(
        ExactTarget::Axis(0),
        ExactAttribute::Dimension,
        Some("z".into()),
    )?;
    edit.update_exact(
        ExactTarget::Categories { axis: 0 },
        ExactAttribute::CellRangeAddress,
        Some("Data.B2:.B6".into()),
    )?;
    edit.update_exact(
        ExactTarget::Grid { axis: 0, index: 0 },
        ExactAttribute::Class,
        Some("minor".into()),
    )?;
    edit.update_exact(
        ExactTarget::Grid { axis: 0, index: 0 },
        ExactAttribute::StyleName,
        Some("grid-new".into()),
    )?;
    edit.update_exact(
        ExactTarget::Legend,
        ExactAttribute::LegendExpansion,
        Some("wide".into()),
    )?;
    edit.update_exact(
        ExactTarget::Wall,
        ExactAttribute::StyleName,
        Some("wall-new".into()),
    )?;
    edit.update_exact(
        ExactTarget::Floor,
        ExactAttribute::StyleName,
        Some("floor-new".into()),
    )?;
    let changed = edit.commit()?;
    assert_eq!(changed.patch().changes().len(), 1);
    assert_eq!(changed.patch().exact_changes().len(), 8);
    let xml = changed.package().chart().content_xml();
    assert!(xml.contains("<!-- retained embedded-chart layout vocabulary -->"));
    assert!(xml.contains("chart:data-source-has-labels=\"row\""));
    assert!(xml.contains("chart:dimension=\"z\""));
    assert!(xml.contains("table:cell-range-address=\"Data.B2:.B6\""));
    assert!(xml.contains("chart:class=\"minor\" chart:style-name=\"grid-new\""));
    assert!(xml.contains("style:legend-expansion=\"wide\""));
    assert!(xml.contains("<chart:wall chart:style-name=\"wall-new\"/>"));
    assert!(xml.contains("<chart:floor chart:style-name=\"floor-new\"/>"));
    assert_eq!(
        changed
            .package()
            .apply(&changed.patch().inverse())?
            .into_package()
            .as_bytes(),
        source.as_bytes()
    );

    let wire = changed.patch().to_bytes();
    let decoded = Patch::from_bytes(&wire, source.chart().options())?;
    assert_eq!(decoded.exact_changes(), changed.patch().exact_changes());
    assert_eq!(
        source.apply(&decoded)?.into_package().as_bytes(),
        changed.package().as_bytes()
    );

    let mut independent_edit = source.edit();
    independent_edit.update_exact(
        ExactTarget::Series(0),
        ExactAttribute::StyleName,
        Some("independent-series".into()),
    )?;
    let independent = independent_edit.commit()?;
    let joined = changed.patch().join(independent.patch())?;
    assert!(joined.is_merged());
    let joined_chart = source
        .apply(
            joined
                .patch()
                .ok_or("missing embedded-vocabulary join patch")?,
        )?
        .into_package();
    assert!(
        joined_chart
            .chart()
            .content_xml()
            .contains("chart:dimension=\"z\"")
    );
    assert!(
        joined_chart
            .chart()
            .content_xml()
            .contains("chart:style-name=\"independent-series\"")
    );

    let destination_xml = noncanonical.replacen(
        "<office:body>",
        "<!-- independently retained destination -->\n<office:body>",
        1,
    );
    let destination = Package::from_bytes(&raw_negative_package(&destination_xml)?)?;
    let transfer = changed.patch().transfer_to(&destination)?;
    assert!(transfer.is_merged());
    let transferred = destination
        .apply(
            transfer
                .patch()
                .ok_or("missing embedded-vocabulary transfer patch")?,
        )?
        .into_package();
    assert!(
        transferred
            .chart()
            .content_xml()
            .contains("<!-- independently retained destination -->")
    );
    assert!(
        transferred
            .chart()
            .content_xml()
            .contains("chart:dimension=\"z\"")
    );
    assert!(
        transferred
            .chart()
            .content_xml()
            .contains("style:legend-expansion=\"wide\"")
    );
    Ok(())
}

#[test]
fn exact_enum_tokens_and_missing_namespace_insertions_fail_closed() -> TestResult<()> {
    let mut authored = definition();
    authored.plot_area.data_source_labels = Some(Labels::Both);
    authored.legend = Some(LegendSpec {
        expansion: Some("high".into()),
        ..LegendSpec::default()
    });
    let source = Package::from_bytes(&raw_negative_package(&litchi_odc::serialize_content(
        &authored,
    )?)?)?;

    for (target, attribute, value) in [
        (
            ExactTarget::PlotArea,
            ExactAttribute::DataSourceHasLabels,
            "rows",
        ),
        (
            ExactTarget::Axis(0),
            ExactAttribute::Dimension,
            "horizontal",
        ),
        (
            ExactTarget::Legend,
            ExactAttribute::LegendExpansion,
            "stretch",
        ),
        (
            ExactTarget::Grid { axis: 0, index: 0 },
            ExactAttribute::Class,
            "medium",
        ),
    ] {
        let mut edit = source.edit();
        assert!(
            edit.update_exact(target, attribute, Some(value.into()))
                .is_err()
        );
    }

    let mut without_expansion_definition = definition();
    without_expansion_definition.legend = Some(LegendSpec::default());
    let without_expansion = Builder::new()
        .with_definition(without_expansion_definition)
        .build()?;
    let mut edit = without_expansion.edit();
    edit.update_exact(
        ExactTarget::Legend,
        ExactAttribute::LegendExpansion,
        Some("wide".into()),
    )?;
    assert!(edit.commit().is_err());
    Ok(())
}

#[test]
fn definition_join_transfer_and_dependency_conflicts_are_deterministic() -> TestResult<()> {
    let source = Builder::new().with_definition(definition()).build()?;

    let mut left_edit = source.edit();
    left_edit
        .definition_edit()?
        .set_style(StyleTarget::Axis(0), Some("left".into()))?;
    let left = left_edit.commit()?;
    let mut right_edit = source.edit();
    right_edit
        .definition_edit()?
        .set_style(StyleTarget::Axis(1), Some("right".into()))?;
    let right = right_edit.commit()?;
    let joined = left.patch().join(right.patch())?;
    assert!(joined.is_merged());
    let joined_target = source
        .apply(joined.patch().ok_or("missing joined patch")?)?
        .into_package();
    let joined_definition = joined_target.chart().definition()?;
    assert_eq!(
        joined_definition.plot_area.axes[0].style_name.as_deref(),
        Some("left")
    );
    assert_eq!(
        joined_definition.plot_area.axes[1].style_name.as_deref(),
        Some("right")
    );

    let mut conflict_edit = source.edit();
    conflict_edit
        .definition_edit()?
        .set_style(StyleTarget::Axis(0), Some("other".into()))?;
    let conflict_commit = conflict_edit.commit()?;
    let conflict = left.patch().join(conflict_commit.patch())?;
    assert!(!conflict.is_merged());
    assert!(!conflict.conflicts().is_empty());

    let mut rename_edit = source.edit();
    let definition_edit = rename_edit.definition_edit()?;
    let mut renamed_axis = definition_edit.definition().plot_area.axes[0].clone();
    renamed_axis.name = Some("renamed".into());
    definition_edit.update_axis(0, renamed_axis)?;
    let mut renamed_series = definition_edit.definition().plot_area.series[0].clone();
    renamed_series.attached_axis = Some("renamed".into());
    definition_edit.update_series(0, renamed_series)?;
    let rename = rename_edit.commit()?;

    let mut destination_definition = definition();
    destination_definition.title = Some(Text::new("destination"));
    let destination = Builder::new()
        .with_definition(destination_definition)
        .build()?;
    let transferred = rename.patch().transfer_to(&destination)?;
    assert!(transferred.is_merged());
    let transferred_target = destination
        .apply(transferred.patch().ok_or("missing transferred patch")?)?
        .into_package();
    let transferred_definition = transferred_target.chart().definition()?;
    assert_eq!(
        transferred_definition.plot_area.axes[0].name.as_deref(),
        Some("renamed")
    );
    assert_eq!(
        transferred_definition
            .title
            .as_ref()
            .map(|text| text.text.as_str()),
        Some("destination")
    );

    let mut dependency_definition = definition();
    dependency_definition.plot_area.series.clear();
    let dependency_source = Builder::new()
        .with_definition(dependency_definition.clone())
        .build()?;
    let mut series_edit = dependency_source.edit();
    series_edit.definition_edit()?.insert_series(
        0,
        SeriesSpec {
            attached_axis: Some("x".into()),
            ..SeriesSpec::default()
        },
    )?;
    let series_commit = series_edit.commit()?;
    let mut no_axes = dependency_definition;
    no_axes.plot_area.axes.clear();
    let no_axes_package = Builder::new().with_definition(no_axes).build()?;
    let dependency_conflict = series_commit.patch().transfer_to(&no_axes_package)?;
    assert!(!dependency_conflict.is_merged());
    assert!(!dependency_conflict.conflicts().is_empty());
    Ok(())
}

#[test]
fn durable_package_patches_join_and_history_are_exact() -> TestResult<()> {
    let source = Builder::new().with_definition(definition()).build()?;
    let mut left_edit = source.edit();
    left_edit.set_style(StyleTarget::Chart, Some("left".into()))?;
    let left = left_edit.commit()?;
    let mut right_edit = source.edit();
    right_edit.add_resource("Pictures/value.bin", "application/octet-stream", b"value")?;
    let right = right_edit.commit()?;

    let merge = left.patch().join(right.patch())?;
    assert!(merge.is_merged());
    let merged_patch = merge.patch().ok_or("missing package patch")?;
    let merged = source.apply(merged_patch)?.into_package();
    assert_eq!(merged.chart().resources().len(), 1);
    assert_eq!(
        merged.chart().definition()?.style_name.as_deref(),
        Some("left")
    );

    let wire = merged_patch.to_bytes();
    assert_eq!(wire, merged_patch.to_bytes());
    let decoded = Patch::from_bytes(&wire, source.chart().options())?;
    assert_eq!(
        source.apply(&decoded)?.into_package().as_bytes(),
        merged.as_bytes()
    );
    assert_eq!(
        merged.apply(&decoded.inverse())?.into_package().as_bytes(),
        source.as_bytes()
    );
    let mut trailing = wire.clone();
    trailing.push(0);
    assert!(Patch::from_bytes(&trailing, source.chart().options()).is_err());

    let limits =
        OpenOptions::new().semantic_limits(SemanticLimits::builder().max_history(1)?.build()?);
    let bounded_source = Builder::new()
        .with_options(limits)
        .with_definition(definition())
        .build()?;
    let mut first_edit = bounded_source.edit();
    first_edit.set_style(StyleTarget::Chart, Some("one".into()))?;
    let first = first_edit.commit()?;
    let mut second_edit = first.package().edit();
    second_edit.set_style(StyleTarget::Chart, Some("two".into()))?;
    let second = second_edit.commit()?;
    let mut history = History::new(bounded_source.clone());
    history.record(&first)?;
    assert!(history.record(&second).is_err());
    assert!(history.undo()?);
    assert_eq!(history.current().as_bytes(), bounded_source.as_bytes());
    assert!(history.redo()?);
    assert_eq!(history.current().as_bytes(), first.package().as_bytes());
    Ok(())
}

#[test]
fn package_transfer_closes_chart_style_data_and_resource_dependencies() -> TestResult<()> {
    let source = Builder::new().with_definition(definition()).build()?;
    let mut changed_edit = source.edit();
    changed_edit.set_style(StyleTarget::Chart, Some("transferred".into()))?;
    let mut table = CachedTable::new("Data", 1);
    table
        .rows
        .push(CachedRow::new(vec![CachedCell::new(CachedValue::String(
            "value".into(),
        ))]));
    changed_edit.set_cached_table(Some(table))?;
    changed_edit.set_styles_xml(STYLES);
    changed_edit.add_resource(
        "Pictures/transferred.bin",
        "application/octet-stream",
        b"transferred",
    )?;
    let changed = changed_edit.commit()?;

    let mut destination_definition = definition();
    destination_definition.title = Some(Text::new("destination"));
    let plain_destination = Builder::new()
        .with_definition(destination_definition)
        .build()?;
    let mut destination_edit = plain_destination.edit();
    destination_edit.add_resource(
        "Pictures/existing.bin",
        "application/octet-stream",
        b"existing",
    )?;
    let destination = destination_edit.commit()?.into_package();

    let transfer = changed.patch().transfer_to(&destination)?;
    assert!(transfer.is_merged());
    let transferred = destination
        .apply(
            transfer
                .patch()
                .ok_or("missing transferred package patch")?,
        )?
        .into_package();
    let reopened = Package::from_bytes(transferred.as_bytes())?;
    let projected = reopened.chart().definition()?;
    assert_eq!(projected.style_name.as_deref(), Some("transferred"));
    assert_eq!(
        projected.title.as_ref().map(|value| value.text.as_str()),
        Some("destination")
    );
    assert_eq!(
        projected
            .cached_table
            .as_ref()
            .map(|value| value.name.as_str()),
        Some("Data")
    );
    assert_eq!(reopened.chart().styles_xml(), Some(STYLES));
    assert_eq!(reopened.chart().resources().len(), 2);

    let mut conflict_edit = plain_destination.edit();
    conflict_edit.add_resource(
        "Pictures/transferred.bin",
        "application/octet-stream",
        b"destination",
    )?;
    let conflict_destination = conflict_edit.commit()?.into_package();
    let conflict = changed.patch().transfer_to(&conflict_destination)?;
    assert_eq!(
        conflict.conflicts()[0].path(),
        "package.resource[Pictures/transferred.bin]"
    );
    Ok(())
}

#[test]
fn row_column_ranges_styles_and_security_policy_are_enforced() -> TestResult<()> {
    validate_range_list("Data.$A:.$C Data.$1:.$9")?;
    assert!(validate_range_list("Data.$A:.$9").is_err());
    assert!(validate_range_list("Data.$A").is_err());
    assert!(validate_range_list("Data.$1").is_err());

    let source = Builder::new().with_definition(definition()).build()?;
    let mut wrong_root = source.edit();
    wrong_root.set_styles_xml("<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>");
    assert!(wrong_root.commit().is_err());
    let mut scripts = source.edit();
    scripts.set_styles_xml(r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:scripts/></office:document-styles>"#);
    assert!(scripts.commit().is_err());
    let mut missing_family = source.edit();
    missing_family.set_styles_xml(r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles><style:style style:name="dangling"/></office:styles></office:document-styles>"#);
    assert!(missing_family.commit().is_err());
    let mut valid = source.edit();
    valid.set_styles_xml(STYLES);
    assert_eq!(valid.commit()?.package().chart().styles_xml(), Some(STYLES));
    let mut noncompact_resource = source.edit();
    noncompact_resource.add_resource(
        "Object.xml",
        "application/xml",
        b"<root>\n<child/></root>",
    )?;
    assert!(noncompact_resource.commit().is_err());
    assert!(
        Builder::new()
            .with_resource("Object.xml", "application/xml", b"<root>\n<child/></root>")
            .build()
            .is_err()
    );

    let content = source.chart().content_xml();
    let signed = Package::from_bytes(&package(
        content,
        Some(("META-INF/vendorsignatures.xml", b"<signatures/>")),
    )?)?;
    assert!(signed.chart().is_signed());
    let mut signed_edit = signed.edit();
    signed_edit.set_style(StyleTarget::Chart, Some("blocked".into()))?;
    assert!(signed_edit.commit().is_err());

    let encrypted = Package::from_bytes(&encrypted_envelope(content)?)?;
    assert!(encrypted.chart().is_encrypted());
    let mut encrypted_edit = encrypted.edit();
    encrypted_edit.set_style(StyleTarget::Chart, Some("blocked".into()))?;
    assert!(encrypted_edit.commit().is_err());

    let mut repeated_data = definition();
    let mut repeated_table = CachedTable::new("Data", 1);
    let mut repeated_row = CachedRow::new(vec![CachedCell::new(CachedValue::Boolean(true))]);
    repeated_row.repeated = 2;
    repeated_table.rows.push(repeated_row);
    repeated_data.cached_table = Some(repeated_table);
    let malformed_boolean = litchi_odc::serialize_content(&repeated_data)?.replace(
        "office:boolean-value=\"true\"",
        "office:boolean-value=\"maybe\"",
    );
    assert!(Package::from_bytes(&raw_negative_package(&malformed_boolean)?).is_err());
    let row_limits =
        OpenOptions::new().semantic_limits(SemanticLimits::builder().max_cached_rows(1)?.build()?);
    assert!(
        Builder::new()
            .with_options(row_limits)
            .with_definition(repeated_data)
            .build()
            .is_err()
    );

    let duplicate_plot = source.chart().content_xml().replacen(
        "</chart:plot-area>",
        "</chart:plot-area><chart:plot-area/>",
        1,
    );
    assert!(Package::from_bytes(&raw_negative_package(&duplicate_plot)?).is_err());

    let mut too_many_ranges = definition();
    too_many_ranges.plot_area.cell_range_address = Some("Data.A1 Data.B1".into());
    let range_limits =
        OpenOptions::new().semantic_limits(SemanticLimits::builder().max_range_items(1)?.build()?);
    assert!(
        Builder::new()
            .with_options(range_limits)
            .with_definition(too_many_ranges)
            .build()
            .is_err()
    );
    Ok(())
}
