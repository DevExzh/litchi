use super::{Part, Storage};
use crate::{Presentation, core::OwnedPackage};
use litchi_core::Result;
use litchi_odf_common::constants::{ODF_CHART, ODF_PRESENTATION};
use litchi_odf_common::core::PackageWriter;

const CHART: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:foo="urn:example:future"><office:body><office:chart><chart:chart chart:class="chart:bar"><foo:future foo:value="kept"/></chart:chart></office:chart></office:body></office:document-content>"#;

fn content(object: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Slide 1"><draw:frame draw:name="Chart A">{object}</draw:frame></draw:page></office:presentation></office:body></office:document-content>"#
    )
}

fn package(inline: bool) -> Result<Vec<u8>> {
    let object = if inline {
        format!(
            "<draw:object>{}</draw:object>",
            super::codec::content_inline(CHART)?
        )
    } else {
        "<draw:object xlink:href=\"./Object_1\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>".to_string()
    };
    let mut writer = PackageWriter::new();
    writer.set_mimetype(ODF_PRESENTATION)?;
    writer.add_file("content.xml", content(&object).as_bytes())?;
    if !inline {
        writer.add_manifest_directory("Object_1/", ODF_CHART)?;
        writer.add_file_with_media_type("Object_1/content.xml", CHART.as_bytes(), "text/xml")?;
    }
    writer.finish_to_bytes()
}

fn open_presentation(bytes: Vec<u8>) -> Result<Presentation> {
    Presentation::from_bytes(bytes)
}

#[test]
fn discovers_typed_part_and_selector_context() -> Result<()> {
    let bytes = package(false)?;
    let presentation = open_presentation(bytes)?;
    let inventory = presentation.charts()?;
    assert_eq!(inventory.len(), 1);
    let chart = inventory.named("Chart A")?.ok_or_else(|| {
        litchi_core::Error::InvalidFormat("chart fixture was not discovered".to_string())
    })?;
    assert_eq!(chart.page(), Some("Slide 1"));
    assert_eq!(chart.storage(), Storage::PackageSubdocument);
    assert_eq!(chart.part().chart().children().len(), 1);
    assert!(chart.part().xml().contains("foo:future"));
    Ok(())
}

#[test]
fn no_op_commit_returns_exact_archive_bytes() -> Result<()> {
    let bytes = package(false)?;
    let presentation = open_presentation(bytes.clone())?;
    let inventory = presentation.charts()?;
    let commit = inventory.transaction().commit()?;
    assert!(!commit.changed());
    assert_eq!(commit.bytes(), bytes.as_slice());
    Ok(())
}

#[test]
fn inline_replace_preserves_unknown_chart_xml() -> Result<()> {
    let bytes = package(true)?;
    let presentation = open_presentation(bytes)?;
    let inventory = presentation.charts()?;
    let mut transaction = inventory.transaction();
    let replacement = Part::from_xml(CHART.replace("kept", "still-kept"))?;
    transaction.replace("Chart A", replacement)?;
    let commit = transaction.commit()?;
    assert!(commit.changed());
    let reparsed = open_presentation(commit.into_owned_bytes())?;
    let charts = reparsed.charts()?;
    let chart = charts.at(0)?.ok_or_else(|| {
        litchi_core::Error::InvalidFormat("replaced chart disappeared".to_string())
    })?;
    assert!(chart.part().xml().contains("still-kept"));
    assert!(chart.part().xml().contains("foo:future"));
    Ok(())
}

#[test]
fn remove_then_add_rebuilds_only_permitted_package_parts() -> Result<()> {
    let bytes = package(false)?;
    let presentation = open_presentation(bytes)?;
    let inventory = presentation.charts()?;
    let mut transaction = inventory.transaction();
    transaction.remove("Chart A")?;
    transaction.add(
        0usize,
        "Chart B",
        Storage::InlineXml,
        Part::from_xml(CHART)?,
    )?;
    let commit = transaction.commit()?;
    assert!(commit.changed());
    let output = commit.into_owned_bytes();
    let archive = OwnedPackage::from_bytes(output.clone())?;
    assert!(!archive.has_file("Object_1/content.xml")?);
    let reparsed = open_presentation(output)?;
    let charts = reparsed.charts()?;
    assert_eq!(charts.len(), 1);
    assert_eq!(
        charts.at(0)?.and_then(|chart| chart.name()),
        Some("Chart B")
    );
    Ok(())
}

#[test]
fn limits_reject_oversized_parts_before_staging() -> Result<()> {
    let bytes = package(false)?;
    let presentation = open_presentation(bytes)?;
    let limits = super::Limits::new(1, CHART.len(), CHART.len())?;
    let inventory = presentation.charts_with(limits)?;
    let mut transaction = inventory.transaction();
    let error = transaction.replace("Chart A", Part::from_xml(format!("{CHART} "))?);
    assert!(error.is_err());
    Ok(())
}
