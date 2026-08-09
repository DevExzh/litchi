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
    let error = transaction.replace(
        "Chart A",
        Part::from_xml(CHART.replace("kept", "too-long"))?,
    );
    assert!(error.is_err());
    Ok(())
}

#[test]
fn owned_snapshot_commit_and_patch_rehydrate_the_presentation() -> Result<()> {
    let source_bytes = package(false)?;
    let presentation = open_presentation(source_bytes.clone())?;
    let source = presentation.chart_snapshot()?;
    let mut edit = source.edit();
    edit.replace(
        "Chart A",
        Part::from_xml(CHART.replace("kept", "snapshot-kept"))?,
    )?;
    let commit = edit.commit()?;

    assert!(commit.changed());
    assert_eq!(commit.diagnostics().charts_before(), 1);
    assert_eq!(commit.diagnostics().charts_after(), 1);
    assert!(commit.diagnostics().changed());
    assert_eq!(source.bytes(), source_bytes.as_slice());
    assert!(
        commit
            .snapshot()
            .get("Chart A")?
            .is_some_and(|chart| chart.part().xml().contains("snapshot-kept"))
    );
    let reopened = commit.snapshot().to_presentation()?;
    assert!(
        reopened
            .chart("Chart A")?
            .is_some_and(|chart| chart.part().xml().contains("snapshot-kept"))
    );

    let applied = commit.patch().apply(&source)?;
    assert_eq!(applied.snapshot().bytes(), commit.snapshot().bytes());
    let restored = commit.patch().inverse().apply(applied.snapshot())?;
    assert_eq!(restored.snapshot().bytes(), source.bytes());
    assert!(commit.patch().is_applicable_to(restored.snapshot()));
    assert!(commit.patch().apply(restored.snapshot()).is_ok());
    let stale = open_presentation(package(true)?)?.chart_snapshot()?;
    assert!(commit.patch().apply(&stale).is_err());
    Ok(())
}

#[test]
fn owned_snapshot_add_remove_and_noop_have_verified_readback() -> Result<()> {
    let presentation = open_presentation(package(false)?)?;
    let source = presentation.chart_snapshot()?;
    let noop = source.edit().commit()?;
    assert!(!noop.changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.snapshot().bytes(), source.bytes());

    let mut edit = source.edit();
    edit.remove("Chart A")?;
    edit.add(
        "Slide 1",
        "Chart B",
        Storage::InlineXml,
        Part::from_xml(CHART)?,
    )?;
    let commit = edit.commit()?;
    let chart = commit
        .snapshot()
        .get("Chart B")?
        .ok_or_else(|| litchi_core::Error::InvalidFormat("added chart disappeared".to_string()))?;
    assert_eq!(chart.storage(), Storage::InlineXml);
    assert!(commit.snapshot().get("Chart A")?.is_none());
    Ok(())
}

#[test]
fn authored_chart_parts_must_be_compact_xml() {
    let formatted = CHART.replace("><", ">\n  <");
    assert!(Part::from_xml(formatted).is_err());
    let spaced = CHART.replacen("><", "> <", 1);
    assert!(Part::from_xml(spaced).is_err());
}
