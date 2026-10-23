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
    let mut writer = PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
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
    let commit = presentation.chart_snapshot()?.edit().commit()?;
    assert_eq!(commit.snapshot().bytes(), bytes.as_slice());
    Ok(())
}

#[test]
fn inline_replace_preserves_unknown_chart_xml() -> Result<()> {
    let bytes = package(true)?;
    let presentation = open_presentation(bytes)?;
    let mut edit = presentation.chart_snapshot()?.edit();
    let replacement = Part::from_xml(CHART.replace("kept", "still-kept"))?;
    edit.replace("Chart A", replacement)?;
    let commit = edit.commit()?;
    let reparsed = open_presentation(commit.snapshot().bytes().to_vec())?;
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
    let mut edit = presentation.chart_snapshot()?.edit();
    edit.remove("Chart A")?;
    edit.add(
        0usize,
        "Chart B",
        Storage::InlineXml,
        Part::from_xml(CHART)?,
    )?;
    let commit = edit.commit()?;
    let output = commit.snapshot().bytes().to_vec();
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
    let mut edit = presentation.chart_snapshot_with(limits)?.edit();
    let error = edit.replace(
        "Chart A",
        Part::from_xml(CHART.replace("kept", "too-long"))?,
    );
    assert!(error.is_err());
    Ok(())
}

#[test]
fn owned_snapshot_commit_rehydrates_the_presentation() -> Result<()> {
    let source_bytes = package(false)?;
    let presentation = open_presentation(source_bytes.clone())?;
    let source = presentation.chart_snapshot()?;
    let mut edit = source.edit();
    edit.replace(
        "Chart A",
        Part::from_xml(CHART.replace("kept", "snapshot-kept"))?,
    )?;
    let commit = edit.commit()?;

    assert_ne!(commit.snapshot().bytes(), source.bytes());
    assert_eq!(commit.snapshot().charts().len(), 1);
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

    Ok(())
}

#[test]
fn owned_snapshot_add_remove_and_noop_have_verified_readback() -> Result<()> {
    let presentation = open_presentation(package(false)?)?;
    let source = presentation.chart_snapshot()?;
    let noop = source.edit().commit()?;
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

/// Record 0358a: the removed inventory transaction's chart edits on the
/// package route, with a durable, reversible, mergeable patch.
#[test]
fn package_chart_edits_have_durable_patches() -> Result<()> {
    let bytes = package(false)?;
    let package = crate::Package::from_bytes(&bytes)
        .map_err(|error| litchi_core::Error::InvalidFormat(error.to_string()))?;
    let mut edit = package
        .edit()
        .map_err(|error| litchi_core::Error::InvalidFormat(error.to_string()))?;
    edit.replace_chart("Chart A", Part::from_xml(CHART.replace("kept", "durable"))?)?;
    let commit = edit
        .commit()
        .map_err(|error| litchi_core::Error::InvalidFormat(error.to_string()))?;
    let wire = commit.patch().to_durable_bytes()?;
    let patch = crate::edit::Patch::from_durable_bytes(&wire)?;
    let replayed = package
        .apply(&patch)
        .map_err(|error| litchi_core::Error::InvalidFormat(error.to_string()))?;
    assert_eq!(replayed.package().as_bytes(), commit.package().as_bytes());
    let restored = commit
        .package()
        .apply(&patch.inverse())
        .map_err(|error| litchi_core::Error::InvalidFormat(error.to_string()))?;
    assert_eq!(restored.package().as_bytes(), bytes.as_slice());
    Ok(())
}
