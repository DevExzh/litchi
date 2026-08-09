use super::{Part, Snapshot};
use litchi_core::Result;
use litchi_odf_common::core::PackageWriter;

const CONTENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:spreadsheet><table:table table:name="Data"><table:shapes><draw:frame draw:name="Sales"><draw:object xlink:href="./Object_1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document-content>"#;
const CHART: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#;
const REPLACEMENT: &str = r#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart chart:class="chart:line"><chart:plot-area/></chart:chart></office:chart></office:body></office:document-content>"#;

fn package() -> Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype("application/vnd.oasis.opendocument.spreadsheet")?;
    writer.add_file("content.xml", CONTENT.as_bytes())?;
    writer.add_manifest_directory("Object_1/", "application/vnd.oasis.opendocument.chart")?;
    writer.add_file_with_media_type("Object_1/content.xml", CHART.as_bytes(), "text/xml")?;
    writer.finish_to_bytes()
}

#[test]
fn owned_snapshot_patch_is_source_checked_reversible_and_rehydrates() -> Result<()> {
    let snapshot = Snapshot::from_bytes(package()?)?;
    assert_eq!(snapshot.charts().len(), 1);
    assert_eq!(
        snapshot
            .get("Sales")?
            .ok_or_else(|| litchi_core::Error::InvalidFormat("missing Sales chart".to_string()))?
            .name(),
        Some("Sales")
    );

    let mut edit = snapshot.edit();
    edit.replace("Sales", Part::from_xml(REPLACEMENT)?)?;
    let commit = edit.commit()?;
    assert!(commit.changed());
    assert!(
        commit.snapshot().charts()[0]
            .content_xml()
            .contains("chart:line")
    );

    let restored = commit.patch().inverse().apply(commit.snapshot())?;
    assert_eq!(restored.snapshot().as_bytes(), snapshot.as_bytes());
    let applied = commit.patch().apply(&snapshot)?;
    assert!(applied.changed());
    assert!(commit.patch().apply(commit.snapshot()).is_err());
    Ok(())
}
