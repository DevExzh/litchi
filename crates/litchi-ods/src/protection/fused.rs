//! Fused single-tokenization protection parse for ODS `content.xml`.
//!
//! [`super::codec::parse`] historically tokenized the same `content.xml`
//! twice: once for the source locator ([`super::codec::Location`]) and once
//! for the semantic protection metadata ([`crate::model::protection`]).
//! This module runs both passes over one shared `NsReader` event stream.
//!
//! Observable behavior is identical to the sequential passes: the locator
//! size limit runs up front, each handler keeps its own error messages, a
//! quick_xml read failure is recorded for the first still-active handler in
//! pass order (both passes historically used the same
//! `"XML parsing error: ..."` mapping), and the final error selection follows
//! the original call order — the source locator ran to completion before the
//! semantic protection parse started.  The caller keeps the sheet-count
//! disagreement check.

use litchi_core::{Error, Result};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::codec::{self, Location, LocationHandler};
use crate::model::protection::{self, Protection, ProtectionHandler, Sheet};

/// Parse the `content.xml` protection snapshot and semantic metadata with a
/// single tokenization.
///
/// This is the fused equivalent of running
/// [`super::codec::Location::parse`] and then
/// [`crate::model::protection::parse_protection`] sequentially: the returned
/// triple, every error message, and the error precedence are unchanged.
pub(crate) fn parse(
    source: &str,
    styles_xml: Option<&str>,
) -> Result<(Location, Protection, Vec<Sheet>)> {
    // The locator is the first pass of the sequential parse, so its size
    // limit runs up front; the semantic pass has no size limit of its own.
    codec::validate_size(source)?;

    let mut reader = NsReader::from_str(source);
    let mut buffer = Vec::new();
    let mut location = LocationHandler::new(source, styles_xml);
    let mut protection = ProtectionHandler::default();
    let mut location_error = None;
    let mut protection_error = None;

    loop {
        let pos_before = reader.buffer_position();
        let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
            Ok(resolved) => resolved,
            Err(error) => {
                // A broken stream stops both passes.  Standalone, the locator
                // reports the failure with its mapping and the semantic pass
                // never runs, so record the error for the first still-active
                // handler in pass order; both mappings are the same string,
                // so only the selection position matters.
                if location_error.is_none() {
                    location_error =
                        Some(Error::InvalidFormat(format!("XML parsing error: {error}")));
                } else if protection_error.is_none() {
                    protection_error =
                        Some(Error::InvalidFormat(format!("XML parsing error: {error}")));
                }
                break;
            },
        };
        // The resolved namespace borrows the reader mutably, so classify it
        // for both passes before touching the reader again, exactly where
        // each historical loop body classified it.
        let location_namespace = codec::location_namespace(&namespace);
        let protection_namespace = protection::classify(&namespace);
        let pos_after = reader.buffer_position();
        let is_eof = matches!(event, Event::Eof);
        if location_error.is_none()
            && let Err(error) = location.on_event(
                location_namespace,
                &event,
                reader.resolver(),
                reader.decoder(),
                pos_before,
                pos_after,
            )
        {
            location_error = Some(error);
        }
        if protection_error.is_none()
            && let Err(error) = protection.on_event(
                protection_namespace,
                &event,
                reader.resolver(),
                reader.decoder(),
            )
        {
            protection_error = Some(error);
        }
        buffer.clear();
        if is_eof {
            break;
        }
    }

    // Final selection in the original call order: the locator ran to
    // completion (including its end-of-stream checks) before the semantic
    // protection parse historically started.
    if let Some(error) = location_error {
        return Err(error);
    }
    let location = location.finish()?;
    if let Some(error) = protection_error {
        return Err(error);
    }
    let (document, sheets) = protection.finish()?;
    Ok((location, document, sheets))
}

#[cfg(test)]
mod tests {
    use super::parse;
    use crate::model::protection::{Protection, Sheet, parse_protection};
    use crate::protection::codec::Location;
    use litchi_core::{Error, Result};
    use std::io::Read as _;
    use std::path::{Path, PathBuf};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const LOEXT: &str = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    const OFFICE_EXT: &str = "http://openoffice.org/2009/office";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const DISAGREE: &str = "ODS protection sheet parser and source locator disagree";

    const STYLES: &str = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#;

    type Compared = (String, Protection, Vec<Sheet>);

    /// The original sequential passes with their exact error precedence.
    fn sequential(source: &str, styles_xml: Option<&str>) -> Result<Compared> {
        let location = Location::parse(source, styles_xml)?;
        let (document, sheets) = parse_protection(source)?;
        if sheets.len() != location.sheets().len() {
            return Err(Error::InvalidFormat(DISAGREE.to_string()));
        }
        Ok((format!("{location:?}"), document, sheets))
    }

    /// The fused parse with the same comparison projection.
    fn fused(source: &str, styles_xml: Option<&str>) -> Result<Compared> {
        let (location, document, sheets) = parse(source, styles_xml)?;
        if sheets.len() != location.sheets().len() {
            return Err(Error::InvalidFormat(DISAGREE.to_string()));
        }
        Ok((format!("{location:?}"), document, sheets))
    }

    fn assert_equivalent(label: &str, source: &str, styles_xml: Option<&str>) {
        let expected = sequential(source, styles_xml).map_err(|error| error.to_string());
        let actual = fused(source, styles_xml).map_err(|error| error.to_string());
        assert_eq!(expected, actual, "{label}: fused and sequential disagree");
    }

    fn document(body: &str) -> String {
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:loext="{LOEXT}" xmlns:office-ext="{OFFICE_EXT}" xmlns:style="{STYLE}"><office:body><office:spreadsheet>{body}</office:spreadsheet></office:body></office:document-content>"#
        )
    }

    fn corpus_files() -> Vec<PathBuf> {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        for base in [
            "test-data/odf",
            "test-data/odfdo",
            "test-data/odfpy",
            "test-data/libreoffice-core/sc/qa/unit/data/ods",
            "test-data/libreoffice-core/sc/qa/extras/testdocuments",
        ] {
            collect_ods(&root.join(base), &mut files);
        }
        files.sort();
        files
    }

    fn collect_ods(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_ods(&path, files);
            } else if path.extension().is_some_and(|extension| extension == "ods") {
                files.push(path);
            }
        }
    }

    fn package_member(path: &Path, name: &str) -> Option<String> {
        let bytes = std::fs::read(path).ok()?;
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).ok()?;
        let mut entry = archive.by_name(name).ok()?;
        let mut content = String::new();
        entry.read_to_string(&mut content).ok()?;
        Some(content)
    }

    #[test]
    fn fused_parse_matches_sequential_passes_on_ods_corpus() {
        let files = corpus_files();
        assert!(!files.is_empty(), "no .ods corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(content) = package_member(path, "content.xml") else {
                continue;
            };
            let styles = package_member(path, "styles.xml");
            assert_equivalent(&path.display().to_string(), &content, styles.as_deref());
            compared += 1;
        }
        assert!(compared > 0, "no .ods corpus fixtures yielded content.xml");
    }

    #[test]
    fn fused_parse_matches_sequential_passes_on_minimal_document() {
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:loext="{LOEXT}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="ce1" style:family="table-cell"/></office:automatic-styles><office:body><office:spreadsheet table:structure-protected="true" table:protection-key="key" table:protection-key-digest-algorithm="urn:sha256"><table:table table:name="Sheet1" table:protected="true" table:protection-key="abc" loext:protection-key-digest-algorithm-2="urn:sha1"><loext:table-protection loext:select-protected-cells="false" loext:insert-rows="true"/></table:table><table:table table:name="Sheet2"/></office:spreadsheet></office:body></office:document-content>"#
        );
        let (location, document, sheets) =
            parse(&xml, Some(STYLES)).expect("minimal document parses");
        assert_eq!(document.structure_protected, Some(true));
        assert_eq!(document.key.digest_algorithm.as_deref(), Some("urn:sha256"));
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].protected, Some(true));
        assert_eq!(
            sheets[0].key.secondary_digest_algorithm.as_deref(),
            Some("urn:sha1")
        );
        assert_eq!(sheets[0].options.select_protected_cells, Some(false));
        assert_eq!(sheets[0].options.insert_rows, Some(true));
        assert_eq!(sheets[1].protected, None);
        assert_eq!(location.sheets()[0].name, "Sheet1");
        assert_eq!(location.sheets()[1].name, "Sheet2");
        assert!(location.automatic_range().is_some());
        assert_eq!(location.styles_xml(), Some(STYLES));
        assert_equivalent("minimal", &xml, Some(STYLES));
        assert_equivalent("minimal-without-styles", &xml, None);
    }

    #[test]
    fn custom_table_and_loext_prefixes() {
        let xml = format!(
            r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:l="{LOEXT}"><o:body><o:spreadsheet><t:table t:name="S" t:protected="1"><l:table-protection l:use-pivot="true"/></t:table></o:spreadsheet></o:body></o:document-content>"#
        );
        let (location, _, sheets) = parse(&xml, None).expect("prefixed document parses");
        assert_eq!(location.table_prefix(), "t");
        assert_eq!(location.loext_prefix(), Some("l"));
        assert_eq!(sheets[0].protected, Some(true));
        assert_eq!(sheets[0].options.use_pivot, Some(true));
        assert_equivalent("custom-prefixes", &xml, None);
    }

    #[test]
    fn mixed_sheets_and_ignored_constructs() {
        let xml = document(concat!(
            r#"<table:table table:name="Empty"/>"#,
            r#"<table:table table:name="A&amp;B" table:protected="false"><table:table-row/></table:table>"#,
            r#"<table:table table:name="Ext"><office-ext:table-protection office-ext:use-autofilter="true"/></table:table>"#,
            r#"<table:table table:name="Tab"><table:table-protection table:delete-rows="false"/></table:table>"#,
            r#"<table:table table:name="Outer"><table:table table:name="Inner"/></table:table>"#,
            r#"<loext:table-protection loext:use-pivot="true"/>"#,
            r#"<table:dde-links><table:dde-link><table:table table:name="Cache" table:protected="true"/></table:dde-link></table:dde-links>"#,
        ));
        let (location, _, sheets) = parse(&xml, None).expect("mixed document parses");
        let names = location
            .sheets()
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<Vec<_>>();
        assert_eq!(names, ["Empty", "A&B", "Ext", "Tab", "Outer"]);
        assert_eq!(sheets.len(), 5);
        assert_eq!(sheets[2].options.use_auto_filter, Some(true));
        assert_eq!(sheets[3].options.delete_rows, Some(false));
        assert_equivalent("mixed", &xml, None);
    }

    #[test]
    fn location_error_beats_protection_error() {
        // The sheet misses table:name (locator failure) and carries an
        // invalid Boolean (semantic failure) at the same event; the locator
        // pass runs first.
        let xml = document(r#"<table:table table:protected="yes"></table:table>"#);
        let located = Location::parse(&xml, None).expect_err("locator must fail");
        assert_eq!(
            located.to_string(),
            "Invalid format: ODS protected table is missing table:name"
        );
        let wire = parse_protection(&xml).expect_err("semantic pass must fail");
        assert_eq!(
            wire.to_string(),
            "Invalid format: invalid protection Boolean value 'yes'"
        );
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(fused.to_string(), located.to_string());
        assert_equivalent("location-beats-protection", &xml, None);
    }

    #[test]
    fn protection_error_surfaces_when_locator_passes() {
        let xml = document(r#"<table:table table:name="S" table:protected="yes"/>"#);
        Location::parse(&xml, None).expect("locator accepts the sheet");
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: invalid protection Boolean value 'yes'"
        );
        assert_equivalent("protection-only-error", &xml, None);
    }

    #[test]
    fn duplicate_table_protection_prefers_the_locator_mapping() {
        // The second table-protection element fails both passes with
        // different messages; the locator pass runs first.
        let xml = document(
            r#"<table:table table:name="S"><loext:table-protection loext:use-pivot="true"/><loext:table-protection loext:use-pivot="false"/></table:table>"#,
        );
        let wire = parse_protection(&xml).expect_err("semantic pass must fail");
        assert_eq!(
            wire.to_string(),
            "Invalid format: duplicate sheet table-protection element"
        );
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: duplicate ODS table-protection element"
        );
        assert_equivalent("duplicate-table-protection", &xml, None);
    }

    #[test]
    fn unterminated_document_prefers_the_locator_check() {
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:table table:name="S">"#
        );
        let wire = parse_protection(&xml).expect_err("semantic pass must fail");
        assert_eq!(
            wire.to_string(),
            "Invalid format: unterminated protected sheet"
        );
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: unterminated ODS protection XML element"
        );
        assert_equivalent("unterminated", &xml, None);
    }

    #[test]
    fn missing_spreadsheet_surfaces_from_the_locator_finish() {
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:body/></office:document-content>"#
        );
        parse_protection(&xml).expect("semantic pass accepts a missing host");
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: missing office:spreadsheet element"
        );
        assert_equivalent("missing-spreadsheet", &xml, None);
    }

    #[test]
    fn duplicate_spreadsheet_and_automatic_styles() {
        // Both passes report the duplicate spreadsheet with the same message.
        let dup_spreadsheet = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:spreadsheet/><office:spreadsheet/></office:body></office:document-content>"#
        );
        let fused = parse(&dup_spreadsheet, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: duplicate office:spreadsheet element"
        );
        assert_equivalent("duplicate-spreadsheet", &dup_spreadsheet, None);

        // Only the locator tracks automatic-styles.
        let dup_automatic = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:automatic-styles/><office:automatic-styles/><office:body><office:spreadsheet/></office:body></office:document-content>"#
        );
        parse_protection(&dup_automatic).expect("semantic pass ignores automatic-styles");
        let fused = parse(&dup_automatic, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: duplicate office:automatic-styles element"
        );
        assert_equivalent("duplicate-automatic-styles", &dup_automatic, None);
    }

    #[test]
    fn malformed_xml_uses_the_shared_read_mapping() {
        // Both passes map quick_xml read failures to the same string, and the
        // locator is the first still-active handler.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:body></office:document-content>"#
        );
        let located = Location::parse(&xml, None).expect_err("locator must fail");
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(fused.to_string(), located.to_string());
        assert!(
            fused
                .to_string()
                .starts_with("Invalid format: XML parsing error: ")
        );
        assert_equivalent("malformed", &xml, None);
    }

    #[test]
    fn size_limit_fires_before_both_passes() {
        let padding = "x".repeat(65 * 1024 * 1024);
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><!--{padding}--><office:body><office:spreadsheet/></office:body></office:document-content>"#
        );
        let fused = parse(&xml, None).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: ODS protection content.xml exceeds the snapshot limit"
        );
        assert_equivalent("size-limit", &xml, None);
    }

    #[test]
    fn standalone_shells_match_their_handlers() {
        // The thin standalone shells drive the same checks as the fused
        // driver, so pin shell behavior against the fused result here.
        let xml = document(r#"<table:table table:name="S" table:protected="true"/>"#);
        let location = Location::parse(&xml, None).expect("locator shell");
        assert_eq!(location.sheets().len(), 1);
        assert_eq!(location.table_prefix(), "table");
        let (document, sheets) = parse_protection(&xml).expect("semantic shell");
        assert_eq!(document.structure_protected, None);
        assert_eq!(sheets[0].protected, Some(true));
        assert_equivalent("shells", &xml, None);
    }
}
