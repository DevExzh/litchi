//! Fused single-tokenization open parse for ODS `content.xml`.
//!
//! [`crate::facade::SourceBackedSpreadsheet`] opening historically tokenized
//! the same `content.xml` once per consumer.  This module runs the
//! package-content structure validation ([`crate::authoring`]), the semantic
//! calculation-settings parse ([`litchi_odf_common::calculation`]), the
//! calculation-settings location scan ([`crate::settings::codec`]), the
//! named-definition scan ([`crate::codec::names`]), and the worksheet parse
//! ([`crate::worksheet::codec`]) over one shared `NsReader` event stream.
//!
//! Observable behavior is identical to the sequential passes: each handler
//! keeps its own limits and error messages, a quick_xml read failure is
//! mapped with the first still-active handler in pass order, and the final
//! error selection follows the original call order — content validation, the
//! semantic calculation-settings parse, settings locate, the semantic/XML
//! disagreement check, named definitions, then worksheets.  The split into
//! [`OpenParse::run`] (the shared tokenization) and [`OpenParse::finish`]
//! (the post-stream pass checks and error selection) keeps the historical
//! interleave in which the styles/meta/metadata package members load after
//! the content validation has completed.

use litchi_core::{Error, Result};
use litchi_odf_common::calculation::{CalculationHandler, Settings};
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use crate::authoring::ValidateHandler;
use crate::codec::names::{self, NamesHandler};
use crate::model::names::{Definition, validate_collection};
use crate::settings::codec::{self, LocateHandler, Location};
use crate::worksheet::Sheet;
use crate::worksheet::codec::WorksheetHandler;

/// The outputs of the fused open passes over `content.xml`.
#[derive(Debug)]
pub(crate) struct OpenOutputs {
    /// The spreadsheet host and calculation-settings spans (pass 2b).
    pub(crate) location: Location,
    /// The validated named-definition catalog (pass 3).
    pub(crate) definitions: Vec<Definition>,
    /// The typed worksheets in document order (pass 4).
    pub(crate) sheets: Vec<Sheet>,
}

/// The two-phase fused open parse over `content.xml`.
///
/// [`OpenParse::run`] performs the size pre-checks and the single shared
/// tokenization driving the content-validation (pass 1), semantic
/// calculation-settings (pass 2a), settings-locate (pass 2b),
/// named-definition (pass 3), and worksheet (pass 4) handlers.
/// [`OpenParse::finish`] runs the post-stream pass checks and the final
/// error selection.  The caller loads the styles/meta/metadata package
/// members between the two phases, exactly where the historical sequential
/// opening loaded them.
pub(crate) struct OpenParse {
    validate: ValidateHandler,
    calculation: CalculationHandler,
    locate: LocateHandler,
    named: NamesHandler,
    worksheet: WorksheetHandler,
    calculation_error: Option<Error>,
    locate_error: Option<Error>,
    names_error: Option<Error>,
    worksheet_error: Option<Error>,
}

impl OpenParse {
    /// Tokenize `content_xml` once, driving the content-validation,
    /// calculation-settings, settings-locate, named-definition, and
    /// worksheet handlers over the shared event stream.
    ///
    /// Content validation is the first pass of the sequential opening, so
    /// its size limit runs up front and its mid-stream errors return
    /// immediately; its end-of-stream checks are reported when the loop
    /// finishes, still before [`OpenParse::finish`] selects among the
    /// pass-2+ results and before the caller loads the styles/meta/metadata
    /// members.  The calculation-settings, settings-locate,
    /// named-definition, and worksheet size limits historically fired only
    /// after the earlier passes had succeeded, so they are recorded as
    /// pre-stream pass errors: the handlers stay inactive and
    /// [`OpenParse::finish`] reports them at their original precedence
    /// positions.
    pub(crate) fn run(content_xml: &str) -> Result<Self> {
        crate::authoring::validate_size(content_xml)?;
        let mut calculation_error =
            litchi_odf_common::calculation::validate_size(content_xml).err();
        let mut locate_error = codec::validate_size(content_xml).err();
        let mut names_error = names::validate_size(content_xml).err();
        let mut worksheet_error =
            crate::worksheet::validation::validate_content_xml_size(content_xml).err();

        let mut reader = NsReader::from_str(content_xml);
        let mut buffer = Vec::new();
        let mut validate = ValidateHandler::default();
        let mut calculation = CalculationHandler::default();
        let mut locate = LocateHandler::default();
        let mut named = NamesHandler::default();
        let mut worksheet = WorksheetHandler::new(true);
        let mut validate_error = None;

        loop {
            let pos_before = reader.buffer_position();
            let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
                Ok(resolved) => resolved,
                Err(error) => {
                    // A broken stream stops every pass.  Standalone, the
                    // earliest still-running pass reports the failure with
                    // its own mapping and later passes never run, so record
                    // the error for the first still-active handler in pass
                    // order using its mapping.  Content validation is active
                    // on every read failure (its mid-stream errors stop the
                    // loop above and its end-of-stream errors end it), so
                    // the later arms exist only for structural completeness.
                    if validate_error.is_none() {
                        validate_error = Some(Error::InvalidFormat(format!(
                            "invalid ODS content.xml: {error}"
                        )));
                    } else if calculation_error.is_none() {
                        calculation_error =
                            Some(Error::InvalidFormat(format!("XML parsing error: {error}")));
                    } else if locate_error.is_none() {
                        locate_error = Some(Error::InvalidFormat(format!(
                            "invalid ODS content.xml: {error}"
                        )));
                    } else if names_error.is_none() {
                        names_error = Some(Error::InvalidFormat(format!(
                            "invalid ODS content.xml: {error}"
                        )));
                    } else if worksheet_error.is_none() {
                        worksheet_error =
                            Some(Error::InvalidFormat(format!("invalid ODS XML: {error}")));
                    }
                    break;
                },
            };
            // The resolved namespace borrows the reader mutably, so classify
            // it for all five passes before touching the reader again,
            // exactly where each historical loop body classified it.
            let validate_office = crate::authoring::is_office_namespace(&namespace);
            let (calculation_table, calculation_office) =
                litchi_odf_common::calculation::classify(&namespace);
            let locate_namespace = codec::namespace_kind(&namespace);
            let (is_table, is_office) = names::classify(&namespace);
            let worksheet_namespace = crate::worksheet::codec::namespace_kind(&namespace);
            let pos_after = reader.buffer_position();
            let is_eof = matches!(event, Event::Eof);
            if validate_error.is_none()
                && let Err(error) = validate.on_event(validate_office, &event)
            {
                // Content validation ran to completion before the later
                // passes in the sequential opening, so its mid-stream errors
                // end the open immediately; the end-of-stream checks are
                // recorded and reported below, still ahead of pass 2a.
                if is_eof {
                    validate_error = Some(error);
                } else {
                    return Err(error);
                }
            }
            if calculation_error.is_none()
                && let Err(error) = calculation.on_event(
                    calculation_table,
                    calculation_office,
                    &event,
                    reader.resolver(),
                    reader.decoder(),
                )
            {
                calculation_error = Some(error);
            }
            if locate_error.is_none()
                && let Err(error) = locate.on_event(
                    locate_namespace,
                    &event,
                    reader.resolver(),
                    reader.decoder(),
                    pos_before,
                    pos_after,
                )
            {
                locate_error = Some(error);
            }
            if names_error.is_none()
                && let Err(error) = named.on_event(
                    is_table,
                    is_office,
                    &event,
                    reader.resolver(),
                    reader.decoder(),
                    pos_before,
                    pos_after,
                )
            {
                names_error = Some(error);
            }
            if worksheet_error.is_none()
                && let Err(error) = worksheet.on_event(
                    worksheet_namespace,
                    &event,
                    reader.resolver(),
                    reader.decoder(),
                    pos_before,
                    pos_after,
                )
            {
                worksheet_error = Some(error);
            }
            buffer.clear();
            if is_eof {
                break;
            }
        }

        // Content validation completed before the later passes historically
        // ran, so its recorded end-of-stream or read-failure error still
        // beats every pass-2 result and the styles/meta/metadata member
        // loads the caller performs next.
        if let Some(error) = validate_error {
            return Err(error);
        }

        Ok(Self {
            validate,
            calculation,
            locate,
            named,
            worksheet,
            calculation_error,
            locate_error,
            names_error,
            worksheet_error,
        })
    }

    /// Run the post-stream pass checks and select the first error in the
    /// original call order: content validation, the semantic
    /// calculation-settings parse, settings locate, the semantic/XML
    /// disagreement check, named definitions, then worksheets.
    pub(crate) fn finish(self) -> Result<(Option<Settings>, OpenOutputs)> {
        // `run` surfaces every content-validation error before returning, so
        // this finish can never fail; it only keeps the handler exhausted.
        self.validate.finish()?;
        if let Some(error) = self.calculation_error {
            return Err(error);
        }
        let calculation = self.calculation.finish()?;
        if let Some(error) = self.locate_error {
            return Err(error);
        }
        let location = self.locate.finish()?;
        if calculation.is_some() != location.calculation.is_some() {
            return Err(Error::InvalidFormat(
                "calculation-settings semantic and XML locations disagree".to_string(),
            ));
        }
        if let Some(error) = self.names_error {
            return Err(error);
        }
        let scan = self.named.finish()?;
        validate_collection(&scan.definitions)?;
        if let Some(error) = self.worksheet_error {
            return Err(error);
        }
        let sheets = self.worksheet.finish()?;

        Ok((
            calculation,
            OpenOutputs {
                location,
                definitions: scan.definitions,
                sheets,
            },
        ))
    }
}

#[cfg(test)]
mod tests {
    use super::{OpenOutputs, OpenParse};
    use litchi_core::{Error, Result};
    use std::io::Read as _;
    use std::path::{Path, PathBuf};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    const DISAGREE: &str = "calculation-settings semantic and XML locations disagree";

    type Compared = (
        String,
        String,
        Vec<crate::model::names::Definition>,
        Vec<crate::worksheet::Sheet>,
    );

    /// The original sequential open passes with their exact error precedence.
    fn sequential(xml: &str) -> Result<Compared> {
        crate::authoring::validate_content_xml(xml)?;
        let calculation = litchi_odf_common::calculation::parse(xml)?;
        let location = crate::settings::codec::locate(xml)?;
        if calculation.is_some() != location.calculation.is_some() {
            return Err(Error::InvalidFormat(DISAGREE.to_string()));
        }
        let definitions = crate::codec::names::parse(xml)?;
        let sheets = crate::worksheet::codec::parse(xml)?;
        Ok((
            format!("{calculation:?}"),
            format!("{location:?}"),
            definitions,
            sheets,
        ))
    }

    /// The fused two-phase open parse with the same comparison projection.
    fn fused(xml: &str) -> Result<Compared> {
        let (calculation, outputs) = OpenParse::run(xml)?.finish()?;
        Ok((
            format!("{calculation:?}"),
            format!("{:?}", outputs.location),
            outputs.definitions,
            outputs.sheets,
        ))
    }

    fn assert_equivalent(label: &str, xml: &str) {
        let expected = sequential(xml).map_err(|error| error.to_string());
        let actual = fused(xml).map_err(|error| error.to_string());
        assert_eq!(expected, actual, "{label}: fused and sequential disagree");
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

    fn content_xml(path: &Path) -> Option<String> {
        let bytes = std::fs::read(path).ok()?;
        let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).ok()?;
        let mut entry = archive.by_name("content.xml").ok()?;
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
            let Some(xml) = content_xml(path) else {
                continue;
            };
            assert_equivalent(&path.display().to_string(), &xml);
            compared += 1;
        }
        assert!(compared > 0, "no .ods corpus fixtures yielded content.xml");
    }

    #[test]
    fn fused_parse_matches_sequential_passes_on_minimal_document() {
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:calculation-settings table:case-sensitive="true"/><table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="string"><text:p>hi</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        let calculation =
            litchi_odf_common::calculation::parse(&xml).expect("minimal document parses");
        assert!(calculation.is_some(), "fixture must declare settings");
        assert_equivalent("minimal", &xml);
    }

    #[test]
    fn validate_error_beats_semantic_parse() {
        // Two office:body elements fail content validation while the
        // duplicate calculation-settings would fail the semantic parse; the
        // content-validation pass runs first.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:calculation-settings/><table:calculation-settings/></office:spreadsheet></office:body><office:body/></office:document-content>"#
        );
        let semantic = litchi_odf_common::calculation::parse(&xml)
            .expect_err("semantic parse must fail on duplicates");
        assert_eq!(
            semantic.to_string(),
            "Invalid format: duplicate table:calculation-settings"
        );
        let fused = OpenParse::run(&xml)
            .err()
            .expect("content validation must fail in run");
        assert_eq!(
            fused.to_string(),
            "Invalid format: ODS content.xml has more than one office:body"
        );
        assert_equivalent("validate-beats-semantic", &xml);
    }

    #[test]
    fn semantic_parse_beats_locate_error() {
        // Duplicate calculation-settings elements: content validation
        // accepts the structure, the semantic parse fails, and locate
        // reports its own duplicate; pass 2a is selected before locate.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:calculation-settings/><table:calculation-settings/></office:spreadsheet></office:body></office:document-content>"#
        );
        crate::authoring::validate_content_xml(&xml).expect("structure is valid");
        let located = crate::settings::codec::locate(&xml).expect_err("locate must fail");
        assert_eq!(
            located.to_string(),
            "Invalid format: ODS content.xml has duplicate table:calculation-settings"
        );
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: duplicate table:calculation-settings"
        );
        assert_equivalent("semantic-beats-locate", &xml);
    }

    #[test]
    fn locate_error_beats_names_error() {
        // An unknown calculation-settings attribute fails locate (content
        // validation and the semantic pass ignore it), and a nested
        // named-expressions host fails names.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:calculation-settings table:bogus="1"/><table:named-expressions><table:named-expressions/></table:named-expressions></office:spreadsheet></office:body></office:document-content>"#
        );
        crate::authoring::validate_content_xml(&xml).expect("structure is valid");
        let calculation = litchi_odf_common::calculation::parse(&xml)
            .expect("unknown attributes are inert to the semantic parse");
        assert!(calculation.is_some(), "fixture must declare settings");
        assert!(
            crate::codec::names::parse(&xml).is_err(),
            "names pass must fail on this fixture"
        );
        let fused = fused(&xml).expect_err("fused parse must fail");
        let located = crate::settings::codec::locate(&xml).expect_err("locate must fail");
        assert_eq!(fused.to_string(), located.to_string());
        assert!(fused.to_string().contains("unknown attribute"));
        assert_equivalent("locate-error-beats-names", &xml);
    }

    #[test]
    fn disagreement_beats_names_error() {
        // The calculation-settings element sits inside an office:text
        // sibling of the spreadsheet host: the semantic pass accepts it as a
        // document-body child while locate classifies the host as foreign,
        // and a nested named-expressions host fails names.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:named-expressions><table:named-expressions/></table:named-expressions></office:spreadsheet><office:text><table:calculation-settings/></office:text></office:body></office:document-content>"#
        );
        crate::authoring::validate_content_xml(&xml).expect("structure is valid");
        let calculation =
            litchi_odf_common::calculation::parse(&xml).expect("sibling settings parse");
        assert!(calculation.is_some(), "fixture must disagree semantically");
        assert!(
            crate::codec::names::parse(&xml).is_err(),
            "names pass must fail on this fixture"
        );
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert_eq!(fused.to_string(), format!("Invalid format: {DISAGREE}"));
        assert_equivalent("disagreement-beats-names", &xml);
    }

    #[test]
    fn names_error_beats_worksheet_error() {
        // A nested named-expressions host fails names, and the stray empty
        // table row fails the worksheet pass; locate accepts the document.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:named-expressions><table:named-expressions/></table:named-expressions><table:table-row/></office:spreadsheet></office:body></office:document-content>"#
        );
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: nested table:named-expressions element"
        );
        assert_equivalent("names-error-beats-worksheet", &xml);
    }

    /// Build a document with a comment padding it past the named-definition
    /// 16 MiB size limit, with `tail` appended inside `office:body`.
    fn oversized_xml(tail: &str) -> String {
        let padding = "x".repeat(17 * 1024 * 1024);
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><!--{padding}-->{tail}</office:body></office:document-content>"#
        )
    }

    /// Build a document padded past the calculation-settings 64 MiB size
    /// limit but under the content-validation 256 MiB limit.
    fn past_calculation_limit_xml(tail: &str) -> String {
        let padding = "x".repeat(65 * 1024 * 1024);
        format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><!--{padding}-->{tail}</office:body></office:document-content>"#
        )
    }

    #[test]
    fn calculation_size_limit_beats_later_passes() {
        // Over the calculation-settings size limit but otherwise
        // well-formed: pass 2a's size error surfaces at its original
        // position, ahead of the locate/names/worksheet results (whose own
        // size limits also trip but lose the selection).
        let xml = past_calculation_limit_xml(r#"<office:spreadsheet/>"#);
        crate::authoring::validate_content_xml(&xml).expect("structure is valid");
        let semantic =
            litchi_odf_common::calculation::parse(&xml).expect_err("semantic size check must fail");
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert_eq!(fused.to_string(), semantic.to_string());
        assert_eq!(
            fused.to_string(),
            "Invalid format: calculation settings XML exceeds the size limit"
        );
        assert_equivalent("calculation-size-limit", &xml);
    }

    #[test]
    fn locate_error_beats_names_size_limit() {
        // Over the named-definition size limit AND with a
        // calculation-settings attribute locate rejects: the sequential
        // passes report the locate failure because it runs before the
        // named-definition size check.
        let xml = oversized_xml(
            r#"<office:spreadsheet><table:calculation-settings table:bogus="1"/></office:spreadsheet>"#,
        );
        crate::authoring::validate_content_xml(&xml).expect("structure is valid");
        let calculation = litchi_odf_common::calculation::parse(&xml)
            .expect("unknown attributes are inert to the semantic parse");
        assert!(calculation.is_some(), "fixture must declare settings");
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert!(
            fused.to_string().contains("unknown attribute"),
            "unexpected error: {fused}"
        );
        assert_equivalent("locate-error-beats-names-size", &xml);
    }

    #[test]
    fn names_size_limit_beats_worksheet_pass() {
        // Over the named-definition size limit but otherwise well-formed: the
        // named-definition size error surfaces at its original position.
        let xml = oversized_xml(r#"<office:spreadsheet/>"#);
        let fused = fused(&xml).expect_err("fused parse must fail");
        assert_eq!(
            fused.to_string(),
            "Invalid format: ODS content.xml exceeds the mutation limit"
        );
        assert_equivalent("names-size-limit", &xml);
    }

    #[test]
    fn malformed_xml_reports_the_validate_mapping() {
        // The read failure mapping belongs to the first still-active pass;
        // with every handler active that is the content-validation mapping
        // (the same "invalid ODS content.xml: ..." string locate uses).
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}"><office:body></office:document-content>"#
        );
        let fused = fused(&xml).expect_err("fused parse must fail");
        let validated =
            crate::authoring::validate_content_xml(&xml).expect_err("validate must fail");
        assert_eq!(fused.to_string(), validated.to_string());
        assert!(
            fused
                .to_string()
                .starts_with("Invalid format: invalid ODS content.xml: ")
        );
    }

    #[test]
    fn standalone_shells_match_their_handlers() {
        // The thin standalone shells drive the same handlers as the fused
        // driver, so pin shell behavior against direct handler use here.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="float" office:value="1.5"><text:p>1.5</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
        );
        crate::authoring::validate_content_xml(&xml).expect("validate shell");
        let located = crate::settings::codec::locate(&xml).expect("locate shell");
        assert!(located.calculation.is_none());
        let definitions = crate::codec::names::parse(&xml).expect("names shell");
        assert!(definitions.is_empty());
        let sheets = crate::worksheet::codec::parse(&xml).expect("worksheet shell");
        assert_eq!(sheets.len(), 1);
        assert_equivalent("shells", &xml);
    }

    #[test]
    fn open_outputs_are_constructible_for_debug() {
        // Keep the output projection shape exercised outside the driver.
        let xml = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet/></office:body></office:document-content>"#
        );
        let (calculation, outputs): (_, OpenOutputs) = OpenParse::run(&xml)
            .and_then(OpenParse::finish)
            .expect("fused parse");
        assert!(calculation.is_none());
        assert!(outputs.definitions.is_empty());
        assert!(outputs.sheets.is_empty());
        assert!(outputs.location.calculation.is_none());
    }
}
