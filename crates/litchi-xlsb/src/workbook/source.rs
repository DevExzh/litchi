//! Deferred, source-backed XLSB worksheet catalog and materialization.

use std::fmt;
use std::io::{self, Cursor, Write};
use std::sync::{Arc, Mutex};

use litchi_core::sheet::{Cell as SheetCell, Worksheet as SheetWorksheet};
use litchi_core::{
    ExecutionContext, ReadAt, SequentialTextWriter, SourceVersion, TextObjectKind, TextOutputError,
    TextOutputOptions, TextOutputReport,
};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{
    PackURI, PartView, ReadLimits, SourceBackedPackage, SourceCacheDiagnostics, SourceCacheLimits,
};
use once_cell::sync::OnceCell;

use super::model::Workbook;
use crate::package::error::{Error, Result};
use crate::package::formula::Context;
use crate::package::shared_strings::SharedString;
use crate::package::styles_table::StylesTable;
use crate::raw::{Records, kind};
use crate::sheet::Worksheet;

const XLSB_WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";
const XLSB_CHARTSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.chartsheet";
const XLSB_DIALOGSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.dialogsheet";
const XLSB_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.macrosheet";
const XLSB_INTL_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.intlmacrosheet";
const XLSB_MAX_ROW_INDEX: u32 = 1_048_575;
const XLSB_MAX_COLUMN_INDEX: u32 = 16_383;
const XLSB_SHARED_STRINGS_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";
const XLSB_STYLES_CONTENT_TYPE: &str = "application/vnd.ms-excel.styles";
const XLSB_COMMENTS_CONTENT_TYPE: &str = "application/vnd.ms-excel.comments";
const XLSB_CHARTSHEET_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const XLSB_STRICT_CHARTSHEET_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const XLSB_DIALOGSHEET_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const XLSB_STRICT_DIALOGSHEET_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet";
const XLSB_MACROSHEET_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const XLSB_INTL_MACROSHEET_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SheetKind {
    Worksheet,
    Chart,
    Dialog,
    Macro,
    IntlMacro,
}

#[derive(Debug)]
struct SheetMetadata {
    name: String,
    workbook_position: usize,
    partname: PackURI,
    kind: SheetKind,
}

struct SourceInner {
    package: SourceBackedPackage,
    sheets: Vec<SheetMetadata>,
    worksheet_positions: Vec<usize>,
    formula_context: Context,
    shared_strings_part: Option<PackURI>,
    styles_part: Option<PackURI>,
    incomplete_formula_context: bool,
    is_1904_date_system: bool,
    shared_strings: SemanticCache<Vec<SharedString>>,
    styles: SemanticCache<StylesTable>,
}

/// A bounded XLSB workbook catalog whose worksheet bodies remain deferred.
///
/// Opening reads OPC metadata and `workbook.bin`, but not ordinary worksheet,
/// shared-string, or style payloads. The handle is immutable and source-bound.
/// `SourceCacheLimits` bounds retained OPC payload bytes; successfully parsed
/// shared-string and style values are retained separately and remain bounded by
/// the configured per-Part read limits.
#[derive(Clone)]
pub struct SourceBackedWorkbook {
    inner: Arc<SourceInner>,
}

/// A source-bound handle for one sheet in an XLSB workbook.
///
/// The handle contains only catalog metadata. Calling [`Self::materialize`]
/// reads the selected BIFF12 worksheet body and its shared-string/style inputs;
/// non-worksheet handles return a typed capability error.
#[derive(Clone)]
pub struct SourceBackedWorksheet {
    inner: Arc<SourceInner>,
    catalog_position: usize,
}

impl SourceBackedWorkbook {
    /// Open with the default bounded OPC read and cache policies.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at(source)?)
    }

    /// Open with an explicit bounded OPC read policy.
    pub fn from_read_at_with_limits(source: Arc<dyn ReadAt>, limits: ReadLimits) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_limits(
            source, limits,
        )?)
    }

    /// Open with an explicit finite deferred-payload cache policy.
    pub fn from_read_at_with_cache_limits(
        source: Arc<dyn ReadAt>,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_cache_limits(
            source,
            cache_limits,
        )?)
    }

    /// Open with explicit bounded OPC read and cache policies.
    pub fn from_read_at_with_limits_and_cache_limits(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits(
                source,
                limits,
                cache_limits,
            )?,
        )
    }

    /// Open with explicit read and caller-owned execution policies.
    pub fn from_read_at_with_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(SourceBackedPackage::from_read_at_with_execution_context(
            source, limits, context,
        )?)
    }

    /// Open with explicit read, cache, and caller-owned execution policies.
    pub fn from_read_at_with_limits_and_cache_limits_and_execution_context(
        source: Arc<dyn ReadAt>,
        limits: ReadLimits,
        cache_limits: SourceCacheLimits,
        context: ExecutionContext,
    ) -> Result<Self> {
        Self::from_source_backed_package(
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
                source,
                limits,
                cache_limits,
                context,
            )?,
        )
    }

    /// Build a lazy XLSB catalog from an already validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        package.check_execution()?;
        let workbook_part = package.main_document_part()?;
        if workbook_part.content_type() != content_type::XLSB_BIN {
            return Err(Error::InvalidContentType {
                expected: content_type::XLSB_BIN.to_string(),
                got: workbook_part.content_type().to_string(),
            });
        }

        let workbook_data = workbook_part.data()?;
        let mut records = Records::new(workbook_data.as_bytes());
        let info = Workbook::read_workbook(&mut records)?;
        if info.worksheet_names.len() != info.worksheet_rel_ids.len() {
            return Err(Error::InvalidRelationship(
                "XLSB sheet names and relationship identifiers have different lengths".to_string(),
            ));
        }

        let shared_strings_part = optional_related_part(
            &workbook_part,
            &[
                relationship_type::SHARED_STRINGS,
                relationship_type::STRICT_SHARED_STRINGS,
            ],
            "shared strings",
        )?;
        let styles_part = optional_related_part(
            &workbook_part,
            &[relationship_type::STYLES, relationship_type::STRICT_STYLES],
            "styles",
        )?;

        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(info.worksheet_names.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB sheet catalog",
                source,
            })?;
        let mut worksheet_positions = Vec::new();
        worksheet_positions
            .try_reserve_exact(info.worksheet_names.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB worksheet catalog",
                source,
            })?;
        let mut sheet_targets = Vec::new();
        sheet_targets
            .try_reserve_exact(info.worksheet_names.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB worksheet target validation",
                source,
            })?;
        let mut incomplete_formula_context = !info.external_link_rel_ids.is_empty();

        for (workbook_position, (name, relationship_id)) in info
            .worksheet_names
            .iter()
            .zip(&info.worksheet_rel_ids)
            .enumerate()
        {
            let relationship_id = relationship_id.as_deref().ok_or_else(|| {
                Error::InvalidRelationship(format!(
                    "XLSB sheet {name:?} has no relationship identifier"
                ))
            })?;
            let relationship = workbook_part.rels().get(relationship_id).ok_or_else(|| {
                Error::InvalidRelationship(format!(
                    "XLSB sheet {name:?} relationship {relationship_id:?} is missing"
                ))
            })?;
            if relationship.is_external() {
                return Err(Error::InvalidRelationship(format!(
                    "XLSB sheet {name:?} has an external relationship"
                )));
            }
            let partname = relationship.target_partname()?;
            if sheet_targets.contains(&partname) {
                return Err(Error::InvalidRelationship(format!(
                    "multiple XLSB sheets resolve to {partname}"
                )));
            }
            sheet_targets.push(partname.clone());
            let sheet_part = package.part(&partname)?;
            let sheet_kind = if matches!(
                relationship.reltype(),
                relationship_type::WORKSHEET | relationship_type::STRICT_WORKSHEET
            ) {
                require_content_type(&sheet_part, XLSB_WORKSHEET_CONTENT_TYPE)?;
                incomplete_formula_context |= sheet_part.rels().iter().any(|relationship| {
                    matches!(
                        relationship.reltype(),
                        relationship_type::TABLE
                            | relationship_type::STRICT_TABLE
                            | relationship_type::PIVOT_TABLE
                            | relationship_type::STRICT_PIVOT_TABLE
                    )
                });
                SheetKind::Worksheet
            } else {
                match relationship.reltype() {
                    XLSB_CHARTSHEET_RELATIONSHIP | XLSB_STRICT_CHARTSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_CHARTSHEET_CONTENT_TYPE)?;
                        SheetKind::Chart
                    },
                    XLSB_DIALOGSHEET_RELATIONSHIP | XLSB_STRICT_DIALOGSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_DIALOGSHEET_CONTENT_TYPE)?;
                        SheetKind::Dialog
                    },
                    XLSB_MACROSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_MACROSHEET_CONTENT_TYPE)?;
                        SheetKind::Macro
                    },
                    XLSB_INTL_MACROSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_INTL_MACROSHEET_CONTENT_TYPE)?;
                        SheetKind::IntlMacro
                    },
                    _ => {
                        return Err(Error::InvalidRelationship(format!(
                            "XLSB sheet {name:?} has unsupported relationship type {:?}",
                            relationship.reltype()
                        )));
                    },
                }
            };
            let catalog_position = sheets.len();
            sheets.push(SheetMetadata {
                name: name.clone(),
                workbook_position,
                partname,
                kind: sheet_kind,
            });
            if matches!(sheet_kind, SheetKind::Worksheet) {
                worksheet_positions.push(catalog_position);
            }
        }

        if let Some(partname) = shared_strings_part.as_ref() {
            let part = package.part(partname)?;
            require_content_type(&part, XLSB_SHARED_STRINGS_CONTENT_TYPE)?;
        }
        if let Some(partname) = styles_part.as_ref() {
            let part = package.part(partname)?;
            require_content_type(&part, XLSB_STYLES_CONTENT_TYPE)?;
        }

        let is_1904_date_system = info.is_1904;
        let formula_context = Context {
            worksheet_names: info.worksheet_names.into(),
            supporting_links: info.supporting_links.into(),
            external_sheets: info.external_sheets.into(),
            external_books: Vec::new().into(),
            defined_names: info.defined_names.into(),
            tables: Vec::new().into(),
            pivot_views: Vec::new().into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        package.source_version()?;

        Ok(Self {
            inner: Arc::new(SourceInner {
                package,
                sheets,
                worksheet_positions,
                formula_context,
                shared_strings_part,
                styles_part,
                incomplete_formula_context,
                is_1904_date_system,
                shared_strings: SemanticCache::new(),
                styles: SemanticCache::new(),
            }),
        })
    }

    /// Return the number of workbook tabs after checking source freshness.
    pub fn sheet_count(&self) -> Result<usize> {
        self.inner.package.source_version()?;
        Ok(self.inner.sheets.len())
    }

    /// Snapshot all workbook tab names in workbook order after checking source freshness.
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        self.inner.package.source_version()?;
        let mut names = Vec::new();
        names
            .try_reserve_exact(self.inner.sheets.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB sheet-name snapshot",
                source,
            })?;
        names.extend(self.inner.sheets.iter().map(|sheet| sheet.name.clone()));
        Ok(names)
    }

    /// Snapshot all checked workbook tab handles in workbook order.
    pub fn sheets(&self) -> Result<Vec<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(self.inner.sheets.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB sheet-handle snapshot",
                source,
            })?;
        sheets.extend(
            (0..self.inner.sheets.len()).map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            }),
        );
        Ok(sheets)
    }

    /// Select a workbook tab by zero-based position in complete workbook order.
    pub fn sheet_by_index(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        Ok(self.inner.sheets.get(index).map(|_| SourceBackedWorksheet {
            inner: Arc::clone(&self.inner),
            catalog_position: index,
        }))
    }

    /// Select a workbook tab by its exact workbook name.
    pub fn sheet_by_name(&self, name: &str) -> Result<Option<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        Ok(self
            .inner
            .sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            }))
    }

    /// Return the number of worksheet owners after checking source freshness.
    pub fn worksheet_count(&self) -> Result<usize> {
        self.inner.package.source_version()?;
        Ok(self.inner.worksheet_positions.len())
    }

    /// Snapshot worksheet names in workbook order after checking source freshness.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        self.inner.package.source_version()?;
        let mut names = Vec::new();
        names
            .try_reserve_exact(self.inner.worksheet_positions.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB worksheet-name snapshot",
                source,
            })?;
        names.extend(
            self.inner
                .worksheet_positions
                .iter()
                .map(|&catalog_position| self.inner.sheets[catalog_position].name.clone()),
        );
        Ok(names)
    }

    /// Snapshot all checked worksheet handles in workbook order.
    pub fn worksheets(&self) -> Result<Vec<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        let mut worksheets = Vec::new();
        worksheets
            .try_reserve_exact(self.inner.worksheet_positions.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB worksheet-handle snapshot",
                source,
            })?;
        worksheets.extend(
            self.inner
                .worksheet_positions
                .iter()
                .copied()
                .map(|catalog_position| SourceBackedWorksheet {
                    inner: Arc::clone(&self.inner),
                    catalog_position,
                }),
        );
        Ok(worksheets)
    }

    /// Select a worksheet by zero-based position in the worksheet-only catalog.
    pub fn worksheet_by_index(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        Ok(self
            .inner
            .worksheet_positions
            .get(index)
            .copied()
            .map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            }))
    }

    /// Select a worksheet by its exact workbook name.
    pub fn worksheet_by_name(&self, name: &str) -> Result<Option<SourceBackedWorksheet>> {
        self.inner.package.source_version()?;
        Ok(self
            .inner
            .worksheet_positions
            .iter()
            .copied()
            .find(|&catalog_position| self.inner.sheets[catalog_position].name == name)
            .map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            }))
    }

    /// Return the exact source identity and revision captured at open.
    pub fn source_version(&self) -> Result<SourceVersion> {
        self.inner.package.source_version().map_err(Into::into)
    }

    /// Return whether this workbook uses Excel's 1904 date system.
    pub fn is_1904_date_system(&self) -> Result<bool> {
        self.inner.package.source_version()?;
        Ok(self.inner.is_1904_date_system)
    }

    /// Return content-free deferred-Part cache diagnostics.
    #[must_use]
    pub fn cache_diagnostics(&self) -> SourceCacheDiagnostics {
        self.inner.package.cache_diagnostics()
    }

    /// Stream ordinary worksheet rows to a caller-owned sequential sink.
    ///
    /// Rows are emitted as paragraph-like objects with tab-separated cells.
    /// The shared output policy controls separators, empty rows, and output
    /// limits. This method never appends a terminal separator and never falls
    /// back to the eager workbook owner.
    pub fn write_text_to<W: Write + ?Sized>(
        &self,
        output: &mut W,
        options: TextOutputOptions<'_>,
    ) -> std::result::Result<TextOutputReport, TextOutputError<Error>> {
        let failure = Arc::new(Mutex::new(None));
        let mut checked_output = SourceCheckedTextSink {
            output,
            owner: &self.inner,
            failure: Arc::clone(&failure),
        };
        let mut writer = SequentialTextWriter::new(&mut checked_output, options);
        let conversion = (|| {
            check_text_state(&self.inner).map_err(|source| writer.document_error(source))?;
            for catalog_position in 0..self.inner.sheets.len() {
                check_text_state(&self.inner).map_err(|source| writer.document_error(source))?;
                let sheet = SourceBackedWorksheet {
                    inner: Arc::clone(&self.inner),
                    catalog_position,
                };
                let worksheet = sheet
                    .materialize()
                    .map_err(|source| writer.document_error(source))?;
                write_text_worksheet(&self.inner, &worksheet, &mut writer, options)?;
                drop(worksheet);
                check_text_state(&self.inner).map_err(|source| writer.document_error(source))?;
            }
            Ok::<(), TextOutputError<Error>>(())
        })();

        let progress = writer.progress();
        if let Some(source) = take_source_text_failure(&failure) {
            return Err(TextOutputError::Document { source, progress });
        }
        if let Err(source) = check_text_state(&self.inner) {
            return Err(TextOutputError::Document { source, progress });
        }
        conversion.map(|()| writer.finish())
    }

    /// Extract ordinary worksheet text while retaining the legacy terminal
    /// newline projection used by the facade.
    pub fn text(&self) -> Result<String> {
        let mut collector = FallibleTextCollector::default();
        let report = match self.write_text_to(&mut collector, TextOutputOptions::default()) {
            Ok(report) => report,
            Err(error) => {
                let allocation = collector.allocation.take();
                return Err(map_text_output_error(error, allocation));
            },
        };
        if report.objects_written() != 0 {
            check_text_state(&self.inner)?;
            if let Err(error) = collector.push_terminal_newline(
                report.bytes_written(),
                TextOutputOptions::default().max_output_bytes(),
            ) {
                return match check_text_state(&self.inner) {
                    Ok(()) => Err(error),
                    Err(source) => Err(source),
                };
            }
            check_text_state(&self.inner)?;
        }
        check_text_state(&self.inner)?;
        String::from_utf8(collector.take_bytes()).map_err(|error| {
            drop(error);
            Error::Encoding("XLSB text output was not valid UTF-8".to_string())
        })
    }
}

impl SourceBackedWorksheet {
    /// Return this sheet's name after checking source freshness.
    pub fn name(&self) -> Result<&str> {
        self.inner.package.source_version()?;
        Ok(&self.metadata().name)
    }

    /// Return this sheet's zero-based position in the complete workbook sheet order.
    pub fn workbook_position(&self) -> Result<usize> {
        self.inner.package.source_version()?;
        Ok(self.metadata().workbook_position)
    }

    /// Parse the selected worksheet BIFF12 stream without reading unselected sheets.
    ///
    /// This first source-backed layer materializes stream-owned worksheet data.
    /// Workbook adjunct owners such as drawings, PivotTables, slicers, and
    /// timelines remain separately deferred rather than forcing eager loading.
    pub fn materialize(&self) -> Result<Worksheet> {
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        let metadata = self.metadata();
        if !matches!(metadata.kind, SheetKind::Worksheet) {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB non-worksheet materialization is not supported".to_string(),
            ));
        }
        if self.inner.incomplete_formula_context {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB worksheet materialization requires deferred external, table, or PivotTable formula owners"
                    .to_string(),
            ));
        }
        let worksheet_part = self.inner.package.part(&metadata.partname)?;
        if worksheet_part.rels().iter().any(|relationship| {
            relationship.reltype().contains("/slicer")
                || relationship.reltype().contains("/timeline")
        }) {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB worksheet materialization does not yet include slicer or timeline views"
                    .to_string(),
            ));
        }
        let comments_part = optional_related_part(
            &worksheet_part,
            &[
                relationship_type::COMMENTS,
                relationship_type::STRICT_COMMENTS,
            ],
            "comments",
        )?;
        let worksheet_data = worksheet_part.data()?;
        for record in Records::new(worksheet_data.as_bytes()) {
            if record?.kind() == kind::BEGIN_SPARKLINE_GROUPS {
                return Err(Error::UnsupportedFeature(
                    "source-backed XLSB worksheet materialization does not yet include sparkline groups"
                        .to_string(),
                ));
            }
        }
        let shared_strings = self.inner.shared_strings()?;
        let styles = self.inner.styles()?;
        let mut worksheet = Workbook::read_worksheet(
            Cursor::new(worksheet_data.as_bytes()),
            metadata.name.clone(),
            shared_strings.as_slice(),
            &self.inner.formula_context,
            metadata.workbook_position,
            styles.cell_xfs.len(),
        )?;
        worksheet.set_scenarios(crate::package::scenarios::parse_worksheet(
            worksheet_data.as_bytes(),
        )?);
        if let Some(partname) = comments_part {
            let comments_part = self.inner.package.part(&partname)?;
            require_content_type(&comments_part, XLSB_COMMENTS_CONTENT_TYPE)?;
            if !comments_part.rels().is_empty() {
                return Err(Error::InvalidRelationship(
                    "XLSB comments parts cannot own relationships".to_string(),
                ));
            }
            let comments_data = comments_part.data()?;
            for comment in crate::comments::read(comments_data.as_bytes())? {
                worksheet.add_comment(comment);
            }
        }
        self.inner.package.source_version()?;
        Ok(worksheet)
    }

    fn metadata(&self) -> &SheetMetadata {
        &self.inner.sheets[self.catalog_position]
    }
}

struct SemanticCache<T> {
    value: OnceCell<Arc<T>>,
}

impl<T> SemanticCache<T> {
    const fn new() -> Self {
        Self {
            value: OnceCell::new(),
        }
    }

    fn get_or_try_init<E, F>(&self, init: F) -> std::result::Result<Arc<T>, E>
    where
        F: FnOnce() -> std::result::Result<Arc<T>, E>,
    {
        self.value.get_or_try_init(init).map(Arc::clone)
    }
}

impl SourceInner {
    fn shared_strings(&self) -> Result<Arc<Vec<SharedString>>> {
        self.package.check_execution()?;
        self.package.source_version()?;
        let strings = self.shared_strings.get_or_try_init(|| {
            self.package.check_execution()?;
            self.package.source_version()?;
            let mut strings = Vec::new();
            if let Some(partname) = self.shared_strings_part.as_ref() {
                let data = self.package.part(partname)?.data()?;
                self.package.check_execution()?;
                self.package.source_version()?;
                let mut records = Records::new(data.as_bytes());
                Workbook::read_shared_strings(&mut records, &mut strings)?;
                self.package.check_execution()?;
                self.package.source_version()?;
            }
            let strings = Arc::new(strings);
            self.package.check_execution()?;
            self.package.source_version()?;
            Ok::<_, Error>(strings)
        })?;
        self.package.check_execution()?;
        self.package.source_version()?;
        Ok(strings)
    }

    fn styles(&self) -> Result<Arc<StylesTable>> {
        self.package.check_execution()?;
        self.package.source_version()?;
        let styles = self.styles.get_or_try_init(|| {
            self.package.check_execution()?;
            self.package.source_version()?;
            let styles = if let Some(partname) = self.styles_part.as_ref() {
                let data = self.package.part(partname)?.data()?;
                self.package.check_execution()?;
                self.package.source_version()?;
                let styles = StylesTable::from_bytes(data.as_bytes())?;
                self.package.check_execution()?;
                self.package.source_version()?;
                styles
            } else {
                StylesTable::default()
            };
            let styles = Arc::new(styles);
            self.package.check_execution()?;
            self.package.source_version()?;
            Ok::<_, Error>(styles)
        })?;
        self.package.check_execution()?;
        self.package.source_version()?;
        Ok(styles)
    }
}

struct SourceCheckedTextSink<'owner, 'output, W: ?Sized> {
    output: &'output mut W,
    owner: &'owner SourceInner,
    failure: Arc<Mutex<Option<Error>>>,
}

impl<'owner, 'output, W: Write + ?Sized> SourceCheckedTextSink<'owner, 'output, W> {
    fn record_failure(&self, error: Error) {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error);
        }
    }

    fn check(&self) -> io::Result<()> {
        match check_text_state(self.owner) {
            Ok(()) => Ok(()),
            Err(error) => {
                let message = error.to_string();
                self.record_failure(error);
                Err(io::Error::other(message))
            },
        }
    }
}

impl<W: Write + ?Sized> Write for SourceCheckedTextSink<'_, '_, W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.check()?;
        let result = self.output.write(bytes);
        let _ = self.check();
        result
    }

    fn flush(&mut self) -> io::Result<()> {
        self.check()?;
        let result = self.output.flush();
        let _ = self.check();
        result
    }
}

fn take_source_text_failure(failure: &Arc<Mutex<Option<Error>>>) -> Option<Error> {
    failure
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take()
}

#[derive(Default)]
struct FallibleTextCollector {
    bytes: Vec<u8>,
    allocation: Option<Error>,
}

impl FallibleTextCollector {
    fn push_terminal_newline(&mut self, current_bytes: u64, max_output_bytes: u64) -> Result<()> {
        let required = current_bytes
            .checked_add(1)
            .ok_or(Error::CapacityOverflow {
                resource: "XLSB text output",
            })?;
        if required > max_output_bytes {
            return Err(Error::CapacityOverflow {
                resource: "XLSB text output",
            });
        }
        self.bytes
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "XLSB text output",
                source,
            })?;
        self.bytes.push(b'\n');
        Ok(())
    }

    fn take_bytes(self) -> Vec<u8> {
        self.bytes
    }
}

impl Write for FallibleTextCollector {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if let Err(source) = self.bytes.try_reserve(bytes.len()) {
            let error = Error::Allocation {
                resource: "XLSB text output",
                source,
            };
            let message = error.to_string();
            self.allocation = Some(error);
            return Err(io::Error::other(message));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn check_text_state(inner: &SourceInner) -> Result<()> {
    inner.package.check_execution()?;
    inner.package.source_version()?;
    Ok(())
}

fn map_text_output_error(error: TextOutputError<Error>, allocation: Option<Error>) -> Error {
    match error {
        TextOutputError::Document { source, .. } => source,
        TextOutputError::Limit { .. } => Error::CapacityOverflow {
            resource: "XLSB text output",
        },
        TextOutputError::Sink { source, .. } => allocation.unwrap_or(Error::Io(source)),
        TextOutputError::NonDeterministicFragments { .. } => {
            Error::InvalidFormat("XLSB text fragments were not deterministic".to_string())
        },
        _ => Error::InvalidFormat("unsupported XLSB text output error".to_string()),
    }
}

fn worksheet_dimensions(worksheet: &Worksheet) -> Result<(usize, usize)> {
    let (min_row, min_column, max_row, max_column) =
        SheetWorksheet::dimensions(worksheet).unwrap_or((0, 0, 0, 0));
    if min_row > max_row || min_column > max_column {
        return Err(Error::InvalidFormat(
            "XLSB worksheet dimensions are reversed".to_string(),
        ));
    }
    if max_row > XLSB_MAX_ROW_INDEX {
        return Err(Error::InvalidFormat(
            "XLSB worksheet row dimension exceeds the BIFF12 bound".to_string(),
        ));
    }
    if max_column > XLSB_MAX_COLUMN_INDEX {
        return Err(Error::InvalidFormat(
            "XLSB worksheet column dimension exceeds the BIFF12 bound".to_string(),
        ));
    }
    let row_count = max_row.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "XLSB text row count",
    })?;
    let column_count = max_column.checked_add(1).ok_or(Error::CapacityOverflow {
        resource: "XLSB text column count",
    })?;
    let row_count = usize::try_from(row_count).map_err(|error| {
        let _ = error;
        Error::CapacityOverflow {
            resource: "XLSB text row count",
        }
    })?;
    let column_count = usize::try_from(column_count).map_err(|error| {
        let _ = error;
        Error::CapacityOverflow {
            resource: "XLSB text column count",
        }
    })?;
    Ok((row_count, column_count))
}

fn count_formatted_text<F>(render: &F) -> Result<usize>
where
    F: Fn(&mut dyn fmt::Write) -> fmt::Result,
{
    let mut counter = TextByteCounter { bytes: 0 };
    render(&mut counter).map_err(|error| {
        let _ = error;
        Error::CapacityOverflow {
            resource: "XLSB text row",
        }
    })?;
    Ok(counter.bytes)
}

fn count_cell_text(value: &litchi_core::sheet::CellValue) -> Result<usize> {
    use litchi_core::sheet::CellValue;

    match value {
        CellValue::Empty => Ok(0),
        CellValue::Bool(value) => Ok(if *value { 4 } else { 5 }),
        CellValue::Int(value) => {
            count_formatted_text(&|writer| fmt::write(writer, format_args!("{value}")))
        },
        CellValue::Float(value) | CellValue::DateTime(value) => {
            count_formatted_text(&|writer| fmt::write(writer, format_args!("{value}")))
        },
        CellValue::String(value) | CellValue::Error(value) => Ok(value.len()),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(CellValue::Empty) | None => count_formatted_text(&|writer| {
                fmt::Write::write_char(writer, '=')?;
                fmt::Write::write_str(writer, formula)
            }),
            Some(value) => count_cell_text(value),
        },
    }
}

fn count_text_row(worksheet: &Worksheet, row: u32, column_count: usize) -> Result<usize> {
    let mut bytes = 0_usize;
    for column_index in 0..column_count {
        if column_index != 0 {
            bytes = bytes.checked_add(1).ok_or(Error::CapacityOverflow {
                resource: "XLSB text row",
            })?;
        }
        if let Some(cell) = worksheet.get_cell(
            row,
            u32::try_from(column_index).map_err(|error| {
                let _ = error;
                Error::CapacityOverflow {
                    resource: "XLSB text column",
                }
            })?,
        ) {
            bytes = bytes
                .checked_add(count_cell_text(SheetCell::value(cell))?)
                .ok_or(Error::CapacityOverflow {
                    resource: "XLSB text row",
                })?;
        }
    }
    Ok(bytes)
}

fn write_text_worksheet<W: Write + ?Sized>(
    owner: &SourceInner,
    worksheet: &Worksheet,
    writer: &mut SequentialTextWriter<'_, '_, W>,
    options: TextOutputOptions<'_>,
) -> std::result::Result<(), TextOutputError<Error>> {
    const MAX_LIMIT_PROBE_ROW_BYTES: usize = 1024 * 1024;

    let (row_count, column_count) =
        worksheet_dimensions(worksheet).map_err(|source| writer.document_error(source))?;
    for row_index in 0..row_count {
        check_text_state(owner).map_err(|source| writer.document_error(source))?;
        let row = u32::try_from(row_index).map_err(|error| {
            let _ = error;
            writer.document_error(Error::CapacityOverflow {
                resource: "XLSB text row",
            })
        })?;
        let row_bytes = count_text_row(worksheet, row, column_count)
            .map_err(|source| writer.document_error(source))?;
        check_text_state(owner).map_err(|source| writer.document_error(source))?;
        if row_bytes == 0 && !options.include_empty_objects() {
            continue;
        }
        let progress = writer.progress();
        let row_bytes_u64 = u64::try_from(row_bytes).map_err(|error| {
            let _ = error;
            writer.document_error(Error::CapacityOverflow {
                resource: "XLSB text row",
            })
        })?;
        let separator_bytes = if progress.objects_written() == 0 {
            0
        } else {
            u64::try_from(options.paragraph_separator().len()).map_err(|error| {
                let _ = error;
                writer.document_error(Error::CapacityOverflow {
                    resource: "XLSB text separator",
                })
            })?
        };
        let required = progress
            .bytes_written()
            .checked_add(separator_bytes)
            .and_then(|bytes| bytes.checked_add(row_bytes_u64))
            .ok_or_else(|| {
                writer.document_error(Error::CapacityOverflow {
                    resource: "XLSB text output",
                })
            })?;
        if required > options.max_output_bytes() && row_bytes > MAX_LIMIT_PROBE_ROW_BYTES {
            return Err(writer.document_error(Error::CapacityOverflow {
                resource: "XLSB text row exceeds configured output capacity",
            }));
        }
        let objects_required = progress.objects_written().checked_add(1).ok_or_else(|| {
            writer.document_error(Error::CapacityOverflow {
                resource: "XLSB text object count",
            })
        })?;
        if objects_required > options.max_objects() && row_bytes > MAX_LIMIT_PROBE_ROW_BYTES {
            return Err(writer.document_error(Error::CapacityOverflow {
                resource: "XLSB text row exceeds configured object capacity",
            }));
        }
        let mut value = String::new();
        value.try_reserve(row_bytes).map_err(|source| {
            writer.document_error(Error::Allocation {
                resource: "XLSB text row",
                source,
            })
        })?;
        for column_index in 0..column_count {
            if column_index != 0 {
                value.push('\t');
            }
            let column = u32::try_from(column_index).map_err(|error| {
                let _ = error;
                writer.document_error(Error::CapacityOverflow {
                    resource: "XLSB text row",
                })
            })?;
            if let Some(cell) = worksheet.get_cell(row, column) {
                append_cell_text(&mut value, SheetCell::value(cell))
                    .map_err(|source| writer.document_error(source))?;
            }
        }
        writer.write_object(TextObjectKind::Paragraph, &value)?;
        check_text_state(owner).map_err(|source| writer.document_error(source))?;
    }
    Ok(())
}

struct TextByteCounter {
    bytes: usize,
}

impl fmt::Write for TextByteCounter {
    fn write_str(&mut self, value: &str) -> fmt::Result {
        self.bytes = self.bytes.checked_add(value.len()).ok_or(fmt::Error)?;
        Ok(())
    }
}

fn append_counted_text<F>(output: &mut String, render: F) -> Result<()>
where
    F: Fn(&mut dyn fmt::Write) -> fmt::Result,
{
    let bytes = count_formatted_text(&render)?;
    output
        .try_reserve(bytes)
        .map_err(|source| Error::Allocation {
            resource: "XLSB text row",
            source,
        })?;
    render(output).map_err(|error| {
        let _ = error;
        Error::InvalidFormat("XLSB text formatting failed".to_string())
    })
}

fn append_cell_text(output: &mut String, value: &litchi_core::sheet::CellValue) -> Result<()> {
    use litchi_core::sheet::CellValue;

    match value {
        CellValue::Empty => Ok(()),
        CellValue::Bool(value) => append_counted_text(output, |writer| {
            fmt::Write::write_str(writer, if *value { "TRUE" } else { "FALSE" })
        }),
        CellValue::Int(value) => {
            append_counted_text(output, |writer| fmt::write(writer, format_args!("{value}")))
        },
        CellValue::Float(value) | CellValue::DateTime(value) => {
            append_counted_text(output, |writer| fmt::write(writer, format_args!("{value}")))
        },
        CellValue::String(value) | CellValue::Error(value) => {
            append_counted_text(output, |writer| fmt::Write::write_str(writer, value))
        },
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(CellValue::Empty) | None => append_counted_text(output, |writer| {
                fmt::Write::write_char(writer, '=')?;
                fmt::Write::write_str(writer, formula)
            }),
            Some(value) => append_cell_text(output, value),
        },
    }
}

fn optional_related_part(
    workbook_part: &PartView<'_>,
    relationship_types: &[&str],
    owner: &str,
) -> Result<Option<PackURI>> {
    let mut relationships = workbook_part
        .rels()
        .iter()
        .filter(|relationship| relationship_types.contains(&relationship.reltype()));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(Error::InvalidRelationship(format!(
            "XLSB workbook has multiple {owner} relationships"
        )));
    }
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(format!(
            "XLSB workbook {owner} relationship is external"
        )));
    }
    relationship.target_partname().map(Some).map_err(Into::into)
}

fn require_content_type(part: &PartView<'_>, expected: &str) -> Result<()> {
    if part.content_type() != expected {
        return Err(Error::InvalidContentType {
            expected: expected.to_string(),
            got: part.content_type().to_string(),
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        clippy::unwrap_used,
        reason = "single-flight tests use panic-on-failure thread extraction"
    )]

    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::sync::{Arc, Barrier};
    use std::thread;

    use super::SemanticCache;

    #[test]
    fn concurrent_callers_share_one_successful_initialization() {
        const CALLERS: usize = 8;
        let cache = Arc::new(SemanticCache::new());
        let ready = Arc::new(Barrier::new(CALLERS));
        let loader_started = Arc::new(Barrier::new(2));
        let load_count = Arc::new(AtomicUsize::new(0));
        let handles = (0..CALLERS)
            .map(|_| {
                let cache = Arc::clone(&cache);
                let ready = Arc::clone(&ready);
                let loader_started = Arc::clone(&loader_started);
                let load_count = Arc::clone(&load_count);
                thread::spawn(move || {
                    ready.wait();
                    cache.get_or_try_init(|| {
                        if load_count.fetch_add(1, Ordering::SeqCst) == 0 {
                            loader_started.wait();
                        }
                        Ok::<Arc<usize>, &'static str>(Arc::new(42))
                    })
                })
            })
            .collect::<Vec<_>>();
        loader_started.wait();

        let values = handles
            .into_iter()
            .map(|handle| handle.join().unwrap().unwrap())
            .collect::<Vec<_>>();
        let first = values.first().expect("callers should return values");
        assert!(values.iter().all(|value| Arc::ptr_eq(first, value)));
        assert!(values.iter().all(|value| **value == 42));
        assert_eq!(load_count.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn failed_initialization_is_not_retained_and_can_retry() {
        let cache = SemanticCache::new();
        let load_count = AtomicUsize::new(0);
        let first = cache.get_or_try_init(|| {
            load_count.fetch_add(1, Ordering::SeqCst);
            Err::<Arc<usize>, _>("synthetic parse failure")
        });
        assert!(matches!(first, Err("synthetic parse failure")));

        let retry = cache
            .get_or_try_init(|| {
                load_count.fetch_add(1, Ordering::SeqCst);
                Ok::<Arc<usize>, &'static str>(Arc::new(9))
            })
            .unwrap();
        assert_eq!(*retry, 9);
        assert_eq!(load_count.load(Ordering::SeqCst), 2);
    }

    #[test]
    fn separate_cache_instances_initialize_independently() {
        let first = SemanticCache::new();
        let second = SemanticCache::new();
        let first_value = first
            .get_or_try_init(|| Ok::<Arc<usize>, &'static str>(Arc::new(1)))
            .unwrap();
        let second_value = second
            .get_or_try_init(|| Ok::<Arc<usize>, &'static str>(Arc::new(2)))
            .unwrap();

        assert_eq!(*first_value, 1);
        assert_eq!(*second_value, 2);
        assert!(!Arc::ptr_eq(&first_value, &second_value));
    }

    #[test]
    fn source_backed_semantic_caches_reuse_production_arcs() {
        let bytes = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/Simple.xlsb"
        ))
        .unwrap();
        let workbook = super::SourceBackedWorkbook::from_read_at(Arc::new(
            litchi_core::OwnedSource::new(bytes),
        ))
        .unwrap();
        let expected_strings = workbook.inner.shared_strings().unwrap();
        let expected_styles = workbook.inner.styles().unwrap();
        let handles = (0..8)
            .map(|_| {
                let inner = Arc::clone(&workbook.inner);
                thread::spawn(move || (inner.shared_strings().unwrap(), inner.styles().unwrap()))
            })
            .collect::<Vec<_>>();

        for handle in handles {
            let (strings, styles) = handle.join().unwrap();
            assert!(Arc::ptr_eq(&expected_strings, &strings));
            assert!(Arc::ptr_eq(&expected_styles, &styles));
        }
    }
}
