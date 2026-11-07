//! Deferred, source-backed XLSB worksheet catalog and materialization.

use std::collections::HashSet;
use std::fmt;
use std::io::{self, Cursor, Read, Write};
use std::ops::ControlFlow;
use std::sync::{Arc, Mutex};

use litchi_core::sheet::{Cell as SheetCell, CellValue, Worksheet as SheetWorksheet};
use litchi_core::{
    ReadAt, SequentialTextWriter, SourceVersion, TextObjectKind, TextOutputError,
    TextOutputOptions, TextOutputReport,
};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{PackURI, PartData, PartView, SourceBackedPackage, SourceCacheDiagnostics};
use once_cell::sync::OnceCell;

use super::cache::{Limits as WorksheetCacheLimits, Stats, WorksheetCache};
use super::catalog::{Kind, State};
use super::model::Workbook;
use crate::external_link::{ExternalLinkLimits, Link};
use crate::leniency::{Leniency, Tolerance, ToleranceReport};
use crate::package::error::{Error, Result};
use crate::package::formula::{Context, Definition, ExternalBook};
use crate::package::shared_strings::SharedString;
use crate::package::styles_table::StylesTable;
use crate::package::table::Table;
use crate::raw::{Records, kind};
use crate::sheet::{Grid, Selector, Worksheet};

const XLSB_WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";
const XLSB_CHARTSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.chartsheet";
const XLSB_DIALOGSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.dialogsheet";
const XLSB_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.macrosheet";
const XLSB_INTL_MACROSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.intlmacrosheet";
const XLSB_MAX_ROW_INDEX: u32 = 1_048_575;
const XLSB_MAX_COLUMN_INDEX: u32 = 16_383;
const XLSB_SHARED_STRINGS_CONTENT_TYPE: &str = "application/vnd.ms-excel.sharedStrings";
const XLSB_STYLES_CONTENT_TYPE: &str = "application/vnd.ms-excel.styles";
const XLSB_TABLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.table";
const XLSB_EXTERNAL_LINK_CONTENT_TYPE: &str = "application/vnd.ms-excel.externalLink";
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

// Scalar-only fallback grids have a separate, deliberately smaller budget
// from complete worksheet snapshots. A source owner therefore cannot retain
// two full default caches merely by requesting both projections.
const CACHED_VALUE_CACHE_LIMITS: WorksheetCacheLimits = WorksheetCacheLimits {
    max_entries: 32,
    max_weight: 16 * 1024 * 1024,
};

#[derive(Debug)]
struct SheetMetadata {
    name: String,
    workbook_position: usize,
    partname: Option<PackURI>,
    kind: Kind,
    state: State,
    /// The compact worksheet repair report survives semantic-cache eviction.
    tolerance: OnceCell<ToleranceReport>,
    /// Repairs from the scalar-only indexed projection. A later successful
    /// full worksheet read supersedes this report in the aggregate accessor.
    cached_value_tolerance: OnceCell<ToleranceReport>,
    /// Serializes builders for this worksheet only; parsing never holds the
    /// owner-wide cache mutex.
    build: Mutex<()>,
    /// The first batch of the first cell cursor that validated this
    /// worksheet (record 0267), which later cursors start from. A few dozen
    /// cells.
    first_batch: OnceCell<Arc<cursor::FirstBatch>>,
}

#[derive(Debug)]
struct TablePartMetadata {
    worksheet_position: usize,
    partname: PackURI,
}

#[derive(Debug)]
struct ExternalLinkPartMetadata {
    partname: PackURI,
}

struct SourceInner {
    package: SourceBackedPackage,
    sheets: Vec<SheetMetadata>,
    worksheet_positions: Vec<usize>,
    table_parts: Vec<TablePartMetadata>,
    external_link_parts: Vec<ExternalLinkPartMetadata>,
    active_catalog_position: Option<usize>,
    formula_context: Context,
    external_link_limits: ExternalLinkLimits,
    raw_limits: crate::raw::Limits,
    shared_strings_part: Option<PackURI>,
    styles_part: Option<PackURI>,
    incomplete_formula_context: bool,
    is_1904_date_system: bool,
    shared_strings: SemanticCache<Vec<SharedString>>,
    styles: SemanticCache<StylesTable>,
    tables: SemanticCache<TableDefinitions>,
    external_books: SemanticCache<ExternalBooks>,
    worksheet_cache: WorksheetCache,
    /// Separate bounded retention for scalar-only indexed fallback grids.
    /// These snapshots never enter the full worksheet cache because they do
    /// not contain formulas, styles, or worksheet adjuncts.
    cached_value_cache: WorksheetCache,
    /// Workbook-scope repairs: the workbook stream at open, then each lazily
    /// parsed workbook-scope part once its semantic cache is initialized.
    tolerance: Mutex<ToleranceReport>,
    /// Opt-in repair policy applied to lazily read worksheet content.
    leniency: Leniency,
    /// The semantic budgets of the family reads of this owner.
    semantic_limits: crate::SemanticLimits,
}

/// What reading one worksheet stream needs ([`SourceBackedWorksheet`]).
struct MaterializationInputs {
    data: PartData,
    name: String,
    shared_strings: Arc<Vec<SharedString>>,
    formula_context: Context,
    workbook_position: usize,
    cell_xf_count: usize,
    comments_part: Option<PackURI>,
}

mod cursor;
mod features;
pub use cursor::CellCursor;

/// A bounded XLSB workbook catalog whose worksheet bodies remain deferred.
///
/// Opening reads OPC metadata and `workbook.bin`, but not ordinary worksheet,
/// shared-string, style, or table-definition payloads. The handle is immutable
/// and source-bound.
/// `SourceCacheLimits` bounds retained OPC payload bytes, and configured
/// per-Part read limits apply to each source payload. Successfully parsed
/// shared-string, style, table-definition, and bounded inert external-link
/// values are retained separately.
#[derive(Clone)]
pub struct SourceBackedWorkbook {
    inner: Arc<SourceInner>,
    /// Document-property edits published over the retained source
    /// ([`crate::metadata::source`]). The workbook parts are unchanged, so
    /// every semantic read and cache stays shared with the source owner.
    pub(crate) metadata: Option<Arc<crate::metadata::source::Revision>>,
}

/// A source-bound handle for one sheet in an XLSB workbook.
///
/// The handle contains only catalog metadata. Calling [`Self::materialize`]
/// reads the selected BIFF12 worksheet body, its shared-string/style inputs,
/// and all workbook table definitions needed for cross-sheet formula context;
/// non-worksheet handles return a typed capability error.
#[derive(Clone)]
pub struct SourceBackedWorksheet {
    inner: Arc<SourceInner>,
    catalog_position: usize,
}

// Shared handles: the unified operations layer keeps these in `Arc`-backed,
// `Send + Sync` file handles (record 0003 risk R3).
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SourceBackedWorkbook>();
    assert_send_sync::<SourceBackedWorksheet>();
    assert_send_sync::<Workbook>();
};

/// Cached cell values accessed through a verified worksheet binary index.
///
/// Construction materializes the selected worksheet and index parts and checks
/// their record framing. The handle retains those managed worksheet bytes;
/// subsequent lookups do not construct a complete worksheet cell table.
/// Formula results are the stored cache, without recalculation or formatting.
#[derive(Clone)]
pub struct SourceBackedIndexedWorksheet {
    inner: Arc<SourceInner>,
    worksheet: PartData,
    index: Arc<crate::binary_index::WorksheetBinaryIndex>,
    limits: crate::binary_index::Limits,
}

impl fmt::Debug for SourceBackedIndexedWorksheet {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedIndexedWorksheet")
            .field("worksheet_bytes", &self.worksheet.as_bytes().len())
            .finish_non_exhaustive()
    }
}

impl SourceBackedIndexedWorksheet {
    /// Read the stored scalar value at a zero-based row and column.
    ///
    /// An absent cell returns `None`; an explicit blank returns `Some(Empty)`.
    /// Formula cells return their cached scalar without reading formula context
    /// or evaluating it. Numeric values are raw numbers, without date or display
    /// formatting. A shared-string lookup loads the workbook's shared-string
    /// table on first use. Every call checks source freshness and cancellation.
    /// The string ceiling bounds the returned shared-string value; loading the
    /// shared-string table uses the workbook's existing ingress policy.
    pub fn cached_value(&self, row: u32, column: u32) -> Result<Option<CellValue>> {
        let _ = preflight_package(&self.inner.package)?;
        let result: Result<(Option<CellValue>, ToleranceReport)> = (|| {
            let record = self
                .index
                .lookup_record(self.worksheet.as_bytes(), row, column)?;
            let (value, report) = match record {
                Some((offset, payload)) => {
                    let mut report = ToleranceReport::default();
                    let value = crate::binary_index::decode_cached_value(
                        crate::raw::Kind::new(offset.record_kind)?,
                        payload,
                        self.limits.raw,
                        row,
                        &mut Tolerance::recording_with(&mut report, self.inner.leniency),
                    )?;
                    (Some(self.scalar_value(value)?), report)
                },
                None => (None, ToleranceReport::default()),
            };
            Ok((value, report))
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        let (value, report) = result?;
        // Publish repairs only after both the lookup and its freshness fence
        // have succeeded. A failed or stale lookup must not mutate the report.
        self.inner.lock_tolerance().append(&report);
        Ok(value)
    }

    /// Visit the stored scalar values in worksheet row-major order without
    /// constructing a complete [`Worksheet`]. Formula tokens remain opaque;
    /// each formula contributes only its stored cached scalar. Returning
    /// [`ControlFlow::Break`] stops the scan after the selected cell.
    pub fn visit_cached_values<B, F>(&self, visit: F) -> Result<ControlFlow<B>>
    where
        F: FnMut(crate::binary_index::CellOffset, CellValue) -> ControlFlow<B>,
    {
        let (result, report) = self.visit_cached_values_with_report(visit)?;
        // The public visitor retains its legacy behavior: repairs discovered
        // by an explicit scan join the workbook report immediately after its
        // own source fence has passed.
        self.inner.lock_tolerance().append(&report);
        Ok(result)
    }

    fn visit_cached_values_with_report<B, F>(
        &self,
        mut visit: F,
    ) -> Result<(ControlFlow<B>, ToleranceReport)>
    where
        F: FnMut(crate::binary_index::CellOffset, CellValue) -> ControlFlow<B>,
    {
        let _ = preflight_package(&self.inner.package)?;
        let mut report = ToleranceReport::default();
        let result = (|| {
            let mut in_sheet_data = false;
            let mut seen_sheet_data = false;
            let mut current_row = None;
            let mut record_count = 0usize;
            for record_result in Records::bounded(self.worksheet.as_bytes(), self.limits.raw)? {
                let record = record_result?;
                record_count = record_count.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "indexed worksheet record count",
                })?;
                if record_count > self.limits.max_records {
                    return Err(Error::limit(
                        crate::workbook::LimitKind::IndexedRecords,
                        record_count,
                        self.limits.max_records,
                    ));
                }
                match record.kind() {
                    kind::BEGIN_SHEET_DATA => {
                        if in_sheet_data || seen_sheet_data || !record.payload().is_empty() {
                            return Err(Error::InvalidFormat(
                                "indexed worksheet has invalid sheet-data framing",
                            ));
                        }
                        in_sheet_data = true;
                        seen_sheet_data = true;
                        current_row = None;
                    },
                    kind::END_SHEET_DATA => {
                        if !in_sheet_data || !record.payload().is_empty() {
                            return Err(Error::InvalidFormat(
                                "indexed worksheet has invalid sheet-data framing",
                            ));
                        }
                        in_sheet_data = false;
                        current_row = None;
                    },
                    kind::ROW_HDR if in_sheet_data => {
                        let bytes = record.payload().get(..4).ok_or(Error::InvalidFormat(
                            "indexed worksheet row header is truncated",
                        ))?;
                        let row = u32::from_le_bytes(bytes.try_into().map_err(|_error| {
                            Error::InvalidFormat("indexed worksheet row header is truncated")
                        })?);
                        if row > crate::binary_index::MAX_ROW_INDEX {
                            return Err(Error::InvalidFormat(
                                "indexed worksheet row is outside the BIFF12 grid",
                            ));
                        }
                        current_row = Some(row);
                    },
                    cell_kind if in_sheet_data && crate::binary_index::is_cell_kind(cell_kind) => {
                        let row = current_row.ok_or(Error::InvalidFormat(
                            "indexed worksheet cell has no row header",
                        ))?;
                        let bytes = record
                            .payload()
                            .get(..4)
                            .ok_or(Error::InvalidFormat("indexed worksheet cell is truncated"))?;
                        let column = u32::from_le_bytes(bytes.try_into().map_err(|_error| {
                            Error::InvalidFormat("indexed worksheet cell is truncated")
                        })?);
                        if column > crate::binary_index::MAX_COLUMN_INDEX {
                            return Err(Error::InvalidFormat(
                                "indexed worksheet column is outside the BIFF12 grid",
                            ));
                        }
                        let value = crate::binary_index::decode_cached_value(
                            cell_kind,
                            record.payload(),
                            self.limits.raw,
                            row,
                            &mut Tolerance::recording_with(&mut report, self.inner.leniency),
                        )?;
                        let value = self.scalar_value(value)?;
                        let offset = crate::binary_index::CellOffset {
                            row,
                            column,
                            offset: u64::try_from(record.offset()).map_err(|_error| {
                                Error::CapacityOverflow {
                                    resource: "indexed worksheet record offset",
                                }
                            })?,
                            record_kind: cell_kind.get(),
                        };
                        if let ControlFlow::Break(value) = visit(offset, value) {
                            return Ok(ControlFlow::Break(value));
                        }
                    },
                    _ => {},
                }
            }
            if !seen_sheet_data || in_sheet_data {
                return Err(Error::InvalidFormat(
                    "indexed worksheet sheet-data framing is incomplete",
                ));
            }
            Ok(ControlFlow::Continue(()))
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        Ok((result?, report))
    }

    fn scalar_value(&self, value: crate::cell_values::Value) -> Result<CellValue> {
        use crate::cell_values::{CellError, Value};

        Ok(match value {
            Value::Blank => CellValue::Empty,
            Value::RkNumber(value) | Value::Number(value) | Value::FormulaNumberCache(value) => {
                CellValue::Float(value)
            },
            Value::Boolean(value) | Value::FormulaBooleanCache(value) => CellValue::Bool(value),
            Value::InlineString(value) | Value::FormulaStringCache(value) => {
                CellValue::String(value)
            },
            Value::RichString(value) => CellValue::String(value.text),
            Value::SharedStringIndex(index) => {
                let strings = self.inner.shared_strings()?;
                let index = usize::try_from(index).map_err(|_error| {
                    Error::InvalidFormat("shared-string index exceeds platform size")
                })?;
                let string = strings.get(index).ok_or(Error::InvalidFormat(
                    "indexed cell refers to a missing shared string",
                ))?;
                let units = string.text.encode_utf16().count();
                if units > self.limits.raw.string_units() {
                    return Err(Error::limit(
                        crate::workbook::LimitKind::IndexedStringUnits,
                        units,
                        self.limits.raw.string_units(),
                    ));
                }
                let mut text = String::new();
                text.try_reserve_exact(string.text.len())
                    .map_err(|source| Error::Allocation {
                        resource: "indexed shared-string value",
                        source,
                    })?;
                text.push_str(&string.text);
                CellValue::String(text)
            },
            Value::Error(error) | Value::FormulaErrorCache(error) => {
                let text = match error {
                    CellError::Null => "#NULL!",
                    CellError::DivisionByZero => "#DIV/0!",
                    CellError::Value => "#VALUE!",
                    CellError::Reference => "#REF!",
                    CellError::Name => "#NAME?",
                    CellError::Number => "#NUM!",
                    CellError::NotAvailable => "#N/A",
                    CellError::GettingData => "#GETTING_DATA",
                };
                CellValue::Error(text.to_string())
            },
        })
    }
}

/// A source-checked, immutable view of one bounded inert external-link model.
///
/// The handle shares the already parsed semantic catalog without cloning DDE
/// or OLE cached matrices. It never follows, opens, refreshes, evaluates, or
/// executes the external target.
#[derive(Clone)]
pub struct SourceBackedExternalLink {
    books: Arc<[ExternalBook]>,
    index: usize,
}

impl SourceBackedExternalLink {
    /// Borrow the complete inert link metadata retained by this snapshot.
    #[must_use]
    pub fn metadata(&self) -> &Link {
        self.books[self.index].metadata_ref()
    }
}

impl SourceBackedWorkbook {
    /// Open with the default bounded OPC read and cache policies.
    pub fn from_read_at(source: Arc<dyn ReadAt>) -> Result<Self> {
        Self::from_read_at_with_options(source, crate::OpenOptions::default())
    }

    /// Open an XLSB package from a sequential reader with the default bounded
    /// OPC and external-link policies.
    ///
    /// The reader is consumed into the source-backed owner's bounded byte
    /// storage. Workbook metadata is validated at open, while worksheet,
    /// shared-string, style, table, and external-link payloads remain
    /// deferred until their selected semantic reads require them. The owned
    /// source has stable freshness identity for this immutable snapshot.
    pub fn from_reader<R: Read>(reader: R) -> Result<Self> {
        Self::from_reader_with_options(reader, crate::OpenOptions::default())
    }

    /// Open an XLSB package from a sequential reader under explicit source
    /// limits, external-link limits and leniency.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedOpenPolicy`] for skipped drawing decoding
    /// and for cache or execution policies (the owned sequential source keeps
    /// its bytes in memory and has no deferred-payload session; open through
    /// [`Self::from_read_at_with_options`] to use them), the OPC limit and
    /// archive errors, and the catalog validation errors.
    pub fn from_reader_with_options<R: Read>(
        reader: R,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_source()?;
        options.check_owning()?;
        Self::build(
            crate::package::ingress::source_from_reader(reader, options.source_limits)?,
            options.external_link_limits,
            options.leniency,
            options.semantic_limits,
            options.raw_limits,
        )
    }

    /// Open a positional source under explicit ingress policies: source,
    /// cache and external-link limits, an execution context and leniency.
    ///
    /// # Errors
    ///
    /// Returns [`Error::UnsupportedOpenPolicy`] for skipped drawing decoding
    /// (this owner defers every drawing on its own), the OPC limit, archive
    /// and execution errors, and the catalog validation errors.
    pub fn from_read_at_with_options(
        source: Arc<dyn ReadAt>,
        options: crate::OpenOptions,
    ) -> Result<Self> {
        options.check_source()?;
        let cache_limits = options.cache_limits.unwrap_or_default();
        let package = crate::package::ingress::source_from_read_at(
            source,
            options.source_limits,
            cache_limits,
            options.execution_context,
        )?;
        Self::build(
            package,
            options.external_link_limits,
            options.leniency,
            options.semantic_limits,
            options.raw_limits,
        )
    }

    /// Build a lazy XLSB catalog from an already validated deferred OPC package.
    pub fn from_source_backed_package(package: SourceBackedPackage) -> Result<Self> {
        Self::build(
            package,
            ExternalLinkLimits::default(),
            Leniency::Strict,
            crate::SemanticLimits::new(),
            crate::raw::Limits::DEFAULT,
        )
    }

    fn build(
        package: SourceBackedPackage,
        external_link_limits: ExternalLinkLimits,
        leniency: Leniency,
        semantic_limits: crate::SemanticLimits,
        raw_limits: crate::raw::Limits,
    ) -> Result<Self> {
        let _ = preflight_package(&package)?;
        let workbook_part = package.main_document_part()?;
        if workbook_part.content_type() != content_type::XLSB_BIN {
            return Err(Error::InvalidContentType {
                expected: content_type::XLSB_BIN,
            });
        }

        let workbook_data = workbook_part.data()?;
        let mut records = Records::new(workbook_data.as_bytes(), raw_limits);
        // The packaging repairs of the OPC read come first in discovery order.
        let mut tolerance = ToleranceReport::from_package(package.tolerance_report());
        let info = Workbook::read_workbook(&mut records, &mut tolerance, raw_limits)?;
        let active_catalog_position =
            (!info.worksheet_names.is_empty()).then_some(info.active_catalog_position.unwrap_or(0));
        if info.worksheet_names.len() != info.worksheet_rel_ids.len() {
            return Err(Error::InvalidRelationship(
                "XLSB sheet names and relationship identifiers have different lengths",
            ));
        }
        if info.worksheet_names.len() != info.worksheet_states.len() {
            return Err(Error::InvalidFormat(
                "XLSB sheet names and visibility states have different lengths",
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

        let external_link_budget = external_link_limits.budget();
        external_link_budget.preflight_links(info.external_link_rel_ids.len())?;
        let mut external_link_parts = Vec::new();
        external_link_parts
            .try_reserve_exact(info.external_link_rel_ids.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB external-link catalog",
                source,
            })?;
        for relationship_id in &info.external_link_rel_ids {
            let _ = preflight_package(&package)?;
            let relationship =
                workbook_part
                    .rels()
                    .get(relationship_id)
                    .ok_or(Error::InvalidRelationship(
                        "BrtSupBookSrc relationship is missing",
                    ))?;
            if relationship.is_external() {
                return Err(Error::InvalidRelationship(
                    "BrtSupBookSrc relationship is external",
                ));
            }
            if !matches!(
                relationship.reltype(),
                relationship_type::EXTERNAL_LINK | relationship_type::STRICT_EXTERNAL_LINK
            ) {
                return Err(Error::InvalidRelationship(
                    "BrtSupBookSrc relationship has invalid type",
                ));
            }
            let partname = relationship.target_partname()?;
            let part = package.part(&partname)?;
            require_content_type(&part, XLSB_EXTERNAL_LINK_CONTENT_TYPE)?;
            external_link_parts.push(ExternalLinkPartMetadata { partname });
            postflight_package(&package)?;
        }

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
        let mut table_parts = Vec::new();
        table_parts
            .try_reserve_exact(info.worksheet_names.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB table relationship catalog",
                source,
            })?;
        let mut sheet_targets = HashSet::new();
        sheet_targets
            .try_reserve(info.worksheet_names.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB worksheet target validation",
                source,
            })?;
        let mut incomplete_formula_context = false;

        for (workbook_position, (name, relationship_id)) in info
            .worksheet_names
            .iter()
            .zip(&info.worksheet_rel_ids)
            .enumerate()
        {
            package.check_execution()?;
            let state_value = info.worksheet_states[workbook_position];
            let state = State::from_hs_state(state_value)?;
            let Some(relationship_id) = relationship_id.as_deref() else {
                if state != State::VeryHidden {
                    return Err(Error::InvalidRelationship(
                        "XLSB unbound sheet tabs must be very hidden",
                    ));
                }
                sheets.push(SheetMetadata {
                    name: name.clone(),
                    workbook_position,
                    partname: None,
                    kind: Kind::Unbound,
                    state,
                    tolerance: OnceCell::new(),
                    cached_value_tolerance: OnceCell::new(),
                    build: Mutex::new(()),
                    first_batch: OnceCell::new(),
                });
                continue;
            };
            let relationship =
                workbook_part
                    .rels()
                    .get(relationship_id)
                    .ok_or(Error::InvalidRelationship(
                        "XLSB sheet relationship is missing",
                    ))?;
            if relationship.is_external() {
                return Err(Error::InvalidRelationship(
                    "XLSB sheet has an external relationship",
                ));
            }
            let partname = relationship.target_partname()?;
            if !sheet_targets.insert(partname.as_str().to_ascii_lowercase()) {
                return Err(Error::InvalidRelationship(
                    "multiple XLSB sheets resolve to one physical target",
                ));
            }
            let sheet_part = package.part(&partname)?;
            let sheet_kind = if matches!(
                relationship.reltype(),
                relationship_type::WORKSHEET | relationship_type::STRICT_WORKSHEET
            ) {
                let worksheet_position = worksheet_positions.len();
                require_content_type(&sheet_part, XLSB_WORKSHEET_CONTENT_TYPE)?;
                for relationship in sheet_part.rels().iter() {
                    if matches!(
                        relationship.reltype(),
                        relationship_type::TABLE | relationship_type::STRICT_TABLE
                    ) {
                        if relationship.is_external() {
                            return Err(Error::InvalidRelationship(
                                "XLSB worksheet has an external table relationship",
                            ));
                        }
                        let table_partname = relationship.target_partname()?;
                        let table_part = package.part(&table_partname)?;
                        require_content_type(&table_part, XLSB_TABLE_CONTENT_TYPE)?;
                        table_parts
                            .try_reserve(1)
                            .map_err(|source| Error::Allocation {
                                resource: "source-backed XLSB table relationship catalog",
                                source,
                            })?;
                        table_parts.push(TablePartMetadata {
                            worksheet_position,
                            partname: table_partname,
                        });
                    }
                    if matches!(
                        relationship.reltype(),
                        relationship_type::PIVOT_TABLE | relationship_type::STRICT_PIVOT_TABLE
                    ) {
                        incomplete_formula_context = true;
                    }
                }
                Kind::Worksheet
            } else {
                match relationship.reltype() {
                    XLSB_CHARTSHEET_RELATIONSHIP | XLSB_STRICT_CHARTSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_CHARTSHEET_CONTENT_TYPE)?;
                        Kind::Chart
                    },
                    XLSB_DIALOGSHEET_RELATIONSHIP | XLSB_STRICT_DIALOGSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_DIALOGSHEET_CONTENT_TYPE)?;
                        Kind::Dialog
                    },
                    XLSB_MACROSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_MACROSHEET_CONTENT_TYPE)?;
                        Kind::Macro
                    },
                    XLSB_INTL_MACROSHEET_RELATIONSHIP => {
                        require_content_type(&sheet_part, XLSB_INTL_MACROSHEET_CONTENT_TYPE)?;
                        Kind::IntlMacro
                    },
                    _ => {
                        return Err(Error::InvalidRelationship(
                            "XLSB sheet relationship type is unsupported",
                        ));
                    },
                }
            };
            let catalog_position = sheets.len();
            sheets.push(SheetMetadata {
                name: name.clone(),
                workbook_position,
                partname: Some(partname),
                kind: sheet_kind,
                state,
                tolerance: OnceCell::new(),
                cached_value_tolerance: OnceCell::new(),
                build: Mutex::new(()),
                first_batch: OnceCell::new(),
            });
            if matches!(sheet_kind, Kind::Worksheet) {
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
        postflight_package(&package)?;

        Ok(Self {
            metadata: None,
            inner: Arc::new(SourceInner {
                package,
                sheets,
                worksheet_positions,
                table_parts,
                external_link_parts,
                active_catalog_position,
                formula_context,
                external_link_limits,
                raw_limits,
                shared_strings_part,
                styles_part,
                incomplete_formula_context,
                is_1904_date_system,
                shared_strings: SemanticCache::new(),
                styles: SemanticCache::new(),
                tables: SemanticCache::new(),
                external_books: SemanticCache::new(),
                worksheet_cache: WorksheetCache::new(),
                cached_value_cache: WorksheetCache::with_limits(CACHED_VALUE_CACHE_LIMITS),
                tolerance: Mutex::new(tolerance),
                leniency,
                semantic_limits,
            }),
        })
    }

    /// The semantic budgets this owner keeps for its family reads.
    #[must_use]
    pub fn semantic_limits(&self) -> crate::SemanticLimits {
        self.inner.semantic_limits
    }

    /// A snapshot of every workbook-scope repair applied so far.
    ///
    /// The packaging repairs of the OPC read that admitted the source come
    /// first, as [`crate::Defect::Package`] entries. The workbook stream is
    /// parsed when the catalog opens; the styles part
    /// is parsed lazily the first time a worksheet materialization needs it,
    /// and its repairs join this report then, exactly once. Worksheet-stream
    /// repairs are reported by [`Worksheet::tolerance_report`] on each
    /// materialized worksheet. A scalar-only indexed fallback retains its
    /// compact report beside the selected catalog entry until a successful
    /// full worksheet read supersedes it. See [`crate::leniency`] for the
    /// repair contract.
    #[must_use]
    pub fn tolerance_report(&self) -> ToleranceReport {
        self.inner.lock_tolerance().clone()
    }

    /// Return the workbook report together with every worksheet repair that
    /// has been successfully published by this source-backed owner.
    ///
    /// Worksheet streams are parsed lazily. Their compact reports are
    /// retained beside the worksheet catalog, so cache eviction and a later
    /// materialization do not append the same report twice. A scalar-only
    /// indexed fallback has a parallel report slot; a successful full read
    /// takes precedence for that worksheet. A failed or cancelled read
    /// publishes neither its semantic value nor its report. The accessor does
    /// not parse source parts. Its source and execution fences run before and
    /// after the report snapshot, so callers cannot observe a report from a
    /// stale or cancelled source.
    #[must_use]
    pub fn read_tolerance_report(&self) -> Result<ToleranceReport> {
        let _ = preflight_package(&self.inner.package)?;
        let mut report = self.inner.lock_tolerance().clone();
        for sheet in &self.inner.sheets {
            if let Some(tolerance) = sheet
                .tolerance
                .get()
                .or_else(|| sheet.cached_value_tolerance.get())
            {
                report.append(tolerance);
            }
        }
        postflight_package(&self.inner.package)?;
        Ok(report)
    }

    /// Return the workbook and sheet-scoped defined names retained by the
    /// catalog. The name records are read from `workbook.bin`; no worksheet
    /// or table part is materialized by this accessor.
    pub fn defined_names(&self) -> Result<Arc<[String]>> {
        let _ = preflight_package(&self.inner.package)?;
        let names = Arc::clone(&self.inner.formula_context.defined_names);
        postflight_package(&self.inner.package)?;
        Ok(names)
    }

    /// Return formula-context table definitions, loading the bounded table
    /// parts through the leaf semantic cache when necessary.
    pub fn tables(&self) -> Result<Arc<[Definition]>> {
        self.inner.tables()
    }

    /// Return parsed ListObject table parts grouped by public worksheet
    /// ordinal. Each selected worksheet loads only its own table parts. The
    /// returned aggregate is caller-owned; the source owner does not retain a
    /// second unbounded all-table cache.
    pub fn structured_tables(&self) -> Result<Arc<[(usize, Table)]>> {
        let _ = preflight_package(&self.inner.package)?;
        let mut tables = Vec::new();
        tables
            .try_reserve_exact(self.inner.table_parts.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB structured tables",
                source,
            })?;
        for (worksheet_position, &catalog_position) in
            self.inner.worksheet_positions.iter().enumerate()
        {
            let sheet = SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            };
            for table in sheet.tables()?.iter() {
                tables.push((worksheet_position, table.clone()));
            }
        }
        postflight_package(&self.inner.package)?;
        Ok(tables.into())
    }

    /// Return the source owner's validated style table. Styles are parsed at
    /// most once and remain owned by the leaf semantic cache.
    pub fn styles(&self) -> Result<Arc<StylesTable>> {
        self.inner.styles()
    }

    /// Content-free diagnostics for the owner-scoped worksheet cache.
    #[must_use]
    pub fn semantic_cache_diagnostics(&self) -> Stats {
        self.inner.worksheet_cache.stats()
    }

    /// Release parsed worksheet values that no grid or legacy handle pins.
    pub fn trim_semantic_cache(&self) -> usize {
        self.inner
            .worksheet_cache
            .trim_clean()
            .saturating_add(self.inner.cached_value_cache.trim_clean())
    }

    /// Return the number of external-link owners without reading their payloads.
    pub fn external_link_count(&self) -> Result<usize> {
        let _ = preflight_package(&self.inner.package)?;
        let count = self.inner.external_link_parts.len();
        postflight_package(&self.inner.package)?;
        Ok(count)
    }

    /// Return one bounded inert external-link snapshot in `BrtSupBookSrc` order.
    pub fn external_link(&self, index: usize) -> Result<Option<SourceBackedExternalLink>> {
        let _ = preflight_package(&self.inner.package)?;
        let books = self.inner.external_books()?;
        let link = if index < books.len() {
            Some(SourceBackedExternalLink { books, index })
        } else {
            None
        };
        postflight_package(&self.inner.package)?;
        Ok(link)
    }

    /// Snapshot all bounded inert external links in `BrtSupBookSrc` order.
    pub fn external_links(&self) -> Result<Vec<SourceBackedExternalLink>> {
        let _ = preflight_package(&self.inner.package)?;
        let books = self.inner.external_books()?;
        let mut links = Vec::new();
        links
            .try_reserve_exact(books.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB external-link handles",
                source,
            })?;
        links.extend((0..books.len()).map(|index| SourceBackedExternalLink {
            books: Arc::clone(&books),
            index,
        }));
        postflight_package(&self.inner.package)?;
        Ok(links)
    }

    /// Return the external-link semantic policy configured for this source-backed workbook.
    #[must_use]
    pub fn external_link_limits(&self) -> ExternalLinkLimits {
        self.inner.external_link_limits
    }

    /// Return the BIFF12 raw policy retained by this source-backed owner.
    #[must_use]
    pub fn raw_limits_value(&self) -> crate::raw::Limits {
        self.inner.raw_limits
    }

    /// Return the number of workbook tabs after checking source freshness.
    pub fn sheet_count(&self) -> Result<usize> {
        let _ = preflight_package(&self.inner.package)?;
        let count = self.inner.sheets.len();
        postflight_package(&self.inner.package)?;
        Ok(count)
    }

    /// Snapshot all workbook tab names in workbook order after checking source freshness.
    pub fn sheet_names(&self) -> Result<Vec<String>> {
        let _ = preflight_package(&self.inner.package)?;
        let mut names = Vec::new();
        names
            .try_reserve_exact(self.inner.sheets.len())
            .map_err(|source| Error::Allocation {
                resource: "source-backed XLSB sheet-name snapshot",
                source,
            })?;
        names.extend(self.inner.sheets.iter().map(|sheet| sheet.name.clone()));
        postflight_package(&self.inner.package)?;
        Ok(names)
    }

    /// Snapshot all checked workbook tab handles in workbook order.
    pub fn sheets(&self) -> Result<Vec<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
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
        postflight_package(&self.inner.package)?;
        Ok(sheets)
    }

    /// Select a workbook tab by its semantic name or checked zero-based
    /// position in complete workbook order.
    ///
    /// Names use the workbook's case-insensitive comparison policy. A missing
    /// name or position returns `Ok(None)`; every source and cancellation
    /// fence is checked around the lookup.
    pub fn sheet<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Option<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let catalog_position = match selector.into() {
            Selector::Position(position) => self
                .inner
                .sheets
                .get(position.get())
                .map(|_| position.get()),
            Selector::Name(name) => self.inner.sheets.iter().position(|sheet| {
                crate::package::formula::excel_name_eq(&sheet.name, name.as_ref())
            }),
        };
        let sheet = catalog_position.map(|catalog_position| SourceBackedWorksheet {
            inner: Arc::clone(&self.inner),
            catalog_position,
        });
        postflight_package(&self.inner.package)?;
        Ok(sheet)
    }

    /// Select a workbook tab by zero-based position in complete workbook order.
    pub fn sheet_by_index(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let sheet = self.inner.sheets.get(index).map(|_| SourceBackedWorksheet {
            inner: Arc::clone(&self.inner),
            catalog_position: index,
        });
        postflight_package(&self.inner.package)?;
        Ok(sheet)
    }

    /// Select a workbook tab by its exact workbook name.
    pub fn sheet_by_name(&self, name: &str) -> Result<Option<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let sheet = self
            .inner
            .sheets
            .iter()
            .position(|sheet| sheet.name == name)
            .map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            });
        postflight_package(&self.inner.package)?;
        Ok(sheet)
    }

    /// Return the number of worksheet owners after checking source freshness.
    pub fn worksheet_count(&self) -> Result<usize> {
        let _ = preflight_package(&self.inner.package)?;
        let count = self.inner.worksheet_positions.len();
        postflight_package(&self.inner.package)?;
        Ok(count)
    }

    /// Snapshot worksheet names in workbook order after checking source freshness.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        let _ = preflight_package(&self.inner.package)?;
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
        postflight_package(&self.inner.package)?;
        Ok(names)
    }

    /// Snapshot all checked worksheet handles in workbook order.
    pub fn worksheets(&self) -> Result<Vec<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
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
        postflight_package(&self.inner.package)?;
        Ok(worksheets)
    }

    /// Select a worksheet by zero-based position in the worksheet-only catalog.
    pub fn worksheet_by_index(&self, index: usize) -> Result<Option<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let worksheet =
            self.inner
                .worksheet_positions
                .get(index)
                .copied()
                .map(|catalog_position| SourceBackedWorksheet {
                    inner: Arc::clone(&self.inner),
                    catalog_position,
                });
        postflight_package(&self.inner.package)?;
        Ok(worksheet)
    }

    /// Select a worksheet by its exact workbook name.
    pub fn worksheet_by_name(&self, name: &str) -> Result<Option<SourceBackedWorksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let worksheet = self
            .inner
            .worksheet_positions
            .iter()
            .copied()
            .find(|&catalog_position| self.inner.sheets[catalog_position].name == name)
            .map(|catalog_position| SourceBackedWorksheet {
                inner: Arc::clone(&self.inner),
                catalog_position,
            });
        postflight_package(&self.inner.package)?;
        Ok(worksheet)
    }

    /// Return the primary workbook view's physical sheet-catalog position.
    pub fn active_catalog_position(&self) -> Result<Option<usize>> {
        let _ = preflight_package(&self.inner.package)?;
        let position = self.inner.active_catalog_position;
        postflight_package(&self.inner.package)?;
        Ok(position)
    }

    /// Return the primary workbook view's public worksheet ordinal, when it
    /// names a worksheet rather than a chart, dialog, or macro sheet.
    pub fn active_worksheet_index(&self) -> Result<Option<usize>> {
        let _ = preflight_package(&self.inner.package)?;
        let index = self
            .inner
            .active_catalog_position
            .and_then(|catalog_position| {
                self.inner
                    .worksheet_positions
                    .iter()
                    .position(|&position| position == catalog_position)
            });
        postflight_package(&self.inner.package)?;
        Ok(index)
    }

    /// Return the exact source identity and revision captured at open.
    pub fn source_version(&self) -> Result<SourceVersion> {
        let version = preflight_package(&self.inner.package)?;
        postflight_package(&self.inner.package)?;
        Ok(version)
    }

    /// Borrow the validated deferred OPC package this catalog reads from.
    ///
    /// This is the package indexed when the workbook was opened: no ZIP
    /// directory, content-type catalog, or relationship graph is rebuilt.
    /// Borrowed package services, such as
    /// [`SourceBackedPackage::source_artifact`], read through the same
    /// positional source and observe the same source-version fence.
    #[must_use]
    pub fn source_backed_package(&self) -> &SourceBackedPackage {
        &self.inner.package
    }

    /// Publish this snapshot to a sequential sink.
    ///
    /// An unedited catalog streams its source artifact byte for byte. After
    /// a document-property edit ([`Self::edit_metadata`]) the property
    /// parts, their package relationships and content types are replaced,
    /// and every other member keeps its source bytes and compression.
    ///
    /// # Errors
    ///
    /// Returns the source-change, execution and sink errors of the OPC
    /// publisher.
    pub fn write_to(&self, writer: impl Write) -> Result<()> {
        match &self.metadata {
            Some(revision) => revision
                .topology(&self.inner.package)
                .map_err(publication_error)?
                .publish_to_stream(writer)?,
            None => self.inner.package.source_artifact().write_to(writer)?,
        }
        Ok(())
    }

    /// The same source owner with another document-property revision.
    pub(crate) fn with_metadata_revision(
        &self,
        metadata: Option<Arc<crate::metadata::source::Revision>>,
    ) -> Self {
        Self {
            inner: Arc::clone(&self.inner),
            metadata,
        }
    }

    /// Whether this snapshot carries document-property edits over its
    /// source ([`Self::edit_metadata`]).
    #[must_use]
    pub fn has_metadata_edits(&self) -> bool {
        self.metadata.is_some()
    }

    /// Materialize the eager, editable [`Workbook`] from this catalog.
    ///
    /// The owning OPC graph is built from the package indexed at open,
    /// through the same positional source: no format detection, ZIP
    /// directory, or relationship index is repeated, and this lazy catalog
    /// stays usable. Ordinary Part payloads are read, so the cost is that of
    /// an eager open's payload reads. The workbook keeps this catalog's
    /// external-link limits and [`Leniency`].
    ///
    /// # Errors
    ///
    /// Returns the typed source-change error when the source changed after
    /// open, `OpcError::ManagedPackageMaterialization` for a catalog opened
    /// with an execution context, and the eager workbook's own validation,
    /// limit and allocation errors otherwise.
    pub fn to_workbook(&self) -> Result<Workbook> {
        preflight_package(&self.inner.package)?;
        let package = match &self.metadata {
            Some(revision) => revision
                .topology(&self.inner.package)
                .map_err(publication_error)?
                .to_opc_package()?,
            None => self.inner.package.to_opc_package()?,
        };
        let mut workbook = Workbook::from_opc_package_with_external_link_limits(
            package,
            self.inner.external_link_limits,
            self.inner.raw_limits,
        )?;
        workbook.leniency = self.inner.leniency;
        postflight_package(&self.inner.package)?;
        Ok(workbook)
    }

    /// Read the workbook's optional theme without materializing worksheets.
    ///
    /// This loads the selected theme XML and returns an immutable metadata
    /// view. Image relationships remain inert; their payloads are not read.
    /// Source freshness and cancellation are checked around the operation.
    /// Retain the returned view for repeated color and font queries.
    ///
    /// The owner's [`crate::SemanticLimits::theme_value`] bound the read: the
    /// selected part's declared size is checked before payload loading, and
    /// decoded XML bounds remain enforced by the theme owner.
    pub fn theme(&self) -> Result<Option<crate::theme::View>> {
        let limits = self.inner.semantic_limits.theme_value();
        let _ = preflight_package(&self.inner.package)?;
        let theme = crate::theme::read_source(&self.inner.package, limits);
        postflight_package(&self.inner.package)?;
        theme
    }

    /// Inventory embedded-object and embedded-package relationships without
    /// materializing the OPC package or any payload bytes.
    ///
    /// The returned catalog entries borrow this source-backed package. Their
    /// payload views remain deferred; reading a payload is an explicit,
    /// fallible operation on the returned entry rather than part of this
    /// inventory call.
    ///
    /// The owner's [`crate::SemanticLimits::embedded_value`] bound the
    /// inventory.
    pub fn embedded(&self) -> Result<Vec<litchi_ooxml_common::embedded::SourceEntry<'_>>> {
        let limits = &self.inner.semantic_limits.embedded_value();
        let _ = preflight_package(&self.inner.package)?;
        let entries = litchi_ooxml_common::embedded::scan_source(&self.inner.package, limits);
        postflight_package(&self.inner.package)?;
        Ok(entries?)
    }

    /// Return whether this workbook uses Excel's 1904 date system.
    pub fn is_1904_date_system(&self) -> Result<bool> {
        let _ = preflight_package(&self.inner.package)?;
        let is_1904 = self.inner.is_1904_date_system;
        postflight_package(&self.inner.package)?;
        Ok(is_1904)
    }

    /// The workbook's defined names in record order, with their scope and
    /// formula text. Only the workbook stream is read.
    ///
    /// # Errors
    ///
    /// Returns the source-change, execution and decoding errors of the
    /// workbook stream.
    pub fn defined_name_entries(&self) -> Result<Vec<crate::named_ranges::Entry>> {
        let _ = preflight_package(&self.inner.package)?;
        let data = self.inner.package.main_document_part()?.data()?;
        let entries = crate::named_ranges::read_entries(
            data.as_bytes(),
            &self.inner.formula_context,
            self.inner.raw_limits,
        );
        postflight_package(&self.inner.package)?;
        entries
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
            for &catalog_position in &self.inner.worksheet_positions {
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
            Error::Encoding("XLSB text output was not valid UTF-8")
        })
    }
}

impl SourceBackedWorksheet {
    /// Open this worksheet's binary index for stored scalar value lookups.
    ///
    /// Returns `None` when the worksheet has no index relationship. An invalid
    /// or stale index is an error, rather than evidence that a cell is absent.
    /// The selected worksheet and index parts are materialized and validated
    /// once. Styles, formulas, drawings, and other sheets remain deferred.
    ///
    /// ```no_run
    /// use litchi_xlsb::SourceBackedWorkbook;
    ///
    /// # fn read_one(path: &str) -> Result<(), Box<dyn std::error::Error>> {
    /// let workbook = SourceBackedWorkbook::from_reader(std::fs::File::open(path)?)?;
    /// if let Some(sheet) = workbook.worksheet_by_name("Sheet1")?
    ///     && let Some(values) = sheet.indexed_values()?
    /// {
    ///     let a1 = values.cached_value(0, 0)?;
    ///     let b2 = values.cached_value(1, 1)?;
    ///     // Reuse `values` for further lookups in this immutable source.
    /// #   let _ = (a1, b2);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// The owner's [`crate::SemanticLimits::binary_index_value`] bound the
    /// worksheet and index validation.
    pub fn indexed_values(&self) -> Result<Option<SourceBackedIndexedWorksheet>> {
        let limits = self.inner.semantic_limits.binary_index_value();
        let _ = preflight_package(&self.inner.package)?;
        let result = (|| {
            let limits = limits.validate()?;
            let metadata = self.metadata();
            if !matches!(metadata.kind, Kind::Worksheet) {
                return Err(Error::UnsupportedFeature(
                    "binary-index value lookup requires an XLSB worksheet",
                ));
            }
            let partname = self.physical_partname()?;
            let part = self.inner.package.part(partname)?;
            let Some(index_uri) = optional_related_part(
                &part,
                &["http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex"],
                "worksheet binary index",
            )?
            else {
                return Ok(None);
            };
            let index_part = self.inner.package.part(&index_uri)?;
            require_content_type(&index_part, "application/vnd.ms-excel.binIndexWs")?;
            if !index_part.rels().is_empty() {
                return Err(Error::InvalidRelationship(
                    "worksheet binary index part must not have relationships",
                ));
            }
            preflight_indexed_part(
                &index_part,
                limits
                    .source_bytes
                    .min(crate::binary_index::MAX_INDEX_STREAM_BYTES),
                crate::binary_index::LimitKind::Bytes.into(),
            )?;
            preflight_indexed_part(
                &part,
                limits.source_bytes,
                crate::workbook::LimitKind::IndexedPartBytes.into(),
            )?;
            let index_data = index_part.data()?;
            let worksheet = part.data()?;
            let _ = preflight_package(&self.inner.package)?;
            let index = crate::binary_index::WorksheetBinaryIndex::from_parts(
                index_data.as_bytes(),
                worksheet.as_bytes(),
                limits,
            )?;
            Ok(Some(SourceBackedIndexedWorksheet {
                inner: Arc::clone(&self.inner),
                worksheet,
                index: Arc::new(index),
                limits,
            }))
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        result
    }

    /// Return this sheet's name after checking source freshness.
    pub fn name(&self) -> Result<&str> {
        let _ = preflight_package(&self.inner.package)?;
        let name = &self.metadata().name;
        postflight_package(&self.inner.package)?;
        Ok(name)
    }

    /// Return this sheet's zero-based position in the complete workbook sheet order.
    pub fn workbook_position(&self) -> Result<usize> {
        let _ = preflight_package(&self.inner.package)?;
        let position = self.metadata().workbook_position;
        postflight_package(&self.inner.package)?;
        Ok(position)
    }

    /// Return this tab's semantic kind after checking source freshness.
    ///
    /// Relationship identifiers and part names used to resolve the kind are
    /// retained only by the source-backed owner; callers receive the semantic
    /// projection in [`crate::sheet::Kind`].
    pub fn kind(&self) -> Result<Kind> {
        let _ = preflight_package(&self.inner.package)?;
        let kind = self.metadata().kind;
        postflight_package(&self.inner.package)?;
        Ok(kind)
    }

    /// Return this tab's semantic visibility state after checking source
    /// freshness.
    pub fn state(&self) -> Result<State> {
        let _ = preflight_package(&self.inner.package)?;
        let state = self.metadata().state;
        postflight_package(&self.inner.package)?;
        Ok(state)
    }

    /// Resolve one materialized cell's local style through the source-owned
    /// stylesheet cache. The cell and stylesheet remain tied to this source
    /// generation; no eager workbook reconstruction is performed.
    /// The workbook styles table, parsed once per owner on first use.
    ///
    /// # Errors
    ///
    /// Returns the source-change, execution and styles-part errors.
    pub fn styles(&self) -> Result<Arc<StylesTable>> {
        let _ = preflight_package(&self.inner.package)?;
        let styles = self.inner.styles()?;
        postflight_package(&self.inner.package)?;
        Ok(styles)
    }

    pub fn style_for_cell(
        &self,
        cell: &crate::package::Cell,
    ) -> Result<Option<crate::package::styles_table::CellFormat>> {
        let _ = preflight_package(&self.inner.package)?;
        let styles = self.inner.styles()?;
        let style = styles.get_cell_format(cell.style_id() as usize).cloned();
        postflight_package(&self.inner.package)?;
        Ok(style)
    }

    /// Read the ListObject table parts owned by this worksheet only.
    ///
    /// Table parts remain deferred and are parsed through the leaf on each
    /// selected read. The source owner deliberately does not retain an
    /// unbounded second table-value cache: the caller owns the returned Arc,
    /// while malformed or unrelated table parts remain untouched. The table
    /// codec's strict parser is used here, so a failed read never publishes a
    /// partial table collection or a duplicate tolerance report.
    pub fn tables(&self) -> Result<Arc<[Table]>> {
        let _ = preflight_package(&self.inner.package)?;
        let result = (|| {
            if !matches!(self.metadata().kind, Kind::Worksheet) {
                return Err(Error::UnsupportedFeature(
                    "source-backed XLSB tables require a worksheet tab",
                ));
            }
            let worksheet_position = self
                .inner
                .worksheet_positions
                .binary_search(&self.catalog_position)
                .map_err(|_error| {
                    Error::InvalidRelationship(
                        "source-backed XLSB worksheet catalog is internally inconsistent",
                    )
                })?;
            let count = self
                .inner
                .table_parts
                .iter()
                .filter(|part| part.worksheet_position == worksheet_position)
                .count();
            let mut tables = Vec::new();
            tables
                .try_reserve_exact(count)
                .map_err(|source| Error::Allocation {
                    resource: "source-backed XLSB selected worksheet tables",
                    source,
                })?;
            for table_part in self
                .inner
                .table_parts
                .iter()
                .filter(|part| part.worksheet_position == worksheet_position)
            {
                let _ = preflight_package(&self.inner.package)?;
                let part = self.inner.package.part(&table_part.partname)?;
                require_content_type(&part, XLSB_TABLE_CONTENT_TYPE)?;
                let data = part.data()?;
                let table = crate::package::table::parse_table_part(
                    data.as_bytes(),
                    self.inner.raw_limits,
                )?;
                postflight_package(&self.inner.package)?;
                tables.push(table);
            }
            Ok(tables.into())
        })();
        let fence = postflight_package(&self.inner.package);
        fence?;
        result
    }

    /// Parse the selected worksheet BIFF12 stream without reading unselected sheets.
    ///
    /// This source-backed layer materializes stream-owned worksheet data and
    /// all table definitions and bounded inert external-link metadata needed
    /// for formula resolution. Other workbook adjunct owners such as drawings,
    /// PivotTables, slicers, and timelines remain separately deferred rather
    /// than forcing eager loading.
    pub fn materialize(&self) -> Result<Worksheet> {
        Ok((*self.cached_worksheet()?).clone())
    }

    /// Return this worksheet as an immutable, cheaply cloned grid snapshot.
    ///
    /// The owner performs source and execution fences before returning the
    /// grid. Subsequent grid reads use only the captured worksheet value and
    /// therefore remain valid after cache eviction or owner drop.
    pub fn grid(&self) -> Result<Grid> {
        self.cached_worksheet().map(Grid::from_arc)
    }

    /// Build a scalar-only grid from the worksheet's binary cached-value
    /// index. This is the explicit source fallback for workbooks whose full
    /// worksheet materialization is refused because of an unsupported
    /// deferred adjunct such as a PivotTable, slicer, or timeline.
    ///
    /// The fallback performs one bounded row-major index scan and creates
    /// ordinary scalar cells with [`crate::package::Cell::new`]. It never invents formula
    /// text, styles, comments, or other worksheet metadata. A missing binary
    /// index is a typed unsupported-feature error.
    pub fn cached_grid(&self) -> Result<Grid> {
        let _ = preflight_package(&self.inner.package)?;
        let metadata = self.metadata();
        if !matches!(metadata.kind, Kind::Worksheet) {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB cached values require a worksheet tab",
            ));
        }
        // The cache hit still crosses the source fence before the snapshot
        // escapes. The build lock serializes full and scalar-only builds for
        // this catalog position without holding the cache mutex during scan.
        let _build = metadata
            .build
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(worksheet) = self
            .inner
            .cached_value_cache
            .get(self.catalog_position, true)
        {
            postflight_package(&self.inner.package)?;
            return Ok(Grid::from_arc(worksheet));
        }
        let indexed = self.indexed_values()?.ok_or({
            Error::UnsupportedFeature(
                "source-backed XLSB worksheet has no indexed cached-value projection",
            )
        })?;
        let mut worksheet = Worksheet::new(metadata.name.clone());
        let (_, report) = indexed.visit_cached_values_with_report(|offset, value| {
            worksheet.add_cell(crate::package::Cell::new(offset.row, offset.column, value));
            ControlFlow::<()>::Continue(())
        })?;
        worksheet.set_tolerance_report(report.clone());
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        let candidate = Arc::new(worksheet);
        let weight = candidate.retained_bytes();
        postflight_package(&self.inner.package)?;
        let canonical = self.inner.cached_value_cache.publish(
            self.catalog_position,
            Arc::clone(&candidate),
            weight,
        );
        if let Err(error) = postflight_package(&self.inner.package) {
            if Arc::ptr_eq(&canonical, &candidate) {
                self.inner
                    .cached_value_cache
                    .discard_if(self.catalog_position, &candidate);
            }
            return Err(error);
        }
        // Keep fallback repairs only after the candidate has crossed the
        // outer publication fence. If a later full worksheet read succeeds,
        // `read_tolerance_report` prefers its report for this sheet.
        self.metadata()
            .cached_value_tolerance
            .get_or_init(|| report);
        Ok(Grid::from_arc(canonical))
    }

    /// Return a pinned, immutable worksheet snapshot from the owner-scoped
    /// weighted cache. The source and execution fences run on cache hits as
    /// well as misses; parsing is never performed while the cache mutex is
    /// held.
    pub(crate) fn cached_worksheet(&self) -> Result<Arc<Worksheet>> {
        let _ = preflight_package(&self.inner.package)?;
        let metadata = self.metadata();
        if !matches!(metadata.kind, Kind::Worksheet) {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB non-worksheet materialization is not supported",
            ));
        }
        let partname = self.physical_partname()?.clone();
        // Resolving the physical part is intentional on hits: it verifies the
        // captured source snapshot before any retained semantic value escapes.
        let _part = self.inner.package.part(&partname)?;
        // The per-sheet build lock also covers the cache hit path. A builder
        // publishes the value before its final fence but initializes the
        // retained report immediately after that fence; taking this lock
        // before looking in the cache prevents a concurrent reader from
        // observing that provisional interval.
        let _build = self
            .metadata()
            .build
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner);
        if let Some(worksheet) = self.inner.worksheet_cache.get(self.catalog_position, true) {
            postflight_package(&self.inner.package)?;
            return Ok(worksheet);
        }

        let mut worksheet = self.parse_materialized(partname)?;
        self.inner.package.check_execution()?;
        self.inner.package.source_version()?;
        let report = worksheet.tolerance_report().clone();
        if let Some(retained_report) = self.metadata().tolerance.get() {
            worksheet.set_tolerance_report(retained_report.clone());
        }
        let candidate = Arc::new(worksheet);
        let weight = candidate.retained_bytes();
        // Check once before publication and once after it. The second fence
        // closes the race in which the source changes between the first check
        // and the cache publication.
        postflight_package(&self.inner.package)?;
        let canonical = self.inner.worksheet_cache.publish(
            self.catalog_position,
            Arc::clone(&candidate),
            weight,
        );
        if let Err(error) = postflight_package(&self.inner.package) {
            if Arc::ptr_eq(&canonical, &candidate) {
                self.inner
                    .worksheet_cache
                    .discard_if(self.catalog_position, &candidate);
            }
            return Err(error);
        }
        // This infallible publication is intentionally after the final
        // source/execution fence, so a refused candidate leaves no aggregate
        // worksheet report behind.
        self.metadata().tolerance.get_or_init(|| report);
        Ok(canonical)
    }

    fn parse_materialized(&self, partname: PackURI) -> Result<Worksheet> {
        let inputs = self.materialization_inputs(partname)?;
        let mut worksheet = Workbook::read_worksheet(
            Cursor::new(inputs.data.as_bytes()),
            inputs.name,
            inputs.shared_strings.as_slice(),
            &inputs.formula_context,
            inputs.workbook_position,
            inputs.cell_xf_count,
            self.inner.raw_limits,
            self.inner.leniency,
        )?;
        worksheet.set_scenarios(crate::package::scenarios::parse_worksheet(
            inputs.data.as_bytes(),
            self.inner.raw_limits,
        )?);
        if let Some(partname) = inputs.comments_part {
            for comment in self.read_comments(&partname, worksheet.tolerance_report_mut())? {
                worksheet.add_comment(comment);
            }
        }
        Ok(worksheet)
    }

    /// What reading the worksheet stream needs, with every refusal the
    /// materialization raises before it reads the stream (record 0267 shares
    /// this with the cell cursor).
    fn materialization_inputs(&self, partname: PackURI) -> Result<MaterializationInputs> {
        if self.inner.incomplete_formula_context {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB worksheet materialization requires deferred PivotTable formula owners",
            ));
        }
        let metadata = self.metadata();
        let external_books = self.inner.external_books()?;
        let tables = self.inner.tables()?;
        let worksheet_part = self.inner.package.part(&partname)?;
        if worksheet_part.rels().iter().any(|relationship| {
            relationship.reltype().contains("/slicer")
                || relationship.reltype().contains("/timeline")
        }) {
            return Err(Error::UnsupportedFeature(
                "source-backed XLSB worksheet materialization does not yet include slicer or timeline views",
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
        for record in Records::new(worksheet_data.as_bytes(), self.inner.raw_limits) {
            if record?.kind() == kind::BEGIN_SPARKLINE_GROUPS {
                return Err(Error::UnsupportedFeature(
                    "source-backed XLSB worksheet materialization does not yet include sparkline groups",
                ));
            }
        }
        let shared_strings = self.inner.shared_strings()?;
        let styles = self.inner.styles()?;
        let mut formula_context = self.inner.formula_context.clone();
        formula_context.external_books = external_books;
        formula_context.tables = Arc::clone(&tables);
        Ok(MaterializationInputs {
            data: worksheet_data,
            name: metadata.name.clone(),
            shared_strings,
            formula_context,
            workbook_position: metadata.workbook_position,
            cell_xf_count: styles.cell_xfs.len(),
            comments_part,
        })
    }

    /// The comments of the worksheet's comments part.
    fn read_comments(
        &self,
        partname: &PackURI,
        report: &mut ToleranceReport,
    ) -> Result<Vec<crate::comments::Record>> {
        let comments_part = self.inner.package.part(partname)?;
        require_content_type(&comments_part, XLSB_COMMENTS_CONTENT_TYPE)?;
        if !comments_part.rels().is_empty() {
            return Err(Error::InvalidRelationship(
                "XLSB comments parts cannot own relationships",
            ));
        }
        let comments_data = comments_part.data()?;
        Ok(crate::comments::read_recording(
            comments_data.as_bytes(),
            self.inner.raw_limits,
            report,
        )?)
    }

    fn metadata(&self) -> &SheetMetadata {
        &self.inner.sheets[self.catalog_position]
    }

    fn physical_partname(&self) -> Result<&PackURI> {
        self.metadata()
            .partname
            .as_ref()
            .ok_or(Error::UnsupportedFeature(
                "selected XLSB sheet tab has no physical sheet part",
            ))
    }
}

struct TableDefinitions {
    definitions: Arc<[Definition]>,
}

struct ExternalBooks {
    books: Arc<[ExternalBook]>,
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
    fn lock_tolerance(&self) -> std::sync::MutexGuard<'_, ToleranceReport> {
        // The report is only ever appended to whole; a panic elsewhere while
        // the lock was held cannot leave it half-written.
        self.tolerance
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    fn external_books(&self) -> Result<Arc<[ExternalBook]>> {
        let _ = preflight_package(&self.package)?;
        let external_books =
            self.external_books.get_or_try_init(|| {
                let _ = preflight_package(&self.package)?;
                let mut budget = self.external_link_limits.budget();
                budget.preflight_links(self.external_link_parts.len())?;
                let mut books = Vec::new();
                books
                    .try_reserve_exact(self.external_link_parts.len())
                    .map_err(|source| Error::Allocation {
                        resource: "source-backed XLSB external books",
                        source,
                    })?;
                for metadata in &self.external_link_parts {
                    let _ = preflight_package(&self.package)?;
                    let part = self.package.part(&metadata.partname)?;
                    require_content_type(&part, XLSB_EXTERNAL_LINK_CONTENT_TYPE)?;
                    let declared_size = usize::try_from(part.declared_uncompressed_size()?)
                        .map_err(|_source| Error::CapacityOverflow {
                            resource: "source-backed XLSB external-link declared size",
                        })?;
                    budget.preflight_link_part(declared_size)?;
                    postflight_package(&self.package)?;
                    let _ = preflight_package(&self.package)?;
                    let data = part.data()?;
                    postflight_package(&self.package)?;
                    let _ = preflight_package(&self.package)?;
                    let parsed = crate::external_link::parse_external_link_with_budget(
                        data.as_bytes(),
                        &mut budget,
                    )?;
                    postflight_package(&self.package)?;
                    let book = Workbook::resolve_external_book(part.rels(), parsed, &mut budget)?;
                    postflight_package(&self.package)?;
                    books.push(book);
                }
                let books = Arc::new(ExternalBooks {
                    books: books.into(),
                });
                postflight_package(&self.package)?;
                Ok::<_, Error>(books)
            })?;
        postflight_package(&self.package)?;
        Ok(Arc::clone(&external_books.books))
    }

    fn shared_strings(&self) -> Result<Arc<Vec<SharedString>>> {
        self.package.check_execution()?;
        self.package.source_version()?;
        let strings = self.shared_strings.get_or_try_init(|| {
            self.package.check_execution()?;
            self.package.source_version()?;
            let mut report = ToleranceReport::default();
            let mut strings = Vec::new();
            if let Some(partname) = self.shared_strings_part.as_ref() {
                let data = self.package.part(partname)?.data()?;
                self.package.check_execution()?;
                self.package.source_version()?;
                let mut records = Records::new(data.as_bytes(), self.raw_limits);
                Workbook::read_shared_strings(
                    &mut records,
                    &mut strings,
                    &mut report,
                    self.raw_limits,
                )?;
                self.package.check_execution()?;
                self.package.source_version()?;
            }
            let strings = Arc::new(strings);
            self.package.check_execution()?;
            self.package.source_version()?;
            // Published only on success; see `styles`.
            self.lock_tolerance().append(&report);
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
            let mut report = ToleranceReport::default();
            self.package.check_execution()?;
            self.package.source_version()?;
            let styles = if let Some(partname) = self.styles_part.as_ref() {
                let data = self.package.part(partname)?.data()?;
                self.package.check_execution()?;
                self.package.source_version()?;
                let styles = StylesTable::from_bytes_with_report(
                    data.as_bytes(),
                    self.raw_limits,
                    &mut report,
                )?;
                self.package.check_execution()?;
                self.package.source_version()?;
                styles
            } else {
                StylesTable::default()
            };
            let styles = Arc::new(styles);
            self.package.check_execution()?;
            self.package.source_version()?;
            // Publish the repairs only once the cache is about to be filled,
            // so a failed (and later retried) initialization never records
            // its defects twice.
            self.lock_tolerance().append(&report);
            Ok::<_, Error>(styles)
        })?;
        self.package.check_execution()?;
        self.package.source_version()?;
        Ok(styles)
    }

    fn tables(&self) -> Result<Arc<[Definition]>> {
        let _ = preflight_package(&self.package)?;
        let tables = self.tables.get_or_try_init(|| {
            let _ = preflight_package(&self.package)?;
            let mut report = ToleranceReport::default();
            let mut tables = Vec::new();
            tables
                .try_reserve_exact(self.table_parts.len())
                .map_err(|source| Error::Allocation {
                    resource: "source-backed XLSB table definitions",
                    source,
                })?;
            for table_part in &self.table_parts {
                let _ = preflight_package(&self.package)?;
                let part = self.package.part(&table_part.partname)?;
                require_content_type(&part, XLSB_TABLE_CONTENT_TYPE)?;
                let _ = preflight_package(&self.package)?;
                let data = part.data()?;
                postflight_package(&self.package)?;
                let _ = preflight_package(&self.package)?;
                // Table XTI references use the complete sheet catalog,
                // including chart tabs, rather than worksheet ordinals.
                let table = Workbook::parse_table_definition(
                    data.as_bytes(),
                    self.sheets[self.worksheet_positions[table_part.worksheet_position]]
                        .workbook_position,
                    &mut Tolerance::recording(&mut report),
                    self.raw_limits,
                );
                postflight_package(&self.package)?;
                let table = table?;
                let duplicate_error = if tables
                    .iter()
                    .any(|existing: &Definition| existing.table_id() == table.table_id())
                {
                    Some(Error::InvalidFormula("duplicate workbook table ID"))
                } else if tables.iter().any(|existing: &Definition| {
                    crate::package::formula::excel_name_eq(
                        existing.display_name(),
                        table.display_name(),
                    )
                }) {
                    Some(Error::InvalidFormula(
                        "duplicate workbook table display name",
                    ))
                } else {
                    None
                };
                postflight_package(&self.package)?;
                if let Some(error) = duplicate_error {
                    return Err(error);
                }
                tables.push(table);
                postflight_package(&self.package)?;
            }
            let tables = TableDefinitions {
                definitions: tables.into(),
            };
            let tables = Arc::new(tables);
            postflight_package(&self.package)?;
            // Published only on success; see `styles`.
            self.lock_tolerance().append(&report);
            Ok::<_, Error>(tables)
        })?;
        postflight_package(&self.package)?;
        Ok(Arc::clone(&tables.definitions))
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

fn preflight_package(package: &SourceBackedPackage) -> Result<SourceVersion> {
    package.check_execution()?;
    package.source_version().map_err(Into::into)
}

fn postflight_package(package: &SourceBackedPackage) -> Result<()> {
    package.source_version()?;
    package.check_execution()?;
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
            Error::InvalidFormat("XLSB text fragments were not deterministic")
        },
        _ => Error::InvalidFormat("unsupported XLSB text output error"),
    }
}

fn worksheet_dimensions(worksheet: &Worksheet) -> Result<(usize, usize)> {
    let (min_row, min_column, max_row, max_column) = SheetWorksheet::dimensions(worksheet)
        .map(|(a, b, c, d)| (a - 1, b - 1, c - 1, d - 1))
        .unwrap_or((0, 0, 0, 0));
    if min_row > max_row || min_column > max_column {
        return Err(Error::InvalidFormat(
            "XLSB worksheet dimensions are reversed",
        ));
    }
    if max_row > XLSB_MAX_ROW_INDEX {
        return Err(Error::InvalidFormat(
            "XLSB worksheet row dimension exceeds the BIFF12 bound",
        ));
    }
    if max_column > XLSB_MAX_COLUMN_INDEX {
        return Err(Error::InvalidFormat(
            "XLSB worksheet column dimension exceeds the BIFF12 bound",
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

fn count_cell_text(value: &CellValue) -> Result<usize> {
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
        Error::InvalidFormat("XLSB text formatting failed")
    })
}

fn append_cell_text(output: &mut String, value: &CellValue) -> Result<()> {
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
    _owner: &str,
) -> Result<Option<PackURI>> {
    let mut relationships = workbook_part
        .rels()
        .iter()
        .filter(|relationship| relationship_types.contains(&relationship.reltype()));
    let Some(relationship) = relationships.next() else {
        return Ok(None);
    };
    if relationships.next().is_some() {
        return Err(Error::InvalidRelationship(
            "XLSB workbook has multiple relationships",
        ));
    }
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(
            "XLSB workbook relationship is external",
        ));
    }
    relationship.target_partname().map(Some).map_err(Into::into)
}

fn require_content_type(part: &PartView<'_>, expected: &'static str) -> Result<()> {
    if part.content_type() != expected {
        return Err(Error::InvalidContentType { expected });
    }
    Ok(())
}

fn preflight_indexed_part(
    part: &PartView<'_>,
    maximum: usize,
    kind: crate::package::LimitKind,
) -> Result<()> {
    let actual = usize::try_from(part.declared_uncompressed_size()?).map_err(|_error| {
        Error::CapacityOverflow {
            resource: "indexed part size",
        }
    })?;
    if actual > maximum {
        return Err(Error::limit(kind, actual, maximum));
    }
    Ok(())
}

/// Classify a failure to plan a document-property revision over the source.
fn publication_error(error: crate::metadata::Error) -> Error {
    match error {
        crate::metadata::Error::Package(litchi_ooxml_common::Error::Opc(error)) => {
            Error::Opc(error)
        },
        crate::metadata::Error::Package(error) => Error::Common(error),
        crate::metadata::Error::Signed => Error::Signed,
        crate::metadata::Error::PatchConflict => Error::PatchConflict {
            part: "the document-property parts",
        },
        crate::metadata::Error::Refused(refusal) => Error::Refused(refusal),
        crate::metadata::Error::SourceChanged { expected, observed } => {
            Error::Opc(litchi_opc::OpcError::SourceChanged {
                expected,
                actual: observed,
            })
        },
        crate::metadata::Error::Limit {
            observed, maximum, ..
        } => Error::limit(crate::opc::LimitKind::PatchBytes, observed, maximum),
        crate::metadata::Error::Internal(reason) => Error::CapacityOverflow { resource: reason },
        crate::metadata::Error::Invalid(_) | _ => {
            Error::InvalidFormat("the package structure is not valid")
        },
    }
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

#[cfg(test)]
mod materialize_tests {
    use std::io::Cursor;

    use litchi_core::sheet::WorkbookTrait;

    use super::{SourceBackedWorkbook, Workbook};

    fn fixture() -> Vec<u8> {
        std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/Simple.xlsb"
        ))
        .expect("read XLSB fixture")
    }

    #[test]
    fn to_workbook_matches_an_eager_open_and_keeps_the_catalog_usable() {
        let bytes = fixture();
        let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
        let eager = Workbook::new(Cursor::new(bytes)).unwrap();

        let materialized = source.to_workbook().expect("materialize from the catalog");
        assert_eq!(materialized.worksheet_names(), eager.worksheet_names());
        assert_eq!(
            materialized.active_worksheet_index(),
            eager.active_worksheet_index()
        );
        assert_eq!(
            materialized.is_1904_date_system(),
            eager.is_1904_date_system()
        );
        for index in 0..eager.worksheet_count() {
            let left = materialized.worksheet(index).unwrap();
            let right = eager.worksheet(index).unwrap();
            assert_eq!(
                litchi_core::sheet::Worksheet::dimensions(&left),
                litchi_core::sheet::Worksheet::dimensions(&right)
            );
        }
        // The materialized owner is the eager edit entry.
        materialized
            .to_package()
            .edit_cell_values(materialized.worksheet_names()[0].as_str())
            .expect("eager cell-value edit")
            .resolve()
            .expect("worksheet resolution")
            .expect("worksheet selector");
        // The lazy catalog stays readable after materialization.
        assert_eq!(source.worksheet_names().unwrap(), eager.worksheet_names());
    }
}
