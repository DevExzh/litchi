//! Deferred, source-backed XLSB worksheet catalog and materialization.

use std::io::Cursor;
use std::sync::{Arc, Mutex};

use litchi_core::{ExecutionContext, ReadAt, SourceVersion};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{
    PackURI, PartView, ReadLimits, SourceBackedPackage, SourceCacheDiagnostics, SourceCacheLimits,
};

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
    shared_strings: Mutex<Option<Arc<Vec<SharedString>>>>,
    styles: Mutex<Option<Arc<StylesTable>>>,
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
                shared_strings: Mutex::new(None),
                styles: Mutex::new(None),
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

impl SourceInner {
    fn shared_strings(&self) -> Result<Arc<Vec<SharedString>>> {
        self.package.check_execution()?;
        self.package.source_version()?;
        {
            let retained = self.shared_strings.lock().map_err(|_error| {
                Error::InvalidFormat(
                    "source-backed XLSB shared-string cache is poisoned".to_string(),
                )
            })?;
            if let Some(strings) = retained.as_ref() {
                return Ok(Arc::clone(strings));
            }
        }
        let mut strings = Vec::new();
        if let Some(partname) = self.shared_strings_part.as_ref() {
            let data = self.package.part(partname)?.data()?;
            let mut records = Records::new(data.as_bytes());
            Workbook::read_shared_strings(&mut records, &mut strings)?;
        }
        self.package.source_version()?;
        let strings = Arc::new(strings);
        let mut retained = self.shared_strings.lock().map_err(|_error| {
            Error::InvalidFormat("source-backed XLSB shared-string cache is poisoned".to_string())
        })?;
        Ok(Arc::clone(retained.get_or_insert(strings)))
    }

    fn styles(&self) -> Result<Arc<StylesTable>> {
        self.package.check_execution()?;
        self.package.source_version()?;
        {
            let retained = self.styles.lock().map_err(|_error| {
                Error::InvalidFormat("source-backed XLSB style cache is poisoned".to_string())
            })?;
            if let Some(styles) = retained.as_ref() {
                return Ok(Arc::clone(styles));
            }
        }
        let styles = if let Some(partname) = self.styles_part.as_ref() {
            let data = self.package.part(partname)?.data()?;
            StylesTable::from_bytes(data.as_bytes())?
        } else {
            StylesTable::default()
        };
        self.package.source_version()?;
        let styles = Arc::new(styles);
        let mut retained = self.styles.lock().map_err(|_error| {
            Error::InvalidFormat("source-backed XLSB style cache is poisoned".to_string())
        })?;
        Ok(Arc::clone(retained.get_or_insert(styles)))
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
