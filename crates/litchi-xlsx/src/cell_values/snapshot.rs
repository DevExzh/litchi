//! Immutable source closure for one value-only worksheet capability.

use std::collections::HashSet;
use std::sync::Arc;

use litchi_core::{ExecutionContext, ExecutionError, Selector as CoreSelector, SourceVersion};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, Part, Relationship, Relationships,
    SourceBackedPackage, SourceLineage, SourceRelationshipTarget, SourceTopologyPlan, TargetMode,
};

use crate::cell::{Cell, SharedFormulaStorage, Store, Value};
use crate::error::{EditBlock, Error, Result, allocation, invalid};
use crate::formula::Kind;
use crate::source_payload::SourcePayload;
use crate::workbook::source::validate_sheet_graph;
use crate::{Selector, WorksheetKind, raw};

use super::validation;

const CHARTSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
const STRICT_CHARTSHEET_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet";
const DIALOGSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet";
const MACROSHEET_REL: &str = "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet";
const INTL_MACROSHEET_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet";
const CHARTSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
const CALCULATION_CHAIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

fn check_execution(context: Option<&ExecutionContext>) -> Result<()> {
    let Some(context) = context else {
        return Ok(());
    };
    context.check().map_err(|error| {
        Error::Package(match error {
            ExecutionError::Cancelled => OpcError::Cancelled,
            error => OpcError::Execution(error),
        })
    })
}

/// Exact worksheet values plus the complete package owner state required to
/// publish a one-Part overlay safely.
#[derive(Clone, Debug)]
pub struct Snapshot {
    sheet_name: Box<str>,
    sheet_position: usize,
    cells: Arc<Store>,
    source: SourceState,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct SharedFormulaGroup {
    pub(crate) storage: SharedFormulaStorage,
    pub(crate) master: litchi_sheet::Cell,
    pub(crate) members: Box<[litchi_sheet::Cell]>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SheetGraphState {
    name: Box<str>,
    relationship_id: Box<str>,
    sheet_id: u32,
    visibility: raw::Visibility,
    kind: WorksheetKind,
    uri: PackURI,
    content_type: Box<str>,
    relationships: Box<[SourceRelationship]>,
    relationship: SourceRelationship,
}

struct SourceCatalogCapture {
    sheets: Vec<raw::Sheet>,
    parts: Vec<crate::workbook::source::SheetPart>,
    workbook_uri: PackURI,
    workbook_content_type: Box<str>,
    workbook_xml: SourcePayload,
    owner_relationship: SourceRelationship,
    package_relationships: Arc<[SourceRelationship]>,
    workbook_relationships: Arc<[SourceRelationship]>,
    calculation_chain: Option<CalculationChainState>,
    style_count: u32,
    auxiliary: Arc<[PartState]>,
    graph: Arc<[SheetGraphState]>,
    source_lineage: SourceLineage,
    source_version: SourceVersion,
}

struct OwnedCatalogCapture {
    sheets: Vec<raw::Sheet>,
    parts: Vec<crate::workbook::source::SheetPart>,
    workbook_uri: PackURI,
    workbook_content_type: Box<str>,
    workbook_xml: SourcePayload,
    owner_relationship: SourceRelationship,
    package_relationships: Arc<[SourceRelationship]>,
    workbook_relationships: Arc<[SourceRelationship]>,
    calculation_chain: Option<CalculationChainState>,
    style_count: u32,
    auxiliary: Arc<[PartState]>,
    graph: Arc<[SheetGraphState]>,
}

pub(super) fn checked_multi_bytes(total: usize, next: usize, maximum: usize) -> Result<usize> {
    let updated = total
        .checked_add(next)
        .ok_or_else(|| invalid("multi-sheet worksheet XML size overflows usize"))?;
    if updated > maximum {
        return Err(invalid(format!(
            "multi-sheet worksheet XML exceeds {maximum} bytes"
        )));
    }
    Ok(updated)
}

/// Maximum worksheet owners in one source-backed scalar transaction.
pub const MAX_SHEET_OWNERS: usize = 64;
/// Maximum aggregate worksheet XML retained by one multi-sheet transaction.
pub const MAX_MULTI_WORKSHEET_BYTES: usize = 64 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SourceProvenance {
    Matched,
    Mismatched,
    Unavailable,
}

/// Immutable source closure for a bounded set of worksheets.
///
/// Selected worksheet payload bytes are retained for editing and reversible
/// publication. The complete workbook sheet graph (descriptors, relationships,
/// targets, kinds, and content types) is retained as shared ownership
/// metadata; unrelated worksheet payloads remain deferred so independent edits
/// can compose without materializing or overwriting them.
#[derive(Clone, Debug)]
pub struct MultiSnapshot {
    sheets: Box<[Snapshot]>,
}

impl MultiSnapshot {
    pub(crate) fn from_sheets(mut sheets: Vec<Snapshot>) -> Result<Self> {
        if sheets.is_empty() {
            return Err(invalid("multi-sheet value edits require one worksheet"));
        }
        if sheets.len() > MAX_SHEET_OWNERS {
            return Err(invalid(format!(
                "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
            )));
        }
        let mut total_bytes = 0usize;
        for snapshot in &sheets {
            total_bytes = checked_multi_bytes(
                total_bytes,
                snapshot.source_xml().len(),
                MAX_MULTI_WORKSHEET_BYTES,
            )?;
        }
        sheets.sort_unstable_by_key(Snapshot::sheet_position);
        if sheets
            .windows(2)
            .any(|pair| pair[0].sheet_position() == pair[1].sheet_position())
        {
            return Err(invalid(
                "multi-sheet value edits contain a duplicate worksheet",
            ));
        }
        Ok(Self {
            sheets: sheets.into_boxed_slice(),
        })
    }

    pub(crate) fn load_source_backed<'a, I>(
        package: &SourceBackedPackage,
        selectors: I,
    ) -> Result<Self>
    where
        I: IntoIterator<Item = Selector<'a>>,
    {
        package.check_execution()?;
        let source_version = package.source_version()?;
        let mut selected = Vec::new();
        selected
            .try_reserve_exact(MAX_SHEET_OWNERS)
            .map_err(|source| allocation("multi-sheet selector snapshot", source))?;
        for selector in selectors {
            package.check_execution()?;
            if selected.len() >= MAX_SHEET_OWNERS {
                return Err(invalid(format!(
                    "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
                )));
            }
            selected.push(selector);
        }
        let selectors = selected;
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(MAX_SHEET_OWNERS)
            .map_err(|source| allocation("multi-sheet value snapshot", source))?;
        if selectors.len() > MAX_SHEET_OWNERS {
            return Err(invalid(format!(
                "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
            )));
        }
        let capture = load_source_catalog(package)?;
        let positions = resolve_selectors(&capture.sheets, selectors)?;
        let mut aggregate_bytes = 0usize;
        for position in positions {
            let remaining = MAX_MULTI_WORKSHEET_BYTES.saturating_sub(aggregate_bytes);
            let snapshot = Snapshot::from_source_selected(package, position, &capture, remaining)?;
            aggregate_bytes = checked_multi_bytes(
                aggregate_bytes,
                snapshot.source_xml().len(),
                MAX_MULTI_WORKSHEET_BYTES,
            )?;
            sheets.push(snapshot);
        }
        let snapshot = Self::from_sheets(sheets)?;
        package.check_execution()?;
        let final_version = package.source_version()?;
        if final_version != source_version {
            return Err(invalid(
                "source-backed multi-sheet snapshot source version changed",
            ));
        }
        Ok(snapshot)
    }

    pub(crate) fn load_owned<I>(package: &OpcPackage, positions: I) -> Result<Self>
    where
        I: IntoIterator<Item = usize>,
    {
        let mut selected = Vec::new();
        selected
            .try_reserve_exact(MAX_SHEET_OWNERS)
            .map_err(|source| allocation("multi-sheet owned selector snapshot", source))?;
        for position in positions {
            if selected.len() >= MAX_SHEET_OWNERS {
                return Err(invalid(format!(
                    "multi-sheet value edits exceed {MAX_SHEET_OWNERS} worksheet owners"
                )));
            }
            if selected.contains(&position) {
                return Err(invalid(
                    "multi-sheet value edits contain a duplicate worksheet",
                ));
            }
            selected.push(position);
        }
        selected.sort_unstable();
        let capture = load_owned_catalog(package)?;
        let mut sheets = Vec::new();
        sheets
            .try_reserve_exact(selected.len())
            .map_err(|source| allocation("multi-sheet owned snapshots", source))?;
        let mut aggregate_bytes = 0usize;
        for position in selected {
            if position >= capture.sheets.len() {
                return Err(invalid("multi-sheet worksheet position did not resolve"));
            }
            let remaining = MAX_MULTI_WORKSHEET_BYTES.saturating_sub(aggregate_bytes);
            let snapshot = Snapshot::from_owned_selected(package, position, &capture, remaining)?;
            aggregate_bytes = checked_multi_bytes(
                aggregate_bytes,
                snapshot.source_xml().len(),
                MAX_MULTI_WORKSHEET_BYTES,
            )?;
            sheets.push(snapshot);
        }
        Self::from_sheets(sheets)
    }

    pub(crate) fn sheets(&self) -> &[Snapshot] {
        &self.sheets
    }

    /// Number of selected worksheet owners.
    #[must_use]
    pub fn len(&self) -> usize {
        self.sheets.len()
    }

    /// Whether no worksheet owners were selected.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.sheets.is_empty()
    }

    /// Name of a selected worksheet in deterministic workbook order.
    #[must_use]
    pub fn sheet_name(&self, index: usize) -> Option<&str> {
        self.sheets.get(index).map(Snapshot::sheet_name)
    }

    /// Zero-based workbook position of a selected worksheet.
    #[must_use]
    pub fn sheet_position(&self, index: usize) -> Option<usize> {
        self.sheets.get(index).map(Snapshot::sheet_position)
    }

    /// Exact scalar value in one selected worksheet.
    #[must_use]
    pub fn value(&self, index: usize, address: litchi_sheet::Cell) -> Option<&Value> {
        self.sheets
            .get(index)
            .and_then(|sheet| sheet.value(address))
    }

    /// Whether an explicit cell owner exists in one selected worksheet.
    #[must_use]
    pub fn contains_cell(&self, index: usize, address: litchi_sheet::Cell) -> bool {
        self.sheets
            .get(index)
            .is_some_and(|sheet| sheet.contains_cell(address))
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.sheets.len() == other.sheets.len()
            && self
                .sheets
                .iter()
                .zip(&other.sheets)
                .all(|(left, right)| left.same_source(right))
    }

    pub(crate) fn check_execution(&self) -> Result<()> {
        for snapshot in &self.sheets {
            snapshot.check_execution()?;
        }
        Ok(())
    }

    /// Check that this source-backed closure still belongs to the exact
    /// opened package and source revision without reloading any Part body.
    pub(crate) fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
    ) -> Result<SourceProvenance> {
        package.check_execution()?;
        let version = package.source_version()?;
        package.check_execution()?;
        let lineage = package.source_lineage();
        let mut unavailable = false;
        for snapshot in &self.sheets {
            match snapshot.source_provenance(&lineage, version) {
                SourceProvenance::Matched => {},
                SourceProvenance::Mismatched => return Ok(SourceProvenance::Mismatched),
                SourceProvenance::Unavailable => unavailable = true,
            }
        }
        Ok(if unavailable {
            SourceProvenance::Unavailable
        } else {
            SourceProvenance::Matched
        })
    }
}

impl Snapshot {
    fn from_owned_selected(
        package: &OpcPackage,
        position: usize,
        capture: &OwnedCatalogCapture,
        remaining_bytes: usize,
    ) -> Result<Self> {
        let sheet = &capture.sheets[position];
        let sheet_part = &capture.parts[position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
        let worksheet = package.get_part(&sheet_part.uri)?;
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let worksheet_xml = worksheet.blob_arc();
        checked_multi_bytes(0, worksheet_xml.len(), remaining_bytes)?;
        validation::worksheet_xml(worksheet_xml.as_slice())?;
        let cells = raw::worksheet::parse(worksheet_xml.as_slice(), || Ok(None))?;
        validate_style_references(&cells, capture.style_count)?;
        validate_scalar_cells(&cells)?;
        let graph = &capture.graph[position];
        Ok(Self {
            sheet_name: copy_boxed(&sheet.name, "value-only sheet name")?,
            sheet_position: position,
            cells: Arc::new(cells),
            source: SourceState {
                workbook: PartState::new(
                    capture.workbook_uri.clone(),
                    capture.workbook_content_type.as_ref(),
                    capture.workbook_xml.clone(),
                )?,
                worksheet: PartState::new(
                    worksheet.partname().clone(),
                    worksheet.content_type(),
                    SourcePayload::Owned(worksheet_xml),
                )?,
                owner_relationship: capture.owner_relationship.clone(),
                sheet_relationship: graph.relationship.clone(),
                package_relationships: Arc::clone(&capture.package_relationships),
                workbook_relationships: Arc::clone(&capture.workbook_relationships),
                calculation_chain: capture.calculation_chain.clone(),
                auxiliary: Arc::clone(&capture.auxiliary),
                graph: Arc::clone(&capture.graph),
                context: None,
                source_lineage: None,
                source_version: None,
            },
        })
    }

    fn from_source_selected(
        package: &SourceBackedPackage,
        position: usize,
        capture: &SourceCatalogCapture,
        remaining_bytes: usize,
    ) -> Result<Self> {
        package.check_execution()?;
        let sheet = &capture.sheets[position];
        let sheet_part = &capture.parts[position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let worksheet_xml = SourcePayload::from_part_data(package, worksheet.data()?)?;
        checked_multi_bytes(0, worksheet_xml.len(), remaining_bytes)?;
        validation::worksheet_xml(worksheet_xml.as_bytes())?;
        let cells = raw::worksheet::parse(worksheet_xml.as_bytes(), || Ok(None))?;
        package.check_execution()?;
        validate_style_references(&cells, capture.style_count)?;
        validate_scalar_cells(&cells)?;
        let graph = &capture.graph[position];
        Ok(Self {
            sheet_name: copy_boxed(&sheet.name, "value-only sheet name")?,
            sheet_position: position,
            cells: Arc::new(cells),
            source: SourceState {
                workbook: PartState::new(
                    capture.workbook_uri.clone(),
                    capture.workbook_content_type.as_ref(),
                    capture.workbook_xml.clone(),
                )?,
                worksheet: PartState::new(
                    worksheet.partname().clone(),
                    worksheet.content_type(),
                    worksheet_xml,
                )?,
                owner_relationship: capture.owner_relationship.clone(),
                sheet_relationship: graph.relationship.clone(),
                package_relationships: Arc::clone(&capture.package_relationships),
                workbook_relationships: Arc::clone(&capture.workbook_relationships),
                calculation_chain: capture.calculation_chain.clone(),
                auxiliary: Arc::clone(&capture.auxiliary),
                graph: Arc::clone(&capture.graph),
                context: package.execution_context(),
                source_lineage: Some(capture.source_lineage.clone()),
                source_version: Some(capture.source_version),
            },
        })
    }

    pub(crate) fn load_source_backed<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
    ) -> Result<Self> {
        Self::load_source_backed_with_sheet_policy(package, selector, true)
    }

    fn load_source_backed_with_sheet_policy<'a>(
        package: &SourceBackedPackage,
        selector: impl Into<Selector<'a>>,
        require_single_sheet: bool,
    ) -> Result<Self> {
        package.check_execution()?;
        let source_version = package.source_version()?;
        let workbook = package.main_document_part()?;
        validate_package_relationships(package.rels())?;
        if workbook.content_type() != ct::SML_SHEET_MAIN {
            return Err(invalid(
                "value-only edits require an ordinary XLSX workbook",
            ));
        }
        let workbook_xml = SourcePayload::from_part_data(package, workbook.data()?)?;
        validation::workbook_xml(workbook_xml.as_bytes())?;
        let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
        let sheet_parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
        if require_single_sheet && catalog.sheets.len() != 1 {
            return Err(invalid(
                "value-only edits currently require exactly one worksheet",
            ));
        }
        validate_workbook_relationships(workbook.rels(), require_single_sheet)?;
        let position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("value-only worksheet selector did not resolve"))?;
        let sheet = &catalog.sheets[position];
        let sheet_part = &sheet_parts[position];
        if sheet_part.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name.clone(),
            });
        }
        let worksheet = package.part(&sheet_part.uri)?;
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let sheet_relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        let worksheet_xml = SourcePayload::from_part_data(package, worksheet.data()?)?;
        validation::worksheet_xml(worksheet_xml.as_bytes())?;
        let (style_count, auxiliary) = capture_auxiliary_source(package, &workbook)?;
        let calculation_chain = capture_calculation_chain_source(package, &workbook)?;
        let owner = unique_owner(package.rels())?;
        let snapshot = Self::from_parts(
            &sheet.name,
            position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner,
            package.rels(),
            workbook.rels(),
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            sheet_relationship,
            style_count,
            calculation_chain,
            auxiliary,
            capture_sheet_graph_source(package, &workbook, &catalog.sheets, &sheet_parts)?,
            package.execution_context(),
            Some(package.source_lineage()),
            Some(source_version),
        )?;
        package.check_execution()?;
        let final_version = package.source_version()?;
        if final_version != source_version {
            return Err(invalid(
                "source-backed worksheet snapshot source version changed",
            ));
        }
        Ok(snapshot)
    }

    /// Load and validate one value-only closure from an owning OPC package.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        Self::load_with_sheet_policy(package, selector, true)
    }

    /// Load one selected worksheet from a multi-worksheet value-only closure.
    pub fn load_multi<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        Self::load_with_sheet_policy(package, selector, false)
    }

    fn load_with_sheet_policy<'a>(
        package: &OpcPackage,
        selector: impl Into<Selector<'a>>,
        require_single_sheet: bool,
    ) -> Result<Self> {
        let workbook = package.main_document_part()?;
        validate_package_relationships(package.rels())?;
        if workbook.content_type() != ct::SML_SHEET_MAIN {
            return Err(invalid(
                "value-only edits require an ordinary XLSX workbook",
            ));
        }
        let workbook_xml = SourcePayload::Owned(workbook.blob_arc());
        validation::workbook_xml(workbook_xml.as_bytes())?;
        let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
        let sheet_parts = validate_sheet_graph_owned(package, workbook, &catalog.sheets)?;
        if require_single_sheet && catalog.sheets.len() != 1 {
            return Err(invalid(
                "value-only edits currently require exactly one worksheet",
            ));
        }
        validate_workbook_relationships(workbook.rels(), require_single_sheet)?;
        let position = resolve_selector(&catalog.sheets, selector.into())?
            .ok_or_else(|| invalid("value-only worksheet selector did not resolve"))?;
        let sheet = &catalog.sheets[position];
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("selected worksheet relationship is missing"))?;
        require_worksheet_relationship(relationship)?;
        let uri = relationship.target_partname()?;
        let worksheet = package.get_part(&uri)?;
        if worksheet.content_type() != ct::SML_WORKSHEET {
            return Err(invalid("selected worksheet content type is invalid"));
        }
        if !worksheet.rels().is_empty() {
            return Err(invalid("value-only edits refuse worksheet relationships"));
        }
        let worksheet_xml = SourcePayload::Owned(worksheet.blob_arc());
        validation::worksheet_xml(worksheet_xml.as_bytes())?;
        let (style_count, auxiliary) = capture_auxiliary(package, workbook)?;
        let calculation_chain = capture_calculation_chain_owned(package, workbook)?;
        let owner = unique_owner(package.rels())?;
        Self::from_parts(
            &sheet.name,
            position,
            workbook.partname().clone(),
            workbook.content_type(),
            workbook_xml,
            owner,
            package.rels(),
            workbook.rels(),
            worksheet.partname().clone(),
            worksheet.content_type(),
            worksheet_xml,
            relationship,
            style_count,
            calculation_chain,
            auxiliary,
            capture_sheet_graph_owned(package, workbook, &catalog.sheets, &sheet_parts)?,
            None,
            None,
            None,
        )
    }

    #[allow(clippy::too_many_arguments)]
    fn from_parts(
        sheet_name: &str,
        sheet_position: usize,
        workbook_uri: PackURI,
        workbook_content_type: &str,
        workbook_xml: SourcePayload,
        owner_relationship: &Relationship,
        package_relationships: &Relationships,
        workbook_relationships: &Relationships,
        worksheet_uri: PackURI,
        worksheet_content_type: &str,
        worksheet_xml: SourcePayload,
        sheet_relationship: &Relationship,
        style_count: u32,
        calculation_chain: Option<CalculationChainState>,
        auxiliary: Box<[PartState]>,
        graph: Box<[SheetGraphState]>,
        context: Option<ExecutionContext>,
        source_lineage: Option<SourceLineage>,
        source_version: Option<SourceVersion>,
    ) -> Result<Self> {
        require_worksheet_relationship(sheet_relationship)?;
        if sheet_relationship.target_partname()? != worksheet_uri {
            return Err(invalid(
                "selected worksheet relationship does not target its captured Part",
            ));
        }
        let cells = raw::worksheet::parse(worksheet_xml.as_bytes(), || Ok(None))?;
        validate_style_references(&cells, style_count)?;
        validate_scalar_cells(&cells)?;
        Ok(Self {
            sheet_name: copy_boxed(sheet_name, "value-only sheet name")?,
            sheet_position,
            cells: Arc::new(cells),
            source: SourceState {
                workbook: PartState::new(workbook_uri, workbook_content_type, workbook_xml)?,
                worksheet: PartState::new(worksheet_uri, worksheet_content_type, worksheet_xml)?,
                owner_relationship: SourceRelationship::capture(owner_relationship)?,
                sheet_relationship: SourceRelationship::capture(sheet_relationship)?,
                package_relationships: Arc::from(capture_relationships(package_relationships)?),
                workbook_relationships: Arc::from(capture_relationships(workbook_relationships)?),
                calculation_chain,
                auxiliary: Arc::from(auxiliary),
                graph: Arc::from(graph),
                context,
                source_lineage,
                source_version,
            },
        })
    }

    pub(crate) fn from_rewritten_source(source: &Self, bytes: Vec<u8>) -> Result<Self> {
        source.source.check_execution()?;
        validation::worksheet_xml(&bytes)?;
        let cells = raw::worksheet::parse(&bytes, || Ok(None))?;
        let mut result = source.clone();
        result.cells = Arc::new(cells);
        result.source.worksheet.bytes = SourcePayload::Owned(Arc::new(bytes));
        result.source.check_execution()?;
        Ok(result)
    }

    pub(super) fn invalidated_workbook_xml(&self) -> Result<Arc<Vec<u8>>> {
        self.source.check_execution()?;
        let bytes = raw::recalc::invalidate(self.source.workbook.bytes.as_bytes())?;
        validation::workbook_xml(&bytes)?;
        self.source.check_execution()?;
        Ok(Arc::new(bytes))
    }

    pub(super) fn with_invalidated_calculation(self) -> Result<Self> {
        let workbook = self.invalidated_workbook_xml()?;
        self.with_invalidated_workbook(workbook)
    }

    pub(super) fn with_invalidated_workbook(mut self, workbook: Arc<Vec<u8>>) -> Result<Self> {
        self.source.check_execution()?;
        validation::workbook_xml(workbook.as_slice())?;
        self.source.workbook.bytes = SourcePayload::Owned(workbook);
        if let Some(chain) = self.source.calculation_chain.take() {
            let mut relationships = self.source.workbook_relationships.to_vec();
            relationships.retain(|relationship| relationship.id != chain.relationship.id);
            self.source.workbook_relationships = Arc::from(relationships.into_boxed_slice());
        }
        self.source.check_execution()?;
        Ok(self)
    }

    /// Rebind worksheet bytes after the row-visibility owner proves that only
    /// direct `hidden` attributes changed.
    ///
    /// Cell values, formulas, styles, metadata, and owners are therefore
    /// byte-identical, so their already validated store remains authoritative.
    /// The candidate XML grammar is still validated independently before the
    /// source-bound snapshot is published.
    pub(crate) fn from_visibility_rewrite<'source>(
        source: &'source Self,
        rewrite: crate::row_visibility::rewrite::VisibilityRewrite<'source>,
    ) -> Result<Self> {
        source.source.check_execution()?;
        let bytes = rewrite.into_bytes_for(source.source_xml())?;
        validation::worksheet_xml(&bytes)?;
        let mut result = source.clone();
        result.source.worksheet.bytes = SourcePayload::Owned(Arc::new(bytes));
        result.source.check_execution()?;
        Ok(result)
    }

    #[cfg(test)]
    pub(crate) fn shares_cell_store_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.cells, &other.cells)
    }

    /// Selected worksheet name.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Selected zero-based sheet position.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    /// Exact scalar value at a stored coordinate.
    #[must_use]
    pub fn value(&self, address: litchi_sheet::Cell) -> Option<&Value> {
        match self.cells.entry(address).map(|entry| &entry.cell) {
            Some(Cell::Value(value)) => Some(value),
            _ => None,
        }
    }

    /// Exact semantic cell stored at a coordinate.
    #[must_use]
    pub fn cell(&self, address: litchi_sheet::Cell) -> Option<&Cell> {
        self.cells.entry(address).map(|entry| &entry.cell)
    }

    pub(super) fn editable_cell(&self, address: litchi_sheet::Cell) -> Option<&Cell> {
        self.cell(address)
    }

    pub(crate) fn shared_formula_group(
        &self,
        address: litchi_sheet::Cell,
        maximum_members: usize,
    ) -> Result<SharedFormulaGroup> {
        let entry = self.cells.entry(address).ok_or_else(|| {
            invalid(format!(
                "shared formula selector '{address}' has no existing cell owner"
            ))
        })?;
        let Some(storage) = entry.shared_formula.as_ref() else {
            return Err(self.edit_blocked(address));
        };
        if !storage.master || storage.range.start() != address {
            return Err(Error::EditBlocked {
                sheet: self.sheet_name().to_owned(),
                address,
                reason: EditBlock::GroupFormula,
            });
        }
        if !matches!(&entry.cell, Cell::Formula(formula) if matches!(formula.kind(), Kind::Scalar))
            || entry.cell_metadata.is_some()
            || entry.value_metadata.is_some()
        {
            return Err(invalid(
                "shared formula group contains a non-scalar or metadata cell",
            ));
        }

        let rows = usize::try_from(storage.range.rows())
            .map_err(|_| invalid("shared formula range row count overflows usize"))?;
        let columns = usize::try_from(storage.range.columns())
            .map_err(|_| invalid("shared formula range column count overflows usize"))?;
        let area = rows
            .checked_mul(columns)
            .ok_or_else(|| invalid("shared formula range area overflows usize"))?;
        if area > maximum_members {
            return Err(invalid(format!(
                "shared formula group exceeds the bounded {maximum_members}-cell edit limit"
            )));
        }

        let mut members = Vec::new();
        members
            .try_reserve_exact(area)
            .map_err(|source| allocation("shared formula group members", source))?;
        let start = storage.range.start();
        let (end_row, end_column) = storage.range.end();
        for row in start.row().get()..end_row {
            for column in start.column().get()..end_column {
                let member = litchi_sheet::Cell::at(row, column)
                    .map_err(|_| invalid("shared formula range contains an invalid cell"))?;
                let member_entry = self.cells.entry(member).ok_or_else(|| {
                    invalid(format!("shared formula group is missing cell '{member}'"))
                })?;
                let Some(member_storage) = member_entry.shared_formula.as_ref() else {
                    return Err(invalid(format!(
                        "shared formula group cell '{member}' has no shared storage"
                    )));
                };
                if member_storage.index != storage.index
                    || member_storage.range != storage.range
                    || member_storage.reference != storage.reference
                {
                    return Err(invalid(
                        "shared formula group members do not share exact storage metadata",
                    ));
                }
                if member_storage.master != (member == start)
                    || !matches!(
                        &member_entry.cell,
                        Cell::Formula(formula) if matches!(formula.kind(), Kind::Scalar)
                    )
                    || member_entry.cell_metadata.is_some()
                    || member_entry.value_metadata.is_some()
                {
                    return Err(invalid(
                        "shared formula group contains a non-scalar or metadata cell",
                    ));
                }
                members.push(member);
            }
        }

        for member_entry in self.cells.entries() {
            let Some(member_storage) = member_entry.shared_formula.as_ref() else {
                continue;
            };
            if member_storage.index != storage.index {
                continue;
            }
            if member_storage.range != storage.range
                || member_storage.reference != storage.reference
                || !storage.range.contains(member_entry.address)
                || member_storage.master != (member_entry.address == start)
            {
                return Err(invalid(
                    "shared formula group contains an outsider or mismatched storage metadata",
                ));
            }
        }

        Ok(SharedFormulaGroup {
            storage: storage.clone(),
            master: start,
            members: members.into_boxed_slice(),
        })
    }

    pub(super) fn edit_blocked(&self, address: litchi_sheet::Cell) -> Error {
        let reason = self
            .cells
            .entry(address)
            .and_then(|entry| entry.formula_range)
            .map_or(EditBlock::UnknownCell, |_| EditBlock::GroupFormula);
        Error::EditBlocked {
            sheet: self.sheet_name().to_owned(),
            address,
            reason,
        }
    }

    pub(super) fn require_formula_target(&self, address: litchi_sheet::Cell) -> Result<()> {
        let entry = self.cells.entry(address).ok_or_else(|| {
            invalid(format!(
                "formula selector '{address}' has no existing cell owner"
            ))
        })?;
        if entry.formula_range.is_some() {
            return Err(Error::EditBlocked {
                sheet: self.sheet_name().to_owned(),
                address,
                reason: EditBlock::GroupFormula,
            });
        }
        if matches!(entry.cell, Cell::Value(_) | Cell::Formula(_)) {
            Ok(())
        } else {
            Err(Error::EditBlocked {
                sheet: self.sheet_name().to_owned(),
                address,
                reason: EditBlock::UnknownCell,
            })
        }
    }

    /// Whether an explicit `<c>` owner exists at this coordinate.
    ///
    /// This distinguishes a cleared cell record from a removed cell record.
    #[must_use]
    pub fn contains_cell(&self, address: litchi_sheet::Cell) -> bool {
        self.cells.entry(address).is_some()
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.source.worksheet.bytes.as_bytes()
    }

    /// Selected worksheet Part URI.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        &self.source.worksheet.uri
    }

    pub(crate) fn same_source(&self, other: &Self) -> bool {
        self.sheet_name == other.sheet_name
            && self.sheet_position == other.sheet_position
            && self.source.same_owner(&other.source)
    }

    pub(crate) fn check_execution(&self) -> Result<()> {
        self.source.check_execution()
    }

    fn source_provenance(
        &self,
        lineage: &SourceLineage,
        version: SourceVersion,
    ) -> SourceProvenance {
        match (
            self.source.source_lineage.as_ref(),
            self.source.source_version,
        ) {
            (Some(expected_lineage), Some(expected_version)) => {
                if expected_lineage == lineage && expected_version == version {
                    SourceProvenance::Matched
                } else {
                    SourceProvenance::Mismatched
                }
            },
            (None, None) => SourceProvenance::Unavailable,
            _ => SourceProvenance::Unavailable,
        }
    }

    /// Check that this source-backed closure still belongs to the exact
    /// opened package and source revision without reloading any Part body.
    pub(crate) fn matches_source_backed(
        &self,
        package: &SourceBackedPackage,
    ) -> Result<SourceProvenance> {
        package.check_execution()?;
        let version = package.source_version()?;
        package.check_execution()?;
        let lineage = package.source_lineage();
        Ok(self.source_provenance(&lineage, version))
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if !self.source.workbook.matches_part(workbook)
            || !unique_owner(package.rels())
                .is_ok_and(|owner| self.source.owner_relationship.matches(owner))
            || !relationships_match(package.rels(), &self.source.package_relationships)
            || !relationships_match(workbook.rels(), &self.source.workbook_relationships)
            || self.source.auxiliary.iter().any(|expected| {
                package.get_part(&expected.uri).map_or(true, |part| {
                    !expected.matches_part(part) || !part.rels().is_empty()
                })
            })
        {
            return false;
        }
        if let Some(chain) = &self.source.calculation_chain {
            let Some(relationship) = workbook.rels().get(chain.relationship.id.as_ref()) else {
                return false;
            };
            if !chain.relationship.matches(relationship)
                || package.get_part(&chain.part.uri).map_or(true, |part| {
                    !chain.part.matches_part(part) || !part.rels().is_empty()
                })
            {
                return false;
            }
        }
        let Ok(workbook_xml) = raw::parse_catalog(workbook.blob()) else {
            return false;
        };
        let Ok(sheet_parts) = validate_sheet_graph_owned(package, workbook, &workbook_xml.sheets)
        else {
            return false;
        };
        let Ok(graph) =
            capture_sheet_graph_owned(package, workbook, &workbook_xml.sheets, &sheet_parts)
        else {
            return false;
        };
        if graph.as_ref() != self.source.graph.as_ref() {
            return false;
        }
        let Some(relationship) = workbook
            .rels()
            .get(self.source.sheet_relationship.id.as_ref())
        else {
            return false;
        };
        if !self.source.sheet_relationship.matches(relationship)
            || relationship.target_partname().ok().as_ref() != Some(&self.source.worksheet.uri)
        {
            return false;
        }
        package
            .get_part(&self.source.worksheet.uri)
            .is_ok_and(|part| self.source.worksheet.matches_part(part) && part.rels().is_empty())
    }

    pub(super) fn topology_plan_from(&self, before: &Self) -> Result<SourceTopologyPlan> {
        let mut plan = SourceTopologyPlan::new();
        append_owner_topology(&mut plan, &before.source, &self.source)?;
        append_worksheet_replacement(&mut plan, &before.source, &self.source)?;
        Ok(plan)
    }

    pub(super) fn apply_owned_target(
        before: &Self,
        after: &Self,
        package: &mut OpcPackage,
        maximum_bytes: Option<usize>,
    ) -> Result<()> {
        let mut materialized = 0usize;
        apply_owner_target(
            &before.source,
            &after.source,
            package,
            maximum_bytes,
            &mut materialized,
        )?;
        apply_worksheet_target(
            &before.source,
            &after.source,
            package,
            maximum_bytes,
            &mut materialized,
        )
    }
}

impl MultiSnapshot {
    pub(super) fn topology_plan_from(&self, before: &Self) -> Result<SourceTopologyPlan> {
        if self.len() != before.len() {
            return Err(invalid(
                "multi-sheet topology snapshots have different owners",
            ));
        }
        let mut plan = SourceTopologyPlan::new();
        let first_before = before
            .sheets()
            .first()
            .ok_or_else(|| invalid("multi-sheet topology has no source owner"))?;
        let first_after = self
            .sheets()
            .first()
            .ok_or_else(|| invalid("multi-sheet topology has no target owner"))?;
        append_owner_topology(&mut plan, &first_before.source, &first_after.source)?;
        for (before, after) in before.sheets().iter().zip(self.sheets()) {
            append_worksheet_replacement(&mut plan, &before.source, &after.source)?;
        }
        Ok(plan)
    }

    pub(super) fn apply_owned_target(
        before: &Self,
        after: &Self,
        package: &mut OpcPackage,
        maximum_bytes: Option<usize>,
    ) -> Result<()> {
        if before.len() != after.len() {
            return Err(invalid("multi-sheet patch owner count changed"));
        }
        let mut materialized = 0usize;
        let first_before = before
            .sheets()
            .first()
            .ok_or_else(|| invalid("multi-sheet patch has no source owner"))?;
        let first_after = after
            .sheets()
            .first()
            .ok_or_else(|| invalid("multi-sheet patch has no target owner"))?;
        apply_owner_target(
            &first_before.source,
            &first_after.source,
            package,
            maximum_bytes,
            &mut materialized,
        )?;
        for (before, after) in before.sheets().iter().zip(after.sheets()) {
            apply_worksheet_target(
                &before.source,
                &after.source,
                package,
                maximum_bytes,
                &mut materialized,
            )?;
        }
        Ok(())
    }
}

fn validate_sheet_graph_owned(
    package: &OpcPackage,
    workbook: &dyn Part,
    sheets: &[raw::Sheet],
) -> Result<Vec<crate::workbook::source::SheetPart>> {
    let mut parts = Vec::new();
    let mut targets = HashSet::new();
    parts
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("owned workbook sheet graph", source))?;
    for sheet in sheets {
        let relationship = workbook.rels().get(&sheet.relationship_id).ok_or_else(|| {
            invalid(format!(
                "sheet '{}' references missing relationship '{}'",
                sheet.name, sheet.relationship_id
            ))
        })?;
        if relationship.is_external() {
            return Err(invalid("worksheet relationship cannot be external"));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        let kind = match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => {
                if part.content_type() != ct::SML_WORKSHEET {
                    return Err(invalid("worksheet content type is invalid"));
                }
                WorksheetKind::Worksheet
            },
            CHARTSHEET_REL | STRICT_CHARTSHEET_REL => {
                if part.content_type() != CHARTSHEET_CONTENT_TYPE {
                    return Err(invalid("chartsheet content type is invalid"));
                }
                WorksheetKind::Chart
            },
            DIALOGSHEET_REL => WorksheetKind::Dialog,
            MACROSHEET_REL | INTL_MACROSHEET_REL => WorksheetKind::Macro,
            _ => WorksheetKind::Unknown,
        };
        if !targets.insert(part.partname().clone()) {
            return Err(invalid(
                "workbook sheet graph contains a duplicate Part target",
            ));
        }
        parts.push(crate::workbook::source::SheetPart {
            kind,
            uri: part.partname().clone(),
        });
    }
    Ok(parts)
}

#[derive(Clone, Debug)]
struct SourceState {
    workbook: PartState,
    worksheet: PartState,
    owner_relationship: SourceRelationship,
    sheet_relationship: SourceRelationship,
    package_relationships: Arc<[SourceRelationship]>,
    workbook_relationships: Arc<[SourceRelationship]>,
    calculation_chain: Option<CalculationChainState>,
    auxiliary: Arc<[PartState]>,
    graph: Arc<[SheetGraphState]>,
    context: Option<ExecutionContext>,
    source_lineage: Option<SourceLineage>,
    source_version: Option<SourceVersion>,
}

impl SourceState {
    fn check_execution(&self) -> Result<()> {
        check_execution(self.context.as_ref())
    }

    fn same_owner(&self, other: &Self) -> bool {
        self.workbook == other.workbook
            && self.worksheet == other.worksheet
            && self.owner_relationship == other.owner_relationship
            && self.sheet_relationship == other.sheet_relationship
            && self.package_relationships == other.package_relationships
            && self.workbook_relationships == other.workbook_relationships
            && self.calculation_chain == other.calculation_chain
            && self.auxiliary == other.auxiliary
            && self.graph == other.graph
            && match (&self.source_lineage, &other.source_lineage) {
                (Some(left), Some(right)) => left == right,
                (None, _) | (_, None) => true,
            }
            && match (self.source_version, other.source_version) {
                (Some(left), Some(right)) => left == right,
                (None, _) | (_, None) => true,
            }
    }
}

#[derive(Clone, Debug)]
struct PartState {
    uri: PackURI,
    content_type: Box<str>,
    bytes: SourcePayload,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct CalculationChainState {
    part: PartState,
    relationship: SourceRelationship,
}

impl PartState {
    fn new(uri: PackURI, content_type: &str, bytes: SourcePayload) -> Result<Self> {
        Ok(Self {
            uri,
            content_type: copy_boxed(content_type, "value-only content type")?,
            bytes,
        })
    }
    fn matches_part(&self, part: &dyn Part) -> bool {
        part.partname() == &self.uri
            && part.content_type() == self.content_type.as_ref()
            && part.blob() == self.bytes.as_bytes()
    }
}

impl PartialEq for PartState {
    fn eq(&self, other: &Self) -> bool {
        self.uri == other.uri
            && self.content_type == other.content_type
            && self.bytes == other.bytes
    }
}

impl Eq for PartState {}

#[derive(Clone, Debug, PartialEq, Eq)]
struct SourceRelationship {
    id: Box<str>,
    kind: Box<str>,
    target: Box<str>,
    mode: TargetMode,
}

impl SourceRelationship {
    fn capture(value: &Relationship) -> Result<Self> {
        Ok(Self {
            id: copy_boxed(value.r_id(), "value-only relationship ID")?,
            kind: copy_boxed(value.reltype(), "value-only relationship type")?,
            target: copy_boxed(value.target_ref(), "value-only relationship target")?,
            mode: value.target_mode(),
        })
    }
    fn matches(&self, value: &Relationship) -> bool {
        value.r_id() == self.id.as_ref()
            && value.reltype() == self.kind.as_ref()
            && value.target_ref() == self.target.as_ref()
            && value.target_mode() == self.mode
    }
}

fn copy_payload(bytes: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| allocation(resource, source))?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn append_worksheet_replacement(
    plan: &mut SourceTopologyPlan,
    before: &SourceState,
    after: &SourceState,
) -> Result<()> {
    if before.worksheet.uri != after.worksheet.uri
        || before.worksheet.content_type != after.worksheet.content_type
    {
        return Err(invalid("worksheet topology changed during a cell edit"));
    }
    if before.worksheet.bytes != after.worksheet.bytes {
        plan.try_replace_part(
            after.worksheet.uri.clone(),
            copy_payload(
                after.worksheet.bytes.as_bytes(),
                "source-backed worksheet topology replacement",
            )?,
        )?;
    }
    Ok(())
}

fn append_owner_topology(
    plan: &mut SourceTopologyPlan,
    before: &SourceState,
    after: &SourceState,
) -> Result<()> {
    if before.workbook.uri != after.workbook.uri
        || before.workbook.content_type != after.workbook.content_type
    {
        return Err(invalid("workbook topology changed during a cell edit"));
    }
    if before.workbook.bytes != after.workbook.bytes {
        plan.try_replace_part(
            after.workbook.uri.clone(),
            copy_payload(
                after.workbook.bytes.as_bytes(),
                "source-backed workbook topology replacement",
            )?,
        )?;
    }
    match (&before.calculation_chain, &after.calculation_chain) {
        (Some(before_chain), None) => {
            plan.try_remove_relationship(
                after.workbook.uri.clone(),
                before_chain.relationship.id.as_ref(),
            )?;
            plan.try_remove_part(before_chain.part.uri.clone())?;
        },
        (None, Some(after_chain)) => {
            plan.try_add_part(
                after_chain.part.uri.clone(),
                after_chain.part.content_type.as_ref(),
                copy_payload(
                    after_chain.part.bytes.as_bytes(),
                    "source-backed calculation-chain addition",
                )?,
            )?;
            plan.try_add_relationship(
                after.workbook.uri.clone(),
                after_chain.relationship.id.as_ref(),
                after_chain.relationship.kind.as_ref(),
                SourceRelationshipTarget::Internal(after_chain.part.uri.clone()),
            )?;
        },
        (Some(before_chain), Some(after_chain)) if before_chain != after_chain => {
            return Err(invalid(
                "calculation-chain replacement is outside the cell transaction",
            ));
        },
        (Some(_), Some(_)) | (None, None) => {},
    }
    Ok(())
}

fn materialized_part_arc(
    part: &PartState,
    maximum_bytes: Option<usize>,
    materialized_bytes: &mut usize,
    resource: &'static str,
) -> Result<Arc<Vec<u8>>> {
    match maximum_bytes {
        Some(maximum) => {
            let prior = *materialized_bytes;
            let updated = prior
                .checked_add(part.bytes.len())
                .ok_or_else(|| invalid("cell patch materialization size overflows usize"))?;
            if updated > maximum {
                return Err(invalid(format!(
                    "cell patch materialization exceeds the explicit bound {maximum} bytes"
                )));
            }
            let output = part.bytes.materialized_arc(maximum - prior, resource)?;
            *materialized_bytes = updated;
            Ok(output)
        },
        None => part.bytes.detached_arc(),
    }
}

fn apply_worksheet_target(
    before: &SourceState,
    after: &SourceState,
    package: &mut OpcPackage,
    maximum_bytes: Option<usize>,
    materialized_bytes: &mut usize,
) -> Result<()> {
    if before.worksheet.uri != after.worksheet.uri
        || before.worksheet.content_type != after.worksheet.content_type
    {
        return Err(invalid(
            "worksheet topology changed during patch application",
        ));
    }
    if before.worksheet.bytes != after.worksheet.bytes {
        let bytes = materialized_part_arc(
            &after.worksheet,
            maximum_bytes,
            materialized_bytes,
            "cell patch worksheet materialization",
        )?;
        package
            .get_part_mut(&after.worksheet.uri)?
            .set_blob_shared(bytes);
    }
    Ok(())
}

fn apply_owner_target(
    before: &SourceState,
    after: &SourceState,
    package: &mut OpcPackage,
    maximum_bytes: Option<usize>,
    materialized_bytes: &mut usize,
) -> Result<()> {
    if before.workbook.uri != after.workbook.uri
        || before.workbook.content_type != after.workbook.content_type
    {
        return Err(invalid(
            "workbook topology changed during patch application",
        ));
    }
    if let (Some(before_chain), None) = (&before.calculation_chain, &after.calculation_chain)
        && calculation_chain_is_referenced_elsewhere(
            package,
            &before_chain.part.uri,
            &before.workbook.uri,
            before_chain.relationship.id.as_ref(),
        )?
    {
        return Err(invalid(
            "calculation-chain Part has another inbound relationship",
        ));
    }
    if before.workbook.bytes != after.workbook.bytes {
        let bytes = materialized_part_arc(
            &after.workbook,
            maximum_bytes,
            materialized_bytes,
            "cell patch workbook materialization",
        )?;
        package
            .get_part_mut(&after.workbook.uri)?
            .set_blob_shared(bytes);
    }
    match (&before.calculation_chain, &after.calculation_chain) {
        (Some(before_chain), None) => {
            let removed = package
                .get_part_mut(&after.workbook.uri)?
                .rels_mut()
                .remove(before_chain.relationship.id.as_ref());
            if removed.is_none() || !package.remove_part(&before_chain.part.uri) {
                return Err(invalid(
                    "calculation-chain topology changed before patch application",
                ));
            }
        },
        (None, Some(after_chain)) => {
            let bytes = materialized_part_arc(
                &after_chain.part,
                maximum_bytes,
                materialized_bytes,
                "cell patch calculation-chain materialization",
            )?;
            package.try_add_part(Box::new(BlobPart::new_shared(
                after_chain.part.uri.clone(),
                after_chain.part.content_type.to_string(),
                bytes,
            )))?;
            package
                .get_part_mut(&after.workbook.uri)?
                .rels_mut()
                .try_add_relationship(
                    after_chain.relationship.kind.to_string(),
                    after_chain.relationship.target.to_string(),
                    after_chain.relationship.id.to_string(),
                    after_chain.relationship.mode,
                )?;
        },
        (Some(before_chain), Some(after_chain)) if before_chain != after_chain => {
            return Err(invalid(
                "calculation-chain replacement is outside the cell patch",
            ));
        },
        (Some(_), Some(_)) | (None, None) => {},
    }
    Ok(())
}

fn calculation_chain_is_referenced_elsewhere(
    package: &OpcPackage,
    target: &PackURI,
    workbook: &PackURI,
    workbook_relationship: &str,
) -> Result<bool> {
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if part.partname() == workbook && relationship.r_id() == workbook_relationship {
                continue;
            }
            if !relationship.is_external()
                && relationship.target_partname()?.is_equivalent_to(target)
            {
                return Ok(true);
            }
        }
    }
    for relationship in package.rels().iter() {
        if !relationship.is_external() && relationship.target_partname()?.is_equivalent_to(target) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn calculation_chain_relationship(relationships: &Relationships) -> Result<Option<&Relationship>> {
    let mut matching = relationships.iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::CALC_CHAIN | rt::STRICT_CALC_CHAIN
        )
    });
    let first = matching.next();
    if matching.next().is_some() {
        return Err(invalid(
            "workbook has multiple calculation-chain relationships",
        ));
    }
    Ok(first)
}

fn capture_calculation_chain_source(
    package: &SourceBackedPackage,
    workbook: &litchi_opc::PartView<'_>,
) -> Result<Option<CalculationChainState>> {
    let Some(relationship) = calculation_chain_relationship(workbook.rels())? else {
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let part = package.part(&relationship.target_partname()?)?;
    if part.content_type() != CALCULATION_CHAIN_CONTENT_TYPE || !part.rels().is_empty() {
        return Err(invalid("calculation-chain Part has unsupported topology"));
    }
    let bytes = SourcePayload::from_part_data(package, part.data()?)?;
    Ok(Some(CalculationChainState {
        part: PartState::new(part.partname().clone(), part.content_type(), bytes)?,
        relationship: SourceRelationship::capture(relationship)?,
    }))
}

fn capture_calculation_chain_owned(
    package: &OpcPackage,
    workbook: &dyn Part,
) -> Result<Option<CalculationChainState>> {
    let Some(relationship) = calculation_chain_relationship(workbook.rels())? else {
        return Ok(None);
    };
    if relationship.is_external() {
        return Err(invalid("calculation-chain relationship cannot be external"));
    }
    let part = package.get_part(&relationship.target_partname()?)?;
    if part.content_type() != CALCULATION_CHAIN_CONTENT_TYPE || !part.rels().is_empty() {
        return Err(invalid("calculation-chain Part has unsupported topology"));
    }
    Ok(Some(CalculationChainState {
        part: PartState::new(
            part.partname().clone(),
            part.content_type(),
            SourcePayload::Owned(part.blob_arc()),
        )?,
        relationship: SourceRelationship::capture(relationship)?,
    }))
}

fn validate_workbook_relationships(
    relationships: &Relationships,
    require_single_sheet: bool,
) -> Result<()> {
    let mut worksheets = 0usize;
    let mut styles = 0usize;
    let mut themes = 0usize;
    let mut calculation_chains = 0usize;
    for relationship in relationships.iter() {
        if relationship.is_external() {
            return Err(invalid(
                "value-only edits refuse external workbook relationships",
            ));
        }
        if !matches!(
            relationship.reltype(),
            rt::WORKSHEET
                | rt::STRICT_WORKSHEET
                | rt::STYLES
                | rt::STRICT_STYLES
                | rt::THEME
                | rt::CALC_CHAIN
                | rt::STRICT_CALC_CHAIN
        ) {
            return Err(invalid(format!(
                "value-only edits refuse workbook relationship '{}'",
                relationship.reltype()
            )));
        }
        match relationship.reltype() {
            rt::WORKSHEET | rt::STRICT_WORKSHEET => worksheets += 1,
            rt::STYLES | rt::STRICT_STYLES => styles += 1,
            rt::THEME => themes += 1,
            rt::CALC_CHAIN | rt::STRICT_CALC_CHAIN => calculation_chains += 1,
            _ => {},
        }
    }
    let worksheets_valid = if require_single_sheet {
        worksheets == 1
    } else {
        worksheets > 0
    };
    if !worksheets_valid || styles > 1 || themes > 1 || calculation_chains > 1 {
        return Err(invalid(
            "cell edits require worksheet relationships and at most one styles, theme, and calculation-chain relationship",
        ));
    }
    Ok(())
}

fn validate_scalar_cells(cells: &Store) -> Result<()> {
    if cells.entries().iter().any(|entry| {
        matches!(entry.cell, Cell::Unknown(_))
            || entry.cell_metadata.is_some()
            || entry.value_metadata.is_some()
    }) {
        return Err(invalid("cell edits refuse unknown cells and cell metadata"));
    }
    Ok(())
}

fn resolve_selectors<'a>(
    sheets: &[raw::Sheet],
    selectors: Vec<Selector<'a>>,
) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    positions
        .try_reserve_exact(selectors.len())
        .map_err(|source| allocation("multi-sheet selector positions", source))?;
    for selector in selectors {
        let position = resolve_selector(sheets, selector)?
            .ok_or_else(|| invalid("value-only worksheet selector did not resolve"))?;
        if positions.contains(&position) {
            return Err(invalid(
                "multi-sheet value edits contain a duplicate worksheet",
            ));
        }
        positions.push(position);
    }
    positions.sort_unstable();
    Ok(positions)
}

fn capture_sheet_graph_source(
    package: &SourceBackedPackage,
    workbook: &litchi_opc::PartView<'_>,
    sheets: &[raw::Sheet],
    parts: &[crate::workbook::source::SheetPart],
) -> Result<Box<[SheetGraphState]>> {
    let mut graph = Vec::new();
    graph
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("source-backed worksheet graph", source))?;
    for (sheet, part_info) in sheets.iter().zip(parts) {
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("worksheet relationship is missing"))?;
        let part = package.part(&part_info.uri)?;
        graph.push(SheetGraphState {
            name: copy_boxed(&sheet.name, "worksheet graph sheet name")?,
            relationship_id: copy_boxed(&sheet.relationship_id, "worksheet graph relationship ID")?,
            sheet_id: sheet.sheet_id,
            visibility: sheet.visibility.clone(),
            kind: part_info.kind,
            uri: part.partname().clone(),
            content_type: copy_boxed(part.content_type(), "worksheet graph content type")?,
            relationships: capture_relationships(part.rels())?,
            relationship: SourceRelationship::capture(relationship)?,
        });
    }
    Ok(graph.into_boxed_slice())
}

fn capture_sheet_graph_owned(
    package: &OpcPackage,
    workbook: &dyn Part,
    sheets: &[raw::Sheet],
    parts: &[crate::workbook::source::SheetPart],
) -> Result<Box<[SheetGraphState]>> {
    let mut graph = Vec::new();
    graph
        .try_reserve_exact(sheets.len())
        .map_err(|source| allocation("owned worksheet graph", source))?;
    for (sheet, part_info) in sheets.iter().zip(parts) {
        let relationship = workbook
            .rels()
            .get(&sheet.relationship_id)
            .ok_or_else(|| invalid("worksheet relationship is missing"))?;
        let part = package.get_part(&part_info.uri)?;
        graph.push(SheetGraphState {
            name: copy_boxed(&sheet.name, "worksheet graph sheet name")?,
            relationship_id: copy_boxed(&sheet.relationship_id, "worksheet graph relationship ID")?,
            sheet_id: sheet.sheet_id,
            visibility: sheet.visibility.clone(),
            kind: part_info.kind,
            uri: part.partname().clone(),
            content_type: copy_boxed(part.content_type(), "worksheet graph content type")?,
            relationships: capture_relationships(part.rels())?,
            relationship: SourceRelationship::capture(relationship)?,
        });
    }
    Ok(graph.into_boxed_slice())
}

fn load_source_catalog(package: &SourceBackedPackage) -> Result<SourceCatalogCapture> {
    package.check_execution()?;
    let workbook = package.main_document_part()?;
    validate_package_relationships(package.rels())?;
    if workbook.content_type() != ct::SML_SHEET_MAIN {
        return Err(invalid(
            "value-only edits require an ordinary XLSX workbook",
        ));
    }
    let workbook_xml = SourcePayload::from_part_data(package, workbook.data()?)?;
    validation::workbook_xml(workbook_xml.as_bytes())?;
    let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
    let parts = validate_sheet_graph(package, &workbook, &catalog.sheets)?;
    validate_workbook_relationships(workbook.rels(), false)?;
    let owner = unique_owner(package.rels())?;
    let (style_count, auxiliary) = capture_auxiliary_source(package, &workbook)?;
    let calculation_chain = capture_calculation_chain_source(package, &workbook)?;
    let graph = capture_sheet_graph_source(package, &workbook, &catalog.sheets, &parts)?;
    package.check_execution()?;
    Ok(SourceCatalogCapture {
        sheets: catalog.sheets.clone(),
        parts,
        workbook_uri: workbook.partname().clone(),
        workbook_content_type: copy_boxed(
            workbook.content_type(),
            "value-only workbook content type",
        )?,
        workbook_xml,
        owner_relationship: SourceRelationship::capture(owner)?,
        package_relationships: Arc::from(capture_relationships(package.rels())?),
        workbook_relationships: Arc::from(capture_relationships(workbook.rels())?),
        calculation_chain,
        style_count,
        auxiliary: Arc::from(auxiliary),
        graph: Arc::from(graph),
        source_lineage: package.source_lineage(),
        source_version: package.source_version()?,
    })
}

fn load_owned_catalog(package: &OpcPackage) -> Result<OwnedCatalogCapture> {
    let workbook = package.main_document_part()?;
    validate_package_relationships(package.rels())?;
    if workbook.content_type() != ct::SML_SHEET_MAIN {
        return Err(invalid(
            "value-only edits require an ordinary XLSX workbook",
        ));
    }
    let workbook_xml = SourcePayload::Owned(workbook.blob_arc());
    validation::workbook_xml(workbook_xml.as_bytes())?;
    let catalog = raw::parse_catalog(workbook_xml.as_bytes())?;
    let parts = validate_sheet_graph_owned(package, workbook, &catalog.sheets)?;
    validate_workbook_relationships(workbook.rels(), false)?;
    let owner = unique_owner(package.rels())?;
    let (style_count, auxiliary) = capture_auxiliary(package, workbook)?;
    let calculation_chain = capture_calculation_chain_owned(package, workbook)?;
    let graph = capture_sheet_graph_owned(package, workbook, &catalog.sheets, &parts)?;
    Ok(OwnedCatalogCapture {
        sheets: catalog.sheets.clone(),
        parts,
        workbook_uri: workbook.partname().clone(),
        workbook_content_type: copy_boxed(
            workbook.content_type(),
            "value-only workbook content type",
        )?,
        workbook_xml,
        owner_relationship: SourceRelationship::capture(owner)?,
        package_relationships: Arc::from(capture_relationships(package.rels())?),
        workbook_relationships: Arc::from(capture_relationships(workbook.rels())?),
        calculation_chain,
        style_count,
        auxiliary: Arc::from(auxiliary),
        graph: Arc::from(graph),
    })
}

fn validate_package_relationships(relationships: &Relationships) -> Result<()> {
    let mut owners = 0usize;
    for relationship in relationships.iter() {
        if relationship.is_external()
            || !matches!(
                relationship.reltype(),
                rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT | rt::DIGITAL_SIGNATURE_ORIGIN
            )
        {
            return Err(invalid(format!(
                "value-only edits refuse package relationship '{}'",
                relationship.reltype()
            )));
        }
        if matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        ) {
            owners += 1;
        }
    }
    if owners != 1 {
        return Err(invalid(
            "value-only edits require exactly one package officeDocument owner",
        ));
    }
    Ok(())
}

fn capture_auxiliary_source(
    package: &SourceBackedPackage,
    workbook: &litchi_opc::PartView<'_>,
) -> Result<(u32, Box<[PartState]>)> {
    let mut style_count = 0;
    let mut auxiliary = Vec::new();
    auxiliary
        .try_reserve_exact(2)
        .map_err(|source| allocation("value-only auxiliary closure", source))?;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::STYLES | rt::STRICT_STYLES | rt::THEME
        )
    }) {
        package.check_execution()?;
        let part = package.part(&relationship.target_partname()?)?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "value-only edits refuse styles and theme relationships",
            ));
        }
        let data = SourcePayload::from_part_data(package, part.data()?)?;
        match relationship.reltype() {
            rt::STYLES | rt::STRICT_STYLES if part.content_type() == ct::SML_STYLES => {
                style_count = raw::styles::parse(data.as_bytes())?.len();
            },
            rt::THEME if part.content_type() == ct::OFC_THEME => {},
            rt::STYLES | rt::STRICT_STYLES => {
                return Err(invalid("styles relationship has the wrong content type"));
            },
            rt::THEME => return Err(invalid("theme relationship has the wrong content type")),
            _ => {},
        }
        auxiliary.push(PartState::new(
            part.partname().clone(),
            part.content_type(),
            data,
        )?);
    }
    package.check_execution()?;
    Ok((style_count, auxiliary.into_boxed_slice()))
}

fn capture_auxiliary(package: &OpcPackage, workbook: &dyn Part) -> Result<(u32, Box<[PartState]>)> {
    let mut style_count = 0;
    let mut auxiliary = Vec::new();
    auxiliary
        .try_reserve_exact(2)
        .map_err(|source| allocation("value-only auxiliary closure", source))?;
    for relationship in workbook.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::STYLES | rt::STRICT_STYLES | rt::THEME
        )
    }) {
        let part = package.get_part(&relationship.target_partname()?)?;
        if !part.rels().is_empty() {
            return Err(invalid(
                "value-only edits refuse styles and theme relationships",
            ));
        }
        match relationship.reltype() {
            rt::STYLES | rt::STRICT_STYLES if part.content_type() == ct::SML_STYLES => {
                style_count = raw::styles::parse(part.blob())?.len();
            },
            rt::THEME if part.content_type() == ct::OFC_THEME => {},
            rt::STYLES | rt::STRICT_STYLES => {
                return Err(invalid("styles relationship has the wrong content type"));
            },
            rt::THEME => return Err(invalid("theme relationship has the wrong content type")),
            _ => {},
        }
        auxiliary.push(PartState::new(
            part.partname().clone(),
            part.content_type(),
            SourcePayload::Owned(part.blob_arc()),
        )?);
    }
    Ok((style_count, auxiliary.into_boxed_slice()))
}

fn validate_style_references(cells: &Store, style_count: u32) -> Result<()> {
    let invalid_cell = cells
        .entries()
        .iter()
        .any(|entry| entry.style.is_some_and(|style| style >= style_count));
    let invalid_row = cells.row_entries().iter().any(|entry| {
        entry
            .properties
            .style
            .is_some_and(|style| style >= style_count)
    });
    let invalid_column = cells.column_entries().iter().any(|entry| {
        entry
            .properties
            .style
            .is_some_and(|style| style >= style_count)
    });
    if invalid_cell || invalid_row || invalid_column {
        return Err(invalid(
            "worksheet references a shared style outside the styles table",
        ));
    }
    Ok(())
}

fn require_worksheet_relationship(relationship: &Relationship) -> Result<()> {
    if relationship.target_mode() != TargetMode::Internal
        || !matches!(relationship.reltype(), rt::WORKSHEET | rt::STRICT_WORKSHEET)
    {
        return Err(invalid("selected worksheet relationship is invalid"));
    }
    Ok(())
}

fn unique_owner(relationships: &Relationships) -> Result<&Relationship> {
    let mut owners = relationships.iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let owner = owners
        .next()
        .ok_or_else(|| invalid("workbook has no officeDocument owner"))?;
    if owners.next().is_some() || owner.is_external() {
        return Err(invalid(
            "workbook officeDocument owner is not unique and internal",
        ));
    }
    Ok(owner)
}

fn resolve_selector(sheets: &[raw::Sheet], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        CoreSelector::Position(position) => {
            Ok((position.get() < sheets.len()).then_some(position.get()))
        },
        CoreSelector::Name(name) => {
            let key = crate::sheet::key(&name);
            Ok(sheets
                .iter()
                .position(|sheet| crate::sheet::key(&sheet.name) == key))
        },
        CoreSelector::Id(never) => match never {},
        _ => Err(Error::UnsupportedSelector),
    }
}

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|source| allocation(resource, source))?;
    output.push_str(value);
    Ok(output.into_boxed_str())
}

fn capture_relationships(values: &Relationships) -> Result<Box<[SourceRelationship]>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(values.len())
        .map_err(|source| allocation("value-only relationship closure", source))?;
    for value in values.iter() {
        output.push(SourceRelationship::capture(value)?);
    }
    output.sort_unstable_by(|left, right| left.id.cmp(&right.id));
    Ok(output.into_boxed_slice())
}

fn relationships_match(values: &Relationships, expected: &[SourceRelationship]) -> bool {
    values.len() == expected.len()
        && values.iter().all(|value| {
            expected
                .binary_search_by(|item| item.id.as_ref().cmp(value.r_id()))
                .is_ok_and(|index| expected[index].matches(value))
        })
}

#[cfg(test)]
mod tests {
    use super::checked_multi_bytes;

    #[test]
    fn aggregate_cap_accepts_exact_total_and_rejects_one_over() {
        assert_eq!(checked_multi_bytes(0, 4, 4).unwrap(), 4);
        assert!(checked_multi_bytes(4, 1, 4).is_err());
    }
}
