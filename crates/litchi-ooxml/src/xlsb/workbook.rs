//! Workbook implementation for XLSB files

use crate::xlsb::XlsbCell;
use crate::xlsb::calculation::CalculationProperties;
use crate::xlsb::error::XlsbResult;
use crate::xlsb::formula::{
    FormulaExternalBook, FormulaExternalSheet, FormulaPivotViewDefinition,
    FormulaResolutionContext, FormulaSupportingLink, FormulaTableDefinition, excel_name_eq,
    XlsbExternalLink, XlsbExternalLinkKind,
};
use crate::xlsb::named_ranges::{NamedRange, validate_defined_name};
use crate::xlsb::merged_cells::{MAX_MERGED_CELL_RANGES, MergedCell};
use crate::xlsb::records::{XlsbRecord, XlsbRecordIter, record_types};
use crate::xlsb::shared_strings::SharedString;
use crate::xlsb::styles_table::{CellFormat, StylesTable};
use crate::xlsb::vba_project::{VbaProject, discover_vba_project};
use crate::xlsb::worksheet::XlsbWorksheet;
use litchi_core::binary;
use litchi_core::sheet::{Result, Worksheet as SheetTrait, WorksheetIterator};
use litchi_opc::OpcPackage;
use litchi_opc::constants::relationship_type;
use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap};
use std::io::{BufReader, Cursor, Read, Seek, Write};

/// External-workbook relationship types documented by MS-XLSB.
const EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES: &[&str] = &[
    relationship_type::EXTERNAL_LINK_PATH,
    relationship_type::STRICT_EXTERNAL_LINK_PATH,
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlLibrary",
    "http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing",
    "http://schemas.microsoft.com/office/2019/04/relationships/externalLinkLongPath",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlStartup",
    "http://schemas.microsoft.com/office/2019/04/relationships/xlExternalLinkLongPath/xlAlternateStartup",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlPathMissing",
    "http://schemas.microsoft.com/office/2009/04/relationships/xlExternalLinkLongPath/xlLibrary",
];

/// OLE data-source relationship types documented by MS-XLSB and MS-OI29500.
const OLE_DATA_SOURCE_RELATIONSHIP_TYPES: &[&str] = &[
    relationship_type::OLE_OBJECT,
    relationship_type::STRICT_OLE_OBJECT,
    "http://schemas.microsoft.com/office/2019/04/relationships/oleObjectLinkLongPath",
];

/// XLSB workbook implementation
#[allow(dead_code)]
pub struct XlsbWorkbook {
    package: OpcPackage,
    worksheets: Vec<XlsbWorksheet>,
    worksheet_rel_ids: Vec<Option<String>>,
    formula_context: FormulaResolutionContext,
    shared_strings: Vec<SharedString>,
    styles: StylesTable,
    calculation_properties: CalculationProperties,
    is_1904: bool,
    pivot_cache_definitions: Vec<(u32, crate::xlsb::pivot::PivotCacheDefinition)>,
}

#[derive(Default)]
struct ParsedWorkbookInfo {
    worksheet_names: Vec<String>,
    worksheet_rel_ids: Vec<Option<String>>,
    supporting_links: Vec<FormulaSupportingLink>,
    external_sheets: Vec<FormulaExternalSheet>,
    external_link_rel_ids: Vec<String>,
    defined_names: Vec<String>,
    is_1904: bool,
    calculation_properties: Option<CalculationProperties>,
}

#[derive(Debug)]
struct MergeBlockLayout {
    ranges: Vec<MergedCell>,
    block_span: Option<(usize, usize)>,
    insertion_offset: usize,
}

impl std::fmt::Debug for XlsbWorkbook {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("XlsbWorkbook")
            .field("worksheet_names", &self.formula_context.worksheet_names)
            .field("worksheet_rel_ids", &self.worksheet_rel_ids)
            .field("shared_strings_count", &self.shared_strings.len())
            .field("cell_xfs_count", &self.styles.cell_xfs.len())
            .field("calculation_properties", &self.calculation_properties)
            .field("is_1904", &self.is_1904)
            .finish()
    }
}

impl XlsbWorkbook {
    /// Get the underlying OPC package.
    pub fn opc_package(&self) -> &OpcPackage {
        &self.package
    }

    /// Get mutable OPC access, dropping signatures that would become stale.
    pub fn opc_package_mut(&mut self) -> &mut OpcPackage {
        let _ = self.package.clear_digital_signatures();
        &mut self.package
    }

    /// Verify package signatures without making a PKI trust determination.
    pub fn verify_digital_signatures(
        &self,
        policy: &litchi_opc::SignatureVerificationPolicy,
    ) -> litchi_opc::signature::Result<Vec<litchi_opc::DigitalSignatureVerification>> {
        self.package.verify_digital_signatures(policy)
    }

    /// Sign the current, fully materialized package while preserving valid signatures.
    pub fn add_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<litchi_opc::PackURI> {
        self.package.add_digital_signature(signer)
    }

    /// Replace all package signatures with one new signature.
    pub fn resign_digital_signature(
        &mut self,
        signer: &litchi_opc::PackageSigner,
    ) -> litchi_opc::signature::Result<litchi_opc::PackURI> {
        self.package.resign_digital_signature(signer)
    }

    /// Remove all package digital signatures.
    pub fn clear_digital_signatures(&mut self) -> litchi_opc::signature::Result<()> {
        self.package.clear_digital_signatures()
    }

    /// Discover the attached MS-XLSB VBA project and declared signatures.
    ///
    /// This validates only the declared OPC relationship graph and content
    /// types. It does not inspect, parse, verify, or execute VBA project or
    /// signature bytes.
    pub fn vba_project(&self) -> crate::error::Result<Option<VbaProject>> {
        let workbook = self.package.main_document_part()?;
        discover_vba_project(&self.package, workbook)
    }

    /// Workbook and sheet-scoped defined names in `PtgName` index order.
    pub fn defined_names(&self) -> &[String] {
        &self.formula_context.defined_names
    }

    /// Return stored external-workbook, DDE, and OLE link metadata.
    ///
    /// The returned values are data-only snapshots in workbook link order.
    /// Litchi never follows, opens, contacts, refreshes, evaluates, or
    /// executes any external-link target.
    pub fn external_links(&self) -> Vec<XlsbExternalLink> {
        self.formula_context
            .external_books
            .iter()
            .map(FormulaExternalBook::metadata)
            .collect()
    }

    /// Structured-table definitions in workbook table-ID order of discovery.
    pub fn tables(&self) -> &[FormulaTableDefinition] {
        &self.formula_context.tables
    }

    /// PivotTable views available as hosts for calculated field/item formulas.
    pub fn pivot_views(&self) -> &[FormulaPivotViewDefinition] {
        &self.formula_context.pivot_views
    }

    /// Typed PivotCache definitions paired with their workbook cache
    /// identifiers, in workbook declaration order (MS-XLSB 2.1.7.38).
    ///
    /// These are inert data snapshots: external connection identifiers,
    /// relationship identifiers, MDX expressions, and formula tokens are
    /// stored verbatim and are never dereferenced, refreshed, or evaluated.
    pub fn pivot_cache_definitions(&self) -> &[(u32, crate::xlsb::pivot::PivotCacheDefinition)] {
        &self.pivot_cache_definitions
    }

    /// Look up a typed PivotCache definition by its workbook cache identifier.
    pub fn pivot_cache_definition(
        &self,
        cache_id: u32,
    ) -> Option<&crate::xlsb::pivot::PivotCacheDefinition> {
        self.pivot_cache_definitions
            .iter()
            .find(|(id, _)| *id == cache_id)
            .map(|(_, definition)| definition)
    }

    /// Workbook style table loaded from `xl/styles.bin`.
    pub fn styles(&self) -> &StylesTable {
        &self.styles
    }

    /// Unique strings loaded from `xl/sharedStrings.bin`, including rich-text
    /// and phonetic metadata when present.
    pub fn shared_strings(&self) -> &[SharedString] {
        &self.shared_strings
    }

    /// Resolve a parsed cell's style reference to its cell XF.
    pub fn style_for_cell(&self, cell: &XlsbCell) -> Option<&CellFormat> {
        self.styles.get_cell_format(cell.style_id() as usize)
    }

    /// Workbook formula calculation policy.
    pub fn calculation_properties(&self) -> &CalculationProperties {
        &self.calculation_properties
    }

    /// Save this parsed workbook, including atomic worksheet-stream mutations.
    pub fn save<W: Write + Seek>(&self, writer: W) -> XlsbResult<()> {
        self.package.to_stream(writer)?;
        Ok(())
    }

    /// List merged ranges in a worksheet selected by zero-based index.
    pub fn merged_cell_ranges(&self, worksheet_index: usize) -> XlsbResult<Vec<MergedCell>> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let part = self.package.get_part(&uri)?;
        Ok(Self::inspect_merge_block(part.blob())?.ranges)
    }

    /// List merged ranges in a worksheet selected by exact name.
    pub fn merged_cell_ranges_by_name(&self, worksheet_name: &str) -> XlsbResult<Vec<MergedCell>> {
        self.merged_cell_ranges(self.worksheet_index(worksheet_name)?)
    }

    /// Atomically replace all merged ranges in a worksheet selected by index.
    pub fn set_merged_cell_ranges(
        &mut self,
        worksheet_index: usize,
        ranges: &[MergedCell],
    ) -> XlsbResult<()> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let original = self.package.get_part(&uri)?.blob().to_vec();
        let layout = Self::inspect_merge_block(&original)?;
        let normalized = Self::normalize_merge_ranges(ranges)?;
        let replacement = Self::serialize_merge_block(&normalized)?;
        let (start, end) = layout
            .block_span
            .unwrap_or((layout.insertion_offset, layout.insertion_offset));
        let capacity = original
            .len()
            .checked_sub(end - start)
            .and_then(|value| value.checked_add(replacement.len()))
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidLength {
                    expected: usize::MAX,
                    found: original.len(),
                }
            })?;
        let mut updated = Vec::with_capacity(capacity);
        updated.extend_from_slice(&original[..start]);
        updated.extend_from_slice(&replacement);
        updated.extend_from_slice(&original[end..]);
        Self::inspect_merge_block(&updated)?;
        self.package.get_part_mut(&uri)?.set_blob(updated);
        Ok(())
    }

    /// Atomically replace all merged ranges in a worksheet selected by name.
    pub fn set_merged_cell_ranges_by_name(
        &mut self,
        worksheet_name: &str,
        ranges: &[MergedCell],
    ) -> XlsbResult<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.set_merged_cell_ranges(index, ranges)
    }

    /// Atomically add one merged range to a worksheet selected by index.
    pub fn add_merged_cell_range(
        &mut self,
        worksheet_index: usize,
        range: MergedCell,
    ) -> XlsbResult<()> {
        let mut ranges = self.merged_cell_ranges(worksheet_index)?;
        ranges.push(range);
        self.set_merged_cell_ranges(worksheet_index, &ranges)
    }

    /// Atomically add one merged range to a worksheet selected by name.
    pub fn add_merged_cell_range_by_name(
        &mut self,
        worksheet_name: &str,
        range: MergedCell,
    ) -> XlsbResult<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.add_merged_cell_range(index, range)
    }

    /// Atomically remove an exact merged range from a worksheet by index.
    pub fn remove_merged_cell_range(
        &mut self,
        worksheet_index: usize,
        range: &MergedCell,
    ) -> XlsbResult<bool> {
        let mut ranges = self.merged_cell_ranges(worksheet_index)?;
        let Some(index) = ranges.iter().position(|candidate| candidate == range) else {
            return Ok(false);
        };
        ranges.remove(index);
        self.set_merged_cell_ranges(worksheet_index, &ranges)?;
        Ok(true)
    }

    /// Atomically remove an exact merged range from a worksheet by name.
    pub fn remove_merged_cell_range_by_name(
        &mut self,
        worksheet_name: &str,
        range: &MergedCell,
    ) -> XlsbResult<bool> {
        let index = self.worksheet_index(worksheet_name)?;
        self.remove_merged_cell_range(index, range)
    }

    /// Atomically clear all merged ranges in a worksheet selected by index.
    pub fn clear_merged_cell_ranges(&mut self, worksheet_index: usize) -> XlsbResult<()> {
        self.set_merged_cell_ranges(worksheet_index, &[])
    }

    /// Atomically clear all merged ranges in a worksheet selected by name.
    pub fn clear_merged_cell_ranges_by_name(&mut self, worksheet_name: &str) -> XlsbResult<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.clear_merged_cell_ranges(index)
    }

    /// Open an XLSB workbook from a reader
    pub fn new<R: Read + Seek>(reader: R) -> XlsbResult<Self> {
        let package = OpcPackage::from_reader(reader)?;
        let mut workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calculation_properties: CalculationProperties::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
        };

        workbook.load_workbook_info()?;
        workbook.load_styles()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    /// Create an XLSB workbook from an already-parsed OPC package.
    ///
    /// This is used for single-pass parsing where the OPC package has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `package` - An already-parsed OPC package
    pub fn from_opc_package(package: OpcPackage) -> XlsbResult<Self> {
        let mut workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calculation_properties: CalculationProperties::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
        };

        workbook.load_workbook_info()?;
        workbook.load_styles()?;
        workbook.load_shared_strings()?;

        Ok(workbook)
    }

    fn worksheet_index(&self, worksheet_name: &str) -> XlsbResult<usize> {
        self.formula_context
            .worksheet_names
            .iter()
            .position(|name| name == worksheet_name)
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::WorksheetNotFound(worksheet_name.to_string())
            })
    }

    fn worksheet_uri(&self, index: usize) -> XlsbResult<litchi_opc::PackURI> {
        let name = self
            .formula_context
            .worksheet_names
            .get(index)
            .ok_or_else(|| {
                crate::error::OoxmlError::InvalidFormat(format!(
                    "Worksheet index {index} out of bounds"
                ))
            })?;
        let rel_id = self
            .worksheet_rel_ids
            .get(index)
            .and_then(Option::as_deref)
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;
        let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
            crate::xlsb::error::XlsbError::FileNotFound(format!(
                "relationship {rel_id:?} for sheet {name:?}"
            ))
        })?;
        if relationship.is_external() {
            return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                "sheet {name:?} has an external worksheet relationship"
            )));
        }
        Ok(relationship.target_partname()?)
    }

    fn merge_range_key(range: &MergedCell) -> (u32, u32, u32, u32) {
        (
            range.row_first,
            range.col_first,
            range.row_last,
            range.col_last,
        )
    }

    fn normalize_merge_ranges(ranges: &[MergedCell]) -> XlsbResult<Vec<MergedCell>> {
        if ranges.len() > MAX_MERGED_CELL_RANGES {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: MAX_MERGED_CELL_RANGES,
                found: ranges.len(),
            });
        }
        let mut normalized = ranges.to_vec();
        for range in &normalized {
            range.validate()?;
        }
        normalized.sort_unstable_by_key(Self::merge_range_key);
        Self::validate_merge_range_collection(&normalized, false)?;
        Ok(normalized)
    }

    fn validate_merge_range_collection(
        ranges: &[MergedCell],
        require_canonical_order: bool,
    ) -> XlsbResult<()> {
        if ranges.len() > MAX_MERGED_CELL_RANGES {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: MAX_MERGED_CELL_RANGES,
                found: ranges.len(),
            });
        }
        let mut active = BTreeMap::<u32, (u32, u32)>::new();
        let mut expirations = BinaryHeap::<Reverse<(u32, u32)>>::new();
        let mut previous = None;
        for range in ranges {
            range.validate()?;
            let key = Self::merge_range_key(range);
            if require_canonical_order && previous.is_some_and(|value| value >= key) {
                return Err(crate::xlsb::error::XlsbError::Unrecognized {
                    typ: "BrtMergeCell collection".to_string(),
                    val: "duplicate or noncanonical range order".to_string(),
                });
            }
            previous = Some(key);
            while let Some(Reverse((row_last, col_first))) = expirations.peek().copied() {
                if row_last >= range.row_first {
                    break;
                }
                expirations.pop();
                if active.get(&col_first).is_some_and(|entry| entry.1 == row_last) {
                    active.remove(&col_first);
                }
            }
            if let Some((&col_first, &(col_last, _))) = active.range(..=range.col_last).next_back() {
                if col_last >= range.col_first {
                    return Err(crate::xlsb::error::XlsbError::InvalidCellReference(format!(
                        "merged range {} overlaps an existing range beginning in column {}",
                        range.to_range_string(), col_first
                    )));
                }
            }
            active.insert(range.col_first, (range.col_last, range.row_last));
            expirations.push(Reverse((range.row_last, range.col_first)));
        }
        Ok(())
    }

    fn is_post_merge_record(record_type: u16) -> bool {
        matches!(
            record_type,
            record_types::PHONETIC_INFO
                | record_types::H_LINK
                | record_types::BEGIN_D_VALS
                | record_types::BEGIN_D_VALS14
                | record_types::BEGIN_COND_FORMATTING
                | record_types::BEGIN_COND_FORMATTING14
                | record_types::MARGINS
                | record_types::PRINT_OPTIONS
                | record_types::PAGE_SETUP
                | record_types::BEGIN_HEADER_FOOTER
                | record_types::DRAWING
                | record_types::LEGACY_DRAWING
                | record_types::LEGACY_DRAWING_HF
        )
    }

    fn inspect_merge_block(data: &[u8]) -> XlsbResult<MergeBlockLayout> {
        let mut cursor = Cursor::new(data);
        let mut begin_offset = None;
        let mut block_span = None;
        let mut declared_count = None;
        let mut ranges = Vec::new();
        let mut in_block = false;
        let mut saw_end_sheet_data = false;
        let mut end_sheet_offset = None;
        let mut first_post_merge_offset = None;
        while (cursor.position() as usize) < data.len() {
            let start = cursor.position() as usize;
            let record = XlsbRecord::read(&mut cursor)?;
            let end = cursor.position() as usize;
            match record.header.record_type {
                record_types::BEGIN_MERGE_CELLS => {
                    if in_block || begin_offset.is_some() || !saw_end_sheet_data {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBeginMergeCells".to_string(),
                            val: "duplicate, nested, or out-of-order record".to_string(),
                        });
                    }
                    if record.data.len() != 4 {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: 4,
                            found: record.data.len(),
                        });
                    }
                    let count = binary::read_u32_le_at(&record.data, 0)? as usize;
                    if count == 0 || count > MAX_MERGED_CELL_RANGES {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: MAX_MERGED_CELL_RANGES,
                            found: count,
                        });
                    }
                    begin_offset = Some(start);
                    declared_count = Some(count);
                    in_block = true;
                },
                record_types::MERGE_CELL => {
                    if !in_block {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtMergeCell".to_string(),
                            val: "record occurs outside BrtBeginMergeCells".to_string(),
                        });
                    }
                    ranges.push(MergedCell::parse(&record.data)?);
                    if ranges.len() > declared_count.unwrap_or_default() {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBeginMergeCells".to_string(),
                            val: "declared count is smaller than the record collection".to_string(),
                        });
                    }
                },
                record_types::END_MERGE_CELLS => {
                    if !in_block || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtEndMergeCells".to_string(),
                            val: "orphan, duplicate, or nonempty record".to_string(),
                        });
                    }
                    if declared_count != Some(ranges.len()) {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBeginMergeCells".to_string(),
                            val: format!(
                                "declared count {:?} disagrees with {} BrtMergeCell records",
                                declared_count,
                                ranges.len()
                            ),
                        });
                    }
                    block_span = Some((begin_offset.expect("merge begin offset"), end));
                    in_block = false;
                },
                record_types::END_SHEET_DATA => {
                    if in_block {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtMergeCells collection".to_string(),
                            val: "noncontiguous record collection".to_string(),
                        });
                    }
                    saw_end_sheet_data = true;
                },
                record_types::END_SHEET => {
                    if in_block || end_sheet_offset.replace(start).is_some() {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtEndSheet".to_string(),
                            val: "duplicate or embedded in merge collection".to_string(),
                        });
                    }
                },
                record_type => {
                    if in_block {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtMergeCells collection".to_string(),
                            val: format!("unexpected record 0x{record_type:04X}"),
                        });
                    }
                    if saw_end_sheet_data
                        && first_post_merge_offset.is_none()
                        && Self::is_post_merge_record(record_type)
                    {
                        first_post_merge_offset = Some(start);
                    }
                },
            }
        }
        if in_block || begin_offset.is_some() != block_span.is_some() {
            return Err(crate::xlsb::error::XlsbError::UnexpectedEndOfStream(
                "BrtMergeCells collection".to_string(),
            ));
        }
        let end_sheet_offset = end_sheet_offset.ok_or_else(|| {
            crate::xlsb::error::XlsbError::UnexpectedEndOfStream("BrtEndSheet".to_string())
        })?;
        if !saw_end_sheet_data {
            return Err(crate::xlsb::error::XlsbError::UnexpectedEndOfStream(
                "BrtEndSheetData".to_string(),
            ));
        }
        if block_span.is_some() {
            Self::validate_merge_range_collection(&ranges, true)?;
            if first_post_merge_offset.is_some_and(|offset| {
                block_span.is_some_and(|(begin, _)| offset < begin)
            }) {
                return Err(crate::xlsb::error::XlsbError::Unrecognized {
                    typ: "BrtMergeCells collection".to_string(),
                    val: "collection occurs after a later worksheet feature".to_string(),
                });
            }
        }
        Ok(MergeBlockLayout {
            ranges,
            block_span,
            insertion_offset: first_post_merge_offset.unwrap_or(end_sheet_offset),
        })
    }

    fn serialize_merge_block(ranges: &[MergedCell]) -> XlsbResult<Vec<u8>> {
        if ranges.is_empty() {
            return Ok(Vec::new());
        }
        let mut output = Vec::with_capacity(10 + ranges.len() * 19);
        let mut writer = crate::xlsb::writer::RecordWriter::new(&mut output);
        writer.write_record(
            record_types::BEGIN_MERGE_CELLS,
            &(ranges.len() as u32).to_le_bytes(),
        )?;
        for range in ranges {
            writer.write_record(record_types::MERGE_CELL, &range.serialize())?;
        }
        writer.write_record(record_types::END_MERGE_CELLS, &[])?;
        Ok(output)
    }

    /// Load workbook information from workbook.bin
    fn load_workbook_info(&mut self) -> XlsbResult<()> {
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;

        let blob = workbook_part.blob();
        let mut iter = XlsbRecordIter::new(BufReader::new(blob));
        let info = Self::read_workbook(&mut iter)?;
        let external_link_uris = info
            .external_link_rel_ids
            .iter()
            .map(|rel_id| {
                let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
                    crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is missing"
                    ))
                })?;
                if relationship.is_external() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "BrtSupBookSrc relationship {rel_id:?} is external"
                    )));
                }
                relationship.target_partname().map_err(Into::into)
            })
            .collect::<XlsbResult<Vec<_>>>()?;
        let external_books = external_link_uris
            .iter()
            .map(|uri| self.load_external_book(uri))
            .collect::<XlsbResult<Vec<_>>>()?;
        let pivot_cache_ids = Self::parse_pivot_cache_ids(workbook_part.blob())?;
        let mut pivot_cache_definitions = Vec::with_capacity(pivot_cache_ids.len());
        for (cache_id, rel_id) in &pivot_cache_ids {
            let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "PivotCache {cache_id} relationship {rel_id:?} is missing"
                ))
            })?;
            if relationship.is_external()
                || !relationship
                    .reltype()
                    .to_ascii_lowercase()
                    .ends_with("/pivotcachedefinition")
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "PivotCache {cache_id} relationship is external or has the wrong type"
                )));
            }
            let part = self.package.get_part(&relationship.target_partname()?)?;
            let definition = crate::xlsb::pivot::parse_pivot_cache_definition(part.blob())?;
            pivot_cache_definitions.push((*cache_id, definition));
        }

        let mut tables = Vec::new();
        let mut pivot_views = Vec::new();
        for (sheet_index, rel_id) in info.worksheet_rel_ids.iter().enumerate() {
            let Some(rel_id) = rel_id else { continue };
            let Some(sheet_relationship) = workbook_part.rels().get(rel_id) else {
                continue;
            };
            if sheet_relationship.is_external() {
                continue;
            }
            let sheet_part = self
                .package
                .get_part(&sheet_relationship.target_partname()?)?;
            for relationship in sheet_part.rels().iter().filter(|relationship| {
                matches!(
                    relationship.reltype(),
                    relationship_type::TABLE | relationship_type::STRICT_TABLE
                ) || relationship.reltype()
                    == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells"
            }) {
                if relationship.is_external() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                        "worksheet has an external table relationship".to_string(),
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let table = Self::parse_table_definition(part.blob(), sheet_index)?;
                if tables.iter().any(|existing: &FormulaTableDefinition| {
                    existing.table_id() == table.table_id()
                }) {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "duplicate workbook table ID {}",
                        table.table_id()
                    )));
                }
                if tables.iter().any(|existing: &FormulaTableDefinition| {
                    excel_name_eq(existing.display_name(), table.display_name())
                }) {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "duplicate workbook table display name {:?}",
                        table.display_name()
                    )));
                }
                tables.push(table);
            }
            for relationship in sheet_part.rels().iter().filter(|relationship| {
                relationship
                    .reltype()
                    .to_ascii_lowercase()
                    .ends_with("/pivottable")
            }) {
                if relationship.is_external() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                        "worksheet has an external PivotTable relationship".to_string(),
                    ));
                }
                let part = self.package.get_part(&relationship.target_partname()?)?;
                let view = Self::parse_pivot_view(part.blob(), sheet_index)?;
                if !pivot_cache_ids
                    .iter()
                    .any(|(cache_id, _)| *cache_id == view.cache_id())
                {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "PivotTable view {:?} references unknown cache {}",
                        view.name(),
                        view.cache_id()
                    )));
                }
                if pivot_views
                    .iter()
                    .any(|existing: &FormulaPivotViewDefinition| {
                        existing.cache_id() == view.cache_id()
                            && existing.sheet_index() == view.sheet_index()
                            && excel_name_eq(existing.name(), view.name())
                    })
                {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "duplicate PivotTable view {:?} for cache {} on sheet {sheet_index}",
                        view.name(),
                        view.cache_id()
                    )));
                }
                pivot_views.push(view);
            }
        }
        self.formula_context = FormulaResolutionContext {
            worksheet_names: info.worksheet_names.into(),
            supporting_links: info.supporting_links.into(),
            external_sheets: info.external_sheets.into(),
            external_books: external_books.into(),
            defined_names: info.defined_names.into(),
            tables: tables.into(),
            pivot_views: pivot_views.into(),
            pivot_name_scopes: Vec::new().into(),
            active_pivot_scope: None,
            current_sheet: None,
        };
        self.worksheet_rel_ids = info.worksheet_rel_ids;
        self.is_1904 = info.is_1904;
        self.calculation_properties = info.calculation_properties.unwrap_or_default();
        self.pivot_cache_definitions = pivot_cache_definitions;

        Ok(())
    }

    /// Load shared strings from xl/sharedStrings.bin
    fn load_shared_strings(&mut self) -> XlsbResult<()> {
        let shared_strings_uri = litchi_opc::PackURI::new("/xl/sharedStrings.bin")?;
        if let Ok(shared_strings_part) = self.package.get_part(&shared_strings_uri) {
            let blob = shared_strings_part.blob();
            let mut iter = XlsbRecordIter::new(BufReader::new(blob));
            Self::read_shared_strings(&mut iter, &mut self.shared_strings)?;
        }

        Ok(())
    }

    /// Load workbook styles. The default table keeps style index zero usable
    /// for minimal producer files that omit the optional styles part.
    fn load_styles(&mut self) -> XlsbResult<()> {
        let styles_uri = litchi_opc::PackURI::new("/xl/styles.bin")?;
        if let Ok(styles_part) = self.package.get_part(&styles_uri) {
            self.styles = StylesTable::from_reader(styles_part.blob())?;
        }
        Ok(())
    }

    /// Load a concrete XLSB worksheet by index, including related comments.
    pub fn worksheet(&self, index: usize) -> XlsbResult<XlsbWorksheet> {
        if index >= self.formula_context.worksheet_names.len() {
            return Err(crate::error::OoxmlError::InvalidFormat(format!(
                "Worksheet index {} out of bounds",
                index
            ))
            .into());
        }

        let name = &self.formula_context.worksheet_names[index];
        let rel_id = self
            .worksheet_rel_ids
            .get(index)
            .and_then(Option::as_deref)
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;
        let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
            crate::xlsb::error::XlsbError::FileNotFound(format!(
                "relationship {rel_id:?} for sheet {name:?}"
            ))
        })?;
        if relationship.is_external() {
            return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(format!(
                "sheet {name:?} has an external worksheet relationship"
            )));
        }
        let sheet_uri = relationship.target_partname()?;

        let sheet_part = self.package.get_part(&sheet_uri)?;
        let comments_uri = {
            let mut relationships = sheet_part
                .rels()
                .iter()
                .filter(|rel| rel.reltype() == relationship_type::COMMENTS);
            let first = relationships.next();
            if relationships.next().is_some() {
                return Err(crate::xlsb::error::XlsbError::Unrecognized {
                    typ: "worksheet comments relationship".to_string(),
                    val: "multiple relationships".to_string(),
                });
            }
            match first {
                Some(rel) if rel.is_external() => {
                    return Err(crate::xlsb::error::XlsbError::UnsupportedFeature(
                        "external XLSB comments part".to_string(),
                    ));
                },
                Some(rel) => Some(rel.target_partname()?),
                None => None,
            }
        };
        let blob = sheet_part.blob();
        let cursor = Cursor::new(blob);
        let mut worksheet = Self::read_worksheet(
            cursor,
            name.clone(),
            &self.shared_strings,
            &self.formula_context,
            index,
            self.styles.cell_xfs.len(),
        )?;
        if let Some(uri) = comments_uri {
            let part = self.package.get_part(&uri)?;
            if !part.rels().is_empty() {
                return Err(crate::xlsb::error::XlsbError::Unrecognized {
                    typ: "Comments part".to_string(),
                    val: "relationships are not permitted".to_string(),
                });
            }
            for comment in crate::xlsb::comments::read_comments(part.blob())? {
                worksheet.add_comment(comment);
            }
        }
        Ok(worksheet)
    }

    /// Read shared strings from SST
    fn read_shared_strings(
        iter: &mut XlsbRecordIter<impl Read>,
        strings: &mut Vec<SharedString>,
    ) -> XlsbResult<()> {
        let initial_count = strings.len();
        let mut expected_unique = None;
        let mut ended = false;
        for record in iter.by_ref() {
            let record = record?;
            match record.header.record_type {
                record_types::BEGIN_SST => {
                    if expected_unique.is_some() {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBeginSst".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    if record.data.len() != 8 {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: 8,
                            found: record.data.len(),
                        });
                    }
                    let total = binary::read_u32_le_at(&record.data, 0)?;
                    let unique = binary::read_u32_le_at(&record.data, 4)?;
                    if total > 0x7FFF_FFFF || unique > total {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBeginSst counts".to_string(),
                            val: format!("total={total}, unique={unique}"),
                        });
                    }
                    expected_unique = Some(unique as usize);
                },
                record_types::SST_ITEM => {
                    let expected = expected_unique.ok_or_else(|| {
                        crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtSSTItem".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found >= expected {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtSSTItem count".to_string(),
                            val: format!("more than declared {expected}"),
                        });
                    }
                    strings.push(SharedString::parse(&record.data)?);
                },
                record_types::END_SST => {
                    if !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: 0,
                            found: record.data.len(),
                        });
                    }
                    let expected = expected_unique.ok_or_else(|| {
                        crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtEndSst".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found != expected {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtSSTItem count".to_string(),
                            val: format!("declared {expected}, found {found}"),
                        });
                    }
                    ended = true;
                    break;
                },
                _ => {
                    // Skip other records
                },
            }
        }
        if expected_unique.is_none() {
            return Err(crate::xlsb::error::XlsbError::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtBeginSst".to_string(),
            });
        }
        if !ended {
            return Err(crate::xlsb::error::XlsbError::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtEndSst".to_string(),
            });
        }
        Ok(())
    }

    /// Read workbook structure
    fn read_workbook(iter: &mut XlsbRecordIter<impl Read>) -> XlsbResult<ParsedWorkbookInfo> {
        let mut info = ParsedWorkbookInfo::default();
        let worksheet_names = &mut info.worksheet_names;
        let worksheet_rel_ids = &mut info.worksheet_rel_ids;
        let supporting_links = &mut info.supporting_links;
        let external_sheets = &mut info.external_sheets;
        let external_link_rel_ids = &mut info.external_link_rel_ids;
        let defined_names = &mut info.defined_names;
        let is_1904 = &mut info.is_1904;
        for record in iter.by_ref() {
            let record = record?;
            match record.header.record_type {
                record_types::WORKBOOK_PROP => {
                    if let Ok(prop) = crate::xlsb::records::WorkbookPropRecord::parse(&record.data)
                    {
                        *is_1904 = prop.is_date1904;
                    }
                },
                record_types::CALC_PROP => {
                    if info.calculation_properties.is_some() {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtCalcProp".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    info.calculation_properties = Some(CalculationProperties::parse(&record.data)?);
                },
                record_types::BUNDLE_SH => {
                    let bundle_sh = crate::xlsb::records::BundleSheetRecord::parse(&record.data)?;
                    if worksheet_names
                        .iter()
                        .any(|name| excel_name_eq(name, &bundle_sh.name))
                    {
                        return Err(crate::xlsb::error::XlsbError::Unrecognized {
                            typ: "BrtBundleSh strName".to_string(),
                            val: format!("duplicate sheet name {:?}", bundle_sh.name),
                        });
                    }
                    worksheet_names.push(bundle_sh.name);
                    worksheet_rel_ids.push(bundle_sh.rel_id);
                },
                record_types::SUP_SELF => {
                    supporting_links.push(FormulaSupportingLink::SelfWorkbook);
                },
                record_types::SUP_SAME => {
                    supporting_links.push(FormulaSupportingLink::SameSheet);
                },
                record_types::SUP_BOOK_SRC => {
                    let (rel_id, consumed) = crate::xlsb::records::wide_str_with_len(&record.data)?;
                    if rel_id.is_empty() || consumed != record.data.len() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupBookSrc has an invalid relationship ID".to_string(),
                        ));
                    }
                    let book_index = u32::try_from(external_link_rel_ids.len()).map_err(|_| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "external-link count overflow".to_string(),
                        )
                    })?;
                    external_link_rel_ids.push(rel_id);
                    supporting_links.push(FormulaSupportingLink::ExternalWorkbook(book_index));
                },
                record_types::SUP_ADDIN => {
                    supporting_links.push(FormulaSupportingLink::AddIn);
                },
                record_types::EXTERN_SHEET => {
                    Self::parse_extern_sheet(&record.data, external_sheets)?;
                },
                record_types::NAME => {
                    let named_range = NamedRange::parse(&record.data)?;
                    if named_range
                        .sheet_id
                        .is_some_and(|index| index as usize >= worksheet_names.len())
                    {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                            "BrtName {} has invalid sheet scope {:?}",
                            named_range.name, named_range.sheet_id
                        )));
                    }
                    defined_names.push(named_range.name);
                },
                _ => {
                    // Skip other records
                },
            }
        }
        Ok(info)
    }

    /// Read a worksheet
    fn read_worksheet(
        cursor: Cursor<&[u8]>,
        name: String,
        shared_strings: &[SharedString],
        formula_context: &FormulaResolutionContext,
        sheet_index: usize,
        cell_xf_count: usize,
    ) -> XlsbResult<XlsbWorksheet> {
        let mut worksheet = XlsbWorksheet::new(name);
        let iter = crate::xlsb::records::RecordIter::<std::io::Cursor<&[u8]>>::from_cursor(cursor);
        let formula_context = formula_context.for_sheet(sheet_index);
        let mut cells_reader = crate::xlsb::cells_reader::XlsbCellsReader::new(
            iter,
            shared_strings,
            &formula_context,
            cell_xf_count,
        )?;

        // Read all cells
        while let Some(cell) = cells_reader.next_cell()? {
            worksheet.add_cell(cell);
        }

        // Transfer advanced features from reader to worksheet
        for merged in cells_reader.merged_cells {
            worksheet.add_merged_cell(merged);
        }
        for hyperlink in cells_reader.hyperlinks {
            worksheet.add_hyperlink(hyperlink);
        }
        worksheet.set_column_infos(cells_reader.column_infos);
        worksheet.set_row_infos(cells_reader.row_infos);
        worksheet.set_auto_filter(cells_reader.auto_filter);
        worksheet.set_sheet_protection(cells_reader.sheet_protection);
        worksheet.set_strong_sheet_protection(cells_reader.strong_sheet_protection);
        worksheet.set_data_validations(
            cells_reader.data_validation_settings,
            cells_reader.data_validation14_settings,
            cells_reader.data_validations,
        );
        worksheet.set_conditional_formattings(cells_reader.conditional_formattings);

        Ok(worksheet)
    }

    fn parse_extern_sheet(
        data: &[u8],
        external_sheets: &mut Vec<FormulaExternalSheet>,
    ) -> XlsbResult<()> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "BrtExternSheet count overflow".to_string(),
            )
        })?;
        if count >= 65_536 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtExternSheet count {count} exceeds 65,535"
            )));
        }
        let expected = 4usize
            .checked_add(count.checked_mul(12).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?)
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?;
        if data.len() != expected {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        external_sheets.reserve(count);
        for chunk in data[4..].chunks_exact(12) {
            external_sheets.push(FormulaExternalSheet {
                external_link: binary::read_u32_le_at(chunk, 0)?,
                first_sheet: binary::read_u32_le_at(chunk, 4)? as i32,
                last_sheet: binary::read_u32_le_at(chunk, 8)? as i32,
            });
        }
        Ok(())
    }

    fn load_external_book(&self, uri: &litchi_opc::PackURI) -> XlsbResult<FormulaExternalBook> {
        let part = self.package.get_part(uri)?;
        let mut iter = XlsbRecordIter::new(BufReader::new(part.blob()));
        let mut link_type = None;
        let mut target_key = String::new();
        let mut target_detail = String::new();
        let mut sheet_names = Vec::new();
        let mut defined_names = Vec::new();
        let mut saw_sup_tabs = false;
        // 0 = outside a name, 1 = expect formula, 2 = expect bits,
        // 3 = expect end (or optional DDE/OLE cached values).
        let mut sup_name_state = 0u8;
        let mut saw_end = false;

        for record in &mut iter {
            let record = record?;
            if saw_end {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "external link has records after BrtEndSupBook".to_string(),
                ));
            }
            if link_type.is_none() && record.header.record_type != record_types::BEGIN_SUP_BOOK {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "external link does not start with BrtBeginSupBook".to_string(),
                ));
            }
            match record.header.record_type {
                record_types::BEGIN_SUP_BOOK => {
                    if link_type.is_some() || record.data.len() < 10 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtBeginSupBook framing".to_string(),
                        ));
                    }
                    let kind = binary::read_u16_le_at(&record.data, 0)?;
                    let (first, consumed) =
                        crate::xlsb::records::wide_str_with_len(&record.data[2..])?;
                    let mut offset = 2 + consumed;
                    let (second, consumed) = if kind == 0 {
                        Self::parse_nullable_wide_string(&record.data[offset..])?
                    } else {
                        let (value, consumed) =
                            crate::xlsb::records::wide_str_with_len(&record.data[offset..])?;
                        (Some(value), consumed)
                    };
                    offset += consumed;
                    if offset != record.data.len() || kind > 2 || first.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtBeginSupBook payload".to_string(),
                        ));
                    }
                    if kind == 0 && second.is_some() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "external workbook BrtBeginSupBook string2 is not NULL".to_string(),
                        ));
                    }
                    link_type = Some(kind);
                    target_key = first;
                    target_detail = second.unwrap_or_default();
                },
                record_types::SUP_TABS => {
                    if link_type != Some(0) || saw_sup_tabs || sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupTabs".to_string(),
                        ));
                    }
                    sheet_names = Self::parse_external_sheet_names(&record.data)?;
                    saw_sup_tabs = true;
                },
                record_types::SUP_NAME_START => {
                    let kind = link_type.ok_or_else(|| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameStart precedes BrtBeginSupBook".to_string(),
                        )
                    })?;
                    if sup_name_state != 0 || (kind == 0 && !saw_sup_tabs) {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameStart".to_string(),
                        ));
                    }
                    let (name, consumed) = crate::xlsb::records::wide_str_with_len(&record.data)?;
                    if consumed != record.data.len() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameStart has trailing bytes".to_string(),
                        ));
                    }
                    if kind == 0 {
                        validate_defined_name(&name)?;
                        sup_name_state = 1;
                    } else {
                        sup_name_state = 2;
                    }
                    defined_names.push(name);
                },
                record_types::SUP_NAME_FORMULA => {
                    if link_type != Some(0) || sup_name_state != 1 || record.data.len() < 4 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameFmla".to_string(),
                        ));
                    }
                    let formula_len = usize::try_from(binary::read_u32_le_at(&record.data, 0)?)
                        .map_err(|_| {
                            crate::xlsb::error::XlsbError::InvalidFormula(
                                "BrtSupNameFmla size overflow".to_string(),
                            )
                        })?;
                    let expected = formula_len.checked_add(4).ok_or_else(|| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtSupNameFmla size overflow".to_string(),
                        )
                    })?;
                    if record.data.len() != expected {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected,
                            found: record.data.len(),
                        });
                    }
                    sup_name_state = 2;
                },
                record_types::SUP_NAME_BITS => {
                    if sup_name_state != 2 || record.data.len() != 7 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected BrtSupNameBits".to_string(),
                        ));
                    }
                    sup_name_state = 3;
                },
                record_types::SUP_NAME_END => {
                    if sup_name_state != 3 || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtSupNameEnd".to_string(),
                        ));
                    }
                    sup_name_state = 0;
                },
                record_types::END_SUP_BOOK => {
                    if !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidLength {
                            expected: 0,
                            found: record.data.len(),
                        });
                    }
                    if sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtEndSupBook occurs inside an external-name block".to_string(),
                        ));
                    }
                    saw_end = true;
                },
                _ => {
                    if link_type == Some(0) && sup_name_state != 0 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unexpected record inside an external defined name".to_string(),
                        ));
                    }
                },
            }
        }
        let kind = link_type.ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "external link has no BrtBeginSupBook".to_string(),
            )
        })?;
        if !saw_end {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "external link has no BrtEndSupBook".to_string(),
            ));
        }
        if kind == 0 && !saw_sup_tabs {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "external workbook link has no BrtSupTabs".to_string(),
            ));
        }
        let (link_kind, source, detail, target) = match kind {
            1 => {
                if !part.rels().is_empty() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                        "DDE external link must not contain relationships".to_string(),
                    ));
                }
                let target = format!("{target_key}:{target_detail}");
                (
                    XlsbExternalLinkKind::Dde,
                    target_key,
                    Some(target_detail),
                    target,
                )
            },
            0 | 2 => {
                let relationship = part.rels().get(&target_key).ok_or_else(|| {
                    crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "external data relationship {target_key:?} is missing"
                    ))
                })?;
                if !relationship.is_external() {
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "external data relationship {target_key:?} is internal"
                    )));
                }
                let allowed_relationship_types = if kind == 0 {
                    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES
                } else {
                    OLE_DATA_SOURCE_RELATIONSHIP_TYPES
                };
                if !allowed_relationship_types.contains(&relationship.reltype()) {
                    let source_kind = if kind == 0 {
                        "external workbook"
                    } else {
                        "OLE data source"
                    };
                    return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                        "{source_kind} relationship {target_key:?} has invalid type {:?}",
                        relationship.reltype()
                    )));
                }
                let target = relationship.target_ref().to_string();
                let link_kind = if kind == 0 {
                    XlsbExternalLinkKind::Workbook
                } else {
                    XlsbExternalLinkKind::Ole
                };
                let detail = if kind == 2 { Some(target_detail) } else { None };
                (link_kind, target.clone(), detail, target)
            },
            _ => unreachable!("external link kind was validated above"),
        };
        Ok(FormulaExternalBook {
            kind: link_kind,
            source,
            detail,
            target,
            sheet_names: sheet_names.into(),
            defined_names: defined_names.into(),
            is_workbook: kind == 0,
        })
    }

    fn parse_external_sheet_names(data: &[u8]) -> XlsbResult<Vec<String>> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "external sheet-name count overflow".to_string(),
            )
        })?;
        if count >= 65_535 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "external sheet-name count {count} exceeds 65,534"
            )));
        }
        let mut names = Vec::with_capacity(count);
        let mut offset = 4;
        for _ in 0..count {
            let (name, consumed) = crate::xlsb::records::wide_str_with_len(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "external sheet-name size overflow".to_string(),
                )
            })?;
            let name_len = name.encode_utf16().count();
            if name_len == 0
                || name_len > 31
                || name.contains(['\0', '\u{0003}', ':', '\\', '*', '?', '/', '[', ']'])
                || name.starts_with('\'')
                || name.ends_with('\'')
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "external sheet name {name:?} does not follow sheet-name grammar"
                )));
            }
            if names
                .iter()
                .any(|existing: &String| excel_name_eq(existing, &name))
            {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "duplicate external sheet name {name:?}"
                )));
            }
            names.push(name);
        }
        if offset != data.len() {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtSupTabs has {} trailing bytes",
                data.len() - offset
            )));
        }
        Ok(names)
    }

    fn parse_nullable_wide_string(data: &[u8]) -> XlsbResult<(Option<String>, usize)> {
        if data.len() < 4 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        if binary::read_u32_le_at(data, 0)? == u32::MAX {
            Ok((None, 4))
        } else {
            let (value, consumed) = crate::xlsb::records::wide_str_with_len(data)?;
            Ok((Some(value), consumed))
        }
    }

    fn parse_pivot_cache_ids(data: &[u8]) -> XlsbResult<Vec<(u32, String)>> {
        const BEGIN_PIVOT_CACHE_IDS: u16 = 384;
        const END_PIVOT_CACHE_IDS: u16 = 385;
        const BEGIN_PIVOT_CACHE_ID: u16 = 386;
        const END_PIVOT_CACHE_ID: u16 = 387;

        let mut in_collection = false;
        let mut open_cache = false;
        let mut ended = false;
        let mut caches = Vec::new();
        for record in XlsbRecordIter::new(data) {
            let record = record?;
            match record.header.record_type {
                BEGIN_PIVOT_CACHE_IDS => {
                    if in_collection || ended {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "duplicate BrtBeginPivotCacheIDs collection".to_string(),
                        ));
                    }
                    in_collection = true;
                },
                BEGIN_PIVOT_CACHE_ID => {
                    if !in_collection || open_cache || record.data.len() < 8 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "malformed BrtBeginPivotCacheID nesting or payload".to_string(),
                        ));
                    }
                    let cache_id = binary::read_u32_le_at(&record.data, 0)?;
                    let (rel_id, consumed) =
                        crate::xlsb::records::wide_str_with_len(&record.data[4..])?;
                    if 4 + consumed != record.data.len()
                        || rel_id.is_empty()
                        || rel_id.encode_utf16().count() > 255
                        || caches
                            .iter()
                            .any(|(existing, _): &(u32, String)| *existing == cache_id)
                    {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                            "invalid or duplicate PivotCache ID {cache_id}"
                        )));
                    }
                    caches.push((cache_id, rel_id));
                    open_cache = true;
                },
                END_PIVOT_CACHE_ID => {
                    if !open_cache || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unbalanced BrtEndPivotCacheID".to_string(),
                        ));
                    }
                    open_cache = false;
                },
                END_PIVOT_CACHE_IDS => {
                    if !in_collection || open_cache || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unbalanced BrtEndPivotCacheIDs".to_string(),
                        ));
                    }
                    in_collection = false;
                    ended = true;
                },
                _ => {},
            }
        }
        if in_collection || open_cache {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "unterminated PivotCache ID collection".to_string(),
            ));
        }
        Ok(caches)
    }

    fn parse_pivot_view(data: &[u8], sheet_index: usize) -> XlsbResult<FormulaPivotViewDefinition> {
        const BEGIN_SX_VIEW: u16 = 280;
        let mut view = None;
        for record in XlsbRecordIter::new(data) {
            let record = record?;
            if record.header.record_type != BEGIN_SX_VIEW {
                continue;
            }
            if view.is_some() || record.data.len() < 36 {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "PivotTable part has duplicate or truncated BrtBeginSXView".to_string(),
                ));
            }
            let cache_id = binary::read_u32_le_at(&record.data, 28)?;
            let (name, consumed) = crate::xlsb::records::wide_str_with_len(&record.data[32..])?;
            if consumed > record.data.len() - 32 {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "PivotTable view name overruns BrtBeginSXView".to_string(),
                ));
            }
            view = Some(FormulaPivotViewDefinition::try_new(
                cache_id,
                sheet_index,
                name,
            )?);
        }
        view.ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "PivotTable part omits BrtBeginSXView".to_string(),
            )
        })
    }

    fn parse_table_definition(
        data: &[u8],
        sheet_index: usize,
    ) -> XlsbResult<FormulaTableDefinition> {
        const BEGIN_LIST: u16 = 343;
        const END_LIST: u16 = 344;
        const BEGIN_LIST_COLS: u16 = 345;
        const END_LIST_COLS: u16 = 346;
        const BEGIN_LIST_COL: u16 = 347;
        const END_LIST_COL: u16 = 348;

        let mut table_header: Option<(u32, String, usize)> = None;
        let mut expected_columns = None;
        let mut columns = Vec::new();
        let mut in_column = false;
        let mut ended_columns = false;
        let mut ended_table = false;
        let mut iter = XlsbRecordIter::new(data);
        for record in iter.by_ref() {
            let record = record?;
            if ended_table {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "XLSB table part contains records after BrtEndList".to_string(),
                ));
            }
            match record.header.record_type {
                BEGIN_LIST => {
                    if table_header.is_some() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "XLSB table part contains duplicate BrtBeginList".to_string(),
                        ));
                    }
                    table_header = Some(Self::parse_table_header(&record.data)?);
                },
                BEGIN_LIST_COLS => {
                    let (_, _, range_columns) = table_header.as_ref().ok_or_else(|| {
                        crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtBeginListCols precedes BrtBeginList".to_string(),
                        )
                    })?;
                    if expected_columns.is_some() || record.data.len() != 4 {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid or duplicate BrtBeginListCols".to_string(),
                        ));
                    }
                    let count = usize::try_from(binary::read_u32_le_at(&record.data, 0)?).map_err(
                        |_| {
                            crate::xlsb::error::XlsbError::InvalidFormula(
                                "table column count overflow".to_string(),
                            )
                        },
                    )?;
                    if count == 0 || count > 16_384 || count != *range_columns {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                            "table column count {count} disagrees with range width {range_columns}"
                        )));
                    }
                    expected_columns = Some(count);
                },
                BEGIN_LIST_COL => {
                    if expected_columns.is_none() || ended_columns || in_column {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "BrtBeginListCol occurs outside its column collection".to_string(),
                        ));
                    }
                    columns.push(Self::parse_table_column(&record.data, columns.len())?);
                    in_column = true;
                },
                END_LIST_COL => {
                    if !in_column || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "unmatched or nonempty BrtEndListCol".to_string(),
                        ));
                    }
                    in_column = false;
                },
                END_LIST_COLS => {
                    if expected_columns.is_none()
                        || in_column
                        || ended_columns
                        || !record.data.is_empty()
                    {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtEndListCols".to_string(),
                        ));
                    }
                    ended_columns = true;
                },
                END_LIST => {
                    if !ended_columns || in_column || !record.data.is_empty() {
                        return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                            "invalid BrtEndList".to_string(),
                        ));
                    }
                    ended_table = true;
                },
                _ => {},
            }
        }
        let (table_id, display_name, _) = table_header.ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "XLSB table part omits BrtBeginList".to_string(),
            )
        })?;
        let expected = expected_columns.ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "XLSB table part omits BrtBeginListCols".to_string(),
            )
        })?;
        if !ended_table || columns.len() != expected {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "XLSB table contains {} of {expected} declared columns or is unterminated",
                columns.len()
            )));
        }
        FormulaTableDefinition::try_new(table_id, sheet_index, display_name, columns)
    }

    fn parse_table_header(data: &[u8]) -> XlsbResult<(u32, String, usize)> {
        if data.len() < 64 {
            return Err(crate::xlsb::error::XlsbError::InvalidLength {
                expected: 64,
                found: data.len(),
            });
        }
        let row_first = binary::read_u32_le_at(data, 0)?;
        let row_last = binary::read_u32_le_at(data, 4)?;
        let col_first = binary::read_u32_le_at(data, 8)?;
        let col_last = binary::read_u32_le_at(data, 12)?;
        if row_first > row_last
            || row_last >= 1_048_576
            || col_first > col_last
            || col_last >= 16_384
        {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                "BrtBeginList contains an invalid table range".to_string(),
            ));
        }
        for offset in [24, 28] {
            if binary::read_u32_le_at(data, offset)? > 1 {
                return Err(crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtBeginList contains a non-Boolean row flag".to_string(),
                ));
            }
        }
        let table_id = binary::read_u32_le_at(data, 20)?;
        let mut offset = 64;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtBeginList string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtBeginList has {} trailing bytes",
                data.len() - offset
            )));
        }
        let display_name = strings[1].clone().ok_or_else(|| {
            crate::xlsb::error::XlsbError::InvalidFormula(
                "BrtBeginList has a NULL display name".to_string(),
            )
        })?;
        Ok((
            table_id,
            display_name,
            usize::try_from(col_last - col_first + 1).expect("bounded table width"),
        ))
    }

    fn parse_table_column(data: &[u8], index: usize) -> XlsbResult<String> {
        if data.len() < 24 || binary::read_u32_le_at(data, 0)? == 0 {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtBeginListCol {index} has an invalid header"
            )));
        }
        let mut offset = 24;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(
                    "BrtBeginListCol string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::xlsb::error::XlsbError::InvalidFormula(format!(
                "BrtBeginListCol has {} trailing bytes",
                data.len() - offset
            )));
        }
        strings[0]
            .clone()
            .or_else(|| strings[1].clone())
            .ok_or_else(|| {
                crate::xlsb::error::XlsbError::InvalidFormula(format!(
                    "BrtBeginListCol {index} has neither a name nor caption"
                ))
            })
    }
}

impl litchi_core::sheet::WorkbookTrait for XlsbWorkbook {
    fn active_sheet_index(&self) -> usize {
        0
    }

    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>> {
        self.worksheet_by_index(0)
    }

    fn worksheet_count(&self) -> usize {
        self.formula_context.worksheet_names.len()
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.formula_context.worksheet_names
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn SheetTrait + '_>> {
        let worksheet = self.worksheet(index)?;
        Ok(Box::new(worksheet))
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn SheetTrait + '_>> {
        for (i, ws_name) in self.formula_context.worksheet_names.iter().enumerate() {
            if ws_name == name {
                return self.worksheet_by_index(i);
            }
        }
        Err(Box::new(crate::error::OoxmlError::InvalidFormat(format!(
            "Worksheet '{}' not found",
            name
        ))))
    }

    fn worksheets<'a>(&'a self) -> Box<dyn WorksheetIterator<'a> + 'a> {
        Box::new(XlsbWorksheetIterator {
            workbook: self,
            index: 0,
        })
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904
    }
}

pub struct XlsbWorksheetIterator<'a> {
    workbook: &'a XlsbWorkbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsbWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>>> {
        if self.index < self.workbook.formula_context.worksheet_names.len() {
            match self.workbook.worksheet(self.index) {
                Ok(worksheet) => {
                    self.index += 1;
                    Some(Ok(Box::new(worksheet)))
                },
                Err(e) => {
                    self.index += 1; // Continue to next worksheet even on error
                    Some(Err(Box::new(e)))
                },
            }
        } else {
            None
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::formula::{FormulaConverter, FormulaParser};
    use crate::xlsb::writer::RecordWriter;
    use litchi_core::sheet::{Cell, Worksheet};
    use litchi_opc::part::Part;
    use litchi_opc::{BlobPart, PackURI};
    use std::fs::File;

    fn wide_string(value: &str) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
        for unit in value.encode_utf16() {
            data.extend_from_slice(&unit.to_le_bytes());
        }
        data
    }

    fn external_link_records(records: &[(u16, Vec<u8>)]) -> Vec<u8> {
        let mut data = Vec::new();
        let mut writer = RecordWriter::new(&mut data);
        for (record_type, payload) in records {
            writer.write_record(*record_type, payload).unwrap();
        }
        data
    }

    fn parse_external_link(records: &[(u16, Vec<u8>)]) -> XlsbResult<FormulaExternalBook> {
        parse_external_link_with_relationship_type(
            records,
            Some(relationship_type::EXTERNAL_LINK_PATH),
        )
    }

    fn parse_external_link_with_relationship_type(
        records: &[(u16, Vec<u8>)],
        target_relationship_type: Option<&str>,
    ) -> XlsbResult<FormulaExternalBook> {
        let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let mut part = BlobPart::new(
            uri.clone(),
            "application/vnd.ms-excel.externalLink".to_string(),
            external_link_records(records),
        );
        if let Some(target_relationship_type) = target_relationship_type {
            part.rels_mut().add_relationship(
                target_relationship_type.to_string(),
                "Book.xlsx".to_string(),
                "rIdPath".to_string(),
                true,
            );
        }
        let mut package = OpcPackage::new();
        package.add_part(Box::new(part));
        let workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calculation_properties: CalculationProperties::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
        };
        workbook.load_external_book(&uri)
    }

    fn parse_shared_string_records(records: &[(u16, Vec<u8>)]) -> XlsbResult<Vec<SharedString>> {
        let data = external_link_records(records);
        let mut iter = XlsbRecordIter::new(data.as_slice());
        let mut strings = Vec::new();
        XlsbWorkbook::read_shared_strings(&mut iter, &mut strings)?;
        Ok(strings)
    }

    fn external_workbook_records() -> Vec<(u16, Vec<u8>)> {
        let mut begin = 0u16.to_le_bytes().to_vec();
        begin.extend_from_slice(&wide_string("rIdPath"));
        begin.extend_from_slice(&u32::MAX.to_le_bytes());
        let mut tabs = 1u32.to_le_bytes().to_vec();
        tabs.extend_from_slice(&wide_string("Data Sheet"));
        vec![
            (record_types::BEGIN_SUP_BOOK, begin),
            (record_types::SUP_TABS, tabs),
            (record_types::SUP_NAME_START, wide_string("Rate")),
            (record_types::SUP_NAME_FORMULA, 0u32.to_le_bytes().to_vec()),
            (record_types::SUP_NAME_BITS, vec![0; 7]),
            (record_types::SUP_NAME_END, Vec::new()),
            (record_types::END_SUP_BOOK, Vec::new()),
        ]
    }

    fn external_data_source_records(
        kind: u16,
        source: &str,
        detail: &str,
        item_name: &str,
    ) -> Vec<(u16, Vec<u8>)> {
        assert!(matches!(kind, 1 | 2));
        let mut begin = kind.to_le_bytes().to_vec();
        begin.extend_from_slice(&wide_string(source));
        begin.extend_from_slice(&wide_string(detail));
        vec![
            (record_types::BEGIN_SUP_BOOK, begin),
            (record_types::SUP_NAME_START, wide_string(item_name)),
            (record_types::SUP_NAME_BITS, vec![0; 7]),
            (record_types::SUP_NAME_END, Vec::new()),
            (record_types::END_SUP_BOOK, Vec::new()),
        ]
    }

    #[test]
    fn reads_formula_records_from_real_workbook_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/universal-content.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(path).unwrap()).unwrap();
        let mut formula_cells = Vec::new();
        for index in 0..workbook.formula_context.worksheet_names.len() {
            let worksheet = workbook.worksheet(index).unwrap();
            if let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() {
                for row in min_row..=max_row {
                    for col in min_col..=max_col {
                        let Some(cell) = worksheet.get_cell(row, col) else {
                            continue;
                        };
                        if cell.is_formula() {
                            formula_cells.push((
                                worksheet.name().to_string(),
                                cell.coordinate(),
                                cell.value().clone(),
                                cell.formula_bytes().unwrap().to_vec(),
                            ));
                        }
                    }
                }
            }
        }
        assert_eq!(formula_cells.len(), 4);
        let formulas: Vec<_> = formula_cells
            .iter()
            .map(|cell| match &cell.2 {
                litchi_core::sheet::CellValue::Formula {
                    formula,
                    cached_value,
                    ..
                } => (cell.1.as_str(), formula.as_str(), cached_value.as_deref()),
                value => panic!("expected decoded formula, found {value:?}"),
            })
            .collect();
        assert_eq!(formulas[0].0, "C1");
        assert_eq!(formulas[0].1, "(2*3)");
        assert_eq!(formulas[1].1, "(2+3)");
        assert_eq!(formulas[2].1, "(2-3)");
        assert_eq!(formulas[3].1, "(C1+C2)");
        assert!(matches!(
            formulas[3].2,
            Some(litchi_core::sheet::CellValue::Float(11.0))
        ));
        assert!(formula_cells.iter().all(|cell| !cell.3.is_empty()));
    }

    #[test]
    fn reads_conditional_formatting_from_real_workbook_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/cond_format.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(path).unwrap()).unwrap();
        let worksheet = workbook.worksheet(0).unwrap();
        let formatting = worksheet.conditional_formattings();
        assert_eq!(formatting.len(), 1);
        assert_eq!(formatting[0].ranges, ["E3:E18"]);
        assert!(!formatting[0].pivot_only);
        assert_eq!(formatting[0].rules.len(), 1);
        let rule = &formatting[0].rules[0];
        assert_eq!(
            rule.rule_type,
            crate::xlsb::conditional_formatting::CfRuleType::CellIs
        );
        assert_eq!(rule.template, 0);
        assert_eq!(rule.dxf_id, Some(0));
        assert_eq!(rule.priority, 1);
        assert_eq!(rule.parameter, 5);
        assert_eq!(rule.formula_texts, ["5"]);
    }

    #[test]
    fn reads_rich_and_phonetic_shared_strings_from_local_fixtures() {
        let rich_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/sample.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(rich_path).unwrap()).unwrap();
        let rich = workbook
            .shared_strings()
            .iter()
            .find(|value| !value.runs.is_empty())
            .expect("sample.xlsb should contain rich shared strings");
        assert_eq!(rich.text, "hello, xssf");
        assert_eq!(rich.runs[0].character_index, 0);
        let mut found_cell_text = false;
        for index in 0..workbook.formula_context.worksheet_names.len() {
            let worksheet = workbook.worksheet(index).unwrap();
            let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() else {
                continue;
            };
            for row in min_row..=max_row {
                for col in min_col..=max_col {
                    found_cell_text |= worksheet.get_cell(row, col).is_some_and(|cell| {
                        matches!(cell.value(), litchi_core::sheet::CellValue::String(value) if value == "hello, xssf")
                    });
                }
            }
        }
        assert!(
            found_cell_text,
            "rich SST text should remain the cell value"
        );

        let phonetic_path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/51519.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(phonetic_path).unwrap()).unwrap();
        let phonetic = workbook
            .shared_strings()
            .iter()
            .find_map(|value| {
                value
                    .phonetic
                    .as_ref()
                    .filter(|value| !value.runs.is_empty())
            })
            .expect("51519.xlsb should contain phonetic shared strings");
        assert_eq!(phonetic.font_id, 1);
        assert_eq!(
            phonetic.phonetic_type,
            crate::xlsb::PhoneticType::FullWidthKatakana
        );
        assert_eq!(phonetic.alignment, crate::xlsb::PhoneticAlignment::Left);
    }

    #[test]
    fn reads_binary_comments_from_real_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/comments.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(path).unwrap()).unwrap();
        let worksheet = workbook.worksheet(0).unwrap();
        assert_eq!(worksheet.comments().len(), 4);
        let first = &worksheet.comments()[0];
        assert_eq!((first.row, first.col), (0, 0));
        assert_eq!(first.author, "Sven Nissel");
        assert!(first.text.contains("comment top row1"));
        assert_eq!(first.runs.len(), 2);
    }

    #[test]
    fn validates_shared_string_stream_structure_and_counts() {
        let mut item = vec![0];
        item.extend_from_slice(&wide_string("value"));
        let valid = vec![
            (
                record_types::BEGIN_SST,
                [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
            ),
            (record_types::SST_ITEM, item.clone()),
            (record_types::END_SST, Vec::new()),
        ];
        let strings = parse_shared_string_records(&valid).unwrap();
        assert_eq!(strings[0].text, "value");

        let invalid_counts = vec![(
            record_types::BEGIN_SST,
            [0u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
        )];
        assert!(matches!(
            parse_shared_string_records(&invalid_counts),
            Err(crate::xlsb::error::XlsbError::Unrecognized { .. })
        ));

        let missing_item = vec![
            (
                record_types::BEGIN_SST,
                [2u32.to_le_bytes(), 2u32.to_le_bytes()].concat(),
            ),
            (record_types::SST_ITEM, item),
            (record_types::END_SST, Vec::new()),
        ];
        assert!(matches!(
            parse_shared_string_records(&missing_item),
            Err(crate::xlsb::error::XlsbError::Unrecognized { .. })
        ));

        let malformed_item = vec![
            (
                record_types::BEGIN_SST,
                [1u32.to_le_bytes(), 1u32.to_le_bytes()].concat(),
            ),
            (record_types::SST_ITEM, vec![1]),
            (record_types::END_SST, Vec::new()),
        ];
        assert!(parse_shared_string_records(&malformed_item).is_err());
    }

    #[test]
    fn resolves_cell_style_references_from_real_fixtures() {
        let mut saw_nondefault_style = false;
        for fixture in [
            "Simple.xlsb",
            "date.xlsb",
            "universal-content.xlsb",
            "cond_format.xlsb",
        ] {
            let path = format!(
                "{}/../../test-data/ooxml/xlsb/{fixture}",
                env!("CARGO_MANIFEST_DIR")
            );
            let workbook = XlsbWorkbook::new(File::open(path).unwrap())
                .unwrap_or_else(|error| panic!("{fixture}: {error}"));
            assert!(!workbook.styles().cell_xfs.is_empty(), "{fixture}");
            for index in 0..workbook.formula_context.worksheet_names.len() {
                let worksheet = workbook.worksheet(index).unwrap();
                if let Some((min_row, min_col, max_row, max_col)) = worksheet.dimensions() {
                    for row in min_row..=max_row {
                        for col in min_col..=max_col {
                            let Some(cell) = worksheet.get_cell(row, col) else {
                                continue;
                            };
                            saw_nondefault_style |= cell.style_id() != 0;
                            assert!(workbook.style_for_cell(cell).is_some(), "{fixture}");
                        }
                    }
                }
            }
        }
        assert!(saw_nondefault_style);
    }

    #[test]
    fn opens_custom_number_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/62815.xlsb"
        );
        let workbook = XlsbWorkbook::new(File::open(path).unwrap()).unwrap();
        assert!(workbook.styles().num_fmts.keys().any(|id| *id >= 164));
        let worksheet = workbook.worksheet(0).unwrap();
        assert!(worksheet.dimensions().is_some());
    }

    #[test]
    fn reads_external_book_metadata_from_local_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/ooxml/xlsb/bug66682.xlsb"
        );

        let package = OpcPackage::open(path).unwrap();
        let workbook = XlsbWorkbook {
            package,
            worksheets: Vec::new(),
            worksheet_rel_ids: Vec::new(),
            formula_context: FormulaResolutionContext::default(),
            shared_strings: Vec::new(),
            styles: StylesTable::default(),
            calculation_properties: CalculationProperties::default(),
            is_1904: false,
            pivot_cache_definitions: Vec::new(),
        };
        let uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let book = workbook.load_external_book(&uri).unwrap();
        assert!(book.is_workbook);
        assert_eq!(book.target, "ab");
        assert_eq!(&*book.sheet_names, &["ab"]);
    }

    #[test]
    fn parses_external_workbook_sheet_and_name_metadata() {
        let book = parse_external_link(&external_workbook_records()).unwrap();
        assert!(book.is_workbook);
        assert_eq!(book.target, "Book.xlsx");
        assert_eq!(&*book.sheet_names, &["Data Sheet"]);
        assert_eq!(&*book.defined_names, &["Rate"]);

        let link = book.metadata();
        assert_eq!(link.kind(), XlsbExternalLinkKind::Workbook);
        assert!(link.is_workbook());
        assert_eq!(link.source(), "Book.xlsx");
        assert_eq!(link.dde_topic(), None);
        assert_eq!(link.ole_program_id(), None);
        assert_eq!(link.sheet_names(), &["Data Sheet".to_string()]);
        assert_eq!(link.declared_names(), &["Rate".to_string()]);
    }

    #[test]
    fn exposes_inert_dde_and_ole_link_metadata() {
        let dde_records = external_data_source_records(1, "Excel", "System", "Rates Item");
        let dde = parse_external_link_with_relationship_type(&dde_records, None)
            .unwrap()
            .metadata();
        assert_eq!(dde.kind(), XlsbExternalLinkKind::Dde);
        assert!(!dde.is_workbook());
        assert_eq!(dde.source(), "Excel");
        assert_eq!(dde.dde_topic(), Some("System"));
        assert_eq!(dde.ole_program_id(), None);
        assert!(dde.sheet_names().is_empty());
        assert_eq!(dde.declared_names(), &["Rates Item".to_string()]);

        let ole_records =
            external_data_source_records(2, "rIdPath", "Acme.Server", "Report Item");
        let ole = parse_external_link_with_relationship_type(
            &ole_records,
            Some(relationship_type::OLE_OBJECT),
        )
        .unwrap()
        .metadata();
        assert_eq!(ole.kind(), XlsbExternalLinkKind::Ole);
        assert!(!ole.is_workbook());
        assert_eq!(ole.source(), "Book.xlsx");
        assert_eq!(ole.dde_topic(), None);
        assert_eq!(ole.ole_program_id(), Some("Acme.Server"));
        assert!(ole.sheet_names().is_empty());
        assert_eq!(ole.declared_names(), &["Report Item".to_string()]);
    }

    #[test]
    fn validates_external_link_relationship_types() {
        assert!(matches!(
            parse_external_link_with_relationship_type(
                &external_workbook_records(),
                Some(relationship_type::OLE_OBJECT),
            ),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let dde_records = external_data_source_records(1, "Excel", "System", "Rates");
        assert!(matches!(
            parse_external_link_with_relationship_type(
                &dde_records,
                Some(relationship_type::EXTERNAL_LINK_PATH),
            ),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let ole_records = external_data_source_records(2, "rIdPath", "Acme.Server", "Report");
        assert!(matches!(
            parse_external_link_with_relationship_type(
                &ole_records,
                Some(relationship_type::EXTERNAL_LINK_PATH),
            ),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));
    }

    #[test]
    fn resolves_external_formula_tokens_from_package_relationships() {
        let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
        let mut workbook_data = Vec::new();
        {
            let mut writer = RecordWriter::new(&mut workbook_data);
            writer
                .write_record(record_types::SUP_BOOK_SRC, &wide_string("rIdExternal"))
                .unwrap();
            let mut extern_sheet = 1u32.to_le_bytes().to_vec();
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            extern_sheet.extend_from_slice(&0u32.to_le_bytes());
            writer
                .write_record(record_types::EXTERN_SHEET, &extern_sheet)
                .unwrap();
        }
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
            workbook_data,
        );
        workbook_part.rels_mut().add_relationship(
            relationship_type::EXTERNAL_LINK.to_string(),
            "externalLinks/externalLink1.bin".to_string(),
            "rIdExternal".to_string(),
            false,
        );

        let external_uri = PackURI::new("/xl/externalLinks/externalLink1.bin").unwrap();
        let mut external_part = BlobPart::new(
            external_uri,
            "application/vnd.ms-excel.externalLink".to_string(),
            external_link_records(&external_workbook_records()),
        );
        external_part.rels_mut().add_relationship(
            relationship_type::EXTERNAL_LINK_PATH.to_string(),
            "Book.xlsx".to_string(),
            "rIdPath".to_string(),
            true,
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(external_part));
        let workbook = XlsbWorkbook::from_opc_package(package).unwrap();

        let links = workbook.external_links();
        assert_eq!(links.len(), 1);
        let link = &links[0];
        assert_eq!(link.kind(), XlsbExternalLinkKind::Workbook);
        assert_eq!(link.source(), "Book.xlsx");
        assert_eq!(link.sheet_names(), &["Data Sheet".to_string()]);
        assert_eq!(link.declared_names(), &["Rate".to_string()]);

        let reference = FormulaParser::new(&[0x5A, 0, 0, 0, 0, 0, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(
                &reference,
                &workbook.formula_context
            )
            .unwrap(),
            "'[Book.xlsx]Data Sheet'!$A$1"
        );
        let name = FormulaParser::new(&[0x59, 0, 0, 1, 0, 0, 0])
            .parse()
            .unwrap();
        assert_eq!(
            FormulaConverter::try_tokens_to_string_with_context(&name, &workbook.formula_context)
                .unwrap(),
            "'[Book.xlsx]'!Rate"
        );
    }

    #[test]
    fn rejects_malformed_external_workbook_record_sequences() {
        let mut duplicate_tabs = external_workbook_records();
        duplicate_tabs.insert(2, duplicate_tabs[1].clone());
        assert!(matches!(
            parse_external_link(&duplicate_tabs),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let mut unclosed_name = external_workbook_records();
        unclosed_name.remove(5);
        assert!(matches!(
            parse_external_link(&unclosed_name),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));

        let mut trailing_record = external_workbook_records();
        trailing_record.push((record_types::SUP_NAME_END, Vec::new()));
        assert!(matches!(
            parse_external_link(&trailing_record),
            Err(crate::xlsb::error::XlsbError::InvalidFormula(_))
        ));
    }

    #[test]
    fn loads_typed_pivot_cache_definitions_from_package_relationships() {
        // workbook.bin declares one PivotCache (idSx 12) related to a
        // pivotCacheDefinition part.
        let mut cache_id = 12u32.to_le_bytes().to_vec();
        cache_id.extend_from_slice(&wide_string("rIdCache"));
        let workbook_data = external_link_records(&[
            (record_types::BEGIN_PIVOT_CACHE_IDS, Vec::new()),
            (record_types::BEGIN_PIVOT_CACHE_ID, cache_id),
            (record_types::END_PIVOT_CACHE_ID, Vec::new()),
            (record_types::END_PIVOT_CACHE_IDS, Vec::new()),
        ]);
        let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
        let mut workbook_part = BlobPart::new(
            workbook_uri,
            "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
            workbook_data,
        );
        workbook_part.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"
                .to_string(),
            "pivotCache/pivotCacheDefinition1.bin".to_string(),
            "rIdCache".to_string(),
            false,
        );

        // Minimal worksheet-range PivotCache definition stream.
        let mut definition = vec![
            3,           // bVerCacheLastRefresh
            0,           // bVerCacheRefreshableMin
            2,           // bVerCacheCreated
            0b0001_0001, // fSaveData | fEnableRefresh
        ];
        definition.extend_from_slice(&(-1i32).to_le_bytes()); // citmGhostMax
        definition.extend_from_slice(&44_000.0f64.to_le_bytes()); // xnumRefreshedDate
        definition.push(0x00); // no optional strings
        definition.extend_from_slice(&5u32.to_le_bytes()); // cRecords
        definition.extend_from_slice(&[0; 4]); // unused (fLoadRefreshedWho = 0)
        let mut source = Vec::new();
        source.extend_from_slice(&0u32.to_le_bytes()); // iSrcType = sheet
        source.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
        let mut range = vec![0x00, 0x00, 0b0000_0010]; // fLoadSheet
        range.extend_from_slice(&wide_string("Data"));
        for value in [0i32, 9, 0, 3] {
            range.extend_from_slice(&value.to_le_bytes());
        }
        let definition_part = BlobPart::new(
            PackURI::new("/xl/pivotCache/pivotCacheDefinition1.bin").unwrap(),
            "application/vnd.ms-excel.pivotCacheDefinition".to_string(),
            external_link_records(&[
                (record_types::BEGIN_PIVOT_CACHE_DEF, definition),
                (record_types::BEGIN_PCD_SOURCE, source),
                (record_types::BEGIN_PCDS_RANGE, range),
                (record_types::END_PCDS_RANGE, Vec::new()),
                (record_types::END_PCD_SOURCE, Vec::new()),
                (record_types::END_PIVOT_CACHE_DEF, Vec::new()),
            ]),
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(workbook_part));
        package.add_part(Box::new(definition_part));
        let workbook = XlsbWorkbook::from_opc_package(package).unwrap();

        let definitions = workbook.pivot_cache_definitions();
        assert_eq!(definitions.len(), 1);
        assert_eq!(definitions[0].0, 12);
        let definition = workbook.pivot_cache_definition(12).unwrap();
        assert!(definition.save_data);
        assert_eq!(definition.record_count, 5);
        let source = definition.source.as_ref().unwrap();
        assert_eq!(
            source.source_type,
            crate::xlsb::pivot::PivotCacheSourceType::Worksheet
        );
        let worksheet = source.worksheet.as_ref().unwrap();
        assert_eq!(worksheet.sheet_name.as_deref(), Some("Data"));
        assert_eq!(
            worksheet.range,
            Some(crate::xlsb::pivot::PivotCacheRange {
                first_row: 0,
                last_row: 9,
                first_column: 0,
                last_column: 3,
            })
        );
        assert!(workbook.pivot_cache_definition(99).is_none());
    }
}
