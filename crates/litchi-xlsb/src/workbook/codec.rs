//! BIFF12/XLSB stream codecs and worksheet materialization.

use super::model::Workbook;
use crate::calc::{self, Props};
use crate::named_ranges::{Definition, validate_name};
use crate::package::error::Result;
use crate::package::external_link::{
    CachedValue, DATA_ITEM_REQUIRED_TRAILING_FLAG, DATA_ITEM_WANT_ADVISE, DATA_ITEM_WANT_PICTURE,
    DDE_ITEM_RESERVED_MASK, DDE_ITEM_SUPPORTS_OLE, DdeItem, DefinedName, EXTERNAL_NAME_BUILT_IN,
    EXTERNAL_NAME_RESERVED_MASK, EXTERNAL_REFERENCE_DDE, EXTERNAL_REFERENCE_OLE,
    EXTERNAL_REFERENCE_WORKBOOK, Entries, ErrorValue, Kind, Link, MAX_XLSB_EXTERNAL_CACHED_VALUES,
    NameFormula, OLE_ITEM_DISPLAY_AS_ICON, OLE_ITEM_REQUIRED_CLASS_FLAG, OLE_ITEM_RESERVED_MASK,
    OleItem, ValueMatrix,
};
use crate::package::formula::{
    Context, ExternalBook, ExternalSheet, SupportingLink, View, excel_name_eq,
    table::Definition as TableDefinition,
};
use crate::package::merged_cells::{MAX_MERGED_CELL_RANGES, MergedCell};
use crate::package::shared_strings::SharedString;
use crate::raw::{Records, kind};
use crate::sheet::Worksheet;
use litchi_core::binary;
use litchi_ooxml_common::external_link::EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES;
use litchi_opc::constants::relationship_type;
use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap};
use std::io::Cursor;

/// OLE data-source relationship types documented by MS-XLSB and MS-OI29500.
const OLE_DATA_SOURCE_RELATIONSHIP_TYPES: &[&str] = &[
    relationship_type::OLE_OBJECT,
    relationship_type::STRICT_OLE_OBJECT,
    "http://schemas.microsoft.com/office/2019/04/relationships/oleObjectLinkLongPath",
];

#[derive(Default)]
pub(super) struct ParsedWorkbookInfo {
    pub(super) worksheet_names: Vec<String>,
    pub(super) worksheet_rel_ids: Vec<Option<String>>,
    pub(super) worksheet_states: Vec<u32>,
    pub(super) supporting_links: Vec<SupportingLink>,
    pub(super) external_sheets: Vec<ExternalSheet>,
    pub(super) external_link_rel_ids: Vec<String>,
    pub(super) defined_names: Vec<String>,
    pub(super) is_1904: bool,
    pub(super) calc: Option<Props>,
}

#[derive(Debug)]
struct MergeBlockLayout {
    ranges: Vec<MergedCell>,
    block_span: Option<(usize, usize)>,
    insertion_offset: usize,
}

impl Workbook {
    pub fn merged_cell_ranges(&self, worksheet_index: usize) -> Result<Vec<MergedCell>> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let part = self.package.get_part(&uri)?;
        Ok(Self::inspect_merge_block(part.blob())?.ranges)
    }

    /// List merged ranges in a worksheet selected by exact name.
    pub fn merged_cell_ranges_by_name(&self, worksheet_name: &str) -> Result<Vec<MergedCell>> {
        self.merged_cell_ranges(self.worksheet_index(worksheet_name)?)
    }

    /// Atomically replace all merged ranges in a worksheet selected by index.
    pub fn set_merged_cell_ranges(
        &mut self,
        worksheet_index: usize,
        ranges: &[MergedCell],
    ) -> Result<()> {
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
            .ok_or(crate::package::error::Error::InvalidLength {
                expected: usize::MAX,
                found: original.len(),
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
    ) -> Result<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.set_merged_cell_ranges(index, ranges)
    }

    /// Atomically add one merged range to a worksheet selected by index.
    pub fn add_merged_cell_range(
        &mut self,
        worksheet_index: usize,
        range: MergedCell,
    ) -> Result<()> {
        let mut ranges = self.merged_cell_ranges(worksheet_index)?;
        ranges.push(range);
        self.set_merged_cell_ranges(worksheet_index, &ranges)
    }

    /// Atomically add one merged range to a worksheet selected by name.
    pub fn add_merged_cell_range_by_name(
        &mut self,
        worksheet_name: &str,
        range: MergedCell,
    ) -> Result<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.add_merged_cell_range(index, range)
    }

    /// Atomically remove an exact merged range from a worksheet by index.
    pub fn remove_merged_cell_range(
        &mut self,
        worksheet_index: usize,
        range: &MergedCell,
    ) -> Result<bool> {
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
    ) -> Result<bool> {
        let index = self.worksheet_index(worksheet_name)?;
        self.remove_merged_cell_range(index, range)
    }

    /// Atomically clear all merged ranges in a worksheet selected by index.
    pub fn clear_merged_cell_ranges(&mut self, worksheet_index: usize) -> Result<()> {
        self.set_merged_cell_ranges(worksheet_index, &[])
    }

    /// Atomically clear all merged ranges in a worksheet selected by name.
    pub fn clear_merged_cell_ranges_by_name(&mut self, worksheet_name: &str) -> Result<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.clear_merged_cell_ranges(index)
    }

    /// Open an XLSB workbook from a reader
    fn worksheet_index(&self, worksheet_name: &str) -> Result<usize> {
        self.formula_context
            .worksheet_names
            .iter()
            .position(|name| name == worksheet_name)
            .ok_or_else(|| {
                crate::package::error::Error::WorksheetNotFound(worksheet_name.to_string())
            })
    }

    fn worksheet_uri(&self, index: usize) -> Result<litchi_opc::PackURI> {
        let name = self
            .formula_context
            .worksheet_names
            .get(index)
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormat(format!(
                    "Worksheet index {index} out of bounds"
                ))
            })?;
        let rel_id = self
            .worksheet_rel_ids
            .get(index)
            .and_then(Option::as_deref)
            .ok_or_else(|| {
                crate::package::error::Error::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;
        let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
            crate::package::error::Error::FileNotFound(format!(
                "relationship {rel_id:?} for sheet {name:?}"
            ))
        })?;
        if relationship.is_external() {
            return Err(crate::package::error::Error::UnsupportedFeature(format!(
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

    fn normalize_merge_ranges(ranges: &[MergedCell]) -> Result<Vec<MergedCell>> {
        if ranges.len() > MAX_MERGED_CELL_RANGES {
            return Err(crate::package::error::Error::InvalidLength {
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
    ) -> Result<()> {
        if ranges.len() > MAX_MERGED_CELL_RANGES {
            return Err(crate::package::error::Error::InvalidLength {
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
                return Err(crate::package::error::Error::Unrecognized {
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
                if active
                    .get(&col_first)
                    .is_some_and(|entry| entry.1 == row_last)
                {
                    active.remove(&col_first);
                }
            }
            if let Some((&col_first, &(col_last, _))) = active.range(..=range.col_last).next_back()
                && col_last >= range.col_first
            {
                return Err(crate::package::error::Error::InvalidCellReference(format!(
                    "merged range {} overlaps an existing range beginning in column {}",
                    range.to_range_string(),
                    col_first
                )));
            }
            active.insert(range.col_first, (range.col_last, range.row_last));
            expirations.push(Reverse((range.row_last, range.col_first)));
        }
        Ok(())
    }

    fn is_post_merge_record(record_type: crate::raw::Kind) -> bool {
        matches!(
            record_type,
            kind::PHONETIC_INFO
                | kind::H_LINK
                | kind::BEGIN_D_VALS
                | kind::BEGIN_D_VALS14
                | kind::BEGIN_COND_FORMATTING
                | kind::BEGIN_COND_FORMATTING14
                | kind::MARGINS
                | kind::PRINT_OPTIONS
                | kind::PAGE_SETUP
                | kind::BEGIN_HEADER_FOOTER
                | kind::DRAWING
                | kind::LEGACY_DRAWING
                | kind::LEGACY_DRAWING_HF
        )
    }

    fn inspect_merge_block(data: &[u8]) -> Result<MergeBlockLayout> {
        let mut records = Records::new(data);
        let mut begin_offset = None;
        let mut block_span = None;
        let mut declared_count = None;
        let mut ranges = Vec::new();
        let mut in_block = false;
        let mut saw_end_sheet_data = false;
        let mut end_sheet_offset = None;
        let mut first_post_merge_offset = None;
        while let Some(record) = records.next() {
            let record = record?;
            let start = record.offset();
            let end = records.offset();
            match record.kind() {
                kind::BEGIN_MERGE_CELLS => {
                    if in_block || begin_offset.is_some() || !saw_end_sheet_data {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginMergeCells".to_string(),
                            val: "duplicate, nested, or out-of-order record".to_string(),
                        });
                    }
                    if record.payload().len() != 4 {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 4,
                            found: record.payload().len(),
                        });
                    }
                    let count = binary::read_u32_le_at(record.payload(), 0)? as usize;
                    if count == 0 || count > MAX_MERGED_CELL_RANGES {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: MAX_MERGED_CELL_RANGES,
                            found: count,
                        });
                    }
                    begin_offset = Some(start);
                    declared_count = Some(count);
                    in_block = true;
                },
                kind::MERGE_CELL => {
                    if !in_block {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtMergeCell".to_string(),
                            val: "record occurs outside BrtBeginMergeCells".to_string(),
                        });
                    }
                    ranges.push(MergedCell::parse(record.payload())?);
                    if ranges.len() > declared_count.unwrap_or_default() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginMergeCells".to_string(),
                            val: "declared count is smaller than the record collection".to_string(),
                        });
                    }
                },
                kind::END_MERGE_CELLS => {
                    if !in_block || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtEndMergeCells".to_string(),
                            val: "orphan, duplicate, or nonempty record".to_string(),
                        });
                    }
                    if declared_count != Some(ranges.len()) {
                        return Err(crate::package::error::Error::Unrecognized {
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
                kind::END_SHEET_DATA => {
                    if in_block {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtMergeCells collection".to_string(),
                            val: "noncontiguous record collection".to_string(),
                        });
                    }
                    saw_end_sheet_data = true;
                },
                kind::END_SHEET => {
                    if in_block || end_sheet_offset.replace(start).is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtEndSheet".to_string(),
                            val: "duplicate or embedded in merge collection".to_string(),
                        });
                    }
                },
                record_type => {
                    if in_block {
                        return Err(crate::package::error::Error::Unrecognized {
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
            return Err(crate::package::error::Error::UnexpectedEndOfStream(
                "BrtMergeCells collection".to_string(),
            ));
        }
        let end_sheet_offset = end_sheet_offset.ok_or_else(|| {
            crate::package::error::Error::UnexpectedEndOfStream("BrtEndSheet".to_string())
        })?;
        if !saw_end_sheet_data {
            return Err(crate::package::error::Error::UnexpectedEndOfStream(
                "BrtEndSheetData".to_string(),
            ));
        }
        if block_span.is_some() {
            Self::validate_merge_range_collection(&ranges, true)?;
            if first_post_merge_offset
                .is_some_and(|offset| block_span.is_some_and(|(begin, _)| offset < begin))
            {
                return Err(crate::package::error::Error::Unrecognized {
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

    fn serialize_merge_block(ranges: &[MergedCell]) -> Result<Vec<u8>> {
        if ranges.is_empty() {
            return Ok(Vec::new());
        }
        let mut output = Vec::with_capacity(10 + ranges.len() * 19);
        let mut writer = crate::raw::Writer::new(&mut output);
        writer.write_record(
            kind::BEGIN_MERGE_CELLS,
            &(ranges.len() as u32).to_le_bytes(),
        )?;
        for range in ranges {
            writer.write_record(kind::MERGE_CELL, &range.serialize())?;
        }
        writer.write_record(kind::END_MERGE_CELLS, &[])?;
        Ok(output)
    }

    /// Load workbook information from workbook.bin
    pub fn worksheet(&self, index: usize) -> Result<Worksheet> {
        if index >= self.formula_context.worksheet_names.len() {
            return Err(crate::package::error::Error::InvalidFormat(format!(
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
                crate::package::error::Error::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_uri = litchi_opc::PackURI::new("/xl/workbook.bin")?;
        let workbook_part = self.package.get_part(&workbook_uri)?;
        let relationship = workbook_part.rels().get(rel_id).ok_or_else(|| {
            crate::package::error::Error::FileNotFound(format!(
                "relationship {rel_id:?} for sheet {name:?}"
            ))
        })?;
        if relationship.is_external() {
            return Err(crate::package::error::Error::UnsupportedFeature(format!(
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
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "worksheet comments relationship".to_string(),
                    val: "multiple relationships".to_string(),
                });
            }
            match first {
                Some(rel) if rel.is_external() => {
                    return Err(crate::package::error::Error::UnsupportedFeature(
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
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "Comments part".to_string(),
                    val: "relationships are not permitted".to_string(),
                });
            }
            for comment in crate::package::comments::read_comments(part.blob())? {
                worksheet.add_comment(comment);
            }
        }
        Ok(worksheet)
    }

    /// Read shared strings from SST
    pub(super) fn read_shared_strings(
        iter: &mut Records<'_>,
        strings: &mut Vec<SharedString>,
    ) -> Result<()> {
        let initial_count = strings.len();
        let mut expected_unique = None;
        let mut ended = false;
        for record in iter.by_ref() {
            let record = record?;
            match record.kind() {
                kind::BEGIN_SST => {
                    if expected_unique.is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginSst".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    if record.payload().len() != 8 {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 8,
                            found: record.payload().len(),
                        });
                    }
                    let total = binary::read_u32_le_at(record.payload(), 0)?;
                    let unique = binary::read_u32_le_at(record.payload(), 4)?;
                    if total > 0x7FFF_FFFF || unique > total {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginSst counts".to_string(),
                            val: format!("total={total}, unique={unique}"),
                        });
                    }
                    expected_unique = Some(unique as usize);
                },
                kind::SST_ITEM => {
                    let expected = expected_unique.ok_or_else(|| {
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtSSTItem".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found >= expected {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtSSTItem count".to_string(),
                            val: format!("more than declared {expected}"),
                        });
                    }
                    strings.push(SharedString::parse(record.payload())?);
                },
                kind::END_SST => {
                    if !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    let expected = expected_unique.ok_or_else(|| {
                        crate::package::error::Error::Unrecognized {
                            typ: "BrtEndSst".to_string(),
                            val: "record before BrtBeginSst".to_string(),
                        }
                    })?;
                    let found = strings.len() - initial_count;
                    if found != expected {
                        return Err(crate::package::error::Error::Unrecognized {
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
            return Err(crate::package::error::Error::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtBeginSst".to_string(),
            });
        }
        if !ended {
            return Err(crate::package::error::Error::Unrecognized {
                typ: "SST stream".to_string(),
                val: "missing BrtEndSst".to_string(),
            });
        }
        Ok(())
    }

    /// Read workbook structure
    pub(super) fn read_workbook(iter: &mut Records<'_>) -> Result<ParsedWorkbookInfo> {
        let mut info = ParsedWorkbookInfo::default();
        let worksheet_names = &mut info.worksheet_names;
        let worksheet_rel_ids = &mut info.worksheet_rel_ids;
        let worksheet_states = &mut info.worksheet_states;
        let supporting_links = &mut info.supporting_links;
        let external_sheets = &mut info.external_sheets;
        let external_link_rel_ids = &mut info.external_link_rel_ids;
        let defined_names = &mut info.defined_names;
        let is_1904 = &mut info.is_1904;
        for record in iter.by_ref() {
            let record = record?;
            match record.kind() {
                kind::WORKBOOK_PROP => {
                    if let Ok(prop) =
                        crate::package::records::WorkbookPropRecord::parse(record.payload())
                    {
                        *is_1904 = prop.is_date1904;
                    }
                },
                kind::CALC_PROP => {
                    if info.calc.is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtCalcProp".to_string(),
                            val: "duplicate record".to_string(),
                        });
                    }
                    info.calc = Some(calc::read(record.payload())?);
                },
                kind::BUNDLE_SH => {
                    let bundle_sh =
                        crate::package::records::BundleSheetRecord::parse(record.payload())?;
                    if worksheet_names
                        .iter()
                        .any(|name| excel_name_eq(name, &bundle_sh.name))
                    {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBundleSh strName".to_string(),
                            val: format!("duplicate sheet name {:?}", bundle_sh.name),
                        });
                    }
                    worksheet_names.push(bundle_sh.name);
                    worksheet_rel_ids.push(bundle_sh.rel_id);
                    worksheet_states.push(bundle_sh.state);
                },
                kind::SUP_SELF => {
                    supporting_links.push(SupportingLink::SelfWorkbook);
                },
                kind::SUP_SAME => {
                    supporting_links.push(SupportingLink::SameSheet);
                },
                kind::SUP_BOOK_SRC => {
                    let (rel_id, consumed) =
                        crate::package::records::decode_string(record.payload())?;
                    if rel_id.is_empty() || consumed != record.payload().len() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtSupBookSrc has an invalid relationship ID".to_string(),
                        ));
                    }
                    let book_index = u32::try_from(external_link_rel_ids.len()).map_err(|_| {
                        crate::package::error::Error::InvalidFormula(
                            "external-link count overflow".to_string(),
                        )
                    })?;
                    external_link_rel_ids.push(rel_id);
                    supporting_links.push(SupportingLink::ExternalWorkbook(book_index));
                },
                kind::SUP_ADDIN => {
                    supporting_links.push(SupportingLink::AddIn);
                },
                kind::EXTERN_SHEET => {
                    Self::parse_extern_sheet(record.payload(), external_sheets)?;
                },
                kind::NAME => {
                    let named_range = Definition::parse(record.payload())?;
                    if named_range
                        .sheet_id
                        .is_some_and(|index| index as usize >= worksheet_names.len())
                    {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
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
        formula_context: &Context,
        sheet_index: usize,
        cell_xf_count: usize,
    ) -> Result<Worksheet> {
        let mut worksheet = Worksheet::new(name);
        let iter = crate::package::records::Stream::new(cursor);
        let formula_context = formula_context.for_sheet(sheet_index);
        let mut cells_reader = crate::package::cells_reader::CellsReader::new(
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
        worksheet.set_web_extension_bindings(cells_reader.web_extension_bindings);
        worksheet.set_sheet_views(cells_reader.sheet_views);

        Ok(worksheet)
    }

    fn parse_extern_sheet(data: &[u8], external_sheets: &mut Vec<ExternalSheet>) -> Result<()> {
        if data.len() < 4 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::package::error::Error::InvalidFormula(
                "BrtExternSheet count overflow".to_string(),
            )
        })?;
        if count >= 65_536 {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtExternSheet count {count} exceeds 65,535"
            )));
        }
        let expected = 4usize
            .checked_add(count.checked_mul(12).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?)
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtExternSheet size overflow".to_string(),
                )
            })?;
        if data.len() != expected {
            return Err(crate::package::error::Error::InvalidLength {
                expected,
                found: data.len(),
            });
        }
        external_sheets.reserve(count);
        for chunk in data[4..].chunks_exact(12) {
            external_sheets.push(ExternalSheet {
                external_link: binary::read_u32_le_at(chunk, 0)?,
                first_sheet: binary::read_u32_le_at(chunk, 4)? as i32,
                last_sheet: binary::read_u32_le_at(chunk, 8)? as i32,
            });
        }
        Ok(())
    }

    pub(super) fn load_external_book(&self, uri: &litchi_opc::PackURI) -> Result<ExternalBook> {
        let part = self.package.get_part(uri)?;
        if part.content_type() != "application/vnd.ms-excel.externalLink" {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "external link part {uri} has invalid content type {:?}",
                part.content_type()
            )));
        }
        let mut iter = Records::new(part.blob());
        let mut link_type = None;
        let mut target_key = String::new();
        let mut target_detail = String::new();
        let mut sheet_names = Vec::new();
        let mut workbook_entries = Vec::new();
        let mut dde_entries = Vec::new();
        let mut ole_entries = Vec::new();
        let mut saw_sup_tabs = false;
        // 0 = outside a name, 1 = expect formula, 2 = expect bits,
        // 3 = expect end/value start, 4 = inside a cached matrix.
        let mut sup_name_state = 0u8;
        let mut current_name = None;
        let mut current_formula = None;
        let mut current_bits = None;
        let mut current_cache = None;
        let mut cache_dimensions = None;
        let mut cache_values = Vec::new();
        let mut saw_end = false;

        for record in &mut iter {
            let record = record?;
            if saw_end {
                return Err(crate::package::error::Error::InvalidFormula(
                    "external link has records after BrtEndSupBook".to_string(),
                ));
            }
            if link_type.is_none() && record.kind() != kind::BEGIN_SUP_BOOK {
                return Err(crate::package::error::Error::InvalidFormula(
                    "external link does not start with BrtBeginSupBook".to_string(),
                ));
            }
            match record.kind() {
                kind::BEGIN_SUP_BOOK => {
                    if link_type.is_some() || record.payload().len() < 10 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtBeginSupBook framing".to_string(),
                        ));
                    }
                    let kind = binary::read_u16_le_at(record.payload(), 0)?;
                    let (first, consumed) =
                        crate::package::records::decode_string(&record.payload()[2..])?;
                    let mut offset = 2 + consumed;
                    let (second, consumed) = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                        Self::parse_nullable_wide_string(&record.payload()[offset..])?
                    } else {
                        let (value, consumed) =
                            crate::package::records::decode_string(&record.payload()[offset..])?;
                        (Some(value), consumed)
                    };
                    offset += consumed;
                    if offset != record.payload().len()
                        || kind > EXTERNAL_REFERENCE_OLE
                        || first.is_empty()
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtBeginSupBook payload".to_string(),
                        ));
                    }
                    if kind == EXTERNAL_REFERENCE_WORKBOOK && second.is_some() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "external workbook BrtBeginSupBook string2 is not NULL".to_string(),
                        ));
                    }
                    link_type = Some(kind);
                    target_key = first;
                    target_detail = second.unwrap_or_default();
                },
                kind::SUP_TABS => {
                    if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK)
                        || saw_sup_tabs
                        || sup_name_state != 0
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupTabs".to_string(),
                        ));
                    }
                    sheet_names = Self::parse_external_sheet_names(record.payload())?;
                    saw_sup_tabs = true;
                },
                kind::SUP_NAME_START => {
                    let kind = link_type.ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "BrtSupNameStart precedes BrtBeginSupBook".to_string(),
                        )
                    })?;
                    if sup_name_state != 0 || (kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs)
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupNameStart".to_string(),
                        ));
                    }
                    let (name, consumed) =
                        crate::package::records::decode_string(record.payload())?;
                    if consumed != record.payload().len() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtSupNameStart has trailing bytes".to_string(),
                        ));
                    }
                    if kind == EXTERNAL_REFERENCE_WORKBOOK {
                        validate_name(&name)?;
                        sup_name_state = 1;
                    } else {
                        validate_name(&name)?;
                        sup_name_state = 2;
                    }
                    current_name = Some(name);
                },
                kind::SUP_NAME_FORMULA => {
                    if link_type != Some(EXTERNAL_REFERENCE_WORKBOOK)
                        || sup_name_state != 1
                        || record.payload().len() < 4
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupNameFmla".to_string(),
                        ));
                    }
                    let formula_len = usize::try_from(binary::read_u32_le_at(record.payload(), 0)?)
                        .map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "BrtSupNameFmla size overflow".to_string(),
                            )
                        })?;
                    let expected = formula_len.checked_add(4).ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "BrtSupNameFmla size overflow".to_string(),
                        )
                    })?;
                    if record.payload().len() != expected {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected,
                            found: record.payload().len(),
                        });
                    }
                    current_formula = if formula_len == 0 {
                        None
                    } else {
                        Some(NameFormula::from_tokens(record.payload()[4..].to_vec())?)
                    };
                    sup_name_state = 2;
                },
                kind::SUP_NAME_BITS => {
                    if sup_name_state != 2 || record.payload().len() != 7 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupNameBits".to_string(),
                        ));
                    }
                    let mut bits = [0u8; 7];
                    bits.copy_from_slice(record.payload());
                    Self::validate_external_name_bits(
                        link_type.expect("external link kind is present"),
                        &bits,
                    )?;
                    current_bits = Some(bits);
                    sup_name_state = 3;
                },
                kind::SUP_NAME_VALUE_START => {
                    if !matches!(
                        link_type,
                        Some(EXTERNAL_REFERENCE_DDE | EXTERNAL_REFERENCE_OLE)
                    ) || sup_name_state != 3
                        || record.payload().len() != 8
                        || current_cache.is_some()
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupNameValueStart".to_string(),
                        ));
                    }
                    let rows = binary::read_u32_le_at(record.payload(), 0)?;
                    let columns = binary::read_u32_le_at(record.payload(), 4)?;
                    let count = usize::try_from(rows)
                        .ok()
                        .and_then(|rows| {
                            usize::try_from(columns)
                                .ok()
                                .and_then(|columns| rows.checked_mul(columns))
                        })
                        .ok_or_else(|| {
                            crate::package::error::Error::InvalidFormula(
                                "external cached-value dimensions overflow".to_string(),
                            )
                        })?;
                    if count > MAX_XLSB_EXTERNAL_CACHED_VALUES {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: MAX_XLSB_EXTERNAL_CACHED_VALUES,
                            found: count,
                        });
                    }
                    cache_values.clear();
                    cache_values.reserve(count);
                    cache_dimensions = Some((rows, columns, count));
                    sup_name_state = 4;
                },
                kind::SUP_NAME_NIL
                | kind::SUP_NAME_NUM
                | kind::SUP_NAME_BOOL
                | kind::SUP_NAME_ERROR
                | kind::SUP_NAME_STRING => {
                    let Some((_, _, count)) = cache_dimensions else {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "cached external value occurs outside its matrix".to_string(),
                        ));
                    };
                    if sup_name_state != 4 || cache_values.len() >= count {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "too many or misplaced cached external values".to_string(),
                        ));
                    }
                    cache_values.push(Self::parse_external_cached_value(
                        record.kind(),
                        record.payload(),
                    )?);
                },
                kind::SUP_NAME_VALUE_END => {
                    let Some((rows, columns, count)) = cache_dimensions.take() else {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected BrtSupNameValueEnd".to_string(),
                        ));
                    };
                    if sup_name_state != 4
                        || !record.payload().is_empty()
                        || cache_values.len() != count
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid cached external value matrix".to_string(),
                        ));
                    }
                    current_cache = Some(ValueMatrix::new(
                        rows,
                        columns,
                        std::mem::take(&mut cache_values),
                    )?);
                    sup_name_state = 3;
                },
                kind::SUP_NAME_END => {
                    if sup_name_state != 3 || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtSupNameEnd".to_string(),
                        ));
                    }
                    let kind = link_type.expect("external link kind is present");
                    let name = current_name.take().ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "external name block has no name".to_string(),
                        )
                    })?;
                    let bits = current_bits.take().ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "external name block has no properties".to_string(),
                        )
                    })?;
                    match kind {
                        EXTERNAL_REFERENCE_WORKBOOK => {
                            let scope = binary::read_u32_le_at(&bits, 2)?;
                            let mut entry = DefinedName::new(name)?
                                .with_built_in(bits[0] & EXTERNAL_NAME_BUILT_IN != 0);
                            if scope != 0 {
                                entry = entry.with_sheet_scope(u16::try_from(scope - 1).map_err(
                                    |_| {
                                        crate::package::error::Error::InvalidFormula(
                                            "external defined-name scope overflow".to_string(),
                                        )
                                    },
                                )?);
                            }
                            if let Some(formula) = current_formula.take() {
                                entry = entry.with_formula(formula);
                            }
                            workbook_entries.push(entry);
                        },
                        EXTERNAL_REFERENCE_DDE => {
                            let mut item = DdeItem::new(name)?
                                .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                                .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                                .with_ole_support(bits[0] & DDE_ITEM_SUPPORTS_OLE != 0);
                            if let Some(cache) = current_cache.take() {
                                item = item.with_cached_values(cache);
                            }
                            dde_entries.push(item);
                        },
                        EXTERNAL_REFERENCE_OLE => {
                            let mut item = OleItem::new(name)?
                                .with_advise(bits[0] & DATA_ITEM_WANT_ADVISE != 0)
                                .with_picture(bits[0] & DATA_ITEM_WANT_PICTURE != 0)
                                .with_icon(bits[0] & OLE_ITEM_DISPLAY_AS_ICON != 0);
                            if let Some(cache) = current_cache.take() {
                                item = item.with_cached_values(cache);
                            }
                            ole_entries.push(item);
                        },
                        _ => unreachable!("external link kind was validated above"),
                    }
                    sup_name_state = 0;
                },
                kind::END_SUP_BOOK => {
                    if !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidLength {
                            expected: 0,
                            found: record.payload().len(),
                        });
                    }
                    if sup_name_state != 0 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtEndSupBook occurs inside an external-name block".to_string(),
                        ));
                    }
                    saw_end = true;
                },
                _ => {
                    if sup_name_state == 4
                        || (link_type == Some(EXTERNAL_REFERENCE_WORKBOOK) && sup_name_state != 0)
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unexpected record inside an external name or cache".to_string(),
                        ));
                    }
                },
            }
        }
        let kind = link_type.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "external link has no BrtBeginSupBook".to_string(),
            )
        })?;
        if !saw_end {
            return Err(crate::package::error::Error::InvalidFormula(
                "external link has no BrtEndSupBook".to_string(),
            ));
        }
        if kind == EXTERNAL_REFERENCE_WORKBOOK && !saw_sup_tabs {
            return Err(crate::package::error::Error::InvalidFormula(
                "external workbook link has no BrtSupTabs".to_string(),
            ));
        }
        let (link_kind, source, detail) = match kind {
            EXTERNAL_REFERENCE_DDE => {
                if !part.rels().is_empty() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "DDE external link must not contain relationships".to_string(),
                    ));
                }
                (Kind::Dde, target_key, Some(target_detail))
            },
            EXTERNAL_REFERENCE_WORKBOOK | EXTERNAL_REFERENCE_OLE => {
                if part.rels().len() != 1 {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "external workbook/OLE link must have exactly one data-source relationship"
                            .to_string(),
                    ));
                }
                let relationship = part.rels().get(&target_key).ok_or_else(|| {
                    crate::package::error::Error::InvalidFormula(format!(
                        "external data relationship {target_key:?} is missing"
                    ))
                })?;
                if !relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "external data relationship {target_key:?} is internal"
                    )));
                }
                let allowed_relationship_types = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES
                } else {
                    OLE_DATA_SOURCE_RELATIONSHIP_TYPES
                };
                if !allowed_relationship_types.contains(&relationship.reltype()) {
                    let source_kind = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                        "external workbook"
                    } else {
                        "OLE data source"
                    };
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "{source_kind} relationship {target_key:?} has invalid type {:?}",
                        relationship.reltype()
                    )));
                }
                let target = relationship.target_ref().to_string();
                let link_kind = if kind == EXTERNAL_REFERENCE_WORKBOOK {
                    Kind::Workbook
                } else {
                    Kind::Ole
                };
                let detail = if kind == EXTERNAL_REFERENCE_OLE {
                    Some(target_detail)
                } else {
                    None
                };
                (link_kind, target, detail)
            },
            _ => unreachable!("external link kind was validated above"),
        };
        let entries = match kind {
            EXTERNAL_REFERENCE_WORKBOOK => Entries::Workbook(workbook_entries),
            EXTERNAL_REFERENCE_DDE => Entries::Dde(dde_entries),
            EXTERNAL_REFERENCE_OLE => Entries::Ole(ole_entries),
            _ => unreachable!("external link kind was validated above"),
        };
        let metadata = Link {
            kind: link_kind,
            source,
            detail,
            sheet_names,
            entries,
        };
        metadata.validate()?;
        Ok(ExternalBook { metadata })
    }

    fn validate_external_name_bits(kind: u16, bits: &[u8; 7]) -> Result<()> {
        let reserved_word = &bits[2..6];
        let valid = match kind {
            EXTERNAL_REFERENCE_WORKBOOK => {
                bits[0] & EXTERNAL_NAME_RESERVED_MASK == 0
                    && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG == 0
            },
            EXTERNAL_REFERENCE_DDE => {
                bits[0] & DDE_ITEM_RESERVED_MASK == 0
                    && reserved_word == [0, 0, 0, 0]
                    && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
            },
            EXTERNAL_REFERENCE_OLE => {
                bits[0] & OLE_ITEM_RESERVED_MASK == 0
                    && bits[0] & OLE_ITEM_REQUIRED_CLASS_FLAG != 0
                    && reserved_word == [0, 0, 0, 0]
                    && bits[6] & DATA_ITEM_REQUIRED_TRAILING_FLAG != 0
            },
            _ => false,
        };
        if !valid {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "invalid BrtSupNameBits properties for external-link kind {kind}"
            )));
        }
        Ok(())
    }

    fn parse_external_cached_value(
        record_type: crate::raw::Kind,
        data: &[u8],
    ) -> Result<CachedValue> {
        match record_type {
            kind::SUP_NAME_NIL if data.is_empty() => Ok(CachedValue::Empty),
            kind::SUP_NAME_NUM if data.len() == 8 => {
                let number = f64::from_le_bytes(data.try_into().expect("length was checked"));
                crate::package::external_link::validate_number(number)?;
                Ok(CachedValue::Number(number))
            },
            kind::SUP_NAME_BOOL if data.len() == 1 && data[0] <= 1 => {
                Ok(CachedValue::Boolean(data[0] != 0))
            },
            kind::SUP_NAME_ERROR if data.len() == 1 => {
                Ok(CachedValue::Error(ErrorValue::from_code(data[0])?))
            },
            kind::SUP_NAME_STRING => {
                let (value, consumed) = crate::package::records::decode_string(data)?;
                if consumed != data.len() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "BrtSupNameSt has trailing bytes".to_string(),
                    ));
                }
                Ok(CachedValue::String(value))
            },
            _ => Err(crate::package::error::Error::InvalidFormula(format!(
                "invalid cached external value record {record_type}"
            ))),
        }
    }

    fn parse_external_sheet_names(data: &[u8]) -> Result<Vec<String>> {
        if data.len() < 4 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        let count = usize::try_from(binary::read_u32_le_at(data, 0)?).map_err(|_| {
            crate::package::error::Error::InvalidFormula(
                "external sheet-name count overflow".to_string(),
            )
        })?;
        if count >= 65_535 {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "external sheet-name count {count} exceeds 65,534"
            )));
        }
        let mut names = Vec::with_capacity(count);
        let mut offset = 4;
        for _ in 0..count {
            let (name, consumed) = crate::package::records::decode_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
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
                return Err(crate::package::error::Error::InvalidFormula(format!(
                    "external sheet name {name:?} does not follow sheet-name grammar"
                )));
            }
            if names
                .iter()
                .any(|existing: &String| excel_name_eq(existing, &name))
            {
                return Err(crate::package::error::Error::InvalidFormula(format!(
                    "duplicate external sheet name {name:?}"
                )));
            }
            names.push(name);
        }
        if offset != data.len() {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtSupTabs has {} trailing bytes",
                data.len() - offset
            )));
        }
        Ok(names)
    }

    fn parse_nullable_wide_string(data: &[u8]) -> Result<(Option<String>, usize)> {
        if data.len() < 4 {
            return Err(crate::package::error::Error::InvalidLength {
                expected: 4,
                found: data.len(),
            });
        }
        if binary::read_u32_le_at(data, 0)? == u32::MAX {
            Ok((None, 4))
        } else {
            let (value, consumed) = crate::package::records::decode_string(data)?;
            Ok((Some(value), consumed))
        }
    }

    pub(super) fn parse_pivot_cache_ids(data: &[u8]) -> Result<Vec<(u32, String)>> {
        let mut in_collection = false;
        let mut open_cache = false;
        let mut ended = false;
        let mut caches = Vec::new();
        for record in Records::new(data) {
            let record = record?;
            match record.kind() {
                kind::BEGIN_PIVOT_CACHE_IDS => {
                    if in_collection || ended {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "duplicate BrtBeginPivotCacheIDs collection".to_string(),
                        ));
                    }
                    in_collection = true;
                },
                kind::BEGIN_PIVOT_CACHE_ID => {
                    if !in_collection || open_cache || record.payload().len() < 8 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "malformed BrtBeginPivotCacheID nesting or payload".to_string(),
                        ));
                    }
                    let cache_id = binary::read_u32_le_at(record.payload(), 0)?;
                    let (rel_id, consumed) =
                        crate::package::records::decode_string(&record.payload()[4..])?;
                    if 4 + consumed != record.payload().len()
                        || rel_id.is_empty()
                        || rel_id.encode_utf16().count() > 255
                        || caches
                            .iter()
                            .any(|(existing, _): &(u32, String)| *existing == cache_id)
                    {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
                            "invalid or duplicate PivotCache ID {cache_id}"
                        )));
                    }
                    caches.push((cache_id, rel_id));
                    open_cache = true;
                },
                kind::END_PIVOT_CACHE_ID => {
                    if !open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unbalanced BrtEndPivotCacheID".to_string(),
                        ));
                    }
                    open_cache = false;
                },
                kind::END_PIVOT_CACHE_IDS => {
                    if !in_collection || open_cache || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
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
            return Err(crate::package::error::Error::InvalidFormula(
                "unterminated PivotCache ID collection".to_string(),
            ));
        }
        Ok(caches)
    }

    pub(super) fn parse_pivot_view(data: &[u8], sheet_index: usize) -> Result<View> {
        let mut view = None;
        for record in Records::new(data) {
            let record = record?;
            if record.kind() != kind::BEGIN_SX_VIEW {
                continue;
            }
            if view.is_some() || record.payload().len() < 36 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable part has duplicate or truncated BrtBeginSXView".to_string(),
                ));
            }
            let cache_id = binary::read_u32_le_at(record.payload(), 28)?;
            let (name, consumed) = crate::package::records::decode_string(&record.payload()[32..])?;
            if consumed > record.payload().len() - 32 {
                return Err(crate::package::error::Error::InvalidFormula(
                    "PivotTable view name overruns BrtBeginSXView".to_string(),
                ));
            }
            view = Some(View::try_new(cache_id, sheet_index, name)?);
        }
        view.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "PivotTable part omits BrtBeginSXView".to_string(),
            )
        })
    }

    pub(super) fn parse_table_definition(
        data: &[u8],
        sheet_index: usize,
    ) -> Result<TableDefinition> {
        let mut table_header: Option<(u32, String, usize)> = None;
        let mut expected_columns = None;
        let mut columns = Vec::new();
        let mut in_column = false;
        let mut ended_columns = false;
        let mut ended_table = false;
        let mut iter = Records::new(data);
        for record in iter.by_ref() {
            let record = record?;
            if ended_table {
                return Err(crate::package::error::Error::InvalidFormula(
                    "XLSB table part contains records after BrtEndList".to_string(),
                ));
            }
            match record.kind() {
                kind::BEGIN_LIST => {
                    if table_header.is_some() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "XLSB table part contains duplicate BrtBeginList".to_string(),
                        ));
                    }
                    table_header = Some(Self::parse_table_header(record.payload())?);
                },
                kind::BEGIN_LIST_COLS => {
                    let (_, _, range_columns) = table_header.as_ref().ok_or_else(|| {
                        crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCols precedes BrtBeginList".to_string(),
                        )
                    })?;
                    if expected_columns.is_some() || record.payload().len() != 4 {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid or duplicate BrtBeginListCols".to_string(),
                        ));
                    }
                    let count = usize::try_from(binary::read_u32_le_at(record.payload(), 0)?)
                        .map_err(|_| {
                            crate::package::error::Error::InvalidFormula(
                                "table column count overflow".to_string(),
                            )
                        })?;
                    if count == 0 || count > 16_384 || count != *range_columns {
                        return Err(crate::package::error::Error::InvalidFormula(format!(
                            "table column count {count} disagrees with range width {range_columns}"
                        )));
                    }
                    expected_columns = Some(count);
                },
                kind::BEGIN_LIST_COL => {
                    if expected_columns.is_none() || ended_columns || in_column {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "BrtBeginListCol occurs outside its column collection".to_string(),
                        ));
                    }
                    columns.push(Self::parse_table_column(record.payload(), columns.len())?);
                    in_column = true;
                },
                kind::END_LIST_COL => {
                    if !in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "unmatched or nonempty BrtEndListCol".to_string(),
                        ));
                    }
                    in_column = false;
                },
                kind::END_LIST_COLS => {
                    if expected_columns.is_none()
                        || in_column
                        || ended_columns
                        || !record.payload().is_empty()
                    {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtEndListCols".to_string(),
                        ));
                    }
                    ended_columns = true;
                },
                kind::END_LIST => {
                    if !ended_columns || in_column || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "invalid BrtEndList".to_string(),
                        ));
                    }
                    ended_table = true;
                },
                _ => {},
            }
        }
        let (table_id, display_name, _) = table_header.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "XLSB table part omits BrtBeginList".to_string(),
            )
        })?;
        let expected = expected_columns.ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "XLSB table part omits BrtBeginListCols".to_string(),
            )
        })?;
        if !ended_table || columns.len() != expected {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "XLSB table contains {} of {expected} declared columns or is unterminated",
                columns.len()
            )));
        }
        TableDefinition::try_new(table_id, sheet_index, display_name, columns)
    }

    fn parse_table_header(data: &[u8]) -> Result<(u32, String, usize)> {
        if data.len() < 64 {
            return Err(crate::package::error::Error::InvalidLength {
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
            return Err(crate::package::error::Error::InvalidFormula(
                "BrtBeginList contains an invalid table range".to_string(),
            ));
        }
        for offset in [24, 28] {
            if binary::read_u32_le_at(data, offset)? > 1 {
                return Err(crate::package::error::Error::InvalidFormula(
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
                crate::package::error::Error::InvalidFormula(
                    "BrtBeginList string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginList has {} trailing bytes",
                data.len() - offset
            )));
        }
        let display_name = strings[1].clone().ok_or_else(|| {
            crate::package::error::Error::InvalidFormula(
                "BrtBeginList has a NULL display name".to_string(),
            )
        })?;
        Ok((
            table_id,
            display_name,
            usize::try_from(col_last - col_first + 1).expect("bounded table width"),
        ))
    }

    fn parse_table_column(data: &[u8], index: usize) -> Result<String> {
        if data.len() < 24 || binary::read_u32_le_at(data, 0)? == 0 {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginListCol {index} has an invalid header"
            )));
        }
        let mut offset = 24;
        let mut strings = Vec::with_capacity(6);
        for _ in 0..6 {
            let (value, consumed) = Self::parse_nullable_wide_string(&data[offset..])?;
            offset = offset.checked_add(consumed).ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(
                    "BrtBeginListCol string size overflow".to_string(),
                )
            })?;
            strings.push(value);
        }
        if offset != data.len() {
            return Err(crate::package::error::Error::InvalidFormula(format!(
                "BrtBeginListCol has {} trailing bytes",
                data.len() - offset
            )));
        }
        strings[0]
            .clone()
            .or_else(|| strings[1].clone())
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormula(format!(
                    "BrtBeginListCol {index} has neither a name nor caption"
                ))
            })
    }
}
