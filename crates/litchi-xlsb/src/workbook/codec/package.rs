//! Package-facing workbook resolution and merged-cell mutation.

use super::super::model::Workbook;
use crate::named_ranges::validate_name;
use crate::package::error::Result;
use crate::package::external_link::{
    DATA_ITEM_WANT_ADVISE, DATA_ITEM_WANT_PICTURE, DDE_ITEM_SUPPORTS_OLE, DdeItem, DefinedName,
    EXTERNAL_NAME_BUILT_IN, EXTERNAL_REFERENCE_DDE, EXTERNAL_REFERENCE_OLE,
    EXTERNAL_REFERENCE_WORKBOOK, Entries, Kind, Link, MAX_XLSB_EXTERNAL_CACHED_VALUES, NameFormula,
    OLE_ITEM_DISPLAY_AS_ICON, OleItem, ValueMatrix,
};
use crate::package::formula::ExternalBook;
use crate::package::merged_cells::{MAX_MERGED_CELL_RANGES, MergedCell};
use crate::raw::{Records, kind};
use litchi_core::binary;
use litchi_ooxml_common::external_link::EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES;
use litchi_opc::constants::relationship_type;
use std::cmp::Reverse;
use std::collections::{BTreeMap, BinaryHeap};

/// OLE data-source relationship types documented by MS-XLSB and MS-OI29500.
const OLE_DATA_SOURCE_RELATIONSHIP_TYPES: &[&str] = &[
    relationship_type::OLE_OBJECT,
    relationship_type::STRICT_OLE_OBJECT,
    "http://schemas.microsoft.com/office/2019/04/relationships/oleObjectLinkLongPath",
];

#[derive(Debug)]
struct MergeBlockLayout {
    ranges: Vec<MergedCell>,
    block_span: Option<(usize, usize)>,
    insertion_offset: usize,
}

impl Workbook {
    /// Return the typed Scenario Manager stored in one worksheet, if any.
    ///
    /// Scenario values are inert snapshots. The reader does not apply them to
    /// worksheet cells and never recalculates formulas.
    pub fn scenarios(
        &self,
        worksheet_index: usize,
    ) -> Result<Option<crate::package::scenarios::Manager>> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let part = self.package.get_part(&uri)?;
        crate::package::scenarios::parse_worksheet(part.blob())
    }

    /// Return the typed Scenario Manager for a worksheet selected by name.
    pub fn scenarios_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::package::scenarios::Manager>> {
        let index = self.worksheet_index(worksheet_name)?;
        self.scenarios(index)
    }

    /// Atomically replace or add a worksheet's Scenario Manager.
    ///
    /// The scenario owner validates the complete candidate worksheet before
    /// this method publishes it. Existing record order and opaque records are
    /// retained by the owner; malformed or ambiguous structure is refused.
    pub fn set_scenarios(
        &mut self,
        worksheet_index: usize,
        scenarios: crate::package::scenarios::Manager,
    ) -> Result<()> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let original = self.package.get_part(&uri)?.blob().to_vec();
        let updated = crate::package::scenarios::replace_worksheet(&original, Some(&scenarios))?;
        crate::package::scenarios::parse_worksheet(&updated)?;
        self.package.unsign();
        self.package.get_part_mut(&uri)?.set_blob(updated);
        Ok(())
    }

    /// Atomically remove a worksheet's Scenario Manager.
    ///
    /// Returns `true` only when a scenario collection was removed. Unknown
    /// worksheet records remain byte-for-byte in their original order.
    pub fn remove_scenarios(&mut self, worksheet_index: usize) -> Result<bool> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let original = self.package.get_part(&uri)?.blob().to_vec();
        let updated = crate::package::scenarios::replace_worksheet(&original, None)?;
        if updated == original {
            return Ok(false);
        }
        crate::package::scenarios::parse_worksheet(&updated)?;
        self.package.unsign();
        self.package.get_part_mut(&uri)?.set_blob(updated);
        Ok(true)
    }

    /// Atomically replace or add a worksheet's Scenario Manager by name.
    pub fn set_scenarios_by_name(
        &mut self,
        worksheet_name: &str,
        scenarios: crate::package::scenarios::Manager,
    ) -> Result<()> {
        let index = self.worksheet_index(worksheet_name)?;
        self.set_scenarios(index, scenarios)
    }

    /// Remove a worksheet's Scenario Manager by name.
    pub fn remove_scenarios_by_name(&mut self, worksheet_name: &str) -> Result<bool> {
        let index = self.worksheet_index(worksheet_name)?;
        self.remove_scenarios(index)
    }

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

    pub(in crate::workbook) fn worksheet_uri(&self, index: usize) -> Result<litchi_opc::PackURI> {
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

    pub(in crate::workbook) fn load_external_book(
        &self,
        uri: &litchi_opc::PackURI,
    ) -> Result<ExternalBook> {
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
}
