#![allow(
    clippy::cast_possible_truncation,
    clippy::expect_used,
    reason = "legacy module confines validated BIFF12 field narrowing or exact signed-bit reinterpretation, extraction after an immediately preceding structural invariant check to this codec boundary"
)]

//! Package-facing workbook resolution and merged-cell mutation.

use super::super::model::Workbook;
use crate::external_link::Kind;
use crate::merged_cells::{MAX_MERGED_CELL_RANGES, MergedCell};
use crate::package::error::Result;
use crate::package::formula::ExternalBook;
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
    /// Return inert slicer cache snapshots in workbook relationship order.
    pub fn slicer_caches(&self) -> Result<Vec<crate::slicer::Cache>> {
        Ok(crate::slicer::package::load_caches(
            &self.package,
            &litchi_opc::PackURI::new("/xl/workbook.bin")?,
        )?
        .into_iter()
        .map(|part| part.cache)
        .collect())
    }

    /// Atomically replace workbook slicer caches.
    pub fn set_slicer_caches(&mut self, caches: Vec<crate::slicer::Cache>) -> Result<()> {
        self.edit_opc(|package| {
            let workbook = litchi_opc::PackURI::new("/xl/workbook.bin")?;
            crate::slicer::package::store_caches(package, &workbook, &caches)
        })
    }

    /// Remove all workbook slicer caches. Returns whether any were present.
    pub fn remove_slicer_caches(&mut self) -> Result<bool> {
        let had = !self.slicer_caches()?.is_empty();
        if had {
            self.set_slicer_caches(Vec::new())?;
        }
        Ok(had)
    }

    /// Return inert slicer views attached to one worksheet.
    pub fn slicers(&self, worksheet_index: usize) -> Result<Option<crate::slicer::Views>> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        Ok(crate::slicer::package::load_views(&self.package, &worksheet)?.map(|part| part.views))
    }

    /// Atomically replace one worksheet's slicer views.
    pub fn set_slicers(
        &mut self,
        worksheet_index: usize,
        views: crate::slicer::Views,
    ) -> Result<()> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        self.edit_opc(|package| crate::slicer::package::store_views(package, &worksheet, &views))
    }

    /// Remove one worksheet's slicer views. Returns whether a part was present.
    pub fn remove_slicers(&mut self, worksheet_index: usize) -> Result<bool> {
        let had = self.slicers(worksheet_index)?.is_some();
        if had {
            self.set_slicers(worksheet_index, crate::slicer::Views::new())?;
        }
        Ok(had)
    }

    /// Return inert timeline cache snapshots in workbook relationship order.
    pub fn timeline_caches(&self) -> Result<Vec<crate::timeline::Cache>> {
        Ok(crate::timeline::package::load_caches(
            &self.package,
            &litchi_opc::PackURI::new("/xl/workbook.bin")?,
        )?
        .into_iter()
        .map(|part| part.cache)
        .collect())
    }

    /// Atomically replace workbook timeline caches.
    pub fn set_timeline_caches(&mut self, caches: Vec<crate::timeline::Cache>) -> Result<()> {
        self.edit_opc(|package| {
            let workbook = litchi_opc::PackURI::new("/xl/workbook.bin")?;
            crate::timeline::package::store_caches(package, &workbook, &caches)
        })
    }

    /// Remove all workbook timeline caches. Returns whether any were present.
    pub fn remove_timeline_caches(&mut self) -> Result<bool> {
        let had = !self.timeline_caches()?.is_empty();
        if had {
            self.set_timeline_caches(Vec::new())?;
        }
        Ok(had)
    }

    /// Return inert timeline views attached to one worksheet.
    pub fn timelines(&self, worksheet_index: usize) -> Result<Option<crate::timeline::Views>> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        Ok(crate::timeline::package::load_views(&self.package, &worksheet)?.map(|part| part.views))
    }

    /// Atomically replace one worksheet's timeline views.
    pub fn set_timelines(
        &mut self,
        worksheet_index: usize,
        views: crate::timeline::Views,
    ) -> Result<()> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        self.edit_opc(|package| crate::timeline::package::store_views(package, &worksheet, &views))
    }

    /// Remove one worksheet's timeline views. Returns whether a part was present.
    pub fn remove_timelines(&mut self, worksheet_index: usize) -> Result<bool> {
        let had = self.timelines(worksheet_index)?.is_some();
        if had {
            self.set_timelines(worksheet_index, crate::timeline::Views::new())?;
        }
        Ok(had)
    }

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
    pub(in crate::workbook) fn worksheet_index(&self, worksheet_name: &str) -> Result<usize> {
        self.worksheet_names
            .iter()
            .position(|name| name == worksheet_name)
            .ok_or_else(|| {
                crate::package::error::Error::WorksheetNotFound(worksheet_name.to_string())
            })
    }

    pub(crate) fn worksheet_uri(&self, index: usize) -> Result<litchi_opc::PackURI> {
        let catalog_position = self.catalog_position_for_worksheet(index)?;
        let name = self
            .formula_context
            .worksheet_names
            .get(catalog_position)
            .ok_or_else(|| {
                crate::package::error::Error::InvalidFormat(format!(
                    "Worksheet catalog position {catalog_position} out of bounds"
                ))
            })?;
        let rel_id = self
            .worksheet_rel_ids
            .get(catalog_position)
            .and_then(Option::as_deref)
            .ok_or_else(|| {
                crate::package::error::Error::UnsupportedFeature(format!(
                    "sheet {name:?} has no worksheet relationship"
                ))
            })?;
        let workbook_part = self.package.main_document_part()?;
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

        let parsed = crate::external_link::parse_external_link(part.blob())?;
        match parsed.link().kind() {
            Kind::Dde => {
                if !part.rels().is_empty() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "DDE external link must not contain relationships".to_string(),
                    ));
                }
                Ok(ExternalBook {
                    metadata: parsed.into_link(),
                })
            },
            Kind::Workbook | Kind::Ole => {
                if part.rels().len() != 1 {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "external workbook/OLE link must have exactly one data-source relationship"
                            .to_string(),
                    ));
                }

                let relationship_id = parsed.relationship_id().ok_or_else(|| {
                    crate::package::error::Error::InvalidFormula(
                        "external workbook/OLE link has no data-source relationship ID".to_string(),
                    )
                })?;
                let relationship = part.rels().get(relationship_id).ok_or_else(|| {
                    crate::package::error::Error::InvalidFormula(format!(
                        "external data relationship {relationship_id:?} is missing"
                    ))
                })?;
                if !relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "external data relationship {relationship_id:?} is internal"
                    )));
                }

                let (allowed_relationship_types, source_kind) = match parsed.link().kind() {
                    Kind::Workbook => (EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES, "external workbook"),
                    Kind::Ole => (OLE_DATA_SOURCE_RELATIONSHIP_TYPES, "OLE data source"),
                    Kind::Dde => {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "DDE link reached relationship resolution".to_string(),
                        ));
                    },
                };
                if !allowed_relationship_types.contains(&relationship.reltype()) {
                    return Err(crate::package::error::Error::InvalidFormula(format!(
                        "{source_kind} relationship {relationship_id:?} has invalid type {:?}",
                        relationship.reltype()
                    )));
                }

                Ok(ExternalBook {
                    metadata: parsed.resolve_source(relationship.target_ref().to_string())?,
                })
            },
        }
    }
}
