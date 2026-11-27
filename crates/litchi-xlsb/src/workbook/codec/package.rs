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
pub(crate) struct MergeBlockLayout {
    pub(crate) ranges: Vec<MergedCell>,
    pub(crate) block_span: Option<(usize, usize)>,
    pub(crate) insertion_offset: usize,
}

impl Workbook {
    /// Return inert slicer cache snapshots in workbook relationship order.
    pub fn slicer_caches(&self) -> Result<Vec<crate::slicer::Cache>> {
        Ok(crate::slicer::package::load_caches(
            &self.package,
            &litchi_opc::PackURI::new("/xl/workbook.bin")?,
            self.raw_limits,
        )?
        .into_iter()
        .map(|part| part.cache)
        .collect())
    }

    /// Return inert slicer views attached to one worksheet.
    pub fn slicers(&self, worksheet_index: usize) -> Result<Option<crate::slicer::Views>> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        Ok(
            crate::slicer::package::load_views(&self.package, &worksheet, self.raw_limits)?
                .map(|part| part.views),
        )
    }

    /// Return inert timeline cache snapshots in workbook relationship order.
    pub fn timeline_caches(&self) -> Result<Vec<crate::timeline::Cache>> {
        Ok(crate::timeline::package::load_caches(
            &self.package,
            &litchi_opc::PackURI::new("/xl/workbook.bin")?,
            self.semantic_limits.xml_parts_value(),
        )?
        .into_iter()
        .map(|part| part.cache)
        .collect())
    }

    /// Return inert timeline views attached to one worksheet.
    pub fn timelines(&self, worksheet_index: usize) -> Result<Option<crate::timeline::Views>> {
        let worksheet = self.worksheet_uri(worksheet_index)?;
        Ok(crate::timeline::package::load_views(
            &self.package,
            &worksheet,
            self.semantic_limits.xml_parts_value(),
        )?
        .map(|part| part.views))
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
        crate::package::scenarios::parse_worksheet(part.blob(), self.raw_limits)
    }

    /// Return the typed Scenario Manager for a worksheet selected by name.
    pub fn scenarios_by_name(
        &self,
        worksheet_name: &str,
    ) -> Result<Option<crate::package::scenarios::Manager>> {
        let index = self.worksheet_index(worksheet_name)?;
        self.scenarios(index)
    }

    /// List the merged ranges of the worksheet at worksheet ordinal
    /// `worksheet_index`.
    pub fn merged_cell_ranges(&self, worksheet_index: usize) -> Result<Vec<MergedCell>> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let part = self.package.get_part(&uri)?;
        Ok(Self::inspect_merge_block(part.blob(), self.raw_limits)?.ranges)
    }

    /// List merged ranges in a worksheet selected by exact name.
    pub fn merged_cell_ranges_by_name(&self, worksheet_name: &str) -> Result<Vec<MergedCell>> {
        self.merged_cell_ranges(self.worksheet_index(worksheet_name)?)
    }

    /// Replace all merged ranges of a worksheet in place (the structural
    /// replay seam of [`crate::cell_values::structure`]).
    pub(crate) fn set_merged_cell_ranges(
        &mut self,
        worksheet_index: usize,
        ranges: &[MergedCell],
    ) -> Result<()> {
        let uri = self.worksheet_uri(worksheet_index)?;
        let original = self.package.get_part(&uri)?.blob().to_vec();
        let layout = Self::inspect_merge_block(&original, self.raw_limits)?;
        let normalized = Self::normalize_merge_ranges(ranges)?;
        let replacement = Self::serialize_merge_block(&normalized, self.raw_limits)?;
        let (start, end) = layout
            .block_span
            .unwrap_or((layout.insertion_offset, layout.insertion_offset));
        let mut updated = Vec::with_capacity(original.len() - (end - start) + replacement.len());
        updated.extend_from_slice(&original[..start]);
        updated.extend_from_slice(&replacement);
        updated.extend_from_slice(&original[end..]);
        Self::inspect_merge_block(&updated, self.raw_limits)?;
        if updated == original {
            return Ok(());
        }
        self.edit_opc(|package| {
            crate::worksheet_index::maintain(package, &uri, &original, &updated)?;
            package.get_part_mut(&uri)?.set_blob(updated);
            Ok(())
        })
    }

    /// Open an XLSB workbook from a reader
    pub(in crate::workbook) fn worksheet_index(&self, worksheet_name: &str) -> Result<usize> {
        self.worksheet_names
            .iter()
            .position(|name| name == worksheet_name)
            .ok_or(crate::package::error::Error::WorksheetNotFound(None))
    }

    pub(crate) fn worksheet_uri(&self, index: usize) -> Result<litchi_opc::PackURI> {
        let catalog_position = self.catalog_position_for_worksheet(index)?;
        let _name = self
            .formula_context
            .worksheet_names
            .get(catalog_position)
            .ok_or({
                crate::package::error::Error::InvalidFormat(
                    "Worksheet catalog position out of bounds",
                )
            })?;
        let rel_id = self
            .worksheet_rel_ids
            .get(catalog_position)
            .and_then(Option::as_deref)
            .ok_or({
                crate::package::error::Error::UnsupportedFeature(
                    "sheet has no worksheet relationship",
                )
            })?;
        let workbook_part = self.package.main_document_part()?;
        let relationship =
            workbook_part
                .rels()
                .get(rel_id)
                .ok_or(crate::package::error::Error::FileNotFound(
                    "relationship for sheet",
                ))?;
        if relationship.is_external() {
            return Err(crate::package::error::Error::UnsupportedFeature(
                "sheet has an external worksheet relationship",
            ));
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

    pub(crate) fn normalize_merge_ranges(ranges: &[MergedCell]) -> Result<Vec<MergedCell>> {
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
                    typ: "BrtMergeCell collection",
                    val: "duplicate or noncanonical range order",
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
            if let Some((&_col_first, &(col_last, _))) = active.range(..=range.col_last).next_back()
                && col_last >= range.col_first
            {
                return Err(crate::package::error::Error::InvalidCellReference(
                    "merged range overlaps an existing range beginning in column",
                ));
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

    pub(crate) fn inspect_merge_block(
        data: &[u8],
        raw_limits: crate::raw::Limits,
    ) -> Result<MergeBlockLayout> {
        let mut records = Records::new(data, raw_limits);
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
                            typ: "BrtBeginMergeCells",
                            val: "duplicate, nested, or out-of-order record",
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
                            typ: "BrtMergeCell",
                            val: "record occurs outside BrtBeginMergeCells",
                        });
                    }
                    ranges.push(MergedCell::parse(record.payload(), raw_limits)?);
                    if ranges.len() > declared_count.unwrap_or_default() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginMergeCells",
                            val: "declared count is smaller than the record collection",
                        });
                    }
                },
                kind::END_MERGE_CELLS => {
                    if !in_block || !record.payload().is_empty() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtEndMergeCells",
                            val: "orphan, duplicate, or nonempty record",
                        });
                    }
                    if declared_count != Some(ranges.len()) {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtBeginMergeCells",
                            val: "declared count disagrees with BrtMergeCell records",
                        });
                    }
                    block_span = Some((begin_offset.expect("merge begin offset"), end));
                    in_block = false;
                },
                kind::END_SHEET_DATA => {
                    if in_block {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtMergeCells collection",
                            val: "noncontiguous record collection",
                        });
                    }
                    saw_end_sheet_data = true;
                },
                kind::END_SHEET => {
                    if in_block || end_sheet_offset.replace(start).is_some() {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtEndSheet",
                            val: "duplicate or embedded in merge collection",
                        });
                    }
                },
                record_type => {
                    if in_block {
                        return Err(crate::package::error::Error::Unrecognized {
                            typ: "BrtMergeCells collection",
                            val: "unexpected record",
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
                "BrtMergeCells collection",
            ));
        }
        let end_sheet_offset = end_sheet_offset.ok_or(
            crate::package::error::Error::UnexpectedEndOfStream("BrtEndSheet"),
        )?;
        if !saw_end_sheet_data {
            return Err(crate::package::error::Error::UnexpectedEndOfStream(
                "BrtEndSheetData",
            ));
        }
        if block_span.is_some() {
            Self::validate_merge_range_collection(&ranges, true)?;
            if first_post_merge_offset
                .is_some_and(|offset| block_span.is_some_and(|(begin, _)| offset < begin))
            {
                return Err(crate::package::error::Error::Unrecognized {
                    typ: "BrtMergeCells collection",
                    val: "collection occurs after a later worksheet feature",
                });
            }
        }
        Ok(MergeBlockLayout {
            ranges,
            block_span,
            insertion_offset: first_post_merge_offset.unwrap_or(end_sheet_offset),
        })
    }

    pub(crate) fn serialize_merge_block(
        ranges: &[MergedCell],
        raw_limits: crate::raw::Limits,
    ) -> Result<Vec<u8>> {
        if ranges.is_empty() {
            return Ok(Vec::new());
        }
        let mut output = Vec::with_capacity(10 + ranges.len() * 19);
        let mut writer = crate::raw::Encoder::new(&mut output, raw_limits);
        writer.write_record(
            kind::BEGIN_MERGE_CELLS,
            &(ranges.len() as u32).to_le_bytes(),
        )?;
        for range in ranges {
            writer.write_record(kind::MERGE_CELL, &range.serialize(raw_limits)?)?;
        }
        writer.write_record(kind::END_MERGE_CELLS, &[])?;
        Ok(output)
    }

    pub(in crate::workbook) fn load_external_book_with_budget(
        &self,
        uri: &litchi_opc::PackURI,
        budget: &mut crate::external_link::Budget,
    ) -> Result<ExternalBook> {
        let part = self.package.get_part(uri)?;
        if part.content_type() != "application/vnd.ms-excel.externalLink" {
            return Err(crate::package::error::Error::InvalidFormula(
                "external link part has invalid content type",
            ));
        }

        let parsed = crate::external_link::parse_external_link_with_budget(part.blob(), budget)?;
        Self::resolve_external_book(part.rels(), parsed, budget)
    }

    pub(in crate::workbook) fn resolve_external_book(
        relationships: &litchi_opc::Relationships,
        parsed: crate::external_link::Parsed,
        budget: &mut crate::external_link::Budget,
    ) -> Result<ExternalBook> {
        match parsed.link().kind() {
            Kind::Dde => {
                if !relationships.is_empty() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "DDE external link must not contain relationships",
                    ));
                }
                Ok(ExternalBook {
                    metadata: parsed.into_link(),
                })
            },
            Kind::Workbook | Kind::Ole => {
                let relationship_id = parsed.relationship_id().ok_or({
                    crate::package::error::Error::InvalidFormula(
                        "external workbook/OLE link has no data-source relationship ID",
                    )
                })?;
                let relationship = relationships.get(relationship_id).ok_or({
                    crate::package::error::Error::InvalidFormula(
                        "external data relationship is missing",
                    )
                })?;
                if !relationship.is_external() {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "external data relationship is internal",
                    ));
                }

                let (allowed_relationship_types, _source_kind) = match parsed.link().kind() {
                    Kind::Workbook => (EXTERNAL_WORKBOOK_RELATIONSHIP_TYPES, "external workbook"),
                    Kind::Ole => (OLE_DATA_SOURCE_RELATIONSHIP_TYPES, "OLE data source"),
                    Kind::Dde => {
                        return Err(crate::package::error::Error::InvalidFormula(
                            "DDE link reached relationship resolution",
                        ));
                    },
                };
                if !allowed_relationship_types.contains(&relationship.reltype()) {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "relationship has invalid type",
                    ));
                }
                // Native Excel can store both relative and absolute data
                // source paths. The BrtBeginSupBook ID selects the target;
                // retain alternate external paths without opening any of them.
                if relationships.iter().any(|rel| {
                    !rel.is_external() || !allowed_relationship_types.contains(&rel.reltype())
                }) {
                    return Err(crate::package::error::Error::InvalidFormula(
                        "external link has an invalid alternate data-source relationship",
                    ));
                }

                let target_ref = relationship.target_ref();
                let target = copy_target_ref(target_ref, budget)?;
                Ok(ExternalBook {
                    metadata: parsed.resolve_source(target)?,
                })
            },
        }
    }
}

fn copy_target_ref(target_ref: &str, budget: &mut crate::external_link::Budget) -> Result<String> {
    budget.string(target_ref.encode_utf16().count(), target_ref.len())?;
    let mut copied = String::new();
    copied
        .try_reserve_exact(target_ref.len())
        .map_err(|source| crate::external_link::Error::Allocation {
            resource: "external-link relationship target",
            source,
        })?;
    copied.push_str(target_ref);
    Ok(copied)
}
