//! Semantic collection of worksheet BIFF records.

use super::super::validation::cell_record_xf;
use super::super::wire::{SharedFormulaTemplate, parse_shared_formula_template};
use crate::cell::Cell;
use crate::error::{Error, Result};
use crate::formula::{FormulaContext, ptg_exp_anchor, render_formula};
use crate::formula_metadata::Cell as FormulaCell;
use crate::formula_metadata::array::{self, Owner as ArrayFormula};
use crate::number_format::Formatting;
use crate::records::{
    BoundSheetRecord, CellRecord, DimensionsRecord, Encoding, FormulaValue, SharedStringProperties,
};
use crate::workbook::model::Workbook;
use crate::worksheet::Worksheet;
use crate::{
    autofilter, comments, conditional_format, hyperlinks, layout, merged_cells, page_setup,
    pivot_table, protection, utils, view,
};
use litchi_biff::Records;
use litchi_core::sheet::Cell as _;
use std::collections::{HashMap, HashSet};
use std::io::{Read, Seek};
use std::sync::Arc;

#[derive(Clone, Copy)]
struct FormulaLink {
    row: u16,
    col: u16,
    anchor: Option<(u16, u16)>,
    shared: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CompanionKind {
    Array,
    Table,
    Shared,
    Sub,
}

impl FormulaLink {
    fn from_record(record: &CellRecord) -> Option<Self> {
        let CellRecord::Formula {
            row,
            col,
            formula,
            metadata,
            ..
        } = record
        else {
            return None;
        };
        Some(Self {
            row: *row,
            col: *col,
            anchor: ptg_exp_anchor(formula),
            shared: metadata.shared_formula(),
        })
    }
}

fn add_cell(
    worksheet: &mut Worksheet,
    cell: Cell,
    duplicate_cells: &mut HashSet<(u16, u16)>,
    maximum: usize,
) -> Result<()> {
    let position = (cell.row() as u16, cell.column() as u16);
    if worksheet
        .get_cell(u32::from(position.0), u32::from(position.1))
        .is_some()
    {
        if !duplicate_cells.contains(&position) && duplicate_cells.len() >= maximum {
            return Err(Error::InvalidRecord {
                record_type: 0x0006,
                message: "duplicate cell count exceeds the configured limit".to_string(),
            });
        }
        duplicate_cells
            .try_reserve(1)
            .map_err(|_| Error::Allocation("tracking duplicate worksheet cells"))?;
        duplicate_cells.insert(position);
    }
    worksheet.add_cell(cell);
    Ok(())
}

fn claim_companion(
    predecessor: Option<FormulaLink>,
    kind: CompanionKind,
    record_type: u16,
    claims: &mut HashMap<(u16, u16), CompanionKind>,
    maximum: usize,
) -> Result<FormulaLink> {
    let predecessor = predecessor.ok_or_else(|| Error::InvalidRecord {
        record_type,
        message: format!("{kind:?} must immediately follow its Formula record"),
    })?;
    let anchor = (predecessor.row, predecessor.col);
    if let Some(existing) = claims.get(&anchor) {
        return Err(Error::InvalidRecord {
            record_type,
            message: format!(
                "Formula at ({}, {}) already owns a {existing:?} companion",
                anchor.0, anchor.1
            ),
        });
    }
    if claims.len() >= maximum {
        return Err(Error::InvalidRecord {
            record_type,
            message: "Formula companion count exceeds the configured limit".to_string(),
        });
    }
    claims
        .try_reserve(1)
        .map_err(|_| Error::Allocation("tracking Formula companion ownership"))?;
    claims.insert(anchor, kind);
    Ok(predecessor)
}

fn column_mask(first: u8, last: u8, word: usize) -> u64 {
    let word_first = word * 64;
    let word_last = word_first + 63;
    let first = usize::from(first).max(word_first);
    let last = usize::from(last).min(word_last);
    if first > last {
        return 0;
    }
    let width = last - first + 1;
    let bits = if width == 64 {
        u64::MAX
    } else {
        (1u64 << width) - 1
    };
    bits << (first - word_first)
}

fn claim_array_range(occupancy: &mut HashMap<u16, [u64; 4]>, owner: &ArrayFormula) -> Result<()> {
    let range = owner.range();
    let first_row = range.first().row();
    let last_row = range.last().row();
    let masks: [u64; 4] =
        std::array::from_fn(|word| column_mask(range.first().col(), range.last().col(), word));
    let mut missing_rows = 0usize;
    for row in first_row..=last_row {
        if let Some(columns) = occupancy.get(&row) {
            if columns
                .iter()
                .zip(masks)
                .any(|(occupied, mask)| occupied & mask != 0)
            {
                return Err(Error::InvalidRecord {
                    record_type: 0x0221,
                    message: "Array owner overlaps an existing Array range".to_string(),
                });
            }
        } else {
            missing_rows = missing_rows
                .checked_add(1)
                .ok_or_else(|| Error::InvalidRecord {
                    record_type: 0x0221,
                    message: "Array occupancy row count overflows".to_string(),
                })?;
        }
    }
    occupancy
        .try_reserve(missing_rows)
        .map_err(|_| Error::Allocation("indexing Array range occupancy"))?;
    for row in first_row..=last_row {
        let columns = occupancy.entry(row).or_insert([0; 4]);
        for (occupied, mask) in columns.iter_mut().zip(masks) {
            *occupied |= mask;
        }
    }
    Ok(())
}

impl<R: Read + Seek> Workbook<R> {
    /// Parse a worksheet from its position in the workbook stream
    pub(crate) fn parse_worksheet_from_position(
        &self,
        workbook_data: &[u8],
        bound_sheet: &BoundSheetRecord,
        encoding: &Encoding,
    ) -> Result<Worksheet> {
        let worksheet_position = usize::try_from(bound_sheet.position).map_err(|_| {
            Error::InvalidData("worksheet position does not fit in usize".to_string())
        })?;
        let worksheet_data = workbook_data.get(worksheet_position..).ok_or_else(|| {
            Error::InvalidData("BoundSheet8 points outside the Workbook stream".to_string())
        })?;
        let mut records = Records::new(worksheet_data);

        // Start directly at the worksheet BOF. The borrowed BIFF iterator
        // reports offsets relative to this slice; the absolute base is passed
        // below for row-block indexing.
        let bof = match records.next() {
            Some(record_result) => record_result?,
            None => return Err(Error::Eof("Expected BOF record for worksheet")),
        };

        if bof.kind().get() != 0x0809 {
            return Err(Error::UnexpectedRecordType {
                expected: 0x0809,
                found: bof.kind().get(),
            });
        }

        // Parse worksheet records (clone Arc is cheap - just increments ref count)
        let shared_strings = self
            .shared_strings
            .clone()
            .unwrap_or_else(|| Arc::new(Vec::new()));
        let shared_string_properties = self
            .shared_string_properties
            .clone()
            .unwrap_or_else(|| Arc::new(Vec::new()));
        let worksheet_position = worksheet_position as u64;
        let current_position = worksheet_position
            .checked_add(bof.encoded().len() as u64)
            .ok_or_else(|| Error::InvalidData("worksheet position overflows".to_string()))?;
        Self::parse_worksheet_records(
            &mut records,
            workbook_data.len() as u64,
            worksheet_position,
            current_position,
            encoding,
            &bound_sheet.name,
            shared_strings,
            shared_string_properties,
            Some(&self.formula_context),
            self.formatting.clone(),
        )
    }

    /// Parse worksheet records sequentially
    pub(crate) fn parse_worksheet_records(
        records_iter: &mut Records<'_>,
        stream_len: u64,
        base_position: u64,
        current_position: u64,
        encoding: &Encoding,
        name: &str,
        shared_strings: Arc<Vec<String>>,
        shared_string_properties: Arc<Vec<Option<Box<SharedStringProperties>>>>,
        formula_context: Option<&FormulaContext>,
        formatting: Arc<Formatting>,
    ) -> Result<Worksheet> {
        let mut worksheet = Worksheet::with_shared_string_properties(
            name.to_string(),
            shared_strings,
            shared_string_properties,
        );
        worksheet.set_formatting(formatting.clone());

        let mut pivot_table_collector = pivot_table::PivotTableCollector::new();

        let mut comment_collector = comments::CommentCollector::new();
        let mut hyperlink_collector = hyperlinks::HyperlinkCollector::new();
        let mut layout_collector = layout::Collector::new();
        let mut sheet_layout_collector = crate::worksheet::layout::Collector::new();
        let mut view_collector = view::ViewCollector::new();
        let mut page_setup_collector = page_setup::PageSetupCollector::new();
        let mut protection_collector = protection::SheetProtectionCollector::new();
        let mut conditional_format_collector =
            conditional_format::ConditionalFormatCollector::new();
        let mut calculation_collector = crate::calculation::WorksheetCalculationCollector::new();
        let mut scenario_collector = crate::scenario::ScenarioCollector::new();
        let mut vba_collector = crate::vba::WorksheetVbaCollector::new();
        let mut consolidation_collector = crate::consolidation::ConsolidationCollector::new();
        let mut formula_error_collector = crate::formula_errors::FormulaErrorCollector::new();
        let mut list_object_collector = crate::list_object::ListObjectCollector::new();
        // The query-table collector must see records before the sort-data
        // collector: a QUERYTABLE sequence may end with a SortData record
        // (0x0895) that must not be mistaken for the sheet-level sort.
        let mut query_table_collector = crate::query_table::QueryTableCollector::new();
        let mut sort_data_collector = crate::sort_data::SortDataCollector::default();
        let mut row_block_index_collector =
            crate::row_block_index::RowBlockIndexCollector::new(stream_len, current_position);
        let mut pending_string_formula: Option<CellRecord> = None;
        let mut shared_formulas = HashMap::<(u16, u16), SharedFormulaTemplate>::new();
        let array_limits = array::Limits::default();
        let tracking_limit = array_limits.max_cells();
        let mut array_formulas = Vec::<Arc<ArrayFormula>>::new();
        let mut array_formula_by_anchor = HashMap::<(u16, u16), Arc<ArrayFormula>>::new();
        let mut array_occupancy = HashMap::<u16, [u64; 4]>::new();
        let mut ptg_exp_cells = HashMap::<(u16, u16), FormulaLink>::new();
        let mut duplicate_cells = HashSet::<(u16, u16)>::new();
        let mut companion_claims = HashMap::<(u16, u16), CompanionKind>::new();
        let mut immediately_preceding_formula: Option<FormulaLink> = None;
        let mut dimensions: Option<DimensionsRecord> = None;
        let mut remaining_data_validations: Option<usize> = None;

        while let Some(record_result) = records_iter.next() {
            let record = record_result?;
            let record_position = base_position + record.offset() as u64;
            sheet_layout_collector.feed_record(record.kind().get(), record.payload())?;
            comment_collector.feed_record(record.kind().get(), record.payload())?;
            hyperlink_collector.feed_record(record.kind().get(), record.payload())?;
            layout_collector.feed_record(record.kind().get(), record.payload(), &formatting)?;
            view_collector.feed_record(record.kind().get(), record.payload())?;
            page_setup_collector.feed_record(record.kind().get(), record.payload())?;
            protection_collector.feed_record(record.kind().get(), record.payload())?;
            conditional_format_collector.feed_record(
                record.kind().get(),
                record.payload(),
                formula_context,
            )?;
            calculation_collector.feed_record(record.kind().get(), record.payload())?;
            scenario_collector.feed_record(record.kind().get(), record.payload())?;
            vba_collector.feed_record(record.kind().get(), record.payload())?;
            consolidation_collector.feed_record(record.kind().get(), record.payload())?;
            formula_error_collector.feed_record(record.kind().get(), record.payload())?;
            list_object_collector.feed_record(record.kind().get(), record.payload())?;
            let query_table_consumed =
                query_table_collector.feed_record(record.kind().get(), record.payload());
            if !query_table_consumed
                && let Some(sort_data) =
                    sort_data_collector.feed_record(record.kind().get(), record.payload())?
            {
                worksheet.set_extended_sort(sort_data);
            }
            row_block_index_collector.feed_record(
                record_position,
                record.kind().get(),
                record.payload(),
            );
            let preceding_formula = immediately_preceding_formula.take();

            if matches!(remaining_data_validations, Some(1..))
                && record.kind().get() != crate::data_validation::DV_RECORD_TYPE
            {
                return Err(Error::InvalidRecord {
                    record_type: record.kind().get(),
                    message: "DVAL must be followed immediately by its declared DV records"
                        .to_string(),
                });
            }

            if let Some(mut formula) = pending_string_formula.take() {
                if record.kind().get() == 0x0207 {
                    // The cached string result may span Continue records
                    // (MS-XLS 2.1: FORMULA = ... [String *Continue]).
                    let mut continues = Vec::new();
                    let text = loop {
                        match utils::decode_string_record(record.payload(), &continues)? {
                            utils::StringRecordDecode::Complete(text) => break text,
                            utils::StringRecordDecode::NeedContinue => {
                                let next =
                                    records_iter.next().ok_or_else(|| Error::InvalidRecord {
                                        record_type: 0x0207,
                                        message: "String result ends before its Continue records"
                                            .to_string(),
                                    })??;
                                if next.kind().get() != 0x003C {
                                    return Err(Error::InvalidRecord {
                                        record_type: next.kind().get(),
                                        message:
                                            "String result continuation must be a Continue record"
                                                .to_string(),
                                    });
                                }
                                continues.push(next.payload().to_vec());
                            },
                        }
                    };
                    if let CellRecord::Formula { value, .. } = &mut formula {
                        *value = FormulaValue::String(text);
                    }
                    formatting.validate_cell_xf(cell_record_xf(&formula))?;
                    if let Some(mut cell) = Cell::from_record_with_formula_context(
                        &formula,
                        worksheet.shared_strings(),
                        formula_context,
                        Some(&formatting),
                    ) {
                        if let CellRecord::Formula {
                            row, col, formula, ..
                        } = &formula
                            && let Some(anchor) = ptg_exp_anchor(formula)
                            && let Some(rendered) = shared_formulas
                                .get(&anchor)
                                .and_then(|template| template.render(formula_context, *row, *col))
                        {
                            cell.set_rendered_formula(Some(rendered));
                        }
                        add_cell(&mut worksheet, cell, &mut duplicate_cells, tracking_limit)?;
                    }
                    continue;
                }
                // Per MS-XLS 2.1 (FORMULA = Formula [Array / Table / ShrFmla /
                // SUB] [String *Continue]) these records legally intervene
                // between the Formula and its String; keep the formula pending
                // and let the records flow through the normal walk.
                if !matches!(record.kind().get(), 0x0221 | 0x0236 | 0x04BC | 0x0091) {
                    return Err(Error::InvalidRecord {
                        record_type: record.kind().get(),
                        message: "String-valued Formula must be followed by a String record"
                            .to_string(),
                    });
                }
                pending_string_formula = Some(formula);
            }

            if !query_table_consumed
                && pivot_table_collector
                    .feed_record(record.kind().get(), record.payload())
                    .map_err(|error| Error::InvalidRecord {
                        record_type: record.kind().get(),
                        message: error.to_string(),
                    })?
            {
                continue;
            }

            match record.kind().get() {
                0x0809 => { // BOF - Beginning of worksheet
                    // This marks the start of a worksheet
                }
                crate::data_validation::DVAL_RECORD_TYPE => {
                    if remaining_data_validations.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::data_validation::DVAL_RECORD_TYPE,
                            message: "worksheet contains more than one DVAL record".to_string(),
                        });
                    }
                    let settings = crate::data_validation::parse_dval(record.payload())?;
                    remaining_data_validations = Some(usize::from(settings.declared_rule_count()));
                    worksheet.set_data_validation_settings(settings);
                }
                crate::data_validation::DV_RECORD_TYPE => {
                    let remaining = remaining_data_validations.as_mut().ok_or_else(|| {
                        Error::InvalidRecord {
                            record_type: crate::data_validation::DV_RECORD_TYPE,
                            message: "DV record appears without a preceding DVAL record".to_string(),
                        }
                    })?;
                    if *remaining == 0 {
                        return Err(Error::InvalidRecord {
                            record_type: crate::data_validation::DV_RECORD_TYPE,
                            message: "DV record exceeds the count declared by DVAL".to_string(),
                        });
                    }
                    worksheet.add_data_validation(crate::data_validation::parse_dv(
                        record.payload(),
                        formula_context,
                    )?);
                    *remaining -= 1;
                }
                0x000A => { // EOF - End of worksheet
                    *worksheet.protection_mut() = protection_collector.finish()?;
                    break;
                }
                crate::sheet_ext::SHEET_EXT_RECORD_TYPE => { // SheetExt
                    if worksheet.sheet_ext().is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::sheet_ext::SHEET_EXT_RECORD_TYPE,
                            message: "worksheet contains more than one SheetExt record".to_string(),
                        });
                    }
                    worksheet.set_sheet_ext(crate::sheet_ext::SheetExt::parse(record.payload())?);
                }
                crate::data_table::TABLE_RECORD_TYPE => { // Table
                    claim_companion(
                        preceding_formula,
                        CompanionKind::Table,
                        crate::data_table::TABLE_RECORD_TYPE,
                        &mut companion_claims,
                        tracking_limit,
                    )?;
                    worksheet.add_data_table(crate::data_table::DataTable::parse(record.payload())?);
                }
                0x0091 => { // SUB
                    claim_companion(
                        preceding_formula,
                        CompanionKind::Sub,
                        0x0091,
                        &mut companion_claims,
                        tracking_limit,
                    )?;
                }
                0x0207 => {
                    return Err(Error::InvalidRecord {
                        record_type: 0x0207,
                        message: "String record has no pending string-valued Formula".to_string(),
                    });
                }
                crate::web_pub::WEB_PUB_RECORD_TYPE => { // WebPub
                    worksheet
                        .add_web_publication(crate::web_pub::WebPub::parse(record.payload())?);
                }
                crate::custom_view::USER_S_VIEW_BEGIN_RECORD_TYPE => { // UserSViewBegin
                    let begin =
                        crate::custom_view::SheetCustomViewBegin::parse(record.payload())?;
                    // CUSTOMVIEW = UserSViewBegin *CUSTOMVIEWCONTENT
                    // UserSViewEnd: the inner records duplicate ordinary
                    // sheet settings for the view, so consume them inertly
                    // instead of feeding them to the sheet collectors.
                    let end = loop {
                        let next = records_iter.next().ok_or_else(|| Error::InvalidRecord {
                            record_type: crate::custom_view::USER_S_VIEW_BEGIN_RECORD_TYPE,
                            message: "UserSViewBegin is not closed by a UserSViewEnd".to_string(),
                        })??;
                        if next.kind().get()
                            == crate::custom_view::USER_S_VIEW_END_RECORD_TYPE
                        {
                            let end = crate::custom_view::SheetCustomViewEnd::parse(
                                next.payload(),
                            )?;
                            // The closing record is consumed by this loop rather than the
                            // normal collector feed path. Close the stateful page-setup
                            // collector explicitly so primary records after the custom-view
                            // bracket are not mistaken for custom-view content.
                            page_setup_collector.feed_record(next.kind().get(), next.payload())?;
                            break end;
                        }
                    };
                    worksheet
                        .add_custom_view(crate::custom_view::SheetCustomView::new(begin, end));
                }
                crate::custom_view::USER_S_VIEW_END_RECORD_TYPE => { // UserSViewEnd
                    return Err(Error::InvalidRecord {
                        record_type: crate::custom_view::USER_S_VIEW_END_RECORD_TYPE,
                        message: "UserSViewEnd without a matching UserSViewBegin".to_string(),
                    });
                }
                crate::phonetic_info::PHONETIC_INFO_RECORD_TYPE => { // PhoneticInfo
                    if worksheet.phonetic_info().is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::phonetic_info::PHONETIC_INFO_RECORD_TYPE,
                            message: "worksheet contains more than one PhoneticInfo record"
                                .to_string(),
                        });
                    }
                    // PHONETICINFO = PhoneticInfo *Continue: pull Continue
                    // records while the declared range list is incomplete.
                    let mut payload = record.payload().to_vec();
                    while payload.len() < 6
                        || payload.len()
                            < 6 + usize::from(u16::from_le_bytes([payload[4], payload[5]])) * 6
                    {
                        let next = records_iter.next().ok_or(Error::InvalidLength {
                            expected: payload.len() + 1,
                            found: payload.len(),
                        })??;
                        if next.kind().get() != 0x003C {
                            return Err(Error::InvalidRecord {
                                record_type: next.kind().get(),
                                message: "PhoneticInfo continuation must be a Continue record"
                                    .to_string(),
                            });
                        }
                        payload.extend_from_slice(next.payload());
                    }
                    worksheet
                        .set_phonetic_info(crate::phonetic_info::PhoneticInfo::parse(&payload)?);
                }
                0x0200 => { // Dimensions
                    if let Ok(parsed_dimensions) = DimensionsRecord::parse(record.payload()) {
                        worksheet.set_dimensions(
                            parsed_dimensions.first_row,
                            parsed_dimensions.last_row,
                            parsed_dimensions.first_col,
                            parsed_dimensions.last_col,
                        );
                        dimensions = Some(parsed_dimensions);
                    }
                }
                // Cell records
                0x0201 | // Blank
                0x0203 | // Number
                0x0204 | // Label
                0x0205 | // BoolErr
                0x027E | // RK
                0x00FD | // LabelSst
                0x0006   // Formula
                => {
                    let cell_record = CellRecord::parse(record.kind().get(), record.payload(), encoding)?;
                    formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                    if let Some(link) = FormulaLink::from_record(&cell_record) {
                        let position = (link.row, link.col);
                        if link.anchor.is_some() {
                            if ptg_exp_cells.len() >= tracking_limit {
                                return Err(Error::InvalidRecord {
                                    record_type: 0x0006,
                                    message: "PtgExp Formula count exceeds the configured limit"
                                        .to_string(),
                                });
                            }
                            ptg_exp_cells
                                .try_reserve(1)
                                .map_err(|_| Error::Allocation("tracking PtgExp Formula cells"))?;
                            ptg_exp_cells.insert(position, link);
                        }
                        immediately_preceding_formula = Some(link);
                    }
                    if matches!(
                        &cell_record,
                        CellRecord::Formula {
                            value: FormulaValue::StringPending,
                            ..
                        }
                    ) {
                        pending_string_formula = Some(cell_record);
                    } else if let Some(mut cell) = Cell::from_record_with_formula_context(
                        &cell_record,
                        worksheet.shared_strings(),
                        formula_context,
                        Some(&formatting),
                    ) {
                        if let CellRecord::Formula {
                            row, col, formula, ..
                        } = &cell_record
                            && let Some(anchor) = ptg_exp_anchor(formula)
                                && let Some(rendered) = shared_formulas
                                    .get(&anchor)
                                    .and_then(|template| {
                                        template.render(formula_context, *row, *col)
                                    })
                                {
                                    cell.set_rendered_formula(Some(rendered));
                                }
                        add_cell(
                            &mut worksheet,
                            cell,
                            &mut duplicate_cells,
                            tracking_limit,
                        )?;
                    }
                }

                0x04BC => { // ShrFmla
                    let predecessor = claim_companion(
                        preceding_formula,
                        CompanionKind::Shared,
                        0x04BC,
                        &mut companion_claims,
                        tracking_limit,
                    )?;
                    let template = parse_shared_formula_template(
                        record.kind().get(),
                        record.payload(),
                    )?;
                    let anchor = (template.first_row, template.first_col);
                    if (predecessor.row, predecessor.col) != anchor
                        || predecessor.anchor != Some(anchor)
                        || !predecessor.shared
                    {
                        return Err(Error::InvalidRecord {
                            record_type: 0x04BC,
                            message: "ShrFmla must follow its fShrFmla/PtgExp anchor Formula"
                                .to_string(),
                        });
                    }
                    if shared_formulas.contains_key(&anchor)
                        || array_formula_by_anchor.contains_key(&anchor)
                    {
                        return Err(Error::InvalidRecord {
                            record_type: 0x04BC,
                            message: "Formula anchor already owns an Array or ShrFmla record"
                                .to_string(),
                        });
                    }
                    if shared_formulas.len() >= tracking_limit {
                        return Err(Error::InvalidRecord {
                            record_type: 0x04BC,
                            message: "ShrFmla owner count exceeds the configured limit".to_string(),
                        });
                    }
                    shared_formulas
                        .try_reserve(1)
                        .map_err(|_| Error::Allocation("tracking ShrFmla owners"))?;
                    let rendered = template.render(formula_context, anchor.0, anchor.1);
                    shared_formulas.insert(anchor, template);
                    if let Some(cell) = worksheet.get_cell_mut(
                        u32::from(anchor.0),
                        u32::from(anchor.1),
                    ) {
                        cell.set_rendered_formula(rendered);
                    }
                }

                0x0221 => { // Array
                    let predecessor = claim_companion(
                        preceding_formula,
                        CompanionKind::Array,
                        0x0221,
                        &mut companion_claims,
                        tracking_limit,
                    )?;
                    let owner = Arc::new(array::parse_payload(
                        record.payload(),
                        array_limits,
                    )?);
                    let anchor = owner.anchor();
                    let anchor_position = (anchor.row(), u16::from(anchor.col()));
                    if (predecessor.row, predecessor.col) != anchor_position
                        || predecessor.anchor != Some(anchor_position)
                    {
                        return Err(Error::InvalidRecord {
                            record_type: 0x0221,
                            message: "Array anchor Formula must be at RefU top-left and contain its exact standalone PtgExp".to_string(),
                        });
                    }
                    if predecessor.shared {
                        return Err(Error::InvalidRecord {
                            record_type: 0x0221,
                            message: "Array anchor Formula must have fShrFmla cleared".to_string(),
                        });
                    }
                    if array_formula_by_anchor.contains_key(&anchor_position)
                        || shared_formulas.contains_key(&anchor_position)
                    {
                        return Err(Error::InvalidRecord {
                            record_type: 0x0221,
                            message: "Array owner duplicates or overlaps an existing Array range"
                                .to_string(),
                        });
                    }
                    if array_formulas.len() >= tracking_limit {
                        return Err(Error::InvalidRecord {
                            record_type: 0x0221,
                            message: "Array owner count exceeds the configured limit".to_string(),
                        });
                    }
                    array_formulas
                        .try_reserve(1)
                        .map_err(|_| Error::Allocation("tracking Array owners"))?;
                    array_formula_by_anchor
                        .try_reserve(1)
                        .map_err(|_| Error::Allocation("indexing Array owners"))?;
                    claim_array_range(&mut array_occupancy, &owner)?;
                    array_formula_by_anchor.insert(anchor_position, Arc::clone(&owner));
                    array_formulas.push(owner);
                }

                0x00BD => { // MulRk
                    for cell_record in CellRecord::parse_mul_rk(record.payload())? {
                        formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                        if let Some(cell) = Cell::from_record_with_formula_context(
                            &cell_record,
                            worksheet.shared_strings(),
                            formula_context,
                            Some(&formatting),
                        ) {
                            add_cell(
                                &mut worksheet,
                                cell,
                                &mut duplicate_cells,
                                tracking_limit,
                            )?;
                        }
                    }
                }

                0x00BE => { // MulBlank
                    for cell_record in CellRecord::parse_mul_blank(record.payload())? {
                        formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                        if let Some(cell) = Cell::from_record_with_formula_context(
                            &cell_record,
                            worksheet.shared_strings(),
                            formula_context,
                            Some(&formatting),
                        ) {
                            add_cell(
                                &mut worksheet,
                                cell,
                                &mut duplicate_cells,
                                tracking_limit,
                            )?;
                        }
                    }
                }

                // --- Merged cells (MERGECELLS 0x00E5) ---
                rt if rt == merged_cells::RECORD_TYPE => {
                    let mut ranges = Vec::new();
                    if merged_cells::parse_mergecells_record(record.payload(), &mut ranges).is_ok() {
                        worksheet.add_merged_cells(&ranges);
                    }
                }

                // --- AutoFilter (AUTOFILTERINFO 0x009D) ---
                rt if rt == autofilter::AUTOFILTERINFO_TYPE => {
                    if let Ok(count) = autofilter::parse_autofilterinfo(record.payload()) {
                        worksheet.set_autofilter_info(count);
                    }
                }

                // --- AutoFilter column (AUTOFILTER 0x009E) ---
                rt if rt == autofilter::AUTOFILTER_TYPE => {
                    if let Ok(col) = autofilter::parse_autofilter(record.payload()) {
                        worksheet.add_autofilter_column(col);
                    }
                }

                // --- Sort (SORT 0x0090) ---
                rt if rt == autofilter::SORT_TYPE => {
                    if let Ok(info) = autofilter::parse_sort(record.payload()) {
                        worksheet.set_sort_info(info);
                    }
                }

                // --- Filter mode (FILTERMODE 0x009B) ---
                rt if rt == autofilter::FILTERMODE_TYPE
                    && autofilter::parse_filtermode(record.payload()).is_ok() =>
                {
                    worksheet.set_filter_mode(true);
                }

                _ => {
                    // Skip other records
                }
            }
        }

        if matches!(remaining_data_validations, Some(1..)) {
            return Err(Error::InvalidRecord {
                record_type: crate::data_validation::DVAL_RECORD_TYPE,
                message: "worksheet ended before all DV records declared by DVAL".to_string(),
            });
        }

        for (position, link) in &ptg_exp_cells {
            let Some(anchor) = link.anchor else {
                continue;
            };
            let resolved_array = !link.shared
                && array_formula_by_anchor.get(&anchor).is_some_and(|owner| {
                    owner.range().contains(FormulaCell::new(
                        position.0,
                        u8::try_from(position.1).unwrap_or(u8::MAX),
                    ))
                });
            let resolved_shared = link.shared
                && shared_formulas
                    .get(&anchor)
                    .is_some_and(|owner| owner.contains(position.0, position.1));
            if !resolved_array && !resolved_shared {
                return Err(Error::InvalidRecord {
                    record_type: 0x0006,
                    message: format!(
                        "Formula at ({}, {}) contains an orphan PtgExp",
                        position.0, position.1
                    ),
                });
            }
        }

        if !array_formulas.is_empty() && dimensions.is_none() {
            return Err(Error::InvalidRecord {
                record_type: 0x0221,
                message: "Array formulas require worksheet Dimensions".to_string(),
            });
        }

        for owner in &array_formulas {
            let range = owner.range();
            if let Some(dimensions) = &dimensions
                && (u32::from(range.first().row()) < dimensions.first_row
                    || u32::from(range.last().row()) >= dimensions.last_row
                    || u32::from(range.first().col()) < dimensions.first_col
                    || u32::from(range.last().col()) >= dimensions.last_col)
            {
                return Err(Error::InvalidRecord {
                    record_type: 0x0221,
                    message: "Array range is outside worksheet Dimensions".to_string(),
                });
            }
            let anchor = owner.anchor();
            let anchor_position = (anchor.row(), u16::from(anchor.col()));
            let rendered: Option<Arc<str>> =
                render_formula(owner.tokens(), formula_context).map(Arc::from);
            for coordinate in owner.cells() {
                let position = (coordinate.row(), u16::from(coordinate.col()));
                let link = ptg_exp_cells
                    .get(&position)
                    .ok_or_else(|| Error::InvalidRecord {
                        record_type: 0x0221,
                        message: format!(
                            "Array range cell ({}, {}) has no Formula/PtgExp record",
                            position.0, position.1
                        ),
                    })?;
                if duplicate_cells.contains(&position)
                    || link.anchor != Some(anchor_position)
                    || link.shared
                {
                    return Err(Error::InvalidRecord {
                        record_type: 0x0221,
                        message: format!(
                            "Array range cell ({}, {}) is not exactly one Formula/PtgExp to its anchor",
                            position.0, position.1
                        ),
                    });
                }
                let cell = worksheet
                    .get_cell_mut(u32::from(position.0), u32::from(position.1))
                    .ok_or_else(|| Error::InvalidRecord {
                        record_type: 0x0221,
                        message: "Array Formula cell was not materialized".to_string(),
                    })?;
                cell.set_rendered_formula_arc(rendered.clone());
                if !cell.set_array_formula(Arc::clone(owner)) {
                    return Err(Error::InvalidRecord {
                        record_type: 0x0221,
                        message: "Array owner cannot attach to a non-Formula cell".to_string(),
                    });
                }
            }
        }
        worksheet.set_array_formulas(array_formulas);
        for table in pivot_table_collector.finish()? {
            worksheet.add_pivot_table(table);
        }
        worksheet.set_query_tables(query_table_collector.finish());
        sort_data_collector.finish()?;

        worksheet.set_comments(comment_collector.finish()?);
        worksheet.set_hyperlinks(hyperlink_collector.finish());
        let (row_layouts, column_layouts) = layout_collector.finish();
        worksheet.set_layouts(row_layouts, column_layouts);
        worksheet.set_layout(sheet_layout_collector.finish());
        worksheet.set_worksheet_views(view_collector.finish()?);
        worksheet.set_page_setup(page_setup_collector.finish()?);
        let (legacy, future, extensions) = conditional_format_collector.finish()?;
        worksheet.set_conditional_formattings(legacy);
        worksheet.set_conditional_formattings12(future);
        worksheet.set_conditional_format_extensions(extensions);
        worksheet.set_calculation(calculation_collector.finish()?);
        worksheet.set_scenario_manager(scenario_collector.finish()?);
        worksheet.set_vba_code_name(vba_collector.finish());
        worksheet.set_consolidation(consolidation_collector.finish());
        worksheet.set_formula_error_features(formula_error_collector.finish()?);
        worksheet.set_list_objects(list_object_collector.finish()?);
        worksheet.set_row_block_index(row_block_index_collector.finish());

        Ok(worksheet)
    }
}
