//! BIFF workbook and worksheet substream decoding.
//!
//! The parser follows the MS-XLS hierarchy: records are decoded inside the
//! workbook globals and sheet substreams, while the owning CFB stream is
//! handled by the sibling package module.

use super::model::Workbook;
use crate::cell::Cell;
use crate::defined_names::{DefinedNameSlot, LBL_RECORD_TYPE, NAME_CMT_RECORD_TYPE};
use crate::error::{Error, Result};
use crate::formula::{FormulaContext, ptg_exp_anchor, render_formula, render_shared_formula};
use crate::leniency::ToleranceLog;
use crate::number_format::{DateSystem, Formatting};
use crate::records::{
    BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, Encoding, FormulaValue,
    SharedStringProperties, SharedStringTable,
};
use crate::worksheet::Worksheet;
use crate::{
    autofilter, comments, conditional_format, hyperlinks, layout, merged_cells, page_setup,
    pivot_table, protection, utils, view,
};
use litchi_biff::Records;
use std::collections::HashMap;
use std::io::{Read, Seek};
use std::sync::Arc;

#[derive(Debug)]
struct SharedFormulaTemplate {
    first_row: u16,
    last_row: u16,
    first_col: u16,
    last_col: u16,
    tokens: Vec<u8>,
    relative: bool,
}

fn cell_record_xf(record: &CellRecord) -> u16 {
    match record {
        CellRecord::Blank { xf_index, .. }
        | CellRecord::Number { xf_index, .. }
        | CellRecord::Label { xf_index, .. }
        | CellRecord::BoolErr { xf_index, .. }
        | CellRecord::Rk { xf_index, .. }
        | CellRecord::LabelSst { xf_index, .. }
        | CellRecord::Formula { xf_index, .. } => *xf_index,
    }
}

impl SharedFormulaTemplate {
    fn contains(&self, row: u16, col: u16) -> bool {
        (self.first_row..=self.last_row).contains(&row)
            && (self.first_col..=self.last_col).contains(&col)
    }

    fn render(&self, context: Option<&FormulaContext>, row: u16, col: u16) -> Option<String> {
        if !self.contains(row, col) {
            return None;
        }
        if self.relative {
            render_shared_formula(&self.tokens, context, row, col)
        } else {
            render_formula(&self.tokens, context)
        }
    }
}

fn parse_shared_formula_template(record_type: u16, data: &[u8]) -> Result<SharedFormulaTemplate> {
    let (fixed_size, length_offset, relative) = match record_type {
        0x04bc => (10usize, 8usize, true),
        0x0221 => (14usize, 12usize, false),
        _ => {
            return Err(Error::UnexpectedRecordType {
                expected: 0x04bc,
                found: record_type,
            });
        },
    };
    if data.len() < fixed_size {
        return Err(Error::InvalidLength {
            expected: fixed_size,
            found: data.len(),
        });
    }
    let first_row = u16::from_le_bytes([data[0], data[1]]);
    let last_row = u16::from_le_bytes([data[2], data[3]]);
    let first_col = u16::from(data[4]);
    let last_col = u16::from(data[5]);
    if first_row > last_row || first_col > last_col {
        return Err(Error::InvalidRecord {
            record_type,
            message: "shared formula range is reversed".to_string(),
        });
    }
    let token_len = usize::from(u16::from_le_bytes([
        data[length_offset],
        data[length_offset + 1],
    ]));
    let end = fixed_size
        .checked_add(token_len)
        .ok_or_else(|| Error::InvalidRecord {
            record_type,
            message: "shared formula token length overflows".to_string(),
        })?;
    let tokens = data
        .get(fixed_size..end)
        .ok_or(Error::InvalidLength {
            expected: end,
            found: data.len(),
        })?
        .to_vec();
    if tokens.is_empty() {
        return Err(Error::InvalidRecord {
            record_type,
            message: "shared formula token stream is empty".to_string(),
        });
    }
    Ok(SharedFormulaTemplate {
        first_row,
        last_row,
        first_col,
        last_col,
        tokens,
        relative,
    })
}

pub(super) fn pivot_cache_stream_paths(
    paths: impl IntoIterator<Item = Vec<String>>,
) -> Vec<(u16, Vec<String>)> {
    let mut cache_paths = paths
        .into_iter()
        .filter_map(|path| {
            if path.len() != 2 || !path[0].eq_ignore_ascii_case("_SX_DB_CUR") {
                return None;
            }
            let stream_id = u16::from_str_radix(&path[1], 16).ok()?;
            Some((stream_id, path))
        })
        .collect::<Vec<_>>();
    cache_paths.sort_unstable_by_key(|(stream_id, _)| *stream_id);
    cache_paths
}

pub(super) struct WorkbookGlobalsSink<'a> {
    /// BoundSheet8 entries in stream order.
    pub(super) bound_sheets: &'a mut Vec<BoundSheetRecord>,
    /// Shared-string table contents.
    pub(super) strings: &'a mut Vec<String>,
    /// Rich-text and phonetic properties parallel to strings.
    pub(super) string_properties: &'a mut Vec<Option<Box<SharedStringProperties>>>,
    /// Lbl records and their trailing optional records.
    pub(super) defined_name_slots: &'a mut Vec<DefinedNameSlot>,
    /// Formatting-defect policy and the repairs it recorded.
    pub(super) tolerance: &'a mut ToleranceLog,
}

impl<R: Read + Seek> Workbook<R> {
    /// Parse workbook globals (SST, bound sheets, etc.)
    pub(super) fn parse_workbook_globals(
        &mut self,
        records_iter: &mut Records<'_>,
        encoding: &mut Encoding,
        sink: WorkbookGlobalsSink<'_>,
    ) -> Result<()> {
        let WorkbookGlobalsSink {
            bound_sheets,
            strings,
            string_properties,
            defined_name_slots,
            tolerance,
        } = sink;
        // Collect all records first for easier processing
        let mut records = Vec::new();
        for record_result in records_iter.by_ref() {
            let record = record_result?;
            let is_globals_eof = record.kind().get() == 0x000A;
            records.push(record);
            if is_globals_eof {
                break;
            }
        }

        self.formatting = Arc::new(Formatting::parse_globals(&records, tolerance)?);
        self.is_1904_date_system = self.formatting.date_system() == DateSystem::Excel1904;

        let mut palette_seen = false;
        let mut protection_collector = protection::WorkbookProtectionCollector::new();
        let mut calculation_collector = crate::calculation::WorkbookCalculationCollector::new();
        let mut vba_collector = crate::vba::WorkbookVbaCollector::new();
        let mut environment_collector = crate::environment::EnvironmentCollector::new();
        let mut write_access_collector = crate::access::WriteAccessCollector::new();
        let mut table_styles_collector = crate::table_styles::TableStylesCollector::new();
        let mut shared_string_index_collector =
            crate::shared_string_index::SharedStringIndexCollector::new();
        let mut workbook_view_collector = crate::workbook_view::WorkbookViewCollector::new();
        let mut function_group_collector = crate::function_group::FunctionGroupCollector::new();
        let mut external_link_collector = crate::external_link::ExternalLinkCollector::new();
        let mut name_optional_target: Option<(usize, u8)> = None;
        let mut i = 0;
        while i < records.len() {
            let record = &records[i];
            if !matches!(
                record.kind().get(),
                LBL_RECORD_TYPE
                    | NAME_CMT_RECORD_TYPE
                    | crate::defined_names::NAME_FN_GRP12_RECORD_TYPE
                    | crate::defined_names::NAME_PUBLISH_RECORD_TYPE
                    | 0x003c
            ) {
                name_optional_target = None;
            }
            if record.kind().get() == 0x003c && name_optional_target.is_some() {
                return Err(Error::InvalidRecord {
                    record_type: 0x003c,
                    message: "CONTINUE is not permitted after a post-Lbl optional record"
                        .to_string(),
                });
            }
            protection_collector.feed_record(record.kind().get(), record.payload())?;
            calculation_collector.feed_record(record.kind().get(), record.payload())?;
            vba_collector.feed_record(record.kind().get(), record.payload())?;
            environment_collector.feed_record(record.kind().get(), record.payload())?;
            write_access_collector.feed_record(record.kind().get(), record.payload());
            table_styles_collector.feed_record(record.kind().get(), record.payload())?;
            shared_string_index_collector.feed_record(record.kind().get(), record.payload());
            workbook_view_collector.feed_record(record.kind().get(), record.payload())?;
            function_group_collector.feed_record(record.kind().get(), record.payload())?;
            external_link_collector.feed_record(record.kind().get(), record.payload())?;

            match record.kind().get() {
                crate::font::FONT_RECORD_TYPE => {
                    let index = crate::font::logical_font_index(self.fonts.len())?;
                    self.fonts
                        .push(crate::font::Font::parse_record(
                            index,
                            record.payload(),
                            tolerance,
                        )?);
                },
                0x0092 => {
                    self.palette = crate::palette::Palette::parse_unique_record(
                        record.payload(),
                        &mut palette_seen,
                    )?;
                },
                crate::theme::THEME_RECORD_TYPE => {
                    if self.theme.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::theme::THEME_RECORD_TYPE,
                            message: "workbook contains more than one Theme record".to_string(),
                        });
                    }
                    let mut continues = Vec::new();
                    while records
                        .get(i + 1)
                        .is_some_and(|next| next.kind().get()
                            == crate::theme::CONTINUE_FRT12_RECORD_TYPE)
                    {
                        i += 1;
                        continues.push(records[i].payload().to_vec());
                    }
                    self.theme = Some(crate::theme::Theme::parse(record.payload(), &continues)?);
                },
                crate::style_ext::STYLE_EXT_RECORD_TYPE => {
                    self.style_extensions
                        .push(crate::style_ext::StyleExt::parse(record.payload())?);
                },
                crate::custom_view::USER_B_VIEW_RECORD_TYPE => {
                    self.custom_views
                        .push(crate::custom_view::WorkbookCustomView::parse(record.payload())?);
                },
                crate::real_time_data::REAL_TIME_DATA_RECORD_TYPE => {
                    // RTD = RealTimeData *ContinueFrt (MS-XLS 2.1): the
                    // logical payload is the record body plus any trailing
                    // ContinueFrt bodies.
                    let mut payload = record.payload().to_vec();
                    while records.get(i + 1).is_some_and(|next| {
                        next.kind().get()
                            == crate::real_time_data::CONTINUE_FRT_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(records[i].payload());
                    }
                    let previous_topic =
                        self.real_time_data.last().map(|topic| topic.topic.as_str());
                    self.real_time_data.push(
                        crate::real_time_data::RealTimeData::parse(&payload, previous_topic)?,
                    );
                },
                crate::web_pub::WEB_PUB_RECORD_TYPE => {
                    self.web_publications
                        .push(crate::web_pub::WebPub::parse(record.payload())?);
                },
                crate::mdx_metadata::MDT_INFO_RECORD_TYPE
                | crate::mdx_metadata::MDX_STR_RECORD_TYPE
                | crate::mdx_metadata::MDX_TUPLE_RECORD_TYPE
                | crate::mdx_metadata::MDX_SET_RECORD_TYPE
                | crate::mdx_metadata::MDX_PROP_RECORD_TYPE
                | crate::mdx_metadata::MDX_KPI_RECORD_TYPE
                | crate::mdx_metadata::MDB_RECORD_TYPE => {
                    // METADATA records are continued by ContinueFrt12 (MS-XLS
                    // 2.1): the logical payload is the record body plus any
                    // trailing ContinueFrt12 bodies.
                    let mut payload = record.payload().to_vec();
                    while records.get(i + 1).is_some_and(|next| {
                        next.kind().get()
                            == crate::mdx_metadata::CONTINUE_FRT12_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(records[i].payload());
                    }
                    self.mdx_metadata
                        .push_record(record.kind().get(), &payload)?;
                },
                crate::book_ext::BOOK_EXT_RECORD_TYPE => {
                    if self.book_ext.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::book_ext::BOOK_EXT_RECORD_TYPE,
                            message: "workbook contains more than one BookExt record".to_string(),
                        });
                    }
                    self.book_ext = Some(crate::book_ext::BookExt::parse(record.payload())?);
                },
                0x0809 => {
                    // BOF
                    let bof = BofRecord::parse(record.payload())?;
                    self.biff_version = bof.version;
                    self.is_1904_date_system = bof.is_1904_date_system;
                },
                0x0042
                    // CodePage
                    if record.payload().len() >= 2 => {
                        let codepage = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                        *encoding = Encoding::from_codepage(codepage)?;
                    },
                0x0022
                    // Date1904
                    if record.payload().len() >= 2 => {
                        let flag = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                        self.is_1904_date_system = flag == 1;
                    },
                0x0085 => {
                    // BoundSheet8
                    let sheet = BoundSheetRecord::parse(record.payload(), encoding)?;
                    bound_sheets.push(sheet);
                },
                0x00D5 => {
                    if record.payload().len() != 2 {
                        return Err(Error::InvalidLength { expected: 2, found: record.payload().len() });
                    }
                    let stream_id = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                    if stream_id == 0 || self.pivot_cache_stream_ids.contains(&stream_id) {
                        return Err(Error::InvalidRecord {
                            record_type: 0x00D5,
                            message: "SXStreamID must be nonzero and unique".to_string(),
                        });
                    }
                    self.pivot_cache_stream_ids.push(stream_id);
                },
                0x01AE => {
                    // SUPBOOK: retain its position and whether it references
                    // this workbook so EXTERNSHEET indices remain stable.
                    self.formula_context.add_sup_book(record.payload());
                },
                0x0017 => {
                    // EXTERNSHEET: XTI entries used by PtgRef3d/PtgArea3d.
                    self.formula_context
                        .add_extern_sheet(record.payload())
                        .map_err(|message| Error::InvalidRecord {
                            record_type: 0x0017,
                            message: message.to_string(),
                        })?;
                },
                LBL_RECORD_TYPE => {
                    if defined_name_slots.len() == usize::from(u16::MAX) {
                        return Err(Error::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "workbook contains more than 65535 Lbl records".to_string(),
                        });
                    }
                    let record_index = u32::try_from(defined_name_slots.len() + 1)
                        .map_err(|_| Error::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "Lbl record index overflows".to_string(),
                        })?;
                    let mut combined = record.payload().to_vec();
                    let mut continuation_chunks = Vec::new();
                    let mut next = i + 1;
                    while next < records.len() && records[next].kind().get() == 0x003c {
                        if combined
                            .len()
                            .checked_add(records[next].payload().len())
                            .is_none_or(|len| len > 1_048_576)
                        {
                            return Err(Error::InvalidRecord {
                                record_type: LBL_RECORD_TYPE,
                                message: "Lbl continuation data exceeds resource bound".to_string(),
                            });
                        }
                        continuation_chunks.push(records[next].payload().to_vec());
                        combined.extend_from_slice(records[next].payload());
                        next += 1;
                    }
                    defined_name_slots.push(DefinedNameSlot::parse_with_continuations(
                        &combined, record_index, continuation_chunks,
                    )?);
                    name_optional_target=Some((defined_name_slots.len()-1,0));
                    i=next-1;
                },
                NAME_CMT_RECORD_TYPE => {
                    let (target,stage)=name_optional_target.ok_or_else(|| Error::InvalidRecord {
                        record_type: NAME_CMT_RECORD_TYPE,
                        message: "NameCmt does not immediately follow a Lbl record".to_string(),
                    })?;
                    if stage!=0{return Err(Error::InvalidRecord{record_type:NAME_CMT_RECORD_TYPE,message:"NameCmt is duplicated or out of order in the Lbl optional-record sequence".to_string()})}
                    defined_name_slots[target].attach_comment(record.payload())?;
                    name_optional_target=Some((target,1));
                },
                crate::defined_names::NAME_FN_GRP12_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||Error::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 does not follow a Lbl record".to_string()})?;if stage>1{return Err(Error::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_function_group(record.payload())?;name_optional_target=Some((target,2));},
                crate::defined_names::NAME_PUBLISH_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||Error::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish does not follow a Lbl record".to_string()})?;if stage>2{return Err(Error::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_publication(record.payload())?;name_optional_target=Some((target,3));},
                0x00FC => {
                    // SST
                    // SST may span multiple records, collect them all
                    let mut sst_records = vec![*record];
                    let mut sst_idx = i + 1;

                    // Collect all following CONTINUE records
                    while sst_idx < records.len() && records[sst_idx].kind().get() == 0x003C {
                        sst_records.push(records[sst_idx]);
                        sst_idx += 1;
                    }

                    let sst = SharedStringTable::parse_from_records(&sst_records, encoding)?;
                    self.shared_string_reference_count = sst.total_count;
                    strings.extend(sst.strings);
                    string_properties.extend(sst.properties);

                    // Skip the CONTINUE records we consumed
                    i = sst_idx - 1;
                },
                0x000A => {
                    // EOF - End of workbook globals
                    crate::font::validate_font_table(&self.fonts)?;
                    self.protection = protection_collector.finish()?;
                    self.calculation = calculation_collector.finish();
                    self.vba_metadata = vba_collector.finish();
                    self.environment = environment_collector.finish()?;
                    self.write_access = write_access_collector.finish();
                    self.table_styles = table_styles_collector
                        .finish(self.formatting.differential_formats().len())?;
                    self.shared_string_index = shared_string_index_collector.finish();
                    self.workbook_view = workbook_view_collector.finish(bound_sheets.len())?;
                    self.function_groups = function_group_collector.finish()?;
                    let extended_count=self.function_groups.as_ref().map_or(0,|groups|groups.extended_categories().len());for slot in defined_name_slots.iter(){slot.validate_extended_category(extended_count)?;}
                    self.external_links = external_link_collector.finish(bound_sheets.len())?;
                    break;
                },
                _ => {
                    // Skip other records for now
                },
            }
            i += 1;
        }

        Ok(())
    }

    /// Parse a worksheet from its position in the workbook stream
    pub(super) fn parse_worksheet_from_position(
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
    pub(super) fn parse_worksheet_records(
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
                        worksheet.add_cell(cell);
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
                    worksheet.add_data_table(crate::data_table::DataTable::parse(record.payload())?);
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
                    if let Ok(dimensions) = DimensionsRecord::parse(record.payload()) {
                        worksheet.set_dimensions(dimensions.first_row, dimensions.last_row,
                                               dimensions.first_col, dimensions.last_col);
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
                        worksheet.add_cell(cell);
                    }
                }

                0x04BC | 0x0221 => { // ShrFmla or Array
                    let template = parse_shared_formula_template(
                        record.kind().get(),
                        record.payload(),
                    )?;
                    let anchor = (template.first_row, template.first_col);
                    let rendered = template.render(formula_context, anchor.0, anchor.1);
                    shared_formulas.insert(anchor, template);
                    if let Some(cell) = worksheet.get_cell_mut(
                        u32::from(anchor.0),
                        u32::from(anchor.1),
                    ) {
                        cell.set_rendered_formula(rendered);
                    }
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
                            worksheet.add_cell(cell);
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
                            worksheet.add_cell(cell);
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
