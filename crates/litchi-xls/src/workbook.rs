//! Workbook implementation for XLS files

use crate::cell::XlsCell;
use crate::defined_names::{
    DefinedNameSlot, LBL_RECORD_TYPE, NAME_CMT_RECORD_TYPE, XlsBuiltInName, XlsDefinedName,
    XlsDefinedNameKind, XlsNameScope,
};
use crate::encryption::prepare_workbook_stream;
use crate::error::{XlsError, XlsResult};
use crate::formula::{FormulaContext, ptg_exp_anchor, render_formula, render_shared_formula};
use crate::leniency::{XlsLeniency, XlsToleranceLog, XlsToleranceReport};
use crate::number_format::{XlsDateSystem, XlsExtendedFormat, XlsFormatting, XlsNumberFormat};
use crate::records::{
    BiffVersion, BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, FormulaValue,
    RecordIter, SharedStringProperties, SharedStringTable, XlsEncoding,
};
use crate::sheet_metadata::XlsSheetMetadata;
use crate::worksheet::XlsWorksheet;
use crate::{
    autofilter, comments, conditional_format, hyperlinks, layout, merged_cells, page_setup,
    pivot_table, protection, utils, view,
};
use litchi_cfb::OleFile;
use litchi_core::sheet::{Result, Worksheet as SheetTrait, WorksheetIterator};
use std::collections::{HashMap, HashSet};
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

fn parse_shared_formula_template(
    record_type: u16,
    data: &[u8],
) -> XlsResult<SharedFormulaTemplate> {
    let (fixed_size, length_offset, relative) = match record_type {
        0x04bc => (10usize, 8usize, true),
        0x0221 => (14usize, 12usize, false),
        _ => {
            return Err(XlsError::UnexpectedRecordType {
                expected: 0x04bc,
                found: record_type,
            });
        },
    };
    if data.len() < fixed_size {
        return Err(XlsError::InvalidLength {
            expected: fixed_size,
            found: data.len(),
        });
    }
    let first_row = u16::from_le_bytes([data[0], data[1]]);
    let last_row = u16::from_le_bytes([data[2], data[3]]);
    let first_col = u16::from(data[4]);
    let last_col = u16::from(data[5]);
    if first_row > last_row || first_col > last_col {
        return Err(XlsError::InvalidRecord {
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
        .ok_or_else(|| XlsError::InvalidRecord {
            record_type,
            message: "shared formula token length overflows".to_string(),
        })?;
    let tokens = data
        .get(fixed_size..end)
        .ok_or(XlsError::InvalidLength {
            expected: end,
            found: data.len(),
        })?
        .to_vec();
    if tokens.is_empty() {
        return Err(XlsError::InvalidRecord {
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

/// XLS workbook implementation
#[derive(Debug)]
pub struct XlsWorkbook<R: Read + Seek> {
    ole_file: OleFile<R>,
    worksheets: Vec<XlsWorksheet>,
    worksheet_names: Vec<String>,
    sheets: Vec<XlsSheetMetadata>,
    /// Shared string table (Arc for zero-copy sharing across worksheets)
    shared_strings: Option<Arc<Vec<String>>>,
    /// Sparse rich-text and phonetic properties parallel to `shared_strings`.
    shared_string_properties: Option<Arc<Vec<Option<Box<SharedStringProperties>>>>>,
    shared_string_reference_count: u32,
    palette: crate::palette::XlsPalette,
    fonts: Vec<crate::font::XlsFont>,
    biff_version: BiffVersion,
    is_1904_date_system: bool,
    formula_context: FormulaContext,
    defined_names: Vec<XlsDefinedName>,
    defined_name_records: Vec<XlsDefinedName>,
    formatting: Arc<XlsFormatting>,
    protection: protection::WorkbookProtection,
    calculation: crate::calculation::XlsWorkbookCalculation,
    vba_metadata: crate::vba::XlsVbaMetadata,
    environment: crate::environment::XlsWorkbookEnvironment,
    book_ext: Option<crate::book_ext::XlsBookExt>,
    style_extensions: Vec<crate::style_ext::XlsStyleExt>,
    theme: Option<crate::theme::XlsTheme>,
    write_access: crate::XlsResult<Option<crate::access::XlsWriteAccess>>,
    table_styles: Option<crate::table_styles::XlsTableStyles>,
    shared_string_index: XlsResult<Option<crate::shared_string_index::XlsSharedStringIndex>>,
    workbook_view: crate::workbook_view::XlsWorkbookView,
    /// Workbook-wide custom views (`UserBView` records), in record order.
    custom_views: Vec<crate::custom_view::XlsWorkbookCustomView>,
    /// Real-time data (RTD) topics (`RealTimeData` records), in record order.
    real_time_data: Vec<crate::real_time_data::XlsRealTimeData>,
    /// MDX (OLAP cube) metadata from the workbook globals `METADATA` production.
    mdx_metadata: crate::mdx_metadata::XlsMdxMetadata,
    /// Published Web pages (`WebPub` records), in record order.
    web_publications: Vec<crate::web_pub::XlsWebPub>,
    function_groups: Option<crate::function_group::XlsFunctionGroups>,
    external_links: crate::external_link::XlsExternalLinks,
    pivot_caches: Vec<crate::PivotCache>,
    /// SXStreamID values in global PivotCache ordinal order.
    pivot_cache_stream_ids: Vec<u16>,
    /// Formatting defects repaired while opening; always empty in strict mode.
    tolerance: XlsToleranceReport,
}

/// Out-parameters filled while scanning the workbook globals substream.
///
/// The globals pass produces several independent collections that the caller
/// consumes afterwards; bundling them keeps the scan a single-responsibility
/// function instead of a long positional parameter list.
struct WorkbookGlobalsSink<'a> {
    /// `BoundSheet8` entries in stream order.
    bound_sheets: &'a mut Vec<BoundSheetRecord>,
    /// Shared-string table contents.
    strings: &'a mut Vec<String>,
    /// Rich-text and phonetic properties parallel to `strings`.
    string_properties: &'a mut Vec<Option<Box<SharedStringProperties>>>,
    /// `Lbl` records and their trailing optional records.
    defined_name_slots: &'a mut Vec<DefinedNameSlot>,
    /// Formatting-defect policy and the repairs it recorded.
    tolerance: &'a mut XlsToleranceLog,
}

/// Options for opening a legacy XLS workbook.
#[derive(Debug, Clone, Copy, Default)]
pub struct XlsOpenOptions<'a> {
    /// Password used for BIFF8 password-to-open encryption.
    pub password: Option<&'a str>,
    /// How non-structural formatting defects are treated.
    ///
    /// Defaults to [`XlsLeniency::Strict`], which rejects any deviation from
    /// MS-XLS. Set [`XlsLeniency::TolerateFormattingDefects`] to open the
    /// widespread real-world workbooks whose cosmetic formatting metadata is
    /// self-contradictory; everything repaired is then enumerable through
    /// [`XlsWorkbook::tolerance_report`]. Structural defects — record framing,
    /// stream grammar, and encryption — remain hard errors either way.
    pub leniency: XlsLeniency,
}

impl<R: Read + Seek> XlsWorkbook<R> {
    /// Open an XLS workbook from a reader
    pub fn new(reader: R) -> XlsResult<Self> {
        Self::new_with_options(reader, XlsOpenOptions::default())
    }

    /// Open an XLS workbook with an explicit password contract.
    pub fn new_with_options(reader: R, options: XlsOpenOptions<'_>) -> XlsResult<Self> {
        let ole_file = OleFile::open(reader)?;

        let mut workbook = XlsWorkbook {
            ole_file,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            sheets: Vec::new(),
            shared_strings: None,
            shared_string_properties: None,
            shared_string_reference_count: 0,
            palette: crate::palette::XlsPalette::default(),
            fonts: Vec::new(),
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
            external_links: crate::XlsExternalLinks::default(),
            pivot_caches: Vec::new(),
            pivot_cache_stream_ids: Vec::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            formatting: Arc::new(XlsFormatting::default()),
            protection: protection::WorkbookProtection::default(),
            calculation: crate::calculation::XlsWorkbookCalculation::default(),
            vba_metadata: crate::vba::XlsVbaMetadata::default(),
            environment: crate::environment::XlsWorkbookEnvironment::default(),
            book_ext: None,
            style_extensions: Vec::new(),
            theme: None,
            write_access: Ok(None),
            table_styles: None,
            shared_string_index: Ok(None),
            workbook_view: crate::workbook_view::XlsWorkbookView::default(),
            custom_views: Vec::new(),
            real_time_data: Vec::new(),
            mdx_metadata: crate::mdx_metadata::XlsMdxMetadata::default(),
            web_publications: Vec::new(),
            function_groups: None,
            tolerance: XlsToleranceReport::default(),
        };

        workbook.parse_workbook(&options)?;
        Ok(workbook)
    }

    /// Create an XLS workbook from an already-parsed OLE file.
    ///
    /// This is used for single-pass parsing where the OLE file has already
    /// been parsed during format detection. It avoids double-parsing.
    ///
    /// # Arguments
    ///
    /// * `ole_file` - An already-parsed OLE file
    pub fn from_ole_file(ole_file: OleFile<R>) -> XlsResult<Self> {
        Self::from_ole_file_with_options(ole_file, XlsOpenOptions::default())
    }

    /// Create a workbook from a parsed OLE file with explicit open options.
    pub fn from_ole_file_with_options(
        ole_file: OleFile<R>,
        options: XlsOpenOptions<'_>,
    ) -> XlsResult<Self> {
        let mut workbook = XlsWorkbook {
            ole_file,
            worksheets: Vec::new(),
            worksheet_names: Vec::new(),
            sheets: Vec::new(),
            shared_strings: None,
            shared_string_properties: None,
            shared_string_reference_count: 0,
            palette: crate::palette::XlsPalette::default(),
            fonts: Vec::new(),
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
            external_links: crate::XlsExternalLinks::default(),
            pivot_caches: Vec::new(),
            pivot_cache_stream_ids: Vec::new(),
            defined_names: Vec::new(),
            defined_name_records: Vec::new(),
            formatting: Arc::new(XlsFormatting::default()),
            protection: protection::WorkbookProtection::default(),
            calculation: crate::calculation::XlsWorkbookCalculation::default(),
            vba_metadata: crate::vba::XlsVbaMetadata::default(),
            environment: crate::environment::XlsWorkbookEnvironment::default(),
            book_ext: None,
            style_extensions: Vec::new(),
            theme: None,
            write_access: Ok(None),
            table_styles: None,
            shared_string_index: Ok(None),
            workbook_view: crate::workbook_view::XlsWorkbookView::default(),
            custom_views: Vec::new(),
            real_time_data: Vec::new(),
            mdx_metadata: crate::mdx_metadata::XlsMdxMetadata::default(),
            web_publications: Vec::new(),
            function_groups: None,
            tolerance: XlsToleranceReport::default(),
        };

        workbook.parse_workbook(&options)?;
        Ok(workbook)
    }

    /// Formatting defects repaired while opening this workbook.
    ///
    /// Always clean under [`XlsLeniency::Strict`], because a strict open either
    /// rejects the defect or never encounters one. Under
    /// [`XlsLeniency::TolerateFormattingDefects`] every repair the reader made
    /// is enumerable here; see [`crate::XlsFormattingDefect`] for the
    /// closed set of defects that can appear and the substitute value each one
    /// produced.
    pub fn tolerance_report(&self) -> &XlsToleranceReport {
        &self.tolerance
    }

    /// Parse the workbook stream
    fn parse_workbook(&mut self, options: &XlsOpenOptions<'_>) -> XlsResult<()> {
        // Find and read the Workbook stream
        let workbook_data = self
            .ole_file
            .open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;
        let workbook_data = prepare_workbook_stream(workbook_data, options.password)?;
        let mut tolerance = XlsToleranceLog::new(options.leniency);

        let mut record_iter = RecordIter::new(std::io::Cursor::new(&workbook_data))?;
        let mut encoding = XlsEncoding::from_codepage(1252)?; // Default codepage
        let mut bound_sheets = Vec::new();
        let mut strings = Vec::new();
        let mut string_properties = Vec::new();

        // Parse workbook globals
        let mut defined_name_slots = Vec::new();
        self.parse_workbook_globals(
            &mut record_iter,
            &mut encoding,
            WorkbookGlobalsSink {
                bound_sheets: &mut bound_sheets,
                strings: &mut strings,
                string_properties: &mut string_properties,
                defined_name_slots: &mut defined_name_slots,
                tolerance: &mut tolerance,
            },
        )?;
        self.tolerance = tolerance.into_report();

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.shared_string_properties = Some(Arc::new(string_properties));
        let all_sheet_names = bound_sheets
            .iter()
            .map(|sheet| sheet.name.clone())
            .collect::<Vec<_>>();
        let mut unique_sheet_names = HashSet::with_capacity(all_sheet_names.len());
        for name in &all_sheet_names {
            if !unique_sheet_names.insert(name.to_lowercase()) {
                return Err(XlsError::InvalidRecord {
                    record_type: 0x0085,
                    message: format!("duplicate case-insensitive BoundSheet8 name: {name:?}"),
                });
            }
        }
        self.sheets = bound_sheets
            .iter()
            .enumerate()
            .map(|(index, sheet)| XlsSheetMetadata::from_bound_sheet(index, sheet))
            .collect();
        self.worksheet_names.clear();
        self.formula_context.set_sheet_names(all_sheet_names);
        self.formula_context.set_scoped_defined_names(
            defined_name_slots
                .iter()
                .map(DefinedNameSlot::formula_symbol)
                .collect(),
        );
        self.formula_context
            .set_external_links(&self.external_links);
        self.defined_name_records = defined_name_slots
            .into_iter()
            .map(|slot| slot.into_public(bound_sheets.len(), &self.formula_context))
            .collect::<XlsResult<Vec<_>>>()?;
        self.defined_names = self
            .defined_name_records
            .iter()
            .filter(|name| !name.is_macro())
            .cloned()
            .collect();

        // Parse worksheets from positions in the workbook stream
        for (sheet_index, bound_sheet) in bound_sheets.iter().enumerate() {
            if bound_sheet.sheet_type != crate::records::SheetType::WorkSheet {
                continue;
            }
            match self.parse_worksheet_from_position(bound_sheet, &encoding, &mut record_iter) {
                Ok(worksheet) => {
                    let worksheet_index = self.worksheets.len();
                    self.worksheet_names.push(bound_sheet.name.clone());
                    self.worksheets.push(worksheet);
                    self.sheets[sheet_index].set_parsed_worksheet_index(worksheet_index);
                },
                Err(error @ XlsError::InvalidRecord { record_type, .. })
                    if pivot_table::is_worksheet_view_record(record_type) =>
                {
                    return Err(error);
                },
                Err(_) => {},
            }
        }

        let mut cache_paths = self
            .ole_file
            .list_streams()
            .into_iter()
            .filter(|path| {
                path.len() == 2
                    && path[0].eq_ignore_ascii_case("_SX_DB_CUR")
                    && u16::from_str_radix(&path[1], 16).is_ok()
            })
            .collect::<Vec<_>>();
        cache_paths.sort_by_key(|path| u16::from_str_radix(&path[1], 16).unwrap());
        let mut pivot_caches = Vec::with_capacity(cache_paths.len());
        for path in cache_paths {
            let expected_stream_id = u16::from_str_radix(&path[1], 16).unwrap();
            if expected_stream_id == 0 {
                return Err(XlsError::InvalidRecord {
                    record_type: 0x00C6,
                    message: "PivotCache storage stream ID must be nonzero".to_string(),
                });
            }
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = self.ole_file.open_stream(&refs)?;
            let cache = crate::pivot_table::parse_pivot_cache_stream(&data)?;
            if cache.stream_id() != expected_stream_id {
                return Err(XlsError::InvalidRecord {
                    record_type: 0x00C6,
                    message: format!(
                        "PivotCache storage stream {:04X} contains stream ID {:04X}",
                        expected_stream_id,
                        cache.stream_id()
                    ),
                });
            }
            pivot_caches.push(cache);
        }
        self.pivot_caches = pivot_caches;
        crate::pivot_table::validate_pivot_cache_links(
            &self.worksheets,
            &self.pivot_caches,
            &self.pivot_cache_stream_ids,
        )?;

        let visible_tabs = bound_sheets
            .iter()
            .map(|sheet| matches!(sheet.visible, crate::records::SheetVisible::Visible))
            .collect::<Vec<_>>();
        let selected_worksheet_tabs = self
            .sheets
            .iter()
            .map(|sheet| {
                sheet
                    .parsed_worksheet_index()
                    .and_then(|index| self.worksheets.get(index))
                    .and_then(|worksheet| worksheet.worksheet_view())
                    .map(|view| view.is_selected())
            })
            .collect::<Vec<_>>();
        self.workbook_view
            .validate_sheet_state(&visible_tabs, &selected_worksheet_tabs)?;

        Ok(())
    }

    /// Parse workbook globals (SST, bound sheets, etc.)
    fn parse_workbook_globals<Reader: Read + Seek>(
        &mut self,
        record_iter: &mut RecordIter<Reader>,
        encoding: &mut XlsEncoding,
        sink: WorkbookGlobalsSink<'_>,
    ) -> XlsResult<()> {
        let WorkbookGlobalsSink {
            bound_sheets,
            strings,
            string_properties,
            defined_name_slots,
            tolerance,
        } = sink;
        // Collect all records first for easier processing
        let mut records = Vec::new();
        for record_result in record_iter.by_ref() {
            let record = record_result?;
            let is_globals_eof = record.header.record_type == 0x000A;
            records.push(record);
            if is_globals_eof {
                break;
            }
        }

        self.formatting = Arc::new(XlsFormatting::parse_globals(&records, tolerance)?);
        self.is_1904_date_system = self.formatting.date_system() == XlsDateSystem::Excel1904;

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
                record.header.record_type,
                LBL_RECORD_TYPE
                    | NAME_CMT_RECORD_TYPE
                    | crate::defined_names::NAME_FN_GRP12_RECORD_TYPE
                    | crate::defined_names::NAME_PUBLISH_RECORD_TYPE
                    | 0x003c
            ) {
                name_optional_target = None;
            }
            if record.header.record_type == 0x003c && name_optional_target.is_some() {
                return Err(XlsError::InvalidRecord {
                    record_type: 0x003c,
                    message: "CONTINUE is not permitted after a post-Lbl optional record"
                        .to_string(),
                });
            }
            protection_collector.feed_record(record.header.record_type, &record.data)?;
            calculation_collector.feed_record(record.header.record_type, &record.data)?;
            vba_collector.feed_record(record.header.record_type, &record.data)?;
            environment_collector.feed_record(record.header.record_type, &record.data)?;
            write_access_collector.feed_record(record.header.record_type, &record.data);
            table_styles_collector.feed_record(record.header.record_type, &record.data)?;
            shared_string_index_collector.feed_record(record.header.record_type, &record.data);
            workbook_view_collector.feed_record(record.header.record_type, &record.data)?;
            function_group_collector.feed_record(record.header.record_type, &record.data)?;
            external_link_collector.feed_record(record.header.record_type, &record.data)?;

            match record.header.record_type {
                crate::font::FONT_RECORD_TYPE => {
                    let index = crate::font::logical_font_index(self.fonts.len())?;
                    self.fonts
                        .push(crate::font::XlsFont::parse_record(
                            index,
                            &record.data,
                            tolerance,
                        )?);
                },
                0x0092 => {
                    self.palette = crate::palette::XlsPalette::parse_unique_record(
                        &record.data,
                        &mut palette_seen,
                    )?;
                },
                crate::theme::THEME_RECORD_TYPE => {
                    if self.theme.is_some() {
                        return Err(XlsError::InvalidRecord {
                            record_type: crate::theme::THEME_RECORD_TYPE,
                            message: "workbook contains more than one Theme record".to_string(),
                        });
                    }
                    let mut continues = Vec::new();
                    while records
                        .get(i + 1)
                        .is_some_and(|next| next.header.record_type
                            == crate::theme::CONTINUE_FRT12_RECORD_TYPE)
                    {
                        i += 1;
                        continues.push(records[i].data.clone());
                    }
                    self.theme = Some(crate::theme::XlsTheme::parse(&record.data, &continues)?);
                },
                crate::style_ext::STYLE_EXT_RECORD_TYPE => {
                    self.style_extensions
                        .push(crate::style_ext::XlsStyleExt::parse(&record.data)?);
                },
                crate::custom_view::USER_B_VIEW_RECORD_TYPE => {
                    self.custom_views
                        .push(crate::custom_view::XlsWorkbookCustomView::parse(&record.data)?);
                },
                crate::real_time_data::REAL_TIME_DATA_RECORD_TYPE => {
                    // RTD = RealTimeData *ContinueFrt (MS-XLS 2.1): the
                    // logical payload is the record body plus any trailing
                    // ContinueFrt bodies.
                    let mut payload = record.data.clone();
                    while records.get(i + 1).is_some_and(|next| {
                        next.header.record_type
                            == crate::real_time_data::CONTINUE_FRT_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(&records[i].data);
                    }
                    let previous_topic =
                        self.real_time_data.last().map(|topic| topic.topic.as_str());
                    self.real_time_data.push(
                        crate::real_time_data::XlsRealTimeData::parse(&payload, previous_topic)?,
                    );
                },
                crate::web_pub::WEB_PUB_RECORD_TYPE => {
                    self.web_publications
                        .push(crate::web_pub::XlsWebPub::parse(&record.data)?);
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
                    let mut payload = record.data.clone();
                    while records.get(i + 1).is_some_and(|next| {
                        next.header.record_type
                            == crate::mdx_metadata::CONTINUE_FRT12_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(&records[i].data);
                    }
                    self.mdx_metadata
                        .push_record(record.header.record_type, &payload)?;
                },
                crate::book_ext::BOOK_EXT_RECORD_TYPE => {
                    if self.book_ext.is_some() {
                        return Err(XlsError::InvalidRecord {
                            record_type: crate::book_ext::BOOK_EXT_RECORD_TYPE,
                            message: "workbook contains more than one BookExt record".to_string(),
                        });
                    }
                    self.book_ext = Some(crate::book_ext::XlsBookExt::parse(&record.data)?);
                },
                0x0809 => {
                    // BOF
                    let bof = BofRecord::parse(&record.data)?;
                    self.biff_version = bof.version;
                    self.is_1904_date_system = bof.is_1904_date_system;
                },
                0x0042
                    // CodePage
                    if record.data.len() >= 2 => {
                        let codepage = litchi_core::binary::read_u16_le_at(&record.data, 0)?;
                        *encoding = XlsEncoding::from_codepage(codepage)?;
                    },
                0x0022
                    // Date1904
                    if record.data.len() >= 2 => {
                        let flag = litchi_core::binary::read_u16_le_at(&record.data, 0)?;
                        self.is_1904_date_system = flag == 1;
                    },
                0x0085 => {
                    // BoundSheet8
                    let sheet = BoundSheetRecord::parse(&record.data, encoding)?;
                    bound_sheets.push(sheet);
                },
                0x00D5 => {
                    if record.data.len() != 2 {
                        return Err(XlsError::InvalidLength { expected: 2, found: record.data.len() });
                    }
                    let stream_id = u16::from_le_bytes(record.data[..2].try_into().expect("length checked"));
                    if stream_id == 0 || self.pivot_cache_stream_ids.contains(&stream_id) {
                        return Err(XlsError::InvalidRecord {
                            record_type: 0x00D5,
                            message: "SXStreamID must be nonzero and unique".to_string(),
                        });
                    }
                    self.pivot_cache_stream_ids.push(stream_id);
                },
                0x01AE => {
                    // SUPBOOK: retain its position and whether it references
                    // this workbook so EXTERNSHEET indices remain stable.
                    self.formula_context.add_sup_book(&record.data);
                },
                0x0017 => {
                    // EXTERNSHEET: XTI entries used by PtgRef3d/PtgArea3d.
                    self.formula_context
                        .add_extern_sheet(&record.data)
                        .map_err(|message| XlsError::InvalidRecord {
                            record_type: 0x0017,
                            message: message.to_string(),
                        })?;
                },
                LBL_RECORD_TYPE => {
                    if defined_name_slots.len() == usize::from(u16::MAX) {
                        return Err(XlsError::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "workbook contains more than 65535 Lbl records".to_string(),
                        });
                    }
                    let record_index = u32::try_from(defined_name_slots.len() + 1)
                        .map_err(|_| XlsError::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "Lbl record index overflows".to_string(),
                        })?;
                    let mut combined = record.data.clone();
                    let mut continuation_chunks = Vec::new();
                    let mut next = i + 1;
                    while next < records.len() && records[next].header.record_type == 0x003c {
                        if combined.len().checked_add(records[next].data.len()).is_none_or(|len| len > 1_048_576) {
                            return Err(XlsError::InvalidRecord {
                                record_type: LBL_RECORD_TYPE,
                                message: "Lbl continuation data exceeds resource bound".to_string(),
                            });
                        }
                        continuation_chunks.push(records[next].data.clone());
                        combined.extend_from_slice(&records[next].data);
                        next += 1;
                    }
                    defined_name_slots.push(DefinedNameSlot::parse_with_continuations(
                        &combined, record_index, continuation_chunks,
                    )?);
                    name_optional_target=Some((defined_name_slots.len()-1,0));
                    i=next-1;
                },
                NAME_CMT_RECORD_TYPE => {
                    let (target,stage)=name_optional_target.ok_or_else(|| XlsError::InvalidRecord {
                        record_type: NAME_CMT_RECORD_TYPE,
                        message: "NameCmt does not immediately follow a Lbl record".to_string(),
                    })?;
                    if stage!=0{return Err(XlsError::InvalidRecord{record_type:NAME_CMT_RECORD_TYPE,message:"NameCmt is duplicated or out of order in the Lbl optional-record sequence".to_string()})}
                    defined_name_slots[target].attach_comment(&record.data)?;
                    name_optional_target=Some((target,1));
                },
                crate::defined_names::NAME_FN_GRP12_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||XlsError::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 does not follow a Lbl record".to_string()})?;if stage>1{return Err(XlsError::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_function_group(&record.data)?;name_optional_target=Some((target,2));},
                crate::defined_names::NAME_PUBLISH_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||XlsError::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish does not follow a Lbl record".to_string()})?;if stage>2{return Err(XlsError::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_publication(&record.data)?;name_optional_target=Some((target,3));},
                0x00FC => {
                    // SST
                    // SST may span multiple records, collect them all
                    let mut sst_records = vec![record.clone()];
                    let mut sst_idx = i + 1;

                    // Collect all following CONTINUE records
                    while sst_idx < records.len() && records[sst_idx].header.record_type == 0x003C {
                        sst_records.push(records[sst_idx].clone());
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
    fn parse_worksheet_from_position<Reader: Read + Seek>(
        &self,
        bound_sheet: &BoundSheetRecord,
        encoding: &XlsEncoding,
        record_iter: &mut RecordIter<Reader>,
    ) -> XlsResult<XlsWorksheet> {
        // Seek to the worksheet position
        record_iter.seek(bound_sheet.position as u64)?;

        // Skip the BOF record at the beginning of the worksheet
        if let Some(record_result) = record_iter.next() {
            let record = record_result?;
            if record.header.record_type != 0x0809 {
                // BOF
                return Err(XlsError::UnexpectedRecordType {
                    expected: 0x0809,
                    found: record.header.record_type,
                });
            }
        } else {
            return Err(XlsError::Eof("Expected BOF record for worksheet"));
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
        Self::parse_worksheet_records(
            record_iter,
            encoding,
            &bound_sheet.name,
            shared_strings,
            shared_string_properties,
            Some(&self.formula_context),
            self.formatting.clone(),
        )
    }

    /// Parse worksheet records sequentially
    fn parse_worksheet_records<Reader: Read + Seek>(
        record_iter: &mut RecordIter<Reader>,
        encoding: &XlsEncoding,
        name: &str,
        shared_strings: Arc<Vec<String>>,
        shared_string_properties: Arc<Vec<Option<Box<crate::records::SharedStringProperties>>>>,
        formula_context: Option<&FormulaContext>,
        formatting: Arc<XlsFormatting>,
    ) -> XlsResult<XlsWorksheet> {
        let mut worksheet = XlsWorksheet::with_shared_string_properties(
            name.to_string(),
            shared_strings,
            shared_string_properties,
        );
        worksheet.set_formatting(formatting.clone());

        let mut pivot_table_collector = pivot_table::PivotTableCollector::new();

        let mut comment_collector = comments::CommentCollector::new();
        let mut hyperlink_collector = hyperlinks::HyperlinkCollector::new();
        let mut layout_collector = layout::LayoutCollector::new();
        let mut sheet_layout_collector = crate::sheet_layout::SheetLayoutCollector::new();
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
        let mut row_block_index_collector = crate::row_block_index::RowBlockIndexCollector::new(
            record_iter.stream_len(),
            record_iter.current_position(),
        );
        let mut pending_string_formula: Option<CellRecord> = None;
        let mut shared_formulas = HashMap::<(u16, u16), SharedFormulaTemplate>::new();
        let mut remaining_data_validations: Option<usize> = None;

        while let Some((record_position, record_result)) = record_iter.next_positioned() {
            let record = record_result?;
            sheet_layout_collector.feed_record(record.header.record_type, &record.data)?;
            comment_collector.feed_record(record.header.record_type, &record.data)?;
            hyperlink_collector.feed_record(record.header.record_type, &record.data)?;
            layout_collector.feed_record(record.header.record_type, &record.data, &formatting)?;
            view_collector.feed_record(record.header.record_type, &record.data)?;
            page_setup_collector.feed_record(record.header.record_type, &record.data)?;
            protection_collector.feed_record(record.header.record_type, &record.data)?;
            conditional_format_collector.feed_record(
                record.header.record_type,
                &record.data,
                formula_context,
            )?;
            calculation_collector.feed_record(record.header.record_type, &record.data)?;
            scenario_collector.feed_record(record.header.record_type, &record.data)?;
            vba_collector.feed_record(record.header.record_type, &record.data)?;
            consolidation_collector.feed_record(record.header.record_type, &record.data)?;
            formula_error_collector.feed_record(record.header.record_type, &record.data)?;
            list_object_collector.feed_record(record.header.record_type, &record.data)?;
            let query_table_consumed =
                query_table_collector.feed_record(record.header.record_type, &record.data);
            if !query_table_consumed
                && let Some(sort_data) =
                    sort_data_collector.feed_record(record.header.record_type, &record.data)?
            {
                worksheet.set_extended_sort(sort_data);
            }
            row_block_index_collector.feed_record(
                record_position,
                record.header.record_type,
                &record.data,
            );

            if matches!(remaining_data_validations, Some(1..))
                && record.header.record_type != super::data_validation::DV_RECORD_TYPE
            {
                return Err(XlsError::InvalidRecord {
                    record_type: record.header.record_type,
                    message: "DVAL must be followed immediately by its declared DV records"
                        .to_string(),
                });
            }

            if let Some(mut formula) = pending_string_formula.take() {
                if record.header.record_type == 0x0207 {
                    // The cached string result may span Continue records
                    // (MS-XLS 2.1: FORMULA = ... [String *Continue]).
                    let mut continues = Vec::new();
                    let text = loop {
                        match utils::decode_string_record(&record.data, &continues)? {
                            utils::StringRecordDecode::Complete(text) => break text,
                            utils::StringRecordDecode::NeedContinue => {
                                let next = record_iter.next().ok_or_else(|| {
                                    XlsError::InvalidRecord {
                                        record_type: 0x0207,
                                        message: "String result ends before its Continue records"
                                            .to_string(),
                                    }
                                })??;
                                if next.header.record_type != 0x003C {
                                    return Err(XlsError::InvalidRecord {
                                        record_type: next.header.record_type,
                                        message:
                                            "String result continuation must be a Continue record"
                                                .to_string(),
                                    });
                                }
                                continues.push(next.data);
                            },
                        }
                    };
                    if let CellRecord::Formula { value, .. } = &mut formula {
                        *value = FormulaValue::String(text);
                    }
                    formatting.validate_cell_xf(cell_record_xf(&formula))?;
                    if let Some(mut cell) = XlsCell::from_record_with_formula_context(
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
                if !matches!(record.header.record_type, 0x0221 | 0x0236 | 0x04BC | 0x0091) {
                    return Err(XlsError::InvalidRecord {
                        record_type: record.header.record_type,
                        message: "String-valued Formula must be followed by a String record"
                            .to_string(),
                    });
                }
                pending_string_formula = Some(formula);
            }

            if !query_table_consumed
                && pivot_table_collector
                    .feed_record(record.header.record_type, &record.data)
                    .map_err(|error| XlsError::InvalidRecord {
                        record_type: record.header.record_type,
                        message: error.to_string(),
                    })?
            {
                continue;
            }

            match record.header.record_type {
                0x0809 => { // BOF - Beginning of worksheet
                    // This marks the start of a worksheet
                }
                super::data_validation::DVAL_RECORD_TYPE => {
                    if remaining_data_validations.is_some() {
                        return Err(XlsError::InvalidRecord {
                            record_type: super::data_validation::DVAL_RECORD_TYPE,
                            message: "worksheet contains more than one DVAL record".to_string(),
                        });
                    }
                    let settings = super::data_validation::parse_dval(&record.data)?;
                    remaining_data_validations = Some(usize::from(settings.declared_rule_count()));
                    worksheet.set_data_validation_settings(settings);
                }
                super::data_validation::DV_RECORD_TYPE => {
                    let remaining = remaining_data_validations.as_mut().ok_or_else(|| {
                        XlsError::InvalidRecord {
                            record_type: super::data_validation::DV_RECORD_TYPE,
                            message: "DV record appears without a preceding DVAL record".to_string(),
                        }
                    })?;
                    if *remaining == 0 {
                        return Err(XlsError::InvalidRecord {
                            record_type: super::data_validation::DV_RECORD_TYPE,
                            message: "DV record exceeds the count declared by DVAL".to_string(),
                        });
                    }
                    worksheet.add_data_validation(super::data_validation::parse_dv(&record.data, formula_context)?);
                    *remaining -= 1;
                }
                0x000A => { // EOF - End of worksheet
                    *worksheet.protection_mut() = protection_collector.finish()?;
                    break;
                }
                crate::sheet_ext::SHEET_EXT_RECORD_TYPE => { // SheetExt
                    if worksheet.sheet_ext().is_some() {
                        return Err(XlsError::InvalidRecord {
                            record_type: crate::sheet_ext::SHEET_EXT_RECORD_TYPE,
                            message: "worksheet contains more than one SheetExt record".to_string(),
                        });
                    }
                    worksheet.set_sheet_ext(crate::sheet_ext::XlsSheetExt::parse(&record.data)?);
                }
                crate::data_table::TABLE_RECORD_TYPE => { // Table
                    worksheet.add_data_table(crate::data_table::XlsDataTable::parse(&record.data)?);
                }
                crate::web_pub::WEB_PUB_RECORD_TYPE => { // WebPub
                    worksheet
                        .add_web_publication(crate::web_pub::XlsWebPub::parse(&record.data)?);
                }
                crate::custom_view::USER_S_VIEW_BEGIN_RECORD_TYPE => { // UserSViewBegin
                    let begin =
                        crate::custom_view::XlsSheetCustomViewBegin::parse(&record.data)?;
                    // CUSTOMVIEW = UserSViewBegin *CUSTOMVIEWCONTENT
                    // UserSViewEnd: the inner records duplicate ordinary
                    // sheet settings for the view, so consume them inertly
                    // instead of feeding them to the sheet collectors.
                    let end = loop {
                        let next = record_iter.next().ok_or_else(|| XlsError::InvalidRecord {
                            record_type: crate::custom_view::USER_S_VIEW_BEGIN_RECORD_TYPE,
                            message: "UserSViewBegin is not closed by a UserSViewEnd".to_string(),
                        })??;
                        if next.header.record_type
                            == crate::custom_view::USER_S_VIEW_END_RECORD_TYPE
                        {
                            break crate::custom_view::XlsSheetCustomViewEnd::parse(
                                &next.data,
                            )?;
                        }
                    };
                    worksheet
                        .add_custom_view(crate::custom_view::XlsSheetCustomView::new(begin, end));
                }
                crate::custom_view::USER_S_VIEW_END_RECORD_TYPE => { // UserSViewEnd
                    return Err(XlsError::InvalidRecord {
                        record_type: crate::custom_view::USER_S_VIEW_END_RECORD_TYPE,
                        message: "UserSViewEnd without a matching UserSViewBegin".to_string(),
                    });
                }
                crate::phonetic_info::PHONETIC_INFO_RECORD_TYPE => { // PhoneticInfo
                    if worksheet.phonetic_info().is_some() {
                        return Err(XlsError::InvalidRecord {
                            record_type: crate::phonetic_info::PHONETIC_INFO_RECORD_TYPE,
                            message: "worksheet contains more than one PhoneticInfo record"
                                .to_string(),
                        });
                    }
                    // PHONETICINFO = PhoneticInfo *Continue: pull Continue
                    // records while the declared range list is incomplete.
                    let mut payload = record.data.clone();
                    while payload.len() < 6
                        || payload.len()
                            < 6 + usize::from(u16::from_le_bytes([payload[4], payload[5]])) * 6
                    {
                        let next = record_iter.next().ok_or(XlsError::InvalidLength {
                            expected: payload.len() + 1,
                            found: payload.len(),
                        })??;
                        if next.header.record_type != 0x003C {
                            return Err(XlsError::InvalidRecord {
                                record_type: next.header.record_type,
                                message: "PhoneticInfo continuation must be a Continue record"
                                    .to_string(),
                            });
                        }
                        payload.extend_from_slice(&next.data);
                    }
                    worksheet
                        .set_phonetic_info(crate::phonetic_info::XlsPhoneticInfo::parse(&payload)?);
                }
                0x0200 => { // Dimensions
                    if let Ok(dimensions) = DimensionsRecord::parse(&record.data) {
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
                    let cell_record = CellRecord::parse(record.header.record_type, &record.data, encoding)?;
                    formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                    if matches!(
                        &cell_record,
                        CellRecord::Formula {
                            value: FormulaValue::StringPending,
                            ..
                        }
                    ) {
                        pending_string_formula = Some(cell_record);
                    } else if let Some(mut cell) = XlsCell::from_record_with_formula_context(
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
                        record.header.record_type,
                        &record.data,
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
                    for cell_record in CellRecord::parse_mul_rk(&record.data)? {
                        formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                        if let Some(cell) = XlsCell::from_record_with_formula_context(
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
                    for cell_record in CellRecord::parse_mul_blank(&record.data)? {
                        formatting.validate_cell_xf(cell_record_xf(&cell_record))?;
                        if let Some(cell) = XlsCell::from_record_with_formula_context(
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
                    if merged_cells::parse_mergecells_record(&record.data, &mut ranges).is_ok() {
                        worksheet.add_merged_cells(&ranges);
                    }
                }

                // --- AutoFilter (AUTOFILTERINFO 0x009D) ---
                rt if rt == autofilter::AUTOFILTERINFO_TYPE => {
                    if let Ok(count) = autofilter::parse_autofilterinfo(&record.data) {
                        worksheet.set_autofilter_info(count);
                    }
                }

                // --- AutoFilter column (AUTOFILTER 0x009E) ---
                rt if rt == autofilter::AUTOFILTER_TYPE => {
                    if let Ok(col) = autofilter::parse_autofilter(&record.data) {
                        worksheet.add_autofilter_column(col);
                    }
                }

                // --- Sort (SORT 0x0090) ---
                rt if rt == autofilter::SORT_TYPE => {
                    if let Ok(info) = autofilter::parse_sort(&record.data) {
                        worksheet.set_sort_info(info);
                    }
                }

                // --- Filter mode (FILTERMODE 0x009B) ---
                rt if rt == autofilter::FILTERMODE_TYPE
                    && autofilter::parse_filtermode(&record.data).is_ok() =>
                {
                    worksheet.set_filter_mode(true);
                }

                _ => {
                    // Skip other records
                }
            }
        }

        if matches!(remaining_data_validations, Some(1..)) {
            return Err(XlsError::InvalidRecord {
                record_type: super::data_validation::DVAL_RECORD_TYPE,
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
        worksheet.set_sheet_layout(sheet_layout_collector.finish());
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

    /// Optional `ExtSST` shared-string lookup index.
    ///
    /// A malformed optional index is reported here without preventing workbook content parsing.
    pub fn shared_string_index(
        &self,
    ) -> std::result::Result<Option<&crate::shared_string_index::XlsSharedStringIndex>, &XlsError>
    {
        match &self.shared_string_index {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(error),
        }
    }

    /// Workbook window state and stable sheet identifiers.
    pub fn workbook_view(&self) -> &crate::workbook_view::XlsWorkbookView {
        &self.workbook_view
    }

    /// Workbook-wide custom views (`UserBView` records), in record order.
    ///
    /// The views are inert: applying one is a UI operation this reader never
    /// performs.
    pub fn custom_views(&self) -> &[crate::custom_view::XlsWorkbookCustomView] {
        &self.custom_views
    }

    /// Real-time data (RTD) topics declared in the workbook globals, in
    /// record order.
    ///
    /// The topics are inert: this reader never locates, launches, or queries
    /// an RTD server; each entry only reports the topic, the last cached
    /// value, and the subscribed cells.
    pub fn real_time_data(&self) -> &[crate::real_time_data::XlsRealTimeData] {
        &self.real_time_data
    }

    /// MDX (OLAP cube) metadata collected from the workbook globals
    /// `METADATA` production; empty when the workbook carries none.
    ///
    /// The metadata is inert: connection names and MDX unique names are stored
    /// verbatim and no OLAP server is ever contacted.
    pub fn mdx_metadata(&self) -> &crate::mdx_metadata::XlsMdxMetadata {
        &self.mdx_metadata
    }

    /// Web pages published from the workbook globals, in record order.
    ///
    /// The records are inert: destination URLs and paths are never opened,
    /// resolved, or fetched.
    pub fn web_publications(&self) -> &[crate::web_pub::XlsWebPub] {
        &self.web_publications
    }

    /// Built-in and custom function categories, when the FNGROUPS collection exists.
    pub fn function_groups(&self) -> Option<&crate::function_group::XlsFunctionGroups> {
        self.function_groups.as_ref()
    }

    /// Inert supporting-book links and cached external cell values.
    pub fn external_links(&self) -> &crate::external_link::XlsExternalLinks {
        &self.external_links
    }

    /// Parsed workbook PivotCache streams ordered by their one-based stream ID.
    pub fn pivot_caches(&self) -> &[crate::PivotCache] {
        &self.pivot_caches
    }

    /// Read the legacy Custom XML Data Storage without resolving schema URIs.
    pub fn custom_xml_data_store(
        &mut self,
    ) -> litchi_ole_common::custom_xml_data::Result<
        Option<litchi_ole_common::custom_xml_data::MsoDataStore>,
    > {
        litchi_ole_common::custom_xml_data::inspect_mso_data_store(&mut self.ole_file)
    }

    pub fn summary_information(&mut self) -> XlsResult<Option<litchi_cfb::PropertySetStream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}SummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    /// Verify workbook XML signatures with the safe strict policy, without
    /// evaluating certificate trust or executing any macro content.
    pub fn signatures(&mut self) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        self.signatures_with(&litchi_sign::Policy::strict())
    }

    /// Verify workbook XML signatures with an explicit trust-neutral policy.
    pub fn signatures_with(
        &mut self,
        policy: &litchi_sign::Policy,
    ) -> litchi_sign::Result<Vec<litchi_sign::cfb::Report>> {
        litchi_sign::cfb::verify(&mut self.ole_file, litchi_sign::cfb::Format::Xls, policy)
    }

    pub fn document_summary_information(
        &mut self,
    ) -> XlsResult<Option<litchi_cfb::PropertySetStream>> {
        match self
            .ole_file
            .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        {
            Ok(value) => Ok(Some(value)),
            Err(litchi_cfb::OleError::StreamNotFound) => Ok(None),
            Err(error) => Err(error.into()),
        }
    }

    pub fn user_defined_properties(&mut self) -> XlsResult<Option<litchi_cfb::PropertySet>> {
        Ok(self.document_summary_information()?.and_then(|stream| {
            stream
                .section(litchi_cfb::USER_DEFINED_PROPERTIES_FMTID)
                .cloned()
        }))
    }

    /// Global PivotCache ordinal-to-storage-stream map from SXStreamID records.
    pub fn pivot_cache_stream_ids(&self) -> &[u16] {
        &self.pivot_cache_stream_ids
    }

    /// Resolves a worksheet PivotTable's global cache link.
    pub fn pivot_cache_for_table(
        &self,
        table: &crate::pivot_table::PivotTable,
    ) -> XlsResult<&crate::PivotCache> {
        let stream_id = *self
            .pivot_cache_stream_ids
            .get(usize::from(table.cache_index()))
            .ok_or_else(|| XlsError::InvalidRecord {
                record_type: crate::pivot_table::SXVIEW_TYPE,
                message: "PivotTable global cache index is out of range".to_string(),
            })?;
        self.pivot_caches
            .iter()
            .find(|cache| cache.stream_id() == stream_id)
            .ok_or_else(|| XlsError::InvalidRecord {
                record_type: crate::pivot_table::SXVIEW_TYPE,
                message: "PivotTable SXStreamID has no matching cache storage".to_string(),
            })
    }

    /// Parsed PivotTables on one worksheet.
    pub fn worksheet_pivot_tables(
        &self,
        index: usize,
    ) -> XlsResult<&[crate::pivot_table::PivotTable]> {
        Ok(self.xls_worksheet(index)?.pivot_tables())
    }

    /// Access the typed `XlsWorksheet` at the given index.
    ///
    /// This provides access to XLS-specific data (protection, comments,
    /// autofilter, pivot tables) that is not exposed through the generic
    /// `WorkbookTrait` / `Worksheet` trait.
    pub fn xls_worksheet(&self, index: usize) -> XlsResult<&XlsWorksheet> {
        self.worksheets
            .get(index)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet index {}", index)))
    }

    /// All workbook sheet directory entries in tab order.
    pub fn sheets(&self) -> &[XlsSheetMetadata] {
        &self.sheets
    }

    /// Sheet directory entry at a workbook tab index.
    pub fn sheet(&self, index: usize) -> Option<&XlsSheetMetadata> {
        self.sheets.get(index)
    }

    /// Case-insensitive sheet directory lookup.
    pub fn sheet_by_name(&self, name: &str) -> Option<&XlsSheetMetadata> {
        self.sheets
            .iter()
            .find(|sheet| sheet.name().eq_ignore_ascii_case(name))
    }

    /// Total number of cell references represented by the workbook SST.
    pub fn shared_string_reference_count(&self) -> u32 {
        self.shared_string_reference_count
    }

    pub fn formatting(&self) -> &XlsFormatting {
        &self.formatting
    }

    pub fn date_system(&self) -> XlsDateSystem {
        self.formatting.date_system()
    }

    pub fn protection(&self) -> &protection::WorkbookProtection {
        &self.protection
    }

    pub fn calculation(&self) -> &crate::calculation::XlsWorkbookCalculation {
        &self.calculation
    }

    pub fn vba_metadata(&self) -> crate::vba::XlsVbaMetadata {
        let mut metadata = self.vba_metadata.clone();
        metadata.set_project_storage_present(self.vba_project_storage().is_some());
        metadata
    }

    /// Discover the MS-XLS `_VBA_PROJECT_CUR` storage without opening macro streams.
    ///
    /// This validates directory names defined by MS-XLS and MS-OVBA only. It
    /// never opens, decompresses, parses, or executes `PROJECT`, `dir`,
    /// `_VBA_PROJECT`, SRP, or module-stream bytes.
    pub fn vba_project_storage(&self) -> Option<crate::XlsVbaProjectStorage> {
        crate::vba::discover_vba_project_storage(&self.ole_file.list_streams())
    }

    /// Parse the `_VBA_PROJECT_CUR` MS-OVBA project with safe default limits.
    ///
    /// The method returns `None` when no structurally complete VBA project is
    /// present. Source is only decompressed and decoded; it is never compiled,
    /// interpreted, or executed.
    pub fn vba(
        &mut self,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        self.vba_with(&litchi_vba::Limits::default())
    }

    /// Parse the `_VBA_PROJECT_CUR` project with explicit resource limits.
    pub fn vba_with(
        &mut self,
        limits: &litchi_vba::Limits,
    ) -> std::result::Result<Option<litchi_vba::project::Project>, litchi_vba::Error> {
        let Some(storage) = self.vba_project_storage() else {
            return Ok(None);
        };
        if !storage.is_structurally_complete() {
            return Ok(None);
        }
        let path: Vec<&str> = storage
            .root_storage_path()
            .iter()
            .map(String::as_str)
            .collect();
        litchi_vba::project::Project::open(&mut self.ole_file, &path, limits).map(Some)
    }

    /// Whether the CFB container holds a shared-workbook `Revision Log` stream
    /// (MS-XLS 2.1.7.14).
    pub fn has_revision_log(&self) -> bool {
        crate::revision_log::find_revision_log_stream(&self.ole_file.list_streams()).is_some()
    }

    /// Parse the shared-workbook `Revision Log` stream, when present.
    ///
    /// The result is a typed, inert model of the RRD revision records.
    /// Parsing never applies, rejects, or replays any recorded revision.
    pub fn revision_log(&mut self) -> XlsResult<Option<crate::revision_log::XlsRevisionLog>> {
        let Some(name) =
            crate::revision_log::find_revision_log_stream(&self.ole_file.list_streams())
                .map(str::to_owned)
        else {
            return Ok(None);
        };
        let data = self.ole_file.open_stream(&[name.as_str()])?;
        crate::revision_log::parse_revision_log_stream(&data).map(Some)
    }

    pub fn environment(&self) -> &crate::environment::XlsWorkbookEnvironment {
        &self.environment
    }

    /// Workbook extension flags from the `BookExt` record, when present.
    pub fn book_ext(&self) -> Option<&crate::book_ext::XlsBookExt> {
        self.book_ext.as_ref()
    }

    /// Cell-style extensions from `StyleExt` records, in record order.
    pub fn style_extensions(&self) -> &[crate::style_ext::XlsStyleExt] {
        &self.style_extensions
    }

    /// The document theme from the `Theme` record, when present.
    pub fn theme(&self) -> Option<&crate::theme::XlsTheme> {
        self.theme.as_ref()
    }

    /// Strictly access the user recorded as last creating, opening, or modifying the workbook.
    ///
    /// Noncanonical legacy producer variants are deferred until this metadata is requested.
    pub fn write_access(&self) -> XlsResult<Option<&crate::access::XlsWriteAccess>> {
        match &self.write_access {
            Ok(value) => Ok(value.as_ref()),
            Err(error) => Err(XlsError::InvalidRecord {
                record_type: crate::access::WRITE_ACCESS_RECORD_TYPE,
                message: format!("invalid WriteAccess metadata: {error}"),
            }),
        }
    }

    /// Default table and PivotTable style catalog, when present.
    pub fn table_styles(&self) -> Option<&crate::table_styles::XlsTableStyles> {
        self.table_styles.as_ref()
    }

    pub fn number_formats(&self) -> &[XlsNumberFormat] {
        self.formatting.number_formats()
    }

    pub fn extended_formats(&self) -> &[XlsExtendedFormat] {
        self.formatting.extended_formats()
    }

    /// Global differential formats referenced by custom table-style elements.
    pub fn differential_formats(&self) -> &[crate::differential_format::XlsDifferentialFormat] {
        self.formatting.differential_formats()
    }

    pub fn differential_format(
        &self,
        id: crate::table_styles::XlsDifferentialFormatId,
    ) -> Option<&crate::differential_format::XlsDifferentialFormat> {
        self.formatting.differential_format(id)
    }

    /// Resolves an XF's effective property families through its parent StyleXF.
    pub fn effective_extended_format(
        &self,
        index: u16,
    ) -> Option<crate::number_format::XlsEffectiveExtendedFormat<'_>> {
        self.formatting.effective_extended_format(index)
    }

    /// Workbook color palette, using BIFF8 defaults when no `Palette` record exists.
    pub fn palette(&self) -> &crate::palette::XlsPalette {
        &self.palette
    }

    /// Font records in physical workbook order.
    pub fn fonts(&self) -> &[crate::font::XlsFont] {
        &self.fonts
    }

    /// Resolve a BIFF8 logical font index. Index 4 is reserved and returns `None`.
    pub fn font(&self, index: u16) -> Option<&crate::font::XlsFont> {
        self.fonts.iter().find(|font| font.index() == index)
    }

    /// Resolve a font's color through the workbook palette.
    pub fn font_color(&self, index: u16) -> Option<crate::palette::XlsColor> {
        self.font(index)
            .and_then(|font| self.palette.color(font.color_index()))
    }

    /// Resolves the global Font record referenced by an XF record.
    pub fn extended_format_font(
        &self,
        format: &crate::number_format::XlsExtendedFormat,
    ) -> Option<&crate::font::XlsFont> {
        self.font(format.font_index())
    }

    /// Rich-text and phonetic properties for a shared-string index.
    ///
    /// Returns `None` for an out-of-range index and for an ordinary string
    /// without either optional BIFF8 payload.
    pub fn shared_string_properties(&self, index: u32) -> Option<&SharedStringProperties> {
        self.shared_string_properties
            .as_ref()?
            .get(index as usize)?
            .as_deref()
    }

    /// Rich-text and phonetic properties for a cell backed by `LabelSst`.
    pub fn shared_string_properties_for_cell(
        &self,
        cell: &XlsCell,
    ) -> Option<&SharedStringProperties> {
        self.shared_string_properties(cell.shared_string_index()?)
    }

    /// Non-macro internal defined names in `Lbl` record order.
    pub fn defined_names(&self) -> &[XlsDefinedName] {
        &self.defined_names
    }

    /// Every internal `Lbl`, including inert macro and procedure metadata.
    pub fn defined_name_records(&self) -> &[XlsDefinedName] {
        &self.defined_name_records
    }

    /// The built-in `_FilterDatabase` defined name scoped to a zero-based
    /// sheet index, if present.
    ///
    /// Its rendered formula describes the AutoFilter cell range.
    pub fn filter_database_name(&self, sheet_index: usize) -> Option<&XlsDefinedName> {
        self.defined_names.iter().find(|defined_name| {
            defined_name.kind
                == crate::defined_names::XlsDefinedNameKind::BuiltIn(
                    crate::defined_names::XlsBuiltInName::FilterDatabase,
                )
                && defined_name.scope == XlsNameScope::Worksheet(sheet_index)
        })
    }

    /// Case-insensitive name lookup with sheet-local-before-workbook precedence.
    /// Duplicate definitions use the last matching `Lbl` record.
    pub fn defined_name(&self, name: &str, sheet_index: Option<usize>) -> Option<&XlsDefinedName> {
        if let Some(sheet_index) = sheet_index
            && let Some(local) = self.defined_names.iter().rev().find(|defined_name| {
                defined_name.scope == XlsNameScope::Worksheet(sheet_index)
                    && names_equal(&defined_name.name, name)
            })
        {
            return Some(local);
        }
        self.defined_names.iter().rev().find(|defined_name| {
            defined_name.scope == XlsNameScope::Workbook && names_equal(&defined_name.name, name)
        })
    }

    /// Built-in print area for a worksheet, if present.
    pub fn print_area(&self, sheet_index: usize) -> Option<&XlsDefinedName> {
        self.built_in_sheet_name(sheet_index, XlsBuiltInName::PrintArea)
    }

    /// Built-in print-title rows/columns for a worksheet, if present.
    pub fn print_titles(&self, sheet_index: usize) -> Option<&XlsDefinedName> {
        self.built_in_sheet_name(sheet_index, XlsBuiltInName::PrintTitles)
    }

    fn built_in_sheet_name(
        &self,
        sheet_index: usize,
        built_in: XlsBuiltInName,
    ) -> Option<&XlsDefinedName> {
        self.defined_names.iter().rev().find(|defined_name| {
            defined_name.scope == XlsNameScope::Worksheet(sheet_index)
                && defined_name.kind == XlsDefinedNameKind::BuiltIn(built_in)
        })
    }
}

fn names_equal(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

#[cfg(test)]
mod defined_name_tests {
    use super::{XlsOpenOptions, XlsWorkbook};
    use crate::{XlsBuiltInName, XlsDefinedNameKind, XlsNameScope};
    use std::fs::File;
    use std::path::{Path, PathBuf};

    fn poi_fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    }

    fn open(name: &str) -> XlsWorkbook<File> {
        XlsWorkbook::new(File::open(poi_fixture(name)).unwrap()).unwrap()
    }

    #[test]
    fn opens_poi_named_input_with_exact_names_and_ranges() {
        let workbook = open("namedinput.xls");
        assert_eq!(workbook.defined_names().len(), 2);
        let first = workbook.defined_name("namedrangename", Some(0)).unwrap();
        assert_eq!(first.name, "NamedRangeName");
        assert_eq!(first.scope, XlsNameScope::Workbook);
        assert!(!first.hidden);
        assert!(first.formula.as_deref().unwrap().contains("$A$1:$D$10"));
        let second = workbook.defined_name("SECONDNAMEDRANGE", None).unwrap();
        assert!(second.formula.as_deref().unwrap().contains("$D$17:$G$27"));
    }

    #[test]
    fn recognizes_deleted_unicode_and_named_formula_fixtures() {
        let deleted = open("24207.xls");
        assert_eq!(deleted.defined_name("a", None).unwrap().name, "a");
        assert!(deleted.defined_name("b", None).unwrap().is_deleted());

        let unicode = open("unicodeNameRecord.xls");
        assert!(!unicode.defined_names().is_empty());

        for fixture in ["named-cell-test.xls", "named-cell-in-formula-test.xls"] {
            let workbook = open(fixture);
            assert!(!workbook.defined_names().is_empty());
        }
    }

    #[test]
    fn recognizes_poi_print_area() {
        let workbook = open("SimpleWithPrintArea.xls");
        let print_area = workbook.print_area(0).unwrap();
        assert_eq!(
            print_area.kind,
            XlsDefinedNameKind::BuiltIn(XlsBuiltInName::PrintArea)
        );
        assert!(print_area.formula.is_some());
    }

    #[test]
    fn lookup_prefers_last_local_name_then_last_global_name() {
        use crate::{XlsDefinedName, XlsDefinedNameKind};
        let mut workbook = open("namedinput.xls");
        let template = workbook.defined_names[0].clone();
        workbook.defined_names.extend([
            XlsDefinedName {
                name: "Rate".to_string(),
                scope: XlsNameScope::Workbook,
                ..template.clone()
            },
            XlsDefinedName {
                name: "RATE".to_string(),
                scope: XlsNameScope::Workbook,
                record_index: 20,
                ..template.clone()
            },
            XlsDefinedName {
                name: "rate".to_string(),
                scope: XlsNameScope::Worksheet(0),
                record_index: 21,
                kind: XlsDefinedNameKind::User,
                ..template
            },
        ]);
        assert_eq!(
            workbook.defined_name("RaTe", None).unwrap().record_index,
            20
        );
        assert_eq!(
            workbook.defined_name("RaTe", Some(0)).unwrap().record_index,
            21
        );
        assert_eq!(
            workbook.defined_name("RaTe", Some(1)).unwrap().record_index,
            20
        );
    }

    #[test]
    fn parses_names_after_xor_decryption() {
        let file = File::open(poi_fixture("xor-encryption-abc.xls")).unwrap();
        let workbook = XlsWorkbook::new_with_options(
            file,
            XlsOpenOptions {
                password: Some("abc"),
                ..XlsOpenOptions::default()
            },
        )
        .unwrap();
        assert!(workbook.defined_names().len() <= usize::from(u16::MAX));
    }
}

impl<R: Read + Seek + std::fmt::Debug + Send + Sync> litchi_core::sheet::WorkbookTrait
    for XlsWorkbook<R>
{
    fn active_worksheet(&self) -> Result<Box<dyn SheetTrait + '_>> {
        if self.worksheets.is_empty() {
            return Err(Box::new(XlsError::WorksheetNotFound(
                "No worksheets found".to_string(),
            )));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[0]))
    }

    fn worksheet_names(&self) -> &[String] {
        // Return slice reference - zero-copy!
        &self.worksheet_names
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Box<dyn SheetTrait + '_>> {
        for worksheet in &self.worksheets {
            if worksheet.name() == name {
                // Return reference instead of clone - zero-copy!
                return Ok(Box::new(worksheet));
            }
        }
        Err(Box::new(XlsError::WorksheetNotFound(name.to_string())))
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Box<dyn SheetTrait + '_>> {
        if index >= self.worksheets.len() {
            return Err(Box::new(XlsError::WorksheetNotFound(format!(
                "Index {} out of bounds",
                index
            ))));
        }
        // Return reference instead of clone - zero-copy!
        Ok(Box::new(&self.worksheets[index]))
    }

    fn worksheets(&self) -> Box<dyn WorksheetIterator<'_> + '_> {
        Box::new(XlsWorksheetIterator {
            worksheets: self.worksheets.iter().collect(),
            index: 0,
        })
    }

    fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    fn active_sheet_index(&self) -> usize {
        0 // Default to first sheet
    }

    fn is_1904_date_system(&self) -> bool {
        self.is_1904_date_system
    }
}

/// Worksheet iterator for XLS workbooks
struct XlsWorksheetIterator<'a> {
    worksheets: Vec<&'a XlsWorksheet>,
    index: usize,
}

impl<'a> WorksheetIterator<'a> for XlsWorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn SheetTrait + 'a>>> {
        if self.index >= self.worksheets.len() {
            None
        } else {
            let worksheet = self.worksheets[self.index];
            self.index += 1;
            // Return reference instead of clone - zero-copy!
            Some(Ok(Box::new(worksheet)))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::sheet::Cell;
    use std::io::Cursor;

    fn push_record(stream: &mut Vec<u8>, record_type: u16, data: &[u8]) {
        stream.extend_from_slice(&record_type.to_le_bytes());
        stream.extend_from_slice(&(data.len() as u16).to_le_bytes());
        stream.extend_from_slice(data);
    }

    fn dval_data(rule_count: u32) -> Vec<u8> {
        let mut data = vec![0; 10];
        data.extend_from_slice(&(-1i32).to_le_bytes());
        data.extend_from_slice(&rule_count.to_le_bytes());
        data
    }

    fn string_formula_data(row: u16, col: u16) -> Vec<u8> {
        let mut data = formula_data(row, col, &[]);
        data[6] = 0;
        data
    }

    fn formula_data(row: u16, col: u16, tokens: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&row.to_le_bytes());
        data.extend_from_slice(&col.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&[3, 0, 0, 0, 0, 0, 0xFF, 0xFF]);
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
        data.extend_from_slice(tokens);
        data
    }

    #[test]
    fn worksheet_expands_packed_numeric_and_blank_cells() {
        let mut mul_rk = Vec::new();
        mul_rk.extend_from_slice(&2u16.to_le_bytes());
        mul_rk.extend_from_slice(&4u16.to_le_bytes());
        mul_rk.extend_from_slice(&0u16.to_le_bytes());
        mul_rk.extend_from_slice(&((7u32 << 2) | 0x02).to_le_bytes());
        mul_rk.extend_from_slice(&1u16.to_le_bytes());
        mul_rk.extend_from_slice(&((250u32 << 2) | 0x03).to_le_bytes());
        mul_rk.extend_from_slice(&5u16.to_le_bytes());

        let mut mul_blank = Vec::new();
        mul_blank.extend_from_slice(&3u16.to_le_bytes());
        mul_blank.extend_from_slice(&1u16.to_le_bytes());
        mul_blank.extend_from_slice(&2u16.to_le_bytes());
        mul_blank.extend_from_slice(&3u16.to_le_bytes());
        mul_blank.extend_from_slice(&2u16.to_le_bytes());

        let mut stream = Vec::new();
        push_record(&mut stream, 0x00BD, &mul_rk);
        push_record(&mut stream, 0x00BE, &mul_blank);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();

        assert!(matches!(
            worksheet.get_cell(2, 4).unwrap().value(),
            litchi_core::sheet::CellValue::Float(value) if *value == 7.0
        ));
        assert!(matches!(
            worksheet.get_cell(2, 5).unwrap().value(),
            litchi_core::sheet::CellValue::Float(value) if *value == 2.5
        ));
        assert!(worksheet.get_cell(3, 1).unwrap().is_empty());
        assert!(worksheet.get_cell(3, 2).unwrap().is_empty());
    }

    #[test]
    fn worksheet_resolves_formula_value_from_following_string_record() {
        let mut string_data = Vec::new();
        string_data.extend_from_slice(&3u16.to_le_bytes());
        string_data.push(1);
        for code_unit in "文😀".encode_utf16() {
            string_data.extend_from_slice(&code_unit.to_le_bytes());
        }

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(4, 5));
        push_record(&mut stream, 0x0207, &string_data);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();
        let cell = worksheet.get_cell(4, 5).unwrap();

        assert!(cell.is_formula());
        assert!(matches!(
            cell.value(),
            litchi_core::sheet::CellValue::String(value) if value == "文😀"
        ));
    }

    #[test]
    fn worksheet_rejects_formula_missing_its_string_record() {
        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(0, 0));
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let result = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        );

        assert!(result.is_err());
    }

    #[test]
    fn worksheet_resolves_string_formula_across_intervening_array_record() {
        // MS-XLS 2.1: FORMULA = Formula [Array / Table / ShrFmla / SUB]
        // [String *Continue], so an Array record may sit between the Formula
        // and its cached String result.
        let mut array = Vec::new();
        array.extend_from_slice(&4u16.to_le_bytes()); // rwFirst
        array.extend_from_slice(&4u16.to_le_bytes()); // rwLast
        array.push(5); // colFirst
        array.push(5); // colLast
        array.extend_from_slice(&[0; 6]); // reserved/options
        array.extend_from_slice(&3u16.to_le_bytes()); // cce
        array.extend_from_slice(&[0x1E, 0x01, 0x00]); // PtgInt 1

        let mut string_data = Vec::new();
        string_data.extend_from_slice(&5u16.to_le_bytes());
        string_data.push(0);
        string_data.extend_from_slice(b"array");

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(4, 5));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x0207, &string_data);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();
        let cell = worksheet.get_cell(4, 5).unwrap();

        assert!(cell.is_formula());
        assert!(matches!(
            cell.value(),
            litchi_core::sheet::CellValue::String(value) if value == "array"
        ));
    }

    #[test]
    fn worksheet_resolves_string_formula_spanning_continue_records() {
        // The declared characters do not fit in the String record; the rest
        // arrive in a Continue record with its own option-flags byte.
        let mut string_data = Vec::new();
        string_data.extend_from_slice(&6u16.to_le_bytes());
        string_data.push(0);
        string_data.extend_from_slice(b"abc");

        let mut continues = Vec::new();
        continues.push(0); // fHighByte = 0 for this chunk
        continues.extend_from_slice(b"def");

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &string_formula_data(1, 2));
        push_record(&mut stream, 0x0207, &string_data);
        push_record(&mut stream, 0x003C, &continues);
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();

        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();
        let cell = worksheet.get_cell(1, 2).unwrap();

        assert!(cell.is_formula());
        assert!(matches!(
            cell.value(),
            litchi_core::sheet::CellValue::String(value) if value == "abcdef"
        ));
    }

    #[test]
    fn worksheet_enforces_dval_dv_ordering() {
        let mut stream = Vec::new();
        push_record(
            &mut stream,
            super::super::data_validation::DVAL_RECORD_TYPE,
            &dval_data(0),
        );
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();
        assert_eq!(
            worksheet
                .data_validation_settings()
                .unwrap()
                .declared_rule_count(),
            0
        );
        assert!(worksheet.data_validations().is_empty());

        let mut stream = Vec::new();
        push_record(
            &mut stream,
            super::super::data_validation::DVAL_RECORD_TYPE,
            &dval_data(1),
        );
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        assert!(
            XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                &mut records,
                &XlsEncoding::Utf16Le,
                "Sheet1",
                Arc::new(Vec::new()),
                Arc::new(Vec::new()),
                None,
                Arc::new(XlsFormatting::default()),
            )
            .is_err()
        );
    }

    #[test]
    fn reads_data_validation_fixtures() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let empty = XlsWorkbook::new(
            std::fs::File::open(root.join("test-data/poi/test-data/spreadsheet/dvEmpty.xls"))
                .unwrap(),
        )
        .unwrap();
        let sheet = empty.xls_worksheet(0).unwrap();
        assert_eq!(
            sheet
                .data_validation_settings()
                .unwrap()
                .declared_rule_count(),
            0
        );
        assert!(sheet.data_validations().is_empty());

        let validation = XlsWorkbook::new(
            std::fs::File::open(
                root.join("test-data/libreoffice-core/sc/qa/unit/data/xls/validation.xls"),
            )
            .unwrap(),
        )
        .unwrap();
        let sheet = validation.xls_worksheet(0).unwrap();
        assert!(!sheet.data_validations().is_empty());
        for rule in sheet.data_validations() {
            assert!(
                rule.formula1().is_some()
                    || rule.kind() == super::super::data_validation::XlsDataValidationKind::Any
            );
            assert!(!rule.ranges().is_empty());
        }
        assert!(
            sheet
                .data_validations()
                .iter()
                .flat_map(|rule| rule.ranges())
                .any(|range| {
                    range.first_row() <= 4
                        && range.last_row() >= 4
                        && range.first_column() <= 3
                        && range.last_column() >= 3
                })
        );
        assert!(
            sheet
                .data_validations()
                .iter()
                .flat_map(|rule| rule.ranges())
                .any(|range| {
                    range.first_row() <= 8
                        && range.last_row() >= 8
                        && range.first_column() <= 5
                        && range.last_column() >= 5
                })
        );
    }

    #[test]
    fn worksheet_rejects_malformed_optional_validation_records_and_continue() {
        for stream in [
            {
                let mut stream = Vec::new();
                push_record(
                    &mut stream,
                    super::super::data_validation::DVAL_RECORD_TYPE,
                    &[0; 17],
                );
                push_record(&mut stream, 0x000A, &[]);
                stream
            },
            {
                let mut stream = Vec::new();
                push_record(
                    &mut stream,
                    super::super::data_validation::DVAL_RECORD_TYPE,
                    &dval_data(1),
                );
                push_record(
                    &mut stream,
                    super::super::data_validation::DV_RECORD_TYPE,
                    &[],
                );
                push_record(&mut stream, 0x000A, &[]);
                stream
            },
            {
                let mut stream = Vec::new();
                push_record(
                    &mut stream,
                    super::super::data_validation::DVAL_RECORD_TYPE,
                    &dval_data(1),
                );
                push_record(&mut stream, 0x003C, &[0]);
                push_record(&mut stream, 0x000A, &[]);
                stream
            },
        ] {
            let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
            assert!(
                XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
                    &mut records,
                    &XlsEncoding::Utf16Le,
                    "Sheet1",
                    Arc::new(Vec::new()),
                    Arc::new(Vec::new()),
                    None,
                    Arc::new(XlsFormatting::default()),
                )
                .is_err()
            );
        }
    }

    #[test]
    fn worksheet_expands_shared_formula_relative_references() {
        let anchor = [0x01, 0, 0, 1, 0];
        let template = [
            0x4c, 0, 0, 0xff, 0xc0, // same row, previous column
            0x1e, 2, 0, 0x05, // * 2
        ];
        let mut shared = Vec::new();
        shared.extend_from_slice(&0u16.to_le_bytes());
        shared.extend_from_slice(&1u16.to_le_bytes());
        shared.extend_from_slice(&[1, 1, 0, 2]); // columns, reserved, cUse
        shared.extend_from_slice(&(template.len() as u16).to_le_bytes());
        shared.extend_from_slice(&template);

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 1, &anchor));
        push_record(&mut stream, 0x04bc, &shared);
        push_record(&mut stream, 0x0006, &formula_data(1, 1, &anchor));
        push_record(&mut stream, 0x000a, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();

        let first = worksheet.get_cell(0, 1).unwrap();
        let second = worksheet.get_cell(1, 1).unwrap();
        assert_eq!(first.formula(), Some("=(A1*2)"));
        assert_eq!(second.formula(), Some("=(A2*2)"));
        assert_eq!(first.formula_bytes(), Some(anchor.as_slice()));
        assert_eq!(second.formula_bytes(), Some(anchor.as_slice()));
    }

    #[test]
    fn worksheet_resolves_array_formula_for_every_cell() {
        let anchor = [0x01, 0, 0, 2, 0];
        let template = [0x1e, 7, 0];
        let mut array = Vec::new();
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&1u16.to_le_bytes());
        array.extend_from_slice(&[2, 2]);
        array.extend_from_slice(&0u16.to_le_bytes());
        array.extend_from_slice(&0u32.to_le_bytes());
        array.extend_from_slice(&(template.len() as u16).to_le_bytes());
        array.extend_from_slice(&template);

        let mut stream = Vec::new();
        push_record(&mut stream, 0x0006, &formula_data(0, 2, &anchor));
        push_record(&mut stream, 0x0221, &array);
        push_record(&mut stream, 0x0006, &formula_data(1, 2, &anchor));
        push_record(&mut stream, 0x000a, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        let worksheet = XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        )
        .unwrap();

        assert_eq!(worksheet.get_cell(0, 2).unwrap().formula(), Some("=7"));
        assert_eq!(worksheet.get_cell(1, 2).unwrap().formula(), Some("=7"));
    }
}
