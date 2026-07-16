//! Workbook implementation for XLS files

use crate::xls::cell::XlsCell;
use crate::xls::error::{XlsError, XlsResult};
use crate::xls::encryption::prepare_workbook_stream;
use crate::xls::defined_names::{
    DefinedNameSlot, LBL_RECORD_TYPE, XlsBuiltInName, XlsDefinedName, XlsDefinedNameKind,
    XlsNameScope,
};
use crate::xls::formula::{FormulaContext, ptg_exp_anchor, render_formula, render_shared_formula};
use crate::xls::number_format::{XlsDateSystem, XlsExtendedFormat, XlsFormatting, XlsNumberFormat};
use crate::xls::pivot_table::PivotTable;
use crate::xls::records::{
    BiffVersion, BofRecord, BoundSheetRecord, CellRecord, DimensionsRecord, FormulaValue,
    RecordIter, SharedStringProperties, SharedStringTable, XlsEncoding,
};
use crate::xls::sheet_metadata::XlsSheetMetadata;
use crate::xls::worksheet::XlsWorksheet;
use crate::xls::{autofilter, comments, conditional_format, hyperlinks, layout, merged_cells, page_setup, pivot_table, protection, utils, view};
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
    palette: crate::xls::palette::XlsPalette,
    fonts: Vec<crate::xls::font::XlsFont>,
    biff_version: BiffVersion,
    is_1904_date_system: bool,
    formula_context: FormulaContext,
    defined_names: Vec<XlsDefinedName>,
    formatting: Arc<XlsFormatting>,
    protection: protection::WorkbookProtection,
    calculation: crate::xls::calculation::XlsWorkbookCalculation,
    vba_metadata: crate::xls::vba::XlsVbaMetadata,
    environment: crate::xls::environment::XlsWorkbookEnvironment,
    workbook_view: crate::xls::workbook_view::XlsWorkbookView,
    function_groups: Option<crate::xls::function_group::XlsFunctionGroups>,
    external_links: crate::xls::external_link::XlsExternalLinks,
}

/// Options for opening a legacy XLS workbook.
#[derive(Debug, Clone, Copy, Default)]
pub struct XlsOpenOptions<'a> {
    /// Password used for BIFF8 password-to-open encryption.
    pub password: Option<&'a str>,
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
            palette: crate::xls::palette::XlsPalette::default(),
            fonts: Vec::new(),
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
            external_links: crate::xls::XlsExternalLinks::default(),
            defined_names: Vec::new(),
            formatting: Arc::new(XlsFormatting::default()),
            protection: protection::WorkbookProtection::default(),
            calculation: crate::xls::calculation::XlsWorkbookCalculation::default(),
            vba_metadata: crate::xls::vba::XlsVbaMetadata::default(),
            environment: crate::xls::environment::XlsWorkbookEnvironment::default(),
            workbook_view: crate::xls::workbook_view::XlsWorkbookView::default(),
            function_groups: None,
        };

        workbook.parse_workbook(options.password)?;
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
            palette: crate::xls::palette::XlsPalette::default(),
            fonts: Vec::new(),
            biff_version: BiffVersion::Biff8,
            is_1904_date_system: false,
            formula_context: FormulaContext::default(),
            external_links: crate::xls::XlsExternalLinks::default(),
            defined_names: Vec::new(),
            formatting: Arc::new(XlsFormatting::default()),
            protection: protection::WorkbookProtection::default(),
            calculation: crate::xls::calculation::XlsWorkbookCalculation::default(),
            vba_metadata: crate::xls::vba::XlsVbaMetadata::default(),
            environment: crate::xls::environment::XlsWorkbookEnvironment::default(),
            workbook_view: crate::xls::workbook_view::XlsWorkbookView::default(),
            function_groups: None,
        };

        workbook.parse_workbook(options.password)?;
        Ok(workbook)
    }

    /// Parse the workbook stream
    fn parse_workbook(&mut self, password: Option<&str>) -> XlsResult<()> {
        // Find and read the Workbook stream
        let workbook_data = self
            .ole_file
            .open_stream(&["Workbook"])
            .or_else(|_| self.ole_file.open_stream(&["Book"]))?;
        let workbook_data = prepare_workbook_stream(workbook_data, password)?;

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
            &mut bound_sheets,
            &mut strings,
            &mut string_properties,
            &mut defined_name_slots,
        )?;

        // Use Arc for zero-copy sharing across worksheets
        self.shared_strings = Some(Arc::new(strings));
        self.shared_string_properties = Some(Arc::new(string_properties));
        let all_sheet_names = bound_sheets.iter().map(|sheet| sheet.name.clone()).collect::<Vec<_>>();
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
        self.formula_context
            .set_sheet_names(all_sheet_names);
        self.formula_context.set_defined_names(
            defined_name_slots
                .iter()
                .map(DefinedNameSlot::symbol)
                .collect(),
        );
        self.defined_names = defined_name_slots
            .into_iter()
            .map(|slot| slot.into_public(bound_sheets.len(), &self.formula_context))
            .collect::<XlsResult<Vec<_>>>()?
            .into_iter()
            .flatten()
            .collect();

        // Parse worksheets from positions in the workbook stream
        for (sheet_index, bound_sheet) in bound_sheets.iter().enumerate() {
            if bound_sheet.sheet_type != crate::xls::records::SheetType::WorkSheet {
                continue;
            }
            match self.parse_worksheet_from_position(bound_sheet, &encoding, &mut record_iter) {
                Ok(worksheet) => {
                    let worksheet_index = self.worksheets.len();
                    self.worksheet_names.push(bound_sheet.name.clone());
                    self.worksheets.push(worksheet);
                    self.sheets[sheet_index].set_parsed_worksheet_index(worksheet_index);
                },
                Err(_e) => {
                    // Failed to parse worksheet, continue with next
                },
            }
        }

        Ok(())
    }

    /// Parse workbook globals (SST, bound sheets, etc.)
    fn parse_workbook_globals<Reader: Read + Seek>(
        &mut self,
        record_iter: &mut RecordIter<Reader>,
        encoding: &mut XlsEncoding,
        bound_sheets: &mut Vec<BoundSheetRecord>,
        strings: &mut Vec<String>,
        string_properties: &mut Vec<Option<Box<SharedStringProperties>>>,
        defined_name_slots: &mut Vec<DefinedNameSlot>,
    ) -> XlsResult<()> {
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

        self.formatting = Arc::new(XlsFormatting::parse_globals(&records)?);
        self.is_1904_date_system = self.formatting.date_system() == XlsDateSystem::Excel1904;

        let mut palette_seen = false;
        let mut protection_collector = protection::WorkbookProtectionCollector::new();
        let mut calculation_collector =
            crate::xls::calculation::WorkbookCalculationCollector::new();
        let mut vba_collector = crate::xls::vba::WorkbookVbaCollector::new();
        let mut environment_collector = crate::xls::environment::EnvironmentCollector::new();
        let mut workbook_view_collector = crate::xls::workbook_view::WorkbookViewCollector::new();
        let mut function_group_collector = crate::xls::function_group::FunctionGroupCollector::new();
        let mut external_link_collector = crate::xls::external_link::ExternalLinkCollector::new();
        let mut i = 0;
        while i < records.len() {
            let record = &records[i];
            protection_collector.feed_record(record.header.record_type, &record.data)?;
            calculation_collector.feed_record(record.header.record_type, &record.data)?;
            vba_collector.feed_record(record.header.record_type, &record.data)?;
            environment_collector.feed_record(record.header.record_type, &record.data)?;
            workbook_view_collector.feed_record(record.header.record_type, &record.data)?;
            function_group_collector.feed_record(record.header.record_type, &record.data)?;
            external_link_collector.feed_record(record.header.record_type, &record.data)?;

            match record.header.record_type {
                0x0031 => {
                    let index = crate::xls::font::logical_font_index(self.fonts.len())?;
                    self.fonts.push(crate::xls::font::XlsFont::parse_record(
                        index,
                        &record.data,
                    )?);
                },
                0x0092 => {
                    self.palette = crate::xls::palette::XlsPalette::parse_unique_record(
                        &record.data,
                        &mut palette_seen,
                    )?;
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
                    defined_name_slots.push(DefinedNameSlot::parse(
                        &record.data,
                        record_index,
                    )?);
                },
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
                    crate::xls::font::validate_font_table(&self.fonts)?;
                    self.protection = protection_collector.finish()?;
                    self.calculation = calculation_collector.finish();
                    self.vba_metadata = vba_collector.finish();
                    self.environment = environment_collector.finish()?;
                    self.workbook_view = workbook_view_collector.finish(bound_sheets.len())?;
                    self.function_groups = function_group_collector.finish()?;
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
        shared_string_properties: Arc<
            Vec<Option<Box<crate::xls::records::SharedStringProperties>>>,
        >,
        formula_context: Option<&FormulaContext>,
        formatting: Arc<XlsFormatting>,
    ) -> XlsResult<XlsWorksheet> {
        let mut worksheet = XlsWorksheet::with_shared_string_properties(
            name.to_string(),
            shared_strings,
            shared_string_properties,
        );
        worksheet.set_formatting(formatting.clone());

        // Accumulator for pivot table records: we collect SX* records in order
        // and assemble complete PivotTable structs when SXVIEW boundaries are hit.
        let mut current_pivot: Option<PivotTable> = None;

        let mut comment_collector = comments::CommentCollector::new();
        let mut hyperlink_collector = hyperlinks::HyperlinkCollector::new();
        let mut layout_collector = layout::LayoutCollector::new();
        let mut sheet_layout_collector = crate::xls::sheet_layout::SheetLayoutCollector::new();
        let mut view_collector = view::ViewCollector::new();
        let mut page_setup_collector = page_setup::PageSetupCollector::new();
        let mut protection_collector = protection::SheetProtectionCollector::new();
        let mut conditional_format_collector = conditional_format::ConditionalFormatCollector::new();
        let mut calculation_collector =
            crate::xls::calculation::WorksheetCalculationCollector::new();
        let mut scenario_collector = crate::xls::scenario::ScenarioCollector::new();
        let mut vba_collector = crate::xls::vba::WorksheetVbaCollector::new();
        let mut pending_string_formula: Option<CellRecord> = None;
        let mut shared_formulas = HashMap::<(u16, u16), SharedFormulaTemplate>::new();
        let mut remaining_data_validations: Option<usize> = None;

        for record_result in record_iter.by_ref() {
            let record = record_result?;
            sheet_layout_collector.feed_record(record.header.record_type, &record.data)?;
            comment_collector.feed_record(record.header.record_type, &record.data)?;
            hyperlink_collector.feed_record(record.header.record_type, &record.data)?;
            layout_collector.feed_record(record.header.record_type, &record.data, &formatting)?;
            view_collector.feed_record(record.header.record_type, &record.data)?;
            page_setup_collector.feed_record(record.header.record_type, &record.data)?;
            protection_collector.feed_record(record.header.record_type, &record.data)?;
            conditional_format_collector.feed_record(record.header.record_type, &record.data)?;
            calculation_collector.feed_record(record.header.record_type, &record.data)?;
            scenario_collector.feed_record(record.header.record_type, &record.data)?;
            vba_collector.feed_record(record.header.record_type, &record.data)?;

            if matches!(remaining_data_validations, Some(1..))
                && record.header.record_type != super::data_validation::DV_RECORD_TYPE
            {
                return Err(XlsError::InvalidRecord {
                    record_type: record.header.record_type,
                    message: "DVAL must be followed immediately by its declared DV records".to_string(),
                });
            }

            if let Some(mut formula) = pending_string_formula.take() {
                if record.header.record_type != 0x0207 {
                    return Err(XlsError::InvalidRecord {
                        record_type: record.header.record_type,
                        message: "String-valued Formula must be followed by a String record"
                            .to_string(),
                    });
                }
                let text = utils::parse_string_record(&record.data, encoding)?;
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
                    {
                        if let Some(anchor) = ptg_exp_anchor(formula) {
                            if let Some(rendered) = shared_formulas
                                .get(&anchor)
                                .and_then(|template| template.render(formula_context, *row, *col))
                            {
                                cell.set_rendered_formula(Some(rendered));
                            }
                        }
                    }
                    worksheet.add_cell(cell);
                }
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
                    worksheet.add_data_validation(super::data_validation::parse_dv(&record.data)?);
                    *remaining -= 1;
                }
                0x000A => { // EOF - End of worksheet
                    // Flush any in-progress pivot table
                    if let Some(pt) = current_pivot.take() {
                        worksheet.add_pivot_table(pt);
                    }
                    *worksheet.protection_mut() = protection_collector.finish()?;
                    break;
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
                        {
                            if let Some(anchor) = ptg_exp_anchor(formula) {
                                if let Some(rendered) = shared_formulas
                                    .get(&anchor)
                                    .and_then(|template| {
                                        template.render(formula_context, *row, *col)
                                    })
                                {
                                    cell.set_rendered_formula(Some(rendered));
                                }
                            }
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

                // --- Pivot table records ---
                rt if rt == pivot_table::SXVIEW_TYPE => {
                    // New SXVIEW starts a new pivot table; flush previous if any
                    if let Some(pt) = current_pivot.take() {
                        worksheet.add_pivot_table(pt);
                    }
                    if let Ok(view) = pivot_table::parse_sxview(&record.data) {
                        current_pivot = Some(PivotTable::new(view));
                    }
                }
                rt if rt == pivot_table::SXVD_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(field) = pivot_table::parse_sxvd(&record.data)
                    {
                        pt.fields.push(field);
                    }
                }
                rt if rt == pivot_table::SXVI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(item) = pivot_table::parse_sxvi(&record.data)
                    {
                        pt.items.push(item);
                    }
                }
                rt if rt == pivot_table::SXDI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(di) = pivot_table::parse_sxdi(&record.data)
                    {
                        pt.data_items.push(di);
                    }
                }
                rt if rt == pivot_table::SXVS_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(src) = pivot_table::parse_sxvs(&record.data)
                    {
                        pt.source_type = src;
                    }
                }
                rt if rt == pivot_table::SXPI_TYPE => {
                    if let Some(ref mut pt) = current_pivot
                        && let Ok(entries) = pivot_table::parse_sxpi(&record.data)
                    {
                        pt.page_entries.extend(entries);
                    }
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

        worksheet.set_comments(comment_collector.finish()?);
        worksheet.set_hyperlinks(hyperlink_collector.finish());
        let (row_layouts, column_layouts) = layout_collector.finish();
        worksheet.set_layouts(row_layouts, column_layouts);
        worksheet.set_sheet_layout(sheet_layout_collector.finish());
        worksheet.set_worksheet_views(view_collector.finish()?);
        worksheet.set_page_setup(page_setup_collector.finish()?);
        worksheet.set_conditional_formattings(conditional_format_collector.finish()?);
        worksheet.set_calculation(calculation_collector.finish()?);
        worksheet.set_scenario_manager(scenario_collector.finish()?);
        worksheet.set_vba_code_name(vba_collector.finish());

        Ok(worksheet)
    }

    /// Workbook window state and stable sheet identifiers.
    pub fn workbook_view(&self) -> &crate::xls::workbook_view::XlsWorkbookView {
        &self.workbook_view
    }

    /// Built-in and custom function categories, when the FNGROUPS collection exists.
    pub fn function_groups(&self) -> Option<&crate::xls::function_group::XlsFunctionGroups> {
        self.function_groups.as_ref()
    }

    /// Inert supporting-book links and cached external cell values.
    pub fn external_links(&self) -> &crate::xls::external_link::XlsExternalLinks {
        &self.external_links
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
        self.sheets.iter().find(|sheet| sheet.name().eq_ignore_ascii_case(name))
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

    pub fn calculation(&self) -> &crate::xls::calculation::XlsWorkbookCalculation {
        &self.calculation
    }

    pub fn vba_metadata(&self) -> crate::xls::vba::XlsVbaMetadata {
        let mut metadata = self.vba_metadata.clone();
        metadata.set_project_storage_present(self.ole_file.list_streams().iter().any(|path| {
            path.first().is_some_and(|name| name.eq_ignore_ascii_case("_VBA_PROJECT_CUR"))
        }));
        metadata
    }

    pub fn environment(&self) -> &crate::xls::environment::XlsWorkbookEnvironment {
        &self.environment
    }

    pub fn number_formats(&self) -> &[XlsNumberFormat] {
        self.formatting.number_formats()
    }

    pub fn extended_formats(&self) -> &[XlsExtendedFormat] {
        self.formatting.extended_formats()
    }

    /// Resolves an XF's effective property families through its parent StyleXF.
    pub fn effective_extended_format(
        &self,
        index: u16,
    ) -> Option<crate::xls::number_format::XlsEffectiveExtendedFormat<'_>> {
        self.formatting.effective_extended_format(index)
    }

    /// Workbook color palette, using BIFF8 defaults when no `Palette` record exists.
    pub fn palette(&self) -> &crate::xls::palette::XlsPalette {
        &self.palette
    }

    /// Font records in physical workbook order.
    pub fn fonts(&self) -> &[crate::xls::font::XlsFont] {
        &self.fonts
    }

    /// Resolve a BIFF8 logical font index. Index 4 is reserved and returns `None`.
    pub fn font(&self, index: u16) -> Option<&crate::xls::font::XlsFont> {
        self.fonts.iter().find(|font| font.index() == index)
    }

    /// Resolve a font's color through the workbook palette.
    pub fn font_color(&self, index: u16) -> Option<crate::xls::palette::XlsColor> {
        self.font(index)
            .and_then(|font| self.palette.color(font.color_index()))
    }

    /// Resolves the global Font record referenced by an XF record.
    pub fn extended_format_font(
        &self,
        format: &crate::xls::number_format::XlsExtendedFormat,
    ) -> Option<&crate::xls::font::XlsFont> {
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

    /// Case-insensitive name lookup with sheet-local-before-workbook precedence.
    /// Duplicate definitions use the last matching `Lbl` record.
    pub fn defined_name(
        &self,
        name: &str,
        sheet_index: Option<usize>,
    ) -> Option<&XlsDefinedName> {
        if let Some(sheet_index) = sheet_index {
            if let Some(local) = self.defined_names.iter().rev().find(|defined_name| {
                defined_name.scope == XlsNameScope::Worksheet(sheet_index)
                    && names_equal(&defined_name.name, name)
            }) {
                return Some(local);
            }
        }
        self.defined_names.iter().rev().find(|defined_name| {
            defined_name.scope == XlsNameScope::Workbook
                && names_equal(&defined_name.name, name)
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
    use crate::xls::{XlsBuiltInName, XlsDefinedNameKind, XlsNameScope};
    use std::fs::File;
    use std::path::{Path, PathBuf};

    fn poi_fixture(name: &str) -> PathBuf {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../3rdparty/poi/test-data/spreadsheet")
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
        use crate::xls::{XlsDefinedName, XlsDefinedNameKind};
        let mut workbook = open("namedinput.xls");
        let template = workbook.defined_names[0].clone();
        workbook.defined_names.extend([
            XlsDefinedName { name: "Rate".to_string(), scope: XlsNameScope::Workbook, ..template.clone() },
            XlsDefinedName { name: "RATE".to_string(), scope: XlsNameScope::Workbook, record_index: 20, ..template.clone() },
            XlsDefinedName { name: "rate".to_string(), scope: XlsNameScope::Worksheet(0), record_index: 21, kind: XlsDefinedNameKind::User, ..template },
        ]);
        assert_eq!(workbook.defined_name("RaTe", None).unwrap().record_index, 20);
        assert_eq!(workbook.defined_name("RaTe", Some(0)).unwrap().record_index, 21);
        assert_eq!(workbook.defined_name("RaTe", Some(1)).unwrap().record_index, 20);
    }

    #[test]
    fn parses_names_after_xor_decryption() {
        let file = File::open(poi_fixture("xor-encryption-abc.xls")).unwrap();
        let workbook = XlsWorkbook::new_with_options(
            file,
            XlsOpenOptions {
                password: Some("abc"),
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
    fn worksheet_enforces_dval_dv_ordering() {
        let mut stream = Vec::new();
        push_record(&mut stream, super::super::data_validation::DVAL_RECORD_TYPE, &dval_data(0));
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
        ).unwrap();
        assert_eq!(worksheet.data_validation_settings().unwrap().declared_rule_count(), 0);
        assert!(worksheet.data_validations().is_empty());

        let mut stream = Vec::new();
        push_record(&mut stream, super::super::data_validation::DVAL_RECORD_TYPE, &dval_data(1));
        push_record(&mut stream, 0x000A, &[]);
        let mut records = RecordIter::new(Cursor::new(stream)).unwrap();
        assert!(XlsWorkbook::<Cursor<Vec<u8>>>::parse_worksheet_records(
            &mut records,
            &XlsEncoding::Utf16Le,
            "Sheet1",
            Arc::new(Vec::new()),
            Arc::new(Vec::new()),
            None,
            Arc::new(XlsFormatting::default()),
        ).is_err());
    }

    #[test]
    fn reads_data_validation_fixtures() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let empty = XlsWorkbook::new(std::fs::File::open(
            root.join("3rdparty/poi/test-data/spreadsheet/dvEmpty.xls"),
        ).unwrap()).unwrap();
        let sheet = empty.xls_worksheet(0).unwrap();
        assert_eq!(sheet.data_validation_settings().unwrap().declared_rule_count(), 0);
        assert!(sheet.data_validations().is_empty());

        let validation = XlsWorkbook::new(std::fs::File::open(
            root.join("3rdparty/libreoffice-core/sc/qa/unit/data/xls/validation.xls"),
        ).unwrap()).unwrap();
        let sheet = validation.xls_worksheet(0).unwrap();
        assert!(!sheet.data_validations().is_empty());
        for rule in sheet.data_validations() {
            assert!(rule.formula1().is_some() || rule.kind() == super::super::data_validation::XlsDataValidationKind::Any);
            assert!(!rule.ranges().is_empty());
        }
        assert!(sheet.data_validations().iter().flat_map(|rule| rule.ranges()).any(|range| {
            range.first_row() <= 4 && range.last_row() >= 4
                && range.first_column() <= 3 && range.last_column() >= 3
        }));
        assert!(sheet.data_validations().iter().flat_map(|rule| rule.ranges()).any(|range| {
            range.first_row() <= 8 && range.last_row() >= 8
                && range.first_column() <= 5 && range.last_column() >= 5
        }));
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
