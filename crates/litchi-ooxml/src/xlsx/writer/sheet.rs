use crate::drawings::blip::write_a_blip_embed_rid_num;
use crate::drawings::ext::write_a16_creation_id_extlst;
use crate::drawings::fill::write_a_stretch_fill_rect;
use crate::xlsx::cell::Cell;
use crate::xlsx::data_validation::{
    DataValidationCollection, DataValidationConformance,
    validate_data_validation_collections, write_data_validation_core,
    write_data_validation_extensions,
};
use crate::xlsx::sort::{SortCondition, SortState};
use crate::xlsx::sparkline::{SparklineGroup, write_sparkline_groups_ext};
use crate::xlsx::sheet_protection::{
    ProtectionPasswordVerifier, WorksheetProtection, WorksheetProtectionConformance,
    WorksheetProtectionMetadata, write_worksheet_protection_core,
    write_worksheet_protection_extensions,
};
use crate::xlsx::table::Table;
use crate::xlsx::views::{SheetPane, SheetSelection, SheetView};
use crate::xlsx::writer::shape::{
    ShapeEmitter, XlsxConnectionShapeSpec, XlsxGroupSpec, XlsxShapeSpec,
};
/// Writer module for creating and modifying Excel worksheets.
use litchi_core::sheet::{CellValue, Result as SheetResult};
use litchi_core::{id::generate_guid_braced, xml::escape::escape_xml};
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

// Import shared formatting types
pub use super::super::format::{
    CellBorder, CellBorderLineStyle, CellBorderSide, CellFill, CellFillPatternType, CellFont,
    CellFormat, DataValidation, DataValidationOperator, DataValidationType,
};

// Import chart types
pub use super::super::chart::{
    ChartAnchor, ChartExternalDataPart, ChartExternalDataTarget, ChartRelationship,
    ChartRelationshipTarget, ChartUserShapesPart, ChartUserShapesRelationship,
    ChartUserShapesRelationshipTarget, WorksheetChart,
};

// Import from other writer modules
use super::strings::MutableSharedStrings;

/// Freeze panes configuration.
///
/// Freezes rows and columns in place while scrolling.
#[derive(Debug, Clone)]
pub struct FreezePanes {
    /// Number of columns to freeze from the left
    pub freeze_cols: u32,
    /// Number of rows to freeze from the top
    pub freeze_rows: u32,
}

/// A standard SpreadsheetML defined name.
///
/// `reference` is retained as inert formula text. Reading or writing this type never
/// evaluates the formula or invokes a macro procedure.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct NamedRange {
    /// Name of the definition (for example, `TaxRate` or `_xlnm.Print_Area`).
    pub name: String,
    /// Inert name-formula text (for example, `Sheet1!$A$1:$B$10` or `42`).
    pub reference: String,
    /// Optional user comment.
    pub comment: Option<String>,
    /// Zero-based workbook sheet index for a local name; `None` means workbook scope.
    pub local_sheet_id: Option<u32>,
    pub custom_menu: Option<String>,
    pub description: Option<String>,
    pub help: Option<String>,
    pub status_bar: Option<String>,
    pub shortcut_key: Option<String>,
    pub hidden: bool,
    pub function: bool,
    pub vb_procedure: bool,
    pub xlm: bool,
    pub function_group_id: Option<u32>,
    pub publish_to_server: bool,
    pub workbook_parameter: bool,
}

impl NamedRange {
    /// Return the standardized built-in represented by this name, if any.
    pub fn built_in(&self) -> Option<DefinedNameBuiltIn> {
        DefinedNameBuiltIn::from_name(&self.name)
    }

    pub fn is_sheet_scoped(&self) -> bool {
        self.local_sheet_id.is_some()
    }
}

/// Built-in names reserved by SpreadsheetML (ECMA-376, `definedName`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum DefinedNameBuiltIn {
    ConsolidateArea,
    AutoOpen,
    AutoClose,
    Extract,
    Database,
    Criteria,
    PrintArea,
    PrintTitles,
    Recorder,
    DataForm,
    AutoActivate,
    AutoDeactivate,
    SheetTitle,
    FilterDatabase,
}

impl DefinedNameBuiltIn {
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ConsolidateArea => "_xlnm.Consolidate_Area",
            Self::AutoOpen => "_xlnm.Auto_Open",
            Self::AutoClose => "_xlnm.Auto_Close",
            Self::Extract => "_xlnm.Extract",
            Self::Database => "_xlnm.Database",
            Self::Criteria => "_xlnm.Criteria",
            Self::PrintArea => "_xlnm.Print_Area",
            Self::PrintTitles => "_xlnm.Print_Titles",
            Self::Recorder => "_xlnm.Recorder",
            Self::DataForm => "_xlnm.Data_Form",
            Self::AutoActivate => "_xlnm.Auto_Activate",
            Self::AutoDeactivate => "_xlnm.Auto_Deactivate",
            Self::SheetTitle => "_xlnm.Sheet_Title",
            Self::FilterDatabase => "_xlnm._FilterDatabase",
        }
    }

    pub fn from_name(name: &str) -> Option<Self> {
        const BUILT_INS: [DefinedNameBuiltIn; 14] = [
            DefinedNameBuiltIn::ConsolidateArea,
            DefinedNameBuiltIn::AutoOpen,
            DefinedNameBuiltIn::AutoClose,
            DefinedNameBuiltIn::Extract,
            DefinedNameBuiltIn::Database,
            DefinedNameBuiltIn::Criteria,
            DefinedNameBuiltIn::PrintArea,
            DefinedNameBuiltIn::PrintTitles,
            DefinedNameBuiltIn::Recorder,
            DefinedNameBuiltIn::DataForm,
            DefinedNameBuiltIn::AutoActivate,
            DefinedNameBuiltIn::AutoDeactivate,
            DefinedNameBuiltIn::SheetTitle,
            DefinedNameBuiltIn::FilterDatabase,
        ];
        BUILT_INS
            .into_iter()
            .find(|built_in| name.eq_ignore_ascii_case(built_in.as_str()))
    }
}

/// Page setup configuration.
#[derive(Debug, Clone)]
pub struct PageSetup {
    /// Orientation: "portrait" or "landscape"
    pub orientation: String,
    /// Paper size (e.g., 1 = Letter, 9 = A4)
    pub paper_size: u32,
    /// Scale percentage (10-400)
    pub scale: Option<u32>,
    /// Fit to page width
    pub fit_to_width: Option<u32>,
    /// Fit to page height
    pub fit_to_height: Option<u32>,
}

/// Optional authoring state for `sheetPr/pageSetUpPr`.
///
/// `None` preserves attribute absence. This is distinct from an explicitly
/// written `false`, which is useful when interoperating with producers that
/// persist the effective defaults.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct PageSetupProperties {
    fit_to_page: Option<bool>,
    automatic_page_breaks: Option<bool>,
}

impl PageSetupProperties {
    /// The explicitly authored `fitToPage` value, if present.
    pub fn fit_to_page(&self) -> Option<bool> {
        self.fit_to_page
    }

    /// The explicitly authored `autoPageBreaks` value, if present.
    pub fn automatic_page_breaks(&self) -> Option<bool> {
        self.automatic_page_breaks
    }

    fn is_empty(&self) -> bool {
        self.fit_to_page.is_none() && self.automatic_page_breaks.is_none()
    }
}

/// Header and footer configuration.
#[derive(Debug, Clone, Default)]
pub struct HeaderFooter {
    /// Header text (left, center, right)
    pub header_left: Option<String>,
    pub header_center: Option<String>,
    pub header_right: Option<String>,
    /// Footer text (left, center, right)
    pub footer_left: Option<String>,
    pub footer_center: Option<String>,
    pub footer_right: Option<String>,
}

/// Manual page break definition.
#[derive(Debug, Clone)]
pub struct PageBreak {
    /// Row or column index (1-based) where the break occurs.
    pub id: u32,
    /// Minimum span index (0-based).
    pub min: u32,
    /// Maximum span index (0-based).
    pub max: u32,
    /// Whether the break is manual.
    pub manual: bool,
}

/// Auto-filter configuration.
#[derive(Debug, Clone)]
pub struct AutoFilter {
    /// Range for the auto-filter (e.g., "A1:D10")
    pub range: String,
    /// Optional sort state associated with the auto-filter.
    pub sort_state: Option<SortState>,
}

/// Hyperlink information for a cell.
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Target URL or internal reference
    pub target: String,
    /// Display text (tooltip)
    pub display: Option<String>,
}

/// Cell comment information.
#[derive(Debug, Clone)]
pub struct CellComment {
    /// Row (1-based)
    pub row: u32,
    /// Column (1-based)
    pub col: u32,
    /// Author of the comment
    pub author: String,
    /// Comment text
    pub text: String,
}

/// Image information for embedding in a worksheet.
#[derive(Debug, Clone)]
pub struct Image {
    /// Image data (raw bytes)
    pub data: Vec<u8>,
    /// Image format extension (e.g., "png", "jpeg", "jpg", "gif")
    pub format: String,
    /// Position: (from_row, from_col, to_row, to_col)
    pub position: (u32, u32, u32, u32),
    /// Optional description/alt text
    pub description: Option<String>,
}

#[derive(Default)]
pub(crate) struct WorksheetPartRelationships<'a> {
    pub(crate) hyperlinks: Option<&'a HashMap<String, String>>,
    pub(crate) vml_drawing: Option<&'a str>,
    pub(crate) pivot_tables: Option<&'a [String]>,
    pub(crate) tables: Option<&'a [String]>,
    pub(crate) drawing: Option<&'a str>,
}

/// Rich text run for a cell.
#[derive(Debug, Clone)]
pub struct RichTextRun {
    /// Text content for this run
    pub text: String,
    /// Font name (optional)
    pub font_name: Option<String>,
    /// Font size in points (optional)
    pub font_size: Option<f64>,
    /// Bold
    pub bold: bool,
    /// Italic
    pub italic: bool,
    /// Underline
    pub underline: bool,
    /// Text color as ARGB hex (e.g., "FF0000FF")
    pub color: Option<String>,
}

/// Conditional formatting rule.
#[derive(Debug, Clone)]
pub struct ConditionalFormat {
    /// Range (e.g., "A1:B10")
    pub range: String,
    /// Rule type
    pub rule_type: ConditionalFormatType,
    /// Priority (lower = higher priority)
    pub priority: u32,
    /// Format to apply (optional - can use built-in formats)
    pub format: Option<CellFormat>,
}

/// Conditional formatting rule types.
#[derive(Debug, Clone)]
pub enum ConditionalFormatType {
    /// Cell value comparison (e.g., greater than, less than)
    CellIs {
        /// Operator (e.g., "greaterThan", "lessThan", "equal")
        operator: String,
        /// Formula to compare against
        formula: String,
    },
    /// Color scale (2 or 3 color gradient)
    ColorScale {
        /// Minimum color (RGB hex)
        min_color: String,
        /// Maximum color (RGB hex)
        max_color: String,
        /// Optional mid color (RGB hex)
        mid_color: Option<String>,
    },
    /// Data bar
    DataBar {
        /// Color (RGB hex)
        color: String,
        /// Show value alongside bar
        show_value: bool,
    },
    /// Icon set
    IconSet {
        /// Icon set name (e.g., "3Arrows", "3TrafficLights")
        icon_set: String,
        /// Show values
        show_value: bool,
    },
    /// Formula-based
    Expression {
        /// Formula that returns TRUE/FALSE
        formula: String,
    },
}

/// A mutable worksheet for writing and modification.
///
/// Provides methods to set cell values, formulas, and formatting.
#[derive(Debug)]
pub struct MutableWorksheet {
    /// Worksheet name
    name: String,
    /// Sheet ID
    sheet_id: u32,
    /// Cell data (row, col) -> value
    cells: HashMap<(u32, u32), CellValue>,
    /// Cell formatting
    cell_formats: HashMap<(u32, u32), CellFormat>,
    /// Merged cell ranges (start_row, start_col, end_row, end_col)
    merged_cells: Vec<(u32, u32, u32, u32)>,
    /// Charts in this worksheet
    charts: Vec<WorksheetChart>,
    /// Data validation rules
    validations: Vec<DataValidation>,
    /// Complete typed core and Office 2010 validation collections.
    typed_data_validations: Option<Vec<DataValidationCollection>>,
    /// Column widths (col -> width in characters)
    column_widths: HashMap<u32, f64>,
    /// Hidden columns
    hidden_columns: std::collections::HashSet<u32>,
    /// Row heights (row -> height in points)
    row_heights: HashMap<u32, f64>,
    /// Hidden rows
    hidden_rows: std::collections::HashSet<u32>,
    /// Freeze panes configuration
    freeze_panes: Option<FreezePanes>,
    /// Whether the worksheet is hidden
    hidden: bool,
    /// Visibility state: "visible", "hidden", or "veryHidden"
    visibility: String,
    /// Whether this worksheet is active
    is_active: bool,
    /// Tab color (RGB hex, e.g., "FF0000" for red)
    tab_color: Option<String>,
    /// Page setup configuration
    page_setup: Option<PageSetup>,
    /// Optional flags serialized as `sheetPr/pageSetUpPr`.
    page_setup_properties: Option<PageSetupProperties>,
    /// Print area
    print_area: Option<String>,
    /// Header and footer
    header_footer: Option<HeaderFooter>,
    /// Repeating rows for printing (e.g., "1:2")
    repeating_rows: Option<String>,
    /// Repeating columns for printing (e.g., "A:B")
    repeating_columns: Option<String>,
    /// Auto-filter configuration
    auto_filter: Option<AutoFilter>,
    /// Sheet view settings
    sheet_view: Option<SheetView>,
    /// Manual row page breaks
    row_breaks: Vec<PageBreak>,
    /// Manual column page breaks
    col_breaks: Vec<PageBreak>,
    /// Sheet protection configuration
    protection: WorksheetProtectionMetadata,
    /// Hyperlinks by cell reference
    hyperlinks: Vec<Hyperlink>,
    /// Cell comments
    comments: Vec<CellComment>,
    /// Conditional formatting rules
    conditional_formats: Vec<ConditionalFormat>,
    /// Images embedded in the worksheet
    images: Vec<Image>,
    /// DrawingML shapes and text boxes authored for the worksheet
    shapes: Vec<XlsxShapeSpec>,
    /// DrawingML shape groups authored for the worksheet
    groups: Vec<XlsxGroupSpec>,
    /// DrawingML connection shapes authored for the worksheet
    connections: Vec<XlsxConnectionShapeSpec>,
    /// Row outline levels (row -> level)
    row_outline_levels: HashMap<u32, u8>,
    /// Column outline levels (col -> level)
    column_outline_levels: HashMap<u32, u8>,
    /// Rich text runs per cell (row, col)
    rich_text_cells: HashMap<(u32, u32), Vec<RichTextRun>>,
    sparkline_groups: Vec<SparklineGroup>,
    /// Tables in this worksheet
    tables: Vec<Table>,
    /// Threaded comments for this worksheet
    threaded_comments: Vec<crate::xlsx::ThreadedComment>,
    /// Whether the worksheet has been modified
    modified: bool,
}

impl MutableWorksheet {
    /// Create a new empty worksheet.
    pub fn new(name: String, sheet_id: u32) -> Self {
        Self {
            name,
            sheet_id,
            cells: HashMap::new(),
            cell_formats: HashMap::new(),
            merged_cells: Vec::new(),
            charts: Vec::new(),
            validations: Vec::new(),
            typed_data_validations: None,
            column_widths: HashMap::new(),
            hidden_columns: std::collections::HashSet::new(),
            row_heights: HashMap::new(),
            hidden_rows: std::collections::HashSet::new(),
            freeze_panes: None,
            hidden: false,
            visibility: "visible".to_string(),
            is_active: false,
            tab_color: None,
            page_setup: None,
            page_setup_properties: None,
            print_area: None,
            header_footer: None,
            repeating_rows: None,
            repeating_columns: None,
            auto_filter: None,
            sheet_view: None,
            row_breaks: Vec::new(),
            col_breaks: Vec::new(),
            protection: WorksheetProtectionMetadata::default(),
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            conditional_formats: Vec::new(),
            images: Vec::new(),
            shapes: Vec::new(),
            groups: Vec::new(),
            connections: Vec::new(),
            row_outline_levels: HashMap::new(),
            column_outline_levels: HashMap::new(),
            rich_text_cells: HashMap::new(),
            sparkline_groups: Vec::new(),
            tables: Vec::new(),
            threaded_comments: Vec::new(),
            modified: false,
        }
    }

    pub fn sparkline_groups(&self) -> &[SparklineGroup] {
        &self.sparkline_groups
    }

    pub fn add_sparkline_group(&mut self, group: SparklineGroup) {
        fn parse_a1_cell_ref(s: &str) -> Option<(u32, u32)> {
            let s = s.trim();
            if s.is_empty() {
                return None;
            }

            let mut col: u32 = 0;
            let mut saw_letter = false;
            let mut i = 0usize;
            for (idx, ch) in s.char_indices() {
                if ch.is_ascii_alphabetic() {
                    saw_letter = true;
                    col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
                    i = idx + ch.len_utf8();
                } else {
                    break;
                }
            }
            if !saw_letter {
                return None;
            }
            if col == 0 {
                return None;
            }
            let col0 = col - 1;

            let row_str = s.get(i..)?;
            let row1: u32 = row_str.parse().ok()?;
            if row1 == 0 {
                return None;
            }
            Some((row1 - 1, col0))
        }

        for sp in &group.sparklines {
            if let Some((r, c)) = parse_a1_cell_ref(&sp.location) {
                self.cells.remove(&(r, c));
                self.cell_formats.remove(&(r, c));
                self.rich_text_cells.remove(&(r, c));
            }
        }

        self.sparkline_groups.push(group);
        self.modified = true;
    }

    /// Get the worksheet name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Set the worksheet name.
    pub fn set_name(&mut self, name: String) {
        self.name = name;
        self.modified = true;
    }

    /// Get the sheet ID.
    pub fn sheet_id(&self) -> u32 {
        self.sheet_id
    }

    /// Set a cell value.
    ///
    /// # Arguments
    /// * `row` - 1-based row number (1 = first row)
    /// * `col` - 1-based column number (1 = column A)
    pub fn set_cell_value<V: Into<CellValue>>(&mut self, row: u32, col: u32, value: V) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.cells.insert((row - 1, col - 1), value.into());
        self.modified = true;
    }

    /// Set a rich text cell composed of multiple formatted runs.
    ///
    /// This also sets the plain string value for the cell by concatenating all
    /// runs, so generic consumers see the combined text via the unified
    /// `CellValue::String` API.
    pub fn set_rich_text_cell(&mut self, row: u32, col: u32, runs: Vec<RichTextRun>) {
        let row_idx = row.saturating_sub(1);
        let col_idx = col.saturating_sub(1);

        // Build plain text for the cell value
        let mut plain = String::new();
        for run in &runs {
            plain.push_str(&run.text);
        }

        self.cells
            .insert((row_idx, col_idx), CellValue::String(plain));
        self.rich_text_cells.insert((row_idx, col_idx), runs);
        self.modified = true;
    }

    /// Get rich text runs for a cell, if present.
    pub fn rich_text_cell(&self, row: u32, col: u32) -> Option<&[RichTextRun]> {
        let row_idx = row.saturating_sub(1);
        let col_idx = col.saturating_sub(1);
        self.rich_text_cells
            .get(&(row_idx, col_idx))
            .map(|runs| runs.as_slice())
    }

    /// Clear rich text formatting for a cell (leaves plain value intact).
    pub fn clear_rich_text_cell(&mut self, row: u32, col: u32) {
        let row_idx = row.saturating_sub(1);
        let col_idx = col.saturating_sub(1);
        if self.rich_text_cells.remove(&(row_idx, col_idx)).is_some() {
            self.modified = true;
        }
    }

    /// Set a cell formula.
    ///
    /// # Arguments
    /// * `row` - 1-based row number
    /// * `col` - 1-based column number
    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.cells.insert(
            (row - 1, col - 1),
            CellValue::Formula {
                formula: formula.to_string(),
                cached_value: None,
                is_array: false,
                array_range: None,
            },
        );
        self.modified = true;
    }

    /// Set a cell formula with a cached result value.
    pub fn set_cell_formula_with_cache<V: Into<CellValue>>(
        &mut self,
        row: u32,
        col: u32,
        formula: &str,
        cached_value: V,
    ) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.cells.insert(
            (row - 1, col - 1),
            CellValue::Formula {
                formula: formula.to_string(),
                cached_value: Some(Box::new(cached_value.into())),
                is_array: false,
                array_range: None,
            },
        );
        self.modified = true;
    }

    /// Set an array formula over a rectangular range.
    ///
    /// # Arguments
    /// * `start_row`, `start_col` - Top-left cell (1-based)
    /// * `end_row`, `end_col` - Bottom-right cell (1-based)
    /// * `formula` - Formula expression (without leading '=')
    pub fn set_array_formula(
        &mut self,
        start_row: u32,
        start_col: u32,
        end_row: u32,
        end_col: u32,
        formula: &str,
    ) {
        if start_row == 0 || start_col == 0 || end_row < start_row || end_col < start_col {
            return;
        }

        let top_row = start_row.saturating_sub(1);
        let left_col = start_col.saturating_sub(1);

        let start_ref = format!("{}{}", Self::column_to_letters(start_col), start_row);
        let end_ref = format!("{}{}", Self::column_to_letters(end_col), end_row);
        let range_ref = format!("{}:{}", start_ref, end_ref);

        self.cells.insert(
            (top_row, left_col),
            CellValue::Formula {
                formula: formula.to_string(),
                cached_value: None,
                is_array: true,
                array_range: Some(range_ref),
            },
        );

        // Ensure all cells in the target range exist so that Excel treats
        // them as part of the array region when it recalculates.
        for r in start_row..=end_row {
            for c in start_col..=end_col {
                let r_idx = r.saturating_sub(1);
                let c_idx = c.saturating_sub(1);
                if r_idx == top_row && c_idx == left_col {
                    continue;
                }
                self.cells.entry((r_idx, c_idx)).or_insert(CellValue::Empty);
            }
        }

        self.modified = true;
    }

    /// Set cell formatting.
    pub fn set_cell_format(&mut self, row: u32, col: u32, format: CellFormat) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.cell_formats.insert((row - 1, col - 1), format);
        self.modified = true;
    }

    /// Merge cells in a rectangular range.
    ///
    /// # Arguments
    /// * `start_row` - 1-based starting row
    /// * `start_col` - 1-based starting column
    /// * `end_row` - 1-based ending row
    /// * `end_col` - 1-based ending column
    pub fn merge_cells(&mut self, start_row: u32, start_col: u32, end_row: u32, end_col: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.merged_cells
            .push((start_row - 1, start_col - 1, end_row - 1, end_col - 1));
        self.modified = true;
    }

    /// Add a chart to the worksheet.
    ///
    /// # Arguments
    ///
    /// * `chart` - The WorksheetChart to add with its positioning
    ///
    /// # Example
    ///
    /// ```rust,ignore
    /// let chart = WorksheetChart::bar_chart(
    ///     "Sales Data",
    ///     "Sheet1!$A$2:$A$10",
    ///     "Sheet1!$B$2:$B$10",
    ///     ChartAnchor::new(1, 1, 7, 14),
    /// )?;
    /// worksheet.add_chart(chart);
    /// ```
    pub fn add_chart(&mut self, chart: WorksheetChart) {
        self.charts.push(chart);
        self.modified = true;
    }

    /// Add a pivot chart bound to a pivot table by name.
    ///
    /// The chart (built with the `WorksheetChart` constructors) is converted
    /// into a pivot chart: its pivot source is set to `pivot_table_name` and
    /// every series without an extension list receives the default
    /// all-visible drop-zone options. The name is validated against the
    /// workbook's pivot tables and normalized to its sheet-qualified form at
    /// save time, so the saved package is valid by construction.
    ///
    /// # Arguments
    ///
    /// * `chart` - The WorksheetChart to add with its positioning
    /// * `pivot_table_name` - Name of the pivot table the chart binds to,
    ///   optionally sheet-qualified (for example `Sheet1!PivotTable1`)
    pub fn add_pivot_chart(
        &mut self,
        chart: WorksheetChart,
        pivot_table_name: &str,
    ) -> SheetResult<()> {
        if pivot_table_name.is_empty() {
            return Err("pivot table name cannot be empty".into());
        }
        self.charts.push(chart.into_pivot_chart(pivot_table_name));
        self.modified = true;
        Ok(())
    }

    /// Get the charts in this worksheet.
    pub fn charts(&self) -> &[WorksheetChart] {
        &self.charts
    }

    /// Get a mutable reference to the charts in this worksheet.
    pub fn charts_mut(&mut self) -> &mut Vec<WorksheetChart> {
        self.modified = true;
        &mut self.charts
    }

    /// Remove all charts from the worksheet.
    pub fn clear_charts(&mut self) {
        self.charts.clear();
        self.modified = true;
    }

    /// Remove a chart at the specified index.
    pub fn remove_chart(&mut self, index: usize) -> Option<WorksheetChart> {
        if index < self.charts.len() {
            self.modified = true;
            Some(self.charts.remove(index))
        } else {
            None
        }
    }

    /// Add data validation to a cell range.
    #[allow(clippy::too_many_arguments)]
    pub fn add_data_validation(
        &mut self,
        range: &str,
        validation_type: DataValidationType,
        show_input_message: bool,
        input_title: Option<&str>,
        input_message: Option<&str>,
        show_error_alert: bool,
        error_title: Option<&str>,
        error_message: Option<&str>,
    ) {
        self.typed_data_validations = None;
        self.validations.push(DataValidation {
            range: range.to_string(),
            validation_type,
            show_input_message,
            input_title: input_title.map(|s| s.to_string()),
            input_message: input_message.map(|s| s.to_string()),
            show_error_alert,
            error_title: error_title.map(|s| s.to_string()),
            error_message: error_message.map(|s| s.to_string()),
        });
        self.modified = true;
    }

    /// Replace validations with complete typed core/x14 collections.
    pub fn set_data_validation_collections(
        &mut self,
        collections: Vec<DataValidationCollection>,
    ) -> SheetResult<()> {
        validate_data_validation_collections(&collections)?;
        self.validations.clear();
        self.typed_data_validations = Some(collections);
        self.modified = true;
        Ok(())
    }

    pub fn data_validation_collections(&self) -> Option<&[DataValidationCollection]> {
        self.typed_data_validations.as_deref()
    }

    pub fn clear_data_validations(&mut self) {
        self.validations.clear();
        self.typed_data_validations = Some(Vec::new());
        self.modified = true;
    }

    /// Get a cell value.
    pub fn cell_value(&self, row: u32, col: u32) -> Option<&CellValue> {
        if row == 0 || col == 0 {
            return None;
        }
        self.cells.get(&(row - 1, col - 1))
    }

    /// Clear a cell.
    pub fn clear_cell(&mut self, row: u32, col: u32) {
        self.cells.remove(&(row, col));
        self.modified = true;
    }

    /// Clear all cells in the worksheet.
    pub fn clear_all(&mut self) {
        self.cells.clear();
        self.modified = true;
    }

    /// Get the number of non-empty cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    /// Set column width in characters (Excel default is 8.43).
    ///
    /// # Arguments
    /// * `col` - 1-based column number (1 = column A)
    pub fn set_column_width(&mut self, col: u32, width: f64) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.column_widths.insert(col - 1, width);
        self.modified = true;
    }

    /// Hide a column.
    ///
    /// # Arguments
    /// * `col` - 1-based column number
    pub fn hide_column(&mut self, col: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.hidden_columns.insert(col - 1);
        self.modified = true;
    }

    /// Show a previously hidden column.
    ///
    /// # Arguments
    /// * `col` - 1-based column number
    pub fn show_column(&mut self, col: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.hidden_columns.remove(&(col - 1));
        self.modified = true;
    }

    /// Set row height in points (Excel default is 15).
    ///
    /// # Arguments
    /// * `row` - 1-based row number (1 = first row)
    pub fn set_row_height(&mut self, row: u32, height: f64) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.row_heights.insert(row - 1, height);
        self.modified = true;
    }

    /// Hide a row.
    ///
    /// # Arguments
    /// * `row` - 1-based row number
    pub fn hide_row(&mut self, row: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.hidden_rows.insert(row - 1);
        self.modified = true;
    }

    /// Show a previously hidden row.
    ///
    /// # Arguments
    /// * `row` - 1-based row number
    pub fn show_row(&mut self, row: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        self.hidden_rows.remove(&(row - 1));
        self.modified = true;
    }

    /// Freeze panes at the specified position.
    pub fn freeze_panes(&mut self, freeze_rows: u32, freeze_cols: u32) {
        if freeze_rows > 0 || freeze_cols > 0 {
            self.freeze_panes = Some(FreezePanes {
                freeze_rows,
                freeze_cols,
            });
            self.modified = true;
        }
    }

    /// Remove freeze panes.
    pub fn unfreeze_panes(&mut self) {
        self.freeze_panes = None;
        self.modified = true;
    }

    /// Check if the worksheet has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified
    }

    // ===== Worksheet Visibility and State =====

    /// Set whether the worksheet is hidden.
    pub fn set_hidden(&mut self, hidden: bool) {
        self.hidden = hidden;
        self.visibility = if hidden {
            "hidden".to_string()
        } else {
            "visible".to_string()
        };
        self.modified = true;
    }

    /// Check if the worksheet is hidden.
    pub fn is_hidden(&self) -> bool {
        self.hidden
    }

    /// Set worksheet visibility state.
    ///
    /// # Arguments
    /// * `visibility` - Visibility state: "visible", "hidden", or "veryHidden"
    pub fn set_visibility(&mut self, visibility: &str) {
        self.visibility = visibility.to_string();
        self.hidden = visibility != "visible";
        self.modified = true;
    }

    /// Get worksheet visibility state.
    pub fn visibility(&self) -> &str {
        &self.visibility
    }

    /// Set whether this worksheet is the active sheet.
    pub fn set_active(&mut self, active: bool) {
        self.is_active = active;
        self.modified = true;
    }

    /// Check if this worksheet is the active sheet.
    pub fn is_active(&self) -> bool {
        self.is_active
    }

    /// Set the tab color for the worksheet.
    ///
    /// # Arguments
    /// * `color` - RGB hex color (e.g., "FF0000" for red)
    pub fn set_tab_color(&mut self, color: &str) {
        self.tab_color = Some(color.to_string());
        self.modified = true;
    }

    /// Remove the tab color from the worksheet.
    pub fn remove_tab_color(&mut self) {
        self.tab_color = None;
        self.modified = true;
    }

    /// Get the tab color for the worksheet.
    pub fn tab_color(&self) -> Option<&str> {
        self.tab_color.as_deref()
    }

    // ===== Hyperlinks =====

    /// Set a hyperlink for a cell.
    ///
    /// # Arguments
    /// * `row` - Row index (1-based)
    /// * `col` - Column index (1-based)
    /// * `url` - The URL or internal reference
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_cell_value(1, 1, "Click here");
    /// ws.set_hyperlink(1, 1, "https://example.com", None);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_hyperlink(&mut self, row: u32, col: u32, url: &str, display: Option<&str>) {
        // Convert row/col to cell reference (API is 1-based, no conversion needed for display)
        let cell_ref = format!("{}{}", Self::column_to_letters(col), row);

        // Remove existing hyperlink for this cell if any
        self.hyperlinks.retain(|h| h.cell_ref != cell_ref);

        // Add new hyperlink
        self.hyperlinks.push(Hyperlink {
            cell_ref,
            target: url.to_string(),
            display: display.map(|s| s.to_string()),
        });

        self.modified = true;
    }

    /// Remove a hyperlink from a cell.
    ///
    /// # Arguments
    /// * `row` - Row index (1-based)
    /// * `col` - Column index (1-based)
    pub fn remove_hyperlink(&mut self, row: u32, col: u32) {
        let cell_ref = format!("{}{}", Self::column_to_letters(col), row);
        let initial_len = self.hyperlinks.len();
        self.hyperlinks.retain(|h| h.cell_ref != cell_ref);

        if self.hyperlinks.len() < initial_len {
            self.modified = true;
        }
    }

    /// Get all hyperlinks in the worksheet.
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    // ===== Comments =====

    /// Add a comment to a cell.
    ///
    /// # Arguments
    /// * `row` - Row index (1-based)
    /// * `col` - Column index (1-based)
    /// * `text` - The comment text
    /// * `author` - Author of the comment
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_cell_value(1, 1, 42);
    /// ws.set_cell_comment(1, 1, "This is important!", "John Doe");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_cell_comment(&mut self, row: u32, col: u32, text: &str, author: &str) {
        // Remove existing comment for this cell if any (row/col are 1-based, stored as-is)
        self.comments.retain(|c| c.row != row || c.col != col);

        // Add new comment (store 1-based for hyperlinks/comments since they reference by cell name)
        self.comments.push(CellComment {
            row,
            col,
            author: author.to_string(),
            text: text.to_string(),
        });

        self.modified = true;
    }

    /// Remove a comment from a cell.
    ///
    /// # Arguments
    /// * `row` - Row index (1-based)
    /// * `col` - Column index (1-based)
    pub fn remove_comment(&mut self, row: u32, col: u32) {
        let initial_len = self.comments.len();
        self.comments.retain(|c| c.row != row || c.col != col);

        if self.comments.len() < initial_len {
            self.modified = true;
        }
    }

    /// Get all comments in the worksheet.
    pub fn comments(&self) -> &[CellComment] {
        &self.comments
    }

    /// Add a threaded comment to a cell.
    ///
    /// Threaded comments support conversation threads, @mentions, and timestamps.
    ///
    /// # Arguments
    /// * `comment` - The threaded comment to add
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, ThreadedComment};
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// let comment = ThreadedComment {
    ///     cell_ref: Some("A1".to_string()),
    ///     id: "{22222222-2222-2222-2222-222222222222}".to_string(),
    ///     person_id: "{11111111-1111-1111-1111-111111111111}".to_string(),
    ///     text: Some("This is a threaded comment".to_string()),
    ///     ..Default::default()
    /// };
    /// ws.add_threaded_comment(comment);
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn add_threaded_comment(&mut self, comment: crate::xlsx::ThreadedComment) {
        self.threaded_comments.push(comment);
        self.modified = true;
    }

    /// Get all threaded comments in the worksheet.
    pub fn threaded_comments(&self) -> &[crate::xlsx::ThreadedComment] {
        &self.threaded_comments
    }

    /// Remove all threaded comments for a specific cell.
    ///
    /// # Arguments
    /// * `cell_ref` - Cell reference (e.g., "A1")
    pub fn remove_threaded_comments(&mut self, cell_ref: &str) {
        let initial_len = self.threaded_comments.len();
        self.threaded_comments
            .retain(|c| c.cell_ref.as_deref() != Some(cell_ref));
        if self.threaded_comments.len() < initial_len {
            self.modified = true;
        }
    }

    /// Generate comments.xml content.
    ///
    /// This generates the XML for all cell comments in the worksheet.
    pub fn generate_comments_xml(&self) -> SheetResult<Option<String>> {
        if self.comments.is_empty() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(1024);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        // Build unique list of authors
        let mut authors: Vec<String> = Vec::new();
        for comment in &self.comments {
            if !authors.contains(&comment.author) {
                authors.push(comment.author.clone());
            }
        }

        // Write authors list
        xml.push_str("<authors>");
        for author in &authors {
            write!(xml, "<author>{}</author>", escape_xml(author))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</authors>");

        // Write comment list
        xml.push_str("<commentList>");
        for comment in &self.comments {
            let cell_ref = format!("{}{}", Self::column_to_letters(comment.col), comment.row);
            let author_id = authors
                .iter()
                .position(|a| a == &comment.author)
                .unwrap_or(0);

            write!(
                xml,
                r#"<comment ref="{}" authorId="{}">"#,
                escape_xml(&cell_ref),
                author_id
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            xml.push_str("<text>");
            // Add author run
            xml.push_str("<r>");
            xml.push_str(r#"<rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>"#);
            write!(xml, "<t>{}:</t>", escape_xml(&comment.author))
                .map_err(|e| format!("XML write error: {}", e))?;
            xml.push_str("</r>");

            // Add text run
            xml.push_str("<r>");
            xml.push_str(r#"<rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>"#);
            write!(
                xml,
                "<t xml:space=\"preserve\">{}</t>",
                escape_xml(&comment.text)
            )
            .map_err(|e| format!("XML write error: {}", e))?;
            xml.push_str("</r>");
            xml.push_str("</text>");

            xml.push_str("</comment>");
        }
        xml.push_str("</commentList>");

        xml.push_str("</comments>");
        Ok(Some(xml))
    }

    /// Generate VML drawing XML for comment indicators.
    ///
    /// This generates the VML drawing file that displays comment indicators in cells.
    pub fn generate_vml_drawing_xml(&self) -> SheetResult<Option<String>> {
        // Generate VML for both legacy comments and threaded comments
        if self.comments.is_empty() && self.threaded_comments.is_empty() {
            return Ok(None);
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str(r#"<xml xmlns:v="urn:schemas-microsoft-com:vml""#);
        xml.push_str(r#" xmlns:o="urn:schemas-microsoft-com:office:office""#);
        xml.push_str(r#" xmlns:x="urn:schemas-microsoft-com:office:excel">"#);

        // VML shape type for comments
        xml.push_str(
            r#"<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>"#,
        );
        xml.push_str(
            r#"<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">"#,
        );
        xml.push_str(
            r#"<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>"#,
        );
        xml.push_str(r#"</v:shapetype>"#);

        // Helper function to parse cell reference like "A1" -> (row, col) 0-indexed
        fn parse_cell_ref(s: &str) -> Option<(u32, u32)> {
            let s = s.trim();
            if s.is_empty() {
                return None;
            }
            let mut col: u32 = 0;
            let mut i = 0usize;
            for (idx, ch) in s.char_indices() {
                if ch.is_ascii_alphabetic() {
                    col = col * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1);
                    i = idx + ch.len_utf8();
                } else {
                    break;
                }
            }
            if col == 0 {
                return None;
            }
            let row_str = s.get(i..)?;
            let row1: u32 = row_str.parse().ok()?;
            if row1 == 0 {
                return None;
            }
            Some((row1 - 1, col - 1))
        }

        let mut shape_idx = 0;

        // Generate a shape for each legacy comment
        for comment in self.comments.iter() {
            let shape_id = shape_idx + 1024; // Start from 1024 to avoid conflicts
            shape_idx += 1;

            write!(
                xml,
                "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\" style=\"position:absolute;",
                shape_id
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            // Position the comment indicator (top-right of cell)
            let margin_left = 48.0 + (comment.col as f64 - 1.0) * 63.0;
            let margin_top = 12.0 + (comment.row as f64 - 1.0) * 15.75;

            write!(
                xml,
                "margin-left:{:.0}pt;margin-top:{:.0}pt;width:108pt;height:59.25pt;z-index:{};visibility:hidden\" ",
                margin_left, margin_top, shape_idx
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            write!(xml, "fillcolor=\"ffffe1\" o:insetmode=\"auto\">")
                .map_err(|e| format!("XML write error: {}", e))?;

            xml.push_str(r#"<v:fill color2="ffffe1"/>"#);
            xml.push_str(r#"<v:shadow on="t" color="black" obscured="t"/>"#);
            xml.push_str(r#"<v:path o:connecttype="none"/>"#);
            xml.push_str(r#"<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>"#);
            write!(
                xml,
                r#"<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{}, 15, {}, 0, {}, 15, {}, 4</x:Anchor>"#,
                comment.col + 1, comment.row - 1, comment.col + 3, comment.row + 2
            )
            .map_err(|e| format!("XML write error: {}", e))?;
            xml.push_str("<x:AutoFill>False</x:AutoFill>");
            write!(xml, "<x:Row>{}</x:Row>", comment.row - 1)
                .map_err(|e| format!("XML write error: {}", e))?;
            write!(xml, "<x:Column>{}</x:Column>", comment.col - 1)
                .map_err(|e| format!("XML write error: {}", e))?;
            xml.push_str("</x:ClientData>");
            xml.push_str("</v:shape>");
        }

        // Generate a shape for each threaded comment
        for comment in self.threaded_comments.iter() {
            if let Some(ref cell_ref) = comment.cell_ref {
                // Parse cell reference to get row/col
                if let Some((row, col)) = parse_cell_ref(cell_ref) {
                    let shape_id = shape_idx + 1024;
                    shape_idx += 1;

                    write!(
                        xml,
                        "<v:shape id=\"_x0000_s{}\" type=\"#_x0000_t202\" style=\"position:absolute;",
                        shape_id
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    // Position the comment indicator
                    let margin_left = 48.0 + (col as f64) * 63.0;
                    let margin_top = 12.0 + (row as f64) * 15.75;

                    write!(
                        xml,
                        "margin-left:{:.0}pt;margin-top:{:.0}pt;width:108pt;height:59.25pt;z-index:{};visibility:hidden\" ",
                        margin_left, margin_top, shape_idx
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    write!(
                        xml,
                        "fillcolor=\"#ECFAD4\" strokecolor=\"#edeaa1\" o:insetmode=\"auto\">"
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    xml.push_str("<v:fill color2=\"#BEFF82\" type=\"gradient\" angle=\"-180\"><o:fill type=\"gradientUnscaled\" v:ext=\"view\"/></v:fill>");
                    xml.push_str(r#"<v:shadow on="t" obscured="t"/>"#);
                    xml.push_str(r#"<v:path o:connecttype="none"/>"#);
                    xml.push_str(r#"<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>"#);
                    write!(
                        xml,
                        r#"<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>{}, 15, {}, 0, {}, 15, {}, 4</x:Anchor>"#,
                        col + 1, row, col + 3, row + 3
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;
                    xml.push_str("<x:AutoFill>False</x:AutoFill>");
                    write!(xml, "<x:Row>{}</x:Row>", row)
                        .map_err(|e| format!("XML write error: {}", e))?;
                    write!(xml, "<x:Column>{}</x:Column>", col)
                        .map_err(|e| format!("XML write error: {}", e))?;
                    xml.push_str("</x:ClientData>");
                    xml.push_str("</v:shape>");
                }
            }
        }

        xml.push_str("</xml>");
        Ok(Some(xml))
    }

    // ===== Images =====

    /// Add an image to the worksheet.
    ///
    /// # Arguments
    /// * `image_data` - Raw image bytes
    /// * `format` - Image format ("png", "jpeg", "jpg", "gif", "bmp")
    /// * `from_row` - Starting row (1-based)
    /// * `from_col` - Starting column (1-based)
    /// * `to_row` - Ending row (1-based)
    /// * `to_col` - Ending column (1-based)
    /// * `description` - Optional alt text/description
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    /// use std::fs;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// let image_data = fs::read("logo.png")?;
    /// ws.add_image(image_data, "png", 1, 1, 5, 5, Some("Company Logo"))?;
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    #[allow(clippy::too_many_arguments)]
    pub fn add_image(
        &mut self,
        image_data: Vec<u8>,
        format: &str,
        from_row: u32,
        from_col: u32,
        to_row: u32,
        to_col: u32,
        description: Option<&str>,
    ) -> SheetResult<()> {
        if image_data.is_empty() {
            return Err("image data cannot be empty".into());
        }
        if from_row == 0 || from_col == 0 || to_row == 0 || to_col == 0 {
            return Err("image anchor coordinates must be one-based and positive".into());
        }
        if to_row < from_row || to_col < from_col {
            return Err("image anchor cannot be descending".into());
        }
        if to_row > 1_048_576 || to_col > 16_384 {
            return Err("image anchor exceeds worksheet bounds".into());
        }
        let format = format.to_ascii_lowercase();
        if !matches!(
            format.as_str(),
            "png" | "jpg" | "jpeg" | "gif" | "bmp" | "svg"
        ) {
            return Err(format!("unsupported worksheet image format '{format}'").into());
        }
        // Convert from 1-based (API) to 0-based (internal storage)
        self.images.push(Image {
            data: image_data,
            format,
            position: (from_row - 1, from_col - 1, to_row - 1, to_col - 1),
            description: description.map(|s| s.to_string()),
        });
        self.modified = true;
        Ok(())
    }

    /// Get all images in the worksheet.
    pub fn images(&self) -> &[Image] {
        &self.images
    }

    /// Add a DrawingML shape or text box to the worksheet's drawing part.
    ///
    /// The spec is validated against worksheet bounds and module limits;
    /// shapes coexist with images and charts in the same drawing part and are
    /// readable back through `Workbook::shapes_on_sheet`.
    pub fn add_shape(&mut self, shape: XlsxShapeSpec) -> SheetResult<()> {
        shape.validate(self.shapes.len())?;
        self.shapes.push(shape);
        self.modified = true;
        Ok(())
    }

    /// Add a plain-text text box with the given preset geometry and anchor.
    ///
    /// Convenience wrapper around [`Self::add_shape`]; `text` is split into
    /// paragraphs on `\n`.
    pub fn add_text_box(
        &mut self,
        name: &str,
        anchor: crate::xlsx::XlsxShapeAnchor,
        preset: crate::xlsx::XlsxShapePreset,
        text: &str,
    ) -> SheetResult<()> {
        self.add_shape(XlsxShapeSpec::text_box(name, anchor, preset, text))
    }

    /// Get all authored shapes in the worksheet.
    pub fn shapes(&self) -> &[XlsxShapeSpec] {
        &self.shapes
    }

    /// Remove the authored shape at `index`, returning it.
    pub fn remove_shape(&mut self, index: usize) -> SheetResult<XlsxShapeSpec> {
        if index >= self.shapes.len() {
            return Err(format!(
                "shape index {index} out of bounds ({} shapes)",
                self.shapes.len()
            )
            .into());
        }
        self.modified = true;
        Ok(self.shapes.remove(index))
    }

    /// Total number of authored top-level drawing objects (shapes, groups,
    /// and connection shapes), used to enforce the per-worksheet limit.
    fn drawing_object_count(&self) -> usize {
        self.shapes.len() + self.groups.len() + self.connections.len()
    }

    /// Add a shape group (`xdr:grpSp`) to the worksheet's drawing part.
    ///
    /// Group children may be text boxes, plain shapes, nested groups, and
    /// connection shapes; anchors of children are ignored inside the group.
    pub fn add_group(&mut self, group: XlsxGroupSpec) -> SheetResult<()> {
        group.validate(self.drawing_object_count())?;
        self.groups.push(group);
        self.modified = true;
        Ok(())
    }

    /// Get all authored shape groups in the worksheet.
    pub fn groups(&self) -> &[XlsxGroupSpec] {
        &self.groups
    }

    /// Remove the authored group at `index`, returning it.
    pub fn remove_group(&mut self, index: usize) -> SheetResult<XlsxGroupSpec> {
        if index >= self.groups.len() {
            return Err(format!(
                "group index {index} out of bounds ({} groups)",
                self.groups.len()
            )
            .into());
        }
        self.modified = true;
        Ok(self.groups.remove(index))
    }

    /// Add a connection shape (`xdr:cxnSp`) to the worksheet's drawing part.
    ///
    /// The start/end sites reference other authored shapes by name; the
    /// references are resolved to drawing object IDs when the drawing XML is
    /// generated, and unknown names fail serialization.
    pub fn add_connection(&mut self, connection: XlsxConnectionShapeSpec) -> SheetResult<()> {
        connection.validate(self.drawing_object_count())?;
        self.connections.push(connection);
        self.modified = true;
        Ok(())
    }

    /// Get all authored connection shapes in the worksheet.
    pub fn connections(&self) -> &[XlsxConnectionShapeSpec] {
        &self.connections
    }

    /// Remove the authored connection shape at `index`, returning it.
    pub fn remove_connection(&mut self, index: usize) -> SheetResult<XlsxConnectionShapeSpec> {
        if index >= self.connections.len() {
            return Err(format!(
                "connection index {index} out of bounds ({} connections)",
                self.connections.len()
            )
            .into());
        }
        self.modified = true;
        Ok(self.connections.remove(index))
    }

    /// Generate drawing XML for images and charts.
    ///
    /// This generates the xl/drawings/drawing{N}.xml file content.
    pub fn generate_drawing_xml(&self) -> SheetResult<Option<String>> {
        if self.images.is_empty()
            && self.charts.is_empty()
            && self.shapes.is_empty()
            && self.groups.is_empty()
            && self.connections.is_empty()
        {
            return Ok(None);
        }
        let mut xml = String::with_capacity(4096);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#);
        xml.push_str(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);

        for (idx, image) in self.images.iter().enumerate() {
            let (from_row, from_col, to_row, to_col) = image.position;

            // Two-cell anchor (positioned relative to cells)
            xml.push_str("<xdr:twoCellAnchor>");

            // From position (position is already 0-based from add_image)
            write!(
                xml,
                "<xdr:from><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>",
                from_col,
                from_row
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            // To position (position is already 0-based from add_image)
            write!(
                xml,
                "<xdr:to><xdr:col>{}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>",
                to_col,
                to_row
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            // Picture element
            write!(
                xml,
                r#"<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{}" name="Picture {}""#,
                idx + 1,
                idx + 1
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            if let Some(description) = &image.description {
                write!(xml, r#" descr="{}""#, escape_xml(description))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push('>');

            if image.description.is_some() {
                let creation_id = generate_guid_braced();
                write_a16_creation_id_extlst(&mut xml, &creation_id)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</xdr:cNvPr>");

            xml.push_str(
                r#"<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>"#,
            );
            xml.push_str("<xdr:blipFill>");
            write_a_blip_embed_rid_num(&mut xml, (idx + 1) as u32, true)
                .map_err(|e| format!("XML write error: {}", e))?;
            write_a_stretch_fill_rect(&mut xml);
            xml.push_str("</xdr:blipFill>");

            // Shape properties
            xml.push_str(
                r#"<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>"#,
            );
            xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>"#);

            xml.push_str("<xdr:clientData/></xdr:twoCellAnchor>");
        }

        let image_count = self.images.len();
        crate::xlsx::chart::write_worksheet_chart_anchors(
            &mut xml,
            &self.charts,
            image_count,
            image_count,
        )?;

        // Add authored shapes, groups, and connectors after pictures and
        // charts; they consume no drawing-part relationships, so the
        // positional rId scheme for images and charts is unaffected. Object
        // IDs continue the same sequence, allocated by the emitter, which
        // also resolves connection-site name references.
        let shape_id_offset = (self.images.len() + self.charts.len()) as u32;
        let mut emitter = ShapeEmitter::new(shape_id_offset + 1);
        for shape in &self.shapes {
            emitter
                .write_anchored_shape(&mut xml, shape)
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        for group in &self.groups {
            emitter
                .write_anchored_group(&mut xml, group)
                .map_err(|e| format!("XML write error: {e}"))?;
        }
        for connection in &self.connections {
            emitter
                .write_anchored_connection(&mut xml, connection)
                .map_err(|e| format!("XML write error: {e}"))?;
        }

        xml.push_str("</xdr:wsDr>");
        Ok(Some(xml))
    }

    // ===== Conditional Formatting =====

    /// Add conditional formatting to a range.
    ///
    /// # Arguments
    /// * `range` - Cell range (e.g., "A1:B10")
    /// * `rule_type` - The formatting rule type
    /// * `priority` - Priority (lower = higher priority, typically start at 1)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, writer::ConditionalFormatType};
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// // Highlight cells greater than 100
    /// ws.add_conditional_formatting(
    ///     "A1:A10",
    ///     ConditionalFormatType::CellIs {
    ///         operator: "greaterThan".to_string(),
    ///         formula: "100".to_string(),
    ///     },
    ///     1,
    ///     None,
    /// );
    ///
    /// // Color scale
    /// ws.add_conditional_formatting(
    ///     "B1:B10",
    ///     ConditionalFormatType::ColorScale {
    ///         min_color: "FF0000".to_string(),
    ///         max_color: "00FF00".to_string(),
    ///         mid_color: None,
    ///     },
    ///     2,
    ///     None,
    /// );
    ///
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn add_conditional_formatting(
        &mut self,
        range: &str,
        rule_type: ConditionalFormatType,
        priority: u32,
        format: Option<CellFormat>,
    ) {
        self.conditional_formats.push(ConditionalFormat {
            range: range.to_string(),
            rule_type,
            priority,
            format,
        });

        self.modified = true;
    }

    /// Remove all conditional formatting from the worksheet.
    pub fn clear_conditional_formatting(&mut self) {
        if !self.conditional_formats.is_empty() {
            self.conditional_formats.clear();
            self.modified = true;
        }
    }

    /// Get all conditional formatting rules.
    pub fn conditional_formatting(&self) -> &[ConditionalFormat] {
        &self.conditional_formats
    }

    // ===== Page Setup =====

    /// Set the worksheet's fit-to-page mode.
    ///
    /// This controls `sheetPr/pageSetUpPr@fitToPage`. It does not alter the
    /// independently configured `pageSetup@fitToWidth` or `@fitToHeight`
    /// dimensions.
    pub fn set_fit_to_page(&mut self, enabled: bool) {
        self.page_setup_properties
            .get_or_insert_with(PageSetupProperties::default)
            .fit_to_page = Some(enabled);
        self.modified = true;
    }

    /// Remove the authored `fitToPage` attribute without discarding dimensions.
    pub fn clear_fit_to_page(&mut self) {
        if let Some(properties) = self.page_setup_properties.as_mut() {
            properties.fit_to_page = None;
            if properties.is_empty() {
                self.page_setup_properties = None;
            }
            self.modified = true;
        }
    }

    /// Get the explicitly authored fit-to-page mode.
    pub fn fit_to_page(&self) -> Option<bool> {
        self.page_setup_properties
            .as_ref()
            .and_then(PageSetupProperties::fit_to_page)
    }

    /// Set whether page breaks are calculated automatically.
    pub fn set_automatic_page_breaks(&mut self, enabled: bool) {
        self.page_setup_properties
            .get_or_insert_with(PageSetupProperties::default)
            .automatic_page_breaks = Some(enabled);
        self.modified = true;
    }

    /// Remove the authored `autoPageBreaks` attribute.
    pub fn clear_automatic_page_breaks(&mut self) {
        if let Some(properties) = self.page_setup_properties.as_mut() {
            properties.automatic_page_breaks = None;
            if properties.is_empty() {
                self.page_setup_properties = None;
            }
            self.modified = true;
        }
    }

    /// Get the explicitly authored automatic-page-break setting.
    pub fn automatic_page_breaks(&self) -> Option<bool> {
        self.page_setup_properties
            .as_ref()
            .and_then(PageSetupProperties::automatic_page_breaks)
    }

    /// Get the complete optional `pageSetUpPr` authoring state.
    pub fn page_setup_properties(&self) -> Option<&PageSetupProperties> {
        self.page_setup_properties.as_ref()
    }

    /// Configure page setup for printing.
    ///
    /// # Arguments
    /// * `orientation` - "portrait" or "landscape"
    /// * `paper_size` - Paper size code (1 = Letter, 9 = A4, etc.)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_page_setup("landscape", 9); // A4 landscape
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_page_setup(&mut self, orientation: &str, paper_size: u32) {
        self.page_setup = Some(PageSetup {
            orientation: orientation.to_string(),
            paper_size,
            scale: None,
            fit_to_width: None,
            fit_to_height: None,
        });
        self.modified = true;
    }

    /// Set page setup with scaling options.
    ///
    /// # Arguments
    /// * `orientation` - "portrait" or "landscape"
    /// * `paper_size` - Paper size code
    /// * `scale` - Scale percentage (10-400), or None
    /// * `fit_to_width` - Fit to N pages wide, or None
    /// * `fit_to_height` - Fit to N pages tall, or None
    pub fn set_page_setup_with_options(
        &mut self,
        orientation: &str,
        paper_size: u32,
        scale: Option<u32>,
        fit_to_width: Option<u32>,
        fit_to_height: Option<u32>,
    ) -> SheetResult<()> {
        if !matches!(orientation, "portrait" | "landscape") {
            return Err(format!(
                "page orientation must be 'portrait' or 'landscape', got '{orientation}'"
            )
            .into());
        }
        if let Some(scale) = scale
            && !(10..=400).contains(&scale)
        {
            return Err(format!("page scale must be between 10 and 400, got {scale}").into());
        }
        const MAX_EXCEL_FIT_PAGES: u32 = 32_767;
        for (name, value) in [
            ("fit-to-width", fit_to_width),
            ("fit-to-height", fit_to_height),
        ] {
            if let Some(value) = value
                && value > MAX_EXCEL_FIT_PAGES
            {
                return Err(format!(
                    "{name} must be between 0 and {MAX_EXCEL_FIT_PAGES}, got {value}"
                )
                .into());
            }
        }

        let setup = PageSetup {
            orientation: orientation.to_string(),
            paper_size,
            scale,
            fit_to_width,
            fit_to_height,
        };
        let activates_fit_to_page = fit_to_width.is_some() || fit_to_height.is_some();

        self.page_setup = Some(setup);
        if activates_fit_to_page {
            self.page_setup_properties
                .get_or_insert_with(PageSetupProperties::default)
                .fit_to_page = Some(true);
        }
        self.modified = true;
        Ok(())
    }

    /// Get the page setup configuration.
    pub fn get_page_setup(&self) -> Option<&PageSetup> {
        self.page_setup.as_ref()
    }

    /// Set the print area for the worksheet.
    ///
    /// # Arguments
    /// * `range` - Cell range (e.g., "A1:D20")
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_print_area("A1:F50");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_print_area(&mut self, range: &str) {
        self.print_area = Some(range.to_string());
        self.modified = true;
    }

    /// Clear the print area.
    pub fn clear_print_area(&mut self) {
        self.print_area = None;
        self.modified = true;
    }

    /// Get the print area.
    pub fn get_print_area(&self) -> Option<&str> {
        self.print_area.as_deref()
    }

    // ===== Headers and Footers =====

    /// Set header and footer for the worksheet.
    ///
    /// # Arguments
    /// * `header_footer` - Header and footer configuration
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::{Workbook, writer::HeaderFooter};
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    ///
    /// let mut hf = HeaderFooter::default();
    /// hf.header_center = Some("Company Name".to_string());
    /// hf.footer_left = Some("&D".to_string()); // Current date
    /// hf.footer_right = Some("Page &P of &N".to_string()); // Page numbers
    ///
    /// ws.set_header_footer(hf);
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_header_footer(&mut self, header_footer: HeaderFooter) {
        self.header_footer = Some(header_footer);
        self.modified = true;
    }

    /// Clear header and footer.
    pub fn clear_header_footer(&mut self) {
        self.header_footer = None;
        self.modified = true;
    }

    /// Get the header and footer configuration.
    pub fn get_header_footer(&self) -> Option<&HeaderFooter> {
        self.header_footer.as_ref()
    }

    // ===== Repeating Rows and Columns =====

    /// Set repeating rows for printing (print titles).
    ///
    /// # Arguments
    /// * `rows` - Row range (e.g., "1:2" for rows 1-2)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_repeating_rows("1:1"); // Repeat row 1 on each printed page
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_repeating_rows(&mut self, rows: &str) {
        self.repeating_rows = Some(rows.to_string());
        self.modified = true;
    }

    /// Clear repeating rows.
    pub fn clear_repeating_rows(&mut self) {
        self.repeating_rows = None;
        self.modified = true;
    }

    /// Get repeating rows.
    pub fn get_repeating_rows(&self) -> Option<&str> {
        self.repeating_rows.as_deref()
    }

    /// Set repeating columns for printing (print titles).
    ///
    /// # Arguments
    /// * `columns` - Column range (e.g., "A:B" for columns A-B)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_repeating_columns("A:A"); // Repeat column A on each printed page
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_repeating_columns(&mut self, columns: &str) {
        self.repeating_columns = Some(columns.to_string());
        self.modified = true;
    }

    /// Clear repeating columns.
    pub fn clear_repeating_columns(&mut self) {
        self.repeating_columns = None;
        self.modified = true;
    }

    /// Get repeating columns.
    pub fn get_repeating_columns(&self) -> Option<&str> {
        self.repeating_columns.as_deref()
    }

    // ===== Auto-filter =====

    /// Add an auto-filter to a range.
    ///
    /// # Arguments
    /// * `range` - Cell range (e.g., "A1:D10")
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.set_auto_filter("A1:D10");
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn set_auto_filter(&mut self, range: &str) {
        self.auto_filter = Some(AutoFilter {
            range: range.to_string(),
            sort_state: None,
        });
        self.modified = true;
    }

    /// Set the auto-filter sort state.
    pub fn set_auto_filter_sort_state(&mut self, sort_state: SortState) {
        if let Some(ref mut filter) = self.auto_filter {
            filter.sort_state = Some(sort_state);
        } else {
            self.auto_filter = Some(AutoFilter {
                range: sort_state.ref_range.clone(),
                sort_state: Some(sort_state),
            });
        }
        self.modified = true;
    }

    /// Clear auto-filter sort state.
    pub fn clear_auto_filter_sort_state(&mut self) {
        if let Some(ref mut filter) = self.auto_filter {
            filter.sort_state = None;
            self.modified = true;
        }
    }

    /// Remove the auto-filter.
    pub fn remove_auto_filter(&mut self) {
        self.auto_filter = None;
        self.modified = true;
    }

    /// Get the auto-filter configuration.
    pub fn get_auto_filter(&self) -> Option<&AutoFilter> {
        self.auto_filter.as_ref()
    }

    // ===== Sheet Views =====

    /// Set sheet view settings.
    pub fn set_sheet_view(&mut self, view: SheetView) {
        self.sheet_view = Some(view);
        self.modified = true;
    }

    /// Clear sheet view settings.
    pub fn clear_sheet_view(&mut self) {
        self.sheet_view = None;
        self.modified = true;
    }

    /// Get sheet view settings.
    pub fn sheet_view(&self) -> Option<&SheetView> {
        self.sheet_view.as_ref()
    }

    // ===== Page Breaks =====

    /// Add a manual row page break.
    pub fn add_row_break(&mut self, row: u32, min_col: u32, max_col: u32) {
        if row == 0 {
            return;
        }
        let break_entry = PageBreak {
            id: row,
            min: min_col,
            max: max_col,
            manual: true,
        };
        self.row_breaks.push(break_entry);
        self.modified = true;
    }

    /// Add a manual column page break.
    pub fn add_column_break(&mut self, col: u32, min_row: u32, max_row: u32) {
        if col == 0 {
            return;
        }
        let break_entry = PageBreak {
            id: col,
            min: min_row,
            max: max_row,
            manual: true,
        };
        self.col_breaks.push(break_entry);
        self.modified = true;
    }

    /// Get all row page breaks.
    pub fn row_breaks(&self) -> &[PageBreak] {
        &self.row_breaks
    }

    /// Get all column page breaks.
    pub fn column_breaks(&self) -> &[PageBreak] {
        &self.col_breaks
    }

    // ===== Tables =====

    /// Add a table to the worksheet.
    ///
    /// Tables provide structured references and enhanced formatting.
    /// Table names must be unique within the workbook and cannot contain spaces.
    pub fn add_table(&mut self, table: Table) {
        self.tables.push(table);
        self.modified = true;
    }

    /// Get all tables in the worksheet.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Get a mutable reference to all tables.
    pub fn tables_mut(&mut self) -> &mut Vec<Table> {
        &mut self.tables
    }

    /// Find a table by name.
    pub fn find_table(&self, name: &str) -> Option<&Table> {
        self.tables
            .iter()
            .find(|t| t.name == name || t.display_name == name)
    }

    /// Find a table by range.
    pub fn find_table_by_range(&self, range: &str) -> Option<&Table> {
        self.tables.iter().find(|t| t.ref_range == range)
    }

    /// Remove a table by name.
    pub fn remove_table(&mut self, name: &str) -> bool {
        if let Some(pos) = self
            .tables
            .iter()
            .position(|t| t.name == name || t.display_name == name)
        {
            self.tables.remove(pos);
            self.modified = true;
            true
        } else {
            false
        }
    }

    // ===== Sheet Protection =====

    /// Protect the worksheet with optional password.
    ///
    /// # Arguments
    /// * `password` - Optional password (will be hashed)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.protect_sheet(Some("password123"));
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn protect_sheet(&mut self, password: Option<&str>) {
        let mut protection = WorksheetProtection::default();
        protection.set_sheet_locked(true);
        protection.set_format_cells_locked(false);
        protection.set_format_columns_locked(false);
        protection.set_format_rows_locked(false);
        protection.set_insert_columns_locked(false);
        protection.set_insert_rows_locked(false);
        protection.set_insert_hyperlinks_locked(false);
        protection.set_delete_columns_locked(false);
        protection.set_delete_rows_locked(false);
        protection.set_sort_locked(false);
        protection.set_auto_filter_locked(false);
        protection.set_pivot_tables_locked(false);
        if let Some(password) = password {
            protection
                .set_verifier(Some(ProtectionPasswordVerifier::Legacy(
                    u16::from_str_radix(&Self::hash_password(password), 16).unwrap(),
                )))
                .expect("legacy verifier is valid");
        }
        self.protection
            .set_sheet_protection(Some(protection))
            .expect("constructed protection is valid");
        self.modified = true;
    }

    /// Protect the worksheet with custom permissions.
    ///
    /// # Arguments
    /// * `password` - Optional password
    /// * `permissions` - Sheet protection settings
    pub fn protect_sheet_with_options(
        &mut self,
        password: Option<&str>,
        mut permissions: WorksheetProtection,
    ) {
        if let Some(password) = password {
            permissions
                .set_verifier(Some(ProtectionPasswordVerifier::Legacy(
                    u16::from_str_radix(&Self::hash_password(password), 16).unwrap(),
                )))
                .expect("legacy verifier is valid");
        }
        self.protection
            .set_sheet_protection(Some(permissions))
            .expect("typed worksheet protection is valid");
        self.modified = true;
    }

    pub fn set_protection_metadata(
        &mut self,
        metadata: WorksheetProtectionMetadata,
    ) -> SheetResult<()> {
        crate::xlsx::sheet_protection::validate_worksheet_protection_metadata(&metadata)?;
        self.protection = metadata;
        self.modified = true;
        Ok(())
    }

    pub fn protection_metadata(&self) -> &WorksheetProtectionMetadata {
        &self.protection
    }

    /// Unprotect the worksheet.
    pub fn unprotect_sheet(&mut self) {
        self.protection.clear_sheet_protection();
        self.modified = true;
    }

    /// Check if the worksheet is protected.
    pub fn is_protected(&self) -> bool {
        self.protection.sheet_protection().is_some()
    }

    /// Get the protection configuration.
    pub fn get_protection(&self) -> Option<&WorksheetProtection> {
        self.protection.sheet_protection()
    }

    /// Simple password hashing for Excel (XOR-based).
    ///
    /// Note: This is NOT a secure hash! Excel uses a very weak password protection.
    /// For production use, consider using stronger protection methods at the file system level.
    pub(crate) fn hash_password(password: &str) -> String {
        let mut hash: u16 = 0;

        for ch in password.chars().rev() {
            let char_code = ch as u16;
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= char_code;
        }

        hash ^= password.len() as u16;
        hash ^= 0xCE4B;

        format!("{:04X}", hash)
    }

    // ===== Row/Column Grouping =====

    /// Group rows (create an outline/collapsible section).
    ///
    /// # Arguments
    /// * `start_row` - Starting row (1-based, inclusive)
    /// * `end_row` - Ending row (1-based, inclusive)
    /// * `level` - Outline level (1-7, where 1 is outermost)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.group_rows(2, 5, 1); // Group rows 2-5 at level 1
    /// ws.group_rows(3, 4, 2); // Nested group rows 3-4 at level 2
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn group_rows(&mut self, start_row: u32, end_row: u32, level: u8) {
        let level = level.clamp(1, 7); // Excel supports levels 1-7
        // Convert from 1-based (API) to 0-based (internal storage)
        for row in (start_row - 1)..=(end_row - 1) {
            self.row_outline_levels.insert(row, level);
        }
        self.modified = true;
    }

    /// Ungroup rows.
    ///
    /// # Arguments
    /// * `start_row` - Starting row (1-based, inclusive)
    /// * `end_row` - Ending row (1-based, inclusive)
    pub fn ungroup_rows(&mut self, start_row: u32, end_row: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        for row in (start_row - 1)..=(end_row - 1) {
            self.row_outline_levels.remove(&row);
        }
        self.modified = true;
    }

    /// Group columns (create an outline/collapsible section).
    ///
    /// # Arguments
    /// * `start_col` - Starting column (1-based, inclusive)
    /// * `end_col` - Ending column (1-based, inclusive)
    /// * `level` - Outline level (1-7, where 1 is outermost)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi_ooxml::xlsx::Workbook;
    ///
    /// let mut wb = Workbook::create()?;
    /// let mut ws = wb.worksheet_mut(0)?;
    /// ws.group_columns(2, 5, 1); // Group columns B-E at level 1
    /// wb.save("output.xlsx")?;
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn group_columns(&mut self, start_col: u32, end_col: u32, level: u8) {
        let level = level.clamp(1, 7); // Excel supports levels 1-7
        // Convert from 1-based (API) to 0-based (internal storage)
        for col in (start_col - 1)..=(end_col - 1) {
            self.column_outline_levels.insert(col, level);
        }
        self.modified = true;
    }

    /// Ungroup columns.
    ///
    /// # Arguments
    /// * `start_col` - Starting column (1-based, inclusive)
    /// * `end_col` - Ending column (1-based, inclusive)
    pub fn ungroup_columns(&mut self, start_col: u32, end_col: u32) {
        // Convert from 1-based (API) to 0-based (internal storage)
        for col in (start_col - 1)..=(end_col - 1) {
            self.column_outline_levels.remove(&col);
        }
        self.modified = true;
    }

    /// Get the outline level for a specific row.
    pub fn get_row_outline_level(&self, row: u32) -> Option<u8> {
        self.row_outline_levels.get(&row).copied()
    }

    /// Get the outline level for a specific column.
    pub fn get_column_outline_level(&self, col: u32) -> Option<u8> {
        self.column_outline_levels.get(&col).copied()
    }

    /// Get the used range dimensions (min_row, min_col, max_row, max_col).
    pub fn used_range(&self) -> Option<(u32, u32, u32, u32)> {
        if self.cells.is_empty() {
            return None;
        }

        let mut min_row = u32::MAX;
        let mut max_row = 0;
        let mut min_col = u32::MAX;
        let mut max_col = 0;

        for &(row, col) in self.cells.keys() {
            min_row = min_row.min(row);
            max_row = max_row.max(row);
            min_col = min_col.min(col);
            max_col = max_col.max(col);
        }

        Some((min_row, min_col, max_row, max_col))
    }

    /// Serialize the worksheet to XML with hyperlink relationship IDs.
    ///
    /// # Arguments
    /// * `shared_strings` - Mutable shared strings table
    /// * `style_indices` - Optional map of cell positions to style indices
    /// * `hyperlink_rel_ids` - Map of cell references to relationship IDs for external hyperlinks
    /// * `vml_rel_id` - Optional relationship ID for VML drawing (for comments)
    pub fn to_xml_with_hyperlink_rels(
        &self,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
        hyperlink_rel_ids: &HashMap<String, String>,
        vml_rel_id: Option<&str>,
        pivot_table_rel_ids: Option<&[String]>,
        table_rel_ids: Option<&[String]>,
    ) -> SheetResult<String> {
        self.to_xml_internal(
            shared_strings,
            style_indices,
            WorksheetPartRelationships {
                hyperlinks: Some(hyperlink_rel_ids),
                vml_drawing: vml_rel_id,
                pivot_tables: pivot_table_rel_ids,
                tables: table_rel_ids,
                drawing: None,
            },
        )
    }

    pub(crate) fn to_xml_with_part_rels(
        &self,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
        relationships: WorksheetPartRelationships<'_>,
    ) -> SheetResult<String> {
        self.to_xml_internal(shared_strings, style_indices, relationships)
    }

    /// Serialize the worksheet to XML.
    ///
    /// # Arguments
    /// * `shared_strings` - Mutable shared strings table
    /// * `style_indices` - Optional map of cell positions to style indices
    pub fn to_xml(
        &self,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
    ) -> SheetResult<String> {
        self.to_xml_internal(
            shared_strings,
            style_indices,
            WorksheetPartRelationships::default(),
        )
    }

    /// Internal method for XML serialization with optional hyperlink relationship IDs.
    fn to_xml_internal(
        &self,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
        relationships: WorksheetPartRelationships<'_>,
    ) -> SheetResult<String> {
        let mut xml = String::with_capacity(4096);
        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        let xr_uid = generate_guid_braced();
        write!(
            xml,
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{}">"#,
            xr_uid
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        // Write sheetPr (sheet properties) if needed - must come BEFORE dimension per OOXML spec
        if self.tab_color.is_some() || self.page_setup_properties.is_some() {
            xml.push_str("<sheetPr>");
            if let Some(ref color) = self.tab_color {
                write!(xml, r#"<tabColor rgb="{}"/>"#, color)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(properties) = self.page_setup_properties {
                xml.push_str("<pageSetUpPr");
                if let Some(automatic_page_breaks) = properties.automatic_page_breaks {
                    write!(
                        xml,
                        r#" autoPageBreaks="{}""#,
                        u8::from(automatic_page_breaks)
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;
                }
                if let Some(fit_to_page) = properties.fit_to_page {
                    write!(xml, r#" fitToPage="{}""#, u8::from(fit_to_page))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
                xml.push_str("/>");
            }
            xml.push_str("</sheetPr>");
        }

        // Write sheet dimensions
        // NOTE: Excel uses 1-based row/column numbering in XML
        if let Some((min_row, min_col, max_row, max_col)) = self.used_range() {
            let min_ref = format!("{}{}", Self::column_to_letters(min_col + 1), min_row + 1);
            let max_ref = format!("{}{}", Self::column_to_letters(max_col + 1), max_row + 1);
            write!(
                xml,
                r#"<dimension ref="{}:{}"/>"#,
                escape_xml(&min_ref),
                escape_xml(&max_ref)
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        } else {
            xml.push_str(r#"<dimension ref="A1"/>"#);
        }

        // Write sheet views (including freeze panes if set)
        xml.push_str("<sheetViews><sheetView");
        let workbook_view_id = self
            .sheet_view
            .as_ref()
            .and_then(|view| view.workbook_view_id)
            .unwrap_or(0);
        write!(xml, r#" workbookViewId="{}""#, workbook_view_id)
            .map_err(|e| format!("XML write error: {}", e))?;
        if self.is_active() {
            xml.push_str(" tabSelected=\"1\"");
        } else if let Some(tab_selected) =
            self.sheet_view.as_ref().and_then(|view| view.tab_selected)
        {
            write!(
                xml,
                r#" tabSelected="{}""#,
                if tab_selected { 1 } else { 0 }
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref view) = self.sheet_view {
            self.write_sheet_view_attributes(&mut xml, view)?;
        }

        // Add freeze panes if configured
        if let Some(ref freeze) = self.freeze_panes {
            if self
                .sheet_view
                .as_ref()
                .is_some_and(|view| view.pane.is_some() || !view.selections.is_empty())
            {
                return Err(
                    "Freeze panes and explicit sheet-view pane selections cannot both be set"
                        .into(),
                );
            }
            xml.push('>');

            let y_split = freeze.freeze_rows;
            let x_split = freeze.freeze_cols;

            let active_pane = match (x_split > 0, y_split > 0) {
                (true, true) => "bottomRight",
                (true, false) => "topRight",
                (false, true) => "bottomLeft",
                (false, false) => "",
            };

            // NOTE: Excel uses 1-based numbering for topLeftCell
            let top_left_cell = format!("{}{}", Self::column_to_letters(x_split + 1), y_split + 1);

            write!(
                xml,
                r#"<pane xSplit="{}" ySplit="{}" topLeftCell="{}" activePane="{}" state="frozen"/>"#,
                x_split, y_split, top_left_cell, active_pane
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            if !active_pane.is_empty() {
                write!(
                    xml,
                    r#"<selection pane="{}" activeCell="{}" sqref="{}"/>"#,
                    active_pane, top_left_cell, top_left_cell
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</sheetView>");
        } else if let Some(view) = self.sheet_view.as_ref()
            && (view.pane.is_some() || !view.selections.is_empty())
        {
            xml.push('>');
            if let Some(pane) = view.pane.as_ref() {
                self.write_sheet_pane(&mut xml, pane)?;
            }
            if view.selections.len() > 4 {
                return Err("A sheet view cannot contain more than four selections".into());
            }
            for selection in &view.selections {
                self.write_sheet_selection(&mut xml, selection)?;
            }
            xml.push_str("</sheetView>");
        } else {
            xml.push_str("/>");
        }

        xml.push_str("</sheetViews>");
        xml.push_str("<sheetFormatPr defaultRowHeight=\"15\"/>");

        // Write column information (widths and hidden columns)
        let has_pivots = relationships
            .pivot_tables
            .map(|ids| !ids.is_empty())
            .unwrap_or(false);
        self.write_cols(&mut xml, has_pivots)?;

        // Write sheet data
        xml.push_str("<sheetData>");
        self.write_sheet_data(&mut xml, shared_strings, style_indices)?;
        xml.push_str("</sheetData>");

        // Write sheet protection if configured (must come right after sheetData per OOXML spec)
        if self.protection.sheet_protection().is_some()
            || self
                .protection
                .protected_range_collections()
                .iter()
                .any(|collection| collection.source() == crate::xlsx::ProtectedRangeSource::Core)
        {
            self.write_sheet_protection(&mut xml)?;
        }

        // Write auto-filter if configured (must come before mergeCells per OOXML spec)
        if self.auto_filter.is_some() {
            self.write_auto_filter(&mut xml)?;
        }

        // Write merged cells
        if !self.merged_cells.is_empty() {
            write!(xml, r#"<mergeCells count="{}">"#, self.merged_cells.len())
                .map_err(|e| format!("XML write error: {}", e))?;

            for (start_row, start_col, end_row, end_col) in &self.merged_cells {
                // NOTE: Excel uses 1-based numbering for cell references
                let start_ref = format!(
                    "{}{}",
                    Self::column_to_letters(start_col + 1),
                    start_row + 1
                );
                let end_ref = format!("{}{}", Self::column_to_letters(end_col + 1), end_row + 1);
                write!(xml, r#"<mergeCell ref="{}:{}"/>"#, start_ref, end_ref)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</mergeCells>");
        }

        // Match Excel's worksheet structure: phoneticPr comes after mergeCells.
        xml.push_str(r#"<phoneticPr fontId="0" type="noConversion"/>"#);

        // Write conditional formatting
        if !self.conditional_formats.is_empty() {
            self.write_conditional_formatting(&mut xml)?;
        }

        // Write data validations
        if let Some(collections) = self.typed_data_validations.as_ref() {
            xml.push_str(&write_data_validation_core(
                collections,
                DataValidationConformance::Transitional,
            )?);
        } else if !self.validations.is_empty() {
            self.write_data_validations(&mut xml)?;
        }

        // Write hyperlinks
        if !self.hyperlinks.is_empty() {
            self.write_hyperlinks(&mut xml, relationships.hyperlinks)?;
        }

        // Write page margins (required by Excel)
        self.write_page_margins(&mut xml)?;

        // Write page setup if configured
        if self.page_setup.is_some() {
            self.write_page_setup(&mut xml)?;
        }

        // Write header and footer if configured
        if self.header_footer.is_some() {
            self.write_header_footer(&mut xml)?;
        }

        // Write manual page breaks
        if !self.row_breaks.is_empty() || !self.col_breaks.is_empty() {
            self.write_page_breaks(&mut xml)?;
        }

        // Write drawing reference for charts and images
        if !self.charts.is_empty() || !self.images.is_empty() {
            write!(
                xml,
                r#"<drawing r:id="{}"/>"#,
                relationships.drawing.unwrap_or("rId1")
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }

        // Write legacyDrawing reference for comments (VML)
        if let Some(vml_rel_id) = relationships.vml_drawing {
            write!(xml, r#"<legacyDrawing r:id="{}"/>"#, vml_rel_id)
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        // Write tableParts if tables are present
        if let Some(table_rels) = relationships.tables.filter(|rels| !rels.is_empty()) {
            write!(xml, r#"<tableParts count="{}">"#, table_rels.len())
                .map_err(|e| format!("XML write error: {}", e))?;
            for rel_id in table_rels {
                write!(xml, r#"<tablePart r:id="{}"/>"#, rel_id)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("</tableParts>");
        }

        if !self.sparkline_groups.is_empty() {
            xml.push_str("<extLst>");
            write_sparkline_groups_ext(&mut xml, &self.sparkline_groups)?;
            xml.push_str("</extLst>");
        }

        if let Some(collections) = self.typed_data_validations.as_ref() {
            xml.push_str(&write_data_validation_extensions(
                collections,
                DataValidationConformance::Transitional,
            )?);
        }
        xml.push_str(&write_worksheet_protection_extensions(
            &self.protection,
            WorksheetProtectionConformance::Transitional,
        )?);
        xml.push_str("</worksheet>");

        Ok(xml)
    }

    /// Get cell formats for all cells (used by workbook to build styles).
    pub fn cell_formats(&self) -> &HashMap<(u32, u32), CellFormat> {
        &self.cell_formats
    }

    /// Write sheet data (rows and cells).
    fn write_sheet_data(
        &self,
        xml: &mut String,
        shared_strings: &mut MutableSharedStrings,
        style_indices: &HashMap<(u32, u32), usize>,
    ) -> SheetResult<()> {
        if self.cells.is_empty() {
            return Ok(());
        }

        let mut rows: HashMap<u32, Vec<(u32, &CellValue)>> = HashMap::new();
        for (&(row, col), value) in &self.cells {
            rows.entry(row).or_default().push((col, value));
        }

        // Sort rows
        let mut row_nums: Vec<u32> = rows.keys().copied().collect();
        row_nums.sort_unstable();

        for row_num in row_nums {
            let mut cells = rows[&row_num].clone();
            cells.sort_unstable_by_key(|(col, _)| *col);

            // NOTE: Excel uses 1-based row numbering
            write!(xml, r#"<row r="{}""#, row_num + 1)
                .map_err(|e| format!("XML write error: {}", e))?;

            // Add custom row height if specified
            if let Some(&height) = self.row_heights.get(&row_num) {
                write!(xml, r#" ht="{}" customHeight="1""#, height)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            // Add hidden attribute if row is hidden
            if self.hidden_rows.contains(&row_num) {
                xml.push_str(r#" hidden="1""#);
            }

            xml.push('>');

            for (col_num, value) in cells {
                // NOTE: Excel uses 1-based numbering for cell references
                let cell_ref = format!("{}{}", Self::column_to_letters(col_num + 1), row_num + 1);
                // Get the style index for this cell (if any)
                let style_index = style_indices.get(&(row_num, col_num)).copied();

                if let Some(runs) = self.rich_text_cells.get(&(row_num, col_num)) {
                    self.write_rich_text_cell(xml, &cell_ref, runs, style_index)?;
                } else {
                    self.write_cell(xml, &cell_ref, value, shared_strings, style_index)?;
                }
            }

            xml.push_str("</row>");
        }

        Ok(())
    }

    /// Write a rich text cell to XML as an inline string.
    fn write_rich_text_cell(
        &self,
        xml: &mut String,
        cell_ref: &str,
        runs: &[RichTextRun],
        style_index: Option<usize>,
    ) -> SheetResult<()> {
        if runs.is_empty() {
            // Fallback: treat as empty cell
            return Ok(());
        }

        let style_attr = if let Some(idx) = style_index {
            format!(r#" s="{}""#, idx)
        } else {
            String::new()
        };

        write!(
            xml,
            r#"<c r="{}"{}  t="inlineStr"><is>"#,
            cell_ref, style_attr
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        for run in runs {
            xml.push_str("<r>");

            // Run properties
            let has_rpr = run.font_name.is_some()
                || run.font_size.is_some()
                || run.bold
                || run.italic
                || run.underline
                || run.color.is_some();

            if has_rpr {
                xml.push_str("<rPr>");

                if let Some(ref name) = run.font_name {
                    write!(xml, "<rFont val=\"{}\"/>", escape_xml(name))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
                if let Some(size) = run.font_size {
                    write!(xml, "<sz val=\"{}\"/>", size)
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
                if run.bold {
                    xml.push_str("<b/>");
                }
                if run.italic {
                    xml.push_str("<i/>");
                }
                if run.underline {
                    xml.push_str("<u/>");
                }
                if let Some(ref color) = run.color {
                    write!(xml, "<color rgb=\"{}\"/>", escape_xml(color))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                xml.push_str("</rPr>");
            }

            // Text content; use xml:space="preserve" to keep leading/trailing spaces
            write!(
                xml,
                "<t xml:space=\"preserve\">{}</t>",
                escape_xml(&run.text)
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            xml.push_str("</r>");
        }

        xml.push_str("</is></c>");

        Ok(())
    }

    /// Write a single cell to XML.
    fn write_cell(
        &self,
        xml: &mut String,
        cell_ref: &str,
        value: &CellValue,
        shared_strings: &mut MutableSharedStrings,
        style_index: Option<usize>,
    ) -> SheetResult<()> {
        // Helper to add style attribute if present
        let style_attr = if let Some(idx) = style_index {
            format!(r#" s="{}""#, idx)
        } else {
            String::new()
        };

        match value {
            CellValue::Empty => {},
            CellValue::String(s) => {
                let string_index = shared_strings.add_string(s);
                write!(
                    xml,
                    r#"<c r="{}"{}  t="s"><v>{}</v></c>"#,
                    cell_ref, style_attr, string_index
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Int(i) => {
                write!(xml, r#"<c r="{}"{}><v>{}</v></c>"#, cell_ref, style_attr, i)
                    .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Float(f) => {
                write!(xml, r#"<c r="{}"{}><v>{}</v></c>"#, cell_ref, style_attr, f)
                    .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Bool(b) => {
                write!(
                    xml,
                    r#"<c r="{}"{}  t="b"><v>{}</v></c>"#,
                    cell_ref,
                    style_attr,
                    if *b { "1" } else { "0" }
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::DateTime(d) => {
                write!(xml, r#"<c r="{}"{}><v>{}</v></c>"#, cell_ref, style_attr, d)
                    .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Error(e) => {
                write!(
                    xml,
                    r#"<c r="{}"{}  t="e"><v>{}</v></c>"#,
                    cell_ref,
                    style_attr,
                    escape_xml(e)
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            },
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => {
                xml.push_str(&format!(r#"<c r="{}"{}>"#, cell_ref, style_attr));
                if *is_array {
                    if let Some(r) = array_range {
                        write!(
                            xml,
                            "<f t=\"array\" ref=\"{}\">{}</f>",
                            escape_xml(r),
                            escape_xml(formula),
                        )
                        .map_err(|e| format!("XML write error: {}", e))?;
                    } else {
                        write!(xml, "<f t=\"array\">{}</f>", escape_xml(formula))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }
                } else {
                    write!(xml, "<f>{}</f>", escape_xml(formula))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                if let Some(cached) = cached_value {
                    match &**cached {
                        CellValue::String(s) => {
                            let string_index = shared_strings.add_string(s);
                            write!(xml, "<v>{}</v>", string_index)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Int(i) => {
                            write!(xml, "<v>{}</v>", i)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Float(f) => {
                            write!(xml, "<v>{}</v>", f)
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        CellValue::Bool(b) => {
                            write!(xml, "<v>{}</v>", if *b { "1" } else { "0" })
                                .map_err(|e| format!("XML write error: {}", e))?;
                        },
                        _ => {},
                    }
                }
                xml.push_str("</c>");
            },
        }

        Ok(())
    }

    /// Write column information (widths and hidden state).
    fn write_cols(&self, xml: &mut String, force: bool) -> SheetResult<()> {
        if self.column_widths.is_empty() && self.hidden_columns.is_empty() {
            if !force {
                return Ok(());
            }
            xml.push_str("<cols>");
            // Minimal valid column definitions to match Excel's pivot sheet structure.
            // Only required attributes are emitted.
            for col in 1..=7u32 {
                write!(
                    xml,
                    r#"<col min="{}" max="{}" width="10" bestFit="1" customWidth="1"/>"#,
                    col, col
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }
            xml.push_str("</cols>");
            return Ok(());
        }

        // Determine the set of columns that need a <col> entry.
        let mut cols_to_write: std::collections::BTreeSet<u32> = std::collections::BTreeSet::new();
        cols_to_write.extend(self.column_widths.keys());
        cols_to_write.extend(&self.hidden_columns);

        if cols_to_write.is_empty() {
            return Ok(());
        }

        xml.push_str("<cols>");

        for &col in &cols_to_write {
            // NOTE: Excel uses 1-based column numbering for min/max attributes
            write!(xml, r#"<col min="{}" max="{}""#, col + 1, col + 1)
                .map_err(|e| format!("XML write error: {}", e))?;

            // Add width if specified
            if let Some(&width) = self.column_widths.get(&col) {
                write!(xml, r#" width="{}" customWidth="1""#, width)
                    .map_err(|e| format!("XML write error: {}", e))?;
            } else {
                // Default Excel column width is 8.43
                xml.push_str(r#" width="8.43""#);
            }

            // Add hidden attribute if column is hidden
            if self.hidden_columns.contains(&col) {
                xml.push_str(r#" hidden="1""#);
            }

            xml.push_str("/>");
        }

        xml.push_str("</cols>");
        Ok(())
    }

    /// Write data validations.
    fn write_data_validations(&self, xml: &mut String) -> SheetResult<()> {
        if self.validations.is_empty() {
            return Ok(());
        }

        write!(
            xml,
            r#"<dataValidations count="{}">"#,
            self.validations.len()
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        for validation in &self.validations {
            xml.push_str(r#"<dataValidation"#);

            // Write type and operator
            match &validation.validation_type {
                DataValidationType::List { values } => {
                    xml.push_str(r#" type="list""#);
                    write!(xml, r#" sqref="{}""#, escape_xml(&validation.range))
                        .map_err(|e| format!("XML write error: {}", e))?;

                    if validation.show_input_message {
                        xml.push_str(r#" showInputMessage="1""#);
                    }
                    if validation.show_error_alert {
                        xml.push_str(r#" showErrorMessage="1""#);
                    }

                    if let Some(ref title) = validation.input_title {
                        write!(xml, r#" promptTitle="{}""#, escape_xml(title))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }
                    if let Some(ref msg) = validation.input_message {
                        write!(xml, r#" prompt="{}""#, escape_xml(msg))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }
                    if let Some(ref title) = validation.error_title {
                        write!(xml, r#" errorTitle="{}""#, escape_xml(title))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }
                    if let Some(ref msg) = validation.error_message {
                        write!(xml, r#" error="{}""#, escape_xml(msg))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }

                    xml.push('>');

                    // Write list values as a comma-separated string in formula1
                    let list_str = values.join(",");
                    write!(xml, "<formula1>\"{}\"</formula1>", escape_xml(&list_str))
                        .map_err(|e| format!("XML write error: {}", e))?;

                    xml.push_str("</dataValidation>");
                },
                DataValidationType::Whole {
                    operator,
                    value1,
                    value2,
                } => {
                    xml.push_str(r#" type="whole""#);
                    write!(
                        xml,
                        r#" operator="{}" sqref="{}""#,
                        operator.as_str(),
                        escape_xml(&validation.range)
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    self.write_validation_attributes(xml, validation)?;

                    xml.push('>');

                    write!(xml, "<formula1>{}</formula1>", value1)
                        .map_err(|e| format!("XML write error: {}", e))?;
                    if let Some(v2) = value2 {
                        write!(xml, "<formula2>{}</formula2>", v2)
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }

                    xml.push_str("</dataValidation>");
                },
                DataValidationType::Decimal {
                    operator,
                    value1,
                    value2,
                } => {
                    xml.push_str(r#" type="decimal""#);
                    write!(
                        xml,
                        r#" operator="{}" sqref="{}""#,
                        operator.as_str(),
                        escape_xml(&validation.range)
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    self.write_validation_attributes(xml, validation)?;

                    xml.push('>');

                    write!(xml, "<formula1>{}</formula1>", value1)
                        .map_err(|e| format!("XML write error: {}", e))?;
                    if let Some(v2) = value2 {
                        write!(xml, "<formula2>{}</formula2>", v2)
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }

                    xml.push_str("</dataValidation>");
                },
                DataValidationType::TextLength {
                    operator,
                    value1,
                    value2,
                } => {
                    xml.push_str(r#" type="textLength""#);
                    write!(
                        xml,
                        r#" operator="{}" sqref="{}""#,
                        operator.as_str(),
                        escape_xml(&validation.range)
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    self.write_validation_attributes(xml, validation)?;

                    xml.push('>');

                    write!(xml, "<formula1>{}</formula1>", value1)
                        .map_err(|e| format!("XML write error: {}", e))?;
                    if let Some(v2) = value2 {
                        write!(xml, "<formula2>{}</formula2>", v2)
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }

                    xml.push_str("</dataValidation>");
                },
                DataValidationType::Date {
                    operator,
                    value1,
                    value2,
                } => {
                    xml.push_str(r#" type="date""#);
                    write!(
                        xml,
                        r#" operator="{}" sqref="{}""#,
                        operator.as_str(),
                        escape_xml(&validation.range)
                    )
                    .map_err(|e| format!("XML write error: {}", e))?;

                    self.write_validation_attributes(xml, validation)?;

                    xml.push('>');

                    write!(xml, "<formula1>{}</formula1>", escape_xml(value1))
                        .map_err(|e| format!("XML write error: {}", e))?;
                    if let Some(v2) = value2 {
                        write!(xml, "<formula2>{}</formula2>", escape_xml(v2))
                            .map_err(|e| format!("XML write error: {}", e))?;
                    }

                    xml.push_str("</dataValidation>");
                },
                DataValidationType::Custom { formula } => {
                    xml.push_str(r#" type="custom""#);
                    write!(xml, r#" sqref="{}""#, escape_xml(&validation.range))
                        .map_err(|e| format!("XML write error: {}", e))?;

                    self.write_validation_attributes(xml, validation)?;

                    xml.push('>');

                    write!(xml, "<formula1>{}</formula1>", escape_xml(formula))
                        .map_err(|e| format!("XML write error: {}", e))?;

                    xml.push_str("</dataValidation>");
                },
            }
        }

        xml.push_str("</dataValidations>");
        Ok(())
    }

    /// Write common validation attributes.
    fn write_validation_attributes(
        &self,
        xml: &mut String,
        validation: &DataValidation,
    ) -> SheetResult<()> {
        if validation.show_input_message {
            xml.push_str(r#" showInputMessage="1""#);
        }
        if validation.show_error_alert {
            xml.push_str(r#" showErrorMessage="1""#);
        }

        if let Some(ref title) = validation.input_title {
            write!(xml, r#" promptTitle="{}""#, escape_xml(title))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref msg) = validation.input_message {
            write!(xml, r#" prompt="{}""#, escape_xml(msg))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref title) = validation.error_title {
            write!(xml, r#" errorTitle="{}""#, escape_xml(title))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref msg) = validation.error_message {
            write!(xml, r#" error="{}""#, escape_xml(msg))
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        Ok(())
    }

    /// Convert column number to Excel column letters (e.g., 1 -> "A", 26 -> "Z", 27 -> "AA").
    pub(crate) fn column_to_letters(col: u32) -> String {
        let mut letters = String::new();
        let mut col = col;

        while col > 0 {
            col -= 1;
            let letter = ((col % 26) as u8 + b'A') as char;
            letters.insert(0, letter);
            col /= 26;
        }

        letters
    }

    /// Write hyperlinks section.
    ///
    /// Note: For external URLs, this requires relationship generation which is handled separately.
    fn write_hyperlinks(
        &self,
        xml: &mut String,
        hyperlink_rel_ids: Option<&HashMap<String, String>>,
    ) -> SheetResult<()> {
        if self.hyperlinks.is_empty() {
            return Ok(());
        }

        xml.push_str("<hyperlinks>");

        for hyperlink in self.hyperlinks.iter() {
            xml.push_str(r#"<hyperlink ref=""#);
            xml.push_str(&escape_xml(&hyperlink.cell_ref));
            xml.push('"');

            // For external URLs, we need a relationship ID (rId)
            // For internal references (e.g., Sheet2!A1), we use location attribute instead
            if hyperlink.target.starts_with("http://")
                || hyperlink.target.starts_with("https://")
                || hyperlink.target.starts_with("ftp://")
                || hyperlink.target.starts_with("mailto:")
            {
                // External link - use the provided relationship ID from the map
                if let Some(rel_ids) = hyperlink_rel_ids
                    && let Some(rel_id) = rel_ids.get(&hyperlink.cell_ref)
                {
                    write!(xml, r#" r:id="{}""#, rel_id)
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
            } else {
                // Internal reference
                write!(xml, r#" location="{}""#, escape_xml(&hyperlink.target))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            // Add display/tooltip if present
            if let Some(ref display) = hyperlink.display {
                write!(xml, r#" display="{}""#, escape_xml(display))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("/>");
        }

        xml.push_str("</hyperlinks>");
        Ok(())
    }

    /// Write conditional formatting section.
    fn write_conditional_formatting(&self, xml: &mut String) -> SheetResult<()> {
        if self.conditional_formats.is_empty() {
            return Ok(());
        }

        // Group conditional formats by range
        let mut formats_by_range: HashMap<String, Vec<&ConditionalFormat>> = HashMap::new();
        for format in &self.conditional_formats {
            formats_by_range
                .entry(format.range.clone())
                .or_default()
                .push(format);
        }

        for (range, formats) in formats_by_range {
            write!(
                xml,
                r#"<conditionalFormatting sqref="{}">"#,
                escape_xml(&range)
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            for format in formats {
                self.write_conditional_format_rule(xml, format)?;
            }

            xml.push_str("</conditionalFormatting>");
        }

        Ok(())
    }

    /// Write a single conditional formatting rule.
    fn write_conditional_format_rule(
        &self,
        xml: &mut String,
        format: &ConditionalFormat,
    ) -> SheetResult<()> {
        match &format.rule_type {
            ConditionalFormatType::CellIs { operator, formula } => {
                write!(
                    xml,
                    r#"<cfRule type="cellIs" priority="{}" operator="{}">"#,
                    format.priority,
                    escape_xml(operator)
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                write!(xml, "<formula>{}</formula>", escape_xml(formula))
                    .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str("</cfRule>");
            },
            ConditionalFormatType::ColorScale {
                min_color,
                max_color,
                mid_color,
            } => {
                write!(
                    xml,
                    r#"<cfRule type="colorScale" priority="{}">"#,
                    format.priority
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str("<colorScale>");

                // OOXML Spec: ALL cfvo elements must come BEFORE color elements

                // All cfvo elements first
                xml.push_str(r#"<cfvo type="min"/>"#);
                if mid_color.is_some() {
                    xml.push_str(r#"<cfvo type="percentile" val="50"/>"#);
                }
                xml.push_str(r#"<cfvo type="max"/>"#);

                // Then all color elements
                write!(xml, r#"<color rgb="{}"/>"#, escape_xml(min_color))
                    .map_err(|e| format!("XML write error: {}", e))?;
                if let Some(mid) = mid_color {
                    write!(xml, r#"<color rgb="{}"/>"#, escape_xml(mid))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }
                write!(xml, r#"<color rgb="{}"/>"#, escape_xml(max_color))
                    .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str("</colorScale>");
                xml.push_str("</cfRule>");
            },
            ConditionalFormatType::DataBar { color, show_value } => {
                write!(
                    xml,
                    r#"<cfRule type="dataBar" priority="{}">"#,
                    format.priority
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                if *show_value {
                    xml.push_str(r#"<dataBar showValue="1">"#);
                } else {
                    xml.push_str("<dataBar>");
                }

                xml.push_str(r#"<cfvo type="min"/>"#);
                xml.push_str(r#"<cfvo type="max"/>"#);
                write!(xml, r#"<color rgb="{}"/>"#, escape_xml(color))
                    .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str("</dataBar>");
                xml.push_str("</cfRule>");
            },
            ConditionalFormatType::IconSet {
                icon_set,
                show_value,
            } => {
                write!(
                    xml,
                    r#"<cfRule type="iconSet" priority="{}">"#,
                    format.priority
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                write!(
                    xml,
                    r#"<iconSet iconSet="{}" showValue="{}">"#,
                    escape_xml(icon_set),
                    if *show_value { "1" } else { "0" }
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str(r#"<cfvo type="percent" val="0"/>"#);
                xml.push_str(r#"<cfvo type="percent" val="33"/>"#);
                xml.push_str(r#"<cfvo type="percent" val="67"/>"#);

                xml.push_str("</iconSet>");
                xml.push_str("</cfRule>");
            },
            ConditionalFormatType::Expression { formula } => {
                write!(
                    xml,
                    r#"<cfRule type="expression" dxfId="0" priority="{}">"#,
                    format.priority
                )
                .map_err(|e| format!("XML write error: {}", e))?;

                write!(xml, "<formula>{}</formula>", escape_xml(formula))
                    .map_err(|e| format!("XML write error: {}", e))?;

                xml.push_str("</cfRule>");
            },
        }

        Ok(())
    }

    /// Write page margins (required by Excel).
    fn write_page_margins(&self, xml: &mut String) -> SheetResult<()> {
        // Default margins in inches: left=0.7, right=0.7, top=0.75, bottom=0.75, header=0.3, footer=0.3
        xml.push_str(
            r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
        );
        Ok(())
    }

    /// Write page setup section.
    fn write_page_setup(&self, xml: &mut String) -> SheetResult<()> {
        if let Some(ref setup) = self.page_setup {
            xml.push_str("<pageSetup");

            // Paper size
            write!(xml, r#" paperSize="{}""#, setup.paper_size)
                .map_err(|e| format!("XML write error: {}", e))?;

            // Orientation
            if setup.orientation == "landscape" {
                xml.push_str(r#" orientation="landscape""#);
            }

            // Scale
            if let Some(scale) = setup.scale {
                write!(xml, r#" scale="{}""#, scale)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            // Fit to page
            if let Some(width) = setup.fit_to_width {
                write!(xml, r#" fitToWidth="{}""#, width)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if let Some(height) = setup.fit_to_height {
                write!(xml, r#" fitToHeight="{}""#, height)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("/>");
        }

        Ok(())
    }

    /// Write header and footer section.
    fn write_header_footer(&self, xml: &mut String) -> SheetResult<()> {
        if let Some(ref hf) = self.header_footer {
            xml.push_str("<headerFooter>");

            // Odd header (default header)
            if hf.header_left.is_some() || hf.header_center.is_some() || hf.header_right.is_some() {
                xml.push_str("<oddHeader>");
                if let Some(ref left) = hf.header_left {
                    xml.push_str("&amp;L");
                    xml.push_str(&escape_xml(left));
                }
                if let Some(ref center) = hf.header_center {
                    xml.push_str("&amp;C");
                    xml.push_str(&escape_xml(center));
                }
                if let Some(ref right) = hf.header_right {
                    xml.push_str("&amp;R");
                    xml.push_str(&escape_xml(right));
                }
                xml.push_str("</oddHeader>");
            }

            // Odd footer (default footer)
            if hf.footer_left.is_some() || hf.footer_center.is_some() || hf.footer_right.is_some() {
                xml.push_str("<oddFooter>");
                if let Some(ref left) = hf.footer_left {
                    xml.push_str("&amp;L");
                    xml.push_str(&escape_xml(left));
                }
                if let Some(ref center) = hf.footer_center {
                    xml.push_str("&amp;C");
                    xml.push_str(&escape_xml(center));
                }
                if let Some(ref right) = hf.footer_right {
                    xml.push_str("&amp;R");
                    xml.push_str(&escape_xml(right));
                }
                xml.push_str("</oddFooter>");
            }

            xml.push_str("</headerFooter>");
        }

        Ok(())
    }

    /// Write auto-filter section.
    fn write_auto_filter(&self, xml: &mut String) -> SheetResult<()> {
        if let Some(ref filter) = self.auto_filter {
            if let Some(ref sort_state) = filter.sort_state {
                write!(xml, r#"<autoFilter ref="{}">"#, escape_xml(&filter.range))
                    .map_err(|e| format!("XML write error: {}", e))?;
                self.write_sort_state(xml, sort_state)?;
                xml.push_str("</autoFilter>");
            } else {
                write!(xml, r#"<autoFilter ref="{}"/>"#, escape_xml(&filter.range))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
        }

        Ok(())
    }

    /// Write sort state for auto-filter.
    fn write_sort_state(&self, xml: &mut String, sort_state: &SortState) -> SheetResult<()> {
        write!(
            xml,
            r#"<sortState ref="{}""#,
            escape_xml(&sort_state.ref_range)
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        if let Some(v) = sort_state.column_sort {
            write!(xml, r#" columnSort="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = sort_state.case_sensitive {
            write!(xml, r#" caseSensitive="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(method) = sort_state.sort_method {
            write!(xml, r#" sortMethod="{}""#, method.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        if sort_state.conditions.is_empty() {
            xml.push_str("/>");
            return Ok(());
        }

        xml.push('>');
        for condition in &sort_state.conditions {
            self.write_sort_condition(xml, condition)?;
        }
        xml.push_str("</sortState>");
        Ok(())
    }

    fn write_sort_condition(&self, xml: &mut String, condition: &SortCondition) -> SheetResult<()> {
        write!(
            xml,
            r#"<sortCondition ref="{}""#,
            escape_xml(&condition.ref_range)
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        if let Some(v) = condition.descending {
            write!(xml, r#" descending="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(sort_by) = condition.sort_by {
            write!(xml, r#" sortBy="{}""#, sort_by.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref list) = condition.custom_list {
            write!(xml, r#" customList="{}""#, escape_xml(list))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(dxf_id) = condition.dxf_id {
            write!(xml, r#" dxfId="{}""#, dxf_id).map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref icon_set) = condition.icon_set {
            write!(xml, r#" iconSet="{}""#, escape_xml(icon_set))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(icon_id) = condition.icon_id {
            write!(xml, r#" iconId="{}""#, icon_id)
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        xml.push_str("/>");
        Ok(())
    }

    fn write_sheet_view_attributes(&self, xml: &mut String, view: &SheetView) -> SheetResult<()> {
        if let Some(v) = view.window_protection {
            write!(xml, r#" windowProtection="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_formulas {
            write!(xml, r#" showFormulas="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_grid_lines {
            write!(xml, r#" showGridLines="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_row_col_headers {
            write!(xml, r#" showRowColHeaders="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_zeros {
            write!(xml, r#" showZeros="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.right_to_left {
            write!(xml, r#" rightToLeft="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_ruler {
            write!(xml, r#" showRuler="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_outline_symbols {
            write!(xml, r#" showOutlineSymbols="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.default_grid_color {
            write!(xml, r#" defaultGridColor="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(v) = view.show_white_space {
            write!(xml, r#" showWhiteSpace="{}""#, if v { 1 } else { 0 })
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(view_type) = view.view_type {
            write!(xml, r#" view="{}""#, view_type.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(ref cell) = view.top_left_cell {
            Cell::reference_to_coords(cell)?;
            write!(xml, r#" topLeftCell="{}""#, escape_xml(cell))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(color_id) = view.color_id {
            write!(xml, r#" colorId="{}""#, color_id)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        for (name, scale) in [
            ("zoomScale", view.zoom_scale),
            ("zoomScaleNormal", view.zoom_scale_normal),
            (
                "zoomScaleSheetLayoutView",
                view.zoom_scale_sheet_layout_view,
            ),
            ("zoomScalePageLayoutView", view.zoom_scale_page_layout_view),
        ] {
            if let Some(scale) = scale {
                if !(10..=400).contains(&scale) {
                    return Err(format!("Sheet-view {name} must be between 10 and 400").into());
                }
                write!(xml, r#" {name}="{scale}""#)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
        }
        Ok(())
    }

    fn write_sheet_pane(&self, xml: &mut String, pane: &SheetPane) -> SheetResult<()> {
        xml.push_str("<pane");
        for (name, split) in [("xSplit", pane.x_split), ("ySplit", pane.y_split)] {
            if let Some(split) = split {
                if !split.is_finite() {
                    return Err(format!("Sheet-view {name} must be finite").into());
                }
                write!(xml, r#" {name}="{split}""#)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
        }
        if let Some(cell) = pane.top_left_cell.as_deref() {
            Cell::reference_to_coords(cell)?;
            write!(xml, r#" topLeftCell="{}""#, escape_xml(cell))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(active_pane) = pane.active_pane {
            write!(xml, r#" activePane="{}""#, active_pane.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(state) = pane.state {
            write!(xml, r#" state="{}""#, state.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("/>");
        Ok(())
    }

    fn write_sheet_selection(
        &self,
        xml: &mut String,
        selection: &SheetSelection,
    ) -> SheetResult<()> {
        xml.push_str("<selection");
        if let Some(pane) = selection.pane {
            write!(xml, r#" pane="{}""#, pane.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(cell) = selection.active_cell.as_deref() {
            Cell::reference_to_coords(cell)?;
            write!(xml, r#" activeCell="{}""#, escape_xml(cell))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(active_cell_id) = selection.active_cell_id {
            write!(xml, r#" activeCellId="{}""#, active_cell_id)
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        if let Some(sqref) = selection.sqref.as_deref() {
            Self::validate_sheet_sqref(sqref)?;
            write!(xml, r#" sqref="{}""#, escape_xml(sqref))
                .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("/>");
        Ok(())
    }

    fn validate_sheet_sqref(sqref: &str) -> SheetResult<()> {
        let mut references = sqref.split_whitespace().peekable();
        if references.peek().is_none() {
            return Err("Sheet-view selection sqref cannot be empty".into());
        }
        for reference in references {
            let mut cells = reference.split(':');
            let start = cells
                .next()
                .ok_or("Sheet-view selection contains an empty reference")?;
            let (start_column, start_row) = Cell::reference_to_coords(start)?;
            if let Some(end) = cells.next() {
                let (end_column, end_row) = Cell::reference_to_coords(end)?;
                if start_column > end_column || start_row > end_row {
                    return Err(format!("Invalid sheet-view selection range '{reference}'").into());
                }
            }
            if cells.next().is_some() {
                return Err(format!("Invalid sheet-view selection range '{reference}'").into());
            }
        }
        Ok(())
    }

    fn write_page_breaks(&self, xml: &mut String) -> SheetResult<()> {
        if !self.row_breaks.is_empty() {
            self.write_break_list(xml, "rowBreaks", &self.row_breaks)?;
        }
        if !self.col_breaks.is_empty() {
            self.write_break_list(xml, "colBreaks", &self.col_breaks)?;
        }
        Ok(())
    }

    fn write_break_list(
        &self,
        xml: &mut String,
        tag: &str,
        breaks: &[PageBreak],
    ) -> SheetResult<()> {
        write!(
            xml,
            r#"<{} count="{}" manualBreakCount="{}">"#,
            tag,
            breaks.len(),
            breaks.len()
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        for brk in breaks {
            write!(
                xml,
                r#"<brk id="{}" min="{}" max="{}" man="{}"/>"#,
                brk.id,
                brk.min,
                brk.max,
                if brk.manual { 1 } else { 0 }
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }

        write!(xml, "</{}>", tag).map_err(|e| format!("XML write error: {}", e))?;
        Ok(())
    }

    /// Write sheet protection section.
    fn write_sheet_protection(&self, xml: &mut String) -> SheetResult<()> {
        xml.push_str(&write_worksheet_protection_core(
            &self.protection,
            WorksheetProtectionConformance::Transitional,
        )?);
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_worksheet() {
        let ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        assert_eq!(ws.name(), "Sheet1");
        assert_eq!(ws.sheet_id(), 1);
        assert_eq!(ws.cell_count(), 0);
    }

    #[test]
    fn test_set_cell_value() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_cell_value(1, 1, "Hello");
        ws.set_cell_value(1, 2, 42);
        ws.set_cell_value(2, 1, 3.15);

        assert_eq!(ws.cell_count(), 3);
        assert!(matches!(ws.cell_value(1, 1), Some(CellValue::String(_))));
    }

    #[test]
    fn rich_text_cell_generates_inline_string() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);

        ws.set_rich_text_cell(
            1,
            1,
            vec![
                RichTextRun {
                    text: "Hello ".to_string(),
                    font_name: None,
                    font_size: None,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                },
                RichTextRun {
                    text: "World".to_string(),
                    font_name: Some("Calibri".to_string()),
                    font_size: Some(11.0),
                    bold: true,
                    italic: false,
                    underline: false,
                    color: Some("FF0000FF".to_string()),
                },
            ],
        );

        let mut shared_strings = MutableSharedStrings::new();
        let styles: HashMap<(u32, u32), usize> = HashMap::new();

        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();

        assert!(xml.contains("t=\"inlineStr\""));
        assert!(xml.contains("<is>"));
        assert!(xml.contains("Hello "));
        assert!(xml.contains("World"));
    }

    #[test]
    fn test_column_to_letters() {
        assert_eq!(MutableWorksheet::column_to_letters(1), "A");
        assert_eq!(MutableWorksheet::column_to_letters(26), "Z");
        assert_eq!(MutableWorksheet::column_to_letters(27), "AA");
        assert_eq!(MutableWorksheet::column_to_letters(702), "ZZ");
    }

    #[test]
    fn phonetic_pr_is_after_merge_cells() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_cell_value(1, 1, "A");
        ws.merge_cells(1, 1, 1, 2);

        let mut shared_strings = MutableSharedStrings::new();
        let styles: HashMap<(u32, u32), usize> = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();

        let merge = xml.find("<mergeCells").unwrap();
        let phonetic = xml.find("<phoneticPr").unwrap();
        assert!(merge < phonetic);
    }

    #[test]
    fn phonetic_pr_is_after_sheet_protection() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_cell_value(1, 1, "A");
        ws.protect_sheet(Some("secret"));

        let mut shared_strings = MutableSharedStrings::new();
        let styles: HashMap<(u32, u32), usize> = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();

        let protection = xml.find("<sheetProtection").unwrap();
        let phonetic = xml.find("<phoneticPr").unwrap();
        assert!(protection < phonetic);
    }

    #[test]
    fn phonetic_pr_is_after_auto_filter() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_cell_value(1, 1, "A");
        ws.set_auto_filter("A1:A10");

        let mut shared_strings = MutableSharedStrings::new();
        let styles: HashMap<(u32, u32), usize> = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();

        let filter = xml.find("<autoFilter").unwrap();
        let phonetic = xml.find("<phoneticPr").unwrap();
        assert!(filter < phonetic);
    }

    #[test]
    fn auto_filter_sort_state_round_trips() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_auto_filter("A1:B5");
        ws.set_auto_filter_sort_state(SortState {
            ref_range: "A2:B5".to_string(),
            column_sort: Some(false),
            case_sensitive: Some(true),
            sort_method: Some(crate::xlsx::sort::SortMethod::PinYin),
            conditions: vec![SortCondition {
                ref_range: "B2:B5".to_string(),
                descending: Some(true),
                sort_by: Some(crate::xlsx::sort::SortBy::Value),
                custom_list: None,
                dxf_id: None,
                icon_set: None,
                icon_id: None,
            }],
        });

        let mut shared_strings = MutableSharedStrings::new();
        let styles = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();
        let parsed = crate::xlsx::parsers::worksheet_parser::parse_worksheet_data(&xml).unwrap();
        let sort = parsed.auto_filter.unwrap().sort_state.unwrap();

        assert_eq!(sort.ref_range, "A2:B5");
        assert_eq!(sort.case_sensitive, Some(true));
        assert_eq!(
            sort.sort_method,
            Some(crate::xlsx::sort::SortMethod::PinYin)
        );
        assert_eq!(sort.conditions[0].ref_range, "B2:B5");
        assert_eq!(sort.conditions[0].descending, Some(true));
    }

    #[test]
    fn rejects_invalid_drawing_anchors_and_image_formats() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        assert!(ws.add_image(vec![1], "png", 0, 1, 2, 2, None).is_err());
        assert!(ws.add_image(vec![1], "png", 3, 1, 2, 2, None).is_err());
        assert!(ws.add_image(vec![1], "webp", 1, 1, 2, 2, None).is_err());

        let mut chart = WorksheetChart::bar_chart(
            "Revenue",
            "Sheet1!$A$1:$A$2",
            "Sheet1!$B$1:$B$2",
            ChartAnchor::new(4, 4, 2, 2),
        )
        .unwrap();
        chart.anchor.from_col_offset = -1;
        ws.add_chart(chart);
        assert!(ws.generate_drawing_xml().is_err());
    }

    #[test]
    fn sheet_view_attributes_round_trip() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_sheet_view(SheetView {
            workbook_view_id: Some(3),
            window_protection: Some(true),
            show_formulas: Some(true),
            show_grid_lines: Some(false),
            show_row_col_headers: Some(false),
            show_zeros: Some(false),
            right_to_left: Some(true),
            tab_selected: Some(false),
            show_ruler: Some(false),
            show_outline_symbols: Some(false),
            default_grid_color: Some(false),
            show_white_space: Some(false),
            view_type: Some(crate::xlsx::SheetViewType::PageLayout),
            top_left_cell: Some("C4".to_string()),
            color_id: Some(12),
            zoom_scale: Some(120),
            zoom_scale_normal: Some(100),
            zoom_scale_sheet_layout_view: Some(60),
            zoom_scale_page_layout_view: Some(80),
            pane: Some(SheetPane {
                x_split: Some(2.0),
                y_split: Some(3.5),
                top_left_cell: Some("C4".to_string()),
                active_pane: Some(crate::xlsx::SheetPanePosition::BottomRight),
                state: Some(crate::xlsx::SheetPaneState::FrozenSplit),
            }),
            selections: vec![SheetSelection {
                pane: Some(crate::xlsx::SheetPanePosition::BottomRight),
                active_cell: Some("D5".to_string()),
                active_cell_id: Some(1),
                sqref: Some("C4:D5 F7".to_string()),
            }],
        });

        let mut shared_strings = MutableSharedStrings::new();
        let styles = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();
        let parsed = crate::xlsx::parsers::worksheet_parser::parse_worksheet_data(&xml).unwrap();
        let view = &parsed.sheet_views[0];

        assert_eq!(view.workbook_view_id, Some(3));
        assert_eq!(view.tab_selected, Some(false));
        assert_eq!(view.window_protection, Some(true));
        assert_eq!(view.view_type, Some(crate::xlsx::SheetViewType::PageLayout));
        assert_eq!(view.top_left_cell.as_deref(), Some("C4"));
        assert_eq!(view.color_id, Some(12));
        assert_eq!(view.zoom_scale_sheet_layout_view, Some(60));
        assert_eq!(view.zoom_scale_page_layout_view, Some(80));
        assert_eq!(view.pane.as_ref().unwrap().y_split, Some(3.5));
        assert_eq!(view.selections.len(), 1);
        assert_eq!(view.selections[0].active_cell.as_deref(), Some("D5"));
    }

    #[test]
    fn rejects_invalid_sheet_view_output() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        let mut shared_strings = MutableSharedStrings::new();
        let styles = HashMap::new();

        ws.set_sheet_view(SheetView {
            zoom_scale: Some(9),
            ..Default::default()
        });
        assert!(ws.to_xml(&mut shared_strings, &styles).is_err());

        ws.set_sheet_view(SheetView {
            top_left_cell: Some("A0".to_string()),
            ..Default::default()
        });
        assert!(ws.to_xml(&mut shared_strings, &styles).is_err());

        ws.set_sheet_view(SheetView {
            pane: Some(SheetPane {
                x_split: Some(f64::NAN),
                ..Default::default()
            }),
            ..Default::default()
        });
        assert!(ws.to_xml(&mut shared_strings, &styles).is_err());

        ws.set_sheet_view(SheetView {
            selections: vec![SheetSelection {
                sqref: Some("B2:A1".to_string()),
                ..Default::default()
            }],
            ..Default::default()
        });
        assert!(ws.to_xml(&mut shared_strings, &styles).is_err());

        ws.set_sheet_view(SheetView {
            pane: Some(SheetPane::default()),
            ..Default::default()
        });
        ws.freeze_panes(2, 1);
        assert!(ws.to_xml(&mut shared_strings, &styles).is_err());
    }

    #[test]
    fn generated_freeze_panes_round_trip() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.freeze_panes(2, 1);

        let mut shared_strings = MutableSharedStrings::new();
        let styles = HashMap::new();
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();
        let parsed = crate::xlsx::parsers::worksheet_parser::parse_worksheet_data(&xml).unwrap();
        let view = &parsed.sheet_views[0];
        let pane = view.pane.as_ref().unwrap();

        assert_eq!(pane.x_split, Some(1.0));
        assert_eq!(pane.y_split, Some(2.0));
        assert_eq!(pane.top_left_cell.as_deref(), Some("B3"));
        assert_eq!(
            pane.active_pane,
            Some(crate::xlsx::SheetPanePosition::BottomRight)
        );
        assert_eq!(pane.state, Some(crate::xlsx::SheetPaneState::Frozen));
        assert_eq!(view.selections[0].active_cell.as_deref(), Some("B3"));
        assert_eq!(view.selections[0].sqref.as_deref(), Some("B3"));
    }

    #[test]
    fn fit_to_page_serializes_in_sheet_properties_order_and_round_trips() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_tab_color("FF336699");
        ws.set_automatic_page_breaks(true);
        ws.set_page_setup_with_options("landscape", 9, None, Some(1), Some(0))
            .unwrap();

        let mut shared_strings = MutableSharedStrings::new();
        let xml = ws.to_xml(&mut shared_strings, &HashMap::new()).unwrap();
        assert!(xml.contains(
            r#"<sheetPr><tabColor rgb="FF336699"/><pageSetUpPr autoPageBreaks="1" fitToPage="1"/></sheetPr>"#
        ));
        assert!(xml.contains(
            r#"<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>"#
        ));

        let properties = crate::xlsx::parse_worksheet_sheet_properties(xml.as_bytes())
            .unwrap()
            .unwrap();
        let page_setup = properties.page_setup_properties().unwrap();
        assert!(page_setup.automatic_page_breaks());
        assert!(page_setup.fit_to_page());
        assert_eq!(properties.tab_color().unwrap().argb(), Some([255, 51, 102, 153]));
    }

    #[test]
    fn fit_to_page_default_omission_explicit_false_and_clear_are_distinct() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        let mut shared_strings = MutableSharedStrings::new();
        let styles = HashMap::new();
        assert!(!ws
            .to_xml(&mut shared_strings, &styles)
            .unwrap()
            .contains("<sheetPr>"));

        ws.set_fit_to_page(false);
        assert_eq!(ws.fit_to_page(), Some(false));
        let xml = ws.to_xml(&mut shared_strings, &styles).unwrap();
        assert!(xml.contains(r#"<sheetPr><pageSetUpPr fitToPage="0"/></sheetPr>"#));

        ws.clear_fit_to_page();
        assert_eq!(ws.fit_to_page(), None);
        assert!(!ws
            .to_xml(&mut shared_strings, &styles)
            .unwrap()
            .contains("<sheetPr>"));
    }

    #[test]
    fn clearing_fit_mode_preserves_dimensions_and_invalid_update_is_atomic() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_page_setup_with_options("portrait", 9, None, Some(1), Some(2))
            .unwrap();
        ws.clear_fit_to_page();
        assert_eq!(ws.fit_to_page(), None);
        assert_eq!(ws.get_page_setup().unwrap().fit_to_width, Some(1));
        assert_eq!(ws.get_page_setup().unwrap().fit_to_height, Some(2));

        assert!(
            ws.set_page_setup_with_options("portrait", 9, None, Some(32_768), Some(1))
                .is_err()
        );
        assert_eq!(ws.get_page_setup().unwrap().fit_to_width, Some(1));
        assert_eq!(ws.get_page_setup().unwrap().fit_to_height, Some(2));
        assert_eq!(ws.fit_to_page(), None);
    }

    #[test]
    fn strict_namespace_parser_oracle_accepts_authored_page_setup_properties() {
        let mut ws = MutableWorksheet::new("Sheet1".to_string(), 1);
        ws.set_fit_to_page(true);
        ws.set_automatic_page_breaks(false);
        let mut shared_strings = MutableSharedStrings::new();
        let xml = ws
            .to_xml(&mut shared_strings, &HashMap::new())
            .unwrap()
            .replace(
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                "http://purl.oclc.org/ooxml/spreadsheetml/main",
            );

        let properties = crate::xlsx::parse_worksheet_sheet_properties(xml.as_bytes())
            .unwrap()
            .unwrap();
        let page_setup = properties.page_setup_properties().unwrap();
        assert!(!page_setup.automatic_page_breaks());
        assert!(page_setup.fit_to_page());
    }
}
