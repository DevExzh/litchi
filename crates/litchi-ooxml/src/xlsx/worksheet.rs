//! Worksheet implementation for Excel files.
//!
//! This module provides the concrete implementation of worksheets
//! for Excel (.xlsx) files.

use std::borrow::Cow;
use std::collections::{HashMap, HashSet};

use litchi_core::sheet::{
    Cell as CellTrait, CellIterator, CellValue, Result, RowIterator, Worksheet as WorksheetTrait,
};
use litchi_core::xml::unescape_xml;
use litchi_opc::Relationships;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};

use super::RichTextRun;
use super::cell::{Cell, CellIterator as XlsxCellIterator, RowIterator as XlsxRowIterator};
use super::comments::parse_comments_xml;
use super::format::{CellBorder, CellFill, CellFont, CellFormat};
use super::parsers::worksheet_parser;
use super::sort::SortState;
use super::sparkline::{SparklineGroup, parse_sparkline_groups_from_worksheet_xml};
use super::table::{Table, parse_table_xml};
use super::views::SheetView;

/// Information about a worksheet
#[derive(Debug, Clone)]
pub struct WorksheetInfo {
    /// Worksheet name
    pub name: String,
    /// Relationship ID for the worksheet
    pub relationship_id: String,
    /// Sheet ID
    pub sheet_id: u32,
    /// Whether this is the active sheet
    pub is_active: bool,
    pub print_area: Option<String>,
    pub repeating_rows: Option<String>,
    pub repeating_columns: Option<String>,
}

/// Column information
#[derive(Debug, Clone)]
pub struct ColumnInfo {
    /// Column width (in character units)
    pub width: Option<f64>,
    /// Whether the column is hidden
    pub hidden: bool,
    /// Custom width set
    pub custom_width: bool,
}

/// Row information
#[derive(Debug, Clone)]
pub struct RowInfo {
    /// Row height (in points)
    pub height: Option<f64>,
    /// Whether the row is hidden
    pub hidden: bool,
    /// Custom height set
    pub custom_height: bool,
}

/// Manual page break definition.
///
/// Excel stores page breaks as row/column break entries with an id and span.
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
    /// Whether the break was created for a pivot-table report.
    pub pivot: bool,
}

/// Hyperlink information
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Target URL or reference
    pub target: String,
    /// Optional location within the target document.
    pub location: Option<String>,
    /// Optional display text.
    pub display: Option<String>,
    /// Optional ScreenTip text.
    pub tooltip: Option<String>,
}

/// Cell comment information
#[derive(Debug, Clone)]
pub struct Comment {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Author of the comment
    pub author: Option<String>,
    /// Zero-based author index stored in the comments part.
    pub author_id: u32,
    /// Comment text
    pub text: String,
    /// Optional comment GUID.
    pub guid: Option<String>,
    /// Optional VML shape identifier.
    pub shape_id: Option<u32>,
}

/// Data validation rule information (parsed from worksheet XML)
#[derive(Debug, Clone)]
pub struct DataValidationRule {
    /// Range (e.g., "A1:B10")
    pub range: String,
    /// Validation type (e.g., "list", "whole", "decimal")
    pub validation_type: String,
    /// Comparison operator, when applicable.
    pub operator: Option<String>,
    /// First validation formula.
    pub formula: Option<String>,
    /// Second validation formula, when applicable.
    pub formula2: Option<String>,
    /// Whether blank cells are allowed.
    pub allow_blank: bool,
    /// Whether the in-cell dropdown arrow is hidden.
    pub show_drop_down: bool,
    /// Whether the input message is shown.
    pub show_input_message: bool,
    /// Whether the error message is shown.
    pub show_error_message: bool,
    /// Error alert style.
    pub error_style: Option<String>,
    /// IME mode.
    pub ime_mode: Option<String>,
    /// Error alert title.
    pub error_title: Option<String>,
    /// Error alert text.
    pub error: Option<String>,
    /// Input prompt title.
    pub prompt_title: Option<String>,
    /// Input prompt text.
    pub prompt: Option<String>,
}

/// Conditional formatting rule
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    /// Range (e.g., "A1:B10")
    pub range: String,
    /// Rule type (e.g., "cellIs", "colorScale")
    pub rule_type: String,
    /// Priority
    pub priority: u32,
}

/// Page setup information
#[derive(Debug, Clone, Default)]
pub struct PageSetup {
    /// Paper size (e.g., 9 for A4)
    pub paper_size: Option<u32>,
    /// Orientation (true = landscape, false = portrait)
    pub landscape: bool,
    /// Scale percentage
    pub scale: Option<u32>,
    /// Fit to page width
    pub fit_to_width: Option<u32>,
    /// Fit to page height
    pub fit_to_height: Option<u32>,
}

/// Auto-filter information
#[derive(Debug, Clone)]
pub struct AutoFilter {
    /// Optional range to which the auto-filter applies (e.g., "A1:D10").
    pub range: Option<String>,
    /// Optional sort state associated with the auto-filter.
    pub sort_state: Option<SortState>,
}

/// Concrete implementation of the Worksheet trait for Excel files.
pub struct Worksheet<'a> {
    /// Reference to the parent workbook
    workbook: &'a Workbook,
    /// Worksheet information
    info: WorksheetInfo,
    /// Cached cell data (row -> column -> value)
    cells: HashMap<u32, HashMap<u32, CellValue>>,
    /// Cell style indices (row -> column -> style_index)
    cell_styles: HashMap<u32, HashMap<u32, u32>>,
    /// Dimensions of the worksheet (min_row, min_col, max_row, max_col)
    dimensions: Option<(u32, u32, u32, u32)>,
    /// Merged cell ranges (start_row, start_col, end_row, end_col)
    merged_regions: Vec<(u32, u32, u32, u32)>,
    /// Hyperlinks by cell reference
    hyperlinks: HashMap<String, Hyperlink>,
    /// Comments by cell reference
    comments: HashMap<String, Comment>,
    /// Column information by column number
    columns: HashMap<u32, ColumnInfo>,
    /// Row information by row number
    rows: HashMap<u32, RowInfo>,
    /// Data validations
    data_validations: Vec<DataValidationRule>,
    /// Conditional formatting rules
    conditional_formats: Vec<ConditionalFormatRule>,
    /// Page setup
    page_setup: PageSetup,
    /// Auto-filter
    auto_filter: Option<AutoFilter>,
    /// Sheet view settings for each workbook window.
    sheet_views: Vec<SheetView>,
    /// Manual row page breaks
    row_breaks: Vec<PageBreak>,
    /// Manual column page breaks
    col_breaks: Vec<PageBreak>,
    rich_text_cells: HashMap<(u32, u32), Vec<RichTextRun>>,
    sparkline_groups: Vec<SparklineGroup>,
    tables: Vec<Table>,
}

impl<'a> Worksheet<'a> {
    /// Create a new worksheet.
    pub fn new(workbook: &'a Workbook, info: WorksheetInfo) -> Self {
        Self {
            workbook,
            info,
            cells: HashMap::new(),
            cell_styles: HashMap::new(),
            dimensions: None,
            merged_regions: Vec::new(),
            hyperlinks: HashMap::new(),
            comments: HashMap::new(),
            columns: HashMap::new(),
            rows: HashMap::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            page_setup: PageSetup::default(),
            auto_filter: None,
            sheet_views: Vec::new(),
            row_breaks: Vec::new(),
            col_breaks: Vec::new(),
            rich_text_cells: HashMap::new(),
            sparkline_groups: Vec::new(),
            tables: Vec::new(),
        }
    }

    /// Load worksheet data from the XML.
    pub fn load_data(&mut self) -> Result<()> {
        let worksheet_uri = self.workbook.worksheet_part_uri(&self.info)?;
        let worksheet_part = self.workbook.package().get_part(&worksheet_uri)?;
        let content = std::str::from_utf8(worksheet_part.blob())?;

        // Parse worksheet data
        self.parse_worksheet_xml(content, worksheet_part.rels())?;

        Ok(())
    }

    /// Parse worksheet XML to extract cell data.
    fn parse_worksheet_xml(&mut self, content: &str, relationships: &Relationships) -> Result<()> {
        let (hyperlinks, table_relationship_ids) = self.parse_sheet_data(content)?;
        self.resolve_hyperlinks(hyperlinks, relationships)?;
        self.load_tables(table_relationship_ids, relationships)?;
        self.load_comments(relationships)?;

        self.sparkline_groups = parse_sparkline_groups_from_worksheet_xml(content)?;

        Ok(())
    }

    fn load_comments(&mut self, relationships: &Relationships) -> Result<()> {
        let mut matching = relationships.iter().filter(|relationship| {
            matches!(relationship.reltype(), rt::COMMENTS | rt::STRICT_COMMENTS)
        });
        let Some(relationship) = matching.next() else {
            self.comments.clear();
            return Ok(());
        };
        if matching.next().is_some() {
            return Err("Worksheet has multiple comments relationships".into());
        }
        if relationship.is_external() {
            return Err("Worksheet comments relationship cannot be external".into());
        }
        let comments_uri = relationship.target_partname()?;
        let comments_part = self.workbook.package().get_part(&comments_uri)?;
        if comments_part.content_type() != ct::SML_COMMENTS {
            return Err(format!(
                "Worksheet comments part '{}' has content type '{}', expected '{}'",
                comments_uri,
                comments_part.content_type(),
                ct::SML_COMMENTS
            )
            .into());
        }
        let content = std::str::from_utf8(comments_part.blob())?;
        self.comments = parse_comments_xml(content)?;
        Ok(())
    }

    pub fn sparkline_groups(&self) -> &[SparklineGroup] {
        &self.sparkline_groups
    }

    /// Structured tables defined on this worksheet.
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Parse sheetData content.
    fn parse_sheet_data(
        &mut self,
        sheet_data: &str,
    ) -> Result<(Vec<worksheet_parser::ParsedHyperlink>, Vec<String>)> {
        let parsed = worksheet_parser::parse_worksheet_data(sheet_data)?;
        self.cells = parsed.cells;
        self.cell_styles = parsed.cell_styles;
        self.rows = parsed.rows;
        self.rich_text_cells = parsed.rich_text_cells;
        self.merged_regions = parsed.merged_regions;
        self.columns = parsed.columns;
        self.data_validations = parsed.data_validations;
        self.conditional_formats = parsed.conditional_formats;
        self.page_setup = parsed.page_setup.unwrap_or_default();
        self.row_breaks = parsed.row_breaks;
        self.col_breaks = parsed.col_breaks;
        self.auto_filter = parsed.auto_filter;
        self.sheet_views = parsed.sheet_views;
        self.dimensions = parsed.dimensions;
        Ok((parsed.hyperlinks, parsed.table_relationship_ids))
    }

    fn load_tables(
        &mut self,
        relationship_ids: Vec<String>,
        relationships: &Relationships,
    ) -> Result<()> {
        let mut tables = Vec::with_capacity(relationship_ids.len());
        let mut table_ids = HashSet::with_capacity(relationship_ids.len());
        let mut table_names = HashSet::with_capacity(relationship_ids.len());
        for relationship_id in relationship_ids {
            let relationship = relationships.get(&relationship_id).ok_or_else(|| {
                format!("Worksheet tablePart references missing relationship '{relationship_id}'")
            })?;
            if !matches!(relationship.reltype(), rt::TABLE | rt::STRICT_TABLE) {
                return Err(format!(
                    "Worksheet tablePart relationship '{relationship_id}' has invalid type '{}'",
                    relationship.reltype()
                )
                .into());
            }
            if relationship.is_external() {
                return Err(format!(
                    "Worksheet tablePart relationship '{relationship_id}' cannot be external"
                )
                .into());
            }
            let table_uri = relationship.target_partname()?;
            let table_part = self.workbook.package().get_part(&table_uri)?;
            if table_part.content_type() != ct::SML_TABLE {
                return Err(format!(
                    "Worksheet table part '{table_uri}' has content type '{}', expected '{}'",
                    table_part.content_type(),
                    ct::SML_TABLE
                )
                .into());
            }
            let xml = std::str::from_utf8(table_part.blob())?;
            let table = parse_table_xml(xml)?
                .ok_or_else(|| format!("Table part '{table_uri}' has no table root"))?;
            if !table_ids.insert(table.id) {
                return Err(format!("Duplicate worksheet table ID {}", table.id).into());
            }
            if !table_names.insert(table.display_name.to_ascii_lowercase()) {
                return Err(format!(
                    "Duplicate worksheet table display name '{}'",
                    table.display_name
                )
                .into());
            }
            tables.push(table);
        }
        self.tables = tables;
        Ok(())
    }

    /// Parse a single row XML.
    #[allow(clippy::type_complexity)]
    #[allow(dead_code)]
    fn parse_row_xml(
        &self,
        row_content: &str,
    ) -> Result<
        Option<(
            u32,
            Option<RowInfo>,
            Vec<(u32, CellValue, Option<u32>, Option<Vec<RichTextRun>>)>,
        )>,
    > {
        // Extract row number
        let row_num = if let Some(r_start) = row_content.find("r=\"") {
            let r_content = &row_content[r_start + 3..];
            if let Some(quote_pos) = r_content.find('"') {
                r_content[..quote_pos].parse::<u32>().ok()
            } else {
                None
            }
        } else {
            None
        };

        let row_num = match row_num {
            Some(r) => r,
            None => return Ok(None),
        };

        // Extract row height and hidden status
        let height = if let Some(ht_start) = row_content.find("ht=\"") {
            let ht_content = &row_content[ht_start + 4..];
            ht_content
                .find('"')
                .and_then(|quote_pos| ht_content[..quote_pos].parse::<f64>().ok())
        } else {
            None
        };

        let hidden = row_content.contains("hidden=\"1\"");
        let custom_height = row_content.contains("customHeight=\"1\"");

        let row_info = if height.is_some() || hidden || custom_height {
            Some(RowInfo {
                height,
                hidden,
                custom_height,
            })
        } else {
            None
        };

        let mut cells = Vec::new();

        // Parse cells in this row
        let mut pos = 0;
        while let Some(c_start) = row_content[pos..].find("<c ") {
            let c_start_pos = pos + c_start;
            if let Some(c_end) = row_content[c_start_pos..].find("</c>") {
                let c_content = &row_content[c_start_pos..c_start_pos + c_end + 4];

                if let Some((col_num, value, style_idx, rich_runs)) =
                    self.parse_cell_xml(c_content)?
                {
                    cells.push((col_num, value, style_idx, rich_runs));
                }

                pos = c_start_pos + c_end + 4;
            } else {
                break;
            }
        }

        Ok(Some((row_num, row_info, cells)))
    }

    /// Parse a single cell XML.
    #[allow(clippy::type_complexity)] // TODO: Refactor the return type
    fn parse_cell_xml(
        &self,
        cell_content: &str,
    ) -> Result<Option<(u32, CellValue, Option<u32>, Option<Vec<RichTextRun>>)>> {
        // Extract cell reference (e.g., "A1")
        let reference = if let Some(r_start) = cell_content.find("r=\"") {
            let r_content = &cell_content[r_start + 3..];
            r_content
                .find('"')
                .map(|quote_pos| r_content[..quote_pos].to_string())
        } else {
            None
        };

        let reference = match reference {
            Some(r) => r,
            None => return Ok(None),
        };

        // Convert reference to row/col numbers
        let (col_num, _row_num) = Cell::reference_to_coords(&reference)?;

        // Extract style index (s attribute)
        let style_idx = if let Some(s_start) = cell_content.find(" s=\"") {
            let s_content = &cell_content[s_start + 4..];
            s_content
                .find('"')
                .and_then(|quote_pos| s_content[..quote_pos].parse::<u32>().ok())
        } else {
            None
        };

        // Extract cell type
        let cell_type = if let Some(t_start) = cell_content.find("t=\"") {
            let t_content = &cell_content[t_start + 3..];
            t_content
                .find('"')
                .map(|quote_pos| t_content[..quote_pos].to_string())
        } else {
            None
        };

        // Inline strings: text is stored inside <is><t>...</t></is> instead of <v>.
        if matches!(cell_type.as_deref(), Some("inlineStr")) || cell_content.contains("<is>") {
            let text = Self::extract_inline_string_text(cell_content).unwrap_or_default();
            let rich_runs = Self::extract_inline_rich_text_runs(cell_content);
            return Ok(Some((
                col_num,
                CellValue::String(text),
                style_idx,
                rich_runs,
            )));
        }

        // Extract formula text (if present) from <f>...</f> and capture
        // array formula attributes when available.
        let mut is_array_formula = false;
        let mut array_ref: Option<String> = None;
        let formula_text = if let Some(f_start) = cell_content.find("<f") {
            let f_content = &cell_content[f_start..];
            if let Some(gt_rel) = f_content.find('>') {
                let tag_end = f_start + gt_rel + 1;
                let f_tag = &cell_content[f_start..tag_end];

                // Detect array formulas: <f t="array" ref="A1:C3">...
                if f_tag.contains("t=\"array\"") {
                    is_array_formula = true;
                }
                if let Some(r) = Self::extract_attribute(f_tag, "ref") {
                    array_ref = Some(r);
                }

                let text_start = tag_end;
                if let Some(end_rel) = cell_content[text_start..].find("</f>") {
                    let raw = &cell_content[text_start..text_start + end_rel];
                    Some(unescape_xml(raw))
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        // Extract value from <v> for normal cells and cached formula results.
        let value = if let Some(v_start) = cell_content.find("<v>") {
            let v_start_pos = v_start + 3;
            cell_content[v_start_pos..]
                .find("</v>")
                .map(|v_end| cell_content[v_start_pos..v_start_pos + v_end].to_string())
        } else {
            None
        };

        // Base value used either as the direct cell value (non-formula) or as the
        // cached result for formula cells.
        let base_value = match (cell_type.as_deref(), value.as_deref()) {
            (Some("str"), Some(v)) => CellValue::String(v.to_string()),
            (Some("s"), Some(v)) => {
                // Shared string reference - parse index and resolve later
                CellValue::String(format!("SHARED_STRING_{}", v))
            },
            (Some("b"), Some(v)) => match v {
                "1" => CellValue::Bool(true),
                "0" => CellValue::Bool(false),
                _ => CellValue::Error("Invalid boolean value".to_string()),
            },
            (_, Some(v)) => {
                // Try to parse as number - use fast parsing
                if let Ok(int_val) = atoi_simd::parse::<_, false, false>(v.as_bytes()) {
                    CellValue::Int(int_val)
                } else if let Ok(float_val) = fast_float2::parse(v) {
                    CellValue::Float(float_val)
                } else {
                    CellValue::String(v.to_string())
                }
            },
            _ => CellValue::Empty,
        };

        let cell_value = if let Some(formula) = formula_text {
            // For formula cells, wrap the parsed formula together with the cached
            // value (if any). A missing <v> is represented as None.
            let cached_value = match base_value {
                CellValue::Empty => None,
                other => Some(Box::new(other)),
            };
            CellValue::Formula {
                formula,
                cached_value,
                is_array: is_array_formula,
                array_range: array_ref,
            }
        } else {
            base_value
        };

        Ok(Some((col_num, cell_value, style_idx, None)))
    }

    /// Extract concatenated text from an inline string cell (<is> ... </is>).
    fn extract_inline_string_text(cell_content: &str) -> Option<String> {
        let bytes = cell_content.as_bytes();
        let mut result = String::new();
        let mut search_start = 0;

        // If there is an <is> element, start after it to avoid matching any
        // unrelated <t> elements (e.g., inside formulas).
        if let Some(is_start) = cell_content.find("<is") {
            let after_is = &bytes[is_start..];
            if let Some(gt_rel) = memchr::memchr(b'>', after_is) {
                search_start = is_start + gt_rel + 1;
            }
        }

        // Concatenate text from all <t> elements within the inline string.
        while let Some(rel_pos) = memchr::memmem::find(&bytes[search_start..], b"<t") {
            let t_pos = search_start + rel_pos;
            let after_t = &bytes[t_pos..];
            let gt_rel = match memchr::memchr(b'>', after_t) {
                Some(p) => p,
                None => break,
            };
            let text_start = t_pos + gt_rel + 1;

            let after_text = &bytes[text_start..];
            if let Some(end_rel) = memchr::memmem::find(after_text, b"</t>") {
                let text_end = text_start + end_rel;
                result.push_str(&cell_content[text_start..text_end]);
                search_start = text_end + 4; // len("</t>")
            } else {
                break;
            }
        }

        if result.is_empty() {
            None
        } else {
            Some(result)
        }
    }

    fn extract_inline_rich_text_runs(cell_content: &str) -> Option<Vec<RichTextRun>> {
        let bytes = cell_content.as_bytes();
        let mut search_start = 0;

        if let Some(is_start) = cell_content.find("<is") {
            let after_is = &bytes[is_start..];
            if let Some(gt_rel) = memchr::memchr(b'>', after_is) {
                search_start = is_start + gt_rel + 1;
            }
        }

        let mut runs: Vec<RichTextRun> = Vec::new();
        let mut pos = search_start;
        let slice = &bytes;

        while let Some(rel) = memchr::memmem::find(&slice[pos..], b"<r") {
            let r_pos = pos + rel;
            let next = slice.get(r_pos + 2).copied();
            if !matches!(next, Some(b'>') | Some(b' ')) {
                pos = r_pos + 2;
                continue;
            }

            let after_r = &slice[r_pos..];
            let gt_rel = match memchr::memchr(b'>', after_r) {
                Some(p) => p,
                None => break,
            };
            let inner_start = r_pos + gt_rel + 1;
            let after_inner = &slice[inner_start..];
            let end_rel = match memchr::memmem::find(after_inner, b"</r>") {
                Some(p) => p,
                None => break,
            };
            let inner_end = inner_start + end_rel;
            let run_inner = &cell_content[inner_start..inner_end];

            if let Some(run) = Self::parse_rich_text_run(run_inner) {
                runs.push(run);
            }

            pos = inner_end + 4; // len("</r>")
        }

        if runs.is_empty() {
            // Fallback: treat entire inline string as single run if we have text but no <r>.
            if let Some(text) = Self::extract_inline_string_text(cell_content)
                && !text.is_empty()
            {
                runs.push(RichTextRun {
                    text,
                    font_name: None,
                    font_size: None,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                });
            }
        }

        if runs.is_empty() { None } else { Some(runs) }
    }

    fn parse_rich_text_run(content: &str) -> Option<RichTextRun> {
        let bytes = content.as_bytes();
        let mut text = String::new();
        let mut search_start = 0;

        while let Some(rel_pos) = memchr::memmem::find(&bytes[search_start..], b"<t") {
            let t_pos = search_start + rel_pos;
            let after_t = &bytes[t_pos..];
            let gt_rel = match memchr::memchr(b'>', after_t) {
                Some(p) => p,
                None => break,
            };
            let text_start = t_pos + gt_rel + 1;

            let after_text = &bytes[text_start..];
            if let Some(end_rel) = memchr::memmem::find(after_text, b"</t>") {
                let text_end = text_start + end_rel;
                text.push_str(&content[text_start..text_end]);
                search_start = text_end + 4;
            } else {
                break;
            }
        }

        if text.is_empty() {
            return None;
        }

        let mut font_name: Option<String> = None;
        let mut font_size: Option<f64> = None;
        let mut bold = false;
        let mut italic = false;
        let mut underline = false;
        let mut color: Option<String> = None;

        if let Some(rpr_start) = content.find("<rPr") {
            let rpr_bytes = &bytes[rpr_start..];
            if let Some(rpr_end_rel) = memchr::memmem::find(rpr_bytes, b"</rPr>") {
                let rpr_end = rpr_start + rpr_end_rel + "</rPr>".len();
                let rpr_content = &content[rpr_start..rpr_end];

                if let Some(pos) = rpr_content.find("<rFont")
                    && let Some(val_pos) = rpr_content[pos..].find("val=\"")
                {
                    let start = pos + val_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"') {
                        font_name = Some(rpr_content[start..start + end_rel].to_string());
                    }
                }

                if let Some(pos) = rpr_content.find("<sz")
                    && let Some(val_pos) = rpr_content[pos..].find("val=\"")
                {
                    let start = pos + val_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"')
                        && let Ok(sz) = rpr_content[start..start + end_rel].parse::<f64>()
                    {
                        font_size = Some(sz);
                    }
                }

                if rpr_content.contains("<b/") || rpr_content.contains("<b ") {
                    bold = true;
                }
                if rpr_content.contains("<i/") || rpr_content.contains("<i ") {
                    italic = true;
                }
                if rpr_content.contains("<u/") || rpr_content.contains("<u ") {
                    underline = true;
                }

                if let Some(pos) = rpr_content.find("<color")
                    && let Some(rgb_pos) = rpr_content[pos..].find("rgb=\"")
                {
                    let start = pos + rgb_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"') {
                        color = Some(rpr_content[start..start + end_rel].to_string());
                    }
                }
            }
        }

        Some(RichTextRun {
            text,
            font_name,
            font_size,
            bold,
            italic,
            underline,
            color,
        })
    }

    fn resolve_hyperlinks(
        &mut self,
        hyperlinks: Vec<worksheet_parser::ParsedHyperlink>,
        relationships: &Relationships,
    ) -> Result<()> {
        use litchi_opc::constants::relationship_type as rt;

        for hyperlink in hyperlinks {
            let target = if let Some(relationship_id) = hyperlink.relationship_id.as_deref() {
                let relationship = relationships.get(relationship_id).ok_or_else(|| {
                    format!(
                        "Worksheet hyperlink '{}' references missing relationship '{}'",
                        hyperlink.cell_ref, relationship_id
                    )
                })?;
                if relationship.reltype() != rt::HYPERLINK
                    && relationship.reltype() != rt::STRICT_HYPERLINK
                {
                    return Err(format!(
                        "Relationship '{}' for worksheet hyperlink '{}' has invalid type '{}'",
                        relationship_id,
                        hyperlink.cell_ref,
                        relationship.reltype()
                    )
                    .into());
                }
                if !relationship.is_external() {
                    return Err(format!(
                        "Relationship '{}' for worksheet hyperlink '{}' is not external",
                        relationship_id, hyperlink.cell_ref
                    )
                    .into());
                }
                relationship.target_ref().to_string()
            } else {
                hyperlink.location.clone().ok_or_else(|| {
                    format!("Worksheet hyperlink '{}' has no target", hyperlink.cell_ref)
                })?
            };
            self.hyperlinks.insert(
                hyperlink.cell_ref.clone(),
                Hyperlink {
                    cell_ref: hyperlink.cell_ref,
                    target,
                    location: hyperlink.location,
                    display: hyperlink.display,
                    tooltip: hyperlink.tooltip,
                },
            );
        }
        Ok(())
    }

    /// Helper method to extract attribute value from XML tag.
    fn extract_attribute(tag: &str, attr: &str) -> Option<String> {
        let search_str = format!("{}=\"", attr);
        if let Some(start) = tag.find(&search_str) {
            let value_start = start + search_str.len();
            tag[value_start..]
                .find('"')
                .map(|end| tag[value_start..value_start + end].to_string())
        } else {
            None
        }
    }

    /// Get cell value at specific coordinates.
    fn get_cell_value(&self, row: u32, col: u32) -> CellValue {
        match self.cells.get(&row).and_then(|row_data| row_data.get(&col)) {
            Some(cell_value) => self.resolve_shared_string(cell_value.clone()),
            None => CellValue::Empty,
        }
    }

    /// Resolve shared string references to actual string values.
    fn resolve_shared_string(&self, cell_value: CellValue) -> CellValue {
        match cell_value {
            CellValue::String(s) if s.starts_with("SHARED_STRING_") => {
                // Extract the index from the shared string reference
                if let Some(index_str) = s.strip_prefix("SHARED_STRING_")
                    && let Ok(index) = atoi_simd::parse::<_, false, false>(index_str.as_bytes())
                    && let Some(shared_string) = self.workbook.shared_strings().get(index)
                {
                    return CellValue::String(shared_string.to_string());
                }
                CellValue::Error("Invalid shared string reference".to_string())
            },
            CellValue::Formula {
                formula,
                cached_value,
                is_array,
                array_range,
            } => {
                let resolved_cached =
                    cached_value.map(|boxed| Box::new(self.resolve_shared_string(*boxed)));
                CellValue::Formula {
                    formula,
                    cached_value: resolved_cached,
                    is_array,
                    array_range,
                }
            },
            other => other,
        }
    }

    /// Get all cells in a specific column.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get all values in column A (column 1)
    /// let column_values = ws.column_values(1)?;
    /// for value in column_values {
    ///     println!("{:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn column_values(&self, column: u32) -> Result<Vec<CellValue>> {
        let mut values = Vec::new();

        if let Some((min_row, _, max_row, _)) = self.dimensions {
            for row in min_row..=max_row {
                values.push(self.get_cell_value(row, column));
            }
        }

        Ok(values)
    }

    /// Get all cells in a specific row.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get all values in row 1
    /// let row_values = ws.row_values(1)?;
    /// for value in row_values {
    ///     println!("{:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn row_values(&self, row: u32) -> Result<Vec<CellValue>> {
        let mut values = Vec::new();

        if let Some((_, min_col, _, max_col)) = self.dimensions {
            for col in min_col..=max_col {
                values.push(self.get_cell_value(row, col));
            }
        }

        Ok(values)
    }

    /// Get a range of cells as a 2D vector.
    ///
    /// # Arguments
    /// * `start_row` - Starting row (1-based, inclusive)
    /// * `start_col` - Starting column (1-based, inclusive)
    /// * `end_row` - Ending row (1-based, inclusive)
    /// * `end_col` - Ending column (1-based, inclusive)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get range A1:C3
    /// let range = ws.range(1, 1, 3, 3)?;
    /// for row in range {
    ///     for cell in row {
    ///         print!("{:?} ", cell);
    ///     }
    ///     println!();
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn range(
        &self,
        start_row: u32,
        start_col: u32,
        end_row: u32,
        end_col: u32,
    ) -> Result<Vec<Vec<CellValue>>> {
        let mut result = Vec::new();

        for row in start_row..=end_row {
            let mut row_data = Vec::new();
            for col in start_col..=end_col {
                row_data.push(self.get_cell_value(row, col));
            }
            result.push(row_data);
        }

        Ok(result)
    }

    /// Get manual row page breaks.
    pub fn row_breaks(&self) -> &[PageBreak] {
        &self.row_breaks
    }

    /// Get manual column page breaks.
    pub fn col_breaks(&self) -> &[PageBreak] {
        &self.col_breaks
    }

    /// Get worksheet view settings, if present.
    pub fn sheet_view(&self) -> Option<&SheetView> {
        self.sheet_views.last()
    }

    /// Get all worksheet views, one for each workbook window.
    pub fn sheet_views(&self) -> &[SheetView] {
        &self.sheet_views
    }

    /// Find cells containing specific text.
    ///
    /// # Arguments
    /// * `query` - Text to search for
    ///
    /// # Returns
    /// Vector of (row, column) tuples where the cell contains the query text
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Find all cells containing "Total"
    /// let matches = ws.find_text("Total")?;
    /// for (row, col) in matches {
    ///     println!("Found at row {}, column {}", row, col);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn find_text(&self, query: &str) -> Result<Vec<(u32, u32)>> {
        let mut matches = Vec::new();

        for (&row, row_data) in &self.cells {
            for (&col, value) in row_data {
                if let CellValue::String(s) = &value
                    && (s.contains(query)
                        || (s.starts_with("SHARED_STRING_") && {
                            let resolved = self.resolve_shared_string(value.clone());
                            matches!(resolved, CellValue::String(ref text) if text.contains(query))
                        }))
                {
                    matches.push((row, col));
                }
            }
        }

        Ok(matches)
    }

    /// Get the used range dimensions (min row, min col, max row, max col).
    ///
    /// Returns None if the worksheet is empty.
    pub fn used_range(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    /// Check if a cell is empty.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_cell_empty(&self, row: u32, column: u32) -> bool {
        matches!(self.get_cell_value(row, column), CellValue::Empty)
    }

    /// Count non-empty cells in the worksheet.
    pub fn non_empty_cell_count(&self) -> usize {
        self.cells
            .values()
            .map(|row| {
                row.values()
                    .filter(|v| !matches!(v, CellValue::Empty))
                    .count()
            })
            .sum()
    }

    /// Get worksheet information.
    pub fn info(&self) -> &WorksheetInfo {
        &self.info
    }

    // ===== Cell Formatting (Reading) =====

    /// Get the cell style for a specific cell.
    ///
    /// Returns the style information including font, fill, border, and number format.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(style) = ws.get_cell_style(1, 1) {
    ///     println!("Cell has custom styling");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_cell_style(&self, row: u32, column: u32) -> Option<&super::styles::CellStyle> {
        self.cell_styles
            .get(&row)
            .and_then(|row_styles| row_styles.get(&column))
            .and_then(|style_idx| self.workbook.styles().get_cell_style(*style_idx as usize))
    }

    /// Get the complete cell format information for a specific cell.
    ///
    /// Returns a `CellFormat` with resolved font, fill, border, and number format.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_cell_format(&self, row: u32, column: u32) -> Option<CellFormat> {
        let style = self.get_cell_style(row, column)?;
        let styles = self.workbook.styles();

        let font = style
            .font_id
            .and_then(|id| styles.get_font(id as usize))
            .map(|f| CellFont {
                name: f.name.clone(),
                size: f.size,
                bold: f.bold,
                italic: f.italic,
                underline: f.underline.is_some(),
                color: f.color.clone(),
            });

        let fill = style
            .fill_id
            .and_then(|id| styles.get_fill(id as usize))
            .and_then(|f| match f {
                super::styles::Fill::Pattern {
                    pattern_type,
                    fg_color,
                    bg_color,
                } => {
                    // Map pattern type string to enum
                    let pattern_enum = match pattern_type.as_str() {
                        "solid" => super::format::CellFillPatternType::Solid,
                        "gray125" => super::format::CellFillPatternType::Gray125,
                        "darkGray" => super::format::CellFillPatternType::DarkGray,
                        "mediumGray" => super::format::CellFillPatternType::MediumGray,
                        "lightGray" => super::format::CellFillPatternType::LightGray,
                        "gray0625" => super::format::CellFillPatternType::Gray0625,
                        "darkHorizontal" => super::format::CellFillPatternType::DarkHorizontal,
                        "darkVertical" => super::format::CellFillPatternType::DarkVertical,
                        "darkDown" => super::format::CellFillPatternType::DarkDown,
                        "darkUp" => super::format::CellFillPatternType::DarkUp,
                        "darkGrid" => super::format::CellFillPatternType::DarkGrid,
                        "darkTrellis" => super::format::CellFillPatternType::DarkTrellis,
                        _ => super::format::CellFillPatternType::None,
                    };
                    Some(CellFill {
                        pattern_type: pattern_enum,
                        fg_color: fg_color.clone(),
                        bg_color: bg_color.clone(),
                    })
                },
                _ => None,
            });

        // Helper function to map border style string to enum
        let map_border_style = |s: &super::styles::BorderStyle| {
            let style = match s.style.as_str() {
                "thin" => super::format::CellBorderLineStyle::Thin,
                "medium" => super::format::CellBorderLineStyle::Medium,
                "dashed" => super::format::CellBorderLineStyle::Dashed,
                "dotted" => super::format::CellBorderLineStyle::Dotted,
                "thick" => super::format::CellBorderLineStyle::Thick,
                "double" => super::format::CellBorderLineStyle::Double,
                "hair" => super::format::CellBorderLineStyle::Hair,
                "mediumDashed" => super::format::CellBorderLineStyle::MediumDashed,
                "dashDot" => super::format::CellBorderLineStyle::DashDot,
                "mediumDashDot" => super::format::CellBorderLineStyle::MediumDashDot,
                "dashDotDot" => super::format::CellBorderLineStyle::DashDotDot,
                "mediumDashDotDot" => super::format::CellBorderLineStyle::MediumDashDotDot,
                "slantDashDot" => super::format::CellBorderLineStyle::SlantDashDot,
                "none" => super::format::CellBorderLineStyle::None,
                _ => super::format::CellBorderLineStyle::Thin, // Default fallback
            };
            super::format::CellBorderSide {
                style,
                color: s.color.clone(),
            }
        };

        let border = style
            .border_id
            .and_then(|id| styles.get_border(id as usize))
            .map(|b| CellBorder {
                left: b.left.as_ref().map(&map_border_style),
                right: b.right.as_ref().map(&map_border_style),
                top: b.top.as_ref().map(&map_border_style),
                bottom: b.bottom.as_ref().map(&map_border_style),
                diagonal: b.diagonal.as_ref().map(&map_border_style),
            });

        let number_format = style
            .num_fmt_id
            .and_then(|id| styles.get_number_format(id))
            .map(|nf| nf.code.clone());

        Some(CellFormat {
            font,
            fill,
            border,
            number_format,
        })
    }

    /// Check if a cell is formatted as a date.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_date_formatted(&self, row: u32, column: u32) -> bool {
        if let Some(style) = self.get_cell_style(row, column)
            && let Some(num_fmt_id) = style.num_fmt_id
            && let Some(num_fmt) = self.workbook.styles().get_number_format(num_fmt_id)
        {
            return num_fmt.is_date_format();
        }
        false
    }

    /// Get the date/time value from a cell formatted as a date.
    ///
    /// Returns None if the cell is not a date or doesn't contain a numeric value.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_date_cell_value(&self, row: u32, column: u32) -> Option<f64> {
        if !self.is_date_formatted(row, column) {
            return None;
        }

        match self.get_cell_value(row, column) {
            CellValue::Float(f) => Some(f),
            CellValue::Int(i) => Some(i as f64),
            _ => None,
        }
    }

    // ===== Merged Regions =====

    /// Get all merged cell regions in the worksheet.
    ///
    /// Returns a slice of tuples (start_row, start_col, end_row, end_col).
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for (start_row, start_col, end_row, end_col) in ws.get_merged_regions() {
    ///     println!("Merged region: ({}, {}) to ({}, {})",
    ///              start_row, start_col, end_row, end_col);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_merged_regions(&self) -> &[(u32, u32, u32, u32)] {
        &self.merged_regions
    }

    /// Check if a cell is part of a merged region.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_merged_cell(&self, row: u32, column: u32) -> bool {
        self.get_merge_region(row, column).is_some()
    }

    /// Get the merged region that contains a specific cell.
    ///
    /// Returns None if the cell is not part of any merged region.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_merge_region(&self, row: u32, column: u32) -> Option<(u32, u32, u32, u32)> {
        self.merged_regions
            .iter()
            .find(|&&(sr, sc, er, ec)| row >= sr && row <= er && column >= sc && column <= ec)
            .copied()
    }

    // ===== Hyperlinks =====

    /// Get the hyperlink for a specific cell.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(hyperlink) = ws.get_hyperlink(1, 1) {
    ///     println!("Cell A1 links to: {}", hyperlink.target);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_hyperlink(&self, row: u32, column: u32) -> Option<&Hyperlink> {
        let cell_ref = format!("{}{}", Cell::column_to_letters(column), row);
        self.hyperlinks.get(&cell_ref).or_else(|| {
            self.hyperlinks
                .values()
                .find(|hyperlink| Self::hyperlink_contains(&hyperlink.cell_ref, row, column))
        })
    }

    fn hyperlink_contains(reference: &str, row: u32, column: u32) -> bool {
        let Some((start, end)) = reference.split_once(':') else {
            return false;
        };
        let Ok((start_column, start_row)) = Cell::reference_to_coords(start) else {
            return false;
        };
        let Ok((end_column, end_row)) = Cell::reference_to_coords(end) else {
            return false;
        };
        row >= start_row && row <= end_row && column >= start_column && column <= end_column
    }

    /// Get all hyperlinks in the worksheet.
    pub fn get_hyperlinks(&self) -> &HashMap<String, Hyperlink> {
        &self.hyperlinks
    }

    // ===== Comments =====

    /// Get the comment for a specific cell.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(comment) = ws.get_cell_comment(1, 1) {
    ///     println!("Comment: {}", comment.text);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_cell_comment(&self, row: u32, column: u32) -> Option<&Comment> {
        let cell_ref = format!("{}{}", Cell::column_to_letters(column), row);
        self.comments.get(&cell_ref)
    }

    /// Get all comments in the worksheet.
    pub fn get_comments(&self) -> &HashMap<String, Comment> {
        &self.comments
    }

    // ===== Column Operations =====

    /// Get the width of a specific column.
    ///
    /// Returns the width in Excel's character units, or None if using default width.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn get_column_width(&self, column: u32) -> Option<f64> {
        self.columns.get(&column).and_then(|info| info.width)
    }

    /// Check if a column is hidden.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn is_column_hidden(&self, column: u32) -> bool {
        self.columns.get(&column).is_some_and(|info| info.hidden)
    }

    /// Get column information.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn get_column_info(&self, column: u32) -> Option<&ColumnInfo> {
        self.columns.get(&column)
    }

    // ===== Row Operations =====

    /// Get the height of a specific row.
    ///
    /// Returns the height in points, or None if using default height.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn get_row_height(&self, row: u32) -> Option<f64> {
        self.rows.get(&row).and_then(|info| info.height)
    }

    /// Check if a row is hidden.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.rows.get(&row).is_some_and(|info| info.hidden)
    }

    /// Get row information.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn get_row_info(&self, row: u32) -> Option<&RowInfo> {
        self.rows.get(&row)
    }

    // ===== Data Validation =====

    /// Get all data validations in the worksheet.
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for validation in ws.get_data_validations() {
    ///     println!("Validation on range: {}", validation.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_data_validations(&self) -> &[DataValidationRule] {
        &self.data_validations
    }

    // ===== Conditional Formatting =====

    /// Get all conditional formatting rules in the worksheet.
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for rule in ws.get_conditional_formatting() {
    ///     println!("Conditional format on range: {}", rule.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_conditional_formatting(&self) -> &[ConditionalFormatRule] {
        &self.conditional_formats
    }

    // ===== Page Setup =====

    /// Get the page setup information.
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// let page_setup = ws.get_page_setup();
    /// if page_setup.landscape {
    ///     println!("Page is in landscape orientation");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_page_setup(&self) -> &PageSetup {
        &self.page_setup
    }

    // ===== Auto-Filter =====

    /// Get the auto-filter information.
    ///
    /// # Examples
    ///
    /// ```ignore
    /// use litchi_ooxml::xlsx::Workbook;
    /// use litchi_core::sheet::WorkbookTrait;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(auto_filter) = ws.get_auto_filter() {
    ///     println!("Auto-filter range: {:?}", auto_filter.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
    /// ```
    pub fn get_auto_filter(&self) -> Option<&AutoFilter> {
        self.auto_filter.as_ref()
    }

    pub fn get_print_area(&self) -> Option<&str> {
        self.info.print_area.as_deref()
    }

    pub fn get_repeating_rows(&self) -> Option<&str> {
        self.info.repeating_rows.as_deref()
    }

    pub fn get_repeating_columns(&self) -> Option<&str> {
        self.info.repeating_columns.as_deref()
    }

    pub fn get_rich_text_cell(&self, row: u32, column: u32) -> Option<&[RichTextRun]> {
        if let Some(runs) = self.rich_text_cells.get(&(row, column)) {
            return Some(runs.as_slice());
        }

        // Fallback: check if this cell (or its cached formula value) references a
        // shared string with rich text runs in the sharedStrings table.
        if let Some(row_data) = self.cells.get(&row)
            && let Some(raw) = row_data.get(&column)
        {
            // Helper to check a CellValue for a shared string reference.
            fn from_shared_string<'a>(
                workbook: &'a Workbook,
                value: &CellValue,
            ) -> Option<&'a [RichTextRun]> {
                if let CellValue::String(s) = value
                    && let Some(index_str) = s.strip_prefix("SHARED_STRING_")
                    && let Ok(index) = atoi_simd::parse::<_, false, false>(index_str.as_bytes())
                {
                    return workbook.shared_strings().rich_text_runs(index);
                }
                None
            }

            if let Some(runs) = from_shared_string(self.workbook, raw) {
                return Some(runs);
            }

            if let CellValue::Formula { cached_value, .. } = raw
                && let Some(cached) = cached_value.as_deref()
                && let Some(runs) = from_shared_string(self.workbook, cached)
            {
                return Some(runs);
            }
        }

        None
    }

    // Previously TODO: Apache POI worksheet-level features - NOW IMPLEMENTED:
    // ✅ Cell formatting (reading): get_cell_style(), get_cell_format()
    // ✅ Cell types (advanced): get_cell_type() via CellValue enum
    // ✅ Date cells: is_date_formatted(), get_date_cell_value()
    // ✅ Cell hyperlinks: get_hyperlink(), get_hyperlinks()
    // ✅ Cell comments: get_cell_comment(), get_comments()
    // ✅ Merged regions: get_merged_regions(), is_merged_cell(), get_merge_region()
    // ✅ Column operations: get_column_width(), is_column_hidden(), get_column_info()
    // ✅ Row operations: is_row_hidden(), get_row_height(), get_row_info()
    // ✅ Auto-filter: get_auto_filter()
    // ✅ Data validation: get_data_validations()
    // ✅ Conditional formatting: get_conditional_formatting()
    // ✅ Page setup: get_page_setup()
    //
    // Still TODO (writing operations and advanced features):
    // - Formula evaluation: evaluate_formula(), get_formula_evaluator()
    // - Array formulas: set_array_formula(), get_array_formulas()
    // - Rich text cells: get_rich_string_cell_value(), set_rich_text_string()
    // - Set operations: set_hyperlink(), remove_hyperlink(), set_cell_comment(), remove_cell_comment()
    // - Column/row mutations: auto_size_column(), set_column_hidden(), set_row_hidden(), set_row_height()
    // - Sheet protection: protect_sheet(), is_protected(), get_protection_info()
    // - Set operations: set_auto_filter(), add_validation_data()
    // - Set operations: set_fit_to_page(), set_header(), set_footer()
    // - Repeating rows/columns: set_repeating_rows(), set_repeating_columns()
}

impl<'a> WorksheetTrait for Worksheet<'a> {
    fn name(&self) -> &str {
        &self.info.name
    }

    fn row_count(&self) -> usize {
        self.dimensions
            .map(|(_, _, max_row, _)| max_row as usize)
            .unwrap_or(0)
    }

    fn column_count(&self) -> usize {
        self.dimensions
            .map(|(_, _, _, max_col)| max_col as usize)
            .unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    fn cell(&self, row: u32, column: u32) -> Result<Box<dyn CellTrait + '_>> {
        let value = self
            .cell_value(row, column)
            .unwrap_or(Cow::Borrowed(CellValue::EMPTY))
            .into_owned();
        Ok(Box::new(Cell::new(row, column, value)))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> Result<Box<dyn CellTrait + '_>> {
        let (col, row) = Cell::reference_to_coords(coordinate)?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        let mut cells = Vec::new();

        for (&row, row_data) in &self.cells {
            for (&col, value) in row_data {
                cells.push(Cell::new(row, col, value.clone()));
            }
        }

        Box::new(XlsxCellIterator::new(cells))
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        let mut rows = Vec::new();

        if let Some((min_row, min_col, max_row, max_col)) = self.dimensions {
            for row in min_row..=max_row {
                let mut row_data = Vec::new();
                for col in min_col..=max_col {
                    let value = self.get_cell_value(row, col).clone();
                    row_data.push(value);
                }
                rows.push(row_data);
            }
        }

        Box::new(XlsxRowIterator::new(rows))
    }

    fn row(&self, row_idx: usize) -> Result<Cow<'_, [CellValue]>> {
        if let Some((min_row, min_col, max_col)) =
            self.dimensions.map(|(mr, mc, _, mc2)| (mr, mc, mc2))
        {
            let mut row_data = Vec::new();
            let row_num = row_idx as u32 + min_row;

            for col in min_col..=max_col {
                row_data.push(self.get_cell_value(row_num, col));
            }

            Ok(Cow::Owned(row_data))
        } else {
            Ok(Cow::Borrowed(&[]))
        }
    }

    fn cell_value(&self, row: u32, column: u32) -> Result<Cow<'_, CellValue>> {
        // XLSX values need shared string resolution, so we return owned
        Ok(Cow::Owned(self.get_cell_value(row, column)))
    }
}

/// Iterator over worksheets in a workbook
pub struct WorksheetIterator<'a> {
    worksheets: Vec<WorksheetInfo>,
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> {
    /// Create a new worksheet iterator.
    pub fn new(worksheets: Vec<WorksheetInfo>, workbook: &'a Workbook) -> Self {
        Self {
            worksheets,
            workbook,
            index: 0,
        }
    }
}

impl<'a> litchi_core::sheet::WorksheetIterator<'a> for WorksheetIterator<'a> {
    fn next(&mut self) -> Option<Result<Box<dyn WorksheetTrait + 'a>>> {
        if self.index >= self.worksheets.len() {
            return None;
        }
        let info = &self.worksheets[self.index];
        let mut ws = Worksheet::new(self.workbook, info.clone());
        self.index += 1;
        if ws.load_data().is_ok() {
            Some(Ok(Box::new(ws)))
        } else {
            // If failed to load, return an error or skip?
            // The trait expects Option<Result<...>>.
            Some(Err("Failed to load worksheet".into()))
        }
    }
}

// Import Workbook from the workbook module
use super::workbook::Workbook;

#[cfg(test)]
mod tests {
    use super::Worksheet;

    #[test]
    fn extract_inline_string_single_t() {
        let xml = r#"<c r=\"A1\" t=\"inlineStr\"><is><t>Hello</t></is></c>"#;
        let text = Worksheet::extract_inline_string_text(xml).unwrap();
        assert_eq!(text, "Hello");
    }

    #[test]
    fn extract_inline_string_multiple_runs() {
        let xml =
            r#"<c r=\"A1\" t=\"inlineStr\"><is><r><t>Hello </t></r><r><t>World</t></r></is></c>"#;
        let text = Worksheet::extract_inline_string_text(xml).unwrap();
        assert_eq!(text, "Hello World");
    }
}
