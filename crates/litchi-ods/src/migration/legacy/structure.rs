//! ODS row and column structural metadata.

use litchi_core::{Error, Result, xml::escape_xml};
use std::ops::Range;

pub(crate) const MAX_EXPANDED_ROWS_PER_SHEET: usize = 1_048_576;
pub(crate) const MAX_EXPANDED_COLUMNS_PER_SHEET: usize = 1_048_576;
pub(crate) const MAX_TABLE_STRUCTURE_DEPTH: usize = 256;

/// Visibility state shared by ODF table rows and columns.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum TableVisibility {
    /// Display normally (the ODF default).
    #[default]
    Visible,
    /// Hidden manually.
    Collapse,
    /// Hidden by filtering.
    Filter,
}

impl TableVisibility {
    pub(crate) fn parse(value: &str) -> Result<Self> {
        match value {
            "visible" => Ok(Self::Visible),
            "collapse" => Ok(Self::Collapse),
            "filter" => Ok(Self::Filter),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table:visibility value '{value}'"
            ))),
        }
    }

    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::Visible => "visible",
            Self::Collapse => "collapse",
            Self::Filter => "filter",
        }
    }
}

/// Structural metadata for one logical ODF table column.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Column {
    /// Zero-based logical column index.
    pub index: usize,
    /// Direct table-column style reference.
    pub style_name: Option<String>,
    /// Default table-cell style for cells in this column.
    pub default_cell_style_name: Option<String>,
    /// Column visibility.
    pub visibility: TableVisibility,
}

/// Optional style-selection flags associated with an ODF table template.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SheetStyleUsage {
    /// Apply the template's first-row style.
    pub use_first_row_styles: Option<bool>,
    /// Apply the template's last-row style.
    pub use_last_row_styles: Option<bool>,
    /// Apply the template's first-column style.
    pub use_first_column_styles: Option<bool>,
    /// Apply the template's last-column style.
    pub use_last_column_styles: Option<bool>,
    /// Apply alternating row styles from the template.
    pub use_banding_row_styles: Option<bool>,
    /// Apply alternating column styles from the template.
    pub use_banding_column_styles: Option<bool>,
}

/// Sheet-level table style and template references.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct SheetStyle {
    /// Direct table-style reference.
    pub style_name: Option<String>,
    /// Table-template reference.
    pub template_name: Option<String>,
    /// Optional template component selection flags.
    pub usage: SheetStyleUsage,
}

/// Sheet printing controls and ODF cell-range addresses.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetPrintSettings {
    /// Whether the sheet participates in printing.
    pub printable: bool,
    /// ODF cell-range addresses printed for this sheet.
    pub ranges: Vec<String>,
}

impl Default for SheetPrintSettings {
    fn default() -> Self {
        Self {
            printable: true,
            ranges: Vec::new(),
        }
    }
}

impl SheetPrintSettings {
    /// Create validated sheet printing settings.
    pub fn new(printable: bool, ranges: Vec<String>) -> Result<Self> {
        validate_cell_range_addresses(&ranges)?;
        Ok(Self { printable, ranges })
    }
}

pub(crate) fn validate_sheet_print_settings(settings: &SheetPrintSettings) -> Result<()> {
    validate_cell_range_addresses(&settings.ranges)
}

pub(crate) fn write_sheet_formatting_attributes(
    out: &mut String,
    style: &SheetStyle,
    print: &SheetPrintSettings,
) -> Result<()> {
    if let Some(value) = &style.style_name {
        write_escaped_attribute(out, "table:style-name", value);
    }
    if let Some(value) = &style.template_name {
        write_escaped_attribute(out, "table:template-name", value);
    }
    write_optional_bool_attribute(
        out,
        "table:use-first-row-styles",
        style.usage.use_first_row_styles,
    );
    write_optional_bool_attribute(
        out,
        "table:use-last-row-styles",
        style.usage.use_last_row_styles,
    );
    write_optional_bool_attribute(
        out,
        "table:use-first-column-styles",
        style.usage.use_first_column_styles,
    );
    write_optional_bool_attribute(
        out,
        "table:use-last-column-styles",
        style.usage.use_last_column_styles,
    );
    write_optional_bool_attribute(
        out,
        "table:use-banding-rows-styles",
        style.usage.use_banding_row_styles,
    );
    write_optional_bool_attribute(
        out,
        "table:use-banding-columns-styles",
        style.usage.use_banding_column_styles,
    );
    if !print.printable {
        out.push_str(" table:print=\"false\"");
    }
    validate_sheet_print_settings(print)?;
    if !print.ranges.is_empty() {
        write_escaped_attribute(out, "table:print-ranges", &print.ranges.join(" "));
    }
    Ok(())
}

pub(crate) fn split_cell_range_addresses(value: &str) -> Result<Vec<String>> {
    let mut ranges = Vec::new();
    let mut current = String::new();
    let mut quoted = false;
    let mut chars = value.chars().peekable();
    while let Some(character) = chars.next() {
        if character == '\'' {
            current.push(character);
            if quoted && chars.peek() == Some(&'\'') {
                current.push(chars.next().expect("peeked quote is present"));
            } else {
                quoted = !quoted;
            }
        } else if character.is_whitespace() && !quoted {
            if !current.is_empty() {
                ranges.push(std::mem::take(&mut current));
            }
        } else {
            current.push(character);
        }
    }
    if quoted {
        return Err(Error::InvalidFormat(
            "table print range contains an unterminated quoted sheet name".to_string(),
        ));
    }
    if !current.is_empty() {
        ranges.push(current);
    }
    Ok(ranges)
}

pub(crate) fn validate_cell_range_addresses(ranges: &[String]) -> Result<()> {
    for range in ranges {
        let parsed = split_cell_range_addresses(range)?;
        if range != range.trim() || parsed.len() != 1 || parsed[0].as_str() != range.as_str() {
            return Err(Error::InvalidFormat(format!(
                "invalid individual table cell range '{range}'"
            )));
        }
    }
    Ok(())
}

fn write_escaped_attribute(out: &mut String, name: &str, value: &str) {
    out.push(' ');
    out.push_str(name);
    out.push_str("=\"");
    out.push_str(&escape_xml(value));
    out.push('"');
}

fn write_optional_bool_attribute(out: &mut String, name: &str, value: Option<bool>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str(if value { "=\"true\"" } else { "=\"false\"" });
    }
}

/// A half-open range of logical rows or columns.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct TableRange {
    /// First included zero-based logical index.
    pub start: usize,
    /// First excluded zero-based logical index.
    pub end: usize,
}

impl TableRange {
    /// Create a non-empty half-open table range.
    pub fn new(start: usize, end: usize) -> Result<Self> {
        if start >= end {
            return Err(Error::InvalidFormat(format!(
                "table structure range {start}..{end} must be non-empty"
            )));
        }
        Ok(Self { start, end })
    }

    fn as_range(self) -> Range<usize> {
        self.start..self.end
    }
}

/// A recursively nested group of adjacent rows or columns.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TableGroup {
    /// Whether the group is expanded/displayed.
    pub display: bool,
    /// Ordered ranges and nested groups contained by this group.
    pub children: Vec<TableStructure>,
}

/// One structural run in an ODF table's row or column layout.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum TableStructure {
    /// An ordinary run of direct rows or columns.
    Range(TableRange),
    /// A run repeated as table headers when the table is printed.
    Header(TableRange),
    /// A potentially collapsed nested group.
    Group(TableGroup),
}

#[derive(Clone, Copy)]
pub(crate) enum TableStructureAxis {
    Rows,
    Columns,
}

impl TableStructureAxis {
    fn group_tag(self) -> &'static str {
        match self {
            Self::Rows => "table:table-row-group",
            Self::Columns => "table:table-column-group",
        }
    }

    fn header_tag(self) -> &'static str {
        match self {
            Self::Rows => "table:table-header-rows",
            Self::Columns => "table:table-header-columns",
        }
    }
}

pub(crate) fn write_table_structure(
    out: &mut String,
    structure: &[TableStructure],
    total: usize,
    axis: TableStructureAxis,
    mut write_range: impl FnMut(&mut String, Range<usize>),
) -> Result<()> {
    let mut cursor = 0usize;
    for entry in structure {
        let (start, end) = structure_bounds(entry, total, 0)?;
        if start < cursor {
            return Err(Error::InvalidFormat(
                "table structure ranges overlap or are out of order".to_string(),
            ));
        }
        if start > cursor {
            write_range(out, cursor..start);
        }
        write_structure_entry(out, entry, total, axis, 0, &mut write_range)?;
        cursor = end;
    }
    if cursor < total {
        write_range(out, cursor..total);
    }
    Ok(())
}

pub(crate) fn validate_table_structure(
    structure: &[TableStructure],
    axis: TableStructureAxis,
) -> Result<usize> {
    let total = structure.iter().try_fold(0usize, |maximum, entry| {
        structure_declared_end(entry, 0).map(|end| maximum.max(end))
    })?;
    write_table_structure(&mut String::new(), structure, total, axis, |_, _| {})?;
    Ok(total)
}

fn structure_declared_end(entry: &TableStructure, depth: usize) -> Result<usize> {
    ensure_structure_depth(depth)?;
    match entry {
        TableStructure::Range(range) | TableStructure::Header(range) => Ok(range.end),
        TableStructure::Group(group) => group
            .children
            .last()
            .ok_or_else(|| {
                Error::InvalidFormat("table groups must contain at least one item".to_string())
            })
            .and_then(|entry| structure_declared_end(entry, depth + 1)),
    }
}

fn write_structure_entry(
    out: &mut String,
    entry: &TableStructure,
    total: usize,
    axis: TableStructureAxis,
    depth: usize,
    write_range: &mut impl FnMut(&mut String, Range<usize>),
) -> Result<()> {
    ensure_structure_depth(depth)?;
    match entry {
        TableStructure::Range(range) => write_range(out, range.as_range()),
        TableStructure::Header(range) => {
            let tag = axis.header_tag();
            out.push('<');
            out.push_str(tag);
            out.push('>');
            write_range(out, range.as_range());
            out.push_str("</");
            out.push_str(tag);
            out.push('>');
        },
        TableStructure::Group(group) => {
            let tag = axis.group_tag();
            out.push('<');
            out.push_str(tag);
            if !group.display {
                out.push_str(" table:display=\"false\"");
            }
            out.push('>');
            let mut cursor = None;
            for child in &group.children {
                let (start, end) = structure_bounds(child, total, depth + 1)?;
                if cursor.is_some_and(|previous| previous != start) {
                    return Err(Error::InvalidFormat(
                        "table group children must cover one contiguous range".to_string(),
                    ));
                }
                write_structure_entry(out, child, total, axis, depth + 1, write_range)?;
                cursor = Some(end);
            }
            out.push_str("</");
            out.push_str(tag);
            out.push('>');
        },
    }
    Ok(())
}

fn structure_bounds(entry: &TableStructure, total: usize, depth: usize) -> Result<(usize, usize)> {
    ensure_structure_depth(depth)?;
    match entry {
        TableStructure::Range(range) | TableStructure::Header(range) => {
            if range.start >= range.end || range.end > total {
                return Err(Error::InvalidFormat(format!(
                    "table structure range {}..{} exceeds logical size {total}",
                    range.start, range.end
                )));
            }
            Ok((range.start, range.end))
        },
        TableStructure::Group(group) => {
            let first = group.children.first().ok_or_else(|| {
                Error::InvalidFormat("table groups must contain at least one item".to_string())
            })?;
            let last = group.children.last().expect("group has a first child");
            let (start, _) = structure_bounds(first, total, depth + 1)?;
            let (_, end) = structure_bounds(last, total, depth + 1)?;
            Ok((start, end))
        },
    }
}

fn ensure_structure_depth(depth: usize) -> Result<()> {
    if depth > MAX_TABLE_STRUCTURE_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "table structure exceeds the {MAX_TABLE_STRUCTURE_DEPTH} level nesting safety limit"
        )));
    }
    Ok(())
}

pub(crate) fn write_columns(out: &mut String, columns: &[Column]) {
    let mut start = 0usize;
    while start < columns.len() {
        let column = &columns[start];
        let mut end = start + 1;
        while end < columns.len()
            && columns[end].style_name == column.style_name
            && columns[end].default_cell_style_name == column.default_cell_style_name
            && columns[end].visibility == column.visibility
        {
            end += 1;
        }
        out.push_str("<table:table-column");
        if end - start > 1 {
            out.push_str(" table:number-columns-repeated=\"");
            out.push_str(&(end - start).to_string());
            out.push('"');
        }
        if let Some(style_name) = &column.style_name {
            out.push_str(" table:style-name=\"");
            out.push_str(&escape_xml(style_name));
            out.push('"');
        }
        if let Some(style_name) = &column.default_cell_style_name {
            out.push_str(" table:default-cell-style-name=\"");
            out.push_str(&escape_xml(style_name));
            out.push('"');
        }
        if column.visibility != TableVisibility::Visible {
            out.push_str(" table:visibility=\"");
            out.push_str(column.visibility.as_str());
            out.push('"');
        }
        out.push_str("/>");
        start = end;
    }
}

pub(crate) fn write_row_attributes(
    out: &mut String,
    style_name: Option<&str>,
    default_cell_style_name: Option<&str>,
    visibility: TableVisibility,
) {
    if let Some(style_name) = style_name {
        out.push_str(" table:style-name=\"");
        out.push_str(&escape_xml(style_name));
        out.push('"');
    }
    if let Some(style_name) = default_cell_style_name {
        out.push_str(" table:default-cell-style-name=\"");
        out.push_str(&escape_xml(style_name));
        out.push('"');
    }
    if visibility != TableVisibility::Visible {
        out.push_str(" table:visibility=\"");
        out.push_str(visibility.as_str());
        out.push('"');
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compacts_equivalent_columns_and_escapes_style_names() {
        let column = Column {
            index: 0,
            style_name: Some("Col&One".to_string()),
            default_cell_style_name: Some("Cell\"Default".to_string()),
            visibility: TableVisibility::Collapse,
        };
        let mut xml = String::new();
        write_columns(&mut xml, &[column.clone(), Column { index: 1, ..column }]);
        assert_eq!(xml.matches("<table:table-column").count(), 1);
        assert!(xml.contains(r#"table:number-columns-repeated="2""#));
        assert!(xml.contains(r#"table:style-name="Col&amp;One""#));
        assert!(xml.contains(r#"table:default-cell-style-name="Cell&quot;Default""#));
        assert!(xml.contains(r#"table:visibility="collapse""#));
    }

    #[test]
    fn splits_quoted_print_ranges_and_rejects_unterminated_names() {
        let ranges = split_cell_range_addresses(
            "$Sheet1.$A$1:$B$2 'Q1 Sales'.$C$3:$D$4 'Bob''s Sheet'.$E$5:$F$6",
        )
        .unwrap();
        assert_eq!(
            ranges,
            [
                "$Sheet1.$A$1:$B$2",
                "'Q1 Sales'.$C$3:$D$4",
                "'Bob''s Sheet'.$E$5:$F$6",
            ]
        );
        assert!(split_cell_range_addresses("'Unclosed Sheet.$A$1").is_err());
    }

    #[test]
    fn writes_sheet_style_and_print_attributes_safely() {
        let style = SheetStyle {
            style_name: Some("Sheet&Style".to_string()),
            template_name: Some("Template\"One".to_string()),
            usage: SheetStyleUsage {
                use_first_row_styles: Some(true),
                use_banding_column_styles: Some(false),
                ..SheetStyleUsage::default()
            },
        };
        let print =
            SheetPrintSettings::new(false, vec!["'Q1 Sales'.$A$1:$B$2".to_string()]).unwrap();
        let mut xml = String::new();
        write_sheet_formatting_attributes(&mut xml, &style, &print).unwrap();
        assert!(xml.contains(r#"table:style-name="Sheet&amp;Style""#));
        assert!(xml.contains(r#"table:template-name="Template&quot;One""#));
        assert!(xml.contains(r#"table:use-first-row-styles="true""#));
        assert!(xml.contains(r#"table:use-banding-columns-styles="false""#));
        assert!(xml.contains(r#"table:print="false""#));
        assert!(xml.contains(r#"table:print-ranges="&apos;Q1 Sales&apos;.$A$1:$B$2""#));
    }
}
