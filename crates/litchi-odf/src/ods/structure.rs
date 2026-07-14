//! ODS row and column structural metadata.

use litchi_core::{Error, Result, xml::escape_xml};

pub(crate) const MAX_EXPANDED_ROWS_PER_SHEET: usize = 1_048_576;
pub(crate) const MAX_EXPANDED_COLUMNS_PER_SHEET: usize = 1_048_576;

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
}
