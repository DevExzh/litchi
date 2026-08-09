//! External linked-table source metadata.

use litchi_core::{Error, Result, xml::escape_xml};
use std::num::NonZeroUsize;

/// Metadata for a rectangular range imported from an external data source.
///
/// The link is preserved as inert metadata. Litchi never dereferences it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CellRange {
    name: String,
    href: String,
    last_row_spanned: NonZeroUsize,
    last_column_spanned: NonZeroUsize,
    actuate_on_request: bool,
    filter_name: Option<String>,
    filter_options: Option<String>,
    refresh_delay: Option<String>,
}

impl CellRange {
    /// Create inert external-range metadata with positive target dimensions.
    pub fn new(
        name: impl Into<String>,
        href: impl Into<String>,
        rows: usize,
        columns: usize,
    ) -> Result<Self> {
        let last_row_spanned = NonZeroUsize::new(rows).ok_or_else(|| {
            Error::InvalidFormat("cell range source row span must be positive".to_string())
        })?;
        let last_column_spanned = NonZeroUsize::new(columns).ok_or_else(|| {
            Error::InvalidFormat("cell range source column span must be positive".to_string())
        })?;
        Ok(Self {
            name: name.into(),
            href: href.into(),
            last_row_spanned,
            last_column_spanned,
            actuate_on_request: false,
            filter_name: None,
            filter_options: None,
            refresh_delay: None,
        })
    }

    /// External source range or object name.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Replace the external source range or object name.
    pub fn set_name(&mut self, name: impl Into<String>) {
        self.name = name.into();
    }

    /// URI of the external source document.
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Replace the source URI without dereferencing it.
    pub fn set_href(&mut self, href: impl Into<String>) {
        self.href = href.into();
    }

    /// Number of rows populated by the imported range.
    pub fn rows(&self) -> usize {
        self.last_row_spanned.get()
    }

    /// Number of columns populated by the imported range.
    pub fn columns(&self) -> usize {
        self.last_column_spanned.get()
    }

    /// Change the positive target dimensions.
    pub fn set_dimensions(&mut self, rows: usize, columns: usize) -> Result<()> {
        let rows = NonZeroUsize::new(rows).ok_or_else(|| {
            Error::InvalidFormat("cell range source row span must be positive".to_string())
        })?;
        let columns = NonZeroUsize::new(columns).ok_or_else(|| {
            Error::InvalidFormat("cell range source column span must be positive".to_string())
        })?;
        self.last_row_spanned = rows;
        self.last_column_spanned = columns;
        Ok(())
    }

    /// Whether the source explicitly uses `xlink:actuate="onRequest"`.
    pub fn actuate_on_request(&self) -> bool {
        self.actuate_on_request
    }

    /// Control preservation of `xlink:actuate="onRequest"`.
    pub fn set_actuate_on_request(&mut self, enabled: bool) {
        self.actuate_on_request = enabled;
    }

    /// Optional import filter name.
    pub fn filter_name(&self) -> Option<&str> {
        self.filter_name.as_deref()
    }

    /// Set or clear the import filter name.
    pub fn set_filter_name(&mut self, value: Option<String>) {
        self.filter_name = value;
    }

    /// Optional filter-specific arguments.
    pub fn filter_options(&self) -> Option<&str> {
        self.filter_options.as_deref()
    }

    /// Set or clear filter-specific arguments.
    pub fn set_filter_options(&mut self, value: Option<String>) {
        self.filter_options = value;
    }

    /// Optional XML Schema duration controlling refresh frequency.
    pub fn refresh_delay(&self) -> Option<&str> {
        self.refresh_delay.as_deref()
    }

    /// Set or clear a validated XML Schema refresh duration.
    pub fn set_refresh_delay(&mut self, value: Option<String>) -> Result<()> {
        if let Some(delay) = &value
            && !is_xsd_duration(delay)
        {
            return Err(Error::InvalidFormat(format!(
                "invalid cell range source refresh duration '{delay}'"
            )));
        }
        self.refresh_delay = value;
        Ok(())
    }
}

/// How an external ODF table source is copied into a sheet.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableMode {
    /// Copy values, formulas, and styles.
    CopyAll,
    /// Copy calculated values without formulas.
    CopyResultsOnly,
}

impl TableMode {
    #[must_use]
    pub fn as_str(self) -> &'static str {
        match self {
            Self::CopyAll => "copy-all",
            Self::CopyResultsOnly => "copy-results-only",
        }
    }

    /// # Errors
    ///
    /// Returns an error when the input is malformed or exceeds the parser's resource limits.
    pub fn parse(value: &str) -> Result<Self> {
        match value {
            "copy-all" => Ok(Self::CopyAll),
            "copy-results-only" => Ok(Self::CopyResultsOnly),
            _ => Err(Error::InvalidFormat(format!(
                "invalid table source mode '{value}'"
            ))),
        }
    }
}

/// Metadata describing an external document linked into an ODF sheet.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SheetTable {
    /// URI of the external document, preserved without dereferencing it.
    pub href: String,
    /// Optional source copying mode.
    pub mode: Option<TableMode>,
    /// Optional source table name.
    pub table_name: Option<String>,
    /// Whether the link explicitly requests `xlink:actuate="onRequest"`.
    pub actuate_on_request: bool,
    /// Optional import filter name.
    pub filter_name: Option<String>,
    /// Optional filter-specific arguments.
    pub filter_options: Option<String>,
    /// Optional XML Schema duration controlling refresh frequency.
    pub refresh_delay: Option<String>,
}

impl SheetTable {
    /// Create linked-table metadata without accessing the target URI.
    pub fn new(href: impl Into<String>) -> Self {
        Self {
            href: href.into(),
            mode: None,
            table_name: None,
            actuate_on_request: false,
            filter_name: None,
            filter_options: None,
            refresh_delay: None,
        }
    }
}

/// # Errors
///
/// Returns an error when a value violates the format or resource constraints.
#[allow(
    clippy::module_name_repetitions,
    reason = "the codec entry point keeps its historical element-qualified name"
)]
pub fn validate_table_source(source: &SheetTable) -> Result<()> {
    if let Some(delay) = &source.refresh_delay
        && !is_xsd_duration(delay)
    {
        return Err(Error::InvalidFormat(format!(
            "invalid table source refresh duration '{delay}'"
        )));
    }
    Ok(())
}

/// # Errors
///
/// Returns an error when the value cannot be serialized.
#[allow(
    clippy::module_name_repetitions,
    reason = "the codec entry point keeps its historical element-qualified name"
)]
pub fn write_table_source(out: &mut String, source: &SheetTable) -> Result<()> {
    validate_table_source(source)?;
    out.push_str("<table:table-source xlink:type=\"simple\" xlink:href=\"");
    out.push_str(&escape_xml(&source.href));
    out.push('"');
    if source.actuate_on_request {
        out.push_str(" xlink:actuate=\"onRequest\"");
    }
    write_optional_attribute(out, "table:mode", source.mode.map(TableMode::as_str));
    write_optional_attribute(out, "table:table-name", source.table_name.as_deref());
    write_optional_attribute(out, "table:filter-name", source.filter_name.as_deref());
    write_optional_attribute(
        out,
        "table:filter-options",
        source.filter_options.as_deref(),
    );
    write_optional_attribute(out, "table:refresh-delay", source.refresh_delay.as_deref());
    out.push_str("/>");
    Ok(())
}

#[allow(
    clippy::module_name_repetitions,
    reason = "the codec entry point keeps its historical element-qualified name"
)]
pub fn write_cell_range_source(out: &mut String, source: &CellRange) {
    out.push_str("<table:cell-range-source table:name=\"");
    out.push_str(&escape_xml(source.name()));
    out.push_str("\" table:last-column-spanned=\"");
    out.push_str(&source.columns().to_string());
    out.push_str("\" table:last-row-spanned=\"");
    out.push_str(&source.rows().to_string());
    out.push_str("\" xlink:type=\"simple\" xlink:href=\"");
    out.push_str(&escape_xml(source.href()));
    out.push('"');
    if source.actuate_on_request() {
        out.push_str(" xlink:actuate=\"onRequest\"");
    }
    write_optional_attribute(out, "table:filter-name", source.filter_name());
    write_optional_attribute(out, "table:filter-options", source.filter_options());
    write_optional_attribute(out, "table:refresh-delay", source.refresh_delay());
    out.push_str("/>");
}

fn write_optional_attribute(out: &mut String, name: &str, value: Option<&str>) {
    if let Some(value) = value {
        out.push(' ');
        out.push_str(name);
        out.push_str("=\"");
        out.push_str(&escape_xml(value));
        out.push('"');
    }
}

fn is_xsd_duration(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return false;
    }
    index += 1;
    let mut any = false;
    any |= consume_integer_component(bytes, &mut index, b'Y');
    any |= consume_integer_component(bytes, &mut index, b'M');
    any |= consume_integer_component(bytes, &mut index, b'D');
    if bytes.get(index) == Some(&b'T') {
        index += 1;
        let mut any_time = false;
        any_time |= consume_integer_component(bytes, &mut index, b'H');
        any_time |= consume_integer_component(bytes, &mut index, b'M');
        any_time |= consume_seconds(bytes, &mut index);
        if !any_time {
            return false;
        }
        any = true;
    }
    any && index == bytes.len()
}

fn consume_integer_component(bytes: &[u8], index: &mut usize, suffix: u8) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end > *index && bytes.get(end) == Some(&suffix) {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn consume_seconds(bytes: &[u8], index: &mut usize) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end == *index {
        return false;
    }
    if bytes.get(end) == Some(&b'.') {
        end += 1;
        let fractional_start = end;
        while bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
        if end == fractional_start {
            return false;
        }
    }
    if bytes.get(end) == Some(&b'S') {
        *index = end + 1;
        true
    } else {
        false
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_complete_duration_grammar() {
        for valid in ["P0D", "PT15M", "P1Y2M3DT4H5M6.25S", "-PT1S"] {
            assert!(is_xsd_duration(valid), "{valid}");
        }
        for invalid in ["", "P", "PT", "1D", "+P1D", "P1H", "PT1.0M", "PT1.S"] {
            assert!(!is_xsd_duration(invalid), "{invalid}");
        }
    }

    #[test]
    fn writes_escaped_inert_link_metadata() {
        let mut source = SheetTable::new("../A&B.ods");
        source.mode = Some(TableMode::CopyResultsOnly);
        source.table_name = Some("Q1 <Q2".to_string());
        source.actuate_on_request = true;
        source.refresh_delay = Some("PT15M".to_string());
        let mut xml = String::new();
        write_table_source(&mut xml, &source).unwrap();
        assert!(xml.contains(r#"xlink:href="../A&amp;B.ods""#));
        assert!(xml.contains(r#"table:mode="copy-results-only""#));
        assert!(xml.contains(r#"table:table-name="Q1 &lt;Q2""#));
        assert!(xml.contains(r#"table:refresh-delay="PT15M""#));
    }

    #[test]
    fn cell_range_source_enforces_positive_dimensions_and_duration() {
        assert!(CellRange::new("Range", "source.ods", 0, 1).is_err());
        assert!(CellRange::new("Range", "source.ods", 1, 0).is_err());

        let mut source = CellRange::new("Data & More", "../A&B.ods", 3, 4).unwrap();
        source.set_actuate_on_request(true);
        source.set_filter_name(Some("calc8".to_string()));
        source.set_refresh_delay(Some("PT5M".to_string())).unwrap();
        assert!(
            source
                .set_refresh_delay(Some("every five minutes".to_string()))
                .is_err()
        );
        assert_eq!(source.refresh_delay(), Some("PT5M"));

        let mut xml = String::new();
        write_cell_range_source(&mut xml, &source);
        assert!(xml.contains(r#"table:name="Data &amp; More""#));
        assert!(xml.contains(r#"table:last-column-spanned="4""#));
        assert!(xml.contains(r#"table:last-row-spanned="3""#));
        assert!(xml.contains(r#"xlink:href="../A&amp;B.ods""#));
        assert!(xml.contains(r#"xlink:actuate="onRequest""#));
    }
}
