//! External linked-table source metadata.

use litchi_core::{Error, Result, xml::escape_xml};

/// How an external ODF table source is copied into a sheet.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableSourceMode {
    /// Copy values, formulas, and styles.
    CopyAll,
    /// Copy calculated values without formulas.
    CopyResultsOnly,
}

impl TableSourceMode {
    pub(crate) fn as_str(self) -> &'static str {
        match self {
            Self::CopyAll => "copy-all",
            Self::CopyResultsOnly => "copy-results-only",
        }
    }

    pub(crate) fn parse(value: &str) -> Result<Self> {
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
pub struct SheetTableSource {
    /// URI of the external document, preserved without dereferencing it.
    pub href: String,
    /// Optional source copying mode.
    pub mode: Option<TableSourceMode>,
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

impl SheetTableSource {
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

pub(crate) fn validate_table_source(source: &SheetTableSource) -> Result<()> {
    if let Some(delay) = &source.refresh_delay
        && !is_xsd_duration(delay)
    {
        return Err(Error::InvalidFormat(format!(
            "invalid table source refresh duration '{delay}'"
        )));
    }
    Ok(())
}

pub(crate) fn write_table_source(out: &mut String, source: &SheetTableSource) -> Result<()> {
    validate_table_source(source)?;
    out.push_str("<table:table-source xlink:type=\"simple\" xlink:href=\"");
    out.push_str(&escape_xml(&source.href));
    out.push('"');
    if source.actuate_on_request {
        out.push_str(" xlink:actuate=\"onRequest\"");
    }
    write_optional_attribute(out, "table:mode", source.mode.map(TableSourceMode::as_str));
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
        let mut source = SheetTableSource::new("../A&B.ods");
        source.mode = Some(TableSourceMode::CopyResultsOnly);
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
}
