//! Streaming ODF worksheet parser and canonical table writer.

use super::{Cell, CellValue, Row, Sheet};
use super::{model::Merge, validation};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{fmt::Write as _, num::NonZeroUsize};

pub(crate) const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(crate) const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_XML_DEPTH: usize = 1_024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Kind {
    Root,
    Body,
    Spreadsheet,
    DdeLink,
    DdeCache,
    Table,
    Row,
    Cell,
    Text,
    Other,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Text,
    Other,
}

struct Attributes {
    name: Option<String>,
    style_name: Option<String>,
    default_cell_style_name: Option<String>,
    value_type: Option<String>,
    value: Option<String>,
    date_value: Option<String>,
    time_value: Option<String>,
    boolean_value: Option<String>,
    currency: Option<String>,
    formula: Option<String>,
    rows_repeated: usize,
    columns_repeated: usize,
    rows_spanned: usize,
    columns_spanned: usize,
    covered: bool,
}

impl Default for Attributes {
    fn default() -> Self {
        Self {
            name: None,
            style_name: None,
            default_cell_style_name: None,
            value_type: None,
            value: None,
            date_value: None,
            time_value: None,
            boolean_value: None,
            currency: None,
            formula: None,
            rows_repeated: 1,
            columns_repeated: 1,
            rows_spanned: 1,
            columns_spanned: 1,
            covered: false,
        }
    }
}

impl Attributes {
    fn from_element(
        element: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        covered: bool,
    ) -> Result<Self> {
        let mut result = Self {
            covered,
            ..Self::default()
        };
        for raw in element.attributes().with_checks(true) {
            let raw = raw
                .map_err(|error| Error::InvalidFormat(format!("invalid ODS attribute: {error}")))?;
            let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
            let local = String::from_utf8_lossy(local.as_ref());
            let value = raw
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODS attribute value: {error}"))
                })?
                .into_owned();
            let namespace = match namespace {
                ResolveResult::Bound(Namespace(uri)) => String::from_utf8_lossy(uri).into_owned(),
                ResolveResult::Unbound => String::new(),
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::InvalidFormat(format!(
                        "unbound ODS attribute prefix '{}'",
                        String::from_utf8_lossy(prefix.as_ref())
                    )));
                },
            };
            if namespace == TABLE_NAMESPACE {
                match local.as_ref() {
                    "name" => result.name = Some(value),
                    "style-name" => result.style_name = Some(value),
                    "default-cell-style-name" => result.default_cell_style_name = Some(value),
                    "formula" => result.formula = Some(value),
                    "number-rows-repeated" => {
                        result.rows_repeated = positive(&value, "number-rows-repeated")?
                    },
                    "number-columns-repeated" => {
                        result.columns_repeated = positive(&value, "number-columns-repeated")?
                    },
                    "number-rows-spanned" => {
                        result.rows_spanned = positive(&value, "number-rows-spanned")?
                    },
                    "number-columns-spanned" => {
                        result.columns_spanned = positive(&value, "number-columns-spanned")?
                    },
                    _ => {},
                }
            } else if namespace == OFFICE_NAMESPACE {
                match local.as_ref() {
                    "value-type" => result.value_type = Some(value),
                    "value" => result.value = Some(value),
                    "date-value" => result.date_value = Some(value),
                    "time-value" => result.time_value = Some(value),
                    "boolean-value" => result.boolean_value = Some(value),
                    "currency" => result.currency = Some(value),
                    _ => {},
                }
            }
        }
        Ok(result)
    }

    fn cell(self, text: String) -> Result<Cell> {
        let value_type = self.value_type.as_deref();
        let value = match value_type {
            None => {
                if text.is_empty() {
                    CellValue::Empty
                } else {
                    CellValue::Text(text.clone())
                }
            },
            Some("string") => CellValue::Text(text.clone()),
            Some("float") | Some("double") | Some("decimal") => CellValue::Number(parse_float(
                self.value.as_deref().unwrap_or_default(),
                "office:value",
            )?),
            Some("currency") => CellValue::Currency {
                value: parse_float(self.value.as_deref().unwrap_or_default(), "office:value")?,
                currency: self.currency.unwrap_or_default(),
            },
            Some("percentage") => CellValue::Percentage(parse_float(
                self.value.as_deref().unwrap_or_default(),
                "office:value",
            )?),
            Some("boolean") => CellValue::Boolean(parse_bool(
                self.boolean_value
                    .as_deref()
                    .or(self.value.as_deref())
                    .unwrap_or_default(),
                "office:boolean-value",
            )?),
            Some("date") => CellValue::Date(self.date_value.or(self.value).ok_or_else(|| {
                Error::InvalidFormat("date cells require office:date-value".to_string())
            })?),
            Some("time") => CellValue::Time(self.time_value.or(self.value).ok_or_else(|| {
                Error::InvalidFormat("time cells require office:time-value".to_string())
            })?),
            Some(kind) => CellValue::Unknown {
                kind: kind.to_string(),
                value: self.value.or(self.date_value).or(self.time_value),
            },
        };
        let mut cell = Cell::repeated(value, text, self.columns_repeated)?;
        cell.formula = self.formula;
        cell.style_name = self.style_name;
        cell.merge = if self.covered {
            Merge::Covered
        } else if self.rows_spanned != 1 || self.columns_spanned != 1 {
            Merge::Span {
                rows: NonZeroUsize::new(self.rows_spanned).expect("positive row span was checked"),
                columns: NonZeroUsize::new(self.columns_spanned)
                    .expect("positive column span was checked"),
            }
        } else {
            Merge::None
        };
        Ok(cell)
    }
}

struct OpenCell {
    attributes: Attributes,
    text: String,
    text_depth: usize,
}

/// Parse all direct spreadsheet tables from an ODS content part.
pub(crate) fn parse(xml: &str) -> Result<Vec<Sheet>> {
    parse_impl(xml, true)
}

/// Parse flat spreadsheet tables while retaining duplicate names for
/// selector-time ambiguity reporting.
pub(crate) fn parse_flat(xml: &str) -> Result<Vec<Sheet>> {
    parse_impl(xml, false)
}

fn parse_impl(xml: &str, require_unique_names: bool) -> Result<Vec<Sheet>> {
    validation::validate_content_xml_size(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<Kind>::new();
    let mut dde_cache_depth = None;
    let mut sheets = Vec::new();
    let mut current_sheet: Option<Sheet> = None;
    let mut current_row: Option<Row> = None;
    let mut current_cell: Option<OpenCell> = None;

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid ODS XML: {error}")))?;
        let namespace = namespace_kind(&namespace);
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"
                    )));
                }
                let local = element.local_name();
                let mut kind = classify(namespace, local.as_ref());
                if dde_cache_depth.is_some() {
                    kind = Kind::Other;
                } else if kind == Kind::Table && stack.last() == Some(&Kind::DdeLink) {
                    kind = Kind::DdeCache;
                    dde_cache_depth = Some(stack.len() + 1);
                }
                match kind {
                    Kind::Root => {},
                    Kind::Body
                    | Kind::Spreadsheet
                    | Kind::DdeLink
                    | Kind::DdeCache
                    | Kind::Other
                    | Kind::Text => {},
                    Kind::Table => {
                        if stack.last() != Some(&Kind::Spreadsheet) {
                            return Err(Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            ));
                        }
                        if current_sheet.is_some() {
                            return Err(Error::InvalidFormat(
                                "ODS worksheet parser encountered a nested table".to_string(),
                            ));
                        }
                        let attributes = Attributes::from_element(&element, &reader, false)?;
                        current_sheet = Some(Sheet {
                            name: attributes.name.unwrap_or_else(|| "Sheet1".to_string()),
                            rows: Vec::new(),
                            style_name: attributes.style_name,
                        });
                    },
                    Kind::Row => {
                        if current_sheet.is_none() || current_row.is_some() {
                            return Err(Error::InvalidFormat(
                                "table:table-row is outside a worksheet row context".to_string(),
                            ));
                        }
                        let attributes = Attributes::from_element(&element, &reader, false)?;
                        current_row = Some(Row {
                            cells: Vec::new(),
                            style_name: attributes.style_name,
                            default_cell_style_name: attributes.default_cell_style_name,
                            repeat: NonZeroUsize::new(attributes.rows_repeated)
                                .expect("positive row repetition was checked"),
                        });
                    },
                    Kind::Cell => {
                        if current_row.is_none() || current_cell.is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell is outside a worksheet row context".to_string(),
                            ));
                        }
                        current_cell = Some(OpenCell {
                            attributes: Attributes::from_element(
                                &element,
                                &reader,
                                is_covered(namespace, local.as_ref()),
                            )?,
                            text: String::new(),
                            text_depth: 0,
                        });
                    },
                }
                if kind == Kind::Text && current_cell.is_some() {
                    current_cell.as_mut().expect("cell is present").text_depth += 1;
                }
                stack.push(kind);
            },
            Event::Empty(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"
                    )));
                }
                let local = element.local_name();
                let mut kind = classify(namespace, local.as_ref());
                if dde_cache_depth.is_some() {
                    kind = Kind::Other;
                } else if kind == Kind::Table && stack.last() == Some(&Kind::DdeLink) {
                    kind = Kind::DdeCache;
                }
                match kind {
                    Kind::Table => {
                        if stack.last() != Some(&Kind::Spreadsheet) {
                            return Err(Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            ));
                        }
                        let attributes = Attributes::from_element(&element, &reader, false)?;
                        sheets.push(Sheet {
                            name: attributes.name.unwrap_or_else(|| "Sheet1".to_string()),
                            rows: Vec::new(),
                            style_name: attributes.style_name,
                        });
                    },
                    Kind::Row => {
                        let attributes = Attributes::from_element(&element, &reader, false)?;
                        let sheet = current_sheet.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "empty table row is outside a worksheet".to_string(),
                            )
                        })?;
                        sheet.rows.push(Row {
                            cells: Vec::new(),
                            style_name: attributes.style_name,
                            default_cell_style_name: attributes.default_cell_style_name,
                            repeat: NonZeroUsize::new(attributes.rows_repeated)
                                .expect("positive row repetition was checked"),
                        });
                    },
                    Kind::Cell => {
                        let row = current_row.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "empty table cell is outside a worksheet row".to_string(),
                            )
                        })?;
                        let attributes = Attributes::from_element(
                            &element,
                            &reader,
                            is_covered(namespace, local.as_ref()),
                        )?;
                        row.cells.push(attributes.cell(String::new())?);
                    },
                    Kind::Text if current_cell.is_some() => {
                        append_empty_text(
                            &mut current_cell,
                            namespace,
                            local.as_ref(),
                            &element,
                            &reader,
                        )?;
                    },
                    _ => {},
                }
            },
            Event::Text(text) => {
                if let Some(cell) = current_cell.as_mut()
                    && cell.text_depth > 0
                {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.text.push_str(&value);
                }
            },
            Event::CData(text) => {
                if let Some(cell) = current_cell.as_mut()
                    && cell.text_depth > 0
                {
                    let value = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.text.push_str(&value);
                }
            },
            Event::End(element) => {
                let kind = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                if kind == Kind::DdeCache {
                    dde_cache_depth = None;
                }
                if kind == Kind::Text && current_cell.is_some() {
                    current_cell.as_mut().expect("cell is present").text_depth = current_cell
                        .as_ref()
                        .expect("cell is present")
                        .text_depth
                        .saturating_sub(1);
                }
                match kind {
                    Kind::Cell => {
                        let open = current_cell.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell close has no open cell".to_string())
                        })?;
                        let row = current_row.as_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell closed outside a row".to_string())
                        })?;
                        row.cells.push(open.attributes.cell(open.text)?);
                    },
                    Kind::Row => {
                        let row = current_row.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS row close has no open row".to_string())
                        })?;
                        current_sheet
                            .as_mut()
                            .ok_or_else(|| {
                                Error::InvalidFormat("ODS row closed outside a sheet".to_string())
                            })?
                            .rows
                            .push(row);
                    },
                    Kind::Table => {
                        let sheet = current_sheet.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS table close has no open table".to_string())
                        })?;
                        sheets.push(sheet);
                    },
                    _ => {},
                }
                let _ = element;
            },
            Event::Eof => break,
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }

    if !stack.is_empty()
        || current_sheet.is_some()
        || current_row.is_some()
        || current_cell.is_some()
    {
        return Err(Error::InvalidFormat(
            "ODS content ended with an unfinished worksheet object".to_string(),
        ));
    }
    if require_unique_names {
        validation::validate_sheets(&sheets)?;
    } else {
        if sheets.len() > validation::MAX_PHYSICAL_RUNS {
            return Err(Error::InvalidFormat(format!(
                "ODS sheet count exceeds the {} safety limit",
                validation::MAX_PHYSICAL_RUNS
            )));
        }
        for sheet in &sheets {
            validation::validate_sheet(sheet)?;
        }
    }
    Ok(sheets)
}

#[cfg(test)]
mod bounded_depth_tests {
    use super::{MAX_XML_DEPTH, parse};

    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="urn:litchi:test"><office:body><office:spreadsheet>"#;
    const SUFFIX: &str = "</office:spreadsheet></office:body></office:document-content>";

    fn nested_unknown(total_depth: usize) -> String {
        let nested = total_depth - 3;
        let mut xml = String::with_capacity(PREFIX.len() + SUFFIX.len() + nested * 7);
        xml.push_str(PREFIX);
        for _ in 0..nested {
            xml.push_str("<x:n>");
        }
        for _ in 0..nested {
            xml.push_str("</x:n>");
        }
        xml.push_str(SUFFIX);
        xml
    }

    fn nested_dde_cache(total_depth: usize) -> String {
        const CACHE_PREFIX: &str = "<table:dde-links><table:dde-link><office:dde-source office:dde-application=\"app\" office:dde-topic=\"topic\" office:dde-item=\"item\"/><table:table>";
        const CACHE_SUFFIX: &str = "</table:table></table:dde-link></table:dde-links>";
        let nested = total_depth - 6;
        let mut xml = String::with_capacity(
            PREFIX.len() + CACHE_PREFIX.len() + CACHE_SUFFIX.len() + SUFFIX.len() + nested * 7,
        );
        xml.push_str(PREFIX);
        xml.push_str(CACHE_PREFIX);
        for _ in 0..nested {
            xml.push_str("<x:n>");
        }
        for _ in 0..nested {
            xml.push_str("</x:n>");
        }
        xml.push_str(CACHE_SUFFIX);
        xml.push_str(SUFFIX);
        xml
    }

    #[test]
    fn accepts_exact_depth_and_rejects_next_depth_outside_dde() {
        assert!(parse(&nested_unknown(MAX_XML_DEPTH)).is_ok());
        assert!(parse(&nested_unknown(MAX_XML_DEPTH + 1)).is_err());
    }

    #[test]
    fn accepts_exact_depth_and_rejects_next_depth_inside_inert_dde_cache() {
        assert!(parse(&nested_dde_cache(MAX_XML_DEPTH)).is_ok());
        assert!(parse(&nested_dde_cache(MAX_XML_DEPTH + 1)).is_err());
    }
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NAMESPACE.as_bytes() => {
            NamespaceKind::Office
        },
        ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE.as_bytes() => {
            NamespaceKind::Table
        },
        ResolveResult::Bound(Namespace(value)) if *value == TEXT_NAMESPACE.as_bytes() => {
            NamespaceKind::Text
        },
        _ => NamespaceKind::Other,
    }
}

fn classify(namespace: NamespaceKind, local: &[u8]) -> Kind {
    if namespace == NamespaceKind::Office && local == b"document-content" {
        Kind::Root
    } else if namespace == NamespaceKind::Office && local == b"body" {
        Kind::Body
    } else if namespace == NamespaceKind::Office && local == b"spreadsheet" {
        Kind::Spreadsheet
    } else if namespace == NamespaceKind::Table && local == b"dde-link" {
        Kind::DdeLink
    } else if namespace == NamespaceKind::Table && local == b"table" {
        Kind::Table
    } else if namespace == NamespaceKind::Table && local == b"table-row" {
        Kind::Row
    } else if namespace == NamespaceKind::Table
        && (local == b"table-cell" || local == b"covered-table-cell")
    {
        Kind::Cell
    } else if namespace == NamespaceKind::Text
        && (local == b"p"
            || local == b"span"
            || local == b"a"
            || local == b"line-break"
            || local == b"tab"
            || local == b"s")
    {
        Kind::Text
    } else {
        Kind::Other
    }
}

fn is_covered(namespace: NamespaceKind, local: &[u8]) -> bool {
    namespace == NamespaceKind::Table && local == b"covered-table-cell"
}

fn append_empty_text(
    current_cell: &mut Option<OpenCell>,
    namespace: NamespaceKind,
    local: &[u8],
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
) -> Result<()> {
    let Some(cell) = current_cell.as_mut() else {
        return Ok(());
    };
    let bound = namespace == NamespaceKind::Text;
    if !bound {
        return Ok(());
    }
    if local == b"tab" {
        cell.text.push('\t');
    } else if local == b"line-break" {
        cell.text.push('\n');
    } else if local == b"s" {
        let mut count = 1usize;
        for raw in element.attributes().with_checks(true) {
            let raw = raw.map_err(|error| {
                Error::InvalidFormat(format!("invalid text:s attribute: {error}"))
            })?;
            let (namespace, local_name) = reader.resolver().resolve_attribute(raw.key);
            let is_count = matches!(namespace, ResolveResult::Bound(Namespace(value))
                if value == TEXT_NAMESPACE.as_bytes())
                && local_name.as_ref() == b"c";
            if is_count {
                let value = raw
                    .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid text:s count: {error}"))
                    })?;
                count = positive(&value, "text:s c")?;
            }
        }
        if count > validation::MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "text:s expansion exceeds the worksheet text safety limit".to_string(),
            ));
        }
        cell.text.extend(std::iter::repeat_n(' ', count));
    }
    Ok(())
}

fn positive(value: &str, name: &str) -> Result<usize> {
    let value = value
        .parse::<usize>()
        .map_err(|_| Error::InvalidFormat(format!("ODS {name} must be a positive integer")))?;
    NonZeroUsize::new(value)
        .map(NonZeroUsize::get)
        .ok_or_else(|| Error::InvalidFormat(format!("ODS {name} must be positive")))
}

fn parse_float(value: &str, name: &str) -> Result<f64> {
    let value = value
        .parse::<f64>()
        .map_err(|_| Error::InvalidFormat(format!("ODS {name} requires a finite decimal value")))?;
    if value.is_finite() {
        Ok(value)
    } else {
        Err(Error::InvalidFormat(format!("ODS {name} must be finite")))
    }
}

fn parse_bool(value: &str, name: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "ODS {name} has invalid Boolean value"
        ))),
    }
}

/// Render one worksheet as a standalone table fragment.
pub(crate) fn write_sheet(sheet: &Sheet) -> Result<String> {
    validation::validate_sheet(sheet)?;
    let mut output = String::new();
    output.push_str("<table:table xmlns:table=\"");
    output.push_str(TABLE_NAMESPACE);
    output.push_str("\" xmlns:office=\"");
    output.push_str(OFFICE_NAMESPACE);
    output.push_str("\" xmlns:text=\"");
    output.push_str(TEXT_NAMESPACE);
    output.push_str("\" table:name=\"");
    output.push_str(&escape_xml(&sheet.name));
    output.push('"');
    if let Some(style_name) = &sheet.style_name {
        output.push_str(" table:style-name=\"");
        output.push_str(&escape_xml(style_name));
        output.push('"');
    }
    if sheet.rows.is_empty() {
        output.push_str("/>");
        return Ok(output);
    }
    output.push('>');
    for row in &sheet.rows {
        write_row(&mut output, row)?;
    }
    output.push_str("</table:table>");
    Ok(output)
}

/// Render row fragments under an exact allocation and output byte bound.
pub(crate) fn write_rows_bounded(rows: &[Row], max_bytes: usize) -> Result<String> {
    let mut output = String::new();
    for row in rows {
        write_row_bounded(&mut output, row, max_bytes)?;
    }
    Ok(output)
}

fn bounded_push(output: &mut String, value: &str, max_bytes: usize) -> Result<()> {
    let next = output.len().checked_add(value.len()).ok_or_else(|| {
        Error::InvalidFormat("flat ODS rendered size overflows usize".to_string())
    })?;
    if next > max_bytes {
        return Err(Error::InvalidFormat(format!(
            "flat ODS rendered rows exceed the {max_bytes} byte limit"
        )));
    }
    output.try_reserve(value.len()).map_err(|_| {
        Error::InvalidFormat("flat ODS row rendering allocation failed".to_string())
    })?;
    output.push_str(value);
    Ok(())
}

fn write_row_bounded(output: &mut String, row: &Row, max_bytes: usize) -> Result<()> {
    validation::validate_cell_runs(&row.cells)?;
    bounded_push(
        output,
        concat!(
            "<table:table-row xmlns:table=\"",
            "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
            "\" xmlns:office=\"",
            "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
            "\" xmlns:text=\"",
            "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
            "\""
        ),
        max_bytes,
    )?;
    if row.repeat() > 1 {
        bounded_push(output, " table:number-rows-repeated=\"", max_bytes)?;
        bounded_push(output, &row.repeat().to_string(), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(value) = &row.style_name {
        bounded_push(output, " table:style-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(value), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(value) = &row.default_cell_style_name {
        bounded_push(output, " table:default-cell-style-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(value), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if row.cells.is_empty() {
        return bounded_push(output, "/>", max_bytes);
    }
    bounded_push(output, ">", max_bytes)?;
    for cell in &row.cells {
        write_cell_bounded(output, cell, max_bytes)?;
    }
    bounded_push(output, "</table:table-row>", max_bytes)
}

fn write_cell_bounded(output: &mut String, cell: &Cell, max_bytes: usize) -> Result<()> {
    validation::validate_cell(cell)?;
    let covered = matches!(cell.merge, Merge::Covered);
    bounded_push(
        output,
        if covered {
            "<table:covered-table-cell"
        } else {
            "<table:table-cell"
        },
        max_bytes,
    )?;
    if cell.repeat() > 1 {
        bounded_push(output, " table:number-columns-repeated=\"", max_bytes)?;
        bounded_push(output, &cell.repeat().to_string(), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Merge::Span { rows, columns } = cell.merge {
        bounded_push(output, " table:number-rows-spanned=\"", max_bytes)?;
        bounded_push(output, &rows.to_string(), max_bytes)?;
        bounded_push(output, "\" table:number-columns-spanned=\"", max_bytes)?;
        bounded_push(output, &columns.to_string(), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(value) = &cell.formula {
        bounded_push(output, " table:formula=\"", max_bytes)?;
        bounded_push(output, &escape_xml(value), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(value) = &cell.style_name {
        bounded_push(output, " table:style-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(value), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if !covered {
        write_value_attributes_bounded(output, &cell.value, max_bytes)?;
    }
    if cell.text.is_empty() && matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
        return bounded_push(output, "/>", max_bytes);
    }
    bounded_push(output, ">", max_bytes)?;
    if !cell.text.is_empty() || matches!(cell.value, CellValue::Text(_)) {
        bounded_push(output, "<text:p>", max_bytes)?;
        bounded_push(output, &escape_xml(&cell.text), max_bytes)?;
        bounded_push(output, "</text:p>", max_bytes)?;
    }
    bounded_push(
        output,
        if covered {
            "</table:covered-table-cell>"
        } else {
            "</table:table-cell>"
        },
        max_bytes,
    )
}

fn write_value_attributes_bounded(
    output: &mut String,
    value: &CellValue,
    max_bytes: usize,
) -> Result<()> {
    match value {
        CellValue::Empty => Ok(()),
        CellValue::Text(_) => bounded_push(output, " office:value-type=\"string\"", max_bytes),
        CellValue::Number(value) => {
            bounded_push(
                output,
                " office:value-type=\"float\" office:value=\"",
                max_bytes,
            )?;
            bounded_push(output, &value.to_string(), max_bytes)?;
            bounded_push(output, "\"", max_bytes)
        },
        CellValue::Currency { value, currency } => {
            bounded_push(
                output,
                " office:value-type=\"currency\" office:value=\"",
                max_bytes,
            )?;
            bounded_push(output, &value.to_string(), max_bytes)?;
            bounded_push(output, "\" office:currency=\"", max_bytes)?;
            bounded_push(output, &escape_xml(currency), max_bytes)?;
            bounded_push(output, "\"", max_bytes)
        },
        CellValue::Percentage(value) => {
            bounded_push(
                output,
                " office:value-type=\"percentage\" office:value=\"",
                max_bytes,
            )?;
            bounded_push(output, &value.to_string(), max_bytes)?;
            bounded_push(output, "\"", max_bytes)
        },
        CellValue::Boolean(value) => bounded_push(
            output,
            if *value {
                " office:value-type=\"boolean\" office:boolean-value=\"true\""
            } else {
                " office:value-type=\"boolean\" office:boolean-value=\"false\""
            },
            max_bytes,
        ),
        CellValue::Date(value) => {
            bounded_push(
                output,
                " office:value-type=\"date\" office:date-value=\"",
                max_bytes,
            )?;
            bounded_push(output, &escape_xml(value), max_bytes)?;
            bounded_push(output, "\"", max_bytes)
        },
        CellValue::Time(value) => {
            bounded_push(
                output,
                " office:value-type=\"time\" office:time-value=\"",
                max_bytes,
            )?;
            bounded_push(output, &escape_xml(value), max_bytes)?;
            bounded_push(output, "\"", max_bytes)
        },
        CellValue::Unknown { kind, value } => {
            bounded_push(output, " office:value-type=\"", max_bytes)?;
            bounded_push(output, &escape_xml(kind), max_bytes)?;
            if let Some(value) = value {
                bounded_push(output, "\" office:value=\"", max_bytes)?;
                bounded_push(output, &escape_xml(value), max_bytes)?;
            }
            bounded_push(output, "\"", max_bytes)
        },
    }
}

fn write_row(output: &mut String, row: &Row) -> Result<()> {
    write_row_inner(output, row, false)
}

fn write_row_inner(output: &mut String, row: &Row, bind_namespaces: bool) -> Result<()> {
    validation::validate_cell_runs(&row.cells)?;
    output.push_str("<table:table-row");
    if bind_namespaces {
        output.push_str(" xmlns:table=\"");
        output.push_str(TABLE_NAMESPACE);
        output.push_str("\" xmlns:office=\"");
        output.push_str(OFFICE_NAMESPACE);
        output.push_str("\" xmlns:text=\"");
        output.push_str(TEXT_NAMESPACE);
        output.push('"');
    }
    if row.repeat() > 1 {
        output.push_str(" table:number-rows-repeated=\"");
        write!(output, "{}", row.repeat()).expect("writing a number to String cannot fail");
        output.push('"');
    }
    if let Some(style_name) = &row.style_name {
        output.push_str(" table:style-name=\"");
        output.push_str(&escape_xml(style_name));
        output.push('"');
    }
    if let Some(style_name) = &row.default_cell_style_name {
        output.push_str(" table:default-cell-style-name=\"");
        output.push_str(&escape_xml(style_name));
        output.push('"');
    }
    if row.cells.is_empty() {
        output.push_str("/>");
        return Ok(());
    }
    output.push('>');
    for cell in &row.cells {
        write_cell(output, cell)?;
    }
    output.push_str("</table:table-row>");
    Ok(())
}

fn write_cell(output: &mut String, cell: &Cell) -> Result<()> {
    validation::validate_cell(cell)?;
    let covered = matches!(cell.merge, Merge::Covered);
    output.push_str(if covered {
        "<table:covered-table-cell"
    } else {
        "<table:table-cell"
    });
    if cell.repeat() > 1 {
        output.push_str(" table:number-columns-repeated=\"");
        write!(output, "{}", cell.repeat()).expect("writing a number to String cannot fail");
        output.push('"');
    }
    if let Merge::Span { rows, columns } = cell.merge {
        output.push_str(" table:number-rows-spanned=\"");
        write!(output, "{}", rows.get()).expect("writing a number to String cannot fail");
        output.push_str("\" table:number-columns-spanned=\"");
        write!(output, "{}", columns.get()).expect("writing a number to String cannot fail");
        output.push('"');
    }
    if let Some(formula) = &cell.formula {
        output.push_str(" table:formula=\"");
        output.push_str(&escape_xml(formula));
        output.push('"');
    }
    if let Some(style_name) = &cell.style_name {
        output.push_str(" table:style-name=\"");
        output.push_str(&escape_xml(style_name));
        output.push('"');
    }
    if !covered {
        write_value_attributes(output, &cell.value);
    }
    if cell.text.is_empty() && matches!(cell.value, CellValue::Empty) && cell.formula.is_none() {
        output.push_str("/>");
        return Ok(());
    }
    output.push('>');
    if !cell.text.is_empty() || matches!(cell.value, CellValue::Text(_)) {
        output.push_str("<text:p>");
        output.push_str(&escape_xml(&cell.text));
        output.push_str("</text:p>");
    }
    output.push_str(if covered {
        "</table:covered-table-cell>"
    } else {
        "</table:table-cell>"
    });
    Ok(())
}

fn write_value_attributes(output: &mut String, value: &CellValue) {
    match value {
        CellValue::Empty => {},
        CellValue::Text(_) => output.push_str(" office:value-type=\"string\""),
        CellValue::Number(value) => {
            output.push_str(" office:value-type=\"float\" office:value=\"");
            write!(output, "{value}").expect("writing a number to String cannot fail");
            output.push('"');
        },
        CellValue::Currency { value, currency } => {
            output.push_str(" office:value-type=\"currency\" office:value=\"");
            write!(output, "{value}").expect("writing a number to String cannot fail");
            output.push_str("\" office:currency=\"");
            output.push_str(&escape_xml(currency));
            output.push('"');
        },
        CellValue::Percentage(value) => {
            output.push_str(" office:value-type=\"percentage\" office:value=\"");
            write!(output, "{value}").expect("writing a number to String cannot fail");
            output.push('"');
        },
        CellValue::Boolean(value) => {
            output.push_str(" office:value-type=\"boolean\" office:boolean-value=\"");
            output.push_str(if *value { "true" } else { "false" });
            output.push('"');
        },
        CellValue::Date(value) => {
            output.push_str(" office:value-type=\"date\" office:date-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Time(value) => {
            output.push_str(" office:value-type=\"time\" office:time-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Unknown { kind, value } => {
            output.push_str(" office:value-type=\"");
            output.push_str(&escape_xml(kind));
            output.push('"');
            if let Some(value) = value {
                output.push_str(" office:value=\"");
                output.push_str(&escape_xml(value));
                output.push('"');
            }
        },
    }
}
