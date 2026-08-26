//! Streaming ODF worksheet parser and canonical table writer.

use super::{Cell, CellValue, Row, Sheet};
use super::{model::Merge, validation};
use crate::model::hyperlink::{Actuate, Link, Show};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{
    XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, ResolveResult},
    reader::NsReader,
};
use std::num::NonZeroUsize;

pub(crate) const OFFICE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
pub(crate) const TABLE_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(crate) const TEXT_NAMESPACE: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
pub(crate) const XLINK_NAMESPACE: &str = "http://www.w3.org/1999/xlink";
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
pub(crate) enum NamespaceKind {
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
    /// Decode from a standalone reader; used by the historical inline loop in
    /// `parse_impl`, which predates the extracted handler.
    fn from_element(
        element: &BytesStart<'_>,
        reader: &NsReader<&[u8]>,
        covered: bool,
    ) -> Result<Self> {
        Self::from_resolved(element, reader.resolver(), reader.decoder(), covered)
    }

    fn from_resolved(
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        covered: bool,
    ) -> Result<Self> {
        let mut result = Self {
            covered,
            ..Self::default()
        };
        for raw in element.attributes().with_checks(true) {
            let raw = raw
                .map_err(|error| Error::InvalidFormat(format!("invalid ODS attribute: {error}")))?;
            let (namespace, local) = resolver.resolve_attribute(raw.key);
            let local = local.as_ref();
            // Decode and normalize every attribute value, including values of
            // attributes this codec ignores: malformed entity and character
            // references are rejected here, ahead of the unknown-prefix error,
            // regardless of whether the value is consumed.  Only consumed
            // values are copied into an owned `String`.
            let value = raw
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODS attribute value: {error}"))
                })?;
            let namespace: &[u8] = match namespace {
                ResolveResult::Bound(Namespace(uri)) => uri,
                ResolveResult::Unbound => b"",
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::InvalidFormat(format!(
                        "unbound ODS attribute prefix '{}'",
                        String::from_utf8_lossy(prefix.as_ref())
                    )));
                },
            };
            if namespace == TABLE_NAMESPACE.as_bytes() {
                match local {
                    b"name" => result.name = Some(value.into_owned()),
                    b"style-name" => result.style_name = Some(value.into_owned()),
                    b"default-cell-style-name" => {
                        result.default_cell_style_name = Some(value.into_owned())
                    },
                    b"formula" => result.formula = Some(value.into_owned()),
                    b"number-rows-repeated" => {
                        result.rows_repeated = positive(&value, "number-rows-repeated")?;
                    },
                    b"number-columns-repeated" => {
                        result.columns_repeated = positive(&value, "number-columns-repeated")?;
                    },
                    b"number-rows-spanned" => {
                        result.rows_spanned = positive(&value, "number-rows-spanned")?;
                    },
                    b"number-columns-spanned" => {
                        result.columns_spanned = positive(&value, "number-columns-spanned")?;
                    },
                    _ => {},
                }
            } else if namespace == OFFICE_NAMESPACE.as_bytes() {
                match local {
                    b"value-type" => result.value_type = Some(value.into_owned()),
                    b"value" => result.value = Some(value.into_owned()),
                    b"date-value" => result.date_value = Some(value.into_owned()),
                    b"time-value" => result.time_value = Some(value.into_owned()),
                    b"boolean-value" => result.boolean_value = Some(value.into_owned()),
                    b"currency" => result.currency = Some(value.into_owned()),
                    _ => {},
                }
            }
        }
        Ok(result)
    }

    fn cell(self, text: String, hyperlinks: Vec<Link>) -> Result<Cell> {
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
            Some("float" | "double" | "decimal") => CellValue::Number(parse_float(
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
        cell.hyperlinks = hyperlinks;
        cell.merge = if self.covered {
            Merge::Covered
        } else if self.rows_spanned != 1 || self.columns_spanned != 1 {
            Merge::Span {
                rows: NonZeroUsize::new(self.rows_spanned).ok_or_else(|| {
                    Error::InvalidFormat("table:number-rows-spanned must be positive".to_string())
                })?,
                columns: NonZeroUsize::new(self.columns_spanned).ok_or_else(|| {
                    Error::InvalidFormat(
                        "table:number-columns-spanned must be positive".to_string(),
                    )
                })?,
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
    paragraph_count: usize,
    paragraph_open: bool,
    text_depth: usize,
    hyperlinks: Vec<Link>,
    open_hyperlink: Option<OpenHyperlink>,
    ignored_anchor_depth: usize,
}

struct OpenHyperlink {
    link: Link,
    start: usize,
}

impl OpenCell {
    fn start_text(
        &mut self,
        local: &[u8],
        parent: Option<Kind>,
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
    ) -> Result<()> {
        if local == b"p" {
            if parent == Some(Kind::Cell) {
                self.paragraph_count = self.paragraph_count.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("ODS cell paragraph count overflows usize".to_string())
                })?;
                self.paragraph_open = true;
            } else if self.paragraph_open {
                return Err(Error::InvalidFormat(
                    "ODS cell paragraphs cannot be nested".to_string(),
                ));
            }
        } else if local == b"a" {
            if self.open_hyperlink.is_some() {
                return Err(Error::InvalidFormat(
                    "ODS cell hyperlinks must be direct text:p children and cannot nest"
                        .to_string(),
                ));
            }
            if self.paragraph_open
                && self.ignored_anchor_depth == 0
                && parent == Some(Kind::Text)
                && self.text_depth == 1
            {
                let link = parse_link(element, resolver, decoder)?;
                self.open_hyperlink = Some(OpenHyperlink {
                    link,
                    start: self.text.len(),
                });
            } else {
                self.ignored_anchor_depth =
                    self.ignored_anchor_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODS cell ignored hyperlink nesting overflows usize".to_string(),
                        )
                    })?;
            }
        } else if self.open_hyperlink.is_some() {
            return Err(Error::InvalidFormat(
                "ODS cell hyperlinks may contain character data only".to_string(),
            ));
        }
        self.text_depth = self.text_depth.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("ODS cell inline text depth overflows usize".to_string())
        })?;
        Ok(())
    }

    fn handle_empty_text(
        &mut self,
        local: &[u8],
        parent: Option<Kind>,
        element: &BytesStart<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
    ) -> Result<()> {
        self.start_text(local, parent, element, resolver, decoder)?;
        self.end_text(local)
    }

    fn append_characters(&mut self, value: &str) -> Result<()> {
        let next = self.text.len().checked_add(value.len()).ok_or_else(|| {
            Error::InvalidFormat("ODS cell text length overflows usize".to_string())
        })?;
        if next > validation::MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "ODS cell text exceeds the worksheet text safety limit".to_string(),
            ));
        }
        self.text.try_reserve(value.len()).map_err(|_error| {
            Error::InvalidFormat("ODS cell text allocation failed".to_string())
        })?;
        self.text.push_str(value);
        if let Some(open) = self.open_hyperlink.as_mut() {
            open.link.text.try_reserve(value.len()).map_err(|_error| {
                Error::InvalidFormat("ODS hyperlink text allocation failed".to_string())
            })?;
            open.link.text.push_str(value);
        }
        Ok(())
    }

    fn end_text(&mut self, local: &[u8]) -> Result<()> {
        if local == b"a" {
            if let Some(mut open) = self.open_hyperlink.take() {
                let range = open.start..self.text.len();
                let anchor = self.text.get(range.clone()).ok_or_else(|| {
                    Error::InvalidFormat("ODS hyperlink range is not valid UTF-8".to_string())
                })?;
                if anchor != open.link.text {
                    return Err(Error::InvalidFormat(
                        "ODS hyperlink text does not match its cell text range".to_string(),
                    ));
                }
                open.link.set_range(range);
                open.link.validate_storage()?;
                if self.hyperlinks.len() >= validation::MAX_PHYSICAL_RUNS {
                    return Err(Error::InvalidFormat(format!(
                        "ODS cell exceeds the {} hyperlink safety limit",
                        validation::MAX_PHYSICAL_RUNS
                    )));
                }
                self.hyperlinks.try_reserve(1).map_err(|_error| {
                    Error::InvalidFormat("ODS cell hyperlink allocation failed".to_string())
                })?;
                self.hyperlinks.push(open.link);
            } else if self.ignored_anchor_depth > 0 {
                self.ignored_anchor_depth -= 1;
            }
        }
        if self.text_depth == 0 {
            return Err(Error::InvalidFormat(
                "ODS cell text element depth underflow".to_string(),
            ));
        }
        self.text_depth -= 1;
        if local == b"p" && self.paragraph_open {
            if self.open_hyperlink.is_some() {
                return Err(Error::InvalidFormat(
                    "ODS cell paragraph closed while a hyperlink is open".to_string(),
                ));
            }
            self.paragraph_open = false;
        }
        Ok(())
    }
}

fn reserve_parser_push<T>(
    collection: &mut Vec<T>,
    limit: usize,
    limit_error: String,
    allocation_error: &str,
) -> Result<()> {
    if collection.len() >= limit {
        return Err(Error::InvalidFormat(limit_error));
    }
    collection
        .try_reserve(1)
        .map_err(|_| Error::InvalidFormat(allocation_error.to_string()))?;
    Ok(())
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

/// Shared worksheet parse over one dedicated reader.
///
/// This standalone path intentionally keeps the historical inline event loop
/// for latency parity on non-fused callers (eager open, commit readback); the
/// fused open parse drives the equivalent [`WorksheetHandler`] instead.
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
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"
                    )));
                }
                let local = element.local_name();
                let parent = stack.last().copied();
                let mut kind = classify(namespace, local.as_ref());
                if dde_cache_depth.is_some() {
                    kind = Kind::Other;
                } else if kind == Kind::Table && stack.last() == Some(&Kind::DdeLink) {
                    kind = Kind::DdeCache;
                    dde_cache_depth = Some(stack.len() + 1);
                }
                match kind {
                    Kind::Root => {},
                    Kind::Body | Kind::Spreadsheet | Kind::DdeLink | Kind::DdeCache => {},
                    Kind::Other => {
                        if current_cell
                            .as_ref()
                            .is_some_and(|cell| cell.open_hyperlink.is_some())
                        {
                            return Err(Error::InvalidFormat(
                                "ODS cell refuses foreign or unsupported markup".to_string(),
                            ));
                        }
                    },
                    Kind::Text => {
                        if let Some(cell) = current_cell.as_mut() {
                            cell.start_text(
                                local.as_ref(),
                                parent,
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                            )?;
                        }
                    },
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
                            repeat: NonZeroUsize::new(attributes.rows_repeated).ok_or_else(
                                || {
                                    Error::InvalidFormat(
                                        "table:number-rows-repeated must be positive".to_string(),
                                    )
                                },
                            )?,
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
                            paragraph_count: 0,
                            paragraph_open: false,
                            text_depth: 0,
                            hyperlinks: Vec::new(),
                            open_hyperlink: None,
                            ignored_anchor_depth: 0,
                        });
                    },
                }
                reserve_parser_push(
                    &mut stack,
                    MAX_XML_DEPTH,
                    format!("ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"),
                    "ODS worksheet XML stack allocation failed",
                )?;
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
                        let sheet = Sheet {
                            name: attributes.name.unwrap_or_else(|| "Sheet1".to_string()),
                            rows: Vec::new(),
                            style_name: attributes.style_name,
                        };
                        reserve_parser_push(
                            &mut sheets,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet sheet allocation failed",
                        )?;
                        sheets.push(sheet);
                    },
                    Kind::Row => {
                        let attributes = Attributes::from_element(&element, &reader, false)?;
                        let sheet = current_sheet.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "empty table row is outside a worksheet".to_string(),
                            )
                        })?;
                        let row = Row {
                            cells: Vec::new(),
                            style_name: attributes.style_name,
                            default_cell_style_name: attributes.default_cell_style_name,
                            repeat: NonZeroUsize::new(attributes.rows_repeated).ok_or_else(
                                || {
                                    Error::InvalidFormat(
                                        "table:number-rows-repeated must be positive".to_string(),
                                    )
                                },
                            )?,
                        };
                        reserve_parser_push(
                            &mut sheet.rows,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet row count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet row allocation failed",
                        )?;
                        sheet.rows.push(row);
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
                        let cell = attributes.cell(String::new(), Vec::new())?;
                        reserve_parser_push(
                            &mut row.cells,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS row cell count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet cell allocation failed",
                        )?;
                        row.cells.push(cell);
                    },
                    Kind::Text if current_cell.is_some() => {
                        let parent = stack.last().copied();
                        if let Some(cell) = current_cell.as_mut() {
                            cell.handle_empty_text(
                                local.as_ref(),
                                parent,
                                &element,
                                reader.resolver(),
                                reader.decoder(),
                            )?;
                        }
                    },
                    Kind::Other => {
                        if current_cell
                            .as_ref()
                            .is_some_and(|cell| cell.open_hyperlink.is_some())
                        {
                            return Err(Error::InvalidFormat(
                                "ODS cell refuses foreign or unsupported markup".to_string(),
                            ));
                        }
                    },
                    Kind::Root
                    | Kind::Body
                    | Kind::Spreadsheet
                    | Kind::DdeLink
                    | Kind::DdeCache
                    | Kind::Text => {},
                }
            },
            Event::Text(text) => {
                if let Some(cell) = current_cell.as_mut() {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.append_characters(&value)?;
                }
            },
            Event::CData(text) => {
                if let Some(cell) = current_cell.as_mut() {
                    let value = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.append_characters(&value)?;
                }
            },
            Event::End(element) => {
                let kind = stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                if kind == Kind::DdeCache {
                    dde_cache_depth = None;
                }
                if kind == Kind::Text
                    && let Some(cell) = current_cell.as_mut()
                {
                    cell.end_text(element.local_name().as_ref())?;
                }
                match kind {
                    Kind::Cell => {
                        let open = current_cell.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell close has no open cell".to_string())
                        })?;
                        let row = current_row.as_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell closed outside a row".to_string())
                        })?;
                        let cell = open.attributes.cell(open.text, open.hyperlinks)?;
                        reserve_parser_push(
                            &mut row.cells,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS row cell count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet cell allocation failed",
                        )?;
                        row.cells.push(cell);
                    },
                    Kind::Row => {
                        let row = current_row.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS row close has no open row".to_string())
                        })?;
                        let sheet = current_sheet.as_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODS row closed outside a sheet".to_string())
                        })?;
                        reserve_parser_push(
                            &mut sheet.rows,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet row count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet row allocation failed",
                        )?;
                        sheet.rows.push(row);
                    },
                    Kind::Table => {
                        let sheet = current_sheet.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS table close has no open table".to_string())
                        })?;
                        reserve_parser_push(
                            &mut sheets,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet sheet allocation failed",
                        )?;
                        sheets.push(sheet);
                    },
                    Kind::Root
                    | Kind::Body
                    | Kind::Spreadsheet
                    | Kind::DdeLink
                    | Kind::DdeCache
                    | Kind::Text
                    | Kind::Other => {},
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

/// Streaming event handler holding the [`parse_impl`] state.
///
/// The fused open parse ([`crate::open_parse`]) drives one shared tokenizer
/// through this handler, while [`parse_impl`] keeps its historical inline
/// loop for latency parity on non-fused callers; both apply the same checks,
/// limits, and error messages at the same events.
pub(crate) struct WorksheetHandler {
    stack: Vec<Kind>,
    dde_cache_depth: Option<usize>,
    sheets: Vec<Sheet>,
    current_sheet: Option<Sheet>,
    current_row: Option<Row>,
    current_cell: Option<OpenCell>,
    require_unique_names: bool,
}

impl WorksheetHandler {
    /// Create a handler with the [`parse_impl`] validation mode.
    pub(crate) fn new(require_unique_names: bool) -> Self {
        Self {
            stack: Vec::new(),
            dde_cache_depth: None,
            sheets: Vec::new(),
            current_sheet: None,
            current_row: None,
            current_cell: None,
            require_unique_names,
        }
    }

    /// Process one resolved event at byte positions `pos_before`/`pos_after`.
    ///
    /// `namespace` is the caller-classified resolution of the event's
    /// namespace; the resolved value borrows the reader mutably, so callers
    /// classify it immediately after the read exactly as the historical loop
    /// body did.
    pub(crate) fn on_event(
        &mut self,
        namespace: NamespaceKind,
        event: &Event<'_>,
        resolver: &NamespaceResolver,
        decoder: Decoder,
        pos_before: u64,
        pos_after: u64,
    ) -> Result<()> {
        let _ = (pos_before, pos_after);
        match event {
            Event::Start(element) => {
                if self.stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"
                    )));
                }
                let local = element.local_name();
                let mut kind = classify(namespace, local.as_ref());
                if self.dde_cache_depth.is_some() {
                    kind = Kind::Other;
                } else if kind == Kind::Table && self.stack.last() == Some(&Kind::DdeLink) {
                    kind = Kind::DdeCache;
                    self.dde_cache_depth = Some(self.stack.len() + 1);
                }
                match kind {
                    Kind::Root => {},
                    Kind::Body | Kind::Spreadsheet | Kind::DdeLink | Kind::DdeCache => {},
                    Kind::Other => {
                        if self
                            .current_cell
                            .as_ref()
                            .is_some_and(|cell| cell.open_hyperlink.is_some())
                        {
                            return Err(Error::InvalidFormat(
                                "ODS cell refuses foreign or unsupported markup".to_string(),
                            ));
                        }
                    },
                    Kind::Text => {
                        if let Some(cell) = self.current_cell.as_mut() {
                            cell.start_text(
                                local.as_ref(),
                                self.stack.last().copied(),
                                element,
                                resolver,
                                decoder,
                            )?;
                        }
                    },
                    Kind::Table => {
                        if self.stack.last() != Some(&Kind::Spreadsheet) {
                            return Err(Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            ));
                        }
                        if self.current_sheet.is_some() {
                            return Err(Error::InvalidFormat(
                                "ODS worksheet parser encountered a nested table".to_string(),
                            ));
                        }
                        let attributes =
                            Attributes::from_resolved(element, resolver, decoder, false)?;
                        self.current_sheet = Some(Sheet {
                            name: attributes.name.unwrap_or_else(|| "Sheet1".to_string()),
                            rows: Vec::new(),
                            style_name: attributes.style_name,
                        });
                    },
                    Kind::Row => {
                        if self.current_sheet.is_none() || self.current_row.is_some() {
                            return Err(Error::InvalidFormat(
                                "table:table-row is outside a worksheet row context".to_string(),
                            ));
                        }
                        let attributes =
                            Attributes::from_resolved(element, resolver, decoder, false)?;
                        self.current_row = Some(Row {
                            cells: Vec::new(),
                            style_name: attributes.style_name,
                            default_cell_style_name: attributes.default_cell_style_name,
                            repeat: NonZeroUsize::new(attributes.rows_repeated).ok_or_else(
                                || {
                                    Error::InvalidFormat(
                                        "table:number-rows-repeated must be positive".to_string(),
                                    )
                                },
                            )?,
                        });
                    },
                    Kind::Cell => {
                        if self.current_row.is_none() || self.current_cell.is_some() {
                            return Err(Error::InvalidFormat(
                                "table cell is outside a worksheet row context".to_string(),
                            ));
                        }
                        self.current_cell = Some(OpenCell {
                            attributes: Attributes::from_resolved(
                                element,
                                resolver,
                                decoder,
                                is_covered(namespace, local.as_ref()),
                            )?,
                            text: String::new(),
                            paragraph_count: 0,
                            paragraph_open: false,
                            text_depth: 0,
                            hyperlinks: Vec::new(),
                            open_hyperlink: None,
                            ignored_anchor_depth: 0,
                        });
                    },
                }
                reserve_parser_push(
                    &mut self.stack,
                    MAX_XML_DEPTH,
                    format!("ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"),
                    "ODS worksheet XML stack allocation failed",
                )?;
                self.stack.push(kind);
            },
            Event::Empty(element) => {
                if self.stack.len() >= MAX_XML_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "ODS worksheet XML nesting exceeds {MAX_XML_DEPTH} elements"
                    )));
                }
                let local = element.local_name();
                let mut kind = classify(namespace, local.as_ref());
                if self.dde_cache_depth.is_some() {
                    kind = Kind::Other;
                } else if kind == Kind::Table && self.stack.last() == Some(&Kind::DdeLink) {
                    kind = Kind::DdeCache;
                }
                match kind {
                    Kind::Table => {
                        if self.stack.last() != Some(&Kind::Spreadsheet) {
                            return Err(Error::InvalidFormat(
                                "table:table must be a direct child of office:spreadsheet"
                                    .to_string(),
                            ));
                        }
                        let attributes =
                            Attributes::from_resolved(element, resolver, decoder, false)?;
                        let sheet = Sheet {
                            name: attributes.name.unwrap_or_else(|| "Sheet1".to_string()),
                            rows: Vec::new(),
                            style_name: attributes.style_name,
                        };
                        reserve_parser_push(
                            &mut self.sheets,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet sheet allocation failed",
                        )?;
                        self.sheets.push(sheet);
                    },
                    Kind::Row => {
                        let attributes =
                            Attributes::from_resolved(element, resolver, decoder, false)?;
                        let sheet = self.current_sheet.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "empty table row is outside a worksheet".to_string(),
                            )
                        })?;
                        let row = Row {
                            cells: Vec::new(),
                            style_name: attributes.style_name,
                            default_cell_style_name: attributes.default_cell_style_name,
                            repeat: NonZeroUsize::new(attributes.rows_repeated).ok_or_else(
                                || {
                                    Error::InvalidFormat(
                                        "table:number-rows-repeated must be positive".to_string(),
                                    )
                                },
                            )?,
                        };
                        reserve_parser_push(
                            &mut sheet.rows,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet row count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet row allocation failed",
                        )?;
                        sheet.rows.push(row);
                    },
                    Kind::Cell => {
                        let row = self.current_row.as_mut().ok_or_else(|| {
                            Error::InvalidFormat(
                                "empty table cell is outside a worksheet row".to_string(),
                            )
                        })?;
                        let attributes = Attributes::from_resolved(
                            element,
                            resolver,
                            decoder,
                            is_covered(namespace, local.as_ref()),
                        )?;
                        let cell = attributes.cell(String::new(), Vec::new())?;
                        reserve_parser_push(
                            &mut row.cells,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS row cell count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet cell allocation failed",
                        )?;
                        row.cells.push(cell);
                    },
                    Kind::Text if self.current_cell.is_some() => {
                        let parent = self.stack.last().copied();
                        if let Some(cell) = self.current_cell.as_mut() {
                            cell.handle_empty_text(
                                local.as_ref(),
                                parent,
                                element,
                                resolver,
                                decoder,
                            )?;
                        }
                    },
                    Kind::Other => {
                        if self
                            .current_cell
                            .as_ref()
                            .is_some_and(|cell| cell.open_hyperlink.is_some())
                        {
                            return Err(Error::InvalidFormat(
                                "ODS cell refuses foreign or unsupported markup".to_string(),
                            ));
                        }
                    },
                    Kind::Root
                    | Kind::Body
                    | Kind::Spreadsheet
                    | Kind::DdeLink
                    | Kind::DdeCache
                    | Kind::Text => {},
                }
            },
            Event::Text(text) => {
                if let Some(cell) = self.current_cell.as_mut() {
                    let value = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.append_characters(&value)?;
                }
            },
            Event::CData(text) => {
                if let Some(cell) = self.current_cell.as_mut() {
                    let value = text.decode().map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS cell text: {error}"))
                    })?;
                    cell.append_characters(&value)?;
                }
            },
            Event::End(element) => {
                let kind = self.stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("ODS XML element stack underflow".to_string())
                })?;
                if kind == Kind::DdeCache {
                    self.dde_cache_depth = None;
                }
                if kind == Kind::Text
                    && let Some(cell) = self.current_cell.as_mut()
                {
                    cell.end_text(element.local_name().as_ref())?;
                }
                match kind {
                    Kind::Cell => {
                        let open = self.current_cell.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell close has no open cell".to_string())
                        })?;
                        let row = self.current_row.as_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODS cell closed outside a row".to_string())
                        })?;
                        let cell = open.attributes.cell(open.text, open.hyperlinks)?;
                        reserve_parser_push(
                            &mut row.cells,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS row cell count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet cell allocation failed",
                        )?;
                        row.cells.push(cell);
                    },
                    Kind::Row => {
                        let row = self.current_row.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS row close has no open row".to_string())
                        })?;
                        let sheet = self.current_sheet.as_mut().ok_or_else(|| {
                            Error::InvalidFormat("ODS row closed outside a sheet".to_string())
                        })?;
                        reserve_parser_push(
                            &mut sheet.rows,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet row count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet row allocation failed",
                        )?;
                        sheet.rows.push(row);
                    },
                    Kind::Table => {
                        let sheet = self.current_sheet.take().ok_or_else(|| {
                            Error::InvalidFormat("ODS table close has no open table".to_string())
                        })?;
                        reserve_parser_push(
                            &mut self.sheets,
                            validation::MAX_PHYSICAL_RUNS,
                            format!(
                                "ODS sheet count exceeds the {} safety limit",
                                validation::MAX_PHYSICAL_RUNS
                            ),
                            "ODS worksheet sheet allocation failed",
                        )?;
                        self.sheets.push(sheet);
                    },
                    Kind::Root
                    | Kind::Body
                    | Kind::Spreadsheet
                    | Kind::DdeLink
                    | Kind::DdeCache
                    | Kind::Text
                    | Kind::Other => {},
                }
                let _ = element;
            },
            Event::Eof => {},
            Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        Ok(())
    }

    /// Validate the end-of-document state and build the worksheet list.
    pub(crate) fn finish(self) -> Result<Vec<Sheet>> {
        if !self.stack.is_empty()
            || self.current_sheet.is_some()
            || self.current_row.is_some()
            || self.current_cell.is_some()
        {
            return Err(Error::InvalidFormat(
                "ODS content ended with an unfinished worksheet object".to_string(),
            ));
        }
        let sheets = self.sheets;
        if self.require_unique_names {
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

#[cfg(test)]
mod attribute_error_order_tests {
    use super::parse;

    const PREFIX: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="urn:litchi:test"><office:body><office:spreadsheet><table:table table:name="S"><table:table-row>"#;
    const SUFFIX: &str = "</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>";

    fn document(cell: &str) -> String {
        format!("{PREFIX}{cell}{SUFFIX}")
    }

    #[test]
    fn ignored_attributes_with_clean_values_are_accepted() {
        // Bound-but-foreign and unbound attribute namespaces are ignored.
        let xml = document(r#"<table:table-cell x:ignored="1" plain="2"/>"#);
        parse(&xml).expect("ignored attributes are inert");
    }

    #[test]
    fn malformed_entity_in_ignored_attribute_is_rejected() {
        // The value decode runs for ignored attributes as well.
        for value in ["&bogus;", "&bogus", "&#xD800;", "&#x110000;"] {
            let xml = document(&format!(r#"<table:table-cell x:ignored="{value}"/>"#));
            let error = parse(&xml).expect_err("malformed value must fail: {value}");
            assert!(
                error
                    .to_string()
                    .starts_with("Invalid format: invalid ODS attribute value: "),
                "unexpected error for {value}: {error}"
            );
        }
    }

    #[test]
    fn value_decode_error_beats_unknown_prefix_error() {
        // Historical per-attribute order is syntax, then value decode, then
        // the unknown-prefix check; an attribute failing both must report the
        // decode error.
        let xml = document(r#"<table:table-cell undeclared:ignored="&bogus;"/>"#);
        let error = parse(&xml).expect_err("malformed value must fail");
        assert!(
            error
                .to_string()
                .starts_with("Invalid format: invalid ODS attribute value: "),
            "unexpected error: {error}"
        );
    }

    #[test]
    fn unknown_prefix_on_ignored_attribute_is_rejected() {
        let xml = document(r#"<table:table-cell undeclared:ignored="clean"/>"#);
        let error = parse(&xml).expect_err("undeclared prefix must fail");
        assert_eq!(
            error.to_string(),
            "Invalid format: unbound ODS attribute prefix 'undeclared'"
        );
    }

    #[test]
    fn consumed_values_keep_entity_normalization() {
        let xml = document(
            r#"<table:table-cell office:value-type="string" table:style-name="a&#x20;b"><text:p>x</text:p></table:table-cell>"#,
        );
        let sheets = parse(&xml).expect("consumed values decode");
        assert_eq!(
            sheets[0].rows[0].cells[0].style_name.as_deref(),
            Some("a b")
        );
    }
}

pub(crate) fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
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
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
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

fn parse_link(
    element: &BytesStart<'_>,
    resolver: &NamespaceResolver,
    decoder: Decoder,
) -> Result<Link> {
    let mut href = None;
    let mut link_type = None;
    let mut show = None;
    let mut actuate = None;
    let mut name = None;
    let mut title = None;
    let mut target_frame_name = None;
    let mut style_name = None;
    let mut visited_style_name = None;

    for raw in element.attributes().with_checks(true) {
        let raw = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODS text:a attribute: {error}"))
        })?;
        if raw.key.as_ref() == b"xmlns" || raw.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        if raw.value.len() > validation::MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "ODS text:a attribute value exceeds the worksheet text safety limit".to_string(),
            ));
        }
        let value = raw
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid ODS text:a attribute value: {error}"))
            })?;
        if value.len() > validation::MAX_TEXT_BYTES {
            return Err(Error::InvalidFormat(
                "ODS text:a attribute value exceeds the worksheet text safety limit".to_string(),
            ));
        }
        let value = value.into_owned();
        let (namespace, local) = resolver.resolve_attribute(raw.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(uri)) => uri,
            ResolveResult::Unbound => {
                return Err(Error::InvalidFormat(format!(
                    "ODS text:a attribute '{}' is unbound",
                    String::from_utf8_lossy(local.as_ref())
                )));
            },
            ResolveResult::Unknown(prefix) => {
                return Err(Error::InvalidFormat(format!(
                    "unbound ODS text:a attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        if namespace == XLINK_NAMESPACE.as_bytes() {
            match local.as_ref() {
                b"href" => href = Some(value),
                b"type" => link_type = Some(value),
                b"show" => {
                    show = Some(Show::parse(&value).ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODS text:a xlink:show must be 'new' or 'replace'".to_string(),
                        )
                    })?);
                },
                b"actuate" => {
                    actuate = Some(Actuate::parse(&value).ok_or_else(|| {
                        Error::InvalidFormat(
                            "ODS text:a xlink:actuate must be 'onRequest'".to_string(),
                        )
                    })?);
                },
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "ODS text:a refuses unknown xlink attribute '{}'",
                        String::from_utf8_lossy(local.as_ref())
                    )));
                },
            }
        } else if namespace == OFFICE_NAMESPACE.as_bytes() {
            match local.as_ref() {
                b"name" => name = Some(value),
                b"title" => title = Some(value),
                b"target-frame-name" => target_frame_name = Some(value),
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "ODS text:a refuses unknown office attribute '{}'",
                        String::from_utf8_lossy(local.as_ref())
                    )));
                },
            }
        } else if namespace == TEXT_NAMESPACE.as_bytes() {
            match local.as_ref() {
                b"style-name" => style_name = Some(value),
                b"visited-style-name" => visited_style_name = Some(value),
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "ODS text:a refuses unknown text attribute '{}'",
                        String::from_utf8_lossy(local.as_ref())
                    )));
                },
            }
        } else {
            return Err(Error::InvalidFormat(format!(
                "ODS text:a refuses unknown attribute '{}'",
                String::from_utf8_lossy(local.as_ref())
            )));
        }
    }

    let href =
        href.ok_or_else(|| Error::InvalidFormat("ODS text:a requires xlink:href".to_string()))?;
    if link_type
        .as_deref()
        .is_some_and(|link_type| link_type != "simple")
    {
        return Err(Error::InvalidFormat(
            "ODS text:a requires xlink:type='simple'".to_string(),
        ));
    }
    let mut link = Link::new(href);
    link.show = show;
    link.actuate = actuate;
    link.name = name;
    link.title = title;
    link.target_frame_name = target_frame_name;
    link.style_name = style_name;
    link.visited_style_name = visited_style_name;
    link.validate_storage()?;
    Ok(link)
}

fn positive(value: &str, name: &str) -> Result<usize> {
    let value = value
        .parse::<usize>()
        .map_err(|_error| Error::InvalidFormat(format!("ODS {name} must be a positive integer")))?;
    NonZeroUsize::new(value)
        .map(NonZeroUsize::get)
        .ok_or_else(|| Error::InvalidFormat(format!("ODS {name} must be positive")))
}

fn parse_float(value: &str, name: &str) -> Result<f64> {
    let value = value.parse::<f64>().map_err(|_error| {
        Error::InvalidFormat(format!("ODS {name} requires a finite decimal value"))
    })?;
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
    output.push_str("\" xmlns:xlink=\"");
    output.push_str(XLINK_NAMESPACE);
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

/// Render one worksheet under an exact allocation and output byte bound.
pub(crate) fn write_sheet_bounded(sheet: &Sheet, max_bytes: usize) -> Result<String> {
    validation::validate_sheet(sheet)?;
    let mut output = String::new();
    bounded_push(&mut output, "<table:table xmlns:table=\"", max_bytes)?;
    bounded_push(&mut output, TABLE_NAMESPACE, max_bytes)?;
    bounded_push(&mut output, "\" xmlns:office=\"", max_bytes)?;
    bounded_push(&mut output, OFFICE_NAMESPACE, max_bytes)?;
    bounded_push(&mut output, "\" xmlns:text=\"", max_bytes)?;
    bounded_push(&mut output, TEXT_NAMESPACE, max_bytes)?;
    bounded_push(&mut output, "\" xmlns:xlink=\"", max_bytes)?;
    bounded_push(&mut output, XLINK_NAMESPACE, max_bytes)?;
    bounded_push(&mut output, "\" table:name=\"", max_bytes)?;
    bounded_push(&mut output, &escape_xml(&sheet.name), max_bytes)?;
    bounded_push(&mut output, "\"", max_bytes)?;
    if let Some(style_name) = &sheet.style_name {
        bounded_push(&mut output, " table:style-name=\"", max_bytes)?;
        bounded_push(&mut output, &escape_xml(style_name), max_bytes)?;
        bounded_push(&mut output, "\"", max_bytes)?;
    }
    if sheet.rows.is_empty() {
        bounded_push(&mut output, "/>", max_bytes)?;
        return Ok(output);
    }
    bounded_push(&mut output, ">", max_bytes)?;
    for row in &sheet.rows {
        write_row_bounded(&mut output, row, max_bytes, false)?;
    }
    bounded_push(&mut output, "</table:table>", max_bytes)?;
    Ok(output)
}

/// Render row fragments under an exact allocation and output byte bound.
pub(crate) fn write_rows_bounded(rows: &[Row], max_bytes: usize) -> Result<String> {
    let mut output = String::new();
    for row in rows {
        write_row_bounded(&mut output, row, max_bytes, true)?;
    }
    Ok(output)
}

fn bounded_push(output: &mut String, value: &str, max_bytes: usize) -> Result<()> {
    let next = output.len().checked_add(value.len()).ok_or_else(|| {
        Error::InvalidFormat("flat ODS rendered size overflows usize".to_string())
    })?;
    if next > max_bytes {
        return Err(Error::InvalidFormat(format!(
            "flat ODS rendered worksheet content exceeds the {max_bytes} byte limit"
        )));
    }
    output.try_reserve(value.len()).map_err(|_error| {
        Error::InvalidFormat("flat ODS worksheet rendering allocation failed".to_string())
    })?;
    output.push_str(value);
    Ok(())
}

fn write_row_bounded(
    output: &mut String,
    row: &Row,
    max_bytes: usize,
    bind_namespaces: bool,
) -> Result<()> {
    validation::validate_cell_runs(&row.cells)?;
    bounded_push(output, "<table:table-row", max_bytes)?;
    if bind_namespaces {
        bounded_push(
            output,
            concat!(
                " xmlns:table=\"",
                "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
                "\" xmlns:office=\"",
                "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
                "\" xmlns:text=\"",
                "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
                "\" xmlns:xlink=\"",
                "http://www.w3.org/1999/xlink",
                "\""
            ),
            max_bytes,
        )?;
    }
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
    if cell.text.is_empty()
        && cell.hyperlinks.is_empty()
        && matches!(cell.value, CellValue::Empty)
        && cell.formula.is_none()
    {
        return bounded_push(output, "/>", max_bytes);
    }
    bounded_push(output, ">", max_bytes)?;
    if !cell.text.is_empty()
        || !cell.hyperlinks.is_empty()
        || matches!(cell.value, CellValue::Text(_))
    {
        bounded_push(
            output,
            if requires_xml_space_preserve(&cell.text) {
                "<text:p xml:space=\"preserve\">"
            } else {
                "<text:p>"
            },
            max_bytes,
        )?;
        write_cell_text_bounded(output, &cell.text, &cell.hyperlinks, max_bytes)?;
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
        output.push_str("\" xmlns:xlink=\"");
        output.push_str(XLINK_NAMESPACE);
        output.push('"');
    }
    if row.repeat() > 1 {
        output.push_str(" table:number-rows-repeated=\"");
        output.push_str(&row.repeat().to_string());
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
    write_cell_inner(output, cell, None, false)
}

pub(crate) fn write_cell_fragment(cell: &Cell, body: Option<&str>) -> Result<String> {
    let mut output = String::new();
    write_cell_inner(&mut output, cell, body, true)?;
    Ok(output)
}

fn write_cell_inner(
    output: &mut String,
    cell: &Cell,
    body: Option<&str>,
    bind_namespaces: bool,
) -> Result<()> {
    validation::validate_cell(cell)?;
    if body.is_some() && !cell.hyperlinks.is_empty() {
        return Err(Error::InvalidFormat(
            "ODS cell fragment refuses to replace modeled hyperlinks with an external body"
                .to_string(),
        ));
    }
    let covered = matches!(cell.merge, Merge::Covered);
    output.push_str(if covered {
        "<table:covered-table-cell"
    } else {
        "<table:table-cell"
    });
    if bind_namespaces {
        output.push_str(" xmlns:table=\"");
        output.push_str(TABLE_NAMESPACE);
        output.push_str("\" xmlns:office=\"");
        output.push_str(OFFICE_NAMESPACE);
        output.push_str("\" xmlns:text=\"");
        output.push_str(TEXT_NAMESPACE);
        output.push_str("\" xmlns:xlink=\"");
        output.push_str(XLINK_NAMESPACE);
        output.push('"');
    }
    if cell.repeat() > 1 {
        output.push_str(" table:number-columns-repeated=\"");
        output.push_str(&cell.repeat().to_string());
        output.push('"');
    }
    if let Merge::Span { rows, columns } = cell.merge {
        output.push_str(" table:number-rows-spanned=\"");
        output.push_str(&rows.get().to_string());
        output.push_str("\" table:number-columns-spanned=\"");
        output.push_str(&columns.get().to_string());
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
    if body.is_none()
        && cell.text.is_empty()
        && cell.hyperlinks.is_empty()
        && matches!(cell.value, CellValue::Empty)
        && cell.formula.is_none()
    {
        output.push_str("/>");
        return Ok(());
    }
    output.push('>');
    if let Some(body) = body {
        output.push_str(body);
    } else if !cell.text.is_empty()
        || !cell.hyperlinks.is_empty()
        || matches!(cell.value, CellValue::Text(_))
    {
        output.push_str(if requires_xml_space_preserve(&cell.text) {
            "<text:p xml:space=\"preserve\">"
        } else {
            "<text:p>"
        });
        write_cell_text(output, &cell.text, &cell.hyperlinks);
        output.push_str("</text:p>");
    }
    output.push_str(if covered {
        "</table:covered-table-cell>"
    } else {
        "</table:table-cell>"
    });
    Ok(())
}

fn write_cell_text(output: &mut String, text: &str, hyperlinks: &[Link]) {
    let mut cursor = 0usize;
    for hyperlink in hyperlinks {
        let range = hyperlink.range();
        output.push_str(&escape_xml(&text[cursor..range.start]));
        hyperlink.write_xml(output);
        cursor = range.end;
    }
    output.push_str(&escape_xml(&text[cursor..]));
}

fn requires_xml_space_preserve(text: &str) -> bool {
    text.bytes()
        .any(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
}

fn write_cell_text_bounded(
    output: &mut String,
    text: &str,
    hyperlinks: &[Link],
    max_bytes: usize,
) -> Result<()> {
    let mut cursor = 0usize;
    for hyperlink in hyperlinks {
        let range = hyperlink.range();
        bounded_push(output, &escape_xml(&text[cursor..range.start]), max_bytes)?;
        write_link_bounded(output, hyperlink, max_bytes)?;
        cursor = range.end;
    }
    bounded_push(output, &escape_xml(&text[cursor..]), max_bytes)
}

fn write_link_bounded(output: &mut String, link: &Link, max_bytes: usize) -> Result<()> {
    bounded_push(
        output,
        "<text:a xlink:type=\"simple\" xlink:href=\"",
        max_bytes,
    )?;
    bounded_push(output, &escape_xml(&link.href), max_bytes)?;
    bounded_push(output, "\"", max_bytes)?;
    if let Some(actuate) = link.actuate {
        bounded_push(output, " xlink:actuate=\"", max_bytes)?;
        bounded_push(output, actuate.as_str(), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(target_frame_name) = &link.target_frame_name {
        bounded_push(output, " office:target-frame-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(target_frame_name), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(show) = link.show {
        bounded_push(output, " xlink:show=\"", max_bytes)?;
        bounded_push(output, show.as_str(), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(name) = &link.name {
        bounded_push(output, " office:name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(name), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(title) = &link.title {
        bounded_push(output, " office:title=\"", max_bytes)?;
        bounded_push(output, &escape_xml(title), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(style_name) = &link.style_name {
        bounded_push(output, " text:style-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(style_name), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    if let Some(visited_style_name) = &link.visited_style_name {
        bounded_push(output, " text:visited-style-name=\"", max_bytes)?;
        bounded_push(output, &escape_xml(visited_style_name), max_bytes)?;
        bounded_push(output, "\"", max_bytes)?;
    }
    bounded_push(output, ">", max_bytes)?;
    bounded_push(output, &escape_xml(&link.text), max_bytes)?;
    bounded_push(output, "</text:a>", max_bytes)
}

fn write_value_attributes(output: &mut String, value: &CellValue) {
    match value {
        CellValue::Empty => {},
        CellValue::Text(_) => output.push_str(" office:value-type=\"string\""),
        CellValue::Number(value) => {
            output.push_str(" office:value-type=\"float\" office:value=\"");
            output.push_str(&value.to_string());
            output.push('"');
        },
        CellValue::Currency { value, currency } => {
            output.push_str(" office:value-type=\"currency\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\" office:currency=\"");
            output.push_str(&escape_xml(currency));
            output.push('"');
        },
        CellValue::Percentage(value) => {
            output.push_str(" office:value-type=\"percentage\" office:value=\"");
            output.push_str(&value.to_string());
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
