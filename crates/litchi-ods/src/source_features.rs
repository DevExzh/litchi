//! Bounded, inert inspection of source-level worksheet extensions.
//!
//! This module deliberately reports only metadata that can be read without
//! evaluating a formula, resolving a style, rendering a drawing, or following
//! an external link. It is separate from the compact worksheet graph because
//! those extension owners cannot be edited through that graph yet.

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const CALCEXT_NAMESPACE: &[u8] =
    b"urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0";
const MAX_DEPTH: usize = 256;
const MAX_SHEETS: usize = 16_384;
const MAX_ITEMS_PER_SHEET: usize = 16_384;
const MAX_TEXT_BYTES: usize = 64 * 1024;

/// Read-only source inventory for worksheet extensions that are intentionally
/// inert in Litchi.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Snapshot {
    sheets: Vec<Sheet>,
}

impl Snapshot {
    /// Inspect a `content.xml` or flat-ODS XML document.
    ///
    /// The parser is bounded and never contacts the URI found in a hyperlink
    /// or drawing. It is not an editing API and does not imply extension
    /// rendering, formula evaluation, or a style calculation.
    pub fn parse(xml: &str) -> Result<Self> {
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack = Vec::<Element>::new();
        let mut sheets = Vec::new();

        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid ODS feature XML: {error}"))
                    })?;
            let namespace = namespace_kind(&namespace);
            let event = event.into_owned();
            match event {
                Event::Start(start) => {
                    if stack.len() >= MAX_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "ODS feature XML nesting exceeds {MAX_DEPTH} elements"
                        )));
                    }
                    stack.push(classify_start(
                        namespace,
                        &start,
                        &reader,
                        &stack,
                        &mut sheets,
                    )?);
                },
                Event::Empty(start) => {
                    if stack.len() >= MAX_DEPTH {
                        return Err(Error::InvalidFormat(format!(
                            "ODS feature XML nesting exceeds {MAX_DEPTH} elements"
                        )));
                    }
                    if let Element::Hyperlink { sheet, text } =
                        classify_start(namespace, &start, &reader, &stack, &mut sheets)?
                    {
                        push_hyperlink(&mut sheets, sheet, text)?;
                    }
                },
                Event::Text(text) => {
                    if let Some(Element::Hyperlink {
                        text: hyperlink, ..
                    }) = stack
                        .iter_mut()
                        .rev()
                        .find(|element| matches!(element, Element::Hyperlink { .. }))
                    {
                        append_text(
                            text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                                Error::InvalidFormat(format!("invalid ODS hyperlink text: {error}"))
                            })?,
                            hyperlink,
                        )?;
                    }
                },
                Event::CData(text) => {
                    if let Some(Element::Hyperlink {
                        text: hyperlink, ..
                    }) = stack
                        .iter_mut()
                        .rev()
                        .find(|element| matches!(element, Element::Hyperlink { .. }))
                    {
                        append_text(
                            text.decode().map_err(|error| {
                                Error::InvalidFormat(format!("invalid ODS hyperlink text: {error}"))
                            })?,
                            hyperlink,
                        )?;
                    }
                },
                Event::End(_) => match stack.pop() {
                    Some(Element::Hyperlink { sheet, text }) => {
                        push_hyperlink(&mut sheets, sheet, text)?;
                    },
                    Some(_) => {},
                    None => {
                        return Err(Error::InvalidFormat(
                            "ODS feature XML element stack underflow".to_string(),
                        ));
                    },
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
        if !stack.is_empty() {
            return Err(Error::InvalidFormat(
                "ODS feature XML ended with an unfinished element".to_string(),
            ));
        }
        Ok(Self { sheets })
    }

    /// Return source features in worksheet document order.
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Select a feature inventory by its exact ODF sheet name.
    pub fn sheet(&self, name: &str) -> Option<&Sheet> {
        self.sheets.iter().find(|sheet| sheet.name == name)
    }
}

/// Inert source features belonging to one worksheet.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Sheet {
    name: String,
    conditional_formats: usize,
    sparkline_groups: usize,
    hyperlinks: Vec<Hyperlink>,
    drawings: Vec<Drawing>,
}

impl Sheet {
    /// Exact ODF `table:name`.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Number of `calcext:conditional-format` elements.
    pub fn conditional_format_count(&self) -> usize {
        self.conditional_formats
    }

    /// Number of `calcext:sparkline-group` elements.
    pub fn sparkline_group_count(&self) -> usize {
        self.sparkline_groups
    }

    /// Inert hyperlinks in source order. Their targets are never fetched.
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }

    /// In-table drawing occurrences in source order. Their sources are never
    /// loaded by this inspection API.
    pub fn drawings(&self) -> &[Drawing] {
        &self.drawings
    }
}

/// One inert `text:a` occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Hyperlink {
    href: String,
    text: String,
}

impl Hyperlink {
    /// Target IRI as written by the producer. Litchi never dereferences it.
    pub fn href(&self) -> &str {
        &self.href
    }

    /// Decoded visible text contained in the anchor.
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// The source element family for an in-table drawing occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DrawingKind {
    /// A `draw:frame` container.
    Frame,
    /// A `draw:image` payload.
    Image,
    /// Another `draw:*` shape element.
    Shape,
}

/// One inert in-table drawing occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Drawing {
    kind: DrawingKind,
    name: Option<String>,
    href: Option<String>,
}

impl Drawing {
    /// The drawing element family.
    pub const fn kind(&self) -> DrawingKind {
        self.kind
    }

    /// Optional producer name (`draw:name`).
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Optional image source (`xlink:href`). It is never dereferenced.
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }
}

#[derive(Debug)]
enum Element {
    Spreadsheet,
    Sheet(usize),
    Hyperlink { sheet: usize, text: Hyperlink },
    Other,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Text,
    Draw,
    Xlink,
    Calcext,
    Other,
}

fn classify_start(
    namespace: NamespaceKind,
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    stack: &[Element],
    sheets: &mut Vec<Sheet>,
) -> Result<Element> {
    let local = start.local_name();
    let local = local.as_ref();
    if namespace_matches(namespace, OFFICE_NAMESPACE) && local == b"spreadsheet" {
        return Ok(Element::Spreadsheet);
    }
    if namespace_matches(namespace, TABLE_NAMESPACE)
        && local == b"table"
        && matches!(stack.last(), Some(Element::Spreadsheet))
    {
        if sheets.len() >= MAX_SHEETS {
            return Err(Error::InvalidFormat(format!(
                "ODS feature inspection exceeds the {MAX_SHEETS} sheet safety limit"
            )));
        }
        let name = attribute(start, reader, TABLE_NAMESPACE, b"name")?
            .unwrap_or_else(|| "Sheet1".to_string());
        sheets.push(Sheet {
            name,
            ..Sheet::default()
        });
        return Ok(Element::Sheet(sheets.len() - 1));
    }
    let Some(sheet) = stack.iter().rev().find_map(|element| match element {
        Element::Sheet(index) => Some(*index),
        _ => None,
    }) else {
        return Ok(Element::Other);
    };
    if namespace_matches(namespace, CALCEXT_NAMESPACE) && local == b"conditional-format" {
        let features = &mut sheets[sheet];
        features.conditional_formats =
            checked_increment(features.conditional_formats, "conditional formats")?;
    } else if namespace_matches(namespace, CALCEXT_NAMESPACE) && local == b"sparkline-group" {
        let features = &mut sheets[sheet];
        features.sparkline_groups =
            checked_increment(features.sparkline_groups, "sparkline groups")?;
    } else if namespace_matches(namespace, TEXT_NAMESPACE) && local == b"a" {
        let href = attribute(start, reader, XLINK_NAMESPACE, b"href")?.unwrap_or_default();
        return Ok(Element::Hyperlink {
            sheet,
            text: Hyperlink {
                href,
                text: String::new(),
            },
        });
    } else if namespace_matches(namespace, DRAW_NAMESPACE) {
        let kind = match local {
            b"frame" => DrawingKind::Frame,
            b"image" => DrawingKind::Image,
            _ => DrawingKind::Shape,
        };
        let drawing = Drawing {
            kind,
            name: attribute(start, reader, DRAW_NAMESPACE, b"name")?,
            href: attribute(start, reader, XLINK_NAMESPACE, b"href")?,
        };
        let features = &mut sheets[sheet];
        if features.drawings.len() >= MAX_ITEMS_PER_SHEET {
            return Err(Error::InvalidFormat(format!(
                "ODS feature inspection exceeds the {MAX_ITEMS_PER_SHEET} drawing safety limit"
            )));
        }
        features.drawings.push(drawing);
    }
    Ok(Element::Other)
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NAMESPACE => {
            NamespaceKind::Office
        },
        ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(value)) if *value == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NAMESPACE => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(value)) if *value == XLINK_NAMESPACE => NamespaceKind::Xlink,
        ResolveResult::Bound(Namespace(value)) if *value == CALCEXT_NAMESPACE => {
            NamespaceKind::Calcext
        },
        _ => NamespaceKind::Other,
    }
}

fn namespace_matches(namespace: NamespaceKind, wanted: &[u8]) -> bool {
    match namespace {
        NamespaceKind::Office => wanted == OFFICE_NAMESPACE,
        NamespaceKind::Table => wanted == TABLE_NAMESPACE,
        NamespaceKind::Text => wanted == TEXT_NAMESPACE,
        NamespaceKind::Draw => wanted == DRAW_NAMESPACE,
        NamespaceKind::Xlink => wanted == XLINK_NAMESPACE,
        NamespaceKind::Calcext => wanted == CALCEXT_NAMESPACE,
        NamespaceKind::Other => false,
    }
}

fn attribute(
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    wanted_namespace: &[u8],
    wanted_local: &[u8],
) -> Result<Option<String>> {
    for raw in start.attributes().with_checks(true) {
        let raw = raw.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODS feature attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        let namespace = namespace_kind(&namespace);
        if namespace_matches(namespace, wanted_namespace) && local.as_ref() == wanted_local {
            let value = raw
                .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODS feature attribute value: {error}"))
                })?
                .into_owned();
            if value.len() > MAX_TEXT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "ODS feature attribute exceeds the {MAX_TEXT_BYTES}-byte safety limit"
                )));
            }
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn checked_increment(value: usize, label: &str) -> Result<usize> {
    value
        .checked_add(1)
        .filter(|value| *value <= MAX_ITEMS_PER_SHEET)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "ODS feature inspection exceeds the {MAX_ITEMS_PER_SHEET} {label} safety limit"
            ))
        })
}

fn append_text(value: std::borrow::Cow<'_, str>, target: &mut Hyperlink) -> Result<()> {
    let next =
        target.text.len().checked_add(value.len()).ok_or_else(|| {
            Error::InvalidFormat("ODS hyperlink text length overflow".to_string())
        })?;
    if next > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "ODS hyperlink text exceeds the {MAX_TEXT_BYTES}-byte safety limit"
        )));
    }
    target.text.push_str(&value);
    Ok(())
}

fn push_hyperlink(sheets: &mut [Sheet], sheet: usize, hyperlink: Hyperlink) -> Result<()> {
    let features = sheets.get_mut(sheet).ok_or_else(|| {
        Error::InvalidFormat("ODS feature inspection lost its sheet context".to_string())
    })?;
    if features.hyperlinks.len() >= MAX_ITEMS_PER_SHEET {
        return Err(Error::InvalidFormat(format!(
            "ODS feature inspection exceeds the {MAX_ITEMS_PER_SHEET} hyperlink safety limit"
        )));
    }
    features.hyperlinks.push(hyperlink);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{DrawingKind, Snapshot};

    #[test]
    fn inventories_inert_source_features_without_contacting_them() {
        let xml = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet 1"><table:table-row><table:table-cell><text:p><text:a xlink:href="https://example.test/never-contact">Link</text:a></text:p></table:table-cell></table:table-row><calcext:conditional-format/><calcext:sparkline-group/><table:shapes><draw:frame draw:name="frame"><draw:image xlink:href="http://192.0.2.1/pixel.png"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document>"#;
        let snapshot = Snapshot::parse(xml).unwrap();
        let sheet = snapshot.sheet("Sheet 1").unwrap();
        assert_eq!(sheet.conditional_format_count(), 1);
        assert_eq!(sheet.sparkline_group_count(), 1);
        assert_eq!(
            sheet.hyperlinks()[0].href(),
            "https://example.test/never-contact"
        );
        assert_eq!(sheet.hyperlinks()[0].text(), "Link");
        assert_eq!(sheet.drawings().len(), 2);
        assert_eq!(sheet.drawings()[0].kind(), DrawingKind::Frame);
        assert_eq!(
            sheet.drawings()[1].href(),
            Some("http://192.0.2.1/pixel.png")
        );
    }
}
