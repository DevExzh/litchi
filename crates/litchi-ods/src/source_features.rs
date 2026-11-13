//! Bounded, inert inspection of source-level worksheet extensions.
//!
//! This module deliberately reports only metadata that can be read without
//! evaluating a formula, resolving a style, rendering a drawing, or following
//! an external link. It is separate from the compact worksheet graph because
//! those extension owners cannot be edited through that graph yet.

use litchi_core::{Error as CoreError, Result as CoreResult};
use litchi_odf_common::Error as OdfError;
use quick_xml::{
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
const MAX_INPUT_BYTES: usize = 16 * 1024 * 1024;
const MAX_EVENTS: usize = 1_000_000;
const MAX_DEPTH: usize = 256;
const MAX_SHEETS: usize = 16_384;
const MAX_ITEMS_PER_SHEET: usize = 16_384;
const MAX_TEXT_BYTES: usize = 64 * 1024;

type FamilyResult<T> = Result<T, OdfError>;

/// Resource budget for [`Snapshot`] source inspection.
///
/// Builder values are clamped to hard ceilings so a caller cannot disable the
/// parser's memory, depth, or event protections.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    events: usize,
    depth: usize,
    sheets: usize,
    items_per_sheet: usize,
    text_bytes: usize,
}

impl Limits {
    /// Cap inspected XML bytes.
    #[must_use]
    pub fn with_input_bytes(mut self, value: usize) -> Self {
        self.input_bytes = value.min(MAX_INPUT_BYTES);
        self
    }

    /// Cap XML events, including declarations and comments.
    #[must_use]
    pub fn with_events(mut self, value: usize) -> Self {
        self.events = value.min(MAX_EVENTS);
        self
    }

    /// Cap element nesting depth.
    #[must_use]
    pub fn with_depth(mut self, value: usize) -> Self {
        self.depth = value.min(MAX_DEPTH);
        self
    }

    /// Cap discovered worksheets.
    #[must_use]
    pub fn with_sheets(mut self, value: usize) -> Self {
        self.sheets = value.min(MAX_SHEETS);
        self
    }

    /// Cap each feature category per worksheet.
    #[must_use]
    pub fn with_items_per_sheet(mut self, value: usize) -> Self {
        self.items_per_sheet = value.min(MAX_ITEMS_PER_SHEET);
        self
    }

    /// Cap decoded attribute and hyperlink-text allocation sizes.
    #[must_use]
    pub fn with_text_bytes(mut self, value: usize) -> Self {
        self.text_bytes = value.min(MAX_TEXT_BYTES);
        self
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            input_bytes: MAX_INPUT_BYTES,
            events: MAX_EVENTS,
            depth: MAX_DEPTH,
            sheets: MAX_SHEETS,
            items_per_sheet: MAX_ITEMS_PER_SHEET,
            text_bytes: MAX_TEXT_BYTES,
        }
    }
}

/// Read-only source inventory for worksheet extensions that are intentionally
/// inert in Litchi.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Snapshot {
    sheets: Vec<Sheet>,
}

impl Snapshot {
    /// Inspect a canonical `office:document-content` XML document.
    ///
    /// The parser is bounded and never contacts the URI found in a hyperlink
    /// or drawing. It is not an editing API and does not imply extension
    /// rendering, formula evaluation, or a style calculation.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse(xml: &str) -> FamilyResult<Self> {
        Ok(Self::parse_with_core(xml, Limits::default())?)
    }

    /// Inspect a canonical content part under an explicit resource budget.
    ///
    /// DTDs and general entity references are rejected. Only the canonical
    /// `office:document-content` -> `office:body` -> `office:spreadsheet`
    /// envelope is accepted; namespace prefixes themselves are unrestricted.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn parse_with(xml: &str, limits: Limits) -> FamilyResult<Self> {
        Ok(Self::parse_with_core(xml, limits)?)
    }

    fn parse_with_core(xml: &str, limits: Limits) -> CoreResult<Self> {
        if xml.len() > limits.input_bytes {
            return Err(CoreError::InvalidFormat(format!(
                "ODS feature XML exceeds the {}-byte input limit",
                limits.input_bytes
            )));
        }
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let mut stack = Vec::<Element>::new();
        let mut sheets = Vec::new();
        let mut events = 0usize;
        let mut phase = RootPhase::Prolog;

        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        CoreError::InvalidFormat(format!("invalid ODS feature XML: {error}"))
                    })?;
            let namespace = namespace_kind(&namespace);
            events = events.checked_add(1).ok_or_else(|| {
                CoreError::InvalidFormat("ODS feature XML event counter overflow".to_string())
            })?;
            if events > limits.events {
                return Err(CoreError::InvalidFormat(format!(
                    "ODS feature XML exceeds the {}-event limit",
                    limits.events
                )));
            }
            let event = event.into_owned();
            match event {
                Event::Start(start) => {
                    if stack.len() >= limits.depth {
                        return Err(CoreError::InvalidFormat(format!(
                            "ODS feature XML nesting exceeds the {}-element limit",
                            limits.depth
                        )));
                    }
                    stack.push(classify_start(
                        namespace,
                        &start,
                        &reader,
                        &stack,
                        &mut sheets,
                        limits,
                        &mut phase,
                    )?);
                },
                Event::Empty(start) => {
                    if stack.len() >= limits.depth {
                        return Err(CoreError::InvalidFormat(format!(
                            "ODS feature XML nesting exceeds the {}-element limit",
                            limits.depth
                        )));
                    }
                    match classify_start(
                        namespace,
                        &start,
                        &reader,
                        &stack,
                        &mut sheets,
                        limits,
                        &mut phase,
                    )? {
                        Element::Hyperlink { sheet, text } => {
                            push_hyperlink(&mut sheets, sheet, text, limits)?;
                        },
                        Element::Root | Element::Body | Element::Workbook | Element::Sheet(_) => {
                            return Err(CoreError::InvalidFormat(
                                "ODS feature XML requires a non-empty canonical envelope"
                                    .to_string(),
                            ));
                        },
                        Element::Cell(_)
                        | Element::ConditionalFormats(_)
                        | Element::SparklineGroups(_)
                        | Element::Shapes(_)
                        | Element::Other => {},
                    }
                },
                Event::Text(text) => {
                    let value = text
                        .xml_content(quick_xml::XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            CoreError::InvalidFormat(format!("invalid ODS hyperlink text: {error}"))
                        })?;
                    if let Some(Element::Hyperlink {
                        text: hyperlink, ..
                    }) = stack
                        .iter_mut()
                        .rev()
                        .find(|element| matches!(element, Element::Hyperlink { .. }))
                    {
                        append_text(value, hyperlink, limits)?;
                    } else if stack.is_empty() && !value.trim().is_empty() {
                        return Err(CoreError::InvalidFormat(
                            "ODS feature XML has text outside its document root".to_string(),
                        ));
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
                                CoreError::InvalidFormat(format!(
                                    "invalid ODS hyperlink text: {error}"
                                ))
                            })?,
                            hyperlink,
                            limits,
                        )?;
                    } else {
                        return Err(CoreError::InvalidFormat(
                            "ODS feature XML has CDATA outside a hyperlink".to_string(),
                        ));
                    }
                },
                Event::End(_) => match stack.pop() {
                    Some(Element::Hyperlink { sheet, text }) => {
                        push_hyperlink(&mut sheets, sheet, text, limits)?;
                    },
                    Some(Element::Root) => phase = RootPhase::Closed,
                    Some(_) => {},
                    None => {
                        return Err(CoreError::InvalidFormat(
                            "ODS feature XML element stack underflow".to_string(),
                        ));
                    },
                },
                Event::DocType(_) => {
                    return Err(CoreError::InvalidFormat(
                        "ODS feature inspection rejects DTD declarations".to_string(),
                    ));
                },
                Event::GeneralRef(_) => {
                    return Err(CoreError::InvalidFormat(
                        "ODS feature inspection rejects general entity references".to_string(),
                    ));
                },
                Event::Eof => break,
                Event::Comment(_) | Event::Decl(_) | Event::PI(_) => {},
            }
            buffer.clear();
        }
        let complete = match phase {
            RootPhase::Prolog | RootPhase::Open => false,
            RootPhase::Closed => stack.is_empty(),
        };
        if !complete {
            return Err(CoreError::InvalidFormat(
                "ODS feature XML is missing a complete canonical document-content envelope"
                    .to_string(),
            ));
        }
        Ok(Self { sheets })
    }

    /// Return source features in worksheet document order.
    #[must_use]
    pub fn sheets(&self) -> &[Sheet] {
        &self.sheets
    }

    /// Select a feature inventory by its exact ODF sheet name.
    #[must_use]
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
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
    /// Number of `calcext:conditional-format` elements.
    #[must_use]
    pub fn conditional_format_count(&self) -> usize {
        self.conditional_formats
    }
    /// Number of `calcext:sparkline-group` elements.
    #[must_use]
    pub fn sparkline_group_count(&self) -> usize {
        self.sparkline_groups
    }
    /// Inert hyperlinks in source order. Their targets are never fetched.
    #[must_use]
    pub fn hyperlinks(&self) -> &[Hyperlink] {
        &self.hyperlinks
    }
    /// In-table drawing occurrences in source order. Their sources are never loaded.
    #[must_use]
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
    #[must_use]
    pub fn href(&self) -> &str {
        &self.href
    }
    /// Decoded visible text contained in the anchor.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// The source element family for an in-table drawing occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DrawingKind {
    Frame,
    Image,
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
    #[must_use]
    pub const fn kind(&self) -> DrawingKind {
        self.kind
    }
    /// Optional producer name (`draw:name`).
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }
    /// Optional image source (`xlink:href`). It is never dereferenced.
    #[must_use]
    pub fn href(&self) -> Option<&str> {
        self.href.as_deref()
    }
}

#[derive(Debug)]
enum Element {
    Root,
    Body,
    Workbook,
    Sheet(usize),
    Cell(usize),
    ConditionalFormats(usize),
    SparklineGroups(usize),
    Shapes(usize),
    Hyperlink { sheet: usize, text: Hyperlink },
    Other,
}

/// Where feature inspection stands relative to `office:document-content`.
///
/// ```text
/// Prolog --office:document-content start tag--> Open
/// Open   --root end tag--> Closed
/// Open | Closed --office:document-content--> error (multiple document roots)
/// Closed --any other element except office:body and office:spreadsheet--> error
/// ```
///
/// `office:body` and `office:spreadsheet` are placed by their parent element
/// alone, and text outside the root is judged by the element stack. A
/// declaration, a comment and a processing instruction are accepted in every
/// phase. The end of input is accepted only in `Closed`.
#[derive(Clone, Copy, PartialEq, Eq)]
enum RootPhase {
    Prolog,
    Open,
    Closed,
}

#[allow(
    clippy::too_many_arguments,
    reason = "parser state is intentionally explicit"
)]
fn classify_start(
    namespace: NamespaceKind,
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    stack: &[Element],
    sheets: &mut Vec<Sheet>,
    limits: Limits,
    phase: &mut RootPhase,
) -> CoreResult<Element> {
    let local = start.local_name();
    let local = local.as_ref();
    if namespace == NamespaceKind::Office && local == b"document-content" {
        if *phase != RootPhase::Prolog || !stack.is_empty() {
            return Err(CoreError::InvalidFormat(
                "ODS feature XML has multiple document roots".to_string(),
            ));
        }
        *phase = RootPhase::Open;
        return Ok(Element::Root);
    }
    if namespace == NamespaceKind::Office && local == b"body" {
        return matches!(stack.last(), Some(Element::Root))
            .then_some(Element::Body)
            .ok_or_else(|| {
                CoreError::InvalidFormat(
                    "office:body must be the direct child of office:document-content".to_string(),
                )
            });
    }
    if namespace == NamespaceKind::Office && local == b"spreadsheet" {
        return matches!(stack.last(), Some(Element::Body))
            .then_some(Element::Workbook)
            .ok_or_else(|| {
                CoreError::InvalidFormat(
                    "office:spreadsheet must be the direct child of office:body".to_string(),
                )
            });
    }
    if stack.is_empty() || *phase == RootPhase::Closed {
        return Err(CoreError::InvalidFormat(
            "ODS feature XML has content outside its canonical document root".to_string(),
        ));
    }
    if namespace == NamespaceKind::Table && local == b"table" {
        if !matches!(stack.last(), Some(Element::Workbook)) {
            return Err(CoreError::InvalidFormat(
                "table:table must be a direct child of office:spreadsheet".to_string(),
            ));
        }
        if sheets.len() >= limits.sheets {
            return Err(CoreError::InvalidFormat(format!(
                "ODS feature inspection exceeds the {}-sheet limit",
                limits.sheets
            )));
        }
        let name = attribute(start, reader, TABLE_NAMESPACE, b"name", limits)?
            .unwrap_or_else(|| "Sheet1".to_string());
        sheets.push(Sheet {
            name,
            ..Sheet::default()
        });
        return Ok(Element::Sheet(sheets.len() - 1));
    }
    let Some(sheet) = stack.iter().rev().find_map(|element| match element {
        Element::Sheet(index) => Some(*index),
        Element::Root
        | Element::Body
        | Element::Workbook
        | Element::Cell(_)
        | Element::ConditionalFormats(_)
        | Element::SparklineGroups(_)
        | Element::Shapes(_)
        | Element::Hyperlink { .. }
        | Element::Other => None,
    }) else {
        return Ok(Element::Other);
    };
    if namespace == NamespaceKind::Table
        && (local == b"table-cell" || local == b"covered-table-cell")
    {
        return Ok(Element::Cell(sheet));
    }
    if namespace == NamespaceKind::Table && local == b"shapes" {
        return matches!(stack.last(), Some(Element::Sheet(index)) if *index == sheet)
            .then_some(Element::Shapes(sheet))
            .ok_or_else(|| {
                CoreError::InvalidFormat(
                    "table:shapes must be a direct child of table:table".to_string(),
                )
            });
    }
    if namespace == NamespaceKind::Calcext && local == b"conditional-formats" {
        return matches!(stack.last(), Some(Element::Sheet(index)) if *index == sheet)
            .then_some(Element::ConditionalFormats(sheet))
            .ok_or_else(|| {
                CoreError::InvalidFormat(
                    "calcext:conditional-formats must be a direct child of table:table".to_string(),
                )
            });
    }
    if namespace == NamespaceKind::Calcext && local == b"conditional-format" {
        let index =
            matches!(stack.last(), Some(Element::ConditionalFormats(index)) if *index == sheet)
                .then_some(sheet)
                .ok_or_else(|| {
                    CoreError::InvalidFormat(
                        "calcext:conditional-format must be inside calcext:conditional-formats"
                            .to_string(),
                    )
                })?;
        let features = &mut sheets[index];
        features.conditional_formats =
            checked_increment(features.conditional_formats, "conditional formats", limits)?;
    } else if namespace == NamespaceKind::Calcext && local == b"sparkline-groups" {
        return matches!(stack.last(), Some(Element::Sheet(index)) if *index == sheet)
            .then_some(Element::SparklineGroups(sheet))
            .ok_or_else(|| {
                CoreError::InvalidFormat(
                    "calcext:sparkline-groups must be a direct child of table:table".to_string(),
                )
            });
    } else if namespace == NamespaceKind::Calcext && local == b"sparkline-group" {
        let index =
            matches!(stack.last(), Some(Element::SparklineGroups(index)) if *index == sheet)
                .then_some(sheet)
                .ok_or_else(|| {
                    CoreError::InvalidFormat(
                        "calcext:sparkline-group must be inside calcext:sparkline-groups"
                            .to_string(),
                    )
                })?;
        let features = &mut sheets[index];
        features.sparkline_groups =
            checked_increment(features.sparkline_groups, "sparkline groups", limits)?;
    } else if namespace == NamespaceKind::Text && local == b"a" {
        if !stack
            .iter()
            .rev()
            .any(|element| matches!(element, Element::Cell(index) if *index == sheet))
        {
            return Err(CoreError::InvalidFormat(
                "text:a must occur inside a table cell".to_string(),
            ));
        }
        let href = attribute(start, reader, XLINK_NAMESPACE, b"href", limits)?.unwrap_or_default();
        return Ok(Element::Hyperlink {
            sheet,
            text: Hyperlink {
                href,
                text: String::new(),
            },
        });
    } else if namespace == NamespaceKind::Draw {
        if !stack
            .iter()
            .rev()
            .any(|element| matches!(element, Element::Shapes(index) if *index == sheet))
        {
            return Err(CoreError::InvalidFormat(
                "draw:* source features must occur inside table:shapes".to_string(),
            ));
        }
        let kind = match local {
            b"frame" => DrawingKind::Frame,
            b"image" => DrawingKind::Image,
            _ => DrawingKind::Shape,
        };
        let drawing = Drawing {
            kind,
            name: attribute(start, reader, DRAW_NAMESPACE, b"name", limits)?,
            href: attribute(start, reader, XLINK_NAMESPACE, b"href", limits)?,
        };
        let features = &mut sheets[sheet];
        if features.drawings.len() >= limits.items_per_sheet {
            return Err(CoreError::InvalidFormat(format!(
                "ODS feature inspection exceeds the {}-drawing limit",
                limits.items_per_sheet
            )));
        }
        features.drawings.push(drawing);
    }
    Ok(Element::Other)
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Table,
    Text,
    Draw,
    Calcext,
    Other,
}

fn namespace_kind(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == OFFICE_NAMESPACE => {
            NamespaceKind::Office
        },
        ResolveResult::Bound(Namespace(value)) if *value == TABLE_NAMESPACE => NamespaceKind::Table,
        ResolveResult::Bound(Namespace(value)) if *value == TEXT_NAMESPACE => NamespaceKind::Text,
        ResolveResult::Bound(Namespace(value)) if *value == DRAW_NAMESPACE => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(value)) if *value == CALCEXT_NAMESPACE => {
            NamespaceKind::Calcext
        },
        ResolveResult::Unbound | ResolveResult::Bound(_) | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn attribute(
    start: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    wanted_namespace: &[u8],
    wanted_local: &[u8],
    limits: Limits,
) -> CoreResult<Option<String>> {
    for raw in start.attributes().with_checks(true) {
        let raw = raw.map_err(|error| {
            CoreError::InvalidFormat(format!("invalid ODS feature attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(raw.key);
        if matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == wanted_namespace)
            && local.as_ref() == wanted_local
        {
            let value = raw
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|error| {
                    CoreError::InvalidFormat(format!(
                        "invalid ODS feature attribute value: {error}"
                    ))
                })?
                .into_owned();
            if value.len() > limits.text_bytes {
                return Err(CoreError::InvalidFormat(format!(
                    "ODS feature attribute exceeds the {}-byte limit",
                    limits.text_bytes
                )));
            }
            return Ok(Some(value));
        }
    }
    Ok(None)
}

fn checked_increment(value: usize, label: &str, limits: Limits) -> CoreResult<usize> {
    value
        .checked_add(1)
        .filter(|value| *value <= limits.items_per_sheet)
        .ok_or_else(|| {
            CoreError::InvalidFormat(format!(
                "ODS feature inspection exceeds the {}-{label} limit",
                limits.items_per_sheet
            ))
        })
}

fn append_text(
    value: std::borrow::Cow<'_, str>,
    target: &mut Hyperlink,
    limits: Limits,
) -> CoreResult<()> {
    let next = target.text.len().checked_add(value.len()).ok_or_else(|| {
        CoreError::InvalidFormat("ODS hyperlink text length overflow".to_string())
    })?;
    if next > limits.text_bytes {
        return Err(CoreError::InvalidFormat(format!(
            "ODS hyperlink text exceeds the {}-byte limit",
            limits.text_bytes
        )));
    }
    target.text.push_str(&value);
    Ok(())
}

fn push_hyperlink(
    sheets: &mut [Sheet],
    sheet: usize,
    hyperlink: Hyperlink,
    limits: Limits,
) -> CoreResult<()> {
    let features = sheets.get_mut(sheet).ok_or_else(|| {
        CoreError::InvalidFormat("ODS feature inspection lost its sheet context".to_string())
    })?;
    if features.hyperlinks.len() >= limits.items_per_sheet {
        return Err(CoreError::InvalidFormat(format!(
            "ODS feature inspection exceeds the {}-hyperlink limit",
            limits.items_per_sheet
        )));
    }
    features.hyperlinks.push(hyperlink);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::{DrawingKind, Limits, Snapshot};

    const XML: &str = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"><office:body><office:spreadsheet><table:table table:name="Sheet 1"><table:table-row><table:table-cell><text:p><text:a xlink:href="https://example.test/never-contact">Link</text:a></text:p></table:table-cell></table:table-row><calcext:conditional-formats><calcext:conditional-format/></calcext:conditional-formats><calcext:sparkline-groups><calcext:sparkline-group/></calcext:sparkline-groups><table:shapes><draw:frame draw:name="frame"><draw:image xlink:href="http://192.0.2.1/pixel.png"/></draw:frame></table:shapes></table:table></office:spreadsheet></office:body></office:document-content>"#;

    #[test]
    fn inventories_inert_source_features_without_contacting_them() {
        let snapshot = Snapshot::parse(XML).expect("test fixture or operation should succeed");
        let sheet = snapshot
            .sheet("Sheet 1")
            .expect("test fixture or operation should succeed");
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

    #[test]
    fn rejects_dtd_wrong_envelopes_and_budget_excesses() {
        assert!(Snapshot::parse("<!DOCTYPE x><office:document-content/>").is_err());
        assert!(Snapshot::parse("<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"/>").is_err());
        assert!(
            Snapshot::parse(
                XML.replace("<text:a", "<calcext:conditional-format><text:a")
                    .as_str()
            )
            .is_err()
        );
        assert!(
            Snapshot::parse_with(XML, Limits::default().with_input_bytes(XML.len() - 1)).is_err()
        );
        assert!(Snapshot::parse_with(XML, Limits::default().with_events(1)).is_err());
    }

    /// Runs `run` over the XML envelope orderings built around one root element
    /// and compares each `label => outcome` line with the frozen expectation.
    fn assert_envelope_matrix(
        open: &str,
        close: &str,
        empty: &str,
        expected: &[&str],
        run: impl Fn(&str) -> String,
    ) {
        const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
        let doc = format!("{open}{close}");
        let rows = [
            ("empty input", String::new()),
            ("whitespace only", " \n".to_owned()),
            ("declaration only", DECL.to_owned()),
            ("empty-element root", empty.to_owned()),
            ("start and end root", doc.clone()),
            ("declaration then root", format!("{DECL}{doc}")),
            ("declaration, whitespace, root", format!("{DECL}\n{doc}")),
            ("root then whitespace", format!("{doc}\n ")),
            ("root then comment", format!("{doc}<!--c-->")),
            ("root then PI", format!("{doc}<?pi x?>")),
            ("root then text", format!("{doc}text")),
            ("root then second root", format!("{doc}{doc}")),
            (
                "root then foreign start",
                format!("{doc}<foreign></foreign>"),
            ),
            ("root then foreign empty", format!("{doc}<foreign/>")),
            (
                "empty-element root then second root",
                format!("{empty}{doc}"),
            ),
            ("empty-element root then text", format!("{empty}text")),
            ("declaration after root", format!("{doc}{DECL}")),
            ("two declarations", format!("{DECL}{DECL}{doc}")),
            ("comment before declaration", format!("<!--c-->{DECL}{doc}")),
            ("comment before root", format!("<!--c-->{doc}")),
            ("PI before root", format!("<?pi x?>{doc}")),
            ("text before root", format!("text{doc}")),
            ("whitespace before declaration", format!(" {DECL}{doc}")),
            ("unterminated root", open.to_owned()),
            ("stray end tag", format!("{doc}</stray>")),
            ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
            ("DOCTYPE after root", format!("{doc}<!DOCTYPE r>")),
            ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
            ("CDATA after root", format!("{doc}<![CDATA[x]]>")),
            ("entity reference before root", format!("&amp;{doc}")),
            ("entity reference after root", format!("{doc}&amp;")),
        ];
        let actual: Vec<String> = rows
            .iter()
            .map(|(label, xml)| format!("{label} => {}", run(xml)))
            .collect();
        let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
        assert!(
            actual
                .iter()
                .map(String::as_str)
                .eq(expected.iter().copied()),
            "actual envelope matrix:\n{}",
            listing.join("\n")
        );
    }

    const MATRIX_ROOT: &str = concat!(
        "office:document-content ",
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
    );
    const MATRIX_BODY: &str =
        "<office:body><office:spreadsheet></office:spreadsheet></office:body>";
    const MATRIX_CLOSE: &str = "</office:document-content>";

    #[test]
    fn source_features_envelope_matrix() {
        assert_envelope_matrix(
            &format!("<{MATRIX_ROOT}>{MATRIX_BODY}"),
            MATRIX_CLOSE,
            &format!("<{MATRIX_ROOT}/>"),
            &[
                "empty input => Core(InvalidFormat(\"ODS feature XML is missing a complete canonical document-content envelope\"))",
                "whitespace only => Core(InvalidFormat(\"ODS feature XML is missing a complete canonical document-content envelope\"))",
                "declaration only => Core(InvalidFormat(\"ODS feature XML is missing a complete canonical document-content envelope\"))",
                "empty-element root => Core(InvalidFormat(\"ODS feature XML requires a non-empty canonical envelope\"))",
                "start and end root => Ok(sheets=0)",
                "declaration then root => Ok(sheets=0)",
                "declaration, whitespace, root => Ok(sheets=0)",
                "root then whitespace => Ok(sheets=0)",
                "root then comment => Ok(sheets=0)",
                "root then PI => Ok(sheets=0)",
                "root then text => Core(InvalidFormat(\"ODS feature XML has text outside its document root\"))",
                "root then second root => Core(InvalidFormat(\"ODS feature XML has multiple document roots\"))",
                "root then foreign start => Core(InvalidFormat(\"ODS feature XML has content outside its canonical document root\"))",
                "root then foreign empty => Core(InvalidFormat(\"ODS feature XML has content outside its canonical document root\"))",
                "empty-element root then second root => Core(InvalidFormat(\"ODS feature XML requires a non-empty canonical envelope\"))",
                "empty-element root then text => Core(InvalidFormat(\"ODS feature XML requires a non-empty canonical envelope\"))",
                "declaration after root => Ok(sheets=0)",
                "two declarations => Ok(sheets=0)",
                "comment before declaration => Ok(sheets=0)",
                "comment before root => Ok(sheets=0)",
                "PI before root => Ok(sheets=0)",
                "text before root => Core(InvalidFormat(\"ODS feature XML has text outside its document root\"))",
                "whitespace before declaration => Ok(sheets=0)",
                "unterminated root => Core(InvalidFormat(\"ODS feature XML is missing a complete canonical document-content envelope\"))",
                "stray end tag => Core(InvalidFormat(\"invalid ODS feature XML: ill-formed document: close tag `</stray>` does not match any open tag\"))",
                "DOCTYPE => Core(InvalidFormat(\"ODS feature inspection rejects DTD declarations\"))",
                "DOCTYPE after root => Core(InvalidFormat(\"ODS feature inspection rejects DTD declarations\"))",
                "top-level CDATA => Core(InvalidFormat(\"ODS feature XML has CDATA outside a hyperlink\"))",
                "CDATA after root => Core(InvalidFormat(\"ODS feature XML has CDATA outside a hyperlink\"))",
                "entity reference before root => Core(InvalidFormat(\"ODS feature inspection rejects general entity references\"))",
                "entity reference after root => Core(InvalidFormat(\"ODS feature inspection rejects general entity references\"))",
            ],
            |xml| match Snapshot::parse(xml) {
                Ok(value) => format!("Ok(sheets={})", value.sheets().len()),
                Err(error) => format!("{error:?}"),
            },
        );
    }
}
