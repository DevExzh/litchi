//! Typed, ordered content extracted from ODF master-page headers and footers.

use std::collections::HashMap;

use litchi_core::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesRef, BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};

use super::header_footer::HeaderFooterKind;

const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const NUMBER_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0";

const MAX_XML_DEPTH: usize = 128;
const MAX_BLOCKS: usize = 4_096;
const MAX_INLINE_TOKENS: usize = 65_536;
const MAX_FIELDS: usize = 16_384;
const MAX_FIELD_ATTRIBUTES: usize = 128;
pub(super) const MAX_EXPANDED_SPACES: usize = 1_000_000;

/// One paragraph or heading in a header/footer region.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HeaderFooterBlock {
    /// The paragraph's `text:style-name`, when present.
    pub style_name: Option<String>,
    /// Inline content in document order.
    pub content: Vec<HeaderFooterInline>,
}

/// Ordered inline header/footer content.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HeaderFooterInline {
    Text(String),
    Space { count: usize },
    Tab,
    LineBreak,
    Field(HeaderFooterField),
}

/// A dynamic ODF text field without evaluating its value.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct HeaderFooterField {
    pub kind: HeaderFooterFieldKind,
    /// Cached/displayed child text stored in the document.
    pub displayed_text: String,
    pub fixed: Option<bool>,
    pub data_style_name: Option<String>,
    /// Attributes in source order with canonical namespace names.
    pub attributes: Vec<(String, String)>,
}

/// Supported field semantics. Fields are read as cached document metadata only.
///
/// The parser never evaluates formulas, resolves DDE connections, loads external
/// content, or invokes macros. Unknown text-namespace fields remain typed and
/// lossless.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum HeaderFooterFieldKind {
    PageNumber,
    PageCount,
    Title,
    Subject,
    AuthorName,
    Date,
    Time,
    FileName,
    Chapter,
    ModificationDate,
    UserDefined,
    /// A sender identity or contact field defined by ODF's `text:sender-*` elements.
    Sender(HeaderFooterSenderFieldKind),
    /// Inert `text:script` metadata. Its URI and payload are never opened or executed.
    Script,
    /// Inert `text:execute-macro` metadata. Its named macro is never invoked.
    ExecuteMacro,
    Unknown {
        namespace: String,
        local_name: String,
    },
}

/// One of the ODF sender identity/contact fields available in header/footer content.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum HeaderFooterSenderFieldKind {
    FirstName,
    LastName,
    Initials,
    Title,
    Position,
    Email,
    PrivatePhone,
    Fax,
    Company,
    WorkPhone,
    Street,
    City,
    PostalCode,
    Country,
    StateOrProvince,
}

impl HeaderFooterSenderFieldKind {
    /// The local name of the corresponding ODF `text:sender-*` element.
    pub const fn element_name(self) -> &'static str {
        match self {
            Self::FirstName => "sender-firstname",
            Self::LastName => "sender-lastname",
            Self::Initials => "sender-initials",
            Self::Title => "sender-title",
            Self::Position => "sender-position",
            Self::Email => "sender-email",
            Self::PrivatePhone => "sender-phone-private",
            Self::Fax => "sender-fax",
            Self::Company => "sender-company",
            Self::WorkPhone => "sender-phone-work",
            Self::Street => "sender-street",
            Self::City => "sender-city",
            Self::PostalCode => "sender-postal-code",
            Self::Country => "sender-country",
            Self::StateOrProvince => "sender-state-or-province",
        }
    }
}

struct Master {
    name: String,
    depth: usize,
}

struct Region {
    master_name: String,
    kind: HeaderFooterKind,
    depth: usize,
    blocks: Vec<HeaderFooterBlock>,
    block: Option<ActiveBlock>,
    field: Option<ActiveField>,
    token_count: usize,
    field_count: usize,
    expanded_spaces: usize,
}

struct ActiveBlock {
    depth: usize,
    block: HeaderFooterBlock,
}

struct ActiveField {
    depth: usize,
    field: HeaderFooterField,
}

pub(super) fn parse_header_footer_blocks(
    xml: &str,
) -> Result<HashMap<(String, HeaderFooterKind), Vec<HeaderFooterBlock>>> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut master: Option<Master> = None;
    let mut region: Option<Region> = None;
    let mut regions = HashMap::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("styles.xml parsing error: {error}")))?;
        let namespace = resolved_namespace(&namespace).map(<[u8]>::to_vec);
        let namespace = namespace.as_deref();
        let event = event.into_owned();
        match event {
            Event::Start(element) => {
                let element_depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if element_depth > MAX_XML_DEPTH {
                    return Err(depth_error());
                }
                if namespace == Some(STYLE_NAMESPACE)
                    && element.local_name().as_ref() == b"master-page"
                {
                    if master.is_some() {
                        return Err(Error::InvalidFormat(
                            "nested style:master-page element".to_string(),
                        ));
                    }
                    let name = namespaced_attr(&reader, &element, STYLE_NAMESPACE, b"name")?
                        .ok_or_else(|| {
                            Error::InvalidFormat(
                                "style:master-page is missing style:name".to_string(),
                            )
                        })?;
                    master = Some(Master {
                        name,
                        depth: element_depth,
                    });
                } else if let Some(master_page) = master.as_ref()
                    && region.is_none()
                    && namespace == Some(STYLE_NAMESPACE)
                    && let Some(kind) = HeaderFooterKind::parse(element.local_name().as_ref())
                {
                    region = Some(Region::new(master_page.name.clone(), kind, element_depth));
                } else if let Some(active) = region.as_mut() {
                    active.start_element(&reader, namespace, &element, element_depth, false)?;
                }
                depth = element_depth;
            },
            Event::Empty(element) => {
                let element_depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if element_depth > MAX_XML_DEPTH {
                    return Err(depth_error());
                }
                if let Some(master_page) = master.as_ref()
                    && region.is_none()
                    && namespace == Some(STYLE_NAMESPACE)
                    && let Some(kind) = HeaderFooterKind::parse(element.local_name().as_ref())
                {
                    insert_region(&mut regions, master_page.name.clone(), kind, Vec::new())?;
                } else if let Some(active) = region.as_mut() {
                    active.start_element(&reader, namespace, &element, element_depth, true)?;
                }
            },
            Event::Text(value) => {
                if let Some(active) = region.as_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid header text: {error}"))
                        })?;
                    active.push_text(&decoded)?;
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(active) = region.as_mut() {
                    active.push_text(&decode_reference(&reference)?)?;
                }
            },
            Event::CData(value) => {
                if let Some(active) = region.as_mut() {
                    let decoded = value
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| {
                            Error::InvalidFormat(format!("invalid header CDATA: {error}"))
                        })?;
                    active.push_text(&decoded)?;
                }
            },
            Event::End(element) => {
                if let Some(active) = region.as_mut() {
                    active.end_element(namespace, element.local_name().as_ref(), depth)?;
                }
                if region.as_ref().is_some_and(|active| active.depth == depth) {
                    let active = region.take().expect("checked header/footer region");
                    if namespace != Some(STYLE_NAMESPACE)
                        || HeaderFooterKind::parse(element.local_name().as_ref())
                            != Some(active.kind)
                    {
                        return Err(Error::InvalidFormat(
                            "malformed header/footer region nesting".to_string(),
                        ));
                    }
                    insert_region(&mut regions, active.master_name, active.kind, active.blocks)?;
                }
                if master.as_ref().is_some_and(|active| active.depth == depth) {
                    master = None;
                }
                depth = depth.checked_sub(1).ok_or_else(depth_error)?;
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(
                    "DTD is not allowed in ODF styles.xml".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if master.is_some() || region.is_some() || depth != 0 {
        return Err(Error::InvalidFormat(
            "unterminated master-page header/footer".to_string(),
        ));
    }
    Ok(regions)
}

impl Region {
    fn new(master_name: String, kind: HeaderFooterKind, depth: usize) -> Self {
        Self {
            master_name,
            kind,
            depth,
            blocks: Vec::new(),
            block: None,
            field: None,
            token_count: 0,
            field_count: 0,
            expanded_spaces: 0,
        }
    }

    fn start_element(
        &mut self,
        reader: &NsReader<&[u8]>,
        namespace: Option<&[u8]>,
        element: &BytesStart<'_>,
        depth: usize,
        empty: bool,
    ) -> Result<()> {
        let local = element.local_name();
        if namespace == Some(TEXT_NAMESPACE) && matches!(local.as_ref(), b"p" | b"h") {
            if self.block.is_some() {
                return Err(Error::InvalidFormat(
                    "nested header/footer paragraph or heading".to_string(),
                ));
            }
            if self.blocks.len() >= MAX_BLOCKS {
                return Err(limit_error("block"));
            }
            self.block = Some(ActiveBlock {
                depth,
                block: HeaderFooterBlock {
                    style_name: namespaced_attr(reader, element, TEXT_NAMESPACE, b"style-name")?,
                    content: Vec::new(),
                },
            });
            if empty {
                self.finish_block()?;
            }
            return Ok(());
        }
        if self.block.is_none() {
            return Ok(());
        }
        if self.field.is_some() {
            if namespace == Some(TEXT_NAMESPACE) {
                self.append_field_control(reader, element)?;
            }
            return Ok(());
        }
        if namespace == Some(TEXT_NAMESPACE) {
            match local.as_ref() {
                b"s" => {
                    let count = text_space_count(reader, element)?.unwrap_or(1);
                    self.add_spaces(count)?;
                    self.push_token(HeaderFooterInline::Space { count })?;
                },
                b"tab" => self.push_token(HeaderFooterInline::Tab)?,
                b"line-break" => self.push_token(HeaderFooterInline::LineBreak)?,
                local if field_kind(local).is_some() || is_unknown_field(local) => {
                    self.field_count = self
                        .field_count
                        .checked_add(1)
                        .ok_or_else(|| limit_error("field"))?;
                    if self.field_count > MAX_FIELDS {
                        return Err(limit_error("field"));
                    }
                    let field = parse_field(reader, element, local)?;
                    if empty {
                        self.push_token(HeaderFooterInline::Field(field))?;
                    } else {
                        self.field = Some(ActiveField { depth, field });
                    }
                },
                _ => {},
            }
        }
        Ok(())
    }

    fn end_element(&mut self, namespace: Option<&[u8]>, local: &[u8], depth: usize) -> Result<()> {
        if self
            .field
            .as_ref()
            .is_some_and(|field| field.depth == depth)
        {
            let field = self.field.take().expect("checked header/footer field");
            self.push_token(HeaderFooterInline::Field(field.field))?;
        }
        if self
            .block
            .as_ref()
            .is_some_and(|block| block.depth == depth)
        {
            if namespace != Some(TEXT_NAMESPACE) || !matches!(local, b"p" | b"h") {
                return Err(Error::InvalidFormat(
                    "malformed header/footer block nesting".to_string(),
                ));
            }
            if self.field.is_some() {
                return Err(Error::InvalidFormat(
                    "unterminated header/footer field".to_string(),
                ));
            }
            self.finish_block()?;
        }
        Ok(())
    }

    fn finish_block(&mut self) -> Result<()> {
        let block = self
            .block
            .take()
            .ok_or_else(|| Error::InvalidFormat("missing header/footer block".to_string()))?;
        self.blocks.push(block.block);
        Ok(())
    }

    fn push_text(&mut self, text: &str) -> Result<()> {
        if text.is_empty() || self.block.is_none() {
            return Ok(());
        }
        if let Some(field) = self.field.as_mut() {
            field.field.displayed_text.push_str(text);
            return Ok(());
        }
        let content = &mut self.block.as_mut().expect("checked block").block.content;
        if let Some(HeaderFooterInline::Text(existing)) = content.last_mut() {
            existing.push_str(text);
            return Ok(());
        }
        self.push_token(HeaderFooterInline::Text(text.to_string()))
    }

    fn append_field_control(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<()> {
        match element.local_name().as_ref() {
            b"s" => {
                let count = text_space_count(reader, element)?.unwrap_or(1);
                self.add_spaces(count)?;
                self.field
                    .as_mut()
                    .expect("checked field")
                    .field
                    .displayed_text
                    .extend(std::iter::repeat_n(' ', count));
            },
            b"tab" => self
                .field
                .as_mut()
                .expect("checked field")
                .field
                .displayed_text
                .push('\t'),
            b"line-break" => self
                .field
                .as_mut()
                .expect("checked field")
                .field
                .displayed_text
                .push('\n'),
            _ => {},
        }
        Ok(())
    }

    fn add_spaces(&mut self, count: usize) -> Result<()> {
        self.expanded_spaces = self
            .expanded_spaces
            .checked_add(count)
            .ok_or_else(|| limit_error("expanded-space"))?;
        if self.expanded_spaces > MAX_EXPANDED_SPACES {
            return Err(limit_error("expanded-space"));
        }
        Ok(())
    }

    fn push_token(&mut self, token: HeaderFooterInline) -> Result<()> {
        self.token_count = self
            .token_count
            .checked_add(1)
            .ok_or_else(|| limit_error("inline-token"))?;
        if self.token_count > MAX_INLINE_TOKENS {
            return Err(limit_error("inline-token"));
        }
        self.block
            .as_mut()
            .expect("token requires active block")
            .block
            .content
            .push(token);
        Ok(())
    }
}

fn parse_field(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
) -> Result<HeaderFooterField> {
    let mut attributes = Vec::new();
    let mut fixed = None;
    let mut data_style_name = None;
    for attribute in element.attributes() {
        if attributes.len() >= MAX_FIELD_ATTRIBUTES {
            return Err(limit_error("field-attribute"));
        }
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid header/footer field attribute: {error}"))
        })?;
        let (namespace, attr_local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = resolved_namespace(&namespace);
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| {
                Error::InvalidFormat(format!("invalid header/footer field attribute: {error}"))
            })?
            .into_owned();
        let key = canonical_attribute_name(namespace, attr_local.as_ref())?;
        if namespace == Some(TEXT_NAMESPACE) && attr_local.as_ref() == b"fixed" {
            fixed = Some(parse_bool(&value)?);
        }
        if namespace == Some(STYLE_NAMESPACE) && attr_local.as_ref() == b"data-style-name" {
            data_style_name = Some(value.clone());
        }
        attributes.push((key, value));
    }
    Ok(HeaderFooterField {
        kind: field_kind(local).unwrap_or_else(|| HeaderFooterFieldKind::Unknown {
            namespace: String::from_utf8_lossy(TEXT_NAMESPACE).into_owned(),
            local_name: String::from_utf8_lossy(local).into_owned(),
        }),
        displayed_text: String::new(),
        fixed,
        data_style_name,
        attributes,
    })
}

fn field_kind(local: &[u8]) -> Option<HeaderFooterFieldKind> {
    Some(match local {
        b"page-number" => HeaderFooterFieldKind::PageNumber,
        b"page-count" => HeaderFooterFieldKind::PageCount,
        b"title" => HeaderFooterFieldKind::Title,
        b"subject" => HeaderFooterFieldKind::Subject,
        b"author-name" => HeaderFooterFieldKind::AuthorName,
        b"date" => HeaderFooterFieldKind::Date,
        b"time" => HeaderFooterFieldKind::Time,
        b"file-name" => HeaderFooterFieldKind::FileName,
        b"chapter" => HeaderFooterFieldKind::Chapter,
        b"modification-date" => HeaderFooterFieldKind::ModificationDate,
        b"user-defined" => HeaderFooterFieldKind::UserDefined,
        b"sender-firstname" => {
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::FirstName)
        },
        b"sender-lastname" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::LastName),
        b"sender-initials" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Initials),
        b"sender-title" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Title),
        b"sender-position" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Position),
        b"sender-email" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Email),
        b"sender-phone-private" => {
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::PrivatePhone)
        },
        b"sender-fax" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Fax),
        b"sender-company" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Company),
        b"sender-phone-work" => {
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::WorkPhone)
        },
        b"sender-street" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Street),
        b"sender-city" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::City),
        b"sender-postal-code" => {
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::PostalCode)
        },
        b"sender-country" => HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Country),
        b"sender-state-or-province" => {
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::StateOrProvince)
        },
        b"script" => HeaderFooterFieldKind::Script,
        b"execute-macro" => HeaderFooterFieldKind::ExecuteMacro,
        _ => return None,
    })
}

fn is_unknown_field(local: &[u8]) -> bool {
    !matches!(
        local,
        b"p" | b"h"
            | b"span"
            | b"a"
            | b"s"
            | b"tab"
            | b"line-break"
            | b"soft-page-break"
            | b"bookmark"
            | b"bookmark-start"
            | b"bookmark-end"
            | b"reference-mark"
            | b"reference-mark-start"
            | b"reference-mark-end"
            | b"alphabetical-index-mark"
            | b"alphabetical-index-mark-start"
            | b"alphabetical-index-mark-end"
            | b"toc-mark"
            | b"toc-mark-start"
            | b"toc-mark-end"
            | b"change"
            | b"change-start"
            | b"change-end"
            | b"note"
            | b"note-citation"
            | b"note-body"
            | b"ruby"
            | b"ruby-base"
            | b"ruby-text"
            | b"meta"
    )
}

fn insert_region(
    regions: &mut HashMap<(String, HeaderFooterKind), Vec<HeaderFooterBlock>>,
    master_name: String,
    kind: HeaderFooterKind,
    blocks: Vec<HeaderFooterBlock>,
) -> Result<()> {
    if regions
        .insert((master_name.clone(), kind), blocks)
        .is_some()
    {
        return Err(Error::InvalidFormat(format!(
            "duplicate {kind:?} in master page '{master_name}'"
        )));
    }
    Ok(())
}

fn namespaced_attr(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    local_name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid header/footer attribute: {error}"))
        })?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        if resolved_namespace(&namespace) == Some(expected_namespace)
            && local.as_ref() == local_name
        {
            return attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid header/footer attribute: {error}"))
                });
        }
    }
    Ok(None)
}

fn text_space_count(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Option<usize>> {
    let value = namespaced_attr(reader, element, TEXT_NAMESPACE, b"c")?;
    value
        .map(|value| {
            value.parse::<usize>().map_err(|_| {
                Error::InvalidFormat("invalid text:c count in header/footer".to_string())
            })
        })
        .transpose()
}

fn canonical_attribute_name(namespace: Option<&[u8]>, local: &[u8]) -> Result<String> {
    let local = std::str::from_utf8(local)
        .map_err(|_| Error::InvalidFormat("invalid field attribute name".to_string()))?;
    Ok(match namespace {
        Some(TEXT_NAMESPACE) => format!("text:{local}"),
        Some(STYLE_NAMESPACE) => format!("style:{local}"),
        Some(OFFICE_NAMESPACE) => format!("office:{local}"),
        Some(NUMBER_NAMESPACE) => format!("number:{local}"),
        Some(namespace) => format!("{{{}}}{local}", String::from_utf8_lossy(namespace)),
        None => local.to_string(),
    })
}

fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid text:fixed value '{value}'"
        ))),
    }
}

fn resolved_namespace<'a>(namespace: &'a ResolveResult<'a>) -> Option<&'a [u8]> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => Some(*value),
        _ => None,
    }
}

fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference.resolve_char_ref().map_err(|error| {
        Error::InvalidFormat(format!("invalid header character reference: {error}"))
    })? {
        return Ok(character.to_string());
    }
    let name = reference
        .decode()
        .map_err(|error| Error::InvalidFormat(format!("invalid header entity: {error}")))?;
    match name.as_ref() {
        "amp" => Ok("&".to_string()),
        "lt" => Ok("<".to_string()),
        "gt" => Ok(">".to_string()),
        "quot" => Ok("\"".to_string()),
        "apos" => Ok("'".to_string()),
        _ => Err(Error::InvalidFormat(format!(
            "unsupported header entity '&{name};'"
        ))),
    }
}

fn depth_error() -> Error {
    Error::InvalidFormat("header/footer XML depth exceeds safety limit".to_string())
}

fn limit_error(resource: &str) -> Error {
    Error::InvalidFormat(format!(
        "header/footer {resource} count exceeds safety limit"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_ordered_blocks_controls_and_fields_with_arbitrary_prefixes() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:header><t:p t:style-name="Header">Page <t:page-number t:select-page="current" t:page-adjust="2" t:fixed="false" s:data-style-name="N1">7</t:page-number><t:s t:c="2"/><t:tab/><t:line-break/><t:sender-company t:fixed="true">Example</t:sender-company></t:p><t:h>Heading</t:h></s:header></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = parse_header_footer_blocks(xml).unwrap();
        let blocks = &regions[&(String::from("A"), HeaderFooterKind::Header)];
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[0].style_name.as_deref(), Some("Header"));
        assert_eq!(
            blocks[0].content[0],
            HeaderFooterInline::Text("Page ".into())
        );
        let HeaderFooterInline::Field(page) = &blocks[0].content[1] else {
            panic!("expected page field");
        };
        assert_eq!(page.kind, HeaderFooterFieldKind::PageNumber);
        assert_eq!(page.displayed_text, "7");
        assert_eq!(page.fixed, Some(false));
        assert_eq!(page.data_style_name.as_deref(), Some("N1"));
        assert!(
            page.attributes
                .contains(&("text:page-adjust".into(), "2".into()))
        );
        assert_eq!(blocks[0].content[2], HeaderFooterInline::Space { count: 2 });
        assert_eq!(blocks[0].content[3], HeaderFooterInline::Tab);
        assert_eq!(blocks[0].content[4], HeaderFooterInline::LineBreak);
        let HeaderFooterInline::Field(sender) = &blocks[0].content[5] else {
            panic!("expected sender field");
        };
        assert_eq!(
            sender.kind,
            HeaderFooterFieldKind::Sender(HeaderFooterSenderFieldKind::Company)
        );
        assert_eq!(sender.displayed_text, "Example");
        assert_eq!(sender.fixed, Some(true));
        assert_eq!(
            blocks[1].content,
            vec![HeaderFooterInline::Text("Heading".into())]
        );
    }

    #[test]
    fn classifies_all_standard_sender_field_names() {
        let cases = [
            ("sender-firstname", HeaderFooterSenderFieldKind::FirstName),
            ("sender-lastname", HeaderFooterSenderFieldKind::LastName),
            ("sender-initials", HeaderFooterSenderFieldKind::Initials),
            ("sender-title", HeaderFooterSenderFieldKind::Title),
            ("sender-position", HeaderFooterSenderFieldKind::Position),
            ("sender-email", HeaderFooterSenderFieldKind::Email),
            (
                "sender-phone-private",
                HeaderFooterSenderFieldKind::PrivatePhone,
            ),
            ("sender-fax", HeaderFooterSenderFieldKind::Fax),
            ("sender-company", HeaderFooterSenderFieldKind::Company),
            ("sender-phone-work", HeaderFooterSenderFieldKind::WorkPhone),
            ("sender-street", HeaderFooterSenderFieldKind::Street),
            ("sender-city", HeaderFooterSenderFieldKind::City),
            (
                "sender-postal-code",
                HeaderFooterSenderFieldKind::PostalCode,
            ),
            ("sender-country", HeaderFooterSenderFieldKind::Country),
            (
                "sender-state-or-province",
                HeaderFooterSenderFieldKind::StateOrProvince,
            ),
        ];

        for (local_name, sender_kind) in cases {
            assert_eq!(
                field_kind(local_name.as_bytes()),
                Some(HeaderFooterFieldKind::Sender(sender_kind))
            );
            assert_eq!(sender_kind.element_name(), local_name);
        }
    }

    #[test]
    fn retains_inert_script_and_macro_metadata() {
        let xml = r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:master-styles><s:master-page s:name="A"><s:header><t:p><t:script x:type="simple" x:href="https://example.invalid/never-open">payload</t:script><t:execute-macro t:name="Standard.Module1.Main">button</t:execute-macro></t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#;
        let regions = parse_header_footer_blocks(xml).unwrap();
        let blocks = &regions[&(String::from("A"), HeaderFooterKind::Header)];
        let fields: Vec<_> = blocks[0]
            .content
            .iter()
            .filter_map(|inline| match inline {
                HeaderFooterInline::Field(field) => Some(field),
                _ => None,
            })
            .collect();

        assert_eq!(fields.len(), 2);
        assert_eq!(fields[0].kind, HeaderFooterFieldKind::Script);
        assert_eq!(fields[0].displayed_text, "payload");
        assert!(fields[0].attributes.contains(&(
            "{http://www.w3.org/1999/xlink}href".into(),
            "https://example.invalid/never-open".into(),
        )));
        assert_eq!(fields[1].kind, HeaderFooterFieldKind::ExecuteMacro);
        assert_eq!(fields[1].displayed_text, "button");
        assert!(
            fields[1]
                .attributes
                .contains(&("text:name".into(), "Standard.Module1.Main".into()))
        );
    }

    #[test]
    fn parses_libreoffice_title_field_regression_fixture() {
        let xml = include_str!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/odf/odt/title-field-invalidate.fodt"
        ));
        let regions = parse_header_footer_blocks(xml).unwrap();
        let blocks = &regions[&(String::from("Standard"), HeaderFooterKind::Footer)];
        let fields: Vec<_> = blocks[0]
            .content
            .iter()
            .filter_map(|inline| match inline {
                HeaderFooterInline::Field(field) => {
                    Some((&field.kind, field.displayed_text.as_str()))
                },
                _ => None,
            })
            .collect();
        assert_eq!(
            fields,
            vec![
                (&HeaderFooterFieldKind::Subject, "mysubject"),
                (&HeaderFooterFieldKind::Title, "mytitle"),
                (&HeaderFooterFieldKind::UserDefined, "1.1"),
                (&HeaderFooterFieldKind::ModificationDate, "May 18, 2021"),
            ]
        );
    }

    #[test]
    fn rejects_invalid_boolean_dtd_depth_and_cumulative_spaces() {
        let wrap = |body: &str| {
            format!(
                r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:master-styles><s:master-page s:name="A"><s:header><t:p>{body}</t:p></s:header></s:master-page></o:master-styles></o:document-styles>"#
            )
        };
        assert!(parse_header_footer_blocks(&wrap("<t:date t:fixed=\"yes\"/>")).is_err());
        assert!(parse_header_footer_blocks(&format!("<!DOCTYPE x>{}", wrap("x"))).is_err());
        assert!(
            parse_header_footer_blocks(&wrap("<t:s t:c=\"600000\"/><t:s t:c=\"600000\"/>"))
                .is_err()
        );
        let nested = format!("{}x{}", "<t:span>".repeat(129), "</t:span>".repeat(129));
        assert!(parse_header_footer_blocks(&wrap(&nested)).is_err());
    }
}
