//! Typed ODF paragraph drop-cap style support.

use crate::{FlatOpenDocument, OpenDocumentPackage, ParagraphStyleTabStops};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::XmlVersion;
use quick_xml::{
    events::{BytesStart, Event},
    name::{QName, ResolveResult},
    reader::NsReader,
};
use std::collections::HashSet;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const STYLE_TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const MAX_XML: usize = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_STYLES: usize = 65_536;
const MAX_VALUE: usize = 4_096;
const MAX_TOTAL: usize = 16 * 1024 * 1024;
const MAX_INTEGER: u32 = 1_000_000;

fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
fn name_ok(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
        return Err(bad(format!("invalid {field}")));
    }
    Ok(())
}

/// ODF `style:length`: the first word or a positive number of characters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DropCapLength {
    Word,
    Characters(u32),
}
impl DropCapLength {
    fn parse(value: &str) -> Result<Self> {
        if value == "word" {
            return Ok(Self::Word);
        }
        let value = value
            .parse::<u32>()
            .map_err(|_| bad("style:length must be 'word' or a positive integer"))?;
        let result = Self::Characters(value);
        result.validate()?;
        Ok(result)
    }
    fn validate(self) -> Result<()> {
        match self {
            Self::Word => Ok(()),
            Self::Characters(value) if (1..=MAX_INTEGER).contains(&value) => Ok(()),
            _ => Err(bad(
                "style:length is outside the supported positive-integer range",
            )),
        }
    }
}

/// Valid ODF physical length lexical value for `style:distance`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DropCapDistance(String);
impl DropCapDistance {
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.len() > MAX_VALUE || !physical_length(&value) {
            return Err(bad("style:distance must be an ODF physical length"));
        }
        Ok(Self(value))
    }
    pub fn as_str(&self) -> &str {
        &self.0
    }
}
fn physical_length(value: &str) -> bool {
    let Some(number) = ["cm", "mm", "in", "pt", "pc", "px"]
        .iter()
        .find_map(|unit| value.strip_suffix(unit))
    else {
        return false;
    };
    let number = number.strip_prefix('-').unwrap_or(number);
    let mut split = number.split('.');
    let whole = split.next().unwrap_or_default();
    let fraction = split.next();
    if split.next().is_some() {
        return false;
    }
    let digits = |part: &str| part.bytes().all(|byte| byte.is_ascii_digit());
    match fraction {
        None => !whole.is_empty() && digits(whole),
        Some(fraction) => {
            digits(whole) && digits(fraction) && (!whole.is_empty() || !fraction.is_empty())
        },
    }
}

/// Complete empty `style:drop-cap` element.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphDropCap {
    pub length: Option<DropCapLength>,
    pub lines: Option<u32>,
    pub distance: Option<DropCapDistance>,
    pub style_name: Option<String>,
}
impl ParagraphDropCap {
    pub fn new() -> Self {
        Self::default()
    }
    pub fn validate(&self) -> Result<()> {
        if let Some(value) = self.length {
            value.validate()?;
        }
        if let Some(value) = self.lines
            && !(1..=MAX_INTEGER).contains(&value)
        {
            return Err(bad(
                "style:lines is outside the supported positive-integer range",
            ));
        }
        if let Some(value) = &self.distance
            && !physical_length(value.as_str())
        {
            return Err(bad("style:distance must be an ODF physical length"));
        }
        if let Some(value) = &self.style_name {
            name_ok(value, "style:style-name")?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(r#"<style:drop-cap xmlns:style="{STYLE_TEXT}""#);
        if let Some(value) = &self.distance {
            xml.push_str(&format!(
                r#" style:distance="{}""#,
                escape_xml(value.as_str())
            ));
        }
        if let Some(value) = self.length {
            let value = match value {
                DropCapLength::Word => "word".to_owned(),
                DropCapLength::Characters(value) => value.to_string(),
            };
            xml.push_str(&format!(r#" style:length="{value}""#));
        }
        if let Some(value) = self.lines {
            xml.push_str(&format!(r#" style:lines="{value}""#));
        }
        if let Some(value) = &self.style_name {
            xml.push_str(&format!(r#" style:style-name="{}""#, escape_xml(value)));
        }
        xml.push_str("/>");
        Ok(xml)
    }
}

/// A named or default paragraph style and its direct optional drop cap.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ParagraphStyleDropCap {
    pub name: Option<String>,
    pub parent_style_name: Option<String>,
    pub is_default_style: bool,
    pub drop_cap: Option<ParagraphDropCap>,
}
impl ParagraphStyleDropCap {
    pub fn named(name: impl Into<String>, drop_cap: Option<ParagraphDropCap>) -> Result<Self> {
        let result = Self {
            name: Some(name.into()),
            parent_style_name: None,
            is_default_style: false,
            drop_cap,
        };
        result.validate()?;
        Ok(result)
    }
    pub fn default_style(drop_cap: Option<ParagraphDropCap>) -> Self {
        Self {
            name: None,
            parent_style_name: None,
            is_default_style: true,
            drop_cap,
        }
    }
    pub fn validate(&self) -> Result<()> {
        match (&self.name, self.is_default_style) {
            (Some(name), false) => name_ok(name, "paragraph style name")?,
            (None, true) => {},
            _ => return Err(bad("paragraph style identity is inconsistent")),
        }
        if let Some(parent) = &self.parent_style_name {
            if self.is_default_style {
                return Err(bad("default paragraph style cannot have a parent"));
            }
            name_ok(parent, "parent paragraph style name")?;
        }
        if let Some(cap) = &self.drop_cap {
            cap.validate()?;
        }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let tag = if self.is_default_style {
            "default-style"
        } else {
            "style"
        };
        let mut xml =
            format!(r#"<style:{tag} xmlns:style="{STYLE_TEXT}" style:family="paragraph""#);
        if let Some(name) = &self.name {
            xml.push_str(&format!(r#" style:name="{}""#, escape_xml(name)));
        }
        if let Some(parent) = &self.parent_style_name {
            xml.push_str(&format!(
                r#" style:parent-style-name="{}""#,
                escape_xml(parent)
            ));
        }
        if let Some(cap) = &self.drop_cap {
            xml.push_str("><style:paragraph-properties>");
            xml.push_str(&cap.to_xml_fragment()?);
            xml.push_str(&format!("</style:paragraph-properties></style:{tag}>"));
        } else {
            xml.push_str("/>");
        }
        Ok(xml)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ParagraphStyleDropCapSet {
    pub styles: Vec<ParagraphStyleDropCap>,
}
impl ParagraphStyleDropCapSet {
    pub fn get(&self, name: &str) -> Option<&ParagraphStyleDropCap> {
        self.styles
            .iter()
            .find(|style| style.name.as_deref() == Some(name))
    }
    pub fn default_style(&self) -> Option<&ParagraphStyleDropCap> {
        self.styles.iter().find(|style| style.is_default_style)
    }
    pub fn resolved_drop_cap(&self, name: &str) -> Result<Option<&ParagraphDropCap>> {
        let mut current = self.get(name);
        let mut seen = HashSet::new();
        while let Some(style) = current {
            let identity = style.name.as_deref().unwrap_or("<default>");
            if !seen.insert(identity) {
                return Err(bad("paragraph style inheritance cycle"));
            }
            if let Some(cap) = &style.drop_cap {
                return Ok(Some(cap));
            }
            current = style
                .parent_style_name
                .as_deref()
                .and_then(|parent| self.get(parent));
            if style.parent_style_name.is_none() {
                break;
            }
        }
        Ok(self
            .default_style()
            .and_then(|style| style.drop_cap.as_ref()))
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum KnownNamespace {
    Office,
    Style,
    Other,
}
fn known(resolve: ResolveResult<'_>) -> KnownNamespace {
    match resolve {
        ResolveResult::Bound(value) if value.as_ref() == OFFICE => KnownNamespace::Office,
        ResolveResult::Bound(value) if value.as_ref() == STYLE => KnownNamespace::Style,
        _ => KnownNamespace::Other,
    }
}
fn element(reader: &NsReader<&[u8]>, name: QName<'_>) -> (KnownNamespace, Vec<u8>) {
    let (namespace, local) = reader.resolver().resolve_element(name);
    (known(namespace), local.as_ref().to_vec())
}
fn value(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    attr: &quick_xml::events::attributes::Attribute<'_>,
) -> Result<String> {
    attr.decoded_and_normalized_value(version, reader.decoder())
        .map(|value| value.into_owned())
        .map_err(|error| bad(format!("invalid attribute value: {error}")))
}
fn style_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<Option<ParagraphStyleDropCap>> {
    let mut name = None;
    let mut parent = None;
    let mut family = None;
    let mut seen = HashSet::new();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|error| bad(format!("invalid style attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        if known(namespace) == KnownNamespace::Style {
            if !seen.insert(local.as_ref().to_vec()) {
                return Err(bad("duplicate paragraph style attribute"));
            }
            match local.as_ref() {
                b"name" => name = Some(value(reader, version, &attr)?),
                b"parent-style-name" => parent = Some(value(reader, version, &attr)?),
                b"family" => family = Some(value(reader, version, &attr)?),
                _ => {},
            }
        } else if matches!(local.as_ref(), b"name" | b"parent-style-name" | b"family") {
            return Err(bad("paragraph style attribute uses wrong namespace"));
        }
    }
    if family.as_deref() != Some("paragraph") {
        return Ok(None);
    }
    let result = ParagraphStyleDropCap {
        name,
        parent_style_name: parent,
        is_default_style: start.local_name().as_ref() == b"default-style",
        drop_cap: None,
    };
    result.validate()?;
    Ok(Some(result))
}
fn cap_attrs(
    reader: &NsReader<&[u8]>,
    version: XmlVersion,
    start: &BytesStart<'_>,
) -> Result<ParagraphDropCap> {
    let mut cap = ParagraphDropCap::new();
    let mut seen = HashSet::new();
    for attr in start.attributes().with_checks(true) {
        let attr = attr.map_err(|error| bad(format!("invalid drop-cap attribute: {error}")))?;
        if attr.key.as_ref() == b"xmlns" || attr.key.as_ref().starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, local) = reader.resolver().resolve_attribute(attr.key);
        if known(namespace) != KnownNamespace::Style || !seen.insert(local.as_ref().to_vec()) {
            return Err(bad("invalid or duplicate style:drop-cap attribute"));
        }
        let value = value(reader, version, &attr)?;
        if value.len() > MAX_VALUE {
            return Err(bad("style:drop-cap attribute is too large"));
        }
        match local.as_ref() {
            b"length" => cap.length = Some(DropCapLength::parse(&value)?),
            b"lines" => {
                let lines = value
                    .parse::<u32>()
                    .map_err(|_| bad("style:lines must be a positive integer"))?;
                cap.lines = Some(lines);
            },
            b"distance" => cap.distance = Some(DropCapDistance::new(value)?),
            b"style-name" => {
                name_ok(&value, "style:style-name")?;
                cap.style_name = Some(value);
            },
            _ => return Err(bad("unknown style:drop-cap attribute")),
        }
    }
    cap.validate()?;
    Ok(cap)
}
struct Active {
    depth: usize,
    style: ParagraphStyleDropCap,
    properties: bool,
    cap: bool,
    open_cap: Option<usize>,
}
fn push_style(
    styles: &mut Vec<ParagraphStyleDropCap>,
    style: ParagraphStyleDropCap,
    total: &mut usize,
) -> Result<()> {
    if styles.len() >= MAX_STYLES {
        return Err(bad("too many paragraph styles"));
    }
    if styles
        .iter()
        .any(|old| old.name == style.name && old.is_default_style == style.is_default_style)
    {
        return Err(bad("duplicate paragraph style identity"));
    }
    *total += style.name.as_deref().map_or(0, str::len)
        + style.parent_style_name.as_deref().map_or(0, str::len)
        + style
            .drop_cap
            .as_ref()
            .and_then(|cap| cap.style_name.as_deref())
            .map_or(0, str::len);
    if *total > MAX_TOTAL {
        return Err(bad("paragraph drop-cap data is too large"));
    }
    styles.push(style);
    Ok(())
}

fn text_is_empty(text: &quick_xml::events::BytesText<'_>) -> bool {
    let bytes: &[u8] = text.as_ref();
    bytes.is_empty()
}

/// Parse ODF paragraph styles and the optional direct `style:drop-cap` child.
pub fn parse_paragraph_style_drop_caps(xml: &str) -> Result<ParagraphStyleDropCapSet> {
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    if !xml.contains("drop-cap") {
        return Ok(ParagraphStyleDropCapSet::default());
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    reader.config_mut().trim_text(false);
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(KnownNamespace, Vec<u8>)> = Vec::new();
    let mut active: Option<Active> = None;
    let mut styles = Vec::new();
    let mut total = 0;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(bad("styles XML nesting is too deep"));
                }
                if active
                    .as_ref()
                    .is_some_and(|state| state.open_cap.is_some())
                {
                    return Err(bad("style:drop-cap must be empty"));
                }
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|(n, l)| {
                    *n == KnownNamespace::Office
                        && matches!(l.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == KnownNamespace::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) = style_attrs(&reader, version, &start)? {
                        active = Some(Active {
                            depth,
                            style,
                            properties: false,
                            cap: false,
                            open_cap: None,
                        });
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    let props = depth == state.depth + 1
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"paragraph-properties";
                    if props {
                        if state.properties {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        state.properties = true;
                    } else if current.1 == b"paragraph-properties"
                        && current.0 != KnownNamespace::Style
                    {
                        return Err(bad("paragraph-properties uses wrong namespace"));
                    }
                    let cap = depth == state.depth + 2
                        && state.properties
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"drop-cap";
                    if cap {
                        if state.cap {
                            return Err(bad("duplicate style:drop-cap"));
                        }
                        state.cap = true;
                        state.style.drop_cap = Some(cap_attrs(&reader, version, &start)?);
                        state.open_cap = Some(depth);
                    } else if current.1 == b"drop-cap" {
                        return Err(bad("style:drop-cap has invalid namespace or parent"));
                    }
                } else if current.1 == b"drop-cap" {
                    return Err(bad("style:drop-cap has invalid parent"));
                }
            },
            Ok(Event::Empty(start)) => {
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|(n, l)| {
                    *n == KnownNamespace::Office
                        && matches!(l.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == KnownNamespace::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                if direct {
                    if let Some(style) = style_attrs(&reader, version, &start)? {
                        push_style(&mut styles, style, &mut total)?;
                    }
                    continue;
                }
                if let Some(state) = active.as_mut() {
                    let depth = stack.len() + 1;
                    let props = depth == state.depth + 1
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"paragraph-properties";
                    if props {
                        if state.properties {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                        state.properties = true;
                    } else if current.1 == b"paragraph-properties"
                        && current.0 != KnownNamespace::Style
                    {
                        return Err(bad("paragraph-properties uses wrong namespace"));
                    }
                    let cap = depth == state.depth + 2
                        && state.properties
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"drop-cap";
                    if cap {
                        if state.cap {
                            return Err(bad("duplicate style:drop-cap"));
                        }
                        state.cap = true;
                        state.style.drop_cap = Some(cap_attrs(&reader, version, &start)?);
                    } else if current.1 == b"drop-cap" {
                        return Err(bad("style:drop-cap has invalid namespace or parent"));
                    }
                } else if current.1 == b"drop-cap" {
                    return Err(bad("style:drop-cap has invalid parent"));
                }
            },
            Ok(Event::End(_)) => {
                let depth = stack.len();
                if let Some(state) = active.as_mut()
                    && state.open_cap == Some(depth)
                {
                    state.open_cap = None;
                }
                if active.as_ref().is_some_and(|state| state.depth == depth) {
                    push_style(&mut styles, active.take().unwrap().style, &mut total)?;
                }
                stack.pop();
            },
            Ok(Event::Text(text))
                if active
                    .as_ref()
                    .is_some_and(|state| state.open_cap.is_some())
                    && !text_is_empty(&text) =>
            {
                return Err(bad("style:drop-cap must be empty"));
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    if !stack.is_empty() || active.is_some() {
        return Err(bad("truncated styles XML"));
    }
    Ok(ParagraphStyleDropCapSet { styles })
}

pub(crate) fn same_style_identity(
    cap: &ParagraphStyleDropCap,
    tabs: &ParagraphStyleTabStops,
) -> bool {
    cap.is_default_style == tabs.is_default_style && cap.name == tabs.name
}
pub(crate) fn merge_with_tab_style(
    tabs: &ParagraphStyleTabStops,
    cap: &ParagraphStyleDropCap,
) -> Result<String> {
    tabs.validate()?;
    cap.validate()?;
    if !same_style_identity(cap, tabs) || cap.parent_style_name != tabs.parent_style_name {
        return Err(bad("paragraph style definitions conflict"));
    }
    let mut xml = tabs.to_xml_fragment()?;
    let Some(drop) = &cap.drop_cap else {
        return Ok(xml);
    };
    let drop = drop.to_xml_fragment()?;
    if let Some(at) = xml.find("<style:paragraph-properties>") {
        xml.insert_str(at + 28, &drop);
    } else if let Some(at) = xml.rfind("/>") {
        let tag = if tabs.is_default_style {
            "default-style"
        } else {
            "style"
        };
        xml.replace_range(
            at..,
            &format!(
                "><style:paragraph-properties>{drop}</style:paragraph-properties></style:{tag}>"
            ),
        );
    }
    Ok(xml)
}

#[derive(Default)]
struct Span {
    start: usize,
    end: usize,
    end_start: usize,
    qname: String,
    empty: bool,
}
#[derive(Default)]
struct Spans {
    style: Span,
    props: Option<Span>,
    cap: Option<Span>,
}
fn boundary(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| bad("invalid XML event boundary"))
}
fn replace(xml: &str, span: &Span, value: &str) -> String {
    format!("{}{}{}", &xml[..span.start], value, &xml[span.end..])
}
fn expand(xml: &str, span: &Span, value: &str) -> Result<String> {
    let raw = &xml[span.start..span.end];
    let slash = raw
        .rfind("/>")
        .ok_or_else(|| bad("invalid empty element"))?;
    Ok(replace(
        xml,
        span,
        &format!("{}>{value}</{}>", &raw[..slash], span.qname),
    ))
}

pub(crate) fn set_paragraph_style_drop_cap_xml(
    xml: &str,
    requested: &ParagraphStyleDropCap,
) -> Result<String> {
    requested.validate()?;
    if xml.len() > MAX_XML {
        return Err(bad("styles XML is too large"));
    }
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut version = XmlVersion::Implicit1_0;
    let mut stack: Vec<(KnownNamespace, Vec<u8>)> = Vec::new();
    let mut active_depth = None;
    let mut active: Option<Spans> = None;
    let mut found: Option<Spans> = None;
    loop {
        match reader.read_event() {
            Ok(Event::Start(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|(n, l)| {
                    *n == KnownNamespace::Office
                        && matches!(l.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == KnownNamespace::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                stack.push(current.clone());
                let depth = stack.len();
                if direct {
                    if let Some(style) = style_attrs(&reader, version, &start)?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
                        }
                        active_depth = Some(depth);
                        active = Some(Spans {
                            style: Span {
                                start: begin,
                                qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                                ..Span::default()
                            },
                            ..Spans::default()
                        });
                    }
                } else if let Some(sd) = active_depth {
                    let span = Span {
                        start: begin,
                        qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                        ..Span::default()
                    };
                    if depth == sd + 1
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"paragraph-properties"
                    {
                        if active.as_mut().unwrap().props.replace(span).is_some() {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                    } else if depth == sd + 2
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"drop-cap"
                        && active.as_mut().unwrap().cap.replace(span).is_some() {
                            return Err(bad("duplicate style:drop-cap"));
                        }
                }
            },
            Ok(Event::Empty(start)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let current = element(&reader, start.name());
                let parent = stack.last();
                let direct = parent.is_some_and(|(n, l)| {
                    *n == KnownNamespace::Office
                        && matches!(l.as_slice(), b"styles" | b"automatic-styles")
                }) && current.0 == KnownNamespace::Style
                    && matches!(current.1.as_slice(), b"style" | b"default-style");
                let span = Span {
                    start: begin,
                    end,
                    end_start: begin,
                    qname: String::from_utf8_lossy(start.name().as_ref()).into_owned(),
                    empty: true,
                };
                if direct {
                    if let Some(style) = style_attrs(&reader, version, &start)?
                        && style.name == requested.name
                        && style.is_default_style == requested.is_default_style
                    {
                        if active.is_some() || found.is_some() {
                            return Err(bad("duplicate target paragraph style"));
                        }
                        found = Some(Spans {
                            style: span,
                            ..Spans::default()
                        });
                    }
                } else if let Some(sd) = active_depth {
                    let depth = stack.len() + 1;
                    if depth == sd + 1
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"paragraph-properties"
                    {
                        if active.as_mut().unwrap().props.replace(span).is_some() {
                            return Err(bad("duplicate style:paragraph-properties"));
                        }
                    } else if depth == sd + 2
                        && current.0 == KnownNamespace::Style
                        && current.1 == b"drop-cap"
                        && active.as_mut().unwrap().cap.replace(span).is_some() {
                            return Err(bad("duplicate style:drop-cap"));
                        }
                }
            },
            Ok(Event::End(_)) => {
                let end = reader.buffer_position() as usize;
                let begin = boundary(xml, end)?;
                let depth = stack.len();
                if let Some(spans) = active.as_mut() {
                    if spans.cap.as_ref().is_some_and(|s| s.end == 0)
                        && depth == active_depth.unwrap() + 2
                    {
                        let s = spans.cap.as_mut().unwrap();
                        s.end_start = begin;
                        s.end = end;
                    }
                    if spans.props.as_ref().is_some_and(|s| s.end == 0)
                        && depth == active_depth.unwrap() + 1
                    {
                        let s = spans.props.as_mut().unwrap();
                        s.end_start = begin;
                        s.end = end;
                    }
                    if active_depth == Some(depth) {
                        spans.style.end_start = begin;
                        spans.style.end = end;
                        found = active.take();
                        active_depth = None;
                    }
                }
                stack.pop();
            },
            Ok(Event::Decl(decl)) => {
                version = decl
                    .xml_version()
                    .map_err(|error| bad(format!("unsupported XML version: {error}")))?
            },
            Ok(Event::DocType(_)) | Ok(Event::PI(_)) => {
                return Err(bad("DTD and processing instructions are not allowed"));
            },
            Ok(Event::Eof) => break,
            Ok(_) => {},
            Err(error) => return Err(bad(format!("invalid styles XML: {error}"))),
        }
    }
    let spans = found.ok_or_else(|| bad("target paragraph style does not exist"))?;
    let replacement = requested
        .drop_cap
        .as_ref()
        .map(ParagraphDropCap::to_xml_fragment)
        .transpose()?;
    if let Some(cap) = &spans.cap {
        return Ok(replace(xml, cap, replacement.as_deref().unwrap_or("")));
    }
    let Some(replacement) = replacement else {
        return Ok(xml.to_owned());
    };
    if let Some(props) = &spans.props {
        if props.empty {
            return expand(xml, props, &replacement);
        }
        let mut out = xml.to_owned();
        out.insert_str(props.end_start, &replacement);
        return Ok(out);
    }
    let props = format!(
        r#"<style:paragraph-properties xmlns:style="{STYLE_TEXT}">{replacement}</style:paragraph-properties>"#
    );
    if spans.style.empty {
        return expand(xml, &spans.style, &props);
    }
    let mut out = xml.to_owned();
    out.insert_str(spans.style.end_start, &props);
    Ok(out)
}

impl OpenDocumentPackage {
    pub fn paragraph_style_drop_caps(&self) -> Result<ParagraphStyleDropCapSet> {
        self.styles_xml()?.map_or_else(
            || Ok(ParagraphStyleDropCapSet::default()),
            |xml| parse_paragraph_style_drop_caps(&xml),
        )
    }
}
impl FlatOpenDocument {
    pub fn paragraph_style_drop_caps(&self) -> Result<ParagraphStyleDropCapSet> {
        parse_paragraph_style_drop_caps(self.xml())
    }
}
