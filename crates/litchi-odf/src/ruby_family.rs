//! Complete typed ODF ruby styles and structure-preserving inline annotations.

use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::{Error, Result, xml::escape_xml};
use quick_xml::{XmlVersion, events::{BytesStart, Event}, name::{Namespace, ResolveResult}, reader::NsReader};
use std::collections::HashSet;

const OFFICE_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const STYLE_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const TEXT_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const DRAW_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const PRESENTATION_URI: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MAX_XML: usize = 32 * 1_048_576;
const MAX_VALUE: usize = 1_048_576;
const MAX_BASE: usize = 4 * 1_048_576;
const MAX_DEPTH: usize = 256;
const MAX_EVENTS: usize = 1_000_000;
const MAX_RUBIES: usize = 65_536;
const MAX_STYLES: usize = 65_536;
const MAX_ATTRIBUTES: usize = 64;

fn bad(message: impl Into<String>) -> Error { Error::InvalidFormat(message.into()) }
fn validate_text(value: &str, context: &str, empty: bool) -> Result<()> {
    if (!empty && value.is_empty()) || value.len() > MAX_VALUE || value.chars().any(|c| matches!(c, '\0'..='\u{8}' | '\u{b}' | '\u{c}' | '\u{e}'..='\u{1f}' | '\u{fffe}' | '\u{ffff}')) {
        return Err(bad(format!("invalid {context}")));
    }
    Ok(())
}
fn ncname_start(c: char) -> bool { c == '_' || c.is_alphabetic() }
fn ncname_continue(c: char) -> bool { ncname_start(c) || c.is_alphanumeric() || matches!(c, '-' | '.' | '\u{b7}' | '\u{300}'..='\u{36f}' | '\u{203f}'..='\u{2040}') }
fn validate_style_name(value: &str, context: &str) -> Result<()> {
    validate_text(value, context, false)?;
    let mut chars = value.chars();
    if !chars.next().is_some_and(ncname_start) || !chars.all(ncname_continue) { return Err(bad(format!("{context} is not an NCName"))); }
    Ok(())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RubyPosition { Above, Below }
impl RubyPosition {
    pub const ALL: [Self; 2] = [Self::Above, Self::Below];
    pub const fn as_str(self) -> &'static str { match self { Self::Above => "above", Self::Below => "below" } }
    fn parse(value: &str) -> Result<Self> { match value { "above" => Ok(Self::Above), "below" => Ok(Self::Below), _ => Err(bad("invalid style:ruby-position")) } }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RubyAlignment { Left, Center, Right, DistributeLetter, DistributeSpace }
impl RubyAlignment {
    pub const ALL: [Self; 5] = [Self::Left, Self::Center, Self::Right, Self::DistributeLetter, Self::DistributeSpace];
    pub const fn as_str(self) -> &'static str { match self { Self::Left => "left", Self::Center => "center", Self::Right => "right", Self::DistributeLetter => "distribute-letter", Self::DistributeSpace => "distribute-space" } }
    fn parse(value: &str) -> Result<Self> { match value { "left" => Ok(Self::Left), "center" => Ok(Self::Center), "right" => Ok(Self::Right), "distribute-letter" => Ok(Self::DistributeLetter), "distribute-space" => Ok(Self::DistributeSpace), _ => Err(bad("invalid style:ruby-align")) } }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RubyProperties { pub position: Option<RubyPosition>, pub alignment: Option<RubyAlignment> }
impl RubyProperties {
    pub fn to_xml_fragment(&self) -> String {
        let mut xml = String::from(r#"<style:ruby-properties xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0""#);
        if let Some(value) = self.position { xml.push_str(&format!(r#" style:ruby-position="{}""#, value.as_str())); }
        if let Some(value) = self.alignment { xml.push_str(&format!(r#" style:ruby-align="{}""#, value.as_str())); }
        xml.push_str("/>"); xml
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RubyStyle { pub name: String, pub display_name: Option<String>, pub parent_style_name: Option<String>, pub properties: Option<RubyProperties> }
impl RubyStyle {
    pub fn new(name: impl Into<String>, properties: Option<RubyProperties>) -> Result<Self> {
        let value = Self { name: name.into(), display_name: None, parent_style_name: None, properties }; value.validate()?; Ok(value)
    }
    pub fn validate(&self) -> Result<()> {
        validate_style_name(&self.name, "ruby style name")?;
        if let Some(value) = &self.display_name { validate_text(value, "ruby style display name", true)?; }
        if let Some(value) = &self.parent_style_name { validate_style_name(value, "ruby parent style name")?; }
        Ok(())
    }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?;
        let mut xml = format!(r#"<style:style xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" style:name="{}" style:family="ruby""#, escape_xml(&self.name));
        if let Some(value) = &self.display_name { xml.push_str(&format!(r#" style:display-name="{}""#, escape_xml(value))); }
        if let Some(value) = &self.parent_style_name { xml.push_str(&format!(r#" style:parent-style-name="{}""#, escape_xml(value))); }
        if let Some(value) = &self.properties { xml.push('>'); xml.push_str(&value.to_xml_fragment()); xml.push_str("</style:style>"); } else { xml.push_str("/>"); }
        Ok(xml)
    }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RubyStyles { pub styles: Vec<RubyStyle> }
impl RubyStyles {
    pub fn get(&self, name: &str) -> Option<&RubyStyle> { self.styles.iter().find(|style| style.name == name) }
    pub fn validate(&self) -> Result<()> {
        if self.styles.len() > MAX_STYLES { return Err(bad("too many ruby styles")); }
        let mut names = HashSet::new();
        for style in &self.styles { style.validate()?; if !names.insert(style.name.as_str()) { return Err(bad("duplicate ruby style name")); } }
        Ok(())
    }
    pub fn to_xml(&self) -> Result<String> {
        self.validate()?;
        let mut xml = String::from(r#"<office:styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0">"#);
        for style in &self.styles { xml.push_str(&style.to_xml_fragment()?); }
        xml.push_str("</office:styles>"); Ok(xml)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
enum Ns { None, Office, Style, Text, Draw, Dr3d, Presentation, Other }
fn ns(value: &ResolveResult<'_>) -> Ns { match value { ResolveResult::Unbound => Ns::None, ResolveResult::Bound(Namespace(v)) if *v == OFFICE_URI => Ns::Office, ResolveResult::Bound(Namespace(v)) if *v == STYLE_URI => Ns::Style, ResolveResult::Bound(Namespace(v)) if *v == TEXT_URI => Ns::Text, ResolveResult::Bound(Namespace(v)) if *v == DRAW_URI => Ns::Draw, ResolveResult::Bound(Namespace(v)) if *v == DR3D_URI => Ns::Dr3d, ResolveResult::Bound(Namespace(v)) if *v == PRESENTATION_URI => Ns::Presentation, _ => Ns::Other } }
include!("ruby_inline_specs.rs");

fn name(value: &[u8]) -> Result<String> { std::str::from_utf8(value).map(str::to_owned).map_err(|_| bad("invalid UTF-8 XML name")) }
type Attrs = Vec<(Ns, String, String)>;
fn attributes(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Attrs> {
    let mut out = Vec::new(); let mut seen = HashSet::new();
    for attribute in start.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| bad(format!("invalid ruby attribute: {error}")))?;
        if attribute.key.as_ref() == b"xmlns" || attribute.key.as_ref().starts_with(b"xmlns:") { continue; }
        if out.len() >= MAX_ATTRIBUTES { return Err(bad("too many ruby attributes")); }
        let (resolved, local) = reader.resolver().resolve_attribute(attribute.key);
        let key = (ns(&resolved), name(local.as_ref())?);
        if !seen.insert(key.clone()) { return Err(bad("duplicate expanded ruby attribute")); }
        let value = attribute.decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder()).map_err(|error| bad(format!("invalid ruby attribute value: {error}")))?.into_owned();
        validate_text(&value, "ruby attribute value", true)?; out.push((key.0, key.1, value));
    }
    Ok(out)
}
fn only_style_ref(reader: &NsReader<&[u8]>, start: &BytesStart<'_>, context: &str) -> Result<Option<String>> {
    let mut style = None;
    for (namespace, local, value) in attributes(reader, start)? {
        if namespace == Ns::Text && local == "style-name" && style.is_none() { validate_style_name(&value, context)?; style = Some(value); }
        else { return Err(bad(format!("unsupported {context} attribute"))); }
    }
    Ok(style)
}
fn require_no_attrs(reader: &NsReader<&[u8]>, start: &BytesStart<'_>, context: &str) -> Result<()> { if attributes(reader, start)?.is_empty() { Ok(()) } else { Err(bad(format!("{context} has attributes"))) } }

fn parse_style_header(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<Option<RubyStyle>> {
    let attributes = attributes(reader, start)?;
    if !attributes.iter().any(|(namespace, local, value)| namespace == &Ns::Style && local == "family" && value == "ruby") {
        return Ok(None);
    }
    let mut family = None; let mut style_name = None; let mut display = None; let mut parent = None;
    for (namespace, local, value) in attributes {
        if namespace != Ns::Style { return Err(bad("ruby style attribute has wrong namespace")); }
        match local.as_str() { "family" => family = Some(value), "name" => style_name = Some(value), "display-name" => display = Some(value), "parent-style-name" => parent = Some(value), _ => return Err(bad("unsupported ruby style attribute")) }
    }
    if family.as_deref() != Some("ruby") { return Ok(None); }
    let value = RubyStyle { name: style_name.ok_or_else(|| bad("ruby style requires style:name"))?, display_name: display, parent_style_name: parent, properties: None }; value.validate()?; Ok(Some(value))
}
fn parse_properties(reader: &NsReader<&[u8]>, start: &BytesStart<'_>) -> Result<RubyProperties> {
    let mut value = RubyProperties::default();
    for (namespace, local, lexical) in attributes(reader, start)? {
        if namespace != Ns::Style { return Err(bad("ruby property has wrong attribute namespace")); }
        match local.as_str() { "ruby-position" if value.position.is_none() => value.position = Some(RubyPosition::parse(&lexical)?), "ruby-align" if value.alignment.is_none() => value.alignment = Some(RubyAlignment::parse(&lexical)?), _ => return Err(bad("unknown or duplicate ruby property")) }
    }
    Ok(value)
}

struct ActiveStyle { depth: usize, value: RubyStyle, property_depth: Option<usize>, seen: bool }
pub fn parse_ruby_styles(xml: &str) -> Result<RubyStyles> {
    if xml.len() > MAX_XML { return Err(bad("ruby styles XML is too large")); }
    let mut reader = NsReader::from_str(xml); reader.config_mut().trim_text(false);
    let mut buffer = Vec::new(); let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new(); let mut active: Option<ActiveStyle> = None; let mut styles = Vec::new(); let mut events = 0usize;
    loop {
        events += 1; if events > MAX_EVENTS { return Err(bad("too many ruby style events")); }
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| bad(format!("invalid ruby styles XML: {error}")))?; let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH { return Err(bad("ruby styles XML is too deep")); }
                let local = start.local_name().as_ref().to_vec(); let direct = matches!(stack.last(), Some((Ns::Office, parent)) if parent.as_slice() == b"styles" || parent.as_slice() == b"automatic-styles") && namespace == Ns::Style && local == b"style";
                stack.push((namespace, local.clone())); let depth = stack.len();
                if direct { if let Some(value) = parse_style_header(&reader, start)? { active = Some(ActiveStyle { depth, value, property_depth: None, seen: false }); } }
                else if namespace == Ns::Style && local == b"ruby-properties" {
                    let Some(style) = active.as_mut() else { return Err(bad("style:ruby-properties has invalid placement")); };
                    if depth != style.depth + 1 || style.seen { return Err(bad("duplicate or nested style:ruby-properties")); }
                    style.seen = true; style.value.properties = Some(parse_properties(&reader, start)?); style.property_depth = Some(depth);
                } else if active.as_ref().is_some_and(|style| depth > style.depth) { return Err(bad("ruby style has unsupported child")); }
            },
            Event::Empty(ref start) => {
                let local = start.local_name().as_ref().to_vec(); let depth = stack.len() + 1; let direct = matches!(stack.last(), Some((Ns::Office, parent)) if parent.as_slice() == b"styles" || parent.as_slice() == b"automatic-styles") && namespace == Ns::Style && local == b"style";
                if direct { if let Some(value) = parse_style_header(&reader, start)? { styles.push(value); } }
                else if namespace == Ns::Style && local == b"ruby-properties" {
                    let Some(style) = active.as_mut() else { return Err(bad("style:ruby-properties has invalid placement")); };
                    if depth != style.depth + 1 || style.seen { return Err(bad("duplicate or nested style:ruby-properties")); }
                    style.seen = true; style.value.properties = Some(parse_properties(&reader, start)?);
                } else if active.as_ref().is_some_and(|style| depth > style.depth) { return Err(bad("ruby style has unsupported child")); }
            },
            Event::Text(ref text) if active.is_some() => { let bytes: &[u8] = text.as_ref(); if !bytes.iter().all(u8::is_ascii_whitespace) { return Err(bad("ruby style cannot contain text")); } },
            Event::CData(_) if active.is_some() => return Err(bad("ruby style cannot contain CDATA")),
            Event::End(_) => {
                let depth = stack.len();
                if active.as_ref().is_some_and(|style| style.property_depth == Some(depth)) { active.as_mut().unwrap().property_depth = None; }
                if active.as_ref().is_some_and(|style| style.depth == depth) { styles.push(active.take().unwrap().value); }
                stack.pop();
            },
            Event::DocType(_) | Event::PI(_) => return Err(bad("DTD and processing instructions are prohibited")), Event::Eof => break, _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() || active.is_some() { return Err(bad("truncated ruby styles XML")); }
    let value = RubyStyles { styles }; value.validate()?; Ok(value)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RubyBase { xml: String }
impl RubyBase {
    pub fn from_text(text: &str) -> Result<Self> { validate_text(text, "ruby base text", true)?; if text.len() > MAX_BASE { return Err(bad("ruby base is too large")); } Ok(Self { xml: escape_xml(text) }) }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> { if fragment.len() > MAX_BASE { return Err(bad("ruby base is too large")); } let ruby = RubyAnnotation::new(None, Self { xml: fragment.to_owned() }, "", None)?; RubyAnnotation::from_xml_fragment(&ruby.to_xml_fragment()?).map(|value| value.base) }
    pub fn xml(&self) -> &str { &self.xml }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RubyAnnotation { pub style_name: Option<String>, pub base: RubyBase, pub text: String, pub text_style_name: Option<String> }
impl RubyAnnotation {
    pub fn new(style_name: Option<String>, base: RubyBase, text: impl Into<String>, text_style_name: Option<String>) -> Result<Self> { let value = Self { style_name, base, text: text.into(), text_style_name }; value.validate()?; Ok(value) }
    pub fn validate(&self) -> Result<()> { if let Some(value) = &self.style_name { validate_style_name(value, "ruby style reference")?; } if let Some(value) = &self.text_style_name { validate_style_name(value, "ruby text style reference")?; } validate_text(&self.text, "ruby pronunciation", true)?; if self.base.xml.len() > MAX_BASE { return Err(bad("ruby base is too large")); } Ok(()) }
    pub fn to_xml_fragment(&self) -> Result<String> {
        self.validate()?; let mut xml = String::from(r#"<text:ruby xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0""#);
        if let Some(value) = &self.style_name { xml.push_str(&format!(r#" text:style-name="{}""#, escape_xml(value))); }
        xml.push_str("><text:ruby-base>"); xml.push_str(&self.base.xml); xml.push_str("</text:ruby-base><text:ruby-text");
        if let Some(value) = &self.text_style_name { xml.push_str(&format!(r#" text:style-name="{}""#, escape_xml(value))); }
        xml.push('>'); xml.push_str(&escape_xml(&self.text)); xml.push_str("</text:ruby-text></text:ruby>"); Ok(xml)
    }
    pub fn from_xml_fragment(fragment: &str) -> Result<Self> { let xml = format!(r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{fragment}</text:p>"#); let mut entries = parse_ruby_entries(&xml)?; entries.sort_by_key(|entry| entry.span.start); entries.into_iter().next().map(|entry| entry.value).ok_or_else(|| bad("fragment contains no text:ruby")) }
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RubyAnnotations { pub annotations: Vec<RubyAnnotation> }
#[derive(Clone)] struct Span { start: usize, end: usize }
struct RubyEntry { value: RubyAnnotation, span: Span }
struct ActiveRuby { depth: usize, start: usize, style_name: Option<String>, base_depth: Option<usize>, base_start: usize, base: Option<(usize, usize)>, text_depth: Option<usize>, text_seen: bool, text_style_name: Option<String>, text: String }
fn event_start(xml: &str, end: usize) -> Result<usize> { xml[..end].rfind('<').ok_or_else(|| bad("invalid ruby XML event boundary")) }
fn ruby_parent(parent: Option<&(Ns, Vec<u8>)>) -> bool { matches!(parent, Some((Ns::Text, local)) if matches!(local.as_slice(), b"p" | b"h" | b"span" | b"a" | b"meta" | b"meta-field" | b"ruby-base")) }
fn parse_ruby_entries(xml: &str) -> Result<Vec<RubyEntry>> {
    if xml.len() > MAX_XML { return Err(bad("ruby XML is too large")); }
    let mut reader = NsReader::from_str(xml); reader.config_mut().trim_text(false); let mut buffer = Vec::new(); let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new(); let mut active = Vec::<ActiveRuby>::new(); let mut entries = Vec::new(); let mut events = 0usize;
    loop {
        events += 1; if events > MAX_EVENTS { return Err(bad("too many ruby XML events")); }
        let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| bad(format!("invalid ruby XML: {error}")))?; let namespace = ns(&resolved);
        match event {
            Event::Start(ref start) => {
                if stack.len() >= MAX_DEPTH { return Err(bad("ruby XML is too deep")); }
                let end = reader.buffer_position() as usize; let begin = event_start(xml, end)?; let local = start.local_name().as_ref().to_vec(); let depth = stack.len() + 1; let parent = stack.last();
                if matches!(local.as_slice(), b"ruby" | b"ruby-base" | b"ruby-text") && namespace != Ns::Text { return Err(bad("ruby element uses wrong namespace")); }
                if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) { return Err(bad("text:ruby-text may contain only text")); }
                if namespace == Ns::Text && local == b"ruby" {
                    if !ruby_parent(parent) { return Err(bad("text:ruby has invalid placement")); }
                    if entries.len() + active.len() >= MAX_RUBIES { return Err(bad("too many ruby annotations")); }
                    active.push(ActiveRuby { depth, start: begin, style_name: only_style_ref(&reader, start, "ruby style reference")?, base_depth: None, base_start: 0, base: None, text_depth: None, text_seen: false, text_style_name: None, text: String::new() });
                } else if let Some(ruby) = active.last_mut() {
                    if depth == ruby.depth + 1 {
                        if ruby.base.is_none() && ruby.base_depth.is_none() && namespace == Ns::Text && local == b"ruby-base" { require_no_attrs(&reader, start, "text:ruby-base")?; ruby.base_depth = Some(depth); ruby.base_start = end; }
                        else if ruby.base.is_some() && !ruby.text_seen && namespace == Ns::Text && local == b"ruby-text" { ruby.text_style_name = only_style_ref(&reader, start, "ruby text style reference")?; ruby.text_depth = Some(depth); ruby.text_seen = true; }
                        else { return Err(bad("text:ruby requires ruby-base then ruby-text")); }
                    } else if ruby.base_depth.is_some() && matches!(parent, Some((Ns::Text, p)) if matches!(p.as_slice(), b"ruby-base" | b"span" | b"meta" | b"meta-field")) && !is_ruby_base_child(namespace, &local) { return Err(bad("unsupported text:ruby-base inline child")); }
                    else if ruby.base_depth.is_some() && matches!(parent, Some((Ns::Text, p)) if p.as_slice() == b"a") && !is_hyperlink_child(namespace, &local) { return Err(bad("unsupported hyperlink inline child")); }
                }
                stack.push((namespace, local));
            },
            Event::Empty(ref start) => {
                let end = reader.buffer_position() as usize; let local = start.local_name().as_ref().to_vec(); let depth = stack.len() + 1; let parent = stack.last();
                if matches!(local.as_slice(), b"ruby" | b"ruby-base" | b"ruby-text") && namespace != Ns::Text { return Err(bad("ruby element uses wrong namespace")); }
                if namespace == Ns::Text && local == b"ruby" { return Err(bad("text:ruby requires ruby-base and ruby-text")); }
                if let Some(ruby) = active.last_mut() {
                    if ruby.text_depth.is_some() { return Err(bad("text:ruby-text may contain only text")); }
                    if depth == ruby.depth + 1 {
                        if ruby.base.is_none() && ruby.base_depth.is_none() && namespace == Ns::Text && local == b"ruby-base" { require_no_attrs(&reader, start, "text:ruby-base")?; ruby.base = Some((end, end)); }
                        else if ruby.base.is_some() && !ruby.text_seen && namespace == Ns::Text && local == b"ruby-text" { ruby.text_style_name = only_style_ref(&reader, start, "ruby text style reference")?; ruby.text_seen = true; }
                        else { return Err(bad("text:ruby requires ruby-base then ruby-text")); }
                    } else if ruby.base_depth.is_some() && matches!(parent, Some((Ns::Text, p)) if matches!(p.as_slice(), b"ruby-base" | b"span" | b"meta" | b"meta-field")) && !is_ruby_base_child(namespace, &local) { return Err(bad("unsupported text:ruby-base inline child")); }
                    else if ruby.base_depth.is_some() && matches!(parent, Some((Ns::Text, p)) if p.as_slice() == b"a") && !is_hyperlink_child(namespace, &local) { return Err(bad("unsupported hyperlink inline child")); }
                }
            },
            Event::Text(ref value) if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) => { let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| bad(format!("invalid ruby text: {error}")))?; let ruby = active.last_mut().unwrap(); if ruby.text.len() + value.len() > MAX_VALUE { return Err(bad("ruby pronunciation is too large")); } ruby.text.push_str(&value); },
            Event::CData(ref value) if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) => { let value = value.xml_content(XmlVersion::Explicit1_0).map_err(|error| bad(format!("invalid ruby CDATA: {error}")))?; active.last_mut().unwrap().text.push_str(&value); },
            Event::GeneralRef(ref value) if active.last().is_some_and(|ruby| ruby.text_depth.is_some()) => { active.last_mut().unwrap().text.push_str(&crate::elements::xml::decode_reference(value, "ruby")?); },
            Event::End(_) => {
                let end = reader.buffer_position() as usize; let begin = event_start(xml, end)?; let depth = stack.len(); let frame = stack.last().ok_or_else(|| bad("ruby XML depth underflow"))?;
                if let Some(ruby) = active.last_mut() {
                    if ruby.base_depth == Some(depth) { if frame.0 != Ns::Text || frame.1 != b"ruby-base" { return Err(bad("invalid ruby-base end")); } if begin < ruby.base_start || begin - ruby.base_start > MAX_BASE { return Err(bad("ruby base is too large")); } ruby.base = Some((ruby.base_start, begin)); ruby.base_depth = None; }
                    if ruby.text_depth == Some(depth) { ruby.text_depth = None; }
                }
                if active.last().is_some_and(|ruby| ruby.depth == depth) {
                    let ruby = active.pop().unwrap(); if frame.0 != Ns::Text || frame.1 != b"ruby" || ruby.base.is_none() || !ruby.text_seen || ruby.base_depth.is_some() || ruby.text_depth.is_some() { return Err(bad("text:ruby requires ruby-base then ruby-text")); }
                    let (base_start, base_end) = ruby.base.unwrap(); let value = RubyAnnotation::new(ruby.style_name, RubyBase { xml: xml[base_start..base_end].to_owned() }, ruby.text, ruby.text_style_name)?; entries.push(RubyEntry { value, span: Span { start: ruby.start, end } });
                }
                stack.pop();
            },
            Event::DocType(_) | Event::PI(_) => return Err(bad("DTD and processing instructions are prohibited")), Event::Eof => break, _ => {}
        }
        buffer.clear();
    }
    if !stack.is_empty() || !active.is_empty() { return Err(bad("truncated ruby XML")); }
    entries.sort_by_key(|entry| entry.span.start); Ok(entries)
}

pub fn parse_ruby_annotations(xml: &str) -> Result<RubyAnnotations> { Ok(RubyAnnotations { annotations: parse_ruby_entries(xml)?.into_iter().map(|entry| entry.value).collect() }) }
pub fn replace_ruby_annotation_xml(xml: &str, index: usize, value: &RubyAnnotation) -> Result<String> { value.validate()?; let entries = parse_ruby_entries(xml)?; let span = &entries.get(index).ok_or_else(|| bad("ruby annotation index does not exist"))?.span; Ok(format!("{}{}{}", &xml[..span.start], value.to_xml_fragment()?, &xml[span.end..])) }
pub fn remove_ruby_annotation_xml(xml: &str, index: usize) -> Result<String> { let entries = parse_ruby_entries(xml)?; let span = &entries.get(index).ok_or_else(|| bad("ruby annotation index does not exist"))?.span; Ok(format!("{}{}", &xml[..span.start], &xml[span.end..])) }
pub fn insert_ruby_annotation_xml(xml: &str, paragraph_index: usize, value: &RubyAnnotation) -> Result<String> {
    value.validate()?; parse_ruby_entries(xml)?; let fragment = value.to_xml_fragment()?; let mut reader = NsReader::from_str(xml); let mut buffer = Vec::new(); let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new(); let mut count = 0usize; let mut target = None::<(usize, String)>;
    loop { let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| bad(format!("invalid ruby insertion XML: {error}")))?; let namespace = ns(&resolved); match event {
        Event::Start(ref start) => { let local = start.local_name().as_ref().to_vec(); if namespace == Ns::Text && local == b"p" { if count == paragraph_index { target = Some((stack.len() + 1, name(start.name().as_ref())?)); } count += 1; } stack.push((namespace, local)); },
        Event::Empty(ref start) if namespace == Ns::Text && start.local_name().as_ref() == b"p" => { let end = reader.buffer_position() as usize; let begin = event_start(xml, end)?; if count == paragraph_index { let raw = &xml[begin..end]; let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty paragraph"))?; let qname = name(start.name().as_ref())?; return Ok(format!("{}{}>{}</{}>{}", &xml[..begin], &raw[..slash], fragment, qname, &xml[end..])); } count += 1; },
        Event::End(_) => { let depth = stack.len(); if target.as_ref().is_some_and(|(target_depth, _)| *target_depth == depth) { let begin = event_start(xml, reader.buffer_position() as usize)?; return Ok(format!("{}{}{}", &xml[..begin], fragment, &xml[begin..])); } stack.pop(); },
        Event::Eof => break, _ => {} } buffer.clear(); }
    Err(bad("paragraph index does not exist"))
}

#[derive(Clone)] enum StyleSite { Content(usize), Empty(Span, String) }
fn locate_ruby_style(xml: &str, target_name: &str) -> Result<(Option<Span>, StyleSite)> {
    parse_ruby_styles(xml)?; let mut reader = NsReader::from_str(xml); let mut buffer = Vec::new(); let mut stack: Vec<(Ns, Vec<u8>)> = Vec::new(); let mut target = None; let mut open = None::<(usize, usize)>; let mut site = None;
    loop { let (resolved, event) = reader.read_resolved_event_into(&mut buffer).map_err(|error| bad(format!("invalid ruby style mutation XML: {error}")))?; let namespace = ns(&resolved); match event {
        Event::Start(ref start) => { let end = reader.buffer_position() as usize; let begin = event_start(xml, end)?; let local = start.local_name().as_ref().to_vec(); let depth = stack.len() + 1; if namespace == Ns::Style && local == b"style" && matches!(stack.last(), Some((Ns::Office, parent)) if parent == b"styles" || parent == b"automatic-styles") { if parse_style_header(&reader, start)?.is_some_and(|style| style.name == target_name) { open = Some((depth, begin)); } } stack.push((namespace, local)); },
        Event::Empty(ref start) => { let end = reader.buffer_position() as usize; let begin = event_start(xml, end)?; let local = start.local_name().as_ref().to_vec(); if namespace == Ns::Style && local == b"style" && matches!(stack.last(), Some((Ns::Office, parent)) if parent == b"styles" || parent == b"automatic-styles") && parse_style_header(&reader, start)?.is_some_and(|style| style.name == target_name) { target = Some(Span { start: begin, end }); } if namespace == Ns::Office && local == b"styles" { site = Some(StyleSite::Empty(Span { start: begin, end }, name(start.name().as_ref())?)); } },
        Event::End(_) => { let depth = stack.len(); let begin = event_start(xml, reader.buffer_position() as usize)?; if open.is_some_and(|(d, _)| d == depth) { let (_, start) = open.take().unwrap(); target = Some(Span { start, end: reader.buffer_position() as usize }); } if matches!(stack.last(), Some((Ns::Office, local)) if local == b"styles") { site = Some(StyleSite::Content(begin)); } stack.pop(); }, Event::Eof => break, _ => {} } buffer.clear(); }
    Ok((target, site.ok_or_else(|| bad("document has no office:styles"))?))
}
pub fn set_ruby_style_xml(xml: &str, style: &RubyStyle) -> Result<String> { style.validate()?; let (target, site) = locate_ruby_style(xml, &style.name)?; let fragment = style.to_xml_fragment()?; if let Some(span) = target { return Ok(format!("{}{}{}", &xml[..span.start], fragment, &xml[span.end..])); } match site { StyleSite::Content(at) => Ok(format!("{}{}{}", &xml[..at], fragment, &xml[at..])), StyleSite::Empty(span, qname) => { let raw = &xml[span.start..span.end]; let slash = raw.rfind("/>").ok_or_else(|| bad("invalid empty office:styles"))?; Ok(format!("{}{}>{}</{}>{}", &xml[..span.start], &raw[..slash], fragment, qname, &xml[span.end..])) } } }
pub fn remove_ruby_style_xml(xml: &str, name: &str) -> Result<String> { validate_style_name(name, "ruby style name")?; let (target, _) = locate_ruby_style(xml, name)?; let Some(span) = target else { return Ok(xml.to_owned()); }; Ok(format!("{}{}", &xml[..span.start], &xml[span.end..])) }

impl OpenDocumentPackage {
    pub fn ruby_styles(&self) -> Result<RubyStyles> { self.styles_xml()?.map_or_else(|| Ok(Default::default()), |xml| parse_ruby_styles(&xml)) }
    pub fn ruby_annotations(&self) -> Result<RubyAnnotations> { parse_ruby_annotations(&self.content_xml()?) }
}
impl FlatOpenDocument {
    pub fn ruby_styles(&self) -> Result<RubyStyles> { parse_ruby_styles(self.xml()) }
    pub fn ruby_annotations(&self) -> Result<RubyAnnotations> { parse_ruby_annotations(self.xml()) }
}

#[cfg(test)] mod tests {
    use super::*;
    const HEAD: &str = r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:styles>"#;
    fn styles(body: &str) -> String { format!("{HEAD}{body}</office:styles><office:body><office:text><text:p/></office:text></office:body></office:document>") }
    #[test] fn exhaustive_properties_round_trip() { for position in RubyPosition::ALL { for alignment in RubyAlignment::ALL { let style = RubyStyle::new(format!("s{}_{}", position.as_str(), alignment.as_str()), Some(RubyProperties { position: Some(position), alignment: Some(alignment) })).unwrap(); let xml = styles(&style.to_xml_fragment().unwrap()); assert_eq!(parse_ruby_styles(&xml).unwrap().styles[0], style); } } }
    #[test] fn mixed_inline_round_trip() { let base = RubyBase::from_xml_fragment(r#" A <text:span text:style-name="Em"> B </text:span><text:a xlink:type="simple" xlink:href="https://example.invalid/">一日</text:a> "#).unwrap(); let ruby = RubyAnnotation::new(Some("Ru1".into()), base, "ついたち", Some("RubyText".into())).unwrap(); let parsed = RubyAnnotation::from_xml_fragment(&ruby.to_xml_fragment().unwrap()).unwrap(); assert_eq!(parsed, ruby); }
    #[test] fn parses_real_libreoffice_fixture() { let xml = include_str!("../../../3rdparty/libreoffice-core/sw/qa/extras/odfexport/data/ruby+hyperlink.fodt"); let styles = parse_ruby_styles(xml).unwrap(); assert_eq!(styles.styles[0].properties.as_ref().unwrap().alignment, Some(RubyAlignment::Left)); let annotations = parse_ruby_annotations(xml).unwrap(); assert_eq!(annotations.annotations[0].text, "ついたち"); assert!(annotations.annotations[0].base.xml().contains("xlink:href")); }
    #[test] fn rejects_malformed_order_namespace_and_lexicals() { for body in [r#"<text:ruby><text:ruby-text>x</text:ruby-text><text:ruby-base>X</text:ruby-base></text:ruby>"#, r#"<text:ruby><text:ruby-base>X</text:ruby-base><text:ruby-text><text:span>x</text:span></text:ruby-text></text:ruby>"#, r#"<text:ruby text:style-name="1bad"><text:ruby-base>X</text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby>"#, r#"<text:ruby><text:ruby-base><text:unknown/></text:ruby-base><text:ruby-text>x</text:ruby-text></text:ruby>"#] { let xml = format!(r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{body}</text:p>"#); assert!(parse_ruby_annotations(&xml).is_err(), "accepted {body}"); } let wrong = styles(r#"<style:style style:name="r" style:family="ruby"><style:ruby-properties style:ruby-align="justify"/></style:style>"#); assert!(parse_ruby_styles(&wrong).is_err()); let duplicate = styles(r#"<style:style style:name="r" style:family="ruby"><style:ruby-properties style:ruby-align="left"/><style:ruby-properties/></style:style>"#); assert!(parse_ruby_styles(&duplicate).is_err()); }
    #[test] fn enforces_caps() { assert!(RubyBase::from_text(&"x".repeat(MAX_BASE + 1)).is_err()); let mut xml = String::from(r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">"#); for _ in 0..=MAX_DEPTH { xml.push_str("<text:span>"); } for _ in 0..=MAX_DEPTH { xml.push_str("</text:span>"); } xml.push_str("</text:p>"); assert!(parse_ruby_annotations(&xml).is_err()); let unit = "<text:ruby><text:ruby-base/><text:ruby-text/></text:ruby>"; let sequential = format!(r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">{}</text:p>"#, unit.repeat(MAX_RUBIES + 1)); assert!(parse_ruby_annotations(&sequential).is_err()); }
    #[test] fn ruby_style_namespace_rejection_is_attribute_order_independent() { let xml = styles(r#"<style:style xmlns:x="urn:wrong" x:foreign="first" style:name="Ru" style:family="ruby"/>"#); assert!(parse_ruby_styles(&xml).is_err()); }
    #[test] fn lossless_style_and_inline_mutation() { let original = styles("<!--keep--><style:style style:name=\"other\" style:family=\"text\"/>"); let style = RubyStyle::new("Ru1", Some(RubyProperties { position: Some(RubyPosition::Above), alignment: Some(RubyAlignment::Center) })).unwrap(); let inserted = set_ruby_style_xml(&original, &style).unwrap(); assert!(inserted.contains("<!--keep--><style:style style:name=\"other\"")); assert_eq!(remove_ruby_style_xml(&inserted, "Ru1").unwrap(), original); let ruby = RubyAnnotation::new(None, RubyBase::from_text("base").unwrap(), "ruby", None).unwrap(); let paragraph = r#"<text:p xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">before</text:p>"#; let with = insert_ruby_annotation_xml(paragraph, 0, &ruby).unwrap(); let replacement = RubyAnnotation::new(None, RubyBase::from_text("B").unwrap(), "R", None).unwrap(); let replaced = replace_ruby_annotation_xml(&with, 0, &replacement).unwrap(); assert!(replaced.contains(">B</text:ruby-base>")); assert_eq!(remove_ruby_annotation_xml(&replaced, 0).unwrap(), paragraph); }
}
