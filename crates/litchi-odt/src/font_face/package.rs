//! Package and flat-document adapters for font-face declarations.

use super::{
    MAX_XML_DEPTH, NamespaceKind,
    codec::{namespace_kind, parse},
    invalid,
    model::Declarations,
    xml_error,
};
use crate::{FlatDocument, Package};
use litchi_core::{Error, Result};
use quick_xml::{events::Event, reader::NsReader};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Part {
    Content,
    Styles,
    Flat,
}

impl Part {
    fn root_local(self) -> &'static [u8] {
        match self {
            Self::Content => b"document-content",
            Self::Styles => b"document-styles",
            Self::Flat => b"document",
        }
    }

    fn root_name(self) -> &'static str {
        match self {
            Self::Content => "office:document-content",
            Self::Styles => "office:document-styles",
            Self::Flat => "office:document",
        }
    }
}

#[derive(Clone)]
struct XmlSpan {
    start: usize,
    end: usize,
}

struct Location {
    target: Option<XmlSpan>,
    insertion: usize,
}

fn parse_font_face_declarations_in_part(xml: &str, part: Part) -> Result<Option<Declarations>> {
    Ok(locate_font_face_declarations(xml, part)?.0)
}

pub(crate) fn parse_content_font_face_declarations(xml: &str) -> Result<Option<Declarations>> {
    parse_font_face_declarations_in_part(xml, Part::Content)
}

pub(crate) fn parse_styles_font_face_declarations(xml: &str) -> Result<Option<Declarations>> {
    parse_font_face_declarations_in_part(xml, Part::Styles)
}

/// Insert or replace content-part font-face declarations without rewriting
/// unrelated content XML.
pub(crate) fn set_content_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
) -> Result<(String, Option<Declarations>)> {
    set_font_face_declarations_xml(xml, declarations, Part::Content)
}

/// Insert or replace styles-part font-face declarations without rewriting
/// unrelated styles XML.
pub(crate) fn set_styles_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
) -> Result<(String, Option<Declarations>)> {
    set_font_face_declarations_xml(xml, declarations, Part::Styles)
}

/// Remove content-part font-face declarations without rewriting unrelated
/// content XML.
pub(crate) fn remove_content_font_face_declarations_xml(
    xml: &str,
) -> Result<(String, Option<Declarations>)> {
    remove_font_face_declarations_xml(xml, Part::Content)
}

/// Remove styles-part font-face declarations without rewriting unrelated
/// styles XML.
pub(crate) fn remove_styles_font_face_declarations_xml(
    xml: &str,
) -> Result<(String, Option<Declarations>)> {
    remove_font_face_declarations_xml(xml, Part::Styles)
}

fn set_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
    part: Part,
) -> Result<(String, Option<Declarations>)> {
    declarations.validate()?;
    let (old, location) = locate_font_face_declarations(xml, part)?;
    let fragment = declarations.to_xml()?;
    let updated = if let Some(target) = location.target {
        replace_span(xml, &target, &fragment)
    } else {
        insert_at(xml, location.insertion, &fragment)
    };
    parse_font_face_declarations_in_part(&updated, part)?;
    Ok((updated, old))
}

fn remove_font_face_declarations_xml(
    xml: &str,
    part: Part,
) -> Result<(String, Option<Declarations>)> {
    let (old, location) = locate_font_face_declarations(xml, part)?;
    let Some(target) = location.target else {
        return Ok((xml.to_owned(), old));
    };
    let updated = replace_span(xml, &target, "");
    parse_font_face_declarations_in_part(&updated, part)?;
    Ok((updated, old))
}

fn locate_font_face_declarations(
    xml: &str,
    part: Part,
) -> Result<(Option<Declarations>, Location)> {
    let declarations = parse(xml)?;
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::<(NamespaceKind, Vec<u8>)>::new();
    let mut root_open_end = None;
    let mut root_closed = false;
    let mut target = None;
    let mut open_target = None::<(usize, usize)>;
    let mut scripts_end = None;
    let mut open_scripts = None::<usize>;

    loop {
        let (resolved, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        let namespace = namespace_kind(&resolved);
        match event {
            Event::Start(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth > MAX_XML_DEPTH {
                    return invalid(format!("ODF XML exceeds the {MAX_XML_DEPTH} depth limit"));
                }
                if depth == 1 {
                    if namespace != NamespaceKind::Office || local != part.root_local() {
                        return invalid(format!(
                            "font-face declarations require a {} root",
                            part.root_name()
                        ));
                    }
                    root_open_end = Some(end);
                } else if depth == 2 {
                    if namespace == NamespaceKind::Office && local == b"font-face-decls" {
                        if target.is_some() || open_target.is_some() {
                            return invalid("ODF XML contains duplicate office:font-face-decls");
                        }
                        open_target = Some((depth, start));
                    } else if part == Part::Content
                        && namespace == NamespaceKind::Office
                        && local == b"scripts"
                    {
                        open_scripts = Some(depth);
                    }
                }
                stack.push((namespace, local));
            },
            Event::Empty(element) => {
                let end = reader.buffer_position() as usize;
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                let depth = stack
                    .len()
                    .checked_add(1)
                    .ok_or_else(|| Error::InvalidFormat("ODF XML depth overflow".to_string()))?;
                if depth == 1 {
                    return invalid(format!(
                        "font-face declarations require a non-empty {} root",
                        part.root_name()
                    ));
                }
                if depth == 2 && namespace == NamespaceKind::Office && local == b"font-face-decls" {
                    if target.is_some() || open_target.is_some() {
                        return invalid("ODF XML contains duplicate office:font-face-decls");
                    }
                    target = Some(XmlSpan { start, end });
                } else if depth == 2
                    && part == Part::Content
                    && namespace == NamespaceKind::Office
                    && local == b"scripts"
                {
                    scripts_end = Some(end);
                }
            },
            Event::End(_) => {
                let end = reader.buffer_position() as usize;
                let depth = stack.len();
                if open_target.is_some_and(|(target_depth, _)| target_depth == depth) {
                    let (_, start) = open_target.take().expect("target depth was checked");
                    target = Some(XmlSpan { start, end });
                }
                if open_scripts.is_some_and(|scripts_depth| scripts_depth == depth) {
                    open_scripts = None;
                    scripts_end = Some(end);
                }
                if depth == 1 {
                    root_closed = true;
                }
                stack.pop().ok_or_else(|| {
                    Error::InvalidFormat("invalid ODF XML font-face element depth".to_string())
                })?;
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if root_open_end.is_none() || !root_closed || !stack.is_empty() || open_target.is_some() {
        return invalid(format!(
            "unterminated {} while locating font-face declarations",
            part.root_name()
        ));
    }
    let insertion = match part {
        Part::Content => scripts_end.unwrap_or_else(|| {
            root_open_end.expect("non-empty document root has an opening event")
        }),
        Part::Styles | Part::Flat => {
            root_open_end.expect("non-empty document root has an opening event")
        },
    };
    Ok((declarations, Location { target, insertion }))
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| Error::InvalidFormat("invalid ODF font-face XML event boundary".to_string()))
}

fn replace_span(xml: &str, span: &XmlSpan, replacement: &str) -> String {
    let mut output = String::with_capacity(xml.len() - (span.end - span.start) + replacement.len());
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    output
}

fn insert_at(xml: &str, insertion: usize, fragment: &str) -> String {
    let mut output = String::with_capacity(xml.len() + fragment.len());
    output.push_str(&xml[..insertion]);
    output.push_str(fragment);
    output.push_str(&xml[insertion..]);
    output
}

impl Package {
    /// Return content-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<Declarations>> {
        let xml = self.content_xml()?;
        parse_content_font_face_declarations(&xml)
    }

    /// Return styles-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<Declarations>> {
        self.styles_xml()?
            .map_or_else(|| Ok(None), |xml| parse_styles_font_face_declarations(&xml))
    }
}

impl FlatDocument {
    /// Return the flat document's font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn font_face_declarations(&self) -> Result<Option<Declarations>> {
        parse_font_face_declarations_in_part(self.xml(), Part::Flat)
    }
}
