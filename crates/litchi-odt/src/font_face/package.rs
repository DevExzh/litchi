//! Package and flat-document adapters for font-face declarations.

use super::{
    MAX_XML_DEPTH, NamespaceKind,
    codec::{namespace_kind, parse},
    invalid,
    model::Declarations,
    xml_error,
};
use crate::{
    FlatDocument, GenericPackage as Package,
    core::{AuthoredXmlFragment, OwnedPackage, XmlSourcePart, XmlSplicePublication},
};
use litchi_core::{Error, Result};
use litchi_odf_common::XmlStreamLimits;
use quick_xml::{events::Event, reader::NsReader};
use std::ops::Range;

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

fn parse_font_face_declarations_in_part(
    xml: &str,
    part: Part,
    limits: &XmlStreamLimits,
) -> Result<Option<Declarations>> {
    Ok(locate_font_face_declarations(xml, part, limits)?.0)
}

pub(crate) fn parse_content_font_face_declarations(
    xml: &str,
    limits: &XmlStreamLimits,
) -> Result<Option<Declarations>> {
    parse_font_face_declarations_in_part(xml, Part::Content, limits)
}

pub(crate) fn parse_styles_font_face_declarations(
    xml: &str,
    limits: &XmlStreamLimits,
) -> Result<Option<Declarations>> {
    parse_font_face_declarations_in_part(xml, Part::Styles, limits)
}

/// Insert or replace content-part font-face declarations without rewriting
/// unrelated content XML.
pub(crate) fn set_content_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    set_font_face_declarations_xml(xml, declarations, Part::Content, limits)
}

/// Insert or replace styles-part font-face declarations without rewriting
/// unrelated styles XML.
pub(crate) fn set_styles_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    set_font_face_declarations_xml(xml, declarations, Part::Styles, limits)
}

/// Remove content-part font-face declarations without rewriting unrelated
/// content XML.
pub(crate) fn remove_content_font_face_declarations_xml(
    xml: &str,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    remove_font_face_declarations_xml(xml, Part::Content, limits)
}

/// Remove styles-part font-face declarations without rewriting unrelated
/// styles XML.
pub(crate) fn remove_styles_font_face_declarations_xml(
    xml: &str,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    remove_font_face_declarations_xml(xml, Part::Styles, limits)
}

pub(crate) fn content_font_face_splice_publication(
    source: &OwnedPackage,
    candidate: &str,
    limits: &XmlStreamLimits,
) -> Result<Option<XmlSplicePublication>> {
    font_face_splice_publication(source, "content.xml", candidate, Part::Content, limits)
}

pub(crate) fn styles_font_face_splice_publication(
    source: &OwnedPackage,
    candidate: &str,
    limits: &XmlStreamLimits,
) -> Result<Option<XmlSplicePublication>> {
    font_face_splice_publication(source, "styles.xml", candidate, Part::Styles, limits)
}

fn font_face_splice_publication(
    package: &OwnedPackage,
    path: &str,
    candidate: &str,
    part: Part,
    limits: &XmlStreamLimits,
) -> Result<Option<XmlSplicePublication>> {
    let source = XmlSourcePart::load(package, path)?;
    if source.bytes() == candidate.as_bytes() {
        return Ok(Some(XmlSplicePublication::new(source)));
    }
    let source_xml = std::str::from_utf8(source.bytes()).map_err(|error| {
        Error::InvalidFormat(format!("invalid UTF-8 in source {path}: {error}"))
    })?;
    let (_, source_location) = locate_font_face_declarations(source_xml, part, limits)?;
    let (_, candidate_location) = locate_font_face_declarations(candidate, part, limits)?;
    let source_range = declaration_or_insertion_range(&source_location);
    let candidate_range = declaration_or_insertion_range(&candidate_location);
    let replacement = candidate
        .as_bytes()
        .get(candidate_range.clone())
        .ok_or_else(|| {
            Error::InvalidFormat("invalid candidate font-face splice range".to_string())
        })?;
    let fragment = if replacement.is_empty() {
        AuthoredXmlFragment::deletion()
    } else {
        AuthoredXmlFragment::markup(replacement.to_vec())?
    };

    if source_xml.get(..source_range.start) == candidate.get(..candidate_range.start)
        && source_xml.get(source_range.end..) == candidate.get(candidate_range.end..)
    {
        let expected = source_xml
            .as_bytes()
            .get(source_range.clone())
            .ok_or_else(|| {
                Error::InvalidFormat("invalid source font-face splice range".to_string())
            })?;
        let proof = source.checked_range(source_range, expected)?;
        let mut publication = XmlSplicePublication::new(source);
        publication.replace(proof, fragment)?;
        return Ok(Some(publication));
    }

    let (Some(source_target), Some(_candidate_target)) =
        (&source_location.target, &candidate_location.target)
    else {
        return Ok(None);
    };
    let source_without = remove_range(source_xml.as_bytes(), source_range.clone())?;
    let candidate_without = remove_range(candidate.as_bytes(), candidate_range.clone())?;
    if source_without != candidate_without {
        return Ok(None);
    }
    let source_insertion = if candidate_range.start < source_target.start {
        candidate_range.start
    } else {
        candidate_range
            .start
            .checked_add(source_range.len())
            .ok_or_else(|| Error::InvalidFormat("font-face splice offset overflow".to_string()))?
    };
    let deletion_bytes = source_xml
        .as_bytes()
        .get(source_range.clone())
        .ok_or_else(|| Error::InvalidFormat("invalid source font-face deletion".to_string()))?;
    let deletion = source.checked_range(source_range, deletion_bytes)?;
    let insertion = source.checked_range(source_insertion..source_insertion, b"")?;
    let mut publication = XmlSplicePublication::new(source);
    publication.replace(deletion, AuthoredXmlFragment::deletion())?;
    publication.replace(insertion, fragment)?;
    Ok(Some(publication))
}

fn remove_range(bytes: &[u8], range: Range<usize>) -> Result<Vec<u8>> {
    let prefix = bytes
        .get(..range.start)
        .ok_or_else(|| Error::InvalidFormat("invalid font-face prefix range".to_string()))?;
    let suffix = bytes
        .get(range.end..)
        .ok_or_else(|| Error::InvalidFormat("invalid font-face suffix range".to_string()))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(prefix.len().saturating_add(suffix.len()))
        .map_err(|source| Error::Allocation {
            resource: "font-face XML splice comparison",
            source,
        })?;
    output.extend_from_slice(prefix);
    output.extend_from_slice(suffix);
    Ok(output)
}

fn declaration_or_insertion_range(location: &Location) -> Range<usize> {
    location
        .target
        .as_ref()
        .map_or(location.insertion..location.insertion, |target| {
            target.start..target.end
        })
}

fn set_font_face_declarations_xml(
    xml: &str,
    declarations: &Declarations,
    part: Part,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    declarations.validate()?;
    let (old, location) = locate_font_face_declarations(xml, part, limits)?;
    let fragment = declarations.to_xml(limits)?;
    let updated = if let Some(target) = location.target {
        replace_span(xml, &target, &fragment, limits)?
    } else {
        insert_at(xml, location.insertion, &fragment, limits)?
    };
    parse_font_face_declarations_in_part(&updated, part, limits)?;
    Ok((updated, old))
}

fn remove_font_face_declarations_xml(
    xml: &str,
    part: Part,
    limits: &XmlStreamLimits,
) -> Result<(String, Option<Declarations>)> {
    let (old, location) = locate_font_face_declarations(xml, part, limits)?;
    let Some(target) = location.target else {
        return Ok((xml.to_owned(), old));
    };
    let updated = replace_span(xml, &target, "", limits)?;
    parse_font_face_declarations_in_part(&updated, part, limits)?;
    Ok((updated, old))
}

fn locate_font_face_declarations(
    xml: &str,
    part: Part,
    limits: &XmlStreamLimits,
) -> Result<(Option<Declarations>, Location)> {
    let declarations = parse(xml, limits)?;
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
                    let (_, start) = open_target.take().ok_or_else(|| {
                        Error::InvalidFormat("missing font-face target start".to_string())
                    })?;
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
    let root_open_end = root_open_end.ok_or_else(|| {
        Error::InvalidFormat("non-empty document root has no opening event".to_string())
    })?;
    let insertion = match part {
        Part::Content => scripts_end.unwrap_or(root_open_end),
        Part::Styles | Part::Flat => root_open_end,
    };
    Ok((declarations, Location { target, insertion }))
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml[..end]
        .rfind('<')
        .ok_or_else(|| Error::InvalidFormat("invalid ODF font-face XML event boundary".to_string()))
}

fn replace_span(
    xml: &str,
    span: &XmlSpan,
    replacement: &str,
    limits: &XmlStreamLimits,
) -> Result<String> {
    let removed = span
        .end
        .checked_sub(span.start)
        .ok_or_else(|| Error::InvalidFormat("invalid font-face replacement span".to_string()))?;
    let final_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or_else(|| Error::InvalidFormat("font-face output size overflow".to_string()))?;
    ensure_output_size(final_len, limits)?;
    let mut output = String::with_capacity(final_len);
    output.push_str(&xml[..span.start]);
    output.push_str(replacement);
    output.push_str(&xml[span.end..]);
    Ok(output)
}

fn insert_at(
    xml: &str,
    insertion: usize,
    fragment: &str,
    limits: &XmlStreamLimits,
) -> Result<String> {
    let final_len = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| Error::InvalidFormat("font-face output size overflow".to_string()))?;
    ensure_output_size(final_len, limits)?;
    let mut output = String::with_capacity(final_len);
    output.push_str(&xml[..insertion]);
    output.push_str(fragment);
    output.push_str(&xml[insertion..]);
    Ok(output)
}

fn ensure_output_size(size: usize, limits: &XmlStreamLimits) -> Result<()> {
    if u64::try_from(size).unwrap_or(u64::MAX) > limits.max_bytes() {
        return invalid("font-face XML output exceeds the configured size limit");
    }
    Ok(())
}

impl Package {
    /// Return content-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn content_font_face_declarations(&self) -> Result<Option<Declarations>> {
        let xml = self.content_xml()?;
        parse_content_font_face_declarations(&xml, &self.xml_limits_value())
    }

    /// Return styles-part font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn styles_font_face_declarations(&self) -> Result<Option<Declarations>> {
        self.styles_xml()?.map_or_else(
            || Ok(None),
            |xml| parse_styles_font_face_declarations(&xml, &self.xml_limits_value()),
        )
    }
}

impl FlatDocument {
    /// Return the flat document's font-face declarations.
    ///
    /// Font resource links are retained as inert metadata only. This method
    /// does not fetch a URI, load a font, or inspect embedded font data.
    pub fn font_face_declarations(&self) -> Result<Option<Declarations>> {
        parse_font_face_declarations_in_part(self.xml(), Part::Flat, &self.xml_limits_value())
    }
}
