//! Namespace-aware, range-preserving XML edits for the transaction root.

use std::collections::{BTreeSet, HashMap, HashSet};
use std::ops::Range;

use litchi_ooxml_common::xml::{DRAWINGML_NAMESPACE, STRICT_DRAWINGML_NAMESPACE};
use quick_xml::Reader;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, NamespaceResolver, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::{Slide, invalid};
use crate::{Error, Result};

const MAX_XML_DEPTH: usize = 256;
const MAX_XML_NODES: usize = 1_000_000;

pub(crate) fn compact_changed_slide_xml(source: &[u8]) -> Result<Vec<u8>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation compact slide XML",
            source,
        })?;
    let mut preserve_space = Vec::new();
    let mut pending_whitespace = Vec::new();
    let mut text_run_has_content = false;
    let mut roots = 0usize;
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                if preserve_space.is_empty() {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide root count overflow"))?;
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                preserve_space.push(element_preserves_space(
                    &element,
                    reader.decoder(),
                    inherited,
                )?);
                write_compact_start(&mut output, &element, reader.decoder(), false)?;
            },
            Event::Empty(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                if preserve_space.is_empty() {
                    roots = roots
                        .checked_add(1)
                        .ok_or_else(|| invalid("opened-presentation slide root count overflow"))?;
                }
                let inherited = preserve_space.last().copied().unwrap_or(false);
                let _preserve = element_preserves_space(&element, reader.decoder(), inherited)?;
                write_compact_start(&mut output, &element, reader.decoder(), true)?;
            },
            Event::End(element) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                preserve_space.pop().ok_or_else(|| {
                    invalid("opened-presentation slide has an unexpected end element")
                })?;
                output.extend_from_slice(b"</");
                output.extend_from_slice(element.name().as_ref());
                output.push(b'>');
            },
            Event::Text(text) => {
                let bytes = text.as_ref();
                if preserve_space.is_empty() {
                    if bytes.iter().all(u8::is_ascii_whitespace) {
                        continue;
                    }
                    return Err(invalid(
                        "opened-presentation slide has text outside its root",
                    ));
                }
                if bytes.iter().all(u8::is_ascii_whitespace)
                    && !preserve_space.last().copied().unwrap_or(false)
                {
                    if text_run_has_content {
                        output.extend_from_slice(bytes);
                    } else {
                        pending_whitespace.extend_from_slice(bytes);
                    }
                } else {
                    output.extend_from_slice(&pending_whitespace);
                    pending_whitespace.clear();
                    output.extend_from_slice(bytes);
                    text_run_has_content = true;
                }
            },
            Event::CData(data) => {
                if preserve_space.is_empty() {
                    return Err(invalid(
                        "opened-presentation slide has CDATA outside its root",
                    ));
                }
                output.extend_from_slice(&pending_whitespace);
                pending_whitespace.clear();
                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(data.as_ref());
                output.extend_from_slice(b"]]>");
                text_run_has_content = true;
            },
            Event::GeneralRef(reference) => {
                if preserve_space.is_empty() {
                    return Err(invalid(
                        "opened-presentation slide has an entity outside its root",
                    ));
                }
                output.extend_from_slice(&pending_whitespace);
                pending_whitespace.clear();
                output.push(b'&');
                output.extend_from_slice(reference.as_ref());
                output.push(b';');
                text_run_has_content = true;
            },
            Event::Decl(declaration) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.extend_from_slice(b"<?");
                output.extend_from_slice(declaration.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::PI(instruction) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.extend_from_slice(b"<?");
                output.extend_from_slice(instruction.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::Comment(comment) => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                output.extend_from_slice(b"<!--");
                output.extend_from_slice(comment.as_ref());
                output.extend_from_slice(b"-->");
            },
            Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation slide document types are not publishable",
                ));
            },
            Event::Eof => {
                finish_compact_text_run(&mut pending_whitespace, &mut text_run_has_content);
                break;
            },
        }
    }
    if !preserve_space.is_empty() || roots != 1 {
        return Err(invalid(
            "opened-presentation slide must contain exactly one closed root",
        ));
    }
    Ok(output)
}

fn finish_compact_text_run(pending_whitespace: &mut Vec<u8>, has_content: &mut bool) {
    pending_whitespace.clear();
    *has_content = false;
}

fn element_preserves_space(
    element: &BytesStart<'_>,
    decoder: Decoder,
    inherited: bool,
) -> Result<bool> {
    let mut preserve = inherited;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.as_ref() != b"xml:space" {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        preserve = match value.as_ref() {
            "preserve" => true,
            "default" => false,
            _ => {
                return Err(invalid(
                    "opened-presentation slide xml:space must be default or preserve",
                ));
            },
        };
    }
    Ok(preserve)
}

fn write_compact_start(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    empty: bool,
) -> Result<()> {
    output.push(b'<');
    output.extend_from_slice(element.name().as_ref());
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        output.push(b' ');
        output.extend_from_slice(attribute.key.as_ref());
        output.extend_from_slice(b"=\"");
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        output.extend_from_slice(quick_xml::escape::escape(value.as_ref()).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(())
}

#[derive(Debug)]
struct SlideIdElement {
    id: u32,
    relationship_id: String,
    span: Range<usize>,
    element_name: Option<Vec<u8>>,
    relationship_attribute_name: Option<Vec<u8>>,
}

#[derive(Debug)]
struct OpenElement {
    id: u32,
    relationship_id: String,
    start: usize,
    depth: usize,
}

#[derive(Debug)]
struct TextElement {
    span: Range<usize>,
    empty_name: Option<Vec<u8>>,
}

struct ShapeTextEdit<'a> {
    index: usize,
    shape: Range<usize>,
    text: &'a str,
    changed: bool,
}

pub(crate) fn reorder_slides(xml: &[u8], current: &[Slide], ordered: &[u32]) -> Result<Vec<u8>> {
    if ordered.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide order is not a complete permutation",
        ));
    }
    let elements = slide_id_elements(xml, false)?;
    if elements.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    for (element, slide) in elements.iter().zip(current) {
        if element.id != slide.id || element.relationship_id != slide.relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
    }
    let mut selected = Vec::new();
    selected
        .try_reserve_exact(ordered.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide permutation",
            source,
        })?;
    let mut seen = HashSet::new();
    for id in ordered {
        if !seen.insert(*id) {
            return Err(invalid(
                "opened-presentation slide order repeats an identity",
            ));
        }
        let index = elements
            .iter()
            .position(|element| element.id == *id)
            .ok_or_else(|| invalid("opened-presentation slide order references an unknown ID"))?;
        selected.push(index);
    }
    if selected
        .iter()
        .enumerate()
        .all(|(left, right)| left == *right)
    {
        return Ok(xml.to_vec());
    }
    let first = elements
        .first()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let last = elements
        .last()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation reordered XML",
            source,
        })?;
    output.extend_from_slice(&xml[..first.span.start]);
    for (position, source_index) in selected.into_iter().enumerate() {
        let source = &elements[source_index];
        output.extend_from_slice(&xml[source.span.clone()]);
        if let Some(next) = elements.get(position + 1) {
            output.extend_from_slice(&xml[elements[position].span.end..next.span.start]);
        }
    }
    output.extend_from_slice(&xml[last.span.end..]);
    Ok(output)
}

pub(crate) fn remove_slide(xml: &[u8], current: &[Slide], id: u32) -> Result<Vec<u8>> {
    let elements = slide_id_elements(xml, false)?;
    if elements.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    for (element, slide) in elements.iter().zip(current) {
        if element.id != slide.id || element.relationship_id != slide.relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
    }
    let target = elements
        .iter()
        .find(|element| element.id == id)
        .ok_or_else(|| invalid("opened-presentation slide removal identity is missing"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_sub(target.span.len()))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation slide removal XML",
            source,
        })?;
    output.extend_from_slice(&xml[..target.span.start]);
    output.extend_from_slice(&xml[target.span.end..]);
    Ok(output)
}

pub(crate) fn insert_slide(
    xml: &[u8],
    current: &[Slide],
    position: usize,
    id: u32,
    relationship_id: &str,
) -> Result<Vec<u8>> {
    if position > current.len() {
        return Err(invalid(
            "opened-presentation slide insertion position is out of bounds",
        ));
    }
    let elements = slide_id_elements(xml, true)?;
    if elements.len() != current.len() {
        return Err(invalid(
            "opened-presentation slide-order XML differs from the semantic graph",
        ));
    }
    for (element, slide) in elements.iter().zip(current) {
        if element.id != slide.id || element.relationship_id != slide.relationship_id {
            return Err(invalid(
                "opened-presentation slide-order binding changed during staging",
            ));
        }
    }
    let exemplar = elements
        .first()
        .ok_or_else(|| invalid("opened-presentation slide list is empty"))?;
    let element_name = exemplar
        .element_name
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation slide entry name is missing"))?;
    let relationship_attribute_name = exemplar
        .relationship_attribute_name
        .as_deref()
        .ok_or_else(|| invalid("opened-presentation relationship attribute name is missing"))?;
    let insertion = elements
        .get(position)
        .map_or_else(
            || elements.last().map(|element| element.span.end),
            |element| Some(element.span.start),
        )
        .ok_or_else(|| invalid("opened-presentation slide insertion point is missing"))?;
    let mut fragment = Vec::new();
    let capacity = element_name
        .len()
        .checked_add(relationship_attribute_name.len())
        .and_then(|value| value.checked_add(64))
        .ok_or_else(|| invalid("opened-presentation slide entry size overflow"))?;
    fragment
        .try_reserve_exact(capacity)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation inserted slide entry",
            source,
        })?;
    fragment.push(b'<');
    fragment.extend_from_slice(element_name);
    fragment.extend_from_slice(b" id=\"");
    fragment.extend_from_slice(id.to_string().as_bytes());
    fragment.extend_from_slice(b"\" ");
    fragment.extend_from_slice(relationship_attribute_name);
    fragment.extend_from_slice(b"=\"");
    fragment.extend_from_slice(relationship_id.as_bytes());
    fragment.extend_from_slice(b"\"/>");
    let output_len = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| invalid("opened-presentation inserted XML size overflow"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation inserted slide XML",
            source,
        })?;
    output.extend_from_slice(&xml[..insertion]);
    output.extend_from_slice(&fragment);
    output.extend_from_slice(&xml[insertion..]);
    Ok(output)
}

pub(crate) fn rewrite_shape_text(
    xml: &[u8],
    shape: Range<usize>,
    text: &str,
    max_text_bytes: usize,
) -> Result<Vec<u8>> {
    let edits = [ShapeTextEdit {
        index: 0,
        shape,
        text,
        changed: true,
    }];
    rewrite_shape_texts(
        xml,
        &edits,
        max_text_bytes,
        crate::shape::Limits::DEFAULT.output_bytes(),
    )?
    .ok_or_else(|| invalid("opened-presentation shape text rewrite disappeared"))
}

pub(crate) fn stage_shape_texts(
    xml: &[u8],
    replacements: &[super::ShapeTextReplacement<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<(Option<Vec<u8>>, usize)> {
    if replacements.is_empty() {
        return Ok((None, 0));
    }
    if replacements.len() > super::MAX_SHAPE_TEXT_REPLACEMENTS {
        return Err(Error::Limit {
            resource: "opened-presentation shape text replacements",
            limit: super::MAX_SHAPE_TEXT_REPLACEMENTS,
        });
    }
    let aggregate_text_bytes = replacements.iter().try_fold(0usize, |total, replacement| {
        total
            .checked_add(replacement.text().len())
            .ok_or_else(|| invalid("opened-presentation aggregate shape text size overflow"))
    })?;
    if aggregate_text_bytes > max_text_bytes {
        return Err(Error::Limit {
            resource: "opened-presentation aggregate shape text bytes",
            limit: max_text_bytes,
        });
    }

    let scene = crate::shape::Scene::read(xml)?;
    if scene.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "set_shape_texts",
            reason: "atomic shape text batches do not support markup-compatibility branch selection",
        });
    }
    if replacements.len() > scene.len() {
        return Err(Error::Limit {
            resource: "opened-presentation selected shape texts",
            limit: scene.len(),
        });
    }
    let mut selected = HashSet::new();
    selected
        .try_reserve(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text identities",
            source,
        })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(replacements.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text edit plan",
            source,
        })?;
    for replacement in replacements {
        let shape = scene.shape(replacement.selector())?;
        let index = shape.common().index();
        if !selected.insert(index) {
            return Err(Error::DuplicateShapeTextSelection { index });
        }
        let current = shape
            .common()
            .text()
            .ok_or_else(|| invalid("opened-presentation selected shape has no text body"))?;
        edits.push(ShapeTextEdit {
            index,
            shape: shape.common().span()?.range(xml.len())?,
            text: replacement.text(),
            changed: current != replacement.text(),
        });
    }
    edits.sort_unstable_by_key(|edit| (edit.shape.start, edit.shape.end));
    for pair in edits.windows(2) {
        if pair[0].shape.end > pair[1].shape.start {
            return Err(invalid("opened-presentation shape text selections overlap"));
        }
    }
    let changed = edits.iter().filter(|edit| edit.changed).count();
    let staged = rewrite_shape_texts(xml, &edits, max_text_bytes, max_output_bytes)?;
    let Some(candidate) = staged.as_deref() else {
        return Ok((None, 0));
    };
    let published = crate::shape::Scene::read(candidate)?;
    if published.is_rewritten() || published.len() != scene.len() {
        return Err(invalid(
            "opened-presentation shape text batch changed the semantic shape graph",
        ));
    }
    for edit in &edits {
        if published.at(edit.index)?.common().text() != Some(edit.text) {
            return Err(invalid(
                "opened-presentation shape text batch did not round-trip semantically",
            ));
        }
    }
    Ok((staged, changed))
}

fn rewrite_shape_texts(
    xml: &[u8],
    edits: &[ShapeTextEdit<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<Option<Vec<u8>>> {
    for edit in edits {
        if edit.text.len() > max_text_bytes {
            return Err(Error::Limit {
                resource: "opened-presentation shape text bytes",
                limit: max_text_bytes,
            });
        }
        if !edit.text.chars().all(is_xml_char) {
            return Err(invalid(
                "opened-presentation shape text contains an invalid XML character",
            ));
        }
        if edit.shape.start >= edit.shape.end || edit.shape.end > xml.len() {
            return Err(invalid("opened-presentation shape range is invalid"));
        }
    }
    let mut owners = Vec::new();
    owners
        .try_reserve_exact(edits.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text owner ranges",
            source,
        })?;
    owners.extend(edits.iter().map(|edit| edit.shape.clone()));
    let elements = drawing_text_elements_for_owners(xml, &owners)?;
    if elements.iter().any(Vec::is_empty) {
        return Err(invalid(
            "opened-presentation selected shape has no DrawingML text run",
        ));
    }
    if !edits.iter().any(|edit| edit.changed) {
        return Ok(None);
    }
    let mut escaped = Vec::new();
    escaped
        .try_reserve_exact(edits.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation escaped shape texts",
            source,
        })?;
    escaped.extend(
        edits
            .iter()
            .map(|edit| quick_xml::escape::escape(edit.text)),
    );
    let mut removed = 0usize;
    let mut emitted = 0usize;
    for ((edit, spans), replacement) in edits.iter().zip(&elements).zip(&escaped) {
        if !edit.changed {
            continue;
        }
        for (position, span) in spans.iter().enumerate() {
            removed = removed
                .checked_add(span.span.len())
                .ok_or_else(|| invalid("opened-presentation shape text size overflow"))?;
            emitted = emitted
                .checked_add(text_element_output_len(
                    xml,
                    span,
                    position,
                    replacement.len(),
                )?)
                .ok_or_else(|| invalid("opened-presentation shape text size overflow"))?;
        }
    }
    let output_len = xml
        .len()
        .checked_sub(removed)
        .and_then(|length| length.checked_add(emitted))
        .ok_or_else(|| invalid("opened-presentation shape text output size overflow"))?;
    if output_len > max_output_bytes {
        return Err(Error::Limit {
            resource: "opened-presentation shape text output bytes",
            limit: max_output_bytes,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text XML",
            source,
        })?;
    let mut cursor = 0usize;
    for ((edit, spans), replacement) in edits.iter().zip(&elements).zip(&escaped) {
        if !edit.changed {
            continue;
        }
        for (position, span) in spans.iter().enumerate() {
            output.extend_from_slice(&xml[cursor..span.span.start]);
            write_text_element(&mut output, xml, span, position, replacement.as_bytes())?;
            cursor = span.span.end;
        }
    }
    output.extend_from_slice(&xml[cursor..]);
    if output.len() != output_len {
        return Err(invalid(
            "opened-presentation shape text output length changed during emission",
        ));
    }
    Ok(Some(output))
}

fn text_element_output_len(
    xml: &[u8],
    element: &TextElement,
    position: usize,
    escaped_len: usize,
) -> Result<usize> {
    if position != 0 {
        return Ok(element
            .empty_name
            .as_ref()
            .map_or(0, |_| element.span.len()));
    }
    let Some(name) = &element.empty_name else {
        return Ok(escaped_len);
    };
    let open_end = empty_text_open_end(&xml[element.span.clone()])?;
    open_end
        .checked_add(escaped_len)
        .and_then(|length| length.checked_add(name.len()))
        .and_then(|length| length.checked_add(4))
        .ok_or_else(|| invalid("opened-presentation empty text expansion size overflow"))
}

fn write_text_element(
    output: &mut Vec<u8>,
    xml: &[u8],
    element: &TextElement,
    position: usize,
    escaped: &[u8],
) -> Result<()> {
    if position != 0 {
        if element.empty_name.is_some() {
            output.extend_from_slice(&xml[element.span.clone()]);
        }
        return Ok(());
    }
    let Some(name) = &element.empty_name else {
        output.extend_from_slice(escaped);
        return Ok(());
    };
    let raw = &xml[element.span.clone()];
    let open_end = empty_text_open_end(raw)?;
    output.extend_from_slice(&raw[..open_end]);
    output.push(b'>');
    output.extend_from_slice(escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(name);
    output.push(b'>');
    Ok(())
}

fn empty_text_open_end(raw: &[u8]) -> Result<usize> {
    let slash = raw
        .iter()
        .rposition(|byte| *byte == b'/')
        .ok_or_else(|| invalid("opened-presentation empty text tag is malformed"))?;
    let mut open_end = slash;
    while open_end > 0 && raw[open_end - 1].is_ascii_whitespace() {
        open_end -= 1;
    }
    Ok(open_end)
}

pub(crate) fn append_shape(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if fragment.is_empty() {
        return Err(invalid(
            "opened-presentation shape fragment cannot be empty",
        ));
    }
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut tree_depth = None;
    let mut trees = 0usize;
    let mut insertion = None;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml_namespace = is_presentation_namespace(&namespace);
        let event = event.into_owned();
        drop(namespace);
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "opened-presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                depth += 1;
                if pml_namespace && element.local_name().as_ref() == b"spTree" {
                    trees = trees.saturating_add(1);
                    if trees != 1 || tree_depth.replace(depth).is_some() {
                        return Err(invalid(
                            "opened-presentation slide has multiple shape trees",
                        ));
                    }
                }
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if pml_namespace && element.local_name().as_ref() == b"spTree" {
                    return Err(invalid(
                        "opened-presentation cannot append to an empty shape-tree element",
                    ));
                }
            },
            Event::End(element) => {
                if tree_depth == Some(depth)
                    && pml_namespace
                    && element.local_name().as_ref() == b"spTree"
                {
                    insertion = Some(start);
                    tree_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || tree_depth.is_some() {
        return Err(invalid("opened-presentation slide XML is unterminated"));
    }
    if trees != 1 {
        return Err(invalid(
            "opened-presentation slide must have one shape tree",
        ));
    }
    let insertion =
        insertion.ok_or_else(|| invalid("opened-presentation slide has no shape tree"))?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(xml.len().saturating_add(fragment.len()))
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation appended shape XML",
            source,
        })?;
    output.extend_from_slice(&xml[..insertion]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[insertion..]);
    Ok(output)
}

pub(crate) fn remap_shape_fragment(
    source: &[u8],
    root_source_id: u32,
    root_name: &str,
    shape_ids: &HashMap<u32, u32>,
    relationships: &HashMap<String, String>,
) -> Result<Vec<u8>> {
    let mut reader = NsReader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut output = Vec::with_capacity(source.len());
    let mut depth = 0usize;
    let mut roots = 0usize;
    let mut remap = ShapeRemap {
        root_source_id,
        root_name,
        shape_ids,
        relationships,
        written_shape_ids: HashSet::new(),
    };
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) => {
                if depth == 0 {
                    roots = roots.saturating_add(1);
                }
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("opened-presentation shape depth overflow"))?;
                write_remapped_shape_start(
                    &mut output,
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    false,
                    &mut remap,
                )?;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    roots = roots.saturating_add(1);
                }
                write_remapped_shape_start(
                    &mut output,
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    true,
                    &mut remap,
                )?;
            },
            Event::End(element) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation shape depth underflow"))?;
                output.extend_from_slice(b"</");
                output.extend_from_slice(element.name().as_ref());
                output.push(b'>');
            },
            Event::Text(text) => output.extend_from_slice(text.as_ref()),
            Event::CData(data) => {
                output.extend_from_slice(b"<![CDATA[");
                output.extend_from_slice(data.as_ref());
                output.extend_from_slice(b"]]");
                output.push(b'>');
            },
            Event::GeneralRef(reference) => {
                output.push(b'&');
                output.extend_from_slice(reference.as_ref());
                output.push(b';');
            },
            Event::Comment(comment) => {
                output.extend_from_slice(b"<!--");
                output.extend_from_slice(comment.as_ref());
                output.extend_from_slice(b"-->");
            },
            Event::PI(instruction) => {
                output.extend_from_slice(b"<?");
                output.extend_from_slice(instruction.as_ref());
                output.extend_from_slice(b"?>");
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape fragment contains document-level markup",
                ));
            },
            Event::Eof => break,
        }
    }
    if depth != 0 || roots != 1 || !remap.written_shape_ids.contains(&root_source_id) {
        return Err(invalid(
            "opened-presentation transferred shape must have one root and a complete identity map",
        ));
    }
    Ok(output)
}

pub(crate) fn connector_connection_ids(source: &[u8]) -> Result<BTreeSet<u32>> {
    let mut reader = Reader::from_reader(source);
    reader.config_mut().trim_text(false);
    let mut connections = BTreeSet::new();
    loop {
        match reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
        {
            Event::Start(element) | Event::Empty(element)
                if matches!(element.local_name().as_ref(), b"stCxn" | b"endCxn") =>
            {
                let mut identity = None;
                for attribute in element.attributes() {
                    let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
                    if attribute.key.as_ref() != b"id" {
                        continue;
                    }
                    if identity.is_some() {
                        return Err(invalid(
                            "opened-presentation connector endpoint repeats its identity attribute",
                        ));
                    }
                    let value = attribute
                        .decoded_and_normalized_value(
                            quick_xml::XmlVersion::Implicit1_0,
                            reader.decoder(),
                        )
                        .map_err(|error| Error::Xml(error.to_string()))?;
                    identity = Some(value.parse::<u32>().map_err(|_err| {
                        invalid("opened-presentation connector endpoint identity is invalid")
                    })?);
                }
                connections.insert(identity.ok_or_else(|| {
                    invalid("opened-presentation connector endpoint identity is missing")
                })?);
            },
            Event::Decl(_) | Event::DocType(_) => {
                return Err(invalid(
                    "opened-presentation shape fragment contains document-level markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    Ok(connections)
}

struct ShapeRemap<'a> {
    root_source_id: u32,
    root_name: &'a str,
    shape_ids: &'a HashMap<u32, u32>,
    relationships: &'a HashMap<String, String>,
    written_shape_ids: HashSet<u32>,
}

fn write_remapped_shape_start(
    output: &mut Vec<u8>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    empty: bool,
    remap: &mut ShapeRemap<'_>,
) -> Result<()> {
    let is_identity = element.local_name().as_ref() == b"cNvPr";
    let is_connection = matches!(element.local_name().as_ref(), b"stCxn" | b"endCxn");
    let mut attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let decoded = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        attributes.push((attribute.key.as_ref().to_vec(), decoded.into_owned()));
    }
    let source_identity = if is_identity {
        shape_identity_from_attributes(&attributes)?
    } else {
        None
    };
    let mapped_identity = source_identity
        .map(|source_id| {
            remap
                .shape_ids
                .get(&source_id)
                .copied()
                .map(|destination_id| (source_id, destination_id))
                .ok_or_else(|| {
                    invalid(format!(
                        "opened-presentation transferred shape identity {source_id} was not remapped"
                    ))
                })
        })
        .transpose()?
        .map(|(source_id, destination_id)| {
            if !remap.written_shape_ids.insert(source_id) {
                return Err(invalid(format!(
                    "opened-presentation transferred shape repeats identity {source_id}"
                )));
            }
            Ok((source_id, destination_id))
        })
        .transpose()?;
    output.push(b'<');
    output.extend_from_slice(element.name().as_ref());
    for (key, decoded) in attributes {
        output.push(b' ');
        output.extend_from_slice(&key);
        output.extend_from_slice(b"=\"");
        let value = if let Some((_source_id, destination_id)) = mapped_identity
            && key == b"id"
        {
            destination_id.to_string()
        } else if let Some((source_id, destination_id)) = mapped_identity
            && key == b"name"
        {
            if source_id == remap.root_source_id {
                remap.root_name.to_owned()
            } else {
                format!("{decoded} Copy {destination_id}")
            }
        } else if is_connection && key == b"id" {
            let connected_id = decoded.parse::<u32>().map_err(|_err| {
                invalid("opened-presentation connector endpoint identity is invalid")
            })?;
            remap
                .shape_ids
                .get(&connected_id)
                .ok_or_else(|| {
                    Error::ShapeTransfer {
                        kind: crate::ShapeTransferRefusal::UnresolvedConnectorEndpoint,
                        detail: format!(
                            "connector endpoint {connected_id} lies outside the planned transfer closure"
                        ),
                    }
                })?
                .to_string()
        } else {
            let (namespace, _) = resolver.resolve_attribute(quick_xml::name::QName(key.as_slice()));
            let is_relationship_namespace = matches!(
                &namespace,
                ResolveResult::Bound(Namespace(value))
                    if *value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                        || *value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
            );
            let is_inherited_relationship_alias = matches!(
                &namespace,
                ResolveResult::Unknown(prefix)
                    if prefix.as_slice() == b"r"
                        || remap.relationships.contains_key(decoded.as_str())
            );
            if is_relationship_namespace || is_inherited_relationship_alias {
                remap
                    .relationships
                    .get(decoded.as_str())
                    .cloned()
                    .ok_or_else(|| {
                        invalid(format!(
                            "opened-presentation transferred relationship {decoded} was not remapped"
                        ))
                    })?
            } else {
                decoded
            }
        };
        output.extend_from_slice(quick_xml::escape::escape(&value).as_bytes());
        output.push(b'"');
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(())
}

fn shape_identity_from_attributes(attributes: &[(Vec<u8>, String)]) -> Result<Option<u32>> {
    let mut identities = attributes
        .iter()
        .filter(|(key, _)| key.as_slice() == b"id")
        .map(|(_, value)| value);
    let Some(value) = identities.next() else {
        return Ok(None);
    };
    if identities.next().is_some() {
        return Err(invalid(
            "opened-presentation non-visual identity has duplicate id attributes",
        ));
    }
    value
        .parse::<u32>()
        .map(Some)
        .map_err(|_err| invalid("opened-presentation non-visual shape identity is invalid"))
}

fn slide_id_elements(xml: &[u8], capture_names: bool) -> Result<Vec<SlideIdElement>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut list_depth = None;
    let mut lists = 0usize;
    let mut open = None;
    let mut elements = Vec::new();
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let pml_namespace = is_presentation_namespace(&namespace);
        let event = event.into_owned();
        drop(namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "opened-presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                if pml_namespace && element.local_name().as_ref() == b"sldIdLst" {
                    lists = lists.saturating_add(1);
                    if lists != 1 {
                        return Err(invalid("opened-presentation has multiple slide-ID lists"));
                    }
                    list_depth = Some(depth + 1);
                } else if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                    if capture_names {
                        return Err(invalid(
                            "slide-copy requires empty presentation slide-ID entries",
                        ));
                    }
                    if open.is_some() {
                        return Err(invalid("opened-presentation slide IDs overlap"));
                    }
                    let (id, relationship_id, _relationship_attribute_name) =
                        parse_slide_id(&element, &reader, false)?;
                    open = Some(OpenElement {
                        id,
                        relationship_id,
                        start,
                        depth: depth + 1,
                    });
                }
                depth += 1;
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                if list_depth == Some(depth) {
                    if !pml_namespace || element.local_name().as_ref() != b"sldId" {
                        return Err(invalid(
                            "opened-presentation slide-ID list has an unsupported child",
                        ));
                    }
                    let capture_element_names = capture_names && elements.is_empty();
                    let (id, relationship_id, relationship_attribute_name) =
                        parse_slide_id(&element, &reader, capture_element_names)?;
                    elements.push(SlideIdElement {
                        id,
                        relationship_id,
                        span: start..end,
                        element_name: capture_element_names
                            .then(|| element.name().as_ref().to_vec()),
                        relationship_attribute_name,
                    });
                }
            },
            Event::End(element) => {
                if let Some(active) = &open
                    && active.depth == depth
                    && pml_namespace
                    && element.local_name().as_ref() == b"sldId"
                {
                    let active = open
                        .take()
                        .ok_or_else(|| invalid("opened-presentation slide ID disappeared"))?;
                    elements.push(SlideIdElement {
                        id: active.id,
                        relationship_id: active.relationship_id,
                        span: active.start..end,
                        element_name: None,
                        relationship_attribute_name: None,
                    });
                }
                if list_depth == Some(depth)
                    && pml_namespace
                    && element.local_name().as_ref() == b"sldIdLst"
                {
                    list_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(value) if list_depth == Some(depth) => {
                if !value
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return Err(invalid("opened-presentation slide-ID list contains text"));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if list_depth == Some(depth) => {
                return Err(invalid(
                    "opened-presentation slide-ID list contains unsupported content",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || open.is_some() || list_depth.is_some() {
        return Err(invalid("opened-presentation XML is unterminated"));
    }
    if lists != 1 {
        return Err(invalid("opened-presentation has no slide-ID list"));
    }
    Ok(elements)
}

fn parse_slide_id(
    element: &BytesStart<'_>,
    reader: &NsReader<&[u8]>,
    capture_relationship_name: bool,
) -> Result<(u32, String, Option<Vec<u8>>)> {
    let value =
        litchi_ooxml_common::xml::unqualified_attribute_value(element, b"id", reader.decoder())?
            .ok_or_else(|| invalid("opened-presentation slide ID lacks id"))?;
    let id = value
        .parse::<u32>()
        .map_err(|_err| invalid("opened-presentation slide ID is invalid"))?;
    let relationship_id = crate::parts::relationship_attribute(element, reader)?
        .ok_or_else(|| invalid("opened-presentation slide ID lacks r:id"))?;
    let mut relationship_attribute_name = None;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        if attribute.key.local_name().as_ref() != b"id" {
            continue;
        }
        if capture_relationship_name
            && matches!(
                reader.resolver().resolve_attribute(attribute.key).0,
                ResolveResult::Bound(Namespace(value))
                    if value == b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        || value == b"http://purl.oclc.org/ooxml/officeDocument/relationships"
            )
        {
            if relationship_attribute_name
                .replace(attribute.key.as_ref().to_vec())
                .is_some()
            {
                return Err(invalid(
                    "opened-presentation slide ID repeats its relationship attribute",
                ));
            }
        }
    }
    if capture_relationship_name && relationship_attribute_name.is_none() {
        return Err(invalid(
            "opened-presentation slide ID relationship name is unresolved",
        ));
    }
    Ok((id, relationship_id, relationship_attribute_name))
}

fn drawing_text_elements_for_owners(
    xml: &[u8],
    owners: &[Range<usize>],
) -> Result<Vec<Vec<TextElement>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut spans = Vec::new();
    spans
        .try_reserve_exact(owners.len())
        .map_err(|source| Error::Allocation {
            resource: "opened-presentation shape text span groups",
            source,
        })?;
    spans.resize_with(owners.len(), Vec::new);
    let mut active: Option<(usize, usize, usize)> = None;
    let mut owner_position = 0usize;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let drawing_namespace = is_drawing(&namespace);
        let event = event.into_owned();
        drop(namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(element) => {
                bump(&mut nodes)?;
                if depth >= MAX_XML_DEPTH {
                    return Err(Error::Limit {
                        resource: "opened-presentation XML depth",
                        limit: MAX_XML_DEPTH,
                    });
                }
                depth += 1;
                while owners
                    .get(owner_position)
                    .is_some_and(|owner| start >= owner.end)
                {
                    owner_position += 1;
                }
                let owner = owners
                    .get(owner_position)
                    .filter(|owner| owner.contains(&start));
                if owner.is_some() && drawing_namespace && element.local_name().as_ref() == b"t" {
                    if active.replace((end, depth, owner_position)).is_some() {
                        return Err(invalid(
                            "opened-presentation DrawingML text elements overlap",
                        ));
                    }
                } else if active.is_some() {
                    return Err(invalid(
                        "opened-presentation DrawingML text contains child markup",
                    ));
                }
            },
            Event::Empty(element) => {
                bump(&mut nodes)?;
                while owners
                    .get(owner_position)
                    .is_some_and(|owner| start >= owner.end)
                {
                    owner_position += 1;
                }
                let owner = owners
                    .get(owner_position)
                    .filter(|owner| owner.contains(&start));
                if owner.is_some() && drawing_namespace && element.local_name().as_ref() == b"t" {
                    let group = spans.get_mut(owner_position).ok_or_else(|| {
                        invalid("opened-presentation shape text span owner disappeared")
                    })?;
                    group.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "opened-presentation shape text spans",
                        source,
                    })?;
                    group.push(TextElement {
                        span: start..end,
                        empty_name: Some(element.name().as_ref().to_vec()),
                    });
                } else if active.is_some() {
                    return Err(invalid(
                        "opened-presentation DrawingML text contains child markup",
                    ));
                }
            },
            Event::End(element) => {
                if let Some((content_start, active_depth, owner)) = active
                    && active_depth == depth
                    && drawing_namespace
                    && element.local_name().as_ref() == b"t"
                {
                    let group = spans.get_mut(owner).ok_or_else(|| {
                        invalid("opened-presentation shape text span owner disappeared")
                    })?;
                    group.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "opened-presentation shape text spans",
                        source,
                    })?;
                    group.push(TextElement {
                        span: content_start..start,
                        empty_name: None,
                    });
                    active = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("opened-presentation XML depth underflow"))?;
            },
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Comment(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains a comment",
                ));
            },
            Event::Decl(_) | Event::PI(_) | Event::DocType(_) if active.is_some() => {
                return Err(invalid(
                    "opened-presentation DrawingML text contains forbidden markup",
                ));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || active.is_some() {
        return Err(invalid("opened-presentation slide XML is unterminated"));
    }
    Ok(spans)
}

fn is_drawing(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE
    )
}

fn is_presentation_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == crate::namespace::PRESENTATIONML_NAMESPACE
                || *value == crate::namespace::STRICT_PRESENTATIONML_NAMESPACE
    )
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_err| invalid("opened-presentation XML position exceeds usize"))
}

fn bump(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| invalid("opened-presentation XML node count overflow"))?;
    if *nodes > MAX_XML_NODES {
        return Err(Error::Limit {
            resource: "opened-presentation XML nodes",
            limit: MAX_XML_NODES,
        });
    }
    Ok(())
}

fn is_xml_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x1_0000..=0x10_FFFF)
}
