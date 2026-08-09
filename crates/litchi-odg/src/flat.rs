//! Lossless flat `OpenDocument` Drawing snapshots and bounded text edits.

use litchi_core::{Error, FileFormat, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::NsReader,
};
use std::{ops::Range, sync::Arc};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_BYTES: usize = 256 * 1024 * 1024;
const MAX_DEPTH: usize = 256;
const MAX_PAGES: usize = 16_384;
const MAX_SHAPES: usize = 1_000_000;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Office,
    Draw,
    Text,
    Other,
}

#[derive(Clone, Debug)]
struct State {
    bytes: Vec<u8>,
    pages: Vec<FlatPage>,
}

/// An immutable, byte-preserving flat ODG snapshot.
#[derive(Clone, Debug)]
pub struct FlatDrawing(Arc<State>);

impl FlatDrawing {
    /// Opens a flat ODG document from owned bytes.
    ///
    /// # Errors
    ///
    /// Returns an error when the bytes are not a bounded, structurally valid FODG document.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        parse(bytes).map(|state| Self(Arc::new(state)))
    }

    /// Returns the parsed drawing pages.
    #[must_use]
    pub fn pages(&self) -> &[FlatPage] {
        &self.0.pages
    }

    /// Returns the original flat XML bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0.bytes
    }

    /// Starts a detached edit without mutating this snapshot.
    #[must_use]
    pub fn edit(&self) -> FlatDrawingEdit {
        FlatDrawingEdit {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    /// Consumes the snapshot and returns its exact flat XML bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }
}

/// One drawing page in a flat ODG snapshot.
#[derive(Clone, Debug)]
pub struct FlatPage {
    name: Option<String>,
    shapes: Vec<FlatShape>,
}

impl FlatPage {
    /// Returns the optional `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns all bounded drawing shapes found below this page.
    #[must_use]
    pub fn shapes(&self) -> &[FlatShape] {
        &self.shapes
    }
}

/// One drawing shape and its extracted paragraph text.
#[derive(Clone, Debug)]
pub struct FlatShape {
    name: Option<String>,
    text: String,
    text_spans: Vec<Range<usize>>,
}

impl FlatShape {
    /// Returns the optional `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns text from `text:p` descendants.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A staged, source-bound flat drawing edit.
pub struct FlatDrawingEdit {
    source: FlatDrawing,
    changes: Vec<TextChange>,
}

impl FlatDrawingEdit {
    /// Stages replacement of a shape whose text occupies one exact XML text span.
    ///
    /// Shapes with empty, split, or mixed-content text are refused rather than
    /// being serialized through a lossy generic XML writer.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector, unsupported source span, or limit violation.
    pub fn set_shape_text(
        &mut self,
        page: usize,
        shape: usize,
        text: impl Into<String>,
    ) -> Result<()> {
        let replacement = text.into();
        if replacement.len() > MAX_TEXT_BYTES {
            return Err(invalid("replacement shape text exceeds the flat ODG limit"));
        }
        let selected = self
            .source
            .pages()
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or_else(|| invalid("flat ODG shape selector is out of bounds"))?;
        if selected.text_spans.len() != 1 {
            return Err(invalid(
                "flat ODG shape text is not one losslessly replaceable XML span",
            ));
        }
        if selected.text == replacement {
            self.changes
                .retain(|change| change.page != page || change.shape != shape);
            return Ok(());
        }
        let change = TextChange {
            page,
            shape,
            before: selected.text.clone(),
            after: replacement,
        };
        if let Some(existing) = self
            .changes
            .iter_mut()
            .find(|existing| existing.page == page && existing.shape == shape)
        {
            *existing = change;
        } else {
            self.changes.push(change);
        }
        Ok(())
    }

    /// Validates and atomically publishes a new immutable snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when staged spans overlap, output parsing fails, or readback differs.
    pub fn commit(self) -> Result<FlatDrawingCommit> {
        let FlatDrawingEdit { source, changes } = self;
        let mut replacements = Vec::with_capacity(changes.len());
        for change in &changes {
            let shape = &source.pages()[change.page].shapes()[change.shape];
            replacements.push((
                shape.text_spans[0].clone(),
                quick_xml::escape::escape(&change.after).into_owned(),
            ));
        }
        replacements.sort_unstable_by_key(|replacement| std::cmp::Reverse(replacement.0.start));
        let mut bytes = source.as_bytes().to_vec();
        let mut previous_start = bytes.len();
        for (span, replacement) in replacements {
            if span.end > previous_start || span.start > span.end || span.end > bytes.len() {
                return Err(invalid("overlapping or invalid flat ODG text patch"));
            }
            bytes.splice(span.clone(), replacement.bytes());
            previous_start = span.start;
        }
        let snapshot = FlatDrawing::from_bytes(bytes)?;
        for change in &changes {
            let actual = snapshot.pages()[change.page].shapes()[change.shape].text();
            if actual != change.after {
                return Err(invalid("flat ODG edit failed typed readback"));
            }
        }
        Ok(FlatDrawingCommit {
            snapshot: snapshot.clone(),
            patch: FlatDrawingPatch {
                source,
                target: snapshot,
                changes,
            },
        })
    }
}

/// A committed flat drawing snapshot and its reversible semantic patch.
pub struct FlatDrawingCommit {
    snapshot: FlatDrawing,
    patch: FlatDrawingPatch,
}

impl FlatDrawingCommit {
    /// Returns the published snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &FlatDrawing {
        &self.snapshot
    }

    /// Returns the semantic patch.
    #[must_use]
    pub fn patch(&self) -> &FlatDrawingPatch {
        &self.patch
    }

    /// Consumes the commit and returns the published snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> FlatDrawing {
        self.snapshot
    }
}

/// A source-checked, reversible list of shape-text changes.
#[derive(Clone, Debug)]
pub struct FlatDrawingPatch {
    source: FlatDrawing,
    target: FlatDrawing,
    changes: Vec<TextChange>,
}

impl FlatDrawingPatch {
    /// Returns whether this patch was committed from the exact supplied bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &FlatDrawing) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not the exact source artifact.
    pub fn apply(&self, source: &FlatDrawing) -> Result<FlatDrawing> {
        if !self.is_applicable_to(source) {
            return Err(invalid("flat ODG patch source does not match"));
        }
        Ok(self.target.clone())
    }

    /// Returns the ordered semantic changes.
    #[must_use]
    pub fn changes(&self) -> &[TextChange] {
        &self.changes
    }

    /// Returns a patch applicable to this patch's exact target snapshot.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self
                .changes
                .iter()
                .map(|change| TextChange {
                    page: change.page,
                    shape: change.shape,
                    before: change.after.clone(),
                    after: change.before.clone(),
                })
                .collect(),
        }
    }
}

/// One selector-bound reversible shape-text change.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextChange {
    page: usize,
    shape: usize,
    before: String,
    after: String,
}

impl TextChange {
    #[must_use]
    pub fn page(&self) -> usize {
        self.page
    }

    #[must_use]
    pub fn shape(&self) -> usize {
        self.shape
    }

    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }
}

fn parse(bytes: Vec<u8>) -> Result<State> {
    if bytes.len() > MAX_BYTES {
        return Err(invalid("flat ODG exceeds the input size limit"));
    }
    if litchi_odf_common::detect::flat(&bytes) != Some(FileFormat::Odg) {
        return Err(invalid("input is not a flat ODG document"));
    }
    let xml = std::str::from_utf8(&bytes).map_err(|_error| invalid("flat ODG is not UTF-8"))?;
    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut body_seen = false;
    let mut drawing_seen = false;
    let mut body_depth = None;
    let mut drawing_depth = None;
    let mut page_open: Option<(usize, usize)> = None;
    let mut active_shapes = Vec::<(usize, usize, usize)>::new();
    let mut paragraph_depths = Vec::<usize>::new();
    let mut pages = Vec::<FlatPage>::new();
    let mut shape_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let event_start = position(&reader)?;
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid flat ODG XML: {error}")))?;
        let namespace = classify(&resolved_namespace);
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth)?;
                observe_start(
                    &reader,
                    namespace,
                    &element,
                    depth,
                    false,
                    &mut root_seen,
                    &mut body_seen,
                    &mut drawing_seen,
                    &mut body_depth,
                    &mut drawing_depth,
                    &mut page_open,
                    &mut active_shapes,
                    &mut paragraph_depths,
                    &mut pages,
                    &mut shape_count,
                )?;
            },
            Event::Empty(element) => {
                let virtual_depth = checked_depth(depth)?;
                observe_start(
                    &reader,
                    namespace,
                    &element,
                    virtual_depth,
                    true,
                    &mut root_seen,
                    &mut body_seen,
                    &mut drawing_seen,
                    &mut body_depth,
                    &mut drawing_depth,
                    &mut page_open,
                    &mut active_shapes,
                    &mut paragraph_depths,
                    &mut pages,
                    &mut shape_count,
                )?;
            },
            Event::Text(text) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                let decoded = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid flat ODG text: {error}")))?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| invalid(format!("invalid flat ODG text escape: {error}")))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("flat ODG text size overflow"))?;
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(invalid("flat ODG text exceeds the extraction limit"));
                }
                pages[page].shapes[shape].text.push_str(&value);
                pages[page].shapes[shape]
                    .text_spans
                    .push(event_start..event_end);
            },
            Event::CData(text) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                let value = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid flat ODG CDATA: {error}")))?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("flat ODG text size overflow"))?;
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(invalid("flat ODG text exceeds the extraction limit"));
                }
                pages[page].shapes[shape].text.push_str(&value);
                pages[page].shapes[shape]
                    .text_spans
                    .push(event_start..event_end);
            },
            Event::GeneralRef(reference) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                let value = resolve_reference(&reference)?;
                text_bytes = text_bytes
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("flat ODG text size overflow"))?;
                if text_bytes > MAX_TEXT_BYTES {
                    return Err(invalid("flat ODG text exceeds the extraction limit"));
                }
                pages[page].shapes[shape].text.push_str(&value);
                pages[page].shapes[shape]
                    .text_spans
                    .push(event_start..event_end);
            },
            Event::End(element) => {
                if paragraph_depths.last() == Some(&depth)
                    && namespace == NamespaceKind::Text
                    && element.local_name().as_ref() == b"p"
                {
                    paragraph_depths.pop();
                }
                if active_shapes.last().is_some_and(|shape| shape.0 == depth) {
                    active_shapes.pop();
                }
                if page_open.is_some_and(|page| page.0 == depth) {
                    page_open = None;
                }
                if drawing_depth == Some(depth) {
                    drawing_depth = None;
                }
                if body_depth == Some(depth) {
                    body_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("flat ODG XML depth underflow"))?;
            },
            Event::DocType(_) => return Err(invalid("DOCTYPE is not allowed in flat ODG")),
            Event::Eof => break,
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
        }
    }
    if depth != 0 || !root_seen || !body_seen || !drawing_seen || page_open.is_some() {
        return Err(invalid("flat ODG has an incomplete drawing structure"));
    }
    Ok(State { bytes, pages })
}

#[allow(
    clippy::too_many_arguments,
    reason = "scanner state is kept explicit and allocation-free"
)]
fn observe_start(
    reader: &NsReader<&[u8]>,
    namespace: NamespaceKind,
    element: &BytesStart<'_>,
    depth: usize,
    empty: bool,
    root_seen: &mut bool,
    body_seen: &mut bool,
    drawing_seen: &mut bool,
    body_depth: &mut Option<usize>,
    drawing_depth: &mut Option<usize>,
    page_open: &mut Option<(usize, usize)>,
    active_shapes: &mut Vec<(usize, usize, usize)>,
    paragraph_depths: &mut Vec<usize>,
    pages: &mut Vec<FlatPage>,
    shape_count: &mut usize,
) -> Result<()> {
    let local = element.local_name();
    if depth == 1 {
        if *root_seen
            || namespace != NamespaceKind::Office
            || local.as_ref() != b"document"
            || empty
        {
            return Err(invalid("flat ODG requires one office:document root"));
        }
        *root_seen = true;
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"body" {
        if *body_seen || depth != 2 || empty {
            return Err(invalid("flat ODG requires one non-empty office:body"));
        }
        *body_seen = true;
        *body_depth = Some(depth);
    } else if namespace == NamespaceKind::Office && local.as_ref() == b"drawing" {
        if *drawing_seen || *body_depth != Some(depth - 1) {
            return Err(invalid("office:drawing is misplaced or duplicated"));
        }
        *drawing_seen = true;
        if !empty {
            *drawing_depth = Some(depth);
        }
    } else if namespace == NamespaceKind::Draw && local.as_ref() == b"page" {
        if *drawing_depth != Some(depth - 1) || page_open.is_some() {
            return Err(invalid("draw:page is outside office:drawing"));
        }
        if pages.len() >= MAX_PAGES {
            return Err(invalid("flat ODG page count exceeds the limit"));
        }
        let index = pages.len();
        pages.push(FlatPage {
            name: attribute(reader, element, DRAW, b"name")?,
            shapes: Vec::new(),
        });
        if !empty {
            *page_open = Some((depth, index));
        }
    } else if page_open.is_some() && namespace == NamespaceKind::Draw && is_shape(local.as_ref()) {
        if *shape_count >= MAX_SHAPES {
            return Err(invalid("flat ODG shape count exceeds the limit"));
        }
        *shape_count += 1;
        let Some((_, page)) = *page_open else {
            return Err(invalid("draw shape is outside draw:page"));
        };
        let shape = pages[page].shapes.len();
        pages[page].shapes.push(FlatShape {
            name: attribute(reader, element, DRAW, b"name")?,
            text: String::new(),
            text_spans: Vec::new(),
        });
        if !empty {
            active_shapes.push((depth, page, shape));
        }
    } else if namespace == NamespaceKind::Draw && is_shape(local.as_ref()) {
        return Err(invalid("flat ODG drawing shape is outside draw:page"));
    } else if !active_shapes.is_empty()
        && namespace == NamespaceKind::Text
        && local.as_ref() == b"p"
        && !empty
    {
        paragraph_depths.push(depth);
    }
    Ok(())
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: &[u8],
    local: &[u8],
) -> Result<Option<String>> {
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| invalid(format!("invalid flat ODG attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&resolved, namespace) && name.as_ref() == local {
            return parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map(|value| Some(value.into_owned()))
                .map_err(|error| invalid(format!("invalid flat ODG attribute value: {error}")));
        }
    }
    Ok(None)
}

fn is_shape(local: &[u8]) -> bool {
    matches!(
        local,
        b"caption"
            | b"circle"
            | b"connector"
            | b"control"
            | b"custom-shape"
            | b"ellipse"
            | b"frame"
            | b"g"
            | b"line"
            | b"measure"
            | b"path"
            | b"page-thumbnail"
            | b"polygon"
            | b"polyline"
            | b"rect"
            | b"regular-polygon"
    )
}

fn resolve_reference(reference: &quick_xml::events::BytesRef<'_>) -> Result<String> {
    if let Some(character) = reference
        .resolve_char_ref()
        .map_err(|error| invalid(format!("invalid flat ODG character reference: {error}")))?
    {
        return Ok(character.to_string());
    }
    match reference
        .decode()
        .map_err(|error| invalid(format!("invalid flat ODG entity reference: {error}")))?
        .as_ref()
    {
        "amp" => Ok("&".into()),
        "lt" => Ok("<".into()),
        "gt" => Ok(">".into()),
        "apos" => Ok("'".into()),
        "quot" => Ok("\"".into()),
        name => Err(invalid(format!(
            "unsupported flat ODG entity reference '&{name};'"
        ))),
    }
}

fn checked_depth(depth: usize) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("flat ODG XML depth overflow"))?;
    if next_depth > MAX_DEPTH {
        return Err(invalid("flat ODG XML depth exceeds the limit"));
    }
    Ok(next_depth)
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_error| invalid("flat ODG XML position exceeds platform limits"))
}

fn resolved_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

fn classify(namespace: &ResolveResult<'_>) -> NamespaceKind {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) if *uri == OFFICE => NamespaceKind::Office,
        ResolveResult::Bound(Namespace(uri)) if *uri == DRAW => NamespaceKind::Draw,
        ResolveResult::Bound(Namespace(uri)) if *uri == TEXT => NamespaceKind::Text,
        ResolveResult::Bound(_) | ResolveResult::Unbound | ResolveResult::Unknown(_) => {
            NamespaceKind::Other
        },
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
