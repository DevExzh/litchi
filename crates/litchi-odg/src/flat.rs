//! Lossless flat (single-file) `OpenDocument` Drawing owners and bounded
//! text edits.
//!
//! [`Drawing`] is the immutable owner of one `.fodg` document. It opens
//! under one [`OpenOptions`] envelope: the source `input_bytes` and the
//! semantic `content_bytes` bound the XML, `depth` its nesting, `units` its
//! pages, `shapes` its drawing shapes and `text_bytes` every text value and
//! replacement. [`Drawing::edit`] stages an [`Edit`]; [`Edit::commit`]
//! publishes a new [`Drawing`] with its [`Patch`] and [`Diagnostics`];
//! [`Drawing::apply`] replays a patch onto its exact source.

use litchi_core::{Error as CoreError, FileFormat, Result};
use litchi_odf_common::{SemanticLimitKind, SemanticLimits};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
    reader::{NsReader, Reader},
};
use std::{borrow::Cow, io::Read, ops::Range, path::Path, sync::Arc};

/// Whether a commit changed the drawing bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits and refusals plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::OpenOptions;

/// A committed flat edit: the published drawing, its patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Drawing, Patch>;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const TEXT: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// The bounds one flat drawing is read and edited under, resolved from its
/// [`OpenOptions`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Bounds {
    /// Bytes of the flat XML (`min(input_bytes, content_bytes)`).
    bytes: usize,
    /// XML element nesting.
    depth: usize,
    /// Drawing pages (`units`).
    pages: usize,
    /// Drawing shapes.
    shapes: usize,
    /// Bytes of one text value and of all extracted text.
    text_bytes: usize,
    limits: SemanticLimits,
    options: OpenOptions,
}

impl Bounds {
    fn new(options: OpenOptions) -> Self {
        let limits = options.semantic_limits_value();
        let input =
            usize::try_from(options.source_limits_value().max_input_bytes()).unwrap_or(usize::MAX);
        Self {
            bytes: input.min(limits.usize_value(SemanticLimitKind::ContentBytes)),
            depth: limits.usize_value(SemanticLimitKind::Depth),
            pages: limits.usize_value(SemanticLimitKind::Units),
            shapes: limits.usize_value(SemanticLimitKind::Shapes),
            text_bytes: limits.usize_value(SemanticLimitKind::TextBytes),
            limits,
            options,
        }
    }

    /// Report an exceeded dimension as the typed limit failure.
    fn exceeded(self, kind: SemanticLimitKind, observed: usize) -> CoreError {
        self.limits
            .check_core(kind, observed)
            .err()
            .unwrap_or_else(|| invalid("flat ODG limit accounting failed"))
    }

    /// The input-bytes failure of an oversized flat document.
    fn input_exceeded(self, observed: usize) -> CoreError {
        if observed > self.limits.usize_value(SemanticLimitKind::ContentBytes) {
            return self.exceeded(SemanticLimitKind::ContentBytes, observed);
        }
        CoreError::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(self.bytes).unwrap_or(u64::MAX),
            scope: "ODF owned package input".into(),
        })
    }
}

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
    pages: Vec<Page>,
    bounds: Bounds,
}

/// The immutable, byte-preserving owner of one flat ODG document.
///
/// Cheap to clone; a commit publishes a new drawing and never changes its
/// source. The drawing keeps the [`OpenOptions`] it was opened with; its
/// edits and replays use them.
#[derive(Clone, Debug)]
pub struct Drawing(Arc<State>);

impl Drawing {
    /// Opens a flat ODG document from a path with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens a flat ODG document from a path under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let file = std::fs::File::open(path)?;
        Self::from_reader_with_options(file, options)
    }

    /// Reads a flat ODG document from a stream with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader(reader: impl Read) -> std::result::Result<Self, Error> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads a flat ODG document from a stream under caller-selected
    /// limits; the stream is read through a `bytes + 1` guard.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader_with_options(
        reader: impl Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let bounds = Bounds::new(options);
        let mut bytes = Vec::new();
        let guard = u64::try_from(bounds.bytes)
            .unwrap_or(u64::MAX)
            .saturating_add(1);
        reader.take(guard).read_to_end(&mut bytes)?;
        if bytes.len() > bounds.bytes {
            return Err(bounds.input_exceeded(bytes.len()).into());
        }
        Ok(Self::parse_owned(bytes, bounds)?)
    }

    /// Opens a flat ODG document with the default [`OpenOptions`]; the
    /// bytes are copied.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens a flat ODG document under caller-selected limits; the bytes
    /// are copied after they are validated.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when a selected dimension is exceeded and
    /// [`Error::Core`] when the bytes are not a structurally valid FODG
    /// document.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        // Validate the borrowed input before copying it, so a refused input
        // never costs an input-sized allocation.
        let bounds = Bounds::new(options);
        let pages = parse(bytes, bounds)?;
        Ok(Self(Arc::new(State {
            bytes: bytes.to_vec(),
            pages,
            bounds,
        })))
    }

    fn parse_owned(bytes: Vec<u8>, bounds: Bounds) -> Result<Self> {
        let pages = parse(&bytes, bounds)?;
        Ok(Self(Arc::new(State {
            bytes,
            pages,
            bounds,
        })))
    }

    /// The options this drawing was opened with.
    pub fn options(&self) -> OpenOptions {
        self.0.bounds.options
    }

    /// Returns the parsed drawing pages.
    #[must_use]
    pub fn pages(&self) -> &[Page] {
        &self.0.pages
    }

    /// Returns the original flat XML bytes exactly.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0.bytes
    }

    /// Starts a detached edit without mutating this drawing.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            changes: Vec::new(),
        }
    }

    /// Replays a patch onto this exact source drawing.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this drawing is not the patch's
    /// exact source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, Error> {
        if !patch.is_applicable_to(self) {
            return Err(Error::patch_conflict(
                "flat.apply",
                "the drawing is not the patch source",
            ));
        }
        let changed = !patch.is_noop();
        Ok(Commit::new(
            patch.target.clone(),
            patch.clone(),
            Diagnostics::new(changed),
        ))
    }

    /// Consumes the snapshot and returns its exact flat XML bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        Arc::try_unwrap(self.0).map_or_else(|state| state.bytes.clone(), |state| state.bytes)
    }
}

/// One drawing page of a flat drawing.
#[derive(Clone, Debug)]
pub struct Page {
    name: Option<String>,
    shapes: Vec<Shape>,
}

impl Page {
    /// Returns the optional `draw:name`.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Returns all bounded drawing shapes found below this page.
    #[must_use]
    pub fn shapes(&self) -> &[Shape] {
        &self.shapes
    }
}

/// One drawing shape and its extracted paragraph text.
#[derive(Clone, Debug)]
pub struct Shape {
    name: Option<String>,
    text: String,
    text_spans: Vec<Range<usize>>,
}

impl Shape {
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
pub struct Edit {
    source: Drawing,
    changes: Vec<TextChange>,
}

impl Edit {
    /// The source drawing this edit stages against.
    #[must_use]
    pub fn package(&self) -> &Drawing {
        &self.source
    }

    /// Discards every staged change in place.
    pub fn reset(&mut self) {
        self.changes.clear();
    }

    /// Stages replacement of a shape whose text occupies one exact XML text span.
    ///
    /// Shapes with empty, split, or mixed-content text are refused rather than
    /// being serialized through a lossy generic XML writer.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Refused`] for a shape whose text is not one plain
    /// span (several paragraphs, spans, links or fields) and for an
    /// out-of-bounds selector, and [`Error::Limit`] when the replacement
    /// exceeds `text_bytes`.
    pub fn set_shape_text(
        &mut self,
        page: usize,
        shape: usize,
        text: impl Into<String>,
    ) -> std::result::Result<(), Error> {
        let replacement = text.into();
        let bounds = self.source.0.bounds;
        if replacement.len() > bounds.text_bytes {
            return Err(bounds
                .exceeded(SemanticLimitKind::TextBytes, replacement.len())
                .into());
        }
        let selected = self
            .source
            .pages()
            .get(page)
            .and_then(|page_value| page_value.shapes().get(shape))
            .ok_or(Error::refused(
                litchi_core::edit::RefusalKind::NotFound,
                "flat.set_shape_text",
                "the shape selector is out of bounds",
            ))?;
        if selected.text_spans.len() != 1 {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "flat.set_shape_text",
                "the shape text is not one losslessly replaceable XML span",
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

    /// Validates and atomically publishes a new immutable drawing.
    ///
    /// An edit with no staged change commits the source itself:
    /// `diagnostics().changed()` is false and `patch().is_noop()` holds.
    ///
    /// # Errors
    ///
    /// Returns an error when staged spans overlap, the output exceeds the
    /// drawing's limits, output parsing fails, or readback differs.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    fn commit_core(self) -> Result<Commit> {
        let Edit { source, changes } = self;
        let bounds = source.0.bounds;
        let mut output_len = source.as_bytes().len();
        for change in &changes {
            let shape = &source.pages()[change.page].shapes()[change.shape];
            let span = shape
                .text_spans
                .first()
                .ok_or_else(|| invalid("flat ODG shape text source span is missing"))?;
            let escaped_len = escaped_xml_len(&change.after)?;
            output_len = output_len
                .checked_sub(
                    span.end
                        .checked_sub(span.start)
                        .ok_or_else(|| invalid("overlapping or invalid flat ODG text patch"))?,
                )
                .and_then(|length| length.checked_add(escaped_len))
                .ok_or_else(|| invalid("flat ODG text output size overflow"))?;
        }
        if output_len > bounds.bytes {
            return Err(bounds.input_exceeded(output_len));
        }
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(changes.len())
            .map_err(|source| CoreError::Allocation {
                resource: "flat ODG replacements",
                source,
            })?;
        for change in &changes {
            let shape = &source.pages()[change.page].shapes()[change.shape];
            replacements.push((
                shape.text_spans[0].clone(),
                escaped_xml_string(&change.after, bounds, "flat ODG replacement")?,
            ));
        }
        replacements.sort_unstable_by_key(|replacement| std::cmp::Reverse(replacement.0.start));
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(output_len.max(source.as_bytes().len()))
            .map_err(|source| CoreError::Allocation {
                resource: "flat ODG edited output",
                source,
            })?;
        bytes.extend_from_slice(source.as_bytes());
        let mut previous_start = bytes.len();
        for (span, replacement) in replacements {
            if span.end > previous_start || span.start > span.end || span.end > bytes.len() {
                return Err(invalid("overlapping or invalid flat ODG text patch"));
            }
            bytes.splice(span.clone(), replacement.bytes());
            previous_start = span.start;
        }
        let pages = parse(&bytes, bounds)?;
        let snapshot = Drawing(Arc::new(State {
            bytes,
            pages,
            bounds,
        }));
        for change in &changes {
            let actual = snapshot.pages()[change.page].shapes()[change.shape].text();
            if actual != change.after {
                return Err(invalid("flat ODG edit failed typed readback"));
            }
        }
        let changed = !changes.is_empty();
        Ok(Commit::new(
            snapshot.clone(),
            Patch {
                source,
                target: snapshot,
                changes,
            },
            Diagnostics::new(changed),
        ))
    }
}

/// A source-checked, reversible list of shape-text changes.
#[derive(Clone, Debug)]
pub struct Patch {
    source: Drawing,
    target: Drawing,
    changes: Vec<TextChange>,
}

impl Patch {
    /// Returns whether this patch was committed from the exact supplied bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Drawing) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Whether replaying this patch keeps the source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        Arc::ptr_eq(&self.source.0, &self.target.0)
            || self.source.as_bytes() == self.target.as_bytes()
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

fn parse(bytes: &[u8], bounds: Bounds) -> Result<Vec<Page>> {
    if bytes.len() > bounds.bytes {
        return Err(bounds.input_exceeded(bytes.len()));
    }
    // Keep the direct UTF-8 failure before format detection. Raw admission
    // must still precede detection because the detector uses `NsReader` and
    // can otherwise own an oversized root namespace declaration; malformed
    // UTF-8 therefore receives this typed source error instead of a generic
    // non-flat classification.
    let xml = std::str::from_utf8(bytes).map_err(|_error| invalid("flat ODG is not UTF-8"))?;
    validate_raw_events(xml, bounds)?;
    if litchi_odf_common::detect::flat(bytes, litchi_odf_common::detect::Limits::default())
        .ok()
        .flatten()
        != Some(FileFormat::Odg)
    {
        return Err(invalid("input is not a flat ODG document"));
    }
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
    let mut pages = Vec::<Page>::new();
    let mut shape_count = 0usize;
    let mut text_bytes = 0usize;

    loop {
        let event_start = position(&reader)?;
        let (resolved_namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("invalid flat ODG XML: {error}")))?;
        validate_borrowed_event(&event, bounds)?;
        let namespace = classify(&resolved_namespace);
        let event_end = position(&reader)?;
        match event {
            Event::Start(element) => {
                depth = checked_depth(depth, bounds)?;
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
                    bounds,
                )?;
            },
            Event::Empty(element) => {
                let virtual_depth = checked_depth(depth, bounds)?;
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
                    bounds,
                )?;
            },
            Event::Text(text) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                admit_text_bytes(text.as_ref(), text_bytes, "flat ODG text", bounds)?;
                let decoded = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid flat ODG text: {error}")))?;
                let decoded = own_bounded_cow(decoded, "flat ODG text", bounds)?;
                let value = quick_xml::escape::unescape(&decoded)
                    .map_err(|error| invalid(format!("invalid flat ODG text escape: {error}")))?;
                let value = own_bounded_cow(value, "flat ODG text", bounds)?;
                append_text(
                    &mut pages,
                    page,
                    shape,
                    &value,
                    &mut text_bytes,
                    event_start..event_end,
                    bounds,
                )?;
            },
            Event::CData(text) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                admit_text_bytes(text.as_ref(), text_bytes, "flat ODG CDATA", bounds)?;
                let value = text
                    .decode()
                    .map_err(|error| invalid(format!("invalid flat ODG CDATA: {error}")))?;
                let value = own_bounded_cow(value, "flat ODG CDATA", bounds)?;
                append_text(
                    &mut pages,
                    page,
                    shape,
                    &value,
                    &mut text_bytes,
                    event_start..event_end,
                    bounds,
                )?;
            },
            Event::GeneralRef(reference) if !paragraph_depths.is_empty() => {
                let Some((_, page, shape)) = active_shapes.last().copied() else {
                    continue;
                };
                admit_text_bytes(reference.as_ref(), text_bytes, "flat ODG reference", bounds)?;
                let value = resolve_reference(&reference, bounds)?;
                append_text(
                    &mut pages,
                    page,
                    shape,
                    &value,
                    &mut text_bytes,
                    event_start..event_end,
                    bounds,
                )?;
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
    Ok(pages)
}

fn validate_borrowed_event(event: &Event<'_>, bounds: Bounds) -> Result<()> {
    match event {
        Event::Start(element) | Event::Empty(element) => {
            validate_raw_value(element.as_ref(), "flat ODG start event", bounds)?;
            for raw_attribute in element.attributes() {
                let attribute = raw_attribute.map_err(|error| {
                    invalid(format!("invalid flat ODG event attribute: {error}"))
                })?;
                validate_raw_value(attribute.key.as_ref(), "flat ODG attribute name", bounds)?;
                validate_raw_value(attribute.value.as_ref(), "flat ODG attribute value", bounds)?;
            }
        },
        Event::End(element) => validate_raw_value(element.as_ref(), "flat ODG end event", bounds)?,
        Event::Text(text) | Event::Comment(text) | Event::DocType(text) => {
            validate_raw_value(text.as_ref(), "flat ODG text event", bounds)?;
        },
        Event::CData(text) => validate_raw_value(text.as_ref(), "flat ODG CDATA event", bounds)?,
        Event::Decl(declaration) => {
            validate_raw_value(declaration.as_ref(), "flat ODG declaration event", bounds)?;
        },
        Event::PI(instruction) => {
            validate_raw_value(
                instruction.as_ref(),
                "flat ODG processing-instruction event",
                bounds,
            )?;
        },
        Event::GeneralRef(reference) => {
            validate_raw_value(reference.as_ref(), "flat ODG reference event", bounds)?;
        },
        Event::Eof => {},
    }
    Ok(())
}

/// Admit raw events before `NsReader` can store namespace declarations or
/// clone an unknown prefix while resolving a qualified name.
fn validate_raw_events(xml: &str, bounds: Bounds) -> Result<()> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    loop {
        let event = reader
            .read_event()
            .map_err(|error| invalid(format!("invalid flat ODG XML: {error}")))?;
        validate_borrowed_event(&event, bounds)?;
        match event {
            Event::Start(_) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("flat ODG XML depth overflow"))?;
                if depth > bounds.depth {
                    return Err(bounds.exceeded(SemanticLimitKind::Depth, depth));
                }
            },
            Event::Empty(_) => {
                let virtual_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("flat ODG XML depth overflow"))?;
                if virtual_depth > bounds.depth {
                    return Err(bounds.exceeded(SemanticLimitKind::Depth, virtual_depth));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::DocType(_)
            | Event::GeneralRef(_)
            | Event::PI(_)
            | Event::Text(_) => {},
            Event::Eof => break,
        }
    }
    Ok(())
}

fn validate_raw_value(value: &[u8], _owner: &str, bounds: Bounds) -> Result<()> {
    if value.len() > bounds.text_bytes {
        return Err(bounds.exceeded(SemanticLimitKind::TextBytes, value.len()));
    }
    Ok(())
}

fn validate_decoded_value(value: &str, owner: &str, bounds: Bounds) -> Result<()> {
    if value.len() > bounds.text_bytes {
        return Err(bounds.exceeded(SemanticLimitKind::TextBytes, value.len()));
    }
    if value.contains('\0') {
        return Err(invalid(format!("{owner} contains a NUL character")));
    }
    Ok(())
}

fn own_bounded_cow(value: Cow<'_, str>, owner: &str, bounds: Bounds) -> Result<String> {
    match value {
        Cow::Borrowed(value) => clone_bounded_string(value, owner, bounds),
        Cow::Owned(value) => {
            validate_decoded_value(&value, owner, bounds)?;
            Ok(value)
        },
    }
}

fn clone_bounded_string(value: &str, owner: &str, bounds: Bounds) -> Result<String> {
    validate_decoded_value(value, owner, bounds)?;
    let mut result = String::new();
    result
        .try_reserve_exact(value.len())
        .map_err(|source| CoreError::Allocation {
            resource: "flat ODG text",
            source,
        })?;
    result.push_str(value);
    Ok(result)
}

fn admit_text_bytes(raw: &[u8], current: usize, owner: &str, bounds: Bounds) -> Result<()> {
    validate_raw_value(raw, owner, bounds)?;
    let next = current
        .checked_add(raw.len())
        .ok_or_else(|| invalid("flat ODG text size overflow"))?;
    if next > bounds.text_bytes {
        return Err(bounds.exceeded(SemanticLimitKind::TextBytes, next));
    }
    Ok(())
}

fn append_text(
    pages: &mut [Page],
    page: usize,
    shape: usize,
    value: &str,
    text_bytes: &mut usize,
    span: Range<usize>,
    bounds: Bounds,
) -> Result<()> {
    let next = text_bytes
        .checked_add(value.len())
        .ok_or_else(|| invalid("flat ODG text size overflow"))?;
    if next > bounds.text_bytes {
        return Err(bounds.exceeded(SemanticLimitKind::TextBytes, next));
    }
    let flat_shape = pages
        .get_mut(page)
        .and_then(|page_value| page_value.shapes.get_mut(shape))
        .ok_or_else(|| invalid("flat ODG text owner is missing"))?;
    flat_shape
        .text
        .try_reserve_exact(value.len())
        .map_err(|source| CoreError::Allocation {
            resource: "flat ODG extracted text",
            source,
        })?;
    flat_shape
        .text_spans
        .try_reserve_exact(1)
        .map_err(|source| CoreError::Allocation {
            resource: "flat ODG text spans",
            source,
        })?;
    flat_shape.text.push_str(value);
    flat_shape.text_spans.push(span);
    *text_bytes = next;
    Ok(())
}

fn escaped_xml_len(value: &str) -> Result<usize> {
    value.chars().try_fold(0usize, |total, character| {
        let extra = match character {
            '&' => 4,
            '<' | '>' => 3,
            '\'' | '"' => 5,
            _ => 0,
        };
        total
            .checked_add(character.len_utf8())
            .and_then(|size| size.checked_add(extra))
            .ok_or_else(|| invalid("flat ODG escaped text size overflow"))
    })
}

fn push_escaped_xml(output: &mut String, value: &str, bounds: Bounds) -> Result<()> {
    let escaped_len = escaped_xml_len(value)?;
    let required = output
        .len()
        .checked_add(escaped_len)
        .ok_or_else(|| invalid("flat ODG escaped text size overflow"))?;
    if required > output.capacity() {
        output
            .try_reserve_exact(required - output.len())
            .map_err(|source| CoreError::Allocation {
                resource: "flat ODG escaped text",
                source,
            })?;
    }
    let mut start = 0usize;
    for (index, character) in value.char_indices() {
        let replacement = match character {
            '&' => "&amp;",
            '<' => "&lt;",
            '>' => "&gt;",
            '\'' => "&apos;",
            '"' => "&quot;",
            _ => continue,
        };
        output.push_str(&value[start..index]);
        output.push_str(replacement);
        start = index + character.len_utf8();
    }
    output.push_str(&value[start..]);
    if output.len() > bounds.bytes {
        return Err(bounds.input_exceeded(output.len()));
    }
    Ok(())
}

fn escaped_xml_string(value: &str, bounds: Bounds, _owner: &str) -> Result<String> {
    let escaped_len = escaped_xml_len(value)?;
    if escaped_len > bounds.bytes {
        return Err(bounds.input_exceeded(escaped_len));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(escaped_len)
        .map_err(|source| CoreError::Allocation {
            resource: "flat ODG escaped text",
            source,
        })?;
    push_escaped_xml(&mut output, value, bounds)?;
    Ok(output)
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
    pages: &mut Vec<Page>,
    shape_count: &mut usize,
    bounds: Bounds,
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
        if pages.len() >= bounds.pages {
            return Err(bounds.exceeded(SemanticLimitKind::Units, pages.len().saturating_add(1)));
        }
        let index = pages.len();
        pages.push(Page {
            name: attribute(reader, element, DRAW, b"name", bounds)?,
            shapes: Vec::new(),
        });
        if !empty {
            *page_open = Some((depth, index));
        }
    } else if page_open.is_some() && namespace == NamespaceKind::Draw && is_shape(local.as_ref()) {
        if *shape_count >= bounds.shapes {
            return Err(bounds.exceeded(SemanticLimitKind::Shapes, shape_count.saturating_add(1)));
        }
        *shape_count += 1;
        let Some((_, page)) = *page_open else {
            return Err(invalid("draw shape is outside draw:page"));
        };
        let shape = pages[page].shapes.len();
        pages[page].shapes.push(Shape {
            name: attribute(reader, element, DRAW, b"name", bounds)?,
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
    bounds: Bounds,
) -> Result<Option<String>> {
    for raw_attribute in element.attributes() {
        let parsed_attribute = raw_attribute
            .map_err(|error| invalid(format!("invalid flat ODG attribute: {error}")))?;
        let (resolved, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
        if resolved_bound(&resolved, namespace) && name.as_ref() == local {
            validate_raw_value(
                parsed_attribute.key.as_ref(),
                "flat ODG attribute name",
                bounds,
            )?;
            validate_raw_value(
                parsed_attribute.value.as_ref(),
                "flat ODG attribute value",
                bounds,
            )?;
            let value = parsed_attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| invalid(format!("invalid flat ODG attribute value: {error}")))?;
            validate_decoded_value(value.as_ref(), "flat ODG attribute value", bounds)?;
            return own_bounded_cow(value, "flat ODG attribute value", bounds).map(Some);
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

fn resolve_reference(
    reference: &quick_xml::events::BytesRef<'_>,
    bounds: Bounds,
) -> Result<String> {
    validate_raw_value(reference.as_ref(), "flat ODG character reference", bounds)?;
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

fn checked_depth(depth: usize, bounds: Bounds) -> Result<usize> {
    let next_depth = depth
        .checked_add(1)
        .ok_or_else(|| invalid("flat ODG XML depth overflow"))?;
    if next_depth > bounds.depth {
        return Err(bounds.exceeded(SemanticLimitKind::Depth, next_depth));
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

fn invalid(message: impl Into<String>) -> CoreError {
    CoreError::InvalidFormat(message.into())
}
