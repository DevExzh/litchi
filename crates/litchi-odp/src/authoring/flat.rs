//! Bounded, source-bound flat (single-file) `OpenDocument` presentation
//! owners and slide edits.
//!
//! [`Presentation`] is the immutable owner of one `.fodp` document, opened
//! under one [`OpenOptions`] envelope: the source `input_bytes` and the
//! semantic `content_bytes` bound the XML, `depth` its nesting, `units` its
//! slides and `text_bytes` one slide replacement. [`Presentation::edit`]
//! stages an [`Edit`]; [`Edit::commit`] publishes a new [`Presentation`] with
//! its [`Patch`] and [`Diagnostics`]; [`Presentation::apply`] replays a patch
//! onto its exact source.

use crate::codec::Parser;
use crate::model::{Settings, Slide, declaration, page_layout, page_metadata, settings};
use litchi_core::xml::escape_xml;
use litchi_core::{Error as CoreError, Result};
use litchi_odf_common::{SemanticLimitKind, SemanticLimits};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;
use std::io::Read;
use std::ops::Range;
use std::path::Path;
use std::sync::Arc;

/// Whether a commit changed the presentation bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits and refusals plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::OpenOptions;

/// A committed flat edit: the published presentation, its patch and
/// diagnostics.
pub type Commit = litchi_core::edit::Commit<Presentation, Patch>;
use xml_minifier::audit;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.presentation";

/// The bounds one flat presentation is read and edited under, resolved from
/// its [`OpenOptions`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Bounds {
    bytes: usize,
    text_bytes: usize,
    depth: usize,
    slides: usize,
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
            text_bytes: limits.usize_value(SemanticLimitKind::TextBytes),
            depth: limits.usize_value(SemanticLimitKind::Depth),
            slides: limits.usize_value(SemanticLimitKind::Units),
            limits,
            options,
        }
    }

    fn exceeded(self, kind: SemanticLimitKind, observed: usize) -> CoreError {
        self.limits
            .check_core(kind, observed)
            .err()
            .unwrap_or_else(|| invalid_error("flat ODP limit accounting failed"))
    }

    fn check_input(self, observed: usize) -> Result<()> {
        if observed <= self.bytes {
            return Ok(());
        }
        if observed > self.limits.usize_value(SemanticLimitKind::ContentBytes) {
            return Err(self.exceeded(SemanticLimitKind::ContentBytes, observed));
        }
        Err(CoreError::ResourceLimit(litchi_core::ResourceLimit {
            resource: litchi_core::Resource::InputBytes,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            limit: u64::try_from(self.bytes).unwrap_or(u64::MAX),
            scope: "ODF owned package input".into(),
        }))
    }
}

/// The immutable owner of one flat `OpenDocument` presentation (`.fodp`).
///
/// Cheap to clone; a commit publishes a new presentation and never changes
/// its source. The presentation keeps the [`OpenOptions`] it was opened
/// with; its edits and replays use them.
#[derive(Clone)]
pub struct Presentation {
    bounds: Bounds,
    bytes: Arc<[u8]>,
    xml: Arc<str>,
    page_ranges: Arc<[Range<usize>]>,
    slides: Arc<[Slide]>,
    tolerance: Arc<litchi_odf_common::leniency::ToleranceReport>,
}

impl Presentation {
    /// Opens a flat presentation from a path with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens a flat presentation from a path under caller-selected limits.
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

    /// Reads a flat presentation from a stream with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader(reader: impl Read) -> std::result::Result<Self, Error> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads a flat presentation from a stream under caller-selected limits;
    /// the stream is read through a `bytes + 1` guard.
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
        Ok(Self::parse(&bytes, bounds)?)
    }

    /// Opens bounded UTF-8 flat-presentation bytes with the default
    /// [`OpenOptions`]; the bytes are copied.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Opens bounded UTF-8 flat-presentation bytes under caller-selected
    /// limits; the bytes are copied.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Limit`] when a selected dimension is exceeded and
    /// [`Error::Core`] when the XML is not a structurally valid flat
    /// presentation.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        Ok(Self::parse(bytes, Bounds::new(options))?)
    }

    fn parse(bytes: &[u8], bounds: Bounds) -> Result<Self> {
        bounds.check_input(bytes.len())?;
        let xml = String::from_utf8(bytes.to_vec())
            .map_err(|error| CoreError::InvalidFormat(format!("flat ODP is not UTF-8: {error}")))?;
        let page_ranges = scan_flat(&xml, bounds)?;
        let (slides, tolerance) = Parser::parse_slides_with_report(&xml, Some(&xml))?;
        if slides.len() != page_ranges.len() {
            return invalid("flat ODP slide projection does not match its page structure");
        }
        Ok(Self {
            bounds,
            bytes: Arc::from(bytes),
            xml: Arc::from(xml),
            page_ranges: Arc::from(page_ranges),
            slides: Arc::from(slides),
            tolerance: Arc::new(tolerance),
        })
    }

    /// Everything a tolerant read repaired while this presentation was
    /// parsed (LENIENCY_AUDIT F9.10).
    ///
    /// An empty report means the document conformed; see
    /// [`litchi_odf_common::leniency::ToleranceReport`].
    #[must_use]
    pub fn tolerance_report(&self) -> &litchi_odf_common::leniency::ToleranceReport {
        &self.tolerance
    }

    /// Borrow the exact source bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the exact source bytes without normalization.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.bytes.to_vec()
    }

    /// Borrow the parsed slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Return the number of slides.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Read inert presentation declarations.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn declarations(&self) -> Result<declaration::Collection> {
        declaration::parse(&self.xml)
    }

    /// Read flat-document presentation page layouts.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn layouts(&self) -> Result<page_layout::Collection> {
        page_layout::parse(&self.xml)
    }

    /// Read static page metadata.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn pages(&self) -> Result<page_metadata::Collection> {
        page_metadata::parse(&self.xml)
    }

    /// Read inert slide-show settings.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn settings(&self) -> Result<Option<Settings>> {
        settings::parse(&self.xml)
    }

    /// The options this presentation was opened with.
    pub fn options(&self) -> OpenOptions {
        self.bounds.options
    }

    /// Starts an isolated, source-bound edit.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            draft: None,
            replacements: BTreeMap::new(),
        }
    }

    /// Replays a patch onto this exact source presentation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this presentation is not the
    /// patch's exact source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, Error> {
        if !patch.is_applicable_to(self) {
            return Err(Error::patch_conflict(
                "flat.apply",
                "the presentation is not the patch source",
            ));
        }
        Ok(Commit::new(
            patch.after.clone(),
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }
}

/// Checked slide selector.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Selector<'a> {
    /// Zero-based position.
    Index(usize),
    /// Exact title; duplicates are an ambiguity error.
    Title(&'a str),
}

impl From<usize> for Selector<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

impl<'a> From<&'a str> for Selector<'a> {
    fn from(value: &'a str) -> Self {
        Self::Title(value)
    }
}

/// An isolated, source-bound replacement edit of a flat presentation.
pub struct Edit {
    source: Presentation,
    draft: Option<Vec<Slide>>,
    replacements: BTreeMap<usize, String>,
}

impl Edit {
    /// The source presentation this edit stages against.
    #[must_use]
    pub fn package(&self) -> &Presentation {
        &self.source
    }

    /// Discards every staged replacement in place.
    pub fn reset(&mut self) {
        self.draft = None;
        self.replacements.clear();
    }

    /// Borrow the staged slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        self.draft
            .as_deref()
            .unwrap_or_else(|| self.source.slides())
    }

    /// Replace one slide's modeled title and drawing payload.
    ///
    /// Page attributes and speaker notes are retained verbatim. Unsupported
    /// direct page children cause a typed refusal before staging.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn replace<'a, S>(&mut self, selector: S, title: &str, text: &str) -> Result<Option<()>>
    where
        S: Into<Selector<'a>>,
    {
        check_text(title, text, self.source.bounds)?;
        let Some(index) = select(self.slides(), selector.into())? else {
            return Ok(None);
        };
        if self.slides()[index].title.as_deref() == Some(title) && self.slides()[index].text == text
        {
            return Ok(Some(()));
        }
        let range = &self.source.page_ranges[index];
        let source_page = &self.source.xml[range.clone()];
        let replacement = replace_page(source_page, title, text)?;
        self.replacements.insert(index, replacement);
        let slide = &mut self.draft_mut()[index];
        slide.title = (!title.is_empty()).then(|| title.to_owned());
        text.clone_into(&mut slide.text);
        slide.shapes.clear();
        slide.animations.clear();
        slide.legacy_animation = None;
        Ok(Some(()))
    }

    /// Validate, serialize, reparse, and atomically publish the staged bytes.
    ///
    /// An edit with no staged replacement commits the source itself:
    /// `diagnostics().changed()` is false and `patch().is_noop()` holds.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    fn commit_core(self) -> Result<Commit> {
        let bounds = self.source.bounds;
        if self.replacements.is_empty() {
            return Ok(Commit::new(
                self.source.clone(),
                Patch {
                    before: self.source.clone(),
                    after: self.source,
                },
                Diagnostics::UNCHANGED,
            ));
        }
        let xml = &self.source.xml;
        let mut capacity = xml.len();
        for (index, replacement) in &self.replacements {
            let original = self.source.page_ranges[*index].len();
            capacity = capacity
                .checked_sub(original)
                .and_then(|size| size.checked_add(replacement.len()))
                .ok_or_else(|| invalid_error("flat ODP output size overflow"))?;
        }
        if capacity > bounds.bytes {
            return Err(bounds.exceeded(SemanticLimitKind::ContentBytes, capacity));
        }
        let mut output = String::with_capacity(capacity);
        let mut cursor = 0;
        for (index, range) in self.source.page_ranges.iter().enumerate() {
            push_bounded(&mut output, &xml[cursor..range.start], bounds)?;
            if let Some(replacement) = self.replacements.get(&index) {
                push_bounded(&mut output, replacement, bounds)?;
            } else {
                push_bounded(&mut output, &xml[range.clone()], bounds)?;
            }
            cursor = range.end;
        }
        push_bounded(&mut output, &xml[cursor..], bounds)?;
        validate_compact_xml(&output, bounds)?;
        let presentation = Presentation::parse(output.as_bytes(), bounds)?;
        if presentation.slides()
            != self
                .draft
                .as_deref()
                .unwrap_or_else(|| self.source.slides())
        {
            return invalid("flat ODP commit readback differs from the staged slide model");
        }
        let changed = presentation.as_bytes() != self.source.as_bytes();
        let patch = Patch {
            before: self.source,
            after: presentation.clone(),
        };
        Ok(Commit::new(presentation, patch, Diagnostics::new(changed)))
    }

    fn draft_mut(&mut self) -> &mut Vec<Slide> {
        self.draft
            .get_or_insert_with(|| self.source.slides().to_vec())
    }
}

/// Exact-source-checked reversible flat-presentation patch.
#[derive(Clone)]
pub struct Patch {
    before: Presentation,
    after: Presentation,
}

impl Patch {
    /// Whether this patch was made from exactly `source`'s bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Presentation) -> bool {
        source.as_bytes() == self.before.as_bytes()
    }

    /// Return the inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Whether applying the patch changes no bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_bytes() == self.after.as_bytes()
    }
}

fn scan_flat(xml: &str, bounds: Bounds) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Option<Vec<u8>>, Vec<u8>, usize)> = Vec::new();
    let mut pages = Vec::new();
    let mut body_count = 0usize;
    let mut presentation_count = 0usize;
    let mut root_seen = false;
    loop {
        let (namespace, event) = {
            let (resolved, event) = reader.read_resolved_event().map_err(|error| {
                CoreError::InvalidFormat(format!("invalid flat ODP XML: {error}"))
            })?;
            (namespace_of(&resolved), event)
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid_error("flat ODP XML position exceeds platform limits"))?;
        match event {
            Event::Start(element) => {
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                validate_open(
                    &reader,
                    &element,
                    namespace.as_deref(),
                    &local,
                    stack.len(),
                    false,
                    &mut root_seen,
                    &mut body_count,
                    &mut presentation_count,
                )?;
                if is(namespace.as_deref(), &local, DRAW, b"page") {
                    validate_page_parent_and_limit(&stack, pages.len(), bounds)?;
                }
                if stack.len() >= bounds.depth {
                    return Err(bounds.exceeded(SemanticLimitKind::Depth, stack.len() + 1));
                }
                stack.push((namespace, local, start));
            },
            Event::Empty(element) => {
                let start = event_start(xml, end)?;
                let local = element.local_name().as_ref().to_vec();
                validate_open(
                    &reader,
                    &element,
                    namespace.as_deref(),
                    &local,
                    stack.len(),
                    true,
                    &mut root_seen,
                    &mut body_count,
                    &mut presentation_count,
                )?;
                if is(namespace.as_deref(), &local, DRAW, b"page") {
                    validate_page_parent_and_limit(&stack, pages.len(), bounds)?;
                    pages.push(start..end);
                }
            },
            Event::End(_) => {
                let (_, local, start) = stack
                    .pop()
                    .ok_or_else(|| invalid_error("flat ODP XML depth underflow"))?;
                if is(namespace.as_deref(), &local, DRAW, b"page")
                    && xml[start..end].starts_with('<')
                {
                    pages.push(start..end);
                }
            },
            Event::Text(text) if stack.is_empty() => {
                if !text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| CoreError::InvalidFormat(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return invalid("flat ODP has text outside its root element");
                }
            },
            // The five predefined entities and character references are
            // plain XML (quick-xml reports them as references, record 0341).
            Event::GeneralRef(reference) if is_plain_reference(&reference) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("DOCTYPE and custom entities are not allowed in flat ODP");
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    if !stack.is_empty() || !root_seen || body_count != 1 || presentation_count != 1 {
        return invalid("flat ODP is missing its unique presentation body");
    }
    Ok(pages)
}

fn validate_page_parent_and_limit(
    stack: &[(Option<Vec<u8>>, Vec<u8>, usize)],
    page_count: usize,
    bounds: Bounds,
) -> Result<()> {
    if stack.len() != 3
        || !stack.get(2).is_some_and(|(namespace, local, _)| {
            is(namespace.as_deref(), local, OFFICE, b"presentation")
        })
    {
        return invalid("draw:page is outside office:presentation");
    }
    if page_count >= bounds.slides {
        return Err(bounds.exceeded(SemanticLimitKind::Units, page_count + 1));
    }
    Ok(())
}

#[allow(
    clippy::too_many_arguments,
    reason = "the scanner threads root/body/presentation bookkeeping through one validation pass; splitting it would obscure the state machine"
)]
fn validate_open(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    namespace: Option<&[u8]>,
    local: &[u8],
    depth: usize,
    empty: bool,
    root_seen: &mut bool,
    body_count: &mut usize,
    presentation_count: &mut usize,
) -> Result<()> {
    if depth == 0 {
        if *root_seen || empty || !is(namespace, local, OFFICE, b"document") {
            return invalid("flat ODP root is not office:document");
        }
        *root_seen = true;
        let mut mimetype = None;
        for attribute in element.attributes() {
            let parsed_attribute = attribute.map_err(|error| {
                CoreError::InvalidFormat(format!("invalid flat ODP attribute: {error}"))
            })?;
            let (resolved, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
            if name.as_ref() == b"mimetype"
                && matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE)
            {
                mimetype = Some(
                    parsed_attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| CoreError::InvalidFormat(error.to_string()))?
                        .into_owned(),
                );
            }
        }
        if mimetype.as_deref() != Some(MIMETYPE) {
            return invalid("flat ODP has the wrong office:mimetype");
        }
    } else if is(namespace, local, OFFICE, b"body") {
        if depth != 1 || empty {
            return invalid("flat ODP has a misplaced or empty office:body");
        }
        *body_count += 1;
    } else if is(namespace, local, OFFICE, b"presentation") {
        if depth != 2 {
            return invalid("flat ODP has a misplaced office:presentation");
        }
        *presentation_count += 1;
    }
    Ok(())
}

fn replace_page(page: &str, title: &str, text: &str) -> Result<String> {
    let open_end = page
        .find('>')
        .ok_or_else(|| invalid_error("flat ODP page start tag is malformed"))?;
    let close_start = page
        .rfind("</")
        .ok_or_else(|| invalid_error("empty flat ODP pages cannot be replaced"))?;
    let notes = direct_notes(page)?;
    let mut output = String::with_capacity(page.len() + title.len() + text.len() + 256);
    output.push_str(&page[..=open_end]);
    if !title.is_empty() {
        output.push_str(
            r#"<draw:frame draw:layer="layout" presentation:class="title"><draw:text-box>"#,
        );
        push_paragraphs(&mut output, title);
        output.push_str("</draw:text-box></draw:frame>");
    }
    if !text.is_empty() {
        output.push_str(
            r#"<draw:frame draw:layer="layout" presentation:class="object"><draw:text-box>"#,
        );
        push_paragraphs(&mut output, text);
        output.push_str("</draw:text-box></draw:frame>");
    }
    if let Some(notes_fragment) = notes {
        output.push_str(notes_fragment);
    }
    output.push_str(&page[close_start..]);
    Ok(output)
}

fn direct_notes(page: &str) -> Result<Option<&str>> {
    let mut reader = NsReader::from_str(page);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Option<Vec<u8>>, Vec<u8>, usize)> = Vec::new();
    let mut notes = None;
    loop {
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|error| CoreError::InvalidFormat(format!("invalid flat ODP page: {error}")))?;
        let namespace = namespace_of(&resolved);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid_error("flat ODP XML position exceeds platform limits"))?;
        match event {
            Event::Start(element) => {
                let start = event_start(page, end)?;
                let local = element.local_name().as_ref().to_vec();
                let inherited = inherited_fragment_namespace(namespace, element.name().as_ref());
                if stack.len() == 1 && !is_modeled_direct_page_child(inherited.as_deref(), &local) {
                    return Err(CoreError::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "odp.authoring.flat",
                        "flat ODP page has an unsupported direct child; replacement refused",
                    )));
                }
                stack.push((inherited, local, start));
            },
            Event::Empty(element) => {
                let inherited = inherited_fragment_namespace(namespace, element.name().as_ref());
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if stack.len() == 1 && !is_modeled_direct_page_child(inherited.as_deref(), local) {
                    return Err(CoreError::Refused(litchi_core::edit::Refusal::new(
                        litchi_core::edit::RefusalKind::UnsupportedOperation,
                        "odp.authoring.flat",
                        "flat ODP page has an unsupported direct child; replacement refused",
                    )));
                }
                if stack.len() == 1 && is(inherited.as_deref(), local, PRESENTATION, b"notes") {
                    if notes.is_some() {
                        return invalid("flat ODP page has duplicate speaker notes");
                    }
                    let start = event_start(page, end)?;
                    notes = Some(&page[start..end]);
                }
            },
            Event::End(_) => {
                let (popped_namespace, local, start) = stack
                    .pop()
                    .ok_or_else(|| invalid_error("flat ODP page depth underflow"))?;
                if stack.len() == 1
                    && is(popped_namespace.as_deref(), &local, PRESENTATION, b"notes")
                {
                    if notes.is_some() {
                        return invalid("flat ODP page has duplicate speaker notes");
                    }
                    notes = Some(&page[start..end]);
                }
            },
            Event::GeneralRef(reference) if is_plain_reference(&reference) => {},
            Event::DocType(_) | Event::GeneralRef(_) => {
                return invalid("DOCTYPE and custom entities are not allowed in flat ODP pages");
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_) => {},
        }
    }
    Ok(notes)
}

/// Whether an entity reference is one of XML's five predefined entities or
/// a character reference to an XML 1.0 character, which need no document
/// type (as `litchi-odf-common`'s package validation admits them).
fn is_plain_reference(reference: &quick_xml::events::BytesRef<'_>) -> bool {
    matches!(
        reference.as_ref(),
        b"amp" | b"lt" | b"gt" | b"apos" | b"quot"
    ) || (reference.is_char_ref()
        && reference
            .resolve_char_ref()
            .ok()
            .flatten()
            .is_some_and(|character| {
                matches!(character, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..)
            }))
}

fn is_modeled_direct_page_child(namespace: Option<&[u8]>, local: &[u8]) -> bool {
    is(namespace, local, DRAW, b"frame")
        || is(namespace, local, DRAW, b"g")
        || is(namespace, local, DRAW, b"custom-shape")
        || is(namespace, local, DRAW, b"rect")
        || is(namespace, local, DRAW, b"ellipse")
        || is(namespace, local, DRAW, b"line")
        || is(namespace, local, DRAW, b"connector")
        || is(namespace, local, PRESENTATION, b"notes")
}

fn push_paragraphs(output: &mut String, value: &str) {
    for line in value.split('\n') {
        output.push_str("<text:p>");
        output.push_str(&escape_xml(line));
        output.push_str("</text:p>");
    }
}

fn push_bounded(output: &mut String, value: &str, bounds: Bounds) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| invalid_error("flat ODP output size overflow"))?;
    if size > bounds.bytes {
        return Err(bounds.exceeded(SemanticLimitKind::ContentBytes, size));
    }
    output.push_str(value);
    Ok(())
}

fn validate_compact_xml(xml: &str, bounds: Bounds) -> Result<()> {
    let limits = audit::Limits::new(
        bounds.bytes,
        bounds.depth,
        1_000_000,
        250_000,
        bounds.text_bytes.min(bounds.bytes),
        bounds.bytes,
    )
    .map_err(|source| invalid_error(format!("invalid flat ODP XML audit limits: {source}")))?;
    let _report = audit::verify(xml.as_bytes(), limits).map_err(|source| match source {
        audit::Error::NotCompact(_) => CoreError::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odp.authoring.flat",
            "flat ODP XML is not compact",
        )),
        audit::Error::Limit { .. }
        | audit::Error::Encoding { .. }
        | audit::Error::Malformed { .. }
        | audit::Error::Doctype { .. }
        | audit::Error::Allocation
        | _ => CoreError::InvalidFormat(format!("flat ODP XML failed audit: {source}")),
    })?;
    Ok(())
}

fn select(slides: &[Slide], selector: Selector<'_>) -> Result<Option<usize>> {
    match selector {
        Selector::Index(index) => Ok((index < slides.len()).then_some(index)),
        Selector::Title(title) => {
            let mut matches = slides
                .iter()
                .enumerate()
                .filter(|(_, slide)| slide.title.as_deref() == Some(title));
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return invalid("flat ODP slide title selector is ambiguous");
            }
            Ok(first)
        },
    }
}

fn check_text(title: &str, text: &str, bounds: Bounds) -> Result<()> {
    let size = title
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid_error("flat ODP slide text size overflow"))?;
    if size > bounds.text_bytes {
        return Err(bounds.exceeded(SemanticLimitKind::TextBytes, size));
    }
    Ok(())
}

fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml.get(..end)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| invalid_error("invalid flat ODP XML event boundary"))
}

fn namespace_of(namespace: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match namespace {
        ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
        ResolveResult::Unbound | ResolveResult::Unknown(_) => None,
    }
}

fn inherited_fragment_namespace(namespace: Option<Vec<u8>>, qualified: &[u8]) -> Option<Vec<u8>> {
    namespace.or_else(|| {
        qualified
            .starts_with(b"draw:")
            .then(|| DRAW.to_vec())
            .or_else(|| {
                qualified
                    .starts_with(b"presentation:")
                    .then(|| PRESENTATION.to_vec())
            })
    })
}

fn is(namespace: Option<&[u8]>, local: &[u8], expected_ns: &[u8], expected_local: &[u8]) -> bool {
    namespace == Some(expected_ns) && local == expected_local
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> CoreError {
    CoreError::InvalidFormat(message.into())
}
