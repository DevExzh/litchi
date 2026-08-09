//! Bounded, source-bound flat `OpenDocument` presentation support.

use crate::codec::Parser;
use crate::model::{Settings, Slide, declaration, page_layout, page_metadata, settings};
use litchi_core::xml::escape_xml;
use litchi_core::{Error, Result};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::BTreeMap;
use std::ops::Range;
use std::sync::Arc;
use xml_minifier::audit;

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const MIMETYPE: &str = "application/vnd.oasis.opendocument.presentation";
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
const MAX_SLIDES: usize = 65_536;

/// An immutable flat `OpenDocument` presentation (`.fodp`).
#[allow(
    clippy::module_name_repetitions,
    reason = "FlatPresentation is the established public type name; renaming it would break the crate API"
)]
#[derive(Clone)]
pub struct FlatPresentation {
    bytes: Arc<[u8]>,
    xml: Arc<str>,
    page_ranges: Arc<[Range<usize>]>,
    slides: Arc<[Slide]>,
}

impl FlatPresentation {
    /// Open bounded UTF-8 flat-presentation bytes.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.len() > MAX_XML_BYTES {
            return invalid("flat ODP exceeds the 64 MiB input limit");
        }
        let xml = String::from_utf8(bytes.clone())
            .map_err(|error| Error::InvalidFormat(format!("flat ODP is not UTF-8: {error}")))?;
        let page_ranges = scan_flat(&xml)?;
        let slides = Parser::parse_slides_with_styles(&xml, Some(&xml))?;
        if slides.len() != page_ranges.len() {
            return invalid("flat ODP slide projection does not match its page structure");
        }
        Ok(Self {
            bytes: Arc::from(bytes),
            xml: Arc::from(xml),
            page_ranges: Arc::from(page_ranges),
            slides: Arc::from(slides),
        })
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

    /// Create an immutable source-bound editing snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Snapshot {
        Snapshot {
            presentation: self.clone(),
        }
    }
}

/// Immutable source identity for flat-presentation edits.
#[derive(Clone)]
pub struct Snapshot {
    presentation: FlatPresentation,
}

impl Snapshot {
    /// Parse owned source bytes as a flat editing snapshot.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Ok(FlatPresentation::from_bytes(bytes)?.snapshot())
    }

    /// Borrow the exact source bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.presentation.as_bytes()
    }

    /// Borrow the immutable slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        self.presentation.slides()
    }

    /// Start an isolated edit transaction.
    #[must_use]
    pub fn transaction(&self) -> Transaction {
        Transaction {
            source: self.clone(),
            draft: self.slides().to_vec(),
            replacements: BTreeMap::new(),
        }
    }

    /// Materialize the ordinary flat reader.
    #[must_use]
    pub fn to_presentation(&self) -> FlatPresentation {
        self.presentation.clone()
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

/// Isolated replacement transaction for a flat presentation.
pub struct Transaction {
    source: Snapshot,
    draft: Vec<Slide>,
    replacements: BTreeMap<usize, String>,
}

impl Transaction {
    /// Borrow the staged slide projection.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.draft
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
        check_text(title, text)?;
        let Some(index) = select(&self.draft, selector.into())? else {
            return Ok(None);
        };
        if self.draft[index].title.as_deref() == Some(title) && self.draft[index].text == text {
            return Ok(Some(()));
        }
        let range = &self.source.presentation.page_ranges[index];
        let source_page = &self.source.presentation.xml[range.clone()];
        let replacement = replace_page(source_page, title, text)?;
        self.replacements.insert(index, replacement);
        let slide = &mut self.draft[index];
        slide.title = (!title.is_empty()).then(|| title.to_owned());
        text.clone_into(&mut slide.text);
        slide.shapes.clear();
        slide.animations.clear();
        slide.legacy_animation = None;
        Ok(Some(()))
    }

    /// Validate, serialize, reparse, and atomically publish the staged bytes.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn commit(self) -> Result<Commit> {
        if self.replacements.is_empty() {
            return Ok(Commit::unchanged(self.source));
        }
        let xml = &self.source.presentation.xml;
        let mut capacity = xml.len();
        for (index, replacement) in &self.replacements {
            let original = self.source.presentation.page_ranges[*index].len();
            capacity = capacity
                .checked_sub(original)
                .and_then(|size| size.checked_add(replacement.len()))
                .ok_or_else(|| invalid_error("flat ODP output size overflow"))?;
        }
        if capacity > MAX_XML_BYTES {
            return invalid("flat ODP commit exceeds the 64 MiB output limit");
        }
        let mut output = String::with_capacity(capacity);
        let mut cursor = 0;
        for (index, range) in self.source.presentation.page_ranges.iter().enumerate() {
            push_bounded(&mut output, &xml[cursor..range.start])?;
            if let Some(replacement) = self.replacements.get(&index) {
                push_bounded(&mut output, replacement)?;
            } else {
                push_bounded(&mut output, &xml[range.clone()])?;
            }
            cursor = range.end;
        }
        push_bounded(&mut output, &xml[cursor..])?;
        validate_compact_xml(&output)?;
        let presentation = FlatPresentation::from_bytes(output.into_bytes())?;
        if presentation.slides() != self.draft {
            return invalid("flat ODP commit readback differs from the staged slide model");
        }
        let snapshot = presentation.snapshot();
        let patch = Patch {
            before: self.source,
            after: snapshot.clone(),
        };
        Ok(Commit {
            snapshot,
            patch,
            changed: true,
        })
    }
}

/// Successful flat transaction publication.
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
    changed: bool,
}

impl Commit {
    fn unchanged(snapshot: Snapshot) -> Self {
        Self {
            patch: Patch {
                before: snapshot.clone(),
                after: snapshot.clone(),
            },
            snapshot,
            changed: false,
        }
    }

    /// Whether bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.changed
    }

    /// Borrow the published snapshot.
    #[must_use]
    pub fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consume the commit into its snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }
}

/// Exact-source-checked reversible flat-presentation patch.
#[derive(Clone)]
pub struct Patch {
    before: Snapshot,
    after: Snapshot,
}

impl Patch {
    /// Apply only to the exact source bytes from which this patch was made.
    ///
    /// # Errors
    /// Returns an error when a configured limit is exceeded or the package cannot be serialized.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.bytes() != self.before.bytes() {
            return invalid("flat ODP patch source does not match");
        }
        Ok(self.after.clone())
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
        self.before.bytes() == self.after.bytes()
    }
}

fn scan_flat(xml: &str) -> Result<Vec<Range<usize>>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<(Option<Vec<u8>>, Vec<u8>, usize)> = Vec::new();
    let mut pages = Vec::new();
    let mut body_count = 0usize;
    let mut presentation_count = 0usize;
    let mut root_seen = false;
    loop {
        let (namespace, event) = {
            let (resolved, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::InvalidFormat(format!("invalid flat ODP XML: {error}")))?;
            (namespace_of(&resolved), event)
        };
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid_error("flat ODP XML position exceeds platform limits"))?;
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
                    validate_page_parent_and_limit(&stack, pages.len())?;
                }
                if stack.len() == MAX_DEPTH {
                    return invalid("flat ODP exceeds the XML nesting limit");
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
                    validate_page_parent_and_limit(&stack, pages.len())?;
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
                    .map_err(|error| Error::InvalidFormat(error.to_string()))?
                    .trim()
                    .is_empty()
                {
                    return invalid("flat ODP has text outside its root element");
                }
            },
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
) -> Result<()> {
    if stack.len() != 3
        || !stack.get(2).is_some_and(|(namespace, local, _)| {
            is(namespace.as_deref(), local, OFFICE, b"presentation")
        })
    {
        return invalid("draw:page is outside office:presentation");
    }
    if page_count >= MAX_SLIDES {
        return invalid("flat ODP exceeds the slide-count limit");
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
                Error::InvalidFormat(format!("invalid flat ODP attribute: {error}"))
            })?;
            let (resolved, name) = reader.resolver().resolve_attribute(parsed_attribute.key);
            if name.as_ref() == b"mimetype"
                && matches!(resolved, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE)
            {
                mimetype = Some(
                    parsed_attribute
                        .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
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
            .map_err(|error| Error::InvalidFormat(format!("invalid flat ODP page: {error}")))?;
        let namespace = namespace_of(&resolved);
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid_error("flat ODP XML position exceeds platform limits"))?;
        match event {
            Event::Start(element) => {
                let start = event_start(page, end)?;
                let local = element.local_name().as_ref().to_vec();
                let inherited = inherited_fragment_namespace(namespace, element.name().as_ref());
                if stack.len() == 1 && !is_modeled_direct_page_child(inherited.as_deref(), &local) {
                    return Err(Error::Unsupported(
                        "flat ODP page has an unsupported direct child; replacement refused"
                            .to_string(),
                    ));
                }
                stack.push((inherited, local, start));
            },
            Event::Empty(element) => {
                let inherited = inherited_fragment_namespace(namespace, element.name().as_ref());
                let local_name = element.local_name();
                let local = local_name.as_ref();
                if stack.len() == 1 && !is_modeled_direct_page_child(inherited.as_deref(), local) {
                    return Err(Error::Unsupported(
                        "flat ODP page has an unsupported direct child; replacement refused"
                            .to_string(),
                    ));
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

fn push_bounded(output: &mut String, value: &str) -> Result<()> {
    let size = output
        .len()
        .checked_add(value.len())
        .ok_or_else(|| invalid_error("flat ODP output size overflow"))?;
    if size > MAX_XML_BYTES {
        return invalid("flat ODP commit exceeds the 64 MiB output limit");
    }
    output.push_str(value);
    Ok(())
}

fn validate_compact_xml(xml: &str) -> Result<()> {
    let limits = audit::Limits::new(
        MAX_XML_BYTES,
        MAX_DEPTH,
        1_000_000,
        250_000,
        MAX_TEXT_BYTES,
        MAX_XML_BYTES,
    )
    .map_err(|source| invalid_error(format!("invalid flat ODP XML audit limits: {source}")))?;
    let _report = audit::verify(xml.as_bytes(), limits).map_err(|source| match source {
        audit::Error::NotCompact(_) => {
            Error::Unsupported(format!("flat ODP XML is not compact: {source}"))
        },
        audit::Error::Limit { .. }
        | audit::Error::Encoding { .. }
        | audit::Error::Malformed { .. }
        | audit::Error::Doctype { .. }
        | audit::Error::Allocation
        | _ => Error::InvalidFormat(format!("flat ODP XML failed audit: {source}")),
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

fn check_text(title: &str, text: &str) -> Result<()> {
    let size = title
        .len()
        .checked_add(text.len())
        .ok_or_else(|| invalid_error("flat ODP slide text size overflow"))?;
    if size > MAX_TEXT_BYTES {
        return invalid("flat ODP slide text exceeds the 16 MiB limit");
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

fn invalid_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}
