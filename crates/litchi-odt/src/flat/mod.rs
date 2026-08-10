//! Immutable flat ODT snapshots and failure-atomic text edits.

use std::ffi::OsString;
use std::io::{Read, Write};
use std::ops::Range;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{Error, Metadata, Resource, ResourceLimit, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::elements::parser::{OrderElement, Parser};
use crate::elements::table::Table;
use crate::generic::{Family, FlatDocument};

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const MAX_TEMP_CREATE_ATTEMPTS: usize = 128;
static NEXT_TEMP_ID: AtomicU64 = AtomicU64::new(0);

/// Finite resource limits for opening and editing a flat ODT snapshot.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct Limits {
    /// Maximum flat XML document size.
    max_document_bytes: usize,
    /// Maximum replacement text size for one paragraph edit.
    max_paragraph_text_bytes: usize,
    /// Maximum number of staged paragraph edits.
    max_edits: usize,
    /// Maximum XML nesting depth inspected by an edit.
    max_xml_depth: usize,
}

impl Limits {
    /// Default flat XML document byte limit.
    pub const DEFAULT_MAX_DOCUMENT_BYTES: usize = 16 * 1024 * 1024;
    /// Default replacement paragraph UTF-8 byte limit.
    pub const DEFAULT_MAX_PARAGRAPH_TEXT_BYTES: usize = 256 * 1024;
    /// Default staged paragraph edit limit.
    pub const DEFAULT_MAX_EDITS: usize = 256;
    /// Default inspected XML nesting limit.
    pub const DEFAULT_MAX_XML_DEPTH: usize = 256;
    /// Hard safety ceiling for flat XML document bytes.
    pub const HARD_MAX_DOCUMENT_BYTES: usize = 64 * 1024 * 1024;
    /// Hard safety ceiling for one replacement paragraph's UTF-8 bytes.
    pub const HARD_MAX_PARAGRAPH_TEXT_BYTES: usize = 1024 * 1024;
    /// Hard safety ceiling for staged paragraph edits.
    pub const HARD_MAX_EDITS: usize = 1_024;
    /// Hard safety ceiling for inspected XML nesting depth.
    pub const HARD_MAX_XML_DEPTH: usize = 4_096;

    /// Returns the production-safe default limits.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_document_bytes: Self::DEFAULT_MAX_DOCUMENT_BYTES,
            max_paragraph_text_bytes: Self::DEFAULT_MAX_PARAGRAPH_TEXT_BYTES,
            max_edits: Self::DEFAULT_MAX_EDITS,
            max_xml_depth: Self::DEFAULT_MAX_XML_DEPTH,
        }
    }

    /// Sets the flat XML byte limit without exceeding the hard safety ceiling.
    pub fn with_max_document_bytes(mut self, value: usize) -> Result<Self> {
        check_limit(
            value,
            Self::HARD_MAX_DOCUMENT_BYTES,
            Resource::InputBytes,
            "flat ODT configured document bytes",
        )?;
        self.max_document_bytes = value;
        Ok(self)
    }

    /// Sets the paragraph replacement byte limit without exceeding the hard ceiling.
    pub fn with_max_paragraph_text_bytes(mut self, value: usize) -> Result<Self> {
        check_limit(
            value,
            Self::HARD_MAX_PARAGRAPH_TEXT_BYTES,
            Resource::InputBytes,
            "flat ODT configured paragraph text bytes",
        )?;
        self.max_paragraph_text_bytes = value;
        Ok(self)
    }

    /// Sets the edit count limit without exceeding the hard safety ceiling.
    pub fn with_max_edits(mut self, value: usize) -> Result<Self> {
        check_limit(
            value,
            Self::HARD_MAX_EDITS,
            Resource::Objects,
            "flat ODT configured edit count",
        )?;
        self.max_edits = value;
        Ok(self)
    }

    /// Sets the XML nesting limit without exceeding the hard safety ceiling.
    pub fn with_max_xml_depth(mut self, value: usize) -> Result<Self> {
        check_limit(
            value,
            Self::HARD_MAX_XML_DEPTH,
            Resource::Depth,
            "flat ODT configured XML depth",
        )?;
        self.max_xml_depth = value;
        Ok(self)
    }

    /// Returns the configured flat XML byte limit.
    #[must_use]
    pub const fn max_document_bytes(self) -> usize {
        self.max_document_bytes
    }

    /// Returns the configured paragraph replacement byte limit.
    #[must_use]
    pub const fn max_paragraph_text_bytes(self) -> usize {
        self.max_paragraph_text_bytes
    }

    /// Returns the configured edit count limit.
    #[must_use]
    pub const fn max_edits(self) -> usize {
        self.max_edits
    }

    /// Returns the configured XML nesting limit.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.max_xml_depth
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

/// Immutable, cheap-to-share flat OpenDocument Text snapshot.
#[derive(Clone)]
pub struct Document {
    inner: Arc<FlatDocument>,
    limits: Limits,
}

impl Document {
    /// Opens and validates a flat ODT document with default finite limits.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        Self::open_with_limits(path, Limits::default())
    }

    /// Opens and validates a flat ODT document with caller-selected finite limits.
    pub fn open_with_limits(path: impl AsRef<Path>, limits: Limits) -> Result<Self> {
        Self::from_reader_with_limits(std::fs::File::open(path)?, limits)
    }

    /// Reads and validates a flat ODT document with default finite limits.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        Self::from_reader_with_limits(reader, Limits::default())
    }

    /// Reads and validates a flat ODT document with caller-selected finite limits.
    pub fn from_reader_with_limits(mut reader: impl Read, limits: Limits) -> Result<Self> {
        validate_limits(limits)?;
        let read_limit = limits.max_document_bytes.checked_add(1).ok_or_else(|| {
            resource_limit_error(
                Resource::InputBytes,
                usize::MAX,
                limits.max_document_bytes,
                "flat ODT input",
            )
        })?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(read_limit.min(64 * 1024))
            .map_err(|source| Error::Allocation {
                resource: "flat ODT input",
                source,
            })?;
        let read_limit = u64::try_from(read_limit).map_err(|_error| {
            resource_limit_error(
                Resource::InputBytes,
                read_limit,
                limits.max_document_bytes,
                "flat ODT input",
            )
        })?;
        reader.by_ref().take(read_limit).read_to_end(&mut bytes)?;
        Self::from_bytes_with_limits(bytes, limits)
    }

    /// Validates owned flat ODT bytes with default finite limits.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        Self::from_bytes_with_limits(bytes, Limits::default())
    }

    /// Validates owned flat ODT bytes with caller-selected finite limits.
    pub fn from_bytes_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        validate_limits(limits)?;
        if bytes.len() > limits.max_document_bytes {
            return Err(resource_limit_error(
                Resource::InputBytes,
                bytes.len(),
                limits.max_document_bytes,
                "flat ODT input",
            ));
        }
        let document = FlatDocument::from_bytes(bytes)?;
        if document.family() != Family::Text {
            return invalid("flat document is not an OpenDocument Text document");
        }
        Ok(Self {
            inner: Arc::new(document),
            limits,
        })
    }

    /// Returns the validated flat XML without reconstructing it.
    #[must_use]
    pub fn xml(&self) -> &str {
        self.inner.xml()
    }

    /// Returns the exact snapshot bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.inner.as_bytes()
    }

    /// Clones the exact snapshot bytes.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.inner.to_bytes()
    }

    /// Returns common flat-document metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.inner.metadata()
    }

    /// Returns the unnamed fallback page layout, when one is declared.
    pub fn default_page_layout(&self) -> Result<Option<crate::page_layout::PageLayout>> {
        self.inner.default_page_layout()
    }

    /// Discovers inert inline and linked embedded objects without fetching them.
    pub fn embedded_objects(&self) -> Result<Vec<crate::Object>> {
        self.inner.embedded_objects()
    }

    /// Inspects ordered variable declarations without evaluating fields.
    pub fn variable_declarations(&self) -> Result<crate::variable_declaration::Declarations> {
        self.inner.variable_declarations()
    }

    /// Inspects inert document script and macro metadata without executing it.
    pub fn document_scripts(&self) -> Result<Option<crate::document_scripts::Scripts>> {
        crate::package::scripts::document_scripts(self.xml())
    }

    /// Returns the finite resource limits attached to this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Parses modeled top-level text elements in document order.
    pub fn elements(&self) -> Result<Vec<OrderElement>> {
        Parser::parse_elements_in_order(self.xml())
    }

    /// Parses modeled tables in document order.
    pub fn tables(&self) -> Result<Vec<Table>> {
        Parser::parse_tables_in_order(self.xml())
    }

    /// Extracts modeled paragraph and heading text.
    pub fn text(&self) -> Result<String> {
        let mut result = String::new();
        for element in self.elements()? {
            let text = match element {
                OrderElement::Paragraph(paragraph) => paragraph.text()?,
                OrderElement::Heading(heading) => heading.text()?,
                _ => continue,
            };
            result
                .try_reserve(text.len())
                .map_err(|source| Error::Allocation {
                    resource: "flat ODT extracted text",
                    source,
                })?;
            result.push_str(&text);
        }
        Ok(result)
    }

    /// Starts a detached failure-atomic edit over this snapshot.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: self.clone(),
            operations: Vec::new(),
            paragraph_sites: None,
        }
    }

    /// Atomically replaces a filesystem destination with these exact bytes.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        let path = path.as_ref();
        let parent = path
            .parent()
            .filter(|value| !value.as_os_str().is_empty())
            .unwrap_or_else(|| Path::new("."));
        validate_destination(path)?;
        let (temporary, mut file) = create_owned_sibling_temp(parent)?;
        let write_result = (|| -> Result<()> {
            file.write_all(self.as_bytes())?;
            file.flush()?;
            file.sync_all()?;
            let expected_len = u64::try_from(self.as_bytes().len()).map_err(|_error| {
                resource_limit_error(
                    Resource::OutputBytes,
                    self.as_bytes().len(),
                    self.limits.max_document_bytes,
                    "flat ODT output",
                )
            })?;
            if file.metadata()?.len() != expected_len {
                return Err(ResourceLimit {
                    resource: Resource::OutputBytes,
                    observed: file.metadata()?.len(),
                    limit: expected_len,
                    scope: Arc::from("flat ODT temporary output"),
                }
                .into());
            }
            drop(file);
            publish_owned_temp(&temporary, path)?;
            sync_parent(parent)?;
            Ok(())
        })();
        if write_result.is_err() {
            drop(std::fs::remove_file(&temporary));
        }
        write_result
    }
}

/// Detached, failure-atomic flat ODT edit.
pub struct Edit {
    source: Document,
    operations: Vec<ParagraphEdit>,
    paragraph_sites: Option<Vec<ParagraphSite>>,
}

impl Edit {
    /// Stages replacement of a zero-based top-level paragraph's text payload.
    pub fn update_paragraph(&mut self, index: usize, text: &str) -> Result<Option<&mut Self>> {
        if self.operations.len() >= self.source.limits.max_edits {
            return Err(resource_limit_error(
                Resource::Objects,
                self.operations.len().saturating_add(1),
                self.source.limits.max_edits,
                "flat ODT paragraph edits",
            ));
        }
        if text.len() > self.source.limits.max_paragraph_text_bytes {
            return Err(resource_limit_error(
                Resource::InputBytes,
                text.len(),
                self.source.limits.max_paragraph_text_bytes,
                "flat ODT paragraph text",
            ));
        }
        validate_xml_text(&text)?;
        if self
            .operations
            .iter()
            .any(|operation| operation.index == index)
        {
            return invalid("flat ODT edit contains a duplicate paragraph selector");
        }
        if self.paragraph_sites.is_none() {
            self.paragraph_sites = Some(index_direct_paragraphs(
                self.source.xml(),
                self.source.limits,
            )?);
        }
        let site = self
            .paragraph_sites
            .as_ref()
            .and_then(|sites| sites.get(index))
            .cloned();
        let Some(site) = site else {
            return Ok(None);
        };
        let ParagraphSite::Plain(range) = site else {
            return invalid("flat ODT paragraph contains structured or opaque inline markup");
        };
        let mut owned_text = String::new();
        owned_text
            .try_reserve_exact(text.len())
            .map_err(|source| Error::Allocation {
                resource: "flat ODT paragraph text",
                source,
            })?;
        owned_text.push_str(text);
        self.operations
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "flat ODT edit operations",
                source,
            })?;
        self.operations.push(ParagraphEdit {
            index,
            text: owned_text,
            range,
        });
        Ok(Some(self))
    }

    /// Validates and publishes all staged operations as one commit.
    pub fn commit(mut self) -> Result<Commit> {
        if self.operations.is_empty() {
            let document = self.source.clone();
            let patch = Patch {
                before: self.source,
                after: document.clone(),
            };
            return Ok(Commit {
                document,
                patch,
                diagnostics: Vec::new(),
            });
        }
        self.operations
            .sort_unstable_by_key(|operation| operation.index);
        let candidate =
            render_paragraph_edits(self.source.xml(), &self.operations, self.source.limits)?;
        let compact_limits = litchi_odf_common::compact_xml::Limits::new(
            self.source.limits.max_document_bytes,
            self.source.limits.max_xml_depth,
        )
        .map_err(Error::from)?;
        litchi_odf_common::compact_xml::validate_with_limits(candidate.as_bytes(), compact_limits)
            .map_err(Error::from)?;
        let document =
            Document::from_bytes_with_limits(candidate.into_bytes(), self.source.limits)?;
        semantic_readback(&document, &self.operations)?;
        let patch = Patch {
            before: self.source,
            after: document.clone(),
        };
        Ok(Commit {
            document,
            patch,
            diagnostics: Vec::new(),
        })
    }
}

struct ParagraphEdit {
    index: usize,
    text: String,
    range: Range<usize>,
}

#[derive(Clone)]
enum ParagraphSite {
    Plain(Range<usize>),
    Opaque,
}

/// A validated flat ODT commit.
pub struct Commit {
    document: Document,
    patch: Patch,
    diagnostics: Vec<Diagnostic>,
}

impl Commit {
    /// Returns the committed immutable snapshot.
    #[must_use]
    pub const fn document(&self) -> &Document {
        &self.document
    }

    /// Returns the reversible exact-byte patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Returns deterministic diagnostics produced during commit validation.
    #[must_use]
    pub fn diagnostics(&self) -> &[Diagnostic] {
        &self.diagnostics
    }

    /// Consumes the commit and returns its snapshot.
    #[must_use]
    pub fn into_document(self) -> Document {
        self.document
    }
}

/// A structured commit diagnostic.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub struct Diagnostic {
    /// Stable diagnostic code.
    pub code: &'static str,
    /// Human-readable diagnostic message.
    pub message: String,
}

/// Reversible, exact-byte flat ODT patch.
#[derive(Clone)]
pub struct Patch {
    before: Document,
    after: Document,
}

impl Patch {
    /// Applies this patch only to its exact immutable source snapshot.
    pub fn apply(&self, source: &Document) -> Result<Document> {
        if source.as_bytes() != self.before.as_bytes() {
            return invalid("flat ODT patch source does not match its expected snapshot");
        }
        Ok(self.after.clone())
    }

    /// Returns a patch that restores the exact source bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

fn render_paragraph_edits(
    xml: &str,
    operations: &[ParagraphEdit],
    limits: Limits,
) -> Result<String> {
    let mut output_len = xml.len();
    for operation in operations {
        let range = &operation.range;
        let escaped_len = escaped_xml_text_len(&operation.text)?;
        output_len = output_len
            .checked_sub(range.len())
            .and_then(|value| value.checked_add(escaped_len))
            .ok_or_else(|| {
                resource_limit_error(
                    Resource::OutputBytes,
                    usize::MAX,
                    limits.max_document_bytes,
                    "flat ODT edited output",
                )
            })?;
    }
    if output_len > limits.max_document_bytes {
        return Err(resource_limit_error(
            Resource::OutputBytes,
            output_len,
            limits.max_document_bytes,
            "flat ODT edited output",
        ));
    }
    let mut output = String::new();
    output
        .try_reserve_exact(output_len)
        .map_err(|source| Error::Allocation {
            resource: "flat ODT paragraph edit",
            source,
        })?;
    let mut cursor = 0usize;
    for operation in operations {
        let range = &operation.range;
        output.push_str(&xml[cursor..range.start]);
        push_escaped_xml_text(&mut output, &operation.text);
        cursor = range.end;
    }
    output.push_str(&xml[cursor..]);
    Ok(output)
}

fn index_direct_paragraphs(xml: &str, limits: Limits) -> Result<Vec<ParagraphSite>> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut office_text_depth = None;
    let mut sites = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid flat ODT XML: {error}")))?;
        let is_office = is_bound(&namespace, OFFICE_NAMESPACE);
        let is_text = is_bound(&namespace, TEXT_NAMESPACE);
        let event = event.into_owned();
        let event_end = usize::try_from(reader.buffer_position()).map_err(|_error| {
            resource_limit_error(
                Resource::InputBytes,
                usize::MAX,
                usize::MAX - 1,
                "flat ODT XML position",
            )
        })?;

        match event {
            Event::Start(ref element) => {
                depth = checked_depth(depth, limits)?;
                let local = element.local_name();
                if is_office && local.as_ref() == b"text" {
                    office_text_depth = Some(depth);
                } else if office_text_depth.is_some()
                    && office_text_depth.and_then(|value| value.checked_add(1)) == Some(depth)
                    && is_text
                    && local.as_ref() == b"p"
                {
                    buffer.clear();
                    let site =
                        classify_paragraph_end(&mut reader, &mut buffer, event_end, depth, limits)?;
                    sites.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "flat ODT paragraph selector index",
                        source,
                    })?;
                    sites.push(site);
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        Error::InvalidFormat("flat ODT XML depth underflow".to_string())
                    })?;
                }
            },
            Event::Empty(ref element) => {
                let local = element.local_name();
                if office_text_depth == Some(depth) && is_text && local.as_ref() == b"p" {
                    sites.try_reserve(1).map_err(|source| Error::Allocation {
                        resource: "flat ODT paragraph selector index",
                        source,
                    })?;
                    sites.push(ParagraphSite::Opaque);
                }
            },
            Event::End(ref element) => {
                let local = element.local_name();
                if office_text_depth == Some(depth) && is_office && local.as_ref() == b"text" {
                    office_text_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("flat ODT XML depth underflow".to_string())
                })?;
            },
            Event::DocType(_) => return invalid("flat ODT edits refuse documents with a doctype"),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    Ok(sites)
}

fn classify_paragraph_end(
    reader: &mut NsReader<&[u8]>,
    buffer: &mut Vec<u8>,
    content_start: usize,
    outer_depth: usize,
    limits: Limits,
) -> Result<ParagraphSite> {
    let mut nested_depth = 0usize;
    let mut opaque = false;
    loop {
        let event_start = usize::try_from(reader.buffer_position()).map_err(|_error| {
            resource_limit_error(
                Resource::InputBytes,
                usize::MAX,
                usize::MAX - 1,
                "flat ODT XML position",
            )
        })?;
        let (_, event) = reader
            .read_resolved_event_into(buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid flat ODT XML: {error}")))?;
        match event {
            Event::End(_) if nested_depth == 0 => {
                return Ok(if opaque {
                    ParagraphSite::Opaque
                } else {
                    ParagraphSite::Plain(content_start..event_start)
                });
            },
            Event::End(_) => nested_depth -= 1,
            Event::Text(_) | Event::CData(_) | Event::GeneralRef(_) => {},
            Event::Start(_) => {
                opaque = true;
                nested_depth = nested_depth.checked_add(1).ok_or_else(|| {
                    resource_limit_error(
                        Resource::Depth,
                        usize::MAX,
                        limits.max_xml_depth,
                        "flat ODT paragraph selector index",
                    )
                })?;
                let observed = outer_depth.saturating_add(nested_depth);
                if observed > limits.max_xml_depth {
                    return Err(resource_limit_error(
                        Resource::Depth,
                        observed,
                        limits.max_xml_depth,
                        "flat ODT paragraph selector index",
                    ));
                }
            },
            Event::Empty(_) | Event::Comment(_) | Event::PI(_) | Event::Decl(_) => opaque = true,
            Event::DocType(_) => return invalid("flat ODT edits refuse documents with a doctype"),
            Event::Eof => return invalid("unterminated flat ODT paragraph"),
        }
        buffer.clear();
    }
}

fn semantic_readback(document: &Document, operations: &[ParagraphEdit]) -> Result<()> {
    let sites = index_direct_paragraphs(document.xml(), document.limits)?;
    for operation in operations {
        let Some(ParagraphSite::Plain(range)) = sites.get(operation.index) else {
            return invalid("committed flat ODT paragraph selector did not round-trip");
        };
        let value =
            quick_xml::escape::unescape(&document.xml()[range.clone()]).map_err(|error| {
                Error::InvalidFormat(format!("invalid committed paragraph: {error}"))
            })?;
        if value.as_ref() != operation.text {
            return invalid("committed flat ODT paragraph value did not round-trip");
        }
    }
    Ok(())
}

fn escaped_xml_text_len(text: &str) -> Result<usize> {
    let extra = text
        .bytes()
        .filter(|byte| matches!(byte, b'&' | b'<' | b'>'))
        .count()
        .checked_mul(4)
        .ok_or_else(|| {
            resource_limit_error(
                Resource::OutputBytes,
                usize::MAX,
                Limits::HARD_MAX_DOCUMENT_BYTES,
                "flat ODT escaped paragraph text",
            )
        })?;
    let capacity = text.len().checked_add(extra).ok_or_else(|| {
        resource_limit_error(
            Resource::OutputBytes,
            usize::MAX,
            Limits::HARD_MAX_DOCUMENT_BYTES,
            "flat ODT escaped paragraph text",
        )
    })?;
    Ok(capacity)
}

fn push_escaped_xml_text(output: &mut String, text: &str) {
    for character in text.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(character),
        }
    }
}

fn validate_xml_text(text: &str) -> Result<()> {
    if text.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        return invalid("flat ODT paragraph text contains an invalid XML character");
    }
    Ok(())
}

fn validate_limits(limits: Limits) -> Result<()> {
    check_limit(
        limits.max_document_bytes,
        Limits::HARD_MAX_DOCUMENT_BYTES,
        Resource::InputBytes,
        "flat ODT configured document bytes",
    )?;
    check_limit(
        limits.max_paragraph_text_bytes,
        Limits::HARD_MAX_PARAGRAPH_TEXT_BYTES,
        Resource::InputBytes,
        "flat ODT configured paragraph text bytes",
    )?;
    check_limit(
        limits.max_edits,
        Limits::HARD_MAX_EDITS,
        Resource::Objects,
        "flat ODT configured edit count",
    )?;
    check_limit(
        limits.max_xml_depth,
        Limits::HARD_MAX_XML_DEPTH,
        Resource::Depth,
        "flat ODT configured XML depth",
    )?;
    Ok(())
}

fn check_limit(
    value: usize,
    hard_max: usize,
    resource: Resource,
    scope: &'static str,
) -> Result<()> {
    if value == 0 || value > hard_max {
        return Err(resource_limit_error(resource, value, hard_max, scope));
    }
    Ok(())
}

fn validate_destination(path: &Path) -> Result<()> {
    match std::fs::symlink_metadata(path) {
        Ok(metadata) if metadata.file_type().is_symlink() => Err(resource_limit_error(
            Resource::Objects,
            1,
            0,
            "flat ODT symbolic-link destination",
        )),
        Ok(metadata) if !metadata.is_file() => Err(resource_limit_error(
            Resource::Objects,
            1,
            0,
            "flat ODT non-file destination",
        )),
        Ok(_) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) => Err(error.into()),
    }
}

fn create_owned_sibling_temp(parent: &Path) -> Result<(PathBuf, std::fs::File)> {
    for _ in 0..MAX_TEMP_CREATE_ATTEMPTS {
        let id = NEXT_TEMP_ID
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |value| {
                value.checked_add(1)
            })
            .map_err(|_error| {
                resource_limit_error(
                    Resource::Work,
                    usize::MAX,
                    usize::MAX - 1,
                    "flat ODT temporary identifiers",
                )
            })?;
        let temporary_name = OsString::from(format!(".litchi-{:x}-{id:x}.tmp", std::process::id()));
        let temporary = parent.join(temporary_name);
        match std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temporary)
        {
            Ok(file) => return Ok((temporary, file)),
            Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => {},
            Err(error) => return Err(error.into()),
        }
    }
    Err(resource_limit_error(
        Resource::Work,
        MAX_TEMP_CREATE_ATTEMPTS.saturating_add(1),
        MAX_TEMP_CREATE_ATTEMPTS,
        "flat ODT sibling temporary collisions",
    ))
}

#[cfg(unix)]
fn publish_owned_temp(temporary: &Path, destination: &Path) -> Result<()> {
    std::fs::rename(temporary, destination)?;
    Ok(())
}

#[cfg(windows)]
fn publish_owned_temp(_temporary: &Path, _destination: &Path) -> Result<()> {
    Err(Error::Unsupported(
        "atomic flat ODT publication is unavailable on Windows".to_string(),
    ))
}

#[cfg(not(any(unix, windows)))]
fn publish_owned_temp(_temporary: &Path, _destination: &Path) -> Result<()> {
    Err(Error::Unsupported(
        "atomic flat ODT publication is unavailable on this platform".to_string(),
    ))
}

#[cfg(unix)]
fn sync_parent(parent: &Path) -> Result<()> {
    std::fs::File::open(parent)?.sync_all()?;
    Ok(())
}

#[cfg(not(unix))]
fn sync_parent(_parent: &Path) -> Result<()> {
    Ok(())
}

fn resource_limit_error(
    resource: Resource,
    observed: usize,
    limit: usize,
    scope: &'static str,
) -> Error {
    ResourceLimit {
        resource,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        limit: u64::try_from(limit).unwrap_or(u64::MAX),
        scope: Arc::from(scope),
    }
    .into()
}

fn checked_depth(depth: usize, limits: Limits) -> Result<usize> {
    let depth = depth.checked_add(1).ok_or_else(|| {
        resource_limit_error(
            Resource::Depth,
            usize::MAX,
            limits.max_xml_depth,
            "flat ODT XML",
        )
    })?;
    if depth > limits.max_xml_depth {
        return Err(resource_limit_error(
            Resource::Depth,
            depth,
            limits.max_xml_depth,
            "flat ODT XML",
        ));
    }
    Ok(depth)
}

fn is_bound(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.as_ref() == expected)
}

fn invalid<T>(message: &str) -> Result<T> {
    Err(Error::InvalidFormat(message.to_string()))
}
