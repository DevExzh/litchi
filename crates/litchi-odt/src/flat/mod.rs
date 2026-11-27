//! Immutable flat (single-file) ODT owners and failure-atomic text edits.
//!
//! [`Document`] is the immutable owner of one `.fodt` document, opened under
//! one [`OpenOptions`] envelope: the source `input_bytes` and the semantic
//! `content_bytes` bound the XML (at most 64 MiB, the compact-XML checker's
//! ceiling), `depth` its nesting, `edits` the paragraphs one edit stages and
//! `text_bytes` (at most 1 MiB) one replacement paragraph. [`Document::edit`]
//! stages an [`Edit`]; [`Edit::commit`] publishes a new [`Document`] with its
//! [`Patch`] and [`Diagnostics`]; [`Document::apply`] replays a patch onto its
//! exact source.

use std::io::Read;
use std::ops::Range;
use std::path::Path;
use std::sync::Arc;

use crate::core::{ResolvedReader, XmlStreamLimits};
use litchi_core::{Error as CoreError, Metadata, Resource, ResourceLimit, Result};
use litchi_odf_common::{Error as OdfError, SemanticLimitKind};
use quick_xml::events::Event;

use crate::elements::parser::{OrderElement, Parser};
use crate::elements::table::Table;
use crate::generic::{Family, FlatDocument};

type FamilyResult<T> = std::result::Result<T, OdfError>;

pub(crate) mod atomic;

/// Whether a commit changed the document bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits and refusals plus the common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;
pub use litchi_odf_common::OpenOptions;

/// A committed flat edit: the published document, its patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Document, Patch>;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";

/// The bounds one flat ODT document is read and edited under, resolved from
/// its [`OpenOptions`] and clamped to the readers' ceilings.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Limits {
    /// Maximum flat XML document size.
    max_document_bytes: usize,
    /// Maximum replacement text size for one paragraph edit.
    max_paragraph_text_bytes: usize,
    /// Maximum number of staged paragraph edits.
    max_edits: usize,
    /// Maximum XML nesting depth inspected by an edit.
    max_xml_depth: usize,
    options: OpenOptions,
}

impl Limits {
    /// Ceiling of the flat XML document bytes (the compact-XML checker's).
    const HARD_MAX_DOCUMENT_BYTES: usize = 64 * 1024 * 1024;
    /// Ceiling of one replacement paragraph's UTF-8 bytes.
    const HARD_MAX_PARAGRAPH_TEXT_BYTES: usize = 1024 * 1024;
    /// Ceiling of the staged paragraph edits.
    const HARD_MAX_EDITS: usize = 1_024;
    /// Ceiling of the inspected XML nesting depth.
    const HARD_MAX_XML_DEPTH: usize = 4_096;

    pub(crate) fn from_options(options: OpenOptions) -> Self {
        let semantic = options.semantic_limits_value();
        let input =
            usize::try_from(options.source_limits_value().max_input_bytes()).unwrap_or(usize::MAX);
        Self {
            max_document_bytes: input
                .min(semantic.usize_value(SemanticLimitKind::ContentBytes))
                .min(Self::HARD_MAX_DOCUMENT_BYTES),
            max_paragraph_text_bytes: semantic
                .usize_value(SemanticLimitKind::TextBytes)
                .min(Self::HARD_MAX_PARAGRAPH_TEXT_BYTES),
            max_edits: semantic
                .usize_value(SemanticLimitKind::Edits)
                .min(Self::HARD_MAX_EDITS),
            max_xml_depth: semantic
                .usize_value(SemanticLimitKind::Depth)
                .min(Self::HARD_MAX_XML_DEPTH),
            options,
        }
    }

    fn input_exceeded(self, observed: usize) -> CoreError {
        let semantic = self.options.semantic_limits_value();
        if let Err(error) = semantic.check_core(SemanticLimitKind::ContentBytes, observed) {
            return error;
        }
        resource_limit_error(
            Resource::InputBytes,
            observed,
            self.max_document_bytes,
            "flat ODT input",
        )
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::from_options(OpenOptions::new())
    }
}

/// The immutable, cheap-to-share owner of one flat OpenDocument Text
/// document.
///
/// A commit publishes a new document and never changes its source. The
/// document keeps the [`OpenOptions`] it was opened with; its edits,
/// replays and republication use them.
#[derive(Clone)]
pub struct Document {
    inner: Arc<FlatDocument>,
    limits: Limits,
}

impl Document {
    /// Opens and validates a flat ODT document with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open(path: impl AsRef<Path>) -> std::result::Result<Self, Error> {
        Self::open_with_options(path, OpenOptions::new())
    }

    /// Opens and validates a flat ODT document under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn open_with_options(
        path: impl AsRef<Path>,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        Self::from_reader_with_options(std::fs::File::open(path)?, options)
    }

    /// Reads and validates a flat ODT document with the default
    /// [`OpenOptions`].
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader(reader: impl Read) -> std::result::Result<Self, Error> {
        Self::from_reader_with_options(reader, OpenOptions::new())
    }

    /// Reads and validates a flat ODT document under caller-selected limits.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`], plus the read failure.
    pub fn from_reader_with_options(
        reader: impl Read,
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        Ok(Self::from_reader_with_limits(
            reader,
            Limits::from_options(options),
        )?)
    }

    fn from_reader_with_limits(reader: impl Read, limits: Limits) -> Result<Self> {
        let maximum = u64::try_from(limits.max_document_bytes).map_err(|_error| {
            resource_limit_error(
                Resource::InputBytes,
                limits.max_document_bytes,
                limits.max_document_bytes,
                "flat ODT input",
            )
        })?;
        let document = FlatDocument::from_reader_with_limit_and_xml_limits(
            reader,
            maximum,
            limits.options.xml_limits_value(),
        )
        .map_err(|error| match error {
            CoreError::ResourceLimit(limit) if limit.resource == Resource::InputBytes => {
                limits.input_exceeded(usize::try_from(limit.observed).unwrap_or(usize::MAX))
            },
            other => other,
        })?;
        if document.family() != Family::Text {
            return invalid("flat document is not an OpenDocument Text document");
        }
        Ok(Self {
            inner: Arc::new(document),
            limits,
        })
    }

    /// Validates flat ODT bytes with the default [`OpenOptions`]; the bytes
    /// are copied.
    ///
    /// # Errors
    ///
    /// As [`Self::from_bytes_with_options`].
    pub fn from_bytes(bytes: &[u8]) -> std::result::Result<Self, Error> {
        Self::from_bytes_with_options(bytes, OpenOptions::new())
    }

    /// Validates flat ODT bytes under caller-selected limits; the bytes are
    /// copied after the size check.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Core`] carrying a resource-limit failure when a
    /// selected dimension is exceeded, or the validation failure of a
    /// malformed or non-text flat document.
    pub fn from_bytes_with_options(
        bytes: &[u8],
        options: OpenOptions,
    ) -> std::result::Result<Self, Error> {
        let limits = Limits::from_options(options);
        if bytes.len() > limits.max_document_bytes {
            return Err(limits.input_exceeded(bytes.len()).into());
        }
        Ok(Self::from_bytes_with_limits(bytes.to_vec(), limits)?)
    }

    pub(crate) fn from_bytes_with_limits(bytes: Vec<u8>, limits: Limits) -> Result<Self> {
        if bytes.len() > limits.max_document_bytes {
            return Err(limits.input_exceeded(bytes.len()));
        }
        let maximum = u64::try_from(limits.max_document_bytes).map_err(|_error| {
            resource_limit_error(
                Resource::InputBytes,
                limits.max_document_bytes,
                limits.max_document_bytes,
                "flat ODT input",
            )
        })?;
        let document = FlatDocument::from_bytes_with_limit_and_xml_limits(
            bytes,
            maximum,
            limits.options.xml_limits_value(),
        )?;
        if document.family() != Family::Text {
            return invalid("flat document is not an OpenDocument Text document");
        }
        Ok(Self {
            inner: Arc::new(document),
            limits,
        })
    }

    /// A packaged copy of the document for the packaged readers: the flat
    /// root's children split into `content.xml`, `styles.xml`, `meta.xml`
    /// and `settings.xml` with their bytes kept
    /// ([`litchi_odf_common::flat::Parts`]). The flat bytes stay this
    /// document's publication; the copy is a read projection.
    ///
    /// # Errors
    ///
    /// Returns an error when the parts cannot be split or the packaged
    /// reader refuses them.
    pub fn to_package(&self) -> std::result::Result<crate::Package, Error> {
        let limits = litchi_odf_common::flat::Limits::from_options(self.options());
        let parts = litchi_odf_common::flat::Parts::split(self.inner.xml().as_bytes(), &limits)?;
        crate::Package::from_bytes_with_options(
            &parts.to_bytes(&limits)?,
            limits.projection_options(),
        )
    }

    /// This document with an edit of its packaged projection
    /// ([`Self::to_package`]) applied: every top-level part element the
    /// edit changed replaces its bytes here, and every other byte is kept
    /// ([`litchi_odf_common::flat::Parts::merge_package_projection`]).
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an edit a flat document cannot hold (a
    /// changed part root, added or removed top-level elements, an added
    /// package member), and the flat reader's error when the result does not
    /// reopen under this document's limits.
    pub fn republish(&self, edited: &crate::Package) -> std::result::Result<Self, Error> {
        let limits = litchi_odf_common::flat::Limits::from_options(self.options());
        let parts = litchi_odf_common::flat::Parts::split(self.as_bytes(), &limits)?;
        let bytes = parts.merge_package_projection(self.as_bytes(), edited.as_bytes(), &limits)?;
        Ok(Self::from_bytes_with_limits(bytes, self.limits)?)
    }

    /// Returns the validated flat XML without reconstructing it.
    #[must_use]
    pub fn xml(&self) -> &str {
        self.inner.xml()
    }

    /// Returns the validated flat document family.
    #[must_use]
    pub fn family(&self) -> Family {
        self.inner.family()
    }

    /// Returns whether this flat text snapshot uses the text-template MIME.
    #[must_use]
    pub fn is_template(&self) -> bool {
        self.inner.is_template()
    }

    /// Returns the canonical root `office:mimetype` value.
    #[must_use]
    pub fn mimetype(&self) -> &str {
        self.inner.mimetype()
    }

    /// Returns the conventional flat extension for this snapshot.
    #[must_use]
    pub fn extension(&self) -> &'static str {
        self.inner.extension()
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
    pub fn metadata(&self) -> FamilyResult<Metadata> {
        self.inner.metadata()
    }

    /// Returns the unnamed fallback page layout, when one is declared.
    pub fn default_page_layout(&self) -> FamilyResult<Option<crate::page_layout::PageLayout>> {
        self.inner.default_page_layout()
    }

    /// Discovers inert inline and linked embedded objects without fetching them.
    pub fn embedded_objects(&self) -> FamilyResult<Vec<crate::Object>> {
        self.inner.embedded_objects()
    }

    /// Inspects ordered variable declarations without evaluating fields.
    pub fn variable_declarations(&self) -> FamilyResult<crate::variable_declaration::Declarations> {
        self.inner.variable_declarations()
    }

    /// Inspects inert document script and macro metadata without executing it.
    pub fn document_scripts(&self) -> FamilyResult<Option<crate::document_scripts::Scripts>> {
        Ok(crate::package::scripts::document_scripts(
            self.xml(),
            &self.limits.options.xml_limits_value(),
        )?)
    }

    /// The options this document was opened with, retained for its edits,
    /// replays and republication.
    pub const fn options(&self) -> OpenOptions {
        self.limits.options
    }

    /// Replays a patch onto this exact source document.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this document is not the
    /// patch's exact source.
    pub fn apply(&self, patch: &Patch) -> std::result::Result<Commit, Error> {
        if self.as_bytes() != patch.before.as_bytes() {
            return Err(Error::patch_conflict(
                "flat.apply",
                "the document is not the patch source",
            ));
        }
        Ok(Commit::new(
            patch.after.clone(),
            patch.clone(),
            Diagnostics::new(!patch.is_noop()),
        ))
    }

    /// Parses modeled top-level text elements in document order.
    pub fn elements(&self) -> FamilyResult<Vec<OrderElement>> {
        Ok(Parser::parse_elements_in_order(self.xml())?)
    }

    /// Parses modeled tables in document order.
    pub fn tables(&self) -> FamilyResult<Vec<Table>> {
        Ok(Parser::parse_tables_in_order(self.xml())?)
    }

    /// Extracts modeled paragraph and heading text.
    pub fn text(&self) -> FamilyResult<String> {
        let mut result = String::new();
        for element in self.elements()? {
            let text = match element {
                OrderElement::Paragraph(paragraph) => paragraph.text()?,
                OrderElement::Heading(heading) => heading.text()?,
                _ => continue,
            };
            result
                .try_reserve(text.len())
                .map_err(|source| CoreError::Allocation {
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
    ///
    /// Unix and Windows use a same-directory temporary and atomic replacement;
    /// other targets return [`CoreError::Unsupported`].
    pub fn save(&self, path: impl AsRef<Path>) -> FamilyResult<()> {
        Ok(self.inner.save_with_attached_context_and_scopes(
            path,
            "flat ODT output",
            "flat ODT symbolic-link or non-file destination",
        )?)
    }

    /// Atomically replaces a filesystem destination using a caller-supplied
    /// cancellation and resource context for staging and publication.
    pub fn save_with_execution_context(
        &self,
        path: impl AsRef<Path>,
        context: litchi_core::ExecutionContext,
    ) -> FamilyResult<()> {
        Ok(self.inner.save_with_execution_context_and_scopes(
            path,
            context,
            "flat ODT output",
            "flat ODT symbolic-link or non-file destination",
        )?)
    }
}

/// Detached, failure-atomic flat ODT edit.
pub struct Edit {
    source: Document,
    operations: Vec<ParagraphEdit>,
    paragraph_sites: Option<Vec<ParagraphSite>>,
}

impl Edit {
    /// The source document this edit stages against.
    #[must_use]
    pub fn package(&self) -> &Document {
        &self.source
    }

    /// Discards every staged paragraph edit in place.
    pub fn reset(&mut self) {
        self.operations.clear();
    }

    /// Stages replacement of a zero-based top-level paragraph's text payload.
    pub fn update_paragraph(
        &mut self,
        index: usize,
        text: &str,
    ) -> FamilyResult<Option<&mut Self>> {
        if self.operations.len() >= self.source.limits.max_edits {
            return Err(resource_limit_error(
                Resource::Objects,
                self.operations.len().saturating_add(1),
                self.source.limits.max_edits,
                "flat ODT paragraph edits",
            )
            .into());
        }
        if text.len() > self.source.limits.max_paragraph_text_bytes {
            return Err(resource_limit_error(
                Resource::InputBytes,
                text.len(),
                self.source.limits.max_paragraph_text_bytes,
                "flat ODT paragraph text",
            )
            .into());
        }
        validate_xml_text(&text)?;
        if self
            .operations
            .iter()
            .any(|operation| operation.index == index)
        {
            return invalid_family("flat ODT edit contains a duplicate paragraph selector");
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
            return invalid_family(
                "flat ODT paragraph contains structured or opaque inline markup",
            );
        };
        let mut owned_text = String::new();
        owned_text
            .try_reserve_exact(text.len())
            .map_err(|source| CoreError::Allocation {
                resource: "flat ODT paragraph text",
                source,
            })?;
        owned_text.push_str(text);
        self.operations
            .try_reserve(1)
            .map_err(|source| CoreError::Allocation {
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
    ///
    /// An edit with no staged operation commits the source itself:
    /// `diagnostics().changed()` is false and `patch().is_noop()` holds.
    ///
    /// # Errors
    ///
    /// Returns an error when the edited document exceeds its limits, is not
    /// compact XML, or fails semantic readback.
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    pub(crate) fn commit_core(mut self) -> Result<Commit> {
        if self.operations.is_empty() {
            let document = self.source.clone();
            let patch = Patch {
                before: self.source,
                after: document.clone(),
            };
            return Ok(Commit::new(document, patch, Diagnostics::UNCHANGED));
        }
        self.operations
            .sort_unstable_by_key(|operation| operation.index);
        let candidate =
            render_paragraph_edits(self.source.xml(), &self.operations, self.source.limits)?;
        let compact_limits = litchi_odf_common::compact_xml::Limits::new(
            self.source.limits.max_document_bytes,
            self.source.limits.max_xml_depth,
        )
        .map_err(CoreError::from)?;
        litchi_odf_common::compact_xml::validate(candidate.as_bytes(), compact_limits)
            .map_err(CoreError::from)?;
        let document =
            Document::from_bytes_with_limits(candidate.into_bytes(), self.source.limits)?;
        semantic_readback(&document, &self.operations)?;
        let changed = self.source.as_bytes() != document.as_bytes();
        let patch = Patch {
            before: self.source,
            after: document.clone(),
        };
        Ok(Commit::new(document, patch, Diagnostics::new(changed)))
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

/// Reversible, exact-byte flat ODT patch.
#[derive(Clone)]
pub struct Patch {
    before: Document,
    after: Document,
}

impl Patch {
    /// Whether this patch was made from exactly `source`'s bytes.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Document) -> bool {
        source.as_bytes() == self.before.as_bytes()
    }

    /// Whether replaying this patch keeps the source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_bytes() == self.after.as_bytes()
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
        .map_err(|source| CoreError::Allocation {
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
    let retained = limits.options.xml_limits_value();
    let max_document_bytes = u64::try_from(limits.max_document_bytes).map_err(|_| {
        CoreError::InvalidFormat("flat ODT input limit exceeds platform limits".to_string())
    })?;
    let xml_limits = XmlStreamLimits::new(
        retained.max_bytes().min(max_document_bytes),
        retained.max_depth().min(limits.max_xml_depth),
        retained.max_events().min(4_000_000),
        retained.max_attributes().min(1_000_000),
        retained.max_token_bytes().min(64 * 1024),
    )?
    .namespace_limits(
        retained
            .max_namespace_scope_depth()
            .min(limits.max_xml_depth),
        retained.max_namespace_declarations().min(256),
        retained.max_namespace_bindings().min(256 * 1024),
        retained.max_namespace_bytes().min(16 * 1024 * 1024),
        retained.max_namespace_prefix_bytes().min(64 * 1024),
        retained.max_namespace_uri_bytes().min(64 * 1024),
    )?;
    let mut reader = ResolvedReader::from_xml(xml, xml_limits)?;
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut office_text_depth = None;
    let mut sites = Vec::new();

    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| {
                crate::core::map_resolved_reader_error(error, "invalid flat ODT XML")
            })?;
        let is_office = crate::elements::xml::normalized_namespace_matches(
            &namespace,
            OFFICE_NAMESPACE,
            "flat ODT",
        )?;
        let is_text = crate::elements::xml::normalized_namespace_matches(
            &namespace,
            TEXT_NAMESPACE,
            "flat ODT",
        )?;
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
                    sites
                        .try_reserve(1)
                        .map_err(|source| CoreError::Allocation {
                            resource: "flat ODT paragraph selector index",
                            source,
                        })?;
                    sites.push(site);
                    depth = depth.checked_sub(1).ok_or_else(|| {
                        CoreError::InvalidFormat("flat ODT XML depth underflow".to_string())
                    })?;
                }
            },
            Event::Empty(ref element) => {
                let local = element.local_name();
                if office_text_depth == Some(depth) && is_text && local.as_ref() == b"p" {
                    sites
                        .try_reserve(1)
                        .map_err(|source| CoreError::Allocation {
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
                    CoreError::InvalidFormat("flat ODT XML depth underflow".to_string())
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
    reader: &mut ResolvedReader<'_>,
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
        let (_, event) = reader.read_resolved_event_into(buffer).map_err(|error| {
            crate::core::map_resolved_reader_error(error, "invalid flat ODT XML")
        })?;
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
                CoreError::InvalidFormat(format!("invalid committed paragraph: {error}"))
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

fn resource_limit_error(
    resource: Resource,
    observed: usize,
    limit: usize,
    scope: &'static str,
) -> CoreError {
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

fn invalid<T>(message: &str) -> Result<T> {
    Err(CoreError::InvalidFormat(message.to_string()))
}

fn invalid_family<T>(message: &str) -> FamilyResult<T> {
    Err(OdfError::from(CoreError::InvalidFormat(
        message.to_string(),
    )))
}
