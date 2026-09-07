//! Bounded, source-backed append of one plain slide to an existing ODP.
//!
//! This module is deliberately separate from [`super::source`] and from the
//! materialized authoring transaction.  It retains the positional package
//! index and one bounded authored page, then asks the common ODF insertion
//! owner to replay `content.xml` into the preserved ZIP layout.  It never
//! constructs a complete `content.xml`, candidate package, `Snapshot`, or
//! `Commit`.

use std::{fmt, io::Write, sync::Arc};

use litchi_core::{
    CancellationToken, Error, ExecutionError, ReadAt, Resource, Result, SourceVersion,
};
use litchi_odf_common::{
    constants::{ODF_CONTENT, ODF_PRESENTATION},
    core::source_publication::{
        SourceContentInsertionError, SourceContentInsertionPlan, SourceContentScanError,
        scan_source_content,
    },
    core::{
        AuthoredXmlFragment, SourceBackedPackage, SourceContentPublicationError,
        SourceContentPublicationOptions, SourceContentPublicationProgress,
        SourceContentPublicationReport, SourceMemberReaderError, XmlStreamLimits,
        private::{BindingTracker, XmlStreamEvent},
    },
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, ResolveResult},
};

use super::ReadLimits;

const FAMILY_NAME: &str = "ODP";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const GENERATED_DOCUMENT_PREFIX: &[u8] = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation>"#;
const GENERATED_DOCUMENT_SUFFIX: &[u8] =
    b"</office:presentation></office:body></office:document-content>";
const GENERATED_READBACK_TEXT_GEOMETRIC_MULTIPLIER: usize = 2;
const GENERATED_READBACK_TEXT_PARALLEL_STRINGS: usize = 3;
const GENERATED_FRAGMENT_GROWTH_MULTIPLIER: usize = 2;
const GENERATED_READBACK_PARSER_BUFFER_MULTIPLIER: usize = 1;
const GENERATED_READBACK_DECODED_TOKEN_MULTIPLIER: usize = 2;
const GENERATED_READBACK_FIXED_STATE_BYTES: usize = 64 * 1024;
const MAX_TEXT_SPACE_COUNT: usize = 1_000_000;
const MAX_SLIDES: usize = 65_536;
const MAX_TEXT_BYTES: usize = 16 * 1024 * 1024;
const MAX_FRAGMENT_BYTES: usize = 512 * 1024;
const MAX_CONTENT_XML_BYTES: u64 = 256 * 1024 * 1024;

/// Finite bounds for a source-backed one-slide append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TailAppendLimits {
    /// Maximum combined UTF-8 bytes in the submitted title and body.
    pub max_text_bytes: usize,
    /// Maximum generated page bytes retained by the plan.
    pub max_fragment_bytes: usize,
    /// Maximum direct pages accepted in the source presentation.
    pub max_slides: usize,
    /// Maximum decoded `content.xml` bytes accepted by the scanner.
    pub max_content_xml_bytes: u64,
}

impl TailAppendLimits {
    /// Create explicit finite append bounds.
    #[must_use]
    pub const fn new(
        max_text_bytes: usize,
        max_fragment_bytes: usize,
        max_slides: usize,
        max_content_xml_bytes: u64,
    ) -> Self {
        Self {
            max_text_bytes,
            max_fragment_bytes,
            max_slides,
            max_content_xml_bytes,
        }
    }
}

impl Default for TailAppendLimits {
    fn default() -> Self {
        Self::new(
            MAX_TEXT_BYTES,
            MAX_FRAGMENT_BYTES,
            MAX_SLIDES,
            MAX_CONTENT_XML_BYTES,
        )
    }
}

/// Failure from the format-owned source append proof or publication handoff.
#[derive(Debug)]
#[non_exhaustive]
pub enum TailAppendError {
    /// The source or authored request is not accepted by the ODP proof.
    Invalid(Error),
    /// A verified source member reader failed while scanning the source.
    SourceMember(SourceMemberReaderError<Error>),
    /// The common insertion owner rejected the bounded plan before output.
    Plan(SourceContentPublicationError),
    /// The common insertion owner reported a typed sink/source failure after
    /// or during publication.
    Publication(SourceContentInsertionError),
}

impl fmt::Display for TailAppendError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(error) => error.fmt(formatter),
            Self::SourceMember(error) => error.fmt(formatter),
            Self::Plan(error) => error.fmt(formatter),
            Self::Publication(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for TailAppendError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Invalid(error) => Some(error),
            Self::SourceMember(error) => Some(error),
            Self::Plan(error) => Some(error),
            Self::Publication(error) => Some(error),
        }
    }
}

impl TailAppendError {
    /// Return sink progress when the common publication layer has observed it.
    #[must_use]
    pub const fn progress(&self) -> Option<SourceContentPublicationProgress> {
        match self {
            Self::Plan(error) => Some(error.progress()),
            Self::Publication(error) => Some(error.progress()),
            Self::Invalid(_) | Self::SourceMember(_) => None,
        }
    }
}

fn invalid(message: impl Into<String>) -> TailAppendError {
    TailAppendError::Invalid(Error::InvalidFormat(message.into()))
}

fn check_plan_options(
    options: &SourceContentPublicationOptions,
) -> std::result::Result<(), TailAppendError> {
    if options
        .cancellation()
        .is_some_and(CancellationToken::is_cancelled)
    {
        return Err(TailAppendError::Plan(
            SourceContentPublicationError::Cancelled {
                progress: SourceContentPublicationProgress::Untouched,
            },
        ));
    }
    if let Some(execution) = options.execution_context() {
        execution
            .check()
            .map_err(map_execution_error)
            .map_err(TailAppendError::Plan)?;
    }
    Ok(())
}

fn map_execution_error(error: ExecutionError) -> SourceContentPublicationError {
    match error {
        ExecutionError::Cancelled => SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Untouched,
        },
        source => SourceContentPublicationError::Execution {
            progress: SourceContentPublicationProgress::Untouched,
            source,
        },
    }
}

fn reserve_fragment_memory(
    options: &SourceContentPublicationOptions,
    maximum: usize,
) -> std::result::Result<Option<litchi_core::Reservation>, TailAppendError> {
    let Some(execution) = options.execution_context() else {
        return Ok(None);
    };
    let amount = u64::try_from(maximum).map_err(|_| {
        invalid("ODP generated page memory bound exceeds the execution accounting range")
    })?;
    execution
        .reserve(Resource::Memory, amount)
        .map(Some)
        .map_err(map_execution_error)
        .map_err(TailAppendError::Plan)
}

/// Bound the format-owned page generation and parser readback window.
///
/// The submitted strings remain live through generated-page readback and are charged by
/// their actual capacity, including spare caller-supplied capacity. The established
/// parser can have two decoded strings for each input channel live at once:
/// the current paragraph plus its retained shape, or the retained body shape
/// plus the slide body projection; the parser's decoded current token is the
/// third.  Each String is charged with geometric capacity slack.  The
/// generated fragment can overlap its old allocation during growth; while
/// readback runs, that allocation overlaps the wrapped document, one parser
/// token buffer, and the decoded current token.  The fixed term covers the
/// one-slide model vectors, namespace stack, and parser bookkeeping; all
/// source-sized buffers are charged separately.  The document wrapper is also
/// larger than the compact wrapper used by the later AuthoredXmlFragment
/// audit.
fn generated_readback_memory_bound(
    title: &str,
    body: &str,
    retained_text_capacity: usize,
    max_fragment_bytes: usize,
) -> Result<usize> {
    let submitted_text_bytes = title
        .len()
        .checked_add(body.len())
        .ok_or_else(|| Error::InvalidFormat("ODP submitted text size overflow".to_string()))?;
    let parser_text_bytes = submitted_text_bytes
        .checked_mul(GENERATED_READBACK_TEXT_GEOMETRIC_MULTIPLIER)
        .and_then(|value| value.checked_mul(GENERATED_READBACK_TEXT_PARALLEL_STRINGS))
        .ok_or_else(|| Error::InvalidFormat("ODP parser text bound overflow".to_string()))?;
    let text_memory = retained_text_capacity
        .checked_add(parser_text_bytes)
        .ok_or_else(|| Error::InvalidFormat("ODP text memory bound overflow".to_string()))?;
    let wrapper_bytes = GENERATED_DOCUMENT_PREFIX
        .len()
        .checked_add(GENERATED_DOCUMENT_SUFFIX.len())
        .ok_or_else(|| {
            Error::InvalidFormat("ODP generated document wrapper size overflow".to_string())
        })?;
    let fragment_growth = max_fragment_bytes
        .checked_mul(GENERATED_FRAGMENT_GROWTH_MULTIPLIER)
        .ok_or_else(|| Error::InvalidFormat("ODP fragment growth bound overflow".to_string()))?;
    let readback_document = max_fragment_bytes
        .checked_add(wrapper_bytes)
        .ok_or_else(|| Error::InvalidFormat("ODP readback document bound overflow".to_string()))?;
    let parser_buffer = max_fragment_bytes
        .checked_mul(GENERATED_READBACK_PARSER_BUFFER_MULTIPLIER)
        .ok_or_else(|| Error::InvalidFormat("ODP parser buffer bound overflow".to_string()))?;
    let decoded_token = max_fragment_bytes
        .checked_mul(GENERATED_READBACK_DECODED_TOKEN_MULTIPLIER)
        .ok_or_else(|| Error::InvalidFormat("ODP decoded token bound overflow".to_string()))?;
    text_memory
        .checked_add(fragment_growth)
        .and_then(|value| value.checked_add(readback_document))
        .and_then(|value| value.checked_add(parser_buffer))
        .and_then(|value| value.checked_add(decoded_token))
        .and_then(|value| value.checked_add(GENERATED_READBACK_FIXED_STATE_BYTES))
        .ok_or_else(|| Error::InvalidFormat("ODP generated page memory bound overflow".to_string()))
}

/// Facts established by the bounded source scan before ZIP preparation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TailAppendSourceProof {
    source_version: SourceVersion,
    content_bytes: u64,
    insert_at: u64,
    slide_count: usize,
    page_name: String,
}

impl TailAppendSourceProof {
    /// Source revision captured by the proof.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.source_version
    }

    /// Verified decoded `content.xml` length.
    #[must_use]
    pub const fn content_bytes(&self) -> u64 {
        self.content_bytes
    }

    /// Exact decoded insertion offset, immediately after the final page.
    #[must_use]
    pub const fn insert_at(&self) -> u64 {
        self.insert_at
    }

    /// Number of direct pages in the source presentation.
    #[must_use]
    pub const fn slide_count(&self) -> usize {
        self.slide_count
    }

    /// Generated page name checked against every existing page during the
    /// candidate preflight pass.
    #[must_use]
    pub fn page_name(&self) -> &str {
        &self.page_name
    }
}

/// A source-backed ODP append edit containing only bounded caller input.
pub struct SourceBackedTailAppendEdit {
    package: Arc<SourceBackedPackage>,
    title: String,
    body: String,
    limits: TailAppendLimits,
}

impl fmt::Debug for SourceBackedTailAppendEdit {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedTailAppendEdit")
            .field("package", &self.package)
            .field("title_bytes", &self.title.len())
            .field("body_bytes", &self.body.len())
            .field("limits", &self.limits)
            .finish()
    }
}

impl SourceBackedTailAppendEdit {
    /// Open an ODP positional source and stage one plain title/body slide.
    pub fn from_read_at(
        source: Arc<dyn ReadAt>,
        title: impl Into<String>,
        body: impl Into<String>,
    ) -> Result<Self> {
        Self::from_read_at_with_limits(source, ReadLimits::default(), title, body)
    }

    /// Open an ODP positional source with explicit package and append bounds.
    pub fn from_read_at_with_limits(
        source: Arc<dyn ReadAt>,
        package_limits: ReadLimits,
        title: impl Into<String>,
        body: impl Into<String>,
    ) -> Result<Self> {
        let package = SourceBackedPackage::from_read_at_with_limits(source, package_limits)?;
        let package = Arc::new(package);
        if package.mimetype()? != ODF_PRESENTATION {
            return Err(Error::InvalidFormat(format!(
                "expected {FAMILY_NAME} package MIME type '{ODF_PRESENTATION}'"
            )));
        }
        let title = title.into();
        let body = body.into();
        validate_text_core(&title, &body, TailAppendLimits::default().max_text_bytes)?;
        Ok(Self {
            package,
            title,
            body,
            limits: TailAppendLimits::default(),
        })
    }

    /// Start from an already indexed, MIME-validated positional package.
    pub fn new(
        package: Arc<SourceBackedPackage>,
        title: impl Into<String>,
        body: impl Into<String>,
    ) -> std::result::Result<Self, TailAppendError> {
        Self::with_limits(package, title, body, TailAppendLimits::default())
    }

    /// Start from an already indexed package with explicit append bounds.
    pub fn with_limits(
        package: Arc<SourceBackedPackage>,
        title: impl Into<String>,
        body: impl Into<String>,
        limits: TailAppendLimits,
    ) -> std::result::Result<Self, TailAppendError> {
        let mimetype = package.mimetype().map_err(TailAppendError::Invalid)?;
        if mimetype != ODF_PRESENTATION {
            return Err(invalid(format!(
                "expected {FAMILY_NAME} package MIME type '{ODF_PRESENTATION}'"
            )));
        }
        let title = title.into();
        let body = body.into();
        validate_text(&title, &body, limits.max_text_bytes)?;
        Ok(Self {
            package,
            title,
            body,
            limits,
        })
    }

    /// Return the indexed source package retained by this edit.
    #[must_use]
    pub fn package(&self) -> &Arc<SourceBackedPackage> {
        &self.package
    }

    /// Return the configured append bounds.
    #[must_use]
    pub const fn limits(&self) -> TailAppendLimits {
        self.limits
    }

    /// Build a source proof and common ZIP insertion plan.
    pub fn plan(
        self,
        options: &SourceContentPublicationOptions,
    ) -> std::result::Result<SourceBackedTailAppendPublicationPlan, TailAppendError> {
        let Self {
            package,
            title,
            body,
            limits,
        } = self;
        check_plan_options(options)?;
        let source_version = package.source_version().map_err(TailAppendError::Invalid)?;
        let content_bytes = package
            .member_materialized_size(ODF_CONTENT)
            .map_err(TailAppendError::Invalid)?
            .ok_or_else(|| invalid("ODP package has no content.xml"))?;
        if content_bytes > limits.max_content_xml_bytes {
            return Err(invalid("ODP content.xml exceeds the source append limit"));
        }
        let first = scan_source(&package, limits, options)?;
        if first.slide_count >= limits.max_slides {
            return Err(invalid("ODP append would exceed the slide-count limit"));
        }
        let page_number = first
            .slide_count
            .checked_add(1)
            .ok_or_else(|| invalid("ODP page-name index overflow"))?;
        let page_name = format!("page{page_number}");
        let authored = {
            let retained_text_capacity = title
                .capacity()
                .checked_add(body.capacity())
                .ok_or_else(|| invalid("ODP retained text capacity overflow"))?;
            let readback_bound = generated_readback_memory_bound(
                &title,
                &body,
                retained_text_capacity,
                limits.max_fragment_bytes,
            )
            .map_err(TailAppendError::Invalid)?;
            let _memory = reserve_fragment_memory(options, readback_bound)?;
            let fragment = build_plain_page(&page_name, &title, &body, limits.max_fragment_bytes)?;
            // Keep the reservation alive while the exact authored fragment is
            // audited by the established ODP parser.  The parser sees only
            // this bounded page wrapped in a small, locally namespace-bound
            // document; the source content remains cold and opaque.
            validate_generated_page_semantics(&fragment, &title, &body)
                .map_err(TailAppendError::Invalid)?;
            let authored =
                AuthoredXmlFragment::markup(fragment).map_err(TailAppendError::Invalid)?;
            // Release submitted allocations while their reservation is still
            // held; only the authored fragment enters common planning.
            drop(title);
            drop(body);
            authored
        };
        let mut validator = CandidateValidator::new(first.slide_count, &page_name);
        let candidate_token_bytes = limits
            .max_fragment_bytes
            .saturating_add(1024)
            .clamp(1, 64 * 1024 * 1024);
        let common = SourceContentInsertionPlan::prepare_with_validator(
            Arc::clone(&package),
            first.insert_at,
            authored,
            XmlStreamLimits::new(
                limits.max_content_xml_bytes,
                256,
                4_000_000,
                1_000_000,
                candidate_token_bytes,
            )
            .map_err(TailAppendError::Invalid)?,
            options,
            |event, tracker| validator.observe(event, tracker),
        )
        .map_err(TailAppendError::Plan)?;

        Ok(SourceBackedTailAppendPublicationPlan {
            common,
            proof: TailAppendSourceProof {
                source_version,
                content_bytes,
                insert_at: first.insert_at,
                slide_count: first.slide_count,
                page_name,
            },
        })
    }
}

/// A fully source-proved ODP append publication plan.
pub struct SourceBackedTailAppendPublicationPlan {
    common: SourceContentInsertionPlan,
    proof: TailAppendSourceProof,
}

impl fmt::Debug for SourceBackedTailAppendPublicationPlan {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SourceBackedTailAppendPublicationPlan")
            .field("proof", &self.proof)
            .finish_non_exhaustive()
    }
}

impl SourceBackedTailAppendPublicationPlan {
    /// Return the ODP source proof.
    #[must_use]
    pub const fn proof(&self) -> &TailAppendSourceProof {
        &self.proof
    }

    /// Return the common insertion plan for diagnostics.
    #[must_use]
    pub fn common_plan(&self) -> &SourceContentInsertionPlan {
        &self.common
    }

    /// Return the source archive version captured by the common plan.
    #[must_use]
    pub const fn source_version(&self) -> SourceVersion {
        self.common.source_version()
    }

    /// Return the source archive length captured by the common plan.
    #[must_use]
    pub const fn source_length(&self) -> u64 {
        self.common.source_length()
    }

    /// Return the decoded insertion offset captured by the common plan.
    #[must_use]
    pub const fn decoded_offset(&self) -> u64 {
        self.common.decoded_offset()
    }

    /// Return the source `content.xml` decoded length.
    #[must_use]
    pub const fn source_content_length(&self) -> u64 {
        self.common.source_content_length()
    }

    /// Return the candidate `content.xml` decoded length.
    #[must_use]
    pub const fn target_content_length(&self) -> u64 {
        self.common.target_content_length()
    }

    /// Return the source `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn source_content_sha256(&self) -> [u8; 32] {
        self.common.source_content_sha256()
    }

    /// Return the candidate `content.xml` SHA-256 proof.
    #[must_use]
    pub const fn target_content_sha256(&self) -> [u8; 32] {
        self.common.target_content_sha256()
    }

    /// Publish to a sequential sink using the common raw-preserving ZIP plan.
    pub fn write_to<W: Write>(
        &self,
        sink: W,
        options: SourceContentPublicationOptions,
    ) -> std::result::Result<SourceContentPublicationReport, TailAppendError> {
        self.common
            .write_to(sink, options)
            .map_err(TailAppendError::Publication)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct ScanFacts {
    content_bytes: u64,
    insert_at: u64,
    slide_count: usize,
}

fn scan_source(
    package: &SourceBackedPackage,
    limits: TailAppendLimits,
    options: &SourceContentPublicationOptions,
) -> std::result::Result<ScanFacts, TailAppendError> {
    let mut facts = ScanFacts {
        content_bytes: 0,
        insert_at: 0,
        slide_count: 0,
    };
    let stream_limits = XmlStreamLimits::new(
        limits.max_content_xml_bytes,
        256,
        4_000_000,
        1_000_000,
        64 * 1024,
    )
    .map_err(TailAppendError::Invalid)?;
    let mut state = OdpScanState::default();
    let report = scan_source_content(
        package,
        stream_limits,
        options,
        |event: &XmlStreamEvent<'_>, tracker: &BindingTracker| {
            state.observe(event, tracker, &mut facts, limits)
        },
    )
    .map_err(|error| match error {
        SourceContentScanError::Publication(error) => TailAppendError::Plan(error),
        SourceContentScanError::SourceMember(error) => TailAppendError::SourceMember(error),
        _ => TailAppendError::Invalid(Error::InvalidFormat(
            "unknown source-content scan failure".to_string(),
        )),
    })?;
    facts.content_bytes = report.bytes();
    facts.slide_count = state.pages_seen;
    state.finish(&facts).map_err(TailAppendError::Invalid)?;
    Ok(facts)
}

#[derive(Debug, Default)]
struct OdpScanState {
    stack: Vec<ElementKind>,
    root_seen: bool,
    body_seen: bool,
    presentation_seen: bool,
    presentation_empty: bool,
    page_open: bool,
    pages_seen: usize,
    trailing_seen: bool,
    content_bytes: u64,
}

#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    #[default]
    Other,
    Root,
    Body,
    Presentation,
    Page,
}

impl OdpScanState {
    fn observe(
        &mut self,
        event: &XmlStreamEvent<'_>,
        tracker: &BindingTracker,
        facts: &mut ScanFacts,
        limits: TailAppendLimits,
    ) -> Result<()> {
        self.content_bytes = self.content_bytes.max(event.end());
        if self.content_bytes > limits.max_content_xml_bytes {
            return Err(candidate_invalid(
                "ODP content.xml exceeds the source append limit",
            ));
        }
        match event.event() {
            Event::Start(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    if self.trailing_seen {
                        return Err(candidate_invalid(
                            "ODP presentation pages are not a contiguous direct prefix",
                        ));
                    }
                    if self.page_open {
                        return Err(candidate_invalid(
                            "ODP presentation contains a nested draw:page",
                        ));
                    }
                    if self.pages_seen >= limits.max_slides {
                        return Err(candidate_invalid(
                            "ODP presentation exceeds the slide-count limit",
                        ));
                    }
                    self.pages_seen = self
                        .pages_seen
                        .checked_add(1)
                        .ok_or_else(|| candidate_invalid("ODP slide-count overflow"))?;
                    self.page_open = true;
                } else if depth == 3
                    && self.presentation_seen
                    && self.pages_seen > 0
                    && !self.page_open
                {
                    self.trailing_seen = true;
                }
                self.stack.push(kind);
            },
            Event::Empty(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    if self.trailing_seen {
                        return Err(candidate_invalid(
                            "ODP presentation pages are not a contiguous direct prefix",
                        ));
                    }
                    if self.pages_seen >= limits.max_slides {
                        return Err(candidate_invalid(
                            "ODP presentation exceeds the slide-count limit",
                        ));
                    }
                    self.pages_seen = self
                        .pages_seen
                        .checked_add(1)
                        .ok_or_else(|| candidate_invalid("ODP slide-count overflow"))?;
                    facts.insert_at = event.end();
                } else if depth == 3 && self.presentation_seen && self.pages_seen > 0 {
                    self.trailing_seen = true;
                }
                if kind == ElementKind::Presentation {
                    self.presentation_empty = true;
                }
            },
            Event::End(_) => {
                let kind = self
                    .stack
                    .pop()
                    .ok_or_else(|| candidate_invalid("ODP content.xml depth underflow"))?;
                match kind {
                    ElementKind::Page => {
                        if !self.page_open {
                            return Err(candidate_invalid("ODP presentation page state underflow"));
                        }
                        self.page_open = false;
                        facts.insert_at = event.end();
                    },
                    ElementKind::Presentation => {
                        if self.presentation_empty {
                            return Err(candidate_invalid(
                                "ODP append does not support self-closing office:presentation",
                            ));
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(text)
                if self.stack.len() == 3 && self.presentation_seen && self.pages_seen > 0 =>
            {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    self.trailing_seen = true;
                }
            },
            Event::Comment(_) | Event::PI(_) if self.stack.len() == 3 && self.pages_seen > 0 => {
                self.trailing_seen = true;
            },
            Event::CData(_) | Event::GeneralRef(_)
                if self.stack.len() == 3 && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::DocType(_) => {
                return Err(candidate_invalid(
                    "ODP content.xml must not contain a doctype",
                ));
            },
            Event::Eof => {},
            _ => {},
        }
        Ok(())
    }

    fn classify_start(
        &mut self,
        depth: usize,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
    ) -> Result<ElementKind> {
        let (namespace, local) = tracker.resolve_element(element.name()).map_err(|error| {
            error.into_litchi_error_with_context(|| "invalid ODP namespace binding".to_string())
        })?;
        let is = |expected_namespace: &[u8], expected_local: &[u8]| {
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == expected_namespace)
                && local.as_ref() == expected_local
        };
        let kind = if depth == 0 {
            if !is(OFFICE_NAMESPACE, b"document-content") {
                return Err(candidate_invalid(
                    "ODP content.xml root is not office:document-content",
                ));
            }
            if self.root_seen {
                return Err(candidate_invalid("ODP content.xml has duplicate root"));
            }
            self.root_seen = true;
            ElementKind::Root
        } else if depth == 1 && is(OFFICE_NAMESPACE, b"body") {
            if self.stack.len() != 1 || self.stack[0] != ElementKind::Root {
                return Err(candidate_invalid(
                    "ODP office:body is outside office:document-content",
                ));
            }
            if self.body_seen {
                return Err(candidate_invalid(
                    "ODP content.xml has duplicate office:body",
                ));
            }
            self.body_seen = true;
            ElementKind::Body
        } else if depth == 2 && is(OFFICE_NAMESPACE, b"presentation") {
            if self.stack.len() != 2
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
            {
                return Err(candidate_invalid(
                    "ODP office:presentation is outside office:body",
                ));
            }
            if self.presentation_seen {
                return Err(candidate_invalid(
                    "ODP content.xml has duplicate office:presentation",
                ));
            }
            self.presentation_seen = true;
            ElementKind::Presentation
        } else if is(DRAW_NAMESPACE, b"page") {
            if depth != 3
                || self.stack.len() != 3
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
                || self.stack[2] != ElementKind::Presentation
            {
                return Err(candidate_invalid(
                    "ODP draw:page must be a direct presentation child",
                ));
            }
            ElementKind::Page
        } else {
            ElementKind::Other
        };
        Ok(kind)
    }

    fn finish(&self, facts: &ScanFacts) -> Result<()> {
        if !self.root_seen || !self.body_seen || !self.presentation_seen {
            return Err(candidate_invalid(
                "ODP content.xml has no complete presentation body",
            ));
        }
        if self.presentation_empty || facts.slide_count == 0 {
            return Err(candidate_invalid(
                "ODP append requires a non-empty, non-self-closing presentation",
            ));
        }
        if !self.stack.is_empty() || self.page_open {
            return Err(candidate_invalid(
                "ODP content.xml has unterminated elements",
            ));
        }
        if facts.insert_at == 0 || facts.insert_at > facts.content_bytes {
            return Err(candidate_invalid(
                "ODP append insertion offset is outside content.xml",
            ));
        }
        Ok(())
    }
}

/// Candidate-side proof state for the one generated page.
///
/// The source scan proves the existing presentation and the common insertion
/// owner proves the complete byte-level splice.  This state adds the
/// format-owned candidate fact that exactly one generated direct page occurs
/// immediately after the source page prefix, before any trailing child.
#[derive(Debug)]
struct CandidateValidator {
    source_pages: usize,
    expected_name: String,
    stack: Vec<ElementKind>,
    root_seen: bool,
    body_seen: bool,
    presentation_seen: bool,
    presentation_empty: bool,
    page_open: bool,
    trailing_seen: bool,
    pages_seen: usize,
    expected_seen: bool,
}

impl CandidateValidator {
    fn new(source_pages: usize, expected_name: &str) -> Self {
        Self {
            source_pages,
            expected_name: expected_name.to_owned(),
            stack: Vec::new(),
            root_seen: false,
            body_seen: false,
            presentation_seen: false,
            presentation_empty: false,
            page_open: false,
            trailing_seen: false,
            pages_seen: 0,
            expected_seen: false,
        }
    }

    fn observe(&mut self, event: &XmlStreamEvent<'_>, tracker: &BindingTracker) -> Result<()> {
        match event.event() {
            Event::Start(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    self.begin_page(element, tracker, event.decoder(), false)?;
                    self.page_open = true;
                } else if self.is_direct_presentation_child(depth)
                    && self.pages_seen > 0
                    && !self.page_open
                {
                    self.trailing_seen = true;
                }
                self.stack.push(kind);
            },
            Event::Empty(element) => {
                let depth = self.stack.len();
                let kind = self.classify_start(depth, element, tracker)?;
                if kind == ElementKind::Page {
                    self.begin_page(element, tracker, event.decoder(), true)?;
                } else if self.is_direct_presentation_child(depth) && self.pages_seen > 0 {
                    self.trailing_seen = true;
                }
                if kind == ElementKind::Presentation {
                    self.presentation_empty = true;
                }
            },
            Event::End(_) => {
                let kind = self
                    .stack
                    .pop()
                    .ok_or_else(|| candidate_invalid("ODP candidate depth underflow"))?;
                match kind {
                    ElementKind::Page => {
                        if !self.page_open {
                            return Err(candidate_invalid("ODP candidate page state underflow"));
                        }
                        self.page_open = false;
                    },
                    ElementKind::Presentation => {
                        if self.presentation_empty {
                            return Err(candidate_invalid(
                                "ODP candidate contains a self-closing presentation",
                            ));
                        }
                    },
                    _ => {},
                }
            },
            Event::Text(text) if self.is_direct_presentation_text() && self.pages_seen > 0 => {
                if !text.iter().all(u8::is_ascii_whitespace) {
                    self.trailing_seen = true;
                }
            },
            Event::Comment(_) | Event::PI(_)
                if self.is_direct_presentation_text() && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::CData(_) | Event::GeneralRef(_)
                if self.is_direct_presentation_text() && self.pages_seen > 0 =>
            {
                self.trailing_seen = true;
            },
            Event::Eof => self.finish()?,
            _ => {},
        }
        Ok(())
    }

    fn begin_page(
        &mut self,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
        decoder: quick_xml::encoding::Decoder,
        empty: bool,
    ) -> Result<()> {
        if self.trailing_seen {
            return Err(candidate_invalid(
                "ODP candidate pages are not a contiguous direct prefix",
            ));
        }
        if self.page_open {
            return Err(candidate_invalid(
                "ODP candidate contains a nested draw:page",
            ));
        }
        let position = self.pages_seen;
        self.pages_seen = self
            .pages_seen
            .checked_add(1)
            .ok_or_else(|| candidate_invalid("ODP candidate slide-count overflow"))?;
        if page_name_matches_core(element, tracker, decoder, &self.expected_name)? {
            if self.expected_seen {
                return Err(candidate_invalid(
                    "ODP candidate contains the generated page more than once",
                ));
            }
            if position != self.source_pages {
                return Err(candidate_invalid(
                    "ODP generated page is not immediately after the source page prefix",
                ));
            }
            self.expected_seen = true;
            if empty {
                return Err(candidate_invalid(
                    "ODP generated page must be a non-empty draw:page",
                ));
            }
        }
        Ok(())
    }

    fn classify_start(
        &mut self,
        depth: usize,
        element: &BytesStart<'_>,
        tracker: &BindingTracker,
    ) -> Result<ElementKind> {
        let (namespace, local) = tracker.resolve_element(element.name()).map_err(|error| {
            error.into_litchi_error_with_context(|| {
                "invalid ODP candidate namespace binding".to_string()
            })
        })?;
        let is = |expected_namespace: &[u8], expected_local: &[u8]| {
            matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == expected_namespace)
                && local.as_ref() == expected_local
        };
        if depth == 0 {
            if !is(OFFICE_NAMESPACE, b"document-content") {
                return Err(candidate_invalid(
                    "ODP candidate root is not office:document-content",
                ));
            }
            if self.root_seen {
                return Err(candidate_invalid("ODP candidate has duplicate root"));
            }
            self.root_seen = true;
            return Ok(ElementKind::Root);
        }
        if depth == 1 && is(OFFICE_NAMESPACE, b"body") {
            if self.stack.len() != 1 || self.stack[0] != ElementKind::Root {
                return Err(candidate_invalid(
                    "ODP candidate office:body is outside office:document-content",
                ));
            }
            if self.body_seen {
                return Err(candidate_invalid("ODP candidate has duplicate office:body"));
            }
            self.body_seen = true;
            return Ok(ElementKind::Body);
        }
        if depth == 2 && is(OFFICE_NAMESPACE, b"presentation") {
            if self.stack.len() != 2
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
            {
                return Err(candidate_invalid(
                    "ODP candidate office:presentation is outside office:body",
                ));
            }
            if self.presentation_seen {
                return Err(candidate_invalid(
                    "ODP candidate has duplicate office:presentation",
                ));
            }
            self.presentation_seen = true;
            return Ok(ElementKind::Presentation);
        }
        if is(DRAW_NAMESPACE, b"page") {
            if depth != 3
                || self.stack.len() != 3
                || self.stack[0] != ElementKind::Root
                || self.stack[1] != ElementKind::Body
                || self.stack[2] != ElementKind::Presentation
            {
                return Err(candidate_invalid(
                    "ODP candidate draw:page must be a direct presentation child",
                ));
            }
            return Ok(ElementKind::Page);
        }
        Ok(ElementKind::Other)
    }

    fn is_direct_presentation_child(&self, depth: usize) -> bool {
        depth == 3
            && self.presentation_seen
            && self.stack.len() == 3
            && self.stack[0] == ElementKind::Root
            && self.stack[1] == ElementKind::Body
            && self.stack[2] == ElementKind::Presentation
    }

    fn is_direct_presentation_text(&self) -> bool {
        self.stack.len() == 3
            && self.presentation_seen
            && self.stack[0] == ElementKind::Root
            && self.stack[1] == ElementKind::Body
            && self.stack[2] == ElementKind::Presentation
    }

    fn finish(&self) -> Result<()> {
        if !self.root_seen || !self.body_seen || !self.presentation_seen {
            return Err(candidate_invalid(
                "ODP candidate has no complete presentation body",
            ));
        }
        if self.presentation_empty || self.pages_seen == 0 {
            return Err(candidate_invalid(
                "ODP candidate requires a non-empty presentation",
            ));
        }
        if !self.stack.is_empty() || self.page_open {
            return Err(candidate_invalid("ODP candidate has unterminated elements"));
        }
        let expected_count = self
            .source_pages
            .checked_add(1)
            .ok_or_else(|| candidate_invalid("ODP candidate slide-count overflow"))?;
        if self.pages_seen != expected_count {
            return Err(candidate_invalid(
                "ODP candidate page count does not match the source plus one generated page",
            ));
        }
        if !self.expected_seen {
            return Err(candidate_invalid(
                "ODP candidate is missing the generated non-empty page",
            ));
        }
        Ok(())
    }
}

fn candidate_invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn page_name_matches_core(
    element: &BytesStart<'_>,
    tracker: &BindingTracker,
    decoder: quick_xml::encoding::Decoder,
    expected: &str,
) -> Result<bool> {
    for attribute in element.attributes().with_checks(false) {
        let attribute = attribute.map_err(|error| {
            Error::InvalidFormat(format!("invalid ODP draw:page attribute: {error}"))
        })?;
        let (namespace, local) = tracker.resolve_attribute(attribute.key).map_err(|error| {
            error.into_litchi_error_with_context(|| "invalid ODP attribute namespace".to_string())
        })?;
        if local.as_ref() == b"name"
            && matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == DRAW_NAMESPACE)
        {
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
                .map_err(|error| Error::InvalidFormat(format!("invalid ODP draw:name: {error}")))?;
            return Ok(value == expected);
        }
    }
    Ok(false)
}

/// Prove the generated page's semantic projection with the same ODP parser
/// used by `Presentation` and `SourceBackedPresentation`.  The fragment is
/// wrapped only for this bounded readback pass; the candidate scanner still
/// proves the exact namespace bindings, page position, count, and name in the
/// actual source splice.
fn validate_generated_page_semantics(fragment: &[u8], title: &str, body: &str) -> Result<()> {
    let length = GENERATED_DOCUMENT_PREFIX
        .len()
        .checked_add(fragment.len())
        .and_then(|length| length.checked_add(GENERATED_DOCUMENT_SUFFIX.len()))
        .ok_or_else(|| Error::InvalidFormat("ODP generated document size overflow".to_string()))?;
    let mut document = Vec::new();
    document
        .try_reserve_exact(length)
        .map_err(|source| Error::Allocation {
            resource: "ODP generated semantic readback document",
            source,
        })?;
    document.extend_from_slice(GENERATED_DOCUMENT_PREFIX);
    document.extend_from_slice(fragment);
    document.extend_from_slice(GENERATED_DOCUMENT_SUFFIX);
    let document = std::str::from_utf8(&document).map_err(|error| {
        Error::InvalidFormat(format!("generated ODP page is not UTF-8: {error}"))
    })?;
    let slides = crate::codec::Parser::parse_slides_with_styles(document, None)?;
    let slide = slides.first().ok_or_else(|| {
        Error::InvalidFormat("ODP generated page semantic readback found no slide".to_string())
    })?;
    if slides.len() != 1 {
        return Err(Error::InvalidFormat(
            "ODP generated page semantic readback found more than one slide".to_string(),
        ));
    }
    if slide.title.as_deref().unwrap_or_default() != title {
        return Err(Error::InvalidFormat(
            "ODP generated title does not survive the established parser projection".to_string(),
        ));
    }
    if body.trim().is_empty() {
        if !slide.text.is_empty() {
            return Err(Error::InvalidFormat(
                "ODP generated whitespace-only body unexpectedly entered the slide text projection"
                    .to_string(),
            ));
        }
        let mut object_shape = None;
        for shape in &slide.shapes {
            if shape.presentation_class() != Some("object") {
                continue;
            }
            if object_shape.replace(shape).is_some() {
                return Err(Error::InvalidFormat(
                    "ODP generated whitespace-only body produced multiple object shapes"
                        .to_string(),
                ));
            }
        }
        if body.is_empty() {
            if object_shape.is_some() {
                return Err(Error::InvalidFormat(
                    "ODP generated empty body produced an object shape".to_string(),
                ));
            }
        } else {
            let shape = object_shape.ok_or_else(|| {
                Error::InvalidFormat(
                    "ODP generated whitespace-only body was not retained as an object shape"
                        .to_string(),
                )
            })?;
            if shape.text()? != body {
                return Err(Error::InvalidFormat(
                    "ODP generated whitespace-only body does not survive object-shape readback"
                        .to_string(),
                ));
            }
        }
    } else if slide.text != body {
        return Err(Error::InvalidFormat(
            "ODP generated body does not survive the established parser projection".to_string(),
        ));
    }
    Ok(())
}

fn validate_text(
    title: &str,
    body: &str,
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    validate_text_core(title, body, maximum).map_err(TailAppendError::Invalid)
}

fn validate_text_core(title: &str, body: &str, maximum: usize) -> Result<()> {
    let bytes = title
        .len()
        .checked_add(body.len())
        .ok_or_else(|| Error::InvalidFormat("ODP slide text size overflow".to_string()))?;
    if bytes > maximum {
        return Err(Error::InvalidFormat(
            "ODP slide text exceeds the append limit".to_string(),
        ));
    }
    if title
        .chars()
        .chain(body.chars())
        .any(|character| character == '\r')
    {
        return Err(Error::InvalidFormat(
            "ODP slide text contains a carriage return; use '\\n' for paragraph breaks".to_string(),
        ));
    }
    if title.chars().chain(body.chars()).any(|character| {
        !matches!(
            character,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(Error::InvalidFormat(
            "ODP slide text contains a character forbidden by XML 1.0".to_string(),
        ));
    }
    if title.chars().chain(body.chars()).any(|character| {
        character.is_whitespace() && !matches!(character, ' ' | '\t' | '\n' | '\r')
    }) {
        return Err(Error::InvalidFormat(
            "ODP slide text contains Unicode whitespace that the established parser normalizes; use spaces, tabs, or newlines".to_string(),
        ));
    }
    Ok(())
}

fn build_plain_page(
    name: &str,
    title: &str,
    body: &str,
    maximum: usize,
) -> std::result::Result<Vec<u8>, TailAppendError> {
    let mut output = Vec::new();
    append_limited(
        &mut output,
        br#"<draw:page xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" draw:name=""#,
        maximum,
    )?;
    append_escaped_text(&mut output, name, maximum)?;
    append_limited(&mut output, b"\">", maximum)?;
    if !title.is_empty() {
        append_limited(
            &mut output,
            br#"<draw:frame presentation:class="title" svg:width="25.199cm" svg:height="3.506cm" svg:x="1.4cm" svg:y="0.962cm"><draw:text-box>"#,
            maximum,
        )?;
        append_text_paragraphs(&mut output, title, maximum)?;
        append_limited(&mut output, b"</draw:text-box></draw:frame>", maximum)?;
    }
    if !body.is_empty() {
        let y = if title.is_empty() { "2.0cm" } else { "5.0cm" };
        append_limited(
            &mut output,
            br#"<draw:frame presentation:class="object" svg:width="25.199cm" svg:height="10cm" svg:x="1.4cm" svg:y=""#,
            maximum,
        )?;
        append_limited(&mut output, y.as_bytes(), maximum)?;
        append_limited(&mut output, b"\"><draw:text-box>", maximum)?;
        append_text_paragraphs(&mut output, body, maximum)?;
        append_limited(&mut output, b"</draw:text-box></draw:frame>", maximum)?;
    }
    append_limited(&mut output, b"</draw:page>", maximum)?;
    Ok(output)
}

fn append_limited(
    output: &mut Vec<u8>,
    bytes: &[u8],
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    let required = output
        .len()
        .checked_add(bytes.len())
        .ok_or_else(|| invalid("generated ODP page length overflow"))?;
    if required > maximum {
        return Err(invalid("generated ODP page exceeds the fragment limit"));
    }
    output
        .try_reserve_exact(bytes.len())
        .map_err(|error| invalid(format!("ODP append fragment allocation failed: {error}")))?;
    output.extend_from_slice(bytes);
    Ok(())
}

fn append_escaped_text(
    output: &mut Vec<u8>,
    text: &str,
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    for character in text.chars() {
        append_escaped_char(output, character, maximum)?;
    }
    Ok(())
}

fn append_escaped_char(
    output: &mut Vec<u8>,
    character: char,
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    match character {
        '&' => append_limited(output, b"&amp;", maximum),
        '<' => append_limited(output, b"&lt;", maximum),
        '>' => append_limited(output, b"&gt;", maximum),
        '"' => append_limited(output, b"&quot;", maximum),
        '\'' => append_limited(output, b"&apos;", maximum),
        character => {
            let mut encoded = [0_u8; 4];
            let encoded = character.encode_utf8(&mut encoded);
            append_limited(output, encoded.as_bytes(), maximum)
        },
    }
}

fn append_text_paragraphs(
    output: &mut Vec<u8>,
    text: &str,
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    for paragraph in text.split('\n') {
        append_limited(output, b"<text:p>", maximum)?;
        append_encoded_paragraph(output, paragraph, maximum)?;
        append_limited(output, b"</text:p>", maximum)?;
    }
    Ok(())
}

fn append_encoded_paragraph(
    output: &mut Vec<u8>,
    paragraph: &str,
    maximum: usize,
) -> std::result::Result<(), TailAppendError> {
    let mut characters = paragraph.chars().peekable();
    let mut encoded_any = false;
    while let Some(character) = characters.next() {
        match character {
            ' ' => {
                let mut count = 1usize;
                while characters.next_if_eq(&' ').is_some() {
                    count = count
                        .checked_add(1)
                        .ok_or_else(|| invalid("ODP generated space count overflow"))?;
                }
                if count == 1 && encoded_any && characters.peek().is_some() {
                    append_limited(output, b" ", maximum)?;
                } else if count == 1 {
                    append_limited(output, b"<text:s/>", maximum)?;
                } else {
                    let mut remaining = count;
                    while remaining != 0 {
                        let chunk_count = remaining.min(MAX_TEXT_SPACE_COUNT);
                        append_limited(output, b"<text:s text:c=\"", maximum)?;
                        let chunk = chunk_count.to_string();
                        append_limited(output, chunk.as_bytes(), maximum)?;
                        append_limited(output, b"\"/>", maximum)?;
                        remaining -= chunk_count;
                    }
                }
                encoded_any = true;
            },
            '\t' => {
                append_limited(output, b"<text:tab/>", maximum)?;
                encoded_any = true;
            },
            '\r' => {
                return Err(invalid(
                    "ODP slide text contains a carriage return; use '\\n' for paragraph breaks",
                ));
            },
            character => {
                append_escaped_char(output, character, maximum)?;
                encoded_any = true;
            },
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_core::{OwnedSource, ReadAt};
    use litchi_odf_common::core::{OwnedPackage, PackageWriter};

    use super::super::{Presentation, SourceBackedPresentation};

    const OFFICE_NAMESPACE_UTF8: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const DRAW_NAMESPACE_UTF8: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    const PRESENTATION_NAMESPACE_UTF8: &str =
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";

    fn source_package(content: &str) -> Arc<dyn ReadAt> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(ODF_PRESENTATION).unwrap();
        writer.add_file(ODF_CONTENT, content.as_bytes()).unwrap();
        Arc::new(OwnedSource::new(writer.finish_to_bytes().unwrap()))
    }

    fn plan_options() -> SourceContentPublicationOptions {
        SourceContentPublicationOptions::new()
    }

    #[test]
    fn empty_presentation_refuses_with_typed_invalid_error() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}"><office:body><office:presentation></office:presentation></office:body></office:document-content>"#,
        );
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", "body")
                .unwrap();
        let error = edit.plan(&plan_options()).unwrap_err();
        assert!(matches!(
            error,
            TailAppendError::Invalid(Error::InvalidFormat(message))
                if message.contains("non-empty")
        ));
    }

    #[test]
    fn self_closing_presentation_refuses_with_typed_invalid_error() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}"><office:body><office:presentation/></office:body></office:document-content>"#,
        );
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", "body")
                .unwrap();
        let error = edit.plan(&plan_options()).unwrap_err();
        assert!(matches!(
            error,
            TailAppendError::Invalid(Error::InvalidFormat(message))
                if message.contains("non-empty") || message.contains("self-closing")
        ));
    }

    #[test]
    fn candidate_validation_refuses_existing_generated_name_without_source_rescan() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page2"/></office:presentation></office:body></office:document-content>"#,
        );
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", "body")
                .unwrap();
        let error = edit.plan(&plan_options()).unwrap_err();
        assert!(matches!(
            error,
            TailAppendError::Plan(SourceContentPublicationError::Core(
                Error::InvalidFormat(message)
            )) if message.contains("generated page")
        ));
    }

    #[test]
    fn publication_inserts_before_native_trailing_presentation_child() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}" xmlns:presentation="{PRESENTATION_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/><presentation:settings presentation:start-page="page1"/><!--tail-token--></office:presentation></office:body></office:document-content>"#,
        );
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", "body")
                .unwrap();
        let options = plan_options();
        let plan = edit.plan(&options).unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output, options).unwrap();

        let package = OwnedPackage::from_bytes(output).unwrap();
        let published = String::from_utf8(package.get_file(ODF_CONTENT).unwrap()).unwrap();
        let generated = published.find(r#"draw:name="page2""#).unwrap();
        let settings = published.find("<presentation:settings").unwrap();
        assert!(generated < settings);
        assert!(published.contains(
            r#"<presentation:settings presentation:start-page="page1"/><!--tail-token-->"#
        ));
    }

    #[test]
    fn generated_story_semantics_match_established_readback() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let title = "A & < B > \"C\" ' 😀\t\nnext";
        let body = "  body  \n第二行";
        let edit = SourceBackedTailAppendEdit::from_read_at(source_package(&content), title, body)
            .unwrap();
        let options = plan_options();
        let plan = edit.plan(&options).unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output, options).unwrap();

        let source_backed =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output.clone())))
                .unwrap();
        let source_slides = source_backed.slides().unwrap();
        let source_slide = source_slides.last().unwrap();
        assert_eq!(source_slide.title().unwrap(), Some(title));
        assert_eq!(source_slide.text().unwrap(), body);

        let presentation = Presentation::from_bytes(output).unwrap();
        let slides = presentation.slides().unwrap();
        let slide = slides.last().unwrap();
        assert_eq!(slide.title().unwrap(), Some(title));
        assert_eq!(slide.text().unwrap(), body);
    }

    #[test]
    fn whitespace_only_body_is_read_back_from_object_shape() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let body = " \t\n  ";
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", body)
                .unwrap();
        let options = plan_options();
        let plan = edit.plan(&options).unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output, options).unwrap();

        let source_backed =
            SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(output.clone())))
                .unwrap();
        let source_slide = source_backed.slides().unwrap().pop().unwrap();
        assert!(source_slide.text().unwrap().is_empty());
        let source_objects: Vec<_> = source_slide
            .shapes()
            .unwrap()
            .iter()
            .filter(|shape| shape.presentation_class() == Some("object"))
            .collect();
        assert_eq!(source_objects.len(), 1);
        assert_eq!(source_objects[0].text().unwrap(), body);

        let presentation = Presentation::from_bytes(output).unwrap();
        let slide = presentation.slides().unwrap().pop().unwrap();
        assert!(slide.text().unwrap().is_empty());
        let objects: Vec<_> = slide
            .shapes()
            .unwrap()
            .iter()
            .filter(|shape| shape.presentation_class() == Some("object"))
            .collect();
        assert_eq!(objects.len(), 1);
        assert_eq!(objects[0].text().unwrap(), body);
    }

    #[test]
    fn unicode_whitespace_is_a_typed_refusal() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let error = SourceBackedTailAppendEdit::from_read_at(
            source_package(&content),
            "title\u{2003}",
            "body",
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidFormat(message) if message.contains("Unicode whitespace")
        ));
    }

    #[test]
    fn large_space_runs_are_split_at_parser_limit() {
        let body = " ".repeat(MAX_TEXT_SPACE_COUNT + 1);
        let fragment = build_plain_page("page1", "", &body, MAX_FRAGMENT_BYTES).unwrap();
        let xml = std::str::from_utf8(&fragment).unwrap();
        assert_eq!(xml.matches(r#"<text:s text:c="#).count(), 2);
        assert!(xml.contains(r#"<text:s text:c="1000000"/>"#));
        assert!(!xml.contains(r#"text:c="1000001""#));
    }

    #[test]
    fn carriage_return_text_is_a_typed_refusal() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}"><office:body><office:presentation><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let error = SourceBackedTailAppendEdit::from_read_at(
            source_package(&content),
            "title\rwith carriage return",
            "body",
        )
        .unwrap_err();
        assert!(matches!(
            error,
            Error::InvalidFormat(message) if message.contains("carriage return")
        ));
    }

    #[test]
    fn leading_presentation_declaration_does_not_block_page_tail_append() {
        let content = format!(
            r#"<office:document-content xmlns:office="{OFFICE_NAMESPACE_UTF8}" xmlns:draw="{DRAW_NAMESPACE_UTF8}" xmlns:presentation="{PRESENTATION_NAMESPACE_UTF8}"><office:body><office:presentation><presentation:settings presentation:start-page="page1"/><draw:page draw:name="page1"/></office:presentation></office:body></office:document-content>"#,
        );
        let edit =
            SourceBackedTailAppendEdit::from_read_at(source_package(&content), "title", "body")
                .unwrap();
        let options = plan_options();
        let plan = edit.plan(&options).unwrap();
        let mut output = Vec::new();
        plan.write_to(&mut output, options).unwrap();

        let package = OwnedPackage::from_bytes(output).unwrap();
        let published = String::from_utf8(package.get_file(ODF_CONTENT).unwrap()).unwrap();
        let settings = published.find("<presentation:settings").unwrap();
        let generated = published.find(r#"draw:name="page2""#).unwrap();
        assert!(settings < generated);
    }
}
