//! Bounded source-backed append of one plain main-story paragraph.
//!
//! The operation is deliberately a small format-owned closure.  It accepts a
//! canonical WordprocessingML main story made from direct plain paragraphs,
//! preserves the final direct `w:sectPr` as an opaque lexical span, and
//! inserts one locally namespace-bound paragraph immediately before that span
//! (or before `w:body`'s closing tag when the section-properties element is
//! absent).  The source and candidate are scanned through a guarded
//! `BufRead`; neither complete XML document is retained or indexed.
//!
//! The source and candidate proofs are compact scalar evidence.  The OPC
//! publication layer performs its independent source/candidate compact audit
//! before emitting the changed archive.  Opaque `sectPr` preservation is thus
//! scoped to source XML admitted by both this grammar and the current compact
//! XML audit policy.  The opaque scope conservatively refuses non-whitespace
//! character events, unsupported XML event forms, and entity-escaped namespace
//! declarations so expanded-name identity remains bounded and authenticated.

use std::borrow::Cow;
use std::collections::TryReserveError;
use std::io::{self, BufRead, Cursor, Read, Write};
use std::mem::size_of;
use std::ops::Range;
use std::sync::{Arc, Mutex};

use litchi_core::{CancellationToken, ExecutionContext, ExecutionError, Resource, SourceVersion};
use litchi_ooxml_common::mce::{Limits as MceLimits, NAMESPACE as MCE_NAMESPACE};
use litchi_ooxml_common::xml_name::is_qualified_name;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::source_backed::VerifiedDecodedReaderError;
use litchi_opc::{
    SourceArtifactFingerprint, SourcePartSpliceFragment, SourcePartSpliceLimits,
    SourcePartSplicePlan, SourcePartSpliceProof, SourcePartSplicePublication,
};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;
use sha2::{Digest as _, Sha256};
use thiserror::Error;

use crate::error::Error as DocumentError;
use crate::namespace::{STRICT_WORDPROCESSINGML_NAMESPACE, WORDPROCESSINGML_NAMESPACE};
use crate::package::validate_document_main_content_type;
use crate::settings::DocumentSettings;
use crate::streaming::{append_character, escaped_character_len, is_plain_text_character};

use super::Package;

mod mce_workspace;
mod settings_workspace;

const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const XMLNS_NAMESPACE: &[u8] = b"http://www.w3.org/2000/xmlns/";
const INITIAL_QUICK_XML_NAMESPACE_BYTES: usize = 73;
const QUICK_XML_NAMESPACE_BINDING_BYTES: usize = size_of::<usize>() * 4;
const STRICT_SETTINGS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const STRICT_FOOTNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes";
const STRICT_ENDNOTES: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes";
const STRICT_CUSTOM_XML: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml";
const STRICT_ALTERNATIVE_FORMAT_IMPORT: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk";
const MAX_QUICK_XML_DEPTH: u64 = (u16::MAX as u64) - 1;

/// Result returned by the bounded tail-append transaction.
pub type Result<T> = std::result::Result<T, Error>;

/// Stable refusal for a package or story outside the one-paragraph closure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Refusal {
    /// The package's main relationship or part shape is outside the closure.
    MainDocumentShape,
    /// The package contains signature material.
    SignatureInfrastructure,
    /// The package contains encrypted ZIP members.
    Encrypted,
    /// The package contains macro-enabled content.
    MacroEnabled,
    /// The package contains an external relationship.
    ExternalRelationship,
    /// The package contains a dependency outside the plain-story closure.
    UnsupportedDependency,
    /// The settings part enables protection or tracked revisions.
    Protection,
    /// The main story is not the admitted direct-paragraph topology.
    ComplexDocument,
    /// A direct paragraph is not plain text.
    ComplexParagraph,
}

impl std::fmt::Display for Refusal {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::MainDocumentShape => "main document is not the canonical ordinary DOCX part",
            Self::SignatureInfrastructure => "package contains digital-signature infrastructure",
            Self::Encrypted => "package contains encrypted ZIP members",
            Self::MacroEnabled => "package is macro-enabled",
            Self::ExternalRelationship => "package contains an external relationship",
            Self::UnsupportedDependency => "package contains an unsupported story dependency",
            Self::Protection => "document or write protection is enforced",
            Self::ComplexDocument => "main story is not a direct plain-paragraph document",
            Self::ComplexParagraph => "paragraph is not composed only of plain runs and text",
        })
    }
}

/// Failure from a bounded source-backed tail append.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum Error {
    /// The underlying DOCX or OPC package failed.
    #[error(transparent)]
    Document(#[from] DocumentError),
    /// The package is outside the deliberately narrow closure.
    #[error("plain paragraph tail append refused: {0}")]
    Refused(Refusal),
    /// A finite operation ceiling was exceeded.
    #[error("plain paragraph tail append {resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        /// Resource whose ceiling was exceeded.
        resource: &'static str,
        /// First observed value beyond the ceiling.
        actual: u64,
        /// Configured ceiling.
        maximum: u64,
    },
    /// A bounded operation allocation failed.
    #[error("plain paragraph tail append allocation failed for {resource}: {source}")]
    Allocation {
        /// Allocation being attempted.
        resource: &'static str,
        /// Allocator failure.
        #[source]
        source: TryReserveError,
    },
    /// An execution context or cancellation token rejected the operation.
    #[error("plain paragraph tail append execution failed: {0}")]
    Execution(#[source] ExecutionError),
    /// A parser or source adapter failed before publication.
    #[error("plain paragraph tail append XML scan failed: {0}")]
    Scan(String),
    /// The bounded reader reported a transport or source-window I/O failure.
    #[error("plain paragraph tail append I/O failed: {0}")]
    Io(#[from] io::Error),
    /// Publication was attempted against a package that is no longer current.
    #[error("plain paragraph tail append source is stale")]
    StaleSource,
}

/// Finite bounds for one source-backed plain-paragraph tail append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum decoded `/word/document.xml` source bytes.
    pub max_source_xml_bytes: u64,
    /// Maximum UTF-8 input text bytes.
    pub max_text_bytes: u64,
    /// Maximum retained generated fragment bytes.
    pub max_fragment_bytes: u64,
    /// Maximum projected candidate XML bytes.
    pub max_candidate_xml_bytes: u64,
    /// Maximum parser events in either semantic pass.
    pub max_events: u64,
    /// Maximum XML nesting depth, including opaque section children.
    pub max_depth: u64,
    /// Maximum direct paragraphs accepted in the source.
    pub max_paragraphs: u64,
    /// Maximum settings XML bytes before and after markup-compatibility processing.
    pub max_settings_xml_bytes: u64,
    /// Maximum semantic parser workspace reservation.
    pub max_workspace_bytes: u64,
    /// Maximum complete physical output bytes delegated to OPC.
    pub max_output_bytes: u64,
    /// Maximum token bytes exposed to quick-xml for one event.
    pub max_token_bytes: u64,
}

impl Limits {
    /// Construct an explicit finite policy.
    #[must_use]
    pub const fn new(
        max_source_xml_bytes: u64,
        max_text_bytes: u64,
        max_fragment_bytes: u64,
        max_candidate_xml_bytes: u64,
        max_events: u64,
        max_depth: u64,
        max_paragraphs: u64,
        max_settings_xml_bytes: u64,
        max_workspace_bytes: u64,
        max_output_bytes: u64,
        max_token_bytes: u64,
    ) -> Self {
        Self {
            max_source_xml_bytes,
            max_text_bytes,
            max_fragment_bytes,
            max_candidate_xml_bytes,
            max_events,
            max_depth,
            max_paragraphs,
            max_settings_xml_bytes,
            max_workspace_bytes,
            max_output_bytes,
            max_token_bytes,
        }
    }

    fn validate(self) -> Result<()> {
        for (resource, value) in [
            ("source XML bytes", self.max_source_xml_bytes),
            ("text bytes", self.max_text_bytes),
            ("fragment bytes", self.max_fragment_bytes),
            ("candidate XML bytes", self.max_candidate_xml_bytes),
            ("XML events", self.max_events),
            ("XML depth", self.max_depth),
            ("paragraphs", self.max_paragraphs),
            ("settings XML bytes", self.max_settings_xml_bytes),
            ("parser workspace", self.max_workspace_bytes),
            ("output bytes", self.max_output_bytes),
            ("XML token bytes", self.max_token_bytes),
        ] {
            if value == 0 || value == u64::MAX {
                return Err(Error::Limit {
                    resource,
                    actual: value,
                    maximum: value.saturating_sub(1),
                });
            }
        }
        if self.max_depth > MAX_QUICK_XML_DEPTH {
            return Err(Error::Limit {
                resource: "XML depth",
                actual: self.max_depth,
                maximum: MAX_QUICK_XML_DEPTH,
            });
        }
        Ok(())
    }

    pub(super) fn validate_for_stream(self) -> Result<()> {
        self.validate()
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_source_xml_bytes: 32 * 1024 * 1024,
            max_text_bytes: 16 * 1024 * 1024,
            max_fragment_bytes: 16 * 1024 * 1024,
            max_candidate_xml_bytes: 48 * 1024 * 1024,
            max_events: 1_000_000,
            max_depth: 256,
            max_paragraphs: 65_536,
            max_settings_xml_bytes: 4 * 1024 * 1024,
            max_workspace_bytes: 128 * 1024 * 1024,
            max_output_bytes: 2 * 1024 * 1024 * 1024,
            max_token_bytes: 64 * 1024,
        }
    }
}

/// Additional cooperative cancellation policy for an edit.
///
/// Configure execution budgets on the source-backed [`Package`]. Its context
/// governs semantic scanning, fragment storage, and OPC publication together.
#[derive(Debug, Clone, Default)]
pub struct Options {
    execution: Option<ExecutionContext>,
    cancellation: Option<CancellationToken>,
}

impl Options {
    /// Attach an additional token checked at main-story and settings preflight
    /// parser events, owning settings-validation boundaries, and publication
    /// sink callbacks. The package's execution context also controls
    /// cancellation inside OPC validation and replay.
    #[must_use]
    pub fn with_cancellation_token(mut self, token: &CancellationToken) -> Self {
        self.cancellation = Some(token.clone());
        self
    }
}

/// Compact source-side semantic evidence for one tail append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceProof {
    /// Source package revision.
    pub source_version: SourceVersion,
    /// Decoded source XML byte length.
    pub source_len: u64,
    /// SHA-256 over exact decoded source XML bytes.
    pub source_sha256: [u8; 32],
    /// Insertion offset before direct `sectPr` or `body` close.
    pub insertion_offset: u64,
    /// Source direct paragraph count.
    pub paragraph_count: u64,
    /// Source parser-event count.
    pub event_count: u64,
    /// Maximum source depth observed.
    pub max_depth: u64,
    /// Whether the source uses strict WordprocessingML.
    pub strict_namespace: bool,
    /// Opaque direct section-properties length, if present.
    pub sect_pr_len: u64,
    /// SHA-256 over the exact opaque section-properties span.
    pub sect_pr_sha256: [u8; 32],
}

/// Compact candidate-side semantic evidence for one tail append.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CandidateProof {
    /// Candidate decoded XML byte length.
    pub candidate_len: u64,
    /// SHA-256 over exact decoded candidate XML bytes.
    pub candidate_sha256: [u8; 32],
    /// Candidate direct paragraph count.
    pub paragraph_count: u64,
    /// Candidate parser-event count.
    pub event_count: u64,
    /// Maximum candidate depth observed.
    pub max_depth: u64,
    /// Offset where the generated paragraph was admitted.
    pub generated_offset: u64,
    /// Whether the generated paragraph was seen exactly once.
    pub generated_once: bool,
    /// Candidate section-properties length, if present.
    pub sect_pr_len: u64,
    /// Candidate section-properties hash.
    pub sect_pr_sha256: [u8; 32],
}

/// A source-backed edit which appends one plain paragraph or performs an
/// explicit exact no-op publication.
pub struct Edit<'package, 'text> {
    package: &'package Package,
    text: Option<Cow<'text, str>>,
    limits: Limits,
    options: Options,
}

impl<'package, 'text> Edit<'package, 'text> {
    /// Replace the edit's finite policy before preparation.
    #[must_use]
    pub fn with_limits(mut self, limits: Limits) -> Self {
        self.limits = limits;
        self
    }

    /// Attach additional cooperative cancellation options.
    #[must_use]
    pub fn with_options(mut self, options: Options) -> Self {
        self.options = options;
        self
    }

    /// Prepare source and candidate semantic proofs and an OPC splice plan.
    pub fn prepare(self) -> Result<Plan<'package>> {
        prepare_edit(self)
    }

    /// Alias for [`Self::prepare`] with explicit options.
    pub fn prepare_with_options(mut self, options: Options) -> Result<Plan<'package>> {
        self.options = options;
        self.prepare()
    }

    /// Prepare the edit as an explicit commit product.
    pub fn commit(self) -> Result<Commit<'package>> {
        Ok(self.prepare()?.commit())
    }

    /// Prepare the edit and immediately retain the commit product under
    /// explicit options.
    pub fn commit_with_options(self, options: Options) -> Result<Commit<'package>> {
        Ok(self.prepare_with_options(options)?.commit())
    }
}

/// Prepared source/candidate semantic proof plus OPC splice plan.
pub struct Plan<'package> {
    splice: SourcePartSplicePlan<'package>,
    source: SourceProof,
    candidate: CandidateProof,
    limits: Limits,
    options: Options,
}

impl<'package> Plan<'package> {
    /// Borrow the source semantic proof.
    #[must_use]
    pub const fn source_proof(&self) -> SourceProof {
        self.source
    }

    /// Borrow the candidate semantic proof.
    #[must_use]
    pub const fn candidate_proof(&self) -> CandidateProof {
        self.candidate
    }

    /// Return the low-level splice proof used by OPC.
    #[must_use]
    pub const fn splice_proof(&self) -> SourcePartSpliceProof {
        self.splice.proof()
    }

    /// Return the operation's finite policy.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return whether this is the explicit exact no-op route.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.splice.is_noop()
    }

    /// Consume the plan into a named commit product.
    #[must_use]
    pub fn commit(self) -> Commit<'package> {
        Commit { plan: self }
    }

    /// Publish the prepared plan to a sequential sink.
    pub fn write_to_stream(self, writer: impl Write) -> Result<Publication> {
        self.commit().write_to_stream(writer)
    }
}

/// Named commit product retaining the reversible source authorization.
pub struct Commit<'package> {
    plan: Plan<'package>,
}

impl<'package> Commit<'package> {
    /// Borrow the prepared plan's source proof.
    #[must_use]
    pub const fn source_proof(&self) -> SourceProof {
        self.plan.source
    }

    /// Borrow the prepared plan's candidate proof.
    #[must_use]
    pub const fn candidate_proof(&self) -> CandidateProof {
        self.plan.candidate
    }

    /// Return whether this commit is the exact no-op route.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.plan.splice.is_noop()
    }

    /// Publish and retain an exact inverse authorization.
    pub fn write_to_stream(self, writer: impl Write) -> Result<Publication> {
        publish_commit(self, writer)
    }
}

/// Exact publication evidence and immediate inverse authorization.
#[derive(Debug)]
pub struct Publication {
    source: SourceProof,
    candidate: CandidateProof,
    splice: SourcePartSplicePublication,
}

impl Publication {
    /// Borrow the source semantic proof.
    #[must_use]
    pub const fn source_proof(&self) -> SourceProof {
        self.source
    }

    /// Borrow the candidate semantic proof.
    #[must_use]
    pub const fn candidate_proof(&self) -> CandidateProof {
        self.candidate
    }

    /// Return the exact published archive fingerprint.
    #[must_use]
    pub const fn candidate_artifact_fingerprint(&self) -> SourceArtifactFingerprint {
        self.splice.candidate_artifact_fingerprint()
    }

    /// Return whether the publication copied the exact source archive.
    #[must_use]
    pub const fn is_noop(&self) -> bool {
        self.splice.is_noop()
    }

    /// Restore the exact source archive after checking the current package's
    /// complete artifact fingerprint. This is a new operation governed by the
    /// current package's execution context, including its cancellation token.
    pub fn write_inverse_to_stream(&self, current: &Package, writer: impl Write) -> Result<()> {
        self.splice
            .write_inverse_to_stream(&current.package, writer)
            .map_err(|error| Error::Document(DocumentError::from(error)))
    }
}

impl Package {
    /// Start a bounded source-backed append of one plain paragraph.
    #[must_use]
    pub fn tail_append_plain_paragraph<'package, 'text, T>(
        &'package self,
        text: T,
    ) -> Edit<'package, 'text>
    where
        T: Into<Cow<'text, str>>,
    {
        Edit {
            package: self,
            text: Some(text.into()),
            limits: Limits::default(),
            options: Options::default(),
        }
    }

    /// Start a bounded source-backed append under an explicit finite policy.
    pub fn tail_append_plain_paragraph_with_limits<'package, 'text, T>(
        &'package self,
        text: T,
        limits: Limits,
    ) -> Result<Edit<'package, 'text>>
    where
        T: Into<Cow<'text, str>>,
    {
        limits.validate()?;
        Ok(Edit {
            package: self,
            text: Some(text.into()),
            limits,
            options: Options::default(),
        })
    }

    /// Start the explicit exact archive no-op route.
    #[must_use]
    pub fn tail_append_noop<'package>(&'package self) -> Edit<'package, 'package> {
        Edit {
            package: self,
            text: None,
            limits: Limits::default(),
            options: Options::default(),
        }
    }
}

fn prepare_edit<'package, 'text>(mut edit: Edit<'package, 'text>) -> Result<Plan<'package>> {
    edit.limits.validate()?;
    let package = edit.package;
    check_options(&edit.options)?;
    let options = effective_options(package, &edit.options);
    check_options(&options)?;
    let source_version = package
        .package
        .source_version()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;

    if edit.text.is_none() {
        return prepare_noop(package, source_version, edit.limits, options);
    }

    let (text_bytes, text_capacity) = match edit.text.as_ref() {
        Some(Cow::Borrowed(text)) => (text.len(), 0),
        Some(Cow::Owned(text)) => (text.len(), text.capacity()),
        None => (0, 0),
    };
    let text_bytes = u64::try_from(text_bytes).map_err(|_| Error::Limit {
        resource: "text bytes",
        actual: u64::MAX,
        maximum: edit.limits.max_text_bytes,
    })?;
    check_limit("text bytes", text_bytes, edit.limits.max_text_bytes)?;
    let text_capacity = u64::try_from(text_capacity).map_err(|_| Error::Limit {
        resource: "text capacity bytes",
        actual: u64::MAX,
        maximum: edit.limits.max_text_bytes,
    })?;
    check_limit(
        "text capacity bytes",
        text_capacity,
        edit.limits.max_text_bytes,
    )?;
    let text = edit.text.take().ok_or(Error::Scan(
        "tail append text disappeared during admission".into(),
    ))?;

    if package.package.has_encrypted_entries() {
        return Err(Error::Refused(Refusal::Encrypted));
    }

    let text_reservation = options
        .execution
        .as_ref()
        .filter(|_| text_capacity != 0)
        .map(|execution| {
            execution
                .reserve(Resource::Memory, text_capacity)
                .map(Arc::new)
                .map_err(Error::Execution)
        })
        .transpose()?;

    validate_topology(
        package,
        edit.limits,
        options.execution.as_ref(),
        options.cancellation.as_ref(),
    )?;
    let main = package
        .package
        .main_document_part()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    validate_document_main_content_type(main.content_type()).map_err(Error::Document)?;
    let source_len =
        checked_declared_size(&main, edit.limits.max_source_xml_bytes, "source XML bytes")?;
    let source_scan = scan_main_part(
        &main,
        source_len,
        edit.limits,
        options.execution.as_ref(),
        options.cancellation.as_ref(),
        None,
    )?;
    if package_strict_dialect(package.package.rels())? != Some(source_scan.strict_namespace) {
        return Err(Error::Refused(Refusal::MainDocumentShape));
    }
    let fragment = encode_fragment(
        &package.package,
        text.as_ref(),
        source_scan.strict_namespace,
        edit.limits,
        options.execution.as_ref(),
        options.cancellation.as_ref(),
    )?;
    drop(text);
    drop(text_reservation);
    let fragment_len = u64::try_from(fragment.as_slice().len()).map_err(|_| Error::Limit {
        resource: "fragment bytes",
        actual: u64::MAX,
        maximum: edit.limits.max_fragment_bytes,
    })?;
    let candidate_len = source_len.checked_add(fragment_len).ok_or(Error::Limit {
        resource: "candidate XML bytes",
        actual: u64::MAX,
        maximum: edit.limits.max_candidate_xml_bytes,
    })?;
    check_limit(
        "candidate XML bytes",
        candidate_len,
        edit.limits.max_candidate_xml_bytes,
    )?;
    let candidate_scan = scan_main_part(
        &main,
        source_len,
        edit.limits,
        options.execution.as_ref(),
        options.cancellation.as_ref(),
        Some((&source_scan, fragment.as_slice())),
    )?;
    validate_candidate(&source_scan, &candidate_scan, fragment_len, edit.limits)?;

    let fragment_hash = digest_bytes(fragment.as_slice());
    let splice_proof = SourcePartSpliceProof {
        source_version,
        source_len,
        source_sha256: source_scan.sha256,
        insertion_offset: source_scan.insertion_offset,
        fragment_len,
        fragment_sha256: fragment_hash,
        candidate_len,
        candidate_sha256: candidate_scan.sha256,
    };
    let splice_limits = make_splice_limits(edit.limits)?;
    let splice = package
        .package
        .prepare_source_part_splice_with_fragment(
            main.partname(),
            splice_proof,
            fragment,
            splice_limits,
        )
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    package
        .package
        .source_version()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    check_options(&options)?;
    Ok(Plan {
        splice,
        source: source_scan.source_proof(source_version),
        candidate: candidate_scan.candidate_proof(),
        limits: edit.limits,
        options,
    })
}

fn effective_options(package: &Package, requested: &Options) -> Options {
    let mut options = requested.clone();
    // One source-owned execution authority covers both format and package
    // work. An edit cannot replace or bypass its hierarchical budget.
    options.execution = package.package.execution_context();
    options
}

fn prepare_noop(
    package: &Package,
    source_version: SourceVersion,
    limits: Limits,
    options: Options,
) -> Result<Plan<'_>> {
    let main = package
        .package
        .main_document_part()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    let source_len = checked_declared_size(&main, limits.max_source_xml_bytes, "source XML bytes")?;
    let charge_hash_work = package.package.execution_context().is_none();
    let hash_context = options.execution.as_ref();
    let hash_cancellation = options.cancellation.as_ref();
    let (source_sha256, observed_len) = main
        .with_verified_decoded_reader(|reader| {
            hash_reader(reader, hash_context, hash_cancellation, charge_hash_work)
        })
        .map_err(map_hash_verified_error)?;
    if observed_len != source_len {
        return Err(Error::Scan(
            "source XML length changed during no-op proof".into(),
        ));
    }
    let splice_proof = SourcePartSpliceProof {
        source_version,
        source_len,
        source_sha256,
        insertion_offset: 0,
        fragment_len: 0,
        fragment_sha256: digest_bytes(&[]),
        candidate_len: source_len,
        candidate_sha256: source_sha256,
    };
    let splice = package
        .package
        .prepare_source_part_splice(
            main.partname(),
            splice_proof,
            Arc::new(Vec::new()),
            make_splice_limits(limits)?,
        )
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    let source = SourceProof {
        source_version,
        source_len,
        source_sha256,
        insertion_offset: 0,
        paragraph_count: 0,
        event_count: 0,
        max_depth: 0,
        strict_namespace: false,
        sect_pr_len: 0,
        sect_pr_sha256: digest_bytes(&[]),
    };
    let candidate = CandidateProof {
        candidate_len: source_len,
        candidate_sha256: source_sha256,
        paragraph_count: 0,
        event_count: 0,
        max_depth: 0,
        generated_offset: 0,
        generated_once: false,
        sect_pr_len: 0,
        sect_pr_sha256: digest_bytes(&[]),
    };
    check_options(&options)?;
    Ok(Plan {
        splice,
        source,
        candidate,
        limits,
        options,
    })
}

fn publish_commit(commit: Commit<'_>, writer: impl Write) -> Result<Publication> {
    let Commit { plan } = commit;
    check_options(&plan.options)?;
    let source = plan.source;
    let candidate = plan.candidate;
    let guarded_result = if plan.options.execution.is_none() && plan.options.cancellation.is_none()
    {
        (plan.splice.write_to_stream(writer), None)
    } else {
        let failure = Arc::new(Mutex::new(None));
        let guarded = PublicationCheckedWriter {
            inner: writer,
            execution: plan.options.execution.clone(),
            cancellation: plan.options.cancellation.clone(),
            failure: Arc::clone(&failure),
        };
        let result = plan.splice.write_to_stream(guarded);
        let failure = failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take();
        (result, failure)
    };
    let splice = match guarded_result.0 {
        Ok(splice) => splice,
        Err(error) => return Err(Error::Document(DocumentError::from(error))),
    };
    if let Some(error) = guarded_result.1 {
        return Err(Error::Execution(error));
    }
    Ok(Publication {
        source,
        candidate,
        splice,
    })
}

struct PublicationCheckedWriter<W> {
    inner: W,
    execution: Option<ExecutionContext>,
    cancellation: Option<CancellationToken>,
    failure: Arc<Mutex<Option<ExecutionError>>>,
}

impl<W> PublicationCheckedWriter<W> {
    fn check(&self) -> io::Result<()> {
        if let Some(execution) = self.execution.as_ref() {
            if let Err(error) = execution.check() {
                return self.fail(error);
            }
        }
        if let Some(cancellation) = self.cancellation.as_ref() {
            if let Err(error) = cancellation.check() {
                return self.fail(error);
            }
        }
        Ok(())
    }

    fn fail(&self, error: ExecutionError) -> io::Result<()> {
        let mut failure = self
            .failure
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if failure.is_none() {
            *failure = Some(error.clone());
        }
        Err(io::Error::other(error))
    }
}

impl<W: Write> Write for PublicationCheckedWriter<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.check()?;
        self.inner.write(bytes)
    }

    fn flush(&mut self) -> io::Result<()> {
        self.check()?;
        self.inner.flush()
    }
}

fn check_options(options: &Options) -> Result<()> {
    if let Some(context) = options.execution.as_ref() {
        context.check().map_err(Error::Execution)?;
    }
    if let Some(token) = options.cancellation.as_ref() {
        token.check().map_err(Error::Execution)?;
    }
    Ok(())
}

fn check_progress(
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    work: u64,
) -> std::result::Result<(), ScanError> {
    if let Some(context) = context {
        context.check().map_err(ScanError::Execution)?;
        if work != 0 {
            context
                .consume(Resource::Work, work)
                .map_err(ScanError::Execution)?;
        }
    }
    if let Some(token) = cancellation {
        token.check().map_err(ScanError::Execution)?;
    }
    Ok(())
}

fn check_limit(resource: &'static str, actual: u64, maximum: u64) -> Result<()> {
    if actual > maximum {
        return Err(Error::Limit {
            resource,
            actual,
            maximum,
        });
    }
    Ok(())
}

/// Source-derived facts collected by the settings preflight.  The subsequent
/// MCE and settings codecs still perform their complete semantic validation;
/// these values only describe the largest owners they may create from the
/// already authenticated source.  In particular, namespace URI references
/// are counted at each element/attribute use, rather than multiplying a
/// configured token ceiling by a source-sized node count.
#[derive(Debug, Clone, Copy, Default)]
struct SettingsGuardFacts {
    max_token: usize,
    max_depth: usize,
    max_namespace_buffer: usize,
    max_namespace_bindings: usize,
    namespace_declaration_count: usize,
    namespace_copy_bytes: usize,
    event_count: u64,
    node_count: usize,
    attribute_count: usize,
    max_attributes_per_event: usize,
    attribute_bearing_nodes: usize,
    semantic_owned_bytes: usize,
    mce_directive_tokens: usize,
    mce_directive_owned_bytes: usize,
}

/// Return the checked allocation envelope for one `quick-xml` namespace pass.
///
/// The reader's event window is only one owner.  `NsReader` also retains the
/// opened names, namespace resolver bytes/bindings, and one-level attribute
/// duplicate scratch while the parser is at its admitted maximum depth.  The
/// source adapter owns a capture and exposed token window, and this route owns
/// the semantic frame stack.  Every term is checked before it is added so a
/// policy overflow is refused before constructing the parser.
fn scanner_workspace_requirement(max_token: usize, max_depth: usize) -> Option<u64> {
    const MAX_NAMESPACE_DECLARATIONS_PER_LEVEL: usize = 256;
    const INITIAL_NAMESPACE_BYTES: usize = 73;

    let token = max_token.checked_add(1)?;
    let levels = max_depth.checked_add(1)?;
    let namespace_declarations = MAX_NAMESPACE_DECLARATIONS_PER_LEVEL.min(token);
    // Opaque `sectPr` attributes are admitted up to the token window and are
    // copied into the duplicate-check scratch vector before validation.
    let attributes = token;
    let usize_bytes = size_of::<usize>();
    let binding_bytes = usize_bytes.checked_mul(4)?;
    let range_bytes = size_of::<Range<usize>>();
    let frame_bytes = size_of::<FrameKind>();
    let hash_entry_bytes = size_of::<u64>().checked_add(size_of::<u8>())?;
    let seen_attribute_bytes = size_of::<(&[u8], &[u8])>();

    let token_windows = token.checked_mul(4)?;
    let opened_names = levels.checked_mul(token)?.checked_mul(2)?.max(8);
    let opened_indexes = 4usize
        .checked_mul(usize_bytes)?
        .max(2usize.checked_mul(levels)?.checked_mul(usize_bytes)?);
    let namespace_bytes = INITIAL_NAMESPACE_BYTES
        .checked_add(levels.checked_mul(token)?)?
        .checked_mul(2)?
        .max(128);
    let namespace_binding_floor = 8usize.checked_mul(binding_bytes)?;
    let namespace_bindings = 2usize
        .checked_add(levels.checked_mul(namespace_declarations)?)?
        .checked_mul(2)?
        .checked_mul(binding_bytes)
        .map(|value| value.max(namespace_binding_floor))?;
    let attribute_ranges = attributes
        .checked_add(1)?
        .checked_mul(range_bytes)?
        .checked_mul(2)?
        .max(4usize.checked_mul(range_bytes)?);
    let attribute_hash = attributes
        .checked_add(1)?
        .checked_mul(hash_entry_bytes)?
        .checked_mul(4)?
        .max(8usize.checked_mul(hash_entry_bytes)?);
    let attribute_seen = attributes
        .checked_add(1)?
        .checked_mul(seen_attribute_bytes)?
        .checked_mul(2)?
        .max(8usize.checked_mul(seen_attribute_bytes)?);
    let scope_stack = levels
        .checked_mul(frame_bytes)?
        .checked_mul(2)?
        .max(8usize.checked_mul(frame_bytes)?);
    let namespace_scope_stack = levels
        .checked_mul(size_of::<(usize, usize)>())?
        .checked_mul(2)?;

    [
        token_windows,
        opened_names,
        opened_indexes,
        namespace_bytes,
        namespace_bindings,
        attribute_ranges,
        attribute_hash,
        attribute_seen,
        scope_stack,
        namespace_scope_stack,
        3,
    ]
    .into_iter()
    .try_fold(0usize, usize::checked_add)
    .and_then(|value| u64::try_from(value).ok())
}

fn bounded_settings_mce_limits(
    limits: Limits,
    source_bytes: usize,
    observed_depth: usize,
    observed_namespace_bindings: usize,
    observed_directive_tokens: usize,
) -> Result<MceLimits> {
    let configured_depth = usize::try_from(limits.max_depth).map_err(|_| Error::Limit {
        resource: "XML depth",
        actual: limits.max_depth,
        maximum: usize::MAX as u64,
    })?;
    let source_depth = source_bytes.checked_div(3).unwrap_or(0).max(1);
    let max_depth = configured_depth
        .min(source_depth)
        .min(observed_depth.max(1));
    let max_namespace_bindings = MceLimits::default()
        .max_namespace_bindings
        .min(source_bytes.max(1))
        .min(observed_namespace_bindings.max(1));
    let max_directive_tokens = MceLimits::default()
        .max_directive_tokens
        .min(source_bytes.max(1))
        .min(observed_directive_tokens.max(1));
    let max_choices_per_alternate = MceLimits::default()
        .max_choices_per_alternate
        .min(source_bytes.max(1));
    // MCE writes inherited namespace declarations onto emitted elements, so
    // valid processed XML can exceed its source length. Apply the explicit
    // settings-byte ceiling to that output and reserve it before processing.
    let max_output_bytes =
        usize::try_from(limits.max_settings_xml_bytes).map_err(|_| Error::Limit {
            resource: "settings XML bytes",
            actual: limits.max_settings_xml_bytes,
            maximum: usize::MAX as u64,
        })?;
    Ok(MceLimits {
        max_input_bytes: source_bytes,
        max_output_bytes,
        max_depth,
        max_namespace_bindings,
        max_directive_tokens,
        max_choices_per_alternate,
    })
}

fn mce_workspace_requirement(
    limits: Limits,
    source_bytes: usize,
    output_bytes: usize,
    facts: SettingsGuardFacts,
) -> Result<u64> {
    let overflow = || Error::Limit {
        resource: "settings XML workspace",
        actual: u64::MAX,
        maximum: limits.max_workspace_bytes,
    };
    let max_token_bytes = facts.max_token.max(1).min(source_bytes.max(1));
    let mce = mce_workspace::memory_requirement(mce_workspace::Profile {
        source_bytes,
        output_bytes,
        max_token_bytes,
        max_depth: facts.max_depth.max(1),
        max_namespace_bindings: facts.max_namespace_bindings.max(2),
        max_directive_tokens_per_event: facts.mce_directive_tokens.max(1),
        max_choices_per_alternate: source_bytes.max(1),
        namespace_bytes: facts
            .max_namespace_buffer
            .max(INITIAL_QUICK_XML_NAMESPACE_BYTES),
        namespace_declarations: facts.namespace_declaration_count,
        directive_tokens: facts.mce_directive_tokens,
        directive_owned_bytes: facts.mce_directive_owned_bytes,
        // Raw MCE attributes include xmlns declarations; the semantic facts
        // exclude those, so use the authenticated token-byte ceiling here.
        max_attributes_per_event: max_token_bytes,
    })
    .ok_or_else(overflow)?;
    Ok(mce)
}

fn settings_model_workspace_requirement(
    limits: Limits,
    source_bytes: usize,
    facts: SettingsGuardFacts,
) -> Result<u64> {
    let overflow = || Error::Limit {
        resource: "settings XML workspace",
        actual: u64::MAX,
        maximum: limits.max_workspace_bytes,
    };
    settings_workspace::model_memory_requirement(source_bytes, &facts).ok_or_else(overflow)
}

fn make_splice_limits(limits: Limits) -> Result<SourcePartSpliceLimits> {
    let max_source = limits.max_source_xml_bytes;
    let max_fragment = limits.max_fragment_bytes;
    let max_candidate = limits.max_candidate_xml_bytes;
    let mut splice = SourcePartSpliceLimits::new(
        max_source,
        max_fragment,
        max_candidate,
        limits.max_output_bytes,
        limits.max_output_bytes,
        limits.max_output_bytes,
    )
    .map_err(|error| Error::Document(DocumentError::from(error)))?;
    let to_usize = |value: u64| usize::try_from(value).unwrap_or(usize::MAX);
    let audit = xml_minifier::audit::Limits::new(
        to_usize(limits.max_candidate_xml_bytes),
        to_usize(limits.max_depth),
        to_usize(limits.max_events),
        // This is an aggregate count, not a per-event count. Each attribute
        // consumes source bytes; retain the audit owner's immutable ceiling.
        // Its duplicate scratch is capped independently by the token window.
        to_usize(limits.max_candidate_xml_bytes)
            .min(xml_minifier::audit::Limits::ATTRIBUTE_CEILING),
        to_usize(limits.max_token_bytes),
        to_usize(limits.max_candidate_xml_bytes),
    )
    .map_err(|error| Error::Scan(error.to_string()))?;
    splice = splice.with_xml_audit_limits(audit);
    splice.max_xml_workspace_bytes = limits.max_workspace_bytes;
    Ok(splice)
}

/// Build the OPC splice policy for the multi-paragraph stream route.
///
/// The stream owner uses the same source/candidate/output and XML-audit
/// ceilings as the fixed one-paragraph operation.  The replay adapter adds
/// its own bounded reader window at the caller boundary.
pub(super) fn make_splice_limits_for_stream(limits: Limits) -> Result<SourcePartSpliceLimits> {
    make_splice_limits(limits)
}

fn checked_declared_size(
    part: &litchi_opc::PartView<'_>,
    maximum: u64,
    resource: &'static str,
) -> Result<u64> {
    let declared = part
        .declared_uncompressed_size()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    check_limit(resource, declared, maximum)?;
    Ok(declared)
}

#[derive(Debug)]
enum HashReaderError {
    Io(io::Error),
    Execution(ExecutionError),
}

impl std::fmt::Display for HashReaderError {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Io(error) => error.fmt(formatter),
            Self::Execution(error) => error.fmt(formatter),
        }
    }
}

impl std::error::Error for HashReaderError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(error) => Some(error),
            Self::Execution(error) => Some(error),
        }
    }
}

fn map_hash_verified_error(error: VerifiedDecodedReaderError<HashReaderError>) -> Error {
    match error {
        VerifiedDecodedReaderError::Callback(HashReaderError::Io(callback)) => Error::Io(callback),
        VerifiedDecodedReaderError::Callback(HashReaderError::Execution(error)) => {
            Error::Execution(error)
        },
        VerifiedDecodedReaderError::Opc {
            error,
            callback_error: _,
        } => Error::Document(DocumentError::from(error)),
        _ => Error::Scan("verified OPC reader returned an unknown failure".into()),
    }
}

fn hash_reader(
    reader: &mut dyn BufRead,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    charge_work: bool,
) -> std::result::Result<([u8; 32], u64), HashReaderError> {
    let mut hasher = Sha256::new();
    let mut length = 0_u64;
    loop {
        if let Some(context) = context {
            context.check().map_err(HashReaderError::Execution)?;
        }
        if let Some(cancellation) = cancellation {
            cancellation.check().map_err(HashReaderError::Execution)?;
        }
        let available = reader.fill_buf().map_err(HashReaderError::Io)?;
        if available.is_empty() {
            break;
        }
        let count = u64::try_from(available.len()).map_err(|_| {
            HashReaderError::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "decoded XML length overflows u64",
            ))
        })?;
        length = length.checked_add(count).ok_or_else(|| {
            HashReaderError::Io(io::Error::new(
                io::ErrorKind::InvalidData,
                "decoded XML length overflows u64",
            ))
        })?;
        hasher.update(available);
        let available_len = available.len();
        reader.consume(available_len);
        if charge_work {
            if let Some(context) = context {
                context
                    .consume(Resource::Work, count)
                    .map_err(HashReaderError::Execution)?;
            }
        }
    }
    Ok((hasher.finalize().into(), length))
}

fn digest_bytes(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

pub(super) fn validate_topology(
    package: &Package,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<()> {
    check_options(&Options {
        execution: context.cloned(),
        cancellation: cancellation.cloned(),
    })?;
    let main = package
        .package
        .main_document_part()
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    if main.partname().as_str() != "/word/document.xml" {
        return Err(Error::Refused(Refusal::MainDocumentShape));
    }
    if package
        .package
        .physical_member_names()
        .any(is_signature_physical_member)
    {
        return Err(Error::Refused(Refusal::SignatureInfrastructure));
    }
    let package_strict = package_strict_dialect(package.package.rels())?;
    if matches!(
        main.content_type(),
        ct::WML_DOCUMENT_MACRO_MAIN | ct::WML_TEMPLATE_MACRO_MAIN
    ) {
        return Err(Error::Refused(Refusal::MacroEnabled));
    }
    if main.content_type() != ct::WML_DOCUMENT_MAIN {
        return Err(Error::Refused(Refusal::MainDocumentShape));
    }

    for relationship in package.package.rels().iter() {
        validate_relationship(relationship)?;
    }
    let mut settings = None;
    for part in package.package.iter_parts() {
        check_progress(context, cancellation, 1).map_err(map_scan_error)?;
        validate_part(part.partname().as_str(), part.content_type())?;
        for relationship in part.rels().iter() {
            validate_relationship(relationship)?;
            if part.partname() == main.partname()
                && matches!(relationship.reltype(), rt::SETTINGS | STRICT_SETTINGS)
            {
                if settings
                    .replace((
                        relationship
                            .target_partname()
                            .map_err(|error| Error::Document(DocumentError::from(error)))?,
                        relationship.reltype() == STRICT_SETTINGS,
                    ))
                    .is_some()
                {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
            }
        }
    }
    if let Some((settings_name, settings_strict)) = settings {
        if package_strict != Some(settings_strict) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let settings_part = package
            .package
            .part(&settings_name)
            .map_err(|error| Error::Document(DocumentError::from(error)))?;
        if settings_part.content_type() != ct::WML_SETTINGS {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        checked_declared_size(
            &settings_part,
            limits.max_settings_xml_bytes,
            "settings XML bytes",
        )?;
        let managed_context = context
            .cloned()
            .or_else(|| package.package.execution_context());
        let settings_data = settings_part
            .data()
            .map_err(|error| Error::Document(DocumentError::from(error)))?;
        let settings_bytes = settings_data.as_bytes();
        let settings_len = u64::try_from(settings_bytes.len()).map_err(|_| Error::Limit {
            resource: "settings XML bytes",
            actual: u64::MAX,
            maximum: limits.max_settings_xml_bytes,
        })?;
        check_limit(
            "settings XML bytes",
            settings_len,
            limits.max_settings_xml_bytes,
        )?;
        if inspect_root_dialect(
            settings_bytes,
            b"settings",
            settings_strict,
            limits,
            managed_context.as_ref(),
            cancellation,
        )? != settings_strict
        {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let guard_facts = guard_settings_xml(
            settings_bytes,
            limits,
            managed_context.as_ref(),
            cancellation,
        )?;
        let mce_limits = bounded_settings_mce_limits(
            limits,
            settings_bytes.len(),
            guard_facts.max_depth,
            // The MCE context includes its implicit XML binding. The guard's
            // live binding count includes implicit bindings as well as source
            // declarations; the declaration-only total omits that owner.
            guard_facts.max_namespace_bindings,
            guard_facts.mce_directive_tokens,
        )?;
        let workspace = mce_workspace_requirement(
            limits,
            settings_bytes.len(),
            mce_limits.max_output_bytes,
            guard_facts,
        )?;
        check_limit(
            "settings XML workspace",
            workspace,
            limits.max_workspace_bytes,
        )?;
        let output_bound = mce_workspace::output_memory_requirement(mce_limits.max_output_bytes)
            .ok_or(Error::Limit {
                resource: "settings XML workspace",
                actual: u64::MAX,
                maximum: limits.max_workspace_bytes,
            })?;
        let scratch_bound = workspace
            .checked_sub(output_bound)
            .ok_or_else(|| Error::Scan("MCE workspace does not include its output owner".into()))?;
        let reserve = |bytes| {
            managed_context
                .as_ref()
                .map(|execution| {
                    execution
                        .reserve(Resource::Memory, bytes)
                        .map_err(Error::Execution)
                })
                .transpose()
        };
        // Separate leases allow the MCE working state to be released while
        // its processed XML remains live during the next guard/model phases.
        let mut output_reservation = reserve(output_bound)?;
        let mce_scratch_reservation = reserve(scratch_bound)?;
        check_progress(
            managed_context.as_ref(),
            cancellation,
            guard_facts.event_count,
        )
        .map_err(map_scan_error)?;
        let processed =
            DocumentSettings::process_bytes_with_mce_limits(settings_bytes, &mce_limits)
                .map_err(Error::Document)?;
        check_progress(managed_context.as_ref(), cancellation, 0).map_err(map_scan_error)?;
        drop(mce_scratch_reservation);
        let retained_output = if matches!(processed, Cow::Borrowed(_)) {
            drop(output_reservation.take());
            0
        } else {
            output_bound
        };
        let mut guard_limits = limits;
        guard_limits.max_workspace_bytes = limits
            .max_workspace_bytes
            .checked_sub(retained_output)
            .ok_or(Error::Limit {
                resource: "settings XML workspace",
                actual: retained_output,
                maximum: limits.max_workspace_bytes,
            })?;
        // MCE may reinject namespace declarations. Recount its actual output
        // instead of using source facts to bound a different parser input.
        let processed_facts = guard_settings_xml(
            processed.as_ref(),
            guard_limits,
            managed_context.as_ref(),
            cancellation,
        )?;
        let model_workspace =
            settings_model_workspace_requirement(limits, processed.len(), processed_facts)?;
        let model_peak = retained_output
            .checked_add(model_workspace)
            .ok_or(Error::Limit {
                resource: "settings XML workspace",
                actual: u64::MAX,
                maximum: limits.max_workspace_bytes,
            })?;
        check_limit(
            "settings XML workspace",
            model_peak,
            limits.max_workspace_bytes,
        )?;
        let _model_reservation = reserve(model_workspace)?;
        // Prepay one event unit for each of the three owning XML passes.
        // Their cancellation boundary is the complete bounded model phase.
        let model_work = processed_facts
            .event_count
            .checked_mul(3)
            .ok_or(Error::Limit {
                resource: "settings validation work",
                actual: u64::MAX,
                maximum: u64::MAX - 1,
            })?;
        check_progress(managed_context.as_ref(), cancellation, model_work)
            .map_err(map_scan_error)?;
        let inspected = DocumentSettings::extract_from_processed_xml_with_relationships(
            processed.as_ref(),
            settings_part.rels(),
        )
        .map_err(Error::Document)?;
        check_progress(managed_context.as_ref(), cancellation, 0).map_err(map_scan_error)?;
        if inspected.is_protected() || inspected.is_write_protected() || inspected.track_revisions()
        {
            return Err(Error::Refused(Refusal::Protection));
        }
    }
    Ok(())
}

fn package_strict_dialect(rels: &litchi_opc::Relationships) -> Result<Option<bool>> {
    let mut result = None;
    for relationship in rels.iter() {
        if matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        ) {
            if result
                .replace(relationship.reltype() == rt::STRICT_OFFICE_DOCUMENT)
                .is_some()
            {
                return Err(Error::Refused(Refusal::MainDocumentShape));
            }
        }
    }
    Ok(result)
}

fn validate_relationship(relationship: &litchi_opc::Relationship) -> Result<()> {
    if relationship.is_external() {
        return Err(Error::Refused(Refusal::ExternalRelationship));
    }
    if is_signature_relationship(relationship.reltype()) {
        return Err(Error::Refused(Refusal::SignatureInfrastructure));
    }
    if matches!(
        relationship.reltype(),
        rt::VBA_PROJECT
            | rt::VBA_PROJECT_SIGNATURE
            | rt::VBA_PROJECT_SIGNATURE_AGILE
            | rt::WORD_VBA_DATA
    ) {
        return Err(Error::Refused(Refusal::MacroEnabled));
    }
    if matches!(
        relationship.reltype(),
        rt::COMMENTS
            | rt::STRICT_COMMENTS
            | rt::FOOTNOTES
            | rt::ENDNOTES
            | rt::CUSTOM_XML
            | rt::ALTERNATIVE_FORMAT_IMPORT
            | rt::MS_ALTERNATIVE_FORMAT_IMPORT
            | STRICT_FOOTNOTES
            | STRICT_ENDNOTES
            | STRICT_CUSTOM_XML
            | STRICT_ALTERNATIVE_FORMAT_IMPORT
    ) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn validate_part(path: &str, content_type: &str) -> Result<()> {
    if path
        .split(['/', '\\'])
        .find(|segment| !segment.is_empty() && *segment != ".")
        .is_some_and(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
        || matches!(
            content_type,
            ct::OPC_DIGITAL_SIGNATURE_ORIGIN
                | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
                | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
        )
    {
        return Err(Error::Refused(Refusal::SignatureInfrastructure));
    }
    if matches!(
        content_type,
        ct::WML_DOCUMENT_MACRO_MAIN
            | ct::WML_TEMPLATE_MACRO_MAIN
            | ct::OFC_VBA_PROJECT
            | ct::OFC_VBA_PROJECT_SIGNATURE
            | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
            | ct::WML_VBA_DATA
    ) {
        return Err(Error::Refused(Refusal::MacroEnabled));
    }
    if matches!(
        content_type,
        ct::WML_COMMENTS | ct::WML_FOOTNOTES | ct::WML_ENDNOTES | ct::OFC_CUSTOM_XML_PROPERTIES
    ) || path.starts_with("/customXml/")
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn is_signature_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_signature_physical_member(path: &str) -> bool {
    path.split(['/', '\\'])
        .find(|segment| !segment.is_empty() && *segment != ".")
        .is_some_and(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
}

fn map_scan_error(error: ScanError) -> Error {
    match error {
        ScanError::Io(error) => Error::Io(error),
        ScanError::Parser => Error::Scan("XML parser rejected the bounded event".into()),
        ScanError::Semantic(message) => Error::Scan(message.into()),
        ScanError::Limit {
            resource,
            actual,
            maximum,
        } => Error::Limit {
            resource,
            actual,
            maximum,
        },
        ScanError::Allocation { resource, source } => Error::Allocation { resource, source },
        ScanError::Execution(error) => Error::Execution(error),
    }
}

fn inspect_root_dialect(
    bytes: &[u8],
    root: &[u8],
    expected_strict: bool,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<bool> {
    let expected_len = u64::try_from(bytes.len()).map_err(|_| Error::Limit {
        resource: "settings XML bytes",
        actual: u64::MAX,
        maximum: limits.max_settings_xml_bytes,
    })?;
    let max_token = usize::try_from(limits.max_token_bytes)
        .map_err(|_| Error::Limit {
            resource: "XML token bytes",
            actual: limits.max_token_bytes,
            maximum: usize::MAX as u64,
        })?
        .min(bytes.len().max(1));
    let workspace = scanner_workspace_requirement(max_token, 0).ok_or(Error::Limit {
        resource: "parser workspace",
        actual: u64::MAX,
        maximum: limits.max_workspace_bytes,
    })?;
    check_limit("parser workspace", workspace, limits.max_workspace_bytes)?;
    let _reservation = context
        .map(|execution| {
            execution
                .reserve(Resource::Memory, workspace)
                .map(Arc::new)
                .map_err(Error::Execution)
        })
        .transpose()?;
    check_options(&Options {
        execution: context.cloned(),
        cancellation: cancellation.cloned(),
    })?;
    let guarded =
        GuardedBufRead::new(Cursor::new(bytes), expected_len, max_token).map_err(map_scan_error)?;
    let mut parser = NsReader::from_reader(guarded);
    parser
        .resolver_mut()
        .set_max_declarations_per_element(bytes.len().clamp(1, 256));
    parser.config_mut().trim_text(false);
    parser.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(max_token.saturating_add(1))
        .map_err(|source| Error::Allocation {
            resource: "settings XML parser buffer",
            source,
        })?;
    let mut saw_declaration = false;
    loop {
        parser.get_mut().begin_token();
        let (namespace, event) = parser
            .read_resolved_event_into(&mut buffer)
            .map_err(map_quick_xml_error)
            .map_err(map_scan_error)?;
        validate_event_name(&event).map_err(map_scan_error)?;
        let dialect = namespace_dialect(&namespace);
        check_progress(context, cancellation, 1).map_err(map_scan_error)?;
        match event {
            Event::Start(start) | Event::Empty(start) if start.local_name().as_ref() == root => {
                let strict = strict_from_dialect(dialect).map_err(map_scan_error)?;
                if strict != expected_strict {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                return Ok(strict);
            },
            Event::Decl(declaration) => {
                if saw_declaration {
                    return Err(Error::Scan("duplicate XML declaration".into()));
                }
                validate_xml_declaration(&declaration).map_err(map_scan_error)?;
                saw_declaration = true;
            },
            Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
            Event::Eof => return Err(Error::Refused(Refusal::UnsupportedDependency)),
            _ => return Err(Error::Refused(Refusal::UnsupportedDependency)),
        }
        buffer.clear();
    }
}

fn settings_lossy_len(bytes: &[u8]) -> usize {
    bytes.iter().fold(0usize, |total, byte| {
        total.saturating_add(if byte.is_ascii() { 1 } else { 3 })
    })
}

fn settings_namespace_len(namespace: ResolveResult<'_>) -> usize {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => settings_lossy_len(value),
        ResolveResult::Unknown(prefix) => settings_lossy_len(&prefix),
        ResolveResult::Unbound => 0,
    }
}

fn settings_namespace_declarations(element: &BytesStart<'_>) -> Result<(usize, usize)> {
    let mut bytes = 0usize;
    let mut count = 0usize;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|_| Error::Scan("settings XML attribute is invalid".into()))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            &[][..]
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            prefix
        } else {
            continue;
        };
        let prefix_bytes = prefix.len();
        bytes = bytes
            .checked_add(prefix_bytes)
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or(Error::Limit {
                resource: "settings XML namespace bytes",
                actual: u64::MAX,
                maximum: u64::MAX - 1,
            })?;
        count = count.checked_add(1).ok_or(Error::Limit {
            resource: "settings XML namespace bindings",
            actual: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    }
    Ok((bytes, count))
}

fn settings_namespace_is_mce(namespace: ResolveResult<'_>) -> Result<bool> {
    let ResolveResult::Bound(Namespace(value)) = namespace else {
        return Ok(false);
    };
    if value == MCE_NAMESPACE.as_bytes() {
        return Ok(true);
    }
    let value = std::str::from_utf8(value)
        .map_err(|_| Error::Scan("settings XML namespace URI is not UTF-8".into()))?;
    let value = quick_xml::escape::unescape(value)
        .map_err(|_| Error::Scan("settings XML namespace URI escape is invalid".into()))?;
    Ok(value.as_bytes() == MCE_NAMESPACE.as_bytes())
}

fn settings_event_owned_bytes(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<(usize, usize)> {
    let namespace = resolver.resolve_element(element.name()).0;
    let mut owned = settings_namespace_len(namespace)
        .checked_add(settings_lossy_len(element.local_name().as_ref()))
        .ok_or(Error::Limit {
            resource: "settings XML semantic bytes",
            actual: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    let mut attributes = 0usize;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|_| Error::Scan("settings XML attribute is invalid".into()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (attribute_namespace, _) = resolver.resolve_attribute(attribute.key);
        owned = owned
            .checked_add(settings_namespace_len(attribute_namespace))
            .and_then(|value| {
                value.checked_add(settings_lossy_len(attribute.key.local_name().as_ref()))
            })
            .and_then(|value| value.checked_add(attribute.value.len()))
            .ok_or(Error::Limit {
                resource: "settings XML semantic bytes",
                actual: u64::MAX,
                maximum: u64::MAX - 1,
            })?;
        attributes = attributes.checked_add(1).ok_or(Error::Limit {
            resource: "settings XML attributes",
            actual: u64::MAX,
            maximum: u64::MAX - 1,
        })?;
    }
    Ok((owned, attributes))
}

fn settings_mce_directive_facts(
    element: &BytesStart<'_>,
    resolver: &quick_xml::name::NamespaceResolver,
    decoder: Decoder,
) -> Result<(usize, usize)> {
    let mut tokens = 0usize;
    let mut owned = 0usize;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|_| Error::Scan("settings XML attribute is invalid".into()))?;
        let (namespace, _) = resolver.resolve_attribute(attribute.key);
        if !settings_namespace_is_mce(namespace)? {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|_| Error::Scan("settings XML MCE directive value is invalid".into()))?;
        for token in value
            .as_ref()
            .as_bytes()
            .split(|byte| byte.is_ascii_whitespace())
        {
            if token.is_empty() {
                continue;
            }
            tokens = tokens.checked_add(1).ok_or(Error::Limit {
                resource: "settings XML MCE directive tokens",
                actual: u64::MAX,
                maximum: u64::MAX - 1,
            })?;
            owned = owned
                .checked_add(settings_lossy_len(token))
                .ok_or(Error::Limit {
                    resource: "settings XML MCE directive bytes",
                    actual: u64::MAX,
                    maximum: u64::MAX - 1,
                })?;
            let separator = token.iter().position(|byte| *byte == b':');
            let prefix = separator.map_or(token, |separator| &token[..separator]);
            let local = separator.and_then(|separator| token.get(separator + 1..));
            let matching_namespace =
                resolver
                    .bindings()
                    .find_map(|(candidate, namespace)| match candidate {
                        PrefixDeclaration::Named(value) if value == prefix => Some(namespace),
                        PrefixDeclaration::Default | PrefixDeclaration::Named(_) => None,
                    });
            if let Some(namespace) = matching_namespace {
                owned = owned
                    .checked_add(settings_lossy_len(namespace.as_ref()))
                    .and_then(|value| {
                        local.map_or(Some(value), |local| {
                            value.checked_add(settings_lossy_len(local))
                        })
                    })
                    .ok_or(Error::Limit {
                        resource: "settings XML MCE directive bytes",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
            }
        }
    }
    Ok((tokens, owned))
}

/// Consume one settings source pass through the guarded reader before MCE or
/// model extraction.  The complete settings parser owns its model semantics;
/// this pass only enforces the caller's event, depth, and pre-read token
/// ceilings so quick-xml cannot materialize an unbounded start tag first.  It
/// also returns source-derived owner facts for the later MCE/model lease.
fn guard_settings_xml(
    bytes: &[u8],
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<SettingsGuardFacts> {
    let expected_len = u64::try_from(bytes.len()).map_err(|_| Error::Limit {
        resource: "settings XML bytes",
        actual: u64::MAX,
        maximum: limits.max_settings_xml_bytes,
    })?;
    let max_token = usize::try_from(limits.max_token_bytes)
        .map_err(|_| Error::Limit {
            resource: "XML token bytes",
            actual: limits.max_token_bytes,
            maximum: usize::MAX as u64,
        })?
        .min(bytes.len().max(1));
    let max_depth = usize::try_from(limits.max_depth)
        .map_err(|_| Error::Limit {
            resource: "XML depth",
            actual: limits.max_depth,
            maximum: usize::MAX as u64,
        })?
        .min(bytes.len().checked_div(3).unwrap_or(0).max(1));
    let workspace = scanner_workspace_requirement(max_token, max_depth).ok_or(Error::Limit {
        resource: "parser workspace",
        actual: u64::MAX,
        maximum: limits.max_workspace_bytes,
    })?;
    check_limit("parser workspace", workspace, limits.max_workspace_bytes)?;
    let _reservation = context
        .map(|execution| {
            execution
                .reserve(Resource::Memory, workspace)
                .map(Arc::new)
                .map_err(Error::Execution)
        })
        .transpose()?;
    check_options(&Options {
        execution: context.cloned(),
        cancellation: cancellation.cloned(),
    })?;
    let guarded =
        GuardedBufRead::new(Cursor::new(bytes), expected_len, max_token).map_err(map_scan_error)?;
    let max_namespace_declarations = bytes.len().clamp(1, 256);
    let mut reader = NsReader::from_reader(guarded);
    reader
        .resolver_mut()
        .set_max_declarations_per_element(max_namespace_declarations);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let token_window = max_token.checked_add(1).ok_or(Error::Limit {
        resource: "XML token bytes",
        actual: u64::MAX,
        maximum: limits.max_token_bytes,
    })?;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(token_window)
        .map_err(|source| Error::Allocation {
            resource: "settings XML parser buffer",
            source,
        })?;
    let mut depth = 0usize;
    let mut events = 0u64;
    let mut saw_declaration = false;
    let mut saw_root = false;
    let mut facts = SettingsGuardFacts::default();
    let mut namespace_buffer = INITIAL_QUICK_XML_NAMESPACE_BYTES;
    let mut namespace_bindings = 2usize;
    let mut namespace_scopes = Vec::<(usize, usize)>::new();
    let namespace_scope_capacity = max_depth.checked_add(1).ok_or(Error::Limit {
        resource: "XML depth",
        actual: u64::MAX,
        maximum: max_depth as u64,
    })?;
    namespace_scopes
        .try_reserve_exact(namespace_scope_capacity)
        .map_err(|source| Error::Allocation {
            resource: "settings XML namespace scope stack",
            source,
        })?;
    loop {
        reader.get_mut().begin_token();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(map_quick_xml_error)
            .map_err(map_scan_error)?;
        events = events.checked_add(1).ok_or(Error::Limit {
            resource: "XML events",
            actual: u64::MAX,
            maximum: limits.max_events,
        })?;
        if events > limits.max_events {
            return Err(Error::Limit {
                resource: "XML events",
                actual: events,
                maximum: limits.max_events,
            });
        }
        facts.event_count = events;
        check_progress(context, cancellation, 1).map_err(map_scan_error)?;
        let raw_len = reader.get_mut().captured().len();
        facts.max_token = facts.max_token.max(raw_len);
        let decoder = reader.decoder();
        let current_bindings = namespace_bindings;
        facts.max_namespace_buffer = facts.max_namespace_buffer.max(namespace_buffer);
        facts.max_namespace_bindings = facts.max_namespace_bindings.max(current_bindings);
        validate_event_name(&event).map_err(map_scan_error)?;
        match &event {
            Event::Decl(declaration) => {
                if saw_declaration || saw_root {
                    return Err(Error::Scan(
                        "settings XML declaration is outside its prolog".into(),
                    ));
                }
                validate_xml_declaration(declaration).map_err(map_scan_error)?;
                saw_declaration = true;
            },
            Event::Start(_) | Event::Empty(_) if depth == 0 => {
                saw_root = true;
            },
            _ => {},
        }
        match event {
            Event::Start(start) => {
                let (declared_bytes, declared_bindings) = settings_namespace_declarations(&start)?;
                namespace_buffer =
                    namespace_buffer
                        .checked_add(declared_bytes)
                        .ok_or(Error::Limit {
                            resource: "settings XML namespace bytes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                namespace_bindings =
                    namespace_bindings
                        .checked_add(declared_bindings)
                        .ok_or(Error::Limit {
                            resource: "settings XML namespace bindings",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.namespace_declaration_count = facts
                    .namespace_declaration_count
                    .checked_add(declared_bindings)
                    .ok_or(Error::Limit {
                        resource: "settings XML namespace bindings",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                facts.max_namespace_buffer = facts.max_namespace_buffer.max(namespace_buffer);
                facts.max_namespace_bindings = facts.max_namespace_bindings.max(namespace_bindings);
                facts.namespace_copy_bytes = facts
                    .namespace_copy_bytes
                    .checked_add(namespace_buffer)
                    .ok_or(Error::Limit {
                        resource: "settings XML namespace copy bytes",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                namespace_scopes.push((declared_bytes, declared_bindings));
                depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "XML depth",
                    actual: u64::MAX,
                    maximum: limits.max_depth,
                })?;
                if depth > max_depth {
                    return Err(Error::Limit {
                        resource: "XML depth",
                        actual: depth as u64,
                        maximum: max_depth as u64,
                    });
                }
                facts.max_depth = facts.max_depth.max(depth);
                facts.node_count = facts.node_count.checked_add(1).ok_or(Error::Limit {
                    resource: "settings XML nodes",
                    actual: u64::MAX,
                    maximum: u64::MAX - 1,
                })?;
                let (owned, attributes) = settings_event_owned_bytes(&start, reader.resolver())?;
                facts.semantic_owned_bytes =
                    facts
                        .semantic_owned_bytes
                        .checked_add(owned)
                        .ok_or(Error::Limit {
                            resource: "settings XML semantic bytes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.attribute_count =
                    facts
                        .attribute_count
                        .checked_add(attributes)
                        .ok_or(Error::Limit {
                            resource: "settings XML attributes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.max_attributes_per_event = facts.max_attributes_per_event.max(attributes);
                if attributes != 0 {
                    facts.attribute_bearing_nodes = facts
                        .attribute_bearing_nodes
                        .checked_add(1)
                        .ok_or(Error::Limit {
                            resource: "settings XML attribute-bearing nodes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                }
                let (directive_tokens, directive_bytes) =
                    settings_mce_directive_facts(&start, reader.resolver(), decoder)?;
                facts.mce_directive_tokens = facts
                    .mce_directive_tokens
                    .checked_add(directive_tokens)
                    .ok_or(Error::Limit {
                        resource: "settings XML MCE directive tokens",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                facts.mce_directive_owned_bytes = facts
                    .mce_directive_owned_bytes
                    .checked_add(directive_bytes)
                    .ok_or(Error::Limit {
                        resource: "settings XML MCE directive bytes",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
            },
            Event::Empty(start) => {
                let child_depth = depth.checked_add(1).ok_or(Error::Limit {
                    resource: "XML depth",
                    actual: u64::MAX,
                    maximum: limits.max_depth,
                })?;
                if child_depth > max_depth {
                    return Err(Error::Limit {
                        resource: "XML depth",
                        actual: child_depth as u64,
                        maximum: max_depth as u64,
                    });
                }
                facts.max_depth = facts.max_depth.max(child_depth);
                let (declared_bytes, declared_bindings) = settings_namespace_declarations(&start)?;
                namespace_buffer =
                    namespace_buffer
                        .checked_add(declared_bytes)
                        .ok_or(Error::Limit {
                            resource: "settings XML namespace bytes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.namespace_declaration_count = facts
                    .namespace_declaration_count
                    .checked_add(declared_bindings)
                    .ok_or(Error::Limit {
                        resource: "settings XML namespace bindings",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                namespace_bindings =
                    namespace_bindings
                        .checked_add(declared_bindings)
                        .ok_or(Error::Limit {
                            resource: "settings XML namespace bindings",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.max_namespace_buffer = facts.max_namespace_buffer.max(namespace_buffer);
                facts.max_namespace_bindings = facts.max_namespace_bindings.max(namespace_bindings);
                facts.namespace_copy_bytes = facts
                    .namespace_copy_bytes
                    .checked_add(namespace_buffer)
                    .ok_or(Error::Limit {
                        resource: "settings XML namespace copy bytes",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                facts.node_count = facts.node_count.checked_add(1).ok_or(Error::Limit {
                    resource: "settings XML nodes",
                    actual: u64::MAX,
                    maximum: u64::MAX - 1,
                })?;
                let (owned, attributes) = settings_event_owned_bytes(&start, reader.resolver())?;
                facts.semantic_owned_bytes =
                    facts
                        .semantic_owned_bytes
                        .checked_add(owned)
                        .ok_or(Error::Limit {
                            resource: "settings XML semantic bytes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.attribute_count =
                    facts
                        .attribute_count
                        .checked_add(attributes)
                        .ok_or(Error::Limit {
                            resource: "settings XML attributes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                facts.max_attributes_per_event = facts.max_attributes_per_event.max(attributes);
                if attributes != 0 {
                    facts.attribute_bearing_nodes = facts
                        .attribute_bearing_nodes
                        .checked_add(1)
                        .ok_or(Error::Limit {
                            resource: "settings XML attribute-bearing nodes",
                            actual: u64::MAX,
                            maximum: u64::MAX - 1,
                        })?;
                }
                let (directive_tokens, directive_bytes) =
                    settings_mce_directive_facts(&start, reader.resolver(), decoder)?;
                facts.mce_directive_tokens = facts
                    .mce_directive_tokens
                    .checked_add(directive_tokens)
                    .ok_or(Error::Limit {
                        resource: "settings XML MCE directive tokens",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                facts.mce_directive_owned_bytes = facts
                    .mce_directive_owned_bytes
                    .checked_add(directive_bytes)
                    .ok_or(Error::Limit {
                        resource: "settings XML MCE directive bytes",
                        actual: u64::MAX,
                        maximum: u64::MAX - 1,
                    })?;
                namespace_buffer = namespace_buffer.saturating_sub(declared_bytes);
                namespace_bindings = namespace_bindings.saturating_sub(declared_bindings);
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or(Error::Scan("settings XML has an unexpected end tag".into()))?;
                let (declared_bytes, declared_bindings) = namespace_scopes
                    .pop()
                    .ok_or(Error::Scan("settings XML has an unexpected end tag".into()))?;
                namespace_buffer = namespace_buffer.saturating_sub(declared_bytes);
                namespace_bindings = namespace_bindings.saturating_sub(declared_bindings);
            },
            Event::Eof => break,
            Event::Text(_)
            | Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
        buffer.clear();
    }
    if depth != 0 {
        return Err(Error::Scan("settings XML nesting is incomplete".into()));
    }
    if reader.get_mut().position() != expected_len {
        return Err(Error::Scan(
            "settings XML guard did not consume its source bound".into(),
        ));
    }
    Ok(facts)
}

fn encode_fragment<'package>(
    package: &'package litchi_opc::SourceBackedPackage,
    text: &str,
    strict: bool,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
) -> Result<SourcePartSpliceFragment<'package>> {
    let input_bytes = u64::try_from(text.len()).map_err(|_| Error::Limit {
        resource: "text bytes",
        actual: u64::MAX,
        maximum: limits.max_text_bytes,
    })?;
    check_limit("text bytes", input_bytes, limits.max_text_bytes)?;
    check_progress(context, cancellation, 0).map_err(map_scan_error)?;
    let namespace = if strict {
        STRICT_WORDPROCESSINGML_NAMESPACE
    } else {
        WORDPROCESSINGML_NAMESPACE
    };
    let prefix_a = b"<w:p xmlns:w=\"";
    let prefix_b = b"\"><w:r><w:t xml:space=\"preserve\">";
    let suffix = b"</w:t></w:r></w:p>";

    let mut encoded_text = 0_u64;
    let mut pending_work = 0_u64;
    for character in text.chars() {
        if !is_plain_text_character(character) {
            return Err(Error::Refused(Refusal::ComplexParagraph));
        }
        let bytes = u64::try_from(escaped_character_len(character)).map_err(|_| Error::Limit {
            resource: "fragment bytes",
            actual: u64::MAX,
            maximum: limits.max_fragment_bytes,
        })?;
        encoded_text = encoded_text.checked_add(bytes).ok_or(Error::Limit {
            resource: "fragment bytes",
            actual: u64::MAX,
            maximum: limits.max_fragment_bytes,
        })?;
        pending_work = pending_work.checked_add(1).ok_or(Error::Limit {
            resource: "work",
            actual: u64::MAX,
            maximum: limits.max_events,
        })?;
        if pending_work == 64 {
            check_progress(context, cancellation, pending_work).map_err(map_scan_error)?;
            pending_work = 0;
        }
    }
    if pending_work != 0 {
        check_progress(context, cancellation, pending_work).map_err(map_scan_error)?;
    }
    let prefix_len = u64::try_from(prefix_a.len())
        .unwrap_or(u64::MAX)
        .checked_add(u64::try_from(namespace.len()).unwrap_or(u64::MAX))
        .and_then(|value| value.checked_add(u64::try_from(prefix_b.len()).unwrap_or(u64::MAX)))
        .ok_or(Error::Limit {
            resource: "fragment bytes",
            actual: u64::MAX,
            maximum: limits.max_fragment_bytes,
        })?;
    let suffix_len = u64::try_from(suffix.len()).unwrap_or(u64::MAX);
    let fragment_len = prefix_len
        .checked_add(encoded_text)
        .and_then(|value| value.checked_add(suffix_len))
        .ok_or(Error::Limit {
            resource: "fragment bytes",
            actual: u64::MAX,
            maximum: limits.max_fragment_bytes,
        })?;
    check_limit("fragment bytes", fragment_len, limits.max_fragment_bytes)?;
    let capacity = usize::try_from(fragment_len).map_err(|_| Error::Limit {
        resource: "fragment bytes",
        actual: fragment_len,
        maximum: limits.max_fragment_bytes,
    })?;
    // The package allocates and reserves the retained fragment as one fixed
    // owner.  The reservation is transferred into the OPC plan, so this
    // encoder does not create a second temporary lease for the same bytes.
    let mut fragment = package
        .allocate_source_part_splice_fragment(fragment_len, limits.max_fragment_bytes)
        .map_err(|error| Error::Document(DocumentError::from(error)))?;
    let output = fragment.as_mut_slice();
    let mut position = 0_usize;
    output[position..position + prefix_a.len()].copy_from_slice(prefix_a);
    position += prefix_a.len();
    output[position..position + namespace.len()].copy_from_slice(namespace);
    position += namespace.len();
    output[position..position + prefix_b.len()].copy_from_slice(prefix_b);
    position += prefix_b.len();
    let mut scratch = [0_u8; 5];
    let mut pending_work = 0_u64;
    for character in text.chars() {
        let needed = escaped_character_len(character);
        let written = append_character(character, &mut scratch[..needed]);
        if written != needed {
            return Err(Error::Scan("plain text encoder scratch overflow".into()));
        }
        output[position..position + written].copy_from_slice(&scratch[..written]);
        position += written;
        pending_work = pending_work.saturating_add(1);
        if pending_work == 64 {
            check_progress(context, cancellation, pending_work).map_err(map_scan_error)?;
            pending_work = 0;
        }
    }
    if pending_work != 0 {
        check_progress(context, cancellation, pending_work).map_err(map_scan_error)?;
    }
    output[position..position + suffix.len()].copy_from_slice(suffix);
    position += suffix.len();
    if position != capacity {
        return Err(Error::Scan(
            "plain paragraph fragment length changed".into(),
        ));
    }
    Ok(fragment)
}

#[derive(Debug, Error)]
enum ScanError {
    #[error("bounded XML reader failed: {0}")]
    Io(#[from] io::Error),
    #[error("bounded XML parser rejected an event")]
    Parser,
    #[error("{0}")]
    Semantic(&'static str),
    #[error("{resource} limit exceeded: {actual} > {maximum}")]
    Limit {
        resource: &'static str,
        actual: u64,
        maximum: u64,
    },
    #[error("allocation failed for {resource}: {source}")]
    Allocation {
        resource: &'static str,
        #[source]
        source: TryReserveError,
    },
    #[error("execution failed: {0}")]
    Execution(#[source] ExecutionError),
}

fn map_quick_xml_error(error: quick_xml::Error) -> ScanError {
    match error {
        quick_xml::Error::Io(source) => {
            ScanError::Io(io::Error::new(source.kind(), SharedIoSource(source)))
        },
        _ => ScanError::Parser,
    }
}

/// A source/candidate reader which exposes the exact prefix, generated
/// fragment, and suffix without retaining either XML member.
struct TailSpliceReader<'a> {
    source: &'a mut dyn BufRead,
    source_len: u64,
    insertion: u64,
    source_position: u64,
    fragment: &'a [u8],
    fragment_position: usize,
}

/// Preserve the original quick-xml I/O error as the source of the adapter's
/// `io::Error`. quick-xml stores that error behind an `Arc` so its parser error
/// remains cloneable; copying only its display text would discard the original
/// kind and source chain.
#[derive(Debug)]
struct SharedIoSource(Arc<io::Error>);

impl std::fmt::Display for SharedIoSource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        self.0.fmt(formatter)
    }
}

impl std::error::Error for SharedIoSource {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        Some(self.0.as_ref())
    }
}

impl<'a> TailSpliceReader<'a> {
    fn new(
        source: &'a mut dyn BufRead,
        source_len: u64,
        insertion: u64,
        fragment: &'a [u8],
    ) -> Self {
        Self {
            source,
            source_len,
            insertion,
            source_position: 0,
            fragment,
            fragment_position: 0,
        }
    }
}

impl Read for TailSpliceReader<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl BufRead for TailSpliceReader<'_> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.source_position < self.insertion {
            let available = self.source.fill_buf()?;
            let remaining =
                usize::try_from(self.insertion - self.source_position).unwrap_or(usize::MAX);
            return Ok(&available[..available.len().min(remaining)]);
        }
        if self.fragment_position < self.fragment.len() {
            return Ok(&self.fragment[self.fragment_position..]);
        }
        if self.source_position >= self.source_len {
            return Ok(&[]);
        }
        let available = self.source.fill_buf()?;
        let remaining =
            usize::try_from(self.source_len - self.source_position).unwrap_or(usize::MAX);
        Ok(&available[..available.len().min(remaining)])
    }

    fn consume(&mut self, amount: usize) {
        if self.source_position < self.insertion {
            let remaining =
                usize::try_from(self.insertion - self.source_position).unwrap_or(usize::MAX);
            let count = amount.min(remaining);
            if count != 0 {
                self.source.consume(count);
                self.source_position = self.source_position.saturating_add(count as u64);
            }
            return;
        }
        if self.fragment_position < self.fragment.len() {
            let count = amount.min(self.fragment.len() - self.fragment_position);
            self.fragment_position += count;
            return;
        }
        let remaining = usize::try_from(self.source_len.saturating_sub(self.source_position))
            .unwrap_or(usize::MAX);
        let count = amount.min(remaining);
        if count != 0 {
            self.source.consume(count);
            self.source_position = self.source_position.saturating_add(count as u64);
        }
    }
}

/// A guarded `BufRead` copied in spirit from the OPC XML audit adapter.  The
/// prefix probe handles a split UTF-8 BOM; each parser event receives at most
/// `max_token + 1` bytes, so quick-xml cannot grow its event buffer in
/// proportion to an unbounded hostile token before the semantic limit fires.
struct GuardedBufRead<R> {
    inner: R,
    prefix: [u8; 3],
    prefix_len: usize,
    prefix_pos: usize,
    saw_bom: bool,
    total: u64,
    token: usize,
    max_total: u64,
    max_token: usize,
    max_token_window: usize,
    captured: Vec<u8>,
    exposed: Vec<u8>,
    exposed_pos: usize,
    exposed_prefix: bool,
}

impl<R: BufRead> GuardedBufRead<R> {
    fn new(inner: R, max_total: u64, max_token: usize) -> std::result::Result<Self, ScanError> {
        let max_token_window = max_token.checked_add(1).ok_or(ScanError::Limit {
            resource: "XML token bytes",
            actual: u64::MAX,
            maximum: max_token as u64,
        })?;
        let mut guarded = Self {
            inner,
            prefix: [0; 3],
            prefix_len: 0,
            prefix_pos: 0,
            saw_bom: false,
            total: 0,
            token: 0,
            max_total,
            max_token,
            max_token_window,
            captured: Vec::new(),
            exposed: Vec::new(),
            exposed_pos: 0,
            exposed_prefix: false,
        };
        guarded
            .captured
            .try_reserve_exact(max_token_window)
            .map_err(|source| ScanError::Allocation {
                resource: "XML token capture",
                source,
            })?;
        guarded
            .exposed
            .try_reserve_exact(max_token_window)
            .map_err(|source| ScanError::Allocation {
                resource: "XML token window",
                source,
            })?;
        let prefix_limit = usize::try_from(max_total.min(3)).unwrap_or(3);
        while guarded.prefix_len < prefix_limit {
            let available = guarded.inner.fill_buf()?;
            if available.is_empty() {
                break;
            }
            let count = available
                .len()
                .min(prefix_limit.saturating_sub(guarded.prefix_len));
            guarded.prefix[guarded.prefix_len..guarded.prefix_len + count]
                .copy_from_slice(&available[..count]);
            guarded.inner.consume(count);
            guarded.prefix_len += count;
        }
        if guarded.prefix_len == 3 && guarded.prefix == [0xEF, 0xBB, 0xBF] {
            guarded.saw_bom = true;
            guarded.prefix_pos = guarded.prefix_len;
            guarded.total = 3;
        }
        Ok(guarded)
    }

    fn begin_token(&mut self) {
        self.token = 0;
        self.captured.clear();
    }

    const fn position(&self) -> u64 {
        self.total
    }

    const fn saw_bom(&self) -> bool {
        self.saw_bom
    }

    fn captured(&self) -> &[u8] {
        &self.captured
    }
}

impl<R: BufRead> Read for GuardedBufRead<R> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let available = self.fill_buf()?;
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

impl<R: BufRead> BufRead for GuardedBufRead<R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.exposed_pos < self.exposed.len() {
            return Ok(&self.exposed[self.exposed_pos..]);
        }
        self.exposed.clear();
        self.exposed_pos = 0;
        if self.token > self.max_token {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "XML token byte limit exceeded",
            ));
        }
        if self.total >= self.max_total {
            if self.prefix_pos < self.prefix_len {
                return Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "XML source byte limit exceeded",
                ));
            }
            let available = self.inner.fill_buf()?;
            if available.is_empty() {
                return Ok(available);
            }
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "XML source byte limit exceeded",
            ));
        }
        let total_remaining = usize::try_from(self.max_total - self.total).unwrap_or(usize::MAX);
        let token_remaining = self.max_token_window.saturating_sub(self.token);
        if token_remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "XML token byte limit exceeded",
            ));
        }
        let pending = self.prefix_pos < self.prefix_len;
        if pending {
            let available = &self.prefix[self.prefix_pos..self.prefix_len];
            let visible = available.len().min(total_remaining).min(token_remaining);
            self.exposed.extend_from_slice(&available[..visible]);
        } else {
            let mut exposed = std::mem::take(&mut self.exposed);
            let available = self.inner.fill_buf()?;
            if available.is_empty() {
                self.exposed = exposed;
                return Ok(&self.exposed);
            }
            let visible = available.len().min(total_remaining).min(token_remaining);
            exposed.extend_from_slice(&available[..visible]);
            self.exposed = exposed;
        }
        self.exposed_prefix = pending;
        Ok(&self.exposed)
    }

    fn consume(&mut self, amount: usize) {
        let available = self.exposed.len().saturating_sub(self.exposed_pos);
        let amount = amount
            .min(available)
            .min(self.max_token_window.saturating_sub(self.token));
        if amount == 0 {
            return;
        }
        self.captured
            .extend_from_slice(&self.exposed[self.exposed_pos..self.exposed_pos + amount]);
        if self.exposed_prefix {
            self.prefix_pos = self.prefix_pos.saturating_add(amount);
        } else {
            self.inner.consume(amount);
        }
        self.exposed_pos += amount;
        self.token = self.token.saturating_add(amount);
        self.total = self.total.saturating_add(amount as u64);
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FrameKind {
    Document,
    Body,
    Paragraph,
    Run,
    Text,
}

impl FrameKind {
    const fn local(self) -> &'static [u8] {
        match self {
            Self::Document => b"document",
            Self::Body => b"body",
            Self::Paragraph => b"p",
            Self::Run => b"r",
            Self::Text => b"t",
        }
    }
}

#[derive(Clone, Copy)]
pub(super) struct ScanFacts {
    pub(super) len: u64,
    pub(super) sha256: [u8; 32],
    pub(super) insertion_offset: u64,
    pub(super) paragraph_count: u64,
    pub(super) event_count: u64,
    pub(super) max_depth: u64,
    pub(super) strict_namespace: bool,
    pub(super) sect_pr_len: u64,
    pub(super) sect_pr_sha256: [u8; 32],
    pub(super) generated_offset: u64,
    pub(super) generated_once: bool,
    pub(super) generated_count: u64,
}

impl ScanFacts {
    pub(super) fn source_proof(self, source_version: SourceVersion) -> SourceProof {
        SourceProof {
            source_version,
            source_len: self.len,
            source_sha256: self.sha256,
            insertion_offset: self.insertion_offset,
            paragraph_count: self.paragraph_count,
            event_count: self.event_count,
            max_depth: self.max_depth,
            strict_namespace: self.strict_namespace,
            sect_pr_len: self.sect_pr_len,
            sect_pr_sha256: self.sect_pr_sha256,
        }
    }

    pub(super) fn candidate_proof(self) -> CandidateProof {
        CandidateProof {
            candidate_len: self.len,
            candidate_sha256: self.sha256,
            paragraph_count: self.paragraph_count,
            event_count: self.event_count,
            max_depth: self.max_depth,
            generated_offset: self.generated_offset,
            generated_once: self.generated_once,
            sect_pr_len: self.sect_pr_len,
            sect_pr_sha256: self.sect_pr_sha256,
        }
    }
}

pub(super) fn scan_main_part(
    part: &litchi_opc::PartView<'_>,
    source_len: u64,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    candidate: Option<(&ScanFacts, &[u8])>,
) -> Result<ScanFacts> {
    let expected_len = if let Some((source, fragment)) = candidate {
        source
            .len
            .checked_add(u64::try_from(fragment.len()).map_err(|_| Error::Limit {
                resource: "candidate XML bytes",
                actual: u64::MAX,
                maximum: limits.max_candidate_xml_bytes,
            })?)
            .ok_or(Error::Limit {
                resource: "candidate XML bytes",
                actual: u64::MAX,
                maximum: limits.max_candidate_xml_bytes,
            })?
    } else {
        source_len
    };
    let scanned = if let Some((source, fragment)) = candidate {
        part.with_verified_decoded_reader(|reader| {
            let mut splice =
                TailSpliceReader::new(reader, source.len, source.insertion_offset, fragment);
            scan_reader(
                &mut splice,
                expected_len,
                limits,
                context,
                cancellation,
                Some((source.insertion_offset, fragment.len() as u64, 1)),
            )
        })
    } else {
        part.with_verified_decoded_reader(|mut reader| {
            scan_reader(
                &mut reader,
                expected_len,
                limits,
                context,
                cancellation,
                None,
            )
        })
    };
    match scanned {
        Ok(facts) => Ok(facts),
        Err(VerifiedDecodedReaderError::Callback(error)) => Err(map_scan_error(error)),
        Err(VerifiedDecodedReaderError::Opc {
            error,
            callback_error: _,
        }) => Err(Error::Document(DocumentError::from(error))),
        _ => Err(Error::Scan(
            "verified OPC reader returned an unknown failure".into(),
        )),
    }
}

fn scan_reader<R: BufRead>(
    reader: &mut R,
    expected_len: u64,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    generated: Option<(u64, u64, u64)>,
) -> std::result::Result<ScanFacts, ScanError> {
    let max_token = usize::try_from(limits.max_token_bytes).map_err(|_| ScanError::Limit {
        resource: "XML token bytes",
        actual: limits.max_token_bytes,
        maximum: usize::MAX as u64,
    })?;
    let max_depth = usize::try_from(limits.max_depth).map_err(|_| ScanError::Limit {
        resource: "XML depth",
        actual: limits.max_depth,
        maximum: usize::MAX as u64,
    })?;
    let token_window = max_token.checked_add(1).ok_or(ScanError::Limit {
        resource: "XML token bytes",
        actual: u64::MAX,
        maximum: limits.max_token_bytes,
    })?;
    let workspace =
        scanner_workspace_requirement(max_token, max_depth).ok_or(ScanError::Limit {
            resource: "parser workspace",
            actual: u64::MAX,
            maximum: limits.max_workspace_bytes,
        })?;
    if workspace > limits.max_workspace_bytes {
        return Err(ScanError::Limit {
            resource: "parser workspace",
            actual: workspace,
            maximum: limits.max_workspace_bytes,
        });
    }
    let _workspace_reservation = context
        .map(|execution| {
            execution
                .reserve(Resource::Memory, workspace)
                .map(Arc::new)
                .map_err(ScanError::Execution)
        })
        .transpose()?;
    check_progress(context, cancellation, 0)?;
    let guarded = GuardedBufRead::new(reader, expected_len, max_token)?;
    let saw_bom = guarded.saw_bom();
    let mut parser = NsReader::from_reader(guarded);
    parser.config_mut().trim_text(false);
    parser.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    buffer
        .try_reserve_exact(token_window)
        .map_err(|source| ScanError::Allocation {
            resource: "quick-xml event buffer",
            source,
        })?;
    let mut stack = Vec::new();
    stack
        .try_reserve_exact(max_depth.saturating_add(1))
        .map_err(|source| ScanError::Allocation {
            resource: "WordprocessingML scope stack",
            source,
        })?;
    let mut hasher = Sha256::new();
    if saw_bom {
        hasher.update([0xEF, 0xBB, 0xBF]);
    }
    let mut section_hasher = Sha256::new();
    let mut section_len = 0_u64;
    let mut section_depth = 0_u64;
    let mut section_seen = false;
    let mut saw_declaration = false;
    let mut saw_document = false;
    let mut saw_body = false;
    let mut finished_document = false;
    let mut strict_namespace = false;
    let mut insertion_offset = None;
    let mut paragraph_count = 0_u64;
    let mut event_count = 0_u64;
    let mut max_depth_seen = 0_u64;
    let mut generated_offset = None;
    let mut generated_once = false;
    let mut generated_open = false;
    let mut generated_count = 0_u64;
    let mut generated_last_end = None;

    loop {
        parser.get_mut().begin_token();
        let event_start = parser.get_mut().position();
        let (namespace, event) = parser
            .read_resolved_event_into(&mut buffer)
            .map_err(map_quick_xml_error)?;
        validate_event_name(&event)?;
        let dialect = namespace_dialect(&namespace);
        let opaque_namespace_known = !matches!(&namespace, ResolveResult::Unknown(_));
        let event_end = parser.get_mut().position();
        event_count = event_count.checked_add(1).ok_or(ScanError::Limit {
            resource: "XML events",
            actual: u64::MAX,
            maximum: limits.max_events,
        })?;
        if event_count > limits.max_events {
            return Err(ScanError::Limit {
                resource: "XML events",
                actual: event_count,
                maximum: limits.max_events,
            });
        }
        check_progress(context, cancellation, 1)?;
        let opaque = section_depth != 0 && !matches!(&event, Event::Eof);
        let direct_section = section_depth == 0
            && stack.as_slice() == [FrameKind::Document, FrameKind::Body]
            && match &event {
                Event::Start(start) | Event::Empty(start) => {
                    start.local_name().as_ref() == b"sectPr"
                },
                _ => false,
            };
        let raw_len = {
            let raw = parser.get_mut().captured();
            hasher.update(raw);
            if opaque || direct_section {
                section_hasher.update(raw);
            }
            raw.len()
        };
        if opaque {
            section_len = section_len
                .checked_add(raw_len as u64)
                .ok_or(ScanError::Limit {
                    resource: "section-properties bytes",
                    actual: u64::MAX,
                    maximum: limits.max_source_xml_bytes,
                })?;
            match event {
                Event::Start(start) => {
                    require_opaque_namespace(opaque_namespace_known)?;
                    validate_opaque_attributes(&parser, &start)?;
                    if start.local_name().as_ref() == b"sectPr" && dialect.is_some() {
                        return Err(ScanError::Semantic(
                            "nested WordprocessingML section-properties element is not admitted",
                        ));
                    }
                    section_depth = section_depth.checked_add(1).ok_or(ScanError::Limit {
                        resource: "XML depth",
                        actual: u64::MAX,
                        maximum: limits.max_depth,
                    })?;
                },
                Event::End(end) => {
                    require_opaque_namespace(opaque_namespace_known)?;
                    if section_depth == 1 && end.local_name().as_ref() != b"sectPr" {
                        return Err(ScanError::Semantic(
                            "section-properties end tag does not match its start",
                        ));
                    }
                    section_depth = section_depth.saturating_sub(1);
                },
                Event::Empty(empty) => {
                    require_opaque_namespace(opaque_namespace_known)?;
                    validate_opaque_attributes(&parser, &empty)?;
                    if empty.local_name().as_ref() == b"sectPr" && dialect.is_some() {
                        return Err(ScanError::Semantic(
                            "nested WordprocessingML section-properties element is not admitted",
                        ));
                    }
                },
                Event::Text(text) if text.as_ref().iter().all(u8::is_ascii_whitespace) => {},
                Event::Text(_) => {
                    return Err(ScanError::Semantic(
                        "non-whitespace text is outside the admitted opaque section policy",
                    ));
                },
                Event::CData(_)
                | Event::Comment(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_)
                | Event::Decl(_) => {
                    return Err(ScanError::Semantic(
                        "unsupported XML event is inside the opaque section",
                    ));
                },
                Event::Eof => unreachable!(),
            }
            max_depth_seen = max_depth_seen.max(
                u64::try_from(stack.len())
                    .unwrap_or(u64::MAX)
                    .saturating_add(section_depth),
            );
            let depth = u64::try_from(stack.len())
                .unwrap_or(u64::MAX)
                .saturating_add(section_depth);
            if depth > limits.max_depth {
                return Err(ScanError::Limit {
                    resource: "XML depth",
                    actual: depth,
                    maximum: limits.max_depth,
                });
            }
            buffer.clear();
            continue;
        }

        match event {
            Event::Start(start) => {
                let local = start.local_name();
                if stack.is_empty() && !saw_document {
                    strict_namespace = strict_from_dialect(dialect)?;
                    if local.as_ref() != b"document" {
                        return Err(ScanError::Semantic(
                            "WordprocessingML document root is missing",
                        ));
                    }
                    saw_document = true;
                    validate_attributes(
                        &start,
                        FrameKind::Document,
                        word_namespace(strict_namespace),
                    )?;
                    stack.push(FrameKind::Document);
                } else if stack.is_empty() {
                    return Err(ScanError::Semantic(
                        "more than one WordprocessingML document root",
                    ));
                } else {
                    require_word_dialect(dialect, strict_namespace)?;
                    if stack.as_slice() == [FrameKind::Document, FrameKind::Body]
                        && local.as_ref() == b"sectPr"
                    {
                        validate_opaque_attributes(&parser, &start)?;
                        if section_seen {
                            return Err(ScanError::Semantic(
                                "more than one direct section-properties element",
                            ));
                        }
                        section_seen = true;
                        insertion_offset = Some(event_start);
                        section_depth = 1;
                        section_len = raw_len as u64;
                        max_depth_seen = max_depth_seen.max(3);
                    } else {
                        if stack.as_slice() == [FrameKind::Document, FrameKind::Body]
                            && section_seen
                        {
                            return Err(ScanError::Semantic(
                                "direct body content follows section-properties",
                            ));
                        }
                        let frame = next_frame(&stack, local.as_ref())?;
                        validate_attributes(&start, frame, word_namespace(strict_namespace))?;
                        if matches!(frame, FrameKind::Body) {
                            if saw_body {
                                return Err(ScanError::Semantic(
                                    "more than one direct body element",
                                ));
                            }
                            saw_body = true;
                        }
                        if matches!(frame, FrameKind::Paragraph) {
                            paragraph_count =
                                paragraph_count.checked_add(1).ok_or(ScanError::Limit {
                                    resource: "paragraphs",
                                    actual: u64::MAX,
                                    maximum: limits.max_paragraphs,
                                })?;
                            let paragraph_limit = limits.max_paragraphs.saturating_add(
                                generated.map_or(0, |(_, _, paragraph_count)| paragraph_count),
                            );
                            if paragraph_count > paragraph_limit {
                                return Err(ScanError::Limit {
                                    resource: "paragraphs",
                                    actual: paragraph_count,
                                    maximum: paragraph_limit,
                                });
                            }
                            if let Some((expected, fragment_len, _)) = generated {
                                let generated_end = expected.saturating_add(fragment_len);
                                if event_start >= expected && event_start < generated_end {
                                    if generated_offset.is_none() {
                                        generated_offset = Some(event_start);
                                    }
                                    generated_open = true;
                                }
                            }
                        }
                        stack.push(frame);
                    }
                }
                let depth = u64::try_from(stack.len())
                    .unwrap_or(u64::MAX)
                    .saturating_add(section_depth);
                if depth > limits.max_depth {
                    return Err(ScanError::Limit {
                        resource: "XML depth",
                        actual: depth,
                        maximum: limits.max_depth,
                    });
                }
                max_depth_seen = max_depth_seen.max(depth);
            },
            Event::Empty(empty) => {
                let local = empty.local_name();
                require_word_dialect(dialect, strict_namespace)?;
                if stack.as_slice() == [FrameKind::Document, FrameKind::Body]
                    && local.as_ref() == b"sectPr"
                {
                    validate_opaque_attributes(&parser, &empty)?;
                    if section_seen {
                        return Err(ScanError::Semantic(
                            "more than one direct section-properties element",
                        ));
                    }
                    section_seen = true;
                    insertion_offset = Some(event_start);
                    section_len = raw_len as u64;
                    if 3 > limits.max_depth {
                        return Err(ScanError::Limit {
                            resource: "XML depth",
                            actual: 3,
                            maximum: limits.max_depth,
                        });
                    }
                    max_depth_seen = max_depth_seen.max(3);
                } else {
                    if stack.as_slice() == [FrameKind::Document, FrameKind::Body] && section_seen {
                        return Err(ScanError::Semantic(
                            "direct body content follows section-properties",
                        ));
                    }
                    let frame = next_frame(&stack, local.as_ref())?;
                    if matches!(
                        frame,
                        FrameKind::Document | FrameKind::Body | FrameKind::Run
                    ) {
                        return Err(ScanError::Semantic(
                            "empty document, body, or run is outside the plain-story closure",
                        ));
                    }
                    validate_attributes(&empty, frame, word_namespace(strict_namespace))?;
                    if matches!(frame, FrameKind::Paragraph) {
                        paragraph_count =
                            paragraph_count.checked_add(1).ok_or(ScanError::Limit {
                                resource: "paragraphs",
                                actual: u64::MAX,
                                maximum: limits.max_paragraphs,
                            })?;
                        let paragraph_limit = limits.max_paragraphs.saturating_add(
                            generated.map_or(0, |(_, _, paragraph_count)| paragraph_count),
                        );
                        if paragraph_count > paragraph_limit {
                            return Err(ScanError::Limit {
                                resource: "paragraphs",
                                actual: paragraph_count,
                                maximum: paragraph_limit,
                            });
                        }
                        if let Some((expected, fragment_len, expected_count)) = generated {
                            let generated_end = expected.saturating_add(fragment_len);
                            if event_start >= expected
                                && event_start < generated_end
                                && event_end <= generated_end
                            {
                                generated_offset.get_or_insert(event_start);
                                generated_count =
                                    generated_count.checked_add(1).ok_or(ScanError::Limit {
                                        resource: "generated paragraphs",
                                        actual: u64::MAX,
                                        maximum: expected_count,
                                    })?;
                                if generated_count > expected_count {
                                    return Err(ScanError::Semantic(
                                        "generated paragraph count exceeds the authored proof",
                                    ));
                                }
                                generated_last_end = Some(event_end);
                                if event_end == generated_end {
                                    generated_once = true;
                                }
                            }
                        }
                    }
                }
            },
            Event::End(end) => {
                require_word_dialect(dialect, strict_namespace)?;
                let frame = stack
                    .pop()
                    .ok_or(ScanError::Semantic("unexpected XML end tag"))?;
                if frame.local() != end.local_name().as_ref() {
                    return Err(ScanError::Semantic(
                        "WordprocessingML end tag does not match its start",
                    ));
                }
                if matches!(frame, FrameKind::Body) {
                    if insertion_offset.is_none() {
                        insertion_offset = Some(event_start);
                    }
                    if !saw_body {
                        return Err(ScanError::Semantic("WordprocessingML body is missing"));
                    }
                } else if matches!(frame, FrameKind::Document) {
                    finished_document = true;
                } else if matches!(frame, FrameKind::Paragraph) && generated_open {
                    if let Some((expected, fragment_len, expected_count)) = generated {
                        let generated_end = expected.saturating_add(fragment_len);
                        if event_end <= generated_end {
                            generated_count =
                                generated_count.checked_add(1).ok_or(ScanError::Limit {
                                    resource: "generated paragraphs",
                                    actual: u64::MAX,
                                    maximum: expected_count,
                                })?;
                            if generated_count > expected_count {
                                return Err(ScanError::Semantic(
                                    "generated paragraph count exceeds the authored proof",
                                ));
                            }
                            generated_last_end = Some(event_end);
                        }
                        if event_end == generated_end {
                            generated_once = true;
                        }
                    }
                    generated_open = false;
                }
                let depth = u64::try_from(stack.len()).unwrap_or(u64::MAX);
                max_depth_seen = max_depth_seen.max(depth);
            },
            Event::Text(text) => {
                if stack.last() != Some(&FrameKind::Text)
                    && text.as_ref().iter().any(|byte| !byte.is_ascii_whitespace())
                {
                    return Err(ScanError::Semantic(
                        "non-whitespace text occurs outside a plain text element",
                    ));
                }
            },
            Event::Decl(declaration) if stack.is_empty() && !saw_document => {
                if saw_declaration {
                    return Err(ScanError::Semantic("duplicate XML declaration"));
                }
                validate_xml_declaration(&declaration)?;
                saw_declaration = true;
            },
            Event::GeneralRef(reference)
                if stack.last() == Some(&FrameKind::Text)
                    && matches!(
                        reference.as_ref(),
                        b"amp" | b"lt" | b"gt" | b"apos" | b"quot"
                    ) => {},
            Event::Eof => break,
            Event::Comment(_)
            | Event::CData(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {
                return Err(ScanError::Semantic("unsupported XML event in plain story"));
            },
            _ => return Err(ScanError::Semantic("unsupported XML event in plain story")),
        }
        buffer.clear();
    }
    if parser.get_mut().position() != expected_len {
        return Err(ScanError::Semantic(
            "decoded candidate length did not reach its bound",
        ));
    }
    if !saw_document || !saw_body || !finished_document || !stack.is_empty() || section_depth != 0 {
        return Err(ScanError::Semantic(
            "plain story XML topology is incomplete",
        ));
    }
    let insertion_offset = insertion_offset.ok_or(ScanError::Semantic(
        "plain story has no direct body insertion point",
    ))?;
    if let Some((expected, fragment_len, expected_count)) = generated {
        if !generated_once
            || generated_offset != Some(expected)
            || generated_last_end != Some(expected.saturating_add(fragment_len))
            || generated_count != expected_count
        {
            return Err(ScanError::Semantic(
                "generated paragraph was not closed at its fragment boundary",
            ));
        }
    }
    Ok(ScanFacts {
        len: parser.get_mut().position(),
        sha256: hasher.finalize().into(),
        insertion_offset,
        paragraph_count,
        event_count,
        max_depth: max_depth_seen,
        strict_namespace,
        sect_pr_len: section_len,
        sect_pr_sha256: section_hasher.finalize().into(),
        generated_offset: generated_offset.unwrap_or(0),
        generated_once,
        generated_count,
    })
}

pub(super) fn scan_reader_checked<R: BufRead>(
    reader: &mut R,
    expected_len: u64,
    limits: Limits,
    context: Option<&ExecutionContext>,
    cancellation: Option<&CancellationToken>,
    generated: Option<(u64, u64, u64)>,
) -> Result<ScanFacts> {
    scan_reader(
        reader,
        expected_len,
        limits,
        context,
        cancellation,
        generated,
    )
    .map_err(map_scan_error)
}

fn validate_event_name(event: &Event<'_>) -> std::result::Result<(), ScanError> {
    match event {
        Event::Start(element) | Event::Empty(element) => validate_qname(element.name().as_ref()),
        Event::End(element) => validate_qname(element.name().as_ref()),
        _ => Ok(()),
    }
}

fn validate_qname(name: &[u8]) -> std::result::Result<(), ScanError> {
    let name = std::str::from_utf8(name)
        .map_err(|_| ScanError::Semantic("XML qualified name is not UTF-8"))?;
    if !is_qualified_name(name) {
        return Err(ScanError::Semantic(
            "XML qualified name is outside XML 1.0 grammar",
        ));
    }
    Ok(())
}

fn validate_xml_declaration(declaration: &BytesDecl<'_>) -> std::result::Result<(), ScanError> {
    let declaration_text =
        std::str::from_utf8(declaration.as_ref()).map_err(|_| ScanError::Parser)?;
    let raw = BytesStart::from_content(declaration_text, 3);
    let mut state = 0_u8;
    for attribute in raw.attributes().with_checks(true) {
        let attribute = attribute.map_err(|_| ScanError::Parser)?;
        if attribute.key.prefix().is_some() {
            return Err(ScanError::Semantic(
                "XML declaration attributes must be unprefixed",
            ));
        }
        std::str::from_utf8(attribute.value.as_ref()).map_err(|_| ScanError::Parser)?;
        state = match (state, attribute.key.as_ref()) {
            (0, b"version") => {
                if attribute.value.as_ref() != b"1.0" {
                    return Err(ScanError::Semantic(
                        "XML declaration must select XML 1.0 for the bounded UTF-8 closure",
                    ));
                }
                1
            },
            (1, b"encoding") => {
                if !attribute.value.as_ref().eq_ignore_ascii_case(b"UTF-8") {
                    return Err(ScanError::Semantic(
                        "XML declaration encoding must be UTF-8 for the authored fragment",
                    ));
                }
                2
            },
            (1 | 2, b"standalone") => {
                if !matches!(attribute.value.as_ref(), b"yes" | b"no") {
                    return Err(ScanError::Semantic(
                        "XML declaration standalone must be 'yes' or 'no'",
                    ));
                }
                3
            },
            _ => {
                return Err(ScanError::Semantic(
                    "XML declaration has duplicate, unknown, or out-of-order attributes",
                ));
            },
        };
    }
    if state == 0 {
        return Err(ScanError::Semantic(
            "XML declaration must contain a version",
        ));
    }
    Ok(())
}

fn namespace_dialect(namespace: &ResolveResult<'_>) -> Option<bool> {
    match namespace {
        ResolveResult::Bound(Namespace(value)) if *value == STRICT_WORDPROCESSINGML_NAMESPACE => {
            Some(true)
        },
        ResolveResult::Bound(Namespace(value)) if *value == WORDPROCESSINGML_NAMESPACE => {
            Some(false)
        },
        ResolveResult::Bound(_) | ResolveResult::Unknown(_) | ResolveResult::Unbound => None,
    }
}

fn require_opaque_namespace(namespace_known: bool) -> std::result::Result<(), ScanError> {
    if !namespace_known {
        return Err(ScanError::Semantic(
            "opaque section element uses an undeclared namespace prefix",
        ));
    }
    Ok(())
}

fn strict_from_dialect(namespace: Option<bool>) -> std::result::Result<bool, ScanError> {
    namespace.ok_or(ScanError::Semantic(
        "WordprocessingML namespace is not admitted",
    ))
}

fn require_word_dialect(
    namespace: Option<bool>,
    strict: bool,
) -> std::result::Result<(), ScanError> {
    if namespace == Some(strict) {
        Ok(())
    } else {
        Err(ScanError::Semantic(
            "WordprocessingML element namespace changed",
        ))
    }
}

const fn word_namespace(strict: bool) -> &'static [u8] {
    if strict {
        STRICT_WORDPROCESSINGML_NAMESPACE
    } else {
        WORDPROCESSINGML_NAMESPACE
    }
}

fn next_frame(stack: &[FrameKind], local: &[u8]) -> std::result::Result<FrameKind, ScanError> {
    match (stack, local) {
        ([], b"document") => Ok(FrameKind::Document),
        ([FrameKind::Document], b"body") => Ok(FrameKind::Body),
        ([FrameKind::Document, FrameKind::Body], b"p") => Ok(FrameKind::Paragraph),
        ([FrameKind::Document, FrameKind::Body, FrameKind::Paragraph], b"r") => Ok(FrameKind::Run),
        (
            [
                FrameKind::Document,
                FrameKind::Body,
                FrameKind::Paragraph,
                FrameKind::Run,
            ],
            b"t",
        ) => Ok(FrameKind::Text),
        _ => Err(ScanError::Semantic(
            if stack.iter().any(|frame| {
                matches!(
                    frame,
                    FrameKind::Paragraph | FrameKind::Run | FrameKind::Text
                )
            }) {
                "complex paragraph content is outside the bounded closure"
            } else {
                "complex document content is outside the bounded closure"
            },
        )),
    }
}

fn validate_opaque_attributes<R>(
    parser: &NsReader<R>,
    element: &BytesStart<'_>,
) -> std::result::Result<(), ScanError> {
    let mut seen = Vec::<(&[u8], &[u8])>::new();
    seen.try_reserve(element.attributes().count())
        .map_err(|source| ScanError::Allocation {
            resource: "opaque section attribute names",
            source,
        })?;
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ScanError::Parser)?;
        validate_qname(attribute.key.as_ref())?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            // NamespaceResolver::push has already authenticated reserved
            // `xml`/`xmlns` bindings before this event was returned.  Keep
            // declarations opaque while checking ordinary qualified attrs.
            let value = attribute.value.as_ref();
            if value.contains(&b'&')
                || (key == b"xmlns" && (value == XML_NAMESPACE || value == XMLNS_NAMESPACE))
                || (key.starts_with(b"xmlns:") && value.is_empty())
            {
                return Err(ScanError::Semantic(
                    "opaque section namespace declaration is outside the admitted policy",
                ));
            }
            continue;
        }
        let namespace = match parser.resolver().resolve_attribute(attribute.key).0 {
            ResolveResult::Bound(Namespace(value)) => value,
            ResolveResult::Unbound => &[],
            ResolveResult::Unknown(_) => {
                return Err(ScanError::Semantic(
                    "opaque section attribute uses an undeclared namespace prefix",
                ));
            },
        };
        let local = attribute.key.local_name().into_inner();
        seen.push((namespace, local));
    }
    seen.sort_unstable();
    if seen.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(ScanError::Semantic(
            "opaque section contains duplicate expanded attributes",
        ));
    }
    Ok(())
}

fn validate_attributes(
    element: &BytesStart<'_>,
    frame: FrameKind,
    word_namespace: &[u8],
) -> std::result::Result<(), ScanError> {
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|_| ScanError::Parser)?;
        validate_qname(attribute.key.as_ref())?;
        let key = attribute.key.as_ref();
        let value = attribute.value.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            if value.contains(&b'&')
                || (key == b"xmlns" && (value == XML_NAMESPACE || value == XMLNS_NAMESPACE))
                || (key.starts_with(b"xmlns:") && value.is_empty())
            {
                return Err(ScanError::Semantic(
                    "namespace declaration is outside the admitted policy",
                ));
            }
            if !matches!(frame, FrameKind::Document)
                && value != word_namespace
                && value != XML_NAMESPACE
                && value != XMLNS_NAMESPACE
            {
                return Err(ScanError::Semantic(
                    "unexpected namespace declaration in plain story",
                ));
            }
            continue;
        }
        if frame == FrameKind::Text && key == b"xml:space" {
            if !matches!(value, b"preserve" | b"default") {
                return Err(ScanError::Semantic("invalid xml:space value in plain text"));
            }
            continue;
        }
        return Err(ScanError::Semantic(
            "attribute is outside the plain-story closure",
        ));
    }
    Ok(())
}

fn validate_candidate(
    source: &ScanFacts,
    candidate: &ScanFacts,
    fragment_len: u64,
    limits: Limits,
) -> Result<()> {
    let expected_len = source.len.checked_add(fragment_len).ok_or(Error::Limit {
        resource: "candidate XML bytes",
        actual: u64::MAX,
        maximum: limits.max_candidate_xml_bytes,
    })?;
    let expected_candidate_anchor =
        source
            .insertion_offset
            .checked_add(fragment_len)
            .ok_or(Error::Limit {
                resource: "candidate XML bytes",
                actual: u64::MAX,
                maximum: limits.max_candidate_xml_bytes,
            })?;
    if candidate.len != expected_len
        || candidate.paragraph_count != source.paragraph_count.saturating_add(1)
        || !candidate.generated_once
        || candidate.generated_offset != source.insertion_offset
        || candidate.insertion_offset != expected_candidate_anchor
        || candidate.strict_namespace != source.strict_namespace
        || candidate.sect_pr_len != source.sect_pr_len
        || candidate.sect_pr_sha256 != source.sect_pr_sha256
        || candidate.event_count <= source.event_count
        || candidate.max_depth < source.max_depth
    {
        return Err(Error::Scan(
            "candidate semantic proof does not match the source insertion".into(),
        ));
    }
    Ok(())
}
