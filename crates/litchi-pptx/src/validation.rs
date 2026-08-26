//! Bounded, non-mutating semantic validation of a PresentationML package.
//!
//! The validator deliberately stops at content-free package facts.  It checks
//! the PresentationML root, the ordered slide relationship closure, and the
//! already indexed OPC relationship/content-type catalog.  It reports the
//! presence of external targets, signature infrastructure, VBA/macro
//! storage, and markup-compatibility input without fetching, executing,
//! repairing, selecting an MCE branch, or retaining document text.

use std::{error::Error as StdError, fmt, sync::Arc};

use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceDigest, EvidenceValue,
    IssueEvidence, IssueLocation, IssueSeverity, ReadAt, RepairAvailability, SourceVersion,
    ValidateReport, ValidationCheck, ValidationIssue, ValidationLimitKind, ValidationLimits,
    ValidationReportError,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcError, ReadLimits, SourceBackedPackage};
use quick_xml::XmlVersion;
use quick_xml::escape::resolve_xml_entity;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::namespace;

const INGRESS: &str = "pptx.package.ingress";
const CATALOG: &str = "pptx.package.loaded_relationships_content_types";
const GRAPH: &str = "pptx.package.relationship_graph";
const ROOT: &str = "pptx.presentation.root";
const SLIDES: &str = "pptx.presentation.ordered_slide_closure";
const EXTERNAL: &str = "pptx.package.external_target_presence";
const SIGNATURES: &str = "pptx.package.signature_presence";
const MACROS: &str = "pptx.package.macro_presence";
const MCE: &str = "pptx.presentation.mce_presence";

const MCE_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const DEFAULT_XML_BYTES: usize = 64 * 1024 * 1024;
const DEFAULT_OWNER_BYTES: usize = 128 * 1024 * 1024;
const DEFAULT_SLIDES: usize = 100_000;
const DEFAULT_GRAPH_NODES: usize = 100_000;
const DEFAULT_XML_EVENTS: usize = 4_000_000;
const DEFAULT_XML_DEPTH: usize = 4_096;
const MAX_READER_DEPTH: usize = u16::MAX as usize - 1;
const MIN_SLIDE_ID: u64 = 256;
const MAX_SLIDE_ID: u64 = 2_147_483_647;

/// Finite semantic limits for one PPTX validation pass.
///
/// OPC archive, member, relationship, and ZIP/XML-manifest limits remain
/// owned by [`ReadLimits`].  These additional ceilings bound the XML parts
/// materialized by this format-level report, the aggregate validated owner
/// bytes, and the raw relationship graph walk.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct PptxValidationLimits {
    read_limits: ReadLimits,
    max_xml_bytes: usize,
    max_owner_bytes: usize,
    max_slides: usize,
    max_graph_nodes: usize,
    max_xml_events: usize,
    max_xml_depth: usize,
}

impl PptxValidationLimits {
    /// Returns the OPC ingress policy used when opening a positional source.
    #[must_use]
    pub const fn read_limits(self) -> ReadLimits {
        self.read_limits
    }

    /// Return a copy using a caller-provided OPC ingress policy.
    #[must_use]
    pub const fn with_read_limits(mut self, limits: ReadLimits) -> Self {
        self.read_limits = limits;
        self
    }

    /// Return a copy with a per-PresentationML XML part ceiling.
    #[must_use]
    pub const fn with_max_xml_bytes(mut self, maximum: usize) -> Self {
        self.max_xml_bytes = maximum;
        self
    }

    /// Return a copy with an aggregate materialized owner-byte ceiling.
    #[must_use]
    pub const fn with_max_owner_bytes(mut self, maximum: usize) -> Self {
        self.max_owner_bytes = maximum;
        self
    }

    /// Return a copy with an ordered slide-reference ceiling.
    #[must_use]
    pub const fn with_max_slides(mut self, maximum: usize) -> Self {
        self.max_slides = maximum;
        self
    }

    /// Return a copy with a relationship-graph node ceiling.
    #[must_use]
    pub const fn with_max_graph_nodes(mut self, maximum: usize) -> Self {
        self.max_graph_nodes = maximum;
        self
    }

    /// Return a copy with an XML event ceiling per inspected XML part.
    #[must_use]
    pub const fn with_max_xml_events(mut self, maximum: usize) -> Self {
        self.max_xml_events = maximum;
        self
    }

    /// Return a copy with an XML nesting-depth ceiling per inspected XML part.
    #[must_use]
    pub const fn with_max_xml_depth(mut self, maximum: usize) -> Self {
        self.max_xml_depth = if maximum > MAX_READER_DEPTH {
            MAX_READER_DEPTH
        } else {
            maximum
        };
        self
    }

    /// Maximum bytes in one inspected PresentationML XML part.
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    /// Maximum aggregate bytes retained while validating presentation owners.
    #[must_use]
    pub const fn max_owner_bytes(self) -> usize {
        self.max_owner_bytes
    }

    /// Maximum ordered slide references.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.max_slides
    }

    /// Maximum relationship graph nodes inspected.
    #[must_use]
    pub const fn max_graph_nodes(self) -> usize {
        self.max_graph_nodes
    }

    /// Maximum XML events inspected in one XML part.
    #[must_use]
    pub const fn max_xml_events(self) -> usize {
        self.max_xml_events
    }

    /// Maximum XML nesting depth inspected in one XML part.
    #[must_use]
    pub const fn max_xml_depth(self) -> usize {
        self.max_xml_depth
    }
}

impl Default for PptxValidationLimits {
    fn default() -> Self {
        Self {
            read_limits: ReadLimits::default(),
            max_xml_bytes: DEFAULT_XML_BYTES,
            max_owner_bytes: DEFAULT_OWNER_BYTES,
            max_slides: DEFAULT_SLIDES,
            max_graph_nodes: DEFAULT_GRAPH_NODES,
            max_xml_events: DEFAULT_XML_EVENTS,
            max_xml_depth: DEFAULT_XML_DEPTH,
        }
    }
}

/// Failure to perform or retain a PPTX validation report.
#[derive(Debug)]
#[non_exhaustive]
pub enum PptxValidationError {
    /// Source I/O, source instability, or a bounded OPC ingress failure.
    Ingress(OpcError),
    /// The shared bounded report rejected the staged result.
    Report(ValidationReportError),
    /// A validator-owned bounded collection could not be grown.
    Allocation {
        /// Content-free validator resource name.
        resource: &'static str,
    },
}

impl fmt::Display for PptxValidationError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Ingress(error) => write!(formatter, "PPTX validation ingress failed: {error}"),
            Self::Report(error) => write!(formatter, "PPTX validation report failed: {error}"),
            Self::Allocation { resource } => {
                write!(
                    formatter,
                    "PPTX validation allocation failed for {resource}"
                )
            },
        }
    }
}

impl StdError for PptxValidationError {
    fn source(&self) -> Option<&(dyn StdError + 'static)> {
        match self {
            Self::Ingress(error) => Some(error),
            Self::Report(error) => Some(error),
            Self::Allocation { .. } => None,
        }
    }
}

impl From<ValidationReportError> for PptxValidationError {
    fn from(error: ValidationReportError) -> Self {
        Self::Report(error)
    }
}

/// Validate a positional PPTX source under default finite input and report
/// limits without changing it, fetching external targets, or executing code.
pub fn validate_source(source: Arc<dyn ReadAt>) -> Result<ValidateReport, PptxValidationError> {
    validate_source_with_limits(
        source,
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
}

/// Validate a positional PPTX source under explicit finite policies.
///
/// Structural OPC rejection and semantic PresentationML rejection are
/// represented as bounded report issues. Source I/O, source mutation,
/// cancellation, and allocation failures remain errors because no honest
/// report can be made for those conditions.
pub fn validate_source_with_limits(
    source: Arc<dyn ReadAt>,
    input_limits: PptxValidationLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport, PptxValidationError> {
    let expected_version = source
        .version()
        .map_err(|error| PptxValidationError::Ingress(OpcError::IoError(error)))?;
    let result = match SourceBackedPackage::from_read_at_with_limits(
        Arc::clone(&source),
        input_limits.read_limits,
    ) {
        Ok(package) => validate_source_backed_with_limits(&package, input_limits, report_limits),
        Err(OpcError::ReadLimit { .. }) => blocked_report(
            "PPTX OPC ingress reached a configured resource ceiling",
            report_limits,
        ),
        Err(error) if is_structural_rejection(&error) => {
            rejected_ingress_report(&error, report_limits)
        },
        Err(error) => Err(PptxValidationError::Ingress(error)),
    };
    let actual_version = source
        .version()
        .map_err(|error| PptxValidationError::Ingress(OpcError::IoError(error)))?;
    if actual_version != expected_version {
        return Err(PptxValidationError::Ingress(OpcError::SourceChanged {
            expected: expected_version,
            actual: actual_version,
        }));
    }
    result
}

/// Validate an already indexed source-backed OPC package.
///
/// This entry point reuses the package's validated part index, relationship
/// manifests, and content-type catalog. It materializes only the mandatory
/// presentation root and ordered slides needed by this report.
pub fn validate_source_backed(
    package: &SourceBackedPackage,
) -> Result<ValidateReport, PptxValidationError> {
    validate_source_backed_with_limits(
        package,
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
}

/// Validate an already indexed source-backed OPC package with explicit format
/// and report limits.
pub fn validate_source_backed_with_limits(
    package: &SourceBackedPackage,
    input_limits: PptxValidationLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport, PptxValidationError> {
    let source_version = package
        .source_version()
        .map_err(PptxValidationError::Ingress)?;

    let mut facts = Facts::new(report_limits, source_version)?;
    let (catalog, graph) = inspect_catalog_and_graph(package, input_limits.max_graph_nodes())?;
    facts.graph = if graph.blocked {
        GraphOutcome::Blocked
    } else {
        GraphOutcome::Complete
    };
    facts.graph_missing_targets = graph.missing_targets;
    facts.graph_invalid_targets = graph.invalid_targets;

    let root = match package.main_document_part() {
        Ok(view) => view,
        Err(error) if is_transient(&error) => return Err(PptxValidationError::Ingress(error)),
        Err(_error) => {
            facts.root = RootOutcome::Malformed;
            facts.root_issue = true;
            facts.slides = SlideOutcome::NotApplicable;
            facts.root_error = Some(RootErrorKind::Missing);
            facts.mce = MceOutcome::NotApplicable;
            return finish(package, input_limits, report_limits, facts, catalog, None);
        },
    };

    if !is_presentation_main_content_type(root.content_type()) {
        facts.root = RootOutcome::Malformed;
        facts.root_issue = true;
        facts.root_error = Some(RootErrorKind::WrongContentType);
        facts.slides = SlideOutcome::NotApplicable;
        facts.mce = MceOutcome::NotApplicable;
        return finish(
            package,
            input_limits,
            report_limits,
            facts,
            catalog,
            Some(root),
        );
    }

    let data = match root.data() {
        Ok(data) => data,
        Err(OpcError::ReadLimit { .. }) => {
            facts.root = RootOutcome::Blocked;
            facts.slides = SlideOutcome::StoppedByRoot;
            facts.mce = MceOutcome::StoppedByRoot;
            return finish(
                package,
                input_limits,
                report_limits,
                facts,
                catalog,
                Some(root),
            );
        },
        Err(error) if is_transient(&error) => return Err(PptxValidationError::Ingress(error)),
        Err(_) => {
            facts.root = RootOutcome::Malformed;
            facts.root_issue = true;
            facts.root_error = Some(RootErrorKind::Unreadable);
            facts.slides = SlideOutcome::NotApplicable;
            facts.mce = MceOutcome::NotApplicable;
            return finish(
                package,
                input_limits,
                report_limits,
                facts,
                catalog,
                Some(root),
            );
        },
    };

    if data.as_bytes().len() > input_limits.max_xml_bytes()
        || data.as_bytes().len() > input_limits.max_owner_bytes()
    {
        facts.root = RootOutcome::Blocked;
        facts.slides = SlideOutcome::StoppedByRoot;
        facts.mce = MceOutcome::StoppedByRoot;
        return finish(
            package,
            input_limits,
            report_limits,
            facts,
            catalog,
            Some(root),
        );
    }
    facts.owner_bytes = data.as_bytes().len();

    let presentation = match inspect_xml(data.as_bytes(), input_limits, XmlOwner::Presentation) {
        Ok(observation) => observation,
        Err(XmlInspectionError::Limit) => {
            facts.root = RootOutcome::Blocked;
            facts.slides = SlideOutcome::StoppedByRoot;
            facts.mce = MceOutcome::StoppedByRoot;
            return finish(
                package,
                input_limits,
                report_limits,
                facts,
                catalog,
                Some(root),
            );
        },
        Err(XmlInspectionError::Allocation { resource }) => {
            return Err(PptxValidationError::Allocation { resource });
        },
        Err(XmlInspectionError::Malformed) => {
            facts.root = RootOutcome::Malformed;
            facts.root_issue = true;
            facts.root_error = Some(RootErrorKind::Malformed);
            facts.slides = SlideOutcome::NotApplicable;
            facts.mce = MceOutcome::NotApplicable;
            return finish(
                package,
                input_limits,
                report_limits,
                facts,
                catalog,
                Some(root),
            );
        },
    };
    facts.root = RootOutcome::Complete;
    facts.mce = if presentation.mce {
        MceOutcome::Present(1)
    } else {
        MceOutcome::NotApplicable
    };

    if presentation.mce {
        facts.slides = SlideOutcome::NotApplicable;
        return finish(
            package,
            input_limits,
            report_limits,
            facts,
            catalog,
            Some(root),
        );
    }

    if presentation.slides.len() > input_limits.max_slides() {
        facts.slides = SlideOutcome::Blocked;
        return finish(
            package,
            input_limits,
            report_limits,
            facts,
            catalog,
            Some(root),
        );
    }

    let mut seen_slides = std::collections::HashSet::<String>::new();
    seen_slides
        .try_reserve(presentation.slides.len())
        .map_err(|_| PptxValidationError::Allocation {
            resource: "ordered slide identity set",
        })?;
    let mut seen_slide_ids = std::collections::HashSet::<u64>::new();
    seen_slide_ids
        .try_reserve(presentation.slides.len())
        .map_err(|_| PptxValidationError::Allocation {
            resource: "ordered slide numeric identity set",
        })?;
    let mut seen_relationship_ids = std::collections::HashSet::<&str>::new();
    seen_relationship_ids
        .try_reserve(presentation.slides.len())
        .map_err(|_| PptxValidationError::Allocation {
            resource: "ordered slide relationship identity set",
        })?;
    let mut duplicate_slides = 0_u64;
    let mut duplicate_slide_ids = 0_u64;
    let mut duplicate_relationship_ids = 0_u64;
    let mut slide_blocked = false;
    for reference in &presentation.slides {
        if !seen_slide_ids.insert(reference.numeric_id) {
            duplicate_slide_ids = duplicate_slide_ids.saturating_add(1);
        }
        if !seen_relationship_ids.insert(reference.relationship_id.as_str()) {
            duplicate_relationship_ids = duplicate_relationship_ids.saturating_add(1);
        }
        let Some(relationship) = root.rels().get(&reference.relationship_id) else {
            block_mce(&mut facts);
            facts.missing_slide_relationships = facts.missing_slide_relationships.saturating_add(1);
            continue;
        };
        if relationship.is_external()
            || !crate::parts::is_relationship_type(relationship.reltype(), rt::SLIDE, "slide")
        {
            block_mce(&mut facts);
            facts.invalid_slide_relationships = facts.invalid_slide_relationships.saturating_add(1);
            continue;
        }
        let target = match relationship.target_partname() {
            Ok(target) => target,
            Err(_) => {
                block_mce(&mut facts);
                facts.invalid_slide_relationships =
                    facts.invalid_slide_relationships.saturating_add(1);
                continue;
            },
        };
        if !record_owned_text(
            &mut seen_slides,
            target.as_str(),
            "ordered slide identity set",
        )? {
            duplicate_slides = duplicate_slides.saturating_add(1);
        }
        let slide = match package.part(&target) {
            Ok(slide) => slide,
            Err(error) if is_transient(&error) => {
                return Err(PptxValidationError::Ingress(error));
            },
            Err(_) => {
                block_mce(&mut facts);
                facts.missing_slide_parts = facts.missing_slide_parts.saturating_add(1);
                continue;
            },
        };
        if slide.content_type() != ct::PML_SLIDE {
            block_mce(&mut facts);
            facts.invalid_slide_parts = facts.invalid_slide_parts.saturating_add(1);
            continue;
        }
        let slide_data = match slide.data() {
            Ok(data) => data,
            Err(OpcError::ReadLimit { .. }) => {
                block_mce(&mut facts);
                slide_blocked = true;
                break;
            },
            Err(error) if is_transient(&error) => {
                return Err(PptxValidationError::Ingress(error));
            },
            Err(_) => {
                block_mce(&mut facts);
                facts.malformed_slides = facts.malformed_slides.saturating_add(1);
                continue;
            },
        };
        if slide_data.as_bytes().len() > input_limits.max_xml_bytes() {
            block_mce(&mut facts);
            slide_blocked = true;
            break;
        }
        facts.owner_bytes = facts
            .owner_bytes
            .checked_add(slide_data.as_bytes().len())
            .unwrap_or(usize::MAX);
        if facts.owner_bytes > input_limits.max_owner_bytes() {
            block_mce(&mut facts);
            slide_blocked = true;
            break;
        }
        match inspect_xml(slide_data.as_bytes(), input_limits, XmlOwner::Slide) {
            Ok(observation) => {
                if observation.mce {
                    note_mce(&mut facts);
                    continue;
                }
            },
            Err(XmlInspectionError::Limit) => {
                block_mce(&mut facts);
                slide_blocked = true;
                break;
            },
            Err(XmlInspectionError::Allocation { resource }) => {
                return Err(PptxValidationError::Allocation { resource });
            },
            Err(XmlInspectionError::Malformed) => {
                facts.malformed_slides = facts.malformed_slides.saturating_add(1);
                block_mce(&mut facts);
            },
        }
    }
    facts.duplicate_slides = duplicate_slides;
    facts.duplicate_slide_ids = duplicate_slide_ids;
    facts.duplicate_relationship_ids = duplicate_relationship_ids;
    if matches!(facts.mce, MceOutcome::Present(_)) {
        facts.slides = SlideOutcome::NotApplicable;
    } else if slide_blocked {
        facts.slides = SlideOutcome::Blocked;
        facts.mce = MceOutcome::Blocked;
    } else {
        facts.slides = SlideOutcome::Complete;
    }

    finish(
        package,
        input_limits,
        report_limits,
        facts,
        catalog,
        Some(root),
    )
}

#[derive(Clone, Copy)]
enum XmlOwner {
    Presentation,
    Slide,
}

#[derive(Debug)]
enum XmlInspectionError {
    Limit,
    Malformed,
    Allocation { resource: &'static str },
}

struct XmlObservation {
    slides: Vec<SlideReferenceObservation>,
    mce: bool,
}

struct SlideReferenceObservation {
    numeric_id: u64,
    relationship_id: String,
}

struct XmlElementFrame {
    name: Vec<u8>,
    local_name: Vec<u8>,
    presentationml: bool,
}

struct XmlAttributeObservation {
    mce: bool,
    numeric_id: Option<u64>,
    relationship_id: Option<String>,
}

fn inspect_xml(
    bytes: &[u8],
    limits: PptxValidationLimits,
    owner: XmlOwner,
) -> Result<XmlObservation, XmlInspectionError> {
    let mut mce = false;
    let mut reader = NsReader::from_reader(bytes);
    let mut stack = Vec::<XmlElementFrame>::new();
    stack
        .try_reserve(1)
        .map_err(|_| XmlInspectionError::Allocation {
            resource: "PPTX XML element stack",
        })?;
    let mut slides = Vec::new();
    if matches!(owner, XmlOwner::Presentation) {
        slides
            .try_reserve(1)
            .map_err(|_| XmlInspectionError::Allocation {
                resource: "PPTX ordered slide references",
            })?;
    }
    let mut events = 0_usize;
    let mut root_seen = false;
    let mut root_is_expected = false;
    let mut direct_slide_lists = 0_usize;
    loop {
        if stack.len() >= limits.max_xml_depth().min(MAX_READER_DEPTH) {
            return Err(XmlInspectionError::Limit);
        }
        events = events.saturating_add(1);
        if events > limits.max_xml_events() {
            return Err(XmlInspectionError::Limit);
        }
        let (resolved, event) = reader
            .read_resolved_event()
            .map_err(|_| XmlInspectionError::Malformed)?;
        match event {
            Event::Start(element) => {
                if root_seen && stack.is_empty() {
                    return Err(XmlInspectionError::Malformed);
                }
                let depth = stack.len().saturating_add(1);
                if depth > limits.max_xml_depth() {
                    return Err(XmlInspectionError::Limit);
                }
                if has_unresolved_prefix(&resolved, element.name()) {
                    return Err(XmlInspectionError::Malformed);
                }
                let local_name = element.name().local_name();
                let presentationml =
                    is_presentationml_name_strict(&resolved, element.name(), local_name.as_ref());
                let is_expected_root = match owner {
                    XmlOwner::Presentation => {
                        presentationml && local_name.as_ref() == b"presentation"
                    },
                    XmlOwner::Slide => presentationml && local_name.as_ref() == b"sld",
                };
                let element_is_mce = is_mce_namespace(&resolved);
                let is_slide_id = matches!(owner, XmlOwner::Presentation)
                    && presentationml
                    && local_name.as_ref() == b"sldId"
                    && is_direct_slide_owner(&stack);
                let is_direct_slide_list = matches!(owner, XmlOwner::Presentation)
                    && presentationml
                    && local_name.as_ref() == b"sldIdLst"
                    && is_direct_slide_list_owner(&stack);
                if is_direct_slide_list {
                    direct_slide_lists = direct_slide_lists.saturating_add(1);
                    if direct_slide_lists > 1 {
                        return Err(XmlInspectionError::Malformed);
                    }
                }
                if is_direct_slide_owner(&stack)
                    && !(presentationml && local_name.as_ref() == b"sldId")
                {
                    return Err(XmlInspectionError::Malformed);
                }
                let attributes = validate_attributes(
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    is_slide_id,
                )?;
                if !root_seen {
                    root_seen = true;
                    root_is_expected = is_expected_root;
                }
                mce |= element_is_mce || attributes.mce;
                if is_slide_id {
                    let numeric_id = attributes.numeric_id.ok_or(XmlInspectionError::Malformed)?;
                    let relationship_id = attributes
                        .relationship_id
                        .ok_or(XmlInspectionError::Malformed)?;
                    if slides.len() >= limits.max_slides() {
                        return Err(XmlInspectionError::Limit);
                    }
                    slides
                        .try_reserve(1)
                        .map_err(|_| XmlInspectionError::Allocation {
                            resource: "PPTX ordered slide references",
                        })?;
                    slides.push(SlideReferenceObservation {
                        numeric_id,
                        relationship_id,
                    });
                }
                stack
                    .try_reserve(1)
                    .map_err(|_| XmlInspectionError::Allocation {
                        resource: "PPTX XML element stack",
                    })?;
                stack.push(XmlElementFrame {
                    name: copy_xml_name(element.name().as_ref(), "PPTX XML element stack")?,
                    local_name: copy_xml_name(local_name.as_ref(), "PPTX XML element stack")?,
                    presentationml,
                });
            },
            Event::Empty(element) => {
                if root_seen && stack.is_empty() {
                    return Err(XmlInspectionError::Malformed);
                }
                let depth = stack.len().saturating_add(1);
                if depth > limits.max_xml_depth() {
                    return Err(XmlInspectionError::Limit);
                }
                if has_unresolved_prefix(&resolved, element.name()) {
                    return Err(XmlInspectionError::Malformed);
                }
                let local_name = element.name().local_name();
                let presentationml =
                    is_presentationml_name_strict(&resolved, element.name(), local_name.as_ref());
                let is_expected_root = match owner {
                    XmlOwner::Presentation => {
                        presentationml && local_name.as_ref() == b"presentation"
                    },
                    XmlOwner::Slide => presentationml && local_name.as_ref() == b"sld",
                };
                let element_is_mce = is_mce_namespace(&resolved);
                let is_slide_id = matches!(owner, XmlOwner::Presentation)
                    && presentationml
                    && local_name.as_ref() == b"sldId"
                    && is_direct_slide_owner(&stack);
                let is_direct_slide_list = matches!(owner, XmlOwner::Presentation)
                    && presentationml
                    && local_name.as_ref() == b"sldIdLst"
                    && is_direct_slide_list_owner(&stack);
                if is_direct_slide_list {
                    direct_slide_lists = direct_slide_lists.saturating_add(1);
                    if direct_slide_lists > 1 {
                        return Err(XmlInspectionError::Malformed);
                    }
                }
                if is_direct_slide_owner(&stack)
                    && !(presentationml && local_name.as_ref() == b"sldId")
                {
                    return Err(XmlInspectionError::Malformed);
                }
                let attributes = validate_attributes(
                    &element,
                    reader.decoder(),
                    reader.resolver(),
                    is_slide_id,
                )?;
                if !root_seen {
                    root_seen = true;
                    root_is_expected = is_expected_root;
                }
                mce |= element_is_mce || attributes.mce;
                if is_slide_id {
                    let numeric_id = attributes.numeric_id.ok_or(XmlInspectionError::Malformed)?;
                    let relationship_id = attributes
                        .relationship_id
                        .ok_or(XmlInspectionError::Malformed)?;
                    if slides.len() >= limits.max_slides() {
                        return Err(XmlInspectionError::Limit);
                    }
                    slides
                        .try_reserve(1)
                        .map_err(|_| XmlInspectionError::Allocation {
                            resource: "PPTX ordered slide references",
                        })?;
                    slides.push(SlideReferenceObservation {
                        numeric_id,
                        relationship_id,
                    });
                }
            },
            Event::End(element) => {
                if has_unresolved_prefix(&resolved, element.name()) {
                    return Err(XmlInspectionError::Malformed);
                }
                let Some(start) = stack.pop() else {
                    return Err(XmlInspectionError::Malformed);
                };
                if start.name.as_slice() != element.name().as_ref() {
                    return Err(XmlInspectionError::Malformed);
                }
            },
            Event::DocType(_) | Event::PI(_) => return Err(XmlInspectionError::Malformed),
            Event::Text(text) => {
                validate_xml_text(&text)?;
                if stack.is_empty() && !text.as_ref().iter().all(u8::is_ascii_whitespace) {
                    return Err(XmlInspectionError::Malformed);
                }
                if is_direct_slide_owner(stack.as_slice())
                    && !text.as_ref().iter().all(u8::is_ascii_whitespace)
                {
                    return Err(XmlInspectionError::Malformed);
                }
            },
            Event::CData(text) => {
                validate_xml_text(&text)?;
                if stack.is_empty() || is_direct_slide_owner(stack.as_slice()) {
                    return Err(XmlInspectionError::Malformed);
                }
            },
            Event::GeneralRef(reference) => {
                if !valid_general_ref(reference.as_ref()) || stack.is_empty() {
                    return Err(XmlInspectionError::Malformed);
                }
            },
            Event::Comment(comment) => {
                if !valid_xml_comment(comment.as_ref()) {
                    return Err(XmlInspectionError::Malformed);
                }
            },
            Event::Decl(_) if root_seen => return Err(XmlInspectionError::Malformed),
            Event::Decl(_) => {},
            Event::Eof => break,
        }
    }
    if !root_seen || !root_is_expected || !stack.is_empty() {
        return Err(XmlInspectionError::Malformed);
    }
    Ok(XmlObservation { slides, mce })
}

fn is_direct_slide_owner(stack: &[XmlElementFrame]) -> bool {
    stack.len() == 2
        && stack[0].presentationml
        && stack[0].local_name.as_slice() == b"presentation"
        && stack[1].presentationml
        && stack[1].local_name.as_slice() == b"sldIdLst"
}

fn is_direct_slide_list_owner(stack: &[XmlElementFrame]) -> bool {
    stack.len() == 1 && stack[0].presentationml && stack[0].local_name.as_slice() == b"presentation"
}

fn validate_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
    direct_slide_owner: bool,
) -> Result<XmlAttributeObservation, XmlInspectionError> {
    let mut observation = XmlAttributeObservation {
        mce: false,
        numeric_id: None,
        relationship_id: None,
    };
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|_| XmlInspectionError::Malformed)?;
        validate_xml_text(attribute.value.as_ref())?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|_| XmlInspectionError::Malformed)?;
        validate_xml_text(value.as_bytes())?;
        if attribute.key.as_namespace_binding().is_some() {
            continue;
        }
        let (resolved, local) = resolver.resolve_attribute(attribute.key);
        if has_unresolved_prefix(&resolved, attribute.key) {
            return Err(XmlInspectionError::Malformed);
        }
        observation.mce |= is_mce_namespace(&resolved);
        if !direct_slide_owner || local.as_ref() != b"id" {
            continue;
        }
        if matches!(resolved, ResolveResult::Unbound) {
            if observation.numeric_id.is_some() {
                return Err(XmlInspectionError::Malformed);
            }
            let numeric_id =
                parse_numeric_id(value.as_ref()).ok_or(XmlInspectionError::Malformed)?;
            observation.numeric_id = Some(numeric_id);
        } else if is_relationship_namespace(&resolved) {
            if observation.relationship_id.is_some() || value.is_empty() {
                return Err(XmlInspectionError::Malformed);
            }
            let mut relationship_id = String::new();
            relationship_id.try_reserve(value.len()).map_err(|_| {
                XmlInspectionError::Allocation {
                    resource: "PPTX slide relationship identifier",
                }
            })?;
            relationship_id.push_str(value.as_ref());
            observation.relationship_id = Some(relationship_id);
        }
    }
    Ok(observation)
}

fn copy_xml_name(value: &[u8], resource: &'static str) -> Result<Vec<u8>, XmlInspectionError> {
    let mut copied = Vec::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|_| XmlInspectionError::Allocation { resource })?;
    copied.extend_from_slice(value);
    Ok(copied)
}

fn parse_numeric_id(value: &str) -> Option<u64> {
    if value.is_empty() || !value.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    let value = value.parse().ok()?;
    (MIN_SLIDE_ID..=MAX_SLIDE_ID)
        .contains(&value)
        .then_some(value)
}

fn record_owned_text(
    seen: &mut std::collections::HashSet<String>,
    value: &str,
    resource: &'static str,
) -> Result<bool, PptxValidationError> {
    if seen.contains(value) {
        return Ok(false);
    }
    let mut owned = String::new();
    owned
        .try_reserve(value.len())
        .map_err(|_| PptxValidationError::Allocation { resource })?;
    owned.push_str(value);
    seen.try_reserve(1)
        .map_err(|_| PptxValidationError::Allocation { resource })?;
    seen.insert(owned);
    Ok(true)
}

fn note_mce(facts: &mut Facts) {
    facts.mce = match facts.mce {
        MceOutcome::NotApplicable => MceOutcome::Present(1),
        MceOutcome::Present(count) => MceOutcome::Present(count.saturating_add(1)),
        MceOutcome::Blocked | MceOutcome::StoppedByRoot => facts.mce,
    };
}

fn block_mce(facts: &mut Facts) {
    if !matches!(facts.mce, MceOutcome::StoppedByRoot) {
        facts.mce = MceOutcome::Blocked;
    }
}

fn is_presentationml_name_strict(
    namespace: &ResolveResult<'_>,
    name: quick_xml::name::QName<'_>,
    local_name: &[u8],
) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value))
                if *value == namespace::PRESENTATIONML_NAMESPACE
                    || *value == namespace::STRICT_PRESENTATIONML_NAMESPACE
        )
}

fn is_relationship_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == litchi_ooxml_common::relationships::TRANSITIONAL_NAMESPACE
                || *value == litchi_ooxml_common::relationships::STRICT_NAMESPACE
    )
}

fn has_unresolved_prefix(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    matches!(namespace, ResolveResult::Unknown(_))
        || (matches!(namespace, ResolveResult::Unbound) && name.prefix().is_some())
}

fn validate_xml_text(text: &[u8]) -> Result<(), XmlInspectionError> {
    let value = std::str::from_utf8(text).map_err(|_| XmlInspectionError::Malformed)?;
    if value.chars().all(is_xml_char) {
        Ok(())
    } else {
        Err(XmlInspectionError::Malformed)
    }
}

fn valid_xml_comment(value: &[u8]) -> bool {
    !value.windows(2).any(|pair| pair == b"--")
        && value.last() != Some(&b'-')
        && std::str::from_utf8(value)
            .map(|text| text.chars().all(is_xml_char))
            .unwrap_or(false)
}

fn is_xml_char(value: char) -> bool {
    matches!(
        value as u32,
        0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
    )
}

fn valid_general_ref(value: &[u8]) -> bool {
    let Ok(value) = std::str::from_utf8(value) else {
        return false;
    };
    if resolve_xml_entity(value).is_some() {
        return true;
    }
    if let Some(hex) = value
        .strip_prefix("#x")
        .or_else(|| value.strip_prefix("#X"))
    {
        return parse_numeric_reference(hex, 16);
    }
    value
        .strip_prefix('#')
        .is_some_and(|decimal| parse_numeric_reference(decimal, 10))
}

fn parse_numeric_reference(value: &str, radix: u32) -> bool {
    if value.is_empty() {
        return false;
    }
    let mut codepoint = 0_u32;
    for digit in value.chars() {
        let Some(digit) = digit.to_digit(radix) else {
            return false;
        };
        let Some(next) = codepoint
            .checked_mul(radix)
            .and_then(|value| value.checked_add(digit))
        else {
            return false;
        };
        codepoint = next;
    }
    is_xml_char(char::from_u32(codepoint).unwrap_or('\0'))
}

fn is_mce_namespace(namespace: &ResolveResult<'_>) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == MCE_NAMESPACE)
}

struct CatalogFacts {
    external_targets: u64,
    signatures: u64,
    macros: u64,
}

fn inspect_catalog_and_graph(
    package: &SourceBackedPackage,
    max_nodes: usize,
) -> Result<(CatalogFacts, GraphFacts), PptxValidationError> {
    let mut catalog = CatalogFacts {
        external_targets: 0,
        signatures: 0,
        macros: 0,
    };
    let mut seen = std::collections::HashSet::<String>::new();
    seen.try_reserve(1)
        .map_err(|_| PptxValidationError::Allocation {
            resource: "PPTX relationship graph set",
        })?;
    let mut missing_targets = 0_u64;
    let mut invalid_targets = 0_u64;
    let mut blocked = false;

    for relationship in package.rels().iter() {
        inspect_catalog_relationship(&mut catalog, relationship);
        if !blocked && !relationship.is_external() {
            blocked = inspect_graph_relationship(
                package,
                relationship,
                &mut seen,
                max_nodes,
                &mut missing_targets,
                &mut invalid_targets,
            )?;
        }
    }

    for part in package.iter_parts() {
        catalog.signatures = catalog.signatures.saturating_add(u64::from(
            is_signature_path(part.partname().as_str())
                || is_signature_content_type(part.content_type()),
        ));
        catalog.macros = catalog.macros.saturating_add(u64::from(
            is_macro_path(part.partname().as_str()) || is_macro_content_type(part.content_type()),
        ));

        if !blocked {
            if record_graph_node(&mut seen, part.partname().as_str(), max_nodes)? {
                blocked = true;
            }
        }
        for relationship in part.rels().iter() {
            inspect_catalog_relationship(&mut catalog, relationship);
            if !blocked && !relationship.is_external() {
                blocked = inspect_graph_relationship(
                    package,
                    relationship,
                    &mut seen,
                    max_nodes,
                    &mut missing_targets,
                    &mut invalid_targets,
                )?;
            }
        }
    }

    Ok((
        catalog,
        GraphFacts {
            blocked,
            missing_targets,
            invalid_targets,
        },
    ))
}

fn inspect_catalog_relationship(
    catalog: &mut CatalogFacts,
    relationship: &litchi_opc::Relationship,
) {
    catalog.external_targets = catalog
        .external_targets
        .saturating_add(u64::from(relationship.is_external()));
    catalog.signatures = catalog
        .signatures
        .saturating_add(u64::from(is_signature_relationship(relationship.reltype())));
    catalog.macros = catalog
        .macros
        .saturating_add(u64::from(is_macro_relationship(relationship.reltype())));
}

struct GraphFacts {
    blocked: bool,
    missing_targets: u64,
    invalid_targets: u64,
}

fn inspect_graph_relationship(
    package: &SourceBackedPackage,
    relationship: &litchi_opc::Relationship,
    seen: &mut std::collections::HashSet<String>,
    max_nodes: usize,
    missing_targets: &mut u64,
    invalid_targets: &mut u64,
) -> Result<bool, PptxValidationError> {
    let target = match relationship.target_partname() {
        Ok(target) => target,
        Err(_) => {
            *invalid_targets = invalid_targets.saturating_add(1);
            return Ok(false);
        },
    };
    match package.part(&target) {
        Ok(_) => record_graph_node(seen, target.as_str(), max_nodes),
        Err(error) if is_transient(&error) => Err(PptxValidationError::Ingress(error)),
        Err(_) => {
            *missing_targets = missing_targets.saturating_add(1);
            Ok(false)
        },
    }
}

fn record_graph_node(
    seen: &mut std::collections::HashSet<String>,
    uri: &str,
    max_nodes: usize,
) -> Result<bool, PptxValidationError> {
    if seen.contains(uri) {
        return Ok(false);
    }
    if seen.len() >= max_nodes {
        return Ok(true);
    }
    let mut owned = String::new();
    owned
        .try_reserve(uri.len())
        .map_err(|_| PptxValidationError::Allocation {
            resource: "PPTX relationship graph set",
        })?;
    owned.push_str(uri);
    seen.try_reserve(1)
        .map_err(|_| PptxValidationError::Allocation {
            resource: "PPTX relationship graph set",
        })?;
    seen.insert(owned);
    Ok(false)
}

struct Facts {
    source_version: SourceVersion,
    root: RootOutcome,
    slides: SlideOutcome,
    graph: GraphOutcome,
    root_issue: bool,
    root_error: Option<RootErrorKind>,
    graph_missing_targets: u64,
    graph_invalid_targets: u64,
    missing_slide_relationships: u64,
    invalid_slide_relationships: u64,
    missing_slide_parts: u64,
    invalid_slide_parts: u64,
    malformed_slides: u64,
    duplicate_slides: u64,
    duplicate_slide_ids: u64,
    duplicate_relationship_ids: u64,
    mce: MceOutcome,
    owner_bytes: usize,
    issue_capacity: usize,
}

#[derive(Clone, Copy)]
enum RootOutcome {
    Complete,
    Malformed,
    Blocked,
}

#[derive(Clone, Copy)]
enum SlideOutcome {
    Complete,
    NotApplicable,
    Blocked,
    StoppedByRoot,
}

#[derive(Clone, Copy)]
enum GraphOutcome {
    Complete,
    Blocked,
}

#[derive(Clone, Copy)]
enum MceOutcome {
    NotApplicable,
    Present(u64),
    Blocked,
    StoppedByRoot,
}

#[derive(Clone, Copy)]
enum RootErrorKind {
    Missing,
    WrongContentType,
    Unreadable,
    Malformed,
}

impl Facts {
    fn new(
        limits: ValidationLimits,
        source_version: SourceVersion,
    ) -> Result<Self, PptxValidationError> {
        Ok(Self {
            source_version,
            root: RootOutcome::Blocked,
            slides: SlideOutcome::Blocked,
            graph: GraphOutcome::Blocked,
            root_issue: false,
            root_error: None,
            graph_missing_targets: 0,
            graph_invalid_targets: 0,
            missing_slide_relationships: 0,
            invalid_slide_relationships: 0,
            missing_slide_parts: 0,
            invalid_slide_parts: 0,
            malformed_slides: 0,
            duplicate_slides: 0,
            duplicate_slide_ids: 0,
            duplicate_relationship_ids: 0,
            mce: MceOutcome::NotApplicable,
            owner_bytes: 0,
            issue_capacity: limits.max_issues(),
        })
    }

    fn issue_capacity_check(&self, observed: usize) -> Result<(), PptxValidationError> {
        if observed >= self.issue_capacity {
            return Err(PptxValidationError::Report(ValidationReportError::Limit {
                kind: ValidationLimitKind::Issues,
                observed: observed.saturating_add(1),
                limit: self.issue_capacity,
            }));
        }
        Ok(())
    }
}

fn finish(
    package: &SourceBackedPackage,
    _input_limits: PptxValidationLimits,
    limits: ValidationLimits,
    facts: Facts,
    catalog: CatalogFacts,
    _root: Option<litchi_opc::PartView<'_>>,
) -> Result<ValidateReport, PptxValidationError> {
    let current_version = package
        .source_version()
        .map_err(PptxValidationError::Ingress)?;
    if current_version != facts.source_version {
        return Err(PptxValidationError::Ingress(OpcError::SourceChanged {
            expected: facts.source_version,
            actual: current_version,
        }));
    }
    let mut issues = Vec::new();
    issues
        .try_reserve_exact(limits.max_issues().min(16))
        .map_err(|_| PptxValidationError::Allocation {
            resource: "PPTX validation issue staging",
        })?;

    if facts.root_issue {
        facts.issue_capacity_check(issues.len())?;
        let (code, message) = match facts.root_error.unwrap_or(RootErrorKind::Malformed) {
            RootErrorKind::Missing => (
                "pptx.presentation.root.missing",
                "The package does not expose one valid PresentationML main part.",
            ),
            RootErrorKind::WrongContentType => (
                "pptx.presentation.root.content_type",
                "The package main part is not a supported PresentationML content type.",
            ),
            RootErrorKind::Unreadable => (
                "pptx.presentation.root.unreadable",
                "The PresentationML main-part payload could not be inspected.",
            ),
            RootErrorKind::Malformed => (
                "pptx.presentation.root.malformed",
                "The PresentationML main part does not have one bounded, well-formed presentation root.",
            ),
        };
        issues.push(issue(
            ROOT,
            code,
            IssueSeverity::Error,
            message,
            "presentation-root",
            None,
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    } else if matches!(facts.root, RootOutcome::Blocked) {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            ROOT,
            "pptx.presentation.root.limit",
            IssueSeverity::Error,
            "The PresentationML root exceeded a configured finite validation ceiling.",
            "presentation-root",
            Some(EvidenceValue::Size(facts.owner_bytes as u64)),
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }

    if facts.graph_missing_targets != 0 || facts.graph_invalid_targets != 0 {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            GRAPH,
            "pptx.relationship_graph.incomplete",
            IssueSeverity::Error,
            "The loaded OPC relationship manifests contain unresolved internal targets.",
            "relationship-graph",
            Some(EvidenceValue::Count(
                facts
                    .graph_missing_targets
                    .saturating_add(facts.graph_invalid_targets),
            )),
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }
    if facts.slides_is_issue() {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            SLIDES,
            "pptx.presentation.slide_closure.incomplete",
            IssueSeverity::Error,
            "The ordered slide closure contains unresolved or malformed slide entries.",
            "ordered-slide-closure",
            Some(EvidenceValue::Count(
                facts
                    .missing_slide_relationships
                    .saturating_add(facts.invalid_slide_relationships)
                    .saturating_add(facts.missing_slide_parts)
                    .saturating_add(facts.invalid_slide_parts)
                    .saturating_add(facts.malformed_slides)
                    .saturating_add(facts.duplicate_slides)
                    .saturating_add(facts.duplicate_slide_ids)
                    .saturating_add(facts.duplicate_relationship_ids),
            )),
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }
    if catalog.external_targets != 0 {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            EXTERNAL,
            "pptx.external_target.present",
            IssueSeverity::Warning,
            "External relationship targets are present; no target was fetched.",
            "external-relationships",
            Some(EvidenceValue::Count(catalog.external_targets)),
            CompatibilityImpact::Security,
            limits,
        )?);
    }
    if catalog.signatures != 0 {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            SIGNATURES,
            "pptx.signature.infrastructure_present",
            IssueSeverity::Info,
            "Digital-signature infrastructure is present; cryptographic validity was not checked.",
            "signature-infrastructure",
            Some(EvidenceValue::Count(catalog.signatures)),
            CompatibilityImpact::None,
            limits,
        )?);
    }
    if catalog.macros != 0 {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            MACROS,
            "pptx.macro.storage_present",
            IssueSeverity::Warning,
            "VBA or macro-enabled storage is present; macro behavior was not executed or assessed.",
            "macro-storage",
            Some(EvidenceValue::Count(catalog.macros)),
            CompatibilityImpact::Security,
            limits,
        )?);
    }
    let mce_count = match facts.mce {
        MceOutcome::Present(count) => count,
        _ => 0,
    };
    if mce_count != 0 {
        facts.issue_capacity_check(issues.len())?;
        issues.push(issue(
            MCE,
            "pptx.mce.present",
            IssueSeverity::Warning,
            "Markup-compatibility input is present; no branch was selected by this report.",
            "markup-compatibility",
            Some(EvidenceValue::Count(mce_count)),
            CompatibilityImpact::Interoperability,
            limits,
        )?);
    }

    let ingress = CheckStatus::Complete;
    let catalog_status = CheckStatus::Complete;
    let root_status = match facts.root {
        RootOutcome::Complete | RootOutcome::Malformed => CheckStatus::Complete,
        RootOutcome::Blocked => CheckStatus::blocked(
            "PresentationML root inspection reached a configured finite ceiling",
            limits,
        )?,
    };
    let root_id = id(ROOT, limits)?;
    let slides_status = match facts.slides {
        SlideOutcome::Complete => CheckStatus::Complete,
        SlideOutcome::NotApplicable => CheckStatus::NotApplicable,
        SlideOutcome::Blocked => CheckStatus::blocked(
            "ordered slide closure reached a configured finite validation ceiling",
            limits,
        )?,
        SlideOutcome::StoppedByRoot => CheckStatus::stopped_by(root_id.clone()),
    };
    let graph_status = match facts.graph {
        GraphOutcome::Complete => CheckStatus::Complete,
        GraphOutcome::Blocked => CheckStatus::blocked(
            "the loaded relationship graph exceeded the configured node ceiling",
            limits,
        )?,
    };
    let presence_status = |count: u64| {
        if count == 0 {
            CheckStatus::NotApplicable
        } else {
            CheckStatus::Complete
        }
    };
    let checks = [
        check(INGRESS, ingress, limits)?,
        check(CATALOG, catalog_status, limits)?,
        check(GRAPH, graph_status, limits)?,
        check(ROOT, root_status, limits)?,
        check(SLIDES, slides_status, limits)?,
        check(EXTERNAL, presence_status(catalog.external_targets), limits)?,
        check(SIGNATURES, presence_status(catalog.signatures), limits)?,
        check(MACROS, presence_status(catalog.macros), limits)?,
        check(
            MCE,
            match facts.mce {
                MceOutcome::NotApplicable => CheckStatus::NotApplicable,
                MceOutcome::Present(_) => CheckStatus::Complete,
                MceOutcome::Blocked => CheckStatus::blocked(
                    "markup-compatibility inspection reached a configured finite validation ceiling or malformed XML",
                    limits,
                )?,
                MceOutcome::StoppedByRoot => CheckStatus::stopped_by(root_id.clone()),
            },
            limits,
        )?,
    ];
    let report =
        ValidateReport::try_new(checks, issues, limits).map_err(PptxValidationError::from)?;
    let current_version = package
        .source_version()
        .map_err(PptxValidationError::Ingress)?;
    if current_version != facts.source_version {
        return Err(PptxValidationError::Ingress(OpcError::SourceChanged {
            expected: facts.source_version,
            actual: current_version,
        }));
    }
    Ok(report)
}

impl Facts {
    fn slides_is_issue(&self) -> bool {
        self.missing_slide_relationships != 0
            || self.invalid_slide_relationships != 0
            || self.missing_slide_parts != 0
            || self.invalid_slide_parts != 0
            || self.malformed_slides != 0
            || self.duplicate_slides != 0
            || self.duplicate_slide_ids != 0
            || self.duplicate_relationship_ids != 0
    }
}

fn issue(
    check_name: &str,
    code: &str,
    severity: IssueSeverity,
    message: &str,
    path: &str,
    evidence: Option<EvidenceValue>,
    compatibility: CompatibilityImpact,
    limits: ValidationLimits,
) -> Result<ValidationIssue, PptxValidationError> {
    let evidence = evidence
        .map(|value| IssueEvidence::try_new("observations", value, limits))
        .transpose()?;
    ValidationIssue::try_new(
        id(check_name, limits)?,
        code,
        severity,
        message,
        [IssueLocation::try_new(
            None,
            Some(path),
            None,
            None,
            None,
            limits,
        )?],
        evidence,
        None,
        compatibility,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn rejected_ingress_report(
    error: &OpcError,
    limits: ValidationLimits,
) -> Result<ValidateReport, PptxValidationError> {
    let ingress = id(INGRESS, limits)?;
    let issue = ValidationIssue::try_new(
        ingress.clone(),
        "pptx.package.invalid",
        IssueSeverity::Fatal,
        "The input is not a structurally readable OPC package for PresentationML validation.",
        [IssueLocation::try_new(
            None,
            Some("package-ingress"),
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            "diagnostic_sha256",
            EvidenceValue::Sha256(EvidenceDigest::of(error.to_string().as_bytes())),
            limits,
        )?],
        None,
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )?;
    let blocked = CheckStatus::blocked(
        "OPC ingress was structurally rejected before the PresentationML catalog could run",
        limits,
    )?;
    let catalog = id(CATALOG, limits)?;
    let graph = CheckStatus::stopped_by(catalog.clone());
    let root = CheckStatus::stopped_by(catalog.clone());
    let slides = CheckStatus::stopped_by(id(ROOT, limits)?);
    ValidateReport::try_new(
        [
            ValidationCheck::new(ingress, CheckStatus::Complete),
            ValidationCheck::new(catalog, blocked),
            check(GRAPH, graph, limits)?,
            check(ROOT, root, limits)?,
            check(SLIDES, slides, limits)?,
            check(
                EXTERNAL,
                CheckStatus::stopped_by(id(CATALOG, limits)?),
                limits,
            )?,
            check(
                SIGNATURES,
                CheckStatus::stopped_by(id(CATALOG, limits)?),
                limits,
            )?,
            check(
                MACROS,
                CheckStatus::stopped_by(id(CATALOG, limits)?),
                limits,
            )?,
            check(MCE, CheckStatus::stopped_by(id(ROOT, limits)?), limits)?,
        ],
        [issue],
        limits,
    )
    .map_err(Into::into)
}

fn blocked_report(
    reason: &str,
    limits: ValidationLimits,
) -> Result<ValidateReport, PptxValidationError> {
    let ingress = id(INGRESS, limits)?;
    let blocked = CheckStatus::blocked(reason, limits)?;
    let dependent = CheckStatus::stopped_by(ingress.clone());
    ValidateReport::try_new(
        [
            ValidationCheck::new(ingress.clone(), blocked),
            check(CATALOG, dependent.clone(), limits)?,
            check(GRAPH, dependent.clone(), limits)?,
            check(ROOT, dependent.clone(), limits)?,
            check(SLIDES, dependent.clone(), limits)?,
            check(EXTERNAL, dependent.clone(), limits)?,
            check(SIGNATURES, dependent.clone(), limits)?,
            check(MACROS, dependent.clone(), limits)?,
            check(MCE, dependent, limits)?,
        ],
        [],
        limits,
    )
    .map_err(Into::into)
}

fn is_presentation_main_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::PML_PRESENTATION_MAIN
            | ct::PML_SLIDESHOW_MAIN
            | ct::PML_TEMPLATE_MAIN
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}

fn is_signature_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn is_signature_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::OPC_DIGITAL_SIGNATURE_ORIGIN
            | ct::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
            | ct::OPC_DIGITAL_SIGNATURE_CERTIFICATE
    )
}

fn is_signature_path(value: &str) -> bool {
    value
        .split('/')
        .any(|segment| segment.eq_ignore_ascii_case("_xmlsignatures"))
}

fn is_macro_relationship(value: &str) -> bool {
    matches!(
        value,
        rt::VBA_PROJECT | rt::VBA_PROJECT_SIGNATURE | rt::VBA_PROJECT_SIGNATURE_AGILE
    ) || value.ends_with("/vbaProject")
        || value.ends_with("/vbaProjectSignature")
        || value.ends_with("/vbaProjectSignatureAgile")
}

fn is_macro_content_type(value: &str) -> bool {
    matches!(
        value,
        ct::OFC_VBA_PROJECT
            | ct::OFC_VBA_PROJECT_SIGNATURE
            | ct::OFC_VBA_PROJECT_SIGNATURE_AGILE
            | ct::PML_PRES_MACRO_MAIN
            | ct::PML_SLIDESHOW_MACRO_MAIN
            | ct::PML_TEMPLATE_MACRO_MAIN
    )
}

fn is_macro_path(value: &str) -> bool {
    value.split('/').any(|segment| {
        segment.eq_ignore_ascii_case("vbaProject.bin")
            || segment.eq_ignore_ascii_case("vbaProjectSignature.bin")
            || segment.eq_ignore_ascii_case("vbaProjectSignatureAgile.bin")
    })
}

fn is_transient(error: &OpcError) -> bool {
    matches!(
        error,
        OpcError::IoError(_)
            | OpcError::SourceChanged { .. }
            | OpcError::Allocation { .. }
            | OpcError::CollectionAllocation { .. }
            | OpcError::Cancelled
            | OpcError::Execution(_)
            | OpcError::ParallelRead(_)
            | OpcError::Committed { .. }
            | OpcError::IncompleteOutput { .. }
            | OpcError::OperationAccountingOverflow { .. }
    )
}

fn is_structural_rejection(error: &OpcError) -> bool {
    !is_transient(error)
        && !matches!(
            error,
            OpcError::ManagedPartDataArcEscape
                | OpcError::SignedSourceRequiresExplicitPolicy
                | OpcError::SourceBackedOverlayUnavailable { .. }
                | OpcError::PreservationUnavailable { .. }
                | OpcError::UnsupportedExecutionAffinity
        )
}

fn id(value: &str, limits: ValidationLimits) -> Result<CheckCapabilityId, PptxValidationError> {
    CheckCapabilityId::try_new(value, limits).map_err(Into::into)
}

fn check(
    value: &str,
    status: CheckStatus,
    limits: ValidationLimits,
) -> Result<ValidationCheck, PptxValidationError> {
    Ok(ValidationCheck::new(id(value, limits)?, status))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn accounting_overflow_is_transient_and_not_structural() {
        let error = OpcError::OperationAccountingOverflow {
            counter: "output_bytes_accepted",
        };

        assert!(is_transient(&error));
        assert!(!is_structural_rejection(&error));
    }
}
