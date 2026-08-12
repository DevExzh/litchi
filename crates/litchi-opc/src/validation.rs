//! Bounded, read-only validation of an OPC package ingress.
//!
//! Validation reuses the source-backed package catalog. Ordinary part
//! payloads stay in the positional source: only ZIP metadata,
//! `[Content_Types].xml`, and relationship manifests reachable through the
//! package graph are read. Unowned physical `.rels` members and ordinary part
//! payloads are outside this report. It does not validate application XML,
//! execute macros, fetch external targets, verify signatures, or offer repairs.

use std::io;
use std::sync::{Arc, Mutex};

use litchi_core::{
    CheckCapabilityId, CheckStatus, CompatibilityImpact, EvidenceValue, IssueEvidence,
    IssueLocation, IssueSeverity, ReadAt, RepairAvailability, SpecCitation, ValidateReport,
    ValidationCheck, ValidationIssue, ValidationLimits,
};

use crate::constants::{content_type, relationship_type};
use crate::{OpcError, ReadLimits, ReadResource, Result, SourceBackedPackage};

const INGRESS: &str = "opc.package.ingress";
const CATALOG: &str = "opc.package.catalog";
const REACHABLE_RELATIONSHIPS: &str = "opc.package.reachable_relationship_graph";
const SIGNATURE_PRESENCE: &str = "opc.package.signature_presence";

/// Validate an immutable positional OPC source under the default finite read
/// and report limits.
///
/// Ordinary part payloads are not materialized. Structural package rejection
/// is returned as a complete ingress check plus a deterministic issue. A
/// resource ceiling is a blocked check. I/O, source mutation, cancellation,
/// allocation, invalid policy, and report-construction failures remain errors.
pub fn validate_read_at(source: Arc<dyn ReadAt>) -> Result<ValidateReport> {
    validate_read_at_with_limits(source, ReadLimits::default(), ValidationLimits::default())
}

/// Validate an immutable positional OPC source with explicit finite policies.
///
/// `read_limits` bounds ZIP indexing, manifest parsing, relationship graph
/// traversal, and all bytes read by OPC ingress. `report_limits` bounds the
/// retained diagnostic value. No operation mutates the source.
pub fn validate_read_at_with_limits(
    source: Arc<dyn ReadAt>,
    read_limits: ReadLimits,
    report_limits: ValidationLimits,
) -> Result<ValidateReport> {
    let expected = source.version()?;
    let tracked = Arc::new(ValidationSource::new(Arc::clone(&source)));
    let ingress_source: Arc<dyn ReadAt> = tracked.clone();
    let ingress = SourceBackedPackage::from_read_at_with_limits(ingress_source, read_limits);
    if let Some(error) = tracked.take_error() {
        return Err(OpcError::IoError(error));
    }
    let report = match ingress {
        Ok(package) => successful_report(&package, report_limits),
        Err(OpcError::ReadLimit { resource, .. }) => blocked_report(resource, report_limits),
        Err(error) if is_structural_rejection(&error) => structural_report(&error, report_limits),
        Err(error) => Err(error),
    }?;
    let actual = source.version()?;
    if actual != expected {
        return Err(OpcError::SourceChanged { expected, actual });
    }
    Ok(report)
}

struct ValidationSource {
    source: Arc<dyn ReadAt>,
    first_io_error: Mutex<Option<io::Error>>,
}

impl ValidationSource {
    fn new(source: Arc<dyn ReadAt>) -> Self {
        Self {
            source,
            first_io_error: Mutex::new(None),
        }
    }

    fn record(&self, error: io::Error) -> io::Error {
        let forwarded = io::Error::new(error.kind(), error.to_string());
        let mut first = self
            .first_io_error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if first.is_none() {
            *first = Some(error);
        }
        forwarded
    }

    fn take_error(&self) -> Option<io::Error> {
        self.first_io_error
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .take()
    }
}

impl ReadAt for ValidationSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len().map_err(|error| self.record(error))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.source
            .read_at(offset, output)
            .map_err(|error| self.record(error))
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        self.source.version()
    }
}

fn successful_report(
    package: &SourceBackedPackage,
    limits: ValidationLimits,
) -> Result<ValidateReport> {
    // The source-backed open brackets every mandatory read with the captured
    // version. Recheck after observing the catalog so a caller never receives
    // a report for a source changed between open and report construction.
    let _version = package.source_version()?;
    let signature_count = signature_infrastructure_count(package);
    let signature_status = if signature_count == 0 {
        CheckStatus::NotApplicable
    } else {
        CheckStatus::Complete
    };
    let checks = vec![
        check(INGRESS, CheckStatus::Complete, limits)?,
        check(CATALOG, CheckStatus::Complete, limits)?,
        check(REACHABLE_RELATIONSHIPS, CheckStatus::Complete, limits)?,
        check(SIGNATURE_PRESENCE, signature_status, limits)?,
    ];
    let issues = if signature_count == 0 {
        Vec::new()
    } else {
        vec![signature_presence_issue(signature_count, limits)?]
    };
    Ok(ValidateReport::try_new(checks, issues, limits)?)
}

fn structural_report(error: &OpcError, limits: ValidationLimits) -> Result<ValidateReport> {
    let phase = structural_phase(error);
    let checks = match phase {
        StructuralPhase::Ingress => vec![
            check(INGRESS, CheckStatus::Complete, limits)?,
            check(
                CATALOG,
                CheckStatus::blocked("ZIP ingress was structurally rejected", limits)?,
                limits,
            )?,
            check(
                REACHABLE_RELATIONSHIPS,
                CheckStatus::blocked("ZIP ingress was structurally rejected", limits)?,
                limits,
            )?,
            check(
                SIGNATURE_PRESENCE,
                CheckStatus::blocked("ZIP ingress was structurally rejected", limits)?,
                limits,
            )?,
        ],
        StructuralPhase::Catalog => vec![
            check(INGRESS, CheckStatus::Complete, limits)?,
            check(CATALOG, CheckStatus::Complete, limits)?,
            check(
                REACHABLE_RELATIONSHIPS,
                CheckStatus::blocked(
                    "catalog rejection prevented relationship-graph completion",
                    limits,
                )?,
                limits,
            )?,
            check(
                SIGNATURE_PRESENCE,
                CheckStatus::blocked(
                    "catalog rejection prevented signature-presence cataloging",
                    limits,
                )?,
                limits,
            )?,
        ],
        StructuralPhase::Relationships => vec![
            check(INGRESS, CheckStatus::Complete, limits)?,
            check(
                CATALOG,
                CheckStatus::blocked(
                    "relationship rejection prevented catalog completion",
                    limits,
                )?,
                limits,
            )?,
            check(REACHABLE_RELATIONSHIPS, CheckStatus::Complete, limits)?,
            check(
                SIGNATURE_PRESENCE,
                CheckStatus::blocked(
                    "relationship rejection prevented signature-presence cataloging",
                    limits,
                )?,
                limits,
            )?,
        ],
    };
    let issue = structural_issue(error, phase, limits)?;
    Ok(ValidateReport::try_new(checks, [issue], limits)?)
}

fn blocked_report(resource: ReadResource, limits: ValidationLimits) -> Result<ValidateReport> {
    let (blocked, reason) = match resource {
        ReadResource::InputBytes
        | ReadResource::ArchiveMembers
        | ReadResource::ArchiveMemberNameBytes
        | ReadResource::ArchiveMetadataBytes
        | ReadResource::ArchiveCompressedBytes
        | ReadResource::ArchiveEntryBytes
        | ReadResource::ArchiveTotalBytes => (INGRESS, "OPC ZIP ingress resource ceiling reached"),
        ReadResource::ContentTypesBytes
        | ReadResource::ContentTypeMappings
        | ReadResource::Parts
        | ReadResource::PartBytes
        | ReadResource::TotalPartBytes => (CATALOG, "OPC catalog resource ceiling reached"),
        ReadResource::RelationshipParts
        | ReadResource::RelationshipXmlBytes
        | ReadResource::TotalRelationshipXmlBytes
        | ReadResource::RelationshipsPerPart
        | ReadResource::TotalRelationships
        | ReadResource::RelationshipGraphNodes
        | ReadResource::XmlEvents
        | ReadResource::TotalRelationshipXmlEvents
        | ReadResource::XmlDepth
        | ReadResource::XmlAttributeBytes
        | ReadResource::RelationshipTargetBytes => (
            REACHABLE_RELATIONSHIPS,
            "OPC reachable-relationship resource ceiling reached",
        ),
    };

    let ingress = if blocked == INGRESS {
        CheckStatus::blocked(reason, limits)?
    } else {
        CheckStatus::Complete
    };
    let catalog = if blocked == CATALOG || blocked == REACHABLE_RELATIONSHIPS {
        CheckStatus::blocked(reason, limits)?
    } else if blocked == INGRESS {
        CheckStatus::stopped_by(id(INGRESS, limits)?)
    } else {
        CheckStatus::Complete
    };
    let relationships = if blocked == REACHABLE_RELATIONSHIPS {
        CheckStatus::blocked(reason, limits)?
    } else if blocked == INGRESS {
        CheckStatus::stopped_by(id(INGRESS, limits)?)
    } else if blocked == CATALOG {
        CheckStatus::stopped_by(id(CATALOG, limits)?)
    } else {
        CheckStatus::Complete
    };
    let signature = CheckStatus::blocked(reason, limits)?;
    let checks = vec![
        check(INGRESS, ingress, limits)?,
        check(CATALOG, catalog, limits)?,
        check(REACHABLE_RELATIONSHIPS, relationships, limits)?,
        check(SIGNATURE_PRESENCE, signature, limits)?,
    ];
    Ok(ValidateReport::try_new(checks, [], limits)?)
}

fn structural_issue(
    error: &OpcError,
    phase: StructuralPhase,
    limits: ValidationLimits,
) -> Result<ValidationIssue> {
    let (code, message, path) = match error {
        OpcError::ZipError(_) => (
            "opc.zip.invalid",
            "The input is not a structurally readable ZIP package.",
            "zip",
        ),
        OpcError::InvalidContentTypesManifest(_)
        | OpcError::ContentTypeNotFound(_)
        | OpcError::InvalidContentType { .. }
        | OpcError::DuplicateContentTypeDefault(_)
        | OpcError::DuplicateContentTypeOverride { .. }
        | OpcError::InvalidContentTypeExtension(_) => (
            "opc.content_types.invalid",
            "The OPC content-type catalog is structurally invalid.",
            "[Content_Types].xml",
        ),
        OpcError::InvalidRelationship(_)
        | OpcError::InvalidRelationshipsManifest(_)
        | OpcError::DuplicateRelationshipId(_)
        | OpcError::InvalidRelationshipTargetMode(_)
        | OpcError::RelationshipPartCannotBeSource(_)
        | OpcError::MultipleCorePropertiesRelationships => (
            "opc.relationships.invalid",
            "A manifest in the reachable OPC relationship graph is structurally invalid.",
            "reachable-relationship-graph",
        ),
        OpcError::DuplicatePartName(_)
        | OpcError::EquivalentPartNames { .. }
        | OpcError::DerivedPartNames { .. }
        | OpcError::InvalidPackUri(_)
        | OpcError::PartNotFound(_) => (
            "opc.catalog.invalid",
            "The OPC part catalog is structurally invalid.",
            "catalog",
        ),
        OpcError::QuickXmlError(_)
        | OpcError::XmlError(_)
        | OpcError::Utf8Error(_)
        | OpcError::ParseIntError(_)
        | OpcError::AttrError(_) => (
            "opc.manifest_xml.invalid",
            "A required OPC XML manifest is structurally invalid.",
            "manifest",
        ),
        _ => (
            "opc.package.invalid",
            "The input is not a structurally valid OPC package.",
            "package",
        ),
    };
    ValidationIssue::try_new(
        id(phase.capability(), limits)?,
        code,
        IssueSeverity::Error,
        message,
        [IssueLocation::try_new(
            None,
            Some(path),
            None,
            None,
            None,
            limits,
        )?],
        [],
        Some(SpecCitation::try_new("ECMA-376", "Part 2", limits)?),
        CompatibilityImpact::Interoperability,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

#[derive(Clone, Copy)]
enum StructuralPhase {
    Ingress,
    Catalog,
    Relationships,
}

impl StructuralPhase {
    const fn capability(self) -> &'static str {
        match self {
            Self::Ingress => INGRESS,
            Self::Catalog => CATALOG,
            Self::Relationships => REACHABLE_RELATIONSHIPS,
        }
    }
}

fn structural_phase(error: &OpcError) -> StructuralPhase {
    match error {
        OpcError::ZipError(_)
        | OpcError::QuickXmlError(_)
        | OpcError::InvalidPackUri(_)
        | OpcError::XmlError(_)
        | OpcError::Utf8Error(_)
        | OpcError::ParseIntError(_)
        | OpcError::AttrError(_) => StructuralPhase::Ingress,
        OpcError::InvalidRelationship(_)
        | OpcError::InvalidRelationshipsManifest(_)
        | OpcError::DuplicateRelationshipId(_)
        | OpcError::InvalidRelationshipTargetMode(_)
        | OpcError::RelationshipPartCannotBeSource(_)
        | OpcError::MultipleCorePropertiesRelationships => StructuralPhase::Relationships,
        _ => StructuralPhase::Catalog,
    }
}

fn signature_presence_issue(count: u64, limits: ValidationLimits) -> Result<ValidationIssue> {
    ValidationIssue::try_new(
        id(SIGNATURE_PRESENCE, limits)?,
        "opc.signature.infrastructure_present",
        IssueSeverity::Info,
        "OPC digital-signature infrastructure is present; signature validity was not checked.",
        [IssueLocation::try_new(
            None,
            Some("digital-signatures"),
            None,
            None,
            None,
            limits,
        )?],
        [IssueEvidence::try_new(
            "infrastructure_observations",
            EvidenceValue::Count(count),
            limits,
        )?],
        Some(SpecCitation::try_new("ECMA-376", "Part 2 §13", limits)?),
        CompatibilityImpact::None,
        RepairAvailability::Unavailable,
        limits,
    )
    .map_err(Into::into)
}

fn signature_infrastructure_count(package: &SourceBackedPackage) -> u64 {
    let package_relationships = package
        .rels()
        .iter()
        .filter(|relationship| signature_relationship(relationship.reltype()))
        .count() as u64;
    package
        .iter_parts()
        .fold(package_relationships, |count, part| {
            let own = u64::from(
                signature_path(part.partname().as_str())
                    || signature_content_type(part.content_type()),
            );
            let relationships = part
                .rels()
                .iter()
                .filter(|relationship| signature_relationship(relationship.reltype()))
                .count() as u64;
            count.saturating_add(own).saturating_add(relationships)
        })
}

fn signature_relationship(kind: &str) -> bool {
    matches!(
        kind,
        relationship_type::DIGITAL_SIGNATURE_ORIGIN
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
            | "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate"
    )
}

fn signature_path(path: &str) -> bool {
    const DIRECTORY: &[u8] = b"/_xmlsignatures/";
    path.as_bytes()
        .get(..DIRECTORY.len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case(DIRECTORY))
}

fn signature_content_type(value: &str) -> bool {
    matches!(
        value,
        content_type::OPC_DIGITAL_SIGNATURE_ORIGIN
            | content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE
            | content_type::OPC_DIGITAL_SIGNATURE_CERTIFICATE
    )
}

fn is_structural_rejection(error: &OpcError) -> bool {
    matches!(
        error,
        OpcError::InvalidPackUri(_)
            | OpcError::PartNotFound(_)
            | OpcError::DuplicatePartName(_)
            | OpcError::EquivalentPartNames { .. }
            | OpcError::DerivedPartNames { .. }
            | OpcError::ContentTypeNotFound(_)
            | OpcError::InvalidContentType { .. }
            | OpcError::InvalidContentTypesManifest(_)
            | OpcError::DuplicateContentTypeDefault(_)
            | OpcError::DuplicateContentTypeOverride { .. }
            | OpcError::InvalidContentTypeExtension(_)
            | OpcError::InvalidRelationship(_)
            | OpcError::InvalidRelationshipsManifest(_)
            | OpcError::DuplicateRelationshipId(_)
            | OpcError::InvalidRelationshipTargetMode(_)
            | OpcError::RelationshipPartCannotBeSource(_)
            | OpcError::MultipleCorePropertiesRelationships
            | OpcError::XmlError(_)
            | OpcError::ZipError(_)
            | OpcError::QuickXmlError(_)
            | OpcError::Utf8Error(_)
            | OpcError::ParseIntError(_)
            | OpcError::AttrError(_)
    )
}

fn id(value: &str, limits: ValidationLimits) -> Result<CheckCapabilityId> {
    CheckCapabilityId::try_new(value, limits).map_err(Into::into)
}

fn check(value: &str, status: CheckStatus, limits: ValidationLimits) -> Result<ValidationCheck> {
    Ok(ValidationCheck::new(id(value, limits)?, status))
}
