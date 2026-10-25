//! Shared, inert MS-OFFMACRO2 package-graph mutation.
//!
//! The `project` payload itself remains owned by `litchi-vba`. This module
//! only validates and mutates the OPC relationship graph shared by DOCX,
//! PPTX, XLSX, and XLSB hosts.

use crate::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{
    OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, SourceLimits, TargetMode,
};
use litchi_vba::{Limits, project::Project};
use sha2::{Digest, Sha256};
use std::sync::Arc;

const WORD_PROJECT_PART: &str = "/word/vbaProject.bin";
const WORD_PROJECT_TARGET: &str = "vbaProject.bin";
const WORD_SUPPLEMENTAL_PART: &str = "/word/vbaData.xml";
const WORD_SUPPLEMENTAL_TARGET: &str = "vbaData.xml";
const EXCEL_PROJECT_PART: &str = "/xl/vbaProject.bin";
const EXCEL_PROJECT_TARGET: &str = "vbaProject.bin";
const POWERPOINT_PROJECT_PART: &str = "/ppt/vbaProject.bin";
const POWERPOINT_PROJECT_TARGET: &str = "vbaProject.bin";

/// OOXML application family that owns a VBA project.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Host {
    Word,
    Excel,
    PowerPoint,
}

/// Resource dimensions checked before a VBA graph is staged or replayed.
///
/// The graph editor retains the OPC policy from its source and intersects it
/// with the selected edit options.  This keeps a patch from widening the
/// recipient's admission policy while still allowing a caller to select a
/// tighter budget for one operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum LimitKind {
    /// The opaque `vbaProject.bin` payload.
    ProjectBytes,
    /// Word's optional `vbaData.xml` payload.
    SupplementalXmlBytes,
    /// The combined bytes introduced by one graph publication.
    GraphBytes,
    /// The number of graph parts introduced by one publication.
    GraphParts,
    /// Aggregate materialized OPC part bytes after publication.
    TotalPartBytes,
    /// One materialized OPC part's bytes.
    PartBytes,
    /// One relationship member's serialized bytes.
    RelationshipXmlBytes,
    /// Aggregate relationship-member bytes after publication.
    TotalRelationshipXmlBytes,
    /// Relationship members in the package.
    RelationshipParts,
    /// One relationship member's edge count.
    RelationshipsPerPart,
    /// Aggregate relationship edge count after publication.
    TotalRelationships,
    /// The serialized content-types manifest bytes.
    ContentTypesBytes,
}

impl LimitKind {
    /// Stable content-free name for diagnostics and facade classification.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::ProjectBytes => "VBA project bytes",
            Self::SupplementalXmlBytes => "VBA supplemental XML bytes",
            Self::GraphBytes => "VBA graph bytes",
            Self::GraphParts => "VBA graph parts",
            Self::TotalPartBytes => "VBA total part bytes",
            Self::PartBytes => "VBA part bytes",
            Self::RelationshipXmlBytes => "VBA relationship XML bytes",
            Self::TotalRelationshipXmlBytes => "VBA total relationship XML bytes",
            Self::RelationshipParts => "VBA relationship parts",
            Self::RelationshipsPerPart => "VBA relationships per part",
            Self::TotalRelationships => "VBA total relationships",
            Self::ContentTypesBytes => "VBA content-types bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Bounds selected for one VBA graph edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct EditOptions {
    source_limits: SourceLimits,
    max_project_bytes: usize,
    max_supplemental_xml_bytes: usize,
    max_graph_bytes: usize,
    max_graph_parts: usize,
}

impl Default for EditOptions {
    fn default() -> Self {
        let limits = SourceLimits::default();
        Self {
            source_limits: limits,
            max_project_bytes: limits.max_part_bytes().min(usize::MAX as u64) as usize,
            max_supplemental_xml_bytes: limits.max_part_bytes().min(usize::MAX as u64) as usize,
            max_graph_bytes: limits.max_total_part_bytes().min(usize::MAX as u64) as usize,
            max_graph_parts: 2,
        }
    }
}

impl EditOptions {
    /// Start with the standard bounded VBA publication policy.
    #[must_use]
    pub fn new() -> Self {
        Self {
            source_limits: SourceLimits::default(),
            max_project_bytes: 512 * 1024 * 1024,
            max_supplemental_xml_bytes: 512 * 1024 * 1024,
            max_graph_bytes: 512 * 1024 * 1024,
            max_graph_parts: 2,
        }
    }

    /// Intersect this operation with a retained OPC source policy.
    #[must_use]
    pub const fn source_limits(self, limits: SourceLimits) -> Self {
        Self {
            source_limits: limits,
            ..self
        }
    }

    /// Tighten the opaque project payload bound.
    #[must_use]
    pub const fn max_project_bytes(self, maximum: usize) -> Self {
        Self {
            max_project_bytes: maximum,
            ..self
        }
    }

    /// Tighten the Word supplemental XML bound.
    #[must_use]
    pub const fn max_supplemental_xml_bytes(self, maximum: usize) -> Self {
        Self {
            max_supplemental_xml_bytes: maximum,
            ..self
        }
    }

    /// Tighten the total graph payload bound.
    #[must_use]
    pub const fn max_graph_bytes(self, maximum: usize) -> Self {
        Self {
            max_graph_bytes: maximum,
            ..self
        }
    }

    /// Tighten the number of graph parts introduced by one operation.
    #[must_use]
    pub const fn max_graph_parts(self, maximum: usize) -> Self {
        Self {
            max_graph_parts: maximum,
            ..self
        }
    }

    /// The selected source policy.
    #[must_use]
    pub const fn source_limits_value(self) -> SourceLimits {
        self.source_limits
    }

    /// The selected project payload ceiling.
    #[must_use]
    pub const fn max_project_bytes_value(self) -> usize {
        self.max_project_bytes
    }

    /// The selected supplemental payload ceiling.
    #[must_use]
    pub const fn max_supplemental_xml_bytes_value(self) -> usize {
        self.max_supplemental_xml_bytes
    }

    /// The selected combined graph payload ceiling.
    #[must_use]
    pub const fn max_graph_bytes_value(self) -> usize {
        self.max_graph_bytes
    }

    /// The selected graph part-count ceiling.
    #[must_use]
    pub const fn max_graph_parts_value(self) -> usize {
        self.max_graph_parts
    }

    fn effective(self, recipient: &OpcPackage) -> Self {
        let limits = self.source_limits.min(recipient.read_limits());
        Self {
            source_limits: limits,
            max_project_bytes: self
                .max_project_bytes
                .min(limits.max_part_bytes().min(usize::MAX as u64) as usize),
            max_supplemental_xml_bytes: self
                .max_supplemental_xml_bytes
                .min(limits.max_part_bytes().min(usize::MAX as u64) as usize),
            max_graph_bytes: self
                .max_graph_bytes
                .min(limits.max_total_part_bytes().min(usize::MAX as u64) as usize),
            max_graph_parts: self.max_graph_parts.min(limits.max_parts()),
        }
    }
}

#[derive(Debug, Clone)]
struct ExistingGraph {
    relationship_id: String,
    project_part: PackURI,
    supplemental: Option<(PackURI, String)>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelationshipState {
    id: String,
    relationship_type: String,
    target: String,
    target_mode: TargetMode,
}

impl RelationshipState {
    fn capture(value: &litchi_opc::Relationship) -> Self {
        Self {
            id: value.r_id().to_owned(),
            relationship_type: value.reltype().to_owned(),
            target: value.target_ref().to_owned(),
            target_mode: value.target_mode(),
        }
    }
}

#[derive(Clone)]
struct PartState {
    name: PackURI,
    content_type: String,
    data: Arc<Vec<u8>>,
    relationships: Vec<RelationshipState>,
    relationships_source: OwnedRelationships,
    /// Keep the concrete OPC part, including source relationship capture,
    /// admitted IDs, payload provenance and any format-owned opaque state.
    /// Rebuilding a `BlobPart` from the visible relationship projection would
    /// silently canonicalize those details during replay.
    part: Box<dyn Part + Send + Sync>,
}

impl PartState {
    fn capture(package: &OpcPackage, part: &dyn Part) -> Result<Self> {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::capture)
            .collect::<Vec<_>>();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Ok(Self {
            name: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships,
            relationships_source: package
                .source_relationships(part.partname(), package.read_limits())?,
            part: part.clone_part(),
        })
    }

    fn clone_part(&self) -> Box<dyn Part + Send + Sync> {
        self.part.clone()
    }
}

impl std::fmt::Debug for PartState {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("PartState")
            .field("name", &self.name)
            .field("content_type", &self.content_type)
            .field("bytes", &self.data.len())
            .field("relationships", &self.relationships)
            .finish_non_exhaustive()
    }
}

impl PartialEq for PartState {
    fn eq(&self, other: &Self) -> bool {
        self.name == other.name
            && self.content_type == other.content_type
            && self.data.as_slice() == other.data.as_slice()
            && self.relationships == other.relationships
            && self.relationships_source == other.relationships_source
    }
}

impl Eq for PartState {}

#[derive(Debug, Clone)]
struct GraphSnapshot {
    host: Host,
    source: PartState,
    project: Option<PartState>,
    supplemental: Option<PartState>,
    /// The current content-types publication image.  A graph patch restores
    /// it only when the recipient's selected base carried the same image;
    /// otherwise unrelated recipient overrides are left untouched.
    content_types: OwnedContentTypes,
}

impl PartialEq for GraphSnapshot {
    fn eq(&self, other: &Self) -> bool {
        self.host == other.host
            && self.source == other.source
            && self.project == other.project
            && self.supplemental == other.supplemental
    }
}

impl Eq for GraphSnapshot {}

impl GraphSnapshot {
    fn capture(package: &OpcPackage, source_part: &PackURI, host: Host) -> Result<Self> {
        let source = package.get_part(source_part)?;
        let existing = inspect_existing_graph(package, source_part, host)?;
        let project = existing
            .as_ref()
            .map(|graph| package.get_part(&graph.project_part))
            .transpose()?;
        let supplemental = existing
            .as_ref()
            .and_then(|graph| graph.supplemental.as_ref().map(|(part, _)| part))
            .map(|part| package.get_part(part))
            .transpose()?;
        Ok(Self {
            host,
            source: PartState::capture(package, source)?,
            project: project
                .map(|part| PartState::capture(package, part))
                .transpose()?,
            supplemental: supplemental
                .map(|part| PartState::capture(package, part))
                .transpose()?,
            content_types: package.source_content_types(package.read_limits())?,
        })
    }
}

/// A bounded, compact proof for an untouched edit.
///
/// An unchanged edit must not retain the selected graph twice merely to make
/// its replayable no-op patch.  The proof hashes only the affected source,
/// project, and supplemental parts together with their retained relationship
/// members; it does not retain their payload allocations. It therefore
/// remains content-based (rather than package-object identity based), while
/// unrelated package members and manifest overrides stay outside the replay
/// scope.
#[derive(Debug, Clone, PartialEq, Eq)]
struct GraphProof {
    host: Host,
    source_part: PackURI,
    digest: [u8; 32],
}

impl GraphProof {
    fn capture(package: &OpcPackage, source_part: &PackURI, host: Host) -> Result<Self> {
        let mut digest = Sha256::new();
        proof_field(&mut digest, host_name(host).as_bytes());
        proof_field(&mut digest, source_part.as_str().as_bytes());
        let source = package.get_part(source_part)?;
        hash_proof_part(&mut digest, package, source)?;
        let existing = inspect_existing_graph(package, source_part, host)?;
        match existing {
            Some(graph) => {
                digest.update([1]);
                let project = package.get_part(&graph.project_part)?;
                hash_proof_part(&mut digest, package, project)?;
                if let Some((supplemental_name, _)) = graph.supplemental {
                    digest.update([1]);
                    let supplemental = package.get_part(&supplemental_name)?;
                    hash_proof_part(&mut digest, package, supplemental)?;
                } else {
                    digest.update([0]);
                }
            },
            None => digest.update([0]),
        }
        Ok(Self {
            host,
            source_part: source_part.clone(),
            digest: digest.finalize().into(),
        })
    }
}

fn host_name(host: Host) -> &'static str {
    match host {
        Host::Word => "word",
        Host::Excel => "excel",
        Host::PowerPoint => "powerpoint",
    }
}

fn proof_field(digest: &mut Sha256, bytes: &[u8]) {
    digest.update((bytes.len() as u64).to_le_bytes());
    digest.update(bytes);
}

fn hash_proof_part(digest: &mut Sha256, package: &OpcPackage, part: &dyn Part) -> Result<()> {
    proof_field(digest, part.partname().as_str().as_bytes());
    proof_field(digest, part.content_type().as_bytes());
    proof_field(digest, part.blob());
    let relationships = package.source_relationships(part.partname(), package.read_limits())?;
    digest.update([u8::from(relationships.member_present())]);
    proof_field(digest, relationships.bytes());
    Ok(())
}

impl Host {
    fn project_part(self) -> &'static str {
        match self {
            Self::Word => WORD_PROJECT_PART,
            Self::Excel => EXCEL_PROJECT_PART,
            Self::PowerPoint => POWERPOINT_PROJECT_PART,
        }
    }

    fn project_target(self) -> &'static str {
        match self {
            Self::Word => WORD_PROJECT_TARGET,
            Self::Excel => EXCEL_PROJECT_TARGET,
            Self::PowerPoint => POWERPOINT_PROJECT_TARGET,
        }
    }

    fn macro_content_type(self, current: &str) -> Result<&'static str> {
        let mapped = match self {
            Self::Word => match current {
                ct::WML_DOCUMENT_MAIN | ct::WML_DOCUMENT_MACRO_MAIN => ct::WML_DOCUMENT_MACRO_MAIN,
                ct::WML_TEMPLATE_MAIN | ct::WML_TEMPLATE_MACRO_MAIN => ct::WML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::Excel => match current {
                ct::SML_SHEET_MAIN | ct::SML_SHEET_MACRO_MAIN => ct::SML_SHEET_MACRO_MAIN,
                ct::SML_TEMPLATE_MAIN | ct::SML_TEMPLATE_MACRO_MAIN => ct::SML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::PowerPoint => match current {
                ct::PML_PRESENTATION_MAIN | ct::PML_PRES_MACRO_MAIN => ct::PML_PRES_MACRO_MAIN,
                ct::PML_SLIDESHOW_MAIN | ct::PML_SLIDESHOW_MACRO_MAIN => {
                    ct::PML_SLIDESHOW_MACRO_MAIN
                },
                ct::PML_TEMPLATE_MAIN | ct::PML_TEMPLATE_MACRO_MAIN => ct::PML_TEMPLATE_MACRO_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
        };
        Ok(mapped)
    }

    fn non_macro_content_type(self, current: &str) -> Result<&'static str> {
        let mapped = match self {
            Self::Word => match current {
                ct::WML_DOCUMENT_MAIN | ct::WML_DOCUMENT_MACRO_MAIN => ct::WML_DOCUMENT_MAIN,
                ct::WML_TEMPLATE_MAIN | ct::WML_TEMPLATE_MACRO_MAIN => ct::WML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::Excel => match current {
                ct::SML_SHEET_MAIN | ct::SML_SHEET_MACRO_MAIN => ct::SML_SHEET_MAIN,
                ct::SML_TEMPLATE_MAIN | ct::SML_TEMPLATE_MACRO_MAIN => ct::SML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
            Self::PowerPoint => match current {
                ct::PML_PRESENTATION_MAIN | ct::PML_PRES_MACRO_MAIN => ct::PML_PRESENTATION_MAIN,
                ct::PML_SLIDESHOW_MAIN | ct::PML_SLIDESHOW_MACRO_MAIN => ct::PML_SLIDESHOW_MAIN,
                ct::PML_TEMPLATE_MAIN | ct::PML_TEMPLATE_MACRO_MAIN => ct::PML_TEMPLATE_MAIN,
                _ => return Err(invalid_main_content_type(self, current)),
            },
        };
        Ok(mapped)
    }
}

/// Parse one validated package part as inert MS-OVBA source.
/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn read_project_part(
    package: &OpcPackage,
    part_name: &PackURI,
    limits: &Limits,
) -> Result<Project> {
    let part = package.get_part(part_name)?;
    Project::read(part.blob(), limits).map_err(Error::from)
}

fn store_project_graph_in_place(
    package: &mut OpcPackage,
    source_part: &PackURI,
    host: Host,
    project_payload: Vec<u8>,
    word_supplemental_xml: Option<Vec<u8>>,
    options: EditOptions,
) -> Result<()> {
    validate_payload_limits(
        package,
        source_part,
        host,
        project_payload.len(),
        word_supplemental_xml.as_ref().map_or(0, Vec::len),
        options,
    )?;
    let macro_content_type = {
        let source = package.get_part(source_part)?;
        host.macro_content_type(source.content_type())?
    };
    if host == Host::Word && word_supplemental_xml.is_none() {
        return Err(Error::Invalid(
            "Word VBA project storage requires supplemental-data XML".to_string(),
        ));
    }
    if host != Host::Word && word_supplemental_xml.is_some() {
        return Err(Error::Invalid(
            "VBA supplemental-data XML is only valid for Word".to_string(),
        ));
    }

    let existing = inspect_existing_graph(package, source_part, host)?;
    ensure_replacement_targets_available(package, host, existing.as_ref())?;
    if let Some(graph) = &existing {
        ensure_exclusive_inbound_reference(
            package,
            &graph.project_part,
            source_part,
            &graph.relationship_id,
        )?;
        if let Some((supplemental, relationship_id)) = &graph.supplemental {
            ensure_exclusive_inbound_reference(
                package,
                supplemental,
                &graph.project_part,
                relationship_id,
            )?;
        }
    }

    refuse_signed_change(package)?;
    if let Some(graph) = existing {
        let expected = package.source_relationships(source_part, package.read_limits())?;
        let replacement = expected.without_relationship(
            &graph.relationship_id,
            options
                .effective(package)
                .source_limits
                .max_relationship_xml_bytes(),
        )?;
        package.try_replace_relationships(
            &expected,
            &replacement,
            options.effective(package).source_limits,
        )?;
        package.remove_part(&graph.project_part);
        if let Some((supplemental, _)) = graph.supplemental {
            package.remove_part(&supplemental);
        }
    }

    let project_name =
        PackURI::new(host.project_part()).map_err(|error| Error::Uri(error.clone()))?;
    let mut project = BlobPart::new(
        project_name,
        ct::OFC_VBA_PROJECT.to_string(),
        project_payload,
    );
    if let Some(xml) = word_supplemental_xml {
        project.relate_to(WORD_SUPPLEMENTAL_TARGET, rt::WORD_VBA_DATA);
        let supplemental_name =
            PackURI::new(WORD_SUPPLEMENTAL_PART).map_err(|error| Error::Uri(error.clone()))?;
        package.try_add_part(Box::new(BlobPart::new(
            supplemental_name,
            ct::WML_VBA_DATA.to_string(),
            xml,
        )))?;
    }
    package.try_add_part(Box::new(project))?;

    let expected = package.source_relationships(source_part, package.read_limits())?;
    let relationship_id = {
        let mut relationships = package.get_part(source_part)?.rels().clone();
        relationships
            .get_or_add(rt::VBA_PROJECT, host.project_target())
            .r_id()
            .to_owned()
    };
    let replacement = expected.with_relationship(
        rt::VBA_PROJECT,
        host.project_target(),
        &relationship_id,
        TargetMode::Internal,
        options
            .effective(package)
            .source_limits
            .max_relationship_xml_bytes(),
    )?;
    package.try_replace_relationships(
        &expected,
        &replacement,
        options.effective(package).source_limits,
    )?;
    let source = package.get_part_mut(source_part)?;
    source.set_content_type(macro_content_type.to_string())?;
    settle_unsigned_change(package)?;
    Ok(())
}

fn validate_payload_limits(
    package: &OpcPackage,
    source_part: &PackURI,
    host: Host,
    project_bytes: usize,
    supplemental_bytes: usize,
    options: EditOptions,
) -> Result<()> {
    let effective = options.effective(package);
    if project_bytes > effective.max_project_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::ProjectBytes),
            observed: project_bytes,
            maximum: effective.max_project_bytes,
        });
    }
    if supplemental_bytes > effective.max_supplemental_xml_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::SupplementalXmlBytes),
            observed: supplemental_bytes,
            maximum: effective.max_supplemental_xml_bytes,
        });
    }
    let graph_bytes = project_bytes.saturating_add(supplemental_bytes);
    if graph_bytes > effective.max_graph_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphBytes),
            observed: graph_bytes,
            maximum: effective.max_graph_bytes,
        });
    }
    let graph_parts = if host == Host::Word { 2 } else { 1 };
    if graph_parts > effective.max_graph_parts {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphParts),
            observed: graph_parts,
            maximum: effective.max_graph_parts,
        });
    }
    let total_part_bytes = package.iter_parts().try_fold(0u64, |total, metadata| {
        let bytes = u64::try_from(metadata.payload_len()?).unwrap_or(u64::MAX);
        total.checked_add(bytes).ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalPartBytes),
            observed: usize::MAX,
            maximum: limit_observed(effective.source_limits.max_total_part_bytes()),
        })
    })?;
    let replaced_part_bytes =
        if let Some(graph) = inspect_existing_graph(package, source_part, host)? {
            let project = package
                .part_metadata(&graph.project_part)
                .ok_or_else(|| Error::Missing(graph.project_part.to_string()))?;
            let mut bytes = u64::try_from(project.payload_len()?).unwrap_or(u64::MAX);
            if let Some((supplemental, _)) = graph.supplemental {
                let supplemental = package
                    .part_metadata(&supplemental)
                    .ok_or_else(|| Error::Missing(supplemental.to_string()))?;
                bytes = bytes
                    .saturating_add(u64::try_from(supplemental.payload_len()?).unwrap_or(u64::MAX));
            }
            bytes
        } else {
            0
        };
    let projected_total_part_bytes = total_part_bytes
        .saturating_sub(replaced_part_bytes)
        .saturating_add(u64::try_from(project_bytes).unwrap_or(u64::MAX))
        .saturating_add(u64::try_from(supplemental_bytes).unwrap_or(u64::MAX));
    check_vba_limit(
        LimitKind::TotalPartBytes,
        projected_total_part_bytes,
        effective.source_limits.max_total_part_bytes(),
    )?;
    Ok(())
}

/// Whether storing these bytes would change the selected graph at all.
///
/// The relationship identifier and target spelling are deliberately not part
/// of this test.  A logically identical store must not remove and recreate a
/// graph merely to canonicalize an authored source: doing so would lose its
/// retained source bytes and needlessly invalidate a signature.
fn store_is_noop(
    package: &OpcPackage,
    source_part: &PackURI,
    host: Host,
    project_payload: &[u8],
    supplemental_xml: Option<&[u8]>,
) -> Result<bool> {
    let source = package.get_part(source_part)?;
    let macro_content_type = host.macro_content_type(source.content_type())?;
    if source.content_type() != macro_content_type {
        return Ok(false);
    }
    let Some(graph) = inspect_existing_graph(package, source_part, host)? else {
        return Ok(false);
    };
    let project = package.get_part(&graph.project_part)?;
    if project.blob() != project_payload {
        return Ok(false);
    }
    match (host, supplemental_xml, graph.supplemental.as_ref()) {
        (Host::Word, Some(expected), Some((part, _))) => {
            Ok(package.get_part(part)?.blob() == expected)
        },
        (Host::Word, _, _) => Ok(false),
        (_, None, None) => Ok(true),
        (_, _, _) => Ok(false),
    }
}

fn remove_project_graph_on_candidate(
    package: &mut OpcPackage,
    source_part: &PackURI,
    host: Host,
) -> Result<bool> {
    let non_macro_content_type = {
        let source = package.get_part(source_part)?;
        host.non_macro_content_type(source.content_type())?
    };
    let existing = inspect_existing_graph(package, source_part, host)?;
    let Some(graph) = existing else {
        return Ok(false);
    };
    ensure_exclusive_inbound_reference(
        package,
        &graph.project_part,
        source_part,
        &graph.relationship_id,
    )?;
    if let Some((supplemental, relationship_id)) = &graph.supplemental {
        ensure_exclusive_inbound_reference(
            package,
            supplemental,
            &graph.project_part,
            relationship_id,
        )?;
    }

    refuse_signed_change(package)?;
    remove_project_graph_in_place(package, source_part, graph, non_macro_content_type)?;
    Ok(true)
}

fn remove_project_graph_in_place(
    package: &mut OpcPackage,
    source_part: &PackURI,
    graph: ExistingGraph,
    non_macro_content_type: &str,
) -> Result<()> {
    let expected = package.source_relationships(source_part, package.read_limits())?;
    let replacement = expected.without_relationship(
        &graph.relationship_id,
        package.read_limits().max_relationship_xml_bytes(),
    )?;
    package.try_replace_relationships(&expected, &replacement, package.read_limits())?;
    let source = package.get_part_mut(source_part)?;
    source.set_content_type(non_macro_content_type.to_string())?;
    package.remove_part(&graph.project_part);
    if let Some((supplemental, _)) = graph.supplemental {
        package.remove_part(&supplemental);
    }
    settle_unsigned_change(package)?;
    Ok(())
}

fn main_document_partname(package: &OpcPackage) -> Result<PackURI> {
    let mut matching = package.rels().iter().filter(|relationship| {
        matches!(
            relationship.reltype(),
            rt::OFFICE_DOCUMENT | rt::STRICT_OFFICE_DOCUMENT
        )
    });
    let relationship = matching.next().ok_or_else(|| {
        Error::Opc(litchi_opc::error::OpcError::InvalidRelationship(
            "main-document relationship is missing".to_string(),
        ))
    })?;
    if matching.next().is_some() {
        return Err(Error::Opc(
            litchi_opc::error::OpcError::InvalidRelationship(
                "package has multiple main-document relationships".to_string(),
            ),
        ));
    }
    if relationship.is_external() {
        return Err(Error::Opc(
            litchi_opc::error::OpcError::InvalidRelationship(
                "main-document relationship cannot be external".to_string(),
            ),
        ));
    }
    let partname = relationship.target_partname()?;
    if package.part_metadata(&partname).is_none() {
        return Err(Error::Opc(litchi_opc::error::OpcError::PartNotFound(
            partname.to_string(),
        )));
    }
    Ok(partname)
}

/// A staged VBA-project graph edit over one immutable OPC package.
///
/// Construction borrows the package and resolves only the main-part name.
/// The selected graph is captured when the first real mutation is requested,
/// and the structural package clone is deferred until that mutation. This
/// keeps an untouched edit O(1) in both work and additional payload storage.
pub struct Edit<'a> {
    source: &'a OpcPackage,
    candidate: Option<OpcPackage>,
    before: Option<GraphSnapshot>,
    source_part: PackURI,
    host: Host,
    options: EditOptions,
}

impl std::fmt::Debug for Edit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("host", &self.host)
            .field("source_part", &self.source_part)
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    /// Capture the selected host's main-part graph and start an isolated edit.
    ///
    /// The package is borrowed until a mutation is staged; unchanged part
    /// payloads remain shared by the one deferred candidate clone. The source
    /// package is never mutated.
    ///
    /// # Errors
    ///
    /// Returns an error when the package's main-document relationship is
    /// missing, ambiguous, external, or dangling.
    pub fn new(package: &'a OpcPackage, host: Host, options: EditOptions) -> Result<Self> {
        let source_part = main_document_partname(package)?;
        Ok(Self {
            source: package,
            candidate: None,
            before: None,
            source_part,
            host,
            options: options.effective(package),
        })
    }

    fn current(&self) -> &OpcPackage {
        self.candidate.as_ref().unwrap_or(self.source)
    }

    fn capture_before(&mut self) -> Result<()> {
        if self.before.is_none() {
            self.before = Some(GraphSnapshot::capture(
                self.source,
                &self.source_part,
                self.host,
            )?);
        }
        Ok(())
    }

    /// Attach or replace the opaque VBA project graph.
    ///
    /// Word requires supplemental XML; Excel and PowerPoint reject it. The
    /// payload is retained byte-for-byte and is never interpreted here.
    ///
    /// # Errors
    ///
    /// Returns graph validation errors and the signed-package refusal.
    pub fn store(
        &mut self,
        project_payload: Vec<u8>,
        word_supplemental_xml: Option<Vec<u8>>,
    ) -> Result<()> {
        validate_payload_limits(
            self.current(),
            &self.source_part,
            self.host,
            project_payload.len(),
            word_supplemental_xml.as_ref().map_or(0, Vec::len),
            self.options,
        )?;
        if store_is_noop(
            self.current(),
            &self.source_part,
            self.host,
            &project_payload,
            word_supplemental_xml.as_deref(),
        )? {
            return Ok(());
        }
        self.capture_before()?;
        // Keep the prior candidate detached until every mutation succeeds.
        // This makes a refused store atomic even when the edit already has
        // earlier staged operations.
        let mut candidate = self.current().clone();
        store_project_graph_in_place(
            &mut candidate,
            &self.source_part,
            self.host,
            project_payload,
            word_supplemental_xml,
            self.options,
        )?;
        self.candidate = Some(candidate);
        Ok(())
    }

    /// Remove the selected host's complete VBA graph.
    ///
    /// Returns `false` for an absent graph; that exact no-op keeps the source
    /// package and its signature unchanged.
    ///
    /// # Errors
    ///
    /// Returns graph validation errors and the signed-package refusal.
    pub fn remove(&mut self) -> Result<bool> {
        if inspect_existing_graph(self.current(), &self.source_part, self.host)?.is_none() {
            return Ok(false);
        }
        self.capture_before()?;
        let mut candidate = self.current().clone();
        let changed =
            remove_project_graph_on_candidate(&mut candidate, &self.source_part, self.host)?;
        self.candidate = Some(candidate);
        Ok(changed)
    }

    /// Discard every staged graph change.
    pub fn reset(&mut self) {
        self.candidate = None;
        self.before = None;
    }

    /// Publish the staged graph as a new package, exact patch, and diagnostics.
    ///
    /// An unchanged edit publishes the source package itself, preserving
    /// signatures and byte identity. A changed edit of a signed package is
    /// refused before publication.
    ///
    /// # Errors
    ///
    /// Returns graph validation errors and the signed-package refusal.
    pub fn commit(self) -> Result<Commit> {
        let Some(candidate) = self.candidate else {
            let patch = Patch {
                before: None,
                after: None,
                proof: Some(GraphProof::capture(
                    self.source,
                    &self.source_part,
                    self.host,
                )?),
                options: self.options,
            };
            return Ok(Commit::new(
                self.source.clone(),
                patch,
                Diagnostics::UNCHANGED,
            ));
        };
        let before = self.before.ok_or_else(|| {
            Error::Invalid("VBA edit candidate has no captured graph base".to_string())
        })?;
        let after = GraphSnapshot::capture(&candidate, &self.source_part, self.host)?;
        let patch = Patch {
            before: Some(before),
            after: Some(after),
            proof: None,
            options: self.options,
        };
        if patch.is_noop() {
            return Ok(Commit::new(
                self.source.clone(),
                patch,
                Diagnostics::UNCHANGED,
            ));
        }
        let before = patch
            .before
            .as_ref()
            .expect("changed VBA patch has a graph base");
        let after = patch
            .after
            .as_ref()
            .expect("changed VBA patch has a graph target");
        validate_snapshot_limits(self.source, before, after, self.options)?;
        refuse_signed_change(self.source)?;
        Ok(Commit::new(candidate, patch, Diagnostics::CHANGED))
    }
}

/// A complete, reversible VBA-project graph patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: Option<GraphSnapshot>,
    after: Option<GraphSnapshot>,
    proof: Option<GraphProof>,
    options: EditOptions,
}

impl Patch {
    /// Whether this patch is an exact graph no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.proof.is_some()
            || self
                .before
                .as_ref()
                .zip(self.after.as_ref())
                .is_some_and(|(before, after)| before == after)
    }

    /// Return the exact inverse patch without reinterpreting VBA bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        if self.proof.is_some() {
            return self.clone();
        }
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            proof: None,
            options: self.options,
        }
    }

    /// Replay this patch onto a package with the recorded graph base.
    ///
    /// The recipient may contain unrelated package changes; the selected VBA
    /// graph must match exactly. A changed patch is published on a cloned
    /// package and the recipient remains unchanged on every error.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the selected graph base differs
    /// and the signed-package refusal for a changed recipient.
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        let Some(before) = self.before.as_ref() else {
            let proof = self
                .proof
                .as_ref()
                .ok_or_else(|| Error::Invalid("VBA patch has no replay base".to_string()))?;
            let current = GraphProof::capture(package, &proof.source_part, proof.host)?;
            if current != *proof {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        };
        let after = self
            .after
            .as_ref()
            .ok_or_else(|| Error::Invalid("VBA patch has no target graph".to_string()))?;
        let current = GraphSnapshot::capture(package, &before.source.name, before.host)?;
        if current != *before {
            return Err(Error::PatchConflict);
        }
        let preserve_content_types = current.content_types.bytes() == before.content_types.bytes();
        validate_snapshot_limits(package, &current, after, self.options)?;
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
        restore_graph_in_place(&mut candidate, after, preserve_content_types)?;
        let resulting = GraphSnapshot::capture(&candidate, &after.source.name, after.host)?;
        if resulting != *after {
            return Err(Error::PatchConflict);
        }
        Ok(Commit::new(candidate, self.clone(), Diagnostics::CHANGED))
    }
}

/// The published VBA graph commit triad.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch>;

/// Minimal content-free commit diagnostics.
pub use litchi_core::edit::Diagnostics;

struct RelationshipTotals {
    xml_bytes: u64,
    relationships: u64,
    members: usize,
}

fn limit_observed(value: u64) -> usize {
    usize::try_from(value).unwrap_or(usize::MAX)
}

fn check_vba_limit(kind: LimitKind, observed: u64, maximum: u64) -> Result<()> {
    if observed > maximum {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(kind),
            observed: limit_observed(observed),
            maximum: limit_observed(maximum),
        });
    }
    Ok(())
}

fn relationship_totals(package: &OpcPackage, limits: SourceLimits) -> Result<RelationshipTotals> {
    let mut totals = RelationshipTotals {
        xml_bytes: 0,
        relationships: 0,
        members: 0,
    };
    let root = PackURI::new("/").map_err(|error| Error::Uri(error.clone()))?;
    add_relationship_totals(&mut totals, package, &root, package.rels(), limits)?;
    for metadata in package.iter_parts() {
        add_relationship_totals(
            &mut totals,
            package,
            metadata.partname(),
            metadata.rels(),
            limits,
        )?;
    }
    Ok(totals)
}

fn add_relationship_totals(
    totals: &mut RelationshipTotals,
    package: &OpcPackage,
    owner: &PackURI,
    relationships: &litchi_opc::rel::Relationships,
    limits: SourceLimits,
) -> Result<()> {
    let (xml_bytes, member_present) =
        if let Some((bytes, present)) = package.retained_relationships_xml(owner)? {
            (bytes.len(), present)
        } else {
            let plan = relationships.plan_canonical(limits)?;
            (
                plan.final_len(),
                owner.as_str() == "/" || !relationships.is_empty(),
            )
        };
    let xml_bytes = if member_present {
        u64::try_from(xml_bytes).unwrap_or(u64::MAX)
    } else {
        0
    };
    check_vba_limit(
        LimitKind::RelationshipXmlBytes,
        xml_bytes,
        limits.max_relationship_xml_bytes() as u64,
    )?;
    check_vba_limit(
        LimitKind::RelationshipXmlBytes,
        xml_bytes,
        limits.max_part_bytes(),
    )?;
    totals.xml_bytes = totals
        .xml_bytes
        .checked_add(xml_bytes)
        .ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalRelationshipXmlBytes),
            observed: usize::MAX,
            maximum: limits.max_total_relationship_xml_bytes(),
        })?;
    totals.relationships = totals
        .relationships
        .checked_add(u64::try_from(relationships.len()).unwrap_or(u64::MAX))
        .ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalRelationships),
            observed: usize::MAX,
            maximum: limits.max_total_relationships(),
        })?;
    totals.members = totals.members.saturating_add(usize::from(member_present));
    Ok(())
}

fn graph_part_states<'a>(graph: &'a GraphSnapshot) -> impl Iterator<Item = &'a PartState> {
    std::iter::once(&graph.source)
        .chain(graph.project.iter())
        .chain(graph.supplemental.iter())
}

fn validate_snapshot_limits(
    package: &OpcPackage,
    current: &GraphSnapshot,
    target: &GraphSnapshot,
    options: EditOptions,
) -> Result<()> {
    let project_bytes = target.project.as_ref().map_or(0, |part| part.data.len());
    let supplemental_bytes = target
        .supplemental
        .as_ref()
        .map_or(0, |part| part.data.len());
    let effective = options.effective(package);
    if project_bytes > effective.max_project_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::ProjectBytes),
            observed: project_bytes,
            maximum: effective.max_project_bytes,
        });
    }
    if supplemental_bytes > effective.max_supplemental_xml_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::SupplementalXmlBytes),
            observed: supplemental_bytes,
            maximum: effective.max_supplemental_xml_bytes,
        });
    }
    let graph_bytes = project_bytes.saturating_add(supplemental_bytes);
    if graph_bytes > effective.max_graph_bytes {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphBytes),
            observed: graph_bytes,
            maximum: effective.max_graph_bytes,
        });
    }
    let graph_parts =
        usize::from(target.project.is_some()) + usize::from(target.supplemental.is_some());
    if graph_parts > effective.max_graph_parts {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphParts),
            observed: graph_parts,
            maximum: effective.max_graph_parts,
        });
    }
    for part in graph_part_states(target) {
        check_vba_limit(
            LimitKind::PartBytes,
            u64::try_from(part.data.len()).unwrap_or(u64::MAX),
            effective.source_limits.max_part_bytes(),
        )?;
        check_vba_limit(
            LimitKind::RelationshipXmlBytes,
            u64::try_from(part.relationships_source.bytes().len()).unwrap_or(u64::MAX),
            effective.source_limits.max_relationship_xml_bytes() as u64,
        )?;
        check_vba_limit(
            LimitKind::RelationshipsPerPart,
            u64::try_from(part.relationships.len()).unwrap_or(u64::MAX),
            effective.source_limits.max_relationships_per_part() as u64,
        )?;
    }
    check_vba_limit(
        LimitKind::ContentTypesBytes,
        u64::try_from(target.content_types.bytes().len()).unwrap_or(u64::MAX),
        effective.source_limits.max_content_types_bytes() as u64,
    )?;

    let package_part_bytes = package.iter_parts().try_fold(0u64, |total, metadata| {
        let bytes = u64::try_from(metadata.payload_len()?).unwrap_or(u64::MAX);
        total.checked_add(bytes).ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalPartBytes),
            observed: usize::MAX,
            maximum: limit_observed(effective.source_limits.max_total_part_bytes()),
        })
    })?;
    let current_part_bytes = graph_part_states(current)
        .map(|part| u64::try_from(part.data.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let target_part_bytes = graph_part_states(target)
        .map(|part| u64::try_from(part.data.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let projected_part_bytes = package_part_bytes
        .saturating_sub(current_part_bytes)
        .saturating_add(target_part_bytes);
    check_vba_limit(
        LimitKind::TotalPartBytes,
        projected_part_bytes,
        effective.source_limits.max_total_part_bytes(),
    )?;

    let current_relationship_totals = relationship_totals(package, effective.source_limits)?;
    let removed_relationship_bytes = graph_part_states(current)
        .map(|part| u64::try_from(part.relationships_source.bytes().len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let added_relationship_bytes = graph_part_states(target)
        .map(|part| u64::try_from(part.relationships_source.bytes().len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let projected_relationship_bytes = current_relationship_totals
        .xml_bytes
        .saturating_sub(removed_relationship_bytes)
        .saturating_add(added_relationship_bytes);
    check_vba_limit(
        LimitKind::TotalRelationshipXmlBytes,
        projected_relationship_bytes,
        effective.source_limits.max_total_relationship_xml_bytes() as u64,
    )?;
    let removed_relationships = graph_part_states(current)
        .map(|part| u64::try_from(part.relationships.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let added_relationships = graph_part_states(target)
        .map(|part| u64::try_from(part.relationships.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let projected_relationships = current_relationship_totals
        .relationships
        .saturating_sub(removed_relationships)
        .saturating_add(added_relationships);
    check_vba_limit(
        LimitKind::TotalRelationships,
        projected_relationships,
        effective.source_limits.max_total_relationships() as u64,
    )?;
    let removed_members = graph_part_states(current)
        .filter(|part| part.relationships_source.member_present())
        .count();
    let added_members = graph_part_states(target)
        .filter(|part| part.relationships_source.member_present())
        .count();
    let projected_members = current_relationship_totals
        .members
        .saturating_sub(removed_members)
        .saturating_add(added_members);
    check_vba_limit(
        LimitKind::RelationshipParts,
        projected_members as u64,
        effective.source_limits.max_relationship_parts() as u64,
    )?;

    // `try_add_part` checks names and reserves each map insertion.  Check the
    // recipient's bounded part budget before cloning it so a replay cannot
    // allocate a candidate that its own source policy would reject.
    let removed = inspect_existing_graph(package, &target.source.name, target.host)?
        .map(|graph| {
            usize::from(package.part_metadata(&graph.project_part).is_some())
                + usize::from(
                    graph
                        .supplemental
                        .as_ref()
                        .is_some_and(|(part, _)| package.part_metadata(part).is_some()),
                )
        })
        .unwrap_or(0);
    let projected_parts = package
        .part_count()
        .saturating_sub(removed)
        .saturating_add(graph_parts);
    if projected_parts > effective.source_limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphParts),
            observed: projected_parts,
            maximum: effective.source_limits.max_parts(),
        });
    }
    Ok(())
}

fn restore_graph_in_place(
    package: &mut OpcPackage,
    target: &GraphSnapshot,
    preserve_content_types: bool,
) -> Result<()> {
    let source_part = &target.source.name;
    let existing = inspect_existing_graph(package, source_part, target.host)?;
    if let Some(graph) = &existing {
        ensure_exclusive_inbound_reference(
            package,
            &graph.project_part,
            source_part,
            &graph.relationship_id,
        )?;
        if let Some((supplemental, relationship_id)) = &graph.supplemental {
            ensure_exclusive_inbound_reference(
                package,
                supplemental,
                &graph.project_part,
                relationship_id,
            )?;
        }
    }
    refuse_signed_change(package)?;
    if let Some(graph) = &existing {
        package
            .get_part_mut(source_part)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.project_part);
        if let Some((supplemental, _)) = &graph.supplemental {
            package.remove_part(supplemental);
        }
    }

    // Replace the affected parts with their captured concrete values.  The
    // source-bound batch API installs the source relationship and
    // content-types tokens atomically; the per-part replacements below do the
    // same for the project and supplemental `.rels` members.  Plain
    // `try_add_part` would drop those package provenance entries even when the
    // cloned part itself still carried a source capture.
    if let Some(supplemental) = &target.supplemental {
        ensure_no_inbound_reference(package, &supplemental.name)?;
    }
    if let Some(project) = &target.project {
        ensure_no_inbound_reference(package, &project.name)?;
    }
    package.remove_part(source_part);
    package.try_add_part(target.source.clone_part())?;

    let limits = package.read_limits();
    let expected_content_types = package.source_content_types(limits)?;
    let replacement_content_types = if preserve_content_types {
        target.content_types.clone()
    } else {
        let mut overrides = Vec::new();
        if let Some(project) = &target.project {
            overrides.push((&project.name, project.content_type.as_str()));
        }
        if let Some(supplemental) = &target.supplemental {
            overrides.push((&supplemental.name, supplemental.content_type.as_str()));
        }
        expected_content_types.with_part_overrides(&overrides, limits.max_content_types_bytes())?
    };
    let expected_relationships = package.source_relationships(source_part, limits)?;
    let mut parts = Vec::new();
    if let Some(supplemental) = &target.supplemental {
        parts.push(supplemental.clone_part());
    }
    if let Some(project) = &target.project {
        parts.push(project.clone_part());
    }
    package.try_add_parts_with_source_tokens(
        expected_content_types.bytes(),
        &replacement_content_types,
        &expected_relationships,
        &target.source.relationships_source,
        parts,
    )?;

    for part in target.supplemental.iter().chain(target.project.iter()) {
        let current = package.source_relationships(&part.name, limits)?;
        if current != part.relationships_source {
            package.try_replace_relationships(&current, &part.relationships_source, limits)?;
        }
    }

    settle_unsigned_change(package)?;
    Ok(())
}

fn inspect_existing_graph(
    package: &OpcPackage,
    source_part: &PackURI,
    host: Host,
) -> Result<Option<ExistingGraph>> {
    let source = package.get_part(source_part)?;
    let mut matches = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::VBA_PROJECT);
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::Invalid(
            "OOXML main part has multiple VBA Project relationships".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(Error::Invalid(
            "VBA Project relationship cannot be external".to_string(),
        ));
    }
    let relationship_id = relationship.r_id().to_string();
    let project_part = relationship.target_partname()?;
    let project = package.get_part(&project_part)?;
    if project.content_type() != ct::OFC_VBA_PROJECT {
        return Err(Error::ContentType {
            expected: ct::OFC_VBA_PROJECT.to_string(),
            actual: project.content_type().to_string(),
        });
    }

    let supplemental = if host == Host::Word {
        let mut supplemental = project
            .rels()
            .iter()
            .filter(|relationship| relationship.reltype() == rt::WORD_VBA_DATA);
        let Some(relationship) = supplemental.next() else {
            return Err(Error::Invalid(
                "Word VBA Project part is missing supplemental data".to_string(),
            ));
        };
        if supplemental.next().is_some()
            || project
                .rels()
                .iter()
                .any(|candidate| candidate.reltype() != rt::WORD_VBA_DATA)
        {
            return Err(Error::Invalid(
                "Word VBA Project part has invalid relationships".to_string(),
            ));
        }
        if relationship.is_external() {
            return Err(Error::Invalid(
                "Word VBA supplemental-data relationship cannot be external".to_string(),
            ));
        }
        let part_name = relationship.target_partname()?;
        let part = package.get_part(&part_name)?;
        if part.content_type() != ct::WML_VBA_DATA {
            return Err(Error::ContentType {
                expected: ct::WML_VBA_DATA.to_string(),
                actual: part.content_type().to_string(),
            });
        }
        if !part.rels().is_empty() {
            return Err(Error::Invalid(
                "Word VBA supplemental-data part cannot have relationships".to_string(),
            ));
        }
        Some((part_name, relationship.r_id().to_string()))
    } else {
        if !project.rels().is_empty() {
            return Err(Error::Invalid(
                "Excel or PowerPoint VBA Project part cannot have relationships".to_string(),
            ));
        }
        None
    };

    Ok(Some(ExistingGraph {
        relationship_id,
        project_part,
        supplemental,
    }))
}

fn ensure_replacement_targets_available(
    package: &OpcPackage,
    host: Host,
    existing: Option<&ExistingGraph>,
) -> Result<()> {
    let canonical_project =
        PackURI::new(host.project_part()).map_err(|error| Error::Uri(error.clone()))?;
    if existing.is_none_or(|graph| graph.project_part != canonical_project) {
        package.validate_new_part_name(&canonical_project)?;
        ensure_no_inbound_reference(package, &canonical_project)?;
    }
    if host == Host::Word {
        let canonical_supplemental =
            PackURI::new(WORD_SUPPLEMENTAL_PART).map_err(|error| Error::Uri(error.clone()))?;
        if existing
            .and_then(|graph| graph.supplemental.as_ref())
            .map(|(part, _)| part)
            != Some(&canonical_supplemental)
        {
            package.validate_new_part_name(&canonical_supplemental)?;
            ensure_no_inbound_reference(package, &canonical_supplemental)?;
        }
    }
    Ok(())
}

/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn ensure_no_inbound_reference(package: &OpcPackage, target: &PackURI) -> Result<()> {
    let package_relationship = package.rels().iter().any(|relationship| {
        !relationship.is_external() && relationship.target_partname().ok().as_ref() == Some(target)
    });
    let part_relationship = package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship.target_partname().ok().as_ref() == Some(target)
        })
    });
    if package_relationship || part_relationship {
        return Err(Error::Invalid(format!(
            "VBA graph target '{}' already has an inbound relationship",
            target.as_str()
        )));
    }
    Ok(())
}

/// # Errors
///
/// Returns an error when input violates OOXML constraints, exceeds a configured
/// bound, or an underlying XML or package operation fails.
pub fn ensure_exclusive_inbound_reference(
    package: &OpcPackage,
    target: &PackURI,
    expected_source: &PackURI,
    expected_relationship_id: &str,
) -> Result<()> {
    let mut unexpected = false;
    for relationship in package.rels().iter() {
        if !relationship.is_external()
            && relationship.target_partname().ok().as_ref() == Some(target)
        {
            unexpected = true;
        }
    }
    for part in package.iter_parts() {
        for relationship in part.rels().iter() {
            if relationship.is_external()
                || relationship.target_partname().ok().as_ref() != Some(target)
            {
                continue;
            }
            if part.partname() != expected_source || relationship.r_id() != expected_relationship_id
            {
                unexpected = true;
            }
        }
    }
    if unexpected {
        return Err(Error::Invalid(format!(
            "VBA graph part '{}' has an unexpected inbound relationship",
            target.as_str()
        )));
    }
    Ok(())
}

fn invalid_main_content_type(host: Host, actual: &str) -> Error {
    Error::Invalid(format!(
        "{host:?} main part has unsupported content type '{actual}'"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackageWriter;

    #[derive(Clone)]
    struct ImmutableContentTypePart {
        inner: BlobPart,
    }

    impl Part for ImmutableContentTypePart {
        fn partname(&self) -> &PackURI {
            self.inner.partname()
        }

        fn content_type(&self) -> &str {
            self.inner.content_type()
        }

        fn blob(&self) -> &[u8] {
            self.inner.blob()
        }

        fn blob_arc(&self) -> Arc<Vec<u8>> {
            self.inner.blob_arc()
        }

        fn set_blob(&mut self, blob: Vec<u8>) {
            self.inner.set_blob(blob);
        }

        fn rels(&self) -> &litchi_opc::rel::Relationships {
            self.inner.rels()
        }

        fn rels_mut(&mut self) -> &mut litchi_opc::rel::Relationships {
            self.inner.rels_mut()
        }
    }

    #[test]
    fn project_part_reads_through_the_bounded_vba_owner() {
        let part_name = PackURI::new("/xl/vbaProject.bin").expect("project URI");
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                part_name.clone(),
                ct::OFC_VBA_PROJECT.to_owned(),
                vec![0; 4],
            )))
            .expect("project part");
        let limits = Limits {
            max_cfb_bytes: 3,
            ..Limits::default()
        };

        assert!(matches!(
            read_project_part(&package, &part_name, &limits),
            Err(Error::Vba(litchi_vba::Error::Limit {
                kind: litchi_vba::LimitKind::CfbBytes,
                observed: 4,
                maximum: 3
            }))
        ));
    }

    #[test]
    fn maps_every_supported_main_part_kind_both_directions() {
        for (host, plain, macro_enabled) in [
            (
                Host::Word,
                ct::WML_DOCUMENT_MAIN,
                ct::WML_DOCUMENT_MACRO_MAIN,
            ),
            (
                Host::Word,
                ct::WML_TEMPLATE_MAIN,
                ct::WML_TEMPLATE_MACRO_MAIN,
            ),
            (Host::Excel, ct::SML_SHEET_MAIN, ct::SML_SHEET_MACRO_MAIN),
            (
                Host::Excel,
                ct::SML_TEMPLATE_MAIN,
                ct::SML_TEMPLATE_MACRO_MAIN,
            ),
            (
                Host::PowerPoint,
                ct::PML_PRESENTATION_MAIN,
                ct::PML_PRES_MACRO_MAIN,
            ),
            (
                Host::PowerPoint,
                ct::PML_SLIDESHOW_MAIN,
                ct::PML_SLIDESHOW_MACRO_MAIN,
            ),
            (
                Host::PowerPoint,
                ct::PML_TEMPLATE_MAIN,
                ct::PML_TEMPLATE_MACRO_MAIN,
            ),
        ] {
            assert_eq!(host.macro_content_type(plain).unwrap(), macro_enabled);
            assert_eq!(
                host.macro_content_type(macro_enabled).unwrap(),
                macro_enabled
            );
            assert_eq!(host.non_macro_content_type(macro_enabled).unwrap(), plain);
            assert_eq!(host.non_macro_content_type(plain).unwrap(), plain);
        }
    }

    #[test]
    fn rejects_dangling_inbound_references_to_canonical_word_targets_atomically() {
        for (target, canonical) in [
            ("../word/vbaProject.bin", WORD_PROJECT_PART),
            ("../word/vbaData.xml", WORD_SUPPLEMENTAL_PART),
        ] {
            let source_name = PackURI::new("/word/document.xml").unwrap();
            let owner_name = PackURI::new("/custom/owner.xml").unwrap();
            let mut owner = BlobPart::new(
                owner_name.clone(),
                "application/xml".to_string(),
                b"owner".to_vec(),
            );
            owner.relate_to(target, "urn:test:dangling");

            let mut package = OpcPackage::new();
            package
                .try_add_part(Box::new(BlobPart::new(
                    source_name.clone(),
                    ct::WML_DOCUMENT_MAIN.to_string(),
                    b"document".to_vec(),
                )))
                .unwrap();
            package.try_add_part(Box::new(owner)).unwrap();
            package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);

            let before_parts = package.part_count();
            let before_source_relationships = package.get_part(&source_name).unwrap().rels().len();
            let before_owner_relationships = package.get_part(&owner_name).unwrap().rels().len();
            assert_eq!(
                package
                    .get_part(&owner_name)
                    .unwrap()
                    .rels()
                    .iter()
                    .next()
                    .unwrap()
                    .target_partname()
                    .unwrap()
                    .as_str(),
                canonical
            );

            assert!(
                Edit::new(&package, Host::Word, EditOptions::new())
                    .and_then(|mut edit| {
                        edit.store(
                            b"validated payload".to_vec(),
                            Some(b"<wne:vbaSuppData/>".to_vec()),
                        )
                    })
                    .is_err()
            );

            assert_eq!(package.part_count(), before_parts);
            let source = package.get_part(&source_name).unwrap();
            assert_eq!(source.content_type(), ct::WML_DOCUMENT_MAIN);
            assert_eq!(source.rels().len(), before_source_relationships);
            assert_eq!(
                package.get_part(&owner_name).unwrap().rels().len(),
                before_owner_relationships
            );
            assert!(package.get_part(&PackURI::new(canonical).unwrap()).is_err());
        }
    }

    #[test]
    fn signed_store_and_remove_are_refused_until_unsigned() {
        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let project_name = PackURI::new(EXCEL_PROJECT_PART).unwrap();
        let refused = |result: Result<_>| {
            matches!(
                result,
                Err(Error::Opc(
                    litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
                ))
            )
        };
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                source_name.clone(),
                ct::SML_SHEET_MAIN.to_string(),
                b"workbook".to_vec(),
            )))
            .unwrap();
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let parts = package.part_count();

        // Removing an absent graph publishes nothing and keeps the signature.
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        assert!(!edit.remove().unwrap());
        let _ = edit.commit().unwrap();
        assert!(package.is_signed());
        // Storing changes the signed package, so it is refused (ADR 0006).
        let store = |package: &OpcPackage| {
            let mut edit = Edit::new(package, Host::Excel, EditOptions::new())?;
            edit.store(b"project".to_vec(), None)?;
            edit.commit()
        };
        assert!(refused(store(&package)));
        assert!(package.is_signed());
        assert_eq!(package.part_count(), parts);
        assert_eq!(
            package.get_part(&source_name).unwrap().content_type(),
            ct::SML_SHEET_MAIN
        );

        package.unsign();
        package = store(&package).unwrap().into_package();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(refused(
            Edit::new(&package, Host::Excel, EditOptions::new()).and_then(|mut edit| {
                edit.remove()?;
                edit.commit().map(|_| ())
            })
        ));
        assert!(package.get_part(&project_name).is_ok());
        package.unsign();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        let changed = edit.remove().unwrap();
        let next = edit.commit().unwrap().into_package();
        assert!(changed);
        package = next;
        assert!(package.get_part(&project_name).is_err());
    }

    #[test]
    fn semantic_helpers_publish_candidates_without_mutating_source() {
        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                source_name.clone(),
                ct::SML_SHEET_MAIN.to_string(),
                b"workbook".to_vec(),
            )))
            .unwrap();

        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let stored = edit.commit().unwrap().into_package();
        assert!(
            package
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_err()
        );
        assert_eq!(package.part_count(), 1);
        assert_eq!(
            stored.get_part(&source_name).unwrap().content_type(),
            ct::SML_SHEET_MACRO_MAIN
        );

        let mut edit = Edit::new(&stored, Host::Excel, EditOptions::new()).unwrap();
        let changed = edit.remove().unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert!(changed);
        assert!(
            stored
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_ok()
        );
        assert!(
            removed
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_err()
        );
        assert_eq!(
            removed.get_part(&source_name).unwrap().content_type(),
            ct::SML_SHEET_MAIN
        );
    }

    fn excel_package() -> OpcPackage {
        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                source_name,
                ct::SML_SHEET_MAIN.to_string(),
                b"workbook".to_vec(),
            )))
            .unwrap();
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn patch_preserves_noncanonical_relationship_bytes_through_replay_and_inverse() {
        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let unrelated_name = PackURI::new("/custom/unrelated.bin").unwrap();
        let mut package = excel_package();
        package
            .try_add_part(Box::new(BlobPart::new(
                unrelated_name,
                "application/octet-stream".to_string(),
                b"unrelated".to_vec(),
            )))
            .unwrap();
        package
            .get_part_mut(&source_name)
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_string(),
                "../custom/unrelated.bin".to_string(),
                "legacyEdge".to_string(),
                TargetMode::Internal,
            )
            .unwrap();
        let limits = package.read_limits();
        let expected = package.source_relationships(&source_name, limits).unwrap();
        let mut noncanonical = expected.bytes().to_vec();
        let marker = b"<Relationship ";
        let offset = noncanonical
            .windows(marker.len())
            .position(|window| window == marker)
            .expect("relationship element");
        noncanonical.splice(
            offset..offset + marker.len(),
            b"\n  <Relationship ".iter().copied(),
        );
        package
            .try_replace_relationships_bytes(
                &expected,
                Arc::new(noncanonical.clone()),
                true,
                limits,
            )
            .unwrap();

        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let commit = edit.commit().unwrap();
        let stored = commit.package().clone();
        let replayed = commit.patch().apply(&package).unwrap();
        let replay_source = replayed
            .package()
            .retained_relationships_xml(&source_name)
            .unwrap()
            .unwrap()
            .0;
        assert!(
            replay_source
                .windows(b"\n  <Relationship ".len())
                .any(|window| { window == b"\n  <Relationship " })
        );

        let reverted = commit.patch().inverse().apply(&stored).unwrap();
        let reverted_source = reverted
            .package()
            .retained_relationships_xml(&source_name)
            .unwrap()
            .unwrap()
            .0;
        assert_eq!(reverted_source.as_slice(), noncanonical.as_slice());
    }

    #[test]
    fn replay_rejects_tighter_recipient_aggregate_limits_atomically() {
        let package = excel_package();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let part_bytes = package
            .iter_parts()
            .map(|part| part.payload_len().unwrap() as u64)
            .sum::<u64>();
        let recipient_limits = SourceLimits::builder()
            .max_total_part_bytes(part_bytes)
            .unwrap();
        let recipient = OpcPackage::from_bytes(&bytes, recipient_limits).unwrap();
        let before = PackageWriter::to_bytes(&recipient).unwrap();
        assert!(matches!(
            patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::TotalPartBytes),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);
    }

    #[test]
    fn replay_rejects_tighter_recipient_relationship_budget_atomically() {
        let package = excel_package();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let bytes = PackageWriter::to_bytes(&package).unwrap();
        let root = PackURI::new("/").unwrap();
        let relationship_bytes = package
            .source_relationships(&root, package.read_limits())
            .unwrap()
            .bytes()
            .len();
        let recipient_limits = SourceLimits::builder()
            .max_total_relationship_xml_bytes(relationship_bytes)
            .unwrap();
        let recipient = OpcPackage::from_bytes(&bytes, recipient_limits).unwrap();
        let before = PackageWriter::to_bytes(&recipient).unwrap();
        assert!(matches!(
            patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::TotalRelationshipXmlBytes),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);
    }

    #[test]
    fn untouched_edit_replays_its_selected_graph_proof_without_snapshot_payloads() {
        let package = excel_package();
        let commit = Edit::new(&package, Host::Excel, EditOptions::new())
            .unwrap()
            .commit()
            .unwrap();
        assert!(commit.patch().is_noop());
        let replayed = commit.patch().apply(&package).unwrap();
        assert!(!replayed.diagnostics().changed());

        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let mut changed = package.clone();
        changed
            .get_part_mut(&source_name)
            .unwrap()
            .set_blob(b"changed".to_vec());
        assert!(matches!(
            commit.patch().apply(&changed),
            Err(Error::PatchConflict)
        ));
    }

    #[test]
    fn patch_replays_on_equal_copy_with_unrelated_changes_and_inverts() {
        let mut package = excel_package();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/octet-stream".to_string(),
            vec![1, 2, 3, 4],
        )));
        package
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_string(),
                "../custom/unrelated.bin".to_string(),
                "legacyEdge".to_string(),
                TargetMode::Internal,
            )
            .unwrap();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let commit = edit.commit().unwrap();
        let stored = commit.package().clone();
        let patch = commit.patch().clone();

        let mut recipient = package.clone();
        recipient.add_part(Box::new(BlobPart::new(
            PackURI::new("/custom/unrelated.bin").unwrap(),
            "application/octet-stream".to_string(),
            vec![9, 8, 7],
        )));
        let replayed = patch.apply(&recipient).unwrap();
        assert!(replayed.diagnostics().changed());
        assert!(
            replayed
                .package()
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_ok()
        );
        assert!(
            replayed
                .package()
                .get_part(&PackURI::new("/custom/unrelated.bin").unwrap())
                .is_ok()
        );
        assert!(
            replayed
                .package()
                .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
                .unwrap()
                .rels()
                .get("legacyEdge")
                .is_some()
        );

        let reverted = patch.inverse().apply(&stored).unwrap();
        assert!(reverted.diagnostics().changed());
        assert!(
            reverted
                .package()
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_err()
        );
        assert_eq!(
            reverted
                .package()
                .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
                .unwrap()
                .content_type(),
            ct::SML_SHEET_MAIN
        );
    }

    #[test]
    fn patch_rejects_a_changed_selected_graph_as_patch_conflict() {
        let package = excel_package();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let patch = edit.commit().unwrap().patch().clone();

        let mut changed = package.clone();
        changed
            .get_part_mut(&PackURI::new("/xl/workbook.xml").unwrap())
            .unwrap()
            .set_blob(b"different".to_vec());
        assert!(matches!(patch.apply(&changed), Err(Error::PatchConflict)));
    }

    #[test]
    fn reset_and_identical_store_are_exact_noops() {
        let package = excel_package();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        edit.reset();
        let reset = edit.commit().unwrap();
        assert!(!reset.diagnostics().changed());
        assert_eq!(reset.package().part_count(), package.part_count());

        let mut retained_limits = Edit::new(
            &package,
            Host::Excel,
            EditOptions::new().max_project_bytes(1),
        )
        .unwrap();
        retained_limits.reset();
        assert!(matches!(
            retained_limits.store(b"project".to_vec(), None),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::ProjectBytes),
                ..
            })
        ));

        let stored = {
            let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
            edit.store(b"project".to_vec(), None).unwrap();
            edit.commit().unwrap().into_package()
        };
        let mut same = Edit::new(&stored, Host::Excel, EditOptions::new()).unwrap();
        same.store(b"project".to_vec(), None).unwrap();
        let noop = same.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert!(
            noop.package()
                .get_part(&PackURI::new(EXCEL_PROJECT_PART).unwrap())
                .is_ok()
        );

        let mut signed = stored;
        signed.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        let mut same = Edit::new(&signed, Host::Excel, EditOptions::new()).unwrap();
        same.store(b"project".to_vec(), None).unwrap();
        let noop = same.commit().unwrap();
        assert!(!noop.diagnostics().changed());
        assert!(noop.package().is_signed());
    }

    #[test]
    fn selected_limits_refuse_before_staging_and_leave_the_edit_atomic() {
        let package = excel_package();
        let options = EditOptions::new().max_project_bytes(1);
        let mut edit = Edit::new(&package, Host::Excel, options).unwrap();
        assert!(matches!(
            edit.store(b"too large".to_vec(), None),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::ProjectBytes),
                ..
            })
        ));
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.package().part_count(), package.part_count());
    }

    #[test]
    fn late_content_type_failure_rolls_back_vba_removal() {
        let source_name = PackURI::new("/xl/workbook.xml").unwrap();
        let project_name = PackURI::new(EXCEL_PROJECT_PART).unwrap();
        let mut source = BlobPart::new(
            source_name.clone(),
            ct::SML_SHEET_MACRO_MAIN.to_string(),
            b"workbook".to_vec(),
        );
        let relationship_id = source.relate_to(EXCEL_PROJECT_TARGET, rt::VBA_PROJECT);

        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(ImmutableContentTypePart { inner: source }))
            .unwrap();
        package
            .try_add_part(Box::new(BlobPart::new(
                project_name.clone(),
                ct::OFC_VBA_PROJECT.to_string(),
                b"project".to_vec(),
            )))
            .unwrap();
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        let before_parts = package.part_count();

        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        assert!(edit.remove().is_err());

        assert_eq!(package.part_count(), before_parts);
        assert!(package.get_part(&project_name).is_ok());
        let source = package.get_part(&source_name).unwrap();
        assert_eq!(source.content_type(), ct::SML_SHEET_MACRO_MAIN);
        assert!(source.rels().get(&relationship_id).is_some());
    }
}
