//! Shared, inert MS-OFFMACRO2 package-graph mutation.
//!
//! The `project` payload itself remains owned by `litchi-vba`. This module
//! only validates and mutates the OPC relationship graph shared by DOCX,
//! PPTX, XLSX, and XLSB hosts.

/// The VBA family uses the shared OOXML error contract.
pub use crate::Error;

use crate::Result;
use crate::signature::{refuse_signed_change, settle_unsigned_change};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{
    OpcPackage, OwnedContentTypes, OwnedRelationships, PackURI, RetainedArchiveMetrics,
    SourceLimits, TargetMode,
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
    /// XML events in one relationship member.
    XmlEvents,
    /// Aggregate XML events across relationship members.
    TotalRelationshipXmlEvents,
    /// Aggregate relationship edge count after publication.
    TotalRelationships,
    /// The serialized content-types manifest bytes.
    ContentTypesBytes,
    /// Content-type default and override mappings.
    ContentTypeMappings,
    /// XML nesting depth.
    XmlDepth,
    /// XML attribute bytes.
    XmlAttributeBytes,
    /// Relationship target bytes.
    RelationshipTargetBytes,
    /// ZIP non-directory members.
    ArchiveMembers,
    /// ZIP central-directory entries.
    ArchiveTotalEntries,
    /// ZIP member-name bytes.
    ArchiveMemberNameBytes,
    /// ZIP central-directory metadata bytes.
    ArchiveMetadataBytes,
    /// One ZIP member's compressed bytes.
    ArchiveCompressedBytes,
    /// Aggregate ZIP uncompressed bytes.
    AggregateBytes,
    /// One ZIP entry's uncompressed bytes.
    EntryBytes,
    /// A source-policy dimension without a more specific VBA name.
    SourcePolicy,
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
            Self::XmlEvents => "VBA relationship XML events",
            Self::TotalRelationshipXmlEvents => "VBA total relationship XML events",
            Self::TotalRelationships => "VBA total relationships",
            Self::ContentTypesBytes => "VBA content-types bytes",
            Self::ContentTypeMappings => "VBA content-type mappings",
            Self::XmlDepth => "VBA XML depth",
            Self::XmlAttributeBytes => "VBA XML attribute bytes",
            Self::RelationshipTargetBytes => "VBA relationship target bytes",
            Self::ArchiveMembers => "VBA archive members",
            Self::ArchiveTotalEntries => "VBA archive total entries",
            Self::ArchiveMemberNameBytes => "VBA archive member-name bytes",
            Self::ArchiveMetadataBytes => "VBA archive metadata bytes",
            Self::ArchiveCompressedBytes => "VBA archive compressed bytes",
            Self::AggregateBytes => "VBA aggregate bytes",
            Self::EntryBytes => "VBA entry bytes",
            Self::SourcePolicy => "VBA source policy",
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
    fn capture(package: &OpcPackage, part: &dyn Part, limits: SourceLimits) -> Result<Self> {
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
            relationships_source: package.source_relationships(part.partname(), limits)?,
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
    content_type_parts: Vec<(PackURI, Option<String>)>,
}

impl PartialEq for GraphSnapshot {
    fn eq(&self, other: &Self) -> bool {
        self.same_selected_graph(other) && self.content_types == other.content_types
    }
}

impl Eq for GraphSnapshot {}

impl GraphSnapshot {
    fn capture(
        package: &OpcPackage,
        source_part: &PackURI,
        host: Host,
        limits: SourceLimits,
    ) -> Result<Self> {
        let existing = inspect_existing_graph(package, source_part, host)?;
        let source = admit_selected_part(package, source_part, limits)?;
        let project = existing
            .as_ref()
            .map(|graph| admit_selected_part(package, &graph.project_part, limits))
            .transpose()?;
        let supplemental = existing
            .as_ref()
            .and_then(|graph| graph.supplemental.as_ref().map(|(part, _)| part))
            .map(|part| admit_selected_part(package, part, limits))
            .transpose()?;
        let source = PartState::capture(package, source, limits)?;
        let project = project
            .map(|part| PartState::capture(package, part, limits))
            .transpose()?;
        let supplemental = supplemental
            .map(|part| PartState::capture(package, part, limits))
            .transpose()?;
        let content_types = package.source_content_types(limits)?;
        let mut content_type_parts = Vec::with_capacity(3);
        for part in [&source]
            .into_iter()
            .chain(project.iter())
            .chain(supplemental.iter())
        {
            content_type_parts.push((
                part.name.clone(),
                content_types
                    .part_content_type(&part.name)
                    .map(str::to_owned),
            ));
        }
        Ok(Self {
            host,
            source,
            project,
            supplemental,
            content_types,
            content_type_parts,
        })
    }

    fn same_selected_graph(&self, other: &Self) -> bool {
        self.host == other.host
            && self.source == other.source
            && self.project == other.project
            && self.supplemental == other.supplemental
            && self.content_type_parts == other.content_type_parts
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
    fn capture(
        package: &OpcPackage,
        source_part: &PackURI,
        host: Host,
        limits: SourceLimits,
    ) -> Result<Self> {
        let mut digest = Sha256::new();
        proof_field(&mut digest, host_name(host).as_bytes());
        proof_field(&mut digest, source_part.as_str().as_bytes());
        let source = admit_selected_part(package, source_part, limits)?;
        hash_proof_part(&mut digest, package, source, limits)?;
        let existing = inspect_existing_graph(package, source_part, host)?;
        let content_types = package.source_content_types(limits)?;
        proof_content_type(&mut digest, &content_types, source_part);
        match existing {
            Some(graph) => {
                digest.update([1]);
                let project = admit_selected_part(package, &graph.project_part, limits)?;
                hash_proof_part(&mut digest, package, project, limits)?;
                proof_content_type(&mut digest, &content_types, &graph.project_part);
                if let Some((supplemental_name, _)) = graph.supplemental {
                    digest.update([1]);
                    let supplemental = admit_selected_part(package, &supplemental_name, limits)?;
                    hash_proof_part(&mut digest, package, supplemental, limits)?;
                    proof_content_type(&mut digest, &content_types, &supplemental_name);
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

fn proof_content_type(digest: &mut Sha256, content_types: &OwnedContentTypes, part: &PackURI) {
    proof_field(digest, part.as_str().as_bytes());
    match content_types.part_content_type(part) {
        Some(content_type) => {
            digest.update([1]);
            proof_field(digest, content_type.as_bytes());
        },
        None => digest.update([0]),
    }
}

fn hash_proof_part(
    digest: &mut Sha256,
    package: &OpcPackage,
    part: &dyn Part,
    limits: SourceLimits,
) -> Result<()> {
    proof_field(digest, part.partname().as_str().as_bytes());
    proof_field(digest, part.content_type().as_bytes());
    proof_field(digest, part.blob());
    let relationships = package.source_relationships(part.partname(), limits)?;
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
        let source = package.part_metadata(source_part).ok_or_else(|| {
            Error::Opc(litchi_opc::OpcError::PartNotFound(source_part.to_string()))
        })?;
        host.macro_content_type(source.content_type())?
    };
    let limits = options.effective(package).source_limits;
    let _ = admit_selected_part(package, source_part, limits)?;
    if host == Host::Word && word_supplemental_xml.is_none() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "store_vba_project",
            "Word VBA project storage requires supplemental-data XML",
        ));
    }
    if host != Host::Word && word_supplemental_xml.is_some() {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "store_vba_project",
            "VBA supplemental-data XML is only valid for Word",
        ));
    }

    let source_content_types = package.source_content_types(limits)?;
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
    let replacement_content_types = content_types_after_graph_store(
        &source_content_types,
        source_part,
        host,
        existing.as_ref(),
        macro_content_type,
        limits,
    )?;
    if let Some(graph) = existing {
        let expected = package.source_relationships(source_part, limits)?;
        let replacement = expected
            .without_relationship(&graph.relationship_id, limits.max_relationship_xml_bytes())?;
        package.try_replace_relationships(&expected, &replacement, limits)?;
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

    let expected = package.source_relationships(source_part, limits)?;
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
        limits.max_relationship_xml_bytes(),
    )?;
    package.try_replace_relationships(&expected, &replacement, limits)?;
    let source = package.get_part_mut(source_part)?;
    source.set_content_type(macro_content_type.to_string())?;
    let current_content_types = package.source_content_types(limits)?;
    replacement_content_types
        .preflight_under(limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;
    package.try_replace_content_types(
        current_content_types.bytes(),
        &replacement_content_types,
        limits,
    )?;
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
    limits: SourceLimits,
) -> Result<bool> {
    let source = package
        .part_metadata(source_part)
        .ok_or_else(|| Error::Opc(litchi_opc::OpcError::PartNotFound(source_part.to_string())))?;
    let macro_content_type = host.macro_content_type(source.content_type())?;
    if source.content_type() != macro_content_type {
        return Ok(false);
    }
    let Some(graph) = inspect_existing_graph(package, source_part, host)? else {
        return Ok(false);
    };
    let project = admit_selected_part(package, &graph.project_part, limits)?;
    if project.blob() != project_payload {
        return Ok(false);
    }
    match (host, supplemental_xml, graph.supplemental.as_ref()) {
        (Host::Word, Some(expected), Some((part, _))) => {
            Ok(admit_selected_part(package, part, limits)?.blob() == expected)
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
    limits: SourceLimits,
) -> Result<bool> {
    let non_macro_content_type = {
        let source = package.part_metadata(source_part).ok_or_else(|| {
            Error::Opc(litchi_opc::OpcError::PartNotFound(source_part.to_string()))
        })?;
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
    remove_project_graph_in_place(package, source_part, graph, non_macro_content_type, limits)?;
    Ok(true)
}

fn remove_project_graph_in_place(
    package: &mut OpcPackage,
    source_part: &PackURI,
    graph: ExistingGraph,
    non_macro_content_type: &str,
    limits: SourceLimits,
) -> Result<()> {
    let _ = admit_selected_part(package, source_part, limits)?;
    let source_content_types = package.source_content_types(limits)?;
    let replacement_content_types = content_types_after_graph_removal(
        &source_content_types,
        source_part,
        &graph,
        non_macro_content_type,
        limits,
    )?;
    let expected = package.source_relationships(source_part, limits)?;
    let replacement = expected
        .without_relationship(&graph.relationship_id, limits.max_relationship_xml_bytes())?;
    package.try_replace_relationships(&expected, &replacement, limits)?;
    let source = package.get_part_mut(source_part)?;
    source.set_content_type(non_macro_content_type.to_string())?;
    package.remove_part(&graph.project_part);
    if let Some((supplemental, _)) = graph.supplemental {
        package.remove_part(&supplemental);
    }
    let current_content_types = package.source_content_types(limits)?;
    replacement_content_types
        .preflight_under(limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;
    package.try_replace_content_types(
        current_content_types.bytes(),
        &replacement_content_types,
        limits,
    )?;
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
    base_admission: Option<RecipientAdmission>,
    source_part: PackURI,
    host: Host,
    options: EditOptions,
}

impl std::fmt::Debug for Edit<'_> {
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
            base_admission: None,
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
                self.options.source_limits,
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
        self.store_inner(project_payload, word_supplemental_xml)
            .map_err(map_source_limit_error)
    }

    fn store_inner(
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
        let current_limits = self.options.effective(self.current()).source_limits;
        let admission = validate_recipient_metadata_limits(self.current(), current_limits)?;
        if self.candidate.is_none() && self.base_admission.is_none() {
            self.base_admission = Some(admission);
        }
        if store_is_noop(
            self.current(),
            &self.source_part,
            self.host,
            &project_payload,
            word_supplemental_xml.as_deref(),
            self.options.source_limits,
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
        self.remove_inner().map_err(map_source_limit_error)
    }

    fn remove_inner(&mut self) -> Result<bool> {
        let current_limits = self.options.effective(self.current()).source_limits;
        let admission = validate_recipient_metadata_limits(self.current(), current_limits)?;
        if self.candidate.is_none() && self.base_admission.is_none() {
            self.base_admission = Some(admission);
        }
        if inspect_existing_graph(self.current(), &self.source_part, self.host)?.is_none() {
            return Ok(false);
        }
        self.capture_before()?;
        let mut candidate = self.current().clone();
        let changed = remove_project_graph_on_candidate(
            &mut candidate,
            &self.source_part,
            self.host,
            self.options.effective(self.current()).source_limits,
        )?;
        self.candidate = Some(candidate);
        Ok(changed)
    }

    /// Discard every staged graph change.
    pub fn reset(&mut self) {
        self.candidate = None;
        self.before = None;
        self.base_admission = None;
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
        self.commit_inner().map_err(map_source_limit_error)
    }

    fn commit_inner(self) -> Result<Commit> {
        let Some(candidate) = self.candidate else {
            let patch = Patch {
                before: None,
                after: None,
                proof: Some(GraphProof::capture(
                    self.source,
                    &self.source_part,
                    self.host,
                    self.options.source_limits,
                )?),
                options: self.options,
            };
            return Ok(Commit::new(
                self.source.clone(),
                patch,
                Diagnostics::UNCHANGED,
            ));
        };
        let before = self.before.ok_or(Error::Internal(
            "VBA edit candidate has no captured graph base",
        ))?;
        let after = GraphSnapshot::capture(
            &candidate,
            &self.source_part,
            self.host,
            self.options.source_limits,
        )?;
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
            .ok_or(Error::Internal("changed VBA patch has no graph base"))?;
        let after = patch
            .after
            .as_ref()
            .ok_or(Error::Internal("changed VBA patch has no graph target"))?;
        let admission = self.base_admission.ok_or(Error::Internal(
            "VBA edit candidate has no recipient admission summary",
        ))?;
        validate_snapshot_limits(
            self.source,
            before,
            after,
            self.options,
            &after.content_types,
            &admission,
        )?;
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
        self.apply_inner(package).map_err(map_source_limit_error)
    }

    fn apply_inner(&self, package: &OpcPackage) -> Result<Commit> {
        let Some(before) = self.before.as_ref() else {
            let proof = self
                .proof
                .as_ref()
                .ok_or(Error::Internal("VBA patch has no replay base"))?;
            let recipient_limits = self.options.effective(package).source_limits;
            let _admission = validate_recipient_metadata_limits(package, recipient_limits)?;
            let current =
                GraphProof::capture(package, &proof.source_part, proof.host, recipient_limits)?;
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
            .ok_or(Error::Internal("VBA patch has no target graph"))?;
        let recipient_limits = self.options.effective(package).source_limits;
        let admission = validate_recipient_metadata_limits(package, recipient_limits)?;
        let current =
            GraphSnapshot::capture(package, &before.source.name, before.host, recipient_limits)?;
        if !current.same_selected_graph(before) {
            return Err(Error::PatchConflict);
        }
        let preserve_content_types = current.content_types.bytes() == before.content_types.bytes();
        let projected_content_types = if preserve_content_types {
            after.content_types.clone()
        } else {
            let existing = inspect_existing_graph(package, &after.source.name, after.host)?;
            content_types_after_graph_replace(
                &current.content_types,
                after,
                existing.as_ref(),
                recipient_limits,
            )?
        };
        validate_snapshot_limits(
            package,
            &current,
            after,
            self.options,
            &projected_content_types,
            &admission,
        )?;
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
        restore_graph_in_place(
            &mut candidate,
            after,
            preserve_content_types,
            self.options.effective(package).source_limits,
        )?;
        let resulting =
            GraphSnapshot::capture(&candidate, &after.source.name, after.host, recipient_limits)?;
        if !resulting.same_selected_graph(after)
            || (preserve_content_types
                && resulting.content_types.bytes() != after.content_types.bytes())
        {
            return Err(Error::Internal(
                "VBA patch publication changed the selected graph",
            ));
        }
        Ok(Commit::new(candidate, self.clone(), Diagnostics::CHANGED))
    }
}

/// The published VBA graph commit triad.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch>;

/// Minimal content-free commit diagnostics.
pub use litchi_core::edit::Diagnostics;

#[derive(Debug, Clone, Copy)]
struct RelationshipTotals {
    xml_bytes: u64,
    events: u64,
    relationships: u64,
    members: usize,
}

/// Metadata-only admission calculated once for a recipient state.
///
/// Staging and publication both need the same physical-source, part, and
/// relationship totals. Carrying this compact summary across that boundary
/// avoids rescanning the retained central directory and serializing every
/// relationship member a second time.
#[derive(Debug, Clone)]
struct RecipientAdmission {
    retained: Option<RetainedArchiveMetrics>,
    parts: usize,
    part_bytes: u64,
    metadata_bytes: u64,
    relationships: RelationshipTotals,
    physical_relationship_members: usize,
}

#[derive(Clone, Copy)]
struct RelationshipMetrics {
    xml_bytes: u64,
    events: u64,
    relationships: u64,
    member_present: bool,
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

/// Admit one selected graph payload from its retained source metadata before
/// asking the OPC package to decode it. A deferred package must not inflate a
/// payload under the wider open-time policy and discover a tighter edit policy
/// only afterwards. The second check also covers a decoder-reported length
/// that differs from its central-directory hint.
fn admit_selected_part<'a>(
    package: &'a OpcPackage,
    part_name: &PackURI,
    limits: SourceLimits,
) -> Result<&'a dyn Part> {
    let metadata = package
        .part_metadata(part_name)
        .ok_or_else(|| Error::Opc(litchi_opc::OpcError::PartNotFound(part_name.to_string())))?;
    let declared = u64::try_from(metadata.payload_len()?).unwrap_or(u64::MAX);
    check_vba_limit(LimitKind::PartBytes, declared, limits.max_part_bytes())?;
    check_vba_limit(LimitKind::EntryBytes, declared, limits.max_entry_bytes())?;
    let part = package.get_part(part_name)?;
    let actual = u64::try_from(part.blob().len()).unwrap_or(u64::MAX);
    check_vba_limit(LimitKind::PartBytes, actual, limits.max_part_bytes())?;
    check_vba_limit(LimitKind::EntryBytes, actual, limits.max_entry_bytes())?;
    Ok(part)
}

fn map_source_limit_error(error: Error) -> Error {
    let Error::Opc(litchi_opc::OpcError::Limit {
        kind,
        observed,
        maximum,
    }) = error
    else {
        return error;
    };
    let kind = match kind {
        litchi_opc::SourceLimitKind::PartBytes => LimitKind::PartBytes,
        litchi_opc::SourceLimitKind::Parts => LimitKind::GraphParts,
        litchi_opc::SourceLimitKind::TotalPartBytes => LimitKind::TotalPartBytes,
        litchi_opc::SourceLimitKind::ContentTypesBytes => LimitKind::ContentTypesBytes,
        litchi_opc::SourceLimitKind::ContentTypeMappings => LimitKind::ContentTypeMappings,
        litchi_opc::SourceLimitKind::RelationshipParts => LimitKind::RelationshipParts,
        litchi_opc::SourceLimitKind::RelationshipXmlBytes => LimitKind::RelationshipXmlBytes,
        litchi_opc::SourceLimitKind::TotalRelationshipXmlBytes => {
            LimitKind::TotalRelationshipXmlBytes
        },
        litchi_opc::SourceLimitKind::RelationshipsPerPart => LimitKind::RelationshipsPerPart,
        litchi_opc::SourceLimitKind::TotalRelationships => LimitKind::TotalRelationships,
        litchi_opc::SourceLimitKind::Events => LimitKind::XmlEvents,
        litchi_opc::SourceLimitKind::TotalRelationshipXmlEvents => {
            LimitKind::TotalRelationshipXmlEvents
        },
        litchi_opc::SourceLimitKind::Depth => LimitKind::XmlDepth,
        litchi_opc::SourceLimitKind::XmlAttributeBytes => LimitKind::XmlAttributeBytes,
        litchi_opc::SourceLimitKind::RelationshipTargetBytes => LimitKind::RelationshipTargetBytes,
        litchi_opc::SourceLimitKind::Members => LimitKind::ArchiveMembers,
        litchi_opc::SourceLimitKind::ArchiveTotalEntries => LimitKind::ArchiveTotalEntries,
        litchi_opc::SourceLimitKind::ArchiveMemberNameBytes => LimitKind::ArchiveMemberNameBytes,
        litchi_opc::SourceLimitKind::ArchiveMetadataBytes => LimitKind::ArchiveMetadataBytes,
        litchi_opc::SourceLimitKind::ArchiveCompressedBytes => LimitKind::ArchiveCompressedBytes,
        litchi_opc::SourceLimitKind::EntryBytes => LimitKind::EntryBytes,
        litchi_opc::SourceLimitKind::AggregateBytes => LimitKind::AggregateBytes,
        _ => LimitKind::SourcePolicy,
    };
    Error::Limit {
        kind: crate::LimitKind::Vba(kind),
        observed: limit_observed(observed),
        maximum: limit_observed(maximum),
    }
}

fn authored_current_archive_metadata_bytes(package: &OpcPackage) -> Result<u64> {
    let mut total = member_name_metadata("[Content_Types].xml")?
        .checked_add(member_name_metadata("_rels/.rels")?)
        .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    for metadata in package.iter_parts() {
        total = total
            .checked_add(member_name_metadata(metadata.partname().membername())?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        if !metadata.rels().is_empty() {
            let relationship_name = metadata
                .partname()
                .rels_uri()
                .map_err(|error| Error::Uri(error.clone()))?;
            total = total
                .checked_add(member_name_metadata(relationship_name.membername())?)
                .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        }
    }
    for member in package.non_part_members() {
        total = total
            .checked_add(member_name_metadata(member.name())?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    }
    Ok(total)
}

/// Re-admit the recipient's retained archive and OPC metadata before any
/// selected VBA payload is decoded.  Patch replay may replace only a small
/// graph, but its writer still carries every retained source member forward;
/// checking those physical and metadata-only dimensions first prevents an
/// unrelated oversized member from forcing selected-graph hydration.
fn validate_recipient_metadata_limits(
    package: &OpcPackage,
    limits: SourceLimits,
) -> Result<RecipientAdmission> {
    let mut parts = 0usize;
    let mut part_bytes = 0u64;
    for metadata in package.iter_parts() {
        parts = parts.saturating_add(1);
        check_vba_limit(
            LimitKind::GraphParts,
            parts as u64,
            limits.max_parts() as u64,
        )?;
        check_vba_limit(
            LimitKind::ArchiveMemberNameBytes,
            u64::try_from(metadata.partname().membername().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
        let relationship_name = metadata
            .partname()
            .rels_uri()
            .map_err(|error| Error::Uri(error.clone()))?;
        check_vba_limit(
            LimitKind::ArchiveMemberNameBytes,
            u64::try_from(relationship_name.membername().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
        let bytes = u64::try_from(metadata.payload_len()?).unwrap_or(u64::MAX);
        check_vba_limit(LimitKind::PartBytes, bytes, limits.max_part_bytes())?;
        check_vba_limit(LimitKind::EntryBytes, bytes, limits.max_entry_bytes())?;
        part_bytes = part_bytes.checked_add(bytes).ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalPartBytes),
            observed: usize::MAX,
            maximum: limit_observed(limits.max_total_part_bytes()),
        })?;
    }

    // Re-admit the retained archive only after all modeled part metadata has
    // established the VBA per-part precedence.  The archive pass remains
    // metadata-only and supplies the opaque/member/aggregate totals below;
    // it does not decode or rescan modeled payloads.
    let retained = package
        .retained_archive_metrics(limits)
        .map_err(Error::from)?;

    // Check retained part metadata before parsing [Content_Types].xml.  This
    // keeps a tighter selected-part policy deterministic for deferred source
    // packages: the first physical payload violation is reported as
    // `PartBytes` without hydrating the graph or making manifest parsing win
    // merely because it happens to be the first metadata helper called.
    let content_types = package.source_content_types(limits)?;
    let content_type_bytes = u64::try_from(content_types.bytes().len()).unwrap_or(u64::MAX);
    check_vba_limit(
        LimitKind::ContentTypesBytes,
        content_type_bytes,
        limits.max_content_types_bytes() as u64,
    )?;
    check_vba_limit(
        LimitKind::ContentTypeMappings,
        u64::try_from(content_types.mapping_count()).unwrap_or(u64::MAX),
        limits.max_content_type_mappings() as u64,
    )?;
    content_types
        .preflight_under(limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;
    check_vba_limit(
        LimitKind::TotalPartBytes,
        part_bytes,
        limits.max_total_part_bytes(),
    )?;

    check_vba_limit(
        LimitKind::ArchiveMemberNameBytes,
        u64::try_from("[Content_Types].xml".len()).unwrap_or(u64::MAX),
        limits.max_archive_member_name_bytes(),
    )?;
    check_vba_limit(
        LimitKind::ArchiveMemberNameBytes,
        u64::try_from("_rels/.rels".len()).unwrap_or(u64::MAX),
        limits.max_archive_member_name_bytes(),
    )?;
    for member in package.non_part_members() {
        check_vba_limit(
            LimitKind::ArchiveMemberNameBytes,
            u64::try_from(member.name().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
    }

    let relationships = relationship_totals(package, limits)?;
    check_vba_limit(
        LimitKind::RelationshipParts,
        relationships.members as u64,
        limits.max_relationship_parts() as u64,
    )?;
    check_vba_limit(
        LimitKind::TotalRelationshipXmlBytes,
        relationships.xml_bytes,
        limits.max_total_relationship_xml_bytes() as u64,
    )?;
    check_vba_limit(
        LimitKind::TotalRelationshipXmlEvents,
        relationships.events,
        limits.max_total_relationship_xml_events() as u64,
    )?;
    check_vba_limit(
        LimitKind::TotalRelationships,
        relationships.relationships,
        limits.max_total_relationships() as u64,
    )?;

    let metadata_bytes = if let Some(metrics) = retained.as_ref() {
        metrics.metadata_bytes()
    } else {
        authored_current_archive_metadata_bytes(package)?
    };
    check_vba_limit(
        LimitKind::ArchiveMetadataBytes,
        metadata_bytes,
        limits.max_archive_metadata_bytes(),
    )?;

    if let Some(metrics) = retained.as_ref() {
        let relationship_members = relationships.members;
        let physical_members = metrics
            .opaque_members()
            // `relationship_members` already includes the package-level
            // `_rels/.rels` member; add only the content-types member here.
            .saturating_add(1)
            .saturating_add(parts)
            .saturating_add(relationship_members);
        check_vba_limit(
            LimitKind::ArchiveMembers,
            physical_members as u64,
            limits.max_members() as u64,
        )?;
        check_vba_limit(
            LimitKind::ArchiveTotalEntries,
            metrics
                .opaque_entries()
                .saturating_add(1)
                .saturating_add(parts)
                .saturating_add(relationship_members) as u64,
            limits.max_archive_total_entries() as u64,
        )?;
        let aggregate = metrics
            .opaque_bytes()
            .saturating_add(part_bytes)
            .saturating_add(relationships.xml_bytes)
            .saturating_add(content_type_bytes);
        check_vba_limit(
            LimitKind::AggregateBytes,
            aggregate,
            limits.max_aggregate_bytes(),
        )?;
    } else {
        let non_part_files = package
            .non_part_members()
            .iter()
            .filter(|member| !member.name().ends_with('/'))
            .count();
        let non_part_entries = package.non_part_members().len();
        let physical_members = 1usize
            .saturating_add(parts)
            .saturating_add(relationships.members)
            .saturating_add(non_part_files);
        check_vba_limit(
            LimitKind::ArchiveMembers,
            physical_members as u64,
            limits.max_members() as u64,
        )?;
        check_vba_limit(
            LimitKind::ArchiveTotalEntries,
            physical_members.saturating_add(non_part_entries.saturating_sub(non_part_files)) as u64,
            limits.max_archive_total_entries() as u64,
        )?;
        let aggregate = part_bytes
            .saturating_add(relationships.xml_bytes)
            .saturating_add(content_type_bytes);
        check_vba_limit(
            LimitKind::AggregateBytes,
            aggregate,
            limits.max_aggregate_bytes(),
        )?;
    }
    Ok(RecipientAdmission {
        retained,
        parts,
        part_bytes,
        metadata_bytes,
        relationships,
        physical_relationship_members: relationships.members,
    })
}

fn relationship_totals(package: &OpcPackage, limits: SourceLimits) -> Result<RelationshipTotals> {
    let mut totals = RelationshipTotals {
        xml_bytes: 0,
        events: 0,
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
    let member_present = owner.as_str() == "/"
        || !relationships.is_empty()
        || package.has_retained_relationship_member(owner);
    let metrics = if member_present {
        relationship_token_metrics(
            &package
                .source_relationships(owner, limits)
                .map_err(|error| map_source_limit_error(Error::from(error)))?,
            limits,
        )?
    } else {
        RelationshipMetrics {
            xml_bytes: 0,
            events: 0,
            relationships: 0,
            member_present: false,
        }
    };
    totals.xml_bytes = totals
        .xml_bytes
        .checked_add(metrics.xml_bytes)
        .ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalRelationshipXmlBytes),
            observed: usize::MAX,
            maximum: limits.max_total_relationship_xml_bytes(),
        })?;
    totals.events = totals
        .events
        .checked_add(metrics.events)
        .ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalRelationshipXmlEvents),
            observed: usize::MAX,
            maximum: limits.max_total_relationship_xml_events(),
        })?;
    totals.relationships = totals
        .relationships
        .checked_add(metrics.relationships)
        .ok_or_else(|| Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::TotalRelationships),
            observed: usize::MAX,
            maximum: limits.max_total_relationships(),
        })?;
    totals.members = totals
        .members
        .saturating_add(usize::from(metrics.member_present));
    Ok(())
}

fn relationship_token_metrics(
    token: &OwnedRelationships,
    limits: SourceLimits,
) -> Result<RelationshipMetrics> {
    if !token.member_present() {
        return Ok(RelationshipMetrics {
            xml_bytes: 0,
            events: 0,
            relationships: 0,
            member_present: false,
        });
    }
    let plan = token
        .plan_edit(&[], &[], limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;
    let xml_bytes = u64::try_from(plan.final_len()).unwrap_or(u64::MAX);
    let events = u64::try_from(plan.event_count()).unwrap_or(u64::MAX);
    let relationships = u64::try_from(plan.relationship_count()).unwrap_or(u64::MAX);
    check_vba_limit(
        LimitKind::RelationshipXmlBytes,
        xml_bytes,
        limits.max_relationship_xml_bytes() as u64,
    )?;
    check_vba_limit(LimitKind::PartBytes, xml_bytes, limits.max_part_bytes())?;
    check_vba_limit(LimitKind::EntryBytes, xml_bytes, limits.max_entry_bytes())?;
    check_vba_limit(LimitKind::XmlEvents, events, limits.max_events() as u64)?;
    check_vba_limit(
        LimitKind::RelationshipsPerPart,
        relationships,
        limits.max_relationships_per_part() as u64,
    )?;
    Ok(RelationshipMetrics {
        xml_bytes,
        events,
        relationships,
        member_present: true,
    })
}

fn graph_part_states(graph: &GraphSnapshot) -> impl Iterator<Item = &PartState> {
    std::iter::once(&graph.source)
        .chain(graph.project.iter())
        .chain(graph.supplemental.iter())
}

fn graph_relationship_totals(
    graph: &GraphSnapshot,
    limits: SourceLimits,
) -> Result<RelationshipTotals> {
    let mut totals = RelationshipTotals {
        xml_bytes: 0,
        events: 0,
        relationships: 0,
        members: 0,
    };
    for part in graph_part_states(graph) {
        let metrics = relationship_token_metrics(&part.relationships_source, limits)?;
        totals.xml_bytes = totals.xml_bytes.saturating_add(metrics.xml_bytes);
        totals.events = totals.events.saturating_add(metrics.events);
        totals.relationships = totals.relationships.saturating_add(metrics.relationships);
        totals.members = totals
            .members
            .saturating_add(usize::from(metrics.member_present));
    }
    Ok(totals)
}

fn validate_archive_structural_limits(
    package: &OpcPackage,
    current: &GraphSnapshot,
    target: &GraphSnapshot,
    current_graph_relationships: RelationshipTotals,
    target_graph_relationships: RelationshipTotals,
    projected_parts: usize,
    projected_part_bytes: u64,
    projected_relationship_bytes: u64,
    projected_content_types: &OwnedContentTypes,
    limits: SourceLimits,
    admission: &RecipientAdmission,
) -> Result<()> {
    // Re-admit every retained physical entry first. This checks opaque ZIP
    // payloads and names that the OPC projection intentionally does not expose
    // without inflating them. The selected graph's prospective aggregate is
    // checked below after replacing its known members.
    let projected_physical_relationship_parts = admission
        .physical_relationship_members
        .checked_sub(current_graph_relationships.members)
        .ok_or(Error::Internal(
            "selected graph relationship members exceed the recipient physical namespace",
        ))?
        .saturating_add(target_graph_relationships.members);
    let projected_content_type_bytes =
        u64::try_from(projected_content_types.bytes().len()).unwrap_or(u64::MAX);
    if let Some(metrics) = admission.retained.as_ref() {
        // The targeted writer always has one content-types member. The
        // projected relationship count already includes the package-level
        // `.rels` member; everything else comes from the current known part
        // map and immutable opaque source entries. No historical source count
        // is used to infer a relationship member.
        let physical_members = metrics
            .opaque_members()
            // `projected_physical_relationship_parts` retains the package
            // root relationship member, so only content types is added here.
            .saturating_add(1)
            .saturating_add(projected_parts)
            .saturating_add(projected_physical_relationship_parts);
        check_vba_limit(
            LimitKind::ArchiveMembers,
            physical_members as u64,
            limits.max_members() as u64,
        )?;
        check_vba_limit(
            LimitKind::ArchiveTotalEntries,
            metrics
                .opaque_entries()
                .saturating_add(1)
                .saturating_add(projected_parts)
                .saturating_add(projected_physical_relationship_parts) as u64,
            limits.max_archive_total_entries() as u64,
        )?;
        let projected_aggregate = metrics
            .opaque_bytes()
            .saturating_add(projected_part_bytes)
            .saturating_add(projected_relationship_bytes)
            .saturating_add(projected_content_type_bytes);
        check_vba_limit(
            LimitKind::AggregateBytes,
            projected_aggregate,
            limits.max_aggregate_bytes(),
        )?;
        check_vba_limit(
            LimitKind::ArchiveMetadataBytes,
            projected_archive_metadata_bytes(
                package,
                current,
                target,
                projected_content_types,
                metrics,
                admission.metadata_bytes,
            )?,
            limits.max_archive_metadata_bytes(),
        )?;
    } else {
        // Authored packages have no retained physical directory. Their
        // publication shape is the ordinary OPC projection plus any opaque
        // member names recorded by an ingress adapter.
        let non_part_files = package
            .non_part_members()
            .iter()
            .filter(|member| !member.name().ends_with('/'))
            .count();
        let physical_members = 1_u64
            .saturating_add(u64::try_from(projected_parts).unwrap_or(u64::MAX))
            .saturating_add(
                u64::try_from(projected_physical_relationship_parts).unwrap_or(u64::MAX),
            )
            .saturating_add(u64::try_from(non_part_files).unwrap_or(u64::MAX));
        check_vba_limit(
            LimitKind::ArchiveMembers,
            physical_members,
            limits.max_members() as u64,
        )?;
        let archive_entries = physical_members.saturating_add(
            u64::try_from(
                package
                    .non_part_members()
                    .iter()
                    .filter(|member| member.name().ends_with('/'))
                    .count(),
            )
            .unwrap_or(u64::MAX),
        );
        check_vba_limit(
            LimitKind::ArchiveTotalEntries,
            archive_entries,
            limits.max_archive_total_entries() as u64,
        )?;
        check_vba_limit(
            LimitKind::AggregateBytes,
            projected_part_bytes
                .saturating_add(projected_content_type_bytes)
                .saturating_add(projected_relationship_bytes),
            limits.max_aggregate_bytes(),
        )?;
        check_vba_limit(
            LimitKind::ArchiveMetadataBytes,
            authored_archive_metadata_bytes(package, current, target, projected_content_types)?,
            limits.max_archive_metadata_bytes(),
        )?;
    }
    for part in graph_part_states(target) {
        check_vba_limit(
            LimitKind::ArchiveMemberNameBytes,
            u64::try_from(part.name.membername().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
        let relationship_name = part
            .name
            .rels_uri()
            .map_err(|error| Error::Uri(error.clone()))?;
        check_vba_limit(
            LimitKind::ArchiveMemberNameBytes,
            u64::try_from(relationship_name.membername().len()).unwrap_or(u64::MAX),
            limits.max_archive_member_name_bytes(),
        )?;
    }
    check_vba_limit(
        LimitKind::ArchiveMemberNameBytes,
        u64::try_from("[Content_Types].xml".len()).unwrap_or(u64::MAX),
        limits.max_archive_member_name_bytes(),
    )?;
    // Recipient-wide part, relationship, opaque-member, and payload checks
    // were completed by `validate_recipient_metadata_limits` and are carried
    // in `admission`. Only graph entries changed by this publication are
    // checked above; repeating the package scan here would rewalk and may
    // hydrate unrelated deferred payloads.
    Ok(())
}

fn member_name_metadata(name: &str) -> Result<u64> {
    u64::try_from(name.len())
        .map_err(|_| Error::Internal("OPC generated member-name metadata overflows"))
}

fn source_or_generated_part_metadata(
    package: &OpcPackage,
    metrics: &RetainedArchiveMetrics,
    part: &PartState,
) -> Result<u64> {
    let generated = member_name_metadata(part.name.membername())?;
    if package.source_part_member_retained(part.part.as_ref()) {
        Ok(metrics
            .source_member_metadata(part.name.membername())
            .unwrap_or(generated))
    } else {
        Ok(generated)
    }
}

fn source_or_generated_relationship_metadata(
    package: &OpcPackage,
    metrics: &RetainedArchiveMetrics,
    part: &PartState,
) -> Result<u64> {
    if !part.relationships_source.member_present() {
        return Ok(0);
    }
    let relationship_name = part
        .name
        .rels_uri()
        .map_err(|error| Error::Uri(error.clone()))?;
    let generated = member_name_metadata(relationship_name.membername())?;
    if package.source_relationship_member_retained(&part.name, part.part.rels())? {
        Ok(metrics
            .source_member_metadata(relationship_name.membername())
            .unwrap_or(generated))
    } else {
        Ok(generated)
    }
}

fn content_types_metadata(
    package: &OpcPackage,
    metrics: &RetainedArchiveMetrics,
    bytes: &[u8],
) -> Result<u64> {
    let generated = member_name_metadata("[Content_Types].xml")?;
    if package.source_content_types_member_retained(bytes) {
        Ok(metrics
            .source_member_metadata("[Content_Types].xml")
            .unwrap_or(generated))
    } else {
        Ok(generated)
    }
}

fn selected_graph_metadata(
    package: &OpcPackage,
    metrics: &RetainedArchiveMetrics,
    graph: &GraphSnapshot,
) -> Result<u64> {
    let mut total = source_or_generated_part_metadata(package, metrics, &graph.source)?;
    total = total
        .checked_add(source_or_generated_relationship_metadata(
            package,
            metrics,
            &graph.source,
        )?)
        .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    for part in graph.project.iter().chain(graph.supplemental.iter()) {
        total = total
            .checked_add(source_or_generated_part_metadata(package, metrics, part)?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        total = total
            .checked_add(source_or_generated_relationship_metadata(
                package, metrics, part,
            )?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    }
    Ok(total)
}

fn projected_archive_metadata_bytes(
    package: &OpcPackage,
    current: &GraphSnapshot,
    target: &GraphSnapshot,
    projected_content_types: &OwnedContentTypes,
    metrics: &RetainedArchiveMetrics,
    current_metadata: u64,
) -> Result<u64> {
    let current_selected = selected_graph_metadata(package, metrics, current)?;
    let target_selected = selected_graph_metadata(package, metrics, target)?;
    let current_manifest = content_types_metadata(package, metrics, current.content_types.bytes())?;
    let target_manifest =
        content_types_metadata(package, metrics, projected_content_types.bytes())?;
    current_metadata
        .checked_sub(current_selected)
        .and_then(|bytes| bytes.checked_sub(current_manifest))
        .and_then(|bytes| bytes.checked_add(target_selected))
        .and_then(|bytes| bytes.checked_add(target_manifest))
        .ok_or(Error::Internal("VBA archive metadata underflows"))
}

fn authored_archive_metadata_bytes(
    package: &OpcPackage,
    current: &GraphSnapshot,
    target: &GraphSnapshot,
    _projected_content_types: &OwnedContentTypes,
) -> Result<u64> {
    let mut total = member_name_metadata("[Content_Types].xml")?
        .checked_add(member_name_metadata("_rels/.rels")?)
        .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    for metadata in package.iter_parts() {
        total = total
            .checked_add(member_name_metadata(metadata.partname().membername())?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        if !metadata.rels().is_empty() {
            let relationship_name = metadata
                .partname()
                .rels_uri()
                .map_err(|error| Error::Uri(error.clone()))?;
            total = total
                .checked_add(member_name_metadata(relationship_name.membername())?)
                .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        }
    }
    for member in package.non_part_members() {
        total = total
            .checked_add(member_name_metadata(member.name())?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
    }
    let current_selected = selected_graph_metadata_authored(current)?;
    let target_selected = selected_graph_metadata_authored(target)?;
    let current_manifest = member_name_metadata("[Content_Types].xml")?;
    let target_manifest = current_manifest;
    Ok(total
        .saturating_sub(current_selected)
        .saturating_sub(current_manifest)
        .saturating_add(target_selected)
        .saturating_add(target_manifest))
}

fn selected_graph_metadata_authored(graph: &GraphSnapshot) -> Result<u64> {
    let mut total = 0u64;
    for part in graph_part_states(graph) {
        total = total
            .checked_add(member_name_metadata(part.name.membername())?)
            .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        if part.relationships_source.member_present() {
            let relationship_name = part
                .name
                .rels_uri()
                .map_err(|error| Error::Uri(error.clone()))?;
            total = total
                .checked_add(member_name_metadata(relationship_name.membername())?)
                .ok_or(Error::Internal("VBA archive metadata overflows"))?;
        }
    }
    Ok(total)
}

fn content_types_after_graph_store(
    source: &OwnedContentTypes,
    source_part: &PackURI,
    host: Host,
    existing: Option<&ExistingGraph>,
    macro_content_type: &str,
    limits: SourceLimits,
) -> Result<OwnedContentTypes> {
    let project_name =
        PackURI::new(host.project_part()).map_err(|error| Error::Uri(error.clone()))?;
    let supplemental_name = if host == Host::Word {
        Some(PackURI::new(WORD_SUPPLEMENTAL_PART).map_err(|error| Error::Uri(error.clone()))?)
    } else {
        None
    };

    // Keep explicit declarations whose names and resolved types remain valid.
    // Remove only declarations that describe a part being replaced or whose
    // type changes.  The host part's relationship override is intentionally
    // not removed: its physical `.rels` member survives the graph edit.
    let mut part_overrides_to_remove = Vec::new();
    if source.has_part_override(source_part)
        && source.part_content_type(source_part) != Some(macro_content_type)
    {
        part_overrides_to_remove.push(source_part.clone());
    }
    let mut removed_graph_parts = Vec::new();
    if let Some(graph) = existing {
        if graph.project_part != project_name {
            part_overrides_to_remove.push(graph.project_part.clone());
            removed_graph_parts.push(graph.project_part.clone());
        }
        if let Some((supplemental, _)) = &graph.supplemental
            && supplemental_name.as_ref() != Some(supplemental)
        {
            part_overrides_to_remove.push(supplemental.clone());
            removed_graph_parts.push(supplemental.clone());
        }
    }

    let mut base = source
        .without_part_overrides_only(&part_overrides_to_remove, limits.max_content_types_bytes())?;
    // A noncanonical old graph part disappears completely.  Its relationship
    // override must be removed as well, while the host relationship override
    // remains untouched in `base`.
    for part in &removed_graph_parts {
        base = base.without_parts(std::slice::from_ref(part), limits.max_content_types_bytes())?;
    }

    let mut additions = Vec::with_capacity(3);
    additions.push((source_part, macro_content_type));
    additions.push((&project_name, ct::OFC_VBA_PROJECT));
    if let Some(supplemental) = supplemental_name.as_ref() {
        additions.push((supplemental, ct::WML_VBA_DATA));
    }

    // If a noncanonical Word project carried an explicit relationship-member
    // override, carry that authored MIME value to the canonical replacement.
    // The ordinary `rels` default remains implicit and is never synthesized.
    let mut replacement_project_relationship_name = None;
    let mut replacement_project_relationship_type = None;
    if host == Host::Word
        && let Some(graph) = existing
        && graph.project_part != project_name
        && let Ok(old_relationship_name) = graph.project_part.rels_uri()
        && source.has_part_override(&old_relationship_name)
    {
        replacement_project_relationship_type = source
            .part_content_type(&old_relationship_name)
            .map(str::to_owned);
        if replacement_project_relationship_type.is_some() {
            replacement_project_relationship_name = project_name.rels_uri().ok();
        }
    }
    if let (Some(name), Some(content_type)) = (
        replacement_project_relationship_name.as_ref(),
        replacement_project_relationship_type.as_deref(),
    ) {
        additions.push((name, content_type));
    }
    compose_content_types_for_parts(&base, &additions, limits.max_content_types_bytes())
}

fn content_types_after_graph_removal(
    source: &OwnedContentTypes,
    source_part: &PackURI,
    graph: &ExistingGraph,
    non_macro_content_type: &str,
    limits: SourceLimits,
) -> Result<OwnedContentTypes> {
    let mut removed = Vec::with_capacity(2);
    removed.push(graph.project_part.clone());
    if let Some((supplemental, _)) = &graph.supplemental {
        removed.push(supplemental.clone());
    }
    // Removing a graph part also removes its relationship-member override.
    // The host relationship member remains present and source-backed.
    let mut base = source.without_parts(&removed, limits.max_content_types_bytes())?;
    if source.has_part_override(source_part)
        && source.part_content_type(source_part) != Some(non_macro_content_type)
    {
        base = base.without_part_overrides_only(
            std::slice::from_ref(source_part),
            limits.max_content_types_bytes(),
        )?;
    }
    compose_content_types_for_parts(
        &base,
        &[(source_part, non_macro_content_type)],
        limits.max_content_types_bytes(),
    )
}

fn content_types_after_graph_replace(
    source: &OwnedContentTypes,
    target: &GraphSnapshot,
    existing: Option<&ExistingGraph>,
    limits: SourceLimits,
) -> Result<OwnedContentTypes> {
    let target_names = graph_part_states(target)
        .map(|part| &part.name)
        .collect::<Vec<_>>();
    let mut part_overrides_to_remove = Vec::new();
    if source.has_part_override(&target.source.name)
        && source.part_content_type(&target.source.name)
            != Some(target.source.content_type.as_str())
    {
        part_overrides_to_remove.push(target.source.name.clone());
    }
    let mut removed_graph_parts = Vec::new();
    if let Some(graph) = existing {
        let old_parts = std::iter::once(&graph.project_part)
            .chain(graph.supplemental.iter().map(|(part, _)| part));
        for old_part in old_parts {
            if !target_names.contains(&old_part) {
                part_overrides_to_remove.push(old_part.clone());
                removed_graph_parts.push(old_part.clone());
            }
        }
    }
    let mut base = source
        .without_part_overrides_only(&part_overrides_to_remove, limits.max_content_types_bytes())?;
    for part in &removed_graph_parts {
        base = base.without_parts(std::slice::from_ref(part), limits.max_content_types_bytes())?;
    }

    let additions = graph_part_states(target)
        .filter(|part| base.part_content_type(&part.name) != Some(part.content_type.as_str()))
        .map(|part| (&part.name, part.content_type.as_str()))
        .collect::<Vec<_>>();
    compose_content_types_for_parts(&base, &additions, limits.max_content_types_bytes())
}

fn compose_content_types_for_parts(
    base: &OwnedContentTypes,
    parts: &[(&PackURI, &str)],
    maximum: usize,
) -> Result<OwnedContentTypes> {
    let mut overrides = Vec::with_capacity(parts.len());
    for (name, content_type) in parts {
        if base.part_content_type(name) != Some(*content_type) {
            overrides.push((*name, *content_type));
        }
    }
    base.with_part_overrides(&overrides, maximum)
        .map_err(Error::from)
}

fn validate_snapshot_limits(
    package: &OpcPackage,
    current: &GraphSnapshot,
    target: &GraphSnapshot,
    options: EditOptions,
    projected_content_types: &OwnedContentTypes,
    admission: &RecipientAdmission,
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
            LimitKind::EntryBytes,
            u64::try_from(part.data.len()).unwrap_or(u64::MAX),
            effective.source_limits.max_entry_bytes(),
        )?;
        relationship_token_metrics(&part.relationships_source, effective.source_limits)?;
    }
    check_vba_limit(
        LimitKind::ContentTypesBytes,
        u64::try_from(projected_content_types.bytes().len()).unwrap_or(u64::MAX),
        effective.source_limits.max_content_types_bytes() as u64,
    )?;
    check_vba_limit(
        LimitKind::ContentTypeMappings,
        u64::try_from(projected_content_types.mapping_count()).unwrap_or(u64::MAX),
        effective.source_limits.max_content_type_mappings() as u64,
    )?;
    projected_content_types
        .preflight_under(effective.source_limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;

    let current_part_bytes = graph_part_states(current)
        .map(|part| u64::try_from(part.data.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let target_part_bytes = graph_part_states(target)
        .map(|part| u64::try_from(part.data.len()).unwrap_or(u64::MAX))
        .sum::<u64>();
    let projected_part_bytes = admission
        .part_bytes
        .saturating_sub(current_part_bytes)
        .saturating_add(target_part_bytes);
    check_vba_limit(
        LimitKind::TotalPartBytes,
        projected_part_bytes,
        effective.source_limits.max_total_part_bytes(),
    )?;

    let current_graph_relationships = graph_relationship_totals(current, effective.source_limits)?;
    let target_graph_relationships = graph_relationship_totals(target, effective.source_limits)?;
    let removed_relationship_bytes = current_graph_relationships.xml_bytes;
    let added_relationship_bytes = target_graph_relationships.xml_bytes;
    let projected_relationship_bytes = admission
        .relationships
        .xml_bytes
        .saturating_sub(removed_relationship_bytes)
        .saturating_add(added_relationship_bytes);
    check_vba_limit(
        LimitKind::TotalRelationshipXmlBytes,
        projected_relationship_bytes,
        effective.source_limits.max_total_relationship_xml_bytes() as u64,
    )?;
    let removed_relationship_events = current_graph_relationships.events;
    let added_relationship_events = target_graph_relationships.events;
    let projected_relationship_events = admission
        .relationships
        .events
        .saturating_sub(removed_relationship_events)
        .saturating_add(added_relationship_events);
    check_vba_limit(
        LimitKind::TotalRelationshipXmlEvents,
        projected_relationship_events,
        effective.source_limits.max_total_relationship_xml_events() as u64,
    )?;
    let removed_relationships = current_graph_relationships.relationships;
    let added_relationships = target_graph_relationships.relationships;
    let projected_relationships = admission
        .relationships
        .relationships
        .saturating_sub(removed_relationships)
        .saturating_add(added_relationships);
    check_vba_limit(
        LimitKind::TotalRelationships,
        projected_relationships,
        effective.source_limits.max_total_relationships() as u64,
    )?;
    let projected_members = admission
        .relationships
        .members
        .saturating_sub(current_graph_relationships.members)
        .saturating_add(target_graph_relationships.members);
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
    let projected_parts = admission
        .parts
        .saturating_sub(removed)
        .saturating_add(graph_parts);
    if projected_parts > effective.source_limits.max_parts() {
        return Err(Error::Limit {
            kind: crate::LimitKind::Vba(LimitKind::GraphParts),
            observed: projected_parts,
            maximum: effective.source_limits.max_parts(),
        });
    }
    validate_archive_structural_limits(
        package,
        current,
        target,
        current_graph_relationships,
        target_graph_relationships,
        projected_parts,
        projected_part_bytes,
        projected_relationship_bytes,
        projected_content_types,
        effective.source_limits,
        admission,
    )?;
    Ok(())
}

fn restore_graph_in_place(
    package: &mut OpcPackage,
    target: &GraphSnapshot,
    preserve_content_types: bool,
    limits: SourceLimits,
) -> Result<()> {
    let source_part = &target.source.name;
    let existing = inspect_existing_graph(package, source_part, target.host)?;
    let source_content_types = package.source_content_types(limits)?;
    let replacement_content_types = if preserve_content_types {
        target.content_types.clone()
    } else {
        content_types_after_graph_replace(&source_content_types, target, existing.as_ref(), limits)?
    };
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

    let expected_content_types = package.source_content_types(limits)?;
    replacement_content_types
        .preflight_under(limits)
        .map_err(|error| map_source_limit_error(Error::from(error)))?;
    for part in graph_part_states(target) {
        part.relationships_source
            .plan_edit(&[], &[], limits)
            .map_err(|error| map_source_limit_error(Error::from(error)))?;
    }
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
    let source = package
        .part_metadata(source_part)
        .ok_or_else(|| Error::Opc(litchi_opc::OpcError::PartNotFound(source_part.to_string())))?;
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
    let project = package
        .part_metadata(&project_part)
        .ok_or_else(|| Error::Opc(litchi_opc::OpcError::PartNotFound(project_part.to_string())))?;
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
        let part = package
            .part_metadata(&part_name)
            .ok_or_else(|| Error::Opc(litchi_opc::OpcError::PartNotFound(part_name.to_string())))?;
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

fn invalid_main_content_type(_host: Host, _actual: &str) -> Error {
    Error::refused(
        litchi_core::edit::RefusalKind::UnsupportedTarget,
        "vba_project",
        "the selected host does not own the main part",
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackageWriter;

    const EXCEL_MAIN_XML: &[u8] =
        br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#;
    const WORD_MAIN_XML: &[u8] =
        br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#;

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
                br#"<owner/>"#.to_vec(),
            );
            owner.relate_to(target, "urn:test:dangling");

            let mut package = OpcPackage::new();
            package
                .try_add_part(Box::new(BlobPart::new(
                    source_name.clone(),
                    ct::WML_DOCUMENT_MAIN.to_string(),
                    WORD_MAIN_XML.to_vec(),
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
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                source_name.clone(),
                ct::SML_SHEET_MAIN.to_string(),
                EXCEL_MAIN_XML.to_vec(),
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
        assert!(matches!(store(&package), Err(Error::Signed)));
        assert!(package.is_signed());
        assert_eq!(package.part_count(), parts);
        assert_eq!(
            package.get_part(&source_name).unwrap().content_type(),
            ct::SML_SHEET_MAIN
        );

        package.unsign();
        package = store(&package).unwrap().into_package();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(matches!(
            Edit::new(&package, Host::Excel, EditOptions::new()).and_then(|mut edit| {
                edit.remove()?;
                edit.commit().map(|_| ())
            }),
            Err(Error::Signed)
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
                EXCEL_MAIN_XML.to_vec(),
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
                EXCEL_MAIN_XML.to_vec(),
            )))
            .unwrap();
        package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
        package
    }

    fn serialized_archive_counts(bytes: &[u8]) -> (usize, usize) {
        let archive = soapberry_zip::ZipArchive::from_slice(bytes).unwrap();
        let mut members = 0usize;
        let mut entries = 0usize;
        for entry in archive.entries() {
            let entry = entry.unwrap();
            entries = entries.saturating_add(1);
            if !entry.is_dir() {
                members = members.saturating_add(1);
            }
        }
        (members, entries)
    }

    fn add_directory_entry(bytes: &[u8]) -> Vec<u8> {
        let archive = soapberry_zip::office::ArchiveReader::new(
            bytes,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        let mut writer = soapberry_zip::ZipArchiveWriter::new(Vec::new());
        writer.new_dir("custom/").create().unwrap();
        for name in archive.file_names() {
            let payload = archive.read(name).unwrap();
            writer.write_stored_file(name, &payload).unwrap();
        }
        writer.finish().unwrap()
    }

    fn excel_source_package_with_existing_vba_and_directory_entry() -> OpcPackage {
        let package = excel_package();
        let mut initial = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        initial.store(b"source".to_vec(), None).unwrap();
        let stored = initial.commit().unwrap().into_package();
        let authored = PackageWriter::to_bytes(&stored).unwrap();
        let source = add_directory_entry(&authored);
        OpcPackage::from_vec(source, SourceLimits::default()).unwrap()
    }

    fn word_package() -> OpcPackage {
        let source_name = PackURI::new("/word/document.xml").unwrap();
        let mut package = OpcPackage::new();
        package
            .try_add_part(Box::new(BlobPart::new(
                source_name,
                ct::WML_DOCUMENT_MAIN.to_string(),
                WORD_MAIN_XML.to_vec(),
            )))
            .unwrap();
        package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
        package
    }

    // This fixture starts with one authored workbook override.  The expected
    // Excel VBA edit is deliberately spelled out here instead of reusing the
    // production content-types composer: remove that known plain-main
    // declaration and insert the required macro-main and VBA-project
    // declarations at the manifest close, leaving every other source byte
    // untouched.
    fn expected_excel_vba_content_types(source: &[u8]) -> Vec<u8> {
        const PLAIN_WORKBOOK_OVERRIDE: &[u8] = br#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"#;
        const VBA_WORKBOOK_OVERRIDES: &[u8] = br#"<Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>"#;
        let old_offset = source
            .windows(PLAIN_WORKBOOK_OVERRIDE.len())
            .position(|window| window == PLAIN_WORKBOOK_OVERRIDE)
            .expect("plain Excel workbook content-type override");
        assert_eq!(
            source
                .windows(PLAIN_WORKBOOK_OVERRIDE.len())
                .filter(|window| *window == PLAIN_WORKBOOK_OVERRIDE)
                .count(),
            1,
            "Excel fixture must contain one plain workbook override"
        );
        let mut expected = source.to_vec();
        expected.drain(old_offset..old_offset + PLAIN_WORKBOOK_OVERRIDE.len());
        let close = b"</Types>";
        let close_offset = expected
            .windows(close.len())
            .position(|window| window == close)
            .expect("content-types close tag");
        expected.splice(
            close_offset..close_offset,
            VBA_WORKBOOK_OVERRIDES.iter().copied(),
        );
        expected
    }

    fn add_noncanonical_content_types(package: &mut OpcPackage) -> Vec<u8> {
        let limits = package.read_limits();
        let expected = package.source_content_types(limits).unwrap();
        let mut noncanonical = expected.bytes().to_vec();
        let marker = b"<Override";
        let offset = noncanonical
            .windows(marker.len())
            .position(|window| window == marker)
            .expect("content-types override");
        noncanonical.splice(
            offset..offset,
            b"\n  <!-- retained VBA manifest spelling -->\n"
                .iter()
                .copied(),
        );
        package
            .try_replace_content_types_bytes(
                expected.bytes(),
                Arc::new(noncanonical.clone()),
                limits,
            )
            .unwrap();
        noncanonical
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
        let noncanonical_content_types = add_noncanonical_content_types(&mut package);
        let expected_stored_content_types =
            expected_excel_vba_content_types(&noncanonical_content_types);

        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let commit = edit.commit().unwrap();
        let stored_bytes = PackageWriter::to_bytes(commit.package()).unwrap();
        let stored = OpcPackage::from_bytes(&stored_bytes, package.read_limits()).unwrap();
        assert_eq!(
            stored
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            expected_stored_content_types.as_slice()
        );
        assert!(
            stored
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .windows(b"retained VBA manifest spelling".len())
                .any(|window| window == b"retained VBA manifest spelling")
        );
        let recipient_bytes = PackageWriter::to_bytes(&package).unwrap();
        let recipient = OpcPackage::from_bytes(&recipient_bytes, package.read_limits()).unwrap();
        let replayed = commit.patch().apply(&recipient).unwrap();
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
        assert_eq!(
            replayed
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            expected_stored_content_types.as_slice()
        );

        let mut divergent = recipient.clone();
        let divergent_expected = divergent.source_content_types(limits).unwrap();
        let mut divergent_content_types = divergent_expected.bytes().to_vec();
        let close = b"</Types>";
        let close_offset = divergent_content_types
            .windows(close.len())
            .position(|window| window == close)
            .expect("content-types close tag");
        divergent_content_types.splice(
            close_offset..close_offset,
            b"<!-- recipient-only manifest spelling -->\n"
                .iter()
                .copied(),
        );
        divergent
            .try_replace_content_types_bytes(
                divergent_expected.bytes(),
                Arc::new(divergent_content_types.clone()),
                limits,
            )
            .unwrap();
        let expected_divergent_content_types =
            expected_excel_vba_content_types(&divergent_content_types);
        let divergent_replayed = commit.patch().apply(&divergent).unwrap();
        assert_eq!(
            divergent_replayed
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            expected_divergent_content_types.as_slice()
        );

        // The recipient's manifest is deliberately divergent.  Derive the
        // exact aggregate from the composed replay image, then prove the
        // operation admits that boundary and refuses one byte below it.
        let projected = divergent_replayed.package();
        let projected_content_types = projected.source_content_types(limits).unwrap();
        let projected_part_bytes = projected
            .iter_parts()
            .map(|part| part.payload_len().unwrap() as u64)
            .sum::<u64>();
        let projected_relationship_bytes =
            relationship_totals(projected, limits).unwrap().xml_bytes;
        let projected_aggregate = projected_part_bytes
            .saturating_add(projected_relationship_bytes)
            .saturating_add(projected_content_types.bytes().len() as u64);
        assert_eq!(
            projected
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            expected_divergent_content_types.as_slice()
        );

        let exact_limits = SourceLimits::builder()
            .max_aggregate_bytes(projected_aggregate)
            .unwrap()
            .build()
            .unwrap();
        let mut exact_patch = commit.patch().clone();
        exact_patch.options = EditOptions::new().source_limits(exact_limits);
        let exact = exact_patch.apply(&divergent).unwrap();
        assert_eq!(
            exact
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            expected_divergent_content_types.as_slice()
        );

        let before_too_tight = PackageWriter::to_bytes(&divergent).unwrap();
        let one_under = SourceLimits::builder()
            .max_aggregate_bytes(projected_aggregate.saturating_sub(1))
            .unwrap()
            .build()
            .unwrap();
        exact_patch.options = EditOptions::new().source_limits(one_under);
        assert!(matches!(
            exact_patch.apply(&divergent),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::AggregateBytes),
                ..
            })
        ));
        assert_eq!(
            PackageWriter::to_bytes(&divergent).unwrap(),
            before_too_tight
        );

        let stored_again = PackageWriter::to_bytes(&stored).unwrap();
        let stored_recipient =
            OpcPackage::from_bytes(&stored_again, package.read_limits()).unwrap();
        let reverted = commit.patch().inverse().apply(&stored_recipient).unwrap();
        let reverted_source = reverted
            .package()
            .retained_relationships_xml(&source_name)
            .unwrap()
            .unwrap()
            .0;
        assert_eq!(reverted_source.as_slice(), noncanonical.as_slice());
        assert_eq!(
            reverted
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            noncanonical_content_types.as_slice()
        );
    }

    #[test]
    fn inverse_preserves_independent_word_project_and_supplemental_sources() {
        let source_name = PackURI::new("/word/document.xml").unwrap();
        let project_name = PackURI::new(WORD_PROJECT_PART).unwrap();
        let supplemental_name = PackURI::new(WORD_SUPPLEMENTAL_PART).unwrap();
        let mut package = word_package();
        let mut edit = Edit::new(&package, Host::Word, EditOptions::new()).unwrap();
        edit.store(b"first-project".to_vec(), Some(b"<first-supp/>".to_vec()))
            .unwrap();
        package = edit.commit().unwrap().into_package();

        let limits = package.read_limits();
        let source_expected = package.source_relationships(&source_name, limits).unwrap();
        let mut source_noncanonical = source_expected.bytes().to_vec();
        let relationship = b"<Relationship ";
        let source_offset = source_noncanonical
            .windows(relationship.len())
            .position(|window| window == relationship)
            .expect("Word source relationship");
        source_noncanonical.splice(
            source_offset..source_offset + relationship.len(),
            b"\n  <Relationship ".iter().copied(),
        );
        package
            .try_replace_relationships_bytes(
                &source_expected,
                Arc::new(source_noncanonical.clone()),
                true,
                limits,
            )
            .unwrap();

        let project_expected = package.source_relationships(&project_name, limits).unwrap();
        let mut project_noncanonical = project_expected.bytes().to_vec();
        let project_offset = project_noncanonical
            .windows(relationship.len())
            .position(|window| window == relationship)
            .expect("Word project relationship");
        project_noncanonical.splice(
            project_offset..project_offset + relationship.len(),
            b"\n  <Relationship ".iter().copied(),
        );
        package
            .try_replace_relationships_bytes(
                &project_expected,
                Arc::new(project_noncanonical.clone()),
                true,
                limits,
            )
            .unwrap();
        let noncanonical_content_types = add_noncanonical_content_types(&mut package);

        let before_bytes = PackageWriter::to_bytes(&package).unwrap();
        let mut edit = Edit::new(&package, Host::Word, EditOptions::new()).unwrap();
        edit.store(b"second-project".to_vec(), Some(b"<second-supp/>".to_vec()))
            .unwrap();
        let commit = edit.commit().unwrap();
        let stored_bytes = PackageWriter::to_bytes(commit.package()).unwrap();
        let stored = OpcPackage::from_bytes(&stored_bytes, limits).unwrap();
        let recipient = OpcPackage::from_bytes(&before_bytes, limits).unwrap();
        let replayed = commit.patch().apply(&recipient).unwrap();
        assert_eq!(
            replayed
                .package()
                .retained_relationships_xml(&source_name)
                .unwrap()
                .unwrap()
                .0
                .as_slice(),
            source_noncanonical.as_slice()
        );
        assert_eq!(
            replayed
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            noncanonical_content_types.as_slice()
        );

        let stored_recipient = {
            let bytes = PackageWriter::to_bytes(&stored).unwrap();
            OpcPackage::from_bytes(&bytes, limits).unwrap()
        };
        let reverted = commit.patch().inverse().apply(&stored_recipient).unwrap();
        assert_eq!(
            reverted
                .package()
                .retained_relationships_xml(&source_name)
                .unwrap()
                .unwrap()
                .0
                .as_slice(),
            source_noncanonical.as_slice()
        );
        assert_eq!(
            reverted
                .package()
                .retained_relationships_xml(&project_name)
                .unwrap()
                .unwrap()
                .0
                .as_slice(),
            project_noncanonical.as_slice()
        );
        assert!(
            reverted
                .package()
                .retained_relationships_xml(&supplemental_name)
                .unwrap()
                .is_none()
        );
        assert_eq!(
            reverted
                .package()
                .retained_content_types_xml()
                .unwrap()
                .unwrap()
                .as_slice(),
            noncanonical_content_types.as_slice()
        );
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
            .unwrap()
            .build()
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
            .unwrap()
            .build()
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
    fn edit_rejects_tighter_structural_limits_before_publication() {
        // The authored package exercises graph growth without opaque ZIP
        // members, so manifest and relationship limits reach their intended
        // refusal kinds before publication.
        let package = excel_package();
        let limits = package.read_limits();
        let content_types = package.source_content_types(limits).unwrap();
        let relationships = relationship_totals(&package, limits).unwrap();
        let root = PackURI::new("/").unwrap();
        let mut expected_relationship_bytes = package
            .source_relationships(&root, limits)
            .unwrap()
            .bytes()
            .len() as u64;
        for metadata in package.iter_parts() {
            let relationships = package
                .source_relationships(metadata.partname(), limits)
                .unwrap();
            if relationships.member_present() {
                expected_relationship_bytes = expected_relationship_bytes
                    .checked_add(relationships.bytes().len() as u64)
                    .unwrap();
            }
        }
        assert_eq!(relationships.xml_bytes, expected_relationship_bytes);
        let before = PackageWriter::to_bytes(&package).unwrap();

        let assert_growth_refused = |operation_limits: SourceLimits, expected: LimitKind| {
            let result = (|| {
                let options = EditOptions::new().source_limits(operation_limits);
                let mut edit = Edit::new(&package, Host::Excel, options)?;
                edit.store(b"project".to_vec(), None)?;
                edit.commit()
            })();
            assert!(matches!(
                result,
                Err(Error::Limit {
                    kind: crate::LimitKind::Vba(kind),
                    ..
                }) if kind == expected
            ));
            assert_eq!(PackageWriter::to_bytes(&package).unwrap(), before);
        };

        assert_growth_refused(
            SourceLimits::builder()
                .max_content_type_mappings(content_types.mapping_count())
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ContentTypeMappings,
        );
        assert_growth_refused(
            SourceLimits::builder()
                .max_total_relationship_xml_events(relationships.events as usize)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::TotalRelationshipXmlEvents,
        );

        // The directory-bearing source already owns the VBA graph. Updating
        // that known graph keeps the targeted preservation route eligible,
        // while its physical directory entry makes the member and total-entry
        // ceilings independently representable.
        let directory_package = excel_source_package_with_existing_vba_and_directory_entry();
        let projected = {
            let mut edit = Edit::new(&directory_package, Host::Excel, EditOptions::new()).unwrap();
            edit.store(b"project".to_vec(), None).unwrap();
            let candidate = edit.commit().unwrap().into_package();
            let bytes = PackageWriter::to_bytes(&candidate).unwrap();
            OpcPackage::from_vec(bytes, SourceLimits::default()).unwrap()
        };
        let projected_bytes = PackageWriter::to_bytes(&projected).unwrap();
        let projected_relationships =
            relationship_totals(&projected, SourceLimits::default()).unwrap();
        let (projected_members, projected_entries) = serialized_archive_counts(&projected_bytes);
        assert!(
            projected_entries > projected_members,
            "the structural fixture must retain directory entries"
        );
        let directory_before = PackageWriter::to_bytes(&directory_package).unwrap();
        let assert_archive_refused = |operation_limits: SourceLimits, expected: LimitKind| {
            let result = (|| {
                let options = EditOptions::new().source_limits(operation_limits);
                let mut edit = Edit::new(&directory_package, Host::Excel, options)?;
                edit.store(b"project".to_vec(), None)?;
                edit.commit()
            })();
            assert!(matches!(
                result,
                Err(Error::Limit {
                    kind: crate::LimitKind::Vba(kind),
                    ..
                }) if kind == expected
            ));
            assert_eq!(
                PackageWriter::to_bytes(&directory_package).unwrap(),
                directory_before
            );
        };

        assert_archive_refused(
            SourceLimits::builder()
                .max_parts(projected.part_count())
                .unwrap()
                .max_relationship_parts(projected_relationships.members)
                .unwrap()
                .max_members(projected_members - 1)
                .unwrap()
                .max_archive_total_entries(projected_entries)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ArchiveMembers,
        );
        assert_archive_refused(
            SourceLimits::builder()
                .max_parts(projected.part_count())
                .unwrap()
                .max_relationship_parts(projected_relationships.members)
                .unwrap()
                .max_members(projected_members)
                .unwrap()
                .max_archive_total_entries(projected_entries - 1)
                .unwrap()
                .build()
                .unwrap(),
            LimitKind::ArchiveTotalEntries,
        );
    }

    #[test]
    fn deferred_graph_payload_is_rejected_before_hydration_under_tighter_policy() {
        let package = excel_package();
        let mut initial = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        initial.store(b"project".to_vec(), None).unwrap();
        let stored = initial.commit().unwrap().into_package();
        let bytes = PackageWriter::to_bytes(&stored).unwrap();
        let opened = OpcPackage::from_vec(bytes, SourceLimits::default()).unwrap();
        assert_eq!(opened.deferred_decode_counters(), Some((0, 0)));

        let tighter = SourceLimits::builder()
            .max_part_bytes(1)
            .unwrap()
            .max_entry_bytes(1)
            .unwrap()
            .build()
            .unwrap();
        let result = Edit::new(
            &opened,
            Host::Excel,
            EditOptions::new().source_limits(tighter),
        )
        .and_then(|mut edit| edit.store(vec![0], None));
        assert!(matches!(
            result,
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::PartBytes),
                ..
            })
        ));
        assert_eq!(opened.deferred_decode_counters(), Some((0, 0)));
    }

    #[test]
    fn replay_rejects_unrelated_oversized_member_before_selected_hydration() {
        let mut recipient = excel_package();
        recipient
            .try_add_part(Box::new(BlobPart::new(
                PackURI::new("/custom/unrelated.bin").unwrap(),
                "application/octet-stream".to_owned(),
                vec![0; 4096],
            )))
            .unwrap();
        let source = PackageWriter::to_bytes(&recipient).unwrap();
        let opened = OpcPackage::from_vec(source, SourceLimits::default()).unwrap();
        assert_eq!(opened.deferred_decode_counters(), Some((0, 0)));

        let package = excel_package();
        let mut edit = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let mut patch = edit.commit().unwrap().patch().clone();
        let tighter = SourceLimits::builder()
            .max_entry_bytes(1024)
            .unwrap()
            .build()
            .unwrap();
        patch.options = EditOptions::new().source_limits(tighter);
        let before = PackageWriter::to_bytes(&opened).unwrap();
        assert!(matches!(
            patch.apply(&opened),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::EntryBytes),
                ..
            })
        ));
        assert_eq!(opened.deferred_decode_counters(), Some((0, 0)));
        assert_eq!(PackageWriter::to_bytes(&opened).unwrap(), before);
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
    fn replay_projects_serialized_source_namespace_after_unrelated_remove_add() {
        let old_name = PackURI::new("/custom/old.bin").unwrap();
        let new_name = PackURI::new("/custom/new.bin").unwrap();
        let package = excel_package();
        let mut initial = Edit::new(&package, Host::Excel, EditOptions::new()).unwrap();
        initial.store(b"source".to_vec(), None).unwrap();
        let stored = initial.commit().unwrap().into_package();

        // Build a small authored source with both replacement candidates
        // already present. The ZIP rewrite adds a real directory entry, but
        // keeps every subsequently edited member modeled by OPC provenance:
        // replay can therefore exercise source-backed preservation without
        // asking the writer to append beside an opaque directory member.
        let mut source = stored.clone();
        let mut old = BlobPart::new(
            old_name.clone(),
            "application/octet-stream".to_owned(),
            vec![0; 4096],
        );
        old.rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_owned(),
                "../xl/workbook.xml".to_owned(),
                "oldEdge".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        source.try_add_part(Box::new(old)).unwrap();
        let mut source_new = BlobPart::new(
            new_name.clone(),
            "application/octet-stream".to_owned(),
            b"source-new".to_vec(),
        );
        source_new
            .rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_owned(),
                "../xl/workbook.xml".to_owned(),
                "sourceNewEdge".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        source.try_add_part(Box::new(source_new)).unwrap();

        // Reopen a real source archive so the recipient retains immutable
        // provenance for the part and its relationships. Then replace that
        // known source part with a new known part carrying its own `.rels`.
        let authored_bytes = PackageWriter::to_bytes(&source).unwrap();
        let source_bytes = add_directory_entry(&authored_bytes);
        let mut recipient = OpcPackage::from_vec(source_bytes, SourceLimits::default()).unwrap();
        assert!(recipient.remove_part(&old_name));
        assert!(recipient.remove_part(&new_name));
        let mut replacement = BlobPart::new(
            new_name.clone(),
            "application/octet-stream".to_owned(),
            b"new".to_vec(),
        );
        replacement
            .rels_mut()
            .try_add_relationship(
                "urn:unrelated".to_owned(),
                "../xl/workbook.xml".to_owned(),
                "newEdge".to_owned(),
                TargetMode::Internal,
            )
            .unwrap();
        recipient.try_add_part(Box::new(replacement)).unwrap();
        assert!(recipient.source_read_limits().is_some());

        let mut edit = Edit::new(&stored, Host::Excel, EditOptions::new()).unwrap();
        edit.store(b"project".to_vec(), None).unwrap();
        let patch = edit.commit().unwrap().patch().clone();

        let broad = patch.apply(&recipient).unwrap();
        let broad_bytes = PackageWriter::to_bytes(broad.package()).unwrap();
        let (final_members, final_entries) = serialized_archive_counts(&broad_bytes);
        let reopened = OpcPackage::from_vec(broad_bytes, SourceLimits::default()).unwrap();
        let projected_metrics = reopened
            .retained_archive_metrics(SourceLimits::default())
            .unwrap()
            .unwrap();
        let final_aggregate = projected_metrics.aggregate_bytes();
        let final_metadata = projected_metrics.metadata_bytes();
        assert!(final_members > 1);
        assert!(final_entries > final_members);
        assert!(final_aggregate > 0);
        assert!(final_metadata > 0);
        let final_relationships = relationship_totals(&reopened, SourceLimits::default()).unwrap();

        let limits = |members, entries, aggregate, metadata| {
            SourceLimits::builder()
                .max_parts(reopened.part_count())
                .unwrap()
                .max_relationship_parts(final_relationships.members)
                .unwrap()
                .max_members(members)
                .unwrap()
                .max_archive_total_entries(entries)
                .unwrap()
                .max_entry_bytes(1024)
                .unwrap()
                .max_archive_compressed_bytes(1024)
                .unwrap()
                .max_aggregate_bytes(aggregate)
                .unwrap()
                .max_archive_metadata_bytes(metadata)
                .unwrap()
                .build()
                .unwrap()
        };
        let exact_limits = limits(
            final_members,
            final_entries,
            final_aggregate,
            final_metadata,
        );
        let mut exact_patch = patch.clone();
        exact_patch.options = EditOptions::new().source_limits(exact_limits);
        let exact = exact_patch.apply(&recipient).unwrap();
        assert!(exact.diagnostics().changed());

        let before = PackageWriter::to_bytes(&recipient).unwrap();
        let members_under = limits(
            final_members - 1,
            final_entries,
            final_aggregate,
            final_metadata,
        );
        exact_patch.options = EditOptions::new().source_limits(members_under);
        assert!(matches!(
            exact_patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::ArchiveMembers),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);

        let entries_under = limits(
            final_members,
            final_entries - 1,
            final_aggregate,
            final_metadata,
        );
        exact_patch.options = EditOptions::new().source_limits(entries_under);
        assert!(matches!(
            exact_patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::ArchiveTotalEntries),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);

        let aggregate_under = limits(
            final_members,
            final_entries,
            final_aggregate - 1,
            final_metadata,
        );
        exact_patch.options = EditOptions::new().source_limits(aggregate_under);
        assert!(matches!(
            exact_patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::AggregateBytes),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);

        let metadata_under = limits(
            final_members,
            final_entries,
            final_aggregate,
            final_metadata - 1,
        );
        exact_patch.options = EditOptions::new().source_limits(metadata_under);
        assert!(matches!(
            exact_patch.apply(&recipient),
            Err(Error::Limit {
                kind: crate::LimitKind::Vba(LimitKind::ArchiveMetadataBytes),
                ..
            })
        ));
        assert_eq!(PackageWriter::to_bytes(&recipient).unwrap(), before);
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
            EXCEL_MAIN_XML.to_vec(),
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
