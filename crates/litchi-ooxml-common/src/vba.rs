//! Shared, inert MS-OFFMACRO2 package-graph mutation.
//!
//! The `project` payload itself remains owned by `litchi-vba`. This module
//! only validates and mutates the OPC relationship graph shared by DOCX,
//! PPTX, XLSX, and XLSB hosts.

use crate::signature::{refuse_signed_change, settle_unsigned_change};
use crate::{Error, Result};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI, TargetMode};
use litchi_vba::{Limits, project::Project};
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

#[derive(Debug, Clone, PartialEq, Eq)]
struct PartState {
    name: PackURI,
    content_type: String,
    data: Arc<Vec<u8>>,
    relationships: Vec<RelationshipState>,
}

impl PartState {
    fn capture(part: &dyn Part) -> Self {
        let mut relationships = part
            .rels()
            .iter()
            .map(RelationshipState::capture)
            .collect::<Vec<_>>();
        relationships.sort_unstable_by(|left, right| left.id.cmp(&right.id));
        Self {
            name: part.partname().clone(),
            content_type: part.content_type().to_owned(),
            data: part.blob_arc(),
            relationships,
        }
    }

    fn to_part(&self) -> Result<BlobPart> {
        let mut part = BlobPart::new_shared(
            self.name.clone(),
            self.content_type.clone(),
            Arc::clone(&self.data),
        );
        for relationship in &self.relationships {
            part.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.target_mode,
            )?;
        }
        Ok(part)
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct GraphSnapshot {
    host: Host,
    source: PartState,
    project: Option<PartState>,
    supplemental: Option<PartState>,
}

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
            source: PartState::capture(source),
            project: project.map(PartState::capture),
            supplemental: supplemental.map(PartState::capture),
        })
    }
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
    project_payload: Arc<Vec<u8>>,
    word_supplemental_xml: Option<Vec<u8>>,
) -> Result<()> {
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
        package
            .get_part_mut(source_part)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.project_part);
        if let Some((supplemental, _)) = graph.supplemental {
            package.remove_part(&supplemental);
        }
    }

    let project_name =
        PackURI::new(host.project_part()).map_err(|error| Error::Uri(error.clone()))?;
    let mut project = BlobPart::new(project_name, ct::OFC_VBA_PROJECT.to_string(), Vec::new());
    project.set_blob_shared(project_payload);
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

    let source = package.get_part_mut(source_part)?;
    source.relate_to(host.project_target(), rt::VBA_PROJECT);
    source.set_content_type(macro_content_type.to_string())?;
    settle_unsigned_change(package)?;
    Ok(())
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
    let source = package.get_part_mut(source_part)?;
    source.rels_mut().remove(&graph.relationship_id);
    source.set_content_type(non_macro_content_type.to_string())?;
    package.remove_part(&graph.project_part);
    if let Some((supplemental, _)) = graph.supplemental {
        package.remove_part(&supplemental);
    }
    settle_unsigned_change(package)?;
    Ok(())
}

/// A staged VBA-project graph edit over one immutable OPC package.
///
/// The edit captures only the selected host graph and shares every package
/// payload through the OPC package clone. Dropping it discards the candidate.
#[derive(Clone)]
pub struct Edit {
    source: OpcPackage,
    candidate: OpcPackage,
    before: GraphSnapshot,
    source_part: PackURI,
    host: Host,
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

impl Edit {
    /// Capture the selected host's main-part graph and start an isolated edit.
    ///
    /// The package is cloned only as a structural candidate; unchanged part
    /// payloads remain shared. The source package is never mutated.
    ///
    /// # Errors
    ///
    /// Returns an error when the host main part or its existing VBA graph is
    /// malformed.
    pub fn new(package: &OpcPackage, host: Host) -> Result<Self> {
        let source_part = package.main_document_part()?.partname().clone();
        let before = GraphSnapshot::capture(package, &source_part, host)?;
        Ok(Self {
            source: package.clone(),
            candidate: package.clone(),
            before,
            source_part,
            host,
        })
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
        project_payload: Arc<Vec<u8>>,
        word_supplemental_xml: Option<Vec<u8>>,
    ) -> Result<()> {
        let mut candidate = self.candidate.clone();
        store_project_graph_in_place(
            &mut candidate,
            &self.source_part,
            self.host,
            project_payload,
            word_supplemental_xml,
        )?;
        self.candidate = candidate;
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
        let mut candidate = self.candidate.clone();
        let changed =
            remove_project_graph_on_candidate(&mut candidate, &self.source_part, self.host)?;
        self.candidate = candidate;
        Ok(changed)
    }

    /// Discard every staged graph change.
    pub fn reset(&mut self) {
        self.candidate = self.source.clone();
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
        let after = GraphSnapshot::capture(&self.candidate, &self.source_part, self.host)?;
        let patch = Patch {
            before: self.before,
            after,
        };
        if patch.is_noop() {
            return Ok(Commit::new(self.source, patch, Diagnostics::UNCHANGED));
        }
        refuse_signed_change(&self.source)?;
        Ok(Commit::new(self.candidate, patch, Diagnostics::CHANGED))
    }
}

/// A complete, reversible VBA-project graph patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Patch {
    before: GraphSnapshot,
    after: GraphSnapshot,
}

impl Patch {
    /// Whether this patch is an exact graph no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Return the exact inverse patch without reinterpreting VBA bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
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
    /// Returns a relationship error when the selected graph base differs and
    /// the signed-package refusal for a changed recipient.
    pub fn apply(&self, package: &OpcPackage) -> Result<Commit> {
        let current = GraphSnapshot::capture(package, &self.before.source.name, self.before.host)?;
        if current != self.before {
            return Err(Error::Relationship(
                "VBA patch source graph changed".to_string(),
            ));
        }
        if self.is_noop() {
            return Ok(Commit::new(
                package.clone(),
                self.clone(),
                Diagnostics::UNCHANGED,
            ));
        }
        refuse_signed_change(package)?;
        let mut candidate = package.clone();
        restore_graph_in_place(&mut candidate, &self.after)?;
        let resulting =
            GraphSnapshot::capture(&candidate, &self.after.source.name, self.after.host)?;
        if resulting != self.after {
            return Err(Error::Relationship(
                "VBA patch publication changed the staged graph".to_string(),
            ));
        }
        Ok(Commit::new(candidate, self.clone(), Diagnostics::CHANGED))
    }
}

/// The published VBA graph commit triad.
pub type Commit = litchi_core::edit::Commit<OpcPackage, Patch>;

/// Minimal content-free commit diagnostics.
pub use litchi_core::edit::Diagnostics;

fn restore_graph_in_place(package: &mut OpcPackage, target: &GraphSnapshot) -> Result<()> {
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
    if let Some(graph) = existing {
        package
            .get_part_mut(source_part)?
            .rels_mut()
            .remove(&graph.relationship_id);
        package.remove_part(&graph.project_part);
        if let Some((supplemental, _)) = graph.supplemental {
            package.remove_part(&supplemental);
        }
    }

    if let Some(supplemental) = &target.supplemental {
        ensure_no_inbound_reference(package, &supplemental.name)?;
        package.try_add_part(Box::new(supplemental.to_part()?))?;
    }
    if let Some(project) = &target.project {
        ensure_no_inbound_reference(package, &project.name)?;
        package.try_add_part(Box::new(project.to_part()?))?;
    }

    let source = package.get_part_mut(source_part)?;
    if source.blob() != target.source.data.as_slice() {
        return Err(Error::Relationship(
            "VBA patch source part bytes changed".to_string(),
        ));
    }
    let current_relationships = source
        .rels()
        .iter()
        .map(RelationshipState::capture)
        .collect::<Vec<_>>();
    for current in current_relationships {
        let retained = target
            .source
            .relationships
            .iter()
            .any(|candidate| candidate == &current);
        if !retained {
            source.rels_mut().remove(&current.id);
        }
    }
    for relationship in &target.source.relationships {
        let present = source
            .rels()
            .get(&relationship.id)
            .map(RelationshipState::capture);
        if present.as_ref() != Some(relationship) {
            source.rels_mut().remove(&relationship.id);
            source.rels_mut().try_add_relationship(
                relationship.relationship_type.clone(),
                relationship.target.clone(),
                relationship.id.clone(),
                relationship.target_mode,
            )?;
        }
    }
    source.set_content_type(target.source.content_type.clone())?;
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
                Edit::new(&package, Host::Word)
                    .and_then(|mut edit| {
                        edit.store(
                            Arc::new(b"validated payload".to_vec()),
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
        let mut edit = Edit::new(&package, Host::Excel).unwrap();
        assert!(!edit.remove().unwrap());
        let _ = edit.commit().unwrap();
        assert!(package.is_signed());
        // Storing changes the signed package, so it is refused (ADR 0006).
        let store = |package: &OpcPackage| {
            let mut edit = Edit::new(package, Host::Excel)?;
            edit.store(Arc::new(b"project".to_vec()), None)?;
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
        assert!(refused(Edit::new(&package, Host::Excel).and_then(
            |mut edit| {
                edit.remove()?;
                edit.commit().map(|_| ())
            }
        )));
        assert!(package.get_part(&project_name).is_ok());
        package.unsign();
        let mut edit = Edit::new(&package, Host::Excel).unwrap();
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
        let mut edit = Edit::new(&package, Host::Excel).unwrap();
        edit.store(Arc::new(b"project".to_vec()), None).unwrap();
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

        let mut edit = Edit::new(&stored, Host::Excel).unwrap();
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

        let mut edit = Edit::new(&package, Host::Excel).unwrap();
        assert!(edit.remove().is_err());

        assert_eq!(package.part_count(), before_parts);
        assert!(package.get_part(&project_name).is_ok());
        let source = package.get_part(&source_name).unwrap();
        assert_eq!(source.content_type(), ct::SML_SHEET_MACRO_MAIN);
        assert!(source.rels().get(&relationship_id).is_some());
    }
}
