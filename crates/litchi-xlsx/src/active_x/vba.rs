//! Inert MS-OFFMACRO2 VBA-project relationship discovery for Excel workbooks.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, decompress, or execute VBA project bytes.

use crate::error::Result;
use crate::part_edit::{Staged, family};
use litchi_ooxml_common::vba::{Edit as VbaEdit, EditOptions, Host, read_project_part};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_vba::{Limits, Payload, project::Project as ParsedProject};

/// Relationship metadata for the VBA project attached to an Excel workbook.
///
/// This describes the MS-OFFMACRO2 package topology only. The
/// `vbaProject.bin` payload remains opaque and inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Project {
    source_part_name: PackURI,
    relationship_id: String,
    project_part_name: PackURI,
}

impl Project {
    /// Return the Workbook part that owns the VBA-project relationship.
    #[must_use]
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the relationship ID from the Workbook part to the VBA Project part.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of the VBA Project binary part.
    #[must_use]
    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Parse the `vbaProject.bin` payload with explicit resource limits.
    pub fn project(&self, package: &OpcPackage, limits: &Limits) -> Result<ParsedProject> {
        Ok(read_project_part(package, &self.project_part_name, limits)?)
    }
}

/// Discover one structurally conforming Workbook VBA-project relationship.
///
/// MS-OFFMACRO2 permits at most one VBA Project relationship from a Workbook
/// part and disallows relationships from the VBA Project part. Its payload
/// stays opaque here.
pub fn discover(package: &OpcPackage, source: &dyn Part) -> Result<Option<Project>> {
    let mut projects = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::VBA_PROJECT);
    let Some(relationship) = projects.next() else {
        return Ok(None);
    };
    if projects.next().is_some() {
        return Err(Error::Invalid(format!(
            "Workbook part '{}' has multiple VBA Project relationships",
            source.partname().as_str()
        )));
    }
    if relationship.is_external() {
        return Err(Error::Invalid(format!(
            "VBA Project relationship '{}' from '{}' cannot be external",
            relationship.r_id(),
            source.partname().as_str()
        )));
    }

    let project_part_name = relationship.target_partname().map_err(|error| {
        Error::Invalid(format!(
            "invalid VBA Project relationship '{}' from '{}': {error}",
            relationship.r_id(),
            source.partname().as_str()
        ))
    })?;
    let project_part = package.get_part(&project_part_name).map_err(|error| {
        Error::Invalid(format!(
            "VBA Project target '{}' from '{}': {error}",
            project_part_name.as_str(),
            source.partname().as_str()
        ))
    })?;
    if project_part.content_type() != content_type::OFC_VBA_PROJECT {
        return Err(Error::Common(litchi_ooxml_common::Error::ContentType {
            expected: content_type::OFC_VBA_PROJECT.to_string(),
            actual: project_part.content_type().to_string(),
        }));
    }
    if project_part.rels().iter().next().is_some() {
        return Err(Error::Invalid(format!(
            "VBA Project part '{}' has a forbidden relationship",
            project_part_name.as_str()
        )));
    }

    Ok(Some(Project {
        source_part_name: source.partname().clone(),
        relationship_id: relationship.r_id().to_string(),
        project_part_name,
    }))
}

/// Store a VBA project for the workbook part `source`, replacing any present
/// graph.
///
/// Storing always republishes the graph, so a signed package is refused with
/// [`crate::Error::Signed`] before it is changed (ADR 0006): remove the
/// signature explicitly first, with [`crate::Package::unsign`].
///
/// # Errors
///
/// Returns [`crate::Error::Signed`] for a signed package, and an error when the
/// graph is invalid or the package cannot be changed.
pub(crate) fn store(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Payload,
) -> Result<Project> {
    let payload = payload.into_bytes();
    let mut edit = VbaEdit::new(package, Host::Excel, EditOptions::new())
        .map_err(crate::signature_policy::from_shared)?;
    edit.store(payload, None)
        .map_err(crate::signature_policy::from_shared)?;
    *package = edit
        .commit()
        .map_err(crate::signature_policy::from_shared)?
        .into_package();
    let source = package.get_part(source)?;
    discover(package, source)?
        .ok_or_else(|| Error::Invalid("stored Excel VBA project was not discoverable".to_string()))
}

/// Remove the workbook part `source`'s VBA project graph.
///
/// An absent graph returns `Ok(false)` and keeps a signed package's
/// signature; removing a present graph from a signed package is refused with
/// [`crate::Error::Signed`] (ADR 0006): remove the signature explicitly first,
/// with [`crate::Package::unsign`].
///
/// # Errors
///
/// Returns [`crate::Error::Signed`] for a change to a signed package, and an
/// error when the graph is invalid or the package cannot be changed.
pub(crate) fn remove(package: &mut OpcPackage, _source: &PackURI) -> Result<bool> {
    let mut edit = VbaEdit::new(package, Host::Excel, EditOptions::new())
        .map_err(crate::signature_policy::from_shared)?;
    let changed = edit
        .remove()
        .map_err(crate::signature_policy::from_shared)?;
    *package = edit
        .commit()
        .map_err(crate::signature_policy::from_shared)?
        .into_package();
    Ok(changed)
}

family!("VBA-project");

pub use crate::opc::LimitKind;

/// A staged edit of the workbook's VBA-project graph of one
/// [`crate::Package`]. Dropping it discards it.
///
/// The `vbaProject.bin` payload stays opaque: it is stored and removed, never
/// parsed or executed.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    workbook: PackURI,
}

impl crate::Package {
    /// Start an edit of the workbook's VBA-project graph.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_vba(&self) -> Result<Edit> {
        let staged = Staged::new(self, "edit_vba")?;
        let workbook = staged.opc().main_document_part()?.partname().clone();
        Ok(Edit { staged, workbook })
    }

    /// Apply a VBA-project patch to a new package; this one is unchanged.
    /// An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_vba(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_vba")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged VBA project relationship metadata, when present.
    ///
    /// # Errors
    ///
    /// Returns the validation errors of the VBA-project graph.
    pub fn project(&self) -> Result<Option<Project>> {
        let source = self.staged.opc().get_part(&self.workbook)?;
        discover(self.staged.opc(), source)
    }

    /// Store `payload` as the workbook's VBA project, replacing any present
    /// graph.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a signed package and the graph's
    /// validation errors.
    pub fn store(&mut self, payload: Payload) -> Result<Project> {
        let workbook = self.workbook.clone();
        self.staged
            .update(|candidate| store(candidate, &workbook, payload))
    }

    /// Remove the workbook's VBA-project graph. Returns whether one was
    /// present.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a change to a signed package and the
    /// graph's validation errors.
    pub fn remove(&mut self) -> Result<bool> {
        let workbook = self.workbook.clone();
        self.staged.update(|candidate| remove(candidate, &workbook))
    }

    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// Publish the staged graph to a new package. An unchanged edit returns
    /// the source package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Signed`] for a changed signed package and the
    /// workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::package::build_minimal_package;
    use litchi_opc::part::BlobPart;
    use litchi_vba::{Limits, Payload, build};

    fn add_vba_project(package: &mut OpcPackage, add_forbidden_relationship: bool) {
        let workbook_name = package.main_document_part().unwrap().partname().clone();
        let project_name = PackURI::new("/xl/vbaProject.bin").unwrap();
        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if add_forbidden_relationship {
            project.relate_to("unexpected.bin", relationship_type::HYPERLINK);
        }

        package.add_part(Box::new(project));
        package
            .get_part_mut(&workbook_name)
            .unwrap()
            .relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);
    }

    #[test]
    fn discovers_macro_project_metadata_without_parsing_payloads() {
        let mut package = build_minimal_package().unwrap();
        add_vba_project(&mut package, false);

        let workbook = package.main_document_part().unwrap();
        let project = discover(&package, workbook).unwrap().unwrap();
        assert_eq!(project.source_part_name().as_str(), "/xl/workbook.xml");
        assert!(project.relationship_id().starts_with("rId"));
        assert_eq!(project.project_part_name().as_str(), "/xl/vbaProject.bin");
    }

    #[test]
    fn rejects_project_parts_with_outbound_relationships() {
        let mut package = build_minimal_package().unwrap();
        add_vba_project(&mut package, true);

        let workbook = package.main_document_part().unwrap();
        assert!(discover(&package, workbook).is_err());
    }

    #[test]
    fn workbooks_without_vba_projects_return_no_metadata() {
        let package = build_minimal_package().unwrap();
        let workbook = package.main_document_part().unwrap();

        assert!(discover(&package, workbook).unwrap().is_none());
    }

    fn authored_project() -> build::Project {
        build::Project::new("ExcelProject").module(build::Module::standard(
            "Module1",
            "Public Sub Hello()\r\nEnd Sub\r\n",
        ))
    }

    #[test]
    fn stores_preserves_replaces_and_removes_authored_project() {
        let mut package = build_minimal_package().unwrap();
        let source = package.main_document_part().unwrap().partname().clone();
        let metadata = store(
            &mut package,
            &source,
            authored_project().finish(&Limits::default()).unwrap(),
        )
        .unwrap();
        assert_eq!(metadata.project_part_name().as_str(), "/xl/vbaProject.bin");
        assert_eq!(
            package.main_document_part().unwrap().content_type(),
            content_type::SML_SHEET_MACRO_MAIN
        );

        let metadata = discover(&package, package.main_document_part().unwrap())
            .unwrap()
            .unwrap();
        let parsed = metadata.project(&package, &Limits::default()).unwrap();
        assert_eq!(parsed.name(), "ExcelProject");
        assert_eq!(parsed.modules().len(), 1);

        let too_small = Limits {
            max_cfb_bytes: 0,
            ..Limits::default()
        };
        assert!(matches!(
            metadata.project(&package, &too_small),
            Err(Error::Common(litchi_ooxml_common::Error::Vba(
                litchi_vba::Error::Limit { .. },
            )))
        ));

        let limits = Limits::default();
        let imported = Payload::read(
            authored_project().finish(&limits).unwrap().into_bytes(),
            &limits,
        )
        .unwrap();
        store(&mut package, &source, imported).unwrap();
        assert!(remove(&mut package, &source).unwrap());
        assert!(
            discover(&package, package.main_document_part().unwrap())
                .unwrap()
                .is_none()
        );
        assert_eq!(
            package.main_document_part().unwrap().content_type(),
            content_type::SML_SHEET_MAIN
        );
        assert!(
            package
                .get_part(&PackURI::new("/xl/vbaProject.bin").unwrap())
                .is_err()
        );
    }

    #[test]
    fn rejects_invalid_project_before_mutating_workbook() {
        let package = build_minimal_package().unwrap();
        assert!(Payload::read(vec![0; 64], &Limits::default()).is_err());
        assert!(
            discover(&package, package.main_document_part().unwrap())
                .unwrap()
                .is_none()
        );
        assert_eq!(
            package.main_document_part().unwrap().content_type(),
            content_type::SML_SHEET_MAIN
        );
    }

    #[test]
    fn conflicting_canonical_part_name_is_rejected_atomically() {
        let mut package = build_minimal_package().unwrap();
        let occupied = PackURI::new("/xl/VBAPROJECT.bin").unwrap();
        package.add_part(Box::new(BlobPart::new(
            occupied.clone(),
            "application/octet-stream".to_string(),
            b"keep me".to_vec(),
        )));

        let source = package.main_document_part().unwrap().partname().clone();
        assert!(
            store(
                &mut package,
                &source,
                authored_project().finish(&Limits::default()).unwrap(),
            )
            .is_err()
        );
        assert_eq!(package.get_part(&occupied).unwrap().blob(), b"keep me");
        assert!(
            discover(&package, package.main_document_part().unwrap())
                .unwrap()
                .is_none()
        );
        assert_eq!(
            package.main_document_part().unwrap().content_type(),
            content_type::SML_SHEET_MAIN
        );
    }

    #[test]
    fn template_kind_survives_attach_and_remove() {
        let mut package = build_minimal_package().unwrap();
        let source = package.main_document_part().unwrap().partname().clone();
        package
            .get_part_mut(&source)
            .unwrap()
            .set_content_type(content_type::SML_TEMPLATE_MAIN.to_string())
            .unwrap();

        store(
            &mut package,
            &source,
            authored_project().finish(&Limits::default()).unwrap(),
        )
        .unwrap();
        assert_eq!(
            package.get_part(&source).unwrap().content_type(),
            content_type::SML_TEMPLATE_MACRO_MAIN
        );
        remove(&mut package, &source).unwrap();
        assert_eq!(
            package.get_part(&source).unwrap().content_type(),
            content_type::SML_TEMPLATE_MAIN
        );
    }
}
