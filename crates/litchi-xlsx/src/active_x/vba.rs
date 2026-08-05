//! Inert MS-OFFMACRO2 VBA-project relationship discovery for Excel workbooks.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, decompress, or execute VBA project bytes.

use crate::error::{Error, Result};
use litchi_ooxml_common::vba::{
    Host, read_project_part, remove_project_graph, store_project_graph,
};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};
use litchi_vba::{Limits, Payload, project::Project as ParsedProject};
use std::sync::Arc;

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
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the relationship ID from the Workbook part to the VBA Project part.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of the VBA Project binary part.
    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Parse the `vbaProject.bin` payload with default resource limits.
    pub fn project(&self, package: &OpcPackage) -> Result<ParsedProject> {
        self.project_with(package, &Limits::default())
    }

    /// Parse the `vbaProject.bin` payload with explicit resource limits.
    pub fn project_with(&self, package: &OpcPackage, limits: &Limits) -> Result<ParsedProject> {
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

pub fn store(package: &mut OpcPackage, source: &PackURI, payload: Payload) -> Result<Project> {
    let payload = Arc::new(payload.into_bytes());
    store_project_graph(package, source, Host::Excel, payload, None)?;
    let source = package.get_part(source)?;
    discover(package, source)?
        .ok_or_else(|| Error::Invalid("stored Excel VBA project was not discoverable".to_string()))
}

pub fn remove(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    Ok(remove_project_graph(package, source, Host::Excel)?)
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
        let parsed = metadata.project(&package).unwrap();
        assert_eq!(parsed.name(), "ExcelProject");
        assert_eq!(parsed.modules().len(), 1);

        let too_small = Limits {
            max_cfb_bytes: 0,
            ..Limits::default()
        };
        assert!(matches!(
            metadata.project_with(&package, &too_small),
            Err(Error::Common(litchi_ooxml_common::Error::Vba(
                litchi_vba::Error::LimitExceeded { .. },
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
