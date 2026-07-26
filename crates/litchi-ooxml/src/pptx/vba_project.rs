//! Inert MS-OFFMACRO2 VBA-project relationship discovery for presentations.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, decompress, or execute VBA project bytes.

use crate::error::{OoxmlError, Result};
use crate::vba_package::{
    VbaPackageHost, remove_vba_project_graph, store_vba_project_graph, validate_vba_project_payload,
};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};

/// Relationship metadata for the VBA project attached to a presentation.
///
/// This describes the MS-OFFMACRO2 package topology only. The
/// `vbaProject.bin` payload remains opaque and inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    source_part_name: PackURI,
    relationship_id: String,
    project_part_name: PackURI,
}

impl VbaProject {
    /// Return the Presentation part that owns the VBA-project relationship.
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the relationship ID from the Presentation part to the VBA Project part.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of the VBA Project binary part.
    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Parse the `vbaProject.bin` CFB payload as inert MS-OVBA source.
    pub fn read_project(
        &self,
        package: &OpcPackage,
        limits: &crate::vba::VbaLimits,
    ) -> std::result::Result<crate::vba::VbaProject, crate::vba::VbaError> {
        crate::vba::read_project_part(package, &self.project_part_name, limits)
    }
}

/// Discover one structurally conforming Presentation VBA-project relationship.
///
/// MS-OFFMACRO2 permits at most one VBA Project relationship from a
/// Presentation part and disallows relationships from the VBA Project part.
/// Its payload stays opaque here.
pub(crate) fn discover_vba_project(
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<Option<VbaProject>> {
    let mut projects = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::VBA_PROJECT);
    let Some(relationship) = projects.next() else {
        return Ok(None);
    };
    if projects.next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "Presentation part '{}' has multiple VBA Project relationships",
            source.partname().as_str()
        )));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project relationship '{}' from '{}' cannot be external",
            relationship.r_id(),
            source.partname().as_str()
        )));
    }

    let project_part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "invalid VBA Project relationship '{}' from '{}': {error}",
            relationship.r_id(),
            source.partname().as_str()
        ))
    })?;
    let project_part = package.get_part(&project_part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "VBA Project target '{}' from '{}': {error}",
            project_part_name.as_str(),
            source.partname().as_str()
        ))
    })?;
    if project_part.content_type() != content_type::OFC_VBA_PROJECT {
        return Err(OoxmlError::InvalidContentType {
            expected: content_type::OFC_VBA_PROJECT.to_string(),
            got: project_part.content_type().to_string(),
        });
    }
    if project_part.rels().iter().next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project part '{}' has a forbidden relationship",
            project_part_name.as_str()
        )));
    }

    Ok(Some(VbaProject {
        source_part_name: source.partname().clone(),
        relationship_id: relationship.r_id().to_string(),
        project_part_name,
    }))
}

pub(crate) fn store_vba_project(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Vec<u8>,
    limits: &crate::vba::VbaLimits,
) -> Result<VbaProject> {
    validate_vba_project_payload(&payload, limits)?;
    store_vba_project_graph(package, source, VbaPackageHost::PowerPoint, payload, None)?;
    let source = package.get_part(source)?;
    discover_vba_project(package, source)?.ok_or_else(|| {
        OoxmlError::InvalidFormat("stored PowerPoint VBA project was not discoverable".to_string())
    })
}

pub(crate) fn remove_vba_project(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    remove_vba_project_graph(package, source, VbaPackageHost::PowerPoint)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::pptx::Package;
    use crate::vba::{VbaLimits, VbaModuleBuilder, VbaProjectBuilder};
    use litchi_opc::part::BlobPart;

    fn package_with_vba_project(
        main_content_type: &str,
        add_forbidden_relationship: bool,
    ) -> OpcPackage {
        let presentation_name = PackURI::new("/ppt/presentation.xml").unwrap();
        let project_name = PackURI::new("/ppt/vbaProject.bin").unwrap();
        let mut presentation = BlobPart::new(
            presentation_name,
            main_content_type.to_string(),
            b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>".to_vec(),
        );
        presentation.relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);

        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if add_forbidden_relationship {
            project.relate_to("unexpected.bin", relationship_type::HYPERLINK);
        }

        let mut package = OpcPackage::new();
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(project));
        package.relate_to("ppt/presentation.xml", relationship_type::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn discovers_macro_project_metadata_without_parsing_payloads() {
        let package = package_with_vba_project(content_type::PML_PRES_MACRO_MAIN, false);
        let presentation = package.main_document_part().unwrap();

        let project = discover_vba_project(&package, presentation)
            .unwrap()
            .unwrap();
        assert_eq!(project.source_part_name().as_str(), "/ppt/presentation.xml");
        assert!(project.relationship_id().starts_with("rId"));
        assert_eq!(project.project_part_name().as_str(), "/ppt/vbaProject.bin");
    }

    #[test]
    fn rejects_project_parts_with_outbound_relationships() {
        let package = package_with_vba_project(content_type::PML_PRES_MACRO_MAIN, true);
        let presentation = package.main_document_part().unwrap();

        assert!(discover_vba_project(&package, presentation).is_err());
    }

    #[test]
    fn presentations_without_vba_projects_return_no_metadata() {
        let package = Package::new().unwrap();

        assert!(package.vba_project().unwrap().is_none());
    }

    #[test]
    fn package_accepts_all_macro_enabled_presentation_main_part_types() {
        for main_content_type in [
            content_type::PML_PRES_MACRO_MAIN,
            content_type::PML_SLIDESHOW_MACRO_MAIN,
            content_type::PML_TEMPLATE_MACRO_MAIN,
        ] {
            let package =
                Package::from_opc_package(package_with_vba_project(main_content_type, false))
                    .unwrap();
            let project = package.vba_project().unwrap().unwrap();

            assert_eq!(project.project_part_name().as_str(), "/ppt/vbaProject.bin");
            assert!(package.presentation().is_ok());
        }
    }

    fn authored_project() -> crate::vba::VbaProjectBinary {
        VbaProjectBuilder::new("PowerPointProject")
            .with_module(VbaModuleBuilder::standard(
                "Module1",
                "Public Sub Hello()\r\nEnd Sub\r\n",
            ))
            .build(&VbaLimits::default())
            .unwrap()
    }

    #[test]
    fn stores_and_preserves_project_across_presentation_materialization() {
        let file = tempfile::NamedTempFile::with_suffix(".pptm").unwrap();
        let mut package = Package::new().unwrap();
        let project = authored_project();
        package.set_vba_project(&project).unwrap();
        package.presentation_mut().unwrap().add_slide().unwrap();
        package.save(file.path()).unwrap();

        let mut reopened = Package::open(file.path()).unwrap();
        let metadata = reopened.vba_project().unwrap().unwrap();
        let parsed = metadata
            .read_project(reopened.opc_package(), &VbaLimits::default())
            .unwrap();
        assert_eq!(parsed.name(), "PowerPointProject");
        assert_eq!(
            reopened
                .opc_package()
                .main_document_part()
                .unwrap()
                .content_type(),
            content_type::PML_PRES_MACRO_MAIN
        );

        assert!(reopened.remove_vba_project().unwrap());
        assert!(reopened.vba_project().unwrap().is_none());
        assert_eq!(
            reopened
                .opc_package()
                .main_document_part()
                .unwrap()
                .content_type(),
            content_type::PML_PRESENTATION_MAIN
        );
    }

    #[test]
    fn rejects_invalid_project_before_mutating_presentation() {
        let mut package = Package::new().unwrap();
        assert!(
            package
                .set_vba_project_bytes(vec![0; 64], &VbaLimits::default())
                .is_err()
        );
        assert!(package.vba_project().unwrap().is_none());
    }

    #[test]
    fn slideshow_and_template_kinds_survive_attach_and_remove() {
        for (plain, macro_enabled) in [
            (
                content_type::PML_SLIDESHOW_MAIN,
                content_type::PML_SLIDESHOW_MACRO_MAIN,
            ),
            (
                content_type::PML_TEMPLATE_MAIN,
                content_type::PML_TEMPLATE_MACRO_MAIN,
            ),
        ] {
            let mut package = Package::new().unwrap();
            let source = package
                .opc_package()
                .main_document_part()
                .unwrap()
                .partname()
                .clone();
            package
                .opc_package_mut()
                .get_part_mut(&source)
                .unwrap()
                .set_content_type(plain.to_string())
                .unwrap();

            package.set_vba_project(&authored_project()).unwrap();
            assert_eq!(
                package
                    .opc_package()
                    .get_part(&source)
                    .unwrap()
                    .content_type(),
                macro_enabled
            );
            package.remove_vba_project().unwrap();
            assert_eq!(
                package
                    .opc_package()
                    .get_part(&source)
                    .unwrap()
                    .content_type(),
                plain
            );
        }
    }
}
