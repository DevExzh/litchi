//! Inert MS-OFFMACRO2 VBA-project relationship discovery for Word packages.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, decompress, or execute VBA project or supplemental
//! data bytes.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};

/// Relationship metadata for the VBA project attached to a Word main document.
///
/// This describes the MS-OFFMACRO2 package topology only. The `vbaProject.bin`
/// payload and Word VBA supplemental-data XML remain opaque and inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    source_part_name: PackURI,
    project_relationship_id: String,
    project_part_name: PackURI,
    supplemental_data_relationship_id: String,
    supplemental_data_part_name: PackURI,
}

impl VbaProject {
    /// Return the Word main part that owns the VBA-project relationship.
    pub fn source_part_name(&self) -> &PackURI {
        &self.source_part_name
    }

    /// Return the relationship ID from the main part to the VBA Project part.
    pub fn project_relationship_id(&self) -> &str {
        &self.project_relationship_id
    }

    /// Return the absolute OPC part name of the VBA Project binary part.
    pub fn project_part_name(&self) -> &PackURI {
        &self.project_part_name
    }

    /// Return the relationship ID from the VBA Project to Word supplemental data.
    pub fn supplemental_data_relationship_id(&self) -> &str {
        &self.supplemental_data_relationship_id
    }

    /// Return the absolute OPC part name of the Word VBA supplemental-data part.
    pub fn supplemental_data_part_name(&self) -> &PackURI {
        &self.supplemental_data_part_name
    }

    /// Parse the `vbaProject.bin` CFB payload as inert MS-OVBA source.
    ///
    /// The relationship graph remains independently inspectable through this
    /// type. This method only decompresses and decodes source; it never
    /// compiles, interprets, or executes VBA.
    pub fn read_project(
        &self,
        package: &OpcPackage,
        limits: &crate::vba::VbaLimits,
    ) -> std::result::Result<crate::vba::VbaProject, crate::vba::VbaError> {
        crate::vba::read_project_part(package, &self.project_part_name, limits)
    }
}

/// Discover one structurally conforming Word VBA-project relationship graph.
///
/// MS-OFFMACRO2 permits at most one VBA Project relationship from a Word main
/// document. Its binary project part must in turn have exactly one relationship
/// to the Word VBA Supplemental Data part. Both payloads stay opaque here.
pub(crate) fn discover_vba_project(
    package: &OpcPackage,
    source: &dyn Part,
) -> Result<Option<VbaProject>> {
    let mut projects = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::VBA_PROJECT);
    let Some(project_relationship) = projects.next() else {
        return Ok(None);
    };
    if projects.next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "Word main part '{}' has multiple VBA Project relationships",
            source.partname().as_str()
        )));
    }
    if project_relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project relationship '{}' from '{}' cannot be external",
            project_relationship.r_id(),
            source.partname().as_str()
        )));
    }

    let project_part_name = project_relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "invalid VBA Project relationship '{}' from '{}': {error}",
            project_relationship.r_id(),
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
    if project_part
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() != relationship_type::WORD_VBA_DATA)
    {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project part '{}' has a forbidden relationship",
            project_part_name.as_str()
        )));
    }

    let mut supplemental_data = project_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::WORD_VBA_DATA);
    let Some(supplemental_data_relationship) = supplemental_data.next() else {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project part '{}' is missing its Word VBA Supplemental Data relationship",
            project_part_name.as_str()
        )));
    };
    if supplemental_data.next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "VBA Project part '{}' has multiple Word VBA Supplemental Data relationships",
            project_part_name.as_str()
        )));
    }
    if supplemental_data_relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(format!(
            "Word VBA Supplemental Data relationship '{}' from '{}' cannot be external",
            supplemental_data_relationship.r_id(),
            project_part_name.as_str()
        )));
    }

    let supplemental_data_part_name =
        supplemental_data_relationship
            .target_partname()
            .map_err(|error| {
                OoxmlError::InvalidFormat(format!(
                    "invalid Word VBA Supplemental Data relationship '{}' from '{}': {error}",
                    supplemental_data_relationship.r_id(),
                    project_part_name.as_str()
                ))
            })?;
    let supplemental_data_part =
        package
            .get_part(&supplemental_data_part_name)
            .map_err(|error| {
                OoxmlError::PartNotFound(format!(
                    "Word VBA Supplemental Data target '{}' from '{}': {error}",
                    supplemental_data_part_name.as_str(),
                    project_part_name.as_str()
                ))
            })?;
    if supplemental_data_part.content_type() != content_type::WML_VBA_DATA {
        return Err(OoxmlError::InvalidContentType {
            expected: content_type::WML_VBA_DATA.to_string(),
            got: supplemental_data_part.content_type().to_string(),
        });
    }

    Ok(Some(VbaProject {
        source_part_name: source.partname().clone(),
        project_relationship_id: project_relationship.r_id().to_string(),
        project_part_name,
        supplemental_data_relationship_id: supplemental_data_relationship.r_id().to_string(),
        supplemental_data_part_name,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::Package;
    use litchi_opc::part::BlobPart;

    fn package_with_vba_project(
        main_content_type: &str,
        include_supplemental_data: bool,
    ) -> OpcPackage {
        let document_name = PackURI::new("/word/document.xml").unwrap();
        let project_name = PackURI::new("/word/vbaProject.bin").unwrap();
        let supplemental_name = PackURI::new("/word/vbaData.xml").unwrap();

        let mut document = BlobPart::new(
            document_name,
            main_content_type.to_string(),
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>".to_vec(),
        );
        document.relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);

        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if include_supplemental_data {
            project.relate_to("vbaData.xml", relationship_type::WORD_VBA_DATA);
        }

        let supplemental_data = BlobPart::new(
            supplemental_name,
            content_type::WML_VBA_DATA.to_string(),
            b"intentionally not XML".to_vec(),
        );

        let mut package = OpcPackage::new();
        package.add_part(Box::new(document));
        package.add_part(Box::new(project));
        package.add_part(Box::new(supplemental_data));
        package.relate_to("word/document.xml", relationship_type::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn discovers_macro_project_metadata_without_parsing_payloads() {
        let package = package_with_vba_project(content_type::WML_DOCUMENT_MACRO_MAIN, true);
        let source = package.main_document_part().unwrap();

        let project = discover_vba_project(&package, source).unwrap().unwrap();
        assert_eq!(project.source_part_name().as_str(), "/word/document.xml");
        assert_eq!(project.project_relationship_id(), "rId1");
        assert_eq!(project.project_part_name().as_str(), "/word/vbaProject.bin");
        assert_eq!(project.supplemental_data_relationship_id(), "rId1");
        assert_eq!(
            project.supplemental_data_part_name().as_str(),
            "/word/vbaData.xml"
        );
    }

    #[test]
    fn rejects_a_vba_project_without_required_supplemental_data() {
        let package = package_with_vba_project(content_type::WML_DOCUMENT_MACRO_MAIN, false);
        let source = package.main_document_part().unwrap();

        assert!(discover_vba_project(&package, source).is_err());
    }

    #[test]
    fn documents_without_vba_projects_return_no_metadata() {
        let package = Package::new().unwrap();

        assert!(package.vba_project().unwrap().is_none());
    }

    #[test]
    fn docx_package_accepts_macro_enabled_document_and_template_main_parts() {
        for main_content_type in [
            content_type::WML_DOCUMENT_MACRO_MAIN,
            content_type::WML_TEMPLATE_MACRO_MAIN,
        ] {
            let package =
                Package::from_opc_package(package_with_vba_project(main_content_type, true))
                    .unwrap();
            let project = package.vba_project().unwrap().unwrap();

            assert_eq!(project.project_part_name().as_str(), "/word/vbaProject.bin");
            assert!(package.document().is_ok());
        }
    }
}
