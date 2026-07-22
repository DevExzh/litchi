//! Inert MS-OFFMACRO2 VBA-project relationship discovery for Excel workbooks.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, decompress, or execute VBA project bytes.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part};

/// Relationship metadata for the VBA project attached to an Excel workbook.
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
}

/// Discover one structurally conforming Workbook VBA-project relationship.
///
/// MS-OFFMACRO2 permits at most one VBA Project relationship from a Workbook
/// part and disallows relationships from the VBA Project part. Its payload
/// stays opaque here.
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
            "Workbook part '{}' has multiple VBA Project relationships",
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::Workbook;
    use litchi_opc::part::BlobPart;

    fn add_vba_project(workbook: &mut Workbook, add_forbidden_relationship: bool) {
        let workbook_name = workbook
            .opc_package()
            .main_document_part()
            .unwrap()
            .partname()
            .clone();
        let project_name = PackURI::new("/xl/vbaProject.bin").unwrap();
        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if add_forbidden_relationship {
            project.relate_to("unexpected.bin", relationship_type::HYPERLINK);
        }

        let package = workbook.opc_package_mut();
        package.add_part(Box::new(project));
        package
            .get_part_mut(&workbook_name)
            .unwrap()
            .relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);
    }

    #[test]
    fn discovers_macro_project_metadata_without_parsing_payloads() {
        let mut workbook = Workbook::create().unwrap();
        add_vba_project(&mut workbook, false);

        let project = workbook.vba_project().unwrap().unwrap();
        assert_eq!(project.source_part_name().as_str(), "/xl/workbook.xml");
        assert!(project.relationship_id().starts_with("rId"));
        assert_eq!(project.project_part_name().as_str(), "/xl/vbaProject.bin");
    }

    #[test]
    fn rejects_project_parts_with_outbound_relationships() {
        let mut workbook = Workbook::create().unwrap();
        add_vba_project(&mut workbook, true);

        assert!(workbook.vba_project().is_err());
    }

    #[test]
    fn workbooks_without_vba_projects_return_no_metadata() {
        let workbook = Workbook::create().unwrap();

        assert!(workbook.vba_project().unwrap().is_none());
    }
}
