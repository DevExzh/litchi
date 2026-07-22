//! Inert MS-XLSB VBA-project and signature relationship discovery.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, verify, or execute VBA project or signature
//! bytes.

use crate::error::{OoxmlError, Result};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};

/// The declared kind of an XLSB VBA project signature part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum VbaProjectSignatureKind {
    /// The legacy VBA project signature format.
    Legacy,
    /// The Agile VBA project signature format.
    Agile,
}

/// Relationship metadata for one declared XLSB VBA project signature part.
///
/// Signature bytes remain opaque: this is not a cryptographic verification
/// result and does not make a trust decision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProjectSignature {
    kind: VbaProjectSignatureKind,
    relationship_id: String,
    part_name: PackURI,
}

impl VbaProjectSignature {
    /// Return the declared signature format.
    pub fn kind(&self) -> VbaProjectSignatureKind {
        self.kind
    }

    /// Return the relationship ID from the VBA Project part to this signature.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of the signature part.
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }
}

/// Relationship metadata for the VBA project attached to an XLSB workbook.
///
/// This describes the MS-XLSB package topology only. The `vbaProject.bin`
/// payload and declared signature payloads remain opaque and inert.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    source_part_name: PackURI,
    relationship_id: String,
    project_part_name: PackURI,
    signatures: Vec<VbaProjectSignature>,
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

    /// Return declared legacy/Agile VBA project signatures in stable kind order.
    pub fn signatures(&self) -> &[VbaProjectSignature] {
        &self.signatures
    }
}

/// Discover one structurally conforming XLSB VBA-project relationship graph.
///
/// MS-XLSB permits at most one VBA Project relationship from a Workbook part,
/// and permits at most one legacy and one Agile signature relationship from
/// the project. All binary contents remain opaque here.
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

    let mut legacy_signature = None;
    let mut agile_signature = None;
    for signature_relationship in project_part.rels().iter() {
        let (kind, expected_content_type) = match signature_relationship.reltype() {
            relationship_type::VBA_PROJECT_SIGNATURE => (
                VbaProjectSignatureKind::Legacy,
                content_type::OFC_VBA_PROJECT_SIGNATURE,
            ),
            relationship_type::VBA_PROJECT_SIGNATURE_AGILE => (
                VbaProjectSignatureKind::Agile,
                content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE,
            ),
            _ => {
                return Err(OoxmlError::InvalidFormat(format!(
                    "VBA Project part '{}' has a forbidden relationship",
                    project_part_name.as_str()
                )));
            },
        };
        let signature = discover_signature(
            package,
            &project_part_name,
            signature_relationship,
            kind,
            expected_content_type,
        )?;
        let slot = match kind {
            VbaProjectSignatureKind::Legacy => &mut legacy_signature,
            VbaProjectSignatureKind::Agile => &mut agile_signature,
        };
        if slot.replace(signature).is_some() {
            return Err(OoxmlError::InvalidFormat(format!(
                "VBA Project part '{}' has multiple {kind:?} signature relationships",
                project_part_name.as_str()
            )));
        }
    }

    let signatures = [legacy_signature, agile_signature]
        .into_iter()
        .flatten()
        .collect();
    Ok(Some(VbaProject {
        source_part_name: source.partname().clone(),
        relationship_id: relationship.r_id().to_string(),
        project_part_name,
        signatures,
    }))
}

fn discover_signature(
    package: &OpcPackage,
    project_part_name: &PackURI,
    relationship: &Relationship,
    kind: VbaProjectSignatureKind,
    expected_content_type: &str,
) -> Result<VbaProjectSignature> {
    if relationship.is_external() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{kind:?} VBA Project Signature relationship '{}' from '{}' cannot be external",
            relationship.r_id(),
            project_part_name.as_str()
        )));
    }
    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "invalid {kind:?} VBA Project Signature relationship '{}' from '{}': {error}",
            relationship.r_id(),
            project_part_name.as_str()
        ))
    })?;
    let part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "{kind:?} VBA Project Signature target '{}' from '{}': {error}",
            part_name.as_str(),
            project_part_name.as_str()
        ))
    })?;
    if part.content_type() != expected_content_type {
        return Err(OoxmlError::InvalidContentType {
            expected: expected_content_type.to_string(),
            got: part.content_type().to_string(),
        });
    }
    if part.rels().iter().next().is_some() {
        return Err(OoxmlError::InvalidFormat(format!(
            "{kind:?} VBA Project Signature part '{}' has a forbidden relationship",
            part_name.as_str()
        )));
    }
    Ok(VbaProjectSignature {
        kind,
        relationship_id: relationship.r_id().to_string(),
        part_name,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::{MutableXlsbWorksheet, XlsbWorkbook, XlsbWorkbookWriter};
    use litchi_opc::part::BlobPart;
    use std::io::Cursor;

    fn package_with_vba_project(
        include_legacy_signature: bool,
        include_agile_signature: bool,
    ) -> OpcPackage {
        let workbook_name = PackURI::new("/xl/workbook.bin").unwrap();
        let project_name = PackURI::new("/xl/vbaProject.bin").unwrap();
        let mut workbook = BlobPart::new(
            workbook_name,
            content_type::XLSB_BIN.to_string(),
            b"not parsed".to_vec(),
        );
        workbook.relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);

        let mut project = BlobPart::new(
            project_name,
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        if include_legacy_signature {
            project.relate_to(
                "vbaProjectSignature.bin",
                relationship_type::VBA_PROJECT_SIGNATURE,
            );
        }
        if include_agile_signature {
            project.relate_to(
                "vbaProjectSignatureAgile.bin",
                relationship_type::VBA_PROJECT_SIGNATURE_AGILE,
            );
        }

        let mut package = OpcPackage::new();
        package.add_part(Box::new(workbook));
        package.add_part(Box::new(project));
        if include_legacy_signature {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/vbaProjectSignature.bin").unwrap(),
                content_type::OFC_VBA_PROJECT_SIGNATURE.to_string(),
                b"not a signature".to_vec(),
            )));
        }
        if include_agile_signature {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/vbaProjectSignatureAgile.bin").unwrap(),
                content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE.to_string(),
                b"not an Agile signature".to_vec(),
            )));
        }
        package.relate_to("xl/workbook.bin", relationship_type::OFFICE_DOCUMENT);
        package
    }

    #[test]
    fn discovers_legacy_and_agile_signature_metadata_without_parsing_payloads() {
        let package = package_with_vba_project(true, true);
        let workbook = package.main_document_part().unwrap();

        let project = discover_vba_project(&package, workbook).unwrap().unwrap();
        assert_eq!(project.source_part_name().as_str(), "/xl/workbook.bin");
        assert_eq!(project.project_part_name().as_str(), "/xl/vbaProject.bin");
        assert_eq!(project.signatures().len(), 2);
        assert_eq!(
            project.signatures()[0].kind(),
            VbaProjectSignatureKind::Legacy
        );
        assert_eq!(
            project.signatures()[0].part_name().as_str(),
            "/xl/vbaProjectSignature.bin"
        );
        assert_eq!(
            project.signatures()[1].kind(),
            VbaProjectSignatureKind::Agile
        );
        assert_eq!(
            project.signatures()[1].part_name().as_str(),
            "/xl/vbaProjectSignatureAgile.bin"
        );
    }

    #[test]
    fn projects_without_signatures_are_valid() {
        let package = package_with_vba_project(false, false);
        let workbook = package.main_document_part().unwrap();

        let project = discover_vba_project(&package, workbook).unwrap().unwrap();
        assert!(project.signatures().is_empty());
    }

    #[test]
    fn parsed_xlsb_workbook_exposes_inert_project_metadata() {
        let mut writer = XlsbWorkbookWriter::new();
        writer.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        let mut bytes = Cursor::new(Vec::new());
        writer.save(&mut bytes).unwrap();

        let mut workbook = XlsbWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
        let workbook_name = workbook
            .opc_package()
            .main_document_part()
            .unwrap()
            .partname()
            .clone();
        let project = BlobPart::new(
            PackURI::new("/xl/vbaProject.bin").unwrap(),
            content_type::OFC_VBA_PROJECT.to_string(),
            b"intentionally not a compound file".to_vec(),
        );
        let package = workbook.opc_package_mut();
        package.add_part(Box::new(project));
        package
            .get_part_mut(&workbook_name)
            .unwrap()
            .relate_to("vbaProject.bin", relationship_type::VBA_PROJECT);

        let project = workbook.vba_project().unwrap().unwrap();
        assert_eq!(project.source_part_name(), &workbook_name);
        assert!(project.relationship_id().starts_with("rId"));
        assert_eq!(project.project_part_name().as_str(), "/xl/vbaProject.bin");
    }
}
