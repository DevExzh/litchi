//! Inert MS-XLSB VBA-project and signature relationship discovery.
//!
//! Discovery validates OPC relationship and content-type metadata only. It
//! does not inspect, parse, verify, or execute VBA project or signature
//! bytes.

use crate::error::{OoxmlError, Result};
use crate::vba_package::{ensure_exclusive_inbound_reference, validate_vba_project_payload};
use litchi_opc::constants::{content_type, relationship_type};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part, Relationship};

const PROJECT_PART: &str = "/xl/vbaProject.bin";
const PROJECT_TARGET: &str = "vbaProject.bin";

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

    /// Parse the `vbaProject.bin` CFB payload as inert MS-OVBA source.
    ///
    /// Declared signature payloads remain separate opaque relationship
    /// metadata and are not treated as trust decisions.
    pub fn read_project(
        &self,
        package: &OpcPackage,
        limits: &crate::vba::VbaLimits,
    ) -> std::result::Result<crate::vba::VbaProject, crate::vba::VbaError> {
        crate::vba::read_project_part(package, &self.project_part_name, limits)
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
    let main = package.main_document_part()?;
    if source.partname() != main.partname() || source.content_type() != content_type::XLSB_BIN {
        return Err(OoxmlError::InvalidFormat(format!(
            "XLSB VBA source '{}' is not the binary Workbook main part",
            source.partname().as_str()
        )));
    }
    let mut projects = source
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type::VBA_PROJECT);
    let Some(relationship) = projects.next() else {
        ensure_no_orphan_vba_parts(package)?;
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
    let project = VbaProject {
        source_part_name: source.partname().clone(),
        relationship_id: relationship.r_id().to_string(),
        project_part_name,
        signatures,
    };
    ensure_graph_parts_are_unique(package, &project)?;
    ensure_exclusive_inbound_reference(
        package,
        &project.project_part_name,
        &project.source_part_name,
        &project.relationship_id,
    )?;
    for signature in &project.signatures {
        ensure_exclusive_inbound_reference(
            package,
            &signature.part_name,
            &project.project_part_name,
            &signature.relationship_id,
        )?;
    }
    Ok(Some(project))
}

/// Attach a validated VBA project, dropping stale project and package signatures.
pub(crate) fn store_vba_project(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Vec<u8>,
    limits: &crate::vba::VbaLimits,
) -> Result<VbaProject> {
    validate_vba_project_payload(&payload, limits)?;
    let existing = {
        let source_part = package.get_part(source)?;
        discover_vba_project(package, source_part)?
    };
    let canonical =
        PackURI::new(PROJECT_PART).map_err(|error| OoxmlError::InvalidUri(error.to_string()))?;
    if existing
        .as_ref()
        .is_none_or(|project| project.project_part_name != canonical)
    {
        package.validate_new_part_name(&canonical)?;
    }

    if let Some(project) = existing {
        remove_discovered_graph(package, &project)?;
    }
    package.try_add_part(Box::new(BlobPart::new(
        canonical,
        content_type::OFC_VBA_PROJECT.to_string(),
        payload,
    )))?;
    package
        .get_part_mut(source)?
        .relate_to(PROJECT_TARGET, relationship_type::VBA_PROJECT);
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "failed to clear signatures after storing an XLSB VBA project: {error}"
        ))
    })?;

    let source_part = package.get_part(source)?;
    discover_vba_project(package, source_part)?.ok_or_else(|| {
        OoxmlError::InvalidFormat("stored XLSB VBA project was not discoverable".to_string())
    })
}

/// Remove an XLSB VBA project and all declared project-signature parts.
pub(crate) fn remove_vba_project(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    let existing = {
        let source_part = package.get_part(source)?;
        discover_vba_project(package, source_part)?
    };
    let Some(project) = existing else {
        return Ok(false);
    };
    remove_discovered_graph(package, &project)?;
    package.clear_digital_signatures().map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "failed to clear signatures after removing an XLSB VBA project: {error}"
        ))
    })?;
    Ok(true)
}

fn remove_discovered_graph(package: &mut OpcPackage, project: &VbaProject) -> Result<()> {
    package
        .get_part_mut(&project.source_part_name)?
        .rels_mut()
        .remove(&project.relationship_id);
    package.remove_part(&project.project_part_name);
    for signature in &project.signatures {
        package.remove_part(&signature.part_name);
    }
    Ok(())
}

fn ensure_no_orphan_vba_parts(package: &OpcPackage) -> Result<()> {
    if package.iter_parts().any(|part| {
        matches!(
            part.content_type(),
            content_type::OFC_VBA_PROJECT
                | content_type::OFC_VBA_PROJECT_SIGNATURE
                | content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE
        )
    }) {
        return Err(OoxmlError::InvalidFormat(
            "XLSB package contains an orphan VBA project or project-signature part".to_string(),
        ));
    }
    Ok(())
}

fn ensure_graph_parts_are_unique(package: &OpcPackage, project: &VbaProject) -> Result<()> {
    let legacy = project
        .signatures
        .iter()
        .find(|signature| signature.kind == VbaProjectSignatureKind::Legacy)
        .map(|signature| &signature.part_name);
    let agile = project
        .signatures
        .iter()
        .find(|signature| signature.kind == VbaProjectSignatureKind::Agile)
        .map(|signature| &signature.part_name);
    for part in package.iter_parts() {
        let expected = match part.content_type() {
            content_type::OFC_VBA_PROJECT => Some(&project.project_part_name),
            content_type::OFC_VBA_PROJECT_SIGNATURE => legacy,
            content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE => agile,
            _ => continue,
        };
        if expected != Some(part.partname()) {
            return Err(OoxmlError::InvalidFormat(format!(
                "XLSB package contains an unexpected VBA graph part '{}'",
                part.partname().as_str()
            )));
        }
    }
    Ok(())
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
    use crate::vba::{VbaLimits, VbaModuleBuilder, VbaProjectBuilder};
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

    fn authored_project() -> crate::vba::VbaProjectBinary {
        VbaProjectBuilder::new("BinaryWorkbookProject")
            .with_module(VbaModuleBuilder::standard(
                "Module1",
                "Public Sub Hello()\r\nEnd Sub\r\n",
            ))
            .build(&VbaLimits::default())
            .unwrap()
    }

    fn generated_workbook() -> XlsbWorkbook {
        let mut writer = XlsbWorkbookWriter::new();
        writer.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        let mut bytes = Cursor::new(Vec::new());
        writer.save(&mut bytes).unwrap();
        XlsbWorkbook::new(Cursor::new(bytes.into_inner())).unwrap()
    }

    #[test]
    fn parsed_workbook_stores_round_trips_and_removes_authored_project() {
        let mut workbook = generated_workbook();
        workbook.set_vba_project(&authored_project()).unwrap();

        let mut bytes = Cursor::new(Vec::new());
        workbook.save(&mut bytes).unwrap();
        let mut reopened = XlsbWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
        let metadata = reopened.vba_project().unwrap().unwrap();
        let parsed = metadata
            .read_project(reopened.opc_package(), &VbaLimits::default())
            .unwrap();
        assert_eq!(parsed.name(), "BinaryWorkbookProject");
        assert_eq!(parsed.modules().len(), 1);

        assert!(reopened.remove_vba_project().unwrap());
        assert!(reopened.vba_project().unwrap().is_none());
        assert!(
            reopened
                .opc_package()
                .get_part(&PackURI::new(PROJECT_PART).unwrap())
                .is_err()
        );
    }

    #[test]
    fn workbook_writer_attaches_validated_project() {
        let mut writer = XlsbWorkbookWriter::new();
        writer.add_worksheet(MutableXlsbWorksheet::new("Sheet1"));
        writer.set_vba_project(&authored_project()).unwrap();
        let mut bytes = Cursor::new(Vec::new());
        writer.save(&mut bytes).unwrap();

        let workbook = XlsbWorkbook::new(Cursor::new(bytes.into_inner())).unwrap();
        let metadata = workbook.vba_project().unwrap().unwrap();
        assert_eq!(metadata.project_part_name().as_str(), PROJECT_PART);
        assert_eq!(
            metadata
                .read_project(workbook.opc_package(), &VbaLimits::default())
                .unwrap()
                .name(),
            "BinaryWorkbookProject"
        );
    }

    #[test]
    fn replacement_removes_stale_legacy_and_agile_signatures() {
        let mut workbook = generated_workbook();
        workbook.set_vba_project(&authored_project()).unwrap();
        let project_name = PackURI::new(PROJECT_PART).unwrap();
        let legacy_name = PackURI::new("/xl/vbaProjectSignature.bin").unwrap();
        let agile_name = PackURI::new("/xl/vbaProjectSignatureAgile.bin").unwrap();
        {
            let package = workbook.opc_package_mut();
            let project = package.get_part_mut(&project_name).unwrap();
            project.relate_to(
                "vbaProjectSignature.bin",
                relationship_type::VBA_PROJECT_SIGNATURE,
            );
            project.relate_to(
                "vbaProjectSignatureAgile.bin",
                relationship_type::VBA_PROJECT_SIGNATURE_AGILE,
            );
            package.add_part(Box::new(BlobPart::new(
                legacy_name.clone(),
                content_type::OFC_VBA_PROJECT_SIGNATURE.to_string(),
                b"opaque legacy signature".to_vec(),
            )));
            package.add_part(Box::new(BlobPart::new(
                agile_name.clone(),
                content_type::OFC_VBA_PROJECT_SIGNATURE_AGILE.to_string(),
                b"opaque agile signature".to_vec(),
            )));
        }
        assert_eq!(
            workbook.vba_project().unwrap().unwrap().signatures().len(),
            2
        );

        workbook.set_vba_project(&authored_project()).unwrap();
        assert!(
            workbook
                .vba_project()
                .unwrap()
                .unwrap()
                .signatures()
                .is_empty()
        );
        assert!(workbook.opc_package().get_part(&legacy_name).is_err());
        assert!(workbook.opc_package().get_part(&agile_name).is_err());
    }

    #[test]
    fn removal_deletes_the_complete_signed_project_graph() {
        let mut package = package_with_vba_project(true, true);
        let source = PackURI::new("/xl/workbook.bin").unwrap();

        assert!(remove_vba_project(&mut package, &source).unwrap());
        assert!(
            discover_vba_project(&package, package.get_part(&source).unwrap())
                .unwrap()
                .is_none()
        );
        for part_name in [
            PROJECT_PART,
            "/xl/vbaProjectSignature.bin",
            "/xl/vbaProjectSignatureAgile.bin",
        ] {
            assert!(package.get_part(&PackURI::new(part_name).unwrap()).is_err());
        }
    }

    #[test]
    fn rejects_invalid_and_orphan_project_data_without_mutation() {
        let mut workbook = generated_workbook();
        assert!(
            workbook
                .set_vba_project_bytes(vec![0; 64], &VbaLimits::default())
                .is_err()
        );
        assert!(workbook.vba_project().unwrap().is_none());

        workbook.opc_package_mut().add_part(Box::new(BlobPart::new(
            PackURI::new("/xl/orphanSignature.bin").unwrap(),
            content_type::OFC_VBA_PROJECT_SIGNATURE.to_string(),
            b"orphan".to_vec(),
        )));
        assert!(workbook.vba_project().is_err());
        assert!(workbook.set_vba_project(&authored_project()).is_err());
        assert!(
            workbook
                .opc_package()
                .get_part(&PackURI::new("/xl/orphanSignature.bin").unwrap())
                .is_ok()
        );
    }

    #[test]
    fn canonical_name_conflict_preserves_an_existing_noncanonical_graph() {
        let mut workbook = generated_workbook();
        workbook.set_vba_project(&authored_project()).unwrap();
        let metadata = workbook.vba_project().unwrap().unwrap();
        let source_name = metadata.source_part_name().clone();
        let relationship_id = metadata.relationship_id().to_string();
        let canonical_name = metadata.project_part_name().clone();
        let payload = workbook
            .opc_package()
            .get_part(&canonical_name)
            .unwrap()
            .blob()
            .to_vec();
        let custom_name = PackURI::new("/xl/customVba.bin").unwrap();
        {
            let package = workbook.opc_package_mut();
            package
                .get_part_mut(&source_name)
                .unwrap()
                .rels_mut()
                .remove(&relationship_id);
            package.remove_part(&canonical_name);
            package.add_part(Box::new(BlobPart::new(
                custom_name.clone(),
                content_type::OFC_VBA_PROJECT.to_string(),
                payload.clone(),
            )));
            package
                .get_part_mut(&source_name)
                .unwrap()
                .relate_to("customVba.bin", relationship_type::VBA_PROJECT);
            package.add_part(Box::new(BlobPart::new(
                canonical_name.clone(),
                "application/octet-stream".to_string(),
                b"occupied".to_vec(),
            )));
        }

        assert!(workbook.set_vba_project(&authored_project()).is_err());
        let preserved = workbook.vba_project().unwrap().unwrap();
        assert_eq!(preserved.project_part_name(), &custom_name);
        assert_eq!(
            workbook
                .opc_package()
                .get_part(&custom_name)
                .unwrap()
                .blob(),
            payload
        );
        assert_eq!(
            workbook
                .opc_package()
                .get_part(&canonical_name)
                .unwrap()
                .blob(),
            b"occupied"
        );
    }
}
