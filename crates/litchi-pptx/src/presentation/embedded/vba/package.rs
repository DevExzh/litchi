use super::MAX_PAYLOAD_BYTES;
use super::model::Project;
use crate::{Error, Result};
use litchi_ooxml_common::vba::{Host, remove_project_graph, store_project_graph};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};
use std::sync::Arc;

/// Discover the one inert VBA graph attached to a presentation main part.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn discover(package: &OpcPackage, source: &PackURI) -> Result<Option<Project>> {
    let source_part = package.get_part(source)?;
    let mut matches = source_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rt::VBA_PROJECT);
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::Invalid(
            "presentation main part has multiple VBA project relationships".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(Error::Relationship(
            "VBA project relationship cannot be external".to_string(),
        ));
    }
    let project_part_name = relationship.target_partname()?;
    let project_part = package.get_part(&project_part_name)?;
    if project_part.content_type() != ct::OFC_VBA_PROJECT {
        return Err(Error::ContentType {
            expected: ct::OFC_VBA_PROJECT.to_string(),
            actual: project_part.content_type().to_string(),
        });
    }
    if project_part.rels().iter().next().is_some() {
        return Err(Error::Invalid(
            "presentation VBA project part must not have outbound relationships".to_string(),
        ));
    }
    Ok(Some(Project {
        source_part_name: source.clone(),
        relationship_id: relationship.r_id().to_string(),
        project_part_name,
    }))
}

/// Store an opaque VBA payload through the shared transactional graph service.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn store(package: &mut OpcPackage, source: &PackURI, payload: Vec<u8>) -> Result<Project> {
    if payload.is_empty() || payload.len() > MAX_PAYLOAD_BYTES {
        return Err(Error::Limit {
            resource: "PPTX VBA payload bytes",
            limit: MAX_PAYLOAD_BYTES,
        });
    }
    store_project_graph(package, source, Host::PowerPoint, Arc::new(payload), None)
        .map_err(map_common)?;
    discover(package, source)?
        .ok_or_else(|| Error::Invalid("stored VBA project is not discoverable".to_string()))
}

/// Remove the complete VBA relationship graph and restore the non-macro main type.
///
/// # Errors
///
/// Returns an error if the operation fails.
pub fn remove(package: &mut OpcPackage, source: &PackURI) -> Result<bool> {
    remove_project_graph(package, source, Host::PowerPoint).map_err(map_common)
}

fn map_common(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Opc(error) => Error::Opc(error),
        litchi_ooxml_common::Error::Xml(message) => Error::Xml(message),
        litchi_ooxml_common::Error::Missing(message) => Error::PartNotFound(message),
        litchi_ooxml_common::Error::ContentType { expected, actual } => {
            Error::ContentType { expected, actual }
        },
        litchi_ooxml_common::Error::Relationship(message) => Error::Relationship(message),
        litchi_ooxml_common::Error::Invalid(message) => Error::Invalid(message),
        litchi_ooxml_common::Error::Limit { resource, max, .. } => Error::Limit {
            resource,
            limit: max,
        },
        litchi_ooxml_common::Error::Uri(message) => Error::Uri(message),
        litchi_ooxml_common::Error::Mce(error) => Error::MarkupCompatibility(error),
        litchi_ooxml_common::Error::Decode(error) => Error::Decode(error),
        other => Error::Invalid(other.to_string()),
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;
    use litchi_opc::Part;
    use litchi_opc::part::BlobPart;

    #[test]
    fn discovers_and_borrows_an_inert_project() {
        let source = PackURI::new("/ppt/presentation.xml").unwrap();
        let project = PackURI::new("/ppt/vbaProject.bin").unwrap();
        let mut package = OpcPackage::new();
        let mut main = BlobPart::new(
            source.clone(),
            ct::PML_PRESENTATION_MAIN.to_string(),
            b"<p:presentation/>".to_vec(),
        );
        main.relate_to("vbaProject.bin", rt::VBA_PROJECT);
        package.add_part(Box::new(main));
        package.add_part(Box::new(BlobPart::new(
            project,
            ct::OFC_VBA_PROJECT.to_string(),
            vec![1, 2, 3],
        )));
        let value = discover(&package, &source).unwrap().unwrap();
        assert_eq!(value.payload(&package).unwrap(), &[1, 2, 3]);
    }
}
