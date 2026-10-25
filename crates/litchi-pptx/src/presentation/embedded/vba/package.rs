use super::MAX_PAYLOAD_BYTES;
use super::model::Project;
use crate::{Error, Result};
use litchi_ooxml_common::vba::{Edit as VbaEdit, EditOptions, Host};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI};

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
/// Storing always republishes the graph, so a signed package is refused with
/// [`Error::Signed`](crate::Error::Signed)
/// (ADR 0006) before it is changed: remove the signature explicitly first,
/// with [`Package::unsign`](crate::Package::unsign) or, on a raw package,
/// [`OpcPackage::unsign`].
///
/// # Errors
///
/// Returns [`Error::Signed`] for a signed package, and an error if the output
/// cannot be encoded or written.
pub(crate) fn store(
    package: &mut OpcPackage,
    source: &PackURI,
    payload: Vec<u8>,
) -> Result<Project> {
    if payload.is_empty() || payload.len() > MAX_PAYLOAD_BYTES {
        return Err(Error::Limit {
            kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::PayloadBytes),
            observed: payload.len(),
            maximum: MAX_PAYLOAD_BYTES,
        });
    }
    let mut edit =
        VbaEdit::new(package, Host::PowerPoint, EditOptions::new()).map_err(map_common)?;
    edit.store(payload, None).map_err(map_common)?;
    *package = edit.commit().map_err(map_common)?.into_package();
    discover(package, source)?
        .ok_or_else(|| Error::Invalid("stored VBA project is not discoverable".to_string()))
}

/// Remove the complete VBA relationship graph and restore the non-macro main type.
///
/// An absent graph returns `Ok(false)` and keeps a signed package's
/// signature; removing a present graph from a signed package is refused with
/// [`Error::Signed`](crate::Error::Signed)
/// (ADR 0006): remove the signature explicitly first, with
/// [`Package::unsign`](crate::Package::unsign) or, on a raw package,
/// [`OpcPackage::unsign`].
///
/// # Errors
///
/// Returns [`Error::Signed`] for a change to a signed package, and an error if
/// the operation fails.
pub(crate) fn remove(package: &mut OpcPackage, _source: &PackURI) -> Result<bool> {
    let mut edit =
        VbaEdit::new(package, Host::PowerPoint, EditOptions::new()).map_err(map_common)?;
    let changed = edit.remove().map_err(map_common)?;
    *package = edit.commit().map_err(map_common)?.into_package();
    Ok(changed)
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
        litchi_ooxml_common::Error::PatchConflict => Error::PatchConflict,
        litchi_ooxml_common::Error::Refused(refusal) => Error::Refused(refusal),
        litchi_ooxml_common::Error::Signed => Error::Signed,
        litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
            Error::SourceChanged { expected, observed }
        },
        litchi_ooxml_common::Error::Internal(reason) => Error::Internal {
            operation: "vba",
            reason,
        },
        litchi_ooxml_common::Error::Invalid(message) => Error::Invalid(message),
        litchi_ooxml_common::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: crate::opened::LimitKind::Common(kind),
            observed,
            maximum,
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
