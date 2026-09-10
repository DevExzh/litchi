//! Atomic package publication for worksheet binary-index maintenance.
//!
//! [`crate::binary_index`] owns the bounded BIFF12 index codec.  This module
//! owns the OPC relationship boundary used by worksheet mutations: it finds
//! the index part only when a worksheet stream actually changed, validates the
//! source relationship and content type, and stages the resulting index on the
//! caller's candidate package.

use crate::binary_index::{self, Limits};
use crate::package::error::{Error, Result};
use litchi_opc::{OpcPackage, PackURI, Part};

const BINARY_INDEX_RELATIONSHIP: &str = binary_index::BINARY_INDEX_RELATIONSHIP;
const BINARY_INDEX_CONTENT_TYPE: &str = binary_index::BINARY_INDEX_CONTENT_TYPE;
const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";

/// Maintain the binary index owned by one worksheet candidate.
///
/// The caller must invoke this before publishing `after_worksheet` into its
/// candidate package.  The candidate is normally a clone owned by the caller,
/// so a malformed or unsupported index leaves the source package untouched.
/// Exact worksheet no-ops return before relationship or index inspection.
///
/// An existing source index is never silently removed.  The byte-level codec
/// retains it exactly when its indexed geometry remains valid, patches known
/// offsets when that is proven safe, and refuses unsupported topology or
/// opaque-record changes.
pub(crate) fn maintain(
    package: &mut OpcPackage,
    worksheet: &PackURI,
    before_worksheet: &[u8],
    after_worksheet: &[u8],
) -> Result<()> {
    if before_worksheet == after_worksheet {
        return Ok(());
    }

    let index_target = find_index_target(package, worksheet)?;
    let Some(index_target) = index_target else {
        // Preserve the source topology.  Package-only edits do not invent a
        // binary-index part when the source worksheet had none.
        return Ok(());
    };

    let source_index = {
        let index_part = package.get_part(&index_target)?;
        require_index_part(index_part)?;
        if !index_part.rels().is_empty() {
            return Err(Error::InvalidRelationship(
                "worksheet binary index part must not have relationships".to_string(),
            ));
        }
        index_part.blob_arc()
    };

    let maintained = binary_index::maintain_index(
        Some(source_index.as_slice()),
        before_worksheet,
        after_worksheet,
        Limits::publication_default(),
    )?
    .ok_or_else(|| {
        Error::InvalidFormat(
            "worksheet binary index maintenance unexpectedly removed an existing index".to_string(),
        )
    })?;

    // Avoid replacing an equal allocation.  This keeps exact source bytes and
    // the OPC part's sharing shape stable when only worksheet bytes changed in
    // a way that left all indexed offsets valid.
    if maintained.as_slice() != source_index.as_slice() {
        package.get_part_mut(&index_target)?.set_blob(maintained);
    }
    Ok(())
}

fn find_index_target(package: &OpcPackage, worksheet: &PackURI) -> Result<Option<PackURI>> {
    let worksheet_part = package.get_part(worksheet)?;
    let mut matches = worksheet_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == BINARY_INDEX_RELATIONSHIP);
    let Some(relationship) = matches.next() else {
        return Ok(None);
    };
    if matches.next().is_some() {
        return Err(Error::InvalidRelationship(
            "worksheet has multiple binary-index relationships".to_string(),
        ));
    }
    if relationship.is_external() {
        return Err(Error::InvalidRelationship(
            "worksheet binary-index relationship cannot be external".to_string(),
        ));
    }
    let target = relationship.target_partname()?;
    if target.is_equivalent_to(worksheet) {
        return Err(Error::InvalidRelationship(
            "worksheet binary-index relationship targets its worksheet".to_string(),
        ));
    }

    // A binary-index part is owned by exactly one worksheet.  Resolving only
    // the selected worksheet's outgoing edge is insufficient: two worksheet
    // parts can happen to point at the same target, and changing one stream
    // would then publish an index that describes the other stream as well.
    // Scan all incoming binary-index edges before reading or replacing the
    // target so an ambiguous graph fails atomically.
    let mut owners = 0usize;
    for owner in package.iter_parts() {
        let owner_is_worksheet = owner.content_type() == WORKSHEET_CONTENT_TYPE;
        for incoming in owner
            .rels()
            .iter()
            .filter(|incoming| incoming.reltype() == BINARY_INDEX_RELATIONSHIP)
        {
            if incoming.is_external() {
                // An external edge cannot share the selected internal target;
                // leave unrelated malformed topology to the owner that edits
                // that relationship.
                continue;
            }
            let incoming_target = incoming.target_partname()?;
            if incoming_target.is_equivalent_to(&target) {
                if !owner_is_worksheet {
                    return Err(Error::InvalidRelationship(
                        "worksheet binary-index target has a non-worksheet owner".to_string(),
                    ));
                }
                owners = owners.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet binary-index owners",
                })?;
            }
        }
    }
    for incoming in package
        .rels()
        .iter()
        .filter(|incoming| incoming.reltype() == BINARY_INDEX_RELATIONSHIP)
    {
        if incoming.is_external() {
            continue;
        }
        if incoming.target_partname()?.is_equivalent_to(&target) {
            return Err(Error::InvalidRelationship(
                "worksheet binary-index target has a package owner".to_string(),
            ));
        }
    }
    if owners != 1 {
        return Err(Error::InvalidRelationship(
            "worksheet binary-index target does not have exactly one worksheet owner".to_string(),
        ));
    }
    Ok(Some(target))
}

fn require_index_part(part: &dyn Part) -> Result<()> {
    if part.content_type() == BINARY_INDEX_CONTENT_TYPE {
        Ok(())
    } else {
        Err(Error::InvalidContentType {
            expected: BINARY_INDEX_CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        })
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::{BlobPart, OpcPackage, Part};

    fn worksheet_uri(name: &str) -> PackURI {
        PackURI::new(format!("/xl/worksheets/{name}.bin")).unwrap()
    }

    fn index_uri() -> PackURI {
        PackURI::new("/xl/worksheets/binaryIndex.bin").unwrap()
    }

    fn worksheet_part(name: &str, bytes: &[u8], target: Option<&str>) -> BlobPart {
        let mut part = BlobPart::new(
            worksheet_uri(name),
            WORKSHEET_CONTENT_TYPE.to_string(),
            bytes.to_vec(),
        );
        if let Some(target) = target {
            part.relate_to(target, BINARY_INDEX_RELATIONSHIP);
        }
        part
    }

    #[test]
    fn shared_index_owner_is_rejected_before_any_publication() {
        let worksheet = worksheet_uri("sheet1");
        let index = index_uri();
        let original_index = vec![0xff, 0x00, 0x01];
        let mut package = OpcPackage::new();
        package.add_part(Box::new(worksheet_part(
            "sheet1",
            b"sheet-one",
            Some("binaryIndex.bin"),
        )));
        package.add_part(Box::new(worksheet_part(
            "sheet2",
            b"sheet-two",
            Some("binaryIndex.bin"),
        )));
        package.add_part(Box::new(BlobPart::new(
            index.clone(),
            BINARY_INDEX_CONTENT_TYPE.to_string(),
            original_index.clone(),
        )));

        let result = maintain(&mut package, &worksheet, b"before", b"after");
        assert!(matches!(
            result,
            Err(Error::InvalidRelationship(message))
                if message.contains("exactly one worksheet owner")
        ));
        assert_eq!(package.get_part(&index).unwrap().blob(), original_index);
        assert_eq!(package.get_part(&worksheet).unwrap().blob(), b"sheet-one");
    }

    #[test]
    fn absent_index_is_not_synthesized_for_a_changed_worksheet() {
        let worksheet = worksheet_uri("sheet1");
        let mut package = OpcPackage::new();
        package.add_part(Box::new(worksheet_part("sheet1", b"before", None)));
        let part_count = package.part_count();

        maintain(&mut package, &worksheet, b"before", b"after").unwrap();

        assert_eq!(package.part_count(), part_count);
        assert_eq!(package.get_part(&worksheet).unwrap().blob(), b"before");
    }

    #[test]
    fn no_op_bypasses_a_malformed_source_index() {
        let worksheet = worksheet_uri("sheet1");
        let index = index_uri();
        let malformed_index = vec![0xff, 0xff, 0xff];
        let mut package = OpcPackage::new();
        package.add_part(Box::new(worksheet_part(
            "sheet1",
            b"same",
            Some("binaryIndex.bin"),
        )));
        package.add_part(Box::new(BlobPart::new(
            index.clone(),
            BINARY_INDEX_CONTENT_TYPE.to_string(),
            malformed_index.clone(),
        )));

        maintain(&mut package, &worksheet, b"same", b"same").unwrap();

        assert_eq!(package.get_part(&index).unwrap().blob(), malformed_index);
    }
}
