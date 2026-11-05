//! Shared helpers of the typed XLSB package edit families.

use std::hash::{Hash, Hasher};

use litchi_opc::{OpcPackage, PackURI};

use super::error::Result;

/// Borrowed ASCII-case-insensitive key without a lowercase allocation.
#[derive(Debug, Clone, Copy)]
pub(crate) struct Caseless<'a>(&'a str);

impl<'a> Caseless<'a> {
    pub(crate) const fn new(value: &'a str) -> Self {
        Self(value)
    }
}

impl PartialEq for Caseless<'_> {
    fn eq(&self, other: &Self) -> bool {
        self.0.eq_ignore_ascii_case(other.0)
    }
}

impl Eq for Caseless<'_> {}

impl Hash for Caseless<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        for byte in self.0.bytes() {
            state.write_u8(byte.to_ascii_lowercase());
        }
    }
}

/// Refuse deletion or replacement when an owned target has another inbound edge.
pub(crate) fn require_exclusive_inbound(
    package: &OpcPackage,
    owner_part: &PackURI,
    targets: &[(String, PackURI)],
    _description: &'static str,
) -> Result<()> {
    for relationship in package.rels().iter() {
        if relationship.is_external() {
            continue;
        }
        let target = relationship.target_partname()?;
        if targets
            .iter()
            .any(|(_, owned_target)| *owned_target == target)
        {
            return Err(super::error::Error::InvalidRelationship(
                "target is also referenced from the package root",
            ));
        }
    }
    for source in package.iter_parts() {
        for relationship in source.rels().iter() {
            if relationship.is_external() {
                continue;
            }
            let target = relationship.target_partname()?;
            let Some((relationship_id, _)) = targets
                .iter()
                .find(|(_, owned_target)| *owned_target == target)
            else {
                continue;
            };
            if source.partname() != owner_part || relationship.r_id() != relationship_id {
                return Err(super::error::Error::InvalidRelationship(
                    "target has a non-owner inbound relationship",
                ));
            }
        }
    }
    Ok(())
}
