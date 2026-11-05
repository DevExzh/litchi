//! The signature policy of changed package-level publications (ADR 0006).
//!
//! A change to a signed package invalidates every package signature, so
//! removing the signature is an explicit operation and never a side effect of
//! an edit. The package-level families (comments, connections, revisions,
//! `ActiveX`, named sheet views, slicers and timelines, volatile dependencies,
//! XML maps, smart tags and query tables) check [`ensure_unsigned`] before
//! they mutate anything, so a changed publication from a signed package is
//! refused with [`Error::Signed`] instead of dropping the signature. The
//! caller authorizes the removal first with [`crate::Package::unsign`],
//! [`crate::Workbook::unsigned`] or [`litchi_opc::OpcPackage::unsign`]. A call
//! that publishes nothing (removing something absent, a commit with no change)
//! leaves the package, signature included, byte for byte as it was.

use litchi_opc::OpcPackage;

use crate::error::{Error, Result};

/// Refuse a changed publication from `package` while it is signed.
///
/// Present signature infrastructure refuses, and so does the OPC edit-policy
/// requirement that survives generic removal of the signature parts: deleting
/// signature parts by hand does not count as removing the signature.
pub(crate) fn ensure_unsigned(package: &OpcPackage) -> Result<()> {
    if package.is_signed() || package.requires_signature_edit_policy() {
        return Err(Error::Signed);
    }
    Ok(())
}

/// Map an error of a shared `litchi-ooxml-common` owner (custom properties,
/// VBA) to this crate's error, keeping a signature refusal typed as
/// [`Error::Signed`] (unified-ops record 0014a, F1b).
#[cfg(feature = "vba-inspection")]
pub(crate) fn from_shared(error: litchi_ooxml_common::Error) -> Error {
    match error {
        litchi_ooxml_common::Error::Refused(refusal) => Error::Refused(refusal),
        litchi_ooxml_common::Error::PatchConflict => Error::PatchConflict {
            part: "shared OOXML graph".to_owned(),
        },
        litchi_ooxml_common::Error::Signed => Error::Signed,
        litchi_ooxml_common::Error::SourceChanged { expected, observed } => {
            Error::SourceChanged { expected, observed }
        },
        litchi_ooxml_common::Error::Internal(reason) => Error::Internal(reason),
        error => Error::Common(error),
    }
}
