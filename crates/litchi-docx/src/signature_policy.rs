//! The signature policy of changed package-level publications (ADR 0006).
//!
//! A change to a signed package invalidates every package signature, so
//! removing the signature is an explicit operation and never a side effect of
//! an edit. The package-level families (mail merge, glossary, web settings,
//! the font table, the managed font settings and the Ribbon) check
//! [`ensure_unsigned`] before they mutate anything, so a changed publication
//! from a signed package is refused with
//! [`OpcError::SignedSourceRequiresExplicitPolicy`] instead of dropping the
//! signature. The caller authorizes the removal first with [`crate::Package::unsign`] or
//! [`litchi_opc::OpcPackage::unsign`]. A call that publishes nothing (removing
//! something absent, storing an equal value) leaves the package, signature
//! included, byte for byte as it was.

use litchi_opc::{OpcError, OpcPackage};

use crate::{Error, Result};

/// Refuse a changed publication from `package` while it is signed.
///
/// Present signature infrastructure refuses, and so does the OPC edit-policy
/// requirement that survives generic removal of the signature parts: deleting
/// signature parts by hand does not count as removing the signature.
pub(crate) fn ensure_unsigned(package: &OpcPackage) -> Result<()> {
    if is_signed(package) {
        return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
    }
    Ok(())
}

/// Whether a change to `package` would need an explicit signature
/// disposition first.
pub(crate) fn is_signed(package: &OpcPackage) -> bool {
    package.is_signed() || package.requires_signature_edit_policy()
}
