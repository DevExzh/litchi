//! The signature policy of changed publications (ADR 0006).
//!
//! A change to a signed package invalidates every package signature. Removing
//! a signature is an explicit operation, [`crate::Workbook::unsign`] or
//! [`litchi_opc::OpcPackage::unsign`] (and, for Volatile Dependencies,
//! `without_signatures`), never a side effect of an edit. Every edit family
//! that publishes a changed package checks [`ensure_unsigned`] before it
//! mutates anything, so a changed publication from a signed package is refused
//! with [`Error::Signed`] instead of dropping the signature. An exact no-op
//! publishes nothing and keeps the signature byte for byte.

use litchi_opc::OpcPackage;

use crate::package::error::{Error, Result};

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
