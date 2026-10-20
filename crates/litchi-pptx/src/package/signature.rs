//! The explicit signature policy of the package-level edit families (ADR
//! 0006; unified-ops record 0014a).
//!
//! A changed edit invalidates every package signature, and removing one is an
//! explicit operation: [`Package::unsign`](crate::Package::unsign), or
//! [`OpcPackage::unsign`] on a raw package. The families that edit an
//! [`OpcPackage`] therefore check, in this order:
//!
//! 1. an exact no-op returns without touching the package, so a signed
//!    package keeps its signature byte for byte;
//! 2. [`refuse_signed_change`] refuses a changed edit of a signed package
//!    before the first mutation;
//! 3. after the mutation, [`settle_unsigned_change`] settles the OPC
//!    signature-policy state of the package, which step 2 proved unsigned.

use litchi_opc::{OpcError, OpcPackage};

use crate::{Error, Result};

/// Whether a changed edit of `package` needs an explicit signature
/// disposition first.
///
/// This holds while the package carries signature infrastructure, and also
/// while its OPC policy still requires a disposition: a signature graph that
/// generic part removal took apart does not count as removed (ADR 0006).
pub(crate) fn requires_disposition(package: &OpcPackage) -> bool {
    package.is_signed() || package.requires_signature_edit_policy()
}

/// Refuse a changed edit of a signed package.
///
/// # Errors
///
/// Returns [`Error::Opc`] with [`OpcError::SignedSourceRequiresExplicitPolicy`]
/// when [`requires_disposition`] holds.
pub(crate) fn refuse_signed_change(package: &OpcPackage) -> Result<()> {
    if requires_disposition(package) {
        return Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy));
    }
    Ok(())
}

/// Settle the OPC signature-policy state of a changed package that carries
/// no signature.
///
/// On such a package [`OpcPackage::unsign`] removes nothing. It marks the
/// absent signature graph as explicitly handled and drops the relationship
/// source captures, which is what every family did unconditionally before
/// record 0014a, so an unsigned package publishes exactly as it did then. A
/// package that still needs a disposition is refused instead: this never
/// removes a signature, even on a path that missed its up-front
/// [`refuse_signed_change`].
///
/// # Errors
///
/// Returns the error of [`refuse_signed_change`].
pub(crate) fn settle_unsigned_change(package: &mut OpcPackage) -> Result<()> {
    refuse_signed_change(package)?;
    package.unsign();
    Ok(())
}

#[cfg(test)]
mod tests {
    use litchi_opc::constants::relationship_type as rt;

    use super::*;

    #[test]
    fn settling_refuses_a_signed_package_and_never_removes_its_signature() {
        let mut package = OpcPackage::new();
        package.relate_to("_xmlsignatures/origin.sigs", rt::DIGITAL_SIGNATURE_ORIGIN);
        assert!(requires_disposition(&package));
        assert!(matches!(
            settle_unsigned_change(&mut package),
            Err(Error::Opc(OpcError::SignedSourceRequiresExplicitPolicy))
        ));
        assert!(package.is_signed());

        package.unsign();
        assert!(!requires_disposition(&package));
        refuse_signed_change(&package).unwrap();
        settle_unsigned_change(&mut package).unwrap();
    }
}
