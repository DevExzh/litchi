//! Inert package protection inventory and explicit rewrite disposition.

/// Exact rewrite disposition implied by package signature and encryption state.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ProtectionDisposition {
    /// No signature or encrypted manifest entry blocks ordinary semantic rewrite.
    RewriteAllowed,
    /// Rewriting would invalidate at least one retained package signature.
    RefuseSignedRewrite,
    /// Rewriting cannot preserve at least one encrypted package member.
    RefuseEncryptedRewrite,
    /// Both signature invalidation and encrypted-member preservation block rewrite.
    RefuseSignedAndEncryptedRewrite,
}

/// Inert package signature and encrypted-member inventory.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ProtectionInventory {
    signature_members: Vec<String>,
    encrypted_members: Vec<String>,
}

impl ProtectionInventory {
    pub(crate) fn new(
        mut signature_members: Vec<String>,
        mut encrypted_members: Vec<String>,
    ) -> Self {
        signature_members.sort_unstable();
        signature_members.dedup();
        encrypted_members.sort_unstable();
        encrypted_members.dedup();
        Self {
            signature_members,
            encrypted_members,
        }
    }

    /// Returns exact safe package paths containing signature metadata.
    #[must_use]
    pub fn signature_members(&self) -> &[String] {
        &self.signature_members
    }

    /// Returns exact manifest paths carrying encryption descriptors.
    #[must_use]
    pub fn encrypted_members(&self) -> &[String] {
        &self.encrypted_members
    }

    /// Returns whether any signature metadata member is present.
    #[must_use]
    pub fn is_signed(&self) -> bool {
        !self.signature_members.is_empty()
    }

    /// Returns whether any manifest member has an encryption descriptor.
    #[must_use]
    pub fn is_encrypted(&self) -> bool {
        !self.encrypted_members.is_empty()
    }

    /// Returns the explicit ordinary-rewrite disposition for this inventory.
    #[must_use]
    pub fn disposition(&self) -> ProtectionDisposition {
        match (self.is_signed(), self.is_encrypted()) {
            (false, false) => ProtectionDisposition::RewriteAllowed,
            (true, false) => ProtectionDisposition::RefuseSignedRewrite,
            (false, true) => ProtectionDisposition::RefuseEncryptedRewrite,
            (true, true) => ProtectionDisposition::RefuseSignedAndEncryptedRewrite,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn protection_disposition_covers_every_inventory_state() {
        assert_eq!(
            ProtectionInventory::default().disposition(),
            ProtectionDisposition::RewriteAllowed
        );
        assert_eq!(
            ProtectionInventory::new(vec!["META-INF/documentsignatures.xml".into()], Vec::new())
                .disposition(),
            ProtectionDisposition::RefuseSignedRewrite
        );
        assert_eq!(
            ProtectionInventory::new(Vec::new(), vec!["content.xml".into()]).disposition(),
            ProtectionDisposition::RefuseEncryptedRewrite
        );
        assert_eq!(
            ProtectionInventory::new(
                vec!["META-INF/macrosignatures.xml".into()],
                vec!["content.xml".into()],
            )
            .disposition(),
            ProtectionDisposition::RefuseSignedAndEncryptedRewrite
        );
    }
}
