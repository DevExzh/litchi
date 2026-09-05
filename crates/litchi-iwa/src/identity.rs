//! Independent document identities for newly-created iWork packages.

/// The three independent UUIDs assigned to a newly-created iWork document.
///
/// Apple stores the public document UUID, the current saved-version UUID, and
/// a private UUID separately. Keeping these values distinct prevents two
/// documents derived from a common source from being treated as revisions of
/// one another by iWork or iCloud.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct IWorkDocumentIdentity {
    document_uuid: String,
    version_uuid: String,
    private_uuid: String,
}

impl IWorkDocumentIdentity {
    /// Generate three fresh, mutually distinct RFC 4122 version 4 UUIDs.
    pub(crate) fn generate() -> Self {
        let document_uuid = generate_uuid_string();
        let version_uuid = generate_distinct_uuid(&[&document_uuid]);
        let private_uuid = generate_distinct_uuid(&[&document_uuid, &version_uuid]);
        Self {
            document_uuid,
            version_uuid,
            private_uuid,
        }
    }

    /// Stable UUID used by `Metadata/DocumentIdentifier` and sharing metadata.
    pub(crate) fn document_uuid(&self) -> &str {
        &self.document_uuid
    }

    /// UUID of the current saved version and package revision.
    pub(crate) fn version_uuid(&self) -> &str {
        &self.version_uuid
    }

    /// Private UUID used by iWork's local document bookkeeping.
    pub(crate) fn private_uuid(&self) -> &str {
        &self.private_uuid
    }
}

fn generate_uuid_string() -> String {
    let braced = litchi_core::id::generate_guid_braced();
    braced[1..braced.len() - 1].to_owned()
}

fn generate_distinct_uuid(existing: &[&str]) -> String {
    loop {
        let candidate = generate_uuid_string();
        if existing.iter().all(|value| *value != candidate) {
            return candidate;
        }
    }
}
