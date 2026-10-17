//! The header and footer parts a document relates, listed without decoding
//! them (record 0319, L-DOCX-R5).

use crate::error::Result;

use super::Role;

/// One header or footer part the main document relates: its role, its
/// relationship identifier and its part name.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Entry {
    role: Role,
    relationship_id: String,
    name: String,
}

impl Entry {
    pub(crate) const fn new(role: Role, relationship_id: String, name: String) -> Self {
        Self {
            role,
            relationship_id,
            name,
        }
    }

    /// Whether the part is a header or a footer.
    #[must_use]
    pub const fn role(&self) -> Role {
        self.role
    }

    /// The main part's relationship identifier of the part, which section
    /// properties reference.
    #[must_use]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// The part name (`/word/header1.xml`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}

/// The header and footer parts `relationships` (the main part's) relate, in
/// relationship order, each part once.
pub(crate) fn entries_of(relationships: &litchi_opc::Relationships) -> Result<Vec<Entry>> {
    let mut entries: Vec<Entry> = Vec::new();
    for relationship in relationships.iter() {
        if relationship.is_external() {
            continue;
        }
        let role = if relationship.reltype().ends_with("/header") {
            Role::Header
        } else if relationship.reltype().ends_with("/footer") {
            Role::Footer
        } else {
            continue;
        };
        let name = relationship.target_partname()?.as_str().to_owned();
        if entries.iter().any(|entry| entry.name == name) {
            continue;
        }
        entries.push(Entry::new(role, relationship.r_id().to_owned(), name));
    }
    Ok(entries)
}
