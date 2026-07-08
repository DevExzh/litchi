//! Person and mention data structures for threaded comments.

/// A person who can author threaded comments.
///
/// Persons are tracked separately in a person list XML part and referenced
/// by their unique ID in threaded comments.
#[derive(Debug, Clone, Default)]
pub struct Person {
    /// Display name of the person
    pub display_name: String,
    /// Unique identifier (GUID)
    pub id: String,
    /// Optional user ID from identity provider
    pub user_id: Option<String>,
    /// Optional provider ID (e.g., Active Directory)
    pub provider_id: Option<String>,
}

/// Collection of persons who can author comments in a workbook.
#[derive(Debug, Clone, Default)]
pub struct PersonList {
    /// List of persons
    pub persons: Vec<Person>,
}

/// An @mention within a threaded comment.
///
/// Mentions allow tagging specific people within comment text using @username syntax.
#[derive(Debug, Clone, Default)]
pub struct Mention {
    /// Person ID being mentioned
    pub mention_person_id: String,
    /// Unique ID for this mention
    pub mention_id: String,
    /// Character offset where mention starts in text
    pub start_index: u32,
    /// Length of the mention text
    pub length: u32,
}
