//! Package-neutral threaded-comments data graph.

/// A person who authored or is mentioned by a threaded comment.
#[derive(Debug, Clone, Default)]
pub struct Person {
    /// Display name of the person.
    pub display_name: String,
    /// Stable GUID identifying the person.
    pub id: String,
    /// Optional provider-issued user identifier.
    pub user_id: Option<String>,
    /// Optional provider identifier.
    pub provider_id: Option<String>,
}

/// The workbook-level collection of people.
#[derive(Debug, Clone, Default)]
pub struct People {
    /// People in document order.
    pub persons: Vec<Person>,
}

/// A person mention within a comment's UTF-16 text.
#[derive(Debug, Clone, Default)]
pub struct Mention {
    /// GUID of the mentioned person.
    pub mention_person_id: String,
    /// Stable GUID identifying the mention.
    pub mention_id: String,
    /// UTF-16 start offset in the containing comment text.
    pub start_index: u32,
    /// UTF-16 length of the mention range.
    pub length: u32,
}

/// One threaded comment, either a thread root or a reply.
#[derive(Debug, Clone, Default)]
pub struct Comment {
    /// Cell reference carried by a thread root.
    pub cell_ref: Option<String>,
    /// Stable GUID identifying the comment.
    pub id: String,
    /// GUID of the parent root for a reply.
    pub parent_id: Option<String>,
    /// GUID of the authoring person.
    pub person_id: String,
    /// Unformatted comment text.
    pub text: Option<String>,
    /// Creation timestamp.
    pub date_time: Option<String>,
    /// Whether the thread is resolved.
    pub done: Option<bool>,
    /// Mentions in document order.
    pub mentions: Vec<Mention>,
}

/// The comments carried by one worksheet part.
#[derive(Debug, Clone, Default)]
pub struct Comments {
    /// Comments in serialized order.
    pub comments: Vec<Comment>,
}

/// A workbook persons part plus its host relationship identity.
#[derive(Debug, Clone)]
pub struct PeoplePart {
    /// Relationship ID on the workbook part.
    pub relationship_id: String,
    /// OPC part name retained by the host adapter.
    pub part_name: String,
    /// Person data in this part.
    pub persons: People,
}

/// A worksheet threaded-comments part plus its host relationship identity.
#[derive(Debug, Clone)]
pub struct CommentsPart {
    /// Worksheet OPC part name retained by the host adapter.
    pub worksheet_part_name: String,
    /// Relationship ID on the worksheet part.
    pub relationship_id: String,
    /// OPC part name retained by the host adapter.
    pub part_name: String,
    /// Comment data in this part.
    pub comments: Comments,
}

/// Complete package-neutral threaded-comments graph.
#[derive(Debug, Clone, Default)]
pub struct Graph {
    /// Optional workbook-level people part.
    pub persons: Option<PeoplePart>,
    /// Worksheet comment parts in deterministic host order.
    pub worksheets: Vec<CommentsPart>,
}
