//! Inert semantic values for XLSB threaded comments.
//!
//! XLSB stores the worksheet stream in BIFF12, but its threaded-comment and
//! persons parts use the SpreadsheetML XML structures described by
//! `[MS-XLSB]` sections 2.1.17--2.1.18 and `[MS-XLSX]` sections 2.3.7 and
//! 2.6.202--2.6.207.  These values intentionally contain no identity provider,
//! network, rendering, or collaboration behavior.

/// An unknown XML attribute retained in its original qualified spelling.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawAttribute {
    /// Qualified attribute name, for example `x14:foo`.
    pub name: String,
    /// Decoded attribute value.
    pub value: String,
}

/// An unknown XML child payload retained as a bounded, well-formed fragment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawXml {
    /// Serialized XML for one unknown element, including its root tag.
    pub bytes: Vec<u8>,
}

impl RawXml {
    /// Wrap an already parsed unknown XML element.
    #[must_use]
    pub fn new(bytes: Vec<u8>) -> Self {
        Self { bytes }
    }
}

/// A person who authored or is mentioned by a threaded comment.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Person {
    /// Display name from `CT_Person/@displayName`.
    pub display_name: String,
    /// Stable GUID from `CT_Person/@id`.
    pub id: String,
    /// Optional provider-issued identity.
    pub user_id: Option<String>,
    /// Optional identity-provider name.
    pub provider_id: Option<String>,
    /// Unknown attributes retained in preserve mode.
    pub attributes: Vec<RawAttribute>,
    /// Unknown child elements, normally extension lists.
    pub extensions: Vec<RawXml>,
}

impl Person {
    /// Create a person with the required identity fields.
    #[must_use]
    pub fn new(id: impl Into<String>, display_name: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            display_name: display_name.into(),
            ..Self::default()
        }
    }
}

/// The workbook-level persons part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct People {
    /// Persons in serialized document order.
    pub persons: Vec<Person>,
    /// Unknown attributes on `personList`.
    pub attributes: Vec<RawAttribute>,
    /// Unknown children on `personList`.
    pub extensions: Vec<RawXml>,
}

impl People {
    /// Create an empty persons collection.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
}

/// A mention within a threaded comment's UTF-16 text.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Mention {
    /// GUID of the mentioned person.
    pub mention_person_id: String,
    /// Stable GUID identifying this mention.
    pub mention_id: String,
    /// UTF-16 start index in the containing comment text.
    pub start_index: u32,
    /// UTF-16 length of the mention range.
    pub length: u32,
    /// Unknown attributes retained in preserve mode.
    pub attributes: Vec<RawAttribute>,
    /// Unknown child elements retained in preserve mode.
    pub extensions: Vec<RawXml>,
}

impl Mention {
    /// Create a mention with its required identity and text range.
    #[must_use]
    pub fn new(
        mention_person_id: impl Into<String>,
        mention_id: impl Into<String>,
        start_index: u32,
        length: u32,
    ) -> Self {
        Self {
            mention_person_id: mention_person_id.into(),
            mention_id: mention_id.into(),
            start_index,
            length,
            ..Self::default()
        }
    }
}

/// One threaded comment, either a thread root or a reply.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Comment {
    /// Cell reference carried by a root comment.
    pub cell_ref: Option<String>,
    /// Stable GUID identifying this comment.
    pub id: String,
    /// GUID of the parent comment for a reply.
    pub parent_id: Option<String>,
    /// GUID of the authoring person.
    pub person_id: String,
    /// Unformatted comment text.
    pub text: Option<String>,
    /// Optional UTC authoring timestamp.
    pub date_time: Option<String>,
    /// Optional resolved state.
    pub done: Option<bool>,
    /// Mentions in document order.
    pub mentions: Vec<Mention>,
    /// Unknown attributes retained in preserve mode.
    pub attributes: Vec<RawAttribute>,
    /// Unknown child elements retained in preserve mode.
    pub extensions: Vec<RawXml>,
}

impl Comment {
    /// Create a root-capable comment with no text or mentions.
    #[must_use]
    pub fn new(id: impl Into<String>, person_id: impl Into<String>) -> Self {
        Self {
            id: id.into(),
            person_id: person_id.into(),
            ..Self::default()
        }
    }

    /// Whether this value is a top-level comment in its thread.
    #[must_use]
    pub fn is_root(&self) -> bool {
        self.parent_id.is_none()
    }
}

/// One worksheet's threaded-comment part.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Comments {
    /// Comments in serialized order.
    pub comments: Vec<Comment>,
    /// Unknown attributes on `ThreadedComments`.
    pub attributes: Vec<RawAttribute>,
    /// Unknown children on `ThreadedComments`.
    pub extensions: Vec<RawXml>,
}

impl Comments {
    /// Create an empty threaded-comment collection.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Borrow all top-level comments without allocating.
    pub fn roots(&self) -> impl Iterator<Item = &Comment> {
        self.comments.iter().filter(|comment| comment.is_root())
    }
}

/// A semantic thread grouped around one root comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Thread {
    /// The top-level comment anchoring the thread.
    pub root: Comment,
    /// Replies in original worksheet-part order.
    pub replies: Vec<Comment>,
}

/// A workbook persons part plus its relationship identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PeoplePart {
    /// Relationship ID on the workbook part.
    pub relationship_id: String,
    /// OPC part name.
    pub part_name: String,
    /// Typed person data.
    pub persons: People,
}

/// A worksheet threaded-comments part plus its relationship identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentsPart {
    /// Worksheet OPC part name.
    pub worksheet_part_name: String,
    /// Relationship ID on the worksheet part.
    pub relationship_id: String,
    /// OPC part name.
    pub part_name: String,
    /// Typed comment data.
    pub comments: Comments,
}

/// Complete package-neutral threaded-comments graph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Graph {
    /// Optional workbook-level persons part.
    pub persons: Option<PeoplePart>,
    /// Worksheet comment parts in deterministic package order.
    pub worksheets: Vec<CommentsPart>,
}
