//! Package-neutral modern Word comment metadata models.

/// Transitional `WordprocessingML` namespace.
pub const WORD_2012_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
/// Comments IDs namespace.
pub const COMMENTS_IDS_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
/// Comments extensible namespace.
pub const COMMENTS_EXTENSIBLE_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2018/wordml/cex";
/// Word 2018 extension namespace.
pub const WORD_2018_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2018/wordml";
/// Reactions namespace.
pub const REACTIONS_NAMESPACE: &str = "http://schemas.microsoft.com/office/comments/2020/reactions";
/// Office extension-list namespace.
pub const OFFICE_EXTENSION_LIST_NAMESPACE: &str = "http://schemas.microsoft.com/office/2019/extlst";
/// Transitional `WordprocessingML` namespace used by generated parts.
pub const TRANSITIONAL_WORD_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/// Strict `WordprocessingML` namespace used by generated parts.
pub const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

/// Maximum accepted modern-comment part size.
pub const MAX_MODERN_COMMENT_PART_BYTES: usize = 4 * 1024 * 1024;
/// Maximum accepted XML nesting depth.
pub const MAX_MODERN_COMMENT_DEPTH: usize = 128;
/// Maximum number of metadata items or XML nodes.
pub const MAX_MODERN_COMMENT_ITEMS: usize = 65_536;
/// Maximum bytes retained for XML strings.
pub const MAX_MODERN_COMMENT_STRING_BYTES: usize = 8 * 1024 * 1024;

/// `WordprocessingML` conformance used when serializing a metadata part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(crate) const fn word_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_WORD_NAMESPACE,
            Self::Strict => STRICT_WORD_NAMESPACE,
        }
    }
}

/// One commentsExtended entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extended {
    pub paragraph_id: u32,
    pub parent_paragraph_id: Option<u32>,
    pub done: bool,
}

/// A commentsIds paragraph-to-durable identifier mapping.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct IdMapping {
    pub paragraph_id: u32,
    pub durable_id: u32,
}

/// A reaction author.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReactionUser {
    pub user_id: String,
    pub user_name: String,
    pub user_provider: String,
    pub extensions: Option<ExtensionList>,
}

/// Reaction metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReactionInfo {
    pub date_utc: Option<String>,
    pub user: Option<ReactionUser>,
    pub extensions: Option<ExtensionList>,
}

/// One reaction type and its entries.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reaction {
    pub reaction_type: u32,
    pub reactions: Vec<ReactionInfo>,
    pub extensions: Option<ExtensionList>,
}

/// One inert MS-OEXTXML extension with exactly one lax child element.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub(crate) uri: Option<String>,
    pub(crate) child_xml: String,
}

/// Bounded ordered MS-OEXTXML extension list.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ExtensionList {
    pub(crate) extensions: Vec<Extension>,
}

/// One commentsExtensible entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub durable_id: u32,
    pub date_utc: Option<String>,
    pub intelligent_placeholder: Option<bool>,
    pub reactions: Vec<Reaction>,
}

/// Presence details for a person.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Presence {
    pub provider_id: String,
    pub user_id: String,
}

/// A person referenced by a modern comment or revision.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Person {
    pub author: String,
    pub presence: Option<Presence>,
}

/// The four optional modern-comment metadata parts.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Metadata {
    pub comments_extended: Option<Vec<Extended>>,
    pub comments_ids: Option<Vec<IdMapping>>,
    pub comments_extensible: Option<Vec<Comment>>,
    pub people: Option<Vec<Person>>,
}

/// Relationship IDs used to connect the four metadata parts to the main document.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RelationshipIds {
    pub comments_extended: Option<String>,
    pub comments_ids: Option<String>,
    pub comments_extensible: Option<String>,
    pub people: Option<String>,
}
