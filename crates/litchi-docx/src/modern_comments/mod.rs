//! Layered modern Word comment metadata owner.

mod codec;
mod model;
mod package;

#[cfg(test)]
mod tests;

pub use codec::{
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    write_comments_extended, write_comments_extensible, write_comments_ids, write_people,
};
pub use model::{
    COMMENTS_EXTENSIBLE_NAMESPACE, COMMENTS_IDS_NAMESPACE, Comment, Conformance, Extended,
    Extension, ExtensionList, IdMapping, MAX_MODERN_COMMENT_DEPTH, MAX_MODERN_COMMENT_ITEMS,
    MAX_MODERN_COMMENT_PART_BYTES, MAX_MODERN_COMMENT_STRING_BYTES, Metadata,
    OFFICE_EXTENSION_LIST_NAMESPACE, Person, Presence, REACTIONS_NAMESPACE, Reaction, ReactionInfo,
    ReactionUser, RelationshipIds, STRICT_WORD_NAMESPACE, TRANSITIONAL_WORD_NAMESPACE,
    WORD_2012_NAMESPACE, WORD_2018_NAMESPACE,
};
pub use package::{
    COMMENTS_EXTENDED_CONTENT_TYPE, COMMENTS_EXTENDED_RELATIONSHIP,
    COMMENTS_EXTENSIBLE_CONTENT_TYPE, COMMENTS_EXTENSIBLE_RELATIONSHIP, COMMENTS_IDS_CONTENT_TYPE,
    COMMENTS_IDS_RELATIONSHIP, PEOPLE_CONTENT_TYPE, PEOPLE_RELATIONSHIP,
    load_modern_comment_metadata, store_modern_comment_metadata,
};
