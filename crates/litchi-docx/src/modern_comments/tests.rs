//! Public API regression tests for the layered modern-comments owner.

use super::{
    Comment as OwnerComment, Conformance as OwnerConformance, Extended as OwnerExtended,
    Extension as OwnerExtension, ExtensionList as OwnerExtensionList, IdMapping as OwnerIdMapping,
    Metadata as OwnerMetadata, Person as OwnerPerson, Presence as OwnerPresence,
    Reaction as OwnerReaction, ReactionInfo as OwnerReactionInfo,
    ReactionUser as OwnerReactionUser, RelationshipIds as OwnerRelationshipIds,
};
use crate as docx;

#[test]
fn canonical_types_are_available_from_the_owner_facade() {
    let _: docx::Comment = OwnerComment {
        durable_id: 1,
        date_utc: None,
        intelligent_placeholder: None,
        reactions: Vec::new(),
    };
    let _: docx::Conformance = OwnerConformance::Strict;
    let _: docx::Extension = OwnerExtension::new(None, r#"<x:value xmlns:x="urn:test"/>"#).unwrap();
    let _: docx::Metadata = OwnerMetadata::default();
    let _: docx::RelationshipIds = OwnerRelationshipIds::default();
    let _: docx::Person = OwnerPerson {
        author: "author".into(),
        presence: None,
    };
    let _: docx::Reaction = OwnerReaction {
        reaction_type: 1,
        reactions: Vec::new(),
        extensions: None,
    };
}

#[test]
fn historical_types_are_aliases_of_canonical_models() {
    let _: docx::ModernCommentConformance = OwnerConformance::Strict;
    let _: docx::CommentExtension = OwnerExtended {
        paragraph_id: 1,
        parent_paragraph_id: None,
        done: false,
    };
    let _: docx::CommentIdMapping = OwnerIdMapping {
        paragraph_id: 1,
        durable_id: 1,
    };
    let _: docx::CommentReaction = OwnerReaction {
        reaction_type: 1,
        reactions: Vec::new(),
        extensions: None,
    };
    let _: docx::CommentReactionInfo = OwnerReactionInfo {
        date_utc: None,
        user: None,
        extensions: None,
    };
    let _: docx::CommentReactionUser = OwnerReactionUser {
        user_id: String::new(),
        user_name: String::new(),
        user_provider: String::new(),
        extensions: None,
    };
    let _: docx::ModernCommentExtension =
        OwnerExtension::new(None, r#"<x:value xmlns:x="urn:test"/>"#).unwrap();
    let _: docx::ModernCommentExtensionList = OwnerExtensionList::default();
    let _: docx::ExtensibleComment = OwnerComment {
        durable_id: 1,
        date_utc: None,
        intelligent_placeholder: None,
        reactions: Vec::new(),
    };
    let _: docx::PresenceInfo = OwnerPresence {
        provider_id: String::new(),
        user_id: String::new(),
    };
    let _: docx::ModernCommentMetadata = OwnerMetadata::default();
    let _: docx::ModernCommentRelationshipIds = OwnerRelationshipIds::default();
}
