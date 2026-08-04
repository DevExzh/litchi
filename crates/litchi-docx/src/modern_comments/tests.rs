//! Public API regression tests for the layered modern-comments owner.

use super::{
    Comment as OwnerComment, Conformance as OwnerConformance, Extension as OwnerExtension,
    Metadata as OwnerMetadata, Person as OwnerPerson, Reaction as OwnerReaction,
    RelationshipIds as OwnerRelationshipIds,
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
