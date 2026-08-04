//! Compatibility adapter for the canonical DOCX modern-comment codec.
//!
//! The typed metadata model, bounded XML/MCE processing, and package graph
//! operations live in `litchi_docx::modern_comments`. This module preserves
//! the historical host path and maps canonical failures back to `OoxmlError`.

use crate::error::{OoxmlError, Result};
use litchi_opc::{OpcPackage, PackURI};

pub use litchi_docx::modern_comments::{
    COMMENTS_EXTENDED_CONTENT_TYPE, COMMENTS_EXTENDED_RELATIONSHIP,
    COMMENTS_EXTENSIBLE_CONTENT_TYPE, COMMENTS_EXTENSIBLE_NAMESPACE,
    COMMENTS_EXTENSIBLE_RELATIONSHIP, COMMENTS_IDS_CONTENT_TYPE, COMMENTS_IDS_NAMESPACE,
    COMMENTS_IDS_RELATIONSHIP, CommentExtension, CommentIdMapping, CommentReaction,
    CommentReactionInfo, CommentReactionUser, ExtensibleComment, MAX_MODERN_COMMENT_DEPTH,
    MAX_MODERN_COMMENT_ITEMS, MAX_MODERN_COMMENT_PART_BYTES, MAX_MODERN_COMMENT_STRING_BYTES,
    ModernCommentConformance, ModernCommentExtension, ModernCommentExtensionList,
    ModernCommentMetadata, ModernCommentRelationshipIds, OFFICE_EXTENSION_LIST_NAMESPACE,
    PEOPLE_CONTENT_TYPE, PEOPLE_RELATIONSHIP, Person, PresenceInfo, REACTIONS_NAMESPACE,
    STRICT_WORD_NAMESPACE, TRANSITIONAL_WORD_NAMESPACE, WORD_2012_NAMESPACE, WORD_2018_NAMESPACE,
};

fn map_docx_error(error: litchi_docx::Error) -> OoxmlError {
    match error {
        litchi_docx::Error::Opc(error) => OoxmlError::Opc(error),
        litchi_docx::Error::Xml(message) => OoxmlError::Xml(message),
        litchi_docx::Error::ContentType { expected, actual } => OoxmlError::InvalidContentType {
            expected,
            got: actual,
        },
        litchi_docx::Error::Invalid(message) if is_missing_part(&message) => {
            OoxmlError::PartNotFound(message)
        },
        litchi_docx::Error::Invalid(message)
            if message.starts_with("invalid modern comment target:") =>
        {
            OoxmlError::InvalidRelationship(message)
        },
        litchi_docx::Error::Invalid(message) => OoxmlError::InvalidFormat(message),
        litchi_docx::Error::Uri(message) => OoxmlError::InvalidUri(message),
        litchi_docx::Error::Mce(error) => {
            OoxmlError::Common(litchi_ooxml_common::Error::Mce(error))
        },
        litchi_docx::Error::Allocation { resource, source } => {
            OoxmlError::Allocation { resource, source }
        },
        other => OoxmlError::Docx(other),
    }
}

fn is_missing_part(message: &str) -> bool {
    (message.starts_with("Word main document '") || message.starts_with("modern comment part '"))
        && message.contains(": Part not found:")
}

pub fn parse_comments_extended(xml: &[u8]) -> Result<Vec<CommentExtension>> {
    litchi_docx::modern_comments::parse_comments_extended(xml).map_err(map_docx_error)
}

pub fn parse_comments_ids(xml: &[u8]) -> Result<Vec<CommentIdMapping>> {
    litchi_docx::modern_comments::parse_comments_ids(xml).map_err(map_docx_error)
}

pub fn parse_comments_extensible(xml: &[u8]) -> Result<Vec<ExtensibleComment>> {
    litchi_docx::modern_comments::parse_comments_extensible(xml).map_err(map_docx_error)
}

pub fn parse_people(xml: &[u8]) -> Result<Vec<Person>> {
    litchi_docx::modern_comments::parse_people(xml).map_err(map_docx_error)
}

pub fn write_comments_extended(
    items: &[CommentExtension],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    litchi_docx::modern_comments::write_comments_extended(items, conformance)
        .map_err(map_docx_error)
}

pub fn write_comments_ids(
    items: &[CommentIdMapping],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    litchi_docx::modern_comments::write_comments_ids(items, conformance).map_err(map_docx_error)
}

pub fn write_comments_extensible(
    comments: &[ExtensibleComment],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    litchi_docx::modern_comments::write_comments_extensible(comments, conformance)
        .map_err(map_docx_error)
}

pub fn write_people(people: &[Person], conformance: ModernCommentConformance) -> Result<Vec<u8>> {
    litchi_docx::modern_comments::write_people(people, conformance).map_err(map_docx_error)
}

pub fn load_modern_comment_metadata(
    package: &OpcPackage,
    document_part_name: &PackURI,
) -> Result<ModernCommentMetadata> {
    litchi_docx::modern_comments::load_modern_comment_metadata(package, document_part_name)
        .map_err(map_docx_error)
}

pub fn store_modern_comment_metadata(
    package: &mut OpcPackage,
    document_part_name: &PackURI,
    metadata: &ModernCommentMetadata,
    relationship_ids: &ModernCommentRelationshipIds,
    conformance: ModernCommentConformance,
) -> Result<()> {
    litchi_docx::modern_comments::store_modern_comment_metadata(
        package,
        document_part_name,
        metadata,
        relationship_ids,
        conformance,
    )
    .map_err(map_docx_error)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn keeps_historical_invalid_format_mapping() {
        let error = parse_comments_extended(b"<!DOCTYPE x><x/>").expect_err("invalid XML");
        assert!(matches!(error, OoxmlError::InvalidFormat(_)));
    }
}
