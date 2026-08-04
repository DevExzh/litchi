//! OPC/package graph lifecycle for modern Word comment metadata.

use super::codec::{
    parse_comments_extended, parse_comments_extensible, parse_comments_ids, parse_people,
    validate_metadata, write_comments_extended, write_comments_extensible, write_comments_ids,
    write_people,
};
use super::model::{Conformance, Metadata, RelationshipIds};
use crate::{Error, Result};
use litchi_opc::part::{Part, XmlPart};
use litchi_opc::{OpcPackage, PackURI};
use std::collections::HashSet;

/// Relationship type for the commentsExtended part.
pub const COMMENTS_EXTENDED_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
/// Relationship type for the commentsIds part.
pub const COMMENTS_IDS_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
/// Relationship type for the commentsExtensible part.
pub const COMMENTS_EXTENSIBLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
/// Relationship type for the people part.
pub const PEOPLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2011/relationships/people";

/// Content type for the commentsExtended part.
pub const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
/// Content type for the commentsIds part.
pub const COMMENTS_IDS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
/// Content type for the commentsExtensible part.
pub const COMMENTS_EXTENSIBLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";
/// Content type for the people part.
pub const PEOPLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";

pub fn load_modern_comment_metadata(
    package: &OpcPackage,
    document_part_name: &PackURI,
) -> Result<Metadata> {
    reject_misplaced_relationships(package, document_part_name)?;
    let document = package.get_part(document_part_name).map_err(|error| {
        Error::Invalid(format!(
            "Word main document '{}': {error}",
            document_part_name.as_str()
        ))
    })?;
    require_main_document_content_type(document.content_type())?;
    let metadata = Metadata {
        comments_extended: load_part(
            package,
            document,
            COMMENTS_EXTENDED_RELATIONSHIP,
            COMMENTS_EXTENDED_CONTENT_TYPE,
            parse_comments_extended,
        )?,
        comments_ids: load_part(
            package,
            document,
            COMMENTS_IDS_RELATIONSHIP,
            COMMENTS_IDS_CONTENT_TYPE,
            parse_comments_ids,
        )?,
        comments_extensible: load_part(
            package,
            document,
            COMMENTS_EXTENSIBLE_RELATIONSHIP,
            COMMENTS_EXTENSIBLE_CONTENT_TYPE,
            parse_comments_extensible,
        )?,
        people: load_part(
            package,
            document,
            PEOPLE_RELATIONSHIP,
            PEOPLE_CONTENT_TYPE,
            parse_people,
        )?,
    };
    validate_metadata(&metadata)?;
    Ok(metadata)
}

pub fn store_modern_comment_metadata(
    package: &mut OpcPackage,
    document_part_name: &PackURI,
    metadata: &Metadata,
    relationship_ids: &RelationshipIds,
    conformance: Conformance,
) -> Result<()> {
    validate_metadata(metadata)?;
    let document = package.get_part(document_part_name).map_err(|error| {
        Error::Invalid(format!(
            "Word main document '{}': {error}",
            document_part_name.as_str()
        ))
    })?;
    require_main_document_content_type(document.content_type())?;
    let specs = [
        (
            metadata
                .comments_extended
                .as_ref()
                .map(|items| write_comments_extended(items, conformance)),
            relationship_ids.comments_extended.as_deref(),
            "/word/commentsExtended.xml",
            COMMENTS_EXTENDED_RELATIONSHIP,
            COMMENTS_EXTENDED_CONTENT_TYPE,
        ),
        (
            metadata
                .comments_ids
                .as_ref()
                .map(|items| write_comments_ids(items, conformance)),
            relationship_ids.comments_ids.as_deref(),
            "/word/commentsIds.xml",
            COMMENTS_IDS_RELATIONSHIP,
            COMMENTS_IDS_CONTENT_TYPE,
        ),
        (
            metadata
                .comments_extensible
                .as_ref()
                .map(|items| write_comments_extensible(items, conformance)),
            relationship_ids.comments_extensible.as_deref(),
            "/word/commentsExtensible.xml",
            COMMENTS_EXTENSIBLE_RELATIONSHIP,
            COMMENTS_EXTENSIBLE_CONTENT_TYPE,
        ),
        (
            metadata
                .people
                .as_ref()
                .map(|items| write_people(items, conformance)),
            relationship_ids.people.as_deref(),
            "/word/people.xml",
            PEOPLE_RELATIONSHIP,
            PEOPLE_CONTENT_TYPE,
        ),
    ];
    let mut pending = Vec::new();
    let mut ids = HashSet::new();
    for (xml, relationship_id, part_name, relationship_type, content_type) in specs {
        match (xml, relationship_id) {
            (None, None) => continue,
            (Some(_), None) => return invalid(format!("missing relationship ID for {part_name}")),
            (None, Some(_)) => {
                return invalid(format!("relationship ID supplied without {part_name}"));
            },
            (Some(xml), Some(relationship_id)) => {
                if relationship_id.is_empty() || !ids.insert(relationship_id) {
                    return invalid(
                        "modern comment relationship IDs must be nonempty and unique".into(),
                    );
                }
                let part_name = PackURI::new(part_name).map_err(Error::Uri)?;
                if package
                    .iter_parts()
                    .any(|part| part.partname() == &part_name)
                    || document.rels().iter().any(|relationship| {
                        relationship.r_id() == relationship_id
                            || relationship.reltype() == relationship_type
                    })
                {
                    return invalid(format!(
                        "modern comment part or relationship already exists: {}",
                        part_name.as_str()
                    ));
                }
                pending.push((
                    part_name,
                    xml?,
                    relationship_id.to_owned(),
                    relationship_type,
                    content_type,
                ));
            },
        }
    }
    for (part_name, xml, relationship_id, relationship_type, content_type) in pending {
        let target = part_name.relative_ref(document_part_name.base_uri());
        package.add_part(Box::new(XmlPart::new(part_name, content_type.into(), xml)));
        package
            .get_part_mut(document_part_name)?
            .rels_mut()
            .add_relationship(relationship_type.into(), target, relationship_id, false);
    }
    Ok(())
}

fn load_part<T>(
    package: &OpcPackage,
    document: &dyn Part,
    relationship_type: &str,
    content_type: &str,
    parser: fn(&[u8]) -> Result<Vec<T>>,
) -> Result<Option<Vec<T>>> {
    let relationships: Vec<_> = document
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == relationship_type)
        .collect();
    if relationships.len() > 1 {
        return invalid(format!("multiple '{relationship_type}' relationships"));
    }
    let Some(relationship) = relationships.first() else {
        return Ok(None);
    };
    if relationship.is_external() {
        return invalid(format!(
            "modern comment relationship '{}' must be internal",
            relationship.r_id()
        ));
    }
    let target = relationship
        .target_partname()
        .map_err(|error| Error::Invalid(format!("invalid modern comment target: {error}")))?;
    let part = package.get_part(&target).map_err(|error| {
        Error::Invalid(format!(
            "modern comment part '{}': {error}",
            target.as_str()
        ))
    })?;
    if part.content_type() != content_type {
        return Err(Error::ContentType {
            expected: content_type.into(),
            actual: part.content_type().into(),
        });
    }
    if part.rels().iter().next().is_some() {
        return invalid(format!(
            "modern comment part '{}' must not have relationships",
            target.as_str()
        ));
    }
    parser(part.blob()).map(Some).map_err(|error| {
        Error::Invalid(format!(
            "invalid modern comment part '{}': {error}",
            target.as_str()
        ))
    })
}

fn reject_misplaced_relationships(package: &OpcPackage, document_name: &PackURI) -> Result<()> {
    if package
        .rels()
        .iter()
        .any(|relationship| is_modern_relationship(relationship.reltype()))
    {
        return invalid("package root cannot source modern Word comment relationships".into());
    }
    for part in package.iter_parts() {
        if part.partname() != document_name
            && part
                .rels()
                .iter()
                .any(|relationship| is_modern_relationship(relationship.reltype()))
        {
            return invalid(format!(
                "modern comment relationship has invalid source '{}'",
                part.partname().as_str()
            ));
        }
    }
    Ok(())
}

fn is_modern_relationship(value: &str) -> bool {
    matches!(
        value,
        COMMENTS_EXTENDED_RELATIONSHIP
            | COMMENTS_IDS_RELATIONSHIP
            | COMMENTS_EXTENSIBLE_RELATIONSHIP
            | PEOPLE_RELATIONSHIP
    )
}

fn require_main_document_content_type(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            | "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
            | "application/vnd.ms-word.document.macroEnabled.main+xml"
            | "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
    ) {
        Ok(())
    } else {
        invalid(format!(
            "'{content_type}' is not a Word main-document content type"
        ))
    }
}

fn invalid<T>(message: String) -> Result<T> {
    Err(Error::Invalid(message))
}

#[cfg(test)]
mod tests {
    use super::super::model::*;
    use super::*;
    use litchi_opc::part::BlobPart;

    const POI_DOCX: &[u8] =
        include_bytes!("../../../../test-data/poi/test-data/document/testComment.docx");
    const LO_DOCX: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/redline-range-comment.docx"
    );

    #[test]
    fn loads_poi_and_libreoffice_reference_packages() {
        for bytes in [POI_DOCX, LO_DOCX] {
            let package = OpcPackage::from_bytes(bytes).unwrap();
            let metadata = load_modern_comment_metadata(
                &package,
                &PackURI::new("/word/document.xml").unwrap(),
            )
            .unwrap();
            assert_eq!(metadata.comments_extended.as_ref().unwrap().len(), 1);
            assert_eq!(metadata.comments_ids.as_ref().unwrap().len(), 1);
            assert_eq!(metadata.comments_extensible.as_ref().unwrap().len(), 1);
        }
        let package = OpcPackage::from_bytes(LO_DOCX).unwrap();
        let metadata =
            load_modern_comment_metadata(&package, &PackURI::new("/word/document.xml").unwrap())
                .unwrap();
        assert_eq!(metadata.people.unwrap()[0].author, "Miklos Vajna");
    }

    #[test]
    fn all_writers_are_deterministic_and_round_trip_strict() {
        let metadata = sample_metadata();
        let extended = write_comments_extended(
            metadata.comments_extended.as_ref().unwrap(),
            ModernCommentConformance::Strict,
        )
        .unwrap();
        assert_eq!(
            extended,
            write_comments_extended(
                metadata.comments_extended.as_ref().unwrap(),
                ModernCommentConformance::Strict
            )
            .unwrap()
        );
        assert!(
            std::str::from_utf8(&extended)
                .unwrap()
                .contains(STRICT_WORD_NAMESPACE)
        );
        assert_eq!(
            parse_comments_extended(&extended).unwrap(),
            metadata.comments_extended.unwrap()
        );

        let ids = write_comments_ids(
            metadata.comments_ids.as_ref().unwrap(),
            ModernCommentConformance::Strict,
        )
        .unwrap();
        assert_eq!(
            parse_comments_ids(&ids).unwrap(),
            metadata.comments_ids.unwrap()
        );
        let extensible = write_comments_extensible(
            metadata.comments_extensible.as_ref().unwrap(),
            ModernCommentConformance::Strict,
        )
        .unwrap();
        assert_eq!(
            parse_comments_extensible(&extensible).unwrap(),
            metadata.comments_extensible.unwrap()
        );
        let people = write_people(
            metadata.people.as_ref().unwrap(),
            ModernCommentConformance::Strict,
        )
        .unwrap();
        assert_eq!(parse_people(&people).unwrap(), metadata.people.unwrap());
    }

    #[test]
    fn mce_selects_fallback_comment() {
        let xml = format!(
            r#"<w15:commentsEx xmlns:w15="{WORD_2012_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w15:commentEx w15:paraId="AAAAAAAA"/></mc:Choice><mc:Fallback><w15:commentEx w15:paraId="BBBBBBBB"/></mc:Fallback></mc:AlternateContent></w15:commentsEx>"#
        );
        assert_eq!(
            parse_comments_extended(xml.as_bytes()).unwrap()[0].paragraph_id,
            0xBBBB_BBBB
        );
    }

    #[test]
    fn package_writer_round_trips_all_parts() {
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
            b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>".to_vec(),
        )));
        let metadata = sample_metadata();
        store_modern_comment_metadata(
            &mut package,
            &PackURI::new("/word/document.xml").unwrap(),
            &metadata,
            &ModernCommentRelationshipIds {
                comments_extended: Some("rIdEx".into()),
                comments_ids: Some("rIdIds".into()),
                comments_extensible: Some("rIdCex".into()),
                people: Some("rIdPeople".into()),
            },
            ModernCommentConformance::Transitional,
        )
        .unwrap();
        assert_eq!(
            load_modern_comment_metadata(&package, &PackURI::new("/word/document.xml").unwrap())
                .unwrap(),
            metadata
        );
    }

    #[test]
    fn rejects_malformed_xml_graph_and_cross_part_ids() {
        assert!(parse_comments_extended(br#"<!DOCTYPE x><x/>"#).is_err());
        let bad_hex = format!(
            r#"<w15:commentsEx xmlns:w15="{WORD_2012_NAMESPACE}"><w15:commentEx w15:paraId="123"/></w15:commentsEx>"#
        );
        assert!(parse_comments_extended(bad_hex.as_bytes()).is_err());
        let bad_date = format!(
            r#"<w16cex:commentsExtensible xmlns:w16cex="{COMMENTS_EXTENSIBLE_NAMESPACE}"><w16cex:commentExtensible w16cex:durableId="00000001" w16cex:dateUtc="2026-01-01T01:00:00+01:00"/></w16cex:commentsExtensible>"#
        );
        assert!(parse_comments_extensible(bad_date.as_bytes()).is_err());

        let mut metadata = sample_metadata();
        metadata.comments_extensible.as_mut().unwrap()[0].durable_id = 2;
        assert!(validate_metadata(&metadata).is_err());

        let mut package = OpcPackage::new();
        let mut document = BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .into(),
            b"<document/>".to_vec(),
        );
        document.rels_mut().add_relationship(
            COMMENTS_EXTENDED_RELATIONSHIP.into(),
            "https://example.invalid/comments".into(),
            "rId1".into(),
            true,
        );
        package.add_part(Box::new(document));
        assert!(
            load_modern_comment_metadata(&package, &PackURI::new("/word/document.xml").unwrap())
                .is_err()
        );
    }

    #[test]
    fn enforces_size_depth_count_and_reaction_constraints() {
        assert!(parse_people(&vec![b' '; MAX_MODERN_COMMENT_PART_BYTES + 1]).is_err());
        let deep = format!(
            "{}{}",
            "<x>".repeat(MAX_MODERN_COMMENT_DEPTH + 1),
            "</x>".repeat(MAX_MODERN_COMMENT_DEPTH + 1)
        );
        assert!(parse_people(deep.as_bytes()).is_err());
        let mut metadata = sample_metadata();
        metadata.comments_extensible.as_mut().unwrap()[0].reactions[0].reaction_type = 0;
        assert!(validate_metadata(&metadata).is_err());
    }

    fn sample_metadata() -> ModernCommentMetadata {
        ModernCommentMetadata {
            comments_extended: Some(vec![CommentExtension {
                paragraph_id: 0x1234_ABCD,
                parent_paragraph_id: None,
                done: true,
            }]),
            comments_ids: Some(vec![CommentIdMapping {
                paragraph_id: 0x1234_ABCD,
                durable_id: 0x0123_4567,
            }]),
            comments_extensible: Some(vec![ExtensibleComment {
                durable_id: 0x0123_4567,
                date_utc: Some("2026-07-17T00:00:00Z".into()),
                intelligent_placeholder: Some(false),
                reactions: vec![CommentReaction {
                    reaction_type: 1,
                    reactions: vec![CommentReactionInfo {
                        date_utc: Some("2026-07-17T01:00:00Z".into()),
                        user: Some(CommentReactionUser {
                            user_id: "alice@example.test".into(),
                            user_name: "Alice & Bob".into(),
                            user_provider: "O365".into(),
                            extensions: None,
                        }),
                        extensions: None,
                    }],
                    extensions: None,
                }],
            }]),
            people: Some(vec![Person {
                author: "Alice & Bob".into(),
                presence: Some(PresenceInfo {
                    provider_id: "O365".into(),
                    user_id: "alice@example.test".into(),
                }),
            }]),
        }
    }
}

#[cfg(test)]
mod reaction_extension_list_tests {
    use super::super::codec::REACTIONS_EXTENSION_URI;
    use super::super::model::*;
    use super::*;

    fn comments_with_reaction(reaction_body: &str) -> String {
        format!(
            r#"<w16cex:commentsExtensible xmlns:w16cex="{COMMENTS_EXTENSIBLE_NAMESPACE}" xmlns:w16="{WORD_2018_NAMESPACE}" xmlns:cr="{REACTIONS_NAMESPACE}" xmlns:oel="{OFFICE_EXTENSION_LIST_NAMESPACE}" xmlns:x="urn:example:unknown"><w16cex:commentExtensible w16cex:durableId="00000001"><w16cex:extLst><w16:ext w16:uri="{REACTIONS_EXTENSION_URI}"><cr:reactions><cr:reaction reactionType="1">{reaction_body}</cr:reaction></cr:reactions></w16:ext></w16cex:extLst></w16cex:commentExtensible></w16cex:commentsExtensible>"#
        )
    }

    #[test]
    fn official_extension_shape_round_trips_at_all_three_levels_and_mutates() {
        let xml = comments_with_reaction(
            r#"<cr:reactionInfo dateUtc="2026-07-17T01:00:00Z"><cr:user userId="alice" userName="Alice" userProvider="O365"><oel:extLst><oel:ext uri="  urn:user   metadata  "><x:userData x:flag="1">opaque &amp; <x:value>mixed</x:value></x:userData></oel:ext></oel:extLst></cr:user><oel:extLst><oel:ext><x:infoData><x:value>42</x:value></x:infoData></oel:ext></oel:extLst></cr:reactionInfo><oel:extLst><oel:ext uri="urn:reaction"><x:reactionData/></oel:ext></oel:extLst>"#,
        );
        let mut comments = parse_comments_extensible(xml.as_bytes()).unwrap();
        let reaction = &comments[0].reactions[0];
        assert_eq!(
            reaction.extensions.as_ref().unwrap().extensions()[0].uri(),
            Some("urn:reaction")
        );
        let info = &reaction.reactions[0];
        assert_eq!(
            info.extensions.as_ref().unwrap().extensions()[0].uri(),
            None
        );
        let user_extension = &info
            .user
            .as_ref()
            .unwrap()
            .extensions
            .as_ref()
            .unwrap()
            .extensions()[0];
        assert_eq!(user_extension.uri(), Some("urn:user metadata"));
        assert!(user_extension.child_xml().contains("opaque &amp;"));
        assert!(
            user_extension
                .child_xml()
                .contains("xmlns:x=\"urn:example:unknown\"")
        );

        let first =
            write_comments_extensible(&comments, ModernCommentConformance::Transitional).unwrap();
        let reparsed = parse_comments_extensible(&first).unwrap();
        let second =
            write_comments_extensible(&reparsed, ModernCommentConformance::Transitional).unwrap();
        assert_eq!(first, second);
        assert_eq!(comments, reparsed);

        let user_list = comments[0].reactions[0].reactions[0]
            .user
            .as_mut()
            .unwrap()
            .extensions
            .as_mut()
            .unwrap();
        user_list.extensions[0].set_uri(Some("  urn:changed\tvalue ".into()));
        user_list.extensions[0]
            .set_child_xml(
                r#"<z:data xmlns:z="urn:mutated">new <z:value>content</z:value></z:data>"#,
            )
            .unwrap();
        user_list
            .push(ModernCommentExtension::new(None, r#"<q:extra xmlns:q="urn:q"/>"#).unwrap())
            .unwrap();
        let removed = user_list.remove(1).unwrap();
        assert_eq!(removed.uri(), None);
        let mutated =
            write_comments_extensible(&comments, ModernCommentConformance::Strict).unwrap();
        let reparsed = parse_comments_extensible(&mutated).unwrap();
        let extension = &reparsed[0].reactions[0].reactions[0]
            .user
            .as_ref()
            .unwrap()
            .extensions
            .as_ref()
            .unwrap()
            .extensions()[0];
        assert_eq!(extension.uri(), Some("urn:changed value"));
        assert!(extension.child_xml().contains("urn:mutated"));
    }

    #[test]
    fn rejects_namespace_sequence_cardinality_duplicate_and_resource_violations() {
        for body in [
            r#"<oel:extLst/><cr:reactionInfo/>"#,
            r#"<oel:extLst/><oel:extLst/>"#,
            r#"<cr:extLst/>"#,
            r#"<oel:extLst><oel:ext/></oel:extLst>"#,
            r#"<oel:extLst><oel:ext>text<x:a/></oel:ext></oel:extLst>"#,
            r#"<oel:extLst><oel:ext><x:a/><x:b/></oel:ext></oel:extLst>"#,
            r#"<oel:extLst><oel:ext oel:uri="urn:qualified"><x:a/></oel:ext></oel:extLst>"#,
            r#"<cr:reactionInfo><oel:extLst/><cr:user userId="a" userName="A" userProvider="P"/></cr:reactionInfo>"#,
            r#"<cr:reactionInfo><cr:user userId="a" userId="b" userName="A" userProvider="P"/></cr:reactionInfo>"#,
        ] {
            assert!(
                parse_comments_extensible(comments_with_reaction(body).as_bytes()).is_err(),
                "accepted {body}"
            );
        }

        let extension = ModernCommentExtension::new(None, r#"<x:a xmlns:x="urn:x"/>"#).unwrap();
        assert!(
            ModernCommentExtensionList::new(vec![extension; MAX_MODERN_COMMENT_ITEMS + 1]).is_err()
        );
        let oversized = format!(
            "<x:a xmlns:x=\"urn:x\">{}</x:a>",
            "x".repeat(MAX_MODERN_COMMENT_PART_BYTES)
        );
        assert!(ModernCommentExtension::new(None, oversized).is_err());
        let deep = format!(
            "<x:a xmlns:x=\"urn:x\">{}{}</x:a>",
            "<x:b>".repeat(MAX_MODERN_COMMENT_DEPTH),
            "</x:b>".repeat(MAX_MODERN_COMMENT_DEPTH)
        );
        assert!(ModernCommentExtension::new(None, deep).is_err());
        assert!(
            ModernCommentExtension::new(None, "<?xml version=\"1.0\"?><x:a xmlns:x=\"urn:x\"/>")
                .is_err()
        );
    }
}
