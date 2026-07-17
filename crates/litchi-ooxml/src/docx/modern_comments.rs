//! Typed metadata for modern Word comment threads.
//!
//! The four metadata parts supplement, but do not replace, ISO/IEC 29500
//! `comments.xml`. No presence service, identity provider, or external content
//! is contacted by this module.

use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use chrono::DateTime;
use litchi_opc::part::XmlPart;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::{HashMap, HashSet};

pub const WORD_2012_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
pub const COMMENTS_IDS_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2016/wordml/cid";
pub const COMMENTS_EXTENSIBLE_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/word/2018/wordml/cex";
pub const WORD_2018_NAMESPACE: &str = "http://schemas.microsoft.com/office/word/2018/wordml";
pub const REACTIONS_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/comments/2020/reactions";
pub const TRANSITIONAL_WORD_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

pub const COMMENTS_EXTENDED_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
pub const COMMENTS_IDS_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
pub const COMMENTS_EXTENSIBLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
pub const PEOPLE_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2011/relationships/people";

pub const COMMENTS_EXTENDED_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
pub const COMMENTS_IDS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
pub const COMMENTS_EXTENSIBLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";
pub const PEOPLE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";

const REACTIONS_EXTENSION_URI: &str = "{CE6994B0-6A32-4C9F-8C6B-6E91EDA988CE}";
pub const MAX_MODERN_COMMENT_PART_BYTES: usize = 4 * 1024 * 1024;
pub const MAX_MODERN_COMMENT_DEPTH: usize = 128;
pub const MAX_MODERN_COMMENT_ITEMS: usize = 65_536;
pub const MAX_MODERN_COMMENT_STRING_BYTES: usize = 8 * 1024 * 1024;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ModernCommentConformance {
    Transitional,
    Strict,
}

impl ModernCommentConformance {
    fn word_namespace(self) -> &'static str {
        match self {
            Self::Transitional => TRANSITIONAL_WORD_NAMESPACE,
            Self::Strict => STRICT_WORD_NAMESPACE,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentExtension {
    pub paragraph_id: u32,
    pub parent_paragraph_id: Option<u32>,
    pub done: bool,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentIdMapping {
    pub paragraph_id: u32,
    pub durable_id: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentReactionUser {
    pub user_id: String,
    pub user_name: String,
    pub user_provider: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentReactionInfo {
    pub date_utc: Option<String>,
    pub user: Option<CommentReactionUser>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentReaction {
    pub reaction_type: u32,
    pub reactions: Vec<CommentReactionInfo>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensibleComment {
    pub durable_id: u32,
    pub date_utc: Option<String>,
    pub intelligent_placeholder: Option<bool>,
    pub reactions: Vec<CommentReaction>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresenceInfo {
    pub provider_id: String,
    pub user_id: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Person {
    pub author: String,
    pub presence: Option<PresenceInfo>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ModernCommentMetadata {
    pub comments_extended: Option<Vec<CommentExtension>>,
    pub comments_ids: Option<Vec<CommentIdMapping>>,
    pub comments_extensible: Option<Vec<ExtensibleComment>>,
    pub people: Option<Vec<Person>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ModernCommentRelationshipIds {
    pub comments_extended: Option<String>,
    pub comments_ids: Option<String>,
    pub comments_extensible: Option<String>,
    pub people: Option<String>,
}

pub fn parse_comments_extended(xml: &[u8]) -> Result<Vec<CommentExtension>> {
    let document = parse_document(xml)?;
    let root = document.root()?;
    require_name(root, WORD_2012_NAMESPACE, "commentsEx")?;
    reject_attributes(root, &[])?;
    enforce_count("commentEx", root.children.len())?;
    let mut items = Vec::with_capacity(root.children.len());
    let mut seen = HashSet::new();
    for child in &root.children {
        require_name(child, WORD_2012_NAMESPACE, "commentEx")?;
        reject_attributes(
            child,
            &[
                (WORD_2012_NAMESPACE, "paraId"),
                (WORD_2012_NAMESPACE, "paraIdParent"),
                (WORD_2012_NAMESPACE, "done"),
            ],
        )?;
        require_empty(child)?;
        let paragraph_id = required_hex(child, WORD_2012_NAMESPACE, "paraId")?;
        if !seen.insert(paragraph_id) {
            return invalid(format!("duplicate commentEx paraId {}", format_hex(paragraph_id)));
        }
        items.push(CommentExtension {
            paragraph_id,
            parent_paragraph_id: optional_hex(child, WORD_2012_NAMESPACE, "paraIdParent")?,
            done: optional_on_off(child, WORD_2012_NAMESPACE, "done")?.unwrap_or(false),
        });
    }
    Ok(items)
}

pub fn parse_comments_ids(xml: &[u8]) -> Result<Vec<CommentIdMapping>> {
    let document = parse_document(xml)?;
    let root = document.root()?;
    require_name(root, COMMENTS_IDS_NAMESPACE, "commentsIds")?;
    reject_attributes(root, &[])?;
    enforce_count("commentId", root.children.len())?;
    let mut items = Vec::with_capacity(root.children.len());
    let mut paragraph_ids = HashSet::new();
    let mut durable_ids = HashSet::new();
    for child in &root.children {
        require_name(child, COMMENTS_IDS_NAMESPACE, "commentId")?;
        reject_attributes(
            child,
            &[
                (COMMENTS_IDS_NAMESPACE, "paraId"),
                (COMMENTS_IDS_NAMESPACE, "durableId"),
            ],
        )?;
        require_empty(child)?;
        let paragraph_id = required_hex(child, COMMENTS_IDS_NAMESPACE, "paraId")?;
        let durable_id = required_hex(child, COMMENTS_IDS_NAMESPACE, "durableId")?;
        validate_durable_id(durable_id)?;
        if !paragraph_ids.insert(paragraph_id) || !durable_ids.insert(durable_id) {
            return invalid("commentsIds contains duplicate paragraph or durable ID".into());
        }
        items.push(CommentIdMapping {
            paragraph_id,
            durable_id,
        });
    }
    Ok(items)
}

pub fn parse_comments_extensible(xml: &[u8]) -> Result<Vec<ExtensibleComment>> {
    let document = parse_document(xml)?;
    let root = document.root()?;
    require_name(root, COMMENTS_EXTENSIBLE_NAMESPACE, "commentsExtensible")?;
    reject_attributes(root, &[])?;
    let mut comments = Vec::new();
    let mut seen = HashSet::new();
    let mut saw_root_extensions = false;
    for child in &root.children {
        if child.namespace == COMMENTS_EXTENSIBLE_NAMESPACE && child.local_name == "extLst" {
            if saw_root_extensions || child != root.children.last().expect("child exists") {
                return invalid("commentsExtensible extLst must occur once at the end".into());
            }
            saw_root_extensions = true;
            continue;
        }
        if saw_root_extensions {
            return invalid("commentExtensible occurs after root extLst".into());
        }
        require_name(child, COMMENTS_EXTENSIBLE_NAMESPACE, "commentExtensible")?;
        reject_attributes(
            child,
            &[
                (COMMENTS_EXTENSIBLE_NAMESPACE, "durableId"),
                (COMMENTS_EXTENSIBLE_NAMESPACE, "dateUtc"),
                (COMMENTS_EXTENSIBLE_NAMESPACE, "intelligentPlaceholder"),
            ],
        )?;
        let durable_id = required_hex(child, COMMENTS_EXTENSIBLE_NAMESPACE, "durableId")?;
        validate_durable_id(durable_id)?;
        if !seen.insert(durable_id) {
            return invalid(format!("duplicate extensible durableId {}", format_hex(durable_id)));
        }
        let date_utc = attr(child, COMMENTS_EXTENSIBLE_NAMESPACE, "dateUtc").map(str::to_owned);
        if let Some(date) = &date_utc {
            validate_utc(date)?;
        }
        let reactions = parse_comment_extensions(child)?;
        comments.push(ExtensibleComment {
            durable_id,
            date_utc,
            intelligent_placeholder: optional_on_off(
                child,
                COMMENTS_EXTENSIBLE_NAMESPACE,
                "intelligentPlaceholder",
            )?,
            reactions,
        });
    }
    enforce_count("commentExtensible", comments.len())?;
    Ok(comments)
}

pub fn parse_people(xml: &[u8]) -> Result<Vec<Person>> {
    let document = parse_document(xml)?;
    let root = document.root()?;
    require_name(root, WORD_2012_NAMESPACE, "people")?;
    reject_attributes(root, &[])?;
    enforce_count("person", root.children.len())?;
    let mut people = Vec::with_capacity(root.children.len());
    let mut authors = HashSet::new();
    for child in &root.children {
        require_name(child, WORD_2012_NAMESPACE, "person")?;
        reject_attributes(child, &[(WORD_2012_NAMESPACE, "author")])?;
        let author = required_attr(child, WORD_2012_NAMESPACE, "author")?.to_owned();
        if !authors.insert(author.clone()) {
            return invalid(format!("duplicate people author '{author}'"));
        }
        let presence = match child.children.as_slice() {
            [] => None,
            [presence] => {
                require_name(presence, WORD_2012_NAMESPACE, "presenceInfo")?;
                reject_attributes(
                    presence,
                    &[
                        (WORD_2012_NAMESPACE, "providerId"),
                        (WORD_2012_NAMESPACE, "userId"),
                    ],
                )?;
                require_empty(presence)?;
                Some(PresenceInfo {
                    provider_id: required_attr(presence, WORD_2012_NAMESPACE, "providerId")?.into(),
                    user_id: required_attr(presence, WORD_2012_NAMESPACE, "userId")?.into(),
                })
            },
            _ => return invalid("person permits at most one presenceInfo".into()),
        };
        people.push(Person { author, presence });
    }
    Ok(people)
}

pub fn write_comments_extended(
    items: &[CommentExtension],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    validate_extended(items)?;
    let mut out = xml_header("w15", WORD_2012_NAMESPACE, "commentsEx", conformance);
    for item in items {
        out.push_str("<w15:commentEx w15:paraId=\"");
        out.push_str(&format_hex(item.paragraph_id));
        if let Some(parent) = item.parent_paragraph_id {
            out.push_str("\" w15:paraIdParent=\"");
            out.push_str(&format_hex(parent));
        }
        out.push_str("\" w15:done=\"");
        out.push_str(if item.done { "1" } else { "0" });
        out.push_str("\"/>");
    }
    out.push_str("</w15:commentsEx>");
    Ok(out.into_bytes())
}

pub fn write_comments_ids(
    items: &[CommentIdMapping],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    validate_ids(items)?;
    let mut out = xml_header("w16cid", COMMENTS_IDS_NAMESPACE, "commentsIds", conformance);
    for item in items {
        out.push_str("<w16cid:commentId w16cid:paraId=\"");
        out.push_str(&format_hex(item.paragraph_id));
        out.push_str("\" w16cid:durableId=\"");
        out.push_str(&format_hex(item.durable_id));
        out.push_str("\"/>");
    }
    out.push_str("</w16cid:commentsIds>");
    Ok(out.into_bytes())
}

pub fn write_comments_extensible(
    comments: &[ExtensibleComment],
    conformance: ModernCommentConformance,
) -> Result<Vec<u8>> {
    validate_extensible(comments)?;
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<w16cex:commentsExtensible xmlns:w16cex=\"");
    out.push_str(COMMENTS_EXTENSIBLE_NAMESPACE);
    out.push_str("\" xmlns:w16=\"");
    out.push_str(WORD_2018_NAMESPACE);
    out.push_str("\" xmlns:cr=\"");
    out.push_str(REACTIONS_NAMESPACE);
    out.push_str("\" xmlns:w=\"");
    out.push_str(conformance.word_namespace());
    out.push_str("\">");
    for comment in comments {
        out.push_str("<w16cex:commentExtensible w16cex:durableId=\"");
        out.push_str(&format_hex(comment.durable_id));
        if let Some(date) = &comment.date_utc {
            out.push_str("\" w16cex:dateUtc=\"");
            escape_attr(&mut out, date);
        }
        if let Some(value) = comment.intelligent_placeholder {
            out.push_str("\" w16cex:intelligentPlaceholder=\"");
            out.push_str(if value { "1" } else { "0" });
        }
        if comment.reactions.is_empty() {
            out.push_str("\"/>");
            continue;
        }
        out.push_str("\"><w16cex:extLst><w16:ext w16:uri=\"");
        out.push_str(REACTIONS_EXTENSION_URI);
        out.push_str("\"><cr:reactions>");
        for reaction in &comment.reactions {
            out.push_str("<cr:reaction reactionType=\"");
            out.push_str(&reaction.reaction_type.to_string());
            out.push_str("\">");
            for info in &reaction.reactions {
                out.push_str("<cr:reactionInfo");
                if let Some(date) = &info.date_utc {
                    out.push_str(" dateUtc=\"");
                    escape_attr(&mut out, date);
                    out.push('"');
                }
                if let Some(user) = &info.user {
                    out.push_str("><cr:user userId=\"");
                    escape_attr(&mut out, &user.user_id);
                    out.push_str("\" userName=\"");
                    escape_attr(&mut out, &user.user_name);
                    out.push_str("\" userProvider=\"");
                    escape_attr(&mut out, &user.user_provider);
                    out.push_str("\"/></cr:reactionInfo>");
                } else {
                    out.push_str("/>");
                }
            }
            out.push_str("</cr:reaction>");
        }
        out.push_str("</cr:reactions></w16:ext></w16cex:extLst></w16cex:commentExtensible>");
    }
    out.push_str("</w16cex:commentsExtensible>");
    Ok(out.into_bytes())
}

pub fn write_people(people: &[Person], conformance: ModernCommentConformance) -> Result<Vec<u8>> {
    validate_people(people)?;
    let mut out = xml_header("w15", WORD_2012_NAMESPACE, "people", conformance);
    for person in people {
        out.push_str("<w15:person w15:author=\"");
        escape_attr(&mut out, &person.author);
        if let Some(presence) = &person.presence {
            out.push_str("\"><w15:presenceInfo w15:providerId=\"");
            escape_attr(&mut out, &presence.provider_id);
            out.push_str("\" w15:userId=\"");
            escape_attr(&mut out, &presence.user_id);
            out.push_str("\"/></w15:person>");
        } else {
            out.push_str("\"/>");
        }
    }
    out.push_str("</w15:people>");
    Ok(out.into_bytes())
}

pub fn load_modern_comment_metadata(
    package: &OpcPackage,
    document_part_name: &PackURI,
) -> Result<ModernCommentMetadata> {
    reject_misplaced_relationships(package, document_part_name)?;
    let document = package.get_part(document_part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!("Word main document '{}': {error}", document_part_name.as_str()))
    })?;
    require_main_document_content_type(document.content_type())?;
    let metadata = ModernCommentMetadata {
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
    metadata: &ModernCommentMetadata,
    relationship_ids: &ModernCommentRelationshipIds,
    conformance: ModernCommentConformance,
) -> Result<()> {
    validate_metadata(metadata)?;
    let document = package.get_part(document_part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!("Word main document '{}': {error}", document_part_name.as_str()))
    })?;
    require_main_document_content_type(document.content_type())?;
    let specs = [
        (
            metadata.comments_extended.as_ref().map(|items| write_comments_extended(items, conformance)),
            relationship_ids.comments_extended.as_deref(),
            "/word/commentsExtended.xml",
            COMMENTS_EXTENDED_RELATIONSHIP,
            COMMENTS_EXTENDED_CONTENT_TYPE,
        ),
        (
            metadata.comments_ids.as_ref().map(|items| write_comments_ids(items, conformance)),
            relationship_ids.comments_ids.as_deref(),
            "/word/commentsIds.xml",
            COMMENTS_IDS_RELATIONSHIP,
            COMMENTS_IDS_CONTENT_TYPE,
        ),
        (
            metadata.comments_extensible.as_ref().map(|items| write_comments_extensible(items, conformance)),
            relationship_ids.comments_extensible.as_deref(),
            "/word/commentsExtensible.xml",
            COMMENTS_EXTENSIBLE_RELATIONSHIP,
            COMMENTS_EXTENSIBLE_CONTENT_TYPE,
        ),
        (
            metadata.people.as_ref().map(|items| write_people(items, conformance)),
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
            (None, Some(_)) => return invalid(format!("relationship ID supplied without {part_name}")),
            (Some(xml), Some(relationship_id)) => {
                if relationship_id.is_empty() || !ids.insert(relationship_id) {
                    return invalid("modern comment relationship IDs must be nonempty and unique".into());
                }
                let part_name = PackURI::new(part_name).map_err(OoxmlError::InvalidUri)?;
                if package.iter_parts().any(|part| part.partname() == &part_name)
                    || document.rels().iter().any(|relationship| {
                        relationship.r_id() == relationship_id
                            || relationship.reltype() == relationship_type
                    })
                {
                    return invalid(format!("modern comment part or relationship already exists: {}", part_name.as_str()));
                }
                pending.push((part_name, xml?, relationship_id.to_owned(), relationship_type, content_type));
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
        return invalid(format!("modern comment relationship '{}' must be internal", relationship.r_id()));
    }
    let target = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!("invalid modern comment target: {error}"))
    })?;
    let part = package.get_part(&target).map_err(|error| {
        OoxmlError::PartNotFound(format!("modern comment part '{}': {error}", target.as_str()))
    })?;
    if part.content_type() != content_type {
        return Err(OoxmlError::InvalidContentType {
            expected: content_type.into(),
            got: part.content_type().into(),
        });
    }
    if part.rels().iter().next().is_some() {
        return invalid(format!("modern comment part '{}' must not have relationships", target.as_str()));
    }
    parser(part.blob()).map(Some).map_err(|error| {
        OoxmlError::InvalidFormat(format!(
            "invalid modern comment part '{}': {error}",
            target.as_str()
        ))
    })
}

fn reject_misplaced_relationships(package: &OpcPackage, document_name: &PackURI) -> Result<()> {
    if package.rels().iter().any(|relationship| is_modern_relationship(relationship.reltype())) {
        return invalid("package root cannot source modern Word comment relationships".into());
    }
    for part in package.iter_parts() {
        if part.partname() != document_name
            && part.rels().iter().any(|relationship| is_modern_relationship(relationship.reltype()))
        {
            return invalid(format!("modern comment relationship has invalid source '{}'", part.partname().as_str()));
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

fn validate_metadata(metadata: &ModernCommentMetadata) -> Result<()> {
    if let Some(items) = &metadata.comments_extended {
        validate_extended(items)?;
    }
    if let Some(items) = &metadata.comments_ids {
        validate_ids(items)?;
    }
    if let Some(items) = &metadata.comments_extensible {
        validate_extensible(items)?;
    }
    if let Some(items) = &metadata.people {
        validate_people(items)?;
    }
    if let (Some(extended), Some(ids)) = (&metadata.comments_extended, &metadata.comments_ids) {
        let paragraphs: HashSet<_> = extended.iter().map(|item| item.paragraph_id).collect();
        if ids.iter().any(|item| !paragraphs.contains(&item.paragraph_id)) {
            return invalid("commentsIds references a paraId absent from commentsExtended".into());
        }
        if extended
            .iter()
            .filter_map(|item| item.parent_paragraph_id)
            .any(|parent| !paragraphs.contains(&parent))
        {
            return invalid("commentEx parent paraId is absent from commentsExtended".into());
        }
    }
    if let (Some(ids), Some(extensible)) = (&metadata.comments_ids, &metadata.comments_extensible) {
        let durable: HashSet<_> = ids.iter().map(|item| item.durable_id).collect();
        if extensible.iter().any(|item| !durable.contains(&item.durable_id)) {
            return invalid("commentsExtensible references a durableId absent from commentsIds".into());
        }
    }
    Ok(())
}

fn validate_extended(items: &[CommentExtension]) -> Result<()> {
    enforce_count("commentEx", items.len())?;
    let ids: HashSet<_> = items.iter().map(|item| item.paragraph_id).collect();
    if ids.len() != items.len() {
        return invalid("duplicate commentEx paraId".into());
    }
    Ok(())
}

fn validate_ids(items: &[CommentIdMapping]) -> Result<()> {
    enforce_count("commentId", items.len())?;
    let mut paragraphs = HashSet::new();
    let mut durable = HashSet::new();
    for item in items {
        validate_durable_id(item.durable_id)?;
        if !paragraphs.insert(item.paragraph_id) || !durable.insert(item.durable_id) {
            return invalid("duplicate commentsIds mapping".into());
        }
    }
    Ok(())
}

fn validate_extensible(items: &[ExtensibleComment]) -> Result<()> {
    enforce_count("commentExtensible", items.len())?;
    let mut ids = HashSet::new();
    for item in items {
        validate_durable_id(item.durable_id)?;
        if !ids.insert(item.durable_id) {
            return invalid("duplicate extensible durableId".into());
        }
        if let Some(date) = &item.date_utc {
            validate_utc(date)?;
        }
        enforce_count("reaction", item.reactions.len())?;
        let mut kinds = HashSet::new();
        for reaction in &item.reactions {
            if reaction.reaction_type == 0 || reaction.reaction_type >= 0x8000_0000 {
                return invalid("reactionType must be between 1 and 2147483647".into());
            }
            if !kinds.insert(reaction.reaction_type) {
                return invalid("duplicate reactionType on one comment".into());
            }
            enforce_count("reactionInfo", reaction.reactions.len())?;
            let mut users = HashSet::new();
            for info in &reaction.reactions {
                if let Some(date) = &info.date_utc {
                    validate_utc(date)?;
                }
                if let Some(user) = &info.user {
                    require_nonempty("reaction userId", &user.user_id)?;
                    require_nonempty("reaction userName", &user.user_name)?;
                    require_nonempty("reaction userProvider", &user.user_provider)?;
                    if !users.insert(user.user_id.clone()) {
                        return invalid("duplicate reaction userId for reactionType".into());
                    }
                }
            }
        }
    }
    Ok(())
}

fn validate_people(people: &[Person]) -> Result<()> {
    enforce_count("person", people.len())?;
    let mut authors = HashSet::new();
    for person in people {
        require_nonempty("person author", &person.author)?;
        if !authors.insert(person.author.clone()) {
            return invalid("duplicate people author".into());
        }
        if let Some(presence) = &person.presence {
            require_nonempty("presence providerId", &presence.provider_id)?;
            require_nonempty("presence userId", &presence.user_id)?;
        }
    }
    Ok(())
}

fn parse_comment_extensions(comment: &Node) -> Result<Vec<CommentReaction>> {
    if comment.children.is_empty() {
        return Ok(Vec::new());
    }
    if comment.children.len() != 1 {
        return invalid("commentExtensible permits at most one extLst".into());
    }
    let list = &comment.children[0];
    require_name(list, COMMENTS_EXTENSIBLE_NAMESPACE, "extLst")?;
    reject_attributes(list, &[])?;
    let mut reactions = Vec::new();
    for extension in &list.children {
        require_name(extension, WORD_2018_NAMESPACE, "ext")?;
        reject_attributes(extension, &[(WORD_2018_NAMESPACE, "uri")])?;
        let uri = required_attr(extension, WORD_2018_NAMESPACE, "uri")?;
        if uri != REACTIONS_EXTENSION_URI || extension.children.len() != 1 {
            return invalid(format!("unsupported commentsExtensible extension '{uri}'"));
        }
        let root = &extension.children[0];
        require_name(root, REACTIONS_NAMESPACE, "reactions")?;
        reject_attributes(root, &[])?;
        enforce_count("reaction", root.children.len())?;
        for reaction in &root.children {
            require_name(reaction, REACTIONS_NAMESPACE, "reaction")?;
            reject_attributes(reaction, &[("", "reactionType")])?;
            let reaction_type = required_attr(reaction, "", "reactionType")?
                .parse::<u32>()
                .map_err(|_| OoxmlError::InvalidFormat("invalid reactionType".into()))?;
            let mut infos = Vec::new();
            for info in &reaction.children {
                if info.local_name == "extLst" {
                    return invalid("reaction extension lists are not supported".into());
                }
                require_name(info, REACTIONS_NAMESPACE, "reactionInfo")?;
                reject_attributes(info, &[("", "dateUtc")])?;
                let date_utc = attr(info, "", "dateUtc").map(str::to_owned);
                if let Some(date) = &date_utc {
                    validate_utc(date)?;
                }
                let user = match info.children.as_slice() {
                    [] => None,
                    [user] => {
                        require_name(user, REACTIONS_NAMESPACE, "user")?;
                        reject_attributes(
                            user,
                            &[("", "userId"), ("", "userName"), ("", "userProvider")],
                        )?;
                        require_empty(user)?;
                        Some(CommentReactionUser {
                            user_id: required_attr(user, "", "userId")?.into(),
                            user_name: required_attr(user, "", "userName")?.into(),
                            user_provider: required_attr(user, "", "userProvider")?.into(),
                        })
                    },
                    _ => return invalid("reactionInfo permits at most one user".into()),
                };
                infos.push(CommentReactionInfo { date_utc, user });
            }
            reactions.push(CommentReaction {
                reaction_type,
                reactions: infos,
            });
        }
    }
    validate_extensible(&[ExtensibleComment {
        durable_id: 1,
        date_utc: None,
        intelligent_placeholder: None,
        reactions: reactions.clone(),
    }])?;
    Ok(reactions)
}

fn xml_header(
    prefix: &str,
    namespace: &str,
    root: &str,
    conformance: ModernCommentConformance,
) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><{prefix}:{root} xmlns:{prefix}=\"{namespace}\" xmlns:w=\"{}\">",
        conformance.word_namespace()
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
        invalid(format!("'{content_type}' is not a Word main-document content type"))
    }
}

fn validate_durable_id(value: u32) -> Result<()> {
    if value == 0 || value >= 0x7fff_ffff {
        invalid("durableId must be greater than 0 and less than 0x7FFFFFFF".into())
    } else {
        Ok(())
    }
}

fn validate_utc(value: &str) -> Result<()> {
    let parsed = DateTime::parse_from_rfc3339(value)
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid UTC dateTime '{value}'")))?;
    if parsed.offset().local_minus_utc() != 0 {
        invalid(format!("dateTime '{value}' is not UTC"))
    } else {
        Ok(())
    }
}

fn parse_hex(value: &str) -> Result<u32> {
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return invalid(format!("'{value}' is not ST_LongHexNumber"));
    }
    u32::from_str_radix(value, 16)
        .map_err(|_| OoxmlError::InvalidFormat(format!("invalid hex number '{value}'")))
}

fn format_hex(value: u32) -> String {
    format!("{value:08X}")
}

fn required_hex(node: &Node, namespace: &str, name: &str) -> Result<u32> {
    parse_hex(required_attr(node, namespace, name)?)
}

fn optional_hex(node: &Node, namespace: &str, name: &str) -> Result<Option<u32>> {
    attr(node, namespace, name).map(parse_hex).transpose()
}

fn optional_on_off(node: &Node, namespace: &str, name: &str) -> Result<Option<bool>> {
    attr(node, namespace, name)
        .map(|value| match value {
            "1" | "true" | "on" => Ok(true),
            "0" | "false" | "off" => Ok(false),
            _ => invalid(format!("invalid ST_OnOff value '{value}'")),
        })
        .transpose()
}

fn enforce_count(label: &str, count: usize) -> Result<()> {
    if count > MAX_MODERN_COMMENT_ITEMS {
        invalid(format!("{label} count exceeds {MAX_MODERN_COMMENT_ITEMS}"))
    } else {
        Ok(())
    }
}

fn require_nonempty(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{label} must not be empty"))
    } else {
        Ok(())
    }
}

fn escape_attr(out: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(character),
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
struct Attribute {
    namespace: String,
    local_name: String,
    value: String,
}

#[derive(Debug, PartialEq, Eq)]
struct Node {
    namespace: String,
    local_name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
}

struct XmlDocument {
    root: Option<Node>,
}

impl XmlDocument {
    fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| OoxmlError::InvalidFormat("modern comment XML has no root".into()))
    }
}

fn parse_document(xml: &[u8]) -> Result<XmlDocument> {
    if xml.len() > MAX_MODERN_COMMENT_PART_BYTES {
        return invalid(format!("modern comment part exceeds {MAX_MODERN_COMMENT_PART_BYTES} bytes"));
    }
    let mut capabilities = MceCapabilities::ooxml_baseline();
    for namespace in [
        WORD_2012_NAMESPACE,
        COMMENTS_IDS_NAMESPACE,
        COMMENTS_EXTENSIBLE_NAMESPACE,
        WORD_2018_NAMESPACE,
        REACTIONS_NAMESPACE,
    ] {
        capabilities.understand_namespace(namespace);
    }
    let limits = MceLimits {
        max_input_bytes: MAX_MODERN_COMMENT_PART_BYTES,
        max_output_bytes: MAX_MODERN_COMMENT_PART_BYTES * 2,
        max_depth: MAX_MODERN_COMMENT_DEPTH,
        max_namespace_bindings: 4096,
        max_directive_tokens: 4096,
        max_choices_per_alternate: 1024,
    };
    let processed = process_markup_compatibility(xml, &capabilities, &limits)?;
    build_dom(processed.xml.as_ref())
}

fn build_dom(xml: &[u8]) -> Result<XmlDocument> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut document = XmlDocument { root: None };
    let mut stack: Vec<(Node, HashMap<String, String>)> = Vec::new();
    let mut version = XmlVersion::Implicit1_0;
    let mut string_bytes = 0usize;
    loop {
        match reader.read_event_into(&mut buffer)? {
            Event::Decl(declaration) => version = declaration.xml_version()?,
            Event::Start(element) => push_node(
                &reader,
                &element,
                &mut document,
                &mut stack,
                version,
                &mut string_bytes,
                false,
            )?,
            Event::Empty(element) => push_node(
                &reader,
                &element,
                &mut document,
                &mut stack,
                version,
                &mut string_bytes,
                true,
            )?,
            Event::End(_) if stack.is_empty() => return invalid("unexpected XML end tag".into()),
            Event::End(_) => {
                let (node, _) = stack.pop().expect("checked above");
                attach_node(&mut document, &mut stack, node)?;
            },
            Event::DocType(_) => return invalid("DTD is forbidden in modern comment parts".into()),
            Event::Text(text) if !is_whitespace(text.as_ref()) => {
                return invalid("text is not permitted in modern comment metadata".into());
            },
            Event::CData(text) if !is_whitespace(text.as_ref()) => {
                return invalid("CDATA is not permitted in modern comment metadata".into());
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed modern comment XML element".into());
    }
    Ok(document)
}

fn push_node(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    document: &mut XmlDocument,
    stack: &mut Vec<(Node, HashMap<String, String>)>,
    version: XmlVersion,
    string_bytes: &mut usize,
    empty: bool,
) -> Result<()> {
    if stack.len() >= MAX_MODERN_COMMENT_DEPTH {
        return invalid(format!("modern comment XML depth exceeds {MAX_MODERN_COMMENT_DEPTH}"));
    }
    let mut namespaces = stack.last().map(|(_, map)| map.clone()).unwrap_or_default();
    namespaces.insert("xml".into(), "http://www.w3.org/XML/1998/namespace".into());
    let mut raw_attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        *string_bytes = string_bytes.saturating_add(name.len() + value.len());
        if *string_bytes > MAX_MODERN_COMMENT_STRING_BYTES {
            return invalid("modern comment strings exceed allocation cap".into());
        }
        if name == "xmlns" {
            namespaces.insert(String::new(), value);
        } else if let Some(prefix) = name.strip_prefix("xmlns:") {
            namespaces.insert(prefix.into(), value);
        } else {
            raw_attributes.push((name, value));
        }
    }
    let qname = element.name();
    let name = std::str::from_utf8(qname.as_ref())
        .map_err(|error| OoxmlError::Xml(error.to_string()))?;
    let (prefix, local_name) = split_name(name);
    let namespace = if prefix.is_empty() {
        namespaces.get("").cloned().unwrap_or_default()
    } else {
        namespaces.get(prefix).cloned().ok_or_else(|| {
            OoxmlError::InvalidFormat(format!("unbound XML prefix '{prefix}'"))
        })?
    };
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (name, value) in raw_attributes {
        let (prefix, local_name) = split_name(&name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces.get(prefix).cloned().ok_or_else(|| {
                OoxmlError::InvalidFormat(format!("unbound attribute prefix '{prefix}'"))
            })?
        };
        if !seen.insert((namespace.clone(), local_name.to_owned())) {
            return invalid(format!("duplicate attribute {{{namespace}}}{local_name}"));
        }
        attributes.push(Attribute {
            namespace,
            local_name: local_name.into(),
            value,
        });
    }
    let node = Node {
        namespace,
        local_name: local_name.into(),
        attributes,
        children: Vec::new(),
    };
    if empty {
        attach_node(document, stack, node)
    } else {
        stack.push((node, namespaces));
        Ok(())
    }
}

fn attach_node(
    document: &mut XmlDocument,
    stack: &mut [(Node, HashMap<String, String>)],
    node: Node,
) -> Result<()> {
    if let Some((parent, _)) = stack.last_mut() {
        parent.children.push(node);
    } else if document.root.replace(node).is_some() {
        return invalid("modern comment XML has multiple roots".into());
    }
    Ok(())
}

fn split_name(value: &str) -> (&str, &str) {
    value.split_once(':').unwrap_or(("", value))
}

fn require_name(node: &Node, namespace: &str, local_name: &str) -> Result<()> {
    if node.namespace == namespace && node.local_name == local_name {
        Ok(())
    } else {
        invalid(format!(
            "expected {{{namespace}}}{local_name}, got {{{}}}{}",
            node.namespace, node.local_name
        ))
    }
}

fn attr<'a>(node: &'a Node, namespace: &str, local_name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.namespace == namespace && attribute.local_name == local_name)
        .map(|attribute| attribute.value.as_str())
}

fn required_attr<'a>(node: &'a Node, namespace: &str, local_name: &str) -> Result<&'a str> {
    attr(node, namespace, local_name).ok_or_else(|| {
        OoxmlError::InvalidFormat(format!(
            "{} requires attribute {{{namespace}}}{local_name}",
            node.local_name
        ))
    })
}

fn reject_attributes(node: &Node, allowed: &[(&str, &str)]) -> Result<()> {
    for attribute in &node.attributes {
        if !allowed.iter().any(|(namespace, name)| {
            attribute.namespace == *namespace && attribute.local_name == *name
        }) {
            return invalid(format!(
                "unexpected attribute {{{}}}{} on {}",
                attribute.namespace, attribute.local_name, node.local_name
            ));
        }
    }
    Ok(())
}

fn require_empty(node: &Node) -> Result<()> {
    if node.children.is_empty() {
        Ok(())
    } else {
        invalid(format!("{} must be empty", node.local_name))
    }
}

fn is_whitespace(value: &[u8]) -> bool {
    value.iter().all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid<T>(message: String) -> Result<T> {
    Err(OoxmlError::InvalidFormat(message))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    const POI_DOCX: &[u8] =
        include_bytes!("../../../../3rdparty/poi/test-data/document/testComment.docx");
    const LO_DOCX: &[u8] = include_bytes!(
        "../../../../3rdparty/libreoffice-core/sw/qa/writerfilter/dmapper/data/redline-range-comment.docx"
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
        let metadata = load_modern_comment_metadata(
            &package,
            &PackURI::new("/word/document.xml").unwrap(),
        )
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
        assert!(std::str::from_utf8(&extended).unwrap().contains(STRICT_WORD_NAMESPACE));
        assert_eq!(parse_comments_extended(&extended).unwrap(), metadata.comments_extended.unwrap());

        let ids = write_comments_ids(metadata.comments_ids.as_ref().unwrap(), ModernCommentConformance::Strict).unwrap();
        assert_eq!(parse_comments_ids(&ids).unwrap(), metadata.comments_ids.unwrap());
        let extensible = write_comments_extensible(metadata.comments_extensible.as_ref().unwrap(), ModernCommentConformance::Strict).unwrap();
        assert_eq!(parse_comments_extensible(&extensible).unwrap(), metadata.comments_extensible.unwrap());
        let people = write_people(metadata.people.as_ref().unwrap(), ModernCommentConformance::Strict).unwrap();
        assert_eq!(parse_people(&people).unwrap(), metadata.people.unwrap());
    }

    #[test]
    fn mce_selects_fallback_comment() {
        let xml = format!(
            r#"<w15:commentsEx xmlns:w15="{WORD_2012_NAMESPACE}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><w15:commentEx w15:paraId="AAAAAAAA"/></mc:Choice><mc:Fallback><w15:commentEx w15:paraId="BBBBBBBB"/></mc:Fallback></mc:AlternateContent></w15:commentsEx>"#
        );
        assert_eq!(parse_comments_extended(xml.as_bytes()).unwrap()[0].paragraph_id, 0xBBBB_BBBB);
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
            load_modern_comment_metadata(
                &package,
                &PackURI::new("/word/document.xml").unwrap()
            )
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
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml".into(),
            b"<document/>".to_vec(),
        );
        document.rels_mut().add_relationship(
            COMMENTS_EXTENDED_RELATIONSHIP.into(),
            "https://example.invalid/comments".into(),
            "rId1".into(),
            true,
        );
        package.add_part(Box::new(document));
        assert!(load_modern_comment_metadata(
            &package,
            &PackURI::new("/word/document.xml").unwrap()
        )
        .is_err());
    }

    #[test]
    fn enforces_size_depth_count_and_reaction_constraints() {
        assert!(parse_people(&vec![b' '; MAX_MODERN_COMMENT_PART_BYTES + 1]).is_err());
        let deep = format!("{}{}", "<x>".repeat(MAX_MODERN_COMMENT_DEPTH + 1), "</x>".repeat(MAX_MODERN_COMMENT_DEPTH + 1));
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
                        }),
                    }],
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
