//! Bounded XML, MCE, and reaction codec for modern Word comments.

use super::model::{
    COMMENTS_EXTENSIBLE_NAMESPACE, COMMENTS_IDS_NAMESPACE, Comment, Conformance, Extended,
    Extension, ExtensionList, IdMapping, MAX_MODERN_COMMENT_DEPTH, MAX_MODERN_COMMENT_ITEMS,
    MAX_MODERN_COMMENT_PART_BYTES, MAX_MODERN_COMMENT_STRING_BYTES, Metadata,
    OFFICE_EXTENSION_LIST_NAMESPACE, Person, Presence, REACTIONS_NAMESPACE, Reaction, ReactionInfo,
    ReactionUser, WORD_2012_NAMESPACE, WORD_2018_NAMESPACE,
};
use crate::{Error, Result};
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};
use std::collections::{HashMap, HashSet};

pub(super) const REACTIONS_EXTENSION_URI: &str = "{CE6994B0-6A32-4C9F-8C6B-6E91EDA988CE}";

impl Extension {
    pub fn new(uri: Option<String>, child_xml: impl Into<String>) -> Result<Self> {
        let uri = uri.map(|value| normalize_xsd_token(&value));
        let child_xml = canonical_extension_child(&child_xml.into())?;
        Ok(Self { uri, child_xml })
    }

    pub fn uri(&self) -> Option<&str> {
        self.uri.as_deref()
    }

    pub fn child_xml(&self) -> &str {
        &self.child_xml
    }

    pub fn set_uri(&mut self, uri: Option<String>) {
        self.uri = uri.map(|value| normalize_xsd_token(&value));
    }

    pub fn set_child_xml(&mut self, child_xml: impl Into<String>) -> Result<()> {
        self.child_xml = canonical_extension_child(&child_xml.into())?;
        Ok(())
    }
}

impl ExtensionList {
    pub fn new(extensions: Vec<Extension>) -> Result<Self> {
        enforce_count("modern comment extension", extensions.len())?;
        let list = Self { extensions };
        validate_extension_list(&list)?;
        Ok(list)
    }

    pub fn extensions(&self) -> &[Extension] {
        &self.extensions
    }

    pub fn push(&mut self, extension: Extension) -> Result<()> {
        enforce_count(
            "modern comment extension",
            self.extensions.len().saturating_add(1),
        )?;
        self.extensions.push(extension);
        Ok(())
    }

    pub fn remove(&mut self, index: usize) -> Option<Extension> {
        (index < self.extensions.len()).then(|| self.extensions.remove(index))
    }
}

pub fn parse_comments_extended(xml: &[u8]) -> Result<Vec<Extended>> {
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
            return invalid(format!(
                "duplicate commentEx paraId {}",
                format_hex(paragraph_id)
            ));
        }
        items.push(Extended {
            paragraph_id,
            parent_paragraph_id: optional_hex(child, WORD_2012_NAMESPACE, "paraIdParent")?,
            done: optional_on_off(child, WORD_2012_NAMESPACE, "done")?.unwrap_or(false),
        });
    }
    Ok(items)
}

pub fn parse_comments_ids(xml: &[u8]) -> Result<Vec<IdMapping>> {
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
        items.push(IdMapping {
            paragraph_id,
            durable_id,
        });
    }
    Ok(items)
}

pub fn parse_comments_extensible(xml: &[u8]) -> Result<Vec<Comment>> {
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
            return invalid(format!(
                "duplicate extensible durableId {}",
                format_hex(durable_id)
            ));
        }
        let date_utc = attr(child, COMMENTS_EXTENSIBLE_NAMESPACE, "dateUtc").map(str::to_owned);
        if let Some(date) = &date_utc {
            validate_utc(date)?;
        }
        let reactions = parse_comment_extensions(child)?;
        comments.push(Comment {
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
                Some(Presence {
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

pub fn write_comments_extended(items: &[Extended], conformance: Conformance) -> Result<Vec<u8>> {
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

pub fn write_comments_ids(items: &[IdMapping], conformance: Conformance) -> Result<Vec<u8>> {
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
    comments: &[Comment],
    conformance: Conformance,
) -> Result<Vec<u8>> {
    validate_extensible(comments)?;
    let mut out = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
    out.push_str("<w16cex:commentsExtensible xmlns:w16cex=\"");
    out.push_str(COMMENTS_EXTENSIBLE_NAMESPACE);
    out.push_str("\" xmlns:w16=\"");
    out.push_str(WORD_2018_NAMESPACE);
    out.push_str("\" xmlns:cr=\"");
    out.push_str(REACTIONS_NAMESPACE);
    out.push_str("\" xmlns:oel=\"");
    out.push_str(OFFICE_EXTENSION_LIST_NAMESPACE);
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
                if info.user.is_none() && info.extensions.is_none() {
                    out.push_str("/>");
                    continue;
                }
                out.push('>');
                if let Some(user) = &info.user {
                    out.push_str("<cr:user userId=\"");
                    escape_attr(&mut out, &user.user_id);
                    out.push_str("\" userName=\"");
                    escape_attr(&mut out, &user.user_name);
                    out.push_str("\" userProvider=\"");
                    escape_attr(&mut out, &user.user_provider);
                    if let Some(extensions) = &user.extensions {
                        out.push_str("\">");
                        write_extension_list(&mut out, extensions);
                        out.push_str("</cr:user>");
                    } else {
                        out.push_str("\"/>");
                    }
                }
                if let Some(extensions) = &info.extensions {
                    write_extension_list(&mut out, extensions);
                }
                out.push_str("</cr:reactionInfo>");
            }
            if let Some(extensions) = &reaction.extensions {
                write_extension_list(&mut out, extensions);
            }
            out.push_str("</cr:reaction>");
        }
        out.push_str("</cr:reactions></w16:ext></w16cex:extLst></w16cex:commentExtensible>");
    }
    out.push_str("</w16cex:commentsExtensible>");
    Ok(out.into_bytes())
}

pub fn write_people(people: &[Person], conformance: Conformance) -> Result<Vec<u8>> {
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

pub(super) fn validate_metadata(metadata: &Metadata) -> Result<()> {
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
        if ids
            .iter()
            .any(|item| !paragraphs.contains(&item.paragraph_id))
        {
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
        if extensible
            .iter()
            .any(|item| !durable.contains(&item.durable_id))
        {
            return invalid(
                "commentsExtensible references a durableId absent from commentsIds".into(),
            );
        }
    }
    Ok(())
}

fn validate_extended(items: &[Extended]) -> Result<()> {
    enforce_count("commentEx", items.len())?;
    let ids: HashSet<_> = items.iter().map(|item| item.paragraph_id).collect();
    if ids.len() != items.len() {
        return invalid("duplicate commentEx paraId".into());
    }
    Ok(())
}

fn validate_ids(items: &[IdMapping]) -> Result<()> {
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

fn validate_extensible(items: &[Comment]) -> Result<()> {
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
            if let Some(extensions) = &reaction.extensions {
                validate_extension_list(extensions)?;
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
                    if let Some(extensions) = &user.extensions {
                        validate_extension_list(extensions)?;
                    }
                }
                if let Some(extensions) = &info.extensions {
                    validate_extension_list(extensions)?;
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

fn parse_comment_extensions(comment: &Node) -> Result<Vec<Reaction>> {
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
                .map_err(|_| Error::Invalid("invalid reactionType".into()))?;
            let mut infos = Vec::new();
            let mut extensions = None;
            for (index, info) in reaction.children.iter().enumerate() {
                if info.namespace == OFFICE_EXTENSION_LIST_NAMESPACE && info.local_name == "extLst"
                {
                    if extensions.is_some() || index + 1 != reaction.children.len() {
                        return invalid("reaction extLst must occur once at the end".into());
                    }
                    extensions = Some(parse_extension_list(info)?);
                    continue;
                }
                if extensions.is_some() {
                    return invalid("reactionInfo occurs after reaction extLst".into());
                }
                require_name(info, REACTIONS_NAMESPACE, "reactionInfo")?;
                reject_attributes(info, &[("", "dateUtc")])?;
                let date_utc = attr(info, "", "dateUtc").map(str::to_owned);
                if let Some(date) = &date_utc {
                    validate_utc(date)?;
                }
                let mut user = None;
                let mut info_extensions = None;
                for (child_index, child) in info.children.iter().enumerate() {
                    if child.namespace == OFFICE_EXTENSION_LIST_NAMESPACE
                        && child.local_name == "extLst"
                    {
                        if info_extensions.is_some() || child_index + 1 != info.children.len() {
                            return invalid(
                                "reactionInfo extLst must occur once at the end".into(),
                            );
                        }
                        info_extensions = Some(parse_extension_list(child)?);
                    } else if user.is_none() && info_extensions.is_none() {
                        user = Some(parse_reaction_user(child)?);
                    } else {
                        return invalid(
                            "reactionInfo permits one user followed by one extLst".into(),
                        );
                    }
                }
                infos.push(ReactionInfo {
                    date_utc,
                    user,
                    extensions: info_extensions,
                });
            }
            reactions.push(Reaction {
                reaction_type,
                reactions: infos,
                extensions,
            });
        }
    }
    validate_extensible(&[Comment {
        durable_id: 1,
        date_utc: None,
        intelligent_placeholder: None,
        reactions: reactions.clone(),
    }])?;
    Ok(reactions)
}

fn parse_reaction_user(node: &Node) -> Result<ReactionUser> {
    require_name(node, REACTIONS_NAMESPACE, "user")?;
    reject_attributes(
        node,
        &[("", "userId"), ("", "userName"), ("", "userProvider")],
    )?;
    let extensions = match node.children.as_slice() {
        [] => None,
        [extensions] => Some(parse_extension_list(extensions)?),
        _ => return invalid("reaction user permits at most one extLst".into()),
    };
    Ok(ReactionUser {
        user_id: required_attr(node, "", "userId")?.into(),
        user_name: required_attr(node, "", "userName")?.into(),
        user_provider: required_attr(node, "", "userProvider")?.into(),
        extensions,
    })
}

fn parse_extension_list(node: &Node) -> Result<ExtensionList> {
    require_name(node, OFFICE_EXTENSION_LIST_NAMESPACE, "extLst")?;
    reject_attributes(node, &[])?;
    enforce_count("modern comment extension", node.children.len())?;
    let mut extensions = Vec::with_capacity(node.children.len());
    for extension in &node.children {
        require_name(extension, OFFICE_EXTENSION_LIST_NAMESPACE, "ext")?;
        reject_attributes(extension, &[("", "uri")])?;
        if extension.has_non_whitespace_text || extension.children.len() != 1 {
            return invalid("oel:ext requires exactly one lax child element".into());
        }
        extensions.push(Extension::new(
            attr(extension, "", "uri").map(str::to_owned),
            extension.children[0].raw_xml.clone(),
        )?);
    }
    ExtensionList::new(extensions)
}

fn write_extension_list(out: &mut String, list: &ExtensionList) {
    out.push_str("<oel:extLst>");
    for extension in list.extensions() {
        out.push_str("<oel:ext");
        if let Some(uri) = extension.uri() {
            out.push_str(" uri=\"");
            escape_attr(out, uri);
            out.push('"');
        }
        out.push('>');
        out.push_str(extension.child_xml());
        out.push_str("</oel:ext>");
    }
    out.push_str("</oel:extLst>");
}

fn validate_extension_list(list: &ExtensionList) -> Result<()> {
    enforce_count("modern comment extension", list.extensions.len())?;
    for extension in &list.extensions {
        if let Some(uri) = &extension.uri
            && uri != &normalize_xsd_token(uri)
        {
            return invalid("extension uri is not a normalized xsd:token".into());
        }
        if extension.child_xml.len() > MAX_MODERN_COMMENT_PART_BYTES {
            return invalid("extension child XML exceeds the part-size bound".into());
        }
        canonical_extension_child(&extension.child_xml)?;
    }
    Ok(())
}

fn normalize_xsd_token(value: &str) -> String {
    value
        .split([' ', '\t', '\r', '\n'])
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join(" ")
}

fn canonical_extension_child(value: &str) -> Result<String> {
    if value.len() > MAX_MODERN_COMMENT_PART_BYTES {
        return invalid("extension child XML exceeds the part-size bound".into());
    }
    if value.trim_start().starts_with("<?xml") {
        return invalid("extension child XML cannot contain an XML declaration".into());
    }
    let document = build_dom(value.as_bytes())?;
    Ok(document.root()?.raw_xml.clone())
}

fn xml_header(prefix: &str, namespace: &str, root: &str, conformance: Conformance) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><{prefix}:{root} xmlns:{prefix}=\"{namespace}\" xmlns:w=\"{}\">",
        conformance.word_namespace()
    )
}

fn validate_durable_id(value: u32) -> Result<()> {
    if value == 0 || value >= 0x7fff_ffff {
        invalid("durableId must be greater than 0 and less than 0x7FFFFFFF".into())
    } else {
        Ok(())
    }
}

fn validate_utc(value: &str) -> Result<()> {
    if value.trim() != value || !value.is_ascii() {
        return invalid(format!("invalid UTC dateTime '{value}'"));
    }
    litchi_ooxml_common::properties::time::DateTime::new(value)
        .map_err(|_| Error::Invalid(format!("invalid UTC dateTime '{value}'")))?;
    if value.ends_with('Z') || value.ends_with("+00:00") || value.ends_with("-00:00") {
        Ok(())
    } else {
        invalid(format!("dateTime '{value}' is not UTC"))
    }
}

fn parse_hex(value: &str) -> Result<u32> {
    if value.len() != 8 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return invalid(format!("'{value}' is not ST_LongHexNumber"));
    }
    u32::from_str_radix(value, 16)
        .map_err(|_| Error::Invalid(format!("invalid hex number '{value}'")))
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
    raw_xml: String,
    has_non_whitespace_text: bool,
}

struct XmlDocument {
    root: Option<Node>,
}

impl XmlDocument {
    fn root(&self) -> Result<&Node> {
        self.root
            .as_ref()
            .ok_or_else(|| Error::Invalid("modern comment XML has no root".into()))
    }
}

fn parse_document(xml: &[u8]) -> Result<XmlDocument> {
    if xml.len() > MAX_MODERN_COMMENT_PART_BYTES {
        return invalid(format!(
            "modern comment part exceeds {MAX_MODERN_COMMENT_PART_BYTES} bytes"
        ));
    }
    let mut capabilities = MceCapabilities::ooxml_baseline();
    for namespace in [
        WORD_2012_NAMESPACE,
        COMMENTS_IDS_NAMESPACE,
        COMMENTS_EXTENSIBLE_NAMESPACE,
        WORD_2018_NAMESPACE,
        REACTIONS_NAMESPACE,
        OFFICE_EXTENSION_LIST_NAMESPACE,
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
    let document = build_dom(processed.xml.as_ref())?;
    validate_metadata_text(document.root()?, false)?;
    Ok(document)
}

struct StackEntry {
    node: Node,
    namespaces: HashMap<String, String>,
    raw: Vec<u8>,
}

fn build_dom(xml: &[u8]) -> Result<XmlDocument> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut document = XmlDocument { root: None };
    let mut stack: Vec<StackEntry> = Vec::new();
    let mut version = XmlVersion::Implicit1_0;
    let mut string_bytes = 0usize;
    let mut node_count = 0usize;
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
                &mut node_count,
                false,
            )?,
            Event::Empty(element) => push_node(
                &reader,
                &element,
                &mut document,
                &mut stack,
                version,
                &mut string_bytes,
                &mut node_count,
                true,
            )?,
            Event::End(_) if stack.is_empty() => return invalid("unexpected XML end tag".into()),
            Event::End(end) => {
                let mut entry = stack.pop().expect("checked above");
                entry
                    .raw
                    .extend_from_slice(&serialize_event(Event::End(end.into_owned()))?);
                entry.node.raw_xml = String::from_utf8(entry.raw.clone())
                    .map_err(|error| Error::Xml(error.to_string()))?;
                attach_node(&mut document, &mut stack, entry.node, entry.raw)?;
            },
            Event::DocType(_) => return invalid("DTD is forbidden in modern comment parts".into()),
            Event::Text(text) => {
                let non_whitespace = !is_whitespace(text.as_ref());
                let Some(entry) = stack.last_mut() else {
                    if non_whitespace {
                        return invalid("text outside the XML root".into());
                    }
                    buffer.clear();
                    continue;
                };
                entry.node.has_non_whitespace_text |= non_whitespace;
                entry
                    .raw
                    .extend_from_slice(&serialize_event(Event::Text(text.into_owned()))?);
            },
            Event::CData(text) => {
                let non_whitespace = !is_whitespace(text.as_ref());
                let Some(entry) = stack.last_mut() else {
                    if non_whitespace {
                        return invalid("CDATA outside the XML root".into());
                    }
                    buffer.clear();
                    continue;
                };
                entry.node.has_non_whitespace_text |= non_whitespace;
                entry
                    .raw
                    .extend_from_slice(&serialize_event(Event::CData(text.into_owned()))?);
            },
            Event::Comment(comment) => {
                if let Some(entry) = stack.last_mut() {
                    entry
                        .raw
                        .extend_from_slice(&serialize_event(Event::Comment(comment.into_owned()))?);
                }
            },
            Event::PI(instruction) => {
                if let Some(entry) = stack.last_mut() {
                    entry
                        .raw
                        .extend_from_slice(&serialize_event(Event::PI(instruction.into_owned()))?);
                }
            },
            Event::GeneralRef(reference) => {
                if let Some(entry) = stack.last_mut() {
                    entry
                        .raw
                        .extend_from_slice(&serialize_event(Event::GeneralRef(
                            reference.into_owned(),
                        ))?);
                }
            },
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return invalid("unclosed modern comment XML element".into());
    }
    Ok(document)
}

#[allow(clippy::too_many_arguments)]
fn push_node(
    reader: &Reader<&[u8]>,
    element: &BytesStart<'_>,
    document: &mut XmlDocument,
    stack: &mut Vec<StackEntry>,
    version: XmlVersion,
    string_bytes: &mut usize,
    node_count: &mut usize,
    empty: bool,
) -> Result<()> {
    if stack.len() >= MAX_MODERN_COMMENT_DEPTH {
        return invalid(format!(
            "modern comment XML depth exceeds {MAX_MODERN_COMMENT_DEPTH}"
        ));
    }
    *node_count = node_count.saturating_add(1);
    if *node_count > MAX_MODERN_COMMENT_ITEMS {
        return invalid(format!(
            "modern comment XML node count exceeds {MAX_MODERN_COMMENT_ITEMS}"
        ));
    }
    let mut namespaces = stack
        .last()
        .map(|entry| entry.namespaces.clone())
        .unwrap_or_default();
    namespaces.insert("xml".into(), "http://www.w3.org/XML/1998/namespace".into());
    let mut raw_attributes = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        let value = attribute
            .decoded_and_normalized_value(version, reader.decoder())
            .map_err(|error| Error::Xml(error.to_string()))?
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
    let name =
        std::str::from_utf8(qname.as_ref()).map_err(|error| Error::Xml(error.to_string()))?;
    let (prefix, local_name) = split_name(name);
    let namespace = if prefix.is_empty() {
        namespaces.get("").cloned().unwrap_or_default()
    } else {
        namespaces
            .get(prefix)
            .cloned()
            .ok_or_else(|| Error::Invalid(format!("unbound XML prefix '{prefix}'")))?
    };
    let mut attributes = Vec::with_capacity(raw_attributes.len());
    let mut seen = HashSet::new();
    for (name, value) in &raw_attributes {
        let (prefix, local_name) = split_name(name);
        let namespace = if prefix.is_empty() {
            String::new()
        } else {
            namespaces
                .get(prefix)
                .cloned()
                .ok_or_else(|| Error::Invalid(format!("unbound attribute prefix '{prefix}'")))?
        };
        if !seen.insert((namespace.clone(), local_name.to_owned())) {
            return invalid(format!("duplicate attribute {{{namespace}}}{local_name}"));
        }
        attributes.push(Attribute {
            namespace,
            local_name: local_name.into(),
            value: value.clone(),
        });
    }
    let raw = canonical_element_start(name, &namespaces, &raw_attributes, empty)?;
    let mut node = Node {
        namespace,
        local_name: local_name.into(),
        attributes,
        children: Vec::new(),
        raw_xml: String::new(),
        has_non_whitespace_text: false,
    };
    if empty {
        node.raw_xml =
            String::from_utf8(raw.clone()).map_err(|error| Error::Xml(error.to_string()))?;
        attach_node(document, stack, node, raw)
    } else {
        stack.push(StackEntry {
            node,
            namespaces,
            raw,
        });
        Ok(())
    }
}

fn attach_node(
    document: &mut XmlDocument,
    stack: &mut [StackEntry],
    node: Node,
    raw: Vec<u8>,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.raw.extend_from_slice(&raw);
        parent.node.children.push(node);
    } else if document.root.replace(node).is_some() {
        return invalid("modern comment XML has multiple roots".into());
    }
    Ok(())
}

fn canonical_element_start(
    name: &str,
    namespaces: &HashMap<String, String>,
    attributes: &[(String, String)],
    empty: bool,
) -> Result<Vec<u8>> {
    let mut element = BytesStart::new(name.to_owned());
    let mut used_prefixes = HashSet::new();
    used_prefixes.insert(split_name(name).0.to_string());
    for (name, _) in attributes {
        let prefix = split_name(name).0;
        if !prefix.is_empty() {
            used_prefixes.insert(prefix.to_string());
        }
    }
    let mut declarations: Vec<_> = namespaces
        .iter()
        .filter(|(prefix, _)| prefix.as_str() != "xml" && used_prefixes.contains(*prefix))
        .map(|(prefix, value)| {
            (
                if prefix.is_empty() {
                    "xmlns".to_string()
                } else {
                    format!("xmlns:{prefix}")
                },
                value.clone(),
            )
        })
        .collect();
    declarations.sort_by(|left, right| left.0.cmp(&right.0));
    for (name, value) in &declarations {
        element.push_attribute((name.as_str(), value.as_str()));
    }
    for (name, value) in attributes {
        element.push_attribute((name.as_str(), value.as_str()));
    }
    serialize_event(if empty {
        Event::Empty(element)
    } else {
        Event::Start(element)
    })
}

fn serialize_event(event: Event<'_>) -> Result<Vec<u8>> {
    let mut writer = Writer::new(Vec::new());
    writer
        .write_event(event)
        .map_err(|error| Error::Xml(error.to_string()))?;
    Ok(writer.into_inner())
}

fn validate_metadata_text(node: &Node, inside_lax_child: bool) -> Result<()> {
    if node.has_non_whitespace_text && !inside_lax_child {
        return invalid("text is not permitted in modern comment metadata".into());
    }
    let children_are_lax = inside_lax_child
        || (node.namespace == OFFICE_EXTENSION_LIST_NAMESPACE && node.local_name == "ext");
    for child in &node.children {
        validate_metadata_text(child, children_are_lax)?;
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
        Error::Invalid(format!(
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
    value
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

fn invalid<T>(message: String) -> Result<T> {
    Err(Error::Invalid(message))
}
