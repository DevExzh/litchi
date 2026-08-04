//! Typed ISO/IEC 29500 PresentationML legacy comment parts and package graph.
//!
//! Comment text and extension payloads are inert data. This module never
//! resolves identities, follows external targets, or executes embedded data.

use crate::{Error, Result};
use chrono::{DateTime, NaiveDateTime};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const COMMENTS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
const STRICT_COMMENTS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/comments";
const AUTHORS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors";
const STRICT_AUTHORS_REL: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors";
const SLIDE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const STRICT_SLIDE_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const AUTHORS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";
const COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const SLIDE_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const MAX_PART_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 128;
const MAX_NODES: usize = 262_144;
const MAX_STRING_BYTES: usize = 1024 * 1024;
const MAX_AUTHORS: usize = 4096;
const MAX_SLIDES: usize = 100_000;
const MAX_COMMENTS_PER_SLIDE: usize = 65_536;
const MAX_TOTAL_COMMENTS: usize = 1_000_000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PresentationCommentConformance {
    Transitional,
    Strict,
}

impl PresentationCommentConformance {
    fn namespace(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }

    fn comments_relationship(self) -> &'static str {
        match self {
            Self::Transitional => COMMENTS_REL,
            Self::Strict => STRICT_COMMENTS_REL,
        }
    }

    fn authors_relationship(self) -> &'static str {
        match self {
            Self::Transitional => AUTHORS_REL,
            Self::Strict => STRICT_AUTHORS_REL,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationCommentAuthor {
    pub id: u32,
    pub name: String,
    pub initials: String,
    pub last_index: u32,
    pub color_index: u32,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationComment {
    pub author_id: u32,
    pub date_time: Option<String>,
    pub index: u32,
    pub x: i64,
    pub y: i64,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SlideCommentList {
    pub slide_part_name: String,
    pub relationship_id: String,
    pub part_name: String,
    pub comments: Vec<PresentationComment>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationComments {
    pub author_relationship_id: String,
    pub author_part_name: String,
    pub authors: Vec<PresentationCommentAuthor>,
    pub slides: Vec<SlideCommentList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Attribute {
    name: String,
    value: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Node {
    namespace: String,
    name: String,
    attributes: Vec<Attribute>,
    children: Vec<Node>,
    text: String,
}

pub fn parse_comment_authors(xml: &[u8]) -> Result<Vec<PresentationCommentAuthor>> {
    let root = parse_document(xml)?;
    let namespace = root_namespace(&root, "cmAuthorLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    if root.children.len() > MAX_AUTHORS {
        return Err(invalid("comment-author count exceeds limit"));
    }
    let mut authors = Vec::with_capacity(root.children.len());
    let mut ids = HashSet::new();
    for node in &root.children {
        require_name(node, namespace, "cmAuthor")?;
        whitespace_only(node)?;
        no_attributes(node, &["id", "name", "initials", "lastIdx", "clrIdx"])?;
        if node.children.len() > 1
            || node
                .children
                .first()
                .is_some_and(|child| child.namespace != namespace || child.name != "extLst")
        {
            return Err(invalid("cmAuthor permits only one trailing extLst"));
        }
        let author = PresentationCommentAuthor {
            id: required_u32(node, "id")?,
            name: required(node, "name")?.to_owned(),
            initials: required(node, "initials")?.to_owned(),
            last_index: required_u32(node, "lastIdx")?,
            color_index: required_u32(node, "clrIdx")?,
        };
        if !ids.insert(author.id) {
            return Err(invalid(format!(
                "duplicate comment author ID {}",
                author.id
            )));
        }
        validate_string("comment author name", &author.name)?;
        validate_string("comment author initials", &author.initials)?;
        authors.push(author);
    }
    Ok(authors)
}

pub fn parse_slide_comments(xml: &[u8]) -> Result<Vec<PresentationComment>> {
    let root = parse_document(xml)?;
    let namespace = root_namespace(&root, "cmLst")?;
    no_attributes(&root, &[])?;
    whitespace_only(&root)?;
    if root.children.len() > MAX_COMMENTS_PER_SLIDE {
        return Err(invalid("slide comment count exceeds limit"));
    }
    let mut comments = Vec::with_capacity(root.children.len());
    let mut keys = HashSet::new();
    for node in &root.children {
        require_name(node, namespace, "cm")?;
        whitespace_only(node)?;
        no_attributes(node, &["authorId", "dt", "idx"])?;
        if !(node.children.len() == 2 || node.children.len() == 3) {
            return Err(invalid(
                "cm requires pos and text, followed optionally by extLst",
            ));
        }
        let position = &node.children[0];
        require_name(position, namespace, "pos")?;
        no_attributes(position, &["x", "y"])?;
        require_empty(position)?;
        let text = &node.children[1];
        require_name(text, namespace, "text")?;
        no_attributes(text, &["xml:space"])?;
        if !text.children.is_empty() {
            return Err(invalid("comment text cannot contain elements"));
        }
        if let Some(extension_list) = node.children.get(2) {
            require_name(extension_list, namespace, "extLst")?;
        }
        let date_time = optional(node, "dt").map(ToOwned::to_owned);
        if let Some(value) = &date_time {
            validate_date_time(value)?;
        }
        let comment = PresentationComment {
            author_id: required_u32(node, "authorId")?,
            date_time,
            index: required_u32(node, "idx")?,
            x: required_i64(position, "x")?,
            y: required_i64(position, "y")?,
            text: text.text.clone(),
        };
        if comment.index == 0 {
            return Err(invalid("comment index must be at least 1"));
        }
        validate_string("comment text", &comment.text)?;
        if !keys.insert((comment.author_id, comment.index)) {
            return Err(invalid("duplicate author/index comment key in slide"));
        }
        comments.push(comment);
    }
    Ok(comments)
}

pub fn write_comment_authors(
    authors: &[PresentationCommentAuthor],
    conformance: PresentationCommentConformance,
) -> Result<Vec<u8>> {
    validate_authors(authors)?;
    let mut output = declaration();
    output.extend_from_slice(b"<p:cmAuthorLst xmlns:p=\"");
    escape(&mut output, conformance.namespace());
    output.extend_from_slice(b"\">");
    for author in authors {
        output.extend_from_slice(b"<p:cmAuthor");
        attribute(&mut output, "id", &author.id.to_string());
        attribute(&mut output, "name", &author.name);
        attribute(&mut output, "initials", &author.initials);
        attribute(&mut output, "lastIdx", &author.last_index.to_string());
        attribute(&mut output, "clrIdx", &author.color_index.to_string());
        output.extend_from_slice(b"/>");
    }
    output.extend_from_slice(b"</p:cmAuthorLst>");
    Ok(output)
}

pub fn write_slide_comments(
    comments: &[PresentationComment],
    conformance: PresentationCommentConformance,
) -> Result<Vec<u8>> {
    validate_comment_list(comments)?;
    let mut output = declaration();
    output.extend_from_slice(b"<p:cmLst xmlns:p=\"");
    escape(&mut output, conformance.namespace());
    output.extend_from_slice(b"\">");
    for comment in comments {
        output.extend_from_slice(b"<p:cm");
        attribute(&mut output, "authorId", &comment.author_id.to_string());
        if let Some(date_time) = &comment.date_time {
            attribute(&mut output, "dt", date_time);
        }
        attribute(&mut output, "idx", &comment.index.to_string());
        output.extend_from_slice(b"><p:pos");
        attribute(&mut output, "x", &comment.x.to_string());
        attribute(&mut output, "y", &comment.y.to_string());
        output.extend_from_slice(b"/><p:text>");
        escape(&mut output, &comment.text);
        output.extend_from_slice(b"</p:text></p:cm>");
    }
    output.extend_from_slice(b"</p:cmLst>");
    Ok(output)
}

pub fn load_presentation_comments(package: &OpcPackage) -> Result<Option<PresentationComments>> {
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().to_string();
    let slide_names = collect_slides(package, presentation)?;
    validate_relationship_sources(package, &presentation_name, &slide_names)?;

    let author_relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| is_authors_relationship(relationship.reltype()))
        .collect();
    if author_relationships.len() > 1 {
        return Err(invalid(
            "presentation has multiple comment-author relationships",
        ));
    }

    let mut slide_lists = Vec::new();
    let mut comment_targets = HashSet::new();
    for slide_name in &slide_names {
        let slide_uri = PackURI::new(slide_name).map_err(invalid)?;
        let slide = package.get_part(&slide_uri)?;
        let relationships: Vec<_> = slide
            .rels()
            .iter()
            .filter(|relationship| is_comments_relationship(relationship.reltype()))
            .collect();
        if relationships.len() > 1 {
            return Err(invalid(format!(
                "slide '{slide_name}' has multiple comments relationships"
            )));
        }
        let Some(relationship) = relationships.first() else {
            continue;
        };
        if relationship.is_external() {
            return Err(invalid(format!(
                "slide comments relationship '{}' cannot be external",
                relationship.r_id()
            )));
        }
        let target = relationship.target_partname()?;
        if !comment_targets.insert(target.to_string()) {
            return Err(invalid(format!(
                "duplicate comments-part target '{target}'"
            )));
        }
        let part = package.get_part(&target)?;
        require_content_type(part, COMMENTS_CONTENT_TYPE)?;
        reject_outbound_relationships(part, "comments")?;
        slide_lists.push(SlideCommentList {
            slide_part_name: slide_name.clone(),
            relationship_id: relationship.r_id().to_owned(),
            part_name: target.to_string(),
            comments: parse_slide_comments(part.blob())?,
        });
    }

    let Some(author_relationship) = author_relationships.first() else {
        reject_orphan_parts(package, None, &comment_targets)?;
        if slide_lists.is_empty() {
            return Ok(None);
        }
        return Err(invalid(
            "slide comments exist without a comment-author part",
        ));
    };
    if author_relationship.is_external() {
        return Err(invalid("comment-author relationship cannot be external"));
    }
    let author_target = author_relationship.target_partname()?;
    let author_part = package.get_part(&author_target)?;
    require_content_type(author_part, AUTHORS_CONTENT_TYPE)?;
    reject_outbound_relationships(author_part, "comment-author")?;
    reject_orphan_parts(package, Some(author_target.as_str()), &comment_targets)?;
    let value = PresentationComments {
        author_relationship_id: author_relationship.r_id().to_owned(),
        author_part_name: author_target.to_string(),
        authors: parse_comment_authors(author_part.blob())?,
        slides: slide_lists,
    };
    validate_package_value(&value)?;
    Ok(Some(value))
}

pub fn store_presentation_comments(
    package: &mut OpcPackage,
    value: &PresentationComments,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    validate_package_value(value)?;
    if load_presentation_comments(package)?.is_some() {
        return Err(invalid("package already contains presentation comments"));
    }
    if value.author_relationship_id.is_empty() {
        return Err(invalid("comment-author relationship ID cannot be empty"));
    }
    let presentation = package.main_document_part()?;
    require_presentation_content_type(presentation.content_type())?;
    let presentation_name = presentation.partname().clone();
    if presentation
        .rels()
        .get(&value.author_relationship_id)
        .is_some()
    {
        return Err(invalid("comment-author relationship ID already exists"));
    }
    let slide_names = collect_slides(package, presentation)?;
    let slide_set: HashSet<_> = slide_names.iter().cloned().collect();
    let author_uri = PackURI::new(&value.author_part_name).map_err(invalid)?;
    let mut part_names = HashSet::new();
    part_names.insert(author_uri.to_string());
    let author_xml = write_comment_authors(&value.authors, conformance)?;
    let mut pending = Vec::with_capacity(value.slides.len());
    for slide_list in &value.slides {
        if !slide_set.contains(&slide_list.slide_part_name) {
            return Err(invalid(format!(
                "comment source '{}' is not a presentation slide",
                slide_list.slide_part_name
            )));
        }
        if slide_list.relationship_id.is_empty() {
            return Err(invalid("slide comments relationship ID cannot be empty"));
        }
        let slide_uri = PackURI::new(&slide_list.slide_part_name).map_err(invalid)?;
        let slide = package.get_part(&slide_uri)?;
        if slide.rels().get(&slide_list.relationship_id).is_some() {
            return Err(invalid(format!(
                "relationship '{}' already exists on slide",
                slide_list.relationship_id
            )));
        }
        let comment_uri = PackURI::new(&slide_list.part_name).map_err(invalid)?;
        if !part_names.insert(comment_uri.to_string()) {
            return Err(invalid(format!(
                "duplicate comment package part '{comment_uri}'"
            )));
        }
        pending.push((
            slide_uri,
            comment_uri,
            slide_list.relationship_id.clone(),
            write_slide_comments(&slide_list.comments, conformance)?,
        ));
    }
    for name in &part_names {
        let uri = PackURI::new(name).map_err(invalid)?;
        if package.iter_parts().any(|part| part.partname() == &uri) {
            return Err(invalid(format!(
                "comment package part '{uri}' already exists"
            )));
        }
    }

    package.add_part(Box::new(BlobPart::new(
        author_uri.clone(),
        AUTHORS_CONTENT_TYPE.into(),
        author_xml,
    )));
    let author_target = author_uri.relative_ref(presentation_name.base_uri());
    package
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .add_relationship(
            conformance.authors_relationship().into(),
            author_target,
            value.author_relationship_id.clone(),
            false,
        );
    for (slide_uri, comment_uri, relationship_id, xml) in pending {
        let target = comment_uri.relative_ref(slide_uri.base_uri());
        package.add_part(Box::new(BlobPart::new(
            comment_uri,
            COMMENTS_CONTENT_TYPE.into(),
            xml,
        )));
        package
            .get_part_mut(&slide_uri)?
            .rels_mut()
            .add_relationship(
                conformance.comments_relationship().into(),
                target,
                relationship_id,
                false,
            );
    }
    Ok(())
}

pub fn find_presentation_comment_author(
    package: &OpcPackage,
    author_id: u32,
) -> Result<Option<PresentationCommentAuthor>> {
    Ok(load_presentation_comments(package)?.and_then(|graph| {
        graph
            .authors
            .into_iter()
            .find(|author| author.id == author_id)
    }))
}

pub fn add_presentation_comment_author(
    package: &mut OpcPackage,
    author: PresentationCommentAuthor,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    let mut graph = if let Some(graph) = load_presentation_comments(package)? {
        graph
    } else {
        let presentation = package.main_document_part()?;
        let relationship_id = next_relationship_id(presentation, "rIdCommentAuthors")?;
        let part_name = next_legacy_author_part_name(package)?.to_string();
        let graph = PresentationComments {
            author_relationship_id: relationship_id,
            author_part_name: part_name,
            authors: vec![author],
            slides: Vec::new(),
        };
        store_presentation_comments(package, &graph, conformance)?;
        package.unsign();
        return Ok(());
    };
    if graph.authors.iter().any(|item| item.id == author.id) {
        return Err(invalid(format!(
            "comment author {} already exists",
            author.id
        )));
    }
    graph.authors.push(author);
    commit_legacy_authors(package, &graph, conformance)
}

pub fn update_presentation_comment_author(
    package: &mut OpcPackage,
    author_id: u32,
    author: PresentationCommentAuthor,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    if author.id != author_id {
        return Err(invalid("replacement author ID must remain stable"));
    }
    let mut graph = load_presentation_comments(package)?
        .ok_or_else(|| invalid("legacy comment graph does not exist"))?;
    let target = graph
        .authors
        .iter_mut()
        .find(|item| item.id == author_id)
        .ok_or_else(|| invalid(format!("comment author {author_id} was not found")))?;
    *target = author;
    commit_legacy_authors(package, &graph, conformance)
}

pub fn replace_presentation_comment_author(
    package: &mut OpcPackage,
    author_id: u32,
    author: PresentationCommentAuthor,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    update_presentation_comment_author(package, author_id, author, conformance)
}

pub fn remove_presentation_comment_author(
    package: &mut OpcPackage,
    author_id: u32,
    conformance: PresentationCommentConformance,
) -> Result<bool> {
    let Some(mut graph) = load_presentation_comments(package)? else {
        return Ok(false);
    };
    if graph.slides.iter().any(|slide| {
        slide
            .comments
            .iter()
            .any(|comment| comment.author_id == author_id)
    }) {
        return Err(invalid("cannot remove an author referenced by a comment"));
    }
    let Some(offset) = graph
        .authors
        .iter()
        .position(|author| author.id == author_id)
    else {
        return Ok(false);
    };
    graph.authors.remove(offset);
    commit_legacy_authors(package, &graph, conformance)?;
    Ok(true)
}

pub fn reorder_presentation_comment_authors(
    package: &mut OpcPackage,
    ordered_ids: &[u32],
    conformance: PresentationCommentConformance,
) -> Result<()> {
    let mut graph = load_presentation_comments(package)?
        .ok_or_else(|| invalid("legacy comment graph does not exist"))?;
    let expected = graph
        .authors
        .iter()
        .map(|author| author.id)
        .collect::<HashSet<_>>();
    let actual = ordered_ids.iter().copied().collect::<HashSet<_>>();
    if expected != actual || actual.len() != ordered_ids.len() {
        return Err(invalid("comment-author reorder is not a permutation"));
    }
    graph.authors = ordered_ids
        .iter()
        .map(|id| {
            graph
                .authors
                .iter()
                .find(|author| author.id == *id)
                .cloned()
                .ok_or_else(|| {
                    invalid(format!(
                        "comment-author reorder references missing author {id}"
                    ))
                })
        })
        .collect::<Result<Vec<_>>>()?;
    commit_legacy_authors(package, &graph, conformance)
}

pub fn find_presentation_comment(
    package: &OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
) -> Result<Option<PresentationComment>> {
    Ok(load_presentation_comments(package)?.and_then(|graph| {
        graph
            .slides
            .into_iter()
            .find(|slide| slide.slide_part_name == slide_part_name)
            .and_then(|slide| {
                slide
                    .comments
                    .into_iter()
                    .find(|comment| comment.author_id == author_id && comment.index == index)
            })
    }))
}

pub fn add_presentation_comment(
    package: &mut OpcPackage,
    slide_part_name: &str,
    comment: PresentationComment,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    let mut graph = load_presentation_comments(package)?
        .ok_or_else(|| invalid("legacy comment graph does not exist"))?;
    let author = graph
        .authors
        .iter_mut()
        .find(|author| author.id == comment.author_id)
        .ok_or_else(|| {
            invalid(format!(
                "comment references missing author {}",
                comment.author_id
            ))
        })?;
    if comment.index > author.last_index {
        author.last_index = comment.index;
    }
    if graph
        .slides
        .iter()
        .flat_map(|slide| &slide.comments)
        .any(|item| item.author_id == comment.author_id && item.index == comment.index)
    {
        return Err(invalid("comment author/index key already exists"));
    }
    if let Some(slide) = graph
        .slides
        .iter_mut()
        .find(|slide| slide.slide_part_name == slide_part_name)
    {
        slide.comments.push(comment);
        let xml = write_slide_comments(&slide.comments, conformance)?;
        parse_slide_comments(&xml)?;
        let uri = PackURI::new(&slide.part_name).map_err(invalid)?;
        package.get_part_mut(&uri)?.set_blob(xml);
    } else {
        let slide_uri = PackURI::new(slide_part_name).map_err(invalid)?;
        let slide = package.get_part(&slide_uri)?;
        require_content_type(slide, SLIDE_CONTENT_TYPE)?;
        let relationship_id = next_relationship_id(slide, "rIdComments")?;
        let part_uri = next_legacy_comment_part_name(package)?;
        let xml = write_slide_comments(std::slice::from_ref(&comment), conformance)?;
        package.try_add_part(Box::new(BlobPart::new(
            part_uri.clone(),
            COMMENTS_CONTENT_TYPE.into(),
            xml,
        )))?;
        package
            .get_part_mut(&slide_uri)?
            .rels_mut()
            .add_relationship(
                conformance.comments_relationship().into(),
                part_uri.relative_ref(slide_uri.base_uri()),
                relationship_id,
                false,
            );
    }
    commit_legacy_authors(package, &graph, conformance)?;
    package.unsign();
    Ok(())
}

pub fn update_presentation_comment(
    package: &mut OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
    replacement: PresentationComment,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    if replacement.author_id != author_id || replacement.index != index {
        return Err(invalid("replacement comment identity must remain stable"));
    }
    let graph = load_presentation_comments(package)?
        .ok_or_else(|| invalid("legacy comment graph does not exist"))?;
    let slide = graph
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_part_name)
        .ok_or_else(|| invalid("slide has no legacy comments"))?;
    let mut comments = slide.comments.clone();
    let target = comments
        .iter_mut()
        .find(|item| item.author_id == author_id && item.index == index)
        .ok_or_else(|| invalid("comment was not found"))?;
    *target = replacement;
    commit_legacy_slide_comments(package, slide, comments, conformance)
}

pub fn replace_presentation_comment(
    package: &mut OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
    replacement: PresentationComment,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    update_presentation_comment(
        package,
        slide_part_name,
        author_id,
        index,
        replacement,
        conformance,
    )
}

pub fn remove_presentation_comment(
    package: &mut OpcPackage,
    slide_part_name: &str,
    author_id: u32,
    index: u32,
    conformance: PresentationCommentConformance,
) -> Result<bool> {
    let Some(graph) = load_presentation_comments(package)? else {
        return Ok(false);
    };
    let Some(slide) = graph
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_part_name)
    else {
        return Ok(false);
    };
    let mut comments = slide.comments.clone();
    let Some(offset) = comments
        .iter()
        .position(|item| item.author_id == author_id && item.index == index)
    else {
        return Ok(false);
    };
    comments.remove(offset);
    if comments.is_empty() {
        let slide_uri = PackURI::new(slide_part_name).map_err(invalid)?;
        let part_uri = PackURI::new(&slide.part_name).map_err(invalid)?;
        package
            .get_part_mut(&slide_uri)?
            .rels_mut()
            .remove(&slide.relationship_id);
        if !part_is_referenced(package, &part_uri) {
            package.remove_part(&part_uri);
        }
        package.unsign();
    } else {
        commit_legacy_slide_comments(package, slide, comments, conformance)?;
    }
    Ok(true)
}

pub fn reorder_presentation_comments(
    package: &mut OpcPackage,
    slide_part_name: &str,
    ordered_keys: &[(u32, u32)],
    conformance: PresentationCommentConformance,
) -> Result<()> {
    let graph = load_presentation_comments(package)?
        .ok_or_else(|| invalid("legacy comment graph does not exist"))?;
    let slide = graph
        .slides
        .iter()
        .find(|slide| slide.slide_part_name == slide_part_name)
        .ok_or_else(|| invalid("slide has no legacy comments"))?;
    let expected = slide
        .comments
        .iter()
        .map(|item| (item.author_id, item.index))
        .collect::<HashSet<_>>();
    let actual = ordered_keys.iter().copied().collect::<HashSet<_>>();
    if expected != actual || actual.len() != ordered_keys.len() {
        return Err(invalid("comment reorder is not a permutation"));
    }
    let comments = ordered_keys
        .iter()
        .map(|key| {
            slide
                .comments
                .iter()
                .find(|item| (item.author_id, item.index) == *key)
                .cloned()
                .ok_or_else(|| {
                    invalid(format!(
                        "comment reorder references missing comment ({}, {})",
                        key.0, key.1
                    ))
                })
        })
        .collect::<Result<Vec<_>>>()?;
    commit_legacy_slide_comments(package, slide, comments, conformance)
}

fn commit_legacy_authors(
    package: &mut OpcPackage,
    graph: &PresentationComments,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    validate_package_value(graph)?;
    let xml = write_comment_authors(&graph.authors, conformance)?;
    parse_comment_authors(&xml)?;
    let uri = PackURI::new(&graph.author_part_name).map_err(invalid)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    package.unsign();
    Ok(())
}

fn commit_legacy_slide_comments(
    package: &mut OpcPackage,
    slide: &SlideCommentList,
    comments: Vec<PresentationComment>,
    conformance: PresentationCommentConformance,
) -> Result<()> {
    let xml = write_slide_comments(&comments, conformance)?;
    parse_slide_comments(&xml)?;
    let uri = PackURI::new(&slide.part_name).map_err(invalid)?;
    package.get_part_mut(&uri)?.set_blob(xml);
    package.unsign();
    Ok(())
}

fn next_relationship_id(part: &dyn Part, prefix: &str) -> Result<String> {
    for suffix in 1..=65_537u32 {
        let id = format!("{prefix}{suffix}");
        if part.rels().get(&id).is_none() {
            return Ok(id);
        }
    }
    Err(invalid("no free comment relationship ID"))
}

fn next_legacy_comment_part_name(package: &OpcPackage) -> Result<PackURI> {
    for suffix in 1..=100_001u32 {
        let uri = PackURI::new(format!("/ppt/comments/comment{suffix}.xml")).map_err(invalid)?;
        if package.get_part(&uri).is_err() {
            return Ok(uri);
        }
    }
    Err(invalid("no free legacy comment part name"))
}

fn next_legacy_author_part_name(package: &OpcPackage) -> Result<PackURI> {
    let canonical = PackURI::new("/ppt/commentAuthors.xml").map_err(invalid)?;
    if package.get_part(&canonical).is_err() {
        return Ok(canonical);
    }
    for suffix in 1..=65_537u32 {
        let candidate =
            PackURI::new(format!("/ppt/commentAuthors{suffix}.xml")).map_err(invalid)?;
        if package.get_part(&candidate).is_err() {
            return Ok(candidate);
        }
    }
    Err(invalid("no free legacy comment-author part name"))
}

fn part_is_referenced(package: &OpcPackage, target: &PackURI) -> bool {
    package.iter_parts().any(|part| {
        part.rels().iter().any(|relationship| {
            !relationship.is_external()
                && relationship
                    .target_partname()
                    .is_ok_and(|name| name == *target)
        })
    }) || package.rels().iter().any(|relationship| {
        !relationship.is_external()
            && relationship
                .target_partname()
                .is_ok_and(|name| name == *target)
    })
}

fn validate_authors(authors: &[PresentationCommentAuthor]) -> Result<()> {
    if authors.len() > MAX_AUTHORS {
        return Err(invalid("comment-author count exceeds limit"));
    }
    let mut ids = HashSet::new();
    for author in authors {
        validate_string("comment author name", &author.name)?;
        validate_string("comment author initials", &author.initials)?;
        if !ids.insert(author.id) {
            return Err(invalid(format!(
                "duplicate comment author ID {}",
                author.id
            )));
        }
    }
    Ok(())
}

fn validate_comment_list(comments: &[PresentationComment]) -> Result<()> {
    if comments.len() > MAX_COMMENTS_PER_SLIDE {
        return Err(invalid("slide comment count exceeds limit"));
    }
    let mut keys = HashSet::new();
    for comment in comments {
        if comment.index == 0 {
            return Err(invalid("comment index must be at least 1"));
        }
        validate_string("comment text", &comment.text)?;
        if let Some(date_time) = &comment.date_time {
            validate_date_time(date_time)?;
        }
        if !keys.insert((comment.author_id, comment.index)) {
            return Err(invalid("duplicate author/index comment key in slide"));
        }
    }
    Ok(())
}

fn validate_package_value(value: &PresentationComments) -> Result<()> {
    validate_authors(&value.authors)?;
    if value.slides.len() > MAX_SLIDES {
        return Err(invalid("commented-slide count exceeds limit"));
    }
    let authors: HashMap<_, _> = value
        .authors
        .iter()
        .map(|author| (author.id, author.last_index))
        .collect();
    let mut slide_names = HashSet::new();
    let mut part_names = HashSet::new();
    let mut keys = HashSet::new();
    let mut total = 0usize;
    for slide in &value.slides {
        if !slide_names.insert(&slide.slide_part_name) {
            return Err(invalid(format!(
                "duplicate commented slide '{}'",
                slide.slide_part_name
            )));
        }
        if !part_names.insert(&slide.part_name) {
            return Err(invalid(format!(
                "duplicate comments part '{}'",
                slide.part_name
            )));
        }
        validate_comment_list(&slide.comments)?;
        total = total
            .checked_add(slide.comments.len())
            .ok_or_else(|| invalid("comment count overflow"))?;
        if total > MAX_TOTAL_COMMENTS {
            return Err(invalid("total presentation comment count exceeds limit"));
        }
        for comment in &slide.comments {
            let Some(last_index) = authors.get(&comment.author_id) else {
                return Err(invalid(format!(
                    "comment references missing author {}",
                    comment.author_id
                )));
            };
            if comment.index > *last_index {
                return Err(invalid(format!(
                    "comment index {} exceeds author {} lastIdx",
                    comment.index, comment.author_id
                )));
            }
            if !keys.insert((comment.author_id, comment.index)) {
                return Err(invalid(
                    "duplicate author/index comment key across presentation",
                ));
            }
        }
    }
    Ok(())
}

fn collect_slides(package: &OpcPackage, presentation: &dyn Part) -> Result<Vec<String>> {
    let mut names = Vec::new();
    let mut seen = HashSet::new();
    for relationship in presentation
        .rels()
        .iter()
        .filter(|relationship| matches!(relationship.reltype(), SLIDE_REL | STRICT_SLIDE_REL))
    {
        if relationship.is_external() {
            return Err(invalid(
                "presentation slide relationship cannot be external",
            ));
        }
        let target = relationship.target_partname()?;
        if !seen.insert(target.to_string()) {
            return Err(invalid(format!(
                "presentation has duplicate slide target '{target}'"
            )));
        }
        let slide = package.get_part(&target)?;
        require_content_type(slide, SLIDE_CONTENT_TYPE)?;
        names.push(target.to_string());
        if names.len() > MAX_SLIDES {
            return Err(invalid("presentation slide count exceeds limit"));
        }
    }
    Ok(names)
}

fn validate_relationship_sources(
    package: &OpcPackage,
    presentation: &str,
    slides: &[String],
) -> Result<()> {
    if package.rels().iter().any(|relationship| {
        is_authors_relationship(relationship.reltype())
            || is_comments_relationship(relationship.reltype())
    }) {
        return Err(invalid(
            "package root cannot source presentation comment relationships",
        ));
    }
    let slide_set: HashSet<_> = slides.iter().map(String::as_str).collect();
    for part in package.iter_parts() {
        let source = part.partname().as_str();
        for relationship in part.rels().iter() {
            if is_authors_relationship(relationship.reltype()) && source != presentation {
                return Err(invalid(format!(
                    "comment-author relationship has invalid source '{source}'"
                )));
            }
            if is_comments_relationship(relationship.reltype()) && !slide_set.contains(source) {
                return Err(invalid(format!(
                    "comments relationship has invalid source '{source}'"
                )));
            }
        }
    }
    Ok(())
}

fn reject_orphan_parts(
    package: &OpcPackage,
    author_target: Option<&str>,
    comment_targets: &HashSet<String>,
) -> Result<()> {
    for part in package.iter_parts() {
        if part.content_type() == AUTHORS_CONTENT_TYPE
            && Some(part.partname().as_str()) != author_target
        {
            return Err(invalid(format!(
                "orphan comment-author part '{}'",
                part.partname()
            )));
        }
        if part.content_type() == COMMENTS_CONTENT_TYPE
            && !comment_targets.contains(part.partname().as_str())
        {
            return Err(invalid(format!(
                "orphan comments part '{}'",
                part.partname()
            )));
        }
    }
    Ok(())
}

fn parse_document(xml: &[u8]) -> Result<Node> {
    if xml.len() > MAX_PART_BYTES {
        return Err(invalid("presentation comment part is too large"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_PART_BYTES {
        return Err(invalid(
            "MCE-expanded presentation comment part is too large",
        ));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut stack = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    let mut strings = 0usize;
    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(xml_error)?;
        match event {
            Event::Start(element) => {
                if stack.len() >= MAX_DEPTH {
                    return Err(invalid("presentation comment XML depth exceeds limit"));
                }
                let node = make_node(&element, namespace, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation comment XML node count exceeds limit"));
                }
                stack.push(node);
            },
            Event::Empty(element) => {
                let node = make_node(&element, namespace, decoder, &mut strings)?;
                nodes += 1;
                if nodes > MAX_NODES {
                    return Err(invalid("presentation comment XML node count exceeds limit"));
                }
                attach(node, &mut stack, &mut root)?;
            },
            Event::End(element) => {
                let node = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected XML end element"))?;
                if node.name.as_bytes() != element.local_name().as_ref() {
                    return Err(invalid("mismatched XML end element"));
                }
                attach(node, &mut stack, &mut root)?;
            },
            Event::Text(text) => {
                let decoded = text.decode().map_err(xml_error)?;
                let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                strings = strings
                    .checked_add(decoded.len())
                    .ok_or_else(|| invalid("XML string size overflow"))?;
                if strings > MAX_STRING_BYTES {
                    return Err(invalid("presentation comment XML strings exceed limit"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&decoded);
                } else if !decoded.trim().is_empty() {
                    return Err(invalid("text outside XML root"));
                }
            },
            Event::CData(_) => {
                return Err(invalid("CDATA is rejected in presentation comment parts"));
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(xml_error)?;
                let value =
                    if let Some(character) = reference.resolve_char_ref().map_err(xml_error)? {
                        character.to_string()
                    } else {
                        match name.as_ref() {
                            "amp" => "&".to_owned(),
                            "lt" => "<".to_owned(),
                            "gt" => ">".to_owned(),
                            "apos" => "'".to_owned(),
                            "quot" => "\"".to_owned(),
                            _ => {
                                return Err(invalid(format!(
                                    "custom XML entity '&{name};' is rejected"
                                )));
                            },
                        }
                    };
                strings = strings
                    .checked_add(value.len())
                    .ok_or_else(|| invalid("XML string size overflow"))?;
                if strings > MAX_STRING_BYTES {
                    return Err(invalid("presentation comment XML strings exceed limit"));
                }
                if let Some(node) = stack.last_mut() {
                    node.text.push_str(&value);
                } else {
                    return Err(invalid("entity reference outside XML root"));
                }
            },
            Event::Decl(_) | Event::Comment(_) => {},
            Event::Eof => break,
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation comment XML"));
    }
    root.ok_or_else(|| invalid("missing presentation comment XML root"))
}

fn make_node(
    element: &BytesStart<'_>,
    namespace: ResolveResult<'_>,
    decoder: Decoder,
    strings: &mut usize,
) -> Result<Node> {
    let namespace = match namespace {
        ResolveResult::Bound(value) => std::str::from_utf8(value.as_ref())
            .map_err(xml_error)?
            .to_owned(),
        ResolveResult::Unbound => String::new(),
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unbound XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    let name = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    let mut attributes = Vec::new();
    let mut seen = HashSet::new();
    for attribute_value in element.attributes().with_checks(true) {
        let attribute_value = attribute_value.map_err(xml_error)?;
        let name = std::str::from_utf8(attribute_value.key.as_ref())
            .map_err(xml_error)?
            .to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate XML attribute"));
        }
        let value = attribute_value
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, decoder)
            .map_err(xml_error)?
            .into_owned();
        *strings = strings
            .checked_add(name.len() + value.len())
            .ok_or_else(|| invalid("XML string size overflow"))?;
        if *strings > MAX_STRING_BYTES {
            return Err(invalid("presentation comment XML strings exceed limit"));
        }
        if name != "xmlns" && !name.starts_with("xmlns:") {
            attributes.push(Attribute { name, value });
        }
    }
    Ok(Node {
        namespace,
        name,
        attributes,
        children: Vec::new(),
        text: String::new(),
    })
}

fn attach(node: Node, stack: &mut [Node], root: &mut Option<Node>) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(node);
    } else if root.replace(node).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

fn root_namespace<'a>(root: &'a Node, name: &str) -> Result<&'a str> {
    if root.name != name || !matches!(root.namespace.as_str(), PML | STRICT_PML) {
        return Err(invalid(format!("expected PresentationML {name} root")));
    }
    Ok(&root.namespace)
}

fn require_name(node: &Node, namespace: &str, name: &str) -> Result<()> {
    if node.namespace == namespace && node.name == name {
        Ok(())
    } else {
        Err(invalid(format!(
            "expected PresentationML {name}, got {}",
            node.name
        )))
    }
}

fn no_attributes(node: &Node, allowed: &[&str]) -> Result<()> {
    if let Some(attribute) = node
        .attributes
        .iter()
        .find(|attribute| !allowed.contains(&attribute.name.as_str()))
    {
        return Err(invalid(format!(
            "unexpected attribute '{}' on {}",
            attribute.name, node.name
        )));
    }
    Ok(())
}

fn required<'a>(node: &'a Node, name: &str) -> Result<&'a str> {
    optional(node, name)
        .ok_or_else(|| invalid(format!("{} requires attribute '{name}'", node.name)))
}

fn optional<'a>(node: &'a Node, name: &str) -> Option<&'a str> {
    node.attributes
        .iter()
        .find(|attribute| attribute.name == name)
        .map(|attribute| attribute.value.as_str())
}

fn required_u32(node: &Node, name: &str) -> Result<u32> {
    required(node, name)?.parse().map_err(|_| {
        invalid(format!(
            "invalid unsigned integer '{name}' on {}",
            node.name
        ))
    })
}

fn required_i64(node: &Node, name: &str) -> Result<i64> {
    required(node, name)?
        .parse()
        .map_err(|_| invalid(format!("invalid coordinate '{name}' on {}", node.name)))
}

fn whitespace_only(node: &Node) -> Result<()> {
    if node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("unexpected text in {}", node.name)))
    }
}

fn require_empty(node: &Node) -> Result<()> {
    if node.children.is_empty() && node.text.trim().is_empty() {
        Ok(())
    } else {
        Err(invalid(format!("{} must be empty", node.name)))
    }
}

fn validate_string(label: &str, value: &str) -> Result<()> {
    if value.len() <= MAX_STRING_BYTES {
        Ok(())
    } else {
        Err(invalid(format!("{label} exceeds size limit")))
    }
}

fn validate_date_time(value: &str) -> Result<()> {
    if DateTime::parse_from_rfc3339(value).is_ok()
        || NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f").is_ok()
    {
        Ok(())
    } else {
        Err(invalid(format!("invalid XML dateTime '{value}'")))
    }
}

fn declaration() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#.to_vec()
}

fn attribute(output: &mut Vec<u8>, name: &str, value: &str) {
    output.push(b' ');
    output.extend_from_slice(name.as_bytes());
    output.extend_from_slice(b"=\"");
    escape(output, value);
    output.push(b'"');
}

fn escape(output: &mut Vec<u8>, value: &str) {
    for character in value.chars() {
        match character {
            '&' => output.extend_from_slice(b"&amp;"),
            '<' => output.extend_from_slice(b"&lt;"),
            '>' => output.extend_from_slice(b"&gt;"),
            '"' => output.extend_from_slice(b"&quot;"),
            '\t' => output.extend_from_slice(b"&#x9;"),
            '\n' => output.extend_from_slice(b"&#xA;"),
            '\r' => output.extend_from_slice(b"&#xD;"),
            _ => {
                let mut bytes = [0; 4];
                output.extend_from_slice(character.encode_utf8(&mut bytes).as_bytes());
            },
        }
    }
}

fn require_presentation_content_type(value: &str) -> Result<()> {
    if matches!(
        value,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
            | "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.slideshow.macroEnabled.main+xml"
            | "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid(format!(
            "'{value}' is not a PresentationML main content type"
        )))
    }
}

fn require_content_type(part: &dyn Part, expected: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(Error::ContentType {
            expected: expected.into(),
            actual: part.content_type().into(),
        })
    }
}

fn reject_outbound_relationships(part: &dyn Part, label: &str) -> Result<()> {
    if part.rels().iter().next().is_none() {
        Ok(())
    } else {
        Err(invalid(format!(
            "{label} part '{}' must not have relationships",
            part.partname()
        )))
    }
}

fn is_comments_relationship(value: &str) -> bool {
    matches!(value, COMMENTS_REL | STRICT_COMMENTS_REL)
}
fn is_authors_relationship(value: &str) -> bool {
    matches!(value, AUTHORS_REL | STRICT_AUTHORS_REL)
}
fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(error.to_string())
}
fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const POI: &[u8] =
        include_bytes!("../../../test-data/poi/test-data/slideshow/45545_Comment.pptx");
    const LIBREOFFICE: &[u8] =
        include_bytes!("../../../test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf89064.pptx");

    fn author() -> PresentationCommentAuthor {
        PresentationCommentAuthor {
            id: 7,
            name: "A & B".into(),
            initials: "AB".into(),
            last_index: 1,
            color_index: 2,
        }
    }

    fn comment() -> PresentationComment {
        PresentationComment {
            author_id: 7,
            date_time: Some("2026-07-17T12:34:56.123+08:00".into()),
            index: 1,
            x: -10,
            y: 20,
            text: "review <this> & that".into(),
        }
    }

    fn package() -> OpcPackage {
        let mut package = OpcPackage::new();
        package.rels_mut().add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                .into(),
            "ppt/presentation.xml".into(),
            "rId1".into(),
            false,
        );
        let mut presentation = BlobPart::new(
            PackURI::new("/ppt/presentation.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml".into(),
            b"<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>".to_vec(),
        );
        presentation.rels_mut().add_relationship(
            SLIDE_REL.into(),
            "slides/slide1.xml".into(),
            "rId1".into(),
            false,
        );
        package.add_part(Box::new(presentation));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/slides/slide1.xml").unwrap(),
            SLIDE_CONTENT_TYPE.into(),
            Vec::new(),
        )));
        package
    }

    fn value() -> PresentationComments {
        PresentationComments {
            author_relationship_id: "rId2".into(),
            author_part_name: "/ppt/commentAuthors.xml".into(),
            authors: vec![author()],
            slides: vec![SlideCommentList {
                slide_part_name: "/ppt/slides/slide1.xml".into(),
                relationship_id: "rId2".into(),
                part_name: "/ppt/comments/comment1.xml".into(),
                comments: vec![comment()],
            }],
        }
    }

    #[test]
    fn loads_poi_and_libreoffice_reference_packages() {
        let poi = load_presentation_comments(&OpcPackage::from_bytes(POI).unwrap())
            .unwrap()
            .unwrap();
        assert_eq!(poi.authors.len(), 1);
        assert_eq!(poi.slides.len(), 2);
        assert_eq!(
            poi.slides
                .iter()
                .map(|slide| slide.comments.len())
                .sum::<usize>(),
            2
        );
        let libreoffice = load_presentation_comments(&OpcPackage::from_bytes(LIBREOFFICE).unwrap())
            .unwrap()
            .unwrap();
        assert_eq!(libreoffice.authors[0].name, "Anonymous");
        assert_eq!(libreoffice.slides[0].comments[0].text, "Comment");
    }

    #[test]
    fn strict_writers_are_deterministic_and_round_trip() {
        let authors = vec![author()];
        let comments = vec![comment()];
        let author_xml =
            write_comment_authors(&authors, PresentationCommentConformance::Strict).unwrap();
        let comment_xml =
            write_slide_comments(&comments, PresentationCommentConformance::Strict).unwrap();
        assert_eq!(
            author_xml,
            write_comment_authors(&authors, PresentationCommentConformance::Strict).unwrap()
        );
        assert_eq!(
            comment_xml,
            write_slide_comments(&comments, PresentationCommentConformance::Strict).unwrap()
        );
        assert!(
            std::str::from_utf8(&comment_xml)
                .unwrap()
                .contains(STRICT_PML)
        );
        assert_eq!(parse_comment_authors(&author_xml).unwrap(), authors);
        assert_eq!(parse_slide_comments(&comment_xml).unwrap(), comments);
    }

    #[test]
    fn mce_fallback_is_selected() {
        let xml = format!(
            r#"<p:cmLst xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:future" mc:Ignorable="x"><mc:AlternateContent><mc:Choice Requires="x"><x:cm/></mc:Choice><mc:Fallback><p:cm authorId="7" idx="1"><p:pos x="1" y="2"/><p:text>fallback</p:text></p:cm></mc:Fallback></mc:AlternateContent></p:cmLst>"#
        );
        assert_eq!(
            parse_slide_comments(xml.as_bytes()).unwrap()[0].text,
            "fallback"
        );
    }

    #[test]
    fn package_writer_round_trips_strict_graph() {
        let mut package = package();
        let expected = value();
        store_presentation_comments(
            &mut package,
            &expected,
            PresentationCommentConformance::Strict,
        )
        .unwrap();
        assert_eq!(
            load_presentation_comments(&package).unwrap().unwrap(),
            expected
        );
    }

    #[test]
    fn rejects_malformed_xml_and_enforces_caps() {
        for xml in [
            format!(
                r#"<p:cmAuthorLst xmlns:p="{PML}"><p:cmAuthor id="0" name="A" initials="A" lastIdx="1"/></p:cmAuthorLst>"#
            ),
            format!(
                r#"<p:cmLst xmlns:p="{PML}"><p:cm authorId="0" idx="0"><p:pos x="1" y="2"/><p:text>x</p:text></p:cm></p:cmLst>"#
            ),
            format!(
                r#"<p:cmLst xmlns:p="{PML}"><p:cm authorId="0" dt="not-a-date" idx="1"><p:pos x="1" y="2"/><p:text>x</p:text></p:cm></p:cmLst>"#
            ),
            format!(r#"<!DOCTYPE x><p:cmLst xmlns:p="{PML}"/>"#),
        ] {
            assert!(if xml.contains("cmAuthorLst") {
                parse_comment_authors(xml.as_bytes()).is_err()
            } else {
                parse_slide_comments(xml.as_bytes()).is_err()
            });
        }
        assert!(parse_slide_comments(&vec![b' '; MAX_PART_BYTES + 1]).is_err());
        let deep = format!(
            r#"<p:cmLst xmlns:p="{PML}">{}{} </p:cmLst>"#,
            "<p:extLst>".repeat(MAX_DEPTH),
            "</p:extLst>".repeat(MAX_DEPTH)
        );
        assert!(parse_slide_comments(deep.as_bytes()).is_err());
        let mut many = format!(r#"<p:cmAuthorLst xmlns:p="{PML}">"#);
        for id in 0..=MAX_AUTHORS {
            many.push_str(&format!(
                r#"<p:cmAuthor id="{id}" name="A" initials="A" lastIdx="0" clrIdx="0"/>"#
            ));
        }
        many.push_str("</p:cmAuthorLst>");
        assert!(parse_comment_authors(many.as_bytes()).is_err());
    }

    #[test]
    fn rejects_graph_and_reference_errors() {
        let mut missing_author = value();
        missing_author.slides[0].comments[0].author_id = 99;
        assert!(
            store_presentation_comments(
                &mut package(),
                &missing_author,
                PresentationCommentConformance::Transitional
            )
            .is_err()
        );

        let mut external = package();
        external
            .get_part_mut(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                COMMENTS_REL.into(),
                "https://invalid.example/comments.xml".into(),
                "rId2".into(),
                true,
            );
        assert!(load_presentation_comments(&external).is_err());

        let mut outbound = package();
        store_presentation_comments(
            &mut outbound,
            &value(),
            PresentationCommentConformance::Transitional,
        )
        .unwrap();
        outbound
            .get_part_mut(&PackURI::new("/ppt/comments/comment1.xml").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "other.xml".into(),
                "rId1".into(),
                false,
            );
        assert!(load_presentation_comments(&outbound).is_err());
    }
}
