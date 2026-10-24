//! Detached semantic projections for source-backed modern comments.
//!
//! Acquisition owns the bounded source graph walk and hands this module the
//! already processed modern-comment wire model.  This module owns the public
//! detached values: no source owner, package cache, or lazy part handle is
//! retained.  The only XML interpreted here is a `p188:txBody`, which is read
//! by the shared [`litchi_drawingml::text::body::reader`] with the namespace
//! declarations retained by the wire codec.
//!
//! Native identifiers, relationship metadata, and preserved wire fragments
//! are available only through [`raw`].  Ordinary diagnostics intentionally
//! omit authored text, names, identifiers, and XML.

use std::collections::HashMap;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_drawingml::text::body::reader as body_reader;
use quick_xml::name::{Namespace, NamespaceResolver, PrefixDeclaration};

use crate::error::{Error, Result};
use crate::modern_comments as codec;

use super::super::related::RelatedPart;

/// A detached modern-comment list selected from one slide.
#[derive(Clone)]
pub struct List {
    data: Arc<ListData>,
}

struct ListData {
    related: RelatedPart,
    author_related: Option<RelatedPart>,
    root_prefix: String,
    namespace_declarations: Box<[codec::NamespaceDeclaration]>,
    entries: Box<[Comment]>,
}

/// One detached semantic modern comment.
#[derive(Clone)]
pub struct Comment {
    data: Arc<CommentData>,
}

struct CommentData {
    id: String,
    author_id: String,
    author_display: Option<Arc<str>>,
    author_initials: Option<Arc<str>>,
    text: Option<String>,
    status: codec::Status,
    status_wire: Option<codec::Status>,
    created: String,
    start_date: Option<String>,
    due_date: Option<String>,
    assigned_to: Option<Box<[String]>>,
    progress: codec::Progress,
    progress_wire: Option<codec::Progress>,
    title: Option<String>,
    position: Option<(i64, i64)>,
    replies: Box<[Reply]>,
    reply_list_present: bool,
    namespace_declarations: Box<[codec::NamespaceDeclaration]>,
    reply_list_namespace_declarations: Box<[codec::NamespaceDeclaration]>,
    anchors: Box<[codec::Anchor]>,
    text_body_xml: Option<Vec<u8>>,
    extension_xml: Option<Vec<u8>>,
}

/// One detached semantic modern-comment reply.
#[derive(Clone)]
pub struct Reply {
    data: Arc<ReplyData>,
}

struct ReplyData {
    id: String,
    author_id: String,
    author_display: Option<Arc<str>>,
    author_initials: Option<Arc<str>>,
    text: Option<String>,
    status: codec::Status,
    status_wire: Option<codec::Status>,
    created: String,
    namespace_declarations: Box<[codec::NamespaceDeclaration]>,
    text_body_xml: Option<Vec<u8>>,
    extension_xml: Option<Vec<u8>>,
}

/// An author's display name and optional initials, by author identifier.
type AuthorIndex<'a> = HashMap<&'a str, (Arc<str>, Option<Arc<str>>)>;

impl fmt::Debug for List {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("List")
            .field("comment_count", &self.data.entries.len())
            .field("has_authors", &self.data.author_related.is_some())
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for Comment {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Comment")
            .field("has_text", &self.data.text.is_some())
            .field("reply_count", &self.data.replies.len())
            .field("has_author_display", &self.data.author_display.is_some())
            .field("status", &self.data.status)
            .field("progress", &self.data.progress)
            .finish_non_exhaustive()
    }
}

impl fmt::Debug for Reply {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Reply")
            .field("has_text", &self.data.text.is_some())
            .field("has_author_display", &self.data.author_display.is_some())
            .field("status", &self.data.status)
            .finish_non_exhaustive()
    }
}

impl List {
    /// Build a detached list from an acquired, already validated wire model.
    ///
    /// The list and every nested comment/reply retain only owned values and
    /// retained package bytes. DrawingML bodies are parsed before this method
    /// publishes the projection, so malformed or unbound body XML refuses the
    /// source read rather than becoming a deferred semantic failure.
    pub(super) fn from_parts(
        related: RelatedPart,
        parsed: codec::List,
        authors: Option<(RelatedPart, codec::Authors)>,
    ) -> Result<Self> {
        let codec::List {
            root_prefix,
            namespace_declarations,
            comments,
        } = parsed;
        let (author_related, authors) = match authors {
            Some((related, authors)) => (Some(related), Some(authors)),
            None => (None, None),
        };
        let author_index = author_index(authors.as_ref())?;

        let mut entries = Vec::new();
        reserve(
            &mut entries,
            comments.len(),
            "source modern comment projections",
        )?;
        for comment in comments {
            entries.push(project_comment(
                &root_prefix,
                &namespace_declarations,
                author_index.as_ref(),
                comment,
            )?);
        }

        Ok(Self {
            data: Arc::new(ListData {
                related,
                author_related,
                root_prefix,
                namespace_declarations: namespace_declarations.into_boxed_slice(),
                entries: entries.into_boxed_slice(),
            }),
        })
    }

    /// Comments in authored document order.
    #[must_use]
    pub fn entries(&self) -> &[Comment] {
        &self.data.entries
    }

    /// Return one comment at a checked zero-based position.
    #[must_use]
    pub fn comment(&self, position: Position) -> Option<&Comment> {
        self.data.entries.get(position.get())
    }

    /// Number of comments in this list.
    #[must_use]
    pub fn len(&self) -> usize {
        self.data.entries.len()
    }

    /// Whether the selected comments part contains no comments.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.data.entries.is_empty()
    }
}

impl Comment {
    /// The comment body text.
    ///
    /// `None` means the `p188:txBody` element was absent.  `Some("")` means
    /// that the body element was present but contained no text.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.data.text.as_deref()
    }

    /// The resolved modern-comment author display name, when the author part
    /// supplied one.
    #[must_use]
    pub fn author_name(&self) -> Option<&str> {
        self.data.author_display.as_deref()
    }

    /// The resolved author's initials, when the author part supplied them
    /// (record 0346).
    #[must_use]
    pub fn author_initials(&self) -> Option<&str> {
        self.data.author_initials.as_deref()
    }

    /// The authored creation timestamp.
    #[must_use]
    pub fn created(&self) -> &str {
        &self.data.created
    }

    /// The optional authored start timestamp.
    #[must_use]
    pub fn start_date(&self) -> Option<&str> {
        self.data.start_date.as_deref()
    }

    /// The optional authored due timestamp.
    #[must_use]
    pub fn due_date(&self) -> Option<&str> {
        self.data.due_date.as_deref()
    }

    /// The effective status. Omitted wire status follows the schema's active
    /// default; the omission itself is exposed through [`raw`].
    #[must_use]
    pub fn status(&self) -> codec::Status {
        self.data.status
    }

    /// The effective completion progress. Omitted wire progress follows the
    /// schema's zero default; the omission itself is exposed through [`raw`].
    #[must_use]
    pub fn progress(&self) -> codec::Progress {
        self.data.progress
    }

    /// The optional authored title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.data.title.as_deref()
    }

    /// The optional slide-coordinate position.
    #[must_use]
    pub fn location(&self) -> Option<(i64, i64)> {
        self.data.position
    }

    /// Replies in authored reply-list order.
    #[must_use]
    pub fn replies(&self) -> &[Reply] {
        &self.data.replies
    }

    /// Return one reply at a checked zero-based position.
    #[must_use]
    pub fn reply(&self, position: Position) -> Option<&Reply> {
        self.data.replies.get(position.get())
    }
}

impl Reply {
    /// The reply body text.
    ///
    /// `None` means the `p188:txBody` element was absent.  `Some("")` means
    /// that the body element was present but contained no text.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.data.text.as_deref()
    }

    /// The resolved modern-comment author display name, when available.
    #[must_use]
    pub fn author_name(&self) -> Option<&str> {
        self.data.author_display.as_deref()
    }

    /// The resolved author's initials, when available (record 0346).
    #[must_use]
    pub fn author_initials(&self) -> Option<&str> {
        self.data.author_initials.as_deref()
    }

    /// The authored creation timestamp.
    #[must_use]
    pub fn created(&self) -> &str {
        &self.data.created
    }

    /// The effective status. Omitted wire status follows the active default.
    #[must_use]
    pub fn status(&self) -> codec::Status {
        self.data.status
    }
}

fn project_comment<'a>(
    root_prefix: &str,
    root_declarations: &[codec::NamespaceDeclaration],
    author_index: Option<&AuthorIndex<'a>>,
    comment: codec::Comment,
) -> Result<Comment> {
    let codec::Comment {
        id,
        author_id,
        status,
        created,
        start_date,
        due_date,
        assigned_to,
        complete,
        title,
        namespace_declarations,
        anchors: comment_anchors,
        position,
        reply_list_namespace_declarations,
        replies: comment_replies,
        reply_list_present,
        text_body_xml,
        extension_xml,
    } = comment;
    let (author_display, author_initials) = author_display(author_index, &author_id);
    let text = body_text(
        text_body_xml.as_deref(),
        root_prefix,
        root_declarations,
        &namespace_declarations,
        &[],
        &[],
    )?;

    let mut replies = Vec::new();
    reserve(
        &mut replies,
        comment_replies.len(),
        "source modern comment replies",
    )?;
    for reply in comment_replies {
        replies.push(project_reply(
            root_prefix,
            root_declarations,
            &namespace_declarations,
            &reply_list_namespace_declarations,
            author_index,
            reply,
        )?);
    }

    let assigned_to = assigned_to.map(Vec::into_boxed_slice);
    let namespace_declarations = namespace_declarations.into_boxed_slice();
    let reply_list_namespace_declarations = reply_list_namespace_declarations.into_boxed_slice();
    let anchors = comment_anchors.into_boxed_slice();
    Ok(Comment {
        data: Arc::new(CommentData {
            id,
            author_id,
            author_display,
            author_initials,
            text,
            status: status.unwrap_or(codec::Status::Active),
            status_wire: status,
            created,
            start_date,
            due_date,
            assigned_to,
            progress: complete.unwrap_or(codec::Progress::ZERO),
            progress_wire: complete,
            title,
            position: position.map(|position| (position.x, position.y)),
            replies: replies.into_boxed_slice(),
            reply_list_present,
            namespace_declarations,
            reply_list_namespace_declarations,
            anchors,
            text_body_xml,
            extension_xml,
        }),
    })
}

fn project_reply<'a>(
    root_prefix: &str,
    root_declarations: &[codec::NamespaceDeclaration],
    comment_declarations: &[codec::NamespaceDeclaration],
    reply_list_declarations: &[codec::NamespaceDeclaration],
    author_index: Option<&AuthorIndex<'a>>,
    reply: codec::Reply,
) -> Result<Reply> {
    let codec::Reply {
        id,
        author_id,
        status,
        created,
        namespace_declarations,
        text_body_xml,
        extension_xml,
    } = reply;
    let (author_display, author_initials) = author_display(author_index, &author_id);
    let text = body_text(
        text_body_xml.as_deref(),
        root_prefix,
        root_declarations,
        comment_declarations,
        reply_list_declarations,
        &namespace_declarations,
    )?;
    Ok(Reply {
        data: Arc::new(ReplyData {
            id,
            author_id,
            author_display,
            author_initials,
            text,
            status: status.unwrap_or(codec::Status::Active),
            status_wire: status,
            created,
            namespace_declarations: namespace_declarations.into_boxed_slice(),
            text_body_xml,
            extension_xml,
        }),
    })
}

fn author_index<'a>(authors: Option<&'a codec::Authors>) -> Result<Option<AuthorIndex<'a>>> {
    let Some(authors) = authors else {
        return Ok(None);
    };
    let mut index = HashMap::new();
    index
        .try_reserve(authors.authors.len())
        .map_err(|source| Error::Allocation {
            resource: "source modern comment author index",
            source,
        })?;
    for author in &authors.authors {
        index.insert(
            author.id.as_str(),
            (
                Arc::<str>::from(author.name.as_str()),
                author.initials.as_deref().map(Arc::<str>::from),
            ),
        );
    }
    Ok(Some(index))
}

fn author_display(
    authors: Option<&AuthorIndex<'_>>,
    id: &str,
) -> (Option<Arc<str>>, Option<Arc<str>>) {
    authors
        .and_then(|authors| authors.get(id).cloned())
        .map_or((None, None), |(name, initials)| (Some(name), initials))
}

fn body_text(
    xml: Option<&[u8]>,
    root_prefix: &str,
    root_declarations: &[codec::NamespaceDeclaration],
    comment_declarations: &[codec::NamespaceDeclaration],
    reply_list_declarations: &[codec::NamespaceDeclaration],
    reply_declarations: &[codec::NamespaceDeclaration],
) -> Result<Option<String>> {
    let Some(xml) = xml else {
        return Ok(None);
    };
    let scope = namespace_scope(
        root_prefix,
        root_declarations,
        comment_declarations,
        reply_list_declarations,
        reply_declarations,
    )?;
    let body = body_reader::read(xml, scope, body_limits()).map_err(body_error)?;
    Ok(Some(body.text()))
}

fn namespace_scope(
    root_prefix: &str,
    root_declarations: &[codec::NamespaceDeclaration],
    comment_declarations: &[codec::NamespaceDeclaration],
    reply_list_declarations: &[codec::NamespaceDeclaration],
    reply_declarations: &[codec::NamespaceDeclaration],
) -> Result<NamespaceResolver> {
    let mut resolver = NamespaceResolver::default();
    add_binding(&mut resolver, root_prefix, codec::P188)?;
    add_declarations(&mut resolver, root_declarations)?;
    add_declarations(&mut resolver, comment_declarations)?;
    add_declarations(&mut resolver, reply_list_declarations)?;
    add_declarations(&mut resolver, reply_declarations)?;
    Ok(resolver)
}

fn add_declarations(
    resolver: &mut NamespaceResolver,
    declarations: &[codec::NamespaceDeclaration],
) -> Result<()> {
    for declaration in declarations {
        add_binding(resolver, &declaration.prefix, &declaration.uri)?;
    }
    Ok(())
}

fn add_binding(resolver: &mut NamespaceResolver, prefix: &str, uri: &str) -> Result<()> {
    let declaration = if prefix.is_empty() {
        PrefixDeclaration::Default
    } else {
        PrefixDeclaration::Named(prefix.as_bytes())
    };
    resolver
        .add(declaration, Namespace(uri.as_bytes()))
        .map_err(|_| Error::Invalid("modern comment namespace scope is invalid".into()))
}

fn body_limits() -> body_reader::Limits {
    body_reader::Limits::new(
        codec::MAX_BYTES,
        codec::MAX_DEPTH,
        codec::MAX_NODES,
        codec::MAX_STRING_BYTES,
    )
    .unwrap_or(body_reader::Limits::DEFAULT)
}

fn body_error(error: litchi_drawingml::Error) -> Error {
    match error {
        typed @ litchi_drawingml::Error::Limit { .. }
        | typed @ litchi_drawingml::Error::InvalidLimit { .. }
        | typed @ litchi_drawingml::Error::PatchConflict => Error::Drawing(typed),
        _ => Error::Invalid("modern comment DrawingML body is invalid".into()),
    }
}

#[cfg(test)]
mod tests {
    use super::body_error;

    #[test]
    fn drawingml_body_errors_keep_typed_budget_and_conflict_meanings() {
        let limit = body_error(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::TextBodyElements,
            observed: 9,
            maximum: 8,
        });
        assert!(matches!(
            limit,
            crate::Error::Drawing(litchi_drawingml::Error::Limit {
                kind: litchi_drawingml::LimitKind::TextBodyElements,
                observed: 9,
                maximum: 8,
            })
        ));

        let invalid = body_error(litchi_drawingml::Error::InvalidLimit {
            kind: litchi_drawingml::LimitKind::TextBodyElements,
            value: 0,
            maximum: 8,
        });
        assert!(matches!(
            invalid,
            crate::Error::Drawing(litchi_drawingml::Error::InvalidLimit {
                kind: litchi_drawingml::LimitKind::TextBodyElements,
                value: 0,
                maximum: 8,
            })
        ));

        assert!(matches!(
            body_error(litchi_drawingml::Error::PatchConflict),
            crate::Error::Drawing(litchi_drawingml::Error::PatchConflict)
        ));
    }
}

fn reserve<T>(values: &mut Vec<T>, count: usize, resource: &'static str) -> Result<()> {
    values
        .try_reserve_exact(count)
        .map_err(|source| Error::Allocation { resource, source })
}

/// Explicit raw access to native modern-comment metadata and retained wire
/// fragments.
pub mod raw {
    use super::{Comment, List, RelatedPart, Reply};
    use crate::modern_comments::{Anchor, NamespaceDeclaration, Status};
    use litchi_opc::PackURI;

    /// The package part name of the selected modern comments part.
    #[must_use]
    pub fn part_uri(list: &List) -> &PackURI {
        &list.data.related.partname
    }

    /// The selecting relationship ID of the modern comments part.
    #[must_use]
    pub fn relationship_id(list: &List) -> &str {
        &list.data.related.relationship_id
    }

    /// The original comments-part payload.
    #[must_use]
    pub fn data(list: &List) -> &[u8] {
        list.data.related.as_bytes()
    }

    /// The package part name of the modern comment author part, if acquired.
    #[must_use]
    pub fn author_part_uri(list: &List) -> Option<&PackURI> {
        list.data.author_related.as_ref().map(|part| &part.partname)
    }

    /// The selecting relationship ID of the modern comment author part.
    #[must_use]
    pub fn author_relationship_id(list: &List) -> Option<&str> {
        list.data
            .author_related
            .as_ref()
            .map(|part| part.relationship_id.as_ref())
    }

    /// The original author-part payload.
    #[must_use]
    pub fn author_data(list: &List) -> Option<&[u8]> {
        list.data.author_related.as_ref().map(RelatedPart::as_bytes)
    }

    /// The producer-supplied comment identifier.
    #[must_use]
    pub fn comment_id(comment: &Comment) -> &str {
        &comment.data.id
    }

    /// The prefix used by the modern-comment root element.
    #[must_use]
    pub fn root_prefix(list: &List) -> &str {
        &list.data.root_prefix
    }

    /// Namespace declarations written on the modern-comment root.
    #[must_use]
    pub fn namespace_declarations(list: &List) -> &[NamespaceDeclaration] {
        &list.data.namespace_declarations
    }

    /// The producer-supplied comment author identifier.
    #[must_use]
    pub fn comment_author_id(comment: &Comment) -> &str {
        &comment.data.author_id
    }

    /// The producer-supplied assignment identifiers, preserving omission
    /// separately from an explicit empty assignment attribute.
    #[must_use]
    pub fn assigned_to(comment: &Comment) -> Option<&[String]> {
        comment.data.assigned_to.as_deref()
    }

    /// The wire status, preserving omission separately from `active`.
    #[must_use]
    pub fn comment_status(comment: &Comment) -> Option<Status> {
        comment.data.status_wire
    }

    /// The wire completion value, preserving omission separately from zero.
    #[must_use]
    pub fn comment_progress(comment: &Comment) -> Option<crate::modern_comments::Progress> {
        comment.data.progress_wire
    }

    /// The retained complete `p188:txBody` fragment, if present.
    #[must_use]
    pub fn comment_text_body_xml(comment: &Comment) -> Option<&[u8]> {
        comment.data.text_body_xml.as_deref()
    }

    /// The retained complete `p188:extLst` fragment, if present.
    #[must_use]
    pub fn comment_extension_xml(comment: &Comment) -> Option<&[u8]> {
        comment.data.extension_xml.as_deref()
    }

    /// The retained anchor fragments in authored order.
    #[must_use]
    pub fn anchors(comment: &Comment) -> &[Anchor] {
        &comment.data.anchors
    }

    /// The retained namespace declarations written on the comment.
    #[must_use]
    pub fn comment_namespace_declarations(comment: &Comment) -> &[NamespaceDeclaration] {
        &comment.data.namespace_declarations
    }

    /// Whether the wire contained a `p188:replyLst` wrapper, including empty.
    #[must_use]
    pub fn reply_list_present(comment: &Comment) -> bool {
        comment.data.reply_list_present
    }

    /// The retained namespace declarations written on the reply list.
    #[must_use]
    pub fn reply_list_namespace_declarations(comment: &Comment) -> &[NamespaceDeclaration] {
        &comment.data.reply_list_namespace_declarations
    }

    /// The producer-supplied reply identifier.
    #[must_use]
    pub fn reply_id(reply: &Reply) -> &str {
        &reply.data.id
    }

    /// The producer-supplied reply author identifier.
    #[must_use]
    pub fn reply_author_id(reply: &Reply) -> &str {
        &reply.data.author_id
    }

    /// The wire reply status, preserving omission separately from `active`.
    #[must_use]
    pub fn reply_status(reply: &Reply) -> Option<Status> {
        reply.data.status_wire
    }

    /// Namespace declarations written on the reply.
    #[must_use]
    pub fn reply_namespace_declarations(reply: &Reply) -> &[NamespaceDeclaration] {
        &reply.data.namespace_declarations
    }

    /// The retained complete reply `p188:txBody` fragment, if present.
    #[must_use]
    pub fn reply_text_body_xml(reply: &Reply) -> Option<&[u8]> {
        reply.data.text_body_xml.as_deref()
    }

    /// The retained complete reply `p188:extLst` fragment, if present.
    #[must_use]
    pub fn reply_extension_xml(reply: &Reply) -> Option<&[u8]> {
        reply.data.extension_xml.as_deref()
    }
}
