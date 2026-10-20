//! Positional modern (threaded) comment authoring for the unified API
//! (record 0346): add a thread or a reply, and change a thread's or a
//! reply's text, author, creation stamp or resolution.
//!
//! Identities stay native: a new comment, reply or author takes a fresh
//! GUID derived from the edit's inputs and the package's existing
//! identities (no clock or random source is read), checked unique in the
//! comment graph. Authors are matched by name and initials, as `PowerPoint`
//! and OfficeCLI do; an author shared with other comments is never renamed
//! for one comment's change.

use litchi_core::Position;
use sha2::{Digest as _, Sha256};

use super::Edit;
use crate::modern_comments::{self as model, Graph, Status};
use crate::{Error, Result};

/// The changes [`Edit::set_modern_comment_at`] applies to a modern comment
/// or reply; `None` keeps a value.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ModernCommentChange<'a> {
    /// The plain text, replacing the body (a line feed starts a paragraph).
    pub text: Option<&'a str>,
    /// The author's display name.
    pub author: Option<&'a str>,
    /// The author's initials.
    pub initials: Option<&'a str>,
    /// The `xsd:dateTime` creation stamp.
    pub created: Option<&'a str>,
    /// The thread's resolution: `true` writes the `resolved` status, `false`
    /// removes the status (the schema default, active). A thread-level
    /// state, refused on a reply.
    pub resolved: Option<bool>,
}

impl Edit {
    /// Append a modern comment thread to a slide, by the author `author`
    /// with `initials` (matched by both; created when the presentation has
    /// none), with `text` and the `xsd:dateTime` stamp `created`, supplied
    /// by the caller (the editor consults no clock). `resolved` writes the
    /// `resolved` status. The thread is anchored at the slide origin.
    /// Returns its position in the slide's modern comment list.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown slide, invalid text or date, a
    /// signed package, or malformed modern comment parts. A failure leaves
    /// the edit unchanged.
    pub fn add_modern_comment_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        author: &str,
        initials: Option<&str>,
        text: &str,
        created: &str,
        resolved: bool,
    ) -> Result<Position> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let graph = model::load_modern_comment_graph(&candidate)?;
        let list = slide_comments(&graph, &selected.part_name);
        let existing = list.map_or(0, |list| list.comments.len());
        let prefix = list.map_or("p188", |list| list.root_prefix.as_str());
        let text_body = text_body(prefix, text)?;
        let author_id = author_for(&mut candidate, author, initials)?;
        let graph = model::load_modern_comment_graph(&candidate)?;
        let id = fresh_guid(
            &graph,
            &[
                b"comment",
                selected.part_name.as_str().as_bytes(),
                author.as_bytes(),
                text.as_bytes(),
                created.as_bytes(),
                &existing.to_le_bytes(),
            ],
        );
        let comment = model::Comment {
            id,
            author_id,
            status: resolved.then_some(Status::Resolved),
            created: created.to_owned(),
            start_date: None,
            due_date: None,
            assigned_to: None,
            complete: None,
            title: None,
            namespace_declarations: Vec::new(),
            anchors: Vec::new(),
            position: Some(model::Position { x: 0, y: 0 }),
            reply_list_namespace_declarations: Vec::new(),
            replies: Vec::new(),
            reply_list_present: false,
            text_body_xml: Some(text_body),
            extension_xml: None,
        };
        model::add_modern_comment(&mut candidate, &selected.part_name, comment)?;
        model::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(Position::new(existing))
    }

    /// Append a reply to the modern comment thread at `comment` of a slide,
    /// by the author `author` with `initials` (matched by both; created when
    /// absent), with `text` and the caller's `xsd:dateTime` stamp `created`.
    /// Returns the reply's position in the thread.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing thread, and the
    /// errors of [`Self::add_modern_comment_text`]. A failure leaves the
    /// edit unchanged.
    pub fn add_modern_comment_reply_text<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: Position,
        author: &str,
        initials: Option<&str>,
        text: &str,
        created: &str,
    ) -> Result<Position> {
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let graph = model::load_modern_comment_graph(&candidate)?;
        let list = slide_comments(&graph, &selected.part_name);
        let comments = list.map_or(&[][..], |list| list.comments.as_slice());
        let thread = comments.get(comment.get()).ok_or(Error::IndexOutOfBounds {
            index: comment.get(),
            len: comments.len(),
        })?;
        let (thread_id, existing) = (thread.id.clone(), thread.replies.len());
        let prefix = list.map_or("p188", |list| list.root_prefix.as_str());
        let text_body = text_body(prefix, text)?;
        let author_id = author_for(&mut candidate, author, initials)?;
        let graph = model::load_modern_comment_graph(&candidate)?;
        let id = fresh_guid(
            &graph,
            &[
                b"reply",
                thread_id.as_bytes(),
                author.as_bytes(),
                text.as_bytes(),
                created.as_bytes(),
                &existing.to_le_bytes(),
            ],
        );
        let reply = model::Reply {
            id,
            author_id,
            status: None,
            created: created.to_owned(),
            namespace_declarations: Vec::new(),
            text_body_xml: Some(text_body),
            extension_xml: None,
        };
        model::add_modern_comment_reply(&mut candidate, &selected.part_name, &thread_id, reply)?;
        model::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(Position::new(existing))
    }

    /// Change the modern comment thread at `comment` of a slide, or its reply
    /// at `reply`. A changed author or initials renames the author when no
    /// other comment or reply of the presentation is theirs; otherwise the
    /// target moves to the author of that name and initials (created when
    /// absent), so other comments keep their attribution. A new text
    /// replaces the body with plain paragraphs. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns [`Error::IndexOutOfBounds`] for a missing thread or reply,
    /// an error for `resolved` on a reply, invalid text or date, a signed
    /// package, or malformed modern comment parts. A failure leaves the edit
    /// unchanged.
    pub fn set_modern_comment_at<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        comment: Position,
        reply: Option<Position>,
        change: &ModernCommentChange<'_>,
    ) -> Result<bool> {
        if reply.is_some() && change.resolved.is_some() {
            return Err(Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "set_modern_comment_at",
                "resolution is a thread-level state; a reply has none",
            ));
        }
        let selected = self.resolve_slide(slide.into())?;
        let mut candidate = self.working.clone();
        let graph = model::load_modern_comment_graph(&candidate)?;
        let list = slide_comments(&graph, &selected.part_name);
        let comments = list.map_or(&[][..], |list| list.comments.as_slice());
        let thread = comments.get(comment.get()).ok_or(Error::IndexOutOfBounds {
            index: comment.get(),
            len: comments.len(),
        })?;
        let target = match reply {
            Some(position) => Some(thread.replies.get(position.get()).ok_or(
                Error::IndexOutOfBounds {
                    index: position.get(),
                    len: thread.replies.len(),
                },
            )?),
            None => None,
        };
        let thread_id = thread.id.clone();
        let (target_id, current_author) = match target {
            Some(reply) => (reply.id.clone(), reply.author_id.clone()),
            None => (thread.id.clone(), thread.author_id.clone()),
        };
        let prefix = list.map_or("p188", |list| list.root_prefix.as_str());
        let text_body = change
            .text
            .map(|text| text_body(prefix, text))
            .transpose()?;
        let mut changed = false;

        // Attribution first: it may add an author, which the comment then
        // names.
        let mut author_id = current_author.clone();
        if change.author.is_some() || change.initials.is_some() {
            let author = graph
                .authors
                .as_ref()
                .and_then(|part| {
                    part.authors
                        .authors
                        .iter()
                        .find(|author| author.id == current_author)
                })
                .cloned()
                .ok_or_else(|| Error::Invalid("the modern comment names no known author".into()))?;
            let name = change.author.unwrap_or(&author.name);
            let initials = change.initials.or(author.initials.as_deref());
            if name != author.name || initials != author.initials.as_deref() {
                if references(&graph, &current_author, &target_id) == 0 {
                    let (name, initials) = (name.to_owned(), initials.map(str::to_owned));
                    changed |= model::update_modern_comment_author(
                        &mut candidate,
                        &current_author,
                        |a| {
                            a.name = name;
                            a.initials = initials;
                        },
                    )?;
                } else {
                    author_id = author_for(&mut candidate, name, initials)?;
                }
            }
        }

        let created = change.created;
        let resolved = change.resolved;
        let new_author = (author_id != current_author).then_some(author_id);
        let mut target_changed = false;
        match reply {
            None => {
                let current = thread.clone();
                let mut replacement = current.clone();
                if let Some(body) = &text_body {
                    replacement.text_body_xml = Some(body.clone());
                }
                if let Some(author) = &new_author {
                    replacement.author_id.clone_from(author);
                }
                if let Some(created) = created {
                    created.clone_into(&mut replacement.created);
                }
                if let Some(resolved) = resolved {
                    replacement.status = resolved.then_some(Status::Resolved);
                }
                if replacement != current {
                    target_changed = model::update_modern_comment(
                        &mut candidate,
                        &selected.part_name,
                        &thread_id,
                        move |comment| *comment = replacement,
                    )?;
                }
            },
            Some(_) => {
                let Some(current) = target.cloned() else {
                    unreachable!("a reply target was resolved above")
                };
                let mut replacement = current.clone();
                if let Some(body) = &text_body {
                    replacement.text_body_xml = Some(body.clone());
                }
                if let Some(author) = &new_author {
                    replacement.author_id.clone_from(author);
                }
                if let Some(created) = created {
                    created.clone_into(&mut replacement.created);
                }
                if replacement != current {
                    target_changed = model::update_modern_comment_reply(
                        &mut candidate,
                        &selected.part_name,
                        &thread_id,
                        &target_id,
                        move |reply| *reply = replacement,
                    )?;
                }
            },
        }
        changed |= target_changed;
        if !changed {
            return Ok(false);
        }
        model::load_modern_comment_graph(&candidate)?;
        self.working = candidate;
        Ok(true)
    }
}

/// The modern comment list of one slide part.
fn slide_comments<'g>(graph: &'g Graph, slide: &litchi_opc::PackURI) -> Option<&'g model::List> {
    graph
        .comments
        .iter()
        .find(|part| part.slide_part_name == slide.as_str())
        .map(|part| &part.comments)
}

/// How many comments and replies other than `exclude` name `author`.
fn references(graph: &Graph, author: &str, exclude: &str) -> usize {
    graph
        .comments
        .iter()
        .flat_map(|part| &part.comments.comments)
        .map(|comment| {
            usize::from(comment.author_id == author && comment.id != exclude)
                + comment
                    .replies
                    .iter()
                    .filter(|reply| reply.author_id == author && reply.id != exclude)
                    .count()
        })
        .sum()
}

/// The identifier of the author named `name` with `initials`, added to the
/// presentation's modern comment authors when it has none.
fn author_for(
    package: &mut litchi_opc::OpcPackage,
    name: &str,
    initials: Option<&str>,
) -> Result<String> {
    let graph = model::load_modern_comment_graph(package)?;
    if let Some(author) = graph.authors.as_ref().and_then(|part| {
        part.authors
            .authors
            .iter()
            .find(|author| author.name == name && author.initials.as_deref() == initials)
    }) {
        return Ok(author.id.clone());
    }
    let id = fresh_guid(
        &graph,
        &[
            b"author",
            name.as_bytes(),
            initials.unwrap_or("").as_bytes(),
        ],
    );
    let author = model::Author {
        id: id.clone(),
        name: name.to_owned(),
        initials: initials.map(str::to_owned),
        // A local account, as `PowerPoint` writes one without a sign-in.
        user_id: name.to_owned(),
        provider_id: "None".to_owned(),
        namespace_declarations: Vec::new(),
        extension_xml: None,
    };
    model::add_modern_comment_author(package, author)?;
    Ok(id)
}

/// A GUID in braces derived from `seed`, unique among the graph's comment,
/// reply and author identifiers: the editor reads no clock or random
/// source, so equal edits of equal packages produce equal bytes.
fn fresh_guid(graph: &Graph, seed: &[&[u8]]) -> String {
    let taken = |candidate: &str| {
        graph.authors.as_ref().is_some_and(|part| {
            part.authors
                .authors
                .iter()
                .any(|author| author.id.eq_ignore_ascii_case(candidate))
        }) || graph.comments.iter().any(|part| {
            part.comments.comments.iter().any(|comment| {
                comment.id.eq_ignore_ascii_case(candidate)
                    || comment
                        .replies
                        .iter()
                        .any(|reply| reply.id.eq_ignore_ascii_case(candidate))
            })
        })
    };
    let mut attempt = 0_u64;
    loop {
        let mut hasher = Sha256::new();
        for part in seed {
            hasher.update(part.len().to_le_bytes());
            hasher.update(part);
        }
        hasher.update(attempt.to_le_bytes());
        let digest = hasher.finalize();
        let mut bytes = [0_u8; 16];
        bytes.copy_from_slice(&digest[..16]);
        // An RFC 4122 version-4 layout, as `PowerPoint`'s identifiers have.
        bytes[6] = (bytes[6] & 0x0F) | 0x40;
        bytes[8] = (bytes[8] & 0x3F) | 0x80;
        let hex = |range: std::ops::Range<usize>| {
            bytes[range]
                .iter()
                .map(|byte| format!("{byte:02X}"))
                .collect::<String>()
        };
        let candidate = format!(
            "{{{}-{}-{}-{}-{}}}",
            hex(0..4),
            hex(4..6),
            hex(6..8),
            hex(8..10),
            hex(10..16)
        );
        if !taken(&candidate) {
            return candidate;
        }
        attempt = attempt.wrapping_add(1);
    }
}

/// A `txBody` fragment in the list's `p188` prefix holding `text` as plain
/// paragraphs (one per line), declaring the `DrawingML` prefix it uses.
fn text_body(prefix: &str, text: &str) -> Result<Vec<u8>> {
    if !text.chars().all(super::is_xml_char) {
        return Err(Error::Invalid(
            "modern comment text contains an invalid XML character".into(),
        ));
    }
    let tag = if prefix.is_empty() {
        "txBody".to_owned()
    } else {
        format!("{prefix}:txBody")
    };
    let mut xml = format!(
        "<{tag} xmlns:a=\"{}\"><a:bodyPr/><a:lstStyle/>",
        crate::modern_comments::A
    );
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        if line.is_empty() {
            xml.push_str("<a:p/>");
        } else {
            xml.push_str("<a:p><a:r><a:t>");
            xml.push_str(&quick_xml::escape::escape(line));
            xml.push_str("</a:t></a:r></a:p>");
        }
    }
    xml.push_str("</");
    xml.push_str(&tag);
    xml.push('>');
    Ok(xml.into_bytes())
}
