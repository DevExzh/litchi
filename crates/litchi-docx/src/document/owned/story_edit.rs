//! Paragraph and run edits of a header, footer or note story with the
//! main-document transaction verbs (record 0325).
//!
//! [`Edit::edit_story`] wraps the story's content in a synthetic
//! `w:document`/`w:body` (keeping the story part root's namespace
//! declarations), runs the caller's verbs on a [`Snapshot`] of it — the
//! same element paths, formatting, text, insertion and removal verbs the
//! main document takes — and splices the edited content back between the
//! story's own tags. Every byte of the part outside the story's content is
//! kept, and an unchanged story stages nothing.

use super::super::raw::Tree as SnapshotEdit;
use super::super::{Error, Refusal, Result, Snapshot};
use super::Edit;
use super::notes::{NoteKind, find_note};
use crate::Package;
use crate::part_scan::{self, qualified};

/// The story an [`Edit::edit_story`] call edits.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum StoryTarget<'a> {
    /// A header or footer part, by part name.
    Part(&'a str),
    /// A footnote or endnote, by identifier.
    Note(NoteKind, u32),
}

fn refused(reason: Refusal) -> Error {
    Error::Refused {
        position: 0,
        reason,
    }
}

impl Edit<Package> {
    /// Edit the paragraphs and runs of a header, footer or note story with
    /// the main-document verbs: `edit` receives a transaction over the
    /// story's content as a document body, whose element paths start at the
    /// story's blocks. The edited content replaces the story's; an
    /// unchanged story stages nothing.
    ///
    /// ```
    /// use litchi_docx::document::{
    ///     ElementAnchor, ElementPath, NewParagraph, NewRun, NewStory, StoryTarget,
    /// };
    /// use litchi_docx::header_footer::{Kind, Role};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Body");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let name = edit.add_header_footer(Role::Header, Kind::Primary, &NewStory::new("Top"))?;
    /// edit.edit_story(StoryTarget::Part(&name), |story| {
    ///     story.insert_paragraph_at(
    ///         &ElementPath::new(Vec::new()),
    ///         ElementAnchor::End,
    ///         &NewParagraph::new().with_run(NewRun::new("Second")),
    ///     )
    /// })?;
    /// let package = edit.commit()?.into_package();
    /// let xml = package.part_xml(&name)?.unwrap();
    /// assert!(String::from_utf8_lossy(&xml).contains("Second"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing part or note,
    /// [`Refusal::ComplexContent`] for an empty story element or a note
    /// that declares its own namespaces, `edit`'s error, and an error when
    /// the edited content does not splice back. The edit is unchanged on
    /// failure.
    pub fn edit_story<T, E: From<Error>>(
        &mut self,
        story: StoryTarget<'_>,
        edit: impl FnOnce(&mut SnapshotEdit) -> std::result::Result<T, E>,
    ) -> std::result::Result<T, E> {
        let (name, xml) = match story {
            StoryTarget::Part(name) => (
                name.to_owned(),
                self.part_payload(name)
                    .map_err(E::from)?
                    .ok_or_else(|| refused(Refusal::BlockNotFound))?
                    .to_vec(),
            ),
            StoryTarget::Note(kind, _) => self.notes_part(kind).map_err(E::from)?,
        };
        let (prefix, inner, suffix, open_end, close) = wrapper(&xml, story).map_err(E::from)?;
        let mut synthetic = Vec::with_capacity(prefix.len() + inner.len() + suffix.len());
        synthetic.extend_from_slice(&prefix);
        synthetic.extend_from_slice(inner);
        synthetic.extend_from_slice(&suffix);
        let snapshot = Snapshot::from_xml(synthetic).map_err(E::from)?;
        let mut staged = snapshot.edit();
        let output = edit(&mut staged)?;
        let projected = staged.projected().xml_bytes();
        if projected.len() < prefix.len() + suffix.len()
            || !projected.starts_with(&prefix)
            || !projected.ends_with(&suffix)
        {
            return Err(E::from(Error::from(crate::Error::InvalidFormat(
                "a story edit changed its synthetic wrapper".into(),
            ))));
        }
        let edited = &projected[prefix.len()..projected.len() - suffix.len()];
        if edited == inner {
            return Ok(output);
        }
        let mut updated = Vec::with_capacity(xml.len() + edited.len());
        updated.extend_from_slice(&xml[..open_end]);
        updated.extend_from_slice(edited);
        updated.extend_from_slice(&xml[close..]);
        self.replace_part(&name, updated).map_err(E::from)?;
        Ok(output)
    }
}

/// The synthetic document's opening and closing markup, the story's
/// content, and the content's bounds in the part.
#[allow(
    clippy::type_complexity,
    reason = "a private five-part split of one part payload"
)]
fn wrapper<'x>(
    xml: &'x [u8],
    story: StoryTarget<'_>,
) -> Result<(Vec<u8>, &'x [u8], Vec<u8>, usize, usize)> {
    let tree = part_scan::scan(xml, 1).map_err(Error::Document)?;
    let root = tree.root().map_err(Error::Document)?;
    let container = match story {
        StoryTarget::Part(_) => {
            if !(root.is(b"hdr") || root.is(b"ftr")) {
                return Err(refused(Refusal::BlockNotFound));
            }
            root
        },
        StoryTarget::Note(kind, id) => {
            let note = find_note(&tree, kind, id).ok_or_else(|| refused(Refusal::BlockNotFound))?;
            let element = &tree.elements[note];
            if !element.declarations.is_empty() {
                return Err(refused(Refusal::ComplexContent));
            }
            element
        },
    };
    if container.empty || root.empty {
        return Err(refused(Refusal::ComplexContent));
    }
    // The part root's start tag after its name: attributes and namespace
    // declarations, through `>`.
    let root_tag = &xml[root.start..root.tag_end];
    let name_end =
        1 + qualified(&root.prefix, std::str::from_utf8(&root.local).unwrap_or("")).len();
    let attributes = root_tag
        .get(name_end..)
        .ok_or_else(|| refused(Refusal::ComplexContent))?;
    let document = qualified(&root.prefix, "document");
    let body = qualified(&root.prefix, "body");
    let mut prefix = format!("<{document}").into_bytes();
    prefix.extend_from_slice(attributes);
    prefix.extend_from_slice(format!("<{body}>").as_bytes());
    let suffix = format!("</{body}></{document}>").into_bytes();
    let content = &xml[container.tag_end..container.close];
    Ok((prefix, content, suffix, container.tag_end, container.close))
}
