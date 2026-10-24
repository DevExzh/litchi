//! Footnotes and endnotes staged with a main-document edit (record 0319,
//! L-DOCX-R5).
//!
//! A note lives in the footnotes or endnotes part and is referenced from the
//! main story by a reference run. Adding one stages the part's replacement
//! (or the new part and its relationship) as a graph change and the
//! reference run as an inline insertion, so one commit publishes both, as
//! comments do (0251c). Removal strips the note and every run holding its
//! reference.

use litchi_core::Position;

use super::super::{ElementPath, Error, Refusal, Result};
use super::Edit;
use super::stories::text_run;
use crate::Package;
use crate::part_scan::{self, Tree, qualified};

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";

/// Which notes part a note belongs to.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum NoteKind {
    /// A footnote (`w:footnote` in the footnotes part).
    Footnote,
    /// An endnote (`w:endnote` in the endnotes part).
    Endnote,
}

impl NoteKind {
    /// The note element's local name.
    #[must_use]
    pub const fn local_name(self) -> &'static str {
        match self {
            Self::Footnote => "footnote",
            Self::Endnote => "endnote",
        }
    }

    const fn part_root(self) -> &'static str {
        match self {
            Self::Footnote => "footnotes",
            Self::Endnote => "endnotes",
        }
    }

    const fn reference(self) -> &'static str {
        match self {
            Self::Footnote => "footnoteReference",
            Self::Endnote => "endnoteReference",
        }
    }

    const fn mark(self) -> &'static str {
        match self {
            Self::Footnote => "footnoteRef",
            Self::Endnote => "endnoteRef",
        }
    }

    const fn text_style(self) -> &'static str {
        match self {
            Self::Footnote => "FootnoteText",
            Self::Endnote => "EndnoteText",
        }
    }

    const fn reference_style(self) -> &'static str {
        match self {
            Self::Footnote => "FootnoteReference",
            Self::Endnote => "EndnoteReference",
        }
    }

    fn relationship(self, strict: bool) -> &'static str {
        match (self, strict) {
            (Self::Footnote, false) => litchi_opc::constants::relationship_type::FOOTNOTES,
            (Self::Endnote, false) => litchi_opc::constants::relationship_type::ENDNOTES,
            (Self::Footnote, true) => {
                "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes"
            },
            (Self::Endnote, true) => {
                "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes"
            },
        }
    }

    const fn content_type(self) -> &'static str {
        match self {
            Self::Footnote => litchi_opc::constants::content_type::WML_FOOTNOTES,
            Self::Endnote => litchi_opc::constants::content_type::WML_ENDNOTES,
        }
    }
}

/// A note to add: one paragraph of text after the note's reference mark.
///
/// ```
/// use litchi_docx::document::NewNote;
///
/// assert_eq!(NewNote::new("See the appendix.").text(), "See the appendix.");
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewNote {
    text: String,
}

impl NewNote {
    /// A note holding `text`; a line feed becomes a break, a tab a tab.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self { text: text.into() }
    }

    /// The text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

impl Edit<Package> {
    /// Add `note` of `kind`, referenced from the paragraph at `paragraph`
    /// before its direct run `at` (after its last child for `None`), and
    /// return the note's identifier. The note and its reference are written
    /// as OfficeCLI writes them: a `FootnoteText` paragraph that starts with
    /// a superscript reference mark, and a `FootnoteReference` run.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep, NewNote, NoteKind};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Claim");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let id = edit.add_note_at(
    ///     NoteKind::Footnote,
    ///     &ElementPath::new([ElementStep::Paragraph(Position::new(0))]),
    ///     None,
    ///     &NewNote::new("Source."),
    /// )?;
    /// let next = edit.commit()?.into_package();
    /// let notes = next.document()?.footnotes()?;
    /// assert_eq!(notes[0].id(), id);
    /// assert_eq!(notes[0].text()?.trim(), "Source.");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing paragraph or run, an insertion
    /// point inside a field, text XML cannot carry, or a malformed part.
    pub fn add_note_at(
        &mut self,
        kind: NoteKind,
        paragraph: &ElementPath,
        at: Option<Position>,
        note: &NewNote,
    ) -> Result<u32> {
        let strict = self.is_strict()?;
        let part = self.related_part(kind.relationship(strict))?;
        let current = match &part {
            Some((name, _)) => self
                .part_payload(name)?
                .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?
                .to_vec(),
            None => empty_part(kind, strict).into_bytes(),
        };
        let tree = part_scan::scan(&current, 1).map_err(Error::Document)?;
        let id = note_ids(&tree, kind)
            .into_iter()
            .max()
            .map_or(Ok(1), |maximum| {
                u32::try_from(maximum.max(0))
                    .ok()
                    .and_then(|maximum| maximum.checked_add(1))
                    .ok_or_else(invalid_value)
            })?;
        let root = tree.root().map_err(Error::Document)?;
        let prefix = root.prefix.to_vec();
        let markup = note_markup(&prefix, kind, id, &note.text)?;
        let updated = if root.empty {
            let mut replacement = current[root.start..root.tag_end - 2].to_vec();
            replacement.push(b'>');
            replacement.extend_from_slice(&markup);
            replacement.extend_from_slice(
                format!("</{}>", qualified(&prefix, kind.part_root())).as_bytes(),
            );
            part_scan::splice(&current, root.start, root.end, &replacement)
        } else {
            part_scan::splice(&current, root.close, root.close, &markup)
        };
        self.inner_mut()?.insert_run_markup_at(paragraph, at, |prefix| {
            let q = |local: &str| qualified(prefix, local);
            Ok(format!(
                r#"<{r}><{rpr}><{rstyle} {val}="{style}"/></{rpr}><{reference} {id_attribute}="{id}"/></{r}>"#,
                r = q("r"),
                rpr = q("rPr"),
                rstyle = q("rStyle"),
                val = q("val"),
                style = kind.reference_style(),
                reference = q(kind.reference()),
                id_attribute = q("id"),
            )
            .into_bytes())
        })?;
        match part {
            Some((name, _)) => {
                self.replace_part(&name, updated)?;
            },
            None => {
                self.stage_related_part(
                    kind.relationship(strict),
                    ("", kind.part_root(), ".xml"),
                    true,
                    kind.content_type(),
                    updated,
                )?;
            },
        }
        Ok(id)
    }

    /// Replace the text of note `id` of `kind`, as OfficeCLI does: its
    /// first run that is not the reference mark holds `text` (keeping its
    /// properties), and its other such runs are removed.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown note,
    /// [`Refusal::ComplexContent`] for a note holding a field, and
    /// [`Refusal::InvalidPropertyValue`] for text XML cannot carry.
    pub fn set_note_text(&mut self, kind: NoteKind, id: u32, text: &str) -> Result<&mut Self> {
        let (name, xml) = self.notes_part(kind)?;
        // Down to the runs' children, where the reference mark is.
        let tree = part_scan::scan(&xml, 4).map_err(Error::Document)?;
        let note = find_note(&tree, kind, id).ok_or_else(not_found)?;
        // A field is refused rather than kept or dropped (record 0304).
        let element = &tree.elements[note];
        if part_scan::holds_any(&xml[element.start..element.end], part_scan::FIELD_MARKUP)
            .map_err(Error::Document)?
        {
            return Err(Error::Refused {
                position: 0,
                reason: Refusal::ComplexContent,
            });
        }
        let mut edits: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        let mut first_content = true;
        let paragraphs: Vec<usize> = tree
            .children(note)
            .filter(|child| tree.elements[*child].is(b"p"))
            .collect();
        for paragraph in &paragraphs {
            for run in tree
                .children(*paragraph)
                .filter(|child| tree.elements[*child].is(b"r"))
                .collect::<Vec<_>>()
            {
                if tree.child(run, kind.mark().as_bytes()).is_some() {
                    continue;
                }
                let element = &tree.elements[run];
                if first_content {
                    first_content = false;
                    let properties = tree.child(run, b"rPr").map(|child| {
                        let rpr = &tree.elements[child];
                        xml[rpr.start..rpr.end].to_vec()
                    });
                    edits.push((
                        element.start,
                        element.end,
                        text_run(&element.prefix, properties.as_deref(), text)?,
                    ));
                } else {
                    edits.push((element.start, element.end, Vec::new()));
                }
            }
        }
        if first_content {
            let paragraph = *paragraphs.first().ok_or_else(not_found)?;
            let element = &tree.elements[paragraph];
            let markup = text_run(&element.prefix, None, text)?;
            if element.empty {
                let mut replacement = xml[element.start..element.tag_end - 2].to_vec();
                replacement.push(b'>');
                replacement.extend_from_slice(&markup);
                replacement.extend_from_slice(
                    format!("</{}>", qualified(&element.prefix, "p")).as_bytes(),
                );
                edits.push((element.start, element.end, replacement));
            } else {
                edits.push((element.close, element.close, markup));
            }
        }
        edits.sort_by_key(|(start, end, _)| (*start, *end != *start));
        let mut output = Vec::with_capacity(xml.len());
        let mut cursor = 0usize;
        for (start, end, bytes) in edits {
            output.extend_from_slice(&xml[cursor..start]);
            output.extend_from_slice(&bytes);
            cursor = end;
        }
        output.extend_from_slice(&xml[cursor..]);
        if output != xml.as_slice() {
            self.replace_part(&name, output)?;
        }
        Ok(self)
    }

    /// Remove note `id` of `kind` and every main-story run holding its
    /// reference.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for an unknown note.
    pub fn remove_note(&mut self, kind: NoteKind, id: u32) -> Result<&mut Self> {
        let (name, xml) = self.notes_part(kind)?;
        let tree = part_scan::scan(&xml, 1).map_err(Error::Document)?;
        let note = find_note(&tree, kind, id).ok_or_else(not_found)?;
        let element = &tree.elements[note];
        let updated = part_scan::splice(&xml, element.start, element.end, &[]);
        self.replace_part(&name, updated)?;
        self.inner_mut()?
            .strip_note_references(kind.reference().as_bytes(), i64::from(id))?;
        Ok(self)
    }

    pub(super) fn notes_part(&self, kind: NoteKind) -> Result<(String, Vec<u8>)> {
        let (name, _) = self
            .related_part(kind.relationship(self.is_strict()?))?
            .ok_or_else(not_found)?;
        let xml = self
            .part_payload(&name)?
            .ok_or_else(|| crate::Error::PartNotFound(name.clone()))?;
        Ok((name, xml.to_vec()))
    }
}

/// The identifiers of every note of `kind` in a notes part.
fn note_ids(tree: &Tree, kind: NoteKind) -> Vec<i64> {
    tree.children(0)
        .filter(|child| tree.elements[*child].is(kind.local_name().as_bytes()))
        .filter_map(|child| tree.elements[child].attribute(b"id")?.trim().parse().ok())
        .collect()
}

/// The element index of note `id` of `kind`; separators never match.
pub(super) fn find_note(tree: &Tree, kind: NoteKind, id: u32) -> Option<usize> {
    tree.children(0).find(|child| {
        let element = &tree.elements[*child];
        element.is(kind.local_name().as_bytes())
            && element
                .attribute(b"id")
                .and_then(|value| value.trim().parse::<i64>().ok())
                == Some(i64::from(id))
            && element
                .attribute(b"type")
                .is_none_or(|kind| kind == "normal")
    })
}

/// A notes part holding only Word's separator notes.
fn empty_part(kind: NoteKind, strict: bool) -> String {
    let namespace = if strict {
        STRICT_WORD_NAMESPACE
    } else {
        WORD_NAMESPACE
    };
    let root = kind.part_root();
    let note = kind.local_name();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{root} xmlns:w="{namespace}"><w:{note} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:{note}><w:{note} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{note}></w:{root}>"#
    )
}

/// The markup of note `id`.
fn note_markup(prefix: &[u8], kind: NoteKind, id: u32, text: &str) -> Result<Vec<u8>> {
    let q = |local: &str| qualified(prefix, local);
    let mut xml = format!(
        r#"<{note} {id_attribute}="{id}"><{p}><{ppr}><{pstyle} {val}="{style}"/></{ppr}><{r}><{rpr}><{valign} {val}="superscript"/></{rpr}><{mark}/></{r}>"#,
        note = q(kind.local_name()),
        id_attribute = q("id"),
        p = q("p"),
        ppr = q("pPr"),
        pstyle = q("pStyle"),
        val = q("val"),
        style = kind.text_style(),
        r = q("r"),
        rpr = q("rPr"),
        valign = q("vertAlign"),
        mark = q(kind.mark()),
    )
    .into_bytes();
    xml.extend_from_slice(&text_run(prefix, None, text)?);
    xml.extend_from_slice(format!("</{}></{}>", q("p"), q(kind.local_name())).as_bytes());
    Ok(xml)
}

const fn invalid_value() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::InvalidPropertyValue,
    }
}

const fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}
