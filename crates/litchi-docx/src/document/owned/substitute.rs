//! Story text substitution on the package-bound edit (records 0274, 0285).

use std::sync::Arc;

use super::super::Result;
use super::super::substitute::{Outcome, Replacement, Story, StoryKind, StoryPart};
use super::super::transaction::substitute as story;
use super::Edit;
use crate::Package;
use crate::properties;

/// Relationship types (by their last segment) of the story parts the main
/// document part relates: headers, footers, notes and comments.
const STORY_RELATIONSHIPS: [(&str, StoryKind); 5] = [
    ("/header", StoryKind::Header),
    ("/footer", StoryKind::Footer),
    ("/footnotes", StoryKind::Footnotes),
    ("/endnotes", StoryKind::Endnotes),
    ("/comments", StoryKind::Comments),
];

/// The story parts a main document's relationships name, in order, each
/// once.
pub(crate) fn story_parts_of(rels: &litchi_opc::Relationships) -> crate::Result<Vec<StoryPart>> {
    let mut parts: Vec<StoryPart> = Vec::new();
    for relationship in rels.iter() {
        if relationship.is_external() {
            continue;
        }
        let Some(kind) = STORY_RELATIONSHIPS
            .iter()
            .find(|(suffix, _)| relationship.reltype().ends_with(suffix))
            .map(|(_, kind)| *kind)
        else {
            continue;
        };
        let name = relationship.target_partname()?.as_str().to_owned();
        if !parts.iter().any(|part| part.name() == name) {
            parts.push(StoryPart::new(name, kind));
        }
    }
    Ok(parts)
}

impl<P: Clone> Edit<P> {
    /// The main document's story over the staged document: the text of
    /// every paragraph in its body paragraphs, tables and block content
    /// controls.
    ///
    /// ```
    /// use litchi_docx::document::substitute::Replacement;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Dear {{name}},");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let story = edit.story()?;
    /// let text = story.paragraphs()[0].text();
    /// let start = text.find("{{").unwrap();
    /// let end = text.find("}}").unwrap() + 2;
    /// edit.substitute(&story, &[Replacement::new(0, start..end, "Ada")])?;
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.document()?.text()?, "Dear Ada,");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an XML error for malformed markup.
    pub fn story(&self) -> Result<Story> {
        self.inner_ref(|inner| inner.story())?
    }

    /// The main document's story within the elements `scopes` name. See
    /// [`crate::document::raw::Tree::story_within`].
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{ElementPath, ElementStep};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("one");
    /// builder.document()?.add_paragraph_with_text("two");
    /// let package = builder.build()?;
    /// let edit = package.edit()?;
    /// let path = ElementPath::new([ElementStep::Paragraph(Position::new(1))]);
    /// let story = edit.story_within(&[path])?;
    /// assert_eq!(story.paragraphs().len(), 1);
    /// assert_eq!(story.paragraphs()[0].text(), "two");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// A typed refusal for a path that names nothing, or an XML error.
    pub fn story_within(&self, scopes: &[super::super::ElementPath]) -> Result<Story> {
        self.inner_ref(|inner| inner.story_within(scopes))?
    }

    /// Apply `replacements`, sorted by paragraph and range and planned over
    /// [`Self::story`] of this edit, to the main document. See
    /// [`super::super::substitute`] for the rewriting rules.
    ///
    /// # Errors
    ///
    /// Returns a stale-source error for a story of other bytes or of a part,
    /// a typed refusal for a range that is invalid, overlaps another, holds
    /// fixed content or names a field that is not replaceable, and a limit
    /// error. The edit is unchanged on failure. Source-backed edits refuse
    /// it as an unmanaged-only operation.
    pub fn substitute(&mut self, story: &Story, replacements: &[Replacement]) -> Result<Outcome> {
        self.inner_mut()?.substitute(story, replacements)
    }

    /// [`Self::substitute`], formatting the runs that
    /// [`Replacement::formatted`] replacements split out. See
    /// [`crate::document::raw::Tree::substitute_formatted`].
    ///
    /// # Errors
    ///
    /// As [`Self::substitute`], and a refusal for run properties a run
    /// cannot hold.
    pub fn substitute_formatted<F: FnMut(&mut properties::Run)>(
        &mut self,
        story: &Story,
        replacements: &[Replacement],
        format: F,
    ) -> Result<Outcome> {
        self.inner_mut()?
            .substitute_formatted(story, replacements, format)
    }
}

impl Edit<Package> {
    /// The header, footer, footnotes, endnotes and comments parts the main
    /// document relates, in relationship order, each once.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable main part.
    pub fn story_parts(&self) -> Result<Vec<StoryPart>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        Ok(story_parts_of(main.rels())?)
    }

    /// The story of one part (a name from [`Self::story_parts`]), staged
    /// changes included.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for a missing part, or an XML error.
    pub fn part_story(&self, part: &str) -> Result<Story> {
        let payload = self
            .part_payload(part)?
            .ok_or_else(|| crate::Error::PartNotFound(part.to_owned()))?;
        story::part_story(part, payload)
    }

    /// Apply `replacements` planned over [`Self::part_story`] to that part,
    /// staged as an exact part replacement that publishes with the
    /// main-document edit.
    ///
    /// # Errors
    ///
    /// As [`Self::substitute`], for the part.
    pub fn substitute_part(
        &mut self,
        story: &Story,
        replacements: &[Replacement],
    ) -> Result<Outcome> {
        self.substitute_part_formatted(story, replacements, |_run| {})
    }

    /// [`Self::substitute_part`], formatting the runs that
    /// [`Replacement::formatted`] replacements split out.
    ///
    /// # Errors
    ///
    /// As [`Self::substitute_part`], and a refusal for run properties a run
    /// cannot hold.
    pub fn substitute_part_formatted<F: FnMut(&mut properties::Run)>(
        &mut self,
        story: &Story,
        replacements: &[Replacement],
        format: F,
    ) -> Result<Outcome> {
        let Some(part) = story.part() else {
            return self.substitute_formatted(story, replacements, format);
        };
        let current: Arc<Vec<u8>> = self
            .part_payload(part)?
            .ok_or_else(|| crate::Error::PartNotFound(part.to_owned()))?;
        let (after, outcome) = story::substitute_part(story, &current, replacements, format)?;
        if let Some(after) = after {
            self.replace_part(part, after)?;
        }
        Ok(outcome)
    }
}
