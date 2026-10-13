//! Story text substitution on the package-bound edit (record 0274).

use std::sync::Arc;

use super::super::Result;
use super::super::substitute::{Outcome, Replacement, Story};
use super::super::transaction::substitute as story;
use super::Edit;
use crate::Package;

/// Relationship types (by their last segment) of the story parts the main
/// document part relates: headers, footers, notes and comments.
const STORY_RELATIONSHIPS: [&str; 5] =
    ["/header", "/footer", "/footnotes", "/endnotes", "/comments"];

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
        self.inner.story()
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
        self.inner.substitute(story, replacements)
    }
}

impl Edit<Package> {
    /// The names of the header, footer, footnotes, endnotes and comments
    /// parts the main document relates, in relationship order, each once.
    ///
    /// # Errors
    ///
    /// Returns a package-graph error for an unreadable main part.
    pub fn story_parts(&self) -> Result<Vec<String>> {
        let main = self
            .package
            .opc_package()
            .main_document_part()
            .map_err(crate::Error::from)?;
        let mut names: Vec<String> = Vec::new();
        for relationship in main.rels().iter() {
            if relationship.is_external()
                || !STORY_RELATIONSHIPS
                    .iter()
                    .any(|suffix| relationship.reltype().ends_with(suffix))
            {
                continue;
            }
            let name = relationship
                .target_partname()
                .map_err(crate::Error::from)?
                .as_str()
                .to_owned();
            if !names.contains(&name) {
                names.push(name);
            }
        }
        Ok(names)
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
        let Some(part) = story.part() else {
            return self.substitute(story, replacements);
        };
        let current: Arc<Vec<u8>> = self
            .part_payload(part)?
            .ok_or_else(|| crate::Error::PartNotFound(part.to_owned()))?;
        let (after, outcome) = story::substitute_part(story, &current, replacements)?;
        if let Some(after) = after {
            self.replace_part(part, after)?;
        }
        Ok(outcome)
    }
}
