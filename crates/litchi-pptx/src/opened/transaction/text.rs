//! Paragraph and run insertion and removal in slide, table-cell and
//! notes-page text bodies (record 0333).

use litchi_core::Position;
use litchi_opc::PackURI;

use super::Edit;
use crate::opened::{TextBody, TextChange, TextProperties};
use crate::{Error, Result};

impl Edit {
    /// Insert a paragraph into a text body before the paragraph at
    /// `position` (the paragraph count appends). The paragraph holds
    /// `properties`' text (a line feed starts a line break) and its
    /// character formatting and alignment; empty text gives an empty
    /// paragraph whose end-of-paragraph properties carry the formatting.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_pptx::opened::{TextBody, TextProperties};
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(Position::new(0), "Blank")?;
    /// edit.add_text_box(slide.name(), "first", (0, 0, 914_400, 914_400))?;
    /// let body = TextBody::Shape { shape: Position::new(0) };
    /// edit.insert_paragraph(0_usize, body, Position::new(1), TextProperties::default().with_text("second"))?;
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.presentation()?.slides()?[0].text()?, "first\nsecond");
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape, cell, notes page or
    /// paragraph position, a target without a text body, an assignment a
    /// new paragraph cannot take (fills, geometry), invalid text, or
    /// markup-compatibility content. A failure leaves the edit unchanged.
    pub fn insert_paragraph<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        body: TextBody,
        position: Position,
        properties: TextProperties<'_>,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &TextChange::InsertParagraph {
                position: position.get(),
                properties,
            },
        )
    }

    /// Remove the paragraph at `paragraph` of a text body, with its runs.
    /// A text body keeps at least one paragraph.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing target or paragraph, the body's last
    /// paragraph, or markup-compatibility content.
    pub fn remove_paragraph<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        body: TextBody,
        paragraph: Position,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &TextChange::RemoveParagraph {
                paragraph: paragraph.get(),
            },
        )
    }

    /// Insert a text run (`a:r`) into paragraph `paragraph` of a text body,
    /// before its run child (`a:r`, `a:br`, `a:fld`) at `position` (the run
    /// count appends), with `properties`' text and character formatting.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing target, paragraph or run position,
    /// text with a line break, an assignment a run cannot take (alignment,
    /// fills, geometry), or markup-compatibility content.
    pub fn insert_run<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        body: TextBody,
        paragraph: Position,
        position: Position,
        properties: TextProperties<'_>,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &TextChange::InsertRun {
                paragraph: paragraph.get(),
                position: position.get(),
                properties,
            },
        )
    }

    /// Remove run child `run` (`a:r`, `a:br` or `a:fld`) of paragraph
    /// `paragraph` of a text body. A paragraph may be left without runs.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing target, paragraph or run, or
    /// markup-compatibility content.
    pub fn remove_run<'s>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        body: TextBody,
        paragraph: Position,
        run: Position,
    ) -> Result<()> {
        self.change_text_body(
            slide.into(),
            body,
            &TextChange::RemoveRun {
                paragraph: paragraph.get(),
                run: run.get(),
            },
        )
    }

    fn change_text_body(
        &mut self,
        slide: crate::slide::Selector<'_>,
        body: TextBody,
        change: &TextChange<'_>,
    ) -> Result<()> {
        let selected_slide = self.resolve_slide(slide)?;
        let (max_text_bytes, presentation_name) = {
            let source = self.source_ref()?;
            (
                source.limits.max_text_bytes(),
                source.presentation_name.clone(),
            )
        };
        let shape_limits = self.semantic_limits().shape();
        let max_output_bytes = shape_limits.output_bytes();
        let (part_name, candidate) = match body {
            TextBody::Notes { .. } => {
                let source = crate::notes::load_snapshot_with_limits(
                    &self.working,
                    &presentation_name,
                    self.notes_limits(),
                )?;
                let Some(note) = source.as_ref().and_then(|source| {
                    source
                        .slides()
                        .iter()
                        .find(|note| note.owner() == selected_slide.part_name.as_str())
                }) else {
                    return Err(Error::refused(
                        litchi_core::edit::RefusalKind::UnsupportedTarget,
                        "notes_text_structure",
                        "the slide has no notes page",
                    ));
                };
                let part_name = PackURI::new(note.part())
                    .map_err(|_error| invalid("opened-presentation notes part name is invalid"))?;
                let candidate = crate::opened::stage_text_structure_with_limits(
                    note.xml(),
                    body,
                    change,
                    max_text_bytes,
                    max_output_bytes,
                    shape_limits,
                )?;
                (part_name, candidate)
            },
            TextBody::Shape { .. } | TextBody::Cell { .. } => {
                let owner = self.working.get_part(&selected_slide.part_name)?;
                crate::parts::validate_content_type(
                    owner,
                    litchi_opc::constants::content_type::PML_SLIDE,
                )?;
                let candidate = crate::opened::stage_text_structure_with_limits(
                    owner.blob(),
                    body,
                    change,
                    max_text_bytes,
                    max_output_bytes,
                    shape_limits,
                )?;
                (selected_slide.part_name.clone(), candidate)
            },
        };
        self.working.get_part_mut(&part_name)?.set_blob(candidate);
        if !matches!(body, TextBody::Notes { .. }) {
            self.note_scene_read(&part_name);
        }
        Ok(())
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}
