//! Office Math equations of an opened document's main story.
//!
//! The verbs are the main-document transaction's; [`Edit::edit_story`] runs
//! them on a header, footer or note.

use litchi_core::Position;

use super::super::{ElementAnchor, ElementPath, ParagraphEquation, Result};
use super::Edit;
use crate::{OfficeMath, OfficeMathParagraph};

impl<P: Clone> Edit<P> {
    /// Insert `math` as an inline equation of the paragraph at `paragraph`;
    /// see [`crate::document::raw::Tree::insert_equation_at`].
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::OfficeMath;
    /// use litchi_docx::document::{ElementPath, ElementStep};
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("Area: ");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.insert_equation_at(&paragraph, None, &OfficeMath::text("πr²"))?;
    /// let package = edit.commit()?.into_package();
    /// let equations = package.document()?.paragraphs()?[0].omml_equations()?;
    /// assert_eq!(equations, ["<m:oMath><m:r><m:t>πr²</m:t></m:r></m:oMath>"]);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::insert_equation_at`].
    pub fn insert_equation_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        math: &OfficeMath,
    ) -> Result<Position> {
        self.inner_mut()?.insert_equation_at(paragraph, at, math)
    }

    /// Insert a display-equation paragraph holding `math`; see
    /// [`crate::document::raw::Tree::insert_display_equation_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::insert_display_equation_at`].
    pub fn insert_display_equation_at(
        &mut self,
        container: &ElementPath,
        anchor: ElementAnchor,
        math: &OfficeMathParagraph,
    ) -> Result<Position> {
        self.inner_mut()?
            .insert_display_equation_at(container, anchor, math)
    }

    /// Replace one equation of a paragraph; see
    /// [`crate::document::raw::Tree::replace_equation_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::replace_equation_at`].
    pub fn replace_equation_at(
        &mut self,
        paragraph: &ElementPath,
        equation: ParagraphEquation,
        math: &OfficeMath,
    ) -> Result<bool> {
        self.inner_mut()?
            .replace_equation_at(paragraph, equation, math)
    }

    /// Replace one math paragraph of a paragraph; see
    /// [`crate::document::raw::Tree::replace_math_paragraph_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::replace_math_paragraph_at`].
    pub fn replace_math_paragraph_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMathParagraph,
    ) -> Result<bool> {
        self.inner_mut()?
            .replace_math_paragraph_at(paragraph, index, math)
    }

    /// Replace the equations of one math paragraph, keeping its properties;
    /// see [`crate::document::raw::Tree::replace_math_paragraph_equations_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::replace_math_paragraph_equations_at`].
    pub fn replace_math_paragraph_equations_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMath,
    ) -> Result<bool> {
        self.inner_mut()?
            .replace_math_paragraph_equations_at(paragraph, index, math)
    }

    /// Remove one math paragraph of a paragraph; see
    /// [`crate::document::raw::Tree::remove_math_paragraph_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::remove_math_paragraph_at`].
    pub fn remove_math_paragraph_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<&mut Self> {
        self.inner_mut()?
            .remove_math_paragraph_at(paragraph, index)?;
        Ok(self)
    }

    /// Remove one equation of a paragraph; see
    /// [`crate::document::raw::Tree::remove_equation_at`].
    ///
    /// # Errors
    ///
    /// As [`crate::document::raw::Tree::remove_equation_at`].
    pub fn remove_equation_at(
        &mut self,
        paragraph: &ElementPath,
        equation: ParagraphEquation,
    ) -> Result<&mut Self> {
        self.inner_mut()?.remove_equation_at(paragraph, equation)?;
        Ok(self)
    }
}
