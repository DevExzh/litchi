//! Dependency-free text transfer into immutable RTF transactions.

use super::{Commit, Edit, Error, HeaderFooterParagraph, Snapshot};
use crate::TableCellPath;

/// A checked, non-applying transfer into one immutable target snapshot.
///
/// Transfer deliberately copies only plain Unicode text. It never copies font,
/// color, style, list, field, object, drawing, or external-resource handles, so
/// the plan has no hidden resource dependencies. Publication still goes
/// through the ordinary atomic transaction and full RTF readback.
pub struct TransferPlan {
    edit: Edit,
}

impl TransferPlan {
    /// Plans insertion of one source paragraph as a plain target paragraph.
    ///
    /// # Errors
    /// Returns an error when the source contains an inline line break or when
    /// the checked target cannot accept the structural insertion.
    pub fn plain_paragraph(
        source: &Snapshot,
        source_position: usize,
        target: &Snapshot,
        insert_after: usize,
    ) -> Result<Self, Error> {
        let paragraphs = source.body().paragraphs().collect::<Vec<_>>();
        let count = paragraphs.len();
        let text = paragraphs
            .get(source_position)
            .ok_or(Error::ParagraphOutOfRange {
                position: source_position,
                count,
            })?
            .to_text();
        if text.contains('\n') {
            return Err(Error::UnsupportedSource(
                "plain paragraph transfer refuses inline line breaks",
            ));
        }
        let mut edit = target.edit();
        edit.insert_paragraph_after(insert_after, text)?;
        Ok(Self { edit })
    }

    /// Plans copying plain text between checked table-cell destinations.
    ///
    /// # Errors
    /// Returns an error when either path is invalid or the target cell's
    /// dependent positional content cannot survive the replacement.
    pub fn table_cell_text(
        source: &Snapshot,
        source_path: &TableCellPath,
        target: &Snapshot,
        target_path: TableCellPath,
    ) -> Result<Self, Error> {
        let text = super::table_cell(source, source_path)?.text().to_string();
        let mut edit = target.edit();
        edit.set_table_cell_text(target_path, text)?;
        Ok(Self { edit })
    }

    /// Plans copying one plain header/footer paragraph into another.
    ///
    /// # Errors
    /// Returns an error when either selector is invalid or the target story
    /// owns positioned content that would acquire stale offsets.
    pub fn header_footer_text(
        source: &Snapshot,
        source_target: HeaderFooterParagraph,
        target: &Snapshot,
        target_destination: HeaderFooterParagraph,
    ) -> Result<Self, Error> {
        let text = super::header_footer(source, source_target)?
            .paragraphs
            .get(source_target.paragraph())
            .ok_or(Error::DestinationOutOfRange("header/footer paragraph"))?
            .text
            .to_string();
        let mut edit = target.edit();
        edit.set_header_footer_text(target_destination, text)?;
        Ok(Self { edit })
    }

    /// This transfer class never imports format resource handles.
    #[must_use]
    pub const fn is_dependency_free(&self) -> bool {
        true
    }

    /// Returns the still-uncommitted target edit.
    #[must_use]
    pub fn into_edit(self) -> Edit {
        self.edit
    }

    /// Validates and publishes the target transaction atomically.
    ///
    /// # Errors
    /// Returns the ordinary transaction refusal without mutating either input.
    pub fn commit(self) -> Result<Commit, Error> {
        self.edit.commit()
    }
}
