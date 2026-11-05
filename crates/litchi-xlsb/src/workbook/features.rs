//! Worksheet feature reads of an eager [`Workbook`] (unified-ops record
//! 0322): the values the worksheet edit families list, read without decoding
//! worksheet cells.

use super::Workbook;
use crate::comments::Record;
use crate::conditional_formatting::edit::DifferentialStyle;
use crate::drawing::edit::Element;
use crate::package::error::Result;
use crate::sheet::{FeatureSet, Features};

impl Workbook {
    /// The `select`ed hyperlinks, auto-filter, classic validations and
    /// conditional formats, and page breaks of the worksheet at
    /// complete-catalog position `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// stream.
    pub fn sheet_features(&self, sheet: usize, select: FeatureSet) -> Result<Features> {
        let worksheet = crate::part_edit::worksheet_uri(self, sheet)?;
        let part = self.package.get_part(&worksheet)?;
        crate::sheet::features::read(
            part.blob(),
            part.rels(),
            &self.formula_context,
            sheet,
            self.leniency,
            select,
            self.raw_limits,
        )
    }

    /// The notes of the worksheet at complete-catalog position `sheet`, in
    /// stream order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// comments part.
    pub fn sheet_comments(&self, sheet: usize) -> Result<Vec<Record>> {
        let worksheet = crate::part_edit::worksheet_uri(self, sheet)?;
        let part = self.package.get_part(&worksheet)?;
        match crate::comments::edit::comments_part(part.rels())? {
            Some((_, name)) => {
                crate::comments::edit::read(self.package.get_part(&name)?.blob(), self.raw_limits)
            },
            None => Ok(Vec::new()),
        }
    }

    /// The top-level objects of the drawing of the worksheet at
    /// complete-catalog position `sheet`, in drawing order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// drawing graph.
    pub fn drawing_elements(&self, sheet: usize) -> Result<Vec<Element>> {
        let worksheet = crate::part_edit::worksheet_uri(self, sheet)?;
        crate::drawing::edit::eager_elements(&self.package, &worksheet)
    }

    /// The charts of the drawing of the worksheet at complete-catalog
    /// position `sheet`, in drawing order, with their chart part markup.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// drawing graph.
    pub fn drawing_charts(&self, sheet: usize) -> Result<Vec<crate::drawing::edit::ChartPart>> {
        let worksheet = crate::part_edit::worksheet_uri(self, sheet)?;
        crate::drawing::edit::eager_charts(&self.package, &worksheet)
    }

    /// The differential formats (`BrtDXF`) of the styles part.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed styles part.
    pub fn differential_formats(&self) -> Result<Vec<DifferentialStyle>> {
        match crate::conditional_formatting::edit::styles_part(&self.package)? {
            Some(styles) => crate::conditional_formatting::edit::read_differentials(
                self.package.get_part(&styles)?.blob(),
                self.raw_limits,
            ),
            None => Ok(Vec::new()),
        }
    }
}
