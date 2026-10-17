//! Table grid columns and rows of a slide graphic frame (records 0321 and
//! 0333).
//!
//! [`Edit::set_table_column_width`], [`Edit::insert_table_column`] and
//! [`Edit::remove_table_column`] change one column of a table: its
//! `a:gridCol` and the cell at its position in every row;
//! [`Edit::insert_table_row`] and [`Edit::remove_table_row`] change one
//! `a:tr`. The frame's extent follows the sum of the column widths or row
//! heights, as OfficeCLI keeps it; merged cells stay consistent where that
//! is safe.

use litchi_core::Position;

use super::{Edit, invalid};
use crate::Result;
use crate::table::columns::Change;

impl Edit {
    /// Set the width, in EMU, of the column at a zero-based position of the
    /// table frame at a depth-first shape position of a slide. Returns
    /// whether the slide changed; an equal width is an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(Position::new(0), "Blank")?;
    /// let cells = vec![vec!["a".to_owned(), "b".to_owned()]];
    /// edit.add_table(slide.name(), &cells, (0, 0, 1_828_800, 914_400))?;
    /// assert!(edit.set_table_column_width(0_usize, 0_usize, Position::new(1), 1_371_600)?);
    /// let package = edit.commit()?.into_package();
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape or column, a shape that
    /// is not a table, a width out of range, or markup-compatibility content.
    pub fn set_table_column_width<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        column: Position,
        width: i64,
    ) -> Result<bool> {
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::Width {
                column: column.get(),
                width,
            },
        )
    }

    /// Insert a column before a zero-based position (the column count
    /// appends) of a table frame, with a width in EMU (`None`: the average
    /// of the existing columns) and seed text for its new cells.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a position past the
    /// column count, a position inside cells merged across columns, invalid
    /// text, or markup-compatibility content.
    pub fn insert_table_column<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        position: Position,
        width: Option<i64>,
        text: &str,
    ) -> Result<()> {
        if !text.chars().all(|ch| {
            matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
        }) {
            return Err(invalid("table cell text contains a character XML cannot carry"));
        }
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::Insert {
                position: position.get(),
                width,
                text,
            },
        )
        .map(|_| ())
    }

    /// Remove the column at a zero-based position of a table frame, with
    /// its cell in every row. The last column is kept. A merge across the
    /// column spans one column less; when its first cell goes, the next
    /// cell becomes its first.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape or column, a column that
    /// takes part in a merge spanning rows and columns, a table with one
    /// column, or markup-compatibility content.
    pub fn remove_table_column<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        column: Position,
    ) -> Result<()> {
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::Remove {
                column: column.get(),
            },
        )
        .map(|_| ())
    }

    /// Insert a row before a zero-based position (the row count appends)
    /// of a table frame, with a height in EMU (`None`: the first row's) and
    /// seed text for its cells by column (missing columns are empty). The
    /// row holds one cell per grid column.
    ///
    /// ```
    /// use litchi_core::Position;
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(Position::new(0), "Blank")?;
    /// let cells = vec![vec!["a".to_owned(), "b".to_owned()]];
    /// edit.add_table(slide.name(), &cells, (0, 0, 1_828_800, 914_400))?;
    /// edit.insert_table_row(0_usize, 0_usize, Position::new(1), None, &["c", "d"])?;
    /// let package = edit.commit()?.into_package();
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a position past the
    /// row count, a position inside cells merged across rows, more seed
    /// texts than columns, invalid text, or markup-compatibility content.
    pub fn insert_table_row<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        position: Position,
        height: Option<i64>,
        texts: &[&str],
    ) -> Result<()> {
        if !texts.iter().all(|text| {
            text.chars().all(|ch| {
                matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
            })
        }) {
            return Err(invalid("table cell text contains a character XML cannot carry"));
        }
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::InsertRow {
                position: position.get(),
                height,
                texts,
            },
        )
        .map(|_| ())
    }

    /// Set the height, in EMU, of the row at a zero-based position of a
    /// table frame (record 0333). Returns whether the slide changed; an
    /// equal height is an exact no-op. The frame's height follows the sum
    /// of the row heights.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape or row, a height out of
    /// range, or markup-compatibility content.
    pub fn set_table_row_height<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        row: Position,
        height: i64,
    ) -> Result<bool> {
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::RowHeight {
                row: row.get(),
                height,
            },
        )
    }

    /// Remove the row at a zero-based position of a table frame. The last
    /// row is kept. A merge across the row spans one row less; when its
    /// first cell goes, the cells below become its first.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide, shape or row, a table with one
    /// row, a merge whose covered cells are not in place, or
    /// markup-compatibility content.
    pub fn remove_table_row<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        row: Position,
    ) -> Result<()> {
        self.change_table_columns(
            slide.into(),
            shape.into(),
            Change::RemoveRow { row: row.get() },
        )
        .map(|_| ())
    }

    fn change_table_columns(
        &mut self,
        slide: crate::slide::Selector<'_>,
        key: crate::shape::Key<'_>,
        change: Change<'_>,
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide)?;
        let owner = self.working.get_part(&selected.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let xml = owner.blob();
        let scene = crate::shape::Scene::read(xml)?;
        if scene.is_rewritten() {
            return Err(crate::Error::UnsafeEdit {
                operation: "table_column",
                reason: "table column edits do not select markup-compatibility branches",
            });
        }
        let shape = scene.shape(key)?;
        let crate::shape::Shape::Table(_) = shape else {
            return Err(invalid("the selected shape is not a table"));
        };
        let span = shape.common().span()?.range(xml.len())?;
        let table = frame_table(xml, span.clone())?;
        let before = table.grid_columns()?.len();
        let cells: Vec<usize> = table.rows()?.iter().map(|row| row.cell_count()).collect();
        let Some(next) = crate::table::columns::change(xml, span, change)? else {
            return Ok(false);
        };
        let staged = crate::shape::Scene::read(&next)?;
        let column_delta = |count: usize| match change {
            Change::Insert { .. } => count.checked_add(1),
            Change::Remove { .. } => count.checked_sub(1),
            Change::Width { .. }
            | Change::InsertRow { .. }
            | Change::RemoveRow { .. }
            | Change::RowHeight { .. } => Some(count),
        };
        let mut expected_cells: Vec<Option<usize>> = cells.into_iter().map(column_delta).collect();
        match change {
            Change::InsertRow { position, .. } => expected_cells.insert(position, Some(before)),
            Change::RemoveRow { row } => {
                expected_cells.remove(row);
            },
            Change::Width { .. }
            | Change::Insert { .. }
            | Change::Remove { .. }
            | Change::RowHeight { .. } => {},
        }
        let after = staged.shape(key)?;
        let after = frame_table(&next, after.common().span()?.range(next.len())?)?;
        let staged_cells: Vec<Option<usize>> = after
            .rows()?
            .iter()
            .map(|row| Some(row.cell_count()))
            .collect();
        if staged.len() != scene.len()
            || Some(after.grid_columns()?.len()) != column_delta(before)
            || staged_cells != expected_cells
        {
            return Err(invalid("the table column edit did not round-trip"));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(next);
        self.note_scene_read(&selected.part_name);
        Ok(true)
    }
}

/// The table of the graphic frame at `span` of `xml`, with the owner's
/// namespace scope.
fn frame_table(xml: &[u8], span: std::ops::Range<usize>) -> Result<crate::table::Table<'_>> {
    let frame = xml
        .get(span.clone())
        .ok_or_else(|| invalid("the table frame span is outside its slide"))?;
    crate::table::Table::from_graphic_frame_with_context(frame, xml, span.start)
}
