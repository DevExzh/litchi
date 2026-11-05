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

use super::picture::{attribute_value, first_start_tag};
use super::{Edit, invalid};
use crate::Result;
use crate::table::Properties;
use crate::table::columns::Change;

/// The `a:tblPr` style switches, by attribute name and accessor.
const SWITCHES: [(&str, fn(&Properties) -> bool); 6] = [
    ("firstRow", |properties| properties.first_row),
    ("firstCol", |properties| properties.first_col),
    ("lastRow", |properties| properties.last_row),
    ("lastCol", |properties| properties.last_col),
    ("bandRow", |properties| properties.band_row),
    ("bandCol", |properties| properties.band_col),
];

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

    /// Change the style switches (`a:tblPr@firstRow`, `firstCol`,
    /// `lastRow`, `lastCol`, `bandRow`, `bandCol`) of the table frame at a
    /// depth-first shape position of a slide (record 0342). A switch turned
    /// on is written `"1"`; one turned off loses its attribute (off is the
    /// default). A missing `a:tblPr` is created. Only the changed attributes
    /// change; equal switches are an exact no-op. Returns whether the slide
    /// changed.
    ///
    /// ```
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// let slide = edit.insert_slide(litchi_core::Position::new(0), "Blank")?;
    /// let cells = vec![vec!["a".to_owned(), "b".to_owned()]];
    /// edit.add_table(slide.name(), &cells, (0, 0, 1_828_800, 914_400))?;
    /// assert!(edit.set_table_properties(0_usize, 0_usize, |properties| {
    ///     properties.band_row = false;
    ///     properties.last_col = true;
    /// })?);
    /// let package = edit.commit()?.into_package();
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing slide or shape, a shape that is not a
    /// table, a change of the table style (not written), or
    /// markup-compatibility content.
    pub fn set_table_properties<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        change: impl FnOnce(&mut Properties),
    ) -> Result<bool> {
        let selected = self.resolve_slide(slide.into())?;
        let key = shape.into();
        let owner = self.working.get_part(&selected.part_name)?;
        crate::parts::validate_content_type(owner, litchi_opc::constants::content_type::PML_SLIDE)?;
        let xml = owner.blob();
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "set_table_properties",
                "table edits do not select markup-compatibility branches",
            ));
        }
        let target_shape = scene.shape(key)?;
        let crate::shape::Shape::Table(_) = target_shape else {
            return Err(invalid("the selected shape is not a table"));
        };
        let span = target_shape.common().span()?.range(xml.len())?;
        let current = frame_table(xml, span.clone())?.properties()?;
        let base = current.clone().unwrap_or_default();
        let mut target = base.clone();
        change(&mut target);
        if target.style_id != base.style_id {
            return Err(invalid("table style writes are not supported"));
        }
        if target == base {
            return Ok(false);
        }
        let frame = &xml[span.clone()];
        let grid = first_start_tag(frame, b"tblGrid")?;
        let properties = first_start_tag(frame, b"tblPr")?
            .filter(|tag| grid.as_ref().is_none_or(|grid| tag.start < grid.start));
        let (tag, replacement) = match properties {
            Some(tag) => {
                let mut rewritten = frame[tag.clone()].to_vec();
                for (name, switch) in SWITCHES {
                    if switch(&base) != switch(&target) {
                        rewritten = set_switch(&rewritten, name.as_bytes(), switch(&target))?;
                    }
                }
                (tag, rewritten)
            },
            None => {
                let table = first_start_tag(frame, b"tbl")?
                    .ok_or_else(|| invalid("the table frame has no a:tbl"))?;
                let open = &frame[table.clone()];
                if open.ends_with(b"/>") {
                    return Err(invalid("the table frame's a:tbl is empty"));
                }
                let name_end = open
                    .iter()
                    .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))
                    .unwrap_or(open.len());
                let qualified = &open[1..name_end];
                let prefix = qualified
                    .iter()
                    .position(|byte| *byte == b':')
                    .map_or(&b""[..], |colon| &qualified[..=colon]);
                let mut element = b"<".to_vec();
                element.extend_from_slice(prefix);
                element.extend_from_slice(b"tblPr");
                for (name, switch) in SWITCHES {
                    if switch(&target) {
                        element.push(b' ');
                        element.extend_from_slice(name.as_bytes());
                        element.extend_from_slice(b"=\"1\"");
                    }
                }
                element.extend_from_slice(b"/>");
                (table.end..table.end, element)
            },
        };
        let start = span.start + tag.start;
        let end = span.start + tag.end;
        let mut next = Vec::with_capacity(xml.len() + replacement.len());
        next.extend_from_slice(&xml[..start]);
        next.extend_from_slice(&replacement);
        next.extend_from_slice(&xml[end..]);
        let staged = crate::shape::Scene::read(&next, shape_limits)?;
        let after = staged.shape(key)?;
        let read = frame_table(&next, after.common().span()?.range(next.len())?)?.properties()?;
        if staged.len() != scene.len() || read.as_ref() != Some(&target) {
            return Err(invalid("the table property edit did not round-trip"));
        }
        self.working
            .get_part_mut(&selected.part_name)?
            .set_blob(next);
        self.note_scene_read(&selected.part_name);
        Ok(true)
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
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(xml, shape_limits)?;
        if scene.is_rewritten() {
            return Err(crate::Error::refused(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "table_column",
                "table column edits do not select markup-compatibility branches",
            ));
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
        let staged = crate::shape::Scene::read(&next, shape_limits)?;
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

/// `tag` (a raw start tag) with the on/off attribute `name` written `"1"`
/// (`on`) or removed.
fn set_switch(tag: &[u8], name: &[u8], on: bool) -> Result<Vec<u8>> {
    let found = attribute_value(tag, name).filter(|(_, key, _)| key.as_slice() == name);
    Ok(match (found, on) {
        (Some((value, _, _)), true) => {
            let mut out = tag[..value.start].to_vec();
            out.push(b'1');
            out.extend_from_slice(&tag[value.end..]);
            out
        },
        (Some((_, _, whole)), false) => {
            let mut start = whole.start;
            while start > 0 && tag[start - 1].is_ascii_whitespace() {
                start -= 1;
            }
            let mut out = tag[..start].to_vec();
            out.extend_from_slice(&tag[whole.end..]);
            out
        },
        (None, true) => {
            let close = if tag.ends_with(b"/>") {
                tag.len() - 2
            } else if tag.ends_with(b">") {
                tag.len() - 1
            } else {
                return Err(invalid("a table property tag is not closed"));
            };
            let mut close = close;
            while close > 0 && tag[close - 1].is_ascii_whitespace() {
                close -= 1;
            }
            let mut out = tag[..close].to_vec();
            out.push(b' ');
            out.extend_from_slice(name);
            out.extend_from_slice(b"=\"1\"");
            out.extend_from_slice(&tag[close..]);
            out
        },
        (None, false) => tag.to_vec(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn switches_are_rewritten_lexically() {
        let tag = br#"<a:tblPr firstRow="1" bandRow="1">"#;
        assert_eq!(
            set_switch(tag, b"bandRow", false).unwrap(),
            br#"<a:tblPr firstRow="1">"#
        );
        assert_eq!(
            set_switch(tag, b"lastCol", true).unwrap(),
            br#"<a:tblPr firstRow="1" bandRow="1" lastCol="1">"#
        );
        assert_eq!(
            set_switch(br#"<a:tblPr firstRow="0"/>"#, b"firstRow", true).unwrap(),
            br#"<a:tblPr firstRow="1"/>"#
        );
    }

    #[test]
    fn table_switches_publish_and_no_ops_keep_the_slide() {
        let source = crate::Package::new().unwrap();
        let mut edit = source.edit().unwrap();
        let slide = edit.insert_slide(Position::new(0), "Blank").unwrap();
        let cells = vec![vec!["a".to_owned(), "b".to_owned()]];
        edit.add_table(slide.name(), &cells, (0, 0, 1_828_800, 914_400))
            .unwrap();
        assert!(!edit.set_table_properties(0_usize, 0_usize, |_| {}).unwrap());
        assert!(
            edit.set_table_properties(0_usize, 0_usize, |properties| {
                properties.last_row = true;
                properties.first_col = true;
                properties.band_row = false;
            })
            .unwrap()
        );
        assert!(
            edit.set_table_properties(0_usize, 0_usize, |properties| {
                properties.style_id = Some("{x}".to_owned());
            })
            .is_err()
        );
        let package = edit.commit().unwrap().into_package();
        let mut edit = package.edit().unwrap();
        assert!(
            !edit
                .set_table_properties(0_usize, 0_usize, |properties| properties.last_row = true)
                .unwrap()
        );
    }
}
