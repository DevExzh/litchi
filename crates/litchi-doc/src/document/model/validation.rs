//! Structural table validation and materialization.
//!
//! The semantic facade delegates table grouping here so the document model has
//! one canonical implementation for nesting, row-end, and cell-end invariants.
//! The helpers remain `pub(super)` because they are implementation seams, not
//! additional public API.
//!
//! Grouping reads each paragraph's derived table depth
//! ([`ParagraphProperties::table_depth`]), never the stored `sprmPItap` value:
//! a Word 97 table paragraph stores no depth, and its tables are top-level
//! tables ([MS-DOC] §2.4.3).
//!
//! [`ParagraphProperties::table_depth`]: crate::parts::pap::ParagraphProperties::table_depth

use crate::cp_range::CpRange;
use crate::paragraph::Paragraph;
use crate::parts::tap::TableProperties;
use crate::table::{Cell, Row, Table};

/// Where one paragraph falls among a story's top-level blocks.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Placement {
    /// A paragraph outside every table (depth 0).
    Paragraph,
    /// The first paragraph of a top-level table (depth 1, no table open).
    TableStart,
    /// A later paragraph of the open top-level table, including the
    /// paragraphs of tables nested in it (depth 1 or more).
    TableContinue,
    /// A nested-table paragraph that no top-level table encloses: it belongs
    /// to no top-level block.
    Orphan,
}

/// Streams a story's paragraphs, in story order, into top-level blocks.
///
/// A top-level table begins at a paragraph of depth 1 and runs through every
/// following paragraph of depth 1 or more, so the paragraphs of nested tables
/// stay inside the table that encloses them. A paragraph of depth 2 or more
/// with no open table is an orphan and forms no block.
#[derive(Debug, Default, Clone, Copy)]
pub(crate) struct BlockGrouper {
    table_open: bool,
}

impl BlockGrouper {
    /// Place the next paragraph, whose table depth is `depth`.
    pub(crate) fn place(&mut self, depth: u32) -> Placement {
        if self.table_open && depth >= 1 {
            return Placement::TableContinue;
        }
        self.table_open = depth == 1;
        match depth {
            0 => Placement::Paragraph,
            1 => Placement::TableStart,
            _ => Placement::Orphan,
        }
    }
}

/// The smallest CP range holding a sequence of ranges, while every one is
/// known.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Extent {
    /// No range seen yet.
    Empty,
    /// The cover of every range seen.
    Known(CpRange),
    /// A value without a range was seen: the cover is unknown.
    Unknown,
}

impl Extent {
    fn add(self, range: Option<CpRange>) -> Self {
        match (self, range) {
            (Self::Unknown, _) | (_, None) => Self::Unknown,
            (Self::Empty, Some(range)) => Self::Known(range),
            (Self::Known(cover), Some(range)) => Self::Known(cover.cover(range)),
        }
    }

    fn range(self) -> Option<CpRange> {
        match self {
            Self::Known(range) => Some(range),
            Self::Empty | Self::Unknown => None,
        }
    }
}

/// Build a table from its paragraphs, or `None` when they form no row.
///
/// `depth` is the table's depth; paragraphs nested deeper are skipped. The
/// table's CP range covers every one of its paragraphs.
pub(super) fn table_from_paragraphs(paragraphs: Vec<Paragraph>, depth: u32) -> Option<Table> {
    let extent = paragraphs.iter().fold(Extent::Empty, |extent, paragraph| {
        extent.add(paragraph.cp_range())
    });
    let rows = rows_from_table_paragraphs(paragraphs, depth);
    if rows.is_empty() {
        return None;
    }
    let properties = rows.first().and_then(|row| row.properties()).cloned();
    let mut table = match properties {
        Some(properties) => Table::with_properties(rows, properties),
        None => Table::new(rows),
    };
    table.set_cp_range(extent.range());
    Some(table)
}

/// Split one table's paragraphs into rows at its row-end marks, while
/// skipping the paragraphs of nested tables and retaining the source TAP
/// metadata used by snapshot edits.
///
/// A row's CP range covers every paragraph from its first one, nested ones
/// included, through its row-end mark.
pub(super) fn rows_from_table_paragraphs(paragraphs: Vec<Paragraph>, depth: u32) -> Vec<Row> {
    let mut rows = Vec::new();
    let mut current_row_paras = Vec::new();
    let mut row_extent = Extent::Empty;

    for para in paragraphs {
        let props = para.properties();
        let paragraph_depth = props.table_depth();
        row_extent = row_extent.add(para.cp_range());

        if paragraph_depth > depth {
            continue;
        }

        if props.is_table_row_end && paragraph_depth == depth {
            let row_properties = props.table_properties.clone();
            let revision = para.table_formatting_revision().cloned();
            let preserved = props.table_properties_preserved_for_revision;
            let range = row_extent.range();
            // The row-end mark carries no cell content; a row holding
            // nothing else still gets its one empty cell.
            let cells = cells_from_row_paragraphs(
                std::mem::take(&mut current_row_paras),
                row_properties.as_ref(),
                range.map(CpRange::start),
            );
            let mut row = Row::with_metadata(cells, row_properties, revision, preserved);
            row.set_cp_range(range);
            rows.push(row);
            row_extent = Extent::Empty;
            continue;
        }

        current_row_paras.push(para);
    }

    // Preserve an incomplete final row instead of silently dropping its
    // paragraphs; malformed snapshots remain inspectable and editable.
    if let Some(last) = current_row_paras.last() {
        let row_properties = last.properties().table_properties.clone();
        let revision = last.table_formatting_revision().cloned();
        let preserved = last.properties().table_properties_preserved_for_revision;
        let range = row_extent.range();
        let cells = cells_from_row_paragraphs(
            current_row_paras,
            row_properties.as_ref(),
            range.map(CpRange::start),
        );
        let mut row = Row::with_metadata(cells, row_properties, revision, preserved);
        row.set_cp_range(range);
        rows.push(row);
    }

    crate::table::apply_table_cell_styles(&mut rows);
    rows
}

/// Split a row's paragraphs at cell marks while preserving cell properties.
///
/// `row_paras` excludes the row's own row-end mark, so it is empty for a row
/// that held only that mark; every row the callers pass held at least one
/// paragraph, and a row that forms no cell gets one empty cell. A row-end
/// mark of a shallower table, which a well-formed row never holds, is
/// skipped as content, as before.
///
/// Cells tile the row: the first starts at `row_start`, each later one where
/// the previous one ends, and each ends with its last paragraph, so a table
/// nested in a cell lies inside the cell's range.
fn cells_from_row_paragraphs(
    row_paras: Vec<Paragraph>,
    table_properties: Option<&TableProperties>,
    row_start: Option<u32>,
) -> Vec<Cell> {
    let mut cells = Vec::new();
    let mut cell_paragraphs = Vec::new();
    let mut cell_start = row_start;
    let mut finish = |cell_paragraphs: Vec<Paragraph>, cells: &mut Vec<Cell>| {
        let properties = table_properties
            .and_then(|tap| tap.cell_properties.get(cells.len()))
            .cloned();
        let end = cell_paragraphs
            .iter()
            .fold(Extent::Empty, |extent, paragraph| {
                extent.add(paragraph.cp_range())
            })
            .range()
            .map(CpRange::end);
        let range = cell_start
            .zip(end)
            .and_then(|(start, end)| CpRange::new(start, end));
        // Once a range is unknown, every later cell's start is unknown too.
        cell_start = cell_start.and(end);
        let mut cell = Cell::with_properties(cell_paragraphs, properties);
        cell.set_cp_range(range);
        cells.push(cell);
    };
    for para in row_paras {
        let props = para.properties();
        if props.is_table_row_end {
            continue;
        }
        let ends_cell = props.is_table_cell_end;
        cell_paragraphs.push(para);
        if ends_cell {
            finish(std::mem::take(&mut cell_paragraphs), &mut cells);
        }
    }

    if !cell_paragraphs.is_empty() {
        finish(cell_paragraphs, &mut cells);
    }

    if cells.is_empty() {
        cells.push(Cell::new(String::new()));
    }

    cells
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parts::pap::ParagraphProperties;

    /// A paragraph with the table flags a PAPX would give it.
    fn paragraph(text: &str, in_table: bool, stored_depth: i32, mark: Mark) -> Paragraph {
        let properties = ParagraphProperties {
            in_table,
            table_nesting_level: stored_depth,
            is_table_cell_end: matches!(mark, Mark::Cell | Mark::Row),
            is_table_row_end: matches!(mark, Mark::Row),
            ..ParagraphProperties::default()
        };
        Paragraph::with_properties(text.to_string(), properties)
    }

    #[derive(Clone, Copy)]
    enum Mark {
        Paragraph,
        Cell,
        Row,
    }

    fn placements(depths: &[u32]) -> Vec<Placement> {
        let mut grouper = BlockGrouper::default();
        depths.iter().map(|&depth| grouper.place(depth)).collect()
    }

    #[test]
    fn grouper_opens_tables_at_depth_one_and_keeps_nested_paragraphs_inside() {
        use Placement::{Orphan, Paragraph, TableContinue, TableStart};
        assert_eq!(
            placements(&[0, 1, 1, 2, 2, 1, 0, 2, 1, 0]),
            [
                Paragraph,
                TableStart,
                TableContinue,
                TableContinue,
                TableContinue,
                TableContinue,
                Paragraph,
                Orphan,
                TableStart,
                Paragraph,
            ]
        );
        // Two adjacent depth-1 runs are one table: grouping never splits on
        // row properties.
        assert_eq!(placements(&[1, 1]), [TableStart, TableContinue]);
        // An orphan never opens a table, so the depth-1 paragraph after it
        // starts one.
        assert_eq!(placements(&[3, 1]), [Orphan, TableStart]);
    }

    #[test]
    fn word97_table_paragraphs_without_a_stored_depth_form_rows_and_cells() {
        // `sprmPFInTable` with no `sprmPItap`: the Word 97 encoding.
        let table = table_from_paragraphs(
            vec![
                paragraph("a", true, 0, Mark::Cell),
                paragraph("b1", true, 0, Mark::Paragraph),
                paragraph("b2", true, 0, Mark::Cell),
                paragraph("", true, 0, Mark::Row),
                paragraph("c", true, 0, Mark::Cell),
                paragraph("d", true, 0, Mark::Cell),
                paragraph("", true, 0, Mark::Row),
            ],
            1,
        )
        .expect("a Word 97 table forms rows");
        let rows = table.rows().unwrap();
        assert_eq!(rows.len(), 2);
        let first = rows[0].cells().unwrap();
        assert_eq!(first.len(), 2);
        assert_eq!(first[0].text().unwrap(), "a");
        assert_eq!(first[1].text().unwrap(), "b1\nb2");
        let second = rows[1].cells().unwrap();
        assert_eq!(
            second
                .iter()
                .map(|cell| cell.text().unwrap().to_string())
                .collect::<Vec<_>>(),
            ["c", "d"]
        );
    }

    #[test]
    fn stored_depth_and_word97_paragraphs_agree() {
        let shape = |stored: i32| {
            let table = table_from_paragraphs(
                vec![
                    paragraph("x", true, stored, Mark::Cell),
                    paragraph("", true, stored, Mark::Row),
                ],
                1,
            )
            .unwrap();
            let rows = table.rows().unwrap();
            (rows.len(), rows[0].cell_count().unwrap())
        };
        assert_eq!(shape(0), shape(1));
        assert_eq!(shape(0), (1, 1));
    }

    #[test]
    fn nested_paragraphs_are_skipped_and_row_only_marks_get_an_empty_cell() {
        let rows = rows_from_table_paragraphs(
            vec![
                paragraph("outer", true, 1, Mark::Paragraph),
                paragraph("inner", true, 2, Mark::Paragraph),
                paragraph("inner end", true, 2, Mark::Paragraph),
                paragraph("", true, 1, Mark::Cell),
                paragraph("", true, 1, Mark::Row),
                paragraph("", true, 1, Mark::Row),
            ],
            1,
        );
        assert_eq!(rows.len(), 2);
        let cells = rows[0].cells().unwrap();
        assert_eq!(cells.len(), 1);
        assert_eq!(cells[0].text().unwrap(), "outer\n");
        let only_mark = rows[1].cells().unwrap();
        assert_eq!(only_mark.len(), 1);
        assert_eq!(only_mark[0].text().unwrap(), "");
    }

    #[test]
    fn an_unterminated_final_row_is_kept() {
        let rows = rows_from_table_paragraphs(
            vec![
                paragraph("a", true, 0, Mark::Cell),
                paragraph("tail", true, 0, Mark::Paragraph),
            ],
            1,
        );
        assert_eq!(rows.len(), 1);
        let cells = rows[0].cells().unwrap();
        assert_eq!(cells.len(), 2);
        assert_eq!(cells[1].text().unwrap(), "tail");
    }
}
