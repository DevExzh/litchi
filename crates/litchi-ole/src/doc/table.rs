/// Table, Row, and Cell structures for legacy Word documents.
use super::package::Result;
use super::paragraph::Paragraph;
use super::parts::tap::{
    CellProperties, TableJustification, TableProperties, TableStyleBorder, TableStyleDefaults,
    TableStyleShading, matching_table_style_conditions,
};
use super::revision::RevisionMark;
use std::sync::Arc;

/// A table in a Word document.
///
/// Represents a table in the binary DOC format.
///
/// # Example
///
/// ```rust,ignore
/// for table in document.tables()? {
///     println!("Table with {} rows", table.row_count()?);
///     for row in table.rows()? {
///         for cell in row.cells()? {
///             println!("Cell: {}", cell.text()?);
///         }
///     }
/// }
/// ```
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around table data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Table {
    /// Rows in the table (shared via Arc for efficient cloning)
    rows: Arc<Vec<Row>>,
    /// Table-level properties (if available)
    properties: Option<TableProperties>,
}

impl Table {
    /// Create a new Table.
    #[allow(dead_code)]
    pub(crate) fn new(rows: Vec<Row>) -> Self {
        Self {
            rows: Arc::new(rows),
            properties: None,
        }
    }

    /// Create a new Table with properties.
    #[allow(dead_code)]
    pub(crate) fn with_properties(rows: Vec<Row>, properties: TableProperties) -> Self {
        Self {
            rows: Arc::new(rows),
            properties: Some(properties),
        }
    }

    /// Get the number of rows in this table.
    pub fn row_count(&self) -> Result<usize> {
        Ok(self.rows.len())
    }

    /// Get the number of columns in this table.
    ///
    /// Returns the column count from the first row, or 0 if the table is empty.
    pub fn column_count(&self) -> Result<usize> {
        if let Some(first_row) = self.rows.first() {
            first_row.cell_count()
        } else {
            Ok(0)
        }
    }

    /// Get all rows in this table.
    ///
    /// Returns a cloned vector. Due to Arc-based sharing in Row/Cell structures,
    /// cloning is relatively cheap (only increments atomic reference counts).
    pub fn rows(&self) -> Result<Vec<Row>> {
        Ok((*self.rows).clone())
    }

    /// Get a specific cell by row and column index.
    ///
    /// Returns `None` if the indices are out of bounds.
    pub fn cell(&self, row_idx: usize, col_idx: usize) -> Result<Option<Cell>> {
        if let Some(row) = self.rows.get(row_idx) {
            let cells = row.cells()?;
            Ok(cells.get(col_idx).cloned())
        } else {
            Ok(None)
        }
    }

    /// Get the table properties.
    ///
    /// Returns the table-level formatting properties if available.
    pub fn properties(&self) -> Option<&TableProperties> {
        self.properties.as_ref()
    }

    /// Get table justification (alignment).
    pub fn justification(&self) -> Option<TableJustification> {
        self.properties.as_ref().map(|p| p.justification)
    }

    /// Check if the first row is a header row.
    pub fn has_header_row(&self) -> bool {
        self.properties.as_ref().is_some_and(|p| p.is_header_row)
    }
}

/// A row in a table.
///
/// Represents a table row in the binary DOC format.
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around row data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Row {
    /// Cells in the row (shared via Arc for efficient cloning)
    cells: Arc<Vec<Cell>>,
    /// Row-level properties (if available)
    row_properties: Option<TableProperties>,
    /// Resolved tracked property revision metadata
    formatting_revision: Option<RevisionMark>,
    /// Whether properties preceding the revision wall are preserved
    properties_preserved_for_revision: bool,
}

impl Row {
    /// Create a new Row.
    #[allow(unused)]
    pub(crate) fn new(cells: Vec<Cell>) -> Self {
        Self {
            cells: Arc::new(cells),
            row_properties: None,
            formatting_revision: None,
            properties_preserved_for_revision: false,
        }
    }

    /// Create a new Row with properties.
    #[allow(unused)]
    pub(crate) fn with_properties(cells: Vec<Cell>, properties: TableProperties) -> Self {
        Self {
            cells: Arc::new(cells),
            formatting_revision: None,
            properties_preserved_for_revision: properties.properties_preserved_for_revision,
            row_properties: Some(properties),
        }
    }

    pub(crate) fn with_metadata(
        cells: Vec<Cell>,
        row_properties: Option<TableProperties>,
        formatting_revision: Option<RevisionMark>,
        properties_preserved_for_revision: bool,
    ) -> Self {
        Self {
            cells: Arc::new(cells),
            row_properties,
            formatting_revision,
            properties_preserved_for_revision,
        }
    }

    /// Get the number of cells in this row.
    pub fn cell_count(&self) -> Result<usize> {
        Ok(self.cells.len())
    }

    /// Get all cells in this row.
    ///
    /// Returns a cloned vector. Due to Rc-based sharing in Cell structures,
    /// cloning is relatively cheap (only increments reference counts).
    pub fn cells(&self) -> Result<Vec<Cell>> {
        Ok((*self.cells).clone())
    }

    /// Get the row properties.
    pub fn properties(&self) -> Option<&TableProperties> {
        self.row_properties.as_ref()
    }

    /// Get the row height in twips (1/1440 inch).
    pub fn height(&self) -> Option<i16> {
        self.row_properties.as_ref().and_then(|p| p.row_height)
    }

    /// Check if this is a header row.
    pub fn is_header(&self) -> bool {
        self.row_properties
            .as_ref()
            .is_some_and(|p| p.is_header_row)
    }

    /// Tracked table-row property revision metadata.
    pub fn formatting_revision(&self) -> Option<&RevisionMark> {
        self.formatting_revision.as_ref()
    }

    /// Whether table properties preceding the revision wall are preserved.
    pub fn properties_preserved_for_revision(&self) -> bool {
        self.properties_preserved_for_revision
    }
}

/// A cell in a table.
///
/// Represents a table cell in the binary DOC format.
///
/// # Performance
///
/// Uses `Arc` for efficient cloning when passing around cell data.
/// Arc provides thread-safe reference counting, enabling Send + Sync.
#[derive(Debug, Clone)]
pub struct Cell {
    /// Cell content (text) - shared via Arc for efficient cloning
    text: Arc<String>,
    /// Cell content (paragraphs) - shared via Arc for efficient cloning
    paragraphs: Arc<Vec<Paragraph>>,
    /// Cell properties (if available)
    properties: Option<CellProperties>,
}

impl Cell {
    /// Create a new Cell.
    #[allow(unused)]
    pub(crate) fn new(text: String) -> Self {
        let para = Paragraph::new(text.clone());
        Self {
            text: Arc::new(text),
            paragraphs: Arc::new(vec![para]),
            properties: None,
        }
    }

    /// Create a new Cell with paragraphs and properties.
    #[allow(unused)]
    pub(crate) fn with_properties(
        paragraphs: Vec<Paragraph>,
        properties: Option<CellProperties>,
    ) -> Self {
        let text = paragraphs
            .iter()
            .filter_map(|p| p.text().ok())
            .collect::<Vec<&str>>()
            .join("\n");
        Self {
            text: Arc::new(text),
            paragraphs: Arc::new(paragraphs),
            properties,
        }
    }

    /// Get the text content of this cell.
    ///
    /// Concatenates all text from all paragraphs in the cell.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get all paragraphs in this cell.
    ///
    /// Returns a cloned vector. Cloning is relatively cheap due to Rc-based sharing.
    pub fn paragraphs(&self) -> Result<Vec<Paragraph>> {
        Ok((*self.paragraphs).clone())
    }

    /// Get the cell properties.
    pub fn properties(&self) -> Option<&CellProperties> {
        self.properties.as_ref()
    }

    /// Get the cell's vertical alignment.
    pub fn vertical_alignment(&self) -> Option<super::parts::tap::VerticalAlignment> {
        self.properties.as_ref().map(|p| p.vertical_alignment)
    }

    /// Get the cell's vertical merge state.
    pub fn vertical_merge_status(&self) -> Option<super::parts::tap::VerticalMergeStatus> {
        self.properties.as_ref().map(|p| p.vertical_merge_status)
    }

    /// Whether the cell stretches its contents to use the full width.
    pub fn fit_text(&self) -> Option<bool> {
        self.properties.as_ref().map(|p| p.fit_text)
    }

    /// Whether the cell prefers a single unwrapped line.
    pub fn no_wrap(&self) -> Option<bool> {
        self.properties.as_ref().map(|p| p.no_wrap)
    }

    /// Whether an otherwise empty row can hide this cell's mark.
    pub fn hide_mark(&self) -> Option<bool> {
        self.properties.as_ref().map(|p| p.hide_mark)
    }

    /// Get the cell's background color as RGB tuple.
    pub fn background_color(&self) -> Option<(u8, u8, u8)> {
        self.properties.as_ref().and_then(|p| p.background_color)
    }

    /// Get the cell's complete legacy shading descriptor.
    pub fn shading(&self) -> Option<super::parts::tap::CellShading> {
        self.properties.as_ref().and_then(|p| p.shading)
    }

    /// Whether the cell's raw `ShdNil` value defers shading to its table style.
    pub fn shading_inherits_from_style(&self) -> Option<bool> {
        self.properties
            .as_ref()
            .map(|p| p.shading_inherits_from_style)
    }

    /// Get the cell's top padding in twips.
    pub fn padding_top(&self) -> Option<i16> {
        self.properties.as_ref().and_then(|p| p.padding_top)
    }

    /// Get the cell's left padding in twips.
    pub fn padding_left(&self) -> Option<i16> {
        self.properties.as_ref().and_then(|p| p.padding_left)
    }

    /// Get the cell's bottom padding in twips.
    pub fn padding_bottom(&self) -> Option<i16> {
        self.properties.as_ref().and_then(|p| p.padding_bottom)
    }

    /// Get the cell's right padding in twips.
    pub fn padding_right(&self) -> Option<i16> {
        self.properties.as_ref().and_then(|p| p.padding_right)
    }
}

pub(crate) fn apply_table_cell_styles(rows: &mut [Row]) {
    let resolved = {
        let Some(properties) = rows
            .iter()
            .map(|row| row.row_properties.as_ref())
            .collect::<Option<Vec<_>>>()
        else {
            return;
        };
        properties
            .iter()
            .enumerate()
            .map(|(row_index, row)| {
                (0..row.cell_count)
                    .map(|cell_index| {
                        let logical_index = if row.right_to_left {
                            row.cell_count - 1 - cell_index
                        } else {
                            cell_index
                        };
                        let mut style = row.style_defaults;
                        style.border_top = if row_index == 0 {
                            row.style_defaults.border_top
                        } else {
                            row.style_defaults.border_inside_horizontal
                        };
                        style.border_bottom = if row_index + 1 == properties.len() {
                            row.style_defaults.border_bottom
                        } else {
                            row.style_defaults.border_inside_horizontal
                        };
                        style.border_left = if logical_index == 0 {
                            row.style_defaults.border_left
                        } else {
                            row.style_defaults.border_inside_vertical
                        };
                        style.border_right = if logical_index + 1 == row.cell_count {
                            row.style_defaults.border_right
                        } else {
                            row.style_defaults.border_inside_vertical
                        };
                        let conditions =
                            matching_table_style_conditions(&properties, row_index, cell_index);
                        for condition in conditions {
                            for conditional in row
                                .conditional_formats
                                .iter()
                                .filter(|format| format.condition == condition)
                            {
                                overlay_style_defaults(&mut style, conditional.properties);
                                if row_index + 1 < properties.len()
                                    && conditional.properties.border_inside_horizontal.is_some()
                                {
                                    style.border_bottom =
                                        conditional.properties.border_inside_horizontal;
                                }
                                if logical_index + 1 < row.cell_count
                                    && conditional.properties.border_inside_vertical.is_some()
                                {
                                    style.border_right =
                                        conditional.properties.border_inside_vertical;
                                }
                            }
                        }
                        style.horizontal_band_size = None;
                        style.vertical_band_size = None;
                        style.border_inside_horizontal = None;
                        style.border_inside_vertical = None;
                        style
                    })
                    .collect::<Vec<_>>()
            })
            .collect::<Vec<_>>()
    };

    for (row, styles) in rows.iter_mut().zip(resolved) {
        let cells = Arc::make_mut(&mut row.cells);
        for (cell, style) in cells.iter_mut().zip(styles) {
            if style == TableStyleDefaults::default() {
                continue;
            }
            let properties = cell.properties.get_or_insert_with(CellProperties::default);
            apply_style_defaults(properties, style);
        }
    }
}

fn overlay_style_defaults(target: &mut TableStyleDefaults, source: TableStyleDefaults) {
    macro_rules! overlay {
        ($($field:ident),+ $(,)?) => {
            $(if source.$field.is_some() { target.$field = source.$field; })+
        };
    }
    overlay!(
        padding_top,
        padding_left,
        padding_bottom,
        padding_right,
        vertical_alignment,
        no_wrap,
        horizontal_band_size,
        vertical_band_size,
        border_top,
        border_bottom,
        border_left,
        border_right,
        border_inside_horizontal,
        border_inside_vertical,
        border_diagonal_down,
        border_diagonal_up,
        shading,
    );
}

fn apply_style_defaults(cell: &mut CellProperties, style: TableStyleDefaults) {
    if !cell.direct_style.vertical_alignment {
        if let Some(value) = style.vertical_alignment {
            cell.vertical_alignment = value;
        }
    }
    if !cell.direct_style.no_wrap {
        if let Some(value) = style.no_wrap {
            cell.no_wrap = value;
        }
    }
    if !cell.direct_style.padding_top {
        cell.padding_top = style.padding_top.map(|value| value as i16);
    }
    if !cell.direct_style.padding_left {
        cell.padding_left = style.padding_left.map(|value| value as i16);
    }
    if !cell.direct_style.padding_bottom {
        cell.padding_bottom = style.padding_bottom.map(|value| value as i16);
    }
    if !cell.direct_style.padding_right {
        cell.padding_right = style.padding_right.map(|value| value as i16);
    }
    if !cell.direct_style.shading || cell.shading_inherits_from_style {
        match style.shading {
            Some(TableStyleShading::NoShading) => {
                cell.shading = None;
                cell.background_color = None;
                cell.shading_inherits_from_style = false;
            },
            Some(TableStyleShading::Shading(shading)) => {
                cell.shading = Some(shading);
                cell.background_color = shading.background_color;
                cell.shading_inherits_from_style = false;
            },
            None => {},
        }
    }
    apply_style_border(
        &mut cell.borders.top,
        style.border_top,
        cell.direct_style.border_top,
        cell.border_type_overrides.top,
    );
    apply_style_border(
        &mut cell.borders.left,
        style.border_left,
        cell.direct_style.border_left,
        cell.border_type_overrides.left,
    );
    apply_style_border(
        &mut cell.borders.bottom,
        style.border_bottom,
        cell.direct_style.border_bottom,
        cell.border_type_overrides.bottom,
    );
    apply_style_border(
        &mut cell.borders.right,
        style.border_right,
        cell.direct_style.border_right,
        cell.border_type_overrides.right,
    );
    apply_style_border(
        &mut cell.borders.diagonal_down,
        style.border_diagonal_down,
        cell.direct_style.border_diagonal_down,
        None,
    );
    apply_style_border(
        &mut cell.borders.diagonal_up,
        style.border_diagonal_up,
        cell.direct_style.border_diagonal_up,
        None,
    );
}

fn apply_style_border(
    target: &mut Option<super::parts::tap::BorderStyle>,
    style: Option<TableStyleBorder>,
    direct: bool,
    border_type: Option<super::parts::tap::BorderType>,
) {
    if !direct {
        if let Some(style) = style {
            *target = match style {
                TableStyleBorder::NoBorder => None,
                TableStyleBorder::Border(border) => Some(border),
            };
        }
    }
    if let (Some(target), Some(border_type)) = (target, border_type) {
        target.border_type = border_type;
    }
}

#[cfg(test)]
mod tests {
    use super::super::parts::tap::{
        BorderStyle, BorderType, CellDirectStyle, CellShading, ShadingPattern,
        TableConditionalFormatting, TableLook, TableLookFlags, TableStyleCondition,
    };
    use super::*;

    #[test]
    fn test_cell_text() {
        let cell = Cell::new("Cell content".to_string());
        assert_eq!(cell.text().unwrap(), "Cell content");
    }

    #[test]
    fn test_row_cell_count() {
        let cells = vec![
            Cell::new("A".to_string()),
            Cell::new("B".to_string()),
            Cell::new("C".to_string()),
        ];
        let row = Row::new(cells);
        assert_eq!(row.cell_count().unwrap(), 3);
    }

    #[test]
    fn test_table_dimensions() {
        let row1 = Row::new(vec![Cell::new("A".to_string()), Cell::new("B".to_string())]);
        let row2 = Row::new(vec![Cell::new("C".to_string()), Cell::new("D".to_string())]);
        let table = Table::new(vec![row1, row2]);

        assert_eq!(table.row_count().unwrap(), 2);
        assert_eq!(table.column_count().unwrap(), 2);
    }

    #[test]
    fn exposes_table_row_revision_metadata() {
        let revision = RevisionMark {
            kind: super::super::revision::RevisionKind::Formatting,
            author_index: 1,
            author: "Editor".to_string(),
            timestamp: None,
            reason: None,
            revision_id: None,
            revision_save_id: None,
        };
        let row = Row::with_metadata(vec![Cell::new("A".to_string())], None, Some(revision), true);
        assert_eq!(row.formatting_revision().unwrap().author, "Editor");
        assert!(row.properties_preserved_for_revision());
    }

    #[test]
    fn applies_table_style_cells_with_conditions_and_direct_overrides() {
        let border = |width| BorderStyle {
            width,
            color: None,
            border_type: BorderType::Single,
            spacing: 0,
            shadow: false,
            frame: false,
        };
        let shading = |color| CellShading {
            foreground_color: None,
            background_color: Some(color),
            pattern: ShadingPattern::Solid,
        };
        let row_properties = || TableProperties {
            cell_count: 2,
            style_defaults: TableStyleDefaults {
                padding_top: Some(12),
                no_wrap: Some(true),
                border_top: Some(TableStyleBorder::Border(border(1))),
                border_bottom: Some(TableStyleBorder::Border(border(2))),
                border_inside_horizontal: Some(TableStyleBorder::Border(border(3))),
                shading: Some(TableStyleShading::Shading(shading((0, 0, 255)))),
                ..TableStyleDefaults::default()
            },
            conditional_formats: vec![TableConditionalFormatting {
                condition: TableStyleCondition::HeaderRow,
                properties: TableStyleDefaults {
                    padding_left: Some(24),
                    border_inside_horizontal: Some(TableStyleBorder::Border(border(4))),
                    border_inside_vertical: Some(TableStyleBorder::Border(border(5))),
                    shading: Some(TableStyleShading::Shading(shading((255, 0, 0)))),
                    ..TableStyleDefaults::default()
                },
                raw_grpprl: Vec::new(),
            }],
            table_look: Some(TableLook {
                autoformat_index: -1,
                flags: TableLookFlags::HEADER_ROW,
            }),
            ..TableProperties::default()
        };
        let direct = CellProperties {
            no_wrap: false,
            padding_top: Some(99),
            borders: super::super::parts::tap::CellBorders {
                top: Some(border(9)),
                ..super::super::parts::tap::CellBorders::default()
            },
            direct_style: CellDirectStyle {
                no_wrap: true,
                padding_top: true,
                border_top: true,
                ..CellDirectStyle::default()
            },
            ..CellProperties::default()
        };
        let styled_cell = || Cell::with_properties(Vec::new(), Some(CellProperties::default()));
        let mut rows = vec![
            Row::with_properties(
                vec![
                    Cell::with_properties(Vec::new(), Some(direct)),
                    styled_cell(),
                ],
                row_properties(),
            ),
            Row::with_properties(vec![styled_cell(), styled_cell()], row_properties()),
        ];

        apply_table_cell_styles(&mut rows);
        let first = &rows[0].cells[0].properties.as_ref().unwrap();
        assert!(!first.no_wrap);
        assert_eq!(first.padding_top, Some(99));
        assert_eq!(first.padding_left, Some(24));
        assert_eq!(first.background_color, Some((255, 0, 0)));
        assert_eq!(first.borders.top.unwrap().width, 9);
        assert_eq!(first.borders.bottom.unwrap().width, 4);
        assert_eq!(first.borders.right.unwrap().width, 5);

        let top_second = rows[0].cells[1].properties.as_ref().unwrap();
        assert!(top_second.no_wrap);
        assert_eq!(top_second.padding_top, Some(12));
        assert_eq!(top_second.background_color, Some((255, 0, 0)));
        assert_eq!(top_second.borders.top.unwrap().width, 1);
        assert_eq!(top_second.borders.bottom.unwrap().width, 4);
        assert!(top_second.borders.right.is_none());

        let bottom = rows[1].cells[0].properties.as_ref().unwrap();
        assert_eq!(bottom.background_color, Some((0, 0, 255)));
        assert_eq!(bottom.borders.top.unwrap().width, 3);
        assert_eq!(bottom.borders.bottom.unwrap().width, 2);
    }
}
