//! Direct table, row and cell properties: the children of a table's own
//! `w:tblPr` and `w:tblGrid`, a row's own `w:trPr`, and a cell's own
//! `w:tcPr`.

use quick_xml::events::BytesStart;

use super::paragraph::TextDirection;
use super::run::{toggle, value};
use super::scan::{self, Context, Site};
use super::value::{
    Border, Shading, Sign, Twips, Unit, hex_short, measure, number, on_off, percent_fiftieths,
    unsigned, vocabulary,
};
use super::{Invalid, collect, optional, parse_border, parse_shading, position};
use crate::error::Result;

/// A table, cell or margin width (`CT_TblWidth`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Width {
    /// `nil`: no width.
    Nil,
    /// `auto`: determined by the layout.
    Auto,
    /// `dxa`, the default when `w:type` is absent: a length.
    Twips(Twips),
    /// `pct`: fiftieths of a percent, so 5000 is 100%. A width written as a
    /// percentage (`"50%"`) is converted to fiftieths.
    Percent(i64),
}

vocabulary! {
    /// Table and row alignment (`ST_JcTable`). `left` and `right` are the
    /// transitional spellings of `start` and `end`.
    pub enum TableJustification {
        /// `center`
        Center = "center",
        /// `end`
        End = "end",
        /// `left`
        Left = "left",
        /// `right`
        Right = "right",
        /// `start`
        Start = "start",
    }
}

vocabulary! {
    /// The table layout algorithm (`ST_TblLayoutType`).
    pub enum TableLayout {
        /// `fixed`: column widths come from the grid and widths.
        Fixed = "fixed",
        /// `autofit`: columns fit their content.
        Autofit = "autofit",
    }
}

vocabulary! {
    /// A cell's vertical alignment (`ST_VerticalJc`).
    pub enum VerticalJc {
        /// `top`
        Top = "top",
        /// `center`
        Center = "center",
        /// `both`
        Both = "both",
        /// `bottom`
        Bottom = "bottom",
    }
}

vocabulary! {
    /// A cell merge marker (`ST_Merge`); a marker without `w:val` continues.
    pub enum Merge {
        /// `continue`: this cell continues the merge before it.
        Continue = "continue",
        /// `restart`: this cell starts a merged region.
        Restart = "restart",
    }
}

vocabulary! {
    /// How a row height is applied (`ST_HeightRule`).
    pub enum HeightRule {
        /// `auto`: the content decides.
        Auto = "auto",
        /// `exact`
        Exact = "exact",
        /// `atLeast`
        AtLeast = "atLeast",
    }
}

/// A row height (`w:trHeight`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Height {
    /// `w:val`; an omitted value is zero.
    pub value: Twips,
    /// `w:hRule`; an omitted rule is `auto`.
    pub rule: Option<HeightRule>,
}

impl Height {
    /// A row height of `value` with `rule` (`None`: no `w:hRule`, which
    /// reads as `auto`).
    ///
    /// ```
    /// use litchi_docx::properties::{Height, HeightRule, Twips};
    ///
    /// let height = Height::new(Twips(400), Some(HeightRule::Exact));
    /// assert_eq!(height.value, Twips(400));
    /// ```
    #[must_use]
    pub const fn new(value: Twips, rule: Option<HeightRule>) -> Self {
        Self { value, rule }
    }
}

/// The conditional table-style formatting a table applies (`w:tblLook`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TableLook {
    /// `w:val`: the transitional bit mask.
    pub value: Option<u16>,
    /// `w:firstRow`
    pub first_row: Option<bool>,
    /// `w:lastRow`
    pub last_row: Option<bool>,
    /// `w:firstColumn`
    pub first_column: Option<bool>,
    /// `w:lastColumn`
    pub last_column: Option<bool>,
    /// `w:noHBand`: no horizontal banding.
    pub no_horizontal_band: Option<bool>,
    /// `w:noVBand`: no vertical banding.
    pub no_vertical_band: Option<bool>,
}

impl TableLook {
    /// The look as the transitional bit mask: `w:val` when stated, otherwise
    /// assembled from the on/off attributes (first row `0x0020`, last row
    /// `0x0040`, first column `0x0080`, last column `0x0100`, no horizontal
    /// band `0x0200`, no vertical band `0x0400`).
    ///
    /// ```
    /// use litchi_docx::Table;
    ///
    /// let table = Table::new(
    ///     br#"<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblPr><w:tblLook w:firstRow="1" w:noVBand="1"/></w:tblPr></w:tbl>"#.to_vec(),
    /// );
    /// let look = table.properties()?.look.unwrap();
    /// assert_eq!(look.mask(), 0x0420);
    /// # Ok::<(), litchi_docx::Error>(())
    /// ```
    #[must_use]
    pub fn mask(&self) -> u16 {
        if let Some(value) = self.value {
            return value;
        }
        [
            (self.first_row, 0x0020),
            (self.last_row, 0x0040),
            (self.first_column, 0x0080),
            (self.last_column, 0x0100),
            (self.no_horizontal_band, 0x0200),
            (self.no_vertical_band, 0x0400),
        ]
        .into_iter()
        .filter(|(flag, _)| *flag == Some(true))
        .fold(0, |mask, (_, bit)| mask | bit)
    }
}

/// Direct table borders (`w:tblBorders`). `start`/`end` and the
/// transitional `left`/`right` are kept apart as authored.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct TableBorders {
    /// `w:top`
    pub top: Option<Border>,
    /// `w:start`
    pub start: Option<Border>,
    /// `w:left`
    pub left: Option<Border>,
    /// `w:bottom`
    pub bottom: Option<Border>,
    /// `w:end`
    pub end: Option<Border>,
    /// `w:right`
    pub right: Option<Border>,
    /// `w:insideH`: between rows.
    pub inside_horizontal: Option<Border>,
    /// `w:insideV`: between columns.
    pub inside_vertical: Option<Border>,
}

/// Direct cell borders (`w:tcBorders`): the table borders plus the
/// diagonals. `start`/`end` and `left`/`right` are kept apart as authored.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct CellBorders {
    /// `w:top`
    pub top: Option<Border>,
    /// `w:start`
    pub start: Option<Border>,
    /// `w:left`
    pub left: Option<Border>,
    /// `w:bottom`
    pub bottom: Option<Border>,
    /// `w:end`
    pub end: Option<Border>,
    /// `w:right`
    pub right: Option<Border>,
    /// `w:insideH`
    pub inside_horizontal: Option<Border>,
    /// `w:insideV`
    pub inside_vertical: Option<Border>,
    /// `w:tl2br`: the diagonal from the top-left corner.
    pub top_left_to_bottom_right: Option<Border>,
    /// `w:tr2bl`: the diagonal from the top-right corner.
    pub top_right_to_bottom_left: Option<Border>,
}

/// Cell margins (`w:tblCellMar` of a table, `w:tcMar` of a cell).
/// `start`/`end` and `left`/`right` are kept apart as authored.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub struct Margins {
    /// `w:top`
    pub top: Option<Width>,
    /// `w:start`
    pub start: Option<Width>,
    /// `w:left`
    pub left: Option<Width>,
    /// `w:bottom`
    pub bottom: Option<Width>,
    /// `w:end`
    pub end: Option<Width>,
    /// `w:right`
    pub right: Option<Width>,
}

/// The table-property elements this reader types. A property's position
/// here is its bit in [`Table::invalid_properties`].
const TABLE_ELEMENTS: [&str; 16] = [
    "tblStyle",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "gridCol",
];

const _: () = assert!(
    TABLE_ELEMENTS.len() <= 64,
    "an Invalid mask holds 64 properties"
);

const TABLE_BORDERS: usize = position(&TABLE_ELEMENTS, "tblBorders");
const TABLE_CELL_MARGINS: usize = position(&TABLE_ELEMENTS, "tblCellMar");
const GRID_COLUMN: usize = position(&TABLE_ELEMENTS, "gridCol");

/// Direct (local) table properties: what one table's own `w:tblPr` and
/// `w:tblGrid` state.
///
/// Read with [`crate::Table::properties`]. Every field is exactly what the
/// table authors; `None` means the table does not state the property, which
/// is then inherited from the table style and the document defaults. The
/// properties of tables nested in its cells, a row's table-property
/// exceptions (`w:tblPrEx`) and the previous formatting of a tracked change
/// (`w:tblPrChange`) are not read. Floating-table positioning (`w:tblpPr`,
/// `w:tblOverlap`) is not typed yet.
///
/// ```
/// use litchi_docx::Table;
/// use litchi_docx::properties::{TableJustification, Twips, Width};
///
/// let table = Table::new(
///     br#"<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/></w:tblPr><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="5760"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>"#.to_vec(),
/// );
/// let properties = table.properties()?;
/// assert_eq!(properties.style.as_deref(), Some("Grid"));
/// assert_eq!(properties.width, Some(Width::Percent(5000)));
/// assert_eq!(properties.justification, Some(TableJustification::Center));
/// assert_eq!(properties.grid, Some(vec![Some(Twips(2880)), Some(Twips(5760))]));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Table {
    /// `w:tblStyle`: the table style identifier.
    pub style: Option<String>,
    /// `w:bidiVisual`: a right-to-left table.
    pub bidi_visual: Option<bool>,
    /// `w:tblStyleRowBandSize`: rows per band of a banded style.
    pub row_band_size: Option<u32>,
    /// `w:tblStyleColBandSize`: columns per band of a banded style.
    pub column_band_size: Option<u32>,
    /// `w:tblW`
    pub width: Option<Width>,
    /// `w:jc`
    pub justification: Option<TableJustification>,
    /// `w:tblCellSpacing`
    pub cell_spacing: Option<Width>,
    /// `w:tblInd`: the indent from the leading margin.
    pub indent: Option<Width>,
    /// `w:tblBorders`
    pub borders: Option<TableBorders>,
    /// `w:shd`
    pub shading: Option<Shading>,
    /// `w:tblLayout`
    pub layout: Option<TableLayout>,
    /// `w:tblCellMar`: the default cell margins.
    pub cell_margins: Option<Margins>,
    /// `w:tblLook`
    pub look: Option<TableLook>,
    /// `w:tblCaption`: the accessible caption.
    pub caption: Option<String>,
    /// `w:tblDescription`: the accessible description.
    pub description: Option<String>,
    /// The `w:gridCol` widths of `w:tblGrid`, in column order; `None` when
    /// the table has no grid column. A column without `w:w` is `None`.
    pub grid: Option<Vec<Option<Twips>>>,
    pub(super) invalid: Invalid,
}

impl Table {
    /// The direct properties the table states with a value outside the
    /// property's schema domain, by element local name in schema order
    /// (`gridCol` for an invalid grid width).
    ///
    /// Each such property is left unset rather than guessed; a border or
    /// margin that cannot be typed is left out of its group.
    pub fn invalid_properties(&self) -> impl Iterator<Item = &'static str> + '_ {
        self.invalid.names(&TABLE_ELEMENTS)
    }
}

/// The row-property elements this reader types. A property's position here
/// is its bit in [`Row::invalid_properties`].
const ROW_ELEMENTS: [&str; 10] = [
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
];

const _: () = assert!(
    ROW_ELEMENTS.len() <= 64,
    "an Invalid mask holds 64 properties"
);

/// Direct (local) row properties: what one row's own `w:trPr` states.
///
/// Read with [`crate::Row::properties`]. The row's table-property
/// exceptions (`w:tblPrEx`), its tracked insertion or deletion and the
/// previous formatting of a tracked change (`w:trPrChange`) are not read.
///
/// ```
/// use litchi_docx::Row;
/// use litchi_docx::properties::{HeightRule, Twips};
///
/// let row = Row::new(
///     br#"<w:tr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:trPr><w:tblHeader/><w:trHeight w:val="400" w:hRule="exact"/></w:trPr><w:tc><w:p/></w:tc></w:tr>"#.to_vec(),
/// );
/// let properties = row.properties()?;
/// assert_eq!(properties.header, Some(true));
/// let height = properties.height.unwrap();
/// assert_eq!((height.value, height.rule), (Twips(400), Some(HeightRule::Exact)));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Row {
    /// `w:gridBefore`: grid columns skipped before the first cell.
    pub grid_before: Option<u32>,
    /// `w:gridAfter`: grid columns skipped after the last cell.
    pub grid_after: Option<u32>,
    /// `w:wBefore`: the preferred width of the skipped leading columns.
    pub width_before: Option<Width>,
    /// `w:wAfter`: the preferred width of the skipped trailing columns.
    pub width_after: Option<Width>,
    /// `w:cantSplit`: keep the row on one page.
    pub cant_split: Option<bool>,
    /// `w:trHeight`
    pub height: Option<Height>,
    /// `w:tblHeader`: repeat the row as a header on every page.
    pub header: Option<bool>,
    /// `w:tblCellSpacing`
    pub cell_spacing: Option<Width>,
    /// `w:jc`
    pub justification: Option<TableJustification>,
    /// `w:hidden`: hide the row's end mark.
    pub hidden: Option<bool>,
    pub(super) invalid: Invalid,
}

impl Row {
    /// The direct properties the row states with a value outside the
    /// property's schema domain, by element local name in schema order.
    pub fn invalid_properties(&self) -> impl Iterator<Item = &'static str> + '_ {
        self.invalid.names(&ROW_ELEMENTS)
    }
}

/// The cell-property elements this reader types. A property's position here
/// is its bit in [`Cell::invalid_properties`].
const CELL_ELEMENTS: [&str; 12] = [
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
];

const _: () = assert!(
    CELL_ELEMENTS.len() <= 64,
    "an Invalid mask holds 64 properties"
);

const CELL_BORDERS: usize = position(&CELL_ELEMENTS, "tcBorders");
const CELL_MARGINS: usize = position(&CELL_ELEMENTS, "tcMar");

/// Direct (local) cell properties: what one cell's own `w:tcPr` states.
///
/// Read with [`crate::Cell::properties`]. The `w:tcPr` of a cell in a table
/// nested inside this cell is never read, nor the previous formatting of a
/// tracked change (`w:tcPrChange`) or a tracked cell insertion, deletion or
/// merge.
///
/// ```
/// use litchi_docx::Cell;
/// use litchi_docx::properties::{Merge, VerticalJc};
///
/// let cell = Cell::new(
///     br#"<w:tc xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/><w:vAlign w:val="center"/></w:tcPr><w:p/></w:tc>"#.to_vec(),
/// );
/// let properties = cell.properties()?;
/// assert_eq!(properties.grid_span, Some(2));
/// assert_eq!(properties.vertical_merge, Some(Merge::Restart));
/// assert_eq!(properties.vertical_align, Some(VerticalJc::Center));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Debug, Clone, Default, PartialEq, Eq)]
#[non_exhaustive]
pub struct Cell {
    /// `w:tcW`
    pub width: Option<Width>,
    /// `w:gridSpan`: grid columns spanned, at least 1.
    pub grid_span: Option<u32>,
    /// `w:hMerge` (transitional).
    pub horizontal_merge: Option<Merge>,
    /// `w:vMerge`
    pub vertical_merge: Option<Merge>,
    /// `w:tcBorders`
    pub borders: Option<CellBorders>,
    /// `w:shd`
    pub shading: Option<Shading>,
    /// `w:noWrap`
    pub no_wrap: Option<bool>,
    /// `w:tcMar`
    pub margins: Option<Margins>,
    /// `w:textDirection`
    pub text_direction: Option<TextDirection>,
    /// `w:tcFitText`
    pub fit_text: Option<bool>,
    /// `w:vAlign`
    pub vertical_align: Option<VerticalJc>,
    /// `w:hideMark`: ignore the end-of-cell mark in the row height.
    pub hide_mark: Option<bool>,
    invalid: Invalid,
}

impl Cell {
    /// The direct properties the cell states with a value outside the
    /// property's schema domain, by element local name in schema order.
    ///
    /// Each such property is left unset rather than guessed; a border or
    /// margin that cannot be typed is left out of its group.
    pub fn invalid_properties(&self) -> impl Iterator<Item = &'static str> + '_ {
        self.invalid.names(&CELL_ELEMENTS)
    }
}

/// Store a parsed value; `None` (invalid) leaves the field unset.
fn set<T>(slot: &mut Option<T>, value: Option<T>) -> bool {
    match value {
        Some(value) => {
            *slot = Some(value);
            true
        },
        None => false,
    }
}

/// Store a parsed value unless the field is already set.
fn fill<T>(slot: &mut Option<T>, value: Option<T>) -> bool {
    match value {
        Some(value) => {
            if slot.is_none() {
                *slot = Some(value);
            }
            true
        },
        None => false,
    }
}

/// `CT_TblWidth`; `None` when the type or width is invalid.
fn parse_width(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<Width>> {
    let [width, kind] = collect(context, element, [b"w", b"type"])?;
    Ok((|| match kind.as_deref() {
        None | Some("dxa") => Some(Width::Twips(Twips(match width.as_deref() {
            None => 0,
            Some(raw) => measure(raw, Unit::Twips, Sign::Signed)?,
        }))),
        Some("pct") => Some(Width::Percent(match width.as_deref() {
            None => 0,
            Some(raw) if raw.trim_end().ends_with('%') => percent_fiftieths(raw)?,
            Some(raw) => number(raw)?,
        })),
        Some("auto") => Some(Width::Auto),
        Some("nil") => Some(Width::Nil),
        Some(_) => None,
    })())
}

/// A required `w:val` string.
fn text(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<String>> {
    let [value] = collect(context, element, [b"val"])?;
    Ok(value.map(std::borrow::Cow::into_owned))
}

/// `CT_HMerge`/`CT_VMerge`: a marker without `w:val` continues.
fn parse_merge(context: &Context<'_>, element: &BytesStart<'_>) -> Result<Option<Merge>> {
    let [value] = collect(context, element, [b"val"])?;
    Ok(value.map_or(Some(Merge::Continue), |raw| Merge::from_xml(&raw)))
}

/// Collects one table's direct properties; repeated elements are read from
/// their first occurrence.
#[derive(Default)]
struct TableReader {
    table: Table,
    seen: u64,
    first: bool,
}

impl TableReader {
    fn property(
        &mut self,
        context: &Context<'_>,
        container: usize,
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if container == 1 {
            // `w:tblGrid`
            self.first = false;
            if local != b"gridCol" {
                return Ok(());
            }
            let [width] = collect(context, element, [b"w"])?;
            let column = optional(width.as_deref(), |raw| {
                measure(raw, Unit::Twips, Sign::Unsigned).map(Twips)
            });
            let grid = self.table.grid.get_or_insert_with(Vec::new);
            match column {
                Some(column) => grid.push(column),
                None => {
                    grid.push(None);
                    self.table.invalid.insert(GRID_COLUMN);
                },
            }
            return Ok(());
        }
        let Some(index) = TABLE_ELEMENTS[..GRID_COLUMN]
            .iter()
            .position(|name| name.as_bytes() == local)
        else {
            self.first = false;
            return Ok(());
        };
        let bit = 1u64 << index;
        let repeated = self.seen & bit != 0;
        self.seen |= bit;
        self.first = !repeated;
        if repeated {
            return Ok(());
        }
        let table = &mut self.table;
        let valid = match local {
            b"tblStyle" => set(&mut table.style, text(context, element)?),
            b"bidiVisual" => set(&mut table.bidi_visual, toggle(context, element)?),
            b"tblStyleRowBandSize" => {
                set(&mut table.row_band_size, value(context, element, unsigned)?)
            },
            b"tblStyleColBandSize" => set(
                &mut table.column_band_size,
                value(context, element, unsigned)?,
            ),
            b"tblW" => set(&mut table.width, parse_width(context, element)?),
            b"jc" => set(
                &mut table.justification,
                value(context, element, TableJustification::from_xml)?,
            ),
            b"tblCellSpacing" => set(&mut table.cell_spacing, parse_width(context, element)?),
            b"tblInd" => set(&mut table.indent, parse_width(context, element)?),
            b"tblBorders" => {
                table.borders = Some(TableBorders::default());
                true
            },
            b"shd" => set(&mut table.shading, parse_shading(context, element)?),
            b"tblLayout" => {
                let [kind] = collect(context, element, [b"type"])?;
                // `w:type` is optional; an omitted type is `autofit`.
                set(
                    &mut table.layout,
                    kind.map_or(Some(TableLayout::Autofit), |raw| {
                        TableLayout::from_xml(&raw)
                    }),
                )
            },
            b"tblCellMar" => {
                table.cell_margins = Some(Margins::default());
                true
            },
            b"tblLook" => {
                let [
                    value,
                    first_row,
                    last_row,
                    first_column,
                    last_column,
                    no_h,
                    no_v,
                ] = collect(
                    context,
                    element,
                    [
                        b"val",
                        b"firstRow",
                        b"lastRow",
                        b"firstColumn",
                        b"lastColumn",
                        b"noHBand",
                        b"noVBand",
                    ],
                )?;
                set(
                    &mut table.look,
                    (|| {
                        Some(TableLook {
                            value: optional(value.as_deref(), hex_short)?,
                            first_row: optional(first_row.as_deref(), on_off)?,
                            last_row: optional(last_row.as_deref(), on_off)?,
                            first_column: optional(first_column.as_deref(), on_off)?,
                            last_column: optional(last_column.as_deref(), on_off)?,
                            no_horizontal_band: optional(no_h.as_deref(), on_off)?,
                            no_vertical_band: optional(no_v.as_deref(), on_off)?,
                        })
                    })(),
                )
            },
            b"tblCaption" => set(&mut table.caption, text(context, element)?),
            b"tblDescription" => set(&mut table.description, text(context, element)?),
            _ => true,
        };
        if !valid {
            self.table.invalid.insert(index);
        }
        Ok(())
    }

    fn nested(
        &mut self,
        context: &Context<'_>,
        parent: &[u8],
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if !self.first {
            return Ok(());
        }
        let (index, valid) = match parent {
            b"tblBorders" => {
                let Some(borders) = self.table.borders.as_mut() else {
                    return Ok(());
                };
                let slot = match local {
                    b"top" => &mut borders.top,
                    b"start" => &mut borders.start,
                    b"left" => &mut borders.left,
                    b"bottom" => &mut borders.bottom,
                    b"end" => &mut borders.end,
                    b"right" => &mut borders.right,
                    b"insideH" => &mut borders.inside_horizontal,
                    b"insideV" => &mut borders.inside_vertical,
                    _ => return Ok(()),
                };
                (TABLE_BORDERS, fill(slot, parse_border(context, element)?))
            },
            b"tblCellMar" => {
                let Some(margins) = self.table.cell_margins.as_mut() else {
                    return Ok(());
                };
                let Some(slot) = margin_slot(margins, local) else {
                    return Ok(());
                };
                (
                    TABLE_CELL_MARGINS,
                    fill(slot, parse_width(context, element)?),
                )
            },
            _ => return Ok(()),
        };
        if !valid {
            self.table.invalid.insert(index);
        }
        Ok(())
    }
}

fn margin_slot<'m>(margins: &'m mut Margins, local: &[u8]) -> Option<&'m mut Option<Width>> {
    Some(match local {
        b"top" => &mut margins.top,
        b"start" => &mut margins.start,
        b"left" => &mut margins.left,
        b"bottom" => &mut margins.bottom,
        b"end" => &mut margins.end,
        b"right" => &mut margins.right,
        _ => return None,
    })
}

/// Read a table fragment's direct properties.
pub(crate) fn parse_table(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Table> {
    let mut reader = TableReader::default();
    scan::walk(
        xml,
        inherited,
        b"tbl",
        &[b"tblPr", b"tblGrid"],
        "table",
        |context, site, local, element| match site {
            Site::Property(container) => reader.property(context, container, local, element),
            Site::Nested(parent) => reader.nested(context, parent, local, element),
        },
    )?;
    Ok(reader.table)
}

/// Read a row fragment's direct properties; repeated elements are read from
/// their first occurrence.
pub(crate) fn parse_row(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Row> {
    let mut row = Row::default();
    let mut seen = 0u64;
    scan::walk(
        xml,
        inherited,
        b"tr",
        &[b"trPr"],
        "table row",
        |context, site, local, element| {
            let Site::Property(_) = site else {
                return Ok(());
            };
            let Some(index) = ROW_ELEMENTS
                .iter()
                .position(|name| name.as_bytes() == local)
            else {
                return Ok(());
            };
            let bit = 1u64 << index;
            if seen & bit != 0 {
                return Ok(());
            }
            seen |= bit;
            let valid = match local {
                b"gridBefore" => set(&mut row.grid_before, value(context, element, unsigned)?),
                b"gridAfter" => set(&mut row.grid_after, value(context, element, unsigned)?),
                b"wBefore" => set(&mut row.width_before, parse_width(context, element)?),
                b"wAfter" => set(&mut row.width_after, parse_width(context, element)?),
                b"cantSplit" => set(&mut row.cant_split, toggle(context, element)?),
                b"trHeight" => {
                    let [height, rule] = collect(context, element, [b"val", b"hRule"])?;
                    set(
                        &mut row.height,
                        (|| {
                            Some(Height {
                                value: Twips(match height.as_deref() {
                                    None => 0,
                                    Some(raw) => measure(raw, Unit::Twips, Sign::Unsigned)?,
                                }),
                                rule: optional(rule.as_deref(), HeightRule::from_xml)?,
                            })
                        })(),
                    )
                },
                b"tblHeader" => set(&mut row.header, toggle(context, element)?),
                b"tblCellSpacing" => set(&mut row.cell_spacing, parse_width(context, element)?),
                b"jc" => set(
                    &mut row.justification,
                    value(context, element, TableJustification::from_xml)?,
                ),
                b"hidden" => set(&mut row.hidden, toggle(context, element)?),
                _ => true,
            };
            if !valid {
                row.invalid.insert(index);
            }
            Ok(())
        },
    )?;
    Ok(row)
}

/// Collects one cell's direct properties; repeated elements are read from
/// their first occurrence.
#[derive(Default)]
struct CellReader {
    cell: Cell,
    seen: u64,
    first: bool,
}

impl CellReader {
    fn property(
        &mut self,
        context: &Context<'_>,
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        let Some(index) = CELL_ELEMENTS
            .iter()
            .position(|name| name.as_bytes() == local)
        else {
            self.first = false;
            return Ok(());
        };
        let bit = 1u64 << index;
        let repeated = self.seen & bit != 0;
        self.seen |= bit;
        self.first = !repeated;
        if repeated {
            return Ok(());
        }
        let cell = &mut self.cell;
        let valid = match local {
            b"tcW" => set(&mut cell.width, parse_width(context, element)?),
            b"gridSpan" => set(
                &mut cell.grid_span,
                value(context, element, |raw| {
                    unsigned(raw).filter(|span| *span > 0)
                })?,
            ),
            b"hMerge" => set(&mut cell.horizontal_merge, parse_merge(context, element)?),
            b"vMerge" => set(&mut cell.vertical_merge, parse_merge(context, element)?),
            b"tcBorders" => {
                cell.borders = Some(CellBorders::default());
                true
            },
            b"shd" => set(&mut cell.shading, parse_shading(context, element)?),
            b"noWrap" => set(&mut cell.no_wrap, toggle(context, element)?),
            b"tcMar" => {
                cell.margins = Some(Margins::default());
                true
            },
            b"textDirection" => set(
                &mut cell.text_direction,
                value(context, element, TextDirection::from_xml)?,
            ),
            b"tcFitText" => set(&mut cell.fit_text, toggle(context, element)?),
            b"vAlign" => set(
                &mut cell.vertical_align,
                value(context, element, VerticalJc::from_xml)?,
            ),
            b"hideMark" => set(&mut cell.hide_mark, toggle(context, element)?),
            _ => true,
        };
        if !valid {
            self.cell.invalid.insert(index);
        }
        Ok(())
    }

    fn nested(
        &mut self,
        context: &Context<'_>,
        parent: &[u8],
        local: &[u8],
        element: &BytesStart<'_>,
    ) -> Result<()> {
        if !self.first {
            return Ok(());
        }
        let (index, valid) = match parent {
            b"tcBorders" => {
                let Some(borders) = self.cell.borders.as_mut() else {
                    return Ok(());
                };
                let slot = match local {
                    b"top" => &mut borders.top,
                    b"start" => &mut borders.start,
                    b"left" => &mut borders.left,
                    b"bottom" => &mut borders.bottom,
                    b"end" => &mut borders.end,
                    b"right" => &mut borders.right,
                    b"insideH" => &mut borders.inside_horizontal,
                    b"insideV" => &mut borders.inside_vertical,
                    b"tl2br" => &mut borders.top_left_to_bottom_right,
                    b"tr2bl" => &mut borders.top_right_to_bottom_left,
                    _ => return Ok(()),
                };
                (CELL_BORDERS, fill(slot, parse_border(context, element)?))
            },
            b"tcMar" => {
                let Some(margins) = self.cell.margins.as_mut() else {
                    return Ok(());
                };
                let Some(slot) = margin_slot(margins, local) else {
                    return Ok(());
                };
                (CELL_MARGINS, fill(slot, parse_width(context, element)?))
            },
            _ => return Ok(()),
        };
        if !valid {
            self.cell.invalid.insert(index);
        }
        Ok(())
    }
}

/// Read a cell fragment's direct properties.
pub(crate) fn parse_cell(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Cell> {
    let mut reader = CellReader::default();
    scan::walk(
        xml,
        inherited,
        b"tc",
        &[b"tcPr"],
        "table cell",
        |context, site, local, element| match site {
            Site::Property(_) => reader.property(context, local, element),
            Site::Nested(parent) => reader.nested(context, parent, local, element),
        },
    )?;
    Ok(reader.cell)
}

impl std::fmt::Debug for Table {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Table").finish_non_exhaustive()
    }
}
