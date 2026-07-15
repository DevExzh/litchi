//! TAP (Table Properties) generation for DOC files
//!
//! TAP structures define table layout, borders, and cell properties.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's TableProperties.

use super::sprm::SprmBuilder;
use crate::doc::parts::tap::{
    BorderStyle, BorderType, CellBorders, CellShading, TableLook, TableWidth, TextDirection,
    VerticalAlignment, VerticalMergeStatus, WidthType,
};

/// Error returned when table row properties cannot be represented in DOC TAP.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum TapBuildError {
    /// Requested row index is not present in the builder.
    RowOutOfBounds(usize),
    /// DOC table rows can contain at most 63 cells.
    InvalidCellCount(usize),
    /// Cumulative cell boundaries exceed the DOC XAS coordinate range.
    CellWidthsOverflow,
    /// DOC row heights use the YAS range of -31680 through 31680 twips.
    InvalidRowHeight(i16),
    /// A merge continuation cannot occur in the first cell.
    MergeWithoutPrecedingCell,
    /// Brc80 spacing is a five-bit value.
    InvalidBorderSpacing(u8),
    /// DOC cell padding cannot exceed 22 inches.
    InvalidCellPadding(u16),
    /// A preferred-width property uses unsupported units or a value outside its context's range.
    InvalidPreferredWidth(&'static str, TableWidth),
    /// TLP contains bits outside the eleven-bit Fatl field.
    InvalidTableLookFlags(u16),
}

impl std::fmt::Display for TapBuildError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::RowOutOfBounds(index) => write!(f, "table row {index} does not exist"),
            Self::InvalidCellCount(count) => {
                write!(
                    f,
                    "DOC table rows must contain between 1 and 63 cells, got {count}"
                )
            },
            Self::CellWidthsOverflow => {
                write!(
                    f,
                    "DOC cell widths exceed the 31680-twip XAS coordinate range"
                )
            },
            Self::InvalidRowHeight(height) => {
                write!(f, "DOC row height {height} is outside the YAS range")
            },
            Self::MergeWithoutPrecedingCell => {
                write!(f, "the first DOC table cell cannot be a merge continuation")
            },
            Self::InvalidBorderSpacing(spacing) => {
                write!(f, "DOC Brc80 spacing {spacing} exceeds 31 points")
            },
            Self::InvalidCellPadding(padding) => {
                write!(f, "DOC cell padding {padding} exceeds 31680 twips")
            },
            Self::InvalidPreferredWidth(property, width) => {
                write!(f, "DOC {property} has an invalid preferred width {width:?}")
            },
            Self::InvalidTableLookFlags(flags) => {
                write!(f, "DOC table look contains reserved flags {flags:#06x}")
            },
        }
    }
}

impl std::error::Error for TapBuildError {}

/// Table cell descriptor
#[derive(Debug, Clone, Default)]
pub struct TableCell {
    /// Cell width (in twips)
    pub width: u16,
    /// This cell is merged into the preceding cell. The preceding cell is
    /// automatically encoded as the start of the horizontal merge.
    pub merged: bool,
    /// Vertical merge state for this cell
    pub vertical_merge: VerticalMergeStatus,
    /// Vertical alignment of cell contents
    pub vertical_alignment: VerticalAlignment,
    /// Cell text flow and rotation
    pub text_direction: TextDirection,
    /// Stretch contents to use the full cell width
    pub fit_text: bool,
    /// Prefer cell contents on a single unwrapped line
    pub no_wrap: bool,
    /// Hide the cell mark when every cell in the row is empty
    pub hide_mark: bool,
    /// Cell edge borders
    pub borders: CellBorders,
    /// Complete legacy cell shading
    pub shading: Option<CellShading>,
    /// Cell padding in twips
    pub padding_top: Option<u16>,
    pub padding_left: Option<u16>,
    pub padding_bottom: Option<u16>,
    pub padding_right: Option<u16>,
}

/// Default borders for a DOC table row.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct TableBorders {
    pub top: Option<BorderStyle>,
    pub left: Option<BorderStyle>,
    pub bottom: Option<BorderStyle>,
    pub right: Option<BorderStyle>,
    pub horizontal: Option<BorderStyle>,
    pub vertical: Option<BorderStyle>,
}

/// Table row properties
#[derive(Debug, Clone)]
pub struct TableRow {
    /// Cells in this row
    pub cells: Vec<TableCell>,
    /// Row height in twips (positive = at least, negative = exact, zero = auto)
    pub height: i16,
    /// Header row flag
    pub is_header: bool,
    /// Whether the row may split across page breaks
    pub allow_break: bool,
    /// Preferred total table width
    pub preferred_width: Option<TableWidth>,
    /// Automatically resize columns to fit table contents
    pub auto_fit: bool,
    /// Preferred leading width before the first cell
    pub width_before: Option<TableWidth>,
    /// Preferred trailing width after the last cell
    pub width_after: Option<TableWidth>,
    /// Preferred leading indentation of the table
    pub preferred_indent: Option<TableWidth>,
    /// Avoid a page break between this row and the next row
    pub keep_with_next: bool,
    /// Table auto-format identity and optional look flags
    pub table_look: Option<TableLook>,
    /// Style-sheet index of the applied table style
    pub table_style_index: Option<u16>,
    /// Lay out the table from right to left
    pub right_to_left: bool,
    /// Whether this floating table may overlap other tables
    pub allow_overlap: bool,
    /// Default outer and inside borders for this row
    pub borders: TableBorders,
}

impl Default for TableRow {
    fn default() -> Self {
        Self {
            cells: Vec::new(),
            height: 0,
            is_header: false,
            allow_break: true,
            preferred_width: None,
            auto_fit: false,
            width_before: None,
            width_after: None,
            preferred_indent: None,
            keep_with_next: false,
            table_look: None,
            table_style_index: None,
            right_to_left: false,
            allow_overlap: true,
            borders: TableBorders::default(),
        }
    }
}

#[derive(Debug, Clone, Copy)]
enum PreferredWidthUsage {
    Table,
    TablePart,
    Indent,
}

fn encode_preferred_width(
    property: &'static str,
    width: Option<TableWidth>,
    usage: PreferredWidthUsage,
) -> Result<Option<[u8; 3]>, TapBuildError> {
    let Some(width) = width else {
        return Ok(None);
    };
    let units = match width.width_type {
        WidthType::Auto if width.value == 0 => 1,
        WidthType::Percentage
            if matches!(usage, PreferredWidthUsage::Table)
                && (0..=30_000).contains(&width.value) =>
        {
            2
        },
        WidthType::Percentage
            if matches!(usage, PreferredWidthUsage::TablePart)
                && (0..=5_000).contains(&width.value) =>
        {
            2
        },
        WidthType::Twips
            if matches!(
                usage,
                PreferredWidthUsage::Table | PreferredWidthUsage::TablePart
            ) && (0..=31_680).contains(&width.value) =>
        {
            3
        },
        WidthType::Twips
            if matches!(usage, PreferredWidthUsage::Indent)
                && (-31_560..=31_680).contains(&width.value) =>
        {
            3
        },
        _ => return Err(TapBuildError::InvalidPreferredWidth(property, width)),
    };
    let value = width.value.to_le_bytes();
    Ok(Some([units, value[0], value[1]]))
}

/// TAP (Table Properties) builder
#[derive(Debug)]
pub struct TapBuilder {
    rows: Vec<TableRow>,
}

impl TapBuilder {
    /// Create a new TAP builder
    pub fn new() -> Self {
        Self { rows: Vec::new() }
    }

    /// Add a row to the table
    pub fn add_row(&mut self, row: TableRow) {
        self.rows.push(row);
    }

    /// Generate TAP SPRMs for a specific row
    pub fn generate_row_sprms(&self, row_index: usize) -> Vec<u8> {
        self.try_generate_row_sprms(row_index).unwrap_or_default()
    }

    /// Generate validated TAP SPRMs for a specific row.
    pub fn try_generate_row_sprms(&self, row_index: usize) -> Result<Vec<u8>, TapBuildError> {
        let row = self
            .rows
            .get(row_index)
            .ok_or(TapBuildError::RowOutOfBounds(row_index))?;
        generate_row_sprms(row)
    }

    /// Borrow the configured rows.
    pub fn rows(&self) -> &[TableRow] {
        &self.rows
    }

    /// Get the number of rows
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }
}

pub(crate) fn generate_row_sprms(row: &TableRow) -> Result<Vec<u8>, TapBuildError> {
    let cell_count = row.cells.len();
    if !(1..=63).contains(&cell_count) {
        return Err(TapBuildError::InvalidCellCount(cell_count));
    }
    if row.cells.first().is_some_and(|cell| cell.merged) {
        return Err(TapBuildError::MergeWithoutPrecedingCell);
    }
    if !(-31_680..=31_680).contains(&row.height) {
        return Err(TapBuildError::InvalidRowHeight(row.height));
    }
    let effective_widths = if row.cells.iter().all(|cell| cell.width == 0) {
        const DEFAULT_TABLE_WIDTH: u32 = 8640;
        (0..cell_count)
            .map(|index| {
                let left = DEFAULT_TABLE_WIDTH * index as u32 / cell_count as u32;
                let right = DEFAULT_TABLE_WIDTH * (index + 1) as u32 / cell_count as u32;
                (right - left) as u16
            })
            .collect::<Vec<_>>()
    } else {
        row.cells.iter().map(|cell| cell.width).collect()
    };

    let mut boundaries = Vec::with_capacity(cell_count + 1);
    boundaries.push(0i16);
    let mut boundary = 0u32;
    for width in &effective_widths {
        boundary = boundary
            .checked_add(u32::from(*width))
            .ok_or(TapBuildError::CellWidthsOverflow)?;
        if boundary > 31_680 {
            return Err(TapBuildError::CellWidthsOverflow);
        }
        boundaries.push(boundary as i16);
    }

    let preferred_width = encode_preferred_width(
        "table width",
        row.preferred_width,
        PreferredWidthUsage::Table,
    )?;
    let width_before = encode_preferred_width(
        "leading table-part width",
        row.width_before,
        PreferredWidthUsage::TablePart,
    )?;
    let width_after = encode_preferred_width(
        "trailing table-part width",
        row.width_after,
        PreferredWidthUsage::TablePart,
    )?;
    let preferred_indent = encode_preferred_width(
        "table indent",
        row.preferred_indent,
        PreferredWidthUsage::Indent,
    )?;
    if let Some(TableWidth {
        value: indent,
        width_type: WidthType::Twips,
    }) = row.preferred_indent
    {
        let layout_width = match row.preferred_width {
            Some(TableWidth {
                value,
                width_type: WidthType::Twips,
            }) => i32::from(value),
            _ => boundary as i32,
        };
        if i32::from(indent) + layout_width > 31_680 {
            return Err(TapBuildError::InvalidPreferredWidth(
                "table indent",
                TableWidth {
                    value: indent,
                    width_type: WidthType::Twips,
                },
            ));
        }
    }

    let mut builder = SprmBuilder::new();
    // Apply the style first so later SPRMs remain direct row formatting.
    if let Some(style_index) = row.table_style_index {
        builder.add_word(0x563A, style_index);
    }
    if !row.allow_break
        || row
            .cells
            .iter()
            .any(|cell| cell.merged || cell.vertical_merge != VerticalMergeStatus::None)
    {
        // Emit the legacy form first for older readers, followed by the
        // authoritative modern form as required for equivalent SPRMs.
        builder.add_bool(0x3403, true);
        builder.add_bool(0x3466, true);
    }
    if row.is_header {
        builder.add_bool(0x3404, true);
    }
    if row.height != 0 {
        builder.add_signed_word(0x9407, row.height);
    }
    if let Some(width) = preferred_width {
        builder.add_three_byte(0xF614, width);
    }
    if row.auto_fit {
        builder.add_bool(0x3615, true);
    }
    if let Some(width) = width_before {
        builder.add_three_byte(0xF617, width);
    }
    if let Some(width) = width_after {
        builder.add_three_byte(0xF618, width);
    }
    if row.keep_with_next {
        builder.add_bool(0x3619, true);
    }
    if let Some(width) = preferred_indent {
        builder.add_three_byte(0xF661, width);
    }
    if let Some(look) = row.table_look {
        let flags = look.flags.bits();
        if flags & !0x07FF != 0 {
            return Err(TapBuildError::InvalidTableLookFlags(flags));
        }
        builder.add_dword(
            0x740A,
            u32::from_le_bytes([
                look.autoformat_index.to_le_bytes()[0],
                look.autoformat_index.to_le_bytes()[1],
                flags.to_le_bytes()[0],
                flags.to_le_bytes()[1],
            ]),
        );
    }
    if row.right_to_left {
        builder.add_word(0x560B, 1);
        builder.add_word(0x5664, 1);
    }
    if !row.allow_overlap {
        builder.add_bool(0x3465, true);
    }
    let mut sprms = builder.build();

    let mut operand = Vec::with_capacity(1 + (cell_count + 1) * 2 + cell_count * 20);
    operand.push(cell_count as u8);
    for boundary in boundaries {
        operand.extend_from_slice(&boundary.to_le_bytes());
    }
    for (index, width) in effective_widths.into_iter().enumerate() {
        let horizontal_merge = if row.cells[index].merged {
            1
        } else if row.cells.get(index + 1).is_some_and(|cell| cell.merged) {
            2
        } else {
            0
        };
        let text_flow = match row.cells[index].text_direction {
            TextDirection::LrTb => 0,
            TextDirection::TbRl => 1,
            TextDirection::BtLr => 3,
            TextDirection::LrBt => 4,
            TextDirection::TbLr => 5,
        };
        let vertical_merge = match row.cells[index].vertical_merge {
            VerticalMergeStatus::None => 0,
            VerticalMergeStatus::Merged => 1,
            VerticalMergeStatus::First => 3,
        };
        let vertical_alignment = match row.cells[index].vertical_alignment {
            VerticalAlignment::Top => 0,
            VerticalAlignment::Center => 1,
            VerticalAlignment::Bottom => 2,
        };
        let mut flags = horizontal_merge
            | (text_flow << 2)
            | (vertical_merge << 5)
            | (vertical_alignment << 7)
            | (3u16 << 9); // ftsDxa
        if row.cells[index].fit_text {
            flags |= 0x1000;
        }
        if row.cells[index].no_wrap {
            flags |= 0x2000;
        }
        if row.cells[index].hide_mark {
            flags |= 0x4000;
        }
        operand.extend_from_slice(&flags.to_le_bytes());
        operand.extend_from_slice(&width.to_le_bytes());
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.top)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.left)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.bottom)?);
        operand.extend_from_slice(&encode_border80_fallback(row.cells[index].borders.right)?);
    }

    let encoded_size =
        u16::try_from(operand.len() + 1).map_err(|_| TapBuildError::CellWidthsOverflow)?;
    sprms.extend_from_slice(&0xD608u16.to_le_bytes());
    sprms.extend_from_slice(&encoded_size.to_le_bytes());
    sprms.extend_from_slice(&operand);

    append_table_borders(&mut sprms, row.borders)?;
    append_cell_borders(&mut sprms, &row.cells)?;
    append_cell_shading(&mut sprms, &row.cells)?;
    append_cell_padding(&mut sprms, &row.cells)?;
    Ok(sprms)
}

fn append_cell_shading(sprms: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    if !cells.iter().any(|cell| cell.shading.is_some()) {
        return Ok(());
    }
    let last_shaded = cells
        .iter()
        .rposition(|cell| cell.shading.is_some())
        .expect("at least one shaded cell was checked above");
    let legacy_cells = &cells[..=last_shaded];
    let legacy = legacy_cells
        .iter()
        .map(|cell| encode_shading80(cell.shading))
        .collect::<Option<Vec<_>>>();
    if let Some(descriptors) = legacy {
        sprms.extend_from_slice(&0xD609u16.to_le_bytes());
        sprms.push((legacy_cells.len() * 2) as u8);
        for descriptor in descriptors {
            sprms.extend_from_slice(&descriptor.to_le_bytes());
        }
    }

    append_full_shading_chunks(sprms, cells, false);
    append_full_shading_chunks(sprms, cells, true);
    Ok(())
}

fn append_full_shading_chunks(sprms: &mut Vec<u8>, cells: &[TableCell], raw: bool) {
    for (chunk_index, chunk) in cells.chunks(22).enumerate() {
        let Some(last_shaded) = chunk.iter().rposition(|cell| cell.shading.is_some()) else {
            continue;
        };
        let chunk = &chunk[..=last_shaded];
        let opcode = match chunk_index {
            0 if raw => 0xD670u16,
            1 if raw => 0xD671u16,
            2 if raw => 0xD672u16,
            0 => 0xD612u16,
            1 => 0xD616u16,
            2 => 0xD60Cu16,
            _ => unreachable!("DOC rows contain at most 63 cells"),
        };
        sprms.extend_from_slice(&opcode.to_le_bytes());
        sprms.push((chunk.len() * 10) as u8);
        for cell in chunk {
            append_shading(sprms, cell.shading, raw);
        }
    }
}

fn encode_shading80(shading: Option<CellShading>) -> Option<u16> {
    let Some(shading) = shading else {
        return Some(u16::MAX);
    };
    let foreground = rgb_to_ico(shading.foreground_color).ok()?;
    let background = rgb_to_ico(shading.background_color).ok()?;
    Some(u16::from(foreground) | (u16::from(background) << 5) | ((shading.pattern as u16) << 10))
}

fn append_shading(output: &mut Vec<u8>, shading: Option<CellShading>, raw: bool) {
    let Some(shading) = shading else {
        if raw {
            output.extend_from_slice(&[0, 0, 0, 0xFF, 0, 0, 0, 0xFF]);
        } else {
            output.extend_from_slice(&[0xFF; 8]);
        }
        output.extend_from_slice(&0u16.to_le_bytes());
        return;
    };
    append_colorref(output, shading.foreground_color);
    append_colorref(output, shading.background_color);
    output.extend_from_slice(&(shading.pattern as u16).to_le_bytes());
}

fn append_colorref(output: &mut Vec<u8>, color: Option<(u8, u8, u8)>) {
    match color {
        Some((red, green, blue)) => output.extend_from_slice(&[red, green, blue, 0]),
        None => output.extend_from_slice(&[0, 0, 0, 0xFF]),
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct PaddingRun {
    first: u8,
    limit: u8,
    sides: u8,
    width: u16,
}

type PaddingGetter = fn(&TableCell) -> Option<u16>;

fn append_cell_padding(sprms: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    let sides: [(u8, PaddingGetter); 4] = [
        (0x01, |cell| cell.padding_top),
        (0x02, |cell| cell.padding_left),
        (0x04, |cell| cell.padding_bottom),
        (0x08, |cell| cell.padding_right),
    ];
    let mut runs = Vec::<PaddingRun>::new();
    for (side, get_width) in sides {
        let mut first = 0;
        while first < cells.len() {
            let Some(width) = get_width(&cells[first]) else {
                first += 1;
                continue;
            };
            if width > 31_680 {
                return Err(TapBuildError::InvalidCellPadding(width));
            }
            let mut limit = first + 1;
            while limit < cells.len() && get_width(&cells[limit]) == Some(width) {
                limit += 1;
            }
            if let Some(existing) = runs.iter_mut().find(|run| {
                run.first == first as u8 && run.limit == limit as u8 && run.width == width
            }) {
                existing.sides |= side;
            } else {
                runs.push(PaddingRun {
                    first: first as u8,
                    limit: limit as u8,
                    sides: side,
                    width,
                });
            }
            first = limit;
        }
    }

    for run in runs {
        sprms.extend_from_slice(&0xD632u16.to_le_bytes());
        sprms.push(6);
        sprms.extend_from_slice(&[run.first, run.limit, run.sides, 0x03]);
        sprms.extend_from_slice(&run.width.to_le_bytes());
    }
    Ok(())
}

fn encode_border80_fallback(border: Option<BorderStyle>) -> Result<[u8; 4], TapBuildError> {
    let Some(border) = border else {
        return Ok([0; 4]);
    };
    if border.border_type == BorderType::None {
        return Ok([0; 4]);
    }
    if border.spacing > 31 {
        return Err(TapBuildError::InvalidBorderSpacing(border.spacing));
    }
    let border_type = border_type_code(border.border_type);
    if matches!(border.border_type, BorderType::Outset | BorderType::Inset) {
        return Ok([0; 4]);
    }
    let ico = rgb_to_ico(border.color).unwrap_or(0);
    let effects = border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6);
    Ok([border.width, border_type, ico, effects])
}

fn border_type_code(border_type: BorderType) -> u8 {
    match border_type {
        BorderType::None => 0,
        BorderType::Single => 1,
        BorderType::Thick => 5,
        BorderType::Double => 3,
        BorderType::Dotted => 6,
        BorderType::Dashed => 7,
        BorderType::DotDash => 8,
        BorderType::DotDotDash => 9,
        BorderType::Triple => 10,
        BorderType::ThinThickSmall => 11,
        BorderType::ThickThinSmall => 12,
        BorderType::ThinThickThinSmall => 13,
        BorderType::ThinThickMedium => 14,
        BorderType::ThickThinMedium => 15,
        BorderType::ThinThickThinMedium => 16,
        BorderType::ThinThickLarge => 17,
        BorderType::ThickThinLarge => 18,
        BorderType::ThinThickThinLarge => 19,
        BorderType::Wave => 20,
        BorderType::DoubleWave => 21,
        BorderType::DashSmall => 22,
        BorderType::DashDotStroked => 23,
        BorderType::Emboss => 24,
        BorderType::Engrave => 25,
        BorderType::Outset => 26,
        BorderType::Inset => 27,
    }
}

fn append_full_border(
    output: &mut Vec<u8>,
    border: Option<BorderStyle>,
    nil: bool,
) -> Result<(), TapBuildError> {
    let Some(border) = border else {
        if nil {
            output.extend_from_slice(&[0; 4]);
            output.extend_from_slice(&[0xFF; 4]);
        } else {
            output.extend_from_slice(&[0; 8]);
        }
        return Ok(());
    };
    if border.border_type == BorderType::None {
        if nil {
            output.extend_from_slice(&[0; 4]);
            output.extend_from_slice(&[0xFF; 4]);
        } else {
            output.extend_from_slice(&[0; 8]);
        }
        return Ok(());
    }
    if border.spacing > 31 {
        return Err(TapBuildError::InvalidBorderSpacing(border.spacing));
    }
    append_colorref(output, border.color);
    output.push(border.width);
    output.push(border_type_code(border.border_type));
    output.push(border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6));
    output.push(0);
    Ok(())
}

fn append_table_borders(output: &mut Vec<u8>, borders: TableBorders) -> Result<(), TapBuildError> {
    let values = [
        borders.top,
        borders.left,
        borders.bottom,
        borders.right,
        borders.horizontal,
        borders.vertical,
    ];
    if values.iter().all(Option::is_none) {
        return Ok(());
    }
    if let Some(legacy) = values
        .iter()
        .map(|border| encode_border80_exact(*border))
        .collect::<Option<Vec<_>>>()
    {
        output.extend_from_slice(&0xD605u16.to_le_bytes());
        output.push(24);
        for border in legacy {
            output.extend_from_slice(&border);
        }
    }
    output.extend_from_slice(&0xD613u16.to_le_bytes());
    output.push(48);
    for border in values {
        append_full_border(output, border, false)?;
    }
    Ok(())
}

fn encode_border80_exact(border: Option<BorderStyle>) -> Option<[u8; 4]> {
    let Some(border) = border else {
        return Some([0; 4]);
    };
    if border.border_type == BorderType::None {
        return Some([0; 4]);
    }
    if border.spacing > 31 || matches!(border.border_type, BorderType::Outset | BorderType::Inset) {
        return None;
    }
    let ico = rgb_to_ico(border.color).ok()?;
    Some([
        border.width,
        border_type_code(border.border_type),
        ico,
        border.spacing | (u8::from(border.shadow) << 5) | (u8::from(border.frame) << 6),
    ])
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BorderRun {
    first: u8,
    limit: u8,
    sides: u8,
    border: BorderStyle,
}

type BorderGetter = fn(&TableCell) -> Option<BorderStyle>;

fn append_cell_borders(output: &mut Vec<u8>, cells: &[TableCell]) -> Result<(), TapBuildError> {
    let sides: [(u8, BorderGetter); 6] = [
        (0x01, |cell| cell.borders.top),
        (0x02, |cell| cell.borders.left),
        (0x04, |cell| cell.borders.bottom),
        (0x08, |cell| cell.borders.right),
        (0x10, |cell| cell.borders.diagonal_down),
        (0x20, |cell| cell.borders.diagonal_up),
    ];
    let mut runs = Vec::<BorderRun>::new();
    for (side, get_border) in sides {
        let mut first = 0;
        while first < cells.len() {
            let Some(border) = get_border(&cells[first]) else {
                first += 1;
                continue;
            };
            let mut limit = first + 1;
            while limit < cells.len() && get_border(&cells[limit]) == Some(border) {
                limit += 1;
            }
            if let Some(run) = runs.iter_mut().find(|run| {
                run.first == first as u8 && run.limit == limit as u8 && run.border == border
            }) {
                run.sides |= side;
            } else {
                runs.push(BorderRun {
                    first: first as u8,
                    limit: limit as u8,
                    sides: side,
                    border,
                });
            }
            first = limit;
        }
    }
    for run in runs {
        output.extend_from_slice(&0xD62Fu16.to_le_bytes());
        output.push(11);
        output.extend_from_slice(&[run.first, run.limit, run.sides]);
        append_full_border(output, Some(run.border), true)?;
    }
    Ok(())
}

fn rgb_to_ico(color: Option<(u8, u8, u8)>) -> Result<u8, (u8, u8, u8)> {
    Ok(match color {
        None => 0,
        Some((0, 0, 0)) => 1,
        Some((0, 0, 255)) => 2,
        Some((0, 255, 255)) => 3,
        Some((0, 255, 0)) => 4,
        Some((255, 0, 255)) => 5,
        Some((255, 0, 0)) => 6,
        Some((255, 255, 0)) => 7,
        Some((255, 255, 255)) => 8,
        Some((0, 0, 128)) => 9,
        Some((0, 128, 128)) => 10,
        Some((0, 128, 0)) => 11,
        Some((128, 0, 128)) => 12,
        Some((128, 0, 0)) => 13,
        Some((128, 128, 0)) => 14,
        Some((128, 128, 128)) => 15,
        Some((192, 192, 192)) => 16,
        Some(color) => return Err(color),
    })
}

impl Default for TapBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Helper to create a simple table
pub fn create_simple_table(rows: usize, cols: usize, cell_width: u16) -> TapBuilder {
    let mut builder = TapBuilder::new();

    for _ in 0..rows {
        let cells = vec![
            TableCell {
                width: cell_width,
                merged: false,
                ..TableCell::default()
            };
            cols
        ];
        builder.add_row(TableRow {
            cells,
            ..TableRow::default()
        });
    }

    builder
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::parts::tap::{ShadingPattern, TableLookFlags};

    #[test]
    fn test_tap_builder() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![
                TableCell {
                    width: 1000,
                    merged: false,
                    vertical_merge: VerticalMergeStatus::First,
                    vertical_alignment: VerticalAlignment::Center,
                    text_direction: TextDirection::TbRl,
                    fit_text: true,
                    no_wrap: true,
                    hide_mark: true,
                    borders: CellBorders {
                        top: Some(BorderStyle {
                            width: 8,
                            color: Some((1, 2, 3)),
                            border_type: BorderType::Single,
                            spacing: 2,
                            shadow: true,
                            frame: false,
                        }),
                        diagonal_down: Some(BorderStyle {
                            width: 4,
                            color: Some((10, 20, 30)),
                            border_type: BorderType::Outset,
                            spacing: 1,
                            shadow: false,
                            frame: true,
                        }),
                        ..CellBorders::default()
                    },
                    shading: Some(CellShading {
                        foreground_color: Some((0, 0, 255)),
                        background_color: Some((255, 255, 0)),
                        pattern: ShadingPattern::DarkCross,
                    }),
                    padding_top: Some(120),
                    padding_left: Some(240),
                    padding_bottom: Some(120),
                    padding_right: Some(240),
                },
                TableCell {
                    width: 1000,
                    merged: false,
                    ..TableCell::default()
                },
            ],
            height: -200,
            is_header: true,
            allow_break: false,
            borders: TableBorders {
                vertical: Some(BorderStyle {
                    width: 6,
                    color: Some((40, 50, 60)),
                    border_type: BorderType::Double,
                    spacing: 0,
                    shadow: false,
                    frame: false,
                }),
                ..TableBorders::default()
            },
            ..TableRow::default()
        });

        let sprms = builder.try_generate_row_sprms(0).unwrap();
        let opcodes = crate::sprm::parse_sprms(&sprms)
            .into_iter()
            .map(|sprm| sprm.opcode)
            .collect::<Vec<_>>();
        let legacy_cant_split = opcodes.iter().position(|opcode| *opcode == 0x3403).unwrap();
        let modern_cant_split = opcodes.iter().position(|opcode| *opcode == 0x3466).unwrap();
        assert!(legacy_cant_split < modern_cant_split);
        let tap = crate::doc::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
        assert_eq!(tap.cell_boundaries, [0, 1000, 2000]);
        assert_eq!(tap.row_height, Some(-200));
        assert!(tap.is_header_row);
        assert!(!tap.allow_row_break);
        assert_eq!(
            tap.cell_properties[0].vertical_merge_status,
            VerticalMergeStatus::First
        );
        assert_eq!(
            tap.cell_properties[0].vertical_alignment,
            VerticalAlignment::Center
        );
        assert_eq!(tap.cell_properties[0].text_direction, TextDirection::TbRl);
        assert!(tap.cell_properties[0].fit_text);
        assert!(tap.cell_properties[0].no_wrap);
        assert!(tap.cell_properties[0].hide_mark);
        let border = tap.cell_properties[0].borders.top.unwrap();
        assert_eq!(border.color, Some((1, 2, 3)));
        assert_eq!(border.spacing, 2);
        assert!(border.shadow);
        let diagonal = tap.cell_properties[0].borders.diagonal_down.unwrap();
        assert_eq!(diagonal.color, Some((10, 20, 30)));
        assert_eq!(diagonal.border_type, BorderType::Outset);
        assert!(diagonal.frame);
        assert_eq!(tap.border_vertical.unwrap().color, Some((40, 50, 60)));
        assert_eq!(
            tap.cell_properties[0].shading,
            Some(CellShading {
                foreground_color: Some((0, 0, 255)),
                background_color: Some((255, 255, 0)),
                pattern: ShadingPattern::DarkCross,
            })
        );
        assert_eq!(tap.cell_properties[0].padding_top, Some(120));
        assert_eq!(tap.cell_properties[0].padding_left, Some(240));
        assert_eq!(tap.cell_properties[0].padding_bottom, Some(120));
        assert_eq!(tap.cell_properties[0].padding_right, Some(240));
        assert_eq!(
            tap.cell_properties[0].preferred_width.unwrap().width_type,
            crate::doc::parts::tap::WidthType::Twips
        );
    }

    #[test]
    fn test_tap_builder_empty() {
        let builder = TapBuilder::new();
        assert_eq!(builder.row_count(), 0);
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::RowOutOfBounds(0))
        );
    }

    #[test]
    fn writes_full_color_shading_across_all_cell_chunks() {
        let shading = CellShading {
            foreground_color: Some((1, 2, 3)),
            background_color: Some((250, 240, 230)),
            pattern: ShadingPattern::Percent42Point5,
        };
        let mut cells = vec![TableCell::default(); 45];
        cells[0].shading = Some(shading);
        cells[22].shading = Some(shading);
        cells[44].shading = Some(shading);
        let row = TableRow {
            cells,
            ..TableRow::default()
        };

        let sprms = generate_row_sprms(&row).unwrap();
        let opcodes = crate::sprm::parse_sprms(&sprms)
            .into_iter()
            .map(|sprm| sprm.opcode)
            .collect::<Vec<_>>();
        assert!(!opcodes.contains(&0xD609));
        assert!(opcodes.contains(&0xD612));
        assert!(opcodes.contains(&0xD616));
        assert!(opcodes.contains(&0xD60C));
        assert!(opcodes.contains(&0xD670));
        assert!(opcodes.contains(&0xD671));
        assert!(opcodes.contains(&0xD672));

        let tap = crate::doc::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
        assert_eq!(tap.cell_properties[0].shading, Some(shading));
        assert_eq!(tap.cell_properties[22].shading, Some(shading));
        assert_eq!(tap.cell_properties[44].shading, Some(shading));
        assert!(tap.cell_properties[43].shading.is_none());
    }

    #[test]
    fn test_tap_builder_multiple_rows() {
        let mut builder = TapBuilder::new();
        for i in 0..5 {
            builder.add_row(TableRow {
                cells: vec![
                    TableCell {
                        width: 1000,
                        merged: false,
                        ..TableCell::default()
                    },
                    TableCell {
                        width: 1000,
                        merged: false,
                        ..TableCell::default()
                    },
                    TableCell {
                        width: 1000,
                        merged: false,
                        ..TableCell::default()
                    },
                ],
                height: 200 + (i as i16 * 50),
                is_header: i == 0,
                ..TableRow::default()
            });
        }

        assert_eq!(builder.row_count(), 5);
        let sprms = builder.generate_row_sprms(0);
        assert!(!sprms.is_empty());
    }

    #[test]
    fn round_trips_table_sizing_and_fit_properties() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![
                TableCell {
                    width: 1000,
                    ..TableCell::default()
                },
                TableCell {
                    width: 1000,
                    ..TableCell::default()
                },
            ],
            preferred_width: Some(TableWidth {
                value: 7_500,
                width_type: WidthType::Percentage,
            }),
            auto_fit: true,
            width_before: Some(TableWidth {
                value: 250,
                width_type: WidthType::Percentage,
            }),
            width_after: Some(TableWidth {
                value: 400,
                width_type: WidthType::Twips,
            }),
            preferred_indent: Some(TableWidth {
                value: -120,
                width_type: WidthType::Twips,
            }),
            keep_with_next: true,
            table_look: Some(TableLook {
                autoformat_index: -1,
                flags: TableLookFlags::BORDERS
                    | TableLookFlags::HEADER_COLUMN
                    | TableLookFlags::NO_COLUMN_BANDING,
            }),
            table_style_index: Some(0x1234),
            right_to_left: true,
            allow_overlap: false,
            ..TableRow::default()
        });

        let sprms = builder.try_generate_row_sprms(0).unwrap();
        let opcodes = crate::sprm::parse_sprms(&sprms)
            .into_iter()
            .map(|sprm| sprm.opcode)
            .collect::<Vec<_>>();
        assert_eq!(opcodes[0], 0x563A);
        assert!(
            opcodes.iter().position(|opcode| *opcode == 0x560B)
                < opcodes.iter().position(|opcode| *opcode == 0x5664)
        );
        let tap = crate::doc::parts::tap::TableProperties::from_sprm(&sprms).unwrap();
        assert_eq!(tap.preferred_width, builder.rows()[0].preferred_width);
        assert!(tap.auto_fit);
        assert_eq!(tap.width_before, builder.rows()[0].width_before);
        assert_eq!(tap.width_after, builder.rows()[0].width_after);
        assert_eq!(tap.preferred_indent, builder.rows()[0].preferred_indent);
        assert!(tap.keep_with_next);
        assert_eq!(tap.table_look, builder.rows()[0].table_look);
        assert_eq!(tap.table_style_index, Some(0x1234));
        assert!(tap.right_to_left);
        assert!(!tap.allow_overlap);
    }

    #[test]
    fn test_create_simple_table() {
        let table = create_simple_table(3, 4, 1440); // 3 rows, 4 cols, 1 inch cells
        assert_eq!(table.row_count(), 3);
    }

    #[test]
    fn test_create_simple_table_single_cell() {
        let table = create_simple_table(1, 1, 1000);
        assert_eq!(table.row_count(), 1);
        assert_eq!(table.rows[0].cells.len(), 1);
    }

    #[test]
    fn test_create_simple_table_large() {
        let table = create_simple_table(10, 10, 500);
        assert_eq!(table.row_count(), 10);
        assert_eq!(table.rows[0].cells.len(), 10);
    }

    #[test]
    fn test_table_row_count() {
        let table = create_simple_table(5, 3, 1000);
        assert_eq!(table.row_count(), 5);
    }

    #[test]
    fn rejects_unrepresentable_rows() {
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                merged: true,
                ..TableCell::default()
            }],
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::MergeWithoutPrecedingCell)
        );

        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: u16::MAX,
                merged: false,
                ..TableCell::default()
            }],
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::CellWidthsOverflow)
        );

        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                ..TableCell::default()
            }],
            height: i16::MIN,
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::InvalidRowHeight(i16::MIN))
        );

        let invalid_width = TableWidth {
            value: 30_001,
            width_type: WidthType::Percentage,
        };
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                ..TableCell::default()
            }],
            preferred_width: Some(invalid_width),
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::InvalidPreferredWidth(
                "table width",
                invalid_width
            ))
        );

        let invalid_indent = TableWidth {
            value: 31_000,
            width_type: WidthType::Twips,
        };
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                ..TableCell::default()
            }],
            preferred_width: Some(TableWidth {
                value: 1_000,
                width_type: WidthType::Twips,
            }),
            preferred_indent: Some(invalid_indent),
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::InvalidPreferredWidth(
                "table indent",
                invalid_indent
            ))
        );

        let invalid_flags = 0x8000;
        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                ..TableCell::default()
            }],
            table_look: Some(TableLook {
                autoformat_index: 0,
                flags: TableLookFlags::from_bits_retain(invalid_flags),
            }),
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::InvalidTableLookFlags(invalid_flags))
        );

        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                borders: CellBorders {
                    top: Some(BorderStyle {
                        width: 8,
                        color: Some((1, 2, 3)),
                        border_type: BorderType::Single,
                        spacing: 0,
                        shadow: false,
                        frame: false,
                    }),
                    ..CellBorders::default()
                },
                ..TableCell::default()
            }],
            ..TableRow::default()
        });
        assert!(builder.try_generate_row_sprms(0).is_ok());

        let mut builder = TapBuilder::new();
        builder.add_row(TableRow {
            cells: vec![TableCell {
                width: 1000,
                padding_left: Some(31_681),
                ..TableCell::default()
            }],
            ..TableRow::default()
        });
        assert_eq!(
            builder.try_generate_row_sprms(0),
            Err(TapBuildError::InvalidCellPadding(31_681))
        );
    }
}
