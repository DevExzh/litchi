//! The last main-document selection (`Selsf`, MS-DOC 2.9.244).
//!
//! Selection state is a passive Word UI cache. This module validates its fixed
//! fields and the `BlockSel`/`TableSel` union without applying, navigating, or
//! changing a document selection. The exact 36-byte source is retained for
//! callers that need lossless inert metadata; no selection is applied to a
//! Word UI and no OLE package is rewritten.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;

/// FIB `FibRgFcLcb` index for `fcWss`/`lcbWss`.
pub const FIB_INDEX_WSS: usize = 30;
/// Serialized size of one `Selsf` record.
pub const SELSF_SIZE: usize = 36;
const MAX_CP: i32 = i32::MAX;
const MIN_TABLE_EDGE: i16 = -31_680;
const MAX_TABLE_EDGE: i16 = 31_680;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid Selsf {field}: {error}")))
}

fn read_i16(data: &[u8], offset: usize, field: &str) -> Result<i16> {
    litchi_core::binary::read_i16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid Selsf {field}: {error}")))
}

fn read_i32(data: &[u8], offset: usize, field: &str) -> Result<i32> {
    litchi_core::binary::read_i32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid Selsf {field}: {error}")))
}

fn read_u32(data: &[u8], offset: usize, field: &str) -> Result<u32> {
    litchi_core::binary::read_u32_le(data, offset)
        .map_err(|error| corrupted(format!("invalid Selsf {field}: {error}")))
}

/// The `Sty` value embedded in `Selsf`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum SelectionStyle {
    /// The selection type is undefined and determined from `Selsf` flags.
    Undefined = 0,
    /// The selection contains characters, an inline picture, or a text frame.
    Character = 1,
    /// The selection contains one or more whole words.
    Word = 2,
    /// The selection is a sentence.
    Sentence = 3,
    /// The selection is a paragraph or table cell.
    Paragraph = 4,
    /// The selection contains one or more whole lines of text.
    Line = 5,
    /// The selection contains one or more whole table cells (`styCol`).
    Column = 0x000C,
    /// The selection contains one or more table rows.
    Row = 0x000D,
    /// The selection contains one or more table columns (`styColAll`).
    ColumnAll = 0x000E,
    /// The selection is the whole table (`styWholeTable`).
    WholeTable = 0x000F,
    /// The selection is a bullet or numbering character.
    Prefix = 0x001B,
}

impl SelectionStyle {
    fn from_raw(value: u16) -> Result<Self> {
        match value {
            0 => Ok(Self::Undefined),
            1 => Ok(Self::Character),
            2 => Ok(Self::Word),
            3 => Ok(Self::Sentence),
            4 => Ok(Self::Paragraph),
            5 => Ok(Self::Line),
            0x000C => Ok(Self::Column),
            0x000D => Ok(Self::Row),
            0x000E => Ok(Self::ColumnAll),
            0x000F => Ok(Self::WholeTable),
            0x001B => Ok(Self::Prefix),
            _ => Err(corrupted(format!(
                "Selsf sty has invalid value {value:#06x}"
            ))),
        }
    }
}

/// The interpretation of `Selsf.blktblSel` selected by `fTable`/`fBlock`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SelectionGeometry {
    /// No geometry is defined for an ordinary non-block text selection.
    None,
    /// A text block's physical left and right pixel boundaries.
    Block { first: i16, limit: i16 },
    /// A table row/cell selection's first and exclusive-limit cell indices.
    Table { first: u16, limit: u16 },
}

/// A validated `Selsf` record for the last main-document selection.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SavedSelection {
    source: [u8; SELSF_SIZE],
    flags: u16,
    direction_flags: u8,
    f_ins_end: u8,
    cp_first: u32,
    cp_lim: u32,
    raw_geometry: u32,
    cp_anchor: u32,
    style: SelectionStyle,
    cp_anchor_shrink: i32,
    xa_table_left: i16,
    xa_table_right: i16,
}

impl SavedSelection {
    /// Parse the optional `Selsf` selected by FIB index 30.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_WSS) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let length = usize::try_from(length)
            .map_err(|_| corrupted("Selsf length does not fit in memory"))?;
        if length != SELSF_SIZE {
            return Err(corrupted(format!(
                "Selsf must be {SELSF_SIZE} bytes, got {length}"
            )));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("Selsf offset does not fit in memory"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("Selsf range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("Selsf extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete 36-byte `Selsf` record.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let source: [u8; SELSF_SIZE] = data
            .try_into()
            .map_err(|_| corrupted(format!("Selsf must be {SELSF_SIZE} bytes")))?;
        let flags = read_u16(data, 0, "flags")?;
        if flags & (1 << 14) != 0 {
            return Err(corrupted("Selsf unused3 bit is set"));
        }
        let direction_flags = data
            .get(2)
            .copied()
            .ok_or_else(|| corrupted("Selsf direction flags are truncated"))?;
        let f_forward = direction_flags & 0x7F;
        if f_forward > 1 {
            return Err(corrupted("Selsf fForward is not 0 or 1"));
        }
        // fInsEnd is undefined for shape selections. Derive fShape before
        // inspecting the byte so that such records preserve the specified
        // ignore semantics, including otherwise out-of-domain values.
        let f_shape = flags & (1 << 8) != 0;
        let f_ins_end = data
            .get(3)
            .copied()
            .ok_or_else(|| corrupted("Selsf fInsEnd is truncated"))?;
        if !f_shape && f_ins_end > 1 {
            return Err(corrupted("Selsf fInsEnd is not 0 or 1"));
        }

        let cp_first = nonnegative_cp(read_i32(data, 4, "cpFirst")?, "cpFirst")?;
        let cp_lim = nonnegative_cp(read_i32(data, 8, "cpLim")?, "cpLim")?;
        if cp_lim < cp_first {
            return Err(corrupted("Selsf cpLim precedes cpFirst"));
        }
        let raw_geometry = read_u32(data, 16, "blktblSel")?;
        let cp_anchor = nonnegative_cp(read_i32(data, 20, "cpAnchor")?, "cpAnchor")?;
        if cp_anchor < cp_first {
            return Err(corrupted("Selsf cpAnchor precedes cpFirst"));
        }
        let style = SelectionStyle::from_raw(read_u16(data, 24, "sty")?)?;
        let cp_anchor_shrink = read_i32(data, 28, "cpAnchorShrink")?;
        let xa_table_left = read_i16(data, 32, "xaTableLeft")?;
        let xa_table_right = read_i16(data, 34, "xaTableRight")?;

        let f_ins = flags & (1 << 15) != 0;
        if f_ins && cp_first != cp_lim {
            return Err(corrupted("Selsf insertion point has unequal CPs"));
        }
        if !f_shape && f_ins_end != 0 && !f_ins {
            return Err(corrupted("Selsf fInsEnd requires fIns"));
        }

        let f_table = flags & (1 << 11) != 0;
        let f_block = flags & (1 << 13) != 0;
        let f_within_cell = flags & (1 << 2) != 0;
        let f_table_sel_non_shrink = flags & (1 << 4) != 0;
        let f_column = flags & (1 << 10) != 0;
        let f_graphics = flags & (1 << 12) != 0;
        if f_table && f_within_cell {
            return Err(corrupted(
                "Selsf fWithinCell must be clear for whole-cell selections",
            ));
        }
        if f_table_sel_non_shrink && !f_table {
            return Err(corrupted(
                "Selsf fTableSelNonShrink requires a table selection",
            ));
        }
        if f_graphics && f_shape {
            return Err(corrupted("Selsf fGraphics and fShape cannot both be set"));
        }
        if f_column && !f_table {
            return Err(corrupted("Selsf fColumn requires fTable"));
        }
        if f_table && !f_block && f_column {
            return Err(corrupted(
                "Selsf fColumn must be clear for whole-row selections",
            ));
        }
        if f_table && f_block && matches!(style, SelectionStyle::Row | SelectionStyle::WholeTable) {
            return Err(corrupted(
                "Selsf fBlock must be clear for row or whole-table selections",
            ));
        }
        let geometry = if f_table {
            let first = (raw_geometry & 0xFFFF) as u16;
            let limit = (raw_geometry >> 16) as u16;
            if first > 63 || limit > 64 || (!f_block && (first != 0 || limit != 64)) {
                return Err(corrupted("Selsf TableSel has invalid cell bounds"));
            }
            if !(MIN_TABLE_EDGE..=MAX_TABLE_EDGE).contains(&xa_table_left)
                || !(MIN_TABLE_EDGE..=MAX_TABLE_EDGE).contains(&xa_table_right)
                || xa_table_right < xa_table_left
            {
                return Err(corrupted("Selsf table edge is outside its twip bounds"));
            }
            if !f_block && (xa_table_left != MIN_TABLE_EDGE || xa_table_right != MAX_TABLE_EDGE) {
                return Err(corrupted(
                    "Selsf whole-row table edges must span the complete row",
                ));
            }
            SelectionGeometry::Table { first, limit }
        } else if f_block {
            let first = (raw_geometry as u16) as i16;
            let limit = ((raw_geometry >> 16) as u16) as i16;
            if limit < first {
                return Err(corrupted("Selsf BlockSel limit precedes first"));
            }
            SelectionGeometry::Block { first, limit }
        } else {
            SelectionGeometry::None
        };
        // Keep the decoded union validated even though the source retains the
        // raw four-byte representation for callers that need exact bytes.
        let _ = geometry;

        Ok(Self {
            source,
            flags,
            direction_flags,
            f_ins_end,
            cp_first,
            cp_lim,
            raw_geometry,
            cp_anchor,
            style,
            cp_anchor_shrink,
            xa_table_left,
            xa_table_right,
        })
    }

    /// Exact source bytes, including undefined fields.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.source
    }

    /// Copy the exact serialized record, including fields this API does not
    /// interpret.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.source.to_vec()
    }

    /// Whether the selection was made from physical left to right.
    #[must_use]
    pub fn is_rightward(&self) -> bool {
        self.flags & (1 << 0) != 0
    }
    /// Whether the selection is content within a table cell.
    #[must_use]
    pub fn is_within_cell(&self) -> bool {
        self.flags & (1 << 2) != 0
    }
    /// Whether the selection began with table content or cells.
    #[must_use]
    pub fn is_table_anchor(&self) -> bool {
        self.flags & (1 << 3) != 0
    }
    /// Whether the selection contains only whole cells selected by mouse.
    #[must_use]
    pub fn is_table_selection_non_shrink(&self) -> bool {
        self.flags & (1 << 4) != 0
    }
    /// Whether the selection was discontiguous; only its most recent range is stored.
    #[must_use]
    pub fn is_discontiguous(&self) -> bool {
        self.flags & (1 << 6) != 0
    }
    /// Whether the selection is a bullet or number prefix.
    #[must_use]
    pub fn is_prefix(&self) -> bool {
        self.flags & (1 << 7) != 0
    }
    /// Whether the selection is a shape or floating picture.
    #[must_use]
    pub fn is_shape(&self) -> bool {
        self.flags & (1 << 8) != 0
    }
    /// Whether the selection is a text frame.
    #[must_use]
    pub fn is_frame(&self) -> bool {
        self.flags & (1 << 9) != 0
    }
    /// Whether the selection contains one or more whole table cells.
    #[must_use]
    pub fn is_column_selection(&self) -> bool {
        self.flags & (1 << 10) != 0
    }
    /// Whether the selection is a table selection.
    #[must_use]
    pub fn is_table_selection(&self) -> bool {
        self.flags & (1 << 11) != 0
    }
    /// Whether the selection is an inline picture.
    #[must_use]
    pub fn is_graphics(&self) -> bool {
        self.flags & (1 << 12) != 0
    }
    /// Whether the selection is a rectangular block.
    #[must_use]
    pub fn is_block_selection(&self) -> bool {
        self.flags & (1 << 13) != 0
    }
    /// Whether the selection is an insertion point.
    #[must_use]
    pub fn is_insertion_point(&self) -> bool {
        self.flags & (1 << 15) != 0
    }
    /// Raw direction byte, including the ignored Word 2007 prefix bit.
    #[must_use]
    pub const fn direction_flags(&self) -> u8 {
        self.direction_flags
    }
    /// The validated logical direction bit (`fForward`).
    #[must_use]
    pub fn is_forward(&self) -> bool {
        self.direction_flags & 0x7F != 0
    }
    /// The ignored Word 2007 prefix bit (`fPrefixW2007`).
    #[must_use]
    pub fn prefix_w2007(&self) -> bool {
        self.direction_flags & 0x80 != 0
    }
    /// Whether the insertion point is at the end of its line.
    ///
    /// Shape selections leave `fInsEnd` undefined, so their raw byte is not
    /// interpreted as an insertion-end flag.
    #[must_use]
    pub const fn is_insertion_end(&self) -> bool {
        self.flags & (1 << 8) == 0 && self.f_ins_end != 0
    }
    /// Start CP of the selection.
    #[must_use]
    pub const fn cp_first(&self) -> u32 {
        self.cp_first
    }
    /// Exclusive end CP of the selection.
    #[must_use]
    pub const fn cp_lim(&self) -> u32 {
        self.cp_lim
    }
    /// Initial selection CP anchor.
    #[must_use]
    pub const fn cp_anchor(&self) -> u32 {
        self.cp_anchor
    }
    /// CP at which a block selection began; undefined for other selections.
    #[must_use]
    pub const fn cp_anchor_shrink(&self) -> i32 {
        self.cp_anchor_shrink
    }
    /// The typed `Sty` selection type.
    #[must_use]
    pub const fn style(&self) -> SelectionStyle {
        self.style
    }
    /// Raw four-byte `blktblSel` union value.
    #[must_use]
    pub const fn raw_geometry(&self) -> u32 {
        self.raw_geometry
    }
    /// Decode `blktblSel` according to the table/block flags.
    #[must_use]
    pub fn geometry(&self) -> SelectionGeometry {
        if self.is_table_selection() {
            SelectionGeometry::Table {
                first: (self.raw_geometry & 0xFFFF) as u16,
                limit: (self.raw_geometry >> 16) as u16,
            }
        } else if self.is_block_selection() {
            SelectionGeometry::Block {
                first: (self.raw_geometry as u16) as i16,
                limit: ((self.raw_geometry >> 16) as u16) as i16,
            }
        } else {
            SelectionGeometry::None
        }
    }
    /// Physical left table-cell edge in twips.
    #[must_use]
    pub const fn table_left(&self) -> i16 {
        self.xa_table_left
    }
    /// Physical right table-cell edge in twips.
    #[must_use]
    pub const fn table_right(&self) -> i16 {
        self.xa_table_right
    }
}

fn nonnegative_cp(value: i32, field: &str) -> Result<u32> {
    if !(0..=MAX_CP).contains(&value) {
        return Err(corrupted(format!("Selsf {field} is negative")));
    }
    Ok(value as u32)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(flags: u16) -> Vec<u8> {
        let mut data = vec![0; SELSF_SIZE];
        data[0..2].copy_from_slice(&flags.to_le_bytes());
        data[2] = 1;
        data[3] = 0;
        data[4..8].copy_from_slice(&4i32.to_le_bytes());
        data[8..12].copy_from_slice(&8i32.to_le_bytes());
        data[16..20].copy_from_slice(&0x0002_0001u32.to_le_bytes());
        data[20..24].copy_from_slice(&4i32.to_le_bytes());
        data[24..26].copy_from_slice(&1u16.to_le_bytes());
        data[32..34].copy_from_slice(&(-100i16).to_le_bytes());
        data[34..36].copy_from_slice(&100i16.to_le_bytes());
        data
    }

    #[test]
    fn parses_typed_text_block_and_table_flags() {
        let text = SavedSelection::parse_bytes(&record(1 << 13)).unwrap();
        assert_eq!(
            text.geometry(),
            SelectionGeometry::Block { first: 1, limit: 2 }
        );
        assert_eq!(text.cp_first(), 4);

        let mut table_data = record(1 << 11);
        table_data[16..20].copy_from_slice(&0x0040_0000u32.to_le_bytes());
        table_data[32..34].copy_from_slice(&MIN_TABLE_EDGE.to_le_bytes());
        table_data[34..36].copy_from_slice(&MAX_TABLE_EDGE.to_le_bytes());
        let table = SavedSelection::parse_bytes(&table_data).unwrap();
        assert_eq!(
            table.geometry(),
            SelectionGeometry::Table {
                first: 0,
                limit: 64
            }
        );
    }

    #[test]
    fn rejects_invalid_insertion_and_reserved_fields() {
        let mut invalid = record(1 << 15);
        invalid[8..12].copy_from_slice(&9i32.to_le_bytes());
        assert!(SavedSelection::parse_bytes(&invalid).is_err());

        let mut reserved = record(1 << 14);
        assert!(SavedSelection::parse_bytes(&reserved).is_err());
        reserved[0] = 0;
        reserved[3] = 2;
        assert!(SavedSelection::parse_bytes(&reserved).is_err());
    }

    #[test]
    fn accepts_every_normative_sty_value() {
        for raw in [
            SelectionStyle::Undefined,
            SelectionStyle::Character,
            SelectionStyle::Word,
            SelectionStyle::Sentence,
            SelectionStyle::Paragraph,
            SelectionStyle::Line,
            SelectionStyle::Column,
            SelectionStyle::Row,
            SelectionStyle::ColumnAll,
            SelectionStyle::WholeTable,
            SelectionStyle::Prefix,
        ] {
            let mut data = record(0);
            data[24..26].copy_from_slice(&(raw as u16).to_le_bytes());
            assert_eq!(SavedSelection::parse_bytes(&data).unwrap().style(), raw);
        }
        let mut invalid = record(0);
        invalid[24..26].copy_from_slice(&6u16.to_le_bytes());
        assert!(SavedSelection::parse_bytes(&invalid).is_err());
    }

    #[test]
    fn rejects_normative_cross_field_violations() {
        for flags in [
            1 << 2 | 1 << 11,
            1 << 4,
            1 << 8 | 1 << 12,
            1 << 10 | 1 << 11,
        ] {
            assert!(SavedSelection::parse_bytes(&record(flags)).is_err());
        }

        let mut row = record(1 << 11);
        row[16..20].copy_from_slice(&0x0040_0000u32.to_le_bytes());
        assert!(SavedSelection::parse_bytes(&row).is_err());
    }

    #[test]
    fn parse_uses_the_fixed_fib_range_and_rejects_bad_lengths() {
        let source = record(0);
        let offset = 8u32;
        let mut fib_data = vec![0; 154 + 31 * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x00C1u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&31u16.to_le_bytes());
        let pointer = 154 + FIB_INDEX_WSS * 8;
        fib_data[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
        fib_data[pointer + 4..pointer + 8].copy_from_slice(&(SELSF_SIZE as u32).to_le_bytes());
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        let mut table_stream = vec![0xCC; 80];
        table_stream[offset as usize..offset as usize + SELSF_SIZE].copy_from_slice(&source);
        assert_eq!(
            SavedSelection::parse(&fib, &table_stream)
                .unwrap()
                .unwrap()
                .bytes(),
            source.as_slice()
        );

        let mut absent = fib_data.clone();
        absent[pointer + 4..pointer + 8].fill(0);
        let absent_fib = FileInformationBlock::parse(&absent).unwrap();
        assert!(
            SavedSelection::parse(&absent_fib, &table_stream)
                .unwrap()
                .is_none()
        );

        let mut bad_length = fib_data.clone();
        bad_length[pointer + 4..pointer + 8].copy_from_slice(&35u32.to_le_bytes());
        let bad_fib = FileInformationBlock::parse(&bad_length).unwrap();
        assert!(SavedSelection::parse(&bad_fib, &table_stream).is_err());

        let mut bad_offset = fib_data;
        bad_offset[pointer..pointer + 4].copy_from_slice(&u32::MAX.to_le_bytes());
        let bad_fib = FileInformationBlock::parse(&bad_offset).unwrap();
        assert!(SavedSelection::parse(&bad_fib, &table_stream).is_err());
    }
}
