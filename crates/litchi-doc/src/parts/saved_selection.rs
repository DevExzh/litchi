//! The last main-document selection (`Selsf`, MS-DOC 2.9.244).
//!
//! Selection state is a passive Word UI cache. This module validates its fixed
//! fields and the `BlockSel`/`TableSel` union without applying, navigating, or
//! changing a document selection. The exact 36-byte source is retained for
//! source-bound edits. The writer applies a checked byte patch to a
//! caller-owned table stream; it does not apply the selection to a Word UI or
//! rewrite an OLE package.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use std::sync::Arc;

/// FIB `FibRgFcLcb` index for `fcWss`/`lcbWss`.
pub const FIB_INDEX_WSS: usize = 30;
/// Serialized size of one `Selsf` record.
pub const SELSF_SIZE: usize = 36;
const MAX_CP: i32 = i32::MAX;

/// Failure classes used when a main-story splice cannot preserve the passive
/// selection cache without guessing which interior CP should survive.
#[derive(Debug)]
pub(crate) enum SavedSelectionSpliceError {
    /// The selected record or its context is malformed.
    Invalid(PackageError),
    /// A CP lies strictly inside replaced text, or an insertion occurs at an
    /// existing CP, so a lossless position mapping is not provable.
    Ambiguous,
}
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
#[derive(Debug, Clone)]
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
    /// Optional complete DOC source used by the high-level edit facade.
    /// Detached component readers intentionally leave this unset.
    owner: Option<Arc<[u8]>>,
}

impl PartialEq for SavedSelection {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for SavedSelection {}

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
            owner: None,
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

    /// Attach the immutable owning DOC source used by the high-level facade.
    ///
    /// The fixed Selsf bytes remain independently retained; cloning this owner
    /// handle does not copy the complete document.
    pub(crate) fn with_owner(mut self, owner: Arc<[u8]>) -> Self {
        self.owner = Some(owner);
        self
    }

    fn with_optional_owner(mut self, owner: Option<Arc<[u8]>>) -> Self {
        self.owner = owner;
        self
    }

    /// Start a source-bound edit of this record.
    #[must_use]
    pub fn transaction(&self) -> SavedSelectionTransaction {
        SavedSelectionTransaction {
            source: self.clone(),
            draft: self.source,
            owner: self.owner.clone(),
        }
    }

    /// Validate the CP fields whose coordinates are relative to the main DOC
    /// story. The raw parser intentionally remains context-free; the DOC
    /// owner supplies `ccpText` at this boundary.
    pub(crate) fn validate_main_story_bound(&self, ccp_text: u32) -> Result<()> {
        if self.cp_first > ccp_text {
            return Err(corrupted(format!(
                "Selsf cpFirst {} exceeds ccpText {ccp_text}",
                self.cp_first
            )));
        }
        if self.cp_lim > ccp_text {
            return Err(corrupted(format!(
                "Selsf cpLim {} exceeds ccpText {ccp_text}",
                self.cp_lim
            )));
        }
        if self.cp_anchor > ccp_text {
            return Err(corrupted(format!(
                "Selsf cpAnchor {} exceeds ccpText {ccp_text}",
                self.cp_anchor
            )));
        }
        if self.is_block_selection() && !self.is_table_selection() {
            let cp_anchor_shrink = u32::try_from(self.cp_anchor_shrink)
                .map_err(|_| corrupted("Selsf cpAnchorShrink is negative"))?;
            if cp_anchor_shrink > ccp_text {
                return Err(corrupted(format!(
                    "Selsf cpAnchorShrink {cp_anchor_shrink} exceeds ccpText {ccp_text}"
                )));
            }
        }
        Ok(())
    }

    /// Remap the three absolute main-story CPs across one text splice.
    ///
    /// CPs before the replaced range remain fixed, CPs after it move by the
    /// signed size delta, and the two exact range boundaries map to the
    /// corresponding new boundaries. An interior CP cannot be mapped without
    /// selecting a semantic point inside replaced text, so the operation is
    /// refused instead of retaining a stale coordinate.
    pub(crate) fn remap_for_splice(
        &self,
        start: u32,
        end: u32,
        added: u32,
        new_ccp: u32,
    ) -> std::result::Result<[u8; SELSF_SIZE], SavedSelectionSpliceError> {
        let cp_first = remap_splice_cp(self.cp_first, start, end, added)?;
        let cp_lim = remap_splice_cp(self.cp_lim, start, end, added)?;
        let cp_anchor = remap_splice_cp(self.cp_anchor, start, end, added)?;
        let cp_anchor_shrink = if self.is_block_selection() && !self.is_table_selection() {
            let cp_anchor_shrink = u32::try_from(self.cp_anchor_shrink).map_err(|_| {
                SavedSelectionSpliceError::Invalid(corrupted("Selsf cpAnchorShrink is negative"))
            })?;
            Some(remap_splice_cp(cp_anchor_shrink, start, end, added)?)
        } else {
            None
        };
        if cp_first > new_ccp || cp_lim > new_ccp || cp_anchor > new_ccp {
            return Err(SavedSelectionSpliceError::Invalid(corrupted(
                "remapped Selsf CP exceeds the new main story",
            )));
        }
        if cp_anchor_shrink.is_some_and(|value| value > new_ccp) {
            return Err(SavedSelectionSpliceError::Invalid(corrupted(
                "remapped Selsf cpAnchorShrink exceeds the new main story",
            )));
        }
        let cp_first = encode_cp_for_splice(cp_first)?;
        let cp_lim = encode_cp_for_splice(cp_lim)?;
        let cp_anchor = encode_cp_for_splice(cp_anchor)?;
        let cp_anchor_shrink = cp_anchor_shrink.map(encode_cp_for_splice).transpose()?;
        let mut replacement = self.source;
        replacement[4..8].copy_from_slice(&cp_first.to_le_bytes());
        replacement[8..12].copy_from_slice(&cp_lim.to_le_bytes());
        replacement[20..24].copy_from_slice(&cp_anchor.to_le_bytes());
        if let Some(cp_anchor_shrink) = cp_anchor_shrink {
            replacement[28..32].copy_from_slice(&cp_anchor_shrink.to_le_bytes());
        }
        SavedSelection::parse_bytes(&replacement)
            .map(|_| replacement)
            .map_err(SavedSelectionSpliceError::Invalid)
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

/// A checked source-bound edit of one [`SavedSelection`] record.
///
/// Each setter stages its change in a private byte copy and reparses the full
/// candidate. This keeps CP, insertion, table, and block invariants together,
/// while bytes outside the field being edited remain untouched.
#[derive(Debug, Clone)]
pub struct SavedSelectionTransaction {
    source: SavedSelection,
    draft: [u8; SELSF_SIZE],
    owner: Option<Arc<[u8]>>,
}

impl SavedSelectionTransaction {
    /// The source snapshot against which this transaction was opened.
    #[must_use]
    pub fn source(&self) -> &SavedSelection {
        &self.source
    }

    /// Return the currently staged snapshot after validating all fields.
    pub fn snapshot(&self) -> Result<SavedSelection> {
        SavedSelection::parse_bytes(&self.draft)
            .map(|snapshot| snapshot.with_optional_owner(self.owner.clone()))
    }

    /// Replace the selected CP range.
    pub fn set_range(&mut self, cp_first: u32, cp_lim: u32) -> Result<&mut Self> {
        let cp_first = encode_cp(cp_first, "cpFirst")?;
        let cp_lim = encode_cp(cp_lim, "cpLim")?;
        if cp_lim < cp_first {
            return Err(corrupted("Selsf cpLim precedes cpFirst"));
        }
        self.stage(|draft| {
            draft[4..8].copy_from_slice(&cp_first.to_le_bytes());
            draft[8..12].copy_from_slice(&cp_lim.to_le_bytes());
        })
    }

    /// Replace the initial selection anchor CP.
    pub fn set_cp_anchor(&mut self, cp_anchor: u32) -> Result<&mut Self> {
        let cp_anchor = encode_cp(cp_anchor, "cpAnchor")?;
        self.stage(|draft| draft[20..24].copy_from_slice(&cp_anchor.to_le_bytes()))
    }

    /// Replace the CP at which a block selection began.
    pub fn set_cp_anchor_shrink(&mut self, cp_anchor_shrink: i32) -> Result<&mut Self> {
        let flags = u16::from_le_bytes([self.draft[0], self.draft[1]]);
        if flags & (1 << 13) == 0 || flags & (1 << 11) != 0 {
            return Err(corrupted(
                "Selsf cpAnchorShrink is only applicable to a text block selection",
            ));
        }
        self.stage(|draft| {
            draft[28..32].copy_from_slice(&cp_anchor_shrink.to_le_bytes());
        })
    }

    /// Set the logical forward direction while retaining the Word 2007 prefix bit.
    pub fn set_forward(&mut self, forward: bool) -> Result<&mut Self> {
        self.stage(|draft| {
            let mut value = draft[2] & 0x80;
            if forward {
                value |= 1;
            }
            draft[2] = value;
        })
    }

    /// Set the ignored Word 2007 prefix bit in the direction byte.
    pub fn set_prefix_w2007(&mut self, enabled: bool) -> Result<&mut Self> {
        self.stage(|draft| {
            if enabled {
                draft[2] |= 0x80;
            } else {
                draft[2] &= 0x7F;
            }
        })
    }

    /// Set or clear the insertion-point flag.
    ///
    /// Enabling insertion-point mode collapses `cpLim` to `cpFirst`; disabling
    /// it clears `fInsEnd` only when that byte is applicable. Shape selections
    /// leave `fInsEnd` undefined, so its source byte is preserved.
    pub fn set_insertion_point(&mut self, enabled: bool) -> Result<&mut Self> {
        self.stage(|draft| {
            let mut flags = u16::from_le_bytes([draft[0], draft[1]]);
            if enabled {
                flags |= 1 << 15;
                let cp_first = [draft[4], draft[5], draft[6], draft[7]];
                draft[8..12].copy_from_slice(&cp_first);
            } else {
                flags &= !(1 << 15);
                if flags & (1 << 8) == 0 {
                    draft[3] = 0;
                }
            }
            draft[0..2].copy_from_slice(&flags.to_le_bytes());
        })
    }

    /// Set the insertion-point-at-line-end flag.
    pub fn set_insertion_end(&mut self, enabled: bool) -> Result<&mut Self> {
        let flags = u16::from_le_bytes([self.draft[0], self.draft[1]]);
        if flags & (1 << 8) != 0 || flags & (1 << 15) == 0 {
            return Err(corrupted(
                "Selsf fInsEnd is only applicable to an ordinary insertion point",
            ));
        }
        self.stage(|draft| draft[3] = u8::from(enabled))
    }

    /// Set the typed selection style.
    pub fn set_style(&mut self, style: SelectionStyle) -> Result<&mut Self> {
        self.stage(|draft| draft[24..26].copy_from_slice(&(style as u16).to_le_bytes()))
    }

    /// Set the block/table union and its corresponding flags.
    pub fn set_geometry(&mut self, geometry: SelectionGeometry) -> Result<&mut Self> {
        if let SelectionGeometry::Table { first, limit } = geometry {
            if first > 63 || limit > 64 || limit < first {
                return Err(corrupted("Selsf TableSel has invalid cell bounds"));
            }
        }
        self.stage(|draft| {
            let mut flags = u16::from_le_bytes([draft[0], draft[1]]);
            let existing_left = i16::from_le_bytes([draft[32], draft[33]]);
            let existing_right = i16::from_le_bytes([draft[34], draft[35]]);
            let whole_row = matches!(
                geometry,
                SelectionGeometry::Table {
                    first: 0,
                    limit: 64
                }
            );
            let raw = match geometry {
                SelectionGeometry::None => {
                    // These flags describe the TableSel/BlockSel union and
                    // must not survive a transition to ordinary text. The
                    // union bytes become ignored here and are preserved.
                    flags &= !((1 << 4) | (1 << 10) | (1 << 11) | (1 << 13));
                    u32::from_le_bytes([draft[16], draft[17], draft[18], draft[19]])
                },
                SelectionGeometry::Block { first, limit } => {
                    flags = (flags | (1 << 13)) & !((1 << 4) | (1 << 10) | (1 << 11));
                    u32::from(u16::from_le_bytes(first.to_le_bytes()))
                        | (u32::from(u16::from_le_bytes(limit.to_le_bytes())) << 16)
                },
                SelectionGeometry::Table { first, limit } => {
                    // fWithinCell is meaningful for text in a cell, while a
                    // TableSel always denotes whole cells/rows.
                    flags = (flags | (1 << 11)) & !(1 << 2);
                    if first == 0 && limit == 64 {
                        // Whole rows/whole table cannot be a column selection
                        // and are represented without the block bit.
                        flags &= !((1 << 10) | (1 << 13));
                    } else {
                        flags |= 1 << 13;
                        // A prior non-table record may contain arbitrary
                        // ignored edge bytes. Give a partial TableSel a valid
                        // starting pair so the result does not depend on an
                        // inapplicable edge setter being called first.
                        if !table_edges_are_valid(existing_left, existing_right) {
                            draft[32..34].copy_from_slice(&MIN_TABLE_EDGE.to_le_bytes());
                            draft[34..36].copy_from_slice(&MAX_TABLE_EDGE.to_le_bytes());
                        }
                    }
                    u32::from(first) | (u32::from(limit) << 16)
                },
            };
            draft[0..2].copy_from_slice(&flags.to_le_bytes());
            draft[16..20].copy_from_slice(&raw.to_le_bytes());
            if whole_row {
                draft[32..34].copy_from_slice(&MIN_TABLE_EDGE.to_le_bytes());
                draft[34..36].copy_from_slice(&MAX_TABLE_EDGE.to_le_bytes());
            }
        })
    }

    /// Set the physical table-cell edges in twips.
    pub fn set_table_edges(&mut self, left: i16, right: i16) -> Result<&mut Self> {
        let flags = u16::from_le_bytes([self.draft[0], self.draft[1]]);
        if flags & (1 << 11) == 0 {
            return Err(corrupted(
                "Selsf table edges are only applicable to a table selection",
            ));
        }
        validate_table_edges(left, right)?;
        if flags & (1 << 13) == 0 && (left != MIN_TABLE_EDGE || right != MAX_TABLE_EDGE) {
            return Err(corrupted(
                "Selsf whole-row table edges must span the complete row",
            ));
        }
        self.stage(|draft| {
            draft[32..34].copy_from_slice(&left.to_le_bytes());
            draft[34..36].copy_from_slice(&right.to_le_bytes());
        })
    }

    /// Commit the staged record and return its reversible source patch.
    pub fn commit(self) -> Result<SavedSelectionCommit> {
        let owner = self.owner;
        let snapshot = SavedSelection::parse_bytes(&self.draft)?.with_optional_owner(owner.clone());
        let patch = SavedSelectionPatch {
            source: self.source.source,
            replacement: snapshot.source,
            owner,
        };
        Ok(SavedSelectionCommit { snapshot, patch })
    }

    fn stage(&mut self, edit: impl FnOnce(&mut [u8; SELSF_SIZE])) -> Result<&mut Self> {
        let mut candidate = self.draft;
        edit(&mut candidate);
        SavedSelection::parse_bytes(&candidate)?;
        self.draft = candidate;
        Ok(self)
    }
}

/// The result of committing a [`SavedSelectionTransaction`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SavedSelectionCommit {
    snapshot: SavedSelection,
    patch: SavedSelectionPatch,
}

impl SavedSelectionCommit {
    /// The validated post-edit record.
    #[must_use]
    pub fn snapshot(&self) -> &SavedSelection {
        &self.snapshot
    }

    /// The source-checked byte patch from the original record to `snapshot`.
    #[must_use]
    pub fn patch(&self) -> &SavedSelectionPatch {
        &self.patch
    }

    /// Whether any serialized byte changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.patch.source != self.patch.replacement
    }

    /// Consume the commit and return the post-edit record.
    #[must_use]
    pub fn into_snapshot(self) -> SavedSelection {
        self.snapshot
    }
}

/// A reversible, source-checked byte patch for one `Selsf` record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SavedSelectionPatch {
    source: [u8; SELSF_SIZE],
    replacement: [u8; SELSF_SIZE],
    owner: Option<Arc<[u8]>>,
}

impl SavedSelectionPatch {
    /// Exact source bytes required before applying this patch.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Exact replacement bytes written by this patch.
    #[must_use]
    pub fn replacement_bytes(&self) -> &[u8] {
        &self.replacement
    }

    /// Whether applying this patch changes any serialized Selsf byte.
    #[must_use]
    pub fn changed(&self) -> bool {
        self.source != self.replacement
    }

    /// Whether this changed patch was produced from the supplied immutable
    /// DOC source. Pointer identity is a fast path; exact bytes also authorize
    /// an independently reopened copy of the same source.
    pub(crate) fn owner_matches(&self, owner: &Arc<[u8]>) -> bool {
        self.owner
            .as_ref()
            .is_some_and(|bound| Arc::ptr_eq(bound, owner) || bound.as_ref() == owner.as_ref())
    }

    /// Apply this patch to a matching parsed record.
    pub fn apply(&self, source: &SavedSelection) -> Result<SavedSelection> {
        if source.source != self.source {
            return Err(corrupted("Selsf patch source does not match the record"));
        }
        SavedSelection::parse_bytes(&self.replacement)
            .map(|selection| selection.with_optional_owner(source.owner.clone()))
    }

    /// Revert this patch from a matching post-edit record.
    pub fn revert(&self, replacement: &SavedSelection) -> Result<SavedSelection> {
        if replacement.source != self.replacement {
            return Err(corrupted(
                "Selsf patch replacement does not match the record",
            ));
        }
        SavedSelection::parse_bytes(&self.source)
            .map(|selection| selection.with_optional_owner(replacement.owner.clone()))
    }

    /// Validate this patch against the FIB-selected range and exact source
    /// bytes without taking a mutable table-stream copy.
    pub fn preflight_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<(usize, usize)> {
        let (start, end) = table_range(fib, table_stream.len())?;
        if table_stream[start..end] != self.source {
            return Err(corrupted(
                "Selsf patch source does not match the table stream",
            ));
        }
        Ok((start, end))
    }

    /// Validate both sides of this patch against the owning DOC main-story
    /// length. Detached record patches intentionally have no such context.
    pub(crate) fn validate_main_story_bound(&self, ccp_text: u32) -> Result<()> {
        SavedSelection::parse_bytes(&self.source)?.validate_main_story_bound(ccp_text)?;
        SavedSelection::parse_bytes(&self.replacement)?.validate_main_story_bound(ccp_text)
    }

    /// Apply this patch in place to the `Selsf` range of a caller-owned table stream.
    pub fn apply_to_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &mut [u8],
    ) -> Result<()> {
        let (start, end) = self.preflight_table_stream(fib, table_stream)?;
        table_stream[start..end].copy_from_slice(&self.replacement);
        Ok(())
    }

    /// Revert this patch in place from a matching caller-owned table stream.
    pub fn revert_to_table_stream(
        &self,
        fib: &FileInformationBlock,
        table_stream: &mut [u8],
    ) -> Result<()> {
        let (start, end) = table_range(fib, table_stream.len())?;
        if table_stream[start..end] != self.replacement {
            return Err(corrupted(
                "Selsf patch replacement does not match the table stream",
            ));
        }
        table_stream[start..end].copy_from_slice(&self.source);
        Ok(())
    }
}

fn encode_cp(value: u32, field: &str) -> Result<i32> {
    i32::try_from(value).map_err(|_| corrupted(format!("Selsf {field} exceeds signed CP range")))
}

fn validate_table_edges(left: i16, right: i16) -> Result<()> {
    if !table_edges_are_valid(left, right) {
        return Err(corrupted("Selsf table edge is outside its twip bounds"));
    }
    Ok(())
}

fn table_edges_are_valid(left: i16, right: i16) -> bool {
    (MIN_TABLE_EDGE..=MAX_TABLE_EDGE).contains(&left)
        && (MIN_TABLE_EDGE..=MAX_TABLE_EDGE).contains(&right)
        && right >= left
}

pub(crate) fn table_range(
    fib: &FileInformationBlock,
    table_stream_len: usize,
) -> Result<(usize, usize)> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_WSS) else {
        return Err(corrupted("Selsf table pointer is unavailable"));
    };
    let length =
        usize::try_from(length).map_err(|_| corrupted("Selsf length does not fit in memory"))?;
    if length != SELSF_SIZE {
        return Err(corrupted(format!(
            "Selsf must be {SELSF_SIZE} bytes, got {length}"
        )));
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted("Selsf offset does not fit in memory"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("Selsf range overflows"))?;
    if end > table_stream_len {
        return Err(corrupted("Selsf extends beyond the table stream"));
    }
    Ok((start, end))
}

fn remap_splice_cp(
    cp: u32,
    start: u32,
    end: u32,
    added: u32,
) -> std::result::Result<u32, SavedSelectionSpliceError> {
    if start == end {
        if cp == start {
            return Err(SavedSelectionSpliceError::Ambiguous);
        }
        return if cp < start {
            Ok(cp)
        } else {
            cp.checked_add(added).ok_or_else(|| {
                SavedSelectionSpliceError::Invalid(corrupted("remapped Selsf CP overflows"))
            })
        };
    }
    if cp < start {
        return Ok(cp);
    }
    if cp == start {
        return Ok(start);
    }
    if cp == end {
        return start.checked_add(added).ok_or_else(|| {
            SavedSelectionSpliceError::Invalid(corrupted("remapped Selsf CP overflows"))
        });
    }
    if cp < end {
        return Err(SavedSelectionSpliceError::Ambiguous);
    }
    let removed = end - start;
    if added >= removed {
        cp.checked_add(added - removed).ok_or_else(|| {
            SavedSelectionSpliceError::Invalid(corrupted("remapped Selsf CP overflows"))
        })
    } else {
        cp.checked_sub(removed - added).ok_or_else(|| {
            SavedSelectionSpliceError::Invalid(corrupted("remapped Selsf CP underflows"))
        })
    }
}

fn encode_cp_for_splice(value: u32) -> std::result::Result<i32, SavedSelectionSpliceError> {
    i32::try_from(value).map_err(|_| {
        SavedSelectionSpliceError::Invalid(corrupted("remapped Selsf CP exceeds signed range"))
    })
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

        let mut shape = record(1 << 8);
        shape[3] = 2;
        let shape_selection = SavedSelection::parse_bytes(&shape).unwrap();
        assert!(!shape_selection.is_insertion_end());
        shape[3] = 1;
        assert!(SavedSelection::parse_bytes(&shape).is_ok());
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
            1 << 10,
            1 << 10 | 1 << 11,
        ] {
            assert!(SavedSelection::parse_bytes(&record(flags)).is_err());
        }

        let mut row = record(1 << 11);
        row[16..20].copy_from_slice(&0x0040_0000u32.to_le_bytes());
        assert!(SavedSelection::parse_bytes(&row).is_err());
    }

    #[test]
    fn transaction_preserves_unknown_bytes_and_checks_ranges() {
        let mut data = record(0);
        data[12..16].copy_from_slice(&[0xA5, 0x5A, 0xC3, 0x3C]);
        data[26..28].copy_from_slice(&[0xD7, 0x7D]);
        let source = SavedSelection::parse_bytes(&data).unwrap();

        let mut transaction = source.transaction();
        assert!(transaction.set_range(10, 9).is_err());
        assert!(
            transaction
                .set_geometry(SelectionGeometry::Table { first: 4, limit: 2 })
                .is_err()
        );
        let mut table_source = record((1 << 10) | (1 << 11) | (1 << 13));
        table_source[16..20].copy_from_slice(&0x0002_0001u32.to_le_bytes());
        let table_source = SavedSelection::parse_bytes(&table_source).unwrap();
        let mut ordinary = table_source.transaction();
        ordinary.set_geometry(SelectionGeometry::None).unwrap();
        assert_eq!(
            ordinary.snapshot().unwrap().geometry(),
            SelectionGeometry::None
        );
        let mut within_cell = record(1 << 2);
        let mut table_transaction = SavedSelection::parse_bytes(&within_cell)
            .unwrap()
            .transaction();
        table_transaction
            .set_geometry(SelectionGeometry::Table {
                first: 0,
                limit: 64,
            })
            .unwrap();
        assert_eq!(
            table_transaction.snapshot().unwrap().geometry(),
            SelectionGeometry::Table {
                first: 0,
                limit: 64
            }
        );
        within_cell.fill(0);
        assert_ne!(within_cell, table_transaction.snapshot().unwrap().bytes());
        assert_eq!(transaction.snapshot().unwrap(), source);
        transaction
            .set_range(2, 6)
            .unwrap()
            .set_cp_anchor(4)
            .unwrap()
            .set_prefix_w2007(true)
            .unwrap();
        let commit = transaction.commit().unwrap();
        assert!(commit.changed());
        assert_eq!(
            &commit.snapshot().bytes()[12..16],
            &[0xA5, 0x5A, 0xC3, 0x3C]
        );
        assert_eq!(&commit.snapshot().bytes()[26..28], &[0xD7, 0x7D]);
        assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
        assert_eq!(commit.patch().revert(commit.snapshot()).unwrap(), source);
    }

    #[test]
    fn semantic_noops_preserve_ignored_union_and_shape_insertion_bytes() {
        let mut ordinary = record(0);
        ordinary[16..20].copy_from_slice(&0xA1B2_C3D4u32.to_le_bytes());
        ordinary[32..34].copy_from_slice(&i16::MIN.to_le_bytes());
        ordinary[34..36].copy_from_slice(&i16::MAX.to_le_bytes());
        let ordinary = SavedSelection::parse_bytes(&ordinary).unwrap();
        let mut ordinary_transaction = ordinary.transaction();
        ordinary_transaction
            .set_geometry(SelectionGeometry::None)
            .unwrap();
        assert_eq!(
            ordinary_transaction.snapshot().unwrap().bytes(),
            ordinary.bytes()
        );

        let mut shape = record(1 << 8);
        shape[3] = 0xA5;
        let shape = SavedSelection::parse_bytes(&shape).unwrap();
        let mut shape_transaction = shape.transaction();
        shape_transaction.set_insertion_point(false).unwrap();
        assert_eq!(shape_transaction.snapshot().unwrap().bytes(), shape.bytes());
    }

    #[test]
    fn table_transition_normalizes_ignored_edges_and_setters_check_applicability() {
        let mut ordinary = record(0);
        ordinary[32..34].copy_from_slice(&i16::MIN.to_le_bytes());
        ordinary[34..36].copy_from_slice(&i16::MAX.to_le_bytes());
        let ordinary = SavedSelection::parse_bytes(&ordinary).unwrap();
        let mut transaction = ordinary.transaction();
        transaction
            .set_geometry(SelectionGeometry::Table { first: 2, limit: 4 })
            .unwrap();
        assert_eq!(
            transaction.snapshot().unwrap().geometry(),
            SelectionGeometry::Table { first: 2, limit: 4 }
        );
        assert_eq!(transaction.snapshot().unwrap().table_left(), MIN_TABLE_EDGE);
        assert_eq!(
            transaction.snapshot().unwrap().table_right(),
            MAX_TABLE_EDGE
        );
        transaction.set_table_edges(-100, 100).unwrap();
        assert_eq!(transaction.snapshot().unwrap().table_left(), -100);
        assert_eq!(transaction.snapshot().unwrap().table_right(), 100);

        let mut ordinary_transaction = ordinary.transaction();
        assert!(ordinary_transaction.set_table_edges(-100, 100).is_err());
        assert!(ordinary_transaction.set_cp_anchor_shrink(4).is_err());

        let block = SavedSelection::parse_bytes(&record(1 << 13)).unwrap();
        let mut block_transaction = block.transaction();
        block_transaction.set_cp_anchor_shrink(3).unwrap();
        assert_eq!(block_transaction.snapshot().unwrap().cp_anchor_shrink(), 3);
    }

    #[test]
    fn block_anchor_shrink_remaps_and_context_bounds_cover_all_active_cps() {
        let mut data = record(1 << 13);
        data[4..8].copy_from_slice(&4i32.to_le_bytes());
        data[8..12].copy_from_slice(&8i32.to_le_bytes());
        data[20..24].copy_from_slice(&4i32.to_le_bytes());
        data[28..32].copy_from_slice(&4i32.to_le_bytes());
        let selection = SavedSelection::parse_bytes(&data).unwrap();
        assert!(selection.validate_main_story_bound(3).is_err());
        assert!(selection.validate_main_story_bound(8).is_ok());
        let remapped = selection.remap_for_splice(4, 8, 2, 10).unwrap();
        assert_eq!(i32::from_le_bytes(remapped[8..12].try_into().unwrap()), 6);
        assert_eq!(i32::from_le_bytes(remapped[28..32].try_into().unwrap()), 4);
    }

    #[test]
    fn patch_updates_only_the_declared_table_range() {
        let source = SavedSelection::parse_bytes(&record(0)).unwrap();
        let mut transaction = source.transaction();
        transaction.set_range(2, 9).unwrap();
        let commit = transaction.commit().unwrap();

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
        table_stream[offset as usize..offset as usize + SELSF_SIZE].copy_from_slice(source.bytes());
        let outside_before = table_stream[..offset as usize].to_vec();
        commit
            .patch()
            .apply_to_table_stream(&fib, &mut table_stream)
            .unwrap();
        assert_eq!(
            &table_stream[offset as usize..offset as usize + SELSF_SIZE],
            commit.snapshot().bytes()
        );
        assert_eq!(&table_stream[..offset as usize], outside_before.as_slice());
        commit
            .patch()
            .revert_to_table_stream(&fib, &mut table_stream)
            .unwrap();
        assert_eq!(
            &table_stream[offset as usize..offset as usize + SELSF_SIZE],
            source.bytes()
        );
    }
}
