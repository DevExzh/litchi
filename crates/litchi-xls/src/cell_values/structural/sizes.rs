//! Row heights and column widths around a row or column shift, for the
//! drawing anchors that keep their place (record 0316).
//!
//! A position-locked anchor (`AnchorBehavior::Absolute`) and the far end of
//! a size-locked one (`MoveWithoutResize`) keep their absolute position
//! while cells move. Their cell coordinates are unchanged exactly when every
//! row (or column) from the shift point to the anchor keeps its size: then
//! the same cell index and the same 1/256 (1/1024) fraction name the same
//! place. That is the common case of a sheet whose rows all have the
//! default height; any other layout is refused, since converting the
//! position into new cells would round.

use std::collections::BTreeMap;

use super::{AxisShift, COL_INFO, ROW, RawRecord, Result, binary, patch_col_info, top_level};

const DEFAULT_ROW_HEIGHT: u16 = 0x0225;

/// One row's height in twips or one column's `COLINFO` width, hidden as
/// zero; `None` is the default size (no `ROW` or `COLINFO` record).
type Size = Option<u16>;

/// The sizes of one axis before a shift and after it.
pub(super) struct SizeCheck {
    rows: bool,
    before: BTreeMap<u16, u16>,
    after: BTreeMap<u16, u16>,
    default: Size,
}

impl SizeCheck {
    /// The sizes of the shifted axis of `worksheet` and the sizes the shift
    /// gives it: rows move with their `ROW` records and inserted rows have
    /// none; `COLINFO` spans shift as `patch_col_info` shifts them.
    pub(super) fn predict(
        worksheet: &[u8],
        records: &[RawRecord],
        shift: AxisShift,
    ) -> Result<Self> {
        let mut before = BTreeMap::new();
        let mut after = BTreeMap::new();
        let mut default = None;
        let rows = matches!(shift, AxisShift::Rows { .. });
        for (_, record) in top_level(records) {
            let Some(payload) = worksheet.get(record.start + 4..record.end) else {
                continue;
            };
            match record.kind {
                ROW if rows && payload.len() >= 16 => {
                    let row = binary::read_u16_le_at(payload, 0)?;
                    let height = binary::read_u16_le_at(payload, 6)? & 0x7fff;
                    let flags = binary::read_u16_le_at(payload, 12)?;
                    let size = if flags & 0x0020 != 0 { 0 } else { height };
                    before.insert(row, size);
                    if let Some(moved) = moved_point(row, shift) {
                        after.insert(moved, size);
                    }
                },
                DEFAULT_ROW_HEIGHT if rows && payload.len() >= 4 => {
                    let flags = binary::read_u16_le_at(payload, 0)?;
                    let height = binary::read_u16_le_at(payload, 2)?;
                    default = Some(if flags & 0x0002 != 0 { 0 } else { height });
                },
                COL_INFO if !rows && payload.len() >= 10 => {
                    let mut shifted = payload.to_vec();
                    patch_col_info(&mut shifted, shift)?;
                    for (map, span) in [(&mut before, payload), (&mut after, &shifted[..])] {
                        let first = binary::read_u16_le_at(span, 0)?;
                        let last = binary::read_u16_le_at(span, 2)?.min(255);
                        let width = binary::read_u16_le_at(span, 4)?;
                        let hidden = binary::read_u16_le_at(span, 8)? & 0x0001 != 0;
                        for column in first..=last {
                            map.insert(column, if hidden { 0 } else { width });
                        }
                    }
                },
                _ => {},
            }
        }
        Ok(Self {
            rows,
            before,
            after,
            default,
        })
    }

    fn size(&self, map: &BTreeMap<u16, u16>, index: u16) -> Size {
        map.get(&index)
            .copied()
            .or(if self.rows { self.default } else { None })
    }

    /// Whether every row or column of `first..=last` (`first` the shift
    /// point) has the same size after the shift as before it.
    pub(super) fn unchanged(&self, first: u16, last: u16) -> bool {
        (first..=last).all(|index| self.size(&self.before, index) == self.size(&self.after, index))
    }
}

/// Where a row with a `ROW` record lands, or `None` when it is deleted.
fn moved_point(value: u16, shift: AxisShift) -> Option<u16> {
    let AxisShift::Rows {
        start,
        count,
        insert,
    } = shift
    else {
        return Some(value);
    };
    if value < start {
        Some(value)
    } else if insert {
        value.checked_add(count)
    } else if value < start.saturating_add(count) {
        None
    } else {
        Some(value - count)
    }
}

#[cfg(test)]
mod tests {
    use super::super::{patch_client_anchor, raw_records};
    use super::*;

    fn record(kind: u16, payload: &[u8]) -> Vec<u8> {
        let mut bytes = kind.to_le_bytes().to_vec();
        bytes.extend_from_slice(&u16::try_from(payload.len()).unwrap().to_le_bytes());
        bytes.extend_from_slice(payload);
        bytes
    }

    fn row(index: u16, height: u16) -> Vec<u8> {
        let mut payload = index.to_le_bytes().to_vec();
        payload.extend_from_slice(&[0, 0, 1, 0]);
        payload.extend_from_slice(&height.to_le_bytes());
        payload.extend_from_slice(&[0; 8]);
        record(ROW, &payload)
    }

    fn worksheet(heights: &[u16]) -> Vec<u8> {
        let mut bytes = record(DEFAULT_ROW_HEIGHT, &[0, 0, 255, 0]);
        for (index, height) in heights.iter().enumerate() {
            bytes.extend(row(u16::try_from(index).unwrap(), *height));
        }
        bytes
    }

    const INSERT: AxisShift = AxisShift::Rows {
        start: 0,
        count: 1,
        insert: true,
    };

    /// A ClientAnchor from row 1 to row 3 with `flags`.
    fn anchor(flags: u16) -> [u8; 18] {
        let mut anchor = [0; 18];
        anchor[0..2].copy_from_slice(&flags.to_le_bytes());
        anchor[6..8].copy_from_slice(&1u16.to_le_bytes());
        anchor[14..16].copy_from_slice(&3u16.to_le_bytes());
        anchor
    }

    fn check(heights: &[u16]) -> SizeCheck {
        let bytes = worksheet(heights);
        SizeCheck::predict(&bytes, &raw_records(&bytes).unwrap(), INSERT).unwrap()
    }

    #[test]
    fn anchors_that_keep_their_place_need_rows_of_one_height() {
        // Rows of the default height: a position-locked anchor keeps its
        // coordinates, and so does the far end of a size-locked one.
        let even = check(&[255, 255, 255, 255]);
        assert!(even.unchanged(0, 3));
        let mut locked = anchor(3);
        patch_client_anchor(&mut locked, INSERT, Some(&even)).unwrap();
        assert_eq!(locked, anchor(3));
        // One taller row moves under the anchor: refused.
        let uneven = check(&[255, 400, 255, 255]);
        assert!(!uneven.unchanged(0, 3));
        assert!(patch_client_anchor(&mut anchor(3), INSERT, Some(&uneven)).is_err());
        assert!(patch_client_anchor(&mut anchor(3), INSERT, None).is_err());
        // `fMove` alone: the top-left keeps its place, the bottom-right
        // moves with its cell.
        let mut resizing = anchor(1);
        patch_client_anchor(&mut resizing, INSERT, Some(&even)).unwrap();
        assert_eq!(&resizing[6..8], &[1, 0]);
        assert_eq!(&resizing[14..16], &[4, 0]);
        let delete = AxisShift::Rows {
            start: 0,
            count: 1,
            insert: false,
        };
        let bytes = worksheet(&[255, 255, 255, 255, 255]);
        let sizes = SizeCheck::predict(&bytes, &raw_records(&bytes).unwrap(), delete).unwrap();
        patch_client_anchor(&mut resizing, delete, Some(&sizes)).unwrap();
        assert_eq!(resizing, anchor(1));
    }
}
