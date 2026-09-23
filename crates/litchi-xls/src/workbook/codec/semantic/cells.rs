//! Where one worksheet parse keeps the cell records it has validated.
//!
//! The worksheet parser frames, decodes and checks every record the same way
//! whichever store it is given: the record grammar, the XF index, the
//! duplicate-cell limit, `Formula` companions, `PtgExp` ownership and `Array`
//! ranges. Those checks ask exactly two questions of earlier cells — is this
//! position already occupied, and is the latest record there a `Formula` —
//! and a store answers them. The stores differ only in whether a validated
//! record is also decoded into a [`Cell`].

use crate::cell::Cell;
use crate::error::{Error, Result};
use crate::worksheet::Worksheet;
use litchi_core::sheet::Cell as _;
use std::collections::HashMap;

/// A cell record's `(row, column)` exactly as the record stores it.
pub(crate) type Position = (u16, u16);

/// What a store knows about one position.
pub(crate) enum CellSlot<'w> {
    /// No cell record occupies the position.
    Vacant,
    /// A cell record occupies the position.
    Occupied {
        /// Whether the latest record at the position is a `Formula`.
        formula: bool,
        /// The decoded cell, when this store keeps the position.
        cell: Option<&'w mut Cell>,
    },
}

/// Keeps the cell records one worksheet parse has validated.
pub(crate) trait CellStore {
    /// Whether an earlier cell record occupies `position`.
    fn is_occupied(&self, worksheet: &Worksheet, position: Position) -> bool;

    /// Keeps one validated cell record. `formula` says whether it is a
    /// `Formula` record; `decode` builds exactly the cell the public reader
    /// keeps for it, and a store that does not keep `position` never calls
    /// it.
    ///
    /// # Errors
    ///
    /// Returns an allocation error when the store cannot record the position.
    fn store(
        &mut self,
        worksheet: &mut Worksheet,
        position: Position,
        formula: bool,
        decode: impl FnOnce() -> Cell,
    ) -> Result<()>;

    /// What this store knows about `position`.
    fn slot<'w>(&self, worksheet: &'w mut Worksheet, position: Position) -> CellSlot<'w>;
}

/// The public reader's store: every validated record is decoded into the
/// worksheet's cell map, and that map answers both questions.
pub(crate) struct DecodeEveryCell;

impl CellStore for DecodeEveryCell {
    fn is_occupied(&self, worksheet: &Worksheet, position: Position) -> bool {
        worksheet
            .get_cell(u32::from(position.0), u32::from(position.1))
            .is_some()
    }

    fn store(
        &mut self,
        worksheet: &mut Worksheet,
        position: Position,
        _formula: bool,
        decode: impl FnOnce() -> Cell,
    ) -> Result<()> {
        let cell = decode();
        debug_assert_eq!(
            (cell.row(), cell.column()),
            (u32::from(position.0), u32::from(position.1))
        );
        worksheet.add_cell(cell)
    }

    fn slot<'w>(&self, worksheet: &'w mut Worksheet, position: Position) -> CellSlot<'w> {
        match worksheet.get_cell_mut(u32::from(position.0), u32::from(position.1)) {
            None => CellSlot::Vacant,
            Some(cell) => CellSlot::Occupied {
                formula: cell.is_formula_record(),
                cell: Some(cell),
            },
        }
    }
}

/// The validation-only store: every validated record is counted in a compact
/// occupancy map, and only the positions a caller asked to keep are decoded.
///
/// A record is decoded into a cell by a conversion that cannot fail and has
/// no side effect, so counting a record instead of decoding it moves no
/// refusal. The occupancy map answers both questions for every position; a
/// kept position is additionally decoded, in record order, exactly as the
/// public reader decodes it, so its cell is the public reader's cell.
pub(crate) struct ValidateCells<'k> {
    kept: &'k [Position],
    occupancy: CellOccupancy,
}

impl<'k> ValidateCells<'k> {
    /// A store that decodes only the positions in `kept`, which must be
    /// sorted and free of duplicates.
    pub(crate) fn new(kept: &'k [Position]) -> Self {
        debug_assert!(kept.windows(2).all(|pair| pair[0] < pair[1]));
        Self {
            kept,
            occupancy: CellOccupancy::default(),
        }
    }

    fn keeps(&self, position: Position) -> bool {
        !self.kept.is_empty() && self.kept.binary_search(&position).is_ok()
    }
}

impl CellStore for ValidateCells<'_> {
    fn is_occupied(&self, _worksheet: &Worksheet, position: Position) -> bool {
        self.occupancy.get(position).is_some()
    }

    fn store(
        &mut self,
        worksheet: &mut Worksheet,
        position: Position,
        formula: bool,
        decode: impl FnOnce() -> Cell,
    ) -> Result<()> {
        self.occupancy.insert(position, formula)?;
        if self.keeps(position) {
            let cell = decode();
            debug_assert_eq!(
                (cell.row(), cell.column()),
                (u32::from(position.0), u32::from(position.1))
            );
            debug_assert_eq!(cell.is_formula_record(), formula);
            worksheet.add_cell(cell)?;
        }
        Ok(())
    }

    fn slot<'w>(&self, worksheet: &'w mut Worksheet, position: Position) -> CellSlot<'w> {
        match self.occupancy.get(position) {
            None => CellSlot::Vacant,
            Some(formula) => CellSlot::Occupied {
                formula,
                cell: if self.keeps(position) {
                    worksheet.get_cell_mut(u32::from(position.0), u32::from(position.1))
                } else {
                    None
                },
            },
        }
    }
}

/// Columns of the BIFF8 worksheet grid, the dense part of the map.
const GRID_COLUMNS: u16 = 256;

/// Two 256-column bitmaps for one row: occupied positions, and positions
/// whose latest record is a `Formula`.
#[derive(Clone, Copy, Default)]
struct RowBits {
    occupied: [u64; 4],
    formula: [u64; 4],
}

impl RowBits {
    /// `None` when `column` (inside the grid) is vacant, otherwise whether its
    /// latest record is a `Formula`.
    #[inline]
    fn get(&self, column: u16) -> Option<bool> {
        let word = usize::from(column / 64);
        let mask = 1_u64 << (column % 64);
        let occupied = self.occupied.get(word).is_some_and(|word| word & mask != 0);
        occupied.then(|| self.formula.get(word).is_some_and(|word| word & mask != 0))
    }

    /// Marks `column` (inside the grid) occupied, its latest record a
    /// `Formula` exactly when `formula`.
    #[inline]
    fn set(&mut self, column: u16, formula: bool) {
        let word = usize::from(column / 64);
        let mask = 1_u64 << (column % 64);
        if let Some(occupied) = self.occupied.get_mut(word) {
            *occupied |= mask;
        }
        if let Some(formulas) = self.formula.get_mut(word) {
            if formula {
                *formulas |= mask;
            } else {
                *formulas &= !mask;
            }
        }
    }
}

/// Occupied positions of one worksheet's cell records.
///
/// Every occupied row owns one [`RowBits`] (64 bytes), allocated fallibly when
/// the row's first record arrives, so the map is proportional to the rows that
/// hold a cell and never to the row range they span. A row is found through
/// the row most recently stored (the cell records of one row arrive together),
/// then through `ascending`, the rows first seen in ascending order, which is
/// how producers write them: a new row above every row seen so far is absent
/// by construction and is appended without a search or a hash. Only a row that
/// first appears below an already seen row — an unusual producer, or a hostile
/// stream — goes to `scattered`, a keyed hash map. The single-cell record
/// decoder does not bound a column to the 256-column grid, and the public
/// reader keeps such a cell like any other, so positions at column 256 or
/// beyond are counted in a small exact map instead.
#[derive(Default)]
struct CellOccupancy {
    /// Occupied and `Formula` masks, one entry per occupied row.
    rows: Vec<RowBits>,
    /// Rows first seen above every earlier row, with their entry in `rows`,
    /// in ascending row order.
    ascending: Vec<(u16, u32)>,
    /// Rows first seen below an earlier row, with their entry in `rows`.
    scattered: HashMap<u16, u32>,
    /// The row most recently stored and its entry in `rows`.
    recent: Option<(u16, u32)>,
    outside_grid: HashMap<Position, bool>,
}

impl CellOccupancy {
    /// `None` when no record occupies `position`, otherwise whether the latest
    /// record there is a `Formula`.
    ///
    /// The row most recently stored is answered inline; everything else takes
    /// the out-of-line lookup, which keeps the code this adds to the per-record
    /// worksheet walk small.
    #[inline]
    fn get(&self, (row, column): Position) -> Option<bool> {
        if column < GRID_COLUMNS
            && let Some((recent, entry)) = self.recent
            && recent == row
        {
            return self.bits(entry, column);
        }
        self.get_elsewhere((row, column))
    }

    /// Counts one record at `position`, replacing the `Formula` flag of an
    /// earlier record there: the latest record at a position is the one the
    /// public reader keeps.
    ///
    /// A record on the row most recently stored is counted inline; a first
    /// record on a row, and every position outside the grid, takes the
    /// out-of-line path.
    #[inline]
    fn insert(&mut self, (row, column): Position, formula: bool) -> Result<()> {
        if column < GRID_COLUMNS
            && let Some((recent, entry)) = self.recent
            && recent == row
            && let Some(bits) = usize::try_from(entry)
                .ok()
                .and_then(|entry| self.rows.get_mut(entry))
        {
            bits.set(column, formula);
            return Ok(());
        }
        self.insert_elsewhere((row, column), formula)
    }

    /// Whether `column` is occupied in the row at `entry`, and if so whether
    /// its latest record is a `Formula`.
    #[inline]
    fn bits(&self, entry: u32, column: u16) -> Option<bool> {
        self.rows.get(usize::try_from(entry).ok()?)?.get(column)
    }

    /// The entry in `rows` of an occupied `row`.
    fn entry(&self, row: u16) -> Option<u32> {
        if let Some((recent, entry)) = self.recent
            && recent == row
        {
            return Some(entry);
        }
        let &(highest, highest_entry) = self.ascending.last()?;
        if row == highest {
            return Some(highest_entry);
        }
        // A row above every ascending row has not been seen: rows reach
        // `scattered` only when they are below an ascending row.
        if row > highest {
            return None;
        }
        match self
            .ascending
            .binary_search_by_key(&row, |&(ascending, _)| ascending)
        {
            Ok(index) => self.ascending.get(index).map(|&(_, entry)| entry),
            Err(_) if self.scattered.is_empty() => None,
            Err(_) => self.scattered.get(&row).copied(),
        }
    }

    #[inline(never)]
    fn get_elsewhere(&self, (row, column): Position) -> Option<bool> {
        if column >= GRID_COLUMNS {
            return self.outside_grid.get(&(row, column)).copied();
        }
        self.bits(self.entry(row)?, column)
    }

    #[inline(never)]
    fn insert_elsewhere(&mut self, (row, column): Position, formula: bool) -> Result<()> {
        if column >= GRID_COLUMNS {
            self.outside_grid
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
            self.outside_grid.insert((row, column), formula);
            return Ok(());
        }
        let entry = match self.entry(row) {
            Some(entry) => entry,
            None => {
                let entry = u32::try_from(self.rows.len())
                    .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
                let ascending = self
                    .ascending
                    .last()
                    .is_none_or(|&(highest, _)| row > highest);
                // Reserve everything first, so a failed reservation leaves the
                // map exactly as it was.
                self.rows
                    .try_reserve(1)
                    .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
                if ascending {
                    self.ascending.try_reserve(1)
                } else {
                    self.scattered.try_reserve(1)
                }
                .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
                self.rows.push(RowBits::default());
                if ascending {
                    self.ascending.push((row, entry));
                } else {
                    self.scattered.insert(row, entry);
                }
                entry
            },
        };
        self.recent = Some((row, entry));
        usize::try_from(entry)
            .ok()
            .and_then(|entry| self.rows.get_mut(entry))
            .ok_or(Error::Allocation("tracking worksheet cell positions"))?
            .set(column, formula);
        Ok(())
    }

    /// Bytes this map holds, capacity included, for resource tests.
    #[cfg(test)]
    fn retained_bytes(&self) -> usize {
        self.rows.capacity() * size_of::<RowBits>()
            + self.ascending.capacity() * size_of::<(u16, u32)>()
            + self.scattered.capacity() * (size_of::<(u16, u32)>() + 1)
            + self.outside_grid.capacity() * (size_of::<(Position, bool)>() + 1)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn occupancy_tracks_the_latest_record_kind_per_position() {
        let mut occupancy = CellOccupancy::default();
        assert_eq!(occupancy.get((0, 0)), None);
        occupancy.insert((0, 0), false).unwrap();
        assert_eq!(occupancy.get((0, 0)), Some(false));
        occupancy.insert((0, 0), true).unwrap();
        assert_eq!(occupancy.get((0, 0)), Some(true));
        occupancy.insert((0, 0), false).unwrap();
        assert_eq!(occupancy.get((0, 0)), Some(false));
        assert_eq!(occupancy.get((0, 1)), None);
        assert_eq!(occupancy.get((1, 0)), None);
    }

    #[test]
    fn occupancy_covers_the_grid_edges_and_columns_outside_it() {
        let mut occupancy = CellOccupancy::default();
        let positions = [
            (0, 63),
            (0, 64),
            (255, 255),
            (256, 0),
            (u16::MAX, 255),
            (u16::MAX, 256),
            (7, u16::MAX),
        ];
        for (index, position) in positions.into_iter().enumerate() {
            occupancy.insert(position, index % 2 == 0).unwrap();
        }
        for (index, position) in positions.into_iter().enumerate() {
            assert_eq!(
                occupancy.get(position),
                Some(index % 2 == 0),
                "{position:?}"
            );
        }
        for vacant in [
            (0, 62),
            (0, 65),
            (255, 254),
            (1, 0),
            (7, 256),
            (u16::MAX, 257),
        ] {
            assert_eq!(occupancy.get(vacant), None, "{vacant:?}");
        }
        // One entry per occupied grid row; beyond-grid columns are exact.
        assert_eq!(occupancy.rows.len(), 4);
        assert_eq!(occupancy.outside_grid.len(), 2);
    }

    /// One cell per 256-row band, the layout that made a banded bitmap
    /// allocate and zero 16 KiB per cell: the map now holds one 64-byte entry
    /// per occupied row.
    #[test]
    fn occupancy_is_proportional_to_occupied_rows_for_sparse_bands() {
        let mut occupancy = CellOccupancy::default();
        for band in 0..256_u16 {
            occupancy.insert((band * 256, 0), band % 3 == 0).unwrap();
        }
        assert_eq!(occupancy.rows.len(), 256);
        assert!(occupancy.scattered.is_empty());
        for band in 0..256_u16 {
            assert_eq!(occupancy.get((band * 256, 0)), Some(band % 3 == 0));
            assert_eq!(occupancy.get((band * 256 + 1, 0)), None);
            assert_eq!(occupancy.get((band * 256, 1)), None);
        }
        // 256 rows of 64-byte masks plus their 8-byte index entries, with at
        // most doubling slack: well under 64 KiB, where the banded bitmap
        // held 256 bands of 16 KiB (4 MiB).
        let retained = occupancy.retained_bytes();
        assert!(retained <= 2 * 256 * (64 + 8), "retained {retained} bytes");
    }

    /// Rows first seen in descending or shuffled order go through the keyed
    /// map and answer exactly like ascending ones; memory stays one entry per
    /// occupied row.
    #[test]
    fn occupancy_answers_rows_in_any_arrival_order() {
        let mut occupancy = CellOccupancy::default();
        let mut rows = (0..2_048_u16).map(|row| row * 31).collect::<Vec<_>>();
        rows.reverse();
        rows.rotate_left(700);
        for (index, row) in rows.iter().enumerate() {
            occupancy.insert((*row, 5), index % 2 == 0).unwrap();
            occupancy.insert((*row, 200), false).unwrap();
        }
        assert_eq!(occupancy.rows.len(), rows.len());
        assert!(!occupancy.scattered.is_empty());
        for (index, row) in rows.iter().enumerate() {
            assert_eq!(occupancy.get((*row, 5)), Some(index % 2 == 0), "{row}");
            assert_eq!(occupancy.get((*row, 200)), Some(false));
            assert_eq!(occupancy.get((*row, 6)), None);
            assert_eq!(occupancy.get((row.wrapping_add(1), 5)), None);
        }
        let retained = occupancy.retained_bytes();
        assert!(
            retained <= 2 * rows.len() * (64 + 16),
            "retained {retained} bytes"
        );
    }

    #[test]
    fn validation_store_decodes_only_kept_positions() {
        let mut worksheet = Worksheet::new("Sheet".to_string());
        let kept = [(1, 1)];
        let mut store = ValidateCells::new(&kept);
        let mut decoded = Vec::new();
        for position in [(0_u16, 0_u16), (1, 1), (2, 2)] {
            store
                .store(&mut worksheet, position, false, || {
                    decoded.push(position);
                    Cell::new(
                        u32::from(position.0),
                        u32::from(position.1),
                        litchi_core::sheet::CellValue::Float(1.0),
                    )
                })
                .unwrap();
        }
        assert_eq!(decoded, vec![(1, 1)]);
        assert!(store.is_occupied(&worksheet, (0, 0)));
        assert!(store.is_occupied(&worksheet, (2, 2)));
        assert!(!store.is_occupied(&worksheet, (3, 3)));
        assert!(worksheet.get_cell(1, 1).is_some());
        assert!(worksheet.get_cell(0, 0).is_none());
        assert!(matches!(
            store.slot(&mut worksheet, (0, 0)),
            CellSlot::Occupied {
                formula: false,
                cell: None
            }
        ));
        assert!(matches!(
            store.slot(&mut worksheet, (1, 1)),
            CellSlot::Occupied {
                formula: false,
                cell: Some(_)
            }
        ));
        assert!(matches!(
            store.slot(&mut worksheet, (9, 9)),
            CellSlot::Vacant
        ));
    }
}
