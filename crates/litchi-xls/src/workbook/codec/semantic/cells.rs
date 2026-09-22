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

/// Rows per lazily allocated occupancy band.
const BAND_ROWS: usize = 256;
/// Columns of the BIFF8 worksheet grid, the dense part of the map.
const GRID_COLUMNS: u16 = 256;

/// Two 256-column bitmaps for one row: occupied positions, and positions
/// whose latest record is a `Formula`.
#[derive(Clone, Copy, Default)]
struct RowBits {
    occupied: [u64; 4],
    formula: [u64; 4],
}

/// Occupied positions of one worksheet's cell records.
///
/// Rows are grouped into 256-row bands of two bitmaps each (16 KiB per band),
/// allocated fallibly when a band is first touched, so a lookup is two
/// indexed loads and the map is bounded by the 65,536-row BIFF8 grid. The
/// single-cell record decoder does not bound a column to the 256-column grid,
/// and the public reader keeps such a cell like any other, so positions at
/// column 256 or beyond are counted in a small exact map instead.
#[derive(Default)]
struct CellOccupancy {
    bands: Vec<Option<Box<[RowBits]>>>,
    outside_grid: HashMap<Position, bool>,
}

impl CellOccupancy {
    /// `None` when no record occupies `position`, otherwise whether the latest
    /// record there is a `Formula`.
    fn get(&self, (row, column): Position) -> Option<bool> {
        if column >= GRID_COLUMNS {
            return self.outside_grid.get(&(row, column)).copied();
        }
        let band = self.bands.get(usize::from(row) / BAND_ROWS)?.as_ref()?;
        let bits = band.get(usize::from(row) % BAND_ROWS)?;
        let word = usize::from(column / 64);
        let mask = 1_u64 << (column % 64);
        let occupied = bits.occupied.get(word).is_some_and(|word| word & mask != 0);
        occupied.then(|| bits.formula.get(word).is_some_and(|word| word & mask != 0))
    }

    /// Counts one record at `position`, replacing the `Formula` flag of an
    /// earlier record there: the latest record at a position is the one the
    /// public reader keeps.
    fn insert(&mut self, (row, column): Position, formula: bool) -> Result<()> {
        if column >= GRID_COLUMNS {
            self.outside_grid
                .try_reserve(1)
                .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
            self.outside_grid.insert((row, column), formula);
            return Ok(());
        }
        let band_index = usize::from(row) / BAND_ROWS;
        if self.bands.len() <= band_index {
            self.bands
                .try_reserve(band_index + 1 - self.bands.len())
                .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
            self.bands.resize_with(band_index + 1, || None);
        }
        let slot = self
            .bands
            .get_mut(band_index)
            .ok_or(Error::Allocation("tracking worksheet cell positions"))?;
        let band = match slot {
            Some(band) => band,
            None => {
                let mut rows = Vec::new();
                rows.try_reserve_exact(BAND_ROWS)
                    .map_err(|_error| Error::Allocation("tracking worksheet cell positions"))?;
                rows.resize(BAND_ROWS, RowBits::default());
                slot.insert(rows.into_boxed_slice())
            },
        };
        let bits = band
            .get_mut(usize::from(row) % BAND_ROWS)
            .ok_or(Error::Allocation("tracking worksheet cell positions"))?;
        let word = usize::from(column / 64);
        let mask = 1_u64 << (column % 64);
        if let Some(occupied) = bits.occupied.get_mut(word) {
            *occupied |= mask;
        }
        if let Some(formulas) = bits.formula.get_mut(word) {
            if formula {
                *formulas |= mask;
            } else {
                *formulas &= !mask;
            }
        }
        Ok(())
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
        assert_eq!(occupancy.bands.len(), 256);
        assert_eq!(
            occupancy.bands.iter().filter(|band| band.is_some()).count(),
            3
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
