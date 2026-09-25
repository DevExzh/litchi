//! The shared-string table staged for one workbook write.
//!
//! The table borrows every string from the worksheet cells for the duration of
//! the write, so no string is copied, and it hashes each string cell once:
//! while it is built, which also records every cell's index for the cell
//! records to reuse.
//!
//! Its order is part of the output and does not depend on any hash seed: a
//! string's index is its first occurrence in worksheet order and, within a
//! worksheet, in row and then column order, which is the order the cell
//! records are written in.

use std::collections::HashMap;
use std::collections::hash_map::{Entry, RandomState};
use std::hash::{BuildHasher, Hash, Hasher};

use crate::writer::string_limits::{SHARED_STRING_UNITS, ensure_utf16_len_within};
use crate::{Error, Result};

use super::super::model::CellValue;
use super::super::worksheet::WritableWorksheet;

/// Unique shared strings in first-occurrence order and their SST indices.
pub(crate) struct SharedStringTable<'a> {
    /// Unique strings in the order the SST record lists them.
    strings: Vec<&'a str>,
    /// Index of every unique string, keyed by its content and cached hash.
    indices: HashMap<HashedStr<'a>, u32, CachedHashState>,
    /// The keyed hash computed once per string; see [`HashedStr`].
    hasher: RandomState,
    /// Every string cell, counting repeats (`SST.cstTotal`).
    total: u32,
    /// Per worksheet, the SST index of each string cell in row and then
    /// column order.
    cell_indices: Vec<Vec<u32>>,
}

impl<'a> SharedStringTable<'a> {
    /// Collects the string cells of `worksheets` in worksheet order and, within
    /// a worksheet, in row and then column order, so each string's index is
    /// its first occurrence in the order the cell records are written.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StringTooLong`] for a string longer than an SST entry
    /// can hold, before anything is written.
    pub(crate) fn build(worksheets: &'a [WritableWorksheet]) -> Result<Self> {
        let mut table = Self {
            strings: Vec::new(),
            indices: HashMap::with_hasher(CachedHashState),
            hasher: RandomState::new(),
            total: 0,
            cell_indices: Vec::with_capacity(worksheets.len()),
        };
        // One worksheet's string cells at a time, reused across worksheets,
        // each keyed by `row << 16 | column`: a column is below 2^16, so the
        // packed keys order exactly as `(row, column)` pairs do, and one
        // integer comparison orders two cells.
        let mut string_cells: Vec<(u64, &'a str)> = Vec::new();
        for worksheet in worksheets {
            string_cells.clear();
            string_cells.extend(worksheet.cells.iter().filter_map(|(&(row, column), cell)| {
                match &cell.value {
                    CellValue::String(text) => {
                        Some(((u64::from(row) << 16) | u64::from(column), text.as_str()))
                    },
                    _ => None,
                }
            }));
            // Keys are distinct, so an unstable sort gives the one row-major
            // order whatever order the cell map iterated in.
            string_cells.sort_unstable_by_key(|&(key, _)| key);

            let mut sheet_indices = Vec::with_capacity(string_cells.len());
            for &(_, text) in &string_cells {
                table.total = table.total.saturating_add(1);
                let next = crate::utils::truncate_usize_to_u32(table.strings.len());
                let key = table.key(text);
                let index = match table.indices.entry(key) {
                    Entry::Vacant(slot) => {
                        ensure_utf16_len_within(text, SHARED_STRING_UNITS, "shared string")?;
                        slot.insert(next);
                        table.strings.push(text);
                        next
                    },
                    Entry::Occupied(entry) => *entry.get(),
                };
                sheet_indices.push(index);
            }
            table.cell_indices.push(sheet_indices);
        }
        Ok(table)
    }

    fn key(&self, text: &'a str) -> HashedStr<'a> {
        HashedStr {
            hash: self.hasher.hash_one(text),
            text,
        }
    }

    /// Unique strings in SST order.
    pub(crate) fn strings(&self) -> &[&'a str] {
        &self.strings
    }

    /// Number of string cells, counting repeats.
    pub(crate) const fn total(&self) -> u32 {
        self.total
    }

    /// Upper bound of the SST and CONTINUE record bytes for this table, used
    /// only to reserve the workbook stream.
    ///
    /// Each string takes a three-byte header and at most 0xFFFF characters
    /// (longer ones are refused), one byte each when ASCII and two per UTF-16
    /// code unit otherwise (never more than two per UTF-8 byte); every
    /// 8,224-byte record adds a four-byte header and at most one continuation
    /// flag byte.
    pub(crate) fn sst_bytes_hint(&self) -> usize {
        let payload = self.strings.iter().fold(8usize, |bytes, text| {
            let characters = if text.is_ascii() {
                text.len().min(0xFFFF)
            } else {
                text.len().min(0xFFFF).saturating_mul(2)
            };
            bytes.saturating_add(3).saturating_add(characters)
        });
        payload.saturating_add((payload / 8_223).saturating_add(1).saturating_mul(5))
    }

    /// Whether worksheet `sheet` had any string cell when the table was built.
    pub(crate) fn has_string_cells(&self, sheet: usize) -> bool {
        self.cell_indices
            .get(sheet)
            .is_some_and(|cells| !cells.is_empty())
    }

    /// Returns the SST index of the `string_ordinal`-th string cell, in row
    /// and then column order, of worksheet `sheet`, whose value is `value`.
    ///
    /// The index recorded for that cell while the table was built is returned
    /// when the table's string at that index equals `value`: strings in the
    /// table are distinct, so the equal one is `value`'s own index. Only if the
    /// caller counted string cells differently, or the cell is not the one
    /// recorded, does this hash `value` again through [`Self::index_of`].
    pub(crate) fn index_for(
        &self,
        sheet: usize,
        string_ordinal: usize,
        value: &str,
    ) -> Result<u32> {
        let recorded = self
            .cell_indices
            .get(sheet)
            .and_then(|cells| cells.get(string_ordinal))
            .copied();
        if let Some(index) = recorded
            && usize::try_from(index)
                .ok()
                .and_then(|index| self.strings.get(index))
                .is_some_and(|entry| *entry == value)
        {
            return Ok(index);
        }
        self.index_of(value)
    }

    /// Returns the SST index of a string cell's value.
    ///
    /// The index is also checked against the table: the entry it names must
    /// be the very string the index was recorded for.
    pub(crate) fn index_of(&self, value: &str) -> Result<u32> {
        let probe = HashedStr {
            hash: self.hasher.hash_one(value),
            text: value,
        };
        let (key, &index) = self.indices.get_key_value(&probe).ok_or_else(|| {
            Error::InvalidData(format!(
                "string cell value {value:?} is missing from the shared string table"
            ))
        })?;
        let table_index = usize::try_from(index).map_err(|_error| {
            Error::InvalidData(format!(
                "shared string index {index} for value {value:?} cannot be represented"
            ))
        })?;
        match self.strings.get(table_index) {
            Some(entry) if std::ptr::eq(*entry, key.text) => Ok(index),
            Some(_) => Err(Error::InvalidData(format!(
                "shared string index {index} for value {value:?} does not match the shared string table"
            ))),
            None => Err(Error::InvalidData(format!(
                "shared string index {index} for value {value:?} is outside the shared string table"
            ))),
        }
    }
}

/// A shared string together with its hash, computed once.
///
/// `hash` is `RandomState::hash_one(text)`: the same randomly keyed SipHash-1-3
/// that a `HashMap<&str, _>` computes for the string on every lookup, insert
/// and resize. The table's map uses this value as the bucket hash as-is
/// ([`CachedHashState`]), so hash-flooding resistance is exactly that of the
/// standard map; what changes is only that a string is hashed once per cell
/// rather than again on each map operation and again when the map grows.
struct HashedStr<'a> {
    hash: u64,
    text: &'a str,
}

impl PartialEq for HashedStr<'_> {
    fn eq(&self, other: &Self) -> bool {
        // Equal strings always have equal hashes (one keyed hasher computes
        // every key of a table), so comparing hashes first is only a shortcut.
        self.hash == other.hash && self.text == other.text
    }
}

impl Eq for HashedStr<'_> {}

impl Hash for HashedStr<'_> {
    fn hash<H: Hasher>(&self, state: &mut H) {
        state.write_u64(self.hash);
    }
}

/// Builds [`CachedHash`] hashers for the table's map.
#[derive(Clone, Copy, Default)]
struct CachedHashState;

impl BuildHasher for CachedHashState {
    type Hasher = CachedHash;

    fn build_hasher(&self) -> CachedHash {
        CachedHash(0)
    }
}

/// Passes a [`HashedStr`]'s cached keyed hash through unchanged.
struct CachedHash(u64);

impl Hasher for CachedHash {
    fn finish(&self) -> u64 {
        self.0
    }

    fn write(&mut self, bytes: &[u8]) {
        // `HashedStr` only ever writes its cached `u64`; fold anything else so
        // the hasher stays total.
        for &byte in bytes {
            self.0 = self.0.rotate_left(8) ^ u64::from(byte);
        }
    }

    fn write_u64(&mut self, value: u64) {
        self.0 = value;
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::super::super::{CellPos, WritableCell};
    use super::*;
    use crate::writer::Writer;

    /// The table this module promises, computed without it: every string
    /// cell in worksheet, row and column order, each distinct string indexed
    /// at its first occurrence.
    fn reference_table(
        worksheets: &[WritableWorksheet],
    ) -> (Vec<String>, HashMap<String, u32>, u32, Vec<Vec<u32>>) {
        let mut strings = Vec::new();
        let mut map = HashMap::new();
        let mut total = 0u32;
        let mut cell_indices = Vec::new();
        for worksheet in worksheets {
            let mut keys: Vec<(u32, u16)> = worksheet.cells.keys().copied().collect();
            keys.sort_unstable();
            let mut sheet = Vec::new();
            for key in keys {
                if let CellValue::String(text) = &worksheet.cells[&key].value {
                    total += 1;
                    let index = *map.entry(text.clone()).or_insert_with(|| {
                        strings.push(text.clone());
                        u32::try_from(strings.len() - 1).unwrap()
                    });
                    sheet.push(index);
                }
            }
            cell_indices.push(sheet);
        }
        (strings, map, total, cell_indices)
    }

    fn mixed_workbook() -> Writer {
        let mut writer = Writer::new();
        let first = writer.add_worksheet("First").unwrap();
        let second = writer.add_worksheet("Second").unwrap();
        let long_ascii = "long ascii payload ".repeat(2_000);
        let long_unicode = "漢字 ✓ 😀 ".repeat(3_000);
        let values = [
            "",
            "repeat",
            "Latin-1 àé",
            "漢字",
            "😀𝄞",
            "repeat",
            "",
            "tail",
        ];
        for (row, value) in values.iter().enumerate() {
            let row = u32::try_from(row).unwrap();
            writer.write_string(first, row, 0, value).unwrap();
            writer.write_string(second, row, 1, value).unwrap();
            writer.write_number(first, row, 2, f64::from(row)).unwrap();
        }
        writer.write_string(first, 20, 3, &long_ascii).unwrap();
        writer.write_string(second, 21, 4, &long_ascii).unwrap();
        writer.write_string(second, 22, 5, &long_unicode).unwrap();
        // Written bottom-up and right to left, so neither insertion order
        // nor any map order is the row-major order.
        for index in (0..300u32).rev() {
            writer
                .write_string(
                    first,
                    100 + index / 3,
                    6 + u16::try_from(index % 3).unwrap(),
                    &format!("distinct {index} ü"),
                )
                .unwrap();
        }
        writer
    }

    #[test]
    fn strings_are_listed_at_their_first_occurrence_in_row_and_column_order() {
        let mut writer = Writer::new();
        let first = writer.add_worksheet("First").unwrap();
        let second = writer.add_worksheet("Second").unwrap();
        // Inserted in neither row-major nor reverse order.
        writer.write_string(first, 2, 5, "c").unwrap();
        writer.write_string(first, 0, 1, "b").unwrap();
        writer.write_string(second, 5, 5, "d").unwrap();
        writer.write_string(first, 1, 3, "a").unwrap();
        writer.write_string(first, 0, 0, "a").unwrap();
        writer.write_number(first, 0, 2, 1.0).unwrap();
        writer.write_string(second, 0, 1, "a").unwrap();
        writer.write_string(first, 1, 0, "b").unwrap();
        writer.write_string(second, 0, 0, "c").unwrap();

        let table = SharedStringTable::build(&writer.worksheets).unwrap();

        // First: A1 "a", B1 "b", A2 "b", D2 "a", F3 "c"; Second: A1 "c",
        // B1 "a", F6 "d".
        assert_eq!(table.strings(), ["a", "b", "c", "d"]);
        assert_eq!(table.total(), 8);
        assert_eq!(table.cell_indices, [vec![0, 1, 1, 0, 2], vec![2, 0, 3]]);
    }

    #[test]
    fn table_matches_the_reference_order_indices_and_total() {
        let writer = mixed_workbook();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();
        let (strings, map, total, cell_indices) = reference_table(&writer.worksheets);

        assert_eq!(
            table.strings(),
            strings.iter().map(String::as_str).collect::<Vec<_>>()
        );
        assert_eq!(table.total(), total);
        assert_eq!(table.total(), 2 * 8 + 3 + 300);
        assert_eq!(table.cell_indices, cell_indices);
        for worksheet in &writer.worksheets {
            for cell in worksheet.cells.values() {
                if let CellValue::String(text) = &cell.value {
                    assert_eq!(table.index_of(text).unwrap(), map[text]);
                }
            }
        }
    }

    #[test]
    fn the_same_cells_in_maps_with_different_seeds_stage_the_same_table() {
        // Each writer's cell maps are seeded independently, so their
        // iteration orders differ; the table must not.
        let tables: Vec<_> = (0..8).map(|_| mixed_workbook()).collect();
        let reference = SharedStringTable::build(&tables[0].worksheets).unwrap();
        for writer in &tables[1..] {
            let table = SharedStringTable::build(&writer.worksheets).unwrap();
            assert_eq!(table.strings(), reference.strings());
            assert_eq!(table.total(), reference.total());
            assert_eq!(table.cell_indices, reference.cell_indices);
        }
    }

    #[test]
    fn recorded_cell_indices_follow_row_and_column_order() {
        let writer = mixed_workbook();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();
        let (_, map, _, _) = reference_table(&writer.worksheets);

        assert_eq!(table.cell_indices.len(), writer.worksheets.len());
        for (sheet, worksheet) in writer.worksheets.iter().enumerate() {
            let mut cells: Vec<_> = worksheet.cells.iter().collect();
            cells.sort_unstable_by_key(|(key, _)| **key);
            let strings: Vec<&str> = cells
                .iter()
                .filter_map(|(_, cell)| match &cell.value {
                    CellValue::String(text) => Some(text.as_str()),
                    _ => None,
                })
                .collect();
            assert_eq!(table.cell_indices[sheet].len(), strings.len());
            for (ordinal, text) in strings.iter().enumerate() {
                assert_eq!(table.cell_indices[sheet][ordinal], map[*text]);
                assert_eq!(table.index_for(sheet, ordinal, text).unwrap(), map[*text]);
            }
        }
    }

    #[test]
    fn a_recorded_index_for_another_value_falls_back_to_the_hashed_lookup() {
        let writer = mixed_workbook();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();
        let (_, map, _, _) = reference_table(&writer.worksheets);

        // Every ordinal of every sheet, asked for every value: the recorded
        // index is used only when it names that value, and the answer is
        // always the value's own index.
        for sheet in 0..writer.worksheets.len() {
            for ordinal in 0..=table.cell_indices[sheet].len() {
                for value in ["repeat", "漢字", "", "tail", "distinct 7 ü"] {
                    assert_eq!(table.index_for(sheet, ordinal, value).unwrap(), map[value]);
                }
            }
        }
        assert_eq!(table.index_for(99, 0, "tail").unwrap(), map["tail"]);
        assert!(matches!(
            table.index_for(0, 0, "absent"),
            Err(Error::InvalidData(message)) if message.contains("missing from the shared string table")
        ));
    }

    #[test]
    fn the_sst_size_hint_bounds_the_written_records() {
        let writer = mixed_workbook();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();
        let mut sst = Vec::new();
        crate::writer::biff::write_sst(&mut sst, table.strings(), table.total()).unwrap();

        assert!(
            sst.len() <= table.sst_bytes_hint(),
            "{} > {}",
            sst.len(),
            table.sst_bytes_hint()
        );
    }

    #[test]
    fn a_string_longer_than_an_sst_entry_is_refused_while_staging() {
        // Placed behind the writer's own check, as a pivot label or table
        // header could be: the table must refuse it rather than stage it.
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Long").unwrap();
        writer.write_string(sheet, 0, 0, "short").unwrap();
        for (units, fits) in [(0xFFFE, true), (0xFFFF, true), (0x1_0000, false)] {
            for value in ["x".repeat(units), "é".repeat(units)] {
                let cell = WritableCell::new(
                    CellPos::try_new(3, 1).unwrap(),
                    CellValue::String(value.clone()),
                    0,
                    None,
                );
                writer.worksheets[sheet].add_cell(cell);
                let result = SharedStringTable::build(&writer.worksheets);
                if fits {
                    let table = result.unwrap();
                    assert_eq!(table.strings(), ["short", value.as_str()]);
                } else {
                    assert!(matches!(
                        result,
                        Err(Error::StringTooLong {
                            field: "shared string",
                            utf16_units: 0x1_0000,
                            limit: 0xFFFF,
                        })
                    ));
                }
            }
        }
    }

    #[test]
    fn repeated_strings_in_distinct_allocations_share_one_entry() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        writer.write_string(sheet, 0, 1, "Hello").unwrap();
        writer.write_string(sheet, 1, 0, "World").unwrap();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();

        assert_eq!(table.strings(), ["Hello", "World"]);
        assert_eq!(table.total(), 3);
        let owned = String::from("Hello");
        let hello = table.index_of(&owned).unwrap();
        assert_eq!(table.strings()[hello as usize], "Hello");
    }

    #[test]
    fn a_workbook_without_string_cells_stages_an_empty_table() {
        let mut writer = Writer::new();
        let sheet = writer.add_worksheet("Numbers").unwrap();
        writer.write_number(sheet, 0, 0, 1.0).unwrap();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();

        assert!(table.strings().is_empty());
        assert_eq!(table.total(), 0);
        assert!(!table.has_string_cells(sheet));
    }

    #[test]
    fn a_missing_value_is_refused() {
        let writer = mixed_workbook();
        let table = SharedStringTable::build(&writer.worksheets).unwrap();

        let result = table.index_of("absent");

        assert!(matches!(
            result,
            Err(Error::InvalidData(message)) if message.contains("missing from the shared string table")
        ));
    }

    fn inconsistent_table<'a>(
        strings: Vec<&'a str>,
        key: &'a str,
        index: u32,
    ) -> SharedStringTable<'a> {
        let mut table = SharedStringTable {
            strings,
            indices: HashMap::with_hasher(CachedHashState),
            hasher: RandomState::new(),
            total: 1,
            cell_indices: Vec::new(),
        };
        let key = table.key(key);
        table.indices.insert(key, index);
        table
    }

    #[test]
    fn an_index_naming_another_entry_is_refused() {
        let first = String::from("first");
        let second = String::from("second");
        let table = inconsistent_table(vec![&first, &second], &second, 0);

        let result = table.index_of("second");

        assert!(matches!(
            result,
            Err(Error::InvalidData(message)) if message.contains("does not match the shared string table")
        ));
    }

    #[test]
    fn an_equal_string_in_another_allocation_does_not_satisfy_the_entry_check() {
        // The check is identity, not content: the index must name the very
        // string it was recorded for.
        let recorded = String::from("same");
        let listed = String::from("same");
        let table = inconsistent_table(vec![&listed], &recorded, 0);

        assert!(table.index_of("same").is_err());
    }

    #[test]
    fn an_index_past_the_table_is_refused() {
        let only = String::from("only");
        let table = inconsistent_table(vec![&only], &only, 7);

        let result = table.index_of("only");

        assert!(matches!(
            result,
            Err(Error::InvalidData(message)) if message.contains("is outside the shared string table")
        ));
    }

    #[test]
    fn colliding_cached_hashes_keep_distinct_strings_apart() {
        let mut map = HashMap::with_hasher(CachedHashState);
        map.insert(HashedStr { hash: 7, text: "a" }, 0u32);
        map.insert(HashedStr { hash: 7, text: "b" }, 1u32);
        map.insert(HashedStr { hash: 7, text: "a" }, 2u32);

        assert_eq!(map.len(), 2);
        assert_eq!(map.get(&HashedStr { hash: 7, text: "a" }), Some(&2));
        assert_eq!(map.get(&HashedStr { hash: 7, text: "b" }), Some(&1));
        assert_eq!(map.get(&HashedStr { hash: 8, text: "a" }), None);
    }

    #[test]
    fn the_cached_hash_is_the_keyed_hash_of_the_string() {
        let table = SharedStringTable::build(&[]).unwrap();
        let key = table.key("value");

        assert_eq!(key.hash, table.hasher.hash_one("value"));
        assert_eq!(CachedHashState.hash_one(&key), key.hash);
    }
}
