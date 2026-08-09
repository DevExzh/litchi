//! Source-bound `BrtCellReal` snapshots and length-stable edits.

use crate::package::error::{Error, Result};
use crate::raw::{Header, Limits, Records, kind};
use litchi_core::binary;
use std::sync::Arc;

const MAX_ROW: u32 = 1_048_575;
const MAX_COLUMN: u32 = 16_383;

/// A checked zero-based worksheet cell reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub struct Reference {
    row: u32,
    column: u32,
}

impl Reference {
    /// Construct a reference in the XLSB worksheet grid.
    ///
    /// # Errors
    ///
    /// Returns an error when either coordinate is outside the XLSB grid.
    pub fn new(row: u32, column: u32) -> Result<Self> {
        if row > MAX_ROW || column > MAX_COLUMN {
            return Err(Error::InvalidCellReference(format!(
                "cell ({row}, {column}) is outside the XLSB worksheet grid"
            )));
        }
        Ok(Self { row, column })
    }

    /// Return the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Return the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }
}

/// An existing `BrtCellReal` value in source order.
#[derive(Debug, Clone, Copy)]
pub struct Number {
    reference: Reference,
    value: f64,
}

impl Number {
    /// Cell location selected by this item.
    #[must_use]
    pub const fn reference(self) -> Reference {
        self.reference
    }

    /// Stored IEEE-754 value, retained exactly on read.
    #[must_use]
    pub const fn value(self) -> f64 {
        self.value
    }
}

impl PartialEq for Number {
    fn eq(&self, other: &Self) -> bool {
        self.reference == other.reference && self.value.to_bits() == other.value.to_bits()
    }
}

/// Immutable, source-bound snapshot of editable numeric cells.
#[derive(Debug, Clone)]
pub struct Snapshot {
    source: Arc<[u8]>,
    numbers: Vec<Entry>,
}

impl Snapshot {
    /// Existing `BrtCellReal` values in worksheet stream order.
    #[must_use]
    pub fn numbers(&self) -> impl ExactSizeIterator<Item = Number> + '_ {
        self.numbers.iter().map(|entry| entry.number)
    }

    /// Look up an existing numeric cell. Duplicate source records are refused
    /// because an edit cannot safely infer which producer record was intended.
    ///
    /// # Errors
    ///
    /// Returns an error when the source has multiple editable numeric records
    /// for the requested coordinate.
    pub fn number(&self, reference: Reference) -> Result<Option<Number>> {
        unique_number(&self.numbers, reference)
    }

    /// Start a detached edit against this exact source stream.
    #[must_use]
    pub fn edit(&self) -> Edit {
        Edit {
            source: Arc::clone(&self.source),
            numbers: self.numbers.clone(),
        }
    }

    /// Exact worksheet bytes guarded by commits and patches.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// Detached edits of existing `BrtCellReal` fields only.
#[derive(Debug, Clone)]
pub struct Edit {
    source: Arc<[u8]>,
    numbers: Vec<Entry>,
}

impl Edit {
    /// Set one existing `BrtCellReal` value.
    ///
    /// The replacement must be finite. Cells encoded as `BrtCellRk`, formulas,
    /// strings, and every other XLSB cell family are intentionally unsupported.
    ///
    /// # Errors
    ///
    /// Returns an error for non-finite values, duplicate source records, or a
    /// coordinate that is absent or represented by another cell-record family.
    pub fn set_number(&mut self, reference: Reference, value: f64) -> Result<()> {
        if !value.is_finite() {
            return Err(Error::UnsupportedFeature(
                "BrtCellReal edits require a finite IEEE-754 value".to_string(),
            ));
        }
        let selected = unique_index(&self.numbers, reference)?;
        let Some(entry_index) = selected else {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) is absent or is not encoded as BrtCellReal",
                reference.row(),
                reference.column()
            )));
        };
        self.numbers[entry_index].number.value = value;
        Ok(())
    }

    /// Validate and publish a source-checked reversible patch.
    ///
    /// # Errors
    ///
    /// Returns an error if the bounded replacement range cannot be represented
    /// or the generated source stream fails structural readback.
    pub fn commit(self) -> Result<Commit> {
        let changed = self.numbers.iter().any(Entry::changed);
        let after = if changed {
            let mut bytes = self.source.to_vec();
            for entry in &self.numbers {
                if entry.changed() {
                    let value_end =
                        entry
                            .value_offset
                            .checked_add(8)
                            .ok_or(Error::CapacityOverflow {
                                resource: "BrtCellReal replacement range",
                            })?;
                    let destination = bytes.get_mut(entry.value_offset..value_end).ok_or(
                        Error::InvalidFormat(
                            "BrtCellReal replacement is outside its source worksheet stream"
                                .to_string(),
                        ),
                    )?;
                    destination.copy_from_slice(&entry.number.value.to_le_bytes());
                }
            }
            Arc::from(bytes)
        } else {
            Arc::clone(&self.source)
        };
        let snapshot = read_shared(Arc::clone(&after))?;
        Ok(Commit {
            snapshot,
            patch: Patch {
                before: self.source,
                after,
            },
        })
    }
}

/// Successful immutable numeric-cell commit.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Planned immutable worksheet snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Split this result into the planned snapshot and patch.
    #[must_use]
    pub fn into_parts(self) -> (Snapshot, Patch) {
        (self.snapshot, self.patch)
    }
}

/// Reversible, source-checked worksheet-stream patch.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Arc<[u8]>,
    after: Arc<[u8]>,
}

impl Patch {
    /// Whether this patch leaves the source stream untouched.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.before == self.after
    }

    /// Exact source required to apply this patch.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        &self.before
    }

    /// Exact stream produced by this patch.
    #[must_use]
    pub fn after(&self) -> &[u8] {
        &self.after
    }

    /// Apply only to the exact source stream.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` does not equal this patch's before image.
    pub fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before.as_ref() {
            return Err(Error::UnsupportedFeature(
                "cell-value patch source snapshot does not match".to_string(),
            ));
        }
        Ok(self.after.to_vec())
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }
}

#[derive(Debug, Clone)]
struct Entry {
    number: Number,
    original_bits: u64,
    value_offset: usize,
}

impl Entry {
    fn changed(&self) -> bool {
        self.original_bits != self.number.value.to_bits()
    }
}

/// Read one complete worksheet stream without interpreting unsupported cells.
pub(super) fn read(data: &[u8]) -> Result<Snapshot> {
    read_shared(Arc::from(data))
}

fn read_shared(source: Arc<[u8]>) -> Result<Snapshot> {
    let mut numbers = Vec::new();
    let mut in_sheet_data = false;
    let mut current_row = None;
    let mut frt_depth = 0usize;

    for item in Records::new(&source) {
        let record = item?;
        match record.kind() {
            kind::BEGIN_SHEET_DATA => {
                if in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "duplicate BrtBeginSheetData record".to_string(),
                    ));
                }
                in_sheet_data = true;
                current_row = None;
                frt_depth = 0;
            },
            kind::END_SHEET_DATA => {
                if !in_sheet_data {
                    return Err(Error::InvalidFormat(
                        "BrtEndSheetData without BrtBeginSheetData".to_string(),
                    ));
                }
                in_sheet_data = false;
                frt_depth = 0;
            },
            kind::FRT_BEGIN if in_sheet_data => {
                frt_depth = frt_depth.checked_add(1).ok_or(Error::CapacityOverflow {
                    resource: "worksheet FRT nesting depth",
                })?;
            },
            kind::FRT_END if in_sheet_data => {
                frt_depth = frt_depth.saturating_sub(1);
            },
            kind::ROW_HDR if in_sheet_data && frt_depth == 0 => {
                if record.payload().len() < 17 {
                    return Err(Error::InvalidLength {
                        expected: 17,
                        found: record.payload().len(),
                    });
                }
                let row = binary::read_u32_le_at(record.payload(), 0)?;
                if row > MAX_ROW {
                    return Err(Error::InvalidCellReference(format!(
                        "BrtRowHdr row {row} is outside the XLSB worksheet grid"
                    )));
                }
                current_row = Some(row);
            },
            kind::CELL_REAL if in_sheet_data && frt_depth == 0 => {
                if record.payload().len() != 16 {
                    return Err(Error::InvalidLength {
                        expected: 16,
                        found: record.payload().len(),
                    });
                }
                let row = current_row.ok_or_else(|| {
                    Error::InvalidFormat("BrtCellReal appears before a BrtRowHdr".to_string())
                })?;
                let column = binary::read_u32_le_at(record.payload(), 0)?;
                let reference = Reference::new(row, column)?;
                let value = binary::read_f64_le_at(record.payload(), 8)?;
                let record_source = source.get(record.offset()..).ok_or(Error::InvalidFormat(
                    "BrtCellReal record offset is outside its source worksheet stream".to_string(),
                ))?;
                let (_, header_len) = Header::parse(record_source, Limits::DEFAULT)?;
                let value_offset = record
                    .offset()
                    .checked_add(header_len)
                    .and_then(|offset| offset.checked_add(8))
                    .ok_or(Error::CapacityOverflow {
                        resource: "BrtCellReal value offset",
                    })?;
                numbers.push(Entry {
                    number: Number { reference, value },
                    original_bits: value.to_bits(),
                    value_offset,
                });
            },
            _ => {},
        }
    }
    if in_sheet_data {
        return Err(Error::InvalidFormat(
            "worksheet stream ended before BrtEndSheetData".to_string(),
        ));
    }
    Ok(Snapshot { source, numbers })
}

fn unique_number(entries: &[Entry], reference: Reference) -> Result<Option<Number>> {
    Ok(unique_index(entries, reference)?.map(|index| entries[index].number))
}

fn unique_index(entries: &[Entry], reference: Reference) -> Result<Option<usize>> {
    let mut found = None;
    for (index, entry) in entries.iter().enumerate() {
        if entry.number.reference != reference {
            continue;
        }
        if found.replace(index).is_some() {
            return Err(Error::UnsupportedFeature(format!(
                "cell ({}, {}) has duplicate BrtCellReal records",
                reference.row(),
                reference.column()
            )));
        }
    }
    Ok(found)
}

#[cfg(test)]
#[allow(
    clippy::expect_used,
    reason = "unit fixtures unwrap only values whose validity is the assertion setup"
)]
mod tests {
    use super::*;
    use crate::raw::{Kind, Writer};

    fn stream() -> Vec<u8> {
        let mut bytes = Vec::new();
        let mut writer = Writer::new(&mut bytes);
        writer
            .write_record(kind::BEGIN_SHEET_DATA, &[])
            .expect("begin");
        let mut row = vec![0; 17];
        row[..4].copy_from_slice(&3_u32.to_le_bytes());
        writer.write_record(kind::ROW_HDR, &row).expect("row");
        writer
            .write_record(Kind::new(0x1234).expect("kind"), &[7, 8, 9])
            .expect("unknown");
        let mut real = vec![0; 16];
        real[..4].copy_from_slice(&2_u32.to_le_bytes());
        real[8..].copy_from_slice(&4.5_f64.to_le_bytes());
        writer.write_record(kind::CELL_REAL, &real).expect("real");
        writer.write_record(kind::END_SHEET_DATA, &[]).expect("end");
        bytes
    }

    #[test]
    fn replaces_only_the_real_value_and_round_trips_patch() {
        let before = stream();
        let snapshot = read(&before).expect("snapshot");
        let reference = Reference::new(3, 2).expect("reference");
        assert_eq!(
            snapshot
                .number(reference)
                .expect("lookup")
                .expect("value")
                .value()
                .to_bits(),
            4.5_f64.to_bits()
        );

        let mut edit = snapshot.edit();
        edit.set_number(reference, 9.25).expect("set");
        let commit = edit.commit().expect("commit");
        let after = commit.patch().apply(&before).expect("apply");
        assert_eq!(
            commit
                .snapshot()
                .number(reference)
                .expect("lookup")
                .expect("value")
                .value()
                .to_bits(),
            9.25_f64.to_bits()
        );
        assert_eq!(
            commit.patch().inverse().apply(&after).expect("revert"),
            before
        );

        let differences = before.iter().zip(&after).filter(|(a, b)| a != b).count();
        assert!(differences > 0 && differences <= 8);
    }

    #[test]
    fn patch_refuses_stale_source() {
        let snapshot = read(&stream()).expect("snapshot");
        let reference = Reference::new(3, 2).expect("reference");
        let mut edit = snapshot.edit();
        edit.set_number(reference, 1.0).expect("set");
        let commit = edit.commit().expect("commit");
        assert!(commit.patch().apply(b"stale").is_err());
    }
}
