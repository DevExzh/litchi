//! Immutable data-validation edits of an XLSB [`Package`] (unified-ops
//! record 0322).
//!
//! A worksheet stores its classic validations in one `BrtBeginDVals`
//! collection: each validation is a `BrtDVal` record, optionally preceded by
//! its `BrtDValList` override. [`Edit`] lists, adds, replaces and removes
//! classic validations one at a time; the collection header's count is
//! updated in place, every other record keeps its bytes, the collection is
//! created with the first validation and removed with the last. Office 2013
//! `BrtDVal14` validations are not listed here and keep their bytes.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_data_validations`] replays a patch.
//!
//! ```
//! use litchi_xlsb::Package;
//! use litchi_xlsb::data_validation::Validation;
//!
//! # fn main() -> Result<(), litchi_xlsb::package::PackageError> {
//! let package = Package::new()?;
//! let mut edit = package.edit_data_validations()?;
//! let mut list = Validation::new(3, "A1:A9".to_owned());
//! list.formula1 = Some("\"a,b\"".to_owned());
//! assert_eq!(edit.add(0, list)?, 0);
//! let after = edit.commit()?.into_package();
//! let validations = after.edit_data_validations()?.validations(0)?;
//! assert_eq!(validations[0].cell_ranges, "A1:A9");
//! # Ok(())
//! # }
//! ```

use std::ops::Range;

use super::{RecordKind, Validation};
use crate::package::Package;
use crate::package::error::Result;
use crate::part_edit::{
    Rank, SheetLayout, Staged, family, invalid, record_with_limits, snapshot_options, splice,
};
use crate::raw::{Records, kind};

family!("data-validation");
snapshot_options!("data-validation");

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Classic validations of one worksheet ([`MAX_VALIDATIONS`]).
    Validations,
}

/// The most classic validations one collection holds.
pub const MAX_VALIDATIONS: usize = 65_534;

/// A staged data-validation edit. Dropping it discards it.
#[derive(Debug, Clone)]
pub struct Edit {
    staged: Staged,
    options: EditOptions,
}

impl Package {
    /// Start a data-validation edit.
    ///
    /// # Errors
    ///
    /// Returns the workbook materialization errors of this package.
    pub fn edit_data_validations(&self) -> Result<Edit> {
        self.edit_data_validations_with_options(EditOptions::default())
    }

    /// Start a data-validation edit with an explicit retained-payload ceiling.
    pub fn edit_data_validations_with_options(&self, options: EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::of_package(self)?,
            options,
        })
    }

    /// Apply a data-validation patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_data_validations(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((patch.0.apply(self)?, patch.0.clone())))
    }
}

/// The record spans of a classic validation collection.
struct Collection {
    unit: Range<usize>,
    header_payload: Vec<u8>,
    items: Vec<Range<usize>>,
    end: Range<usize>,
}

fn collection(
    bytes: &[u8],
    layout: &SheetLayout,
    raw_limits: crate::raw::Limits,
) -> Result<Option<Collection>> {
    let mut units = layout.units_of(kind::BEGIN_D_VALS);
    let Some(unit) = units.next() else {
        return Ok(None);
    };
    if units.next().is_some() {
        return Err(invalid(
            "a worksheet has more than one validation collection",
        ));
    }
    let base = unit.span.start;
    let mut records = Records::new(&bytes[unit.span.clone()], raw_limits);
    let header = records
        .next()
        .transpose()?
        .ok_or_else(|| invalid("an empty validation collection"))?;
    let header_payload = header.payload().to_vec();
    let mut items = Vec::new();
    let mut pending: Option<usize> = None;
    let mut end = None;
    while let Some(item) = records.next() {
        let item = item?;
        let start = base + item.offset();
        let stop = base + records.offset();
        if item.kind() == kind::D_VAL_LIST {
            pending.get_or_insert(start);
        } else if item.kind() == kind::D_VAL {
            items.push(pending.take().unwrap_or(start)..stop);
        } else if item.kind() == kind::END_D_VALS {
            end = Some(start..stop);
        } else {
            return Err(invalid(
                "a validation collection holds an unexpected record",
            ));
        }
    }
    let end = end.ok_or_else(|| invalid("a validation collection has no end"))?;
    if header_payload.len() != 18 {
        return Err(invalid("a validation collection header is not 18 bytes"));
    }
    Ok(Some(Collection {
        unit: unit.span.clone(),
        header_payload,
        items,
        end,
    }))
}

impl Edit {
    /// Discard every staged change.
    pub fn reset(&mut self) {
        self.staged.reset();
    }

    /// The staged classic validations of the worksheet at complete-catalog
    /// position `sheet`, in stream order.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing or non-worksheet tab or a malformed
    /// stream.
    pub fn validations(&self, sheet: usize) -> Result<Vec<Validation>> {
        let worksheet = self.staged.worksheet(sheet)?;
        Ok(classic(
            self.staged.feature_pass(&worksheet, sheet)?.validations,
        ))
    }

    /// Rewrite the collection with `items` (the exact bytes of each kept or
    /// new validation).
    fn publish(&mut self, sheet: usize, change: Change) -> Result<bool> {
        let worksheet = self.staged.worksheet(sheet)?;
        let bytes = self.staged.blob(&worksheet)?;
        let raw_limits = self.staged.raw_limits();
        let layout = SheetLayout::scan_with_limits(bytes, raw_limits)?;
        let current = collection(bytes, &layout, raw_limits)?;
        let mut items: Vec<Vec<u8>> = current
            .as_ref()
            .map(|collection| {
                collection
                    .items
                    .iter()
                    .map(|span| bytes[span.clone()].to_vec())
                    .collect()
            })
            .unwrap_or_default();
        match change {
            Change::Add(encoded) => {
                if items.len() >= MAX_VALIDATIONS {
                    return Err(invalid("the worksheet holds the most validations it can"));
                }
                items.push(encoded);
            },
            Change::Set(index, encoded) => {
                let slot = items
                    .get_mut(index)
                    .ok_or_else(|| invalid("the validation is missing"))?;
                if *slot == encoded {
                    return Ok(false);
                }
                *slot = encoded;
            },
            Change::Remove(index) => {
                if index >= items.len() {
                    return Err(invalid("the validation is missing"));
                }
                items.remove(index);
            },
        }
        let replacement = if items.is_empty() {
            Vec::new()
        } else {
            let count =
                u32::try_from(items.len()).map_err(|_overflow| invalid("too many validations"))?;
            let (header, end) = match &current {
                Some(collection) => {
                    let mut payload = collection.header_payload.clone();
                    payload[14..18].copy_from_slice(&count.to_le_bytes());
                    (
                        record_with_limits(kind::BEGIN_D_VALS, &payload, raw_limits)?,
                        bytes[collection.end.clone()].to_vec(),
                    )
                },
                None => {
                    let mut payload = vec![0u8; 18];
                    payload[14..18].copy_from_slice(&count.to_le_bytes());
                    (
                        record_with_limits(kind::BEGIN_D_VALS, &payload, raw_limits)?,
                        record_with_limits(kind::END_D_VALS, &[], raw_limits)?,
                    )
                },
            };
            // The header keeps every field but its count.
            let mut out = header;
            for item in &items {
                out.extend_from_slice(item);
            }
            out.extend(end);
            out
        };
        let span = match &current {
            Some(collection) => collection.unit.clone(),
            None => {
                let at = layout.insertion(Rank::DataValidations);
                at..at
            },
        };
        let updated = splice(bytes, span, &replacement);
        self.staged.replace_worksheet(&worksheet, updated)
    }

    /// Append a classic validation; returns its position.
    ///
    /// # Errors
    ///
    /// Returns an error for a validation the record cannot encode (an
    /// Office 2013 record kind, an invalid range list or formula), more
    /// than [`MAX_VALIDATIONS`], or a changed edit of a signed workbook.
    pub fn add(&mut self, sheet: usize, validation: Validation) -> Result<usize> {
        let encoded = encode(&validation, self.staged.raw_limits())?;
        let index = self.validations(sheet)?.len();
        self.publish(sheet, Change::Add(encoded))?;
        Ok(index)
    }

    /// Replace the classic validation at `index`. Returns whether anything
    /// changed.
    ///
    /// # Errors
    ///
    /// The errors of [`Self::add`], and an error for a missing position.
    pub fn set(&mut self, sheet: usize, index: usize, validation: Validation) -> Result<bool> {
        let encoded = encode(&validation, self.staged.raw_limits())?;
        self.publish(sheet, Change::Set(index, encoded))
    }

    /// Remove the classic validation at `index`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing position, a malformed stream or a
    /// changed edit of a signed workbook.
    pub fn remove(&mut self, sheet: usize, index: usize) -> Result<()> {
        self.publish(sheet, Change::Remove(index)).map(|_| ())
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit_with_snapshot_limit(
            self.options.max_snapshot_bytes_value(),
        )?))
    }
}

enum Change {
    Add(Vec<u8>),
    Set(usize, Vec<u8>),
    Remove(usize),
}

pub(crate) fn classic(validations: Vec<Validation>) -> Vec<Validation> {
    validations
        .into_iter()
        .filter(|validation| validation.record_kind == RecordKind::Classic)
        .collect()
}

fn encode(validation: &Validation, limits: crate::raw::Limits) -> Result<Vec<u8>> {
    if validation.record_kind != RecordKind::Classic {
        return Err(invalid(
            "only classic validations are authored; BrtDVal14 records are kept as they are",
        ));
    }
    crate::writer::data_validation::encode_classic_records(validation, limits)
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::data_validation::Validation;

    #[test]
    fn add_set_remove_replay() {
        let package = Package::new().unwrap();
        let mut edit = package.edit_data_validations().unwrap();
        let mut whole = Validation::new(1, "B2:B5".to_owned());
        whole.operator = 4;
        whole.formula1 = Some("10".to_owned());
        let mut list = Validation::new(3, "A1:A9".to_owned());
        list.formula1 = Some("\"a,b\"".to_owned());
        assert_eq!(edit.add(0, whole).unwrap(), 0);
        assert_eq!(edit.add(0, list).unwrap(), 1);
        let (after, patch) = edit.commit().unwrap().into_parts();
        let read = after
            .edit_data_validations()
            .unwrap()
            .validations(0)
            .unwrap();
        assert_eq!(read.len(), 2);
        assert_eq!(read[0].cell_ranges, "B2:B5");
        assert_eq!(read[0].formula1.as_deref(), Some("10"));
        assert_eq!(read[1].validation_type, 3);
        assert_eq!(
            package
                .apply_data_validations(&patch)
                .map(|commit| commit.into_package())
                .unwrap()
                .to_bytes()
                .unwrap(),
            after.to_bytes().unwrap()
        );
        let workbook = after.workbook().unwrap();
        let grid = workbook.grid(0usize).unwrap().unwrap();
        assert_eq!(grid.worksheet().data_validations().len(), 2);

        let mut edit = after.edit_data_validations().unwrap();
        let mut changed = read[0].clone();
        changed.error_text = Some("too small".to_owned());
        assert!(edit.set(0, 0, changed).unwrap());
        assert!(!edit.set(0, 1, read[1].clone()).unwrap());
        let set = edit.commit().unwrap().into_package();
        let read = set.edit_data_validations().unwrap().validations(0).unwrap();
        assert_eq!(read[0].error_text.as_deref(), Some("too small"));

        let mut edit = set.edit_data_validations().unwrap();
        edit.remove(0, 0).unwrap();
        edit.remove(0, 0).unwrap();
        let removed = edit.commit().unwrap().into_package();
        assert_eq!(removed.to_bytes().unwrap(), package.to_bytes().unwrap());
    }
}
