//! Immutable, source-bound worksheet page-break edits of a
//! [`crate::Package`] (record 0360).
//!
//! [`crate::Package::edit_page_breaks`] starts an [`Edit`] of one
//! worksheet's row and column breaks; [`Edit::commit`] returns a [`Commit`]
//! holding the new package, the exact [`Patch`] and the [`Diagnostics`];
//! [`crate::Package::apply_page_breaks`] replays a patch onto an equal
//! package. The source package never changes; an exact no-op returns it, a
//! signed one included, and a changed edit of a signed workbook is refused
//! with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::page_breaks::{Break, Collection};
//!
//! # fn main() -> litchi_xlsx::Result<()> {
//! let package = Package::new()?;
//! let mut edit = package.edit_page_breaks("Sheet1")?;
//! edit.set_horizontal(Collection::horizontal([Break::new(7, 0, 0)?])?)?;
//! let (after, patch) = edit.commit()?.into_parts();
//! assert!(after.page_breaks("Sheet1")?.page_breaks().horizontal().is_some());
//! let restored = after.apply_page_breaks(&patch.inverse())?.into_package();
//! assert!(restored.page_breaks("Sheet1")?.page_breaks().horizontal().is_none());
//! # Ok(())
//! # }
//! ```

use std::borrow::Cow;

use litchi_opc::OpcPackage;

use super::{Collection, PageBreaks, Snapshot};
use crate::Selector;
use crate::error::{Result, invalid};

pub use super::LimitKind;

crate::record::record_family!("page-break");

/// The exact, reversible change of a page-break edit: the worksheet part's
/// before and after states.
#[derive(Clone, Debug)]
pub struct Patch(pub(crate) super::Patch);

impl Patch {
    /// Whether the patch changes nothing.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.0.is_empty()
    }

    /// The reverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self(self.0.inverse())
    }

    /// The worksheet state the patch requires before application.
    #[must_use]
    pub const fn before(&self) -> &Snapshot {
        self.0.before()
    }

    /// The worksheet state the patch produces.
    #[must_use]
    pub const fn after(&self) -> &Snapshot {
        self.0.after()
    }

    /// The semantic payload retained by this reversible patch.
    #[must_use]
    pub fn retained_bytes(&self) -> usize {
        let before = self
            .0
            .before()
            .source_arc()
            .map_or(0, |bytes| bytes.capacity());
        let after = self
            .0
            .after()
            .source_arc()
            .map_or(0, |bytes| bytes.capacity());
        before.saturating_add(after)
    }
}

impl crate::record::Record for Patch {
    fn is_noop(&self) -> bool {
        Self::is_noop(self)
    }

    fn apply_to(&self, package: &mut OpcPackage) -> Result<()> {
        self.0.apply(package)
    }
}

/// A staged page-break edit of one worksheet of a [`crate::Package`].
/// Dropping it discards it.
pub struct Edit {
    target: crate::record::Source,
    selector: Selector<'static>,
    before: Option<Snapshot>,
    staged: Option<PageBreaks>,
    max_retained_bytes: Option<usize>,
}

impl crate::Package {
    /// Start a page-break edit of one worksheet.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused, and an error when the selector does not select a worksheet
    /// or its page-break XML is invalid.
    pub fn edit_page_breaks<'a>(&self, selector: impl Into<Selector<'a>>) -> Result<Edit> {
        self.edit_page_breaks_with_options(selector, crate::EditOptions::default())
    }

    /// Start a page-break edit with an explicit retained-patch budget.
    pub fn edit_page_breaks_with_options<'a>(
        &self,
        selector: impl Into<Selector<'a>>,
        options: crate::EditOptions,
    ) -> Result<Edit> {
        let target = crate::record::Source::new(self, "edit_page_breaks")?;
        Ok(Edit {
            target,
            selector: owned_selector(selector.into())?,
            before: None,
            staged: None,
            max_retained_bytes: options.max_retained_bytes_value(),
        })
    }

    /// Apply an exact page-break patch to a new package; this one is
    /// unchanged.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_page_breaks(&self, patch: &Patch) -> Result<Commit> {
        crate::record::replay(self, patch, "apply_page_breaks")
    }
}

impl Edit {
    /// Materialize the selected worksheet only when the edit first needs its
    /// semantic state. Selector ownership is bounded to the caller's input;
    /// failed reads leave no partial staged value.
    fn ensure_prepared(&mut self) -> Result<()> {
        if self.before.is_some() {
            return Ok(());
        }
        let before = Snapshot::load(&self.target, self.selector.clone())?;
        let staged = before.page_breaks().clone();
        self.staged = Some(staged);
        self.before = Some(before);
        Ok(())
    }

    fn snapshot(&self) -> Result<&Snapshot> {
        self.before
            .as_ref()
            .ok_or_else(|| invalid("page-break edit is not prepared"))
    }

    fn staged(&self) -> Result<&PageBreaks> {
        self.staged
            .as_ref()
            .ok_or_else(|| invalid("page-break edit is not prepared"))
    }

    /// Exact source state captured at transaction start.
    #[must_use = "use the source snapshot or handle the error"]
    pub fn before(&mut self) -> Result<&Snapshot> {
        self.ensure_prepared()?;
        self.snapshot()
    }

    /// Currently staged page-break state.
    #[must_use = "use the staged page breaks or handle the error"]
    pub fn page_breaks(&mut self) -> Result<&PageBreaks> {
        self.ensure_prepared()?;
        self.staged()
    }

    /// Replace the complete staged page-break state.
    ///
    /// # Errors
    ///
    /// Returns an error when `value` violates page-break invariants.
    pub fn set(&mut self, value: PageBreaks) -> Result<bool> {
        self.ensure_prepared()?;
        value.validate()?;
        if self.staged()? == &value {
            return Ok(false);
        }
        self.staged = Some(value);
        Ok(true)
    }

    /// Clone-edit the complete staged value; a failed closure changes nothing.
    ///
    /// # Errors
    ///
    /// Returns the closure's error or a page-break validation error.
    pub fn edit(&mut self, edit: impl FnOnce(&mut PageBreaks) -> Result<()>) -> Result<bool> {
        self.ensure_prepared()?;
        let mut draft = self.staged()?.clone();
        edit(&mut draft)?;
        self.set(draft)
    }

    /// Replace horizontal row breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not horizontal or is invalid.
    pub fn set_horizontal(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_horizontal(collection))
    }

    /// Replace vertical column breaks.
    ///
    /// # Errors
    ///
    /// Returns an error when `collection` is not vertical or is invalid.
    pub fn set_vertical(&mut self, collection: Collection) -> Result<bool> {
        self.edit(|value| value.set_vertical(collection))
    }

    /// Remove horizontal row breaks.
    pub fn remove_horizontal(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        Ok(self
            .staged
            .as_mut()
            .ok_or_else(|| invalid("page-break edit is not prepared"))?
            .remove_horizontal())
    }

    /// Remove vertical column breaks.
    pub fn remove_vertical(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        Ok(self
            .staged
            .as_mut()
            .ok_or_else(|| invalid("page-break edit is not prepared"))?
            .remove_vertical())
    }

    /// Remove both authored collections.
    pub fn clear(&mut self) -> Result<bool> {
        self.ensure_prepared()?;
        Ok(self
            .staged
            .as_mut()
            .ok_or_else(|| invalid("page-break edit is not prepared"))?
            .clear())
    }

    /// Discard every staged change.
    pub fn reset(&mut self) -> Result<()> {
        self.ensure_prepared()?;
        let staged = self.snapshot()?.page_breaks().clone();
        self.staged = Some(staged);
        Ok(())
    }

    /// Validate, rewrite the worksheet, read it back and publish it as a new
    /// package. An unchanged edit returns the source package.
    ///
    /// # Errors
    ///
    /// Returns an error for stale source bytes, signed packages, invalid
    /// staged state, rewrite failures, or semantic readback failures.
    pub fn commit(mut self) -> Result<Commit> {
        self.ensure_prepared()?;
        let before = self
            .before
            .take()
            .ok_or_else(|| invalid("page-break edit is not prepared"))?;
        let staged = self
            .staged
            .take()
            .ok_or_else(|| invalid("page-break edit is not prepared"))?;
        if before.page_breaks() == &staged {
            let patch = Patch(super::Patch::new(before.clone(), before));
            if let Some(maximum) = self.max_retained_bytes {
                let endpoint_bytes = self.target.retained_bytes();
                let observed = endpoint_bytes
                    .saturating_add(endpoint_bytes)
                    .saturating_add(patch.retained_bytes());
                if observed > maximum {
                    return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                        resource: litchi_core::Resource::Memory,
                        observed: u64::try_from(observed).unwrap_or(u64::MAX),
                        limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                        scope: "XLSX page-break retained package".into(),
                    }));
                }
            }
            return Ok(Commit::new(
                self.target.into_package(),
                patch,
                Diagnostics::new(false),
            ));
        }
        if self.target.is_signed() {
            return Err(Error::Signed);
        }
        if !before.matches_current_source(&self.target) {
            return Err(Error::PatchConflict {
                part: before.worksheet_part_name().to_string(),
            });
        }
        let output = super::replace(before.source_xml(), &staged)?;
        let estimated_patch_bytes = before
            .source_arc()
            .map(|before| before.capacity())
            .unwrap_or(0)
            .saturating_add(output.capacity());
        let mut candidate = OpcPackage::clone(&self.target);
        candidate
            .get_part_mut(before.worksheet_part_name())?
            .set_blob(output);
        let snapshot = Snapshot::load(&candidate, before.sheet_position())?;
        if snapshot.page_breaks() != &staged {
            return Err(invalid(
                "page-break publication changed the staged semantic state",
            ));
        }
        let package = self.target.publish(candidate)?;
        let endpoint_charge = self
            .max_retained_bytes
            .map(|_| (self.target.retained_bytes(), package.retained_bytes()));
        if let Some(maximum) = self.max_retained_bytes {
            let (source_bytes, target_bytes) =
                endpoint_charge.expect("retained endpoint charge established with the budget");
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(estimated_patch_bytes);
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX page-break retained package".into(),
                }));
            }
        }
        let patch = Patch(super::Patch::new(before, snapshot));
        if let Some(maximum) = self.max_retained_bytes {
            let (source_bytes, target_bytes) =
                endpoint_charge.expect("retained endpoint charge established with the budget");
            let observed = source_bytes
                .saturating_add(target_bytes)
                .saturating_add(patch.retained_bytes());
            if observed > maximum {
                return Err(Error::ResourceLimit(litchi_core::ResourceLimit {
                    resource: litchi_core::Resource::Memory,
                    observed: u64::try_from(observed).unwrap_or(u64::MAX),
                    limit: u64::try_from(maximum).unwrap_or(u64::MAX),
                    scope: "XLSX page-break retained package".into(),
                }));
            }
        }
        Ok(Commit::new(package, patch, Diagnostics::new(true)))
    }
}

fn owned_selector(selector: Selector<'_>) -> Result<Selector<'static>> {
    match selector {
        Selector::Name(name) => Ok(Selector::Name(Cow::Owned(name.into_owned()))),
        Selector::Position(position) => Ok(Selector::Position(position)),
        Selector::Id(id) => match id {},
        _ => Err(Error::UnsupportedSelector),
    }
}

#[cfg(test)]
mod lazy_tests {
    use litchi_opc::PackURI;

    #[test]
    fn edit_entry_defers_page_break_xml() {
        let package = crate::Package::new().expect("minimal XLSX");
        let mut edit = package
            .edit_page_breaks(0usize)
            .expect("worksheet selector");
        assert!(edit.before.is_none());
        assert!(edit.staged.is_none());
        assert!(edit.page_breaks().is_ok());
        assert!(edit.before.is_some());
    }

    #[test]
    fn failed_page_break_read_is_retryable_without_partial_state() {
        let source = crate::Package::new().expect("minimal XLSX");
        let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").expect("worksheet URI");
        let mut package = source.opc().clone();
        package
            .get_part_mut(&worksheet)
            .expect("worksheet")
            .set_blob(
                br#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:rowBreaks count="2"><x:brk/></x:rowBreaks></x:worksheet>"#.to_vec(),
            );
        let package = crate::Package::from_opc(package).expect("defer worksheet XML");
        let mut edit = package
            .edit_page_breaks(0usize)
            .expect("worksheet selector");
        assert!(edit.page_breaks().is_err());
        assert!(edit.before.is_none());
        assert!(edit.page_breaks().is_err());
        assert!(edit.before.is_none());
    }
}
