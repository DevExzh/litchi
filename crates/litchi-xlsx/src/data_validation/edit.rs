//! Immutable data-validation edits of a [`Package`] (unified-ops record
//! 0320).
//!
//! [`Package::edit_data_validations`] starts an [`Edit`] over an immutable
//! snapshot. [`Edit::validations`] reads a worksheet's complete core and
//! Office 2010 collections and [`Edit::set_validations`] replaces them; the
//! worksheet bytes outside the validation owners are copied exactly.
//! [`Edit::commit`] returns a [`Commit`] with the new package, the
//! reversible [`Patch`] and the [`Diagnostics`];
//! [`Package::apply_data_validations`] replays a patch onto an equal
//! package. An exact no-op keeps every byte, and a changed edit of a signed
//! workbook is refused with [`Error::Signed`].
//!
//! ```
//! use litchi_xlsx::Package;
//! use litchi_xlsx::data_validation::{
//!     Collection, Formula, ListSource, Source, Sqref, Validation, ValidationType,
//! };
//!
//! # fn main() -> Result<(), litchi_xlsx::Error> {
//! let package = Package::new()?;
//! let mut edit = package.edit_data_validations()?;
//! let mut validation = Validation::new(Source::Core, ValidationType::List, Sqref::parse("A2:A20")?);
//! validation.set_formula1(Some(ListSource::Formula(Formula::new("\"Draft,Review\"")?)))?;
//! edit.set_validations(0, vec![Collection::new(Source::Core, vec![validation])?])?;
//! let after = edit.commit()?.into_package();
//! let sheet = after.workbook().sheet(0)?.unwrap();
//! assert_eq!(sheet.data_validations()?[0].validations().len(), 1);
//! # Ok(())
//! # }
//! ```

use super::Collection;
use crate::Package;
use crate::error::Result;
use crate::part_edit::{Staged, family};

family!("data-validation");

/// The most validations one worksheet may carry after an edit (Excel's
/// limit for one `dataValidations` collection).
pub const MAX_VALIDATIONS: usize = 65_534;

/// A bounded dimension of this family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Validations of one worksheet ([`MAX_VALIDATIONS`]).
    Validations,
}

/// A staged data-validation edit of one [`Package`]. Dropping it discards
/// it.
#[derive(Debug)]
pub struct Edit {
    staged: Staged,
}

impl Package {
    /// Start a data-validation edit.
    ///
    /// # Errors
    ///
    /// Returns the encryption-policy error for a package whose mutations are
    /// refused.
    pub fn edit_data_validations(&self) -> Result<Edit> {
        self.edit_data_validations_with_options(crate::EditOptions::default())
    }

    /// Start a data-validation edit with an explicit retained-patch budget.
    pub fn edit_data_validations_with_options(&self, options: crate::EditOptions) -> Result<Edit> {
        Ok(Edit {
            staged: Staged::new_with_options(self, "edit_data_validations", options)?,
        })
    }

    /// Apply a data-validation patch to a new package; this one is
    /// unchanged. An empty patch returns this package.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the patch's
    /// before state and [`Error::Signed`] for a changed patch of a signed
    /// workbook.
    pub fn apply_data_validations(&self, patch: &Patch) -> Result<Commit> {
        Ok(Commit::new((
            patch.0.apply(self, "apply_data_validations")?,
            patch.0.clone(),
        )))
    }
}

impl Edit {
    /// The staged validation collections of the worksheet at `sheet`.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, or malformed
    /// markup.
    pub fn validations(&self, sheet: usize) -> Result<Vec<Collection>> {
        let worksheet = self.staged.worksheet(sheet)?;
        super::parse_data_validation_collections(self.staged.blob(&worksheet)?)
    }

    /// Replace the validation collections of the worksheet at `sheet`.
    /// Returns whether anything changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing sheet or a chartsheet, invalid
    /// collections, too many validations, or a changed edit of a signed
    /// workbook.
    pub fn set_validations(&mut self, sheet: usize, values: Vec<Collection>) -> Result<bool> {
        let count: usize = values.iter().map(|value| value.validations().len()).sum();
        if count > MAX_VALIDATIONS {
            return Err(Error::limit(LimitKind::Validations, count, MAX_VALIDATIONS));
        }
        let worksheet = self.staged.worksheet(sheet)?;
        let xml =
            super::replace_data_validation_collections(self.staged.blob(&worksheet)?, &values)?;
        // A prefixed root (`x:worksheet`) keeps the owner in its prefix, so
        // every worksheet writer still reads it as a direct child.
        let xml = crate::part_edit::WorksheetLayout::scan(&xml)?
            .unprefixed_in_root_prefix(&xml, "dataValidations");
        self.staged.replace_blob(&worksheet, xml)
    }

    /// Validate the staged package and publish it.
    ///
    /// # Errors
    ///
    /// Returns the workbook validation errors of the staged package.
    pub fn commit(self) -> Result<Commit> {
        Ok(Commit::new(self.staged.commit()?))
    }
}

#[cfg(test)]
mod tests {
    use crate::Package;
    use crate::data_validation::{
        Collection, Formula, ListSource, Source, Sqref, Validation, ValidationOperator,
        ValidationType,
    };

    #[test]
    fn set_replay_inverse_and_no_op() {
        let package = Package::new().unwrap();
        let mut validation = Validation::new(
            Source::Core,
            ValidationType::Whole,
            Sqref::parse("A2:A50").unwrap(),
        );
        validation.set_operator(ValidationOperator::Between);
        validation
            .set_formula1(Some(ListSource::Formula(Formula::new("1").unwrap())))
            .unwrap();
        validation
            .set_formula2(Some(Formula::new("100").unwrap()))
            .unwrap();
        let values = vec![Collection::new(Source::Core, vec![validation]).unwrap()];
        let mut edit = package.edit_data_validations().unwrap();
        assert!(edit.set_validations(0, values.clone()).unwrap());
        let (after, patch) = edit.commit().unwrap().into_parts();
        assert_eq!(
            after
                .edit_data_validations()
                .unwrap()
                .validations(0)
                .unwrap(),
            values
        );
        assert!(
            package
                .apply_data_validations(&patch)
                .unwrap()
                .package()
                .same_content(&after)
                .unwrap()
        );
        assert!(
            after
                .apply_data_validations(&patch.inverse())
                .unwrap()
                .package()
                .same_content(&package)
                .unwrap()
        );
        let noop = package.edit_data_validations().unwrap().commit().unwrap();
        assert!(noop.patch().is_noop());
    }

    /// Record 0320 gap 7: on a worksheet whose root carries a prefix, the
    /// validation owner is written in that prefix, so the hyperlink writer
    /// still recognizes every direct child.
    #[test]
    fn a_prefixed_root_keeps_the_worksheet_writable() {
        let package = Package::new().unwrap();
        let mut opc = package.opc().clone();
        let uri = litchi_opc::PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        opc.get_part_mut(&uri).unwrap().set_blob(
            br#"<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:sheetData/></x:worksheet>"#
                .to_vec(),
        );
        let package = Package::from_opc(opc).unwrap();
        let mut validation = Validation::new(
            Source::Core,
            ValidationType::List,
            Sqref::parse("A1:A5").unwrap(),
        );
        validation
            .set_formula1(Some(ListSource::Formula(Formula::new("\"a,b\"").unwrap())))
            .unwrap();
        let values = vec![Collection::new(Source::Core, vec![validation]).unwrap()];
        let mut edit = package.edit_data_validations().unwrap();
        assert!(edit.set_validations(0, values.clone()).unwrap());
        let package = edit.commit().unwrap().into_package();
        let xml = String::from_utf8(package.opc().get_part(&uri).unwrap().blob().to_vec()).unwrap();
        assert!(xml.contains("<x:dataValidations"), "{xml}");
        assert_eq!(
            package
                .edit_data_validations()
                .unwrap()
                .validations(0)
                .unwrap(),
            values
        );
        let link = crate::hyperlinks::Hyperlink::new(
            crate::hyperlinks::HyperlinkReference::new("B2").unwrap(),
            None,
            Some("https://example.test/".to_owned()),
            None,
            None,
        )
        .unwrap();
        let mut edit = package.edit().unwrap();
        edit.put_hyperlink(0, link).unwrap();
        let package = edit.commit().unwrap().into_package();
        assert_eq!(
            package
                .workbook()
                .sheet(0)
                .unwrap()
                .unwrap()
                .hyperlinks()
                .unwrap()
                .len(),
            1
        );
    }
}
