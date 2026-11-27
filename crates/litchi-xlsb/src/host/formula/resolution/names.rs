#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! Workbook-defined and external-name resolution.

use crate::package::error::{Error, Result};

use super::Context;
use super::relationships::{SupportingLink, format_formula_prefix};

impl Context {
    pub(super) fn resolve_defined_name(&self, index: u32) -> Result<String> {
        let index = usize::try_from(index)
            .ok()
            .and_then(|index| index.checked_sub(1))
            .ok_or(Error::InvalidFormula(
                "PtgName index is one-based and cannot be zero",
            ))?;
        self.defined_names
            .get(index)
            .cloned()
            .ok_or(Error::UnresolvedDependency(
                "PtgName index exceeds workbook names",
            ))
    }

    pub(super) fn resolve_external_name(&self, xti_index: u16, name_index: u32) -> Result<String> {
        if name_index == 0 {
            return Err(Error::InvalidFormula(
                "PtgNameX name index is one-based and cannot be zero",
            ));
        }
        let xti =
            self.external_sheets
                .get(usize::from(xti_index))
                .ok_or(Error::UnresolvedDependency(
                    "PtgNameX Xti index exceeds entries",
                ))?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow"))?;
        let link = self
            .supporting_links
            .get(link_index)
            .ok_or(Error::UnresolvedDependency(
                "PtgNameX refers to missing supporting link",
            ))?;
        let book_index = match link {
            SupportingLink::ExternalWorkbook(book_index) => book_index,
            SupportingLink::AddIn => {
                return Err(Error::UnsupportedFeature("PtgNameX refers to an add-in"));
            },
            SupportingLink::SelfWorkbook | SupportingLink::SameSheet => {
                return Err(Error::InvalidFormula(
                    "PtgNameX does not refer to an external workbook",
                ));
            },
        };
        let external_book_index = usize::try_from(*book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow"))?;
        let book = self
            .external_books
            .get(external_book_index)
            .ok_or(Error::UnresolvedDependency("missing external book"))?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(
                "PtgNameX refers to a DDE or OLE data source",
            ));
        }
        let index = usize::try_from(name_index - 1)
            .map_err(|_| Error::InvalidFormula("external name index overflow"))?;
        let names = book.metadata.defined_names();
        let name = names.get(index).ok_or(Error::UnresolvedDependency(
            "external name index exceeds names",
        ))?;
        Ok(format!(
            "{}!{}",
            format_formula_prefix(&format!("[{}]", book.metadata.source())),
            name.name()
        ))
    }
}
