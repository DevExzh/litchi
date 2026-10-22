#![allow(
    clippy::map_err_ignore,
    reason = "legacy module confines normalization into the module's stable typed public error to this codec boundary"
)]

//! XTI, worksheet, and external-workbook relationship resolution.

use crate::package::error::{Error, Result};
use crate::package::formula::ExternalSheet;

use super::Context;

/// Kind of supporting link referenced by `BrtExternSheet` entries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SupportingLink {
    SelfWorkbook,
    SameSheet,
    ExternalWorkbook(u32),
    AddIn,
}

impl Context {
    /// Whether an XTI resolves to exactly one worksheet in this workbook.
    pub(crate) fn is_internal_single_sheet_xti(&self, index: u16) -> bool {
        let Some(xti) = self.external_sheets.get(usize::from(index)) else {
            return false;
        };
        let Some(link_index) = usize::try_from(xti.external_link).ok() else {
            return false;
        };
        let Some(SupportingLink::SelfWorkbook) = self.supporting_links.get(link_index) else {
            return false;
        };
        xti.first_sheet >= 0
            && xti.first_sheet == xti.last_sheet
            && usize::try_from(xti.first_sheet)
                .is_ok_and(|sheet| sheet < self.worksheet_names.len())
    }

    pub(super) fn resolve_table_sheet(&self, index: u16) -> Result<usize> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "structured reference uses invalid Xti index 0xFFFF",
            ));
        }
        let xti =
            self.external_sheets
                .get(usize::from(index))
                .ok_or(Error::UnresolvedDependency(
                    "structured-reference Xti index exceeds entries",
                ))?;
        let link = self
            .supporting_links
            .get(
                usize::try_from(xti.external_link)
                    .map_err(|_| Error::InvalidFormula("table external-link index overflow"))?,
            )
            .ok_or({
                Error::UnresolvedDependency(
                    "structured-reference Xti refers to missing supporting link",
                )
            })?;
        match link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                    return Err(Error::InvalidFormula(
                        "structured-reference Xti must select exactly one worksheet",
                    ));
                }
                let sheet = usize::try_from(xti.first_sheet)
                    .map_err(|_| Error::InvalidFormula("table worksheet index overflow"))?;
                if sheet >= self.worksheet_names.len() {
                    return Err(Error::UnresolvedDependency(
                        "structured-reference worksheet exceeds worksheets",
                    ));
                }
                Ok(sheet)
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(
                        "same-sheet structured-reference Xti must use -2/-2",
                    ));
                }
                self.current_sheet.ok_or({
                    Error::UnresolvedDependency(
                        "same-sheet structured reference has no consuming worksheet",
                    )
                })
            },
            SupportingLink::ExternalWorkbook(_) => Err(Error::InvalidFormula(
                "resident structured reference points to an external workbook",
            )),
            SupportingLink::AddIn => Err(Error::UnsupportedFeature(
                "structured reference points to an add-in",
            )),
        }
    }

    pub(super) fn resolve_external_table_prefix(&self, index: u16) -> Result<String> {
        let xti =
            self.external_sheets
                .get(usize::from(index))
                .ok_or(Error::UnresolvedDependency(
                    "external structured-reference Xti index exceeds entries",
                ))?;
        let link = self
            .supporting_links
            .get(
                usize::try_from(xti.external_link)
                    .map_err(|_| Error::InvalidFormula("table external-link index overflow"))?,
            )
            .ok_or({
                Error::UnresolvedDependency(
                    "external structured-reference Xti has no supporting link",
                )
            })?;
        match link {
            SupportingLink::ExternalWorkbook(_) => self.resolve_sheet_prefix(index),
            SupportingLink::AddIn => Err(Error::UnsupportedFeature(
                "nonresident structured reference points to an add-in",
            )),
            SupportingLink::SelfWorkbook | SupportingLink::SameSheet => Err(Error::InvalidFormula(
                "nonresident structured reference does not point to an external workbook",
            )),
        }
    }

    pub(super) fn resolve_sheet_prefix(&self, index: u16) -> Result<String> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "3D reference uses invalid Xti index 0xFFFF",
            ));
        }
        let xti =
            self.external_sheets
                .get(usize::from(index))
                .ok_or(Error::UnresolvedDependency(
                    "Xti index exceeds extern-sheet entries",
                ))?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow"))?;
        let supporting_link =
            self.supporting_links
                .get(link_index)
                .ok_or(Error::UnresolvedDependency(
                    "Xti index refers to missing supporting link",
                ))?;
        let (first_index, last_index) = match supporting_link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
                    return Err(Error::InvalidFormula(
                        "Xti index has invalid self-reference sheet range",
                    ));
                }
                (
                    usize::try_from(xti.first_sheet)
                        .map_err(|_| Error::InvalidFormula("first sheet index overflow"))?,
                    usize::try_from(xti.last_sheet)
                        .map_err(|_| Error::InvalidFormula("last sheet index overflow"))?,
                )
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(
                        "same-sheet Xti index must use workbook scope -2/-2",
                    ));
                }
                let sheet = self.current_sheet.ok_or({
                    Error::UnresolvedDependency(
                        "same-sheet reference requires a consuming worksheet",
                    )
                })?;
                (sheet, sheet)
            },
            SupportingLink::ExternalWorkbook(book_index) => {
                return self.resolve_external_sheet_prefix(index, xti, *book_index);
            },
            SupportingLink::AddIn => {
                return Err(Error::UnsupportedFeature("Xti index refers to an add"));
            },
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula("Xti index has invalid sheet range"));
        }
        let first = self
            .worksheet_names
            .get(first_index)
            .ok_or(Error::UnresolvedDependency(
                "Xti first sheet exceeds worksheets",
            ))?;
        let last = self
            .worksheet_names
            .get(last_index)
            .ok_or(Error::UnresolvedDependency(
                "Xti last sheet exceeds worksheets",
            ))?;
        let unquoted = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&unquoted))
    }

    fn resolve_external_sheet_prefix(
        &self,
        _xti_index: u16,
        xti: &ExternalSheet,
        book_index: u32,
    ) -> Result<String> {
        let book_index = usize::try_from(book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow"))?;
        let book = self
            .external_books
            .get(book_index)
            .ok_or(Error::UnresolvedDependency(
                "Xti index refers to missing external book",
            ))?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(
                "Xti index refers to a DDE or OLE data source",
            ));
        }
        if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
            return Err(Error::InvalidFormula(
                "Xti index has invalid external sheet range",
            ));
        }
        let first_index = usize::try_from(xti.first_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow"))?;
        let last_index = usize::try_from(xti.last_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow"))?;
        let first =
            book.metadata
                .sheet_names()
                .get(first_index)
                .ok_or(Error::UnresolvedDependency(
                    "external sheet exceeds cached names",
                ))?;
        let last =
            book.metadata
                .sheet_names()
                .get(last_index)
                .ok_or(Error::UnresolvedDependency(
                    "external sheet exceeds cached names",
                ))?;
        let sheets = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&format!(
            "[{}]{sheets}",
            book.metadata.source()
        )))
    }
}

/// Quote a workbook/sheet prefix only when Excel's grammar requires it.
pub(super) fn format_formula_prefix(value: &str) -> String {
    if value
        .chars()
        .all(|ch| ch.is_ascii_alphanumeric() || matches!(ch, '_' | '.'))
        && !value.chars().next().is_some_and(|ch| ch.is_ascii_digit())
    {
        value.to_string()
    } else {
        format!("'{}'", value.replace('\'', "''"))
    }
}
