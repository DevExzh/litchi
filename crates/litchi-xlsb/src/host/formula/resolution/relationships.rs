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
        let Some(SupportingLink::SelfWorkbook) =
            self.supporting_links.get(xti.external_link as usize)
        else {
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
                "structured reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "structured-reference Xti {index} refers to missing supporting link {}",
                    xti.external_link
                ))
            })?;
        match link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.first_sheet != xti.last_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference Xti {index} must select exactly one worksheet"
                    )));
                }
                let sheet = usize::try_from(xti.first_sheet).map_err(|_| {
                    Error::InvalidFormula("table worksheet index overflow".to_string())
                })?;
                if sheet >= self.worksheet_names.len() {
                    return Err(Error::InvalidFormula(format!(
                        "structured-reference worksheet {} exceeds {} worksheets",
                        xti.first_sheet,
                        self.worksheet_names.len()
                    )));
                }
                Ok(sheet)
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet structured-reference Xti {index} must use -2/-2"
                    )));
                }
                self.current_sheet.ok_or_else(|| {
                    Error::InvalidFormula(
                        "same-sheet structured reference has no consuming worksheet".to_string(),
                    )
                })
            },
            SupportingLink::ExternalWorkbook(_) => Err(Error::InvalidFormula(
                "resident structured reference points to an external workbook".to_string(),
            )),
            SupportingLink::AddIn => Err(Error::InvalidFormula(
                "structured reference points to an add-in".to_string(),
            )),
        }
    }

    pub(super) fn resolve_external_table_prefix(&self, index: u16) -> Result<String> {
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti index {index} exceeds {} entries",
                    self.external_sheets.len()
                ))
            })?;
        let link = self
            .supporting_links
            .get(usize::try_from(xti.external_link).map_err(|_| {
                Error::InvalidFormula("table external-link index overflow".to_string())
            })?)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external structured-reference Xti {index} has no supporting link"
                ))
            })?;
        if !matches!(link, SupportingLink::ExternalWorkbook(_)) {
            return Err(Error::InvalidFormula(
                "nonresident structured reference does not point to an external workbook"
                    .to_string(),
            ));
        }
        self.resolve_sheet_prefix(index)
    }

    pub(super) fn resolve_sheet_prefix(&self, index: u16) -> Result<String> {
        if index == u16::MAX {
            return Err(Error::InvalidFormula(
                "3D reference uses invalid Xti index 0xFFFF".to_string(),
            ));
        }
        let xti = self
            .external_sheets
            .get(usize::from(index))
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "Xti index {index} exceeds {} extern-sheet entries",
                    self.external_sheets.len()
                ))
            })?;
        let link_index = usize::try_from(xti.external_link)
            .map_err(|_| Error::InvalidFormula("external-link index overflow".to_string()))?;
        let supporting_link = self.supporting_links.get(link_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {index} refers to missing supporting link {}",
                xti.external_link
            ))
        })?;
        let (first_index, last_index) = match supporting_link {
            SupportingLink::SelfWorkbook => {
                if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
                    return Err(Error::InvalidFormula(format!(
                        "Xti index {index} has invalid self-reference sheet range {}..={}",
                        xti.first_sheet, xti.last_sheet
                    )));
                }
                (
                    usize::try_from(xti.first_sheet).map_err(|_| {
                        Error::InvalidFormula("first sheet index overflow".to_string())
                    })?,
                    usize::try_from(xti.last_sheet).map_err(|_| {
                        Error::InvalidFormula("last sheet index overflow".to_string())
                    })?,
                )
            },
            SupportingLink::SameSheet => {
                if xti.first_sheet != -2 || xti.last_sheet != -2 {
                    return Err(Error::InvalidFormula(format!(
                        "same-sheet Xti index {index} must use workbook scope -2/-2"
                    )));
                }
                let sheet = self.current_sheet.ok_or_else(|| {
                    Error::UnsupportedFeature(
                        "same-sheet reference requires a consuming worksheet".to_string(),
                    )
                })?;
                (sheet, sheet)
            },
            SupportingLink::ExternalWorkbook(book_index) => {
                return self.resolve_external_sheet_prefix(index, xti, *book_index);
            },
            SupportingLink::AddIn => {
                return Err(Error::UnsupportedFeature(format!(
                    "Xti index {index} refers to an add-in"
                )));
            },
        };
        if last_index < first_index {
            return Err(Error::InvalidFormula(format!(
                "Xti index {index} has invalid sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first = self.worksheet_names.get(first_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti first sheet {} exceeds {} worksheets",
                xti.first_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let last = self.worksheet_names.get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti last sheet {} exceeds {} worksheets",
                xti.last_sheet,
                self.worksheet_names.len()
            ))
        })?;
        let unquoted = if first_index == last_index {
            first.clone()
        } else {
            format!("{first}:{last}")
        };
        Ok(format_formula_prefix(&unquoted))
    }

    fn resolve_external_sheet_prefix(
        &self,
        xti_index: u16,
        xti: &ExternalSheet,
        book_index: u32,
    ) -> Result<String> {
        let book_index = usize::try_from(book_index)
            .map_err(|_| Error::InvalidFormula("external book index overflow".to_string()))?;
        let book = self.external_books.get(book_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "Xti index {xti_index} refers to missing external book {book_index}"
            ))
        })?;
        if !book.metadata.is_workbook() {
            return Err(Error::UnsupportedFeature(format!(
                "Xti index {xti_index} refers to a DDE or OLE data source"
            )));
        }
        if xti.first_sheet < 0 || xti.last_sheet < xti.first_sheet {
            return Err(Error::InvalidFormula(format!(
                "Xti index {xti_index} has invalid external sheet range {}..={}",
                xti.first_sheet, xti.last_sheet
            )));
        }
        let first_index = usize::try_from(xti.first_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let last_index = usize::try_from(xti.last_sheet)
            .map_err(|_| Error::InvalidFormula("external sheet index overflow".to_string()))?;
        let first = book
            .metadata
            .sheet_names()
            .get(first_index)
            .ok_or_else(|| {
                Error::InvalidFormula(format!(
                    "external sheet {} exceeds {} cached names",
                    xti.first_sheet,
                    book.metadata.sheet_names().len()
                ))
            })?;
        let last = book.metadata.sheet_names().get(last_index).ok_or_else(|| {
            Error::InvalidFormula(format!(
                "external sheet {} exceeds {} cached names",
                xti.last_sheet,
                book.metadata.sheet_names().len()
            ))
        })?;
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
