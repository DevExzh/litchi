//! Immutable source-snapshot context used by one workbook transaction.

use crate::error::{Error, Result};
use crate::workbook::{Selector, Workbook, Worksheet, WorksheetKind};

/// Borrowed view of the immutable workbook that a transaction edits.
///
/// Keeping source resolution behind this context makes it explicit that
/// selectors are evaluated against the transaction's source snapshot. Pending
/// additions and mutations never participate in selector lookup.
#[derive(Debug, Clone, Copy)]
pub(super) struct Snapshot<'a> {
    workbook: &'a Workbook,
}

impl<'a> Snapshot<'a> {
    pub(super) const fn new(workbook: &'a Workbook) -> Self {
        Self { workbook }
    }

    pub(super) fn tab<'s>(&self, selector: impl Into<Selector<'s>>) -> Result<Option<Worksheet>> {
        self.workbook.sheet(selector)
    }

    pub(super) fn worksheet<'s>(
        &self,
        selector: impl Into<Selector<'s>>,
    ) -> Result<Option<Worksheet>> {
        let Some(sheet) = self.tab(selector)? else {
            return Ok(None);
        };
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        Ok(Some(sheet))
    }

    pub(super) fn pair<'left, 'right>(
        &self,
        left: impl Into<Selector<'left>>,
        right: impl Into<Selector<'right>>,
    ) -> Result<Option<(Worksheet, Worksheet)>> {
        let Some(left) = self.tab(left)? else {
            return Ok(None);
        };
        let Some(right) = self.tab(right)? else {
            return Ok(None);
        };
        Ok(Some((left, right)))
    }
}
