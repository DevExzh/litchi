//! Worksheet tables on the source-backed owner.

use litchi_opc::constants::content_type as ct;

use super::{SourceBackedWorkbook, SourceWorksheet};
use crate::error::{Error, Result, invalid};
use crate::workbook::WorksheetKind;
use crate::workbook::worksheet::{TablePart, collect_tables, find_table, find_workbook_table};

impl SourceWorksheet {
    /// Every table owned by this worksheet, in the order
    /// [`crate::Worksheet::tables`] returns them for the same package.
    ///
    /// Only the worksheet's relationships, which were validated at open, and
    /// its table parts are read; the worksheet payload is not.
    ///
    /// # Errors
    ///
    /// Returns an error for a non-cell sheet, an external or mistyped table
    /// relationship, a malformed table part, or a changed source or cancelled
    /// read.
    pub fn tables(&self) -> Result<Vec<TablePart>> {
        let tables = self.read_tables();
        self.finish_result(tables)
    }

    /// The table whose display name, or failing that whose name, equals
    /// `name` without regard to case.
    ///
    /// Display names are unique in a workbook, and table names within a
    /// sheet, both compared without regard to case, so a well-formed
    /// package has at most one match.
    ///
    /// # Errors
    ///
    /// As for [`Self::tables`], or [`Error::Ambiguous`] when multiple names match.
    pub fn table(&self, name: &str) -> Result<Option<TablePart>> {
        find_table(self.tables()?, name)
    }

    fn read_tables(&self) -> Result<Vec<TablePart>> {
        self.owner.execution_check()?;
        if self.data.kind != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: self.data.name().to_owned(),
            });
        }
        let package = &self.owner.package;
        let limits = package.read_limits();
        let worksheet = package.part(&self.data.part_uri)?;
        collect_tables(worksheet.rels(), |part_name| {
            // A table is published in the result only after its complete
            // source read and parse. Fence each closure independently so a
            // cancellation or source replacement cannot turn one stale table
            // into an apparently successful collection.
            self.owner.execution_check()?;
            package.source_version()?;
            let part = package.part(part_name)?;
            if part.content_type() != ct::SML_TABLE {
                return Err(invalid(format!(
                    "worksheet table part '{part_name}' has content type '{}', expected '{}'",
                    part.content_type(),
                    ct::SML_TABLE
                )));
            }
            let data = part.data()?;
            crate::source_limits::validate_xml_limits(data.as_bytes(), limits)?;
            let table = crate::table::parse_table_xml(data.as_bytes())?;
            self.owner.execution_check()?;
            package.source_version()?;
            Ok(table)
        })
    }
}

impl SourceBackedWorkbook {
    /// The worksheet table whose display name, or failing that whose name,
    /// equals `name` without regard to case, with the worksheet that owns
    /// it.
    ///
    /// Every worksheet is checked for ambiguous matches. A name match is
    /// used only when no worksheet has a display-name match. Worksheets are
    /// searched in workbook order and non-cell sheets, which own no tables,
    /// are skipped. Each table part is read at most once.
    ///
    /// # Errors
    ///
    /// As for [`SourceWorksheet::tables`] on each worksheet searched, or
    /// [`Error::Ambiguous`] when multiple names match.
    pub fn table(&self, name: &str) -> Result<Option<(SourceWorksheet, TablePart)>> {
        let this = self.view()?;
        this.inner.execution_check()?;
        this.inner.package.source_version()?;
        let result = find_workbook_table(
            this.sheets()?,
            |sheet| {
                if sheet.kind() == WorksheetKind::Worksheet {
                    sheet.tables().map(Some)
                } else {
                    Ok(None)
                }
            },
            name,
        );
        this.inner.execution_check()?;
        this.inner.package.source_version()?;
        result
    }
}
