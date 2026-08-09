//! Immutable source-bound worksheet page-break state.

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};

use super::PageBreaks;
use crate::error::{Error, Result, invalid};
use crate::{Selector, Workbook, WorksheetKind};

/// Semantic page breaks plus the exact worksheet owner bytes they came from.
#[derive(Clone, Debug)]
pub struct Snapshot {
    value: PageBreaks,
    sheet_name: Box<str>,
    sheet_position: usize,
    worksheet_uri: PackURI,
    worksheet_xml: Arc<Vec<u8>>,
    workbook_uri: PackURI,
    workbook_xml: Arc<Vec<u8>>,
}

impl Snapshot {
    /// Resolve and read one worksheet by its ordinary semantic selector.
    ///
    /// # Errors
    ///
    /// Returns an error when the package or selector is invalid, the selected
    /// sheet is not a worksheet, or its page-break XML is invalid.
    pub fn load<'a>(package: &OpcPackage, selector: impl Into<Selector<'a>>) -> Result<Self> {
        let workbook = Workbook::from_package(package.clone())?;
        let sheet = workbook
            .sheet(selector)?
            .ok_or_else(|| invalid("page-break worksheet selector did not resolve"))?;
        if sheet.kind() != WorksheetKind::Worksheet {
            return Err(Error::NotWorksheet {
                sheet: sheet.name().to_owned(),
            });
        }
        let worksheet = package.get_part(sheet.part_uri())?;
        let workbook_part = package.main_document_part()?;
        Ok(Self {
            value: super::parse(worksheet.blob())?,
            sheet_name: copy_boxed(sheet.name(), "page-break sheet name")?,
            sheet_position: sheet.position(),
            worksheet_uri: worksheet.partname().clone(),
            worksheet_xml: worksheet.blob_arc(),
            workbook_uri: workbook_part.partname().clone(),
            workbook_xml: workbook_part.blob_arc(),
        })
    }

    /// Typed direct page-break state.
    #[must_use]
    pub const fn page_breaks(&self) -> &PageBreaks {
        &self.value
    }

    /// Developer-facing worksheet name captured at ingress.
    #[must_use]
    pub fn sheet_name(&self) -> &str {
        &self.sheet_name
    }

    /// Checked zero-based worksheet position captured at ingress.
    #[must_use]
    pub const fn sheet_position(&self) -> usize {
        self.sheet_position
    }

    /// Resolved worksheet part name.
    #[must_use]
    pub const fn worksheet_part_name(&self) -> &PackURI {
        &self.worksheet_uri
    }

    /// Exact source worksheet XML.
    #[must_use]
    pub fn source_xml(&self) -> &[u8] {
        self.worksheet_xml.as_slice()
    }

    /// Shared exact source worksheet XML.
    #[must_use]
    pub fn source_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.worksheet_xml)
    }

    pub(super) fn same_source(&self, other: &Self) -> bool {
        self.worksheet_uri == other.worksheet_uri
            && self.worksheet_xml == other.worksheet_xml
            && self.workbook_uri == other.workbook_uri
            && self.workbook_xml == other.workbook_xml
    }

    pub(super) fn matches_current_source(&self, package: &OpcPackage) -> bool {
        let Ok(workbook) = package.main_document_part() else {
            return false;
        };
        if workbook.partname() != &self.workbook_uri
            || workbook.blob() != self.workbook_xml.as_slice()
        {
            return false;
        }
        package
            .get_part(&self.worksheet_uri)
            .is_ok_and(|part| part.blob() == self.worksheet_xml.as_slice())
    }
}

fn copy_boxed(value: &str, resource: &'static str) -> Result<Box<str>> {
    let mut copied = String::new();
    copied
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    copied.push_str(value);
    Ok(copied.into_boxed_str())
}
