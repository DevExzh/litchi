//! The workbook family view, ported from the facade's `sheet::Workbook`
//! reads, with the generic [`WorkbookTrait`] the facade's boxed openers
//! provided.

use litchi_core::FileFormat;
use litchi_core::sheet::WorkbookTrait;

use super::sheet::{Worksheet, Worksheets};

use crate::Result;
#[cfg(feature = "ods")]
use crate::format::ods::Ods;
#[cfg(feature = "xls")]
use crate::format::xls::Xls;
#[cfg(feature = "xlsb")]
use crate::format::xlsb::Xlsb;
#[cfg(feature = "xlsx")]
use crate::format::xlsx::Xlsx;

/// The workbook view of a [`File`](crate::File): the worksheets of XLS,
/// XLSX, XLSB and ODS workbooks.
///
/// [`File::workbook`](crate::File::workbook) returns it. Its inherent reads
/// keep each format's behavior, including sheet-name matching: XLSX matches
/// names case-insensitively, the other formats exactly (record 0003 F15).
///
/// The view also implements [`WorkbookTrait`] by static dispatch to the
/// format handle, so a formula evaluator generic over the trait reads it
/// directly. ODS has no sheet adapter yet: through the trait, an ODS view
/// reports its worksheet names and count, and refuses worksheet access with
/// [`Code::UnsupportedCapability`](crate::Code::UnsupportedCapability).
///
/// ```no_run
/// use litchi_core::sheet::WorkbookTrait;
///
/// # fn main() -> litchi_ops::Result<()> {
/// let file = litchi_ops::open("budget.xlsx")?;
/// let workbook = file.workbook()?;
/// println!("{:?}", workbook.worksheet_names()?);
/// // The same view through the generic sheet trait.
/// let first = WorkbookTrait::worksheet_by_index(&workbook, 0);
/// # let _ = first;
/// # Ok(())
/// # }
/// ```
#[derive(Clone, Copy)]
pub struct WorkbookView<'a>(Inner<'a>);

#[derive(Clone, Copy)]
enum Inner<'a> {
    #[cfg(feature = "xlsx")]
    Xlsx(&'a Xlsx),
    #[cfg(feature = "xlsb")]
    Xlsb(&'a Xlsb),
    #[cfg(feature = "xls")]
    Xls(&'a Xls),
    #[cfg(feature = "ods")]
    Ods(&'a Ods),
}

impl<'a> WorkbookView<'a> {
    #[cfg(feature = "xlsx")]
    pub(crate) fn xlsx(handle: &'a Xlsx) -> Self {
        Self(Inner::Xlsx(handle))
    }

    #[cfg(feature = "xlsb")]
    pub(crate) fn xlsb(handle: &'a Xlsb) -> Self {
        Self(Inner::Xlsb(handle))
    }

    #[cfg(feature = "xls")]
    pub(crate) fn xls(handle: &'a Xls) -> Self {
        Self(Inner::Xls(handle))
    }

    #[cfg(feature = "ods")]
    pub(crate) fn ods(handle: &'a Ods) -> Self {
        Self(Inner::Ods(handle))
    }

    /// The format of the workbook.
    #[must_use]
    pub fn format(&self) -> FileFormat {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(_) => FileFormat::Xlsx,
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(_) => FileFormat::Xlsb,
            #[cfg(feature = "xls")]
            Inner::Xls(_) => FileFormat::Xls,
            #[cfg(feature = "ods")]
            Inner::Ods(_) => FileFormat::Ods,
        }
    }

    /// The worksheet names, in workbook order. XLSX lists every sheet kind;
    /// XLSB and XLS list worksheets only, as their readers do.
    ///
    /// # Errors
    ///
    /// Returns [`Code::IoError`](crate::Code::IoError) when a source-backed
    /// workbook changed after open, and the format's typed error when its
    /// catalog cannot be read.
    pub fn worksheet_names(&self) -> Result<Vec<String>> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(xlsx) => xlsx.worksheet_names(),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(xlsb) => xlsb.worksheet_names(),
            #[cfg(feature = "xls")]
            Inner::Xls(xls) => xls.worksheet_names(),
            #[cfg(feature = "ods")]
            Inner::Ods(ods) => Ok(ods
                .worksheets()?
                .iter()
                .map(|sheet| sheet.name.clone())
                .collect()),
        }
    }

    /// The number of worksheets, counted as [`Self::worksheet_names`] lists
    /// them.
    ///
    /// # Errors
    ///
    /// As [`Self::worksheet_names`].
    pub fn worksheet_count(&self) -> Result<usize> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(xlsx) => xlsx.worksheet_count(),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(xlsb) => xlsb.worksheet_count(),
            #[cfg(feature = "xls")]
            Inner::Xls(xls) => xls.worksheet_count(),
            #[cfg(feature = "ods")]
            Inner::Ods(ods) => Ok(ods.worksheets()?.len()),
        }
    }
}

/// The refusal of access to `what` on an ODS view through the trait.
#[cfg(feature = "ods")]
fn ods_unsupported() -> crate::Error {
    crate::Error::new(
        crate::Code::UnsupportedCapability,
        "ODS worksheets are not available through the generic sheet traits",
    )
    .with_format(FileFormat::Ods)
    .with_suggestion("read ODS worksheets through `WorkbookView::worksheet_names` and `File::text`")
}

/// Static dispatch to the format handle's [`WorkbookTrait`]: XLSX and XLSB
/// as their handles implement it, XLS through its eager reader (read once
/// from the file on first use when the workbook was opened from a path), and
/// ODS refused as described on [`WorkbookView`].
impl WorkbookTrait for WorkbookView<'_> {
    type Error = crate::Error;
    type Worksheet<'b>
        = Worksheet<'b>
    where
        Self: 'b;
    type Worksheets<'b>
        = Worksheets<'b>
    where
        Self: 'b;

    fn active_sheet_index(&self) -> usize {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::active_sheet_index(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::active_sheet_index(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::active_sheet_index(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => 0,
        }
    }

    fn active_worksheet(&self) -> Result<Worksheet<'_>> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::active_worksheet(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::active_worksheet(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::active_worksheet(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => Err(ods_unsupported()),
        }
    }

    fn is_1904_date_system(&self) -> bool {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::is_1904_date_system(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::is_1904_date_system(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::is_1904_date_system(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => false,
        }
    }

    fn worksheet_by_index(&self, index: usize) -> Result<Worksheet<'_>> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::worksheet_by_index(handle, index),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::worksheet_by_index(handle, index),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::worksheet_by_index(handle, index),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => {
                let _ = index;
                Err(ods_unsupported())
            },
        }
    }

    fn worksheet_by_name(&self, name: &str) -> Result<Worksheet<'_>> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::worksheet_by_name(handle, name),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::worksheet_by_name(handle, name),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::worksheet_by_name(handle, name),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => {
                let _ = name;
                Err(ods_unsupported())
            },
        }
    }

    fn worksheet_count(&self) -> usize {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::worksheet_count(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::worksheet_count(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::worksheet_count(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(handle) => handle.trait_names().len(),
        }
    }

    fn worksheet_names(&self) -> &[String] {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::worksheet_names(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::worksheet_names(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::worksheet_names(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(handle) => handle.trait_names(),
        }
    }

    fn worksheets(&self) -> Worksheets<'_> {
        match self.0 {
            #[cfg(feature = "xlsx")]
            Inner::Xlsx(handle) => WorkbookTrait::worksheets(handle),
            #[cfg(feature = "xlsb")]
            Inner::Xlsb(handle) => WorkbookTrait::worksheets(handle),
            #[cfg(feature = "xls")]
            Inner::Xls(handle) => WorkbookTrait::worksheets(handle),
            #[cfg(feature = "ods")]
            Inner::Ods(_) => Worksheets::new(super::sheet::WorksheetsInner::Refused(Some(
                ods_unsupported(),
            ))),
        }
    }
}

impl std::fmt::Debug for WorkbookView<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("WorkbookView")
            .field("format", &self.format())
            .finish_non_exhaustive()
    }
}
