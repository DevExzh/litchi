//! The flat [`File`] enum and its format-neutral read surface (design
//! section 4; record 0003 §2.1, §2.2 and §2.4).

#[cfg(feature = "doc")]
use crate::format::doc::Doc;
#[cfg(feature = "docx")]
use crate::format::docx::Docx;
#[cfg(feature = "odp")]
use crate::format::odp::Odp;
#[cfg(feature = "ods")]
use crate::format::ods::Ods;
#[cfg(feature = "odt")]
use crate::format::odt::Odt;
#[cfg(feature = "ppt")]
use crate::format::ppt::Ppt;
#[cfg(feature = "pptx")]
use crate::format::pptx::Pptx;
#[cfg(feature = "rtf")]
use crate::format::rtf::Rtf;
#[cfg(feature = "xls")]
use crate::format::xls::Xls;
#[cfg(feature = "xlsb")]
use crate::format::xlsb::Xlsb;
#[cfg(feature = "xlsx")]
use crate::format::xlsx::Xlsx;
use litchi_core::{FileFormat, Metadata};

use crate::error::{Error, Result};
#[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
use crate::view::DocumentView;
#[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
use crate::view::PresentationView;
#[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
use crate::view::WorkbookView;
use crate::{Locked, Signed, ToleranceReport};

/// The broad family of a format, reported as data rather than as an enum
/// layer around the per-format variants (record 0003 §2.1, decision D1).
///
/// ```
/// use litchi_core::FileFormat;
/// use litchi_ops::Family;
///
/// assert_eq!(Family::of(FileFormat::Docx), Family::Document);
/// assert_eq!(Family::of(FileFormat::Odp), Family::Presentation);
/// assert_eq!(Family::of(FileFormat::Xlsb), Family::Workbook);
/// assert_eq!(Family::of(FileFormat::Odg), Family::Other);
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Family {
    /// Word-processing documents: DOC, DOCX, RTF, ODT, ODM, OTH, Pages.
    Document,
    /// Presentations: PPT, PPTX, ODP, Keynote.
    Presentation,
    /// Spreadsheets: XLS, XLSX, XLSB, ODS, Numbers.
    Workbook,
    /// Formats outside the three families: drawings, charts, images,
    /// databases and formulas (ODG, ODC, ODI, ODB, ODF formula).
    Other,
}

impl Family {
    /// The family `format` belongs to.
    ///
    /// ```
    /// use litchi_core::FileFormat;
    /// use litchi_ops::Family;
    ///
    /// assert_eq!(Family::of(FileFormat::Rtf), Family::Document);
    /// ```
    #[must_use]
    pub const fn of(format: FileFormat) -> Self {
        match format {
            FileFormat::Doc
            | FileFormat::Docx
            | FileFormat::Rtf
            | FileFormat::Odt
            | FileFormat::Odm
            | FileFormat::Oth
            | FileFormat::Pages => Self::Document,
            FileFormat::Ppt | FileFormat::Pptx | FileFormat::Odp | FileFormat::Keynote => {
                Self::Presentation
            },
            FileFormat::Xls
            | FileFormat::Xlsx
            | FileFormat::Xlsb
            | FileFormat::Ods
            | FileFormat::Numbers => Self::Workbook,
            _ => Self::Other,
        }
    }
}

/// One opened input: a flat enum with one variant per format, plus the
/// not-yet-editable [`Locked`] and [`Signed`] states (ADR 0033 §4, record
/// 0003 §2.1).
///
/// Every variant is one `Arc`-backed handle, so a `File` is a tag and one
/// pointer, is `Send + Sync`, and clones with one reference-count increment.
/// `open`, `open_bytes` and their
/// variants produce it; detection decides the variant from the content. A
/// file name never selects a variant: an OOXML name only refuses content
/// that is not an OOXML package, and names the format of an encrypted one,
/// which its content hides until it is unlocked.
///
/// Tier 1 and tier 2 variants (DOCX, XLSX, PPTX, XLSB, DOC, XLS, PPT) retain
/// the leaf's editable owner. RTF retains its native package and supports the
/// paragraph-text write subset; ODT, ODS and ODP remain read-only for now.
/// The enum is `#[non_exhaustive]`: further formats join as variants. Each
/// variant exists only when its cargo feature is enabled.
///
/// The format-neutral reads are methods on `File` itself; family-specific
/// reads go through `File::document`, `File::presentation` and
/// `File::workbook`, and format-specific ones through the handle a variant
/// carries.
///
/// ```no_run
/// use litchi_ops::File;
///
/// # #[cfg(feature = "docx")]
/// # fn main() -> litchi_ops::Result<()> {
/// let file = litchi_ops::open("report.docx")?;
/// println!("{:?}: {}", file.format(), file.text()?);
/// match &file {
///     File::Docx(docx) => println!("{} sections", docx.section_inventory()?.sections().len()),
///     File::Locked(locked) => println!("encrypted with {:?}", locked.encryption()),
///     _ => {},
/// }
/// # Ok(())
/// # }
/// # #[cfg(not(feature = "docx"))]
/// # fn main() {}
/// ```
#[derive(Clone)]
#[non_exhaustive]
pub enum File {
    /// A WordprocessingML (`.docx`) document.
    #[cfg(feature = "docx")]
    Docx(Docx),
    /// A SpreadsheetML (`.xlsx`) workbook.
    #[cfg(feature = "xlsx")]
    Xlsx(Xlsx),
    /// A PresentationML (`.pptx`) presentation.
    #[cfg(feature = "pptx")]
    Pptx(Pptx),
    /// A binary SpreadsheetML (`.xlsb`) workbook.
    #[cfg(feature = "xlsb")]
    Xlsb(Xlsb),
    /// A legacy Word (`.doc`) document.
    #[cfg(feature = "doc")]
    Doc(Doc),
    /// A legacy Excel (`.xls`) workbook.
    #[cfg(feature = "xls")]
    Xls(Xls),
    /// A legacy PowerPoint (`.ppt`) presentation.
    #[cfg(feature = "ppt")]
    Ppt(Ppt),
    /// A Rich Text Format (`.rtf`) document with paragraph-text writes.
    #[cfg(feature = "rtf")]
    Rtf(Rtf),
    /// An OpenDocument Text (`.odt`) document; read-only.
    #[cfg(feature = "odt")]
    Odt(Odt),
    /// An OpenDocument Spreadsheet (`.ods`) workbook; read-only.
    #[cfg(feature = "ods")]
    Ods(Ods),
    /// An OpenDocument Presentation (`.odp`) presentation; read-only.
    #[cfg(feature = "odp")]
    Odp(Odp),
    /// An encrypted input: neither readable nor editable until it is
    /// unlocked with its password.
    Locked(Locked),
    /// A signed input: readable through the file it wraps; editing must
    /// consume the signature explicitly.
    Signed(Signed),
}

const _: () = {
    const fn assert_handle<T: Send + Sync + Clone>() {}
    assert_handle::<File>();
    assert_handle::<Locked>();
    assert_handle::<Signed>();
    assert_handle::<Error>();
    // A tag plus one pointer (record 0003 decision D2).
    assert!(size_of::<File>() == 2 * size_of::<usize>());
};
#[cfg(target_pointer_width = "64")]
const _: () = assert!(size_of::<File>() == 16);

impl File {
    /// The format of the content.
    ///
    /// A [`File::Signed`] reports the format of the file it wraps. Only a
    /// [`File::Locked`] OOXML package whose format cannot be known before it
    /// is unlocked returns `None`: OOXML encryption covers the whole package,
    /// so only the file name can name its format, and an in-memory input has
    /// none.
    ///
    /// ```
    /// # #[cfg(feature = "rtf")]
    /// # {
    /// use litchi_core::FileFormat;
    ///
    /// let file = litchi_ops::open_bytes(br"{\rtf1 Hello}".to_vec()).unwrap();
    /// assert_eq!(file.format(), Some(FileFormat::Rtf));
    /// # }
    /// ```
    #[must_use]
    pub fn format(&self) -> Option<FileFormat> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(_) => Some(FileFormat::Docx),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(_) => Some(FileFormat::Xlsx),
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => Some(FileFormat::Pptx),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(_) => Some(FileFormat::Xlsb),
            #[cfg(feature = "doc")]
            Self::Doc(_) => Some(FileFormat::Doc),
            #[cfg(feature = "xls")]
            Self::Xls(_) => Some(FileFormat::Xls),
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => Some(FileFormat::Ppt),
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => Some(FileFormat::Rtf),
            #[cfg(feature = "odt")]
            Self::Odt(_) => Some(FileFormat::Odt),
            #[cfg(feature = "ods")]
            Self::Ods(_) => Some(FileFormat::Ods),
            #[cfg(feature = "odp")]
            Self::Odp(_) => Some(FileFormat::Odp),
            Self::Locked(locked) => locked.format(),
            Self::Signed(signed) => Some(signed.format()),
        }
    }

    /// Whether the file is still signed: a [`File::Signed`], or the file one
    /// wraps ([`Signed::file`]). Editing it is refused until
    /// [`Signed::unsign`] consumes the signature.
    ///
    /// The file [`Signed::unsign`] returns is not signed in this sense,
    /// although its bytes keep the signature until its first write removes
    /// it. Nothing is verified: the signature was only detected at open.
    ///
    /// ```
    /// # #[cfg(feature = "rtf")]
    /// # {
    /// let file = litchi_ops::open_bytes(br"{\rtf1 Hello}".to_vec()).unwrap();
    /// assert!(!file.is_signed());
    /// # }
    /// ```
    #[must_use]
    pub fn is_signed(&self) -> bool {
        #[cfg(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "pptx",
            feature = "xls",
            feature = "xlsx",
            feature = "xlsb"
        ))]
        use crate::format::signature::Signature;

        match self {
            Self::Signed(_) => true,
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle.signature() == Signature::Signed,
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.signature() == Signature::Signed,
            _ => false,
        }
    }

    /// This file under another signature disposition (record 0013). Only
    /// the tier 1 and tier 2 variants carry one; the openers never sign
    /// another variant.
    #[cfg(any(
        feature = "doc",
        feature = "docx",
        feature = "ppt",
        feature = "pptx",
        feature = "xls",
        feature = "xlsx",
        feature = "xlsb"
    ))]
    pub(crate) fn with_signature(&self, signature: crate::format::signature::Signature) -> Self {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => Self::Docx(handle.with_signature(signature)),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => Self::Xlsx(handle.with_signature(signature)),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => Self::Pptx(handle.with_signature(signature)),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => Self::Xlsb(handle.with_signature(signature)),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => Self::Doc(handle.with_signature(signature)),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => Self::Xls(handle.with_signature(signature)),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => Self::Ppt(handle.with_signature(signature)),
            // `Locked`, `Signed` and the tier 3 variants carry none.
            other => other.clone(),
        }
    }

    /// The family of [`Self::format`]; `None` exactly when the format is.
    ///
    /// ```
    /// # #[cfg(feature = "rtf")]
    /// # {
    /// use litchi_ops::Family;
    ///
    /// let file = litchi_ops::open_bytes(br"{\rtf1 Hello}".to_vec()).unwrap();
    /// assert_eq!(file.family(), Some(Family::Document));
    /// # }
    /// ```
    #[must_use]
    pub fn family(&self) -> Option<Family> {
        self.format().map(Family::of)
    }

    /// The plain text of the content, in the format's reading order.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// for a [`File::Locked`], [`Code::IoError`](crate::Code::IoError) when a
    /// source-backed file changed after open, and the format's typed error
    /// when its content cannot be decoded.
    ///
    /// ```
    /// # #[cfg(feature = "rtf")]
    /// # {
    /// let file = litchi_ops::open_bytes(br"{\rtf1 Hello}".to_vec()).unwrap();
    /// assert_eq!(file.text().unwrap().trim(), "Hello");
    /// # }
    /// ```
    pub fn text(&self) -> Result<String> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle.text(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle.text(),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle.text(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle.text(),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.text(),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle.text(),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.text(),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle.text(),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => handle.text(),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle.text(),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle.text(),
            Self::Locked(locked) => Err(locked.read_error()),
            Self::Signed(signed) => signed.file().text(),
        }
    }

    /// The document properties: core properties for OOXML, OLE property
    /// sets for the legacy binary formats, `meta.xml` for OpenDocument and
    /// the `\info` group for RTF.
    ///
    /// `None` means the format reports no metadata at all; PPTX and PPT do so
    /// when the package carries no core properties with data. The other
    /// formats report empty metadata instead, as their facade readers did.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// for a [`File::Locked`], [`Code::IoError`](crate::Code::IoError) when a
    /// source-backed file changed after open, and the format's typed error
    /// when a properties part is present but cannot be read.
    pub fn metadata(&self) -> Result<Option<Metadata>> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle.metadata(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle.metadata(),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle.metadata(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle.metadata(),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.metadata(),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle.metadata(),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.metadata(),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle.metadata(),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => handle.metadata(),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle.metadata(),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle.metadata(),
            Self::Locked(locked) => Err(locked.read_error()),
            Self::Signed(signed) => signed.file().metadata(),
        }
    }

    /// Everything the reader repaired so far, as the owning format crate
    /// recorded it. See [`ToleranceReport`] and each handle's
    /// `tolerance_report` for what a format covers.
    ///
    /// # Errors
    ///
    /// Returns [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// for a [`File::Locked`], [`Code::IoError`](crate::Code::IoError) when a
    /// source-backed file changed after open, and the format's typed error
    /// when the parts this report reads cannot be parsed.
    pub fn tolerance_report(&self) -> Result<ToleranceReport> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle.tolerance_report().map(ToleranceReport::Docx),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle.tolerance_report().map(ToleranceReport::Xlsx),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle.tolerance_report().map(ToleranceReport::Pptx),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle.tolerance_report().map(ToleranceReport::Xlsb),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.tolerance_report().map(ToleranceReport::Doc),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle.tolerance_report().map(ToleranceReport::Xls),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.tolerance_report().map(ToleranceReport::Ppt),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle.tolerance_report().map(ToleranceReport::Rtf),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => handle.tolerance_report().map(ToleranceReport::Odt),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle.tolerance_report().map(ToleranceReport::Ods),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle.tolerance_report().map(ToleranceReport::Odp),
            Self::Locked(locked) => Err(locked.read_error()),
            Self::Signed(signed) => signed.file().tolerance_report(),
        }
    }

    /// The word-processing view: paragraphs, tables and elements. Available
    /// when a document format (`doc`, `docx`, `rtf`, `odt`) is enabled.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat)
    /// naming the format ([`Error::detected`]) when the file is not a
    /// document, and [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// for a [`File::Locked`] that is, or may be, one.
    ///
    /// ```no_run
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = litchi_ops::open("report.docx")?;
    /// for paragraph in file.document()?.paragraphs()? {
    ///     println!("{}", paragraph.text()?);
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[cfg(any(feature = "doc", feature = "docx", feature = "rtf", feature = "odt"))]
    pub fn document(&self) -> Result<DocumentView<'_>> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => Ok(DocumentView::docx(handle)),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => Ok(DocumentView::doc(handle)),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => Ok(DocumentView::rtf(handle)),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => Ok(DocumentView::odt(handle)),
            Self::Locked(locked) => Err(locked.view_error(Family::Document)),
            Self::Signed(signed) => signed.file().document(),
            #[allow(
                unreachable_patterns,
                reason = "every variant is a document in builds with only document formats"
            )]
            other => Err(Error::unsupported_format(other.format())),
        }
    }

    /// The presentation view: slides and slide count. Available when a
    /// presentation format (`ppt`, `pptx`, `odp`) is enabled.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat)
    /// naming the format ([`Error::detected`]) when the file is not a
    /// presentation, and
    /// [`Code::DocumentProtected`](crate::Code::DocumentProtected) for a
    /// [`File::Locked`] that is, or may be, one.
    ///
    /// ```no_run
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = litchi_ops::open("deck.pptx")?;
    /// let presentation = file.presentation()?;
    /// println!("{} slides", presentation.slide_count()?);
    /// # Ok(())
    /// # }
    /// ```
    #[cfg(any(feature = "ppt", feature = "pptx", feature = "odp"))]
    pub fn presentation(&self) -> Result<PresentationView<'_>> {
        match self {
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => Ok(PresentationView::pptx(handle)),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => Ok(PresentationView::ppt(handle)),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => Ok(PresentationView::odp(handle)),
            Self::Locked(locked) => Err(locked.view_error(Family::Presentation)),
            Self::Signed(signed) => signed.file().presentation(),
            #[allow(
                unreachable_patterns,
                reason = "every variant is a presentation in builds with only presentation formats"
            )]
            other => Err(Error::unsupported_format(other.format())),
        }
    }

    /// The workbook view: worksheet names and count, and the generic
    /// [`litchi_core::sheet::WorkbookTrait`]. Available when a workbook
    /// format (`xls`, `xlsx`, `xlsb`, `ods`) is enabled.
    ///
    /// # Errors
    ///
    /// Returns [`Code::UnsupportedFormat`](crate::Code::UnsupportedFormat)
    /// naming the format ([`Error::detected`]) when the file is not a
    /// workbook, and [`Code::DocumentProtected`](crate::Code::DocumentProtected)
    /// for a [`File::Locked`] that is, or may be, one.
    ///
    /// ```no_run
    /// # fn main() -> litchi_ops::Result<()> {
    /// let file = litchi_ops::open("budget.xlsx")?;
    /// println!("{:?}", file.workbook()?.worksheet_names()?);
    /// # Ok(())
    /// # }
    /// ```
    #[cfg(any(feature = "xls", feature = "xlsx", feature = "xlsb", feature = "ods"))]
    pub fn workbook(&self) -> Result<WorkbookView<'_>> {
        match self {
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => Ok(WorkbookView::xlsx(handle)),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => Ok(WorkbookView::xlsb(handle)),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => Ok(WorkbookView::xls(handle)),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => Ok(WorkbookView::ods(handle)),
            Self::Locked(locked) => Err(locked.view_error(Family::Workbook)),
            Self::Signed(signed) => signed.file().workbook(),
            #[allow(
                unreachable_patterns,
                reason = "every variant is a workbook in builds with only workbook formats"
            )]
            other => Err(Error::unsupported_format(other.format())),
        }
    }
}

impl std::fmt::Debug for File {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle.fmt(formatter),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle.fmt(formatter),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle.fmt(formatter),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle.fmt(formatter),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.fmt(formatter),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle.fmt(formatter),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.fmt(formatter),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle.fmt(formatter),
            #[cfg(feature = "odt")]
            Self::Odt(handle) => handle.fmt(formatter),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle.fmt(formatter),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle.fmt(formatter),
            Self::Locked(locked) => locked.fmt(formatter),
            Self::Signed(signed) => signed.fmt(formatter),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn family_covers_every_detected_format() {
        use FileFormat as F;
        let cases = [
            (F::Doc, Family::Document),
            (F::Docx, Family::Document),
            (F::Rtf, Family::Document),
            (F::Odt, Family::Document),
            (F::Odm, Family::Document),
            (F::Oth, Family::Document),
            (F::Pages, Family::Document),
            (F::Ppt, Family::Presentation),
            (F::Pptx, Family::Presentation),
            (F::Odp, Family::Presentation),
            (F::Keynote, Family::Presentation),
            (F::Xls, Family::Workbook),
            (F::Xlsx, Family::Workbook),
            (F::Xlsb, Family::Workbook),
            (F::Ods, Family::Workbook),
            (F::Numbers, Family::Workbook),
            (F::Odg, Family::Other),
            (F::Odc, Family::Other),
            (F::Odi, Family::Other),
            (F::Odb, Family::Other),
            (F::Odf, Family::Other),
        ];
        for (format, family) in cases {
            assert_eq!(Family::of(format), family, "{format:?}");
        }
    }
}
