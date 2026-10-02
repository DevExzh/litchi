//! Format-neutral access to the repairs a facade read performed.
//!
//! Every format crate applies reference-compatible repairs to non-structural
//! deviations and *records* each one in its own bounded `ToleranceReport`
//! (see `docs/LENIENCY_AUDIT.md`, "The tolerance contract"). The unified
//! facades — [`Document`](crate::Document), [`Presentation`](crate::Presentation)
//! and [`Workbook`](crate::Workbook) — expose that report through
//! `tolerance_report()`, wrapped in [`ToleranceReport`] so one call works for
//! every format while the detailed, per-format defect classes stay one
//! `match` away.
//!
//! Package metadata the facade reads alongside the document is part of the
//! report: OOXML core-properties repairs (`litchi_ooxml_common`) appear as the
//! owning format's `Defect::Properties` entries, and OLE property-set repairs
//! as `Defect::PropertySet` entries of the legacy formats.

use litchi_core::detection::FileFormat;

/// Everything a unified facade's read repaired, as the owning format crate
/// recorded it.
///
/// An empty report ([`ToleranceReport::is_clean`]) means the document
/// conformed and nothing was substituted. Each variant carries the format
/// crate's own report, whose entries name the repaired defect class and its
/// location; see that crate's `leniency` module for the closed set of
/// repairs. Variants exist only for the formats enabled in this build.
///
/// # When lazily parsed parts contribute
///
/// Formats that parse parts on demand report what has been parsed when the
/// report is taken, plus everything parsed at open. The facades' own
/// documentation lists, per format, which parts are covered.
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub enum ToleranceReport {
    /// Legacy Word binary (`.doc`).
    #[cfg(feature = "doc")]
    Doc(litchi_doc::ToleranceReport),
    /// WordprocessingML (`.docx`).
    #[cfg(feature = "docx")]
    Docx(litchi_docx::ToleranceReport),
    /// Rich Text Format (`.rtf`).
    #[cfg(feature = "rtf")]
    Rtf(litchi_rtf::leniency::ToleranceReport),
    /// OpenDocument Text (`.odt`).
    #[cfg(feature = "odt")]
    Odt(litchi_odf_common::ToleranceReport),
    /// Legacy PowerPoint binary (`.ppt`).
    #[cfg(feature = "ppt")]
    Ppt(litchi_ppt::ToleranceReport),
    /// PresentationML (`.pptx`).
    #[cfg(feature = "pptx")]
    Pptx(litchi_pptx::ToleranceReport),
    /// OpenDocument Presentation (`.odp`).
    #[cfg(feature = "odp")]
    Odp(litchi_odf_common::ToleranceReport),
    /// Legacy Excel binary (`.xls`).
    #[cfg(feature = "xls")]
    Xls(litchi_xls::ToleranceReport),
    /// SpreadsheetML (`.xlsx`).
    #[cfg(feature = "xlsx")]
    Xlsx(litchi_xlsx::ToleranceReport),
    /// Binary SpreadsheetML (`.xlsb`).
    #[cfg(feature = "xlsb")]
    Xlsb(litchi_xlsb::ToleranceReport),
    /// OpenDocument Spreadsheet (`.ods`).
    #[cfg(feature = "ods")]
    Ods(litchi_odf_common::ToleranceReport),
    /// A format whose reader keeps no repair report (the Apple iWork
    /// formats).
    ///
    /// Conformance is unknown, so such a report is never
    /// [`clean`](ToleranceReport::is_clean); its counts are zero.
    Untracked(FileFormat),
}

impl ToleranceReport {
    /// The format whose reader produced this report.
    #[must_use]
    pub const fn format(&self) -> FileFormat {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(_) => FileFormat::Doc,
            #[cfg(feature = "docx")]
            Self::Docx(_) => FileFormat::Docx,
            #[cfg(feature = "rtf")]
            Self::Rtf(_) => FileFormat::Rtf,
            #[cfg(feature = "odt")]
            Self::Odt(_) => FileFormat::Odt,
            #[cfg(feature = "ppt")]
            Self::Ppt(_) => FileFormat::Ppt,
            #[cfg(feature = "pptx")]
            Self::Pptx(_) => FileFormat::Pptx,
            #[cfg(feature = "odp")]
            Self::Odp(_) => FileFormat::Odp,
            #[cfg(feature = "xls")]
            Self::Xls(_) => FileFormat::Xls,
            #[cfg(feature = "xlsx")]
            Self::Xlsx(_) => FileFormat::Xlsx,
            #[cfg(feature = "xlsb")]
            Self::Xlsb(_) => FileFormat::Xlsb,
            #[cfg(feature = "ods")]
            Self::Ods(_) => FileFormat::Ods,
            Self::Untracked(format) => *format,
        }
    }

    /// Whether the format's reader records repairs at all.
    ///
    /// `false` only for [`ToleranceReport::Untracked`].
    #[must_use]
    pub const fn is_tracked(&self) -> bool {
        !matches!(self, Self::Untracked(_))
    }

    /// Whether the read conformed: a tracked report with no repair, recorded
    /// or unrecorded.
    #[must_use]
    pub fn is_clean(&self) -> bool {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(report) => report.is_clean(),
            #[cfg(feature = "docx")]
            Self::Docx(report) => report.is_clean(),
            #[cfg(feature = "rtf")]
            Self::Rtf(report) => report.is_clean(),
            #[cfg(feature = "odt")]
            Self::Odt(report) => report.is_clean(),
            #[cfg(feature = "ppt")]
            Self::Ppt(report) => report.is_clean(),
            #[cfg(feature = "pptx")]
            Self::Pptx(report) => report.is_clean(),
            #[cfg(feature = "odp")]
            Self::Odp(report) => report.is_clean(),
            #[cfg(feature = "xls")]
            Self::Xls(report) => report.is_clean(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(report) => report.is_clean(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(report) => report.is_clean(),
            #[cfg(feature = "ods")]
            Self::Ods(report) => report.is_clean(),
            Self::Untracked(_) => false,
        }
    }

    /// Total repairs performed, recorded or not.
    #[must_use]
    pub fn total(&self) -> u64 {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(report) => report.total(),
            #[cfg(feature = "docx")]
            Self::Docx(report) => report.total(),
            #[cfg(feature = "rtf")]
            Self::Rtf(report) => report.total(),
            #[cfg(feature = "odt")]
            Self::Odt(report) => report.total(),
            #[cfg(feature = "ppt")]
            Self::Ppt(report) => report.total(),
            #[cfg(feature = "pptx")]
            Self::Pptx(report) => report.total(),
            #[cfg(feature = "odp")]
            Self::Odp(report) => report.total(),
            #[cfg(feature = "xls")]
            Self::Xls(report) => report.total(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(report) => report.total(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(report) => report.total(),
            #[cfg(feature = "ods")]
            Self::Ods(report) => report.total(),
            Self::Untracked(_) => 0,
        }
    }

    /// Repairs performed but not individually stored once the format
    /// report's bound was reached.
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(report) => report.unrecorded(),
            #[cfg(feature = "docx")]
            Self::Docx(report) => report.unrecorded(),
            #[cfg(feature = "rtf")]
            Self::Rtf(report) => report.unrecorded(),
            #[cfg(feature = "odt")]
            Self::Odt(report) => report.unrecorded(),
            #[cfg(feature = "ppt")]
            Self::Ppt(report) => report.unrecorded(),
            #[cfg(feature = "pptx")]
            Self::Pptx(report) => report.unrecorded(),
            #[cfg(feature = "odp")]
            Self::Odp(report) => report.unrecorded(),
            #[cfg(feature = "xls")]
            Self::Xls(report) => report.unrecorded(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(report) => report.unrecorded(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(report) => report.unrecorded(),
            #[cfg(feature = "ods")]
            Self::Ods(report) => report.unrecorded(),
            Self::Untracked(_) => 0,
        }
    }

    /// The stable description of every individually recorded repair, in
    /// discovery order.
    ///
    /// This is the format-neutral view; match on the variant for the defect
    /// classes and locations.
    #[must_use]
    pub fn repairs(&self) -> Vec<&'static str> {
        match self {
            #[cfg(feature = "doc")]
            Self::Doc(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "docx")]
            Self::Docx(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "rtf")]
            Self::Rtf(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "odt")]
            Self::Odt(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "ppt")]
            Self::Ppt(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "pptx")]
            Self::Pptx(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "odp")]
            Self::Odp(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "xls")]
            Self::Xls(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect().description())
                .collect(),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect().repair())
                .collect(),
            #[cfg(feature = "ods")]
            Self::Ods(report) => report
                .defects()
                .iter()
                .map(|entry| entry.defect.repair())
                .collect(),
            Self::Untracked(_) => Vec::new(),
        }
    }
}
