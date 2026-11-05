//! Format-neutral access to the repairs a read performed.
//!
//! Every format crate applies reference-compatible repairs to non-structural
//! deviations and *records* each one in its own bounded `ToleranceReport`
//! (see `docs/LENIENCY_AUDIT.md`, "The tolerance contract").
//! [`File::tolerance_report`](crate::File::tolerance_report) returns the
//! report the opened file's format handle keeps, wrapped in
//! [`ToleranceReport`], so one call works for every format while the
//! detailed, per-format defect classes stay one `match` away.
//!
//! ```
//! # #[cfg(feature = "docx")]
//! # {
//! let path = concat!(
//!     env!("CARGO_MANIFEST_DIR"),
//!     "/../../test-data/libreoffice-core/oox/qa/unit/data/chart-data-label-char-color.docx"
//! );
//! let report = litchi::open(path)?.tolerance_report()?;
//! assert_eq!(report.format(), litchi::FileFormat::Docx);
//! assert!(report.is_clean());
//! # }
//! # Ok::<(), litchi::Error>(())
//! ```
//!
//! Package metadata read alongside the document is part of the report: OOXML
//! core-properties repairs (`litchi_ooxml_common`) appear as the owning
//! format's `Defect::Properties` entries, and OLE property-set repairs as
//! `Defect::PropertySet` entries of the legacy formats.
//!
//! The type is `litchi_ops::ToleranceReport`, re-exported here.
//!
//! Apple iWork documents, which [`litchi::open`](crate::open) refuses and the
//! `iwork` module reads, keep no repair report;
//! [`ToleranceReport::Untracked`] names such a format.

pub use litchi_ops::ToleranceReport;
