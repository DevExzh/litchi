//! Repair reporting for non-structural defects in `OpenDocument` packages.
//!
//! Two kinds of repair can happen while a package is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed defect enumerations changes: ZIP framing, XML well-formedness,
//! manifest structure, and decryption remain hard errors under every variant.
//!
//! **LibreOffice-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover container
//! deviations ODF 1.2 Part 3 constrains for *writers* but real producers emit
//! anyway — layouts LibreOffice reads past without complaint, so rejecting
//! them would cost the caller the entire document over packaging that cannot
//! change what a conforming file means. Each substitutes the value documented
//! on the corresponding [`PackageDefect`] variant. This class implements the
//! `mimetype`-member grammar repairs of LENIENCY_AUDIT F9.1: every package
//! open path (owned, positional source-backed, and the prepared detection
//! handoff) accepts a compressed, extra-field-carrying, or relocated
//! `mimetype` member, reads it by name, validates its content exactly, and
//! records the deviation. It also covers the manifest root-entry fallback of
//! LENIENCY_AUDIT F9.2: a manifest without a `/` entry, or whose `/` entry
//! has an empty media type, adopts the `mimetype` member's value, exactly as
//! LibreOffice's media-type fallback does; two non-empty values that disagree
//! stay a hard error. A conforming package yields an empty report.
//!
//! **LibreOffice-compatible content repairs** ([`ContentDefect`]) apply the
//! same posture inside a document's XML parts: one malformed value or one
//! unknown element degrades or is skipped locally — exactly as LibreOffice's
//! import does — instead of failing the whole document, and each is
//! recorded with a [`DefectLocation`] where one exists. Well-formedness,
//! entity safety, and resource limits are never relaxed.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. No ODF repair currently requires a judgment call,
//! so [`Leniency::Strict`] is the only variant; opt-in variants land here as
//! the roadmap adds repairs that need one.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` established
//! for legacy Word documents so every reader in the workspace reports
//! tolerated defects the same way.
//!
//! # Detection and the report
//!
//! Format detection ([`crate::detect`]) is document-less: its probes return a
//! bare classification and cannot hand back a [`ToleranceReport`]. A tolerant
//! detection therefore signals a repaired deviation by *succeeding* — the
//! classification itself is the signal — and the record is carried by the
//! subsequent open. Callers that need the record open through
//! `detect::prepared_or_original` (see [`crate::core::PreparedPackage`]) or a
//! family facade, all of which retain the opening report on the package or
//! document handle they return.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous package must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural package defects.
///
/// Structural defects — malformed ZIP framing, broken XML grammar, an
/// unreadable manifest, or unsupported/absent decryption material — are hard
/// errors under every variant. The LibreOffice-compatible repairs listed by
/// [`PackageDefect`] are applied and recorded under every variant, including
/// [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default and currently the only variant: the
    /// LibreOffice-compatible [`PackageDefect`] repairs are still applied and
    /// recorded in the [`ToleranceReport`], while no opt-in repair class has
    /// been needed yet.
    #[default]
    Strict,
}

/// The class of package-container defect the reader repairs the way
/// LibreOffice repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each one matches what LibreOffice does with package
/// layouts ODF 1.2 Part 3 §3.3 constrains only for producers, so applying it
/// never changes what a conforming document means.
///
/// The `Mimetype*` variants are the `mimetype`-member grammar repairs of
/// LENIENCY_AUDIT F9.1. One package can record several of them (for example a
/// deflated member that is also not first), in the order listed here.
/// Everything else about the member stays strict: its content must name the
/// expected MIME type exactly, and an encrypted or unreadable member is a hard
/// error. [`PackageDefect::ManifestRootMediaTypeMissing`] is the manifest
/// root-entry fallback of LENIENCY_AUDIT F9.2.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PackageDefect {
    /// The `mimetype` member declared a ZIP compression method other than
    /// Store — almost always Deflate from a package re-zipped with an
    /// ordinary archiver (`zip -r`, PHP `ZipArchive`).
    ///
    /// ODF 1.2 Part 3 §3.3 requires the `mimetype` stream to be uncompressed,
    /// but that is a producer rule: LibreOffice reads the stream through
    /// ordinary ZIP APIs, which inflate any compression method
    /// (`ZipPackage::parseManifest`, ZipPackage.cxx). The repair inflates the
    /// member through the central directory and uses its decoded bytes
    /// exactly as a conforming stored member's bytes. The defect context is
    /// the declared ZIP compression method ID (for example 8 for Deflate).
    MimetypeDeflated,
    /// The local header of the `mimetype` member carried an extra field, such
    /// as an Info-ZIP Extended Timestamp left by re-zipping a package.
    ///
    /// ODF 1.2 Part 3 §3.3 requires the `mimetype` stream to carry no extra
    /// field; LibreOffice never inspects the local extra field on the read
    /// path. The repair ignores the extra field and reads the member payload
    /// through the central directory. (The crate also ships an opt-in
    /// file-rewriting repair for this class,
    /// `odf.repair.mimetype_local_extra` in [`crate::repair`]; read-side
    /// tolerance is independent of it.) The defect context is 0.
    MimetypeExtraField,
    /// The `mimetype` member was not the first local entry at offset 0.
    ///
    /// ODF 1.2 Part 3 §3.3 requires the `mimetype` stream to be the first
    /// stream so detectors can sniff it in place; LibreOffice locates it by
    /// name like any other member. The repair reads the member by name from
    /// the central directory. The defect context is 0.
    MimetypeNotFirst,
    /// The manifest had no `/` root entry, or its root entry declared an
    /// empty `manifest:media-type`, so the package media type came from the
    /// `mimetype` member alone.
    ///
    /// ODF 1.2 Part 3 §4.2 says the manifest *should* describe the root;
    /// minimal third-party writers omit it. LibreOffice treats an empty root
    /// media type as a fallback, not a conflict: it adopts the `mimetype`
    /// stream value (`m_bMediaTypeFallbackUsed`, `ZipPackage::parseManifest`
    /// in ZipPackage.cxx) and throws only when both values are non-empty and
    /// differ. The repair adopts the `mimetype` member's value as the
    /// manifest root media type; a non-empty conflicting root stays a hard
    /// error. The defect context is 0 when the root entry is absent and 1
    /// when it is present with an empty media type.
    ManifestRootMediaTypeMissing,
}

impl PackageDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::MimetypeDeflated => {
                "inflated the mimetype member through ordinary ZIP decompression and read it by name from the central directory"
            },
            Self::MimetypeExtraField => {
                "ignored the mimetype local-header extra field; the member was read by name from the central directory"
            },
            Self::MimetypeNotFirst => {
                "located the mimetype member by name in the central directory instead of requiring the offset-0 local header"
            },
            Self::ManifestRootMediaTypeMissing => {
                "adopted the mimetype member's value as the package media type because the manifest root entry was absent or had an empty media type"
            },
        }
    }
}

/// The class of document-content defect the reader repairs the way
/// LibreOffice's import repairs it.
///
/// This enumeration is closed by design. Each variant documents its repair,
/// its [`ToleratedDefect::context`], and its [`DefectLocation`]. Applying one
/// never changes what a conforming document means: the repair engages only
/// on content LibreOffice itself reads past.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContentDefect {
    /// An ODS cell's typed value could not be read: a `float`, `currency`, or
    /// `percentage` cell whose `office:value` is absent or not a finite
    /// decimal (a locale-comma `"1,5"`, `"NaN"`, an empty string), a
    /// `boolean` cell whose value is not `true`/`false`/`1`/`0`, or a `date`
    /// or `time` cell without its value attribute (LENIENCY_AUDIT F9.8).
    ///
    /// LibreOffice's Calc import never aborts on one bad cell value: a failed
    /// numeric conversion leaves the cell at a default value and the import
    /// continues (`ScXMLTableRowCellContext`, xmlcelli.cxx). The repair keeps
    /// the cell, its displayed text, formula, style and span, and stores the
    /// value as `CellValue::Unknown { kind, value }` with the declared value
    /// type and the raw lexical (or `None` when the attribute was absent), so
    /// nothing is invented and an edit can never silently rewrite it. The
    /// context is the declared value type: 0 `float`/`double`/`decimal`, 1
    /// `currency`, 2 `percentage`, 3 `boolean`, 4 `date`, 5 `time`. The
    /// location is the cell ([`DefectLocation::Cell`]).
    CellValueUnparseable,
    /// A child of `table:calculation-settings` other than `table:null-date`
    /// and `table:iteration` — a vendor or LibreOffice extension element, or
    /// an element of a later ODF version (LENIENCY_AUDIT F9.9).
    ///
    /// ODF 1.2 Part 1 §3.3 makes foreign elements ignorable, and LibreOffice's
    /// `ScXMLCalculationSettingsContext::createFastChildContext` returns no
    /// context for any other child, skipping it. The repair skips the child
    /// with its whole subtree (depth-balanced; only predefined entity and
    /// character references are accepted inside it) and keeps every other
    /// setting. The context is 0 for a child outside the table namespace and
    /// 1 for an unknown table-namespace child. There is no location.
    CalculationSettingsUnknownChild,
    /// `table:null-date` followed `table:iteration` inside
    /// `table:calculation-settings`, against the schema's sequence order
    /// (LENIENCY_AUDIT F9.9).
    ///
    /// LibreOffice dispatches each child by name and reads either order. The
    /// repair accepts both settings exactly as in schema order; a duplicate
    /// child stays a hard error. The context is 0. There is no location.
    CalculationSettingsReordered,
}

impl ContentDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::CellValueUnparseable => {
                "kept the cell and its displayed text and stored its unreadable typed value as an uninterpreted value instead of failing the spreadsheet"
            },
            Self::CalculationSettingsUnknownChild => {
                "skipped an unknown child of table:calculation-settings and its subtree"
            },
            Self::CalculationSettingsReordered => {
                "accepted table:null-date after table:iteration in table:calculation-settings"
            },
        }
    }
}

/// Where in a document a [`ContentDefect`] was found.
///
/// Every index is zero-based and counts logical positions, so repeated rows
/// and cells (`table:number-rows-repeated`,
/// `table:number-columns-repeated`) are expanded when counting.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DefectLocation {
    /// A spreadsheet cell.
    Cell {
        /// Sheet index in document order.
        sheet: u32,
        /// Logical row within the sheet.
        row: u64,
        /// Logical column within the row.
        column: u64,
    },
}

/// One class of defect repaired while a package was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a document. [`Defect::Package`] and
/// [`Defect::Content`] classes are repaired under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A package-container defect repaired the way LibreOffice repairs it.
    Package(PackageDefect),
    /// A document-content defect repaired the way LibreOffice repairs it.
    Content(ContentDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Package(defect) => defect.repair(),
            Self::Content(defect) => defect.repair(),
        }
    }
}

impl From<PackageDefect> for Defect {
    fn from(defect: PackageDefect) -> Self {
        Self::Package(defect)
    }
}

impl From<ContentDefect> for Defect {
    fn from(defect: ContentDefect) -> Self {
        Self::Content(defect)
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// What was found: the declared ZIP compression method ID for
    /// [`PackageDefect::MimetypeDeflated`]; 0 (root entry absent) or 1 (root
    /// entry with an empty media type) for
    /// [`PackageDefect::ManifestRootMediaTypeMissing`]; 0 for the other
    /// [`Defect::Package`] classes. Each [`ContentDefect`] variant documents
    /// its own context.
    pub context: u16,
    /// Where the defect was found inside the document, for the
    /// [`ContentDefect`] classes that have a location; `None` otherwise.
    pub location: Option<DefectLocation>,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the document conformed and no substitution was made.
/// After a [`Leniency::Strict`] read the report can still list
/// [`Defect::Package`] repairs, which LibreOffice itself applies.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the document parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at `MAX_RECORDED_DEFECTS` stored entries.
    ///
    /// This is format-family plumbing used by the ODF crates while opening a
    /// package; it is hidden because recording is part of a read, not caller
    /// API.
    #[doc(hidden)]
    pub fn record(&mut self, defect: impl Into<Defect>, context: u16) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            context,
            location: None,
        });
    }

    /// Record one repair found at `location`, saturating like
    /// [`Self::record`].
    ///
    /// This is format-family plumbing; it is hidden because recording is
    /// part of a read, not caller API.
    #[doc(hidden)]
    pub fn record_at(&mut self, defect: impl Into<Defect>, context: u16, location: DefectLocation) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            context,
            location: Some(location),
        });
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    ///
    /// This is format-family plumbing used to aggregate partial reports at an
    /// open boundary; it is hidden because merging is part of a read, not
    /// caller API.
    #[doc(hidden)]
    pub fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            self.push(*tolerated);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(PackageDefect::MimetypeDeflated, 8);
        report.record(PackageDefect::MimetypeExtraField, 0);
        report.record(PackageDefect::MimetypeNotFirst, 0);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Package(PackageDefect::MimetypeDeflated), 8),
                (Defect::Package(PackageDefect::MimetypeExtraField), 0),
                (Defect::Package(PackageDefect::MimetypeNotFirst), 0),
            ]
        );
    }

    /// A hostile package must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for _ in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(PackageDefect::MimetypeExtraField, 0);
        }
        for _ in 0..3 {
            report.record(PackageDefect::MimetypeDeflated, 8);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(PackageDefect::MimetypeNotFirst, 0);

        let mut parsed = ToleranceReport::default();
        parsed.record(PackageDefect::MimetypeDeflated, 8);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(PackageDefect::MimetypeExtraField, 0);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Package(PackageDefect::MimetypeNotFirst),
                context: 0,
                location: None,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn located_defects_keep_their_location_across_a_merge() {
        let cell = DefectLocation::Cell {
            sheet: 1,
            row: 7,
            column: 3,
        };
        let mut parsed = ToleranceReport::default();
        parsed.record_at(ContentDefect::CellValueUnparseable, 3, cell);
        let mut base = ToleranceReport::default();
        base.record(PackageDefect::MimetypeNotFirst, 0);
        base.append(&parsed);
        assert_eq!(base.defects()[0].location, None);
        assert_eq!(
            base.defects()[1],
            ToleratedDefect {
                defect: Defect::Content(ContentDefect::CellValueUnparseable),
                context: 3,
                location: Some(cell),
            }
        );
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(PackageDefect::MimetypeDeflated),
            Defect::from(PackageDefect::MimetypeExtraField),
            Defect::from(PackageDefect::MimetypeNotFirst),
            Defect::from(PackageDefect::ManifestRootMediaTypeMissing),
            Defect::from(ContentDefect::CellValueUnparseable),
            Defect::from(ContentDefect::CalculationSettingsUnknownChild),
            Defect::from(ContentDefect::CalculationSettingsReordered),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Package(PackageDefect::MimetypeDeflated).repair(),
            PackageDefect::MimetypeDeflated.repair()
        );
    }
}
