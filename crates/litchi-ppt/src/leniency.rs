//! Repair reporting for non-structural defects in legacy presentations.
//!
//! Two kinds of repair can happen while a presentation is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: record framing, the UserEdit/persist
//! chain grammar, container structure, and encryption remain hard errors
//! under every variant.
//!
//! **`PowerPoint`-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover bytes
//! MS-PPT leaves undefined or reserved but real producers emit anyway —
//! bytes `PowerPoint`, POI, and LibreOffice read past without complaint, so
//! rejecting them would cost the caller the entire presentation over
//! metadata that cannot change what a conforming file means. Each
//! substitutes the value documented on the corresponding [`TextPropDefect`]
//! variant.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. No repair currently requires opting in; the first
//! candidate on the roadmap is breaking a cyclic `UserEdit` chain (finding
//! F5.9 in `docs/LENIENCY_AUDIT.md`), and the [`Leniency`] variant and
//! open-options plumbing for it land with that finding.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` and
//! `crates/litchi-xls/src/leniency.rs` provide for the other legacy readers
//! so all three behave the same way.
//!
//! # Roadmap state
//!
//! Two findings from `docs/LENIENCY_AUDIT.md` record into this report today,
//! both `PowerPoint`-compatible (default) repairs on the text-box read path:
//!
//! - **F5.1** ([`TextPropDefect::StyleTextPropCoverage`]) — the read path
//!   parses `StyleTextPropAtom` with the clamping POI/LibreOffice posture
//!   (coverage clamped to the text length + 1, parsing stopped at the first
//!   unreadable run, trailing bytes dropped) and records each clamp. The
//!   exact-coverage parser
//!   ([`crate::text_prop::parse_style_text_prop_atom_strict`]) and the
//!   strict [`crate::StyleTextPropAtom`] codec stay in use by the writer and
//!   the text-edit/reparse-verification path, which must not rewrite bytes
//!   it cannot frame exactly.
//! - **F5.2** ([`TextPropDefect::ReservedBulletFlagBits`],
//!   [`TextPropDefect::ReservedWrapFlagBits`]) — reserved
//!   `BulletFlags`/`PFWrapFlags` bits are masked to the defined set instead
//!   of rejecting the paragraph; the defined bits are still applied.
//!
//! Record framing, the `UserEdit`/persist chain, container structure,
//! encryption, and the configured record/interaction resource limits are not
//! touched by either repair and stay fatal.
//!
//! # Where the report lives
//!
//! Text-style repairs fire during *lazy* parsing, so the report is threaded
//! explicitly (`&mut ToleranceReport`, no interior mutability in the report
//! itself) through the text-box parse and attached to the lazily cached value
//! it describes:
//!
//! - [`crate::Slide::tolerance_report`] and
//!   [`crate::SpeakerNotes::tolerance_report`] return the repairs made while
//!   that slide's or notes page's shapes were materialized. The report is
//!   cached in the same `OnceCell` as the shapes, so it is complete no matter
//!   which accessor (`shapes`, `shapes_with_report`, `tolerance_report`)
//!   triggered the parse.
//! - `shapes_with_report` appends that cached report to a caller-owned one,
//!   on every call, so one report can be threaded across many slides.
//! - [`crate::Presentation::tolerance_report`] is the open-time aggregate.
//!   No open-time parse repairs anything yet (the record tree, persist chain,
//!   and slide directory take `&mut ToleranceReport` when findings
//!   F5.4–F5.10 land), so it is clean for every presentation that opens.
//! - [`crate::Presentation::collect_tolerance_report`] returns the open-time
//!   aggregate plus the repairs of every slide's and notes page's shapes —
//!   the one-call answer to "what did the reader change about this file?".
//! - Lower-level readers ([`crate::EscherTextboxWrapper::new_with_report`],
//!   [`crate::TextRunExtractor::extract_from_records_with_report`], and the
//!   `Presentation::*_with_report` stream walks) record into a caller-owned
//!   report; their report-less conveniences apply the same repairs without
//!   enumerating them.
//!
//! Every recorded defect carries a [`DefectLocation`]: the persist ID of the
//! slide or notes page and the `OfficeArt` shape ID of the text box, filled
//! in by the layer that knows each one, plus the per-class `context` value.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous style atom must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural presentation defects.
///
/// Structural defects — malformed record framing or container structure, a
/// broken `UserEdit`/persist-chain grammar, or unsupported/absent decryption
/// material — are hard errors under every variant. The
/// `PowerPoint`-compatible repairs listed by [`TextPropDefect`] are applied
/// and recorded under every variant, including [`Leniency::Strict`].
///
/// The enumeration is non-exhaustive because opt-in variants arrive with the
/// roadmap findings that need them (the first candidate is F5.9); new
/// variants are never a breaking change for callers.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: the `PowerPoint`-compatible [`TextPropDefect`]
    /// repairs are still applied and recorded in the [`ToleranceReport`].
    #[default]
    Strict,
}

/// The class of text-property defect the reader repairs the way `PowerPoint`
/// repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// the text-property repairs a read applies under every [`Leniency`]
/// variant, including [`Leniency::Strict`]. Each one matches what
/// `PowerPoint`, POI, and LibreOffice do with bytes MS-PPT leaves undefined
/// or reserved, so applying it never changes what a conforming presentation
/// means.
///
/// Each variant implements one finding of `docs/LENIENCY_AUDIT.md` (F5.1,
/// F5.2) and is recorded by the text-box read path.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextPropDefect {
    /// A `StyleTextPropAtom` run's `charactersCovered` did not tile exactly
    /// the text length plus one, a run was truncated, or the atom had
    /// trailing bytes (finding F5.1).
    ///
    /// MS-PPT 2.9.40–2.9.41 require coverage equal to the character count
    /// plus one, but nothing prescribes rejecting the document when a
    /// producer miscounts, and miscounted coverage is one of the most common
    /// .ppt deviations. POI only logs and truncates
    /// (`StyleTextPropAtom.java`), and LibreOffice clamps (`svdfppt.cxx`),
    /// so the reader parses with the clamping
    /// [`crate::text_prop::parse_style_text_prop_atom`] posture on the read
    /// path: coverage is clamped to what remains, parsing stops at the first
    /// unreadable run, and trailing bytes are dropped. The strict parser
    /// ([`crate::text_prop::parse_style_text_prop_atom_strict`]) is retained
    /// for the writer and reparse-verification path.
    StyleTextPropCoverage,
    /// A `TextPFRun`'s two-byte `bulletFlags` carried bits outside the set
    /// MS-PPT 2.9.22 defines (finding F5.2).
    ///
    /// MS-PPT 2.9.22 defines bits 0x0001–0x0008 and states the remaining
    /// twelve bits "MUST be zero and MUST be ignored". `PowerPoint`, POI,
    /// and LibreOffice read only the defined bits while third-party
    /// producers routinely leak garbage into the high ones, so the reader
    /// masks `bulletFlags` down to 0x000F and still applies every defined
    /// bit that survives.
    ReservedBulletFlagBits,
    /// A `TextPFRun`'s two-byte `wrapFlags` carried bits outside the set
    /// MS-PPT 2.9.25 defines (finding F5.2).
    ///
    /// MS-PPT 2.9.25 defines bits 0x0001–0x0004 and states the remaining
    /// thirteen bits "MUST be zero and MUST be ignored". As with
    /// [`TextPropDefect::ReservedBulletFlagBits`], the reader masks
    /// `wrapFlags` down to 0x0007 and still applies every defined bit that
    /// survives.
    ReservedWrapFlagBits,
}

impl TextPropDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::StyleTextPropCoverage => {
                "clamped StyleTextPropAtom run coverage to the text length + 1, stopped at the first unreadable run, and dropped trailing bytes; surviving runs were still applied"
            },
            Self::ReservedBulletFlagBits => {
                "masked bulletFlags to the bits MS-PPT 2.9.22 defines (0x000F); defined bits were still applied"
            },
            Self::ReservedWrapFlagBits => {
                "masked wrapFlags to the bits MS-PPT 2.9.25 defines (0x0007); defined bits were still applied"
            },
        }
    }
}

/// One class of defect repaired while a presentation was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a read may change about a presentation. [`Defect::TextProp`] classes
/// are repaired under every [`Leniency`] variant; opt-in classes join this
/// enumeration together with the [`Leniency`] variants that gate them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A text-property defect repaired the way `PowerPoint` repairs it.
    TextProp(TextPropDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::TextProp(defect) => defect.repair(),
        }
    }
}

impl From<TextPropDefect> for Defect {
    fn from(defect: TextPropDefect) -> Self {
        Self::TextProp(defect)
    }
}

/// Where in the presentation a repaired defect was found.
///
/// Each field is filled in by the parse layer that knows it; a field stays
/// `None` when the defect was found by a reader that has no such context
/// (for example [`crate::EscherTextboxWrapper::new_with_report`] called on a
/// detached text-box payload, which knows neither its slide nor its shape).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct DefectLocation {
    /// Persist ID of the slide or notes record whose drawing holds the text.
    pub persist_id: Option<u32>,
    /// `OfficeArt` shape ID (`spid`) of the text box holding the text.
    pub shape_id: Option<u32>,
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Class-specific detail. For
    /// [`TextPropDefect::StyleTextPropCoverage`] this is the zero-based index
    /// of the offending run within its `StyleTextPropAtom` (paragraph runs
    /// first, then character runs; a dropped tail records the index where
    /// another run would have started). The reserved-bit classes store the
    /// raw, unmasked 16-bit flags value.
    pub context: u32,
    /// The slide/notes page and shape the defect was found in.
    pub location: DefectLocation,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the presentation conformed and no substitution was
/// made. After a [`Leniency::Strict`] read the report can still list
/// [`Defect::TextProp`] repairs, which `PowerPoint` itself applies.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the presentation parsed without any repair.
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

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u32) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                context,
                location: DefectLocation::default(),
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order
    /// and each entry's location.
    ///
    /// The stored-entry bound applies to the merged report as a whole;
    /// entries that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            if self.defects.len() < MAX_RECORDED_DEFECTS {
                self.defects.push(*tolerated);
            } else {
                self.unrecorded = self.unrecorded.saturating_add(1);
            }
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }

    /// Fill in location fields that are still unknown on every defect
    /// recorded at or after index `start`.
    ///
    /// Callers capture `start` as `self.defects().len()` before handing the
    /// report to a lower layer and stamp what they know afterwards. Fields a
    /// lower layer already set are kept.
    pub(crate) fn locate_from(&mut self, start: usize, location: DefectLocation) {
        for tolerated in self.defects.iter_mut().skip(start) {
            if tolerated.location.persist_id.is_none() {
                tolerated.location.persist_id = location.persist_id;
            }
            if tolerated.location.shape_id.is_none() {
                tolerated.location.shape_id = location.shape_id;
            }
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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
        report.record(TextPropDefect::StyleTextPropCoverage, 3);
        report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        report.record(TextPropDefect::ReservedWrapFlagBits, 0xFFF8);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::TextProp(TextPropDefect::StyleTextPropCoverage), 3),
                (
                    Defect::TextProp(TextPropDefect::ReservedBulletFlagBits),
                    0x00F0
                ),
                (
                    Defect::TextProp(TextPropDefect::ReservedWrapFlagBits),
                    0xFFF8
                ),
            ]
        );
    }

    /// A hostile style atom must not be able to grow the report without
    /// limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TextPropDefect::StyleTextPropCoverage, index as u32);
        }
        for _ in 0..3 {
            report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TextPropDefect::StyleTextPropCoverage, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TextPropDefect::ReservedWrapFlagBits, 0x00F8);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TextPropDefect::ReservedBulletFlagBits, 0x00F0);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::TextProp(TextPropDefect::StyleTextPropCoverage),
                context: 7,
                location: DefectLocation::default(),
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn locating_fills_only_unknown_fields_of_later_defects() {
        let mut report = ToleranceReport::default();
        report.record(TextPropDefect::StyleTextPropCoverage, 0);
        let start = report.defects().len();
        report.record(TextPropDefect::ReservedBulletFlagBits, 0x00F1);
        report.locate_from(
            start,
            DefectLocation {
                persist_id: None,
                shape_id: Some(1025),
            },
        );
        report.locate_from(
            0,
            DefectLocation {
                persist_id: Some(256),
                shape_id: Some(9999),
            },
        );

        let locations: Vec<_> = report.defects().iter().map(|d| d.location).collect();
        assert_eq!(
            locations,
            vec![
                DefectLocation {
                    persist_id: Some(256),
                    shape_id: Some(9999),
                },
                DefectLocation {
                    persist_id: Some(256),
                    shape_id: Some(1025),
                },
            ]
        );

        // Appending keeps each entry's location.
        let mut merged = ToleranceReport::default();
        merged.append(&report);
        assert_eq!(merged, report);
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(TextPropDefect::StyleTextPropCoverage),
            Defect::from(TextPropDefect::ReservedBulletFlagBits),
            Defect::from(TextPropDefect::ReservedWrapFlagBits),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::TextProp(TextPropDefect::StyleTextPropCoverage).repair(),
            TextPropDefect::StyleTextPropCoverage.repair()
        );
    }

    /// A conformant presentation opens through the default
    /// ([`Leniency::Strict`]) path with an empty report.
    #[test]
    fn a_conformant_presentation_opens_with_a_clean_report() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt");
        let mut package = crate::Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();

        let report = presentation.tolerance_report();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());

        // Reading every slide's and notes page's text boxes repairs nothing.
        let collected = presentation.collect_tolerance_report().unwrap();
        assert!(collected.is_clean());
        for slide in presentation.slides().unwrap() {
            assert!(slide.tolerance_report().unwrap().is_clean());
        }
    }
}
