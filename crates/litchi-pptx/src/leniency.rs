//! Repair reporting for non-structural defects in `PresentationML` packages.
//!
//! Two kinds of repair can happen while a presentation is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: ZIP and XML well-formedness, OPC package
//! grammar, bounded resource limits, and encryption remain hard errors under
//! every variant.
//!
//! **`PowerPoint`-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover markup
//! ECMA-376 leaves unconstrained but real producers emit anyway — markup
//! `PowerPoint` and LibreOffice read past without complaint, so rejecting it
//! would cost the caller the whole slide or timing tree over metadata that
//! cannot change what a conforming deck means. Each substitutes the value
//! documented on the corresponding [`TimingDefect`] variant. Every
//! `PresentationML` repair implemented so far is in this class; the
//! first members are the timing-tree cross-reference repairs of audit finding
//! 6.1 (orphaned animation/build targets and duplicated shape IDs).
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. No `PresentationML` repair currently needs that
//! judgment call, so [`Leniency::Strict`] is the only variant; the enum is
//! where such a gate would appear.
//!
//! # Where the report lives
//!
//! This crate parses lazily: opening a package validates the OPC catalog and
//! the presentation root, and slide payloads — including the `p:timing`
//! animation tree — parse only when a caller asks for them, through `&self`
//! views and free entry points such as [`crate::animations::Sequence`]'s XML
//! parsers. The report is therefore aggregated at two levels:
//!
//! * The owned document handles, [`crate::Package`] and
//!   [`crate::SourceBackedPresentation`], carry the report assembled while
//!   the package was opened and expose it as `tolerance_report()`. No
//!   open-time repair exists yet, so for now these reports are always empty;
//!   open-time recording sites will thread `&mut ToleranceReport` down the
//!   constructor.
//! * A lazy parse carries the repairs it performed itself, the way
//!   `PapBinTable::tolerance_report` lets litchi-doc aggregate part reports
//!   after the fact. [`crate::Slide::animations`] (and the free
//!   [`crate::animations::parse_package_slide`]) returns a
//!   [`crate::animations::Sequence`] whose `tolerance_report()` lists every
//!   finding-6.1 repair, each stamped with the slide's part name
//!   ([`ToleratedDefect::part`]) and the shape ID concerned
//!   ([`ToleratedDefect::context`]). The parse entry points have
//!   `*_with_tolerance` companions taking `&mut ToleranceReport` (explicit
//!   threading, no interior mutability) so a caller can merge per-slide
//!   reports into one deck-level aggregate.
//!
//! # Timing-tree repairs (audit finding 6.1)
//!
//! [`crate::animations::Sequence::parse_slide_xml`] — the read default
//! behind every package-aware entry point — no longer fails the whole slide
//! timing tree over a stale cross-reference:
//!
//! * an animation, interactive trigger, or build-list entry (`p:bldP`,
//!   `p:bldDgm`, `p:bldGraphic`, `p:bldOleChart`) whose `spid` names no shape
//!   on the slide, or a shape unable to host it, is skipped and recorded as
//!   [`TimingDefect::OrphanedTarget`]; the rest of the tree is kept, click
//!   order is renumbered contiguously, and a build group (`grpId`) a skipped
//!   entry carried is left out of the effect/build group-pairing check, whose
//!   pairing the stale reference itself broke. The verbatim timing
//!   subtree is not preserved for a repaired slide, because it still contains
//!   the skipped entries; serialization regenerates from the repaired model;
//! * a `p:cNvPr@id` repeated on the slide resolves to the first shape
//!   carrying it and is recorded as [`TimingDefect::DuplicateShapeId`] once
//!   per repeat; the package-relationship host scan applies the same
//!   first-wins rule.
//!
//! XML well-formedness, the timing-tree grammar (balanced nodes, required
//! attributes, valid token values, duplicate build pairs), the size/depth/
//! count limits, and host-relationship validation of the surviving build
//! entries stay hard errors. Callers that must prove a timing tree is fully
//! self-consistent before a lossless rewrite keep the old contract through
//! [`crate::animations::Sequence::parse_slide_xml_strict`], which rejects
//! any slide whose parse needed a repair.
//!
//! This mirrors the contract `crates/litchi-doc/src/leniency.rs` establishes
//! for legacy Word documents so the readers behave the same way.

use std::sync::Arc;

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous presentation must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural presentation defects.
///
/// Structural defects — malformed ZIP or XML framing, a broken OPC package
/// grammar, unsupported/absent decryption material, or an exceeded resource
/// limit — are hard errors under every variant. The `PowerPoint`-compatible
/// repairs listed by [`TimingDefect`] are applied and recorded under every
/// variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default and currently the only variant: every
    /// `PresentationML` repair on the leniency roadmap is
    /// `PowerPoint`-compatible, so the [`TimingDefect`] repairs are applied
    /// and recorded here too, and no opt-in gate exists yet.
    #[default]
    Strict,
}

/// The class of slide timing-tree defect the reader repairs the way
/// `PowerPoint` repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each one matches what `PowerPoint` and LibreOffice
/// do with markup ECMA-376 does not constrain, so applying it never changes
/// what a conforming presentation means.
///
/// The variants are the recording contract of leniency-audit finding 6.1: the
/// slide timing parser records them, and the repairs above are applied.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TimingDefect {
    /// An animation, an interactive-sequence trigger, or a build-list entry
    /// targeted a shape ID that no shape on the slide carries, or whose shape
    /// cannot host the entry (for example a chart build aimed at a picture).
    ///
    /// ECMA-376-1 §19.5 (`p:spTgt`, `p:bldP`) defines no cross-reference
    /// validity requirement: orphaned timing entries are representable, and
    /// `PowerPoint` routinely leaves them behind when shapes are deleted.
    /// `PowerPoint` and LibreOffice skip the dead entries and play the rest,
    /// so the reader skips the orphaned animation or build entry and keeps
    /// the rest of the timing tree.
    OrphanedTarget,
    /// Two non-visual shape property elements (`p:cNvPr`) on one slide carry
    /// the same shape ID.
    ///
    /// `ST_DrawingElementId` uniqueness is a schema expectation, not a rule
    /// consumers enforce; duplicate IDs survive copy/paste and third-party
    /// round-trips. `PowerPoint` resolves the ID to the first shape carrying
    /// it, so the reader keeps the first occurrence: target resolution sees
    /// the ID once, as the first shape.
    DuplicateShapeId,
}

impl TimingDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::OrphanedTarget => {
                "skipped the orphaned animation or build entry; the rest of the timing tree was kept"
            },
            Self::DuplicateShapeId => {
                "kept the first shape carrying the duplicated ID; later duplicates were ignored for target resolution"
            },
        }
    }
}

/// One class of defect repaired while a presentation was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a read may change about a presentation. [`Defect::Timing`] classes
/// are repaired under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A slide timing-tree defect repaired the way `PowerPoint` repairs it.
    Timing(TimingDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Timing(defect) => defect.repair(),
        }
    }
}

impl From<TimingDefect> for Defect {
    fn from(defect: TimingDefect) -> Self {
        Self::Timing(defect)
    }
}

/// One repaired defect, with enough context to locate it in the package.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found within its part: the shape ID
    /// (`p:cNvPr@id`, `ST_DrawingElementId`) the defect concerns — the
    /// orphaned `spid` for [`TimingDefect::OrphanedTarget`], the duplicated
    /// `id` for [`TimingDefect::DuplicateShapeId`].
    pub context: u32,
    /// The package part the defect was found in, such as
    /// `/ppt/slides/slide3.xml`.
    ///
    /// Package-aware entry points ([`crate::Slide::animations`],
    /// [`crate::animations::parse_package_slide`]) always stamp it. A parse of
    /// bare slide XML ([`crate::animations::Sequence::parse_slide_xml`]) has
    /// no part to name, so it leaves this `None`.
    pub part: Option<Arc<str>>,
}

impl ToleratedDefect {
    /// The package part the defect was found in, when the parse knew it.
    #[inline]
    #[must_use]
    pub fn part(&self) -> Option<&str> {
        self.part.as_deref()
    }
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the presentation conformed and no substitution was
/// made. After a [`Leniency::Strict`] read the report can still list
/// [`Defect::Timing`] repairs, which `PowerPoint` itself applies.
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
        self.push(ToleratedDefect {
            defect: defect.into(),
            context,
            part: None,
        });
    }

    fn push(&mut self, tolerated: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(tolerated);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Stamp `part` on every recorded defect that does not name its part yet.
    pub(crate) fn locate_in_part(&mut self, part: &str) {
        if self
            .defects
            .iter()
            .all(|tolerated| tolerated.part.is_some())
        {
            return;
        }
        let part: Arc<str> = Arc::from(part);
        for tolerated in &mut self.defects {
            if tolerated.part.is_none() {
                tolerated.part = Some(Arc::clone(&part));
            }
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            self.push(tolerated.clone());
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
        report.record(TimingDefect::OrphanedTarget, 7);
        report.record(TimingDefect::DuplicateShapeId, 42);
        report.record(TimingDefect::OrphanedTarget, 9);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Timing(TimingDefect::OrphanedTarget), 7),
                (Defect::Timing(TimingDefect::DuplicateShapeId), 42),
                (Defect::Timing(TimingDefect::OrphanedTarget), 9),
            ]
        );
    }

    /// A hostile timing tree must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TimingDefect::OrphanedTarget, index as u32);
        }
        for _ in 0..3 {
            report.record(TimingDefect::DuplicateShapeId, 2);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TimingDefect::DuplicateShapeId, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TimingDefect::OrphanedTarget, 12);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TimingDefect::OrphanedTarget, 13);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Timing(TimingDefect::DuplicateShapeId),
                context: 7,
                part: None,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(TimingDefect::OrphanedTarget),
            Defect::from(TimingDefect::DuplicateShapeId),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Timing(TimingDefect::OrphanedTarget).repair(),
            TimingDefect::OrphanedTarget.repair()
        );
        assert_eq!(
            Defect::Timing(TimingDefect::DuplicateShapeId).repair(),
            TimingDefect::DuplicateShapeId.repair()
        );
    }
}
