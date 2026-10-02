//! Repair reporting for non-structural defects in WordprocessingML documents.
//!
//! Two kinds of repair can happen while a document is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: OPC packaging, XML well-formedness, part
//! grammar, and decryption remain hard errors under every variant.
//!
//! **Word-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover markup the schema
//! constrains but real producers emit anyway — markup Word and LibreOffice
//! read past without complaint, so rejecting it would cost the caller the
//! whole part over bytes that cannot change what a conforming file means.
//! Each substitutes the behavior documented on the corresponding defect
//! variant. [`TextDefect`] holds this class for semantic text extraction; its
//! first variant is the roadmap contract for the foreign-subtree skip that
//! `extract_word_text` already performs and the typed text codec will adopt.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. There are no opt-in repairs yet: [`Leniency::Strict`]
//! is the only variant, so every read applies the same contract. Judgment-call
//! repairs land behind explicit variants as the roadmap findings that need
//! them are implemented.
//!
//! This mirrors the contract `litchi_doc::leniency` provides for legacy
//! binary documents so the two Word readers behave the same way.
//!
//! # Carrier status
//!
//! The plumbing is in place — [`Package::document`](crate::Package::document)
//! aggregates the report and [`Document::tolerance_report`](crate::Document::tolerance_report)
//! exposes it — and the first recording site is live: the semantic text codec
//! (`paragraph/codec/text.rs`, finding 2.1) skips foreign `*:t` subtrees such
//! as OMML `m:t` instead of failing the extraction and records
//! [`TextDefect::SkippedForeignMarkup`] with the byte offset of the element's
//! start tag within the visible part XML. Because the codec runs lazily under
//! the text-sink facades, each [`Document::write_text_to`](crate::Document::write_text_to)
//! call parses into its own report and merges it into the document-level
//! aggregation with `ToleranceReport::append` — the idiom `litchi-doc`
//! established at open time. The aggregation sits behind a mutex so the sinks
//! keep their `&self` receivers; `tolerance_report` returns a snapshot. The
//! source-backed package facade carries the same surface; the pinned
//! source-backed `Document` and story snapshots extract through the tolerant
//! `extract_word_text` scanner and have no defect to record.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous part must not be able to make the report grow
/// without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural document defects.
///
/// Structural defects — malformed OPC packaging, ill-formed XML, broken part
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant. The Word-compatible repairs listed by [`TextDefect`] are
/// applied and recorded under every variant, including [`Leniency::Strict`].
///
/// No opt-in repairs exist yet, so `Strict` is currently the only variant and
/// the default: every read applies the same contract. Variants gate
/// judgment-call repairs as the roadmap findings that need them land.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: there are no opt-in repairs to reject, while the
    /// Word-compatible [`TextDefect`] repairs are still applied and recorded
    /// in the [`ToleranceReport`].
    #[default]
    Strict,
}

/// The class of semantic text-extraction defect the reader repairs the way
/// Word repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// text-extraction repairs a read applies under every [`Leniency`] variant,
/// including [`Leniency::Strict`]. Each one matches what Word and LibreOffice
/// do with markup no reader-rejection rule covers, so applying it never
/// changes what a conforming document means.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextDefect {
    /// A text-bearing subtree bound to a namespace other than WordprocessingML
    /// appeared where the semantic text codec walks the paragraph: an OMML
    /// `m:t` inside `m:oMath` (ECMA-376 Part 1 §22), or a DrawingML `a:t` in
    /// an inline shape.
    ///
    /// No spec text sanctions refusing the extraction over this: Word,
    /// LibreOffice, and POI all extract the surrounding text, and the crate's
    /// own `extract_word_text` scanner already treats foreign subtrees as
    /// opaque. The typed codec does the same — the subtree contributes no
    /// text and the walk resumes after it — instead of failing the whole
    /// document. The semantic text sink records one defect per foreign `*:t`
    /// element, with the byte offset of the element's start tag within the
    /// visible part XML as the context.
    SkippedForeignMarkup,
}

impl TextDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::SkippedForeignMarkup => {
                "skipped the foreign subtree as opaque; surrounding WordprocessingML text was still extracted"
            },
        }
    }
}

/// One class of defect repaired while a document was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a document. [`Defect::Text`] classes are repaired
/// under every [`Leniency`] variant; opt-in classes join the union together
/// with the [`Leniency`] variants that gate them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A semantic text-extraction defect repaired the way Word repairs it.
    Text(TextDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Text(defect) => defect.repair(),
        }
    }
}

impl From<TextDefect> for Defect {
    fn from(defect: TextDefect) -> Self {
        Self::Text(defect)
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found: the byte offset of the offending element
    /// within the part's XML payload.
    pub context: u64,
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means every read so far conformed and no substitution was
/// made. A [`Leniency::Strict`] report lists only the Word-compatible repairs
/// Word itself applies. The semantic text sink merges each call's repairs
/// into the report, so defects accumulate across calls in call order.
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

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u64) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                context,
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other.defects() {
            self.record(tolerated.defect, tolerated.context);
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
        report.record(TextDefect::SkippedForeignMarkup, 1024);
        report.record(TextDefect::SkippedForeignMarkup, 4096);
        report.record(TextDefect::SkippedForeignMarkup, 8192);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Text(TextDefect::SkippedForeignMarkup), 1024),
                (Defect::Text(TextDefect::SkippedForeignMarkup), 4096),
                (Defect::Text(TextDefect::SkippedForeignMarkup), 8192),
            ]
        );
    }

    /// A hostile part must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(TextDefect::SkippedForeignMarkup, index as u64);
        }
        for _ in 0..3 {
            report.record(TextDefect::SkippedForeignMarkup, 0);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 8);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 8);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(TextDefect::SkippedForeignMarkup, 7);

        let mut parsed = ToleranceReport::default();
        parsed.record(TextDefect::SkippedForeignMarkup, 1024);
        for _ in 1..MAX_RECORDED_DEFECTS {
            parsed.record(TextDefect::SkippedForeignMarkup, 4096);
        }
        assert_eq!(parsed.unrecorded(), 0);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 1);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 1);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Text(TextDefect::SkippedForeignMarkup),
                context: 7,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [Defect::from(TextDefect::SkippedForeignMarkup)];
        for defect in defects {
            assert!(!defect.repair().is_empty());
        }
        assert_eq!(
            Defect::Text(TextDefect::SkippedForeignMarkup).repair(),
            TextDefect::SkippedForeignMarkup.repair()
        );
    }

    /// The carrier is wired through the open path: opening a conformant
    /// fixture yields a document whose report exists and is empty.
    #[test]
    fn conformant_fixture_opens_with_a_clean_report() {
        let fixture = include_bytes!("../../../test-data/ooxml/docx/Hyperlink.docx");
        let package =
            crate::Package::from_reader(std::io::Cursor::new(fixture.as_slice())).unwrap();
        let document = package.document().unwrap();

        let report = document.tolerance_report();
        assert!(report.is_clean());
        assert!(report.defects().is_empty());
        assert_eq!(report.unrecorded(), 0);
        assert_eq!(report.total(), 0);
    }
}
