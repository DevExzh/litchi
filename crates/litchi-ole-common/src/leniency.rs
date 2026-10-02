//! Repair reporting for non-structural OLE Property Set defects.
//!
//! [MS-OLEPS] states its string, padding, and header constraints as producer
//! requirements without prescribing reader recovery. Windows (ole32 over
//! `MultiByteToWideChar`), Word's property pages, and Apache POI read past a
//! known set of such deviations, so rejecting them would cost the caller all
//! document metadata over a byte that cannot change what a conforming
//! property set means. This module is the contract for the same repairs here.
//!
//! **Reference-compatible repairs are applied by default and recorded.** The
//! parser applies the substitution documented on each [`StringDefect`],
//! [`PropertyDefect`], and [`StreamDefect`] variant under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`], and records it in the
//! [`ToleranceReport`] carried by the parsed stream
//! ([`crate::property_set::Stream::tolerance_report`]). Nothing is swallowed
//! silently: an empty report means the stream conformed, and a conformant
//! stream parses exactly as before.
//!
//! **Opt-in repairs** would trade strictness for reach behind explicit
//! [`Leniency`] variants; [`StreamDefect::UnsupportedVersion`] and
//! [`StreamDefect::SectionCount`] are that class and stay rejected for now,
//! since a future property-set version might need new semantics.
//!
//! **Structural corruption and resource limits stay hard errors**: byte order,
//! descriptor and section framing that leaves the stream, offset arithmetic
//! overflow, and every allocation or size ceiling.
//!
//! The defect enumerations are closed and sized to the property-set findings
//! of the leniency audit (F7.8–F7.12). A variant whose repair is not applied
//! says so in its documentation; the parser rejects that deviation and the
//! variant is never recorded.
//!
//! # Merging into a format crate's report
//!
//! Format crates (`litchi-doc`, `litchi-xls`, `litchi-ppt`) read
//! `SummaryInformation` and `DocumentSummaryInformation` through this crate
//! and own their own closed `Defect` enumerations and bounded reports. To
//! surface property-set repairs, a format crate adds one wrapping variant and
//! copies each parsed stream's report:
//!
//! ```ignore
//! // In the format crate's leniency module:
//! pub enum Defect {
//!     // ...format-specific families...
//!     /// A property-set repair applied while reading document metadata.
//!     PropertySet(litchi_ole_common::leniency::Defect),
//! }
//!
//! // After `Stream::parse` (or a `PropertySetReader` call) succeeds:
//! let parsed = stream.tolerance_report();
//! for tolerated in parsed {
//!     report.record(Defect::PropertySet(tolerated.defect), tolerated.context);
//! }
//! // Carry the stream's overflow count into the format report's own.
//! report.count_unrecorded(parsed.unrecorded());
//! ```
//!
//! The same pattern applies to the compound-file container's
//! `litchi_cfb::ToleranceReport`. Reports from several streams of one file
//! merge directly with [`ToleranceReport::append`], which preserves discovery
//! order and the [`MAX_RECORDED_DEFECTS`] bound.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous property set must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the parser treats non-structural property-set defects.
///
/// Every variant applies and records the reference-compatible repairs listed
/// by [`Defect`]. Opt-in variants may be added for repairs that involve a
/// judgment call; structural corruption and resource limits are hard errors
/// under every variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Apply only the default, reference-compatible repairs.
    #[default]
    Strict,
}

/// A property-set string deviation (MS-OLEPS 2.5, 2.15).
///
/// The recorded [`ToleratedDefect::context`] is the property identifier
/// whose value (or dictionary entry) carried the string.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StringDefect {
    /// An ANSI string held a malformed multibyte sequence, or a UTF-16
    /// string an unpaired surrogate (F7.8).
    ///
    /// The string is decoded with U+FFFD substituted for each malformed
    /// sequence, as `MultiByteToWideChar` and Java's replacing decoders do;
    /// one entry is recorded per affected property (PID 0 for the
    /// dictionary). The original bytes are not retained, so the property-set
    /// editor refuses to rewrite such a stream.
    LossyDecode,
    /// A string was not NUL-terminated within its declared size (F7.10).
    ///
    /// The declared bytes are decoded up to the first terminator, or whole
    /// when none is present, as POI does; the size field alone bounds the
    /// string, so nothing is lost and a rewrite adds the terminator. One
    /// entry is recorded per affected property (PID 0 for the dictionary).
    MissingTerminator,
    /// PID 1 named a code page outside the supported table (F7.9).
    ///
    /// The section's strings are decoded as Windows-1252, the default when
    /// PID 1 is absent (LibreOffice likewise keeps its default encoding for
    /// an unknown page). The unrepresentable PID 1 is dropped from the
    /// parsed section, so the property-set editor refuses to rewrite such a
    /// stream. The context is the stored code-page identifier.
    UnsupportedCodePage,
}

impl StringDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::LossyDecode => "decoded the string lossily, substituting U+FFFD",
            Self::MissingTerminator => {
                "decoded the declared bytes; the absent terminator was not required"
            },
            Self::UnsupportedCodePage => "decoded the section's strings as Windows-1252",
        }
    }
}

/// A per-property deviation (MS-OLEPS 2.15, MS-OSHARED composite layouts).
///
/// The recorded [`ToleratedDefect::context`] is the property identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PropertyDefect {
    /// A nested reserved word or trailing padding was not zero (F7.11).
    ///
    /// The bytes are ignored, as the parser already did for the top-level
    /// reserved word; the offset table bounds every value, so they carry
    /// nothing and a rewrite zeroes them.
    NonZeroPadding,
    /// One property's value could not be parsed (F7.11).
    ///
    /// The property (and any dictionary name for it) is dropped so every
    /// other property in the section stays readable, as POI's per-property
    /// resolution and Word's property pages allow. Allocation failures and
    /// resource limits stay fatal. The drop loses content, so the
    /// property-set editor refuses to rewrite such a stream.
    MalformedValue,
}

impl PropertyDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::NonZeroPadding => "ignored the nonzero reserved or padding bytes",
            Self::MalformedValue => {
                "dropped the malformed property; the rest of the section was kept"
            },
        }
    }
}

/// A property-set stream header or section-framing deviation (MS-OLEPS
/// 2.20, 2.21).
///
/// The recorded [`ToleratedDefect::context`] is the stored field value for
/// [`Self::UnsupportedVersion`] and [`Self::SectionCount`], and the section
/// index for [`Self::SectionSizeMismatch`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StreamDefect {
    /// The stream version was neither 0 nor 1 (F7.12, opt-in class).
    ///
    /// Not repaired: the parser rejects this deviation. POI rejects any
    /// nonzero format word outright, and a newer version may carry semantics
    /// this grammar cannot represent, so a repair would be a judgment call
    /// for a future opt-in [`Leniency`] variant.
    UnsupportedVersion,
    /// The section count was outside 1..=2 (F7.12, opt-in class).
    ///
    /// Not repaired: the parser rejects this deviation. Skipping sections
    /// beyond the two the model represents would silently lose content, so
    /// it is left to a future opt-in [`Leniency`] variant.
    SectionCount,
    /// A section's size field overran the range its descriptor implies: the
    /// next section's offset, or the stream end (F7.12).
    ///
    /// Sections are located by offset, as POI and LibreOffice read them; the
    /// size is clamped to the implied range. Duplicate and out-of-range
    /// offsets stay fatal. The context is the section's descriptor index.
    SectionSizeMismatch,
}

impl StreamDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnsupportedVersion => "parsed the stream with version 1 semantics",
            Self::SectionCount => "parsed the sections the grammar supports and skipped the rest",
            Self::SectionSizeMismatch => {
                "trusted the descriptor-implied section range, clamped to the stream"
            },
        }
    }
}

/// One class of defect repaired while a property set was parsed.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a parse may change about a property set.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A string decoding deviation.
    String(StringDefect),
    /// A per-property deviation.
    Property(PropertyDefect),
    /// A stream header or section-framing deviation.
    Stream(StreamDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::String(defect) => defect.repair(),
            Self::Property(defect) => defect.repair(),
            Self::Stream(defect) => defect.repair(),
        }
    }

    /// Whether the repaired value carries exactly the source's content.
    ///
    /// A content-preserving repair only changes framing a re-serialization
    /// writes canonically anyway. A content-changing one (a lossy decode, a
    /// code-page fallback, a dropped property, a skipped section) must not be
    /// written back as if it were the source; the property-set editor
    /// refuses to rewrite a stream that needed one.
    #[must_use]
    pub const fn preserves_content(self) -> bool {
        match self {
            Self::String(StringDefect::MissingTerminator)
            | Self::Property(PropertyDefect::NonZeroPadding)
            | Self::Stream(StreamDefect::SectionSizeMismatch) => true,
            Self::String(StringDefect::LossyDecode | StringDefect::UnsupportedCodePage)
            | Self::Property(PropertyDefect::MalformedValue)
            | Self::Stream(StreamDefect::UnsupportedVersion | StreamDefect::SectionCount) => false,
        }
    }
}

impl From<StringDefect> for Defect {
    fn from(defect: StringDefect) -> Self {
        Self::String(defect)
    }
}

impl From<PropertyDefect> for Defect {
    fn from(defect: PropertyDefect) -> Self {
        Self::Property(defect)
    }
}

impl From<StreamDefect> for Defect {
    fn from(defect: StreamDefect) -> Self {
        Self::Stream(defect)
    }
}

/// One repaired defect, with enough context to locate it in the stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found, as documented on each defect family: a
    /// property identifier, a stored field value, or a section index.
    pub context: u32,
}

/// Everything a parse repaired, in the order the defects were found.
///
/// An empty report means the property set conformed and no substitution was
/// made.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the property set parsed without any repair.
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

    /// Iterates the individually recorded defects, in discovery order.
    #[inline]
    pub fn iter(&self) -> std::slice::Iter<'_, ToleratedDefect> {
        self.defects.iter()
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

    /// Whether every repair kept the source's content (see
    /// [`Defect::preserves_content`]).
    ///
    /// Once defects overflowed the record bound their classes are unknown,
    /// so an overflowed report is conservatively not content-preserving.
    #[must_use]
    pub fn preserves_content(&self) -> bool {
        self.unrecorded == 0
            && self
                .defects
                .iter()
                .all(|tolerated| tolerated.defect.preserves_content())
    }

    /// Whether any recorded defect is of class `defect`.
    #[must_use]
    pub fn contains(&self, defect: impl Into<Defect>) -> bool {
        let defect = defect.into();
        self.defects
            .iter()
            .any(|tolerated| tolerated.defect == defect)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored
    /// entries.
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, context: u32) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                context,
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole;
    /// entries that no longer fit join the unrecorded count.
    pub fn append(&mut self, other: &ToleranceReport) {
        for tolerated in other {
            self.record(tolerated.defect, tolerated.context);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

impl<'a> IntoIterator for &'a ToleranceReport {
    type Item = &'a ToleratedDefect;
    type IntoIter = std::slice::Iter<'a, ToleratedDefect>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
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
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(StringDefect::LossyDecode, 2);
        report.record(PropertyDefect::MalformedValue, 0x0d);
        report.record(StreamDefect::SectionSizeMismatch, 1);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert!(report.contains(PropertyDefect::MalformedValue));
        assert!(!report.contains(PropertyDefect::NonZeroPadding));
        assert_eq!(
            (&report)
                .into_iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::String(StringDefect::LossyDecode), 2),
                (Defect::Property(PropertyDefect::MalformedValue), 0x0d),
                (Defect::Stream(StreamDefect::SectionSizeMismatch), 1),
            ]
        );
    }

    /// A hostile stream must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for pid in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(StringDefect::MissingTerminator, u32::try_from(pid).unwrap());
        }
        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 5);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(StringDefect::UnsupportedCodePage, 28591);

        let mut other = ToleranceReport::default();
        for _ in 0..MAX_RECORDED_DEFECTS {
            other.record(PropertyDefect::NonZeroPadding, 2);
        }
        other.unrecorded = 2;

        base.append(&other);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 3);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 3);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::String(StringDefect::UnsupportedCodePage),
                context: 28591,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(StringDefect::LossyDecode),
            Defect::from(StringDefect::MissingTerminator),
            Defect::from(StringDefect::UnsupportedCodePage),
            Defect::from(PropertyDefect::NonZeroPadding),
            Defect::from(PropertyDefect::MalformedValue),
            Defect::from(StreamDefect::UnsupportedVersion),
            Defect::from(StreamDefect::SectionCount),
            Defect::from(StreamDefect::SectionSizeMismatch),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty(), "{defect:?}");
        }
        assert_eq!(
            Defect::String(StringDefect::LossyDecode).repair(),
            StringDefect::LossyDecode.repair()
        );
    }
}
