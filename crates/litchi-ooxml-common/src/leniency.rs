//! Repair reporting for non-structural defects in shared OOXML metadata parts.
//!
//! The package-level metadata parts this crate reads — core properties
//! (`docProps/core.xml`, ECMA-376 Part 2 §8.3) and custom properties
//! (`docProps/custom.xml`, ECMA-376 Part 1 §22.3) — are non-structural: a
//! document means the same thing with or without them. Word and LibreOffice
//! read them leniently and never refuse a document over one. Rejecting a
//! whole document because one property carries a producer-rule deviation is
//! the worst possible ratio, so the readers here repair such deviations and
//! *record* each repair in a [`ToleranceReport`]. Nothing is swallowed
//! silently: an empty report means the metadata conformed.
//!
//! **Reference-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. Each substitutes the
//! behavior documented on its defect variant: [`CorePropertiesDefect`] and
//! [`CustomPropertiesDefect`]. A repair never invents content — it keeps,
//! ignores, or drops what the producer wrote.
//!
//! **Opt-in repairs** stay gated by [`Leniency`] variants; there are none yet.
//!
//! **Hard errors** remain hard under every variant: XML well-formedness, a
//! missing or wrong document element, DTDs, the Markup Compatibility
//! namespace in core properties (OPC M4.2), and every byte/event/depth limit.
//! Writes keep validating authored values strictly.
//!
//! # Carriers
//!
//! * Core properties: [`properties::read_with_report`](crate::properties::read_with_report)
//!   and [`properties::read_source_backed_with_report`](crate::properties::read_source_backed_with_report)
//!   append into a caller's report; the host [`Slot`](crate::properties::Slot)
//!   keeps the report of its load. There is no report-less core-properties
//!   reader: repaired properties cannot be obtained without their report.
//! * Custom properties: a read [`custom::Props`](crate::custom::Props)
//!   carries its own report ([`custom::Props::tolerance_report`](crate::custom::Props::tolerance_report)).
//!
//! Format crates fold these reports — and the packaging report of
//! `litchi_opc::leniency` — into their document-level report by iterating
//! them (each [`ToleratedDefect`] carries its class and location).
//!
//! # Status
//!
//! Live for findings 8.1 (core-properties body), 8.2 (core-properties graph),
//! 8.3 (custom-properties records), and 8.4 (reserved sensitivity-label
//! properties, advisory on read) of `docs/LENIENCY_AUDIT.md`.

use std::fmt;

/// Upper bound on individually recorded defects.
///
/// Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`]. Each stored string is bounded by the
/// metadata parts' own attribute and text limits.
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the readers treat non-structural metadata defects.
///
/// The reference-compatible repairs listed by [`Defect`] are applied and
/// recorded under every variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: there are no opt-in repairs to reject.
    #[default]
    Strict,
}

/// A core-properties defect (finding 8.1 body rules, finding 8.2 graph rules)
/// repaired the way Word and LibreOffice repair it.
///
/// The recorded [`ToleratedDefect::part`] is the core-properties part (or,
/// for graph defects, the part the defect is about), [`ToleratedDefect::item`]
/// the property's local name or the relationship `Id`, and
/// [`ToleratedDefect::value`] the offending raw value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CorePropertiesDefect {
    /// The part is not valid UTF-8 (for example a Latin-1 producer).
    ///
    /// Decoded lossily — invalid sequences become U+FFFD — and parsed.
    NonUtf8,
    /// The XML namespace (transitional or strict) disagrees with the dialect
    /// of the selecting relationship type.
    ///
    /// The XML's own namespace wins; the relationship type is the value.
    DialectMismatch,
    /// `dcterms:created` or `dcterms:modified` lacks
    /// `xsi:type="dcterms:W3CDTF"` (OPC M4.5, a producer rule).
    ///
    /// The value is read as W3CDTF anyway.
    MissingDateType,
    /// An attribute OPC M4.4/M4.5 forbids or the schema does not define:
    /// `xml:lang` on a property, `xsi:type` on anything but
    /// created/modified (or with another value), or any other non-namespace
    /// attribute.
    ///
    /// Ignored. The attribute's qualified name is the value.
    UnexpectedAttribute,
    /// A property element occurs more than once.
    ///
    /// The first occurrence is kept; later ones are skipped.
    DuplicateProperty,
    /// An element that is not a known core property (including `dcterms:*`
    /// terms other than created/modified, OPC M4.3).
    ///
    /// The element and its subtree are skipped. Its qualified name is the
    /// item.
    UnknownElement,
    /// A date property's value is not a valid W3CDTF / `xsd:dateTime`
    /// lexical.
    ///
    /// The property is dropped; the raw text is the value.
    InvalidValue,
    /// A part carries the core-properties content type but no package
    /// relationship selects it.
    ///
    /// Ignored: the package has no core properties (Word, LibreOffice and POI
    /// locate core.xml only through the relationship).
    OrphanedPart,
    /// The package core-properties relationship targets a part that does not
    /// exist.
    ///
    /// The package has no core properties. The target is the value.
    DanglingRelationship,
    /// More than one part carries the core-properties content type.
    ///
    /// The relationship-selected part is read; the others (recorded as the
    /// part) are dead weight.
    DuplicatePart,
    /// A core-properties relationship is owned by a part instead of the
    /// package.
    ///
    /// Ignored; only the package-level relationship selects core properties.
    MisplacedRelationship,
}

impl CorePropertiesDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::NonUtf8 => "decoded the non-UTF-8 core properties lossily",
            Self::DialectMismatch => {
                "read the core properties in their XML namespace's dialect despite the relationship type"
            },
            Self::MissingDateType => "read the date without xsi:type as W3CDTF",
            Self::UnexpectedAttribute => "ignored the unexpected core-property attribute",
            Self::DuplicateProperty => "kept the first occurrence of the core property",
            Self::UnknownElement => "skipped the unknown core-properties element",
            Self::InvalidValue => "dropped the core property with an invalid date value",
            Self::OrphanedPart => "ignored the core-properties part no relationship selects",
            Self::DanglingRelationship => {
                "read no core properties: the relationship target does not exist"
            },
            Self::DuplicatePart => "ignored an extra core-properties part",
            Self::MisplacedRelationship => "ignored a core-properties relationship owned by a part",
        }
    }
}

/// A custom-properties defect (findings 8.3 and 8.4) repaired the way Word
/// repairs it: per property, never for the whole set.
///
/// The recorded [`ToleratedDefect::part`] is the custom-properties part,
/// [`ToleratedDefect::item`] the property name when it is known, and
/// [`ToleratedDefect::value`] the offending raw value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CustomPropertiesDefect {
    /// The value element is a variant type outside the supported set (for
    /// example `vt:date`, `vt:cy`, `vt:ui4`, `vt:bstr`, `vt:vector`), or not
    /// in the variant-types namespace.
    ///
    /// The property is skipped. The value element's name is the value.
    UnsupportedValueType,
    /// The value's lexical form does not parse for its type (boolean,
    /// RFC 3339 filetime, non-finite float, unknown entity, invalid text).
    ///
    /// The property is skipped; the raw text is the value.
    InvalidValue,
    /// The property lacks a usable `name`, `pid`, or `fmtid`: one is missing,
    /// the PID is not an integer of at least 2, the format ID is not a braced
    /// GUID or is a summary-information format ID, or an unexpected attribute
    /// appears.
    ///
    /// The property is skipped.
    InvalidIdentity,
    /// The property's structure is malformed: no value, more than one value,
    /// attributes on the value element, or nested elements.
    ///
    /// The property is skipped.
    MalformedProperty,
    /// A property reuses an earlier property's PID or (case-insensitively)
    /// name.
    ///
    /// The first property is kept; the later one is skipped.
    DuplicateProperty,
    /// An element other than `property` appears under the root.
    ///
    /// The element and its subtree are skipped.
    UnknownElement,
    /// A processing instruction appears in the part.
    ///
    /// Ignored.
    ProcessingInstruction,
    /// A reserved sensitivity-label property (MS-OI29500 3.11) is scoped to
    /// another host or its value does not have the inferred format.
    ///
    /// The property is skipped on read; writes still validate it strictly.
    ReservedProperty,
}

impl CustomPropertiesDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::UnsupportedValueType => {
                "skipped the custom property with an unsupported value type"
            },
            Self::InvalidValue => "skipped the custom property with an unparseable value",
            Self::InvalidIdentity => "skipped the custom property without a usable name/pid/fmtid",
            Self::MalformedProperty => "skipped the malformed custom property",
            Self::DuplicateProperty => "kept the first custom property; skipped the duplicate",
            Self::UnknownElement => "skipped the unknown custom-properties element",
            Self::ProcessingInstruction => "ignored the processing instruction",
            Self::ReservedProperty => "skipped the invalid reserved sensitivity-label property",
        }
    }
}

/// One class of defect repaired while metadata was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a metadata read may change.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A core-properties defect.
    CoreProperties(CorePropertiesDefect),
    /// A custom-properties defect.
    CustomProperties(CustomPropertiesDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::CoreProperties(defect) => defect.repair(),
            Self::CustomProperties(defect) => defect.repair(),
        }
    }
}

impl From<CorePropertiesDefect> for Defect {
    fn from(defect: CorePropertiesDefect) -> Self {
        Self::CoreProperties(defect)
    }
}

impl From<CustomPropertiesDefect> for Defect {
    fn from(defect: CustomPropertiesDefect) -> Self {
        Self::CustomProperties(defect)
    }
}

impl fmt::Display for Defect {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.repair())
    }
}

/// One repaired defect, with enough context to locate it.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    defect: Defect,
    part: Box<str>,
    item: Option<Box<str>>,
    value: Option<Box<str>>,
}

impl ToleratedDefect {
    /// Which defect class was repaired.
    #[inline]
    #[must_use]
    pub const fn defect(&self) -> Defect {
        self.defect
    }

    /// The part name the defect belongs to.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &str {
        &self.part
    }

    /// The property name, element name, or relationship `Id`, when known.
    #[inline]
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// The raw value the repair ignored or discarded, as documented on each
    /// defect variant.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }
}

/// Everything a metadata read repaired, in discovery order.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the metadata was read without any repair.
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

    /// Iterate the individually recorded defects, in discovery order.
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

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(
        &mut self,
        defect: impl Into<Defect>,
        part: &str,
        item: Option<&str>,
        value: Option<&str>,
    ) {
        if self.defects.len() < MAX_RECORDED_DEFECTS && self.defects.try_reserve(1).is_ok() {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                part: part.into(),
                item: item.map(Into::into),
                value: value.map(Into::into),
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another read's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole; entries
    /// that no longer fit join the unrecorded count.
    pub fn append(&mut self, other: &ToleranceReport) {
        for tolerated in &other.defects {
            if self.defects.len() < MAX_RECORDED_DEFECTS && self.defects.try_reserve(1).is_ok() {
                self.defects.push(tolerated.clone());
            } else {
                self.unrecorded = self.unrecorded.saturating_add(1);
            }
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
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
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
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_keep_discovery_order_and_context() {
        let mut report = ToleranceReport::default();
        report.record(
            CorePropertiesDefect::MissingDateType,
            "/docProps/core.xml",
            Some("created"),
            None,
        );
        report.record(
            CustomPropertiesDefect::UnsupportedValueType,
            "/docProps/custom.xml",
            Some("Due"),
            Some("vt:date"),
        );
        let seen = (&report)
            .into_iter()
            .map(|d| (d.defect(), d.part(), d.item(), d.value()))
            .collect::<Vec<_>>();
        assert_eq!(
            seen,
            vec![
                (
                    Defect::CoreProperties(CorePropertiesDefect::MissingDateType),
                    "/docProps/core.xml",
                    Some("created"),
                    None,
                ),
                (
                    Defect::CustomProperties(CustomPropertiesDefect::UnsupportedValueType),
                    "/docProps/custom.xml",
                    Some("Due"),
                    Some("vt:date"),
                ),
            ]
        );
        assert_eq!(report.total(), 2);
        assert!(!report.is_clean());
    }

    #[test]
    fn defects_past_the_bound_are_counted_and_merged() {
        let mut report = ToleranceReport::default();
        for _ in 0..(MAX_RECORDED_DEFECTS + 3) {
            report.record(CorePropertiesDefect::UnknownElement, "/c.xml", None, None);
        }
        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 3);

        let mut base = ToleranceReport::default();
        base.record(CorePropertiesDefect::NonUtf8, "/c.xml", None, None);
        base.append(&report);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 4);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 4);
    }

    /// The carriers are wired through the read paths: conformant fixtures
    /// with core and custom properties read with clean reports.
    #[test]
    fn conformant_fixtures_read_with_clean_reports() {
        for fixture in [
            "ooxml/docx/Hyperlink.docx",
            "ooxml/pptx/shape-soft-edges.pptx",
            "ooxml/xlsx/ConditionalFormattingSamples.xlsx",
        ] {
            let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data")
                .join(fixture);
            let bytes = std::fs::read(&path).expect("fixture");
            let package = litchi_opc::OpcPackage::from_bytes(&bytes).expect("open");
            assert!(package.tolerance_report().is_clean(), "{fixture}");

            let mut report = ToleranceReport::default();
            let core = crate::properties::read_with_report(&package, &mut report).expect("core");
            assert!(core.is_some(), "{fixture}");
            assert!(report.is_clean(), "{fixture}: {report:?}");
            let slot = crate::properties::Slot::load(&package).expect("slot");
            assert!(slot.tolerance_report().is_clean(), "{fixture}");

            let custom = crate::custom::Props::read(&package).expect("custom");
            assert!(custom.tolerance_report().is_clean(), "{fixture}");
        }
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        use CorePropertiesDefect as C;
        use CustomPropertiesDefect as U;
        let defects = [
            Defect::from(C::NonUtf8),
            C::DialectMismatch.into(),
            C::MissingDateType.into(),
            C::UnexpectedAttribute.into(),
            C::DuplicateProperty.into(),
            C::UnknownElement.into(),
            C::InvalidValue.into(),
            C::OrphanedPart.into(),
            C::DanglingRelationship.into(),
            C::DuplicatePart.into(),
            C::MisplacedRelationship.into(),
            U::UnsupportedValueType.into(),
            U::InvalidValue.into(),
            U::InvalidIdentity.into(),
            U::MalformedProperty.into(),
            U::DuplicateProperty.into(),
            U::UnknownElement.into(),
            U::ProcessingInstruction.into(),
            U::ReservedProperty.into(),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
            assert_eq!(defect.to_string(), defect.repair());
        }
    }
}
