//! Repair reporting for non-structural defects in OPC packages.
//!
//! Two kinds of repair can happen while a package is read, and both are
//! *recorded*, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently, and nothing outside
//! the closed enumerations changes: ZIP framing, ZIP encryption, XML
//! well-formedness, manifest grammar outside the listed defects, part-name
//! conflicts, and every [`ReadLimits`](crate::ReadLimits) bound remain hard
//! errors under every variant.
//!
//! **Reference-compatible repairs** are applied under every [`Leniency`]
//! variant, including the default [`Leniency::Strict`]. They cover manifest
//! content ECMA-376 Part 2 constrains only as a *producer* rule — content
//! Word, LibreOffice, and Apache POI read past without complaint, so
//! rejecting it would cost the caller the whole package over bytes that
//! cannot change what a conforming package means. Each substitutes the
//! behavior documented on the corresponding defect variant:
//! [`RelationshipDefect`] for relationship manifests (`_rels/*.rels`) and the
//! relationship graph, [`ContentTypeDefect`] for `[Content_Types].xml`.
//!
//! **Opt-in repairs** trade strict conformance for reach and stay gated by
//! [`Leniency`] variants. There are no opt-in repairs yet:
//! [`Leniency::Strict`] is the only variant, so every open applies the same
//! contract.
//!
//! This mirrors the contract `litchi_doc::leniency` established for legacy
//! binary documents.
//!
//! # Carrier
//!
//! Every package ingress records into one report, which the opened package
//! carries: [`OpcPackage::tolerance_report`](crate::OpcPackage::tolerance_report)
//! and [`SourceBackedPackage::tolerance_report`](crate::SourceBackedPackage::tolerance_report).
//! A source-backed package materialized into an [`OpcPackage`](crate::OpcPackage)
//! hands its report over. A conformant package reports clean, and package
//! edits never add to the report: it describes the bytes that were opened.
//!
//! Format crates (docx, xlsx, pptx) open their package through this crate and
//! own a format-level report of their own. They fold the packaging defects
//! into it by iterating this report (each [`ToleratedDefect`] carries its
//! class and location) — or, where a format keeps the packaging report
//! alongside its own, by [`ToleranceReport::append`]-ing the reports of
//! several packages. [`ToleranceReport::unrecorded`] must be carried over
//! too, so the merged total stays exact.
//!
//! # Status
//!
//! Live for findings 8.5–8.10 of `docs/LENIENCY_AUDIT.md`: duplicate
//! relationship `Id` (8.5), unresolvable internal target (8.6), `TargetMode`
//! casing and value (8.7), missing `Target`/`Type` (8.8), duplicate
//! content-type mappings (8.9), and a relationships part declared with
//! another content type (8.10). The P2 manifest findings 8.11–8.12 remain
//! strict.
//!
//! Source-preserving edits re-read the opened manifests under the same
//! tolerant rules, so a package that opened with recorded repairs stays
//! editable and an untouched manifest still publishes its exact source bytes.

use std::fmt;

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous package must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`]. Each stored entry's strings are bounded
/// by the attribute and part-name read limits that admitted them.
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural package defects.
///
/// Structural defects — ZIP framing, encryption, ill-formed XML, broken
/// manifest grammar, part-name conflicts, and configured read limits — are
/// hard errors under every variant. The reference-compatible repairs listed by
/// [`Defect`] are applied and recorded under every variant, including
/// [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: there are no opt-in repairs to reject, while the
    /// reference-compatible repairs are still applied and recorded in the
    /// [`ToleranceReport`].
    #[default]
    Strict,
}

/// A relationship-manifest or relationship-graph defect repaired the way the
/// reference applications repair it.
///
/// Every variant is a producer rule of ECMA-376 Part 2 §6.5 that defines no
/// consumer behavior. The recorded [`ToleratedDefect::part`] is the
/// relationship *source* (`/` for package relationships),
/// [`ToleratedDefect::item`] the relationship `Id` when one is known, and
/// [`ToleratedDefect::value`] the offending raw value where one exists.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RelationshipDefect {
    /// Two `Relationship` elements in one manifest share an `Id`
    /// (§6.5.3.4 requires uniqueness of the producer).
    ///
    /// The first occurrence is kept and later duplicates are dropped —
    /// LibreOffice resolves by first match, and first-wins is what most
    /// producers' `r:id` references were emitted against. The dropped
    /// relationship's target is recorded as the value.
    DuplicateId,
    /// An internal target cannot be resolved to a part name: a literal space
    /// or backslash, an absolute URI written as an internal target, or a
    /// `..` that escapes the package root (§9.1.1.1 part-name grammar).
    ///
    /// The relationship stays visible on its source, exactly like a dangling
    /// target, but is not followed while the graph is walked (Apache POI
    /// substitutes a dummy URI; Word opens the file with the link dead).
    /// Resolving it later through `target_partname` still reports the error.
    /// The raw target is recorded as the value.
    UnresolvableTarget,
    /// `TargetMode` is `Internal` or `External` in a non-canonical ASCII case
    /// (for example `external`).
    ///
    /// Matched case-insensitively, as Apache POI does. The raw value is
    /// recorded.
    NonCanonicalTargetMode,
    /// `TargetMode` is neither `Internal` nor `External` in any case.
    ///
    /// Read as `External`, Apache POI's semantics: an unknown mode never
    /// causes part resolution, so it cannot bind a bogus target. The raw value
    /// is recorded.
    UnrecognizedTargetMode,
    /// The `Target` attribute is missing or empty.
    ///
    /// The relationship cannot reach anything, so it is skipped: it is not
    /// part of the source's relationship collection (Word keeps it as a dead
    /// link; LibreOffice only requires `Id`). Any recorded `Type` is the
    /// value.
    MissingTarget,
    /// The `Type` attribute is missing or empty.
    ///
    /// A relationship without a type cannot be looked up by any consumer, so
    /// it is skipped. Any recorded `Target` is the value.
    MissingType,
}

impl RelationshipDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::DuplicateId => "kept the first relationship with this Id; dropped the duplicate",
            Self::UnresolvableTarget => {
                "kept the relationship as dangling; its unresolvable internal target was not followed"
            },
            Self::NonCanonicalTargetMode => "matched TargetMode case-insensitively",
            Self::UnrecognizedTargetMode => "read the unrecognized TargetMode as External",
            Self::MissingTarget => "skipped the relationship with a missing or empty Target",
            Self::MissingType => "skipped the relationship with a missing or empty Type",
        }
    }
}

/// A `[Content_Types].xml` defect repaired the way the reference applications
/// repair it.
///
/// The recorded [`ToleratedDefect::part`] is `/[Content_Types].xml`,
/// [`ToleratedDefect::item`] the extension or part name of the mapping
/// concerned, and [`ToleratedDefect::value`] the raw value documented on each
/// variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ContentTypeDefect {
    /// Two `Default` elements declare the same extension (compared ASCII
    /// case-insensitively, §9.1.2.2).
    ///
    /// The first mapping is kept and the later one dropped — LibreOffice's
    /// first-match lookup (Apache POI's map insertion would keep the last).
    DuplicateDefault,
    /// Two `Override` elements declare equivalent part names (compared ASCII
    /// case-insensitively, §9.1.2.2).
    ///
    /// The first mapping is kept and the later one dropped, as for
    /// [`ContentTypeDefect::DuplicateDefault`].
    DuplicateOverride,
    /// A relationships part (`_rels/*.rels`) is mapped to a content type other
    /// than `application/vnd.openxmlformats-package.relationships+xml`
    /// (§9.2 fixes it as a producer rule), for example by
    /// `<Default Extension="rels" ContentType="application/xml"/>`.
    ///
    /// The part is identified by its reserved name, exactly as LibreOffice
    /// and Apache POI identify it, and is typed with the fixed relationships
    /// content type — the same type this reader already assumes when the
    /// mapping is missing. The item is the relationships part name, the
    /// value the declared content type.
    RelationshipPartType,
}

impl ContentTypeDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::DuplicateDefault => {
                "kept the first Default mapping for the extension; dropped the duplicate"
            },
            Self::DuplicateOverride => {
                "kept the first Override mapping for the part name; dropped the duplicate"
            },
            Self::RelationshipPartType => {
                "typed the relationships part with the fixed relationships content type"
            },
        }
    }
}

/// One class of defect repaired while a package was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a package. Every class is repaired under every
/// [`Leniency`] variant; opt-in classes join the union together with the
/// [`Leniency`] variants that gate them.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A relationship-manifest or relationship-graph defect.
    Relationship(RelationshipDefect),
    /// A content-types manifest defect.
    ContentType(ContentTypeDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Relationship(defect) => defect.repair(),
            Self::ContentType(defect) => defect.repair(),
        }
    }
}

impl From<RelationshipDefect> for Defect {
    fn from(defect: RelationshipDefect) -> Self {
        Self::Relationship(defect)
    }
}

impl From<ContentTypeDefect> for Defect {
    fn from(defect: ContentTypeDefect) -> Self {
        Self::ContentType(defect)
    }
}

impl fmt::Display for Defect {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.repair())
    }
}

/// One repaired defect, with enough context to locate it in the package.
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

    /// The part name the defect belongs to: the relationship source part
    /// (`/` for package relationships) or `/[Content_Types].xml`.
    #[inline]
    #[must_use]
    pub fn part(&self) -> &str {
        &self.part
    }

    /// The item within the part: a relationship `Id`, or the extension or
    /// part name of a content-type mapping. `None` when the item has no key
    /// (for example a relationship without an `Id` value is never admitted).
    #[inline]
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.item.as_deref()
    }

    /// The raw value the repair replaced or discarded, as documented on each
    /// defect variant.
    #[inline]
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }
}

/// Everything a package read repaired, in the order the defects were found.
///
/// An empty report means the package conformed and no substitution was made.
/// Content-types manifest defects are found first, then relationship
/// defects in the order the relationship manifests were read, then
/// [`ContentTypeDefect::RelationshipPartType`] entries as the package's
/// members are classified.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the package was read without any repair.
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

    /// Move another report's entries into this one, keeping discovery order.
    pub(crate) fn absorb(&mut self, other: ToleranceReport) {
        if self.is_clean() {
            *self = other;
        } else {
            self.append(&other);
        }
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
        assert!(report.defects().is_empty());
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(
            ContentTypeDefect::DuplicateDefault,
            "/[Content_Types].xml",
            Some("xml"),
            Some("text/xml"),
        );
        report.record(
            RelationshipDefect::DuplicateId,
            "/word/document.xml",
            Some("rId1"),
            None,
        );

        assert!(!report.is_clean());
        assert_eq!(report.total(), 2);
        let seen = (&report)
            .into_iter()
            .map(|d| (d.defect(), d.part(), d.item(), d.value()))
            .collect::<Vec<_>>();
        assert_eq!(
            seen,
            vec![
                (
                    Defect::ContentType(ContentTypeDefect::DuplicateDefault),
                    "/[Content_Types].xml",
                    Some("xml"),
                    Some("text/xml"),
                ),
                (
                    Defect::Relationship(RelationshipDefect::DuplicateId),
                    "/word/document.xml",
                    Some("rId1"),
                    None,
                ),
            ]
        );
    }

    /// A hostile package must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for _ in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(RelationshipDefect::MissingTarget, "/", None, None);
        }
        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 5);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(RelationshipDefect::DuplicateId, "/", Some("rId1"), None);

        let mut parsed = ToleranceReport::default();
        for _ in 0..MAX_RECORDED_DEFECTS {
            parsed.record(RelationshipDefect::MissingType, "/a.xml", None, None);
        }
        parsed.record(RelationshipDefect::MissingType, "/a.xml", None, None);
        assert_eq!(parsed.unrecorded(), 1);

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 2);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 2);
        assert_eq!(
            base.defects()[0].defect(),
            Defect::Relationship(RelationshipDefect::DuplicateId)
        );

        let mut absorbed = ToleranceReport::default();
        absorbed.absorb(parsed.clone());
        assert_eq!(absorbed, parsed);

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(RelationshipDefect::DuplicateId),
            Defect::from(RelationshipDefect::UnresolvableTarget),
            Defect::from(RelationshipDefect::NonCanonicalTargetMode),
            Defect::from(RelationshipDefect::UnrecognizedTargetMode),
            Defect::from(RelationshipDefect::MissingTarget),
            Defect::from(RelationshipDefect::MissingType),
            Defect::from(ContentTypeDefect::DuplicateDefault),
            Defect::from(ContentTypeDefect::DuplicateOverride),
            Defect::from(ContentTypeDefect::RelationshipPartType),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty());
            assert_eq!(defect.to_string(), defect.repair());
        }
    }
}
