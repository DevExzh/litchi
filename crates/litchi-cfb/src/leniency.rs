//! Repair reporting for non-structural compound-file defects.
//!
//! [MS-CFB] states many of its constraints as producer requirements ("MUST")
//! without prescribing what a reader does when a file breaks one. Word,
//! ole32, LibreOffice, and Apache POI read past a well-known set of such
//! deviations — the residue of interrupted saves, incremental writers, and
//! third-party CFB libraries — because none of them changes which bytes a
//! stream holds. This module is the contract for the same repairs here.
//!
//! **Reference-compatible repairs are applied by default and recorded.** The
//! reader applies the substitution documented on each [`HeaderDefect`],
//! [`AllocationDefect`], and [`DirectoryDefect`] variant under every
//! [`Leniency`] variant, including the default [`Leniency::Strict`], and
//! records it in the [`ToleranceReport`] carried by the opened file
//! ([`crate::OleFile::tolerance_report`],
//! [`crate::SharedOleFile::tolerance_report`]). Nothing is swallowed
//! silently: an empty report means the file conformed, and a conformant file
//! is read exactly as before.
//!
//! **Structural corruption and security limits stay hard errors.** Header
//! framing (signature, byte order, version, sector shifts), sector-chain
//! cycles, out-of-range sector indexes and markers, chains shorter than their
//! declared size, cross-stream sector sharing, directory-tree cycles and
//! repeated ownership, and every configurable resource limit
//! ([`crate::OleFileLimits`]) are rejected under every variant.
//!
//! The defect enumerations are closed and sized to the compound-file findings
//! of the leniency audit (F7.1–F7.7). A variant whose repair has not landed
//! yet says so in its documentation; until then the reader still rejects that
//! deviation and the variant is never recorded.
//!
//! # Merging into a format crate's report
//!
//! Format crates built on this substrate (`litchi-doc`, `litchi-xls`,
//! `litchi-ppt`) own their own closed `Defect` enumerations and bounded
//! reports. To surface container repairs to their callers, a format crate
//! adds one wrapping variant and copies the container report after opening:
//!
//! ```ignore
//! // In the format crate's leniency module:
//! pub enum Defect {
//!     // ...format-specific families...
//!     /// A compound-file repair applied while opening the container.
//!     Container(litchi_cfb::Defect),
//! }
//!
//! // After `OleFile::open` succeeds:
//! let container = ole.tolerance_report();
//! for tolerated in container {
//!     report.record(Defect::Container(tolerated.defect), tolerated.context);
//! }
//! // Carry the container's overflow count into the format report's own.
//! report.count_unrecorded(container.unrecorded());
//! ```
//!
//! The container context is a `u32` (a SID or sector index, see
//! [`ToleratedDefect::context`]); a format crate whose own context is
//! narrower should widen it or keep the container context alongside. Two
//! container reports (for example an outer file and an embedded OLE object)
//! merge directly with [`ToleranceReport::append`], which preserves discovery
//! order and the [`MAX_RECORDED_DEFECTS`] bound.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous compound file must not be able to make the
/// report grow without limit. Defects beyond this bound are counted, not
/// stored; see [`ToleranceReport::unrecorded`].
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural compound-file defects.
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

/// A CFB header field deviation (MS-CFB 2.2).
///
/// The recorded [`ToleratedDefect::context`] is the header field's stored
/// value where one is meaningful (the start sector for the cross-checks),
/// and zero otherwise.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum HeaderDefect {
    /// The six reserved header bytes at offset 0x22 were not zero (F7.6).
    ///
    /// Not yet repaired: the reader currently rejects this deviation. The
    /// repair ignores the bytes, as POI, LibreOffice, and ole32 do.
    ReservedBytes,
    /// The MiniFAT start sector and sector count disagreed (F7.5).
    ///
    /// Not yet repaired: the reader currently rejects this deviation. The
    /// repair treats the count as authoritative, as POI does.
    MiniFatStartCount,
    /// The DIFAT start sector and sector count disagreed (F7.5).
    ///
    /// Not yet repaired: the reader currently rejects this deviation. The
    /// repair treats the count as authoritative, as POI does.
    DifatStartCount,
}

impl HeaderDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedBytes => "ignored the reserved header bytes, which carry no semantics",
            Self::MiniFatStartCount => {
                "trusted the MiniFAT sector count and ignored the disagreeing start sector"
            },
            Self::DifatStartCount => {
                "trusted the DIFAT sector count and ignored the disagreeing start sector"
            },
        }
    }
}

/// A sector-allocation deviation in the FAT or MiniFAT (MS-CFB 2.3–2.6.1).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum AllocationDefect {
    /// A stream's FAT or MiniFAT chain continued past the sectors its
    /// declared size needs (F7.1).
    ///
    /// MS-CFB 2.6.1 makes the stream size authoritative. POI
    /// (`DocumentInputStream`) and LibreOffice (`StgDataStrm`) read exactly
    /// that many bytes and never follow the surplus links, so the reader
    /// terminates the chain after the last sector the size needs; the surplus
    /// sectors are left alone. The context is the stream's directory SID
    /// (SID 0 for the root mini stream).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    ChainLongerThanSize,
    /// A physical sector no validated structure or stream claims carried an
    /// allocation marker instead of `FREESECT` (F7.1).
    ///
    /// This is what surplus chain links and interrupted saves leave behind.
    /// MS-CFB defines no reader behavior for unreferenced sectors and no
    /// reference reader inspects them; the sector is never read. The
    /// context is the physical sector index.
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    UnclaimedAllocatedSector,
    /// A zero-length stream's start sector was not `ENDOFCHAIN` (F7.2).
    ///
    /// A zero-byte read is unaffected by the field; POI and LibreOffice
    /// return an empty stream regardless of it. The reader treats the stream
    /// as empty and never follows the stale start sector. The context is the
    /// stream's directory SID.
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    EmptyStreamStartSector,
}

impl AllocationDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ChainLongerThanSize => {
                "ended the chain at the declared stream size; surplus links were not followed"
            },
            Self::UnclaimedAllocatedSector => {
                "left the unreferenced allocated sector unread; no stream or structure claims it"
            },
            Self::EmptyStreamStartSector => {
                "read the zero-length stream as empty; the stale start sector was not followed"
            },
        }
    }
}

/// A directory-entry or directory-tree deviation (MS-CFB 2.6).
///
/// The recorded [`ToleratedDefect::context`] is the affected directory SID.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum DirectoryDefect {
    /// A sibling tree violated the MS-CFB 2.6.4 name ordering (F7.3).
    ///
    /// The ordering constrains writers; LibreOffice never checks it and POI
    /// looks children up by a linear name scan. Bounds, acyclicity, and
    /// unique ownership are still enforced. Name lookups in the affected
    /// file fall back to an exhaustive, bounded scan of each sibling tree,
    /// returning the first match in tree order.
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    SiblingNameOrder,
    /// A non-empty directory entry was not reachable from the root (F7.4).
    ///
    /// POI and LibreOffice materialize only entries reachable from the root;
    /// orphans are the classic residue of crash-interrupted saves. The entry
    /// is treated as unused (`STGTY_EMPTY`) and its sectors are not claimed.
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    UnreachableEntry,
    /// A directory node color byte was neither red nor black (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    NodeColor,
    /// A storage entry carried a nonzero start sector or size (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    StorageAllocationFields,
    /// A stream entry carried a child link (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    StreamChildLink,
    /// The root entry's name or sibling links deviated from MS-CFB 2.6.2
    /// (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    RootEntryFields,
    /// An unused (`STGTY_EMPTY`) entry carried leftover name bytes (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    EmptyEntryResidue,
    /// A directory name contained a character MS-CFB 2.6.1 forbids (F7.7).
    ///
    /// Not yet repaired: the reader currently rejects this deviation.
    NameCharacters,
}

impl DirectoryDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::SiblingNameOrder => {
                "kept the misordered sibling tree; lookups scan it exhaustively"
            },
            Self::UnreachableEntry => "treated the unreachable entry as unused",
            Self::NodeColor => "ignored the undefined node color; color never affects lookup",
            Self::StorageAllocationFields => {
                "ignored the storage's start sector and size, which storages do not use"
            },
            Self::StreamChildLink => "ignored the stream's child link; streams have no children",
            Self::RootEntryFields => "accepted the root entry by position, as SID 0",
            Self::EmptyEntryResidue => "treated the unused entry as empty despite leftover bytes",
            Self::NameCharacters => "kept the name as an opaque lookup key",
        }
    }
}

/// One class of defect repaired while a compound file was opened.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what a read may change about a compound file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A header field deviation.
    Header(HeaderDefect),
    /// A FAT or MiniFAT allocation deviation.
    Allocation(AllocationDefect),
    /// A directory-entry or directory-tree deviation.
    Directory(DirectoryDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::Header(defect) => defect.repair(),
            Self::Allocation(defect) => defect.repair(),
            Self::Directory(defect) => defect.repair(),
        }
    }
}

impl From<HeaderDefect> for Defect {
    fn from(defect: HeaderDefect) -> Self {
        Self::Header(defect)
    }
}

impl From<AllocationDefect> for Defect {
    fn from(defect: AllocationDefect) -> Self {
        Self::Allocation(defect)
    }
}

impl From<DirectoryDefect> for Defect {
    fn from(defect: DirectoryDefect) -> Self {
        Self::Directory(defect)
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: Defect,
    /// Where the defect was found: a directory SID for stream and directory
    /// defects, a physical sector index for
    /// [`AllocationDefect::UnclaimedAllocatedSector`], and the field value
    /// documented on each [`HeaderDefect`].
    pub context: u32,
}

/// Everything an open repaired, in the order the defects were found.
///
/// An empty report means the compound file conformed and no substitution
/// was made.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the file opened without any repair.
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
        report.record(AllocationDefect::ChainLongerThanSize, 3);
        report.record(DirectoryDefect::UnreachableEntry, 7);
        report.record(HeaderDefect::ReservedBytes, 0);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert!(report.contains(DirectoryDefect::UnreachableEntry));
        assert!(!report.contains(DirectoryDefect::SiblingNameOrder));
        assert_eq!(
            (&report)
                .into_iter()
                .map(|d| (d.defect, d.context))
                .collect::<Vec<_>>(),
            vec![
                (Defect::Allocation(AllocationDefect::ChainLongerThanSize), 3),
                (Defect::Directory(DirectoryDefect::UnreachableEntry), 7),
                (Defect::Header(HeaderDefect::ReservedBytes), 0),
            ]
        );
    }

    /// A hostile file must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for sector in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(
                AllocationDefect::UnclaimedAllocatedSector,
                u32::try_from(sector).unwrap(),
            );
        }
        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 5);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(DirectoryDefect::SiblingNameOrder, 1);

        let mut other = ToleranceReport::default();
        for _ in 0..MAX_RECORDED_DEFECTS {
            other.record(AllocationDefect::UnclaimedAllocatedSector, 9);
        }
        other.unrecorded = 2;

        base.append(&other);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 3);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 3);
        assert_eq!(
            base.defects()[0],
            ToleratedDefect {
                defect: Defect::Directory(DirectoryDefect::SiblingNameOrder),
                context: 1,
            }
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(HeaderDefect::ReservedBytes),
            Defect::from(HeaderDefect::MiniFatStartCount),
            Defect::from(HeaderDefect::DifatStartCount),
            Defect::from(AllocationDefect::ChainLongerThanSize),
            Defect::from(AllocationDefect::UnclaimedAllocatedSector),
            Defect::from(AllocationDefect::EmptyStreamStartSector),
            Defect::from(DirectoryDefect::SiblingNameOrder),
            Defect::from(DirectoryDefect::UnreachableEntry),
            Defect::from(DirectoryDefect::NodeColor),
            Defect::from(DirectoryDefect::StorageAllocationFields),
            Defect::from(DirectoryDefect::StreamChildLink),
            Defect::from(DirectoryDefect::RootEntryFields),
            Defect::from(DirectoryDefect::EmptyEntryResidue),
            Defect::from(DirectoryDefect::NameCharacters),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty(), "{defect:?}");
        }
        assert_eq!(
            Defect::Allocation(AllocationDefect::ChainLongerThanSize).repair(),
            AllocationDefect::ChainLongerThanSize.repair()
        );
    }
}
