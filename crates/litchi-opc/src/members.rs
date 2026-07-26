//! Classification of physical ZIP items into OPC parts and non-part members.
//!
//! ECMA-376 Part 2 (OPC) §9.1.1 defines what a *part* is: an item with a valid
//! part name and a content type. §10.1.3 defines the mapping between part names
//! and ZIP item names, but a ZIP archive is free to carry items that are not
//! parts at all — directory markers, editor scratch files, and the leftovers of
//! whichever tool last rewrote the archive. Those items are outside the OPC
//! object model, so the conformance rules that govern parts (M1.10 reserved
//! names, M1.11 derived names, M1.12 equivalent names, M1.14 content types) do
//! not reach them and a reader must not reject the package over them.
//!
//! This module owns that boundary: it turns ZIP item names into part names,
//! records the items it declined to promote, and detects the name conflicts that
//! genuinely make a package ambiguous.

use crate::error::{OpcError, Result};
use crate::packuri::{PackURI, PartNameConflict};
use std::collections::HashMap;

/// Longest ZIP item name this reader maps onto an OPC part name.
///
/// ZIP already bounds item names to 65,535 bytes through the `u16` name-length
/// field, so this is a tighter working limit: real OPC part names are two orders
/// of magnitude shorter, and the cap keeps the per-name ancestor walk in
/// [`PartNameIndex`] cheap for hostile central directories.
const MAX_PART_NAME_BYTES: usize = 4_096;

/// Why a ZIP item present in the archive was not loaded as an OPC part.
///
/// Non-part items are reported rather than discarded so that a caller can tell
/// the archive carried something this reader chose not to model.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum NonPartReason {
    /// The ZIP item name does not map onto a valid OPC part name.
    ///
    /// Part names are constrained by ECMA-376 Part 2 §9.1.1.1 to canonical,
    /// segment-wise IRI paths built from RFC 3986 `pchar` characters. Names such
    /// as `[trash]/0000.dat`, which Excel and several third-party writers leave
    /// behind, cannot denote a part, so the item is not one.
    UnmappablePartName,

    /// The item has a usable part name but no declared content type, and no
    /// relationship in the package refers to it.
    ///
    /// ECMA-376 Part 2 §10.1.2.2 requires every *part* to have a content type;
    /// an unreferenced, untyped item is not reachable through the package's
    /// object model and is therefore treated as archive junk rather than as a
    /// non-conforming part. Editor scratch files (`.swp`, `~`), `.DS_Store`, and
    /// directory markers stored without a trailing slash land here.
    UntypedAndUnreferenced,
}

impl NonPartReason {
    /// Return a short, stable description of the reason.
    pub fn as_str(&self) -> &'static str {
        match self {
            Self::UnmappablePartName => "ZIP item name is not a valid OPC part name",
            Self::UntypedAndUnreferenced => {
                "ZIP item has no content type and no relationship refers to it"
            },
        }
    }
}

impl std::fmt::Display for NonPartReason {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// A ZIP item that was present in the archive but not exposed as an OPC part.
///
/// The item's bytes are never decompressed by the reader. A caller that needs
/// them can still reach them through
/// [`PhysPkgReader::archive`](crate::phys_pkg::PhysPkgReader::archive) using
/// [`name`](Self::name), which is the raw ZIP item name.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NonPartMember {
    name: String,
    reason: NonPartReason,
}

impl NonPartMember {
    /// Record a ZIP item that was not promoted to a part.
    pub(crate) fn new(name: &str, reason: NonPartReason) -> Self {
        Self {
            name: name.to_string(),
            reason,
        }
    }

    /// The raw ZIP item name, without a leading slash.
    #[inline]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Why the item was not treated as a part.
    #[inline]
    pub fn reason(&self) -> NonPartReason {
        self.reason
    }
}

/// Map a ZIP item name onto an OPC part name.
///
/// Returns `None` when the item cannot denote a part, which makes it archive
/// junk rather than a malformed package (ECMA-376 Part 2 §9.1.1.1).
pub(crate) fn part_name_for_member(member_name: &str) -> Option<PackURI> {
    if member_name.len() > MAX_PART_NAME_BYTES {
        return None;
    }
    let mut absolute = String::with_capacity(member_name.len() + 1);
    absolute.push('/');
    absolute.push_str(member_name);
    PackURI::new(absolute).ok()
}

/// Collection of accepted part names, checked for OPC name conflicts.
///
/// Enforces the three rules that make a part-name collection unambiguous:
/// duplicates, ASCII-case-equivalent names (M1.12), and derived names where one
/// part name is a folder prefix of another (M1.11). Lookups are hashed rather
/// than pairwise so a central directory with tens of thousands of entries cannot
/// turn package opening into a quadratic scan.
pub(crate) struct PartNameIndex {
    /// Accepted part names in insertion order; other maps index into this.
    names: Vec<PackURI>,
    /// ASCII-folded part name to its index in `names`.
    by_folded: HashMap<String, usize>,
    /// ASCII-folded ancestor path to the index of the first part underneath it.
    descendant_by_folded_ancestor: HashMap<String, usize>,
}

impl PartNameIndex {
    /// Create an index sized for a typical Office package.
    pub(crate) fn with_capacity(capacity: usize) -> Self {
        Self {
            names: Vec::with_capacity(capacity),
            by_folded: HashMap::with_capacity(capacity),
            descendant_by_folded_ancestor: HashMap::with_capacity(capacity),
        }
    }

    /// Accept a part name, rejecting it when it conflicts with one already held.
    pub(crate) fn insert(&mut self, partname: &PackURI) -> Result<()> {
        let folded = partname.as_str().to_ascii_lowercase();

        if let Some(existing) = self
            .by_folded
            .get(&folded)
            .and_then(|at| self.names.get(*at))
        {
            let conflict = if existing.as_str() == partname.as_str() {
                PartNameConflict::Duplicate
            } else {
                PartNameConflict::Equivalent
            };
            return Err(conflict_error(existing, partname, conflict));
        }

        // A part name that is a folder prefix of an accepted name, or that an
        // accepted name is a folder prefix of, is derived (M1.11).
        for (boundary, _) in folded.match_indices('/').skip(1) {
            if let Some(existing) = self
                .by_folded
                .get(&folded[..boundary])
                .and_then(|at| self.names.get(*at))
            {
                return Err(conflict_error(
                    existing,
                    partname,
                    PartNameConflict::Derived,
                ));
            }
        }
        if let Some(existing) = self
            .descendant_by_folded_ancestor
            .get(&folded)
            .and_then(|at| self.names.get(*at))
        {
            return Err(conflict_error(
                existing,
                partname,
                PartNameConflict::Derived,
            ));
        }

        let at = self.names.len();
        self.names.push(partname.clone());
        for (boundary, _) in folded.match_indices('/').skip(1) {
            self.descendant_by_folded_ancestor
                .entry(folded[..boundary].to_string())
                .or_insert(at);
        }
        self.by_folded.insert(folded, at);
        Ok(())
    }
}

/// Build the error for a rejected part-name pair.
fn conflict_error(existing: &PackURI, candidate: &PackURI, conflict: PartNameConflict) -> OpcError {
    match conflict {
        PartNameConflict::Duplicate => OpcError::DuplicatePartName(candidate.to_string()),
        PartNameConflict::Equivalent => OpcError::EquivalentPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
        PartNameConflict::Derived => OpcError::DerivedPartNames {
            existing: existing.to_string(),
            candidate: candidate.to_string(),
        },
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn uri(value: &str) -> PackURI {
        PackURI::new(value).expect("valid part name")
    }

    #[test]
    fn maps_ordinary_members_and_rejects_junk_names() {
        assert_eq!(
            part_name_for_member("word/document.xml").map(|uri| uri.as_str().to_string()),
            Some("/word/document.xml".to_string())
        );
        assert!(part_name_for_member("[trash]/0000.dat").is_none());
        assert!(part_name_for_member("word/my document.xml").is_none());
        assert!(part_name_for_member(&"a".repeat(MAX_PART_NAME_BYTES + 1)).is_none());
    }

    #[test]
    fn accepts_distinct_names_and_rejects_the_three_conflicts() {
        let mut index = PartNameIndex::with_capacity(4);
        index.insert(&uri("/word/document.xml")).unwrap();
        index.insert(&uri("/word/documents.xml")).unwrap();
        index.insert(&uri("/word/media/image1.png")).unwrap();

        assert!(matches!(
            index.insert(&uri("/word/document.xml")),
            Err(OpcError::DuplicatePartName(_))
        ));
        assert!(matches!(
            index.insert(&uri("/WORD/Document.XML")),
            Err(OpcError::EquivalentPartNames { .. })
        ));
        assert!(matches!(
            index.insert(&uri("/word/document.xml/image1.gif")),
            Err(OpcError::DerivedPartNames { .. })
        ));
    }

    #[test]
    fn detects_derived_names_declared_in_either_order() {
        let mut index = PartNameIndex::with_capacity(2);
        index.insert(&uri("/xl/theme/theme1.xml")).unwrap();
        assert!(matches!(
            index.insert(&uri("/xl/THEME")),
            Err(OpcError::DerivedPartNames { .. })
        ));
    }

    #[test]
    fn sibling_folders_sharing_a_prefix_do_not_conflict() {
        let mut index = PartNameIndex::with_capacity(2);
        index.insert(&uri("/xl/worksheets/sheet1.xml")).unwrap();
        index.insert(&uri("/xl/worksheetsExtra.xml")).unwrap();
    }
}
