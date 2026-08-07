//! Source-bound worksheet snapshots and transactional sparkline edits.

use super::{Error, Groups, Limits, Result, encode_block, parse_block};
use crate::raw::{Records, kind};
use std::ops::Range;

/// An immutable view of the optional sparkline block in one worksheet stream.
///
/// The complete source allocation is retained so records outside the owned
/// block can be reproduced without decoding, normalizing, or copying them.
#[derive(Debug)]
pub struct Snapshot {
    source: Vec<u8>,
    groups: Option<Groups>,
    block: Option<Range<usize>>,
    insertion_anchor: Option<usize>,
    limits: Limits,
}

impl Snapshot {
    /// Read a complete worksheet using safe default sparkline limits.
    pub fn read(source: &[u8]) -> Result<Self> {
        Self::read_with_limits(source, Limits::DEFAULT)
    }

    /// Read a complete worksheet using an explicit sparkline policy.
    pub fn read_with_limits(source: &[u8], limits: Limits) -> Result<Self> {
        limits.validate()?;
        enforce_worksheet_bytes(source.len(), limits)?;
        let mut retained = Vec::new();
        retained
            .try_reserve_exact(source.len())
            .map_err(|_| Error::Allocation {
                resource: "worksheet source",
            })?;
        retained.extend_from_slice(source);
        scan(retained, limits)
    }

    /// The ordered sparkline groups, if the worksheet owns a block.
    #[must_use]
    pub const fn groups(&self) -> Option<&Groups> {
        self.groups.as_ref()
    }

    /// Whether the source worksheet contains a sparkline block.
    #[must_use]
    pub const fn has_block(&self) -> bool {
        self.block.is_some()
    }

    /// Exact complete worksheet bytes.
    #[must_use]
    pub fn source_bytes(&self) -> &[u8] {
        &self.source
    }

    /// Limits used to validate this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Start a detached edit using this snapshot's policy.
    #[must_use]
    pub fn edit(self) -> Edit {
        Edit {
            source: self.source,
            groups: self.groups,
            block: self.block,
            insertion_anchor: self.insertion_anchor,
            limits: self.limits,
        }
    }

    /// Start a detached edit using a replacement policy.
    #[must_use]
    pub fn edit_with_limits(self, limits: Limits) -> Edit {
        Edit {
            source: self.source,
            groups: self.groups,
            block: self.block,
            insertion_anchor: self.insertion_anchor,
            limits,
        }
    }
}

impl PartialEq for Snapshot {
    fn eq(&self, other: &Self) -> bool {
        self.source == other.source
    }
}

impl Eq for Snapshot {}

/// A detached optional-block edit.
#[derive(Debug)]
pub struct Edit {
    source: Vec<u8>,
    groups: Option<Groups>,
    block: Option<Range<usize>>,
    insertion_anchor: Option<usize>,
    limits: Limits,
}

impl Edit {
    /// The currently staged groups.
    #[must_use]
    pub const fn groups(&self) -> Option<&Groups> {
        self.groups.as_ref()
    }

    /// Insert or replace the complete sparkline collection.
    pub fn set(&mut self, groups: Groups) -> &mut Self {
        self.groups = Some(groups);
        self
    }

    /// Remove the complete sparkline collection.
    pub fn remove(&mut self) -> bool {
        self.groups.take().is_some()
    }

    /// Validate and fully encode the staged value before publishing a commit.
    pub fn commit(self) -> Result<Commit> {
        self.limits.validate()?;
        enforce_worksheet_bytes(self.source.len(), self.limits)?;
        let encoded = self
            .groups
            .as_ref()
            .map(|groups| encode_block(groups, self.limits))
            .transpose()?;
        let after = render(
            &self.source,
            self.block.as_ref(),
            self.insertion_anchor,
            encoded.as_deref(),
            self.limits,
        )?;
        let (groups, after) = if let Some(after) = after {
            let snapshot = scan(after, self.limits)?;
            (snapshot.groups, Some(snapshot.source))
        } else {
            (self.groups, None)
        };
        Ok(Commit {
            groups,
            limits: self.limits,
            patch: Patch {
                before: self.source,
                after,
            },
        })
    }
}

/// A structurally validated, unpublished worksheet edit.
///
/// The encoded worksheet bytes remain crate-private because formulas still
/// require workbook-owned name and `Xti` validation. Publish this value only
/// through [`crate::Workbook::apply_sparklines`].
pub struct Commit {
    groups: Option<Groups>,
    limits: Limits,
    patch: Patch,
}

impl Commit {
    /// The staged semantic groups, if the edit retains a sparkline block.
    #[must_use]
    pub const fn groups(&self) -> Option<&Groups> {
        self.groups.as_ref()
    }

    /// Whether publication would leave the worksheet byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.patch.is_empty()
    }

    pub(crate) fn into_publication(self) -> (Option<Groups>, Limits, Patch) {
        (self.groups, self.limits, self.patch)
    }

    #[cfg(test)]
    pub(crate) const fn patch(&self) -> &Patch {
        &self.patch
    }

    #[cfg(test)]
    pub(crate) fn into_patch(self) -> Patch {
        self.patch
    }
}

impl std::fmt::Debug for Commit {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Commit")
            .field("groups", &self.groups())
            .field("is_noop", &self.is_noop())
            .finish()
    }
}

/// A reversible patch guarded by the exact complete worksheet before-image.
#[derive(Debug)]
pub(crate) struct Patch {
    before: Vec<u8>,
    after: Option<Vec<u8>>,
}

impl Patch {
    /// Whether source bytes are unchanged.
    #[must_use]
    pub(crate) fn is_empty(&self) -> bool {
        self.after.is_none()
    }

    /// Exact source after-image.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn after(&self) -> &[u8] {
        self.after.as_deref().unwrap_or(&self.before)
    }

    /// Apply only when `source` exactly matches the complete before-image.
    ///
    /// Test-only materialization remains fallible and byte-exact for no-ops.
    #[cfg(test)]
    pub(crate) fn apply(&self, source: &[u8]) -> Result<Vec<u8>> {
        if source != self.before {
            return Err(invalid(
                "worksheet sparkline patch",
                "source snapshot does not match",
            ));
        }
        fallible_copy(
            self.after.as_deref().unwrap_or(&self.before),
            "worksheet patch",
        )
    }

    /// Exact inverse patch.
    #[must_use]
    #[cfg(test)]
    pub(crate) fn inverse(self) -> Self {
        match self.after {
            Some(after) => Self {
                before: after,
                after: Some(self.before),
            },
            None => Self {
                before: self.before,
                after: None,
            },
        }
    }

    pub(crate) fn apply_owned(self, source: &[u8]) -> Result<(bool, Vec<u8>)> {
        if source != self.before {
            return Err(invalid(
                "worksheet sparkline patch",
                "source snapshot does not match",
            ));
        }
        match self.after {
            Some(after) => Ok((true, after)),
            None => Ok((false, self.before)),
        }
    }
}

/// Read a complete worksheet with default limits.
pub fn read(source: &[u8]) -> Result<Snapshot> {
    Snapshot::read(source)
}

/// Read a complete worksheet with explicit limits.
pub fn read_with_limits(source: &[u8], limits: Limits) -> Result<Snapshot> {
    Snapshot::read_with_limits(source, limits)
}

pub(crate) fn read_owned(source: Vec<u8>, limits: Limits) -> Result<Snapshot> {
    scan(source, limits)
}

fn scan(source: Vec<u8>, limits: Limits) -> Result<Snapshot> {
    limits.validate()?;
    enforce_worksheet_bytes(source.len(), limits)?;
    let mut records = Records::new(&source);
    let mut family = Vec::new();
    family.try_reserve_exact(3).map_err(|_| Error::Allocation {
        resource: "worksheet family stack",
    })?;
    let mut outer_start = None;
    let mut block = None;
    let mut end_sheet = None;
    let mut tail_boundary = None;
    let mut tail_depth = 0usize;
    let mut boundary_top_level_proven = true;
    let mut after_sheet_data = false;

    while let Some(record) = records.next() {
        let record = record?;
        let record_kind = record.kind();
        let start = record.offset();
        let end = records.offset();

        if tail_boundary.is_some() && is_sparkline_family(record_kind) {
            return Err(invalid(
                "sparkline record family",
                "record occurs after the worksheet FRT tail boundary",
            ));
        }

        match record_kind {
            kind::BEGIN_SPARKLINE_GROUPS => {
                if block.is_some() || outer_start.is_some() || !family.is_empty() {
                    return Err(invalid(
                        "BrtBeginSparklineGroups",
                        "duplicate or nested outer collection",
                    ));
                }
                if !after_sheet_data || end_sheet.is_some() {
                    return Err(invalid(
                        "BrtBeginSparklineGroups",
                        "collection is outside the post-cell-table worksheet region",
                    ));
                }
                if !boundary_top_level_proven {
                    return Err(invalid(
                        "BrtBeginSparklineGroups",
                        "collection placement is not provably top-level",
                    ));
                }
                outer_start = Some(start);
                family.push(kind::END_SPARKLINE_GROUPS);
            },
            kind::BEGIN_SPARKLINE_GROUP => {
                expect_parent(
                    &family,
                    kind::END_SPARKLINE_GROUPS,
                    "BrtBeginSparklineGroup",
                )?;
                family.push(kind::END_SPARKLINE_GROUP);
            },
            kind::BEGIN_SPARKLINES => {
                expect_parent(&family, kind::END_SPARKLINE_GROUP, "BrtBeginSparklines")?;
                family.push(kind::END_SPARKLINES);
            },
            kind::SPARKLINE => {
                expect_parent(&family, kind::END_SPARKLINES, "BrtSparkline")?;
            },
            kind::END_SPARKLINES | kind::END_SPARKLINE_GROUP | kind::END_SPARKLINE_GROUPS => {
                let expected = family.pop().ok_or_else(|| {
                    invalid(
                        "sparkline record family",
                        "end record occurs outside the block",
                    )
                })?;
                if expected != record_kind {
                    return Err(invalid(
                        "sparkline record family",
                        format!("expected {expected}, found {record_kind}"),
                    ));
                }
                if record_kind == kind::END_SPARKLINE_GROUPS {
                    let start = outer_start.take().ok_or_else(|| {
                        invalid("BrtEndSparklineGroups", "matching begin is missing")
                    })?;
                    block = Some(start..end);
                }
            },
            kind::END_SHEET_DATA => {
                after_sheet_data = true;
                boundary_top_level_proven = true;
            },
            kind::FRT_BEGIN if after_sheet_data && family.is_empty() && end_sheet.is_none() => {
                if tail_boundary.is_none() {
                    tail_boundary = Some(start);
                }
                tail_depth = tail_depth
                    .checked_add(1)
                    .ok_or_else(|| invalid("worksheet FRT tail", "collection depth overflow"))?;
            },
            kind::FRT_END if tail_boundary.is_some() => {
                tail_depth = tail_depth.checked_sub(1).ok_or_else(|| {
                    invalid("worksheet FRT tail", "BrtFRTEnd has no matching begin")
                })?;
            },
            kind::FRT_END if after_sheet_data => {
                return Err(invalid(
                    "worksheet FRT tail",
                    "BrtFRTEnd occurs before the first tail boundary",
                ));
            },
            kind::END_SHEET => {
                if end_sheet.replace(start).is_some() {
                    return Err(invalid("BrtEndSheet", "duplicate record"));
                }
                if !family.is_empty() {
                    return Err(invalid(
                        "sparkline record family",
                        "collection is not closed before BrtEndSheet",
                    ));
                }
            },
            _ if after_sheet_data
                && tail_boundary.is_none()
                && family.is_empty()
                && end_sheet.is_none() =>
            {
                // There is no complete worksheet-collection stack in the raw
                // kernel. A non-family record between BrtEndSheetData and the
                // first FRT wrapper could open a nested collection, so absent
                // insertion at that boundary must refuse rather than guess.
                boundary_top_level_proven = false;
            },
            _ => {},
        }
    }

    if !family.is_empty() || outer_start.is_some() {
        return Err(invalid(
            "sparkline record family",
            "collection is not closed at end of worksheet",
        ));
    }
    if tail_boundary.is_some() && tail_depth != 0 {
        return Err(invalid(
            "worksheet FRT tail",
            "BrtFRTBegin collection is not closed at end of worksheet",
        ));
    }
    let end_sheet = end_sheet.ok_or_else(|| invalid("worksheet", "BrtEndSheet is missing"))?;
    let insertion_anchor = match tail_boundary {
        Some(anchor) if boundary_top_level_proven => Some(anchor),
        Some(_) => None,
        None => Some(end_sheet),
    };
    let groups = if let Some(span) = &block {
        let bytes = source
            .get(span.clone())
            .ok_or_else(|| invalid("sparkline block", "source span is out of bounds"))?;
        let (groups, consumed) = parse_block(bytes, limits)?;
        if consumed != bytes.len() {
            return Err(invalid(
                "sparkline block",
                "parser did not consume the exact outer collection",
            ));
        }
        Some(groups)
    } else {
        None
    };

    Ok(Snapshot {
        source,
        groups,
        block,
        insertion_anchor,
        limits,
    })
}

fn expect_parent(
    family: &[crate::raw::Kind],
    expected: crate::raw::Kind,
    record: &'static str,
) -> Result<()> {
    if family.last().copied() == Some(expected) {
        Ok(())
    } else {
        Err(invalid(
            record,
            "record occurs outside its required collection",
        ))
    }
}

fn is_sparkline_family(record: crate::raw::Kind) -> bool {
    matches!(
        record,
        kind::BEGIN_SPARKLINE_GROUPS
            | kind::END_SPARKLINE_GROUPS
            | kind::BEGIN_SPARKLINE_GROUP
            | kind::END_SPARKLINE_GROUP
            | kind::BEGIN_SPARKLINES
            | kind::END_SPARKLINES
            | kind::SPARKLINE
    )
}

fn render(
    source: &[u8],
    block: Option<&Range<usize>>,
    insertion_anchor: Option<usize>,
    encoded: Option<&[u8]>,
    limits: Limits,
) -> Result<Option<Vec<u8>>> {
    let replacement = encoded.unwrap_or_default();
    let (start, end) = if let Some(span) = block {
        (span.start, span.end)
    } else if encoded.is_none() {
        return Ok(None);
    } else {
        let anchor = insertion_anchor.ok_or_else(|| {
            invalid(
                "sparkline worksheet render",
                "the first FRT tail boundary is not provably top-level",
            )
        })?;
        (anchor, anchor)
    };
    if source.get(start..end) == Some(replacement) {
        return Ok(None);
    }
    let prefix = source
        .get(..start)
        .ok_or_else(|| invalid("sparkline worksheet render", "prefix span is out of bounds"))?;
    let suffix = source
        .get(end..)
        .ok_or_else(|| invalid("sparkline worksheet render", "suffix span is out of bounds"))?;
    let capacity = prefix
        .len()
        .checked_add(replacement.len())
        .and_then(|length| length.checked_add(suffix.len()))
        .ok_or(Error::Allocation {
            resource: "worksheet render",
        })?;
    enforce_worksheet_bytes(capacity, limits)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|_| Error::Allocation {
            resource: "worksheet render",
        })?;
    output.extend_from_slice(prefix);
    output.extend_from_slice(replacement);
    output.extend_from_slice(suffix);
    Ok(Some(output))
}

fn enforce_worksheet_bytes(actual: usize, limits: Limits) -> Result<()> {
    if actual > limits.worksheet_bytes() {
        Err(Error::Limit {
            resource: "worksheet bytes",
            actual,
            maximum: limits.worksheet_bytes(),
        })
    } else {
        Ok(())
    }
}

#[cfg(test)]
fn fallible_copy(source: &[u8], resource: &'static str) -> Result<Vec<u8>> {
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| Error::Allocation { resource })?;
    output.extend_from_slice(source);
    Ok(output)
}

fn invalid(record: &'static str, reason: impl Into<String>) -> Error {
    Error::Value {
        record,
        reason: reason.into(),
    }
}
