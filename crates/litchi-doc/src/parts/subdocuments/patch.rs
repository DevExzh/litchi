//! Source-checked table-stream patches for `PlcfWKB` and `SttbFnm`.
//!
//! This module deliberately stops at the table-stream seam. It does not edit
//! FIB pointers, `WordDocument`, or an OLE compound file. A caller that owns
//! those layers can replace the returned table stream and update its pointers
//! using the independently encoded table lengths.

use super::codec::{encode_plcf_wkb, encode_sttb_fnm};
use super::model::Collection;
use crate::package::Result as PackageResult;
use std::fmt;

/// One exact byte range in a caller-provided table stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct TableRange {
    offset: usize,
    length: usize,
}

impl TableRange {
    pub(crate) const fn new(offset: usize, length: usize) -> Self {
        Self { offset, length }
    }

    /// Byte offset of the range in the table stream.
    pub const fn offset(self) -> usize {
        self.offset
    }

    /// Exact byte length of the range.
    pub const fn length(self) -> usize {
        self.length
    }

    /// Exclusive byte end of the range.
    pub fn end(self) -> Option<usize> {
        self.offset.checked_add(self.length)
    }
}

/// The two FIB-addressed ranges owned by this editor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SourceRanges {
    pub(super) referenced_files: Option<TableRange>,
    pub(super) subdocuments: Option<TableRange>,
}

impl SourceRanges {
    /// The `SttbFnm` range, when its FIB length is nonzero.
    pub const fn referenced_files(self) -> Option<TableRange> {
        self.referenced_files
    }

    /// The `PlcfWKB` range, when its FIB length is nonzero.
    pub const fn subdocuments(self) -> Option<TableRange> {
        self.subdocuments
    }
}

/// The immutable stream/context identity captured by a subdocument snapshot.
///
/// The fingerprint covers the complete caller-provided table stream, while
/// the ranges and main-document length capture the structural context needed
/// to interpret the two tables. The fingerprint is a bounded, deterministic
/// integrity check, not a cryptographic signature.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SourceContext {
    main_document_chars: u32,
    table_stream_length: usize,
    fingerprint: Fingerprint,
    ranges: SourceRanges,
    gaps: [Fingerprint; 3],
}

impl SourceContext {
    /// The main-document `ccpText` used to validate the terminal WKB CP.
    pub const fn main_document_chars(self) -> u32 {
        self.main_document_chars
    }

    /// Length of the complete source table stream.
    pub const fn table_stream_length(self) -> usize {
        self.table_stream_length
    }

    /// Exact FIB-addressed ranges captured with the source.
    pub const fn ranges(self) -> SourceRanges {
        self.ranges
    }
}

/// A bounded, reversible replacement of the exact source table slices.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TablePatch {
    before: WireImage,
    after: WireImage,
}

impl TablePatch {
    pub(super) fn new(before: WireImage, after: WireImage) -> Self {
        Self { before, after }
    }

    /// The source context required by this patch.
    pub const fn before_context(&self) -> SourceContext {
        self.before.context
    }

    /// The context produced after this patch is applied.
    pub const fn after_context(&self) -> SourceContext {
        self.after.context
    }

    /// The independently encoded `SttbFnm` replacement bytes.
    pub fn after_referenced_files(&self) -> Option<&[u8]> {
        self.after.referenced_files.as_deref()
    }

    /// The independently encoded `PlcfWKB` replacement bytes.
    pub fn after_subdocuments(&self) -> Option<&[u8]> {
        self.after.subdocuments.as_deref()
    }

    /// The exact inverse table replacement.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Whether both owned table slices are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }

    /// Replace only the two owned table slices in a caller-provided stream.
    ///
    /// The complete source stream fingerprint, exact table ranges, exact table
    /// bytes, and `ccpText` are checked before any allocation is published.
    /// Every byte outside the owned ranges is copied unchanged. The returned
    /// vector may grow or shrink when a string or table count changes. FIB
    /// pointer updates remain the caller's responsibility.
    pub fn apply(&self, table_stream: &[u8], main_document_chars: u32) -> Result<Vec<u8>, Error> {
        self.before
            .context
            .check_stream(table_stream, main_document_chars)?;
        self.before.check_source_slices(table_stream)?;

        let replacements = self.replacements()?;
        let output_len =
            replacements
                .iter()
                .try_fold(table_stream.len(), |length, replacement| {
                    length
                        .checked_sub(replacement.range.length)
                        .and_then(|length| length.checked_add(replacement.after.len()))
                        .ok_or_else(|| {
                            Error::Invalid("patched table stream length overflows".to_string())
                        })
                })?;
        let mut output = Vec::with_capacity(output_len);
        let mut cursor = 0usize;
        for replacement in replacements {
            output.extend_from_slice(
                table_stream
                    .get(cursor..replacement.range.offset)
                    .ok_or_else(|| {
                        Error::Invalid("table patch range is not ordered".to_string())
                    })?,
            );
            output.extend_from_slice(&replacement.after);
            cursor = replacement
                .range
                .end()
                .ok_or_else(|| Error::Invalid("table patch range overflows".to_string()))?;
        }
        output.extend_from_slice(
            table_stream
                .get(cursor..)
                .ok_or_else(|| Error::Invalid("table patch cursor is invalid".to_string()))?,
        );

        self.after
            .context
            .check_stream(&output, self.after.context.main_document_chars)?;
        self.after.check_source_slices(&output)?;
        Ok(output)
    }

    fn replacements(&self) -> Result<Vec<Replacement>, Error> {
        let mut replacements = Vec::with_capacity(2);
        if let (Some(range), Some(_before), Some(after)) = (
            self.before.context.ranges.referenced_files,
            self.before.referenced_files.as_ref(),
            self.after.referenced_files.as_ref(),
        ) {
            replacements.push(Replacement {
                range,
                after: after.clone(),
            });
        }
        if let (Some(range), Some(_before), Some(after)) = (
            self.before.context.ranges.subdocuments,
            self.before.subdocuments.as_ref(),
            self.after.subdocuments.as_ref(),
        ) {
            replacements.push(Replacement {
                range,
                after: after.clone(),
            });
        }
        replacements.sort_by_key(|replacement| replacement.range.offset);
        for pair in replacements.windows(2) {
            let previous_end = pair[0]
                .range
                .end()
                .ok_or_else(|| Error::Invalid("table patch range overflows".to_string()))?;
            if previous_end > pair[1].range.offset {
                return Err(Error::Invalid(
                    "PlcfWKB and SttbFnm source ranges overlap".to_string(),
                ));
            }
        }
        Ok(replacements)
    }
}

/// A checked failure while applying a bounded table patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The caller supplied a stream, range, fingerprint, or `ccpText` that is
    /// not the source captured by this patch.
    SourceConflict(String),
    /// The patch itself cannot be represented safely.
    Invalid(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::SourceConflict(message) => {
                write!(formatter, "subdocument patch conflict: {message}")
            },
            Self::Invalid(message) => {
                write!(formatter, "invalid subdocument table patch: {message}")
            },
        }
    }
}

impl std::error::Error for Error {}

#[derive(Debug, Clone)]
struct Replacement {
    range: TableRange,
    after: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct WireImage {
    pub(super) context: SourceContext,
    pub(super) referenced_files: Option<Vec<u8>>,
    pub(super) subdocuments: Option<Vec<u8>>,
}

impl WireImage {
    pub(super) fn capture(
        table_stream: &[u8],
        main_document_chars: u32,
        ranges: SourceRanges,
        referenced_files: Option<&[u8]>,
        subdocuments: Option<&[u8]>,
    ) -> Result<Self, Error> {
        let context = SourceContext::capture(table_stream, main_document_chars, ranges)?;
        Ok(Self {
            context,
            referenced_files: referenced_files.map(ToOwned::to_owned),
            subdocuments: subdocuments.map(ToOwned::to_owned),
        })
    }

    pub(super) fn reencode(
        source: &Self,
        collection: &Collection,
        main_document_chars: u32,
    ) -> PackageResult<Self> {
        let referenced_files = if source.referenced_files.is_some() {
            Some(encode_sttb_fnm(&collection.referenced_files)?)
        } else if collection.referenced_files.is_empty() {
            None
        } else {
            return Err(crate::package::Error::Corrupted(
                "SttbFnm is absent, so a referenced file cannot be encoded".to_string(),
            ));
        };
        let subdocuments = if source.subdocuments.is_some() {
            Some(encode_plcf_wkb(
                &collection.subdocuments,
                main_document_chars,
                &collection.referenced_files,
            )?)
        } else if collection.subdocuments.is_empty() {
            None
        } else {
            return Err(crate::package::Error::Corrupted(
                "PlcfWKB is absent, so a subdocument reference cannot be encoded".to_string(),
            ));
        };
        let ranges = shifted_ranges(
            source.context.ranges,
            source.referenced_files.as_deref(),
            source.subdocuments.as_deref(),
            referenced_files.as_deref(),
            subdocuments.as_deref(),
        )
        .map_err(|message| crate::package::Error::Corrupted(message.to_string()))?;
        let layout = table_stream_layout(source, &referenced_files, &subdocuments)
            .map_err(|message| crate::package::Error::Corrupted(message.to_string()))?;
        let context = source
            .context
            .after(layout, ranges, main_document_chars)
            .map_err(|message| crate::package::Error::Corrupted(message.to_string()))?;
        Ok(Self {
            context,
            referenced_files,
            subdocuments,
        })
    }

    fn check_source_slices(&self, table_stream: &[u8]) -> Result<(), Error> {
        for (range, expected, name) in [
            (
                self.context.ranges.referenced_files,
                self.referenced_files.as_deref(),
                "SttbFnm",
            ),
            (
                self.context.ranges.subdocuments,
                self.subdocuments.as_deref(),
                "PlcfWKB",
            ),
        ] {
            match (range, expected) {
                (Some(range), Some(expected)) => {
                    let actual = table_stream.get(
                        range.offset..range.end().ok_or_else(|| {
                            Error::Invalid(format!("{name} source range overflows"))
                        })?,
                    );
                    if actual != Some(expected) {
                        return Err(Error::SourceConflict(format!("{name} source bytes differ")));
                    }
                },
                (None, None) => {},
                _ => {
                    return Err(Error::SourceConflict(format!(
                        "{name} source presence differs"
                    )));
                },
            }
        }
        Ok(())
    }
}

impl SourceContext {
    fn capture(
        table_stream: &[u8],
        main_document_chars: u32,
        ranges: SourceRanges,
    ) -> Result<Self, Error> {
        let ordered = ordered_ranges(ranges)?;
        for pair in ordered.windows(2) {
            let previous_end = pair[0]
                .1
                .end()
                .ok_or_else(|| Error::Invalid("table source range overflows".to_string()))?;
            if previous_end > pair[1].1.offset {
                return Err(Error::Invalid(
                    "PlcfWKB and SttbFnm source ranges overlap".to_string(),
                ));
            }
        }
        for (_, range) in &ordered {
            let end = range
                .end()
                .ok_or_else(|| Error::Invalid("table source range overflows".to_string()))?;
            if end > table_stream.len() {
                return Err(Error::Invalid(
                    "table source range exceeds the table stream".to_string(),
                ));
            }
        }
        let gaps = gap_fingerprints(table_stream, &ordered)?;
        Ok(Self {
            main_document_chars,
            table_stream_length: table_stream.len(),
            fingerprint: fingerprint(table_stream),
            ranges,
            gaps,
        })
    }

    fn after(
        &self,
        layout: Vec<u8>,
        ranges: SourceRanges,
        main_document_chars: u32,
    ) -> Result<Self, String> {
        let ordered = ordered_ranges(ranges).map_err(|error| error.to_string())?;
        let gaps = self.gaps;
        let fingerprint =
            compose_fingerprint(&gaps, &ordered, &layout).map_err(|error| error.to_string())?;
        let table_stream_length = fingerprint.length;
        Ok(Self {
            main_document_chars,
            table_stream_length,
            fingerprint,
            ranges,
            gaps,
        })
    }

    fn check_stream(self, table_stream: &[u8], main_document_chars: u32) -> Result<(), Error> {
        if self.main_document_chars != main_document_chars {
            return Err(Error::SourceConflict(
                "main-document character count differs".to_string(),
            ));
        }
        if self.table_stream_length != table_stream.len()
            || self.fingerprint != fingerprint(table_stream)
        {
            return Err(Error::SourceConflict(
                "table-stream context differs".to_string(),
            ));
        }
        Ok(())
    }
}

fn ordered_ranges(ranges: SourceRanges) -> Result<Vec<(&'static str, TableRange)>, Error> {
    let mut ordered = Vec::with_capacity(2);
    if let Some(range) = ranges.referenced_files {
        ordered.push(("SttbFnm", range));
    }
    if let Some(range) = ranges.subdocuments {
        ordered.push(("PlcfWKB", range));
    }
    ordered.sort_by_key(|(_, range)| range.offset);
    Ok(ordered)
}

fn gap_fingerprints(
    table_stream: &[u8],
    ordered: &[(&'static str, TableRange)],
) -> Result<[Fingerprint; 3], Error> {
    let mut gaps = [Fingerprint::empty(); 3];
    let mut cursor = 0usize;
    for (index, (_, range)) in ordered.iter().enumerate() {
        let end = range
            .end()
            .ok_or_else(|| Error::Invalid("table source range overflows".to_string()))?;
        if index >= gaps.len() || range.offset < cursor {
            return Err(Error::Invalid(
                "table source ranges are not disjoint".to_string(),
            ));
        }
        gaps[index] = fingerprint(&table_stream[cursor..range.offset]);
        cursor = end;
    }
    if ordered.len() >= gaps.len() {
        return Err(Error::Invalid("too many owned table ranges".to_string()));
    }
    gaps[ordered.len()] = fingerprint(&table_stream[cursor..]);
    Ok(gaps)
}

fn table_stream_layout(
    source: &WireImage,
    referenced_files: &Option<Vec<u8>>,
    subdocuments: &Option<Vec<u8>>,
) -> Result<Vec<u8>, &'static str> {
    let ordered = ordered_ranges(source.context.ranges).map_err(|_| "invalid table ranges")?;
    let mut values = Vec::with_capacity(2);
    for (_, range) in ordered {
        if Some(range) == source.context.ranges.referenced_files {
            values.push(
                referenced_files
                    .as_deref()
                    .ok_or("SttbFnm presence changed")?,
            );
        } else if Some(range) == source.context.ranges.subdocuments {
            values.push(subdocuments.as_deref().ok_or("PlcfWKB presence changed")?);
        }
    }
    let mut layout = Vec::with_capacity(values.iter().map(|value| value.len()).sum());
    for value in values {
        layout.extend_from_slice(value);
    }
    Ok(layout)
}

fn shifted_ranges(
    before: SourceRanges,
    before_referenced_files: Option<&[u8]>,
    before_subdocuments: Option<&[u8]>,
    after_referenced_files: Option<&[u8]>,
    after_subdocuments: Option<&[u8]>,
) -> Result<SourceRanges, &'static str> {
    let ordered = ordered_ranges(before).map_err(|_| "invalid table ranges")?;
    let mut delta: isize = 0;
    let mut result = SourceRanges::default();
    for (_, range) in ordered {
        let (before_len, after_len) = if Some(range) == before.referenced_files {
            (
                before_referenced_files
                    .ok_or("SttbFnm source presence changed")?
                    .len(),
                after_referenced_files
                    .ok_or("SttbFnm result presence changed")?
                    .len(),
            )
        } else if Some(range) == before.subdocuments {
            (
                before_subdocuments
                    .ok_or("PlcfWKB source presence changed")?
                    .len(),
                after_subdocuments
                    .ok_or("PlcfWKB result presence changed")?
                    .len(),
            )
        } else {
            return Err("unknown table range");
        };
        let offset = add_delta(range.offset, delta)?;
        let after_range = TableRange::new(offset, after_len);
        if Some(range) == before.referenced_files {
            result.referenced_files = Some(after_range);
        } else {
            result.subdocuments = Some(after_range);
        }
        let difference = after_len as isize - before_len as isize;
        delta = delta
            .checked_add(difference)
            .ok_or("table range delta overflows")?;
    }
    Ok(result)
}

fn add_delta(value: usize, delta: isize) -> Result<usize, &'static str> {
    if delta >= 0 {
        value
            .checked_add(delta as usize)
            .ok_or("table range offset overflows")
    } else {
        value
            .checked_sub(delta.unsigned_abs())
            .ok_or("table range offset underflows")
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Fingerprint {
    first: u64,
    second: u64,
    length: usize,
}

impl Fingerprint {
    const fn empty() -> Self {
        Self {
            first: 0,
            second: 0,
            length: 0,
        }
    }
}

const HASH_BASE_A: u64 = 0x100000001B3;
const HASH_BASE_B: u64 = 0x9E3779B185EBCA87;

fn fingerprint(data: &[u8]) -> Fingerprint {
    let mut result = Fingerprint::empty();
    for byte in data {
        result.first = result
            .first
            .wrapping_mul(HASH_BASE_A)
            .wrapping_add(u64::from(*byte) + 1);
        result.second = result
            .second
            .wrapping_mul(HASH_BASE_B)
            .wrapping_add(u64::from(*byte) + 1);
        result.length += 1;
    }
    result
}

fn compose_fingerprint(
    gaps: &[Fingerprint; 3],
    ordered: &[(&'static str, TableRange)],
    layout: &[u8],
) -> Result<Fingerprint, Error> {
    let mut result = Fingerprint::empty();
    let mut layout_offset = 0usize;
    for (index, (_, range)) in ordered.iter().enumerate() {
        result = concat(result, gaps[index])?;
        let table_length = range.length;
        let table = layout
            .get(layout_offset..layout_offset + table_length)
            .ok_or_else(|| Error::Invalid("encoded table layout is inconsistent".to_string()))?;
        result = concat(result, fingerprint(table))?;
        layout_offset = layout_offset
            .checked_add(table_length)
            .ok_or_else(|| Error::Invalid("encoded table layout overflows".to_string()))?;
    }
    concat(result, gaps[ordered.len()])
}

fn concat(left: Fingerprint, right: Fingerprint) -> Result<Fingerprint, Error> {
    let pow_a = power(HASH_BASE_A, right.length);
    let pow_b = power(HASH_BASE_B, right.length);
    Ok(Fingerprint {
        first: left.first.wrapping_mul(pow_a).wrapping_add(right.first),
        second: left.second.wrapping_mul(pow_b).wrapping_add(right.second),
        length: left
            .length
            .checked_add(right.length)
            .ok_or_else(|| Error::Invalid("fingerprint length overflows".to_string()))?,
    })
}

fn power(base: u64, mut exponent: usize) -> u64 {
    let mut result = 1u64;
    let mut factor = base;
    while exponent != 0 {
        if exponent & 1 != 0 {
            result = result.wrapping_mul(factor);
        }
        factor = factor.wrapping_mul(factor);
        exponent >>= 1;
    }
    result
}
