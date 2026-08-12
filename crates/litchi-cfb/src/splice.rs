//! Bounded stream-relative splices backed by the validated overlay publisher.

use crate::SharedOleFile;
use crate::consts::{ENDOFCHAIN, MAXREGSECT, STGTY_STREAM};
use crate::file::ParsedOleIndex;
use crate::overlay::{
    OverlayError, OverlayLimits, PhysicalSpan, SourceSnapshot, ValidatedOverlayPlan,
    collect_chain_exact, finish_overlay_plan, path_refs, sector_offset, unavailable,
    validate_and_coalesce_spans,
};
use std::ops::Range;
use std::sync::Arc;

/// One checked, same-length edit within an existing logical CFB stream.
#[derive(Clone, Debug)]
pub struct SameLengthStreamSplice {
    path: Vec<String>,
    offset: u64,
    expected: Arc<[u8]>,
    replacement: Arc<[u8]>,
}

impl SameLengthStreamSplice {
    /// Creates a splice request without copying either byte allocation.
    #[must_use]
    pub fn new(
        path: Vec<String>,
        offset: u64,
        expected: Arc<[u8]>,
        replacement: Arc<[u8]>,
    ) -> Self {
        Self {
            path,
            offset,
            expected,
            replacement,
        }
    }

    /// Selected CFB path.
    #[must_use]
    pub fn path(&self) -> &[String] {
        &self.path
    }

    /// Stream-relative byte offset.
    #[must_use]
    pub const fn offset(&self) -> u64 {
        self.offset
    }

    /// Bytes that must be present in the source range.
    #[must_use]
    pub fn expected(&self) -> &Arc<[u8]> {
        &self.expected
    }

    /// Equal-length bytes to expose in the composed target.
    #[must_use]
    pub fn replacement(&self) -> &Arc<[u8]> {
        &self.replacement
    }
}

/// Finite input and derived-work bounds for stream splice planning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct StreamSpliceLimits {
    max_streams: usize,
    max_splices: usize,
    max_splice_bytes: u64,
    max_physical_spans: usize,
    max_path_bytes: usize,
}

impl StreamSpliceLimits {
    /// Creates a non-zero bounded splice policy.
    ///
    /// `max_splice_bytes` counts each logical range once; expected and
    /// replacement lengths must be equal. `max_path_bytes` counts all UTF-8
    /// path component bytes supplied by the caller.
    pub fn new(
        max_streams: usize,
        max_splices: usize,
        max_splice_bytes: u64,
        max_physical_spans: usize,
        max_path_bytes: usize,
    ) -> Result<Self, OverlayError> {
        if max_streams == 0
            || max_splices == 0
            || max_splice_bytes == 0
            || max_physical_spans == 0
            || max_path_bytes == 0
        {
            return Err(unavailable("stream splice limits must be non-zero"));
        }
        Ok(Self {
            max_streams,
            max_splices,
            max_splice_bytes,
            max_physical_spans,
            max_path_bytes,
        })
    }

    /// Maximum number of distinct logical streams.
    #[must_use]
    pub const fn max_streams(self) -> usize {
        self.max_streams
    }

    /// Maximum number of caller-supplied splices.
    #[must_use]
    pub const fn max_splices(self) -> usize {
        self.max_splices
    }

    /// Maximum aggregate logical splice bytes.
    #[must_use]
    pub const fn max_splice_bytes(self) -> u64 {
        self.max_splice_bytes
    }

    /// Maximum number of mapped physical fragments, including byte no-ops.
    #[must_use]
    pub const fn max_physical_spans(self) -> usize {
        self.max_physical_spans
    }

    /// Maximum aggregate UTF-8 path bytes.
    #[must_use]
    pub const fn max_path_bytes(self) -> usize {
        self.max_path_bytes
    }
}

impl Default for StreamSpliceLimits {
    fn default() -> Self {
        Self {
            max_streams: 64,
            max_splices: 4_096,
            max_splice_bytes: 64 * 1024 * 1024,
            max_physical_spans: 65_536,
            max_path_bytes: 256 * 1024,
        }
    }
}

struct CheckedSplice {
    offset: u64,
    expected: Arc<[u8]>,
    replacement: Arc<[u8]>,
}

impl CheckedSplice {
    fn end(&self) -> Result<u64, OverlayError> {
        self.offset
            .checked_add(self.expected.len() as u64)
            .ok_or_else(|| unavailable("stream splice end overflows u64"))
    }
}

struct StreamSelection {
    path: Vec<String>,
    sid: u32,
    start_sector: u32,
    size: u64,
    is_minifat: bool,
    splices: Vec<CheckedSplice>,
}

impl SharedOleFile {
    /// Plans bounded checked edits within existing CFB streams.
    ///
    /// Every expected range is compared before a target plan is returned.
    /// Ranges must be non-empty, equal-length, in bounds, and non-overlapping
    /// within a stream. The resulting physical spans are passed through the
    /// same full-artifact fingerprinting, CFB reopen, and opt-in publication
    /// pipeline as whole-stream same-length overlays.
    pub fn plan_same_length_stream_splices(
        &self,
        splices: Vec<SameLengthStreamSplice>,
        limits: StreamSpliceLimits,
    ) -> Result<ValidatedOverlayPlan, OverlayError> {
        self.check_source_version()?;
        if splices.len() > limits.max_splices {
            return Err(unavailable(format!(
                "splice count {} exceeds limit {}",
                splices.len(),
                limits.max_splices
            )));
        }

        let mut aggregate_bytes = 0_u64;
        let mut aggregate_path_bytes = 0_usize;
        for splice in &splices {
            if splice.path.is_empty() || splice.path.iter().any(String::is_empty) {
                return Err(unavailable("splice path must contain non-empty names"));
            }
            if splice.expected.is_empty() {
                return Err(unavailable("stream splice range must be non-empty"));
            }
            if splice.expected.len() != splice.replacement.len() {
                return Err(unavailable(format!(
                    "splice for {:?} changes range length from {} to {}",
                    splice.path,
                    splice.expected.len(),
                    splice.replacement.len()
                )));
            }
            aggregate_bytes = aggregate_bytes
                .checked_add(splice.expected.len() as u64)
                .ok_or_else(|| unavailable("aggregate splice bytes overflow u64"))?;
            if aggregate_bytes > limits.max_splice_bytes {
                return Err(unavailable(format!(
                    "aggregate splice bytes {aggregate_bytes} exceed limit {}",
                    limits.max_splice_bytes
                )));
            }
            for component in &splice.path {
                aggregate_path_bytes = aggregate_path_bytes
                    .checked_add(component.len())
                    .ok_or_else(|| unavailable("aggregate splice path bytes overflow usize"))?;
            }
            if aggregate_path_bytes > limits.max_path_bytes {
                return Err(unavailable(format!(
                    "aggregate splice path bytes {aggregate_path_bytes} exceed limit {}",
                    limits.max_path_bytes
                )));
            }
        }

        let source = SourceSnapshot {
            source: Arc::clone(&self.source),
            version: self.expected_version,
            length: self.index.file_size,
        };
        source.ensure_length()?;

        let mut selections: Vec<StreamSelection> = Vec::new();
        selections
            .try_reserve_exact(limits.max_streams.min(splices.len()))
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB splice stream selections",
                source,
            })?;
        for splice in splices {
            let refs = path_refs(&splice.path)?;
            let entry = self.find_entry(&refs)?;
            if entry.entry_type != STGTY_STREAM {
                return Err(unavailable(format!(
                    "splice path {:?} does not identify a stream",
                    splice.path
                )));
            }
            let end = splice
                .offset
                .checked_add(splice.expected.len() as u64)
                .ok_or_else(|| unavailable("stream splice end overflows u64"))?;
            if end > entry.size {
                return Err(unavailable(format!(
                    "splice range {}..{end} exceeds stream {:?} length {}",
                    splice.offset, splice.path, entry.size
                )));
            }
            let checked = CheckedSplice {
                offset: splice.offset,
                expected: splice.expected,
                replacement: splice.replacement,
            };
            if let Some(selection) = selections.iter_mut().find(|item| item.sid == entry.sid) {
                selection
                    .splices
                    .try_reserve(1)
                    .map_err(|source| OverlayError::Allocation {
                        resource: "CFB per-stream splice selections",
                        source,
                    })?;
                selection.splices.push(checked);
            } else {
                if selections.len() == limits.max_streams {
                    return Err(unavailable(format!(
                        "distinct splice stream count exceeds limit {}",
                        limits.max_streams
                    )));
                }
                let mut stream_splices = Vec::new();
                stream_splices
                    .try_reserve_exact(1)
                    .map_err(|source| OverlayError::Allocation {
                        resource: "CFB per-stream splice selections",
                        source,
                    })?;
                stream_splices.push(checked);
                selections.push(StreamSelection {
                    path: splice.path,
                    sid: entry.sid,
                    start_sector: entry.start_sector,
                    size: entry.size,
                    is_minifat: entry.is_minifat,
                    splices: stream_splices,
                });
            }
        }

        for selection in &mut selections {
            selection
                .splices
                .sort_unstable_by_key(|splice| splice.offset);
            for pair in selection.splices.windows(2) {
                if pair[0].end()? > pair[1].offset {
                    return Err(unavailable(format!(
                        "splice ranges overlap in stream {:?}",
                        selection.path
                    )));
                }
            }
        }

        let mut spans = Vec::new();
        let mut mapped_spans = 0_usize;
        let mut comparison = Vec::new();
        comparison
            .try_reserve_exact(self.index.sector_size)
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB splice comparison buffer",
                source,
            })?;
        comparison.resize(self.index.sector_size, 0);
        for selection in &selections {
            derive_stream_splice_spans(
                &source,
                &self.index,
                selection,
                limits.max_physical_spans,
                &mut mapped_spans,
                &mut comparison,
                &mut spans,
            )?;
        }
        spans.sort_unstable_by_key(|span| span.offset);
        let overlay_limits = OverlayLimits::new(
            limits.max_streams,
            limits.max_physical_spans,
            limits.max_splice_bytes,
        )?;
        validate_and_coalesce_spans(&source, overlay_limits, &mut spans)?;

        let verification_bytes = selections
            .iter()
            .flat_map(|selection| &selection.splices)
            .map(|splice| splice.replacement.len())
            .max()
            .unwrap_or(0);
        let mut verification = Vec::new();
        verification
            .try_reserve_exact(verification_bytes)
            .map_err(|source| OverlayError::Allocation {
                resource: "CFB splice candidate verification",
                source,
            })?;
        verification.resize(verification_bytes, 0);

        finish_overlay_plan(source, spans, |_source, candidate| {
            for selection in &selections {
                let refs = path_refs(&selection.path)?;
                if candidate.stream_len(&refs)? != selection.size {
                    return Err(unavailable(format!(
                        "composed stream {:?} changed length after CFB reopen",
                        selection.path
                    )));
                }
                for splice in &selection.splices {
                    let observed = &mut verification[..splice.replacement.len()];
                    candidate.read_stream_range(&refs, splice.offset, observed)?;
                    if observed != splice.replacement.as_ref() {
                        return Err(unavailable(format!(
                            "composed stream {:?} range {}..{} differs after CFB reopen",
                            selection.path,
                            splice.offset,
                            splice.end()?
                        )));
                    }
                    self.read_stream_range(&refs, splice.offset, observed)?;
                    if observed != splice.expected.as_ref() {
                        let mismatch = observed
                            .iter()
                            .zip(splice.expected.iter())
                            .position(|(left, right)| left != right)
                            .unwrap_or(0);
                        return Err(OverlayError::PreconditionFailed {
                            path: selection.path.clone(),
                            offset: splice.offset.checked_add(mismatch as u64).ok_or_else(
                                || unavailable("precondition failure offset overflow"),
                            )?,
                        });
                    }
                }
            }
            Ok(())
        })
    }
}

fn derive_stream_splice_spans(
    source: &SourceSnapshot,
    index: &ParsedOleIndex,
    selection: &StreamSelection,
    maximum: usize,
    mapped_spans: &mut usize,
    comparison: &mut [u8],
    spans: &mut Vec<PhysicalSpan>,
) -> Result<(), OverlayError> {
    let logical_size = usize::try_from(selection.size)
        .map_err(|_| unavailable("stream length does not fit this platform"))?;
    if selection.is_minifat {
        let root = index
            .root
            .as_ref()
            .ok_or_else(|| unavailable("mini stream has no root entry"))?;
        let root_size = usize::try_from(root.size)
            .map_err(|_| unavailable("root mini-stream length does not fit this platform"))?;
        let root_chain = collect_chain_exact(
            &index.fat,
            root.start_sector,
            root_size.div_ceil(index.sector_size),
            "FAT",
        )?;
        let mini_chain = collect_chain_exact(
            &index.minifat,
            selection.start_sector,
            logical_size.div_ceil(index.mini_sector_size),
            "MiniFAT",
        )?;
        for splice in &selection.splices {
            map_splice(
                source,
                selection,
                splice,
                index.mini_sector_size,
                maximum,
                mapped_spans,
                comparison,
                spans,
                |ordinal, within| {
                    let mini_sector = *mini_chain
                        .get(ordinal)
                        .ok_or_else(|| unavailable("splice exceeds MiniFAT chain"))?;
                    let mini_offset = usize::try_from(mini_sector)
                        .map_err(|_| unavailable("mini-sector index does not fit usize"))?
                        .checked_mul(index.mini_sector_size)
                        .and_then(|value| value.checked_add(within))
                        .ok_or_else(|| unavailable("mini-sector offset overflow"))?;
                    if mini_offset >= root_size {
                        return Err(unavailable("splice exceeds root mini stream"));
                    }
                    let root_ordinal = mini_offset / index.sector_size;
                    let root_within = mini_offset % index.sector_size;
                    let root_sector = *root_chain
                        .get(root_ordinal)
                        .ok_or_else(|| unavailable("mini-sector is outside root FAT chain"))?;
                    sector_offset(root_sector, index.sector_size)?
                        .checked_add(root_within as u64)
                        .ok_or_else(|| unavailable("mini-sector physical offset overflow"))
                },
            )?;
        }
    } else {
        // Splices are sorted, so one forward cursor reaches every selected FAT
        // sector without allocating a chain proportional to the whole stream.
        let mut chain = ChainCursor::new(&index.fat, selection.start_sector, "FAT")?;
        for splice in &selection.splices {
            map_splice(
                source,
                selection,
                splice,
                index.sector_size,
                maximum,
                mapped_spans,
                comparison,
                spans,
                |ordinal, within| {
                    let sector = chain.sector_at(ordinal)?;
                    sector_offset(sector, index.sector_size)?
                        .checked_add(within as u64)
                        .ok_or_else(|| unavailable("FAT stream physical offset overflow"))
                },
            )?;
        }
    }
    Ok(())
}

struct ChainCursor<'a> {
    table: &'a [u32],
    sector: u32,
    ordinal: usize,
    name: &'static str,
}

impl<'a> ChainCursor<'a> {
    fn new(table: &'a [u32], start: u32, name: &'static str) -> Result<Self, OverlayError> {
        if start >= MAXREGSECT {
            return Err(unavailable(format!("invalid {name} chain start")));
        }
        Ok(Self {
            table,
            sector: start,
            ordinal: 0,
            name,
        })
    }

    fn sector_at(&mut self, target: usize) -> Result<u32, OverlayError> {
        if target < self.ordinal {
            return Err(unavailable(format!(
                "{} splice ranges are not monotonic",
                self.name
            )));
        }
        while self.ordinal < target {
            let index = usize::try_from(self.sector)
                .map_err(|_| unavailable(format!("{} sector does not fit usize", self.name)))?;
            let next = *self
                .table
                .get(index)
                .ok_or_else(|| unavailable(format!("{} sector is outside its table", self.name)))?;
            if next == ENDOFCHAIN || next >= MAXREGSECT {
                return Err(unavailable(format!(
                    "{} chain ends before a selected splice range",
                    self.name
                )));
            }
            self.sector = next;
            self.ordinal += 1;
        }
        Ok(self.sector)
    }
}

#[allow(clippy::too_many_arguments)]
fn map_splice<F>(
    source: &SourceSnapshot,
    selection: &StreamSelection,
    splice: &CheckedSplice,
    unit_size: usize,
    maximum: usize,
    mapped_spans: &mut usize,
    comparison: &mut [u8],
    spans: &mut Vec<PhysicalSpan>,
    mut physical_offset: F,
) -> Result<(), OverlayError>
where
    F: FnMut(usize, usize) -> Result<u64, OverlayError>,
{
    let mut logical = usize::try_from(splice.offset)
        .map_err(|_| unavailable("splice offset does not fit this platform"))?;
    let mut relative = 0_usize;
    while relative < splice.expected.len() {
        let ordinal = logical / unit_size;
        let within = logical % unit_size;
        let count = (unit_size - within).min(splice.expected.len() - relative);
        *mapped_spans = mapped_spans
            .checked_add(1)
            .ok_or_else(|| unavailable("mapped splice span count overflow"))?;
        if *mapped_spans > maximum {
            return Err(unavailable(format!(
                "mapped physical splice span count exceeds limit {maximum}"
            )));
        }
        let physical = physical_offset(ordinal, within)?;
        source.read_exact(physical, &mut comparison[..count])?;
        let expected = &splice.expected[relative..relative + count];
        if comparison[..count] != *expected {
            let mismatch = comparison[..count]
                .iter()
                .zip(expected)
                .position(|(left, right)| left != right)
                .unwrap_or(0);
            return Err(OverlayError::PreconditionFailed {
                path: selection.path.clone(),
                offset: splice
                    .offset
                    .checked_add((relative + mismatch) as u64)
                    .ok_or_else(|| unavailable("precondition failure offset overflow"))?,
            });
        }
        let replacement_range: Range<usize> = relative..relative + count;
        if splice.replacement[replacement_range.clone()] != *expected {
            spans
                .try_reserve(1)
                .map_err(|source| OverlayError::Allocation {
                    resource: "CFB physical splice spans",
                    source,
                })?;
            spans.push(PhysicalSpan {
                offset: physical,
                replacement: Arc::clone(&splice.replacement),
                replacement_range,
            });
        }
        logical = logical
            .checked_add(count)
            .ok_or_else(|| unavailable("logical splice cursor overflow"))?;
        relative += count;
    }
    Ok(())
}
