//! Sparse-table allocation planning for the exact table-cell writer.
//!
//! This module deliberately knows only the four embedded `TST.DataStore`
//! messages which change when a write crosses a row-strip boundary.  It does
//! not own an archive or allocate object identifiers: callers first obtain a
//! [`SparsePlan`], reserve the returned [`NewObjectRequest`]s in their single
//! package-metadata transaction, and then pass those identifiers to
//! [`rewrite_data_store`].  Keeping that split is what makes an allocation
//! failure atomic with respect to package publication.

use core::{fmt, mem::size_of};

#[cfg(test)]
use std::cell::Cell as TestCell;

#[cfg(test)]
thread_local! {
    static PREPARED_HEADER_EXECUTION_ALLOCATIONS: TestCell<usize> = const { TestCell::new(0) };
}

#[cfg(test)]
fn reset_prepared_header_execution_allocations() {
    PREPARED_HEADER_EXECUTION_ALLOCATIONS.set(0);
}

#[cfg(test)]
fn prepared_header_execution_allocations() -> usize {
    PREPARED_HEADER_EXECUTION_ALLOCATIONS.get()
}

#[cfg(test)]
fn record_prepared_header_execution_allocation() {
    PREPARED_HEADER_EXECUTION_ALLOCATIONS.set(PREPARED_HEADER_EXECUTION_ALLOCATIONS.get() + 1);
}

/// Logical rows represented by one positional row-header bucket.
pub(super) const HEADER_BUCKET_ROWS: u32 = 65_536;

/// A source cell whose row may require a sparse tile and row header.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
pub(super) struct Cell {
    /// Zero-based table row.
    pub row: u32,
    /// Zero-based table column.
    pub column: u32,
}

/// One externally owned header-bucket object.
///
/// `payload` is the exact uncompressed IWA message data, not an archive
/// envelope.  Sources must be in the same order as the references in
/// `DataStore.rowHeaders`; the plan rejects a mismatch rather than guessing.
#[derive(Debug, Clone, Copy)]
pub(super) struct HeaderBucketSource<'a> {
    /// Identifier named by the `HeaderStorage` reference.
    pub object_id: u64,
    /// Exact `TST.HeaderStorageBucket` payload.
    pub payload: &'a [u8],
}

/// Bounded planning input.  `cells` must be strictly sorted by `(row,column)`
/// and contain no duplicates; the transaction's compact mutation buffer has
/// already established that invariant.
#[derive(Debug, Clone, Copy)]
pub(super) struct SparseRequest<'a> {
    /// Existing embedded `TST.DataStore` message bytes.
    pub data_store: &'a [u8],
    /// Sorted distinct changed coordinates.
    pub cells: &'a [Cell],
    /// Number of logical table columns, used by newly materialised headers.
    pub columns: u32,
    /// Row-strip size supplied by the resolved native tile storage.
    pub tile_size: u32,
    /// Payloads for all existing row-header bucket references.
    pub row_header_buckets: &'a [HeaderBucketSource<'a>],
}

/// Finite limits used before retaining plan-owned vectors or encoded output.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct SparseLimits {
    /// Maximum source protobuf fields parsed in one embedded message.
    pub max_fields: usize,
    /// Maximum changed cells accepted by this phase.
    pub max_cells: usize,
    /// Maximum new tile/header/object records retained by the plan.
    pub max_records: usize,
    /// Maximum output bytes for one rewritten embedded `DataStore`.
    pub max_output_bytes: usize,
    /// Maximum aggregate logical work reported by one sparse leaf call.
    pub max_work: usize,
    /// Maximum elements retained by one sparse leaf result.
    pub max_retained_elements: usize,
    /// Maximum bytes retained by one sparse leaf result.
    pub max_retained_bytes: usize,
    /// Maximum logical scratch extent used by one sparse leaf call.
    pub max_scratch_bytes: usize,
    /// Maximum fallible allocation sites entered by one sparse leaf call.
    pub max_allocation_events: usize,
    /// Maximum decoded or newly introduced object references.
    pub max_references: usize,
}

impl Default for SparseLimits {
    fn default() -> Self {
        Self {
            max_fields: 1_000_000,
            max_cells: 1_000_000,
            max_records: 1_000_000,
            max_output_bytes: 64 * 1024 * 1024,
            max_work: 256 * 1024 * 1024,
            max_retained_elements: 2_000_000,
            max_retained_bytes: 128 * 1024 * 1024,
            max_scratch_bytes: 128 * 1024 * 1024,
            max_allocation_events: 4_000_000,
            max_references: 1_000_000,
        }
    }
}

/// Exact logical resource evidence for one or more sparse leaf calls.
///
/// Bytes and fields count every caller-visible source or result consumed by
/// the leaf. `work` is the checked sum of those byte/field observations and
/// every retained record/reference visit. Retained and scratch byte counts
/// use Rust representation sizes, making them stable across allocator growth
/// strategies. Reports can therefore be accumulated before the grouped
/// writer is authorized without consulting allocator-specific capacities.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct SparseReport {
    pub(super) input_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    /// Existing local reference items decoded from sparse storage.
    pub(super) reference_reads: usize,
    /// Newly assigned local reference items published into sparse storage.
    pub(super) reference_writes: usize,
    /// `reference_reads + reference_writes`, retained for aggregate budgets.
    pub(super) references: usize,
    pub(super) retained_elements: usize,
    pub(super) retained_bytes: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocation_events: usize,
    pub(super) records: usize,
    /// Existing row-header records inspected.
    pub(super) header_reads: usize,
    /// Newly encoded row-header records.
    pub(super) header_writes: usize,
    /// `header_reads + header_writes`, retained for compatibility.
    pub(super) headers: usize,
    pub(super) objects: usize,
}

impl SparseReport {
    /// Checked cumulative accounting used by the transaction owner.
    pub(super) fn merge(&mut self, other: Self) -> Result<(), SparseError> {
        macro_rules! add {
            ($field:ident) => {
                self.$field = self
                    .$field
                    .checked_add(other.$field)
                    .ok_or(SparseError::Overflow)?;
            };
        }
        add!(input_bytes);
        add!(output_bytes);
        add!(fields);
        add!(work);
        add!(reference_reads);
        add!(reference_writes);
        add!(references);
        add!(retained_elements);
        add!(retained_bytes);
        self.peak_scratch_bytes = self.peak_scratch_bytes.max(other.peak_scratch_bytes);
        add!(allocation_events);
        add!(records);
        add!(header_reads);
        add!(header_writes);
        add!(headers);
        add!(objects);
        Ok(())
    }
}

/// A typed sparse-planning failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub(super) enum SparseError {
    /// The caller did not provide sorted distinct source cells.
    UnsortedCells,
    /// A zero strip width cannot describe a tile store.
    ZeroTileSize,
    /// A wire message is malformed or does not have the required shape.
    InvalidSource,
    /// A source has ambiguous duplicate scalar fields or conflicting keys.
    AmbiguousSource,
    /// An existing tile/tree/header relationship is internally inconsistent.
    InconsistentSource,
    /// A finite planning or output limit was exceeded.
    LimitExceeded { observed: usize, maximum: usize },
    /// A checked integer calculation overflowed.
    Overflow,
    /// A fallible allocation failed before publication.
    Allocation { requested: usize },
    /// Assigned object identifiers do not exactly match the plan slots.
    InvalidAssignments,
}

impl fmt::Display for SparseError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("invalid or unbounded Numbers sparse-table allocation")
    }
}

impl std::error::Error for SparseError {}

/// A row-strip entry in the logical `TableRBTree`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct RowStrip {
    /// First row in the strip.
    pub row: u32,
    /// Tile key stored as the tree value.
    pub tile_id: u32,
}

/// One missing tile which must receive an IWA object identifier from the
/// grouped writer before the `DataStore` bytes are encoded.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NewTile {
    /// `TST.TileStorage.Tile.tileid` and row-tree value.
    pub tile_id: u32,
    /// First logical row belonging to this tile.
    pub row_start: u32,
}

/// A newly required row-header record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NewRowHeader {
    /// Logical row index.
    pub row: u32,
    /// Header bucket position in `HeaderStorage.buckets`.
    pub bucket_index: u32,
    /// Native `Header.numberOfCells` value for a newly materialised row.
    pub number_of_cells: u32,
}

/// Final materialized-cell count for one logical row.
///
/// The tile leaf reports these after all changes for the row have been
/// applied. Sparse header encoding consumes the complete, sorted set so a
/// header can never publish a table-width placeholder as its cell count.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct FinalRowCount {
    /// Zero-based logical table row.
    pub row: u32,
    /// Exact populated-cell count in the final tile row.
    pub number_of_cells: u32,
}

/// A stable logical object slot, assigned only by the grouped package writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum NewObjectKind {
    /// An initially empty tile; the tile writer replaces its payload in the
    /// same grouped archive rewrite.
    Tile { tile_id: u32 },
    /// A newly created row-header bucket.
    RowHeaderBucket { bucket_index: u32 },
}

/// An object which the grouped package writer must allocate and register in
/// `PackageMetadata` before this plan can be materialised.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct NewObjectRequest {
    /// Stable plan-local object slot.
    pub slot: u32,
    /// Intended object content/role.
    pub kind: NewObjectKind,
}

/// Identifier assignments made in one package-metadata transaction.
#[derive(Debug, Clone, Copy)]
pub(super) struct ObjectAssignment {
    /// Plan-local slot.
    pub slot: u32,
    /// Newly registered IWA object identifier.
    pub object_id: u64,
    /// The allocation request kind, repeated so callers cannot accidentally
    /// exchange a header-bucket id with a tile id.
    pub kind: NewObjectKind,
    /// Marker that the package-metadata registration was staged.
    pub metadata_registered: bool,
}

/// The allocation plan retained by the grouped table-cell transaction.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct SparsePlan {
    columns: u32,
    row_strips: Vec<RowStrip>,
    new_tiles: Vec<NewTile>,
    new_headers: Vec<NewRowHeader>,
    new_header_buckets: Vec<u32>,
    new_objects: Vec<NewObjectRequest>,
    next_row_strip_id: u32,
    row_header_bucket_hash_function: u32,
    header_counts_synchronized: bool,
}

impl SparsePlan {
    /// Build a plan and return bounded cumulative resource evidence.
    pub(super) fn build_with_report(
        request: SparseRequest<'_>,
        limits: SparseLimits,
    ) -> Result<(Self, SparseReport), SparseError> {
        let input_bytes = request.row_header_buckets.iter().try_fold(
            request.data_store.len(),
            |total, source| {
                total
                    .checked_add(source.payload.len())
                    .ok_or(SparseError::Overflow)
            },
        )?;
        preflight_input(input_bytes, limits)?;
        let minimum_allocations = 8usize
            .checked_add(request.row_header_buckets.len())
            .ok_or(SparseError::Overflow)?;
        limit(minimum_allocations, limits.max_allocation_events)?;
        limit(request.row_header_buckets.len(), limits.max_references)?;
        let maximum_record_size = [
            size_of::<RowStrip>(),
            size_of::<NewTile>(),
            size_of::<NewRowHeader>(),
            size_of::<u32>(),
            size_of::<NewObjectRequest>(),
        ]
        .into_iter()
        .max()
        .ok_or(SparseError::Overflow)?;
        let mut bounded = limits;
        bounded.max_records = bounded
            .max_records
            .min(bounded.max_retained_elements)
            .min(bounded.max_retained_bytes / maximum_record_size.max(1));
        let plan = Self::build(request, bounded)?;
        let fields = request.row_header_buckets.iter().try_fold(
            count_fields(request.data_store, limits.max_fields)?,
            |total, source| {
                total
                    .checked_add(count_fields(source.payload, limits.max_fields)?)
                    .ok_or(SparseError::Overflow)
            },
        )?;
        let retained_elements = plan
            .row_strips
            .len()
            .checked_add(plan.new_tiles.len())
            .and_then(|value| value.checked_add(plan.new_headers.len()))
            .and_then(|value| value.checked_add(plan.new_header_buckets.len()))
            .and_then(|value| value.checked_add(plan.new_objects.len()))
            .ok_or(SparseError::Overflow)?;
        let retained_bytes = plan
            .row_strips
            .len()
            .checked_mul(size_of::<RowStrip>())
            .and_then(|value| {
                value.checked_add(plan.new_tiles.len().checked_mul(size_of::<NewTile>())?)
            })
            .and_then(|value| {
                value.checked_add(
                    plan.new_headers
                        .len()
                        .checked_mul(size_of::<NewRowHeader>())?,
                )
            })
            .and_then(|value| {
                value.checked_add(
                    plan.new_header_buckets
                        .len()
                        .checked_mul(size_of::<u32>())?,
                )
            })
            .and_then(|value| {
                value.checked_add(
                    plan.new_objects
                        .len()
                        .checked_mul(size_of::<NewObjectRequest>())?,
                )
            })
            .ok_or(SparseError::Overflow)?;
        let parsed_data_store = parse_data_store(request.data_store, limits.max_fields)?;
        let reference_reads = parse_tile_storage(parsed_data_store.tiles, limits.max_fields)?
            .len()
            .checked_add(
                parse_header_storage(parsed_data_store.row_headers, limits.max_fields)?
                    .bucket_ids
                    .len(),
            )
            .ok_or(SparseError::Overflow)?;
        let header_reads =
            request
                .row_header_buckets
                .iter()
                .try_fold(0usize, |total, source| {
                    total
                        .checked_add(
                            parse_header_bucket(source.payload, limits.max_fields)?
                                .rows
                                .len(),
                        )
                        .ok_or(SparseError::Overflow)
                })?;
        let records = retained_elements;
        let work = input_bytes
            .checked_add(fields)
            .and_then(|value| value.checked_add(request.cells.len()))
            .and_then(|value| value.checked_add(records))
            .and_then(|value| value.checked_add(reference_reads))
            .and_then(|value| value.checked_add(header_reads))
            .ok_or(SparseError::Overflow)?;
        let report = SparseReport {
            input_bytes,
            fields,
            work,
            reference_reads,
            references: reference_reads,
            retained_elements,
            retained_bytes,
            peak_scratch_bytes: fields
                .checked_mul(size_of::<Field<'_>>())
                .ok_or(SparseError::Overflow)?,
            allocation_events: minimum_allocations,
            records,
            header_reads,
            headers: header_reads,
            objects: plan.new_objects.len(),
            ..SparseReport::default()
        };
        validate_report(report, limits)?;
        Ok((plan, report))
    }

    /// Build a deterministic plan without mutating archive or metadata state.
    ///
    /// The returned rows are sorted and contain every original tree node plus
    /// every newly needed strip.  Consequently a 513-row write beginning at
    /// row zero yields `(0,0), (256,1), (512,2)` and `next_row_strip_id == 3`
    /// for the usual initially materialised tile zero.
    pub(super) fn build(
        request: SparseRequest<'_>,
        limits: SparseLimits,
    ) -> Result<Self, SparseError> {
        if request.tile_size == 0 {
            return Err(SparseError::ZeroTileSize);
        }
        limit(request.cells.len(), limits.max_cells)?;
        ensure_sorted_cells(request.cells)?;

        let data_store = parse_data_store(request.data_store, limits.max_fields)?;
        let tiles = parse_tile_storage(data_store.tiles, limits.max_fields)?;
        let row_strips = parse_row_tree(data_store.row_tree, limits.max_fields)?;
        ensure_sorted_strips(&row_strips)?;
        for strip in &row_strips {
            if strip
                .tile_id
                .checked_mul(request.tile_size)
                .is_none_or(|row| row != strip.row)
                || find_tile(&tiles, strip.tile_id).is_none()
            {
                return Err(SparseError::InconsistentSource);
            }
        }
        let header_storage = parse_header_storage(data_store.row_headers, limits.max_fields)?;
        limit(request.row_header_buckets.len(), limits.max_records)?;
        if usize::try_from(header_storage.bucket_count).map_err(|_| SparseError::Overflow)?
            != request.row_header_buckets.len()
        {
            return Err(SparseError::InconsistentSource);
        }
        let mut bucket_rows = Vec::new();
        reserve(&mut bucket_rows, request.row_header_buckets.len())?;
        for (index, source) in request.row_header_buckets.iter().enumerate() {
            if source.object_id != header_storage.bucket_ids[index] {
                return Err(SparseError::InconsistentSource);
            }
            let bucket = parse_header_bucket(source.payload, limits.max_fields)?;
            if bucket.hash_function != header_storage.hash_function {
                return Err(SparseError::InconsistentSource);
            }
            bucket_rows.push(bucket);
        }

        let (distinct_tiles, distinct_rows) =
            distinct_cell_counts(request.cells, request.tile_size)?;
        let required_bucket_count = request
            .cells
            .last()
            .map(|cell| cell.row / HEADER_BUCKET_ROWS)
            .map_or(Ok(0), |index| {
                index.checked_add(1).ok_or(SparseError::Overflow)
            })?;
        let missing_bucket_count = if required_bucket_count > header_storage.bucket_count {
            required_bucket_count
                .checked_sub(header_storage.bucket_count)
                .ok_or(SparseError::Overflow)?
        } else {
            0
        };
        let mut new_tiles = Vec::new();
        let mut new_headers = Vec::new();
        let mut new_header_buckets = Vec::new();
        let mut new_objects = Vec::new();
        let mut missing_strips = Vec::new();
        reserve(&mut new_tiles, distinct_tiles)?;
        reserve(&mut new_headers, distinct_rows)?;
        reserve(
            &mut new_header_buckets,
            usize::try_from(missing_bucket_count).map_err(|_| SparseError::Overflow)?,
        )?;
        reserve(&mut missing_strips, distinct_tiles)?;
        reserve(
            &mut new_objects,
            distinct_tiles
                .checked_add(
                    usize::try_from(missing_bucket_count).map_err(|_| SparseError::Overflow)?,
                )
                .ok_or(SparseError::Overflow)?,
        )?;
        for bucket_index in header_storage.bucket_count..required_bucket_count {
            new_header_buckets.push(bucket_index);
        }

        let mut last_tile = None;
        let mut last_row = None;
        for cell in request.cells {
            let tile_id = cell.row / request.tile_size;
            if last_tile != Some(tile_id) {
                if find_tile(&tiles, tile_id).is_none() {
                    new_tiles.push(NewTile {
                        tile_id,
                        row_start: tile_id
                            .checked_mul(request.tile_size)
                            .ok_or(SparseError::Overflow)?,
                    });
                }
                let row_start = tile_id
                    .checked_mul(request.tile_size)
                    .ok_or(SparseError::Overflow)?;
                match find_strip(&row_strips, row_start) {
                    Some(existing) if existing.tile_id != tile_id => {
                        return Err(SparseError::InconsistentSource);
                    },
                    Some(_) => {},
                    None => missing_strips.push(RowStrip {
                        row: row_start,
                        tile_id,
                    }),
                }
                last_tile = Some(tile_id);
            }
            if last_row != Some(cell.row) {
                let bucket_index = cell.row / HEADER_BUCKET_ROWS;
                if !header_exists(&bucket_rows, bucket_index, cell.row)? {
                    new_headers.push(NewRowHeader {
                        row: cell.row,
                        bucket_index,
                        number_of_cells: 0,
                    });
                }
                last_row = Some(cell.row);
            }
        }
        limit(new_tiles.len(), limits.max_records)?;
        limit(new_headers.len(), limits.max_records)?;
        let row_strips = merge_row_strips(row_strips, missing_strips)?;

        let mut next_slot = 0_u32;
        for tile in &new_tiles {
            new_objects.push(NewObjectRequest {
                slot: next_slot,
                kind: NewObjectKind::Tile {
                    tile_id: tile.tile_id,
                },
            });
            next_slot = next_slot.checked_add(1).ok_or(SparseError::Overflow)?;
        }
        for &bucket_index in &new_header_buckets {
            new_objects.push(NewObjectRequest {
                slot: next_slot,
                kind: NewObjectKind::RowHeaderBucket { bucket_index },
            });
            next_slot = next_slot.checked_add(1).ok_or(SparseError::Overflow)?;
        }

        let maximum_tree_tile = row_strips.iter().map(|strip| strip.tile_id).max();
        let maximum_source_tile = tiles.iter().map(|tile| tile.tile_id).max();
        let maximum_tile = maximum_tree_tile
            .into_iter()
            .chain(maximum_source_tile)
            .max();
        let derived_next_row_strip_id = match maximum_tile {
            Some(value) => value.checked_add(1).ok_or(SparseError::Overflow)?,
            None => 0,
        };
        let next_row_strip_id = data_store.next_row_strip_id.max(derived_next_row_strip_id);
        let header_counts_synchronized = new_headers.is_empty();
        Ok(Self {
            columns: request.columns,
            row_strips,
            new_tiles,
            new_headers,
            new_header_buckets,
            new_objects,
            next_row_strip_id,
            row_header_bucket_hash_function: header_storage.hash_function,
            header_counts_synchronized,
        })
    }

    /// Bind every newly required row header to its exact final tile count.
    ///
    /// `counts` must be strictly row-sorted and must name each row in
    /// [`Self::new_headers`] exactly once. A zero final count is rejected:
    /// such a row must remain absent from both the tile and header stores.
    pub(super) fn synchronize_new_header_counts(
        &mut self,
        counts: &[FinalRowCount],
    ) -> Result<(), SparseError> {
        if counts.len() != self.new_headers.len()
            || counts.windows(2).any(|pair| pair[0].row >= pair[1].row)
        {
            return Err(SparseError::InvalidAssignments);
        }
        for (header, count) in self.new_headers.iter().zip(counts) {
            if header.row != count.row
                || count.number_of_cells == 0
                || count.number_of_cells > self.columns
            {
                return Err(SparseError::InvalidAssignments);
            }
        }
        for (header, count) in self.new_headers.iter_mut().zip(counts) {
            header.number_of_cells = count.number_of_cells;
        }
        self.header_counts_synchronized = true;
        Ok(())
    }

    /// Synchronize header counts and report the exact record pass.
    pub(super) fn synchronize_new_header_counts_with_report(
        &mut self,
        counts: &[FinalRowCount],
        limits: SparseLimits,
    ) -> Result<SparseReport, SparseError> {
        let report = SparseReport {
            work: counts.len(),
            records: counts.len(),
            ..SparseReport::default()
        };
        validate_report(report, limits)?;
        self.synchronize_new_header_counts(counts)?;
        Ok(report)
    }

    /// Fully sorted replacement tree entries.
    #[must_use]
    pub(super) fn row_strips(&self) -> &[RowStrip] {
        &self.row_strips
    }
    /// Missing tile keys, in ascending tile-id order.
    #[must_use]
    pub(super) fn new_tiles(&self) -> &[NewTile] {
        &self.new_tiles
    }
    /// Missing logical row headers, in ascending row order.
    #[must_use]
    pub(super) fn new_headers(&self) -> &[NewRowHeader] {
        &self.new_headers
    }
    /// Missing positional row-header buckets, in ascending bucket order.
    #[must_use]
    pub(super) fn new_header_buckets(&self) -> &[u32] {
        &self.new_header_buckets
    }
    /// Plan-local metadata allocation requests.
    #[must_use]
    pub(super) fn new_objects(&self) -> &[NewObjectRequest] {
        &self.new_objects
    }
    /// Replacement `DataStore.nextRowStripID`.
    #[must_use]
    pub(super) const fn next_row_strip_id(&self) -> u32 {
        self.next_row_strip_id
    }
}

/// Borrow the single embedded `TST.DataStore` payload from a table model.
///
/// The resolver already establishes the semantic model shape; this focused
/// adapter gives the grouped sparse writer the exact field-4 bytes on which
/// [`SparsePlan::build`] operates. Duplicate field 4 occurrences, a wrong
/// wire type, or a malformed outer message are refused.
pub(super) fn table_model_data_store(
    source: &[u8],
    limits: SparseLimits,
) -> Result<&[u8], SparseError> {
    let fields = parse_fields(source, limits.max_fields)?;
    unique_bytes(&fields, 4)?.ok_or(SparseError::InvalidSource)
}

/// Borrow the data store and return exact bounded scan evidence.
pub(super) fn table_model_data_store_with_report(
    source: &[u8],
    limits: SparseLimits,
) -> Result<(&[u8], SparseReport), SparseError> {
    preflight_input(source.len(), limits)?;
    let fields = count_fields(source, limits.max_fields)?;
    let value = table_model_data_store(source, limits)?;
    let scratch = fields
        .checked_mul(size_of::<Field<'_>>())
        .ok_or(SparseError::Overflow)?;
    let report = scan_report(source.len(), fields, scratch, usize::from(fields != 0))?;
    validate_report(report, limits)?;
    Ok((value, report))
}

/// Replace the single embedded `TST.DataStore` in a table-model payload.
///
/// Every outer field other than field 4 is copied byte-for-byte and in source
/// order. This is the final sparse leaf step before the grouped writer updates
/// the model object's aggregate `ArchiveInfo` references once.
pub(super) fn rewrite_table_model_data_store(
    source: &[u8],
    data_store: &[u8],
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    let fields = parse_fields(source, limits.max_fields)?;
    let _ = unique_bytes(&fields, 4)?.ok_or(SparseError::InvalidSource)?;
    let replacement_len = length_field_len(4, data_store.len())?;
    let source_field_len = fields
        .iter()
        .find(|field| field.number == 4)
        .ok_or(SparseError::InvalidSource)?
        .raw
        .len();
    let output_len = source
        .len()
        .checked_sub(source_field_len)
        .and_then(|length| length.checked_add(replacement_len))
        .ok_or(SparseError::Overflow)?;
    limit(output_len, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    for field in fields {
        if field.number == 4 {
            append_length(&mut output, 4, data_store)?;
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    if output.len() != output_len {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

/// Replace the embedded data store and return exact bounded scan/output evidence.
pub(super) fn rewrite_table_model_data_store_with_report(
    source: &[u8],
    data_store: &[u8],
    limits: SparseLimits,
) -> Result<(Vec<u8>, SparseReport), SparseError> {
    preflight_input(
        source
            .len()
            .checked_add(data_store.len())
            .ok_or(SparseError::Overflow)?,
        limits,
    )?;
    let bounded = output_bounded_limits(limits);
    let output = rewrite_table_model_data_store(source, data_store, bounded)?;
    let fields = count_fields(source, limits.max_fields)?;
    let retained_elements = usize::from(!output.is_empty());
    let retained_bytes = output.len();
    let work = source
        .len()
        .checked_add(data_store.len())
        .and_then(|value| value.checked_add(output.len()))
        .and_then(|value| value.checked_add(fields))
        .ok_or(SparseError::Overflow)?;
    let report = SparseReport {
        input_bytes: source
            .len()
            .checked_add(data_store.len())
            .ok_or(SparseError::Overflow)?,
        output_bytes: output.len(),
        fields,
        work,
        retained_elements,
        retained_bytes,
        peak_scratch_bytes: output.len().max(
            fields
                .checked_mul(size_of::<Field<'_>>())
                .ok_or(SparseError::Overflow)?,
        ),
        allocation_events: retained_elements
            .checked_add(usize::from(fields != 0))
            .ok_or(SparseError::Overflow)?,
        ..SparseReport::default()
    };
    validate_report(report, limits)?;
    Ok((output, report))
}

/// A rewritten embedded `DataStore` plus the payload of one newly created
/// header bucket, if the source had none.  Existing bucket objects are
/// intentionally not rewritten here: `rewrite_header_bucket` is called only
/// for buckets that actually received a new header.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct DataStoreRewrite {
    /// Exact replacement `TST.DataStore` bytes.
    pub data_store: Vec<u8>,
    /// Brand-new bucket payloads keyed by their plan-local allocation slots.
    pub new_header_buckets: Vec<(u32, Vec<u8>)>,
}

/// Rewrite the embedded `DataStore`, preserving every unselected field byte.
///
/// `assignments` must contain exactly the requests returned by the plan, must
/// be metadata-registered, and supplies the only new object references which
/// appear in the output.  This function allocates no archive state.
pub(super) fn rewrite_data_store(
    source: &[u8],
    plan: &SparsePlan,
    assignments: &[ObjectAssignment],
    limits: SparseLimits,
) -> Result<DataStoreRewrite, SparseError> {
    if !plan.header_counts_synchronized {
        return Err(SparseError::InvalidAssignments);
    }
    validate_assignments(plan, assignments)?;
    let parsed = parse_data_store(source, limits.max_fields)?;
    let source_fields = parse_fields(source, limits.max_fields)?;
    let _ = unique_varint(&source_fields, 7)?.ok_or(SparseError::InvalidSource)?;
    let tiles = rewrite_tile_storage(parsed.tiles, plan.new_tiles(), assignments, limits)?;
    let tree = rewrite_row_tree(parsed.row_tree, plan.row_strips(), limits)?;

    let mut new_header_buckets = Vec::new();
    reserve(&mut new_header_buckets, plan.new_header_buckets().len())?;
    let row_headers = if plan.new_header_buckets().is_empty() {
        copy_bytes(parsed.row_headers)?
    } else {
        let first = plan.new_tiles().len();
        let end = first
            .checked_add(plan.new_header_buckets().len())
            .ok_or(SparseError::Overflow)?;
        let header_assignments = assignments
            .get(first..end)
            .ok_or(SparseError::InvalidAssignments)?;
        let mut header_start = 0usize;
        for (&bucket_index, assignment) in plan.new_header_buckets().iter().zip(header_assignments)
        {
            let headers = plan.new_headers();
            if header_start < headers.len() && headers[header_start].bucket_index < bucket_index {
                return Err(SparseError::InconsistentSource);
            }
            let header_end = header_start
                .checked_add(
                    headers[header_start..]
                        .iter()
                        .take_while(|header| header.bucket_index == bucket_index)
                        .count(),
                )
                .ok_or(SparseError::Overflow)?;
            let bucket = encode_header_bucket(
                plan.row_header_bucket_hash_function,
                &headers[header_start..header_end],
                limits,
            )?;
            new_header_buckets.push((assignment.slot, bucket));
            header_start = header_end;
        }
        if header_start != plan.new_headers().len() {
            return Err(SparseError::InconsistentSource);
        }
        rewrite_header_storage_append(parsed.row_headers, header_assignments, limits)?
    };

    let fields = source_fields;
    let mut output_len = 0_usize;
    for field in &fields {
        let replacement = match field.number {
            1 => Some(&row_headers),
            3 => Some(&tiles),
            9 => Some(&tree),
            _ => None,
        };
        let field_len = if field.number == 7 {
            varint_field_len(7, u64::from(plan.next_row_strip_id()))?
        } else {
            match replacement {
                Some(value) => length_field_len(field.number, value.len())?,
                None => field.raw.len(),
            }
        };
        output_len = output_len
            .checked_add(field_len)
            .ok_or(SparseError::Overflow)?;
    }
    limit(output_len, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    let mut next_written = false;
    for field in fields {
        match field.number {
            1 => append_length(&mut output, 1, &row_headers)?,
            3 => append_length(&mut output, 3, &tiles)?,
            7 => {
                append_varint_field(&mut output, 7, u64::from(plan.next_row_strip_id()))?;
                next_written = true;
            },
            9 => append_length(&mut output, 9, &tree)?,
            _ => output.extend_from_slice(field.raw),
        }
    }
    if !next_written {
        return Err(SparseError::InvalidSource);
    }
    if output.len() != output_len {
        return Err(SparseError::InconsistentSource);
    }
    Ok(DataStoreRewrite {
        data_store: output,
        new_header_buckets,
    })
}

/// Rewrite a data store and return exact bounded output/reference evidence.
pub(super) fn rewrite_data_store_with_report(
    source: &[u8],
    plan: &SparsePlan,
    assignments: &[ObjectAssignment],
    limits: SparseLimits,
) -> Result<(DataStoreRewrite, SparseReport), SparseError> {
    preflight_input(source.len(), limits)?;
    limit(assignments.len(), limits.max_references)?;
    let output = rewrite_data_store(source, plan, assignments, output_bounded_limits(limits))?;
    let bucket_bytes =
        output
            .new_header_buckets
            .iter()
            .try_fold(0usize, |total, (_slot, payload)| {
                total
                    .checked_add(payload.len())
                    .ok_or(SparseError::Overflow)
            })?;
    let output_bytes = output
        .data_store
        .len()
        .checked_add(bucket_bytes)
        .ok_or(SparseError::Overflow)?;
    let fields = count_fields(source, limits.max_fields)?;
    let retained_elements = 1usize
        .checked_add(output.new_header_buckets.len())
        .ok_or(SparseError::Overflow)?;
    let retained_bytes = output_bytes
        .checked_add(
            output
                .new_header_buckets
                .len()
                .checked_mul(size_of::<(u32, Vec<u8>)>())
                .ok_or(SparseError::Overflow)?,
        )
        .ok_or(SparseError::Overflow)?;
    let records = plan
        .row_strips
        .len()
        .checked_add(plan.new_tiles.len())
        .and_then(|value| value.checked_add(plan.new_headers.len()))
        .ok_or(SparseError::Overflow)?;
    let header_writes = plan
        .new_headers
        .iter()
        .filter(|header| {
            plan.new_header_buckets
                .binary_search(&header.bucket_index)
                .is_ok()
        })
        .count();
    let work = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(fields))
        .and_then(|value| value.checked_add(records))
        .and_then(|value| value.checked_add(assignments.len()))
        .ok_or(SparseError::Overflow)?;
    let report = SparseReport {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work,
        reference_reads: 0,
        reference_writes: assignments.len(),
        references: assignments.len(),
        retained_elements,
        retained_bytes,
        peak_scratch_bytes: retained_bytes,
        allocation_events: 5usize
            .checked_add(
                output
                    .new_header_buckets
                    .len()
                    .checked_mul(2)
                    .ok_or(SparseError::Overflow)?,
            )
            .ok_or(SparseError::Overflow)?,
        records,
        header_reads: 0,
        header_writes,
        headers: header_writes,
        objects: assignments.len(),
    };
    validate_report(report, limits)?;
    Ok((output, report))
}

/// Append only the missing plan headers to one selected existing bucket.
///
/// Callers use `bucket_index` from [`NewRowHeader`] and publish the returned
/// payload as a replacement for that one bucket object; untouched bucket
/// objects remain byte-identical.
#[cfg(test)]
pub(super) fn rewrite_header_bucket(
    source: &[u8],
    bucket_index: u32,
    plan: &SparsePlan,
    limits: SparseLimits,
) -> Result<Option<Vec<u8>>, SparseError> {
    if !plan.header_counts_synchronized {
        return Err(SparseError::InvalidAssignments);
    }
    let headers = plan.new_headers();
    let addition_start = headers.partition_point(|header| header.bucket_index < bucket_index);
    let addition_end = addition_start
        .checked_add(
            headers[addition_start..]
                .iter()
                .take_while(|header| header.bucket_index == bucket_index)
                .count(),
        )
        .ok_or(SparseError::Overflow)?;
    let additions = &headers[addition_start..addition_end];
    if additions.is_empty() {
        return Ok(None);
    }
    let fields = parse_fields(source, limits.max_fields)?;
    let hash = unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?;
    let hash = u32::try_from(hash).map_err(|_| SparseError::InvalidSource)?;
    let mut output_len = source.len();
    for header in additions {
        output_len = output_len
            .checked_add(length_field_len(2, header_wire_len(*header)?)?)
            .ok_or(SparseError::Overflow)?;
    }
    limit(output_len, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    output.extend_from_slice(source);
    for header in additions {
        append_length_prefix(&mut output, 2, header_wire_len(*header)?)?;
        append_header(&mut output, *header)?;
    }
    if output.len() != output_len || hash != plan.row_header_bucket_hash_function {
        return Err(SparseError::InconsistentSource);
    }
    Ok(Some(output))
}

/// Rewrite one bucket and return exact bounded scan/output/header evidence.
#[cfg(test)]
pub(super) fn rewrite_header_bucket_with_report(
    source: &[u8],
    bucket_index: u32,
    plan: &SparsePlan,
    limits: SparseLimits,
) -> Result<(Option<Vec<u8>>, SparseReport), SparseError> {
    preflight_input(source.len(), limits)?;
    let output = rewrite_header_bucket(source, bucket_index, plan, output_bounded_limits(limits))?;
    let fields = count_fields(source, limits.max_fields)?;
    let headers = plan
        .new_headers
        .iter()
        .filter(|header| header.bucket_index == bucket_index)
        .count();
    let header_reads = if output.is_some() {
        parse_header_bucket(source, limits.max_fields)?.rows.len()
    } else {
        0
    };
    let header_writes = if output.is_some() { headers } else { 0 };
    let header_items = header_reads
        .checked_add(header_writes)
        .ok_or(SparseError::Overflow)?;
    let output_bytes = output.as_ref().map_or(0, Vec::len);
    let retained_elements = usize::from(output.is_some());
    let retained_bytes = output_bytes;
    let work = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(fields))
        .and_then(|value| value.checked_add(header_items))
        .ok_or(SparseError::Overflow)?;
    let report = SparseReport {
        input_bytes: source.len(),
        output_bytes,
        fields,
        work,
        retained_elements,
        retained_bytes,
        peak_scratch_bytes: output_bytes.max(
            headers
                .checked_mul(size_of::<NewRowHeader>())
                .ok_or(SparseError::Overflow)?,
        ),
        allocation_events: if headers == 0 { 0 } else { 2 },
        records: header_items,
        header_reads,
        header_writes,
        headers: header_items,
        ..SparseReport::default()
    };
    validate_report(report, limits)?;
    Ok((output, report))
}

/// Output-free evidence retained by one prepared existing-bucket rewrite.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct HeaderBucketPrepareReport {
    report: SparseReport,
}

impl HeaderBucketPrepareReport {
    #[must_use]
    pub(super) const fn report(self) -> SparseReport {
        self.report
    }

    #[must_use]
    pub(super) const fn output_bytes(self) -> usize {
        self.report.output_bytes
    }
}

/// Exact independent limits needed after an existing header bucket is planned.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub(super) struct HeaderBucketExecutionRequirements {
    pub(super) input_bytes: usize,
    pub(super) output_bytes: usize,
    pub(super) retained_bytes: usize,
    pub(super) retained_elements: usize,
    pub(super) peak_scratch_bytes: usize,
    pub(super) allocation_events: usize,
    pub(super) fields: usize,
    pub(super) work: usize,
    pub(super) headers: usize,
    pub(super) header_reads: usize,
    pub(super) header_writes: usize,
}

impl HeaderBucketExecutionRequirements {
    #[must_use]
    pub(super) const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[must_use]
    pub(super) const fn exact_limits(self) -> HeaderBucketExecutionLimits {
        HeaderBucketExecutionLimits {
            max_input_bytes: self.input_bytes,
            max_output_bytes: self.output_bytes,
            max_retained_bytes: self.retained_bytes,
            max_retained_elements: self.retained_elements,
            max_peak_scratch_bytes: self.peak_scratch_bytes,
            max_allocation_events: self.allocation_events,
            max_fields: self.fields,
            max_work: self.work,
            max_headers: self.headers,
        }
    }
}

/// Execution-only ceilings. Every axis is checked before the candidate Vec.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct HeaderBucketExecutionLimits {
    pub(super) max_input_bytes: usize,
    pub(super) max_output_bytes: usize,
    pub(super) max_retained_bytes: usize,
    pub(super) max_retained_elements: usize,
    pub(super) max_peak_scratch_bytes: usize,
    pub(super) max_allocation_events: usize,
    pub(super) max_fields: usize,
    pub(super) max_work: usize,
    pub(super) max_headers: usize,
}

#[derive(Debug, Clone, Copy)]
enum PreparedHeaderField<'source> {
    Preserve(&'source [u8]),
    HeaderPreserve {
        raw: &'source [u8],
        row: u32,
        count: u32,
        payload_len: usize,
        nested_fields: usize,
    },
    HeaderDelete,
    HeaderReplace {
        source: &'source [u8],
        row: u32,
        count: u32,
        encoded_len: usize,
        payload_len: usize,
        nested_fields: usize,
    },
}

impl PreparedHeaderField<'_> {
    const fn retained_source_bytes(self) -> usize {
        match self {
            Self::Preserve(raw) | Self::HeaderPreserve { raw, .. } => raw.len(),
            Self::HeaderDelete => 0,
            Self::HeaderReplace { source, .. } => source.len(),
        }
    }
}

/// A complete logical rewrite for an already materialized header bucket.
///
/// It borrows raw source fields and owns only exact-capacity field actions and
/// appended row facts. No candidate or per-header replacement bytes exist
/// until [`PreparedExistingHeaderBucketRewrite::execute`].
pub(super) struct PreparedExistingHeaderBucketRewrite<'source> {
    bucket_index: u32,
    hash: u32,
    fields: Vec<PreparedHeaderField<'source>>,
    appended: Vec<NewRowHeader>,
    output_len: usize,
    output_headers: usize,
    prepare_report: HeaderBucketPrepareReport,
    requirements: HeaderBucketExecutionRequirements,
}

impl PreparedExistingHeaderBucketRewrite<'_> {
    #[must_use]
    pub(super) const fn prepare_report(&self) -> HeaderBucketPrepareReport {
        self.prepare_report
    }

    #[must_use]
    pub(super) const fn execution_requirements(&self) -> HeaderBucketExecutionRequirements {
        self.requirements
    }

    pub(super) fn execute(
        self,
        limits: HeaderBucketExecutionLimits,
    ) -> Result<(Option<Vec<u8>>, SparseReport), SparseError> {
        ensure_header_execution_limits(self.requirements, limits)?;
        if self.prepare_report.output_bytes() != 0 {
            return Err(SparseError::InconsistentSource);
        }
        if self.output_len == 0 {
            return Ok((None, header_execution_report(self.requirements)));
        }

        #[cfg(test)]
        record_prepared_header_execution_allocation();
        let mut output = Vec::new();
        reserve_exact_capacity(&mut output, self.output_len)?;
        for field in &self.fields {
            match *field {
                PreparedHeaderField::Preserve(raw)
                | PreparedHeaderField::HeaderPreserve { raw, .. } => {
                    output.extend_from_slice(raw);
                },
                PreparedHeaderField::HeaderDelete => {},
                PreparedHeaderField::HeaderReplace {
                    source,
                    count,
                    payload_len,
                    ..
                } => {
                    append_length_prefix(&mut output, 2, payload_len)?;
                    append_rewritten_header(&mut output, source, count)?;
                },
            }
        }
        for header in &self.appended {
            append_length_prefix(&mut output, 2, header_wire_len(*header)?)?;
            append_header(&mut output, *header)?;
        }
        if output.len() != self.output_len || output.capacity() != self.output_len {
            return Err(SparseError::InconsistentSource);
        }
        reopen_prepared_header_bucket(&output, &self)?;
        Ok((Some(output), header_execution_report(self.requirements)))
    }
}

#[derive(Debug)]
enum HeaderAction {
    Preserve,
    Delete,
    Replace(Vec<u8>),
}

/// Synchronize touched existing and new rows in one exact header bucket.
///
/// `final_rows` must be strictly sorted, contain only rows belonging to
/// `bucket_index`, and include every newly planned header for that bucket.
/// A zero count removes an existing header; a nonzero changed count rewrites
/// only the header's field 4; and a nonzero missing row appends one canonical
/// header. Untouched header records and every bucket/header unknown field are
/// copied byte-for-byte.
fn rewrite_header_bucket_rows_with_report(
    source: &[u8],
    bucket_index: u32,
    columns: u32,
    final_rows: &[FinalRowCount],
    expected_hash: Option<u32>,
    limits: SparseLimits,
) -> Result<(Option<Vec<u8>>, SparseReport), SparseError> {
    limit(final_rows.len(), limits.max_records)?;
    if final_rows.windows(2).any(|pair| pair[0].row >= pair[1].row)
        || final_rows.iter().any(|row| {
            row.row / HEADER_BUCKET_ROWS != bucket_index || row.number_of_cells > columns
        })
    {
        return Err(SparseError::InvalidAssignments);
    }

    preflight_input(source.len(), limits)?;
    let fields = parse_fields(source, limits.max_fields)?;
    let hash = unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?;
    let hash = u32::try_from(hash).map_err(|_| SparseError::InvalidSource)?;
    if expected_hash.is_some_and(|expected| expected != hash) {
        return Err(SparseError::InconsistentSource);
    }
    let source_count = fields.iter().filter(|field| field.number == 2).count();
    limit(source_count, limits.max_records)?;
    let mut source_headers = Vec::new();
    reserve(&mut source_headers, source_count)?;
    let mut fields_read = fields.len();
    for field in fields.iter().filter(|field| field.number == 2) {
        if field.wire != 2 {
            return Err(SparseError::InvalidSource);
        }
        let nested = parse_fields(field.value, limits.max_fields)?;
        fields_read = fields_read
            .checked_add(nested.len())
            .ok_or(SparseError::Overflow)?;
        validate_header_fields(&nested)?;
        let row = u32::try_from(unique_varint(&nested, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        let count = u32::try_from(unique_varint(&nested, 4)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        if row / HEADER_BUCKET_ROWS != bucket_index || count > columns {
            return Err(SparseError::InconsistentSource);
        }
        source_headers.push((row, count, field.value));
    }
    let mut existing_rows = Vec::new();
    reserve(&mut existing_rows, source_headers.len())?;
    existing_rows.extend(source_headers.iter().map(|(row, _count, _raw)| *row));
    existing_rows.sort_unstable();
    if existing_rows.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(SparseError::AmbiguousSource);
    }

    let mut actions = Vec::new();
    reserve(&mut actions, source_headers.len())?;
    let mut header_writes = 0usize;
    let mut replacement_bytes = 0usize;
    for &(row, old_count, source_header) in &source_headers {
        let desired = final_rows
            .binary_search_by_key(&row, |candidate| candidate.row)
            .ok()
            .and_then(|index| final_rows.get(index));
        let action = match desired {
            Some(desired) if desired.number_of_cells == 0 => {
                header_writes = header_writes.checked_add(1).ok_or(SparseError::Overflow)?;
                HeaderAction::Delete
            },
            Some(desired) if desired.number_of_cells != old_count => {
                let replacement =
                    rewrite_header_cell_count(source_header, desired.number_of_cells, limits)?;
                replacement_bytes = replacement_bytes
                    .checked_add(replacement.len())
                    .ok_or(SparseError::Overflow)?;
                header_writes = header_writes.checked_add(1).ok_or(SparseError::Overflow)?;
                HeaderAction::Replace(replacement)
            },
            Some(_) | None => HeaderAction::Preserve,
        };
        actions.push(action);
    }
    let appended = final_rows
        .iter()
        .filter(|row| row.number_of_cells != 0 && existing_rows.binary_search(&row.row).is_err())
        .count();
    header_writes = header_writes
        .checked_add(appended)
        .ok_or(SparseError::Overflow)?;

    let output = if header_writes == 0 {
        None
    } else {
        let mut output_len = 0usize;
        let mut header_position = 0usize;
        for field in &fields {
            let field_len = if field.number != 2 {
                field.raw.len()
            } else {
                let action = actions
                    .get(header_position)
                    .ok_or(SparseError::InconsistentSource)?;
                header_position = header_position
                    .checked_add(1)
                    .ok_or(SparseError::Overflow)?;
                match action {
                    HeaderAction::Preserve => field.raw.len(),
                    HeaderAction::Delete => 0,
                    HeaderAction::Replace(payload) => length_field_len(2, payload.len())?,
                }
            };
            output_len = output_len
                .checked_add(field_len)
                .ok_or(SparseError::Overflow)?;
        }
        for row in final_rows.iter().filter(|row| {
            row.number_of_cells != 0 && existing_rows.binary_search(&row.row).is_err()
        }) {
            let header = NewRowHeader {
                row: row.row,
                bucket_index,
                number_of_cells: row.number_of_cells,
            };
            output_len = output_len
                .checked_add(length_field_len(2, header_wire_len(header)?)?)
                .ok_or(SparseError::Overflow)?;
        }
        limit(output_len, output_bounded_limits(limits).max_output_bytes)?;
        let mut output = Vec::new();
        reserve(&mut output, output_len)?;
        let mut header_position = 0usize;
        for field in &fields {
            if field.number != 2 {
                output.extend_from_slice(field.raw);
                continue;
            }
            match actions
                .get(header_position)
                .ok_or(SparseError::InconsistentSource)?
            {
                HeaderAction::Preserve => output.extend_from_slice(field.raw),
                HeaderAction::Delete => {},
                HeaderAction::Replace(payload) => append_length(&mut output, 2, payload)?,
            }
            header_position = header_position
                .checked_add(1)
                .ok_or(SparseError::Overflow)?;
        }
        for row in final_rows.iter().filter(|row| {
            row.number_of_cells != 0 && existing_rows.binary_search(&row.row).is_err()
        }) {
            let header = NewRowHeader {
                row: row.row,
                bucket_index,
                number_of_cells: row.number_of_cells,
            };
            append_length_prefix(&mut output, 2, header_wire_len(header)?)?;
            append_header(&mut output, header)?;
        }
        if output.len() != output_len {
            return Err(SparseError::InconsistentSource);
        }
        Some(output)
    };

    let header_reads = source_headers.len();
    let header_items = header_reads
        .checked_add(header_writes)
        .ok_or(SparseError::Overflow)?;
    let output_bytes = output.as_ref().map_or(0, Vec::len);
    let scratch_bytes = source_headers
        .len()
        .checked_mul(size_of::<(u32, u32, &[u8])>())
        .and_then(|value| value.checked_add(existing_rows.len().checked_mul(size_of::<u32>())?))
        .and_then(|value| value.checked_add(actions.len().checked_mul(size_of::<HeaderAction>())?))
        .and_then(|value| value.checked_add(replacement_bytes))
        .ok_or(SparseError::Overflow)?;
    let work = source
        .len()
        .checked_add(output_bytes)
        .and_then(|value| value.checked_add(fields_read))
        .and_then(|value| value.checked_add(final_rows.len()))
        .and_then(|value| value.checked_add(header_items))
        .ok_or(SparseError::Overflow)?;
    let report = SparseReport {
        input_bytes: source.len(),
        output_bytes,
        fields: fields_read,
        work,
        retained_elements: usize::from(output.is_some()),
        retained_bytes: output_bytes,
        peak_scratch_bytes: scratch_bytes,
        allocation_events: 3usize
            .checked_add(header_writes)
            .and_then(|value| value.checked_add(usize::from(output.is_some())))
            .ok_or(SparseError::Overflow)?,
        records: header_items,
        header_reads,
        header_writes,
        headers: header_items,
        ..SparseReport::default()
    };
    validate_report(report, limits)?;
    Ok((output, report))
}

/// Synchronize a bucket that also contains newly allocated sparse rows.
pub(super) fn rewrite_header_bucket_final_rows_with_report(
    source: &[u8],
    bucket_index: u32,
    plan: &SparsePlan,
    final_rows: &[FinalRowCount],
    limits: SparseLimits,
) -> Result<(Option<Vec<u8>>, SparseReport), SparseError> {
    if !plan.header_counts_synchronized {
        return Err(SparseError::InvalidAssignments);
    }
    for planned in plan
        .new_headers
        .iter()
        .filter(|header| header.bucket_index == bucket_index)
    {
        let final_count = final_rows
            .binary_search_by_key(&planned.row, |row| row.row)
            .ok()
            .and_then(|index| final_rows.get(index));
        if final_count.is_none_or(|count| {
            count.number_of_cells == 0 || count.number_of_cells != planned.number_of_cells
        }) {
            return Err(SparseError::InvalidAssignments);
        }
    }
    rewrite_header_bucket_rows_with_report(
        source,
        bucket_index,
        plan.columns,
        final_rows,
        Some(plan.row_header_bucket_hash_function),
        limits,
    )
}

/// Synchronize final cell counts for rows in one already materialized tile.
///
/// Zero counts delete an existing header, nonzero counts update it, and a
/// nonzero missing row is appended. All unselected bucket and header fields
/// remain byte-exact. Unlike the sparse-allocation path, no [`SparsePlan`] is
/// required because every tile and row header object already exists.
#[cfg(test)]
pub(super) fn rewrite_existing_header_bucket_final_rows_with_report(
    source: &[u8],
    bucket_index: u32,
    columns: u32,
    final_rows: &[FinalRowCount],
    limits: SparseLimits,
) -> Result<(Option<Vec<u8>>, SparseReport), SparseError> {
    let prepared = prepare_existing_header_bucket_final_rows(
        source,
        bucket_index,
        columns,
        final_rows,
        limits,
    )?;
    let prepare = prepared.prepare_report().report();
    let requirements = prepared.execution_requirements();
    let mut legacy_execution_limits = requirements.exact_limits();
    legacy_execution_limits.max_input_bytes = limits.max_work;
    legacy_execution_limits.max_output_bytes = limits.max_output_bytes;
    legacy_execution_limits.max_retained_bytes = limits.max_retained_bytes;
    legacy_execution_limits.max_retained_elements = limits.max_retained_elements;
    legacy_execution_limits.max_peak_scratch_bytes = limits.max_scratch_bytes;
    legacy_execution_limits.max_allocation_events = limits.max_allocation_events;
    legacy_execution_limits.max_fields = limits.max_fields;
    legacy_execution_limits.max_work = limits.max_work;
    legacy_execution_limits.max_headers = limits.max_records;
    let mut admitted = prepare;
    admitted.merge(header_execution_report(requirements))?;
    validate_report(admitted, limits)?;
    let (output, execution) = prepared.execute(legacy_execution_limits)?;
    let mut report = prepare;
    report.merge(execution)?;
    if report != admitted {
        return Err(SparseError::InconsistentSource);
    }
    if requirements.output_bytes() != output.as_ref().map_or(0, Vec::len) {
        return Err(SparseError::InconsistentSource);
    }
    Ok((output, report))
}

#[derive(Clone, Copy)]
struct SourceHeader<'source> {
    row: u32,
    count: u32,
    raw: &'source [u8],
    nested_fields: usize,
}

/// Build the logical existing-bucket rewrite without allocating output bytes.
pub(super) fn prepare_existing_header_bucket_final_rows<'source>(
    source: &'source [u8],
    bucket_index: u32,
    columns: u32,
    final_rows: &[FinalRowCount],
    limits: SparseLimits,
) -> Result<PreparedExistingHeaderBucketRewrite<'source>, SparseError> {
    limit(final_rows.len(), limits.max_records)?;
    if final_rows.windows(2).any(|pair| pair[0].row >= pair[1].row)
        || final_rows.iter().any(|row| {
            row.row / HEADER_BUCKET_ROWS != bucket_index || row.number_of_cells > columns
        })
    {
        return Err(SparseError::InvalidAssignments);
    }
    preflight_input(source.len(), limits)?;

    let top_count = count_fields(source, limits.max_fields)?;
    let source_count = count_numbered_fields(source, 2, limits.max_fields)?;
    limit(source_count, limits.max_records)?;
    let mut source_headers = exact_vec::<SourceHeader<'source>>(source_count)?;
    let mut hash = None;
    let mut top_position = 0usize;
    let mut fields_read = top_count;
    visit_fields(source, limits.max_fields, |field| {
        top_position = top_position.checked_add(1).ok_or(SparseError::Overflow)?;
        if field.number == 1 {
            if field.wire != 0 || hash.replace(decode_whole_varint(field.value)?).is_some() {
                return Err(SparseError::AmbiguousSource);
            }
        }
        if field.number != 2 {
            return Ok(());
        }
        if field.wire != 2 {
            return Err(SparseError::InvalidSource);
        }
        let (row, count, nested_fields) = scan_header_facts(field.value, limits.max_fields)?;
        fields_read = fields_read
            .checked_add(nested_fields)
            .ok_or(SparseError::Overflow)?;
        if row / HEADER_BUCKET_ROWS != bucket_index || count > columns {
            return Err(SparseError::InconsistentSource);
        }
        source_headers.push(SourceHeader {
            row,
            count,
            raw: field.value,
            nested_fields,
        });
        Ok(())
    })?;
    let hash = u32::try_from(hash.ok_or(SparseError::InvalidSource)?)
        .map_err(|_| SparseError::InvalidSource)?;
    if top_position != top_count || source_headers.len() != source_count {
        return Err(SparseError::InconsistentSource);
    }

    let mut sorted_rows = exact_vec::<u32>(source_count)?;
    sorted_rows.extend(source_headers.iter().map(|header| header.row));
    sorted_rows.sort_unstable();
    if sorted_rows.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(SparseError::AmbiguousSource);
    }

    let appended_count = final_rows
        .iter()
        .filter(|row| row.number_of_cells != 0 && sorted_rows.binary_search(&row.row).is_err())
        .count();
    let mut appended = exact_vec::<NewRowHeader>(appended_count)?;
    for row in final_rows
        .iter()
        .filter(|row| row.number_of_cells != 0 && sorted_rows.binary_search(&row.row).is_err())
    {
        appended.push(NewRowHeader {
            row: row.row,
            bucket_index,
            number_of_cells: row.number_of_cells,
        });
    }

    let mut fields = exact_vec::<PreparedHeaderField<'source>>(top_count)?;
    let mut source_header_position = 0usize;
    let mut header_writes = appended.len();
    let mut output_len = 0usize;
    visit_fields(source, limits.max_fields, |field| {
        if field.number != 2 {
            output_len = output_len
                .checked_add(field.raw.len())
                .ok_or(SparseError::Overflow)?;
            fields.push(PreparedHeaderField::Preserve(field.raw));
            return Ok(());
        }
        let header = *source_headers
            .get(source_header_position)
            .ok_or(SparseError::InconsistentSource)?;
        source_header_position = source_header_position
            .checked_add(1)
            .ok_or(SparseError::Overflow)?;
        let desired = final_rows
            .binary_search_by_key(&header.row, |candidate| candidate.row)
            .ok()
            .and_then(|index| final_rows.get(index));
        match desired {
            Some(desired) if desired.number_of_cells == 0 => {
                header_writes = header_writes.checked_add(1).ok_or(SparseError::Overflow)?;
                fields.push(PreparedHeaderField::HeaderDelete);
            },
            Some(desired) if desired.number_of_cells != header.count => {
                let nested_len = rewritten_header_length(header.raw, desired.number_of_cells)?;
                let encoded_len = length_field_len(2, nested_len)?;
                output_len = output_len
                    .checked_add(encoded_len)
                    .ok_or(SparseError::Overflow)?;
                header_writes = header_writes.checked_add(1).ok_or(SparseError::Overflow)?;
                fields.push(PreparedHeaderField::HeaderReplace {
                    source: header.raw,
                    row: header.row,
                    count: desired.number_of_cells,
                    encoded_len,
                    payload_len: nested_len,
                    nested_fields: header.nested_fields,
                });
            },
            Some(_) | None => {
                output_len = output_len
                    .checked_add(field.raw.len())
                    .ok_or(SparseError::Overflow)?;
                fields.push(PreparedHeaderField::HeaderPreserve {
                    raw: field.raw,
                    row: header.row,
                    count: header.count,
                    payload_len: header.raw.len(),
                    nested_fields: header.nested_fields,
                });
            },
        }
        Ok(())
    })?;
    if source_header_position != source_headers.len() {
        return Err(SparseError::InconsistentSource);
    }
    for header in &appended {
        output_len = output_len
            .checked_add(length_field_len(2, header_wire_len(*header)?)?)
            .ok_or(SparseError::Overflow)?;
    }
    if header_writes == 0 {
        output_len = 0;
    }

    let retained_elements = fields
        .len()
        .checked_add(appended.len())
        .ok_or(SparseError::Overflow)?;
    let retained_bytes = fields
        .len()
        .checked_mul(size_of::<PreparedHeaderField<'_>>())
        .and_then(|value| value.checked_add(appended.len().checked_mul(size_of::<NewRowHeader>())?))
        .ok_or(SparseError::Overflow)?;
    let temporary_bytes = source_headers
        .len()
        .checked_mul(size_of::<SourceHeader<'_>>())
        .and_then(|value| value.checked_add(sorted_rows.len().checked_mul(size_of::<u32>())?))
        .ok_or(SparseError::Overflow)?;
    let prepare_work = source
        .len()
        .checked_mul(6)
        .and_then(|value| value.checked_add(fields_read))
        .and_then(|value| value.checked_add(final_rows.len()))
        .and_then(|value| value.checked_add(source_count))
        .ok_or(SparseError::Overflow)?;
    let prepare_report = HeaderBucketPrepareReport {
        report: SparseReport {
            input_bytes: source.len(),
            fields: fields_read,
            work: prepare_work,
            retained_elements,
            retained_bytes,
            peak_scratch_bytes: retained_bytes
                .checked_add(temporary_bytes)
                .ok_or(SparseError::Overflow)?,
            allocation_events: 2usize
                .checked_mul(usize::from(source_count != 0))
                .and_then(|value| value.checked_add(usize::from(appended_count != 0)))
                .and_then(|value| value.checked_add(usize::from(top_count != 0)))
                .ok_or(SparseError::Overflow)?,
            records: source_count,
            header_reads: source_count,
            headers: source_count,
            ..SparseReport::default()
        },
    };
    validate_report(prepare_report.report(), limits)?;

    let output_headers = source_count
        .checked_sub(
            fields
                .iter()
                .filter(|field| matches!(field, PreparedHeaderField::HeaderDelete))
                .count(),
        )
        .and_then(|value| value.checked_add(appended.len()))
        .ok_or(SparseError::Overflow)?;
    let candidate_fields = if output_len == 0 {
        0
    } else {
        fields
            .len()
            .checked_sub(
                fields
                    .iter()
                    .filter(|field| matches!(field, PreparedHeaderField::HeaderDelete))
                    .count(),
            )
            .and_then(|value| value.checked_add(appended.len()))
            .and_then(|value| {
                fields
                    .iter()
                    .try_fold(value, |total, field| match field {
                        PreparedHeaderField::HeaderPreserve { nested_fields, .. }
                        | PreparedHeaderField::HeaderReplace { nested_fields, .. } => total
                            .checked_add(*nested_fields)
                            .ok_or(SparseError::Overflow),
                        PreparedHeaderField::Preserve(_) | PreparedHeaderField::HeaderDelete => {
                            Ok(total)
                        },
                    })
                    .ok()
            })
            .and_then(|value| value.checked_add(appended.len().checked_mul(4)?))
            .ok_or(SparseError::Overflow)?
    };
    let replacement_fields = fields.iter().try_fold(0usize, |total, field| match field {
        PreparedHeaderField::HeaderReplace { nested_fields, .. } => total
            .checked_add(*nested_fields)
            .ok_or(SparseError::Overflow),
        PreparedHeaderField::Preserve(_)
        | PreparedHeaderField::HeaderPreserve { .. }
        | PreparedHeaderField::HeaderDelete => Ok(total),
    })?;
    let execution_fields = candidate_fields
        .checked_add(replacement_fields)
        .ok_or(SparseError::Overflow)?;
    let candidate_nested_bytes = fields
        .iter()
        .try_fold(0usize, |total, field| match field {
            PreparedHeaderField::HeaderPreserve { payload_len, .. }
            | PreparedHeaderField::HeaderReplace { payload_len, .. } => {
                total.checked_add(*payload_len).ok_or(SparseError::Overflow)
            },
            PreparedHeaderField::Preserve(_) | PreparedHeaderField::HeaderDelete => Ok(total),
        })?
        .checked_add(appended.iter().try_fold(0usize, |total, header| {
            total
                .checked_add(header_wire_len(*header)?)
                .ok_or(SparseError::Overflow)
        })?)
        .ok_or(SparseError::Overflow)?;
    let input_bytes = if output_len == 0 {
        0
    } else {
        fields
            .iter()
            .try_fold(0usize, |total, field| {
                total
                    .checked_add(field.retained_source_bytes())
                    .ok_or(SparseError::Overflow)
            })?
            .checked_add(output_len)
            .and_then(|value| value.checked_add(candidate_nested_bytes))
            .ok_or(SparseError::Overflow)?
    };
    let execution_work = input_bytes
        .checked_add(execution_fields)
        .and_then(|value| value.checked_add(header_writes))
        .ok_or(SparseError::Overflow)?;
    let requirements = HeaderBucketExecutionRequirements {
        input_bytes,
        output_bytes: output_len,
        retained_bytes: output_len,
        retained_elements: usize::from(output_len != 0),
        peak_scratch_bytes: output_len,
        allocation_events: usize::from(output_len != 0),
        fields: execution_fields,
        work: execution_work,
        header_reads: if output_len == 0 { 0 } else { output_headers },
        header_writes: if output_len == 0 { 0 } else { header_writes },
        headers: if output_len == 0 {
            0
        } else {
            output_headers
                .checked_add(header_writes)
                .ok_or(SparseError::Overflow)?
        },
    };
    Ok(PreparedExistingHeaderBucketRewrite {
        bucket_index,
        hash,
        fields,
        appended,
        output_len,
        output_headers,
        prepare_report,
        requirements,
    })
}

fn rewrite_header_cell_count(
    source: &[u8],
    number_of_cells: u32,
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    let fields = parse_fields(source, limits.max_fields)?;
    validate_header_fields(&fields)?;
    let source_count = fields
        .iter()
        .find(|field| field.number == 4)
        .ok_or(SparseError::InvalidSource)?;
    let replacement_len = varint_field_len(4, u64::from(number_of_cells))?;
    let output_len = source
        .len()
        .checked_sub(source_count.raw.len())
        .and_then(|value| value.checked_add(replacement_len))
        .ok_or(SparseError::Overflow)?;
    limit(output_len, output_bounded_limits(limits).max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    for field in fields {
        if field.number == 4 {
            append_varint_field(&mut output, 4, u64::from(number_of_cells))?;
        } else {
            output.extend_from_slice(field.raw);
        }
    }
    if output.len() != output_len {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

// The remaining routines are small, generated-free protobuf primitives. They
// intentionally retain raw source field slices so unknown fields remain exact.

#[derive(Clone, Copy)]
struct Field<'a> {
    number: u32,
    wire: u8,
    raw: &'a [u8],
    value: &'a [u8],
}
#[derive(Clone, Copy)]
struct TileRef {
    tile_id: u32,
}
#[derive(Clone, Copy)]
struct DataStore<'a> {
    row_headers: &'a [u8],
    tiles: &'a [u8],
    row_tree: &'a [u8],
    next_row_strip_id: u32,
}
#[derive(Clone)]
struct HeaderStorage {
    hash_function: u32,
    bucket_ids: Vec<u64>,
    bucket_count: u32,
}
#[derive(Clone)]
struct HeaderBucketRows {
    hash_function: u32,
    rows: Vec<u32>,
}

fn parse_data_store(source: &[u8], max: usize) -> Result<DataStore<'_>, SparseError> {
    let fields = parse_fields(source, max)?;
    let next_row_strip_id =
        u32::try_from(unique_varint(&fields, 7)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
    Ok(DataStore {
        row_headers: unique_bytes(&fields, 1)?.ok_or(SparseError::InvalidSource)?,
        tiles: unique_bytes(&fields, 3)?.ok_or(SparseError::InvalidSource)?,
        row_tree: unique_bytes(&fields, 9)?.ok_or(SparseError::InvalidSource)?,
        next_row_strip_id,
    })
}

fn parse_tile_storage(source: &[u8], max: usize) -> Result<Vec<TileRef>, SparseError> {
    let fields = parse_fields(source, max)?;
    let mut tiles = Vec::new();
    reserve(&mut tiles, fields.len())?;
    for field in fields.into_iter().filter(|field| field.number == 1) {
        let nested = parse_fields(field.value, max)?;
        let tile_id = u32::try_from(unique_varint(&nested, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        let reference = unique_bytes(&nested, 2)?.ok_or(SparseError::InvalidSource)?;
        let _object_id = parse_local_reference(reference, max)?;
        tiles.push(TileRef { tile_id });
    }
    tiles.sort_unstable_by_key(|tile| tile.tile_id);
    if tiles
        .windows(2)
        .any(|pair| pair[0].tile_id == pair[1].tile_id)
    {
        return Err(SparseError::AmbiguousSource);
    }
    Ok(tiles)
}

fn parse_row_tree(source: &[u8], max: usize) -> Result<Vec<RowStrip>, SparseError> {
    let fields = parse_fields(source, max)?;
    let mut strips = Vec::new();
    reserve(&mut strips, fields.len())?;
    for field in fields.into_iter().filter(|field| field.number == 1) {
        let nested = parse_fields(field.value, max)?;
        let row = u32::try_from(unique_varint(&nested, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        let tile_id = u32::try_from(unique_varint(&nested, 2)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        strips.push(RowStrip { row, tile_id });
    }
    strips.sort_unstable_by_key(|strip| strip.row);
    Ok(strips)
}

fn parse_header_storage(source: &[u8], max: usize) -> Result<HeaderStorage, SparseError> {
    let fields = parse_fields(source, max)?;
    let hash_function =
        u32::try_from(unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
    let mut bucket_ids = Vec::new();
    reserve(&mut bucket_ids, fields.len())?;
    for field in fields.into_iter().filter(|field| field.number == 2) {
        bucket_ids.push(parse_local_reference(field.value, max)?);
    }
    let bucket_count = u32::try_from(bucket_ids.len()).map_err(|_| SparseError::Overflow)?;
    Ok(HeaderStorage {
        hash_function,
        bucket_ids,
        bucket_count,
    })
}

fn parse_header_bucket(source: &[u8], max: usize) -> Result<HeaderBucketRows, SparseError> {
    let fields = parse_fields(source, max)?;
    let hash_function =
        u32::try_from(unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
    let mut rows = Vec::new();
    reserve(&mut rows, fields.len())?;
    for field in fields.into_iter().filter(|field| field.number == 2) {
        let fields = parse_fields(field.value, max)?;
        validate_header_fields(&fields)?;
        let row = u32::try_from(unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        rows.push(row);
    }
    rows.sort_unstable();
    if rows.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(SparseError::AmbiguousSource);
    }
    Ok(HeaderBucketRows {
        hash_function,
        rows,
    })
}

fn parse_local_reference(source: &[u8], max: usize) -> Result<u64, SparseError> {
    let fields = parse_fields(source, max)?;
    let identifier = unique_varint(&fields, 1)?.ok_or(SparseError::InvalidSource)?;
    if identifier == 0 {
        return Err(SparseError::InvalidSource);
    }
    if let Some(external) = unique_varint(&fields, 3)? {
        if external > 1 {
            return Err(SparseError::InvalidSource);
        }
        if external == 1 {
            return Err(SparseError::InvalidSource);
        }
    }
    Ok(identifier)
}

fn validate_header_fields(fields: &[Field<'_>]) -> Result<(), SparseError> {
    let _ = unique_varint(fields, 1)?.ok_or(SparseError::InvalidSource)?;
    if fields
        .iter()
        .filter(|field| field.number == 2 && field.wire == 5)
        .count()
        != 1
    {
        return Err(SparseError::InvalidSource);
    }
    let _ = unique_varint(fields, 3)?.ok_or(SparseError::InvalidSource)?;
    let _ = unique_varint(fields, 4)?.ok_or(SparseError::InvalidSource)?;
    Ok(())
}

fn header_exists(buckets: &[HeaderBucketRows], bucket: u32, row: u32) -> Result<bool, SparseError> {
    let source = match buckets.get(usize::try_from(bucket).map_err(|_| SparseError::Overflow)?) {
        Some(value) => value,
        None => return Ok(false),
    };
    Ok(source.rows.binary_search(&row).is_ok())
}

fn ensure_sorted_cells(cells: &[Cell]) -> Result<(), SparseError> {
    if cells.windows(2).all(|pair| pair[0] < pair[1]) {
        Ok(())
    } else if cells.len() <= 1 {
        Ok(())
    } else {
        Err(SparseError::UnsortedCells)
    }
}
fn distinct_cell_counts(cells: &[Cell], tile_size: u32) -> Result<(usize, usize), SparseError> {
    let mut tiles = 0_usize;
    let mut rows = 0_usize;
    let mut previous_tile = None;
    let mut previous_row = None;
    for cell in cells {
        // The concrete tile id is irrelevant to capacity; a new tile begins
        // whenever the caller's sorted positions cross a strip boundary.  The
        // build loop later uses the checked actual tile id.
        let tile_marker = cell.row / tile_size;
        if previous_tile != Some(tile_marker) {
            tiles = tiles.checked_add(1).ok_or(SparseError::Overflow)?;
            previous_tile = Some(tile_marker);
        }
        if previous_row != Some(cell.row) {
            rows = rows.checked_add(1).ok_or(SparseError::Overflow)?;
            previous_row = Some(cell.row);
        }
    }
    Ok((tiles, rows))
}
fn ensure_sorted_strips(strips: &[RowStrip]) -> Result<(), SparseError> {
    if strips.windows(2).all(|pair| pair[0].row < pair[1].row) {
        Ok(())
    } else if strips.len() <= 1 {
        Ok(())
    } else {
        Err(SparseError::AmbiguousSource)
    }
}
fn find_tile(tiles: &[TileRef], tile_id: u32) -> Option<TileRef> {
    tiles
        .binary_search_by_key(&tile_id, |tile| tile.tile_id)
        .ok()
        .map(|index| tiles[index])
}
fn find_strip(strips: &[RowStrip], row: u32) -> Option<RowStrip> {
    strips
        .binary_search_by_key(&row, |strip| strip.row)
        .ok()
        .map(|index| strips[index])
}
fn merge_row_strips(
    existing: Vec<RowStrip>,
    missing: Vec<RowStrip>,
) -> Result<Vec<RowStrip>, SparseError> {
    let total = existing
        .len()
        .checked_add(missing.len())
        .ok_or(SparseError::Overflow)?;
    let mut merged = Vec::new();
    reserve(&mut merged, total)?;
    let mut existing_index = 0;
    let mut missing_index = 0;
    while existing_index < existing.len() || missing_index < missing.len() {
        match (existing.get(existing_index), missing.get(missing_index)) {
            (Some(left), Some(right)) if left.row < right.row => {
                merged.push(*left);
                existing_index += 1;
            },
            (Some(left), Some(right)) if left.row > right.row => {
                merged.push(*right);
                missing_index += 1;
            },
            (Some(_), Some(_)) => return Err(SparseError::AmbiguousSource),
            (Some(left), None) => {
                merged.push(*left);
                existing_index += 1;
            },
            (None, Some(right)) => {
                merged.push(*right);
                missing_index += 1;
            },
            (None, None) => break,
        }
    }
    Ok(merged)
}

fn rewrite_tile_storage(
    source: &[u8],
    new_tiles: &[NewTile],
    assignments: &[ObjectAssignment],
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    if new_tiles.len() > assignments.len() {
        return Err(SparseError::InvalidAssignments);
    }
    let extra = new_tiles
        .iter()
        .enumerate()
        .try_fold(0_usize, |total, (index, tile)| {
            total
                .checked_add(length_field_len(
                    1,
                    tile_record_len(tile.tile_id, assignments[index].object_id)?,
                )?)
                .ok_or(SparseError::Overflow)
        })?;
    let output_len = source
        .len()
        .checked_add(extra)
        .ok_or(SparseError::Overflow)?;
    limit(output_len, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    output.extend_from_slice(source);
    for (index, tile) in new_tiles.iter().enumerate() {
        append_length_prefix(
            &mut output,
            1,
            tile_record_len(tile.tile_id, assignments[index].object_id)?,
        )?;
        append_tile_record(&mut output, tile.tile_id, assignments[index].object_id)?;
    }
    if output.len() != output_len {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

fn rewrite_header_storage_append(
    source: &[u8],
    assignments: &[ObjectAssignment],
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    let mut output_len = source.len();
    for assignment in assignments {
        let reference_len = varint_field_len(1, assignment.object_id)?;
        output_len = output_len
            .checked_add(length_field_len(2, reference_len)?)
            .ok_or(SparseError::Overflow)?;
    }
    limit(output_len, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, output_len)?;
    output.extend_from_slice(source);
    for assignment in assignments {
        append_length_prefix(&mut output, 2, varint_field_len(1, assignment.object_id)?)?;
        append_varint_field(&mut output, 1, assignment.object_id)?;
    }
    if output.len() != output_len {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

fn rewrite_row_tree(
    source: &[u8],
    strips: &[RowStrip],
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    let fields = parse_fields(source, limits.max_fields)?;
    let mut existing: Vec<(RowStrip, &[u8])> = Vec::new();
    reserve(
        &mut existing,
        fields.iter().filter(|field| field.number == 1).count(),
    )?;
    for field in fields.iter().filter(|field| field.number == 1) {
        if field.wire != 2 {
            return Err(SparseError::InvalidSource);
        }
        let nested = parse_fields(field.value, limits.max_fields)?;
        let row = u32::try_from(unique_varint(&nested, 1)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        let tile_id = u32::try_from(unique_varint(&nested, 2)?.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?;
        if existing
            .last()
            .is_some_and(|(previous, _raw)| previous.row >= row)
        {
            return Err(SparseError::AmbiguousSource);
        }
        existing.push((RowStrip { row, tile_id }, field.raw));
    }
    let mut total =
        fields
            .iter()
            .filter(|field| field.number != 1)
            .try_fold(0usize, |length, field| {
                length
                    .checked_add(field.raw.len())
                    .ok_or(SparseError::Overflow)
            })?;
    let mut existing_index = 0usize;
    for strip in strips {
        let field_len = match existing.get(existing_index) {
            Some((candidate, raw)) if candidate == strip => {
                existing_index = existing_index.checked_add(1).ok_or(SparseError::Overflow)?;
                raw.len()
            },
            Some((candidate, _raw)) if candidate.row < strip.row => {
                return Err(SparseError::InconsistentSource);
            },
            Some(_) | None => length_field_len(1, row_strip_len(*strip)?)?,
        };
        total = total.checked_add(field_len).ok_or(SparseError::Overflow)?;
    }
    if existing_index != existing.len() {
        return Err(SparseError::InconsistentSource);
    }
    limit(total, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, total)?;
    let mut source_index = 0usize;
    let mut plan_index = 0usize;
    for field in fields {
        if field.number != 1 {
            output.extend_from_slice(field.raw);
            continue;
        }
        let (source_strip, raw) = existing
            .get(source_index)
            .ok_or(SparseError::InconsistentSource)?;
        while let Some(strip) = strips.get(plan_index) {
            if strip.row >= source_strip.row {
                break;
            }
            append_length_prefix(&mut output, 1, row_strip_len(*strip)?)?;
            append_row_strip(&mut output, *strip)?;
            plan_index = plan_index.checked_add(1).ok_or(SparseError::Overflow)?;
        }
        if strips.get(plan_index) != Some(source_strip) {
            return Err(SparseError::InconsistentSource);
        }
        output.extend_from_slice(raw);
        plan_index = plan_index.checked_add(1).ok_or(SparseError::Overflow)?;
        source_index = source_index.checked_add(1).ok_or(SparseError::Overflow)?;
        if source_index == existing.len() {
            for strip in &strips[plan_index..] {
                append_length_prefix(&mut output, 1, row_strip_len(*strip)?)?;
                append_row_strip(&mut output, *strip)?;
            }
            plan_index = strips.len();
        }
    }
    if source_index != existing.len() {
        return Err(SparseError::InconsistentSource);
    }
    for strip in &strips[plan_index..] {
        append_length_prefix(&mut output, 1, row_strip_len(*strip)?)?;
        append_row_strip(&mut output, *strip)?;
    }
    if output.len() != total {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

fn encode_header_bucket(
    hash: u32,
    headers: &[NewRowHeader],
    limits: SparseLimits,
) -> Result<Vec<u8>, SparseError> {
    let mut total = varint_field_len(1, u64::from(hash))?;
    for header in headers {
        total = total
            .checked_add(length_field_len(2, header_wire_len(*header)?)?)
            .ok_or(SparseError::Overflow)?;
    }
    limit(total, limits.max_output_bytes)?;
    let mut output = Vec::new();
    reserve(&mut output, total)?;
    append_varint_field(&mut output, 1, u64::from(hash))?;
    for header in headers {
        append_length_prefix(&mut output, 2, header_wire_len(*header)?)?;
        append_header(&mut output, *header)?;
    }
    if output.len() != total {
        return Err(SparseError::InconsistentSource);
    }
    Ok(output)
}

#[cfg(test)]
fn encode_tile_record(tile_id: u32, object_id: u64) -> Result<Vec<u8>, SparseError> {
    let total = tile_record_len(tile_id, object_id)?;
    let mut output = Vec::new();
    reserve(&mut output, total)?;
    append_tile_record(&mut output, tile_id, object_id)?;
    Ok(output)
}
#[cfg(test)]
fn encode_reference(object_id: u64) -> Result<Vec<u8>, SparseError> {
    let total = varint_field_len(1, object_id)?;
    let mut output = Vec::new();
    reserve(&mut output, total)?;
    append_varint_field(&mut output, 1, object_id)?;
    Ok(output)
}
#[cfg(test)]
fn encode_header(header: NewRowHeader) -> Result<Vec<u8>, SparseError> {
    let total = header_wire_len(header)?;
    let mut output = Vec::new();
    reserve(&mut output, total)?;
    append_header(&mut output, header)?;
    Ok(output)
}
fn append_tile_record(
    output: &mut Vec<u8>,
    tile_id: u32,
    object_id: u64,
) -> Result<(), SparseError> {
    append_varint_field(output, 1, u64::from(tile_id))?;
    append_length_prefix(output, 2, varint_field_len(1, object_id)?)?;
    append_varint_field(output, 1, object_id)
}
fn append_row_strip(output: &mut Vec<u8>, strip: RowStrip) -> Result<(), SparseError> {
    append_varint_field(output, 1, u64::from(strip.row))?;
    append_varint_field(output, 2, u64::from(strip.tile_id))
}
fn append_header(output: &mut Vec<u8>, header: NewRowHeader) -> Result<(), SparseError> {
    append_varint_field(output, 1, u64::from(header.row))?;
    append_fixed32_field(output, 2, 0)?;
    append_varint_field(output, 3, 0)?;
    append_varint_field(output, 4, u64::from(header.number_of_cells))
}
fn tile_record_len(tile_id: u32, object_id: u64) -> Result<usize, SparseError> {
    let reference_len = varint_field_len(1, object_id)?;
    varint_field_len(1, u64::from(tile_id))?
        .checked_add(length_field_len(2, reference_len)?)
        .ok_or(SparseError::Overflow)
}
fn row_strip_len(strip: RowStrip) -> Result<usize, SparseError> {
    varint_field_len(1, u64::from(strip.row))?
        .checked_add(varint_field_len(2, u64::from(strip.tile_id))?)
        .ok_or(SparseError::Overflow)
}
fn header_wire_len(header: NewRowHeader) -> Result<usize, SparseError> {
    varint_field_len(1, u64::from(header.row))?
        .checked_add(5)
        .ok_or(SparseError::Overflow)?
        .checked_add(varint_field_len(3, 0)?)
        .ok_or(SparseError::Overflow)?
        .checked_add(varint_field_len(4, u64::from(header.number_of_cells))?)
        .ok_or(SparseError::Overflow)
}

fn validate_assignments(
    plan: &SparsePlan,
    assignments: &[ObjectAssignment],
) -> Result<(), SparseError> {
    if assignments.len() != plan.new_objects().len() {
        return Err(SparseError::InvalidAssignments);
    }
    for (index, (request, assignment)) in plan.new_objects().iter().zip(assignments).enumerate() {
        if assignment.slot != request.slot
            || assignment.kind != request.kind
            || !assignment.metadata_registered
            || assignment.object_id == 0
            || (index != 0 && assignments[index - 1].object_id >= assignment.object_id)
        {
            return Err(SparseError::InvalidAssignments);
        }
    }
    Ok(())
}
fn parse_fields(source: &[u8], max: usize) -> Result<Vec<Field<'_>>, SparseError> {
    let mut fields = Vec::new();
    let mut offset = 0_usize;
    while offset < source.len() {
        limit(
            fields.len().checked_add(1).ok_or(SparseError::Overflow)?,
            max,
        )?;
        let start = offset;
        let key = read_varint(source, &mut offset)?;
        let number = u32::try_from(key >> 3).map_err(|_| SparseError::InvalidSource)?;
        let wire = u8::try_from(key & 7).map_err(|_| SparseError::InvalidSource)?;
        if number == 0 {
            return Err(SparseError::InvalidSource);
        }
        let value_start = offset;
        match wire {
            0 => {
                let _ = read_varint(source, &mut offset)?;
            },
            1 => offset = offset.checked_add(8).ok_or(SparseError::Overflow)?,
            2 => {
                let len = usize::try_from(read_varint(source, &mut offset)?)
                    .map_err(|_| SparseError::Overflow)?;
                let end = offset.checked_add(len).ok_or(SparseError::Overflow)?;
                if end > source.len() {
                    return Err(SparseError::InvalidSource);
                }
                offset = end;
            },
            5 => offset = offset.checked_add(4).ok_or(SparseError::Overflow)?,
            _ => return Err(SparseError::InvalidSource),
        }
        if offset > source.len() {
            return Err(SparseError::InvalidSource);
        }
        let value = match wire {
            2 => {
                let mut length_offset = value_start;
                let length = usize::try_from(read_varint(source, &mut length_offset)?)
                    .map_err(|_| SparseError::Overflow)?;
                &source[length_offset
                    ..length_offset
                        .checked_add(length)
                        .ok_or(SparseError::Overflow)?]
            },
            _ => &source[value_start..offset],
        };
        fields
            .try_reserve(1)
            .map_err(|_| SparseError::Allocation { requested: 1 })?;
        fields.push(Field {
            number,
            wire,
            raw: &source[start..offset],
            value,
        });
    }
    Ok(fields)
}
fn unique_varint(fields: &[Field<'_>], number: u32) -> Result<Option<u64>, SparseError> {
    let mut found = None;
    for field in fields.iter().filter(|field| field.number == number) {
        if field.wire != 0 || found.replace(decode_whole_varint(field.value)?).is_some() {
            return Err(SparseError::AmbiguousSource);
        }
    }
    Ok(found)
}
fn unique_bytes<'source>(
    fields: &[Field<'source>],
    number: u32,
) -> Result<Option<&'source [u8]>, SparseError> {
    let mut found = None;
    for field in fields.iter().filter(|field| field.number == number) {
        if field.wire != 2 || found.replace(field.value).is_some() {
            return Err(SparseError::AmbiguousSource);
        }
    }
    Ok(found)
}
fn read_varint(source: &[u8], offset: &mut usize) -> Result<u64, SparseError> {
    let mut value = 0_u64;
    for shift in (0..64).step_by(7) {
        let byte = *source.get(*offset).ok_or(SparseError::InvalidSource)?;
        *offset = offset.checked_add(1).ok_or(SparseError::Overflow)?;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return Ok(value);
        }
        if shift == 63 {
            break;
        }
    }
    Err(SparseError::InvalidSource)
}
fn decode_whole_varint(source: &[u8]) -> Result<u64, SparseError> {
    let mut offset = 0;
    let value = read_varint(source, &mut offset)?;
    if offset == source.len() {
        Ok(value)
    } else {
        Err(SparseError::InvalidSource)
    }
}
fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) -> Result<(), SparseError> {
    append_varint(
        output,
        u64::from(number)
            .checked_shl(3)
            .ok_or(SparseError::Overflow)?,
    );
    append_varint(output, value);
    Ok(())
}
fn append_fixed32_field(output: &mut Vec<u8>, number: u32, value: u32) -> Result<(), SparseError> {
    append_varint(output, (u64::from(number) << 3) | 5);
    output.extend_from_slice(&value.to_le_bytes());
    Ok(())
}
fn append_length(output: &mut Vec<u8>, number: u32, value: &[u8]) -> Result<(), SparseError> {
    append_length_prefix(output, number, value.len())?;
    output.extend_from_slice(value);
    Ok(())
}
fn append_length_prefix(
    output: &mut Vec<u8>,
    number: u32,
    value_len: usize,
) -> Result<(), SparseError> {
    append_varint(output, (u64::from(number) << 3) | 2);
    append_varint(
        output,
        u64::try_from(value_len).map_err(|_| SparseError::Overflow)?,
    );
    Ok(())
}
fn append_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        output.push(if value == 0 { byte } else { byte | 0x80 });
        if value == 0 {
            return;
        }
    }
}
fn varint_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}
fn varint_field_len(number: u32, value: u64) -> Result<usize, SparseError> {
    varint_len(
        u64::from(number)
            .checked_shl(3)
            .ok_or(SparseError::Overflow)?,
    )
    .checked_add(varint_len(value))
    .ok_or(SparseError::Overflow)
}
fn length_field_len(number: u32, value_len: usize) -> Result<usize, SparseError> {
    varint_len((u64::from(number) << 3) | 2)
        .checked_add(varint_len(
            u64::try_from(value_len).map_err(|_| SparseError::Overflow)?,
        ))
        .and_then(|length| length.checked_add(value_len))
        .ok_or(SparseError::Overflow)
}
fn reserve<T>(target: &mut Vec<T>, requested: usize) -> Result<(), SparseError> {
    target
        .try_reserve_exact(requested)
        .map_err(|_| SparseError::Allocation { requested })
}

fn exact_vec<T>(requested: usize) -> Result<Vec<T>, SparseError> {
    let mut target = Vec::new();
    reserve_exact_capacity(&mut target, requested)?;
    Ok(target)
}

fn reserve_exact_capacity<T>(target: &mut Vec<T>, requested: usize) -> Result<(), SparseError> {
    target
        .try_reserve_exact(requested)
        .map_err(|_| SparseError::Allocation { requested })?;
    if size_of::<T>() != 0 && target.capacity() != requested {
        return Err(SparseError::Allocation { requested });
    }
    Ok(())
}

fn ensure_header_execution_limits(
    requirements: HeaderBucketExecutionRequirements,
    limits: HeaderBucketExecutionLimits,
) -> Result<(), SparseError> {
    for (observed, maximum) in [
        (requirements.input_bytes, limits.max_input_bytes),
        (requirements.output_bytes, limits.max_output_bytes),
        (requirements.retained_bytes, limits.max_retained_bytes),
        (requirements.retained_elements, limits.max_retained_elements),
        (
            requirements.peak_scratch_bytes,
            limits.max_peak_scratch_bytes,
        ),
        (requirements.allocation_events, limits.max_allocation_events),
        (requirements.fields, limits.max_fields),
        (requirements.work, limits.max_work),
        (requirements.headers, limits.max_headers),
    ] {
        limit(observed, maximum)?;
    }
    Ok(())
}

fn header_execution_report(requirements: HeaderBucketExecutionRequirements) -> SparseReport {
    SparseReport {
        input_bytes: requirements.input_bytes,
        output_bytes: requirements.output_bytes,
        fields: requirements.fields,
        work: requirements.work,
        retained_elements: requirements.retained_elements,
        retained_bytes: requirements.retained_bytes,
        peak_scratch_bytes: requirements.peak_scratch_bytes,
        allocation_events: requirements.allocation_events,
        records: requirements.headers,
        header_reads: requirements.header_reads,
        header_writes: requirements.header_writes,
        headers: requirements.headers,
        ..SparseReport::default()
    }
}

fn visit_fields<'source>(
    source: &'source [u8],
    maximum: usize,
    mut visit: impl FnMut(Field<'source>) -> Result<(), SparseError>,
) -> Result<(), SparseError> {
    let mut offset = 0usize;
    let mut count = 0usize;
    while offset < source.len() {
        count = count.checked_add(1).ok_or(SparseError::Overflow)?;
        limit(count, maximum)?;
        visit(next_field(source, &mut offset)?)?;
    }
    Ok(())
}

fn count_numbered_fields(source: &[u8], number: u32, maximum: usize) -> Result<usize, SparseError> {
    let mut count = 0usize;
    visit_fields(source, maximum, |field| {
        if field.number == number {
            count = count.checked_add(1).ok_or(SparseError::Overflow)?;
        }
        Ok(())
    })?;
    Ok(count)
}

fn next_field<'source>(
    source: &'source [u8],
    offset: &mut usize,
) -> Result<Field<'source>, SparseError> {
    let start = *offset;
    let key = read_varint(source, offset)?;
    let number = u32::try_from(key >> 3).map_err(|_| SparseError::InvalidSource)?;
    let wire = u8::try_from(key & 7).map_err(|_| SparseError::InvalidSource)?;
    if number == 0 {
        return Err(SparseError::InvalidSource);
    }
    let value_start = *offset;
    match wire {
        0 => {
            let _ = read_varint(source, offset)?;
        },
        1 => *offset = offset.checked_add(8).ok_or(SparseError::Overflow)?,
        2 => {
            let length =
                usize::try_from(read_varint(source, offset)?).map_err(|_| SparseError::Overflow)?;
            *offset = offset.checked_add(length).ok_or(SparseError::Overflow)?;
        },
        5 => *offset = offset.checked_add(4).ok_or(SparseError::Overflow)?,
        _ => return Err(SparseError::InvalidSource),
    }
    if *offset > source.len() {
        return Err(SparseError::InvalidSource);
    }
    let value = match wire {
        2 => {
            let mut length_offset = value_start;
            let length = usize::try_from(read_varint(source, &mut length_offset)?)
                .map_err(|_| SparseError::Overflow)?;
            &source[length_offset
                ..length_offset
                    .checked_add(length)
                    .ok_or(SparseError::Overflow)?]
        },
        _ => &source[value_start..*offset],
    };
    Ok(Field {
        number,
        wire,
        raw: &source[start..*offset],
        value,
    })
}

fn scan_header_facts(source: &[u8], maximum: usize) -> Result<(u32, u32, usize), SparseError> {
    let mut row = None;
    let mut count = None;
    let mut zero = 0usize;
    let mut zero_type = 0usize;
    let mut fields = 0usize;
    visit_fields(source, maximum, |field| {
        fields = fields.checked_add(1).ok_or(SparseError::Overflow)?;
        match field.number {
            1 => {
                if field.wire != 0 || row.replace(decode_whole_varint(field.value)?).is_some() {
                    return Err(SparseError::AmbiguousSource);
                }
            },
            2 if field.wire == 5 => {
                zero = zero.checked_add(1).ok_or(SparseError::Overflow)?;
            },
            2 => {},
            3 => {
                if field.wire != 0
                    || zero_type
                        .checked_add(1)
                        .ok_or(SparseError::Overflow)
                        .is_err()
                {
                    return Err(SparseError::InvalidSource);
                }
                let _ = decode_whole_varint(field.value)?;
                zero_type = zero_type.checked_add(1).ok_or(SparseError::Overflow)?;
            },
            4 => {
                if field.wire != 0 || count.replace(decode_whole_varint(field.value)?).is_some() {
                    return Err(SparseError::AmbiguousSource);
                }
            },
            _ => {},
        }
        Ok(())
    })?;
    if zero != 1 || zero_type != 1 {
        return Err(SparseError::InvalidSource);
    }
    Ok((
        u32::try_from(row.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?,
        u32::try_from(count.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?,
        fields,
    ))
}

fn rewritten_header_length(source: &[u8], count: u32) -> Result<usize, SparseError> {
    let mut output_len = 0usize;
    let mut count_fields_seen = 0usize;
    visit_fields(source, usize::MAX, |field| {
        let field_len = if field.number == 4 {
            count_fields_seen = count_fields_seen
                .checked_add(1)
                .ok_or(SparseError::Overflow)?;
            varint_field_len(4, u64::from(count))?
        } else {
            field.raw.len()
        };
        output_len = output_len
            .checked_add(field_len)
            .ok_or(SparseError::Overflow)?;
        Ok(())
    })?;
    if count_fields_seen != 1 {
        return Err(SparseError::InvalidSource);
    }
    Ok(output_len)
}

fn append_rewritten_header(
    output: &mut Vec<u8>,
    source: &[u8],
    count: u32,
) -> Result<(), SparseError> {
    visit_fields(source, usize::MAX, |field| {
        if field.number == 4 {
            append_varint_field(output, 4, u64::from(count))?;
        } else {
            output.extend_from_slice(field.raw);
        }
        Ok(())
    })
}

fn reopen_prepared_header_bucket(
    candidate: &[u8],
    prepared: &PreparedExistingHeaderBucketRewrite<'_>,
) -> Result<(), SparseError> {
    let mut field_position = 0usize;
    let mut appended_position = 0usize;
    let mut hash = None;
    let mut observed_headers = 0usize;
    visit_fields(candidate, prepared.requirements.fields, |candidate_field| {
        if candidate_field.number == 1 {
            if candidate_field.wire != 0
                || hash
                    .replace(decode_whole_varint(candidate_field.value)?)
                    .is_some()
            {
                return Err(SparseError::AmbiguousSource);
            }
        }
        while matches!(
            prepared.fields.get(field_position),
            Some(PreparedHeaderField::HeaderDelete)
        ) {
            field_position = field_position.checked_add(1).ok_or(SparseError::Overflow)?;
        }
        if let Some(expected) = prepared.fields.get(field_position).copied() {
            field_position = field_position.checked_add(1).ok_or(SparseError::Overflow)?;
            match expected {
                PreparedHeaderField::Preserve(raw) => {
                    if candidate_field.raw != raw {
                        return Err(SparseError::InconsistentSource);
                    }
                },
                PreparedHeaderField::HeaderPreserve {
                    raw, row, count, ..
                } => {
                    if candidate_field.raw != raw {
                        return Err(SparseError::InconsistentSource);
                    }
                    let (observed_row, observed_count, _) =
                        scan_header_facts(candidate_field.value, prepared.requirements.fields)?;
                    if candidate_field.number != 2 || observed_row != row || observed_count != count
                    {
                        return Err(SparseError::InconsistentSource);
                    }
                    observed_headers = observed_headers
                        .checked_add(1)
                        .ok_or(SparseError::Overflow)?;
                },
                PreparedHeaderField::HeaderReplace {
                    row,
                    count,
                    encoded_len,
                    ..
                } => {
                    let (observed_row, observed_count, _) =
                        scan_header_facts(candidate_field.value, prepared.requirements.fields)?;
                    if candidate_field.number != 2
                        || candidate_field.raw.len() != encoded_len
                        || observed_row != row
                        || observed_count != count
                    {
                        return Err(SparseError::InconsistentSource);
                    }
                    observed_headers = observed_headers
                        .checked_add(1)
                        .ok_or(SparseError::Overflow)?;
                },
                PreparedHeaderField::HeaderDelete => {
                    return Err(SparseError::InconsistentSource);
                },
            }
            return Ok(());
        }
        let header = *prepared
            .appended
            .get(appended_position)
            .ok_or(SparseError::InconsistentSource)?;
        appended_position = appended_position
            .checked_add(1)
            .ok_or(SparseError::Overflow)?;
        let (observed_row, observed_count, _) =
            scan_header_facts(candidate_field.value, prepared.requirements.fields)?;
        if candidate_field.number != 2
            || observed_row != header.row
            || observed_count != header.number_of_cells
            || observed_row / HEADER_BUCKET_ROWS != prepared.bucket_index
        {
            return Err(SparseError::InconsistentSource);
        }
        observed_headers = observed_headers
            .checked_add(1)
            .ok_or(SparseError::Overflow)?;
        Ok(())
    })?;
    while matches!(
        prepared.fields.get(field_position),
        Some(PreparedHeaderField::HeaderDelete)
    ) {
        field_position = field_position.checked_add(1).ok_or(SparseError::Overflow)?;
    }
    if field_position != prepared.fields.len()
        || appended_position != prepared.appended.len()
        || u32::try_from(hash.ok_or(SparseError::InvalidSource)?)
            .map_err(|_| SparseError::InvalidSource)?
            != prepared.hash
        || observed_headers != prepared.output_headers
        || candidate.len() != prepared.output_len
    {
        return Err(SparseError::InconsistentSource);
    }
    Ok(())
}
fn copy_bytes(source: &[u8]) -> Result<Vec<u8>, SparseError> {
    let mut output = Vec::new();
    reserve(&mut output, source.len())?;
    output.extend_from_slice(source);
    Ok(output)
}
fn limit(observed: usize, maximum: usize) -> Result<(), SparseError> {
    if observed > maximum {
        Err(SparseError::LimitExceeded { observed, maximum })
    } else {
        Ok(())
    }
}

/// Count one raw message's fields without retaining a second field vector.
fn count_fields(source: &[u8], maximum: usize) -> Result<usize, SparseError> {
    let mut count = 0usize;
    let mut offset = 0usize;
    while offset < source.len() {
        count = count.checked_add(1).ok_or(SparseError::Overflow)?;
        limit(count, maximum)?;
        let key = read_varint(source, &mut offset)?;
        if key >> 3 == 0 {
            return Err(SparseError::InvalidSource);
        }
        match key & 7 {
            0 => {
                let _ = read_varint(source, &mut offset)?;
            },
            1 => offset = offset.checked_add(8).ok_or(SparseError::Overflow)?,
            2 => {
                let length = usize::try_from(read_varint(source, &mut offset)?)
                    .map_err(|_| SparseError::Overflow)?;
                offset = offset.checked_add(length).ok_or(SparseError::Overflow)?;
            },
            5 => offset = offset.checked_add(4).ok_or(SparseError::Overflow)?,
            _ => return Err(SparseError::InvalidSource),
        }
        if offset > source.len() {
            return Err(SparseError::InvalidSource);
        }
    }
    Ok(count)
}

fn validate_report(report: SparseReport, limits: SparseLimits) -> Result<(), SparseError> {
    if report.references
        != report
            .reference_reads
            .checked_add(report.reference_writes)
            .ok_or(SparseError::Overflow)?
        || report.headers
            != report
                .header_reads
                .checked_add(report.header_writes)
                .ok_or(SparseError::Overflow)?
    {
        return Err(SparseError::InconsistentSource);
    }
    for (observed, maximum) in [
        (report.work, limits.max_work),
        (report.retained_elements, limits.max_retained_elements),
        (report.retained_bytes, limits.max_retained_bytes),
        (report.peak_scratch_bytes, limits.max_scratch_bytes),
        (report.allocation_events, limits.max_allocation_events),
        (report.references, limits.max_references),
    ] {
        limit(observed, maximum)?;
    }
    Ok(())
}

fn preflight_input(input_bytes: usize, limits: SparseLimits) -> Result<(), SparseError> {
    limit(input_bytes, limits.max_work)
}

fn output_bounded_limits(mut limits: SparseLimits) -> SparseLimits {
    limits.max_output_bytes = limits
        .max_output_bytes
        .min(limits.max_retained_bytes)
        .min(limits.max_scratch_bytes)
        .min(limits.max_work);
    limits
}

fn scan_report(
    input_bytes: usize,
    fields: usize,
    scratch_bytes: usize,
    allocation_events: usize,
) -> Result<SparseReport, SparseError> {
    Ok(SparseReport {
        input_bytes,
        fields,
        work: input_bytes
            .checked_add(fields)
            .ok_or(SparseError::Overflow)?,
        peak_scratch_bytes: scratch_bytes,
        allocation_events,
        ..SparseReport::default()
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const TILE_SIZE: u32 = 256;

    fn reference(identifier: u64) -> Vec<u8> {
        encode_reference(identifier).expect("small reference")
    }

    fn minimal_source() -> (Vec<u8>, Vec<u8>) {
        let mut header_storage = Vec::new();
        append_varint_field(&mut header_storage, 1, 0).expect("hash");
        append_length(&mut header_storage, 2, &reference(77)).expect("bucket ref");
        let mut bucket = Vec::new();
        append_varint_field(&mut bucket, 1, 0).expect("bucket hash");

        let mut tile_storage = Vec::new();
        append_length(
            &mut tile_storage,
            1,
            &encode_tile_record(0, 30).expect("tile record"),
        )
        .expect("tile field");
        append_varint_field(&mut tile_storage, 2, u64::from(TILE_SIZE)).expect("tile size");

        let mut data_store = Vec::new();
        append_length(&mut data_store, 1, &header_storage).expect("headers");
        append_length(&mut data_store, 3, &tile_storage).expect("tiles");
        append_varint_field(&mut data_store, 7, 1).expect("next strip");
        append_length(&mut data_store, 9, &[]).expect("row tree");
        (data_store, bucket)
    }

    #[test]
    fn allocates_native_513_row_strip_shape() {
        let (data_store, bucket) = minimal_source();
        let mut cells = Vec::new();
        cells.try_reserve_exact(513).expect("test cells");
        for row in 0..=512 {
            cells.push(Cell { row, column: 0 });
        }
        let mut plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &cells,
                columns: 4,
                tile_size: TILE_SIZE,
                row_header_buckets: &[HeaderBucketSource {
                    object_id: 77,
                    payload: &bucket,
                }],
            },
            SparseLimits::default(),
        )
        .expect("valid sparse plan");
        let counts = (0..=512)
            .map(|row| FinalRowCount {
                row,
                number_of_cells: 1,
            })
            .collect::<Vec<_>>();
        plan.synchronize_new_header_counts(&counts)
            .expect("final row counts");
        assert_eq!(
            plan.row_strips(),
            &[
                RowStrip { row: 0, tile_id: 0 },
                RowStrip {
                    row: 256,
                    tile_id: 1,
                },
                RowStrip {
                    row: 512,
                    tile_id: 2,
                },
            ]
        );
        assert_eq!(
            plan.new_tiles(),
            &[
                NewTile {
                    tile_id: 1,
                    row_start: 256,
                },
                NewTile {
                    tile_id: 2,
                    row_start: 512,
                },
            ]
        );
        assert_eq!(plan.next_row_strip_id(), 3);
        assert_eq!(plan.new_headers().len(), 513);
        assert!(
            plan.new_headers()
                .iter()
                .all(|header| header.number_of_cells == 1)
        );
    }

    #[test]
    fn rows_256_and_512_keep_two_column_headers_in_lockstep() {
        let (data_store, bucket) = minimal_source();
        let cells = [
            Cell {
                row: 256,
                column: 0,
            },
            Cell {
                row: 256,
                column: 1,
            },
            Cell {
                row: 512,
                column: 0,
            },
            Cell {
                row: 512,
                column: 1,
            },
        ];
        let mut plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &cells,
                columns: 2,
                tile_size: TILE_SIZE,
                row_header_buckets: &[HeaderBucketSource {
                    object_id: 77,
                    payload: &bucket,
                }],
            },
            SparseLimits::default(),
        )
        .expect("513x2 boundary plan");
        assert_eq!(plan.new_headers().len(), 2);
        assert!(rewrite_header_bucket(&bucket, 0, &plan, SparseLimits::default()).is_err());
        plan.synchronize_new_header_counts(&[
            FinalRowCount {
                row: 256,
                number_of_cells: 2,
            },
            FinalRowCount {
                row: 512,
                number_of_cells: 2,
            },
        ])
        .expect("final counts");
        let rewritten = rewrite_header_bucket(&bucket, 0, &plan, SparseLimits::default())
            .expect("header rewrite")
            .expect("touched bucket");
        let parsed = parse_header_bucket(&rewritten, usize::MAX).expect("valid bucket");
        assert_eq!(parsed.rows, [256, 512]);
        let fields = parse_fields(&rewritten, usize::MAX).expect("bucket fields");
        let counts = fields
            .iter()
            .filter(|field| field.number == 2)
            .map(|field| {
                let header = parse_fields(field.value, usize::MAX).expect("header");
                u32::try_from(
                    unique_varint(&header, 4)
                        .expect("cell count")
                        .expect("required count"),
                )
                .expect("u32 count")
            })
            .collect::<Vec<_>>();
        assert_eq!(counts, [2, 2]);
    }

    #[test]
    fn report_separates_header_io_and_reference_items() {
        let (data_store, mut bucket) = minimal_source();
        append_length(
            &mut bucket,
            2,
            &encode_header(NewRowHeader {
                row: 0,
                bucket_index: 0,
                number_of_cells: 1,
            })
            .expect("existing header"),
        )
        .expect("existing bucket row");
        let cells = [
            Cell {
                row: 256,
                column: 0,
            },
            Cell {
                row: 512,
                column: 0,
            },
        ];
        let buckets = [HeaderBucketSource {
            object_id: 77,
            payload: &bucket,
        }];
        let (mut plan, mut report) = SparsePlan::build_with_report(
            SparseRequest {
                data_store: &data_store,
                cells: &cells,
                columns: 2,
                tile_size: TILE_SIZE,
                row_header_buckets: &buckets,
            },
            SparseLimits::default(),
        )
        .expect("reported plan");
        assert_eq!((report.reference_reads, report.reference_writes), (2, 0));
        assert_eq!(report.references, 2);
        assert_eq!((report.header_reads, report.header_writes), (1, 0));

        let synchronized = plan
            .synchronize_new_header_counts_with_report(
                &[
                    FinalRowCount {
                        row: 256,
                        number_of_cells: 1,
                    },
                    FinalRowCount {
                        row: 512,
                        number_of_cells: 1,
                    },
                ],
                SparseLimits::default(),
            )
            .expect("header synchronization");
        assert_eq!(
            (synchronized.header_reads, synchronized.header_writes),
            (0, 0)
        );
        report.merge(synchronized).expect("sync report");

        let (_header, header_report) =
            rewrite_header_bucket_with_report(&bucket, 0, &plan, SparseLimits::default())
                .expect("header rewrite");
        assert_eq!(
            (header_report.header_reads, header_report.header_writes),
            (1, 2)
        );
        report.merge(header_report).expect("header report");

        let assignments = plan
            .new_objects()
            .iter()
            .map(|request| ObjectAssignment {
                slot: request.slot,
                object_id: 100 + u64::from(request.slot),
                kind: request.kind,
                metadata_registered: true,
            })
            .collect::<Vec<_>>();
        let (_store, store_report) = rewrite_data_store_with_report(
            &data_store,
            &plan,
            &assignments,
            SparseLimits::default(),
        )
        .expect("data store rewrite");
        assert_eq!(
            (store_report.reference_reads, store_report.reference_writes),
            (0, 2)
        );
        assert_eq!(
            (store_report.header_reads, store_report.header_writes),
            (0, 0)
        );
        report.merge(store_report).expect("store report");
        assert_eq!(report.references, 4);
        assert_eq!(
            report.references,
            report.reference_reads + report.reference_writes
        );
        assert_eq!((report.header_reads, report.header_writes), (2, 2));
        assert_eq!(report.headers, report.header_reads + report.header_writes);
    }

    #[test]
    fn final_rows_materialize_missing_existing_tile_header_raw_exactly() {
        let (data_store, mut bucket) = minimal_source();
        append_length(&mut bucket, 91, b"bucket-opaque").expect("unknown bucket field");
        let unknown = parse_fields(&bucket, usize::MAX)
            .expect("bucket fields")
            .into_iter()
            .find(|field| field.number == 91)
            .expect("unknown field")
            .raw
            .to_vec();
        let buckets = [HeaderBucketSource {
            object_id: 77,
            payload: &bucket,
        }];
        let (direct, direct_report) = rewrite_existing_header_bucket_final_rows_with_report(
            &bucket,
            0,
            2,
            &[FinalRowCount {
                row: 0,
                number_of_cells: 1,
            }],
            SparseLimits::default(),
        )
        .expect("existing tile final rows");
        assert_eq!(
            (direct_report.header_reads, direct_report.header_writes),
            (1, 1)
        );
        assert_eq!(
            parse_header_bucket(&direct.expect("direct materialization"), usize::MAX)
                .expect("direct bucket")
                .rows,
            [0]
        );
        let (mut plan, _report) = SparsePlan::build_with_report(
            SparseRequest {
                data_store: &data_store,
                cells: &[Cell { row: 0, column: 0 }],
                columns: 2,
                tile_size: TILE_SIZE,
                row_header_buckets: &buckets,
            },
            SparseLimits::default(),
        )
        .expect("missing header plan");
        plan.synchronize_new_header_counts(&[FinalRowCount {
            row: 0,
            number_of_cells: 1,
        }])
        .expect("new header count");
        let (rewritten, report) = rewrite_header_bucket_final_rows_with_report(
            &bucket,
            0,
            &plan,
            &[FinalRowCount {
                row: 0,
                number_of_cells: 1,
            }],
            SparseLimits::default(),
        )
        .expect("final row header rewrite");
        let rewritten = rewritten.expect("materialized header");
        assert_eq!((report.header_reads, report.header_writes), (0, 1));
        assert_eq!(
            parse_header_bucket(&rewritten, usize::MAX)
                .expect("rewritten bucket")
                .rows,
            [0]
        );
        let rewritten_unknown = parse_fields(&rewritten, usize::MAX)
            .expect("rewritten fields")
            .into_iter()
            .find(|field| field.number == 91)
            .expect("preserved unknown")
            .raw;
        assert_eq!(rewritten_unknown, unknown);
    }

    #[test]
    fn final_rows_update_existing_row_after_missing_cell_becomes_stored() {
        let (data_store, _empty_bucket) = minimal_source();
        let mut header = encode_header(NewRowHeader {
            row: 0,
            bucket_index: 0,
            number_of_cells: 1,
        })
        .expect("existing header");
        append_length(&mut header, 90, b"header-opaque").expect("header unknown");
        let header_unknown = parse_fields(&header, usize::MAX)
            .expect("header fields")
            .into_iter()
            .find(|field| field.number == 90)
            .expect("header unknown")
            .raw
            .to_vec();
        let mut bucket = Vec::new();
        append_varint_field(&mut bucket, 1, 0).expect("hash");
        append_length(&mut bucket, 2, &header).expect("header record");
        let buckets = [HeaderBucketSource {
            object_id: 77,
            payload: &bucket,
        }];
        let plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &[Cell { row: 0, column: 1 }],
                columns: 2,
                tile_size: TILE_SIZE,
                row_header_buckets: &buckets,
            },
            SparseLimits::default(),
        )
        .expect("existing row plan");
        assert!(plan.new_headers().is_empty());
        let (rewritten, report) = rewrite_header_bucket_final_rows_with_report(
            &bucket,
            0,
            &plan,
            &[FinalRowCount {
                row: 0,
                number_of_cells: 2,
            }],
            SparseLimits::default(),
        )
        .expect("existing count update");
        let rewritten = rewritten.expect("changed header");
        assert_eq!((report.header_reads, report.header_writes), (1, 1));
        let record = parse_fields(&rewritten, usize::MAX)
            .expect("bucket fields")
            .into_iter()
            .find(|field| field.number == 2)
            .expect("header record");
        let fields = parse_fields(record.value, usize::MAX).expect("rewritten header fields");
        assert_eq!(unique_varint(&fields, 4).expect("count"), Some(2));
        assert_eq!(
            fields
                .iter()
                .find(|field| field.number == 90)
                .expect("preserved header unknown")
                .raw,
            header_unknown
        );

        let limits = SparseLimits {
            max_output_bytes: rewritten.len() - 1,
            ..SparseLimits::default()
        };
        assert!(matches!(
            rewrite_header_bucket_final_rows_with_report(
                &bucket,
                0,
                &plan,
                &[FinalRowCount {
                    row: 0,
                    number_of_cells: 2,
                }],
                limits,
            ),
            Err(SparseError::LimitExceeded { .. })
        ));
    }

    #[test]
    fn final_rows_delete_existing_header_and_preserve_unrelated_header_bytes() {
        let (data_store, _empty_bucket) = minimal_source();
        let mut first = encode_header(NewRowHeader {
            row: 0,
            bucket_index: 0,
            number_of_cells: 1,
        })
        .expect("first header");
        append_length(&mut first, 90, b"removed-opaque").expect("first unknown");
        let mut second = encode_header(NewRowHeader {
            row: 1,
            bucket_index: 0,
            number_of_cells: 1,
        })
        .expect("second header");
        append_length(&mut second, 90, b"preserved-opaque").expect("second unknown");
        let mut bucket = Vec::new();
        append_varint_field(&mut bucket, 1, 0).expect("hash");
        append_length(&mut bucket, 2, &first).expect("first record");
        append_length(&mut bucket, 2, &second).expect("second record");
        append_varint_field(&mut bucket, 91, 17).expect("bucket unknown");
        let before = parse_fields(&bucket, usize::MAX).expect("before fields");
        let second_raw = before
            .iter()
            .filter(|field| field.number == 2)
            .nth(1)
            .expect("second raw")
            .raw
            .to_vec();
        let bucket_unknown = before
            .iter()
            .find(|field| field.number == 91)
            .expect("bucket raw")
            .raw
            .to_vec();
        let buckets = [HeaderBucketSource {
            object_id: 77,
            payload: &bucket,
        }];
        let plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &[Cell { row: 0, column: 0 }],
                columns: 2,
                tile_size: TILE_SIZE,
                row_header_buckets: &buckets,
            },
            SparseLimits::default(),
        )
        .expect("existing header plan");
        let (rewritten, report) = rewrite_header_bucket_final_rows_with_report(
            &bucket,
            0,
            &plan,
            &[FinalRowCount {
                row: 0,
                number_of_cells: 0,
            }],
            SparseLimits::default(),
        )
        .expect("header deletion");
        let rewritten = rewritten.expect("changed bucket");
        assert_eq!((report.header_reads, report.header_writes), (2, 1));
        assert_eq!(
            parse_header_bucket(&rewritten, usize::MAX)
                .expect("deleted bucket")
                .rows,
            [1]
        );
        let after = parse_fields(&rewritten, usize::MAX).expect("after fields");
        assert_eq!(
            after
                .iter()
                .find(|field| field.number == 2)
                .expect("preserved second")
                .raw,
            second_raw
        );
        assert_eq!(
            after
                .iter()
                .find(|field| field.number == 91)
                .expect("preserved bucket unknown")
                .raw,
            bucket_unknown
        );
    }

    #[test]
    fn header_synchronization_refuses_duplicate_missing_and_width_counts() {
        let (data_store, bucket) = minimal_source();
        let cells = [
            Cell {
                row: 256,
                column: 0,
            },
            Cell {
                row: 512,
                column: 0,
            },
        ];
        let buckets = [HeaderBucketSource {
            object_id: 77,
            payload: &bucket,
        }];
        let request = || SparseRequest {
            data_store: &data_store,
            cells: &cells,
            columns: 2,
            tile_size: TILE_SIZE,
            row_header_buckets: &buckets,
        };
        let mut missing = SparsePlan::build(request(), SparseLimits::default()).expect("plan");
        assert_eq!(
            missing.synchronize_new_header_counts(&[FinalRowCount {
                row: 256,
                number_of_cells: 1,
            }]),
            Err(SparseError::InvalidAssignments)
        );
        let mut duplicate = SparsePlan::build(request(), SparseLimits::default()).expect("plan");
        assert_eq!(
            duplicate.synchronize_new_header_counts(&[
                FinalRowCount {
                    row: 256,
                    number_of_cells: 1,
                },
                FinalRowCount {
                    row: 256,
                    number_of_cells: 1,
                },
            ]),
            Err(SparseError::InvalidAssignments)
        );
        let mut too_wide = SparsePlan::build(request(), SparseLimits::default()).expect("plan");
        assert_eq!(
            too_wide.synchronize_new_header_counts(&[
                FinalRowCount {
                    row: 256,
                    number_of_cells: 3,
                },
                FinalRowCount {
                    row: 512,
                    number_of_cells: 1,
                },
            ]),
            Err(SparseError::InvalidAssignments)
        );
    }

    #[test]
    fn table_model_wrapper_and_empty_tile_preserve_unrelated_bytes() {
        let (data_store, _bucket) = minimal_source();
        let mut model = Vec::new();
        append_varint_field(&mut model, 91, 300).expect("unknown before");
        append_length(&mut model, 4, &data_store).expect("data store");
        append_length(&mut model, 92, b"opaque").expect("unknown after");
        assert_eq!(
            table_model_data_store(&model, SparseLimits::default()).expect("field 4"),
            data_store
        );
        let replacement = [0x08, 0x00];
        let rewritten =
            rewrite_table_model_data_store(&model, &replacement, SparseLimits::default())
                .expect("model rewrite");
        let fields = parse_fields(&rewritten, usize::MAX).expect("model fields");
        assert_eq!(
            fields[0].raw,
            parse_fields(&model, usize::MAX).unwrap()[0].raw
        );
        assert_eq!(fields[1].value, replacement);
        assert_eq!(
            fields[2].raw,
            parse_fields(&model, usize::MAX).unwrap()[2].raw
        );
    }

    #[test]
    fn assignments_refuse_duplicates_and_accept_checked_maximum_suffix() {
        let (data_store, bucket) = minimal_source();
        let cells = [
            Cell {
                row: 256,
                column: 0,
            },
            Cell {
                row: 512,
                column: 0,
            },
        ];
        let mut plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &cells,
                columns: 1,
                tile_size: TILE_SIZE,
                row_header_buckets: &[HeaderBucketSource {
                    object_id: 77,
                    payload: &bucket,
                }],
            },
            SparseLimits::default(),
        )
        .expect("plan");
        plan.synchronize_new_header_counts(&[
            FinalRowCount {
                row: 256,
                number_of_cells: 1,
            },
            FinalRowCount {
                row: 512,
                number_of_cells: 1,
            },
        ])
        .expect("counts");
        let requests = plan.new_objects();
        assert_eq!(requests.len(), 2);
        let duplicate = [
            ObjectAssignment {
                slot: requests[0].slot,
                object_id: 90,
                kind: requests[0].kind,
                metadata_registered: true,
            },
            ObjectAssignment {
                slot: requests[1].slot,
                object_id: 90,
                kind: requests[1].kind,
                metadata_registered: true,
            },
        ];
        assert_eq!(
            rewrite_data_store(&data_store, &plan, &duplicate, SparseLimits::default()),
            Err(SparseError::InvalidAssignments)
        );
        let maximum_suffix = [
            ObjectAssignment {
                slot: requests[0].slot,
                object_id: u64::MAX - 1,
                kind: requests[0].kind,
                metadata_registered: true,
            },
            ObjectAssignment {
                slot: requests[1].slot,
                object_id: u64::MAX,
                kind: requests[1].kind,
                metadata_registered: true,
            },
        ];
        rewrite_data_store(&data_store, &plan, &maximum_suffix, SparseLimits::default())
            .expect("u64 maximum suffix remains representable");
    }

    #[test]
    fn nonlocal_and_duplicate_reference_envelopes_are_refused() {
        let (_source, bucket) = minimal_source();
        let mut external_reference = reference(30);
        append_varint_field(&mut external_reference, 3, 1).expect("external marker");
        let mut record = Vec::new();
        append_varint_field(&mut record, 1, 0).expect("tile key");
        append_length(&mut record, 2, &external_reference).expect("tile reference");
        let mut tiles = Vec::new();
        append_length(&mut tiles, 1, &record).expect("tile");
        append_varint_field(&mut tiles, 2, u64::from(TILE_SIZE)).expect("tile size");
        let mut headers = Vec::new();
        append_varint_field(&mut headers, 1, 0).expect("hash");
        append_length(&mut headers, 2, &reference(77)).expect("bucket");
        let mut data_store = Vec::new();
        append_length(&mut data_store, 1, &headers).expect("headers");
        append_length(&mut data_store, 3, &tiles).expect("tiles");
        append_varint_field(&mut data_store, 7, 1).expect("next strip");
        append_length(&mut data_store, 9, &[]).expect("tree");
        assert!(matches!(
            SparsePlan::build(
                SparseRequest {
                    data_store: &data_store,
                    cells: &[],
                    columns: 1,
                    tile_size: TILE_SIZE,
                    row_header_buckets: &[HeaderBucketSource {
                        object_id: 77,
                        payload: &bucket,
                    }],
                },
                SparseLimits::default(),
            ),
            Err(SparseError::InvalidSource)
        ));

        let mut duplicate_reference = reference(30);
        append_varint_field(&mut duplicate_reference, 1, 31).expect("duplicate identifier");
        let mut duplicate_record = Vec::new();
        append_varint_field(&mut duplicate_record, 1, 0).expect("tile key");
        append_length(&mut duplicate_record, 2, &duplicate_reference).expect("tile reference");
        let mut duplicate_tiles = Vec::new();
        append_length(&mut duplicate_tiles, 1, &duplicate_record).expect("tile");
        append_varint_field(&mut duplicate_tiles, 2, u64::from(TILE_SIZE)).expect("tile size");
        assert!(matches!(
            parse_tile_storage(&duplicate_tiles, usize::MAX),
            Err(SparseError::AmbiguousSource)
        ));
    }

    #[test]
    fn row_tree_preserves_existing_and_unknown_bytes_while_appending() {
        let (_minimal, bucket) = minimal_source();
        let mut headers = Vec::new();
        append_varint_field(&mut headers, 1, 0).expect("hash");
        append_length(&mut headers, 2, &reference(77)).expect("bucket");
        let mut tiles = Vec::new();
        append_length(&mut tiles, 1, &encode_tile_record(0, 30).unwrap()).unwrap();
        append_varint_field(&mut tiles, 2, u64::from(TILE_SIZE)).unwrap();
        let mut node = Vec::new();
        append_row_strip(&mut node, RowStrip { row: 0, tile_id: 0 }).unwrap();
        append_length(&mut node, 90, b"node-opaque").unwrap();
        let mut tree = Vec::new();
        append_length(&mut tree, 1, &node).unwrap();
        append_varint_field(&mut tree, 91, 17).unwrap();
        let source_tree = tree.clone();
        let mut data_store = Vec::new();
        append_length(&mut data_store, 1, &headers).unwrap();
        append_length(&mut data_store, 3, &tiles).unwrap();
        append_varint_field(&mut data_store, 7, 1).unwrap();
        append_length(&mut data_store, 9, &tree).unwrap();
        let mut plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &[Cell {
                    row: 256,
                    column: 0,
                }],
                columns: 1,
                tile_size: TILE_SIZE,
                row_header_buckets: &[HeaderBucketSource {
                    object_id: 77,
                    payload: &bucket,
                }],
            },
            SparseLimits::default(),
        )
        .expect("plan");
        plan.synchronize_new_header_counts(&[FinalRowCount {
            row: 256,
            number_of_cells: 1,
        }])
        .unwrap();
        let request = plan.new_objects()[0];
        let rewritten = rewrite_data_store(
            &data_store,
            &plan,
            &[ObjectAssignment {
                slot: request.slot,
                object_id: 90,
                kind: request.kind,
                metadata_registered: true,
            }],
            SparseLimits::default(),
        )
        .expect("rewrite");
        let rewritten_store = parse_data_store(&rewritten.data_store, usize::MAX).unwrap();
        let before_fields = parse_fields(&source_tree, usize::MAX).unwrap();
        let after_fields = parse_fields(rewritten_store.row_tree, usize::MAX).unwrap();
        assert_eq!(after_fields[0].raw, before_fields[0].raw);
        assert_eq!(after_fields[2].raw, before_fields[1].raw);
        assert_eq!(
            parse_row_tree(rewritten_store.row_tree, usize::MAX).unwrap(),
            [
                RowStrip { row: 0, tile_id: 0 },
                RowStrip {
                    row: 256,
                    tile_id: 1,
                },
            ]
        );
    }

    #[test]
    fn allocates_positional_header_bucket_after_65536_rows() {
        let (data_store, bucket) = minimal_source();
        let mut plan = SparsePlan::build(
            SparseRequest {
                data_store: &data_store,
                cells: &[Cell {
                    row: HEADER_BUCKET_ROWS,
                    column: 0,
                }],
                columns: 1,
                tile_size: TILE_SIZE,
                row_header_buckets: &[HeaderBucketSource {
                    object_id: 77,
                    payload: &bucket,
                }],
            },
            SparseLimits::default(),
        )
        .expect("plan");
        assert_eq!(plan.new_header_buckets(), [1]);
        assert_eq!(plan.new_headers()[0].bucket_index, 1);
        plan.synchronize_new_header_counts(&[FinalRowCount {
            row: HEADER_BUCKET_ROWS,
            number_of_cells: 1,
        }])
        .unwrap();
        let assignments = plan
            .new_objects()
            .iter()
            .enumerate()
            .map(|(index, request)| ObjectAssignment {
                slot: request.slot,
                object_id: 90 + u64::try_from(index).unwrap(),
                kind: request.kind,
                metadata_registered: true,
            })
            .collect::<Vec<_>>();
        let rewritten =
            rewrite_data_store(&data_store, &plan, &assignments, SparseLimits::default())
                .expect("rewrite");
        assert_eq!(rewritten.new_header_buckets.len(), 1);
        let parsed = parse_header_bucket(&rewritten.new_header_buckets[0].1, usize::MAX).unwrap();
        assert_eq!(parsed.rows, [HEADER_BUCKET_ROWS]);
    }

    #[test]
    fn report_is_cumulative_bounded_and_linear_across_4k_to_8k() {
        fn model_with_unknown(size: usize) -> Vec<u8> {
            let mut model = Vec::new();
            append_length(&mut model, 4, &[1, 2, 3]).expect("data store");
            append_length(&mut model, 99, &vec![7; size]).expect("unknown bytes");
            model
        }
        let replacement = [8, 9, 10, 11];
        let small_source = model_with_unknown(4 * 1024);
        let large_source = model_with_unknown(8 * 1024);
        let (small, small_report) = rewrite_table_model_data_store_with_report(
            &small_source,
            &replacement,
            SparseLimits::default(),
        )
        .expect("small report");
        let (_large, large_report) = rewrite_table_model_data_store_with_report(
            &large_source,
            &replacement,
            SparseLimits::default(),
        )
        .expect("large report");
        assert_eq!(small_report.output_bytes, small.len());
        assert!(large_report.work <= small_report.work * 220 / 100);

        let mut total = small_report;
        total
            .merge(large_report)
            .expect("checked cumulative report");
        assert_eq!(
            total.input_bytes,
            small_report.input_bytes + large_report.input_bytes
        );
        assert_eq!(
            total.output_bytes,
            small_report.output_bytes + large_report.output_bytes
        );
    }

    #[test]
    fn report_max_minus_one_refuses_rewrite() {
        let source = {
            let mut value = Vec::new();
            append_length(&mut value, 4, &[1, 2, 3]).expect("data store");
            value
        };
        let replacement = [4, 5, 6, 7, 8];
        let (output, report) = rewrite_table_model_data_store_with_report(
            &source,
            &replacement,
            SparseLimits::default(),
        )
        .expect("baseline");
        assert_eq!(report.retained_bytes, output.len());
        let limits = SparseLimits {
            max_retained_bytes: report.retained_bytes - 1,
            ..SparseLimits::default()
        };
        assert!(matches!(
            rewrite_table_model_data_store_with_report(&source, &replacement, limits),
            Err(SparseError::LimitExceeded { .. })
        ));
    }

    fn prepared_existing_header_fixture() -> (Vec<u8>, [FinalRowCount; 3]) {
        let mut first = encode_header(NewRowHeader {
            row: 0,
            bucket_index: 0,
            number_of_cells: 1,
        })
        .expect("first header");
        append_length(&mut first, 90, b"removed-opaque").expect("first unknown");
        let mut second = encode_header(NewRowHeader {
            row: 1,
            bucket_index: 0,
            number_of_cells: 1,
        })
        .expect("second header");
        append_length(&mut second, 91, b"preserved-opaque").expect("second unknown");
        let mut bucket = Vec::new();
        append_varint_field(&mut bucket, 1, 0).expect("hash");
        append_length(&mut bucket, 2, &first).expect("first row");
        append_length(&mut bucket, 2, &second).expect("second row");
        append_varint_field(&mut bucket, 92, 19).expect("bucket unknown");
        (
            bucket,
            [
                FinalRowCount {
                    row: 0,
                    number_of_cells: 0,
                },
                FinalRowCount {
                    row: 1,
                    number_of_cells: 2,
                },
                FinalRowCount {
                    row: 2,
                    number_of_cells: 1,
                },
            ],
        )
    }

    #[test]
    fn prepared_existing_header_is_output_free_and_executes_exactly() {
        let (source, final_rows) = prepared_existing_header_fixture();
        let prepared = prepare_existing_header_bucket_final_rows(
            &source,
            0,
            3,
            &final_rows,
            SparseLimits::default(),
        )
        .expect("output-free header plan");
        let prepare = prepared.prepare_report().report();
        assert_eq!(prepare.output_bytes, 0);
        assert!(prepare.retained_bytes > 0);
        let requirements = prepared.execution_requirements();
        assert!(requirements.input_bytes > 0);
        assert!(requirements.output_bytes > 0);
        assert!(requirements.retained_bytes > 0);
        assert!(requirements.retained_elements > 0);
        assert!(requirements.peak_scratch_bytes > 0);
        assert!(requirements.allocation_events > 0);
        assert!(requirements.fields > 0);
        assert!(requirements.work > 0);
        assert!(requirements.headers > 0);

        reset_prepared_header_execution_allocations();
        let (output, report) = prepared
            .execute(requirements.exact_limits())
            .expect("exact execution");
        assert_eq!(prepared_header_execution_allocations(), 1);
        let output = output.expect("changed bucket");
        assert_eq!(output.len(), requirements.output_bytes);
        assert_eq!(report.output_bytes, requirements.output_bytes);
        assert_eq!(report.retained_bytes, requirements.retained_bytes);
        assert_eq!(report.retained_elements, requirements.retained_elements);
        assert_eq!(report.peak_scratch_bytes, requirements.peak_scratch_bytes);
        assert_eq!(report.allocation_events, requirements.allocation_events);
        assert_eq!(report.fields, requirements.fields);
        assert_eq!(report.work, requirements.work);
        assert_eq!(report.header_writes, requirements.header_writes);
        let fields = parse_fields(&output, usize::MAX).expect("candidate fields");
        let rows_and_counts = fields
            .iter()
            .filter(|field| field.number == 2)
            .map(|field| {
                let (row, count, _) =
                    scan_header_facts(field.value, usize::MAX).expect("header facts");
                (row, count)
            })
            .collect::<Vec<_>>();
        assert_eq!(rows_and_counts, [(1, 2), (2, 1)]);
    }

    #[test]
    fn prepared_existing_header_every_axis_max_minus_one_is_preallocation() {
        let (source, final_rows) = prepared_existing_header_fixture();
        let baseline = prepare_existing_header_bucket_final_rows(
            &source,
            0,
            3,
            &final_rows,
            SparseLimits::default(),
        )
        .expect("baseline plan")
        .execution_requirements();
        for axis in 0..9 {
            let prepared = prepare_existing_header_bucket_final_rows(
                &source,
                0,
                3,
                &final_rows,
                SparseLimits::default(),
            )
            .expect("axis plan");
            let mut limits = baseline.exact_limits();
            match axis {
                0 => limits.max_input_bytes -= 1,
                1 => limits.max_output_bytes -= 1,
                2 => limits.max_retained_bytes -= 1,
                3 => limits.max_retained_elements -= 1,
                4 => limits.max_peak_scratch_bytes -= 1,
                5 => limits.max_allocation_events -= 1,
                6 => limits.max_fields -= 1,
                7 => limits.max_work -= 1,
                8 => limits.max_headers -= 1,
                _ => unreachable!(),
            }
            reset_prepared_header_execution_allocations();
            assert!(matches!(
                prepared.execute(limits),
                Err(SparseError::LimitExceeded { .. })
            ));
            assert_eq!(
                prepared_header_execution_allocations(),
                0,
                "axis {axis} allocated a candidate"
            );
        }
    }
}
