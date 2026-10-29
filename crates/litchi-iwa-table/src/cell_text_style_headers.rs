//! Bounded row/column header-count publication for cell text-style writes.
//!
//! The tile writer reports the previous bytes of every selected cell.  A
//! `None` previous value means that a style-only BNC cell was materialized;
//! existing cells do not change the external row/column cell counts.  This
//! module turns that fact into source-relative header deltas and applies the
//! existing raw `HeaderStorageBucket` writer under the same residual budget.
//! The package owner still locates and replaces the selected objects.

use litchi_iwa_common::{LimitKind, wire::RawWireFields, wire::RawWireLimits};
use litchi_numbers_wire::{Limits as WireProfile, table_write};

use crate::{cell_text_style, cell_text_style_writer::Update};

/// Sorted, unique cell-count changes for the selected table's row and column
/// header buckets.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MaterializedHeaderDeltas {
    rows: Vec<(u32, i64)>,
    columns: Vec<(u32, i64)>,
}

impl MaterializedHeaderDeltas {
    /// Return global row-index deltas.
    #[must_use]
    pub fn rows(&self) -> &[(u32, i64)] {
        &self.rows
    }

    /// Return table-column-index deltas.
    #[must_use]
    pub fn columns(&self) -> &[(u32, i64)] {
        &self.columns
    }

    /// Whether no external header count changes are required.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.rows.is_empty() && self.columns.is_empty()
    }
}

/// Header-bucket payloads that must be replaced by the package owner.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct HeaderBucketRewrite {
    row: Option<Vec<u8>>,
    column: Option<Vec<u8>>,
}

impl HeaderBucketRewrite {
    /// Rewritten row-header payload, if a materialized cell changed it.
    #[must_use]
    pub fn row(&self) -> Option<&[u8]> {
        self.row.as_deref()
    }

    /// Rewritten column-header payload, if a materialized cell changed it.
    #[must_use]
    pub fn column(&self) -> Option<&[u8]> {
        self.column.as_deref()
    }

    /// Whether either header bucket needs replacement.
    #[must_use]
    pub const fn changed(&self) -> bool {
        self.row.is_some() || self.column.is_some()
    }

    /// Consume the rewrite and return its optional payloads.
    #[must_use]
    pub fn into_parts(self) -> (Option<Vec<u8>>, Option<Vec<u8>>) {
        (self.row, self.column)
    }
}

fn invalid(reason: &'static str) -> cell_text_style::Error {
    cell_text_style::Error::InvalidSource(reason)
}

fn map_common(error: litchi_iwa_common::Error) -> cell_text_style::Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => cell_text_style::Error::Limit {
            kind,
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            cell_text_style::Error::Allocation { amount }
        },
        _ => invalid("malformed header-bucket wire payload"),
    }
}

fn map_wire_kind(kind: litchi_numbers_wire::LimitKind) -> LimitKind {
    match kind {
        litchi_numbers_wire::LimitKind::InputBytes | litchi_numbers_wire::LimitKind::TailBytes => {
            LimitKind::InputBytes
        },
        litchi_numbers_wire::LimitKind::OutputBytes => LimitKind::OutputBytes,
        litchi_numbers_wire::LimitKind::Fields => LimitKind::Fields,
        litchi_numbers_wire::LimitKind::Items => LimitKind::RewriteWork,
        _ => LimitKind::RewriteWork,
    }
}

fn map_table_write(error: table_write::WriteError) -> cell_text_style::Error {
    match error {
        table_write::WriteError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => cell_text_style::Error::Limit {
            kind: map_wire_kind(kind),
            observed: observed as u64,
            maximum: maximum as u64,
        },
        table_write::WriteError::Allocation { requested } => {
            cell_text_style::Error::Allocation { amount: requested }
        },
        table_write::WriteError::Invalid | table_write::WriteError::Unsupported => {
            invalid("header-bucket shape is outside the shared writer")
        },
        _ => invalid("header-bucket writer failed"),
    }
}

fn checked_add(
    left: usize,
    right: usize,
    reason: &'static str,
) -> Result<usize, cell_text_style::Error> {
    left.checked_add(right).ok_or_else(|| invalid(reason))
}

fn checked_mul(
    left: usize,
    right: usize,
    reason: &'static str,
) -> Result<usize, cell_text_style::Error> {
    left.checked_mul(right).ok_or_else(|| invalid(reason))
}

fn aggregate(
    deltas: &mut Vec<(u32, i64)>,
    limits: &cell_text_style::Limits,
) -> Result<(), cell_text_style::Error> {
    if deltas.len() > 1 {
        let sort_work = checked_mul(
            deltas.len(),
            usize::BITS as usize,
            "header delta work overflow",
        )?;
        limits.admit_work(sort_work)?;
    }
    deltas.sort_unstable_by_key(|(index, _)| *index);
    limits.admit_work(deltas.len())?;
    let mut write = 0;
    for read in 0..deltas.len() {
        let (index, delta) = deltas[read];
        if write != 0 && deltas[write - 1].0 == index {
            deltas[write - 1].1 = deltas[write - 1]
                .1
                .checked_add(delta)
                .ok_or_else(|| invalid("header delta overflow"))?;
        } else {
            deltas[write] = (index, delta);
            write += 1;
        }
    }
    deltas.truncate(write);
    Ok(())
}

/// Convert tile-local style-key materializations into global row and column
/// header deltas.
///
/// `row_offset` is the global row represented by tile-local row zero.  A
/// delta is emitted only for an absent previous cell receiving a non-empty
/// style key.  Existing cells, clears of absent cells, and style-key changes
/// do not change header counts.
pub fn materialized_deltas(
    updates: &[Update],
    previous: &[Option<Vec<u8>>],
    row_offset: u32,
    limits: &cell_text_style::Limits,
) -> Result<MaterializedHeaderDeltas, cell_text_style::Error> {
    if updates.len() != previous.len() {
        return Err(invalid("style update and previous-cell lengths differ"));
    }
    if updates.is_empty() {
        return Ok(MaterializedHeaderDeltas::default());
    }

    // The count pass and the coordinate/materialization pass are distinct
    // bounded traversals.  Keep both debits so a caller cannot spend one
    // pass's budget on the other.
    limits.admit_work(updates.len())?;
    let mut coordinates = Vec::new();
    limits.reserve(&mut coordinates, updates.len())?;
    let materialized = updates
        .iter()
        .zip(previous)
        .filter(|(update, before)| before.is_none() && update.key.is_some())
        .count();
    limits.admit_work(updates.len())?;
    let mut rows = Vec::new();
    let mut columns = Vec::new();
    limits.reserve(&mut rows, materialized)?;
    limits.reserve(&mut columns, materialized)?;

    for (update, before) in updates.iter().zip(previous) {
        let row = row_offset
            .checked_add(update.row)
            .ok_or_else(|| invalid("global header row overflows u32"))?;
        coordinates.push((row, update.column));
        if before.is_none() && update.key.is_some() {
            rows.push((row, 1));
            columns.push((update.column, 1));
        }
    }
    let coordinate_sort_work = checked_mul(
        coordinates.len(),
        usize::BITS as usize,
        "header coordinate sort work overflow",
    )?;
    limits.admit_work(coordinate_sort_work)?;
    coordinates.sort_unstable();
    limits.admit_work(coordinates.len())?;
    if coordinates.windows(2).any(|pair| pair[0] == pair[1]) {
        return Err(invalid("duplicate style update coordinate"));
    }
    aggregate(&mut rows, limits)?;
    aggregate(&mut columns, limits)?;
    Ok(MaterializedHeaderDeltas { rows, columns })
}

#[derive(Debug, Clone, Copy, Default)]
struct WireCensus {
    fields: usize,
    work: usize,
}

fn census_bucket(
    bucket: &[u8],
    limits: &cell_text_style::Limits,
) -> Result<WireCensus, cell_text_style::Error> {
    let wire = limits.residual_wire()?;
    limits.charge_input(bucket.len())?;
    let raw_limits = RawWireLimits::new(
        wire.max_input_bytes(),
        wire.max_fields(),
        wire.max_nesting(),
        wire.max_rewrite_work(),
    )
    .map_err(map_common)?;
    let mut root = RawWireFields::new(bucket, raw_limits);
    let mut fields = 0usize;
    let mut work = 0usize;
    while let Some(field) = root.next().map_err(map_common)? {
        if field.number() != 2 {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(invalid("header bucket entry has the wrong wire type"));
        }
        let mut header = RawWireFields::new(field.payload(), raw_limits);
        while header.next().map_err(map_common)?.is_some() {}
        fields = checked_add(fields, header.fields(), "header field count overflow")?;
        work = checked_add(work, header.work(), "header wire work overflow")?;
    }
    fields = checked_add(fields, root.fields(), "header field count overflow")?;
    work = checked_add(work, root.work(), "header wire work overflow")?;
    limits.charge_fields(fields)?;
    limits.admit_work(work)?;
    Ok(WireCensus { fields, work })
}

fn rewrite_one(
    bucket: &[u8],
    deltas: &[(u32, i64)],
    limits: &cell_text_style::Limits,
) -> Result<Vec<u8>, cell_text_style::Error> {
    if deltas.is_empty() {
        return Err(invalid("empty header rewrite"));
    }
    let census = census_bucket(bucket, limits)?;
    // This is the profile reserved for the raw helper.  Capture it after
    // preflight, before debiting the exact execution requirements below; the
    // helper's own finite checks are covered by those debits and must not be
    // given the further-depleted leftover profile.
    let execution_wire = limits.residual_wire()?;
    let growth = checked_mul(
        deltas.len(),
        table_write::HEADER_BYTES,
        "header output bound overflow",
    )?;
    let output_bound = checked_add(bucket.len(), growth, "header output bound overflow")?;
    let execution_work = checked_add(census.work, census.fields, "header rewrite work overflow")?;
    let execution_work = checked_add(execution_work, output_bound, "header rewrite work overflow")?;
    limits.charge_input(bucket.len())?;
    limits.charge_fields(census.fields)?;
    limits.admit_work(execution_work)?;
    limits.admit_output(output_bound)?;
    limits.admit_copies(output_bound)?;
    limits.admit_allocations(1)?;
    let profile = WireProfile::from_wire_limits(execution_wire);
    table_write::header_adjust_many(bucket, deltas, &profile).map_err(map_table_write)
}

/// Rewrite the row and column header buckets using the same residual budget
/// as the tile write.  Empty deltas return without parsing or allocating, so
/// existing-cell style changes remain exact header no-ops.
pub fn rewrite_header_buckets(
    row_bucket: &[u8],
    column_bucket: &[u8],
    deltas: &MaterializedHeaderDeltas,
    limits: &cell_text_style::Limits,
) -> Result<HeaderBucketRewrite, cell_text_style::Error> {
    let row = (!deltas.rows.is_empty())
        .then(|| rewrite_one(row_bucket, &deltas.rows, limits))
        .transpose()?;
    let column = (!deltas.columns.is_empty())
        .then(|| rewrite_one(column_bucket, &deltas.columns, limits))
        .transpose()?;
    Ok(HeaderBucketRewrite { row, column })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::TableSource;
    use litchi_iwa_common::{
        WireLimits,
        wire::{append_length_delimited_field, append_varint_field},
    };
    use prost::Message as _;

    fn updates() -> Vec<Update> {
        vec![
            Update {
                row: 2,
                column: 4,
                key: Some(2),
            },
            Update {
                row: 2,
                column: 3,
                key: Some(2),
            },
            Update {
                row: 3,
                column: 3,
                key: None,
            },
        ]
    }

    #[test]
    fn delta_extraction_preserves_global_rows_and_aggregates_columns() {
        let updates = updates();
        let previous = vec![None, None, None];
        // This test intentionally uses a local finite profile only for the
        // delta phase; the rewrite tests below exercise the full source-based
        // residual profile.
        let source = test_source();
        let limits = cell_text_style::Limits::from_source(&source).expect("limits");
        let before = limits.report().work;
        let deltas = materialized_deltas(&updates, &previous, 256, &limits).expect("deltas");
        let after = limits.report().work;
        assert!(after >= before + updates.len() * 2);
        assert_eq!(deltas.rows(), &[(258, 2)]);
        assert_eq!(deltas.columns(), &[(3, 1), (4, 1)]);
        let second = materialized_deltas(&updates, &previous, 256, &limits).expect("deltas");
        assert_eq!(second, deltas);
        assert!(limits.report().work > after);
    }

    #[test]
    fn existing_cells_and_absent_clears_are_header_noops() {
        let updates = updates();
        let previous = vec![Some(vec![1]), Some(vec![2]), None];
        let source = test_source();
        let limits = cell_text_style::Limits::from_source(&source).expect("limits");
        let deltas = materialized_deltas(&updates, &previous, 0, &limits).expect("deltas");
        assert!(deltas.is_empty());
    }

    #[test]
    fn rewrite_preserves_unknowns_and_creates_missing_headers() {
        let source = test_source();
        let limits = cell_text_style::Limits::from_source(&source).expect("limits");
        let bucket = header_bucket(&[(258, 1)], true);
        let deltas = MaterializedHeaderDeltas {
            rows: vec![(258, 1), (259, 1)],
            columns: Vec::new(),
        };
        let output = rewrite_header_buckets(&bucket, &bucket, &deltas, &limits)
            .expect("rewrite")
            .row
            .expect("row rewrite");
        let decoded = litchi_iwa_protos::tst::HeaderStorageBucket::decode(output.as_slice())
            .expect("header bucket");
        assert_eq!(
            decoded
                .headers
                .iter()
                .map(|header| (header.index, header.number_of_cells))
                .collect::<Vec<_>>(),
            vec![(258, 2), (259, 1)]
        );
        assert!(output.windows(3).any(|window| window == [0xd0, 0x05, 0x7b]));
        assert!(output.windows(3).any(|window| window == [0xd8, 0x05, 0x07]));
    }

    #[test]
    fn overflow_and_zero_budget_fail_before_rewrite() {
        let source = test_source();
        let limits = cell_text_style::Limits::from_source(&source).expect("limits");
        let updates = [Update {
            row: 1,
            column: 0,
            key: Some(1),
        }];
        let previous = [None];
        assert!(matches!(
            materialized_deltas(&updates, &previous, u32::MAX, &limits),
            Err(cell_text_style::Error::InvalidSource(_))
        ));

        let zero = cell_text_style::Limits::from_source(&source).expect("limits");
        zero.admit_work(zero.max_work).expect("exhaust work");
        let deltas = MaterializedHeaderDeltas {
            rows: vec![(1, 1)],
            columns: Vec::new(),
        };
        assert!(matches!(
            rewrite_header_buckets(&header_bucket(&[], false), &[], &deltas, &zero),
            Err(cell_text_style::Error::Limit {
                kind: LimitKind::RewriteWork,
                ..
            })
        ));
    }

    #[test]
    fn cumulative_allocation_budget_covers_repeated_bucket_rewrites() {
        let source = test_source();
        let limits = cell_text_style::Limits::from_source(&source).expect("limits");
        limits
            .admit_allocations(limits.max_allocations.saturating_sub(1))
            .expect("leave one rewrite allocation");
        let deltas = MaterializedHeaderDeltas {
            rows: vec![(1, 1)],
            columns: Vec::new(),
        };
        let bucket = header_bucket(&[], false);
        rewrite_header_buckets(&bucket, &bucket, &deltas, &limits).expect("first rewrite");
        assert!(matches!(
            rewrite_header_buckets(&bucket, &bucket, &deltas, &limits),
            Err(cell_text_style::Error::Allocation { .. })
        ));
    }

    fn header_bucket(headers: &[(u32, u64)], unknown: bool) -> Vec<u8> {
        let mut bucket = Vec::new();
        append_varint_field(&mut bucket, 1, 1).expect("bucket hash");
        for &(index, cells) in headers {
            let mut header = Vec::new();
            append_varint_field(&mut header, 1, u64::from(index)).expect("index");
            header.extend_from_slice(&[0x15, 0, 0, 0, 0]);
            append_varint_field(&mut header, 3, 0).expect("hiding");
            append_varint_field(&mut header, 4, cells).expect("count");
            if unknown {
                append_varint_field(&mut header, 90, 123).expect("header unknown");
            }
            append_length_delimited_field(&mut bucket, 2, &header, WireLimits::DEFAULT)
                .expect("header");
        }
        if unknown {
            append_varint_field(&mut bucket, 91, 7).expect("unknown");
        }
        bucket
    }

    struct TestSource {
        catalog: litchi_iwa_archive::SourceCatalog,
    }

    impl TableSource for TestSource {
        fn catalog(&self) -> &litchi_iwa_archive::ComponentCatalog {
            self.catalog.components()
        }

        fn physical(&self) -> Option<&litchi_iwa_archive::SourceCatalog> {
            Some(&self.catalog)
        }

        fn resolve_object(
            &self,
            _identifier: u64,
        ) -> Result<Option<crate::ResolvedObject<'_>>, crate::ResolveError> {
            Ok(None)
        }

        fn source_bytes(&self) -> &[u8] {
            self.catalog.source_bytes()
        }

        fn archive_limits(&self) -> litchi_iwa_archive::Limits {
            litchi_iwa_archive::Limits::default()
        }

        fn max_objects(&self) -> usize {
            1_000_000
        }

        fn max_references(&self) -> usize {
            1_000_000
        }
    }

    fn test_source() -> TestSource {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers");
        let bytes = std::fs::read(path).expect("fixture");
        TestSource {
            catalog: litchi_iwa_archive::SourceCatalog::from_bytes(
                &bytes,
                litchi_iwa_archive::Limits::default(),
            )
            .expect("catalog"),
        }
    }
}
