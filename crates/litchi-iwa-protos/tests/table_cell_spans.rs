//! Independent coverage for the borrowed Numbers tile-row cell-span helper.
//!
//! These tests exercise the shared offset contract without constructing an
//! IWA archive.  The helper must validate the complete offset buffer before a
//! caller can traverse a span, retain only borrowed source slices, and keep
//! the bounded work report deterministic for both narrow and wide rows.

use litchi_iwa_protos::numbers_table_cell_storage_codec::{
    CellSpans, DecodeError, DecodeLimit, DecodeOptions, decode_tile_row_info,
};

fn options(source: &[u8]) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        usize::MAX,
        usize::MAX,
        64,
        usize::MAX,
        usize::MAX,
    )
}

fn limited_options(source: &[u8], fields: usize, work: usize) -> DecodeOptions {
    DecodeOptions::new(
        source.len().max(1),
        fields,
        work,
        64,
        usize::MAX,
        usize::MAX,
    )
}

fn offsets(values: &[Option<u16>]) -> Vec<u8> {
    let mut output = Vec::with_capacity(values.len() * 2);
    for value in values {
        output.extend_from_slice(&value.unwrap_or(u16::MAX).to_le_bytes());
    }
    output
}

fn assert_spans(spans: CellSpans<'_>, expected: &[(usize, usize, usize)]) {
    let actual: Vec<_> = spans
        .iter()
        .map(|span| (span.column(), span.start(), span.end()))
        .collect();
    assert_eq!(actual, expected);
    assert_eq!(spans.len(), expected.len());
    assert_eq!(spans.is_empty(), expected.is_empty());
}

fn push_varint_value(bytes: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        bytes.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn push_varint(bytes: &mut Vec<u8>, field: u32, value: u64) {
    push_varint_value(bytes, u64::from(field) << 3);
    push_varint_value(bytes, value);
}

fn push_bytes(bytes: &mut Vec<u8>, field: u32, value: &[u8]) {
    push_varint_value(bytes, (u64::from(field) << 3) | 2);
    push_varint_value(
        bytes,
        u64::try_from(value.len()).expect("test payload length fits in u64"),
    );
    bytes.extend_from_slice(value);
}

fn tile_row(
    cell_count: u32,
    pre_storage: &[u8],
    pre_offsets: &[u8],
    modern_storage: Option<&[u8]>,
    modern_offsets: Option<&[u8]>,
    wide: Option<bool>,
) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_varint(&mut bytes, 1, 0);
    push_varint(&mut bytes, 2, u64::from(cell_count));
    push_bytes(&mut bytes, 3, pre_storage);
    push_bytes(&mut bytes, 4, pre_offsets);
    if let Some(storage) = modern_storage {
        push_bytes(&mut bytes, 6, storage);
    }
    if let Some(offsets) = modern_offsets {
        push_bytes(&mut bytes, 7, offsets);
    }
    if let Some(wide) = wide {
        push_varint(&mut bytes, 8, u64::from(wide));
    }
    bytes
}

fn assert_invalid(
    source: &[u8],
    storage_length: usize,
    wide: bool,
    expected: usize,
    columns: usize,
) {
    assert!(
        CellSpans::parse(
            source,
            storage_length,
            wide,
            expected,
            columns,
            options(source)
        )
        .is_err(),
        "malformed cell offsets were accepted: {source:?}"
    );
}

#[test]
fn narrow_spans_cover_holes_and_borrow_storage_without_allocating_ranges() {
    let offsets = offsets(&[None, Some(0), Some(24), Some(48), None]);
    let storage: Vec<u8> = (0..72).collect();
    let (spans, report) =
        CellSpans::parse(&offsets, storage.len(), false, 3, 5, options(&offsets)).unwrap();

    assert_spans(spans, &[(1, 0, 24), (2, 24, 48), (3, 48, 72)]);
    assert_eq!(spans.column_count(), 5);
    assert_eq!(spans.storage_length(), 72);
    assert_eq!(report.source_bytes(), offsets.len());
    assert_eq!(report.fields(), 0);
    assert_eq!(report.work_bytes(), offsets.len() * 2);
    assert_eq!(report.references(), 0);
    assert_eq!(report.text_bytes(), 0);

    assert_eq!(spans.get(0), None);
    assert_eq!(
        spans
            .get(2)
            .map(|span| (span.column(), span.start(), span.end())),
        Some((2, 24, 48))
    );

    for span in spans.iter() {
        let payload = span.bytes(&storage).expect("validated span fits storage");
        assert_eq!(payload.len(), span.end() - span.start());
        assert_eq!(
            payload.as_ptr() as usize,
            storage.as_ptr() as usize + span.start()
        );
    }
}

#[test]
fn wide_spans_scale_offsets_in_four_byte_units() {
    let offsets = offsets(&[Some(0), None, Some(6), None]);
    let storage = [0_u8; 48];
    let (spans, report) =
        CellSpans::parse(&offsets, storage.len(), true, 2, 4, options(&offsets)).unwrap();

    assert_spans(spans, &[(0, 0, 24), (2, 24, 48)]);
    assert_eq!(report.work_bytes(), offsets.len() * 2);
    let mut iter = spans.iter();
    assert_eq!(iter.len(), 2);
    assert_eq!(iter.next().map(|span| span.range()), Some(0..24));
    assert_eq!(iter.len(), 1);
    assert_eq!(iter.next().map(|span| span.range()), Some(24..48));
    assert_eq!(iter.len(), 0);
    assert_eq!(iter.next(), None);
}

#[test]
fn empty_rows_and_trailing_missing_slots_are_valid() {
    let empty = offsets(&[None, None, None]);
    let (spans, report) = CellSpans::parse(&empty, 0, false, 0, 3, options(&empty)).unwrap();
    assert_spans(spans, &[]);
    assert_eq!(report.work_bytes(), empty.len() * 2);
    assert_eq!(spans.get(0), None);

    let padded = offsets(&[Some(0), None, None, None]);
    let (spans, _) = CellSpans::parse(&padded, 1, false, 1, 2, options(&padded)).unwrap();
    assert_spans(spans, &[(0, 0, 1)]);
}

#[test]
fn complete_offset_validation_happens_before_any_span_is_published() {
    let malformed_later_slot = offsets(&[Some(0), Some(2), Some(1)]);
    let mut published = None;
    let result = CellSpans::parse(
        &malformed_later_slot,
        3,
        false,
        3,
        3,
        options(&malformed_later_slot),
    )
    .map(|(spans, _report)| {
        let cells: Vec<_> = spans.iter().collect();
        published = Some(cells);
    });

    assert!(result.is_err());
    assert!(
        published.is_none(),
        "a malformed tail must publish no prefix"
    );
}

#[test]
fn malformed_offset_shapes_are_rejected() {
    assert_invalid(&[0], 1, false, 1, 1);
    assert_invalid(&offsets(&[Some(0), None]), 1, false, 2, 2);
    assert_invalid(&offsets(&[Some(0)]), 1, false, 1, 0);
    assert_invalid(&offsets(&[Some(0), None, Some(0)]), 1, false, 2, 2);
    assert_invalid(&offsets(&[Some(2), Some(1)]), 3, false, 2, 2);
    assert_invalid(&offsets(&[Some(3)]), 3, false, 1, 1);
    assert_invalid(&offsets(&[Some(0), Some(0)]), 1, false, 2, 2);
    assert_invalid(&offsets(&[Some(0), None, Some(1)]), 2, false, 2, 2);
}

#[test]
fn largest_non_sentinel_wide_offset_does_not_wrap() {
    let offsets = offsets(&[Some(u16::MAX - 1)]);
    let start = usize::from(u16::MAX - 1) * 4;
    let storage = vec![0_u8; start + 1];
    let (spans, _) =
        CellSpans::parse(&offsets, storage.len(), true, 1, 1, options(&offsets)).unwrap();
    assert_spans(spans, &[(0, start, start + 1)]);
}

#[test]
fn work_budget_is_precharged_and_reports_zero_wire_fields() {
    let offsets = offsets(&[Some(0), Some(4)]);
    let work = offsets.len() * 2;
    let (_, report) =
        CellSpans::parse(&offsets, 8, false, 2, 2, limited_options(&offsets, 0, work)).unwrap();
    assert_eq!(report.fields(), 0);
    assert_eq!(report.work_bytes(), work);

    let error = CellSpans::parse(
        &offsets,
        8,
        false,
        2,
        2,
        limited_options(&offsets, 0, work - 1),
    )
    .expect_err("one byte below the precharged span scan must fail");
    assert_eq!(
        error.resource_limit(),
        Some(DecodeLimit::Work {
            observed: work,
            maximum: work - 1,
        })
    );
}

#[test]
fn byte_budget_is_checked_before_offset_validation() {
    let offsets = offsets(&[Some(0)]);
    let error = CellSpans::parse(
        &offsets,
        1,
        false,
        1,
        1,
        DecodeOptions::new(1, usize::MAX, usize::MAX, 64, usize::MAX, usize::MAX),
    )
    .expect_err("the source byte ceiling must reject before reading offsets");
    assert_eq!(
        error.resource_limit(),
        Some(DecodeLimit::Bytes {
            observed: offsets.len(),
            maximum: 1,
        })
    );
}

#[test]
fn tile_row_uses_modern_buffers_only_as_a_complete_pair() {
    let pre_storage = b"pre";
    let pre_offsets = offsets(&[Some(0)]);
    let modern_storage = b"modern";
    let modern_offsets = offsets(&[Some(2)]);

    let modern_source = tile_row(
        1,
        pre_storage,
        &pre_offsets,
        Some(modern_storage),
        Some(&modern_offsets),
        None,
    );
    let modern = decode_tile_row_info(&modern_source, options(&modern_source)).unwrap();
    assert_eq!(
        modern.cell_storage_and_offsets(),
        (&modern_storage[..], &modern_offsets[..])
    );
    let (modern_spans, _) = modern.cell_spans(1, options(&modern_source)).unwrap();
    assert_eq!(
        modern_spans.get(0).unwrap().bytes(modern_storage),
        Some(b"dern".as_slice())
    );

    let partial_source = tile_row(
        1,
        pre_storage,
        &pre_offsets,
        Some(modern_storage),
        None,
        None,
    );
    let partial = decode_tile_row_info(&partial_source, options(&partial_source)).unwrap();
    assert_eq!(
        partial.cell_storage_and_offsets(),
        (&pre_storage[..], &pre_offsets[..])
    );
    let (partial_spans, _) = partial.cell_spans(1, options(&partial_source)).unwrap();
    assert_eq!(
        partial_spans.get(0).unwrap().bytes(pre_storage),
        Some(b"pre".as_slice())
    );
}

#[test]
fn cell_span_bytes_returns_none_for_a_shorter_unrelated_buffer() {
    let offsets = offsets(&[Some(0), Some(2)]);
    let (spans, _) = CellSpans::parse(&offsets, 4, false, 2, 2, options(&offsets)).unwrap();
    let span = spans.get(1).unwrap();
    assert_eq!(span.bytes(&[0_u8; 1]), None);
}

#[test]
fn invalid_shape_has_no_resource_limit_disguise() {
    let offsets = offsets(&[Some(0), Some(0)]);
    let error = CellSpans::parse(&offsets, 1, false, 2, 2, options(&offsets))
        .expect_err("equal starts are structural invalidity");
    assert_eq!(error.resource_limit(), None);
    assert_eq!(error, DecodeError::invalid_visitor_result());
}
