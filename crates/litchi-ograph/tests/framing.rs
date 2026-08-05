use litchi_biff::{Encoder, Error as BiffError, Kind, Limits as BiffLimits, Resource};
use litchi_ograph::chart::{self, Kind as ChartKind};
use litchi_ograph::{Error, Limits};

const BOF: Kind = Kind::from_wire(0x0809);
const EOF: Kind = Kind::from_wire(0x000A);
const UNKNOWN: Kind = Kind::from_wire(0x7777);

fn bof() -> [u8; 16] {
    let mut payload = [0; 16];
    payload[0..2].copy_from_slice(&0x0600_u16.to_le_bytes());
    payload[2..4].copy_from_slice(&0x0020_u16.to_le_bytes());
    payload
}

fn chart_bytes() -> Vec<u8> {
    let mut output = Encoder::new();
    output.push(BOF, &bof()).expect("BOF");
    output.push(UNKNOWN, &[0xAA, 0xBB]).expect("unknown");
    output.push(EOF, &[]).expect("EOF");
    output.finish()
}

#[test]
fn canonical_frames_preserve_exact_bytes_and_offsets_through_ograph() {
    let bytes = chart_bytes();
    let chart = chart::Ref::open(&bytes).expect("chart");
    assert_eq!(chart.kind(), ChartKind::Excel);
    let records = chart
        .records()
        .map(|record| record.expect("valid BIFF record"))
        .collect::<Vec<_>>();

    assert_eq!(
        records
            .iter()
            .map(|record| record.offset())
            .collect::<Vec<_>>(),
        [0, 20, 26]
    );
    assert_eq!(records[1].encoded(), &bytes[20..26]);
    assert_eq!(
        records
            .iter()
            .flat_map(|record| record.encoded().iter().copied())
            .collect::<Vec<_>>(),
        bytes
    );
}

#[test]
fn malformed_frames_remain_typed_biff_errors_at_their_wire_offset() {
    let malformed = [0x09, 0x08, 0x04, 0x00, 0x01, 0x02];
    let mut records = chart::Refs::with_limits(&malformed, Limits::default()).expect("scanner");

    assert!(matches!(
        records.next(),
        Some(Err(Error::Biff(BiffError::TruncatedPayload {
            offset: 0,
            declared: 4,
            available: 2,
            ..
        })))
    ));
    assert!(records.next().is_none());
}

#[test]
fn biff_record_bounds_are_enforced_without_a_second_ograph_frame_policy() {
    let bytes = chart_bytes();
    let limits = Limits {
        biff: BiffLimits {
            max_records: 2,
            ..BiffLimits::default()
        },
        ..Limits::default()
    };
    let mut records = chart::Refs::with_limits(&bytes, limits).expect("scanner");

    assert!(matches!(
        records.next(),
        Some(Err(Error::Biff(BiffError::LimitExceeded {
            resource: Resource::RecordCount,
            observed: 3,
            maximum: 2,
        })))
    ));
    assert!(records.next().is_none());
}

#[test]
fn chart_context_still_rejects_non_chart_bofs_after_frame_validation() {
    let mut output = Encoder::new();
    let mut payload = bof();
    payload[2..4].copy_from_slice(&0x0005_u16.to_le_bytes());
    output.push(BOF, &payload).expect("BOF");
    output.push(EOF, &[]).expect("EOF");

    let error = chart::Book::open(output.finish()).expect_err("not a chart workbook");
    assert!(matches!(
        error,
        Error::InvalidChart {
            reason: "Workbook contains no chart substream",
            ..
        }
    ));
}
