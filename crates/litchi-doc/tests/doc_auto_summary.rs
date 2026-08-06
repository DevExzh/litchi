//! Tests for the `PlcfAsumy` AutoSummary priority table (MS-DOC 2.8.4).

use litchi_cfb::OleFile;
use litchi_doc::parts::fib::FileInformationBlock;
use litchi_doc::{AutoSummaryRange, DocumentAutoSummary};
use std::fs::File;
use std::path::PathBuf;

fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../..")
        .join(relative)
}

fn parse_auto_summary(relative: &str) -> Option<DocumentAutoSummary> {
    let mut ole = OleFile::open(File::open(fixture(relative)).unwrap()).unwrap();
    let word_document = ole.open_stream(&["WordDocument"]).unwrap();
    let fib = FileInformationBlock::parse(&word_document).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_stream = ole.open_stream(&[table_name]).unwrap();
    DocumentAutoSummary::parse(&fib, &table_stream).unwrap()
}

#[test]
fn documents_without_asumy_report_none() {
    // None of the checked-in fixtures carry AutoSummary priorities; the
    // parser must report None rather than erroring.
    assert!(parse_auto_summary("test-data/ole/doc/ThreeColHeadFoot.doc").is_none());
    assert!(parse_auto_summary("test-data/poi/test-data/document/47950_normal.doc").is_none());
}

#[test]
fn document_api_reports_none_without_asumy() {
    let mut package = litchi_doc::Package::from_reader(
        File::open(fixture("test-data/ole/doc/ThreeColHeadFoot.doc")).unwrap(),
    )
    .unwrap();
    let document = package.document().unwrap();
    assert!(document.auto_summary().is_none());
}

fn range(start: u32, end: u32, level: u32) -> AutoSummaryRange {
    AutoSummaryRange::new(start, end, level).unwrap()
}

fn authored_summary() -> DocumentAutoSummary {
    DocumentAutoSummary::try_new(vec![range(0, 12, 1), range(12, 28, 3), range(28, 45, 2)]).unwrap()
}

fn plcf_asumy(cps: &[u32], levels: &[i32]) -> Vec<u8> {
    let mut data = Vec::with_capacity((cps.len() + levels.len()) * 4);
    for cp in cps {
        data.extend_from_slice(&cp.to_le_bytes());
    }
    for level in levels {
        data.extend_from_slice(&level.to_le_bytes());
    }
    data
}

#[test]
fn authoring_round_trips_the_complete_plcf_asumy() {
    let expected = authored_summary();
    let bytes = expected.to_bytes().unwrap();

    assert_eq!(DocumentAutoSummary::parse_bytes(&bytes).unwrap(), expected);
    assert_eq!(
        DocumentAutoSummary::parse_bytes(&bytes)
            .unwrap()
            .to_bytes()
            .unwrap(),
        bytes
    );
}

#[test]
fn constructors_enforce_positive_levels_and_signed_wire_bounds() {
    assert!(AutoSummaryRange::new(0, 1, 0).is_err());

    // CP has a format-defined upper bound and ASUMY.lLevel is signed on the
    // wire. The largest representable values remain authorable.
    assert!(AutoSummaryRange::new(2_147_483_645, 2_147_483_646, 2_147_483_647).is_ok());

    // Negative CPs and levels are represented by their two's-complement wire
    // values and must not be accepted by the PlcfAsumy parser.
    assert!(
        DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0xFFFF_FFFF, 0xFFFF_FFFF], &[1],)).is_err()
    );
    assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 1], &[-1])).is_err());
    assert!(
        DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0x7FFF_FFFE, 0x7FFF_FFFF], &[1],)).is_err()
    );
}

#[test]
fn parser_rejects_duplicate_and_non_increasing_cps() {
    assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 10, 10], &[1, 2])).is_err());
    assert!(DocumentAutoSummary::parse_bytes(&plcf_asumy(&[0, 10, 9], &[1, 2])).is_err());
    assert!(DocumentAutoSummary::try_new(vec![range(0, 10, 1), range(11, 20, 2),]).is_err());
}

#[test]
fn checked_edits_cover_push_insert_replace_remove() {
    let mut summary = authored_summary();
    let appended = range(45, 60, 4);
    summary.push(appended).unwrap();
    let inserted = range(60, 72, 5);
    summary.insert(summary.ranges().len(), inserted).unwrap();
    let replacement = range(12, 28, 6);
    summary.replace(1, replacement).unwrap();
    assert_eq!(summary.remove(4).unwrap(), inserted);
    assert_eq!(
        summary.ranges(),
        &[range(0, 12, 1), replacement, range(28, 45, 2), appended]
    );
}

#[test]
fn failed_edits_are_atomic() {
    let mut summary = authored_summary();

    let before = summary.clone();
    assert!(summary.push(range(44, 55, 4)).is_err());
    assert_eq!(summary, before);

    let before = summary.clone();
    assert!(summary.insert(1, range(9, 11, 4)).is_err());
    assert_eq!(summary, before);

    let before = summary.clone();
    assert!(summary.replace(1, range(11, 27, 4)).is_err());
    assert_eq!(summary, before);

    let before = summary.clone();
    assert!(summary.remove(summary.ranges().len()).is_err());
    assert_eq!(summary, before);
}

#[test]
fn constructor_enforces_the_one_million_entry_cap() {
    let ranges = (0..1_000_001)
        .map(|index| AutoSummaryRange::new(index, index + 1, 1).unwrap())
        .collect();
    assert!(DocumentAutoSummary::try_new(ranges).is_err());
}
