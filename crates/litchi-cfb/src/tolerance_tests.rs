//! Default-repair fixtures for the compound-file tolerance contract.
//!
//! Each defect class has a hand-built minimal file that carries the
//! deviation and must open by default with the repair recorded; the same
//! file without the deviation must open with a clean report.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::sync::Arc;

use litchi_core::SourceVersion;

use crate::{OleFile, SharedOleFile, ToleranceReport, writer::OleWriter};

fn build(streams: &[(&[&str], &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        writer.create_stream(path, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Opens `bytes` through both readers and returns the cursor reader's report
/// after checking that the positional reader agrees.
fn open_report(bytes: &[u8]) -> ToleranceReport {
    let file = OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let shared =
        SharedOleFile::open_owned(Arc::from(bytes.to_vec()), SourceVersion::new(0x7a, 0)).unwrap();
    assert_eq!(file.tolerance_report(), shared.tolerance_report());
    file.tolerance_report().clone()
}

#[test]
fn conformant_files_open_with_a_clean_report() {
    let small = build(&[(&["Data"], b"payload")]);
    let regular = build(&[
        (&["Data"], &[0x5a; 4096]),
        (&["Store", "Inner"], &[1; 9000]),
    ]);
    let empty = build(&[(&["Empty"], b"")]);
    for bytes in [small, regular, empty] {
        assert!(open_report(&bytes).is_clean());
    }
}

#[test]
fn real_world_compound_files_open_with_a_clean_report() {
    for bytes in [
        include_bytes!("../../../test-data/ole/doc/FancyFoot.doc").as_slice(),
        include_bytes!("../../../test-data/ole/doc/Lists.doc").as_slice(),
    ] {
        assert!(open_report(bytes).is_clean());
    }
}
