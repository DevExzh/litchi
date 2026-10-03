//! Focused tests for the full-cell runner's post-timing worksheet guard.

use std::{error::Error, sync::Arc};

use litchi_core::OwnedSource;
use litchi_xlsx::SourceBackedWorkbook;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::*;

fn rewrite_archive_member(
    archive: &[u8],
    target: &str,
    rewrite: impl FnOnce(Vec<u8>) -> Vec<u8>,
) -> Vec<u8> {
    let reader = ArchiveReader::new(archive).expect("archive reader");
    let mut writer = StreamingArchiveWriter::new();
    let mut rewrite = Some(rewrite);
    for name in reader.file_names() {
        let payload = reader.read(name).expect("archive member");
        let payload = if name == target {
            rewrite
                .take()
                .expect("archive member rewrite was used once")(payload)
        } else {
            payload
        };
        writer
            .write_deflated(name, &payload)
            .expect("rewritten archive member");
    }
    assert!(rewrite.is_none(), "target archive member was not found");
    writer.finish_to_bytes().expect("rewritten archive")
}

fn producer_read_archive() -> Vec<u8> {
    build_xlsx_producer_corpus(XlsxProducerShape::Medium, XlsxProducerVariant::Read)
        .expect("producer XLSX corpus builds")
        .corpus
        .archive
}

fn source_workbook(archive: Vec<u8>) -> SourceBackedWorkbook {
    SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(archive)))
        .expect("source-backed workbook opens")
}

#[test]
fn forced_guard_rejects_unpublished_dirty_worksheet_after_clean_aggregate() {
    let original = producer_read_archive();
    // A dirty first sheet alone would not catch a guard that accidentally
    // checks only the first worksheet. Keep every preceding sheet clean and
    // require rejection at each of the four positions independently.
    for index in 0..4 {
        let member = format!("xl/worksheets/sheet{}.xml", index + 1);
        let archive = rewrite_archive_member(&original, &member, |payload| {
            let xml = String::from_utf8(payload).expect("worksheet XML");
            // A misplaced dimension is tolerated and recorded only when this
            // lazy worksheet store is materialized.
            let changed = xml.replacen(
                "</sheetData>",
                "</sheetData><dimension ref=\"A1:AF32\"/>",
                1,
            );
            assert_ne!(changed, xml, "worksheet anchor was not found");
            changed.into_bytes()
        });
        let workbook = source_workbook(archive);
        assert!(
            workbook.tolerance_report().is_clean(),
            "worksheet {index} has no published store before the guard"
        );
        let sheets = workbook.sheets().collect::<Vec<_>>();
        let error = force_full_cell_worksheet_tolerance_reports(&sheets)
            .expect_err("dirty worksheet passed the forced tolerance guard");
        assert!(
            error
                .to_string()
                .contains(&format!("worksheet {index} required leniency repairs")),
            "unexpected dirty worksheet error: {error}"
        );
        assert!(
            !workbook.tolerance_report().is_clean(),
            "the aggregate should expose dirty worksheet {index} after forcing"
        );
    }
}

#[test]
fn forced_guard_propagates_malformed_worksheet_store_error() {
    let archive = rewrite_archive_member(
        &producer_read_archive(),
        "xl/worksheets/sheet1.xml",
        |payload| {
            let xml = String::from_utf8(payload).expect("worksheet XML");
            // Leave the package catalog valid while making the worksheet XML
            // itself unterminated.  The lazy store must return its typed parse
            // error; the guard must not turn it into clean evidence.
            let changed = xml.replacen("</row>", "<row>", 1);
            assert_ne!(changed, xml, "worksheet row anchor was not found");
            changed.into_bytes()
        },
    );
    let workbook = source_workbook(archive);
    assert!(workbook.tolerance_report().is_clean());
    let sheets = workbook.sheets().collect::<Vec<_>>();
    let error = force_full_cell_worksheet_tolerance_reports(&sheets)
        .expect_err("malformed worksheet store was accepted");
    assert!(
        error
            .to_string()
            .contains("worksheet 0 tolerance report failed"),
        "unexpected malformed worksheet error: {error}"
    );
    let mut cause = error.as_ref() as &(dyn Error + 'static);
    let mut saw_typed_source = false;
    while let Some(next) = cause.source() {
        if next.is::<litchi_xlsx::Error>() {
            saw_typed_source = true;
            break;
        }
        cause = next;
    }
    assert!(
        saw_typed_source,
        "malformed worksheet error lost its typed litchi-xlsx source: {error:?}"
    );
}
