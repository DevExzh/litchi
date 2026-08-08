#![no_main]

use std::hint::black_box;

use libfuzzer_sys::fuzz_target;
use litchi::iwork::{Document, Format, Options, Snapshot, SnapshotLimits, SourceLimits};

const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024;
const MAX_ENTRIES: usize = 512;
const MAX_ENTRY_BYTES: u64 = 8 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 32 * 1024 * 1024;
const MAX_DECODED_BYTES_PER_ITEM: usize = 8 * 1024 * 1024;
const MAX_TABLES: usize = 4_096;
const MAX_SLIDES: usize = 4_096;
const MAX_SECTIONS: usize = 4_096;
const MAX_TEXT_BYTES: usize = 4 * 1024 * 1024;

fuzz_target!(|data: &[u8]| {
    let Ok(document) = Document::from_bytes_with_options(data, fuzz_options()) else {
        return;
    };

    let snapshot = document.snapshot();
    assert_eq!(document.format(), snapshot.format());
    assert_snapshot(&snapshot);
});

fn fuzz_options() -> Options {
    let source = SourceLimits::new(
        MAX_INPUT_BYTES,
        MAX_ENTRIES,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_DECODED_BYTES_PER_ITEM,
    )
    .unwrap_or_else(|error| unreachable!("valid fuzz source limits: {error}"));
    let snapshot = SnapshotLimits::new(MAX_TABLES, MAX_SLIDES, MAX_SECTIONS, MAX_TEXT_BYTES)
        .unwrap_or_else(|error| unreachable!("valid fuzz snapshot limits: {error}"));
    Options::new(source, snapshot)
}

fn assert_snapshot(snapshot: &Snapshot) {
    let summary = snapshot.summary();
    assert_eq!(summary.tables(), snapshot.table_count());
    assert_eq!(summary.slides(), snapshot.slide_count());
    assert_eq!(summary.sections(), snapshot.section_count());
    assert_eq!(
        snapshot.is_empty(),
        summary.tables() == 0 && summary.slides() == 0 && summary.sections() == 0
    );

    assert!(snapshot.table(snapshot.table_count()).is_none());
    assert!(snapshot.slide(snapshot.slide_count()).is_none());
    assert!(snapshot.section(snapshot.section_count()).is_none());

    match snapshot.format() {
        Format::Numbers => {
            assert_eq!(snapshot.slide_count(), 0);
            assert_eq!(snapshot.section_count(), 0);
        },
        Format::Keynote => {
            assert_eq!(snapshot.table_count(), 0);
            assert_eq!(snapshot.section_count(), 0);
        },
        Format::Pages => {
            assert_eq!(snapshot.table_count(), 0);
            assert_eq!(snapshot.slide_count(), 0);
        },
        _ => {},
    }

    for (position, table) in snapshot.tables().enumerate() {
        assert_eq!(table.position(), position);
        assert!(table.non_empty_cell_count() <= table.cell_count());
        black_box(table.name());
        for cell in table.cells() {
            assert!(cell.row() < table.row_count());
            assert!(cell.column() < table.column_count());
            black_box(cell.value().kind());
            black_box(table.cell(cell.row(), cell.column()));
        }
        for text in table.iter_text() {
            black_box((text.role(), text.value()));
        }
    }

    for (position, slide) in snapshot.slides().enumerate() {
        assert_eq!(slide.position(), position);
        black_box(slide.is_skipped());
        black_box(slide.name());
        black_box(slide.title());
        black_box(slide.text_blocks().count());
        black_box(slide.additional_text().count());
        black_box(slide.notes());
        black_box(slide.build_count());
        black_box(slide.has_transition());
        for text in slide.iter_text() {
            black_box((text.role(), text.value()));
        }
    }

    for (position, section) in snapshot.sections().enumerate() {
        assert_eq!(section.position(), position);
        black_box(section.kind());
        black_box(section.name());
        black_box(section.heading());
        black_box(section.paragraphs().count());
        black_box(section.additional_text().count());
        black_box(section.page_count());
        for text in section.iter_text() {
            black_box((text.role(), text.value()));
        }
    }

    let clone = snapshot.clone();
    assert_eq!(snapshot.summary(), clone.summary());
    assert!(snapshot.iter_text().eq(clone.iter_text()));
}
