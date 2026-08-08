#![cfg(feature = "iwork")]

use std::sync::{Arc, Weak};

use litchi::iwork::{
    Cell, CellView, Document, Error, ErrorKind, Format, Options, Resource, Section, SectionKind,
    Slide, Snapshot, SnapshotLimits, SourceLimits, Stage, Summary, Table, Text, TextRole, Value,
    ValueKind,
};

const PAGES: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");
const KEYNOTE: &[u8] = include_bytes!("../../../test-data/iwork/keynote/basic.key");
const NUMBERS: &[u8] = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");

fn assert_send_sync_static<T: Send + Sync + 'static>() {}

#[test]
fn public_values_have_the_expected_thread_and_lifetime_bounds() {
    assert_send_sync_static::<Document>();
    assert_send_sync_static::<Snapshot>();
    assert_send_sync_static::<Table>();
    assert_send_sync_static::<Slide>();
    assert_send_sync_static::<Section>();
    assert_send_sync_static::<Summary>();
    assert_send_sync_static::<Format>();
    assert_send_sync_static::<Options>();
    assert_send_sync_static::<SourceLimits>();
    assert_send_sync_static::<SnapshotLimits>();
    assert_send_sync_static::<Error>();
    assert_send_sync_static::<ErrorKind>();
    assert_send_sync_static::<Stage>();
    assert_send_sync_static::<Resource>();
    assert_send_sync_static::<SectionKind>();
    assert_send_sync_static::<TextRole>();
    assert_send_sync_static::<ValueKind>();
    assert_send_sync_static::<Cell<'static>>();
    assert_send_sync_static::<CellView<'static>>();
    assert_send_sync_static::<Text<'static>>();
    assert_send_sync_static::<Value<'static>>();
    assert_send_sync_static::<litchi::iwork::Result<Snapshot>>();
}

#[test]
fn semantic_handles_outlive_documents_and_intermediate_snapshots() {
    fn detached_snapshot(value: &[u8]) -> Snapshot {
        let document = Document::from_bytes(value)
            .unwrap_or_else(|error| panic!("native fixture must decode: {error}"));
        document.snapshot()
    }

    fn detached_table() -> Table {
        let snapshot = detached_snapshot(NUMBERS);
        snapshot
            .table(0)
            .unwrap_or_else(|| panic!("Numbers fixture must have one table"))
    }

    fn detached_slide() -> Slide {
        let snapshot = detached_snapshot(KEYNOTE);
        snapshot
            .slide(0)
            .unwrap_or_else(|| panic!("Keynote fixture must have one slide"))
    }

    fn detached_section() -> Section {
        let snapshot = detached_snapshot(PAGES);
        snapshot
            .section(0)
            .unwrap_or_else(|| panic!("Pages fixture must have one section"))
    }

    let table = detached_table();
    let slide = detached_slide();
    let section = detached_section();

    assert_eq!(table.name(), "Table 1");
    assert_eq!(
        table.cell(1, 1),
        Some(CellView::Stored(Value::Text(
            "Litchi native Numbers fixture"
        )))
    );
    assert_eq!(slide.title(), Some("Litchi native Keynote fixture"));
    assert_eq!(section.name(), Some("Blank"));

    let table_reader = std::thread::spawn(move || assert_eq!(table.position(), 0));
    let slide_reader = std::thread::spawn(move || assert_eq!(slide.position(), 0));
    let section_reader = std::thread::spawn(move || assert_eq!(section.position(), 0));
    for reader in [table_reader, slide_reader, section_reader] {
        reader
            .join()
            .unwrap_or_else(|_panic| panic!("detached semantic handle reader panicked"));
    }
}

#[test]
fn pages_fixture_has_focused_semantics_and_detached_handles() {
    let document = Document::from_bytes(PAGES)
        .unwrap_or_else(|error| panic!("native Pages fixture must decode: {error}"));
    assert_eq!(document.format(), Format::Pages);

    let snapshot = document.snapshot();
    assert_eq!(snapshot.table_count(), 0);
    assert_eq!(snapshot.slide_count(), 0);
    assert_eq!(snapshot.section_count(), 1);
    assert_eq!(
        snapshot.summary().to_string(),
        "Tables: 0, Slides: 0, Sections: 1"
    );

    let section = snapshot
        .section(0)
        .unwrap_or_else(|| panic!("fixture must have one section"));
    assert_eq!(section.position(), 0);
    assert_eq!(section.kind(), SectionKind::Body);
    assert_eq!(section.name(), Some("Blank"));
    assert_eq!(section.heading(), None);
    assert_eq!(section.paragraphs().count(), 0);
    assert_eq!(
        section.additional_text().collect::<Vec<_>>(),
        ["Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07"]
    );
    assert_eq!(
        snapshot
            .iter_text()
            .map(|item| (item.role(), item.value()))
            .collect::<Vec<_>>(),
        [(
            TextRole::SectionAdditional,
            "Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07"
        )]
    );

    drop(snapshot);
    drop(document);
    assert_eq!(section.name(), Some("Blank"));
}

#[test]
fn keynote_fixture_preserves_visible_text_order_and_role() {
    let document = Document::from_bytes(KEYNOTE)
        .unwrap_or_else(|error| panic!("native Keynote fixture must decode: {error}"));
    assert_eq!(document.format(), Format::Keynote);
    let snapshot = document.snapshot();
    assert_eq!(snapshot.table_count(), 0);
    assert_eq!(snapshot.slide_count(), 1);
    assert_eq!(snapshot.section_count(), 0);

    let slide = snapshot
        .slide(0)
        .unwrap_or_else(|| panic!("fixture must have one slide"));
    assert_eq!(slide.position(), 0);
    assert_eq!(slide.title(), Some("Litchi native Keynote fixture"));
    assert!(!slide.is_skipped());
    assert_eq!(slide.text_blocks().count(), 0);
    assert_eq!(slide.notes(), None);
    assert_eq!(
        slide
            .iter_text()
            .map(|item| (item.role(), item.value()))
            .collect::<Vec<_>>(),
        [
            (TextRole::SlideTitle, "Litchi native Keynote fixture"),
            (
                TextRole::SlideAdditional,
                "Buffa lazy-view migration verification"
            ),
            (TextRole::SlideAdditional, "2026-08-07"),
        ]
    );
}

#[test]
fn numbers_fixture_uses_global_compatibility_projection_and_typed_cells() {
    let document = Document::from_bytes(NUMBERS)
        .unwrap_or_else(|error| panic!("native Numbers fixture must decode: {error}"));
    assert_eq!(document.format(), Format::Numbers);
    let snapshot = document.snapshot();
    assert_eq!(snapshot.table_count(), 1);
    assert_eq!(snapshot.slide_count(), 0);
    assert_eq!(snapshot.section_count(), 0);

    let table = snapshot
        .table(0)
        .unwrap_or_else(|| panic!("fixture must have one table"));
    assert_eq!(table.position(), 0);
    assert_eq!(table.name(), "Table 1");
    assert_eq!((table.row_count(), table.column_count()), (22, 7));
    assert_eq!(table.cell_count(), 2);
    assert_eq!(table.non_empty_cell_count(), 2);
    assert_eq!(
        table.cell(1, 1),
        Some(CellView::Stored(Value::Text(
            "Litchi native Numbers fixture"
        )))
    );
    assert_eq!(
        table.cell(2, 1),
        Some(CellView::Stored(Value::Number(42.0)))
    );
    assert_eq!(table.cell(0, 0), Some(CellView::Missing));
    assert_eq!(table.cell(22, 0), None);
    assert_eq!(
        table
            .iter_text()
            .map(|item| (item.role(), item.value()))
            .collect::<Vec<_>>(),
        [
            (TextRole::TableName, "Table 1"),
            (TextRole::TableCell, "Litchi native Numbers fixture"),
        ]
    );

    let focused = litchi_numbers::compatibility_tables_from_bytes(NUMBERS)
        .unwrap_or_else(|error| panic!("focused compatibility projection must decode: {error}"));
    assert_eq!(focused.len(), snapshot.table_count());
    assert_eq!(focused[0].name(), table.name());
    assert_eq!(focused[0].row_count(), table.row_count());
    assert_eq!(focused[0].column_count(), table.column_count());
    assert_eq!(focused[0].cell_count(), table.cell_count());
}

#[test]
fn shared_package_allocation_is_released_after_eager_projection() {
    fn read_and_release(value: &[u8], expected: Format) -> Weak<[u8]> {
        let value: Arc<[u8]> = Arc::from(value);
        let weak = Arc::downgrade(&value);
        let document = Document::from_shared_bytes(value)
            .unwrap_or_else(|error| panic!("fixture must decode: {error}"));
        assert_eq!(document.format(), expected);
        assert!(!document.snapshot().is_empty());
        weak
    }

    assert!(read_and_release(PAGES, Format::Pages).upgrade().is_none());
    assert!(
        read_and_release(KEYNOTE, Format::Keynote)
            .upgrade()
            .is_none()
    );
    assert!(
        read_and_release(NUMBERS, Format::Numbers)
            .upgrade()
            .is_none()
    );
}

#[test]
fn concurrent_snapshot_reads_are_deterministic() {
    let document = Arc::new(
        Document::from_bytes(KEYNOTE)
            .unwrap_or_else(|error| panic!("native Keynote fixture must decode: {error}")),
    );
    let expected = document.snapshot().all_text();

    std::thread::scope(|scope| {
        let readers = (0..16)
            .map(|_| {
                let document = Arc::clone(&document);
                let expected = &expected;
                scope.spawn(move || {
                    for _ in 0..64 {
                        let snapshot = document.snapshot();
                        assert_eq!(snapshot.format(), Format::Keynote);
                        assert_eq!(snapshot.summary().slides(), 1);
                        assert_eq!(snapshot.all_text(), *expected);
                    }
                })
            })
            .collect::<Vec<_>>();
        for reader in readers {
            reader
                .join()
                .unwrap_or_else(|_panic| panic!("concurrent snapshot reader panicked"));
        }
    });
}

#[test]
fn physical_and_text_limits_are_inclusive_and_report_one_over() {
    let exact_source = SourceLimits::new(
        PAGES.len() as u64,
        SourceLimits::HARD_MAX_ENTRIES,
        SourceLimits::HARD_MAX_ENTRY_BYTES,
        SourceLimits::HARD_MAX_EXPANDED_BYTES,
        SourceLimits::HARD_MAX_DECODED_BYTES_PER_ITEM,
    )
    .unwrap_or_else(|error| panic!("exact input profile must be valid: {error}"));
    Document::from_bytes_with_options(PAGES, Options::default().with_source(exact_source))
        .unwrap_or_else(|error| panic!("exact input limit must pass: {error}"));

    let one_under_source = SourceLimits::new(
        PAGES.len() as u64 - 1,
        SourceLimits::HARD_MAX_ENTRIES,
        SourceLimits::HARD_MAX_ENTRY_BYTES,
        SourceLimits::HARD_MAX_EXPANDED_BYTES,
        SourceLimits::HARD_MAX_DECODED_BYTES_PER_ITEM,
    )
    .unwrap_or_else(|error| panic!("tight input profile must be valid: {error}"));
    let input_error =
        Document::from_bytes_with_options(PAGES, Options::default().with_source(one_under_source))
            .err()
            .unwrap_or_else(|| panic!("one-over input must fail"));
    assert_eq!(input_error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(input_error.stage(), Stage::Input);
    assert_eq!(input_error.resource(), Some(Resource::InputBytes));
    assert_eq!(input_error.observed(), Some(PAGES.len() as u64));
    assert_eq!(input_error.maximum(), Some(PAGES.len() as u64 - 1));

    let text = "Litchi native Pages fixture\nBuffa lazy-view migration verification\n2026-08-07";
    let retained_text_bytes = text.len() + "Blank".len();
    let exact_text = SnapshotLimits::new(
        SnapshotLimits::HARD_MAX_TABLES,
        SnapshotLimits::HARD_MAX_SLIDES,
        SnapshotLimits::HARD_MAX_SECTIONS,
        retained_text_bytes,
    )
    .unwrap_or_else(|error| panic!("exact text profile must be valid: {error}"));
    Document::from_bytes_with_options(PAGES, Options::default().with_snapshot(exact_text))
        .unwrap_or_else(|error| panic!("exact semantic text limit must pass: {error}"));

    let one_under_text = SnapshotLimits::new(
        SnapshotLimits::HARD_MAX_TABLES,
        SnapshotLimits::HARD_MAX_SLIDES,
        SnapshotLimits::HARD_MAX_SECTIONS,
        retained_text_bytes - 1,
    )
    .unwrap_or_else(|error| panic!("tight text profile must be valid: {error}"));
    let text_error =
        Document::from_bytes_with_options(PAGES, Options::default().with_snapshot(one_under_text))
            .err()
            .unwrap_or_else(|| panic!("one-over semantic text must fail"));
    assert_eq!(text_error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(text_error.stage(), Stage::Semantic);
    assert_eq!(text_error.format(), Some(Format::Pages));
    assert_eq!(text_error.resource(), Some(Resource::TextBytes));
    assert_eq!(text_error.observed(), Some(retained_text_bytes as u64));
    assert_eq!(text_error.maximum(), Some(retained_text_bytes as u64 - 1));
}
