use litchi_rtf::write::Writer;
use litchi_rtf::{Document, read};

fn assert_snapshot_traits<T: Clone + Send + Sync>() {}

#[test]
fn document_is_a_small_shared_snapshot() {
    assert_snapshot_traits::<Document>();
    assert_eq!(
        std::mem::size_of::<Document>(),
        std::mem::size_of::<usize>()
    );

    let document = Document::parse(r"{\rtf1\ansi first\par second}").unwrap();
    let shared = document.clone();

    assert!(document.same_snapshot(&shared));
    assert_eq!(document.paragraph_count(), 2);
    assert_eq!(document.text(), shared.text());
    assert!(std::ptr::eq(document.text(), shared.text()));
}

#[test]
fn contextual_reader_uses_checked_limits() {
    let limits = read::Limits::new().with_max_source_bytes(8);
    assert!(read::Document::parse_with_limits(r"{\rtf1 too large}", limits).is_err());
}

#[test]
fn snapshot_round_trips_through_concise_writer() {
    let document =
        Document::parse(r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello!\par}")
            .unwrap();

    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write(&document).unwrap();
    let reparsed = Document::from_bytes(&bytes).unwrap();

    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.fonts().len(), 1);
    assert!(!reparsed.is_empty());
}

#[test]
fn snapshot_to_bytes_matches_streaming_facade() {
    let document = Document::parse(r"{\rtf1\ansi concise}").unwrap();
    let direct = document.to_bytes().unwrap();
    let reparsed = Document::from_bytes(&direct).unwrap();
    assert_eq!(reparsed.text(), "concise");
}
