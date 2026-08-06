use super::super::support::*;

#[test]
fn standard_bookmarks_round_trip_through_both_output_paths() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main text").unwrap();
    writer.add_bookmark(BookmarkEntry::new("Outer", 2, 5));
    writer.add_bookmark(
        BookmarkEntry::new("_Cell", 0, 8)
            .with_native_export(false)
            .with_column_range(1, 3),
    );

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let bookmarks = package.document().unwrap().bookmarks().unwrap();
    assert_eq!(bookmarks.len(), 2);
    assert_eq!(bookmarks[0].name, "_Cell");
    assert_eq!((bookmarks[0].start, bookmarks[0].end), (0, 8));
    assert_eq!(bookmarks[0].column_range, Some((1, 3)));
    assert!(!bookmarks[0].is_native);
    assert_eq!(bookmarks[1].name, "Outer");
    assert_eq!((bookmarks[1].start, bookmarks[1].end), (2, 5));

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-bookmarks-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    assert_eq!(package.document().unwrap().bookmarks().unwrap(), bookmarks);
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rejects_invalid_standard_bookmarks() {
    let write_error = |entries: Vec<BookmarkEntry>| {
        let mut writer = Writer::new();
        writer.add_paragraph("Main").unwrap();
        for entry in entries {
            writer.add_bookmark(entry);
        }
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };
    assert!(write_error(vec![BookmarkEntry::new("", 0, 1)]).contains("names"));
    assert!(
        write_error(vec![
            BookmarkEntry::new("Same", 0, 1),
            BookmarkEntry::new("Same", 1, 2),
        ])
        .contains("unique")
    );
    assert!(write_error(vec![BookmarkEntry::new("Range", 4, 2)]).contains("range"));
    assert!(
        write_error(vec![
            BookmarkEntry::new("Column", 0, 1).with_column_range(3, 2)
        ])
        .contains("column")
    );
}
