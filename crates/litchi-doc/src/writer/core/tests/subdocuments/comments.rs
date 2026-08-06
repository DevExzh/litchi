use super::super::support::*;

#[test]
fn comments_round_trip_with_other_subdocuments() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main 😀").unwrap();
    writer.add_footnote(FootnoteEntry::new(0, "Footnote", 1));
    writer.add_comment(
        CommentEntry::new(1, "Review 🦀", "Alice 😀", "A😀")
            .with_range(2, 6)
            .with_extended_metadata(crate::CommentExtendedMetadata {
                modified_at: Some(CommentDateTime {
                    year: 2026,
                    month: 7,
                    day: 15,
                    hour: 14,
                    minute: 30,
                    weekday: 3,
                }),
                depth: 0,
                parent_index: None,
                is_ink: false,
            }),
    );
    writer.add_comment(
        CommentEntry::new(3, "Second review", "Alice 😀", "AL")
            .with_range(0, 7)
            .with_extended_metadata(crate::CommentExtendedMetadata {
                modified_at: None,
                depth: 1,
                parent_index: Some(0),
                is_ink: true,
            }),
    );
    writer.add_endnote(FootnoteEntry::new(2, "Endnote", 1));
    writer.set_odd_header("Header");

    let mut cursor = Cursor::new(Vec::new());
    writer.write_to(&mut cursor).unwrap();
    let mut package = crate::Package::from_reader(Cursor::new(cursor.into_inner())).unwrap();
    let document = package.document().unwrap();

    assert_eq!(document.footnotes().unwrap().len(), 1);
    assert_eq!(document.headers().unwrap().len(), 1);
    assert_eq!(document.endnotes().unwrap().len(), 1);
    let comments = document.comments().unwrap();
    assert_eq!(comments.len(), 2);
    assert_eq!(comments[0].author, "Alice 😀");
    assert_eq!(comments[0].initials, "A😀");
    assert_eq!(comments[0].bookmark_tag, Some(0));
    assert_eq!(
        (comments[0].range_start, comments[0].range_end),
        (Some(2), Some(6))
    );
    let first_metadata = comments[0].extended_metadata.unwrap();
    assert_eq!(first_metadata.depth, 0);
    assert_eq!(first_metadata.parent_index, None);
    assert_eq!(
        first_metadata.modified_at,
        Some(CommentDateTime {
            year: 2026,
            month: 7,
            day: 15,
            hour: 14,
            minute: 30,
            weekday: 3,
        })
    );
    assert!(comments[0].text().contains("Review 🦀"));
    assert_eq!(comments[0].paragraphs().unwrap().len(), 1);
    assert_eq!(comments[1].author, "Alice 😀");
    assert_eq!(comments[1].initials, "AL");
    assert_eq!(
        (comments[1].range_start, comments[1].range_end),
        (Some(0), Some(7))
    );
    assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
    assert!(comments[1].extended_metadata.unwrap().is_ink);
    assert!(comments[1].text().contains("Second review"));

    let path = std::env::temp_dir().join(format!(
        "litchi-doc-comments-{}-{}.doc",
        std::process::id(),
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos()
    ));
    writer.save(&path).unwrap();
    let mut package = crate::Package::open(&path).unwrap();
    let comments = package.document().unwrap().comments().unwrap();
    assert_eq!(comments.len(), 2);
    assert_eq!(
        (comments[0].range_start, comments[0].range_end),
        (Some(2), Some(6))
    );
    assert_eq!(comments[1].extended_metadata.unwrap().parent_index, Some(0));
    std::fs::remove_file(path).unwrap();
}

#[test]
fn rejects_comment_metadata_outside_binary_limits() {
    let mut writer = Writer::new();
    writer.add_paragraph("Main").unwrap();
    writer.add_comment(CommentEntry::new(0, "Body", "Author", "0123456789"));

    let error = writer.write_to(&mut Cursor::new(Vec::new())).unwrap_err();
    assert!(error.to_string().contains("at most nine"));
}

#[test]
fn rejects_invalid_comment_ranges_timestamps_and_reply_trees() {
    let write_error = |entry: CommentEntry| {
        let mut writer = Writer::new();
        writer.add_paragraph("Main").unwrap();
        writer.add_comment(entry);
        writer
            .write_to(&mut Cursor::new(Vec::new()))
            .unwrap_err()
            .to_string()
    };

    let error = write_error(CommentEntry::new(0, "Body", "Author", "A").with_range(4, 2));
    assert!(error.contains("range must be ordered"));

    let error = write_error(
        CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
            crate::CommentExtendedMetadata {
                modified_at: Some(CommentDateTime {
                    year: 2026,
                    month: 13,
                    day: 1,
                    hour: 0,
                    minute: 0,
                    weekday: 0,
                }),
                depth: 0,
                parent_index: None,
                is_ink: false,
            },
        ),
    );
    assert!(error.contains("DTTM"));

    let error = write_error(
        CommentEntry::new(0, "Body", "Author", "A").with_extended_metadata(
            crate::CommentExtendedMetadata {
                modified_at: None,
                depth: 1,
                parent_index: Some(0),
                is_ink: false,
            },
        ),
    );
    assert!(error.contains("pre-order"));
}
