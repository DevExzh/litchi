#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

//! Opt-in tolerance for non-structural DOC stylesheet defects.

use litchi_doc::tracked_revision::Limits;
use litchi_doc::writer::{BookmarkEntry, CommentEntry};
use litchi_doc::{
    AuxiliaryTableDefect, BookmarkDefect, CommentDefect, Defect, Leniency, OpenOptions, Package,
    StshifDefect, StyleDefect, StylesheetDefect, Writer,
};
use litchi_ole_common::object::{Editor as PackageEditor, Targets};
use std::io::Cursor;
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name)
}

/// MS-DOC 2.9 requires style names to be unique, and by default this reader
/// enforces that. The fixture is a real Word document that violates it.
#[test]
fn duplicate_style_names_are_rejected_by_default() {
    let mut package = Package::open(fixture("duplicate-style-names.doc")).expect("container opens");

    match package.document() {
        Err(error) => assert!(
            error.to_string().contains("unique"),
            "unexpected error: {error}"
        ),
        Ok(_) => panic!("a duplicated style name must be fatal under the default contract"),
    }
}

/// Opting in recovers the document. The name is only a label — every stored
/// reference resolves a style by index — so the text and structure are intact.
#[test]
fn opting_in_recovers_the_document_and_records_the_repair() {
    let mut package = Package::open(fixture("duplicate-style-names.doc")).expect("container opens");

    let document = package
        .document_with_options(
            OpenOptions::default().with_leniency(Leniency::TolerateStylesheetDefects),
        )
        .expect("a duplicated style name is repairable");

    let text = document.text().expect("text is extractable");
    assert!(
        !text.trim().is_empty(),
        "expected the document body to be readable once the stylesheet is tolerated"
    );

    let stylesheet = document.stylesheet().expect("the fixture has a stylesheet");
    let report = stylesheet.tolerance_report();
    assert!(
        !report.is_clean(),
        "the repair must be recorded, never silent"
    );
    assert!(
        report
            .defects()
            .iter()
            .any(|d| d.defect == Defect::Stylesheet(StylesheetDefect::DuplicateStyleName)),
        "expected a duplicate-style-name defect, got {:?}",
        report.defects()
    );
    assert_eq!(
        report.total(),
        report.defects().len() as u64,
        "the fixture is far below the recording bound, so nothing should be unrecorded"
    );

    // The document-level channel reports the same repair, so a caller that
    // never touches the stylesheet still sees everything that was tolerated.
    assert!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .any(|d| d.defect == Defect::Stylesheet(StylesheetDefect::DuplicateStyleName)),
        "the document report must include the stylesheet repair, got {:?}",
        document.tolerance_report().defects()
    );
}

/// Leniency is scoped to the stylesheet: a structurally broken document stays
/// a hard error in both modes, so opting in cannot mask real corruption.
#[test]
fn structural_defects_stay_fatal_when_tolerating_stylesheet_defects() {
    // Word 6.0 predates the table stream entirely; no stylesheet tolerance can
    // make it readable.
    let mut package = Package::open(fixture("word6-no-table-stream.doc")).expect("container opens");

    assert!(
        package
            .document_with_options(
                OpenOptions::default().with_leniency(Leniency::TolerateStylesheetDefects),
            )
            .is_err(),
        "leniency must not extend past the stylesheet"
    );
}

/// A conforming document reports a clean slate under either contract.
#[test]
fn a_conforming_document_records_no_repairs() {
    for leniency in [Leniency::Strict, Leniency::TolerateStylesheetDefects] {
        let mut package = Package::open(fixture("Lists.doc")).expect("container opens");
        let document = package
            .document_with_options(OpenOptions::default().with_leniency(leniency))
            .expect("a conforming document parses under either contract");

        if let Some(stylesheet) = document.stylesheet() {
            assert!(
                stylesheet.tolerance_report().is_clean(),
                "{leniency:?} reported a repair on a conforming document"
            );
        }
        assert!(
            document.tolerance_report().is_clean(),
            "{leniency:?} reported a repair on a conforming document"
        );
    }
}

// ──────────────────────────────────────────────────────────────────────────
// Optional auxiliary tables and Stshif reserved bytes (default repairs)
// ──────────────────────────────────────────────────────────────────────────

fn writer_document(build: impl FnOnce(&mut Writer)) -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph("Body text for the leniency fixtures.")
        .expect("fixture paragraph");
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("fixture DOC");
    output.into_inner()
}

/// Rewrite the `WordDocument` stream inside a DOC package.
fn edit_word_stream(bytes: &[u8], edit: impl FnOnce(&mut Vec<u8>)) -> Vec<u8> {
    let mut package = PackageEditor::open(bytes.to_vec(), Targets::default(), Limits::default())
        .expect("fixture package");
    let word_path = ["WordDocument".to_string()];
    let mut word = package
        .stream(&word_path)
        .expect("WordDocument stream")
        .to_vec();
    edit(&mut word);
    package
        .put_stream(&word_path, word)
        .expect("edited WordDocument stream");
    package.finish().expect("edited package")
}

/// Point one FIB table entry at a one-byte range, which every auxiliary
/// table parser rejects. The declared pointer count is raised when the index
/// lies past the Writer's Word 2002 layout; the added entries land in the
/// zeroed slack between the FIB and the text, never in document content.
fn break_table_pointer(word: &mut [u8], index: usize) {
    let declared = usize::from(u16::from_le_bytes([word[152], word[153]]));
    if declared < index + 1 {
        word[152..154].copy_from_slice(&u16::try_from(index + 1).unwrap().to_le_bytes());
    }
    let pair = 154 + index * 8;
    word[pair..pair + 4].copy_from_slice(&0u32.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&1u32.to_le_bytes());
}

fn open_degraded(bytes: Vec<u8>) -> litchi_doc::Document {
    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package opens");
    package
        .document()
        .expect("degraded optional tables must not fail the open")
}

/// MS-DOC 2.5 makes every FIB-pointed auxiliary table independently optional
/// metadata, so one malformed table — or all of them — must not cost the
/// document. Each failure degrades to the table's absent value and is
/// recorded with the table's FIB pointer index.
#[test]
fn malformed_auxiliary_tables_degrade_and_are_recorded() {
    let bytes = writer_document(|writer| {
        writer.add_comment(CommentEntry::new(4, "Review", "Alice", "A"));
    });
    const TABLES: &[(usize, AuxiliaryTableDefect)] = &[
        (16, AuxiliaryTableDefect::Fields),
        (4, AuxiliaryTableDefect::Comments),
        (21, AuxiliaryTableDefect::Bookmarks),
        (114, AuxiliaryTableDefect::SmartTags),
        (113, AuxiliaryTableDefect::Rsids),
        (94, AuxiliaryTableDefect::RmdThreading),
        (61, AuxiliaryTableDefect::EmbeddedFonts),
        (97, AuxiliaryTableDefect::Envelope),
        (89, AuxiliaryTableDefect::AutoSummary),
        (141, AuxiliaryTableDefect::ProtectedRanges),
        (120, AuxiliaryTableDefect::FormatConsistency),
        (137, AuxiliaryTableDefect::StructuredTags),
        (136, AuxiliaryTableDefect::XmlSchemas),
        (140, AuxiliaryTableDefect::CustomXmlTransform),
        (85, AuxiliaryTableDefect::OleControls),
        (44, AuxiliaryTableDefect::MailMerge),
        (54, AuxiliaryTableDefect::Subdocuments),
        (51, AuxiliaryTableDefect::RevisionAuthors),
        (32, AuxiliaryTableDefect::AssociatedStrings),
        (91, AuxiliaryTableDefect::ListNames),
        (96, AuxiliaryTableDefect::ListTemplates),
        (6, AuxiliaryTableDefect::Sections),
    ];
    let bytes = edit_word_stream(&bytes, |word| {
        for &(index, _) in TABLES {
            break_table_pointer(word, index);
        }
    });

    let document = open_degraded(bytes);

    // The primary content and the structural tables are fully readable.
    let text = document.text().expect("text extraction");
    assert!(text.contains("Body text for the leniency fixtures."));
    assert!(
        !document.paragraphs().expect("paragraphs").is_empty(),
        "paragraph extraction survives auxiliary degradation"
    );
    assert!(document.stylesheet().is_some());

    // Every degraded table reports the same absent value as a document that
    // never carried it — no panics, no invented content.
    assert!(document.fields_table().is_none());
    assert!(document.comments().expect("comments").is_empty());
    assert!(document.bookmarks().expect("bookmarks").is_empty());
    assert!(document.smart_tags().is_none());
    assert!(document.rsids().is_none());
    assert!(document.rmd_threading().is_none());
    assert!(document.embedded_fonts().is_none());
    assert!(document.envelope().is_none());
    assert!(document.auto_summary().is_none());
    assert!(document.protected_ranges().is_none());
    assert!(document.format_consistency_marks().is_none());
    assert!(document.structured_tags().is_none());
    assert!(document.xml_schemas().is_none());
    assert!(document.custom_xml_transform_path().is_none());
    assert!(document.ole_controls().is_none());
    assert!(document.mail_merge().is_none());
    assert!(document.subdocuments().is_none());
    assert!(document.revision_authors().is_empty());
    assert!(document.associated_strings().is_none());
    assert!(document.list_names().is_none());
    assert!(document.list_templates().is_none());
    assert!(document.sections().is_empty());

    // Every degradation is enumerated, each under its table's FIB index.
    let report = document.tolerance_report();
    for &(index, defect) in TABLES {
        assert!(
            report.defects().iter().any(|tolerated| {
                tolerated.defect == Defect::AuxiliaryTable(defect)
                    && usize::from(tolerated.context) == index
            }),
            "missing {defect:?} at FIB index {index}: {:?}",
            report.defects()
        );
    }
}

/// A malformed `SttbfRMark` degrades only itself: sections and styles parse
/// against the empty table without cascading failures or invented authors.
#[test]
fn a_malformed_revision_author_table_does_not_cascade() {
    let bytes = writer_document(|_| {});
    let bytes = edit_word_stream(&bytes, |word| break_table_pointer(word, 51));

    let document = open_degraded(bytes);
    assert!(document.revision_authors().is_empty());
    assert!(
        !document.sections().is_empty(),
        "an unrelated table must survive the degradation"
    );
    assert!(
        document
            .text()
            .expect("text extraction")
            .contains("Body text")
    );

    let report = document.tolerance_report();
    assert_eq!(
        report.defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::AuxiliaryTable(AuxiliaryTableDefect::RevisionAuthors),
            context: 51,
        }],
        "exactly one repair is recorded for exactly one broken table"
    );
}

/// The stylesheet is mandatory in every Word 97+ file, so its Stshif
/// reserved bytes — which MS-DOC 2.9.274 marks "MUST be ignored" — must not
/// cost the document even under the strict default contract.
#[test]
fn stshif_reserved_bytes_are_ignored_and_recorded() {
    let bytes = writer_document(|_| {});
    let mut package =
        PackageEditor::open(bytes, Targets::default(), Limits::default()).expect("fixture package");
    let word_path = ["WordDocument".to_string()];
    let word = package
        .stream(&word_path)
        .expect("WordDocument stream")
        .to_vec();
    // FIB pointer 1 addresses the stylesheet; fWhichTblStm (bit 9 of the
    // flags word at offset 10) selects the table stream.
    let stsh_offset =
        usize::try_from(u32::from_le_bytes(word[162..166].try_into().unwrap())).unwrap();
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    // Stshif follows the two-byte cbStshi: flags at +4, fixed count at +8.
    table[stsh_offset + 6..stsh_offset + 8].copy_from_slice(&0xBEEFu16.to_le_bytes());
    table[stsh_offset + 10..stsh_offset + 12].copy_from_slice(&12u16.to_le_bytes());
    package
        .put_stream(&table_path, table)
        .expect("edited table stream");
    let bytes = package.finish().expect("edited package");

    let document = open_degraded(bytes);
    let stylesheet = document.stylesheet().expect("stylesheet parses");
    assert_eq!(stylesheet.header().fixed_style_count, 12);
    assert!(
        !stylesheet.styles().is_empty(),
        "style records are read past the ignored bytes"
    );
    assert!(
        document
            .text()
            .expect("text extraction")
            .contains("Body text")
    );

    let report = document.tolerance_report();
    for (defect, context) in [
        (Defect::Stshif(StshifDefect::StyleNameFlags), 0xBEEF),
        (Defect::Stshif(StshifDefect::FixedStyleCount), 12),
    ] {
        assert!(
            report
                .defects()
                .iter()
                .any(|tolerated| tolerated.defect == defect && tolerated.context == context),
            "missing {defect:?} with context {context:#06x}: {:?}",
            report.defects()
        );
    }
}

/// Reserved-byte tolerance stops at framing: a stylesheet too short to hold
/// its own header is structural corruption and stays a hard error.
#[test]
fn stylesheet_framing_stays_fatal() {
    let bytes = writer_document(|_| {});
    let bytes = edit_word_stream(&bytes, |word| {
        // Truncate the stylesheet to four bytes: the STSHI header is
        // incomplete, which is framing, not a reserved-bit deviation.
        let pair = 154 + 8;
        word[pair + 4..pair + 8].copy_from_slice(&4u32.to_le_bytes());
    });

    let mut package = Package::from_reader(Cursor::new(bytes)).expect("package opens");
    assert!(
        package.document().is_err(),
        "stylesheet framing must stay fatal under the default contract"
    );
}

/// Rewrite the stylesheet bytes of a DOC package in place. The closure gets
/// the table stream and the stylesheet's offset within it.
fn edit_stylesheet(bytes: Vec<u8>, edit: impl FnOnce(&mut [u8], usize)) -> Vec<u8> {
    let mut package =
        PackageEditor::open(bytes, Targets::default(), Limits::default()).expect("fixture package");
    let word = package
        .stream(&["WordDocument".to_string()])
        .expect("WordDocument stream")
        .to_vec();
    let stsh_offset =
        usize::try_from(u32::from_le_bytes(word[162..166].try_into().unwrap())).unwrap();
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    edit(&mut table, stsh_offset);
    package
        .put_stream(&table_path, table)
        .expect("edited table stream");
    package.finish().expect("edited package")
}

/// Offset of the first style record (`STD`, istd 0) within the table stream.
fn first_std_offset(table: &[u8], stsh_offset: usize) -> usize {
    let cb_stshi = usize::from(u16::from_le_bytes([
        table[stsh_offset],
        table[stsh_offset + 1],
    ]));
    let lpstd = stsh_offset + 2 + cb_stshi;
    assert_ne!(
        u16::from_le_bytes([table[lpstd], table[lpstd + 1]]),
        0,
        "the Writer always emits the Normal style"
    );
    lpstd + 2
}

/// MS-DOC 2.9.112 marks the reserved GRFSTD bits "MUST be ignored": a style
/// carrying them opens by default with only the defined flags decoded.
#[test]
fn reserved_grfstd_bits_are_ignored_and_recorded() {
    let clean = writer_document(|_| {});
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    let original_flags = original.stylesheet().unwrap().get(0).unwrap().flags;

    let bytes = edit_stylesheet(clean, |table, stsh| {
        let grfstd = first_std_offset(table, stsh) + 8;
        let word = u16::from_le_bytes([table[grfstd], table[grfstd + 1]]) | 0xE080;
        table[grfstd..grfstd + 2].copy_from_slice(&word.to_le_bytes());
    });
    let document = open_degraded(bytes);
    let normal = document.stylesheet().unwrap().get(0).unwrap();
    assert_eq!(
        normal.flags, original_flags,
        "only defined bits are decoded"
    );
    assert!(document.text().unwrap().contains("Body text"));
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::Style(StyleDefect::ReservedStyleFlags),
            context: 0,
        }]
    );
}

/// Rewrite the table stream of a DOC package in place. The closure gets the
/// `WordDocument` stream (for FIB pointers) and the table stream.
fn edit_table_stream(bytes: Vec<u8>, edit: impl FnOnce(&[u8], &mut [u8])) -> Vec<u8> {
    let mut package =
        PackageEditor::open(bytes, Targets::default(), Limits::default()).expect("fixture package");
    let word = package
        .stream(&["WordDocument".to_string()])
        .expect("WordDocument stream")
        .to_vec();
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut table = package.stream(&table_path).expect("table stream").to_vec();
    edit(&word, &mut table);
    package
        .put_stream(&table_path, table)
        .expect("edited table stream");
    package.finish().expect("edited package")
}

/// The table-stream offset stored for one FIB `rgFcLcb` pointer.
fn table_pointer(word: &[u8], index: usize) -> usize {
    let pair = 154 + index * 8;
    usize::try_from(u32::from_le_bytes(word[pair..pair + 4].try_into().unwrap())).unwrap()
}

/// MS-DOC 2.9.7 marks the ATRDPre10 reserved fields "MUST be ignored", and
/// an author index or bookmark tag out of its domain is a cross-reference
/// defect, not framing: each repair stays on the comment carrying it.
#[test]
fn comment_descriptor_defects_are_repaired_per_comment() {
    let clean = writer_document(|writer| {
        writer.add_comment(CommentEntry::new(4, "Review", "Alice", "A"));
    });
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    let original_comments = original.comments().expect("comments");
    assert_eq!(original_comments.len(), 1);
    assert_eq!(original_comments[0].author, "Alice");
    assert_eq!(original_comments[0].bookmark_tag, None);

    let bytes = edit_table_stream(clean, |word, table| {
        // PlcfandRef (FIB pointer 4) holds two CPs, then the ATRDPre10.
        let atrd = table_pointer(word, 4) + 8;
        table[atrd + 20..atrd + 22].copy_from_slice(&9u16.to_le_bytes()); // ibst
        table[atrd + 22..atrd + 24].copy_from_slice(&0xBEEFu16.to_le_bytes()); // bitsNotUsed
        table[atrd + 26..atrd + 30].copy_from_slice(&(-7i32).to_le_bytes()); // lTagBkmk
    });
    let document = open_degraded(bytes);
    let comments = document.comments().expect("comments");
    assert_eq!(
        comments.len(),
        1,
        "the comment survives its descriptor defects"
    );
    assert_eq!(comments[0].text, original_comments[0].text);
    assert_eq!(comments[0].initials, "A");
    assert_eq!(
        comments[0].author, "",
        "an unresolvable author is left empty"
    );
    assert_eq!(comments[0].bookmark_tag, None);
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| (tolerated.defect, tolerated.context))
            .collect::<Vec<_>>(),
        vec![
            (Defect::Comment(CommentDefect::AuthorIndex), 0),
            (Defect::Comment(CommentDefect::ReservedFields), 0),
            (Defect::Comment(CommentDefect::BookmarkTag), 0),
        ]
    );
}

/// MS-DOC 2.9.8 marks BKC.fPub "MUST be ignored", and a broken ibkl
/// cross-reference costs only the bookmark carrying it, never the table.
#[test]
fn bookmark_defects_are_repaired_per_bookmark() {
    let clean = writer_document(|writer| {
        writer.add_bookmark(BookmarkEntry::new("First", 0, 4));
        writer.add_bookmark(BookmarkEntry::new("Second", 5, 9));
    });
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    let original_bookmarks = original.bookmarks().expect("bookmarks");
    assert_eq!(original_bookmarks.len(), 2);

    let bytes = edit_table_stream(clean, |word, table| {
        // PlcfBkf (FIB pointer 22): three CPs, then one FBKF per bookmark.
        let fbkf = table_pointer(word, 22) + 12;
        table[fbkf + 2] |= 0x80; // First: BKC.fPub
        let first_ibkl = [table[fbkf], table[fbkf + 1]];
        table[fbkf + 4..fbkf + 6].copy_from_slice(&first_ibkl); // Second: duplicate ibkl
    });
    let document = open_degraded(bytes);
    let bookmarks = document.bookmarks().expect("bookmarks");
    assert_eq!(bookmarks.len(), 1, "only the broken bookmark is dropped");
    assert_eq!(bookmarks[0], original_bookmarks[0]);
    assert!(document.text().unwrap().contains("Body text"));
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| (tolerated.defect, tolerated.context))
            .collect::<Vec<_>>(),
        vec![
            (Defect::Bookmark(BookmarkDefect::ReservedPublicFlag), 0),
            (Defect::Bookmark(BookmarkDefect::EndIndex), 1),
        ]
    );
}
