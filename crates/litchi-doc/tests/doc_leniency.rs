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
    PapxDefect, PieceDefect, SectionDefect, SprmDefect, StshifDefect, StyleDefect,
    StylesheetDefect, TableDefinitionDefect, Writer,
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
                    && usize::try_from(tolerated.context).unwrap() == index
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

/// Locate the first TC80 of the first `sprmTDefTable` (0xD608) in the
/// `WordDocument` stream: opcode, two-byte size, itcMac, rgdxaCenter.
fn first_tc80_offset(word: &[u8]) -> usize {
    let sprm = word
        .windows(2)
        .position(|pair| pair == [0x08, 0xD6])
        .expect("the Writer emits sprmTDefTable for a table row");
    let itc_mac = usize::from(word[sprm + 4]);
    sprm + 5 + (itc_mac + 1) * 2
}

/// MS-DOC 2.9.17's brcType restriction is a writer constraint and an Ico
/// above 16 names no color: a table border carrying either reads as no line
/// / automatic color, the way LibreOffice reads it, instead of failing the
/// paragraph properties and the document.
#[test]
fn undefined_table_border_bytes_are_repaired_and_recorded() {
    let clean = writer_document(|writer| {
        writer.add_table(1, 2).expect("fixture table");
    });
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    assert_eq!(original.tables().expect("tables").len(), 1);

    let bytes = edit_word_stream(&clean, |word| {
        let brc_top = first_tc80_offset(word) + 4;
        word[brc_top..brc_top + 4].copy_from_slice(&[8, 0x1B, 17, 0]);
    });
    let document = open_degraded(bytes);
    assert_eq!(document.tables().expect("tables").len(), 1);
    let defects = document
        .tolerance_report()
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect::<Vec<_>>();
    for expected in [
        (Defect::Sprm(SprmDefect::BorderType), 0xD608),
        (Defect::Sprm(SprmDefect::BorderColor), 0xD608),
    ] {
        assert!(
            defects.contains(&expected),
            "missing {expected:?}: {defects:?}"
        );
    }
    assert!(
        defects
            .iter()
            .all(|(defect, _)| matches!(defect, Defect::Sprm(_))),
        "only the border repairs are recorded: {defects:?}"
    );
}

/// MS-DOC 2.9.313 places no reader-side rejection on unused TC80 encodings:
/// a cell carrying one reads as the neutral member, recorded with its cell
/// index, instead of failing the paragraph properties and the document.
#[test]
fn unused_tc80_encodings_are_repaired_per_cell() {
    let clean = writer_document(|writer| {
        writer.add_table(1, 2).expect("fixture table");
    });
    let bytes = edit_word_stream(&clean, |word| {
        let tc = first_tc80_offset(word);
        let flags = u16::from_le_bytes([word[tc], word[tc + 1]]) | (3 << 7); // vertAlign 3
        word[tc..tc + 2].copy_from_slice(&flags.to_le_bytes());
        let tc = tc + 20;
        let flags = (u16::from_le_bytes([word[tc], word[tc + 1]]) & !(7 << 2)) | (6 << 2); // textFlow 6
        word[tc..tc + 2].copy_from_slice(&flags.to_le_bytes());
    });
    let document = open_degraded(bytes);
    assert_eq!(document.tables().expect("tables").len(), 1);
    let defects = document
        .tolerance_report()
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect::<Vec<_>>();
    for expected in [
        (
            Defect::TableDefinition(TableDefinitionDefect::VerticalAlignment),
            0,
        ),
        (Defect::TableDefinition(TableDefinitionDefect::TextFlow), 1),
    ] {
        assert!(
            defects.contains(&expected),
            "missing {expected:?}: {defects:?}"
        );
    }
    assert!(
        defects
            .iter()
            .all(|(defect, _)| matches!(defect, Defect::TableDefinition(_))),
        "only the TC80 repairs are recorded: {defects:?}"
    );
}

/// The `WordDocument` offset of the first FKP page named by the bin table at
/// FIB pointer `index` (12 = `PlcfBteChpx`, 13 = `PlcfBtePapx`), plus that
/// pointer's table-stream offset.
fn first_bin_table_page(bytes: &[u8], index: usize) -> (usize, usize) {
    let package = PackageEditor::open(bytes.to_vec(), Targets::default(), Limits::default())
        .expect("fixture package");
    let word = package
        .stream(&["WordDocument".to_string()])
        .expect("WordDocument stream");
    let table_name = if u16::from_le_bytes([word[10], word[11]]) & 0x0200 != 0 {
        "1Table"
    } else {
        "0Table"
    };
    let table = package
        .stream(&[table_name.to_string()])
        .expect("table stream");
    let bte = table_pointer(word, index);
    // One page: two FCs, then the page number.
    let pn = u32::from_le_bytes(table[bte + 8..bte + 12].try_into().unwrap());
    (usize::try_from(pn).unwrap() * 512, bte)
}

/// MS-DOC 2.4.3 defines each `PapxFkp` run independently: one run whose
/// SPRMs cannot be read gets default paragraph properties, the way Word and
/// LibreOffice render it, instead of failing every paragraph property.
#[test]
fn a_malformed_papx_run_reads_with_default_properties() {
    let clean = writer_document(|_| {});
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());

    let (page, _) = first_bin_table_page(&clean, 13);
    let bytes = edit_word_stream(&clean, |word| {
        let page = &mut word[page..page + 512];
        assert_eq!(page[511], 1, "the fixture page holds one run");
        // Re-point the run at a PapxInFkp whose SPRM (sprmPDyaLine, 0x6412)
        // is missing its four-byte operand.
        page[8] = 200;
        page[400..406].copy_from_slice(&[0, 2, 0, 0, 0x12, 0x64]);
    });
    let document = open_degraded(bytes);
    assert!(document.text().unwrap().contains("Body text"));
    assert!(!document.paragraphs().expect("paragraphs").is_empty());
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::Papx(PapxDefect::Run),
            context: 0,
        }]
    );
}

/// A `PlcfBtePapx` entry naming a page outside the stream costs only that
/// page's paragraph properties, not the document.
#[test]
fn an_unreachable_papx_page_is_skipped_and_recorded() {
    let clean = writer_document(|_| {});
    let bytes = edit_table_stream(clean, |word, table| {
        let bte = table_pointer(word, 13);
        table[bte + 8..bte + 12].copy_from_slice(&0x003F_FFFFu32.to_le_bytes());
    });
    let document = open_degraded(bytes);
    assert!(document.text().unwrap().contains("Body text"));
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::Papx(PapxDefect::Page),
            context: 0,
        }]
    );
}

/// The repeated producer shape — one zero pad byte after a complete
/// odd-length PAPX SPRM sequence — now opens by default, with the trim
/// recorded rather than gated behind an opt-in profile.
#[test]
fn papx_alignment_padding_is_trimmed_by_default_and_recorded() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document/test.doc");
    let mut package = Package::open(path).expect("container opens");
    let document = package
        .document()
        .expect("the producer pad no longer needs an opt-in");
    assert!(!document.text().expect("text").is_empty());
    let report = document.tolerance_report();
    assert!(
        report
            .defects()
            .iter()
            .any(|tolerated| tolerated.defect == Defect::Papx(PapxDefect::AlignmentPadding)),
        "the trim must be recorded: {:?}",
        report.defects()
    );
    assert!(
        report
            .defects()
            .iter()
            .all(|tolerated| tolerated.defect != Defect::Papx(PapxDefect::Run)),
        "the padded runs are read, not defaulted: {:?}",
        report.defects()
    );
}

/// MS-DOC 2.9.17 marks the Brc80 reserved bit "MUST be ignored", and an Ico
/// above 16 names no color: a section page border carrying either reads with
/// the bit masked and automatic color, recorded against the SPRM, instead of
/// losing the whole section table.
#[test]
fn section_page_border_defects_are_repaired_and_recorded() {
    use litchi_doc::{Border, Borders, Color, Style};

    let border = Border {
        style: Style::Single,
        width_eighth_points: 8,
        color: Color::Red,
        spacing_points: 4,
        shadow: false,
        frame: false,
    };
    let clean = writer_document(|writer| {
        writer
            .set_section_page_borders(Borders {
                top: Some(border),
                ..Borders::default()
            })
            .expect("fixture borders");
    });
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    assert_eq!(original.sections()[0].page_borders.top, Some(border));

    let bytes = edit_word_stream(&clean, |word| {
        // sprmSBrcTop80 (0x702B): width, brcType, ico, effects.
        let sprm = word
            .windows(4)
            .position(|bytes| bytes == [0x2B, 0x70, 8, 0x01])
            .expect("the Writer emits sprmSBrcTop80");
        word[sprm + 4] = 0x11; // ico 17
        word[sprm + 5] |= 0x80; // reserved effect bit
    });
    let document = open_degraded(bytes);
    assert_eq!(document.sections().len(), 1, "the section table survives");
    assert_eq!(
        document.sections()[0].page_borders.top,
        Some(Border {
            color: Color::Automatic,
            ..border
        })
    );
    assert_eq!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .map(|tolerated| (tolerated.defect, tolerated.context))
            .collect::<Vec<_>>(),
        vec![
            (Defect::Section(SectionDefect::BorderColor), 0x702B),
            (Defect::Section(SectionDefect::ReservedBits), 0x702B),
        ]
    );
}

/// One `Pcd.prm` that cannot be resolved costs only that piece's modifier.
/// Previously it silently discarded the whole piece table and with it every
/// character and paragraph property; now the formatting survives and the
/// piece is recorded.
#[test]
fn an_unresolvable_piece_modifier_is_ignored_and_recorded() {
    let clean = writer_document(|_| {});
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());
    let original_paragraphs = format!("{:?}", original.paragraphs().expect("paragraphs"));

    let bytes = edit_table_stream(clean, |word, table| {
        let clx = table_pointer(word, 33);
        assert_eq!(table[clx], 0x02, "the Writer emits no RgPrc");
        let lcb = usize::try_from(u32::from_le_bytes(
            table[clx + 1..clx + 5].try_into().unwrap(),
        ))
        .unwrap();
        let pieces = (lcb - 4) / 12;
        let pcd = clx + 5 + (pieces + 1) * 4;
        // Prm1 naming RgPrc entry 9, which does not exist.
        table[pcd + 6..pcd + 8].copy_from_slice(&((9u16 << 1) | 1).to_le_bytes());
    });
    let document = open_degraded(bytes);
    assert_eq!(
        format!("{:?}", document.paragraphs().expect("paragraphs")),
        original_paragraphs,
        "the piece keeps its text and FKP formatting"
    );
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::Piece(PieceDefect::PropertyModifier),
            context: 0,
        }]
    );
}

/// Container repairs surface through the document report: a non-critical
/// stream whose FAT chain runs past its declared size opens (POI and
/// LibreOffice read exactly the declared bytes) and the compound-file repair
/// is listed as a `Defect::Container` entry ahead of every DOC-level repair.
#[test]
fn container_repairs_join_the_document_report() {
    let clean = writer_document(|_| {});
    let mut package =
        PackageEditor::open(clean, Targets::default(), Limits::default()).expect("fixture package");
    package
        .add_stream(vec!["LeniencyProbe".to_string()], vec![0xAB; 5000])
        .expect("probe stream");
    let with_probe = package.finish().expect("package with probe");

    let conformant = open_degraded(with_probe.clone());
    assert!(
        conformant.tolerance_report().is_clean(),
        "a conformant container records nothing: {:?}",
        conformant.tolerance_report().defects()
    );

    // Shrink the probe's declared size from 5000 to 4096 bytes: its ten-sector
    // chain now runs two sectors past the eight the size needs.
    let name = "LeniencyProbe"
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    let mut bytes = with_probe;
    let entry = bytes
        .windows(name.len())
        .position(|window| window == name.as_slice())
        .expect("probe directory entry");
    assert_eq!(
        u32::from_le_bytes(bytes[entry + 120..entry + 124].try_into().unwrap()),
        5000
    );
    bytes[entry + 120..entry + 124].copy_from_slice(&4096u32.to_le_bytes());

    let document = open_degraded(bytes);
    assert!(document.text().unwrap().contains("Body text"));
    let defects = document.tolerance_report().defects();
    assert!(
        defects.iter().any(|tolerated| tolerated.defect
            == Defect::Container(litchi_cfb::Defect::Allocation(
                litchi_cfb::AllocationDefect::ChainLongerThanSize
            ))),
        "missing the container repair: {defects:?}"
    );
    assert!(
        defects
            .iter()
            .all(|tolerated| matches!(tolerated.defect, Defect::Container(_))),
        "only container repairs are recorded: {defects:?}"
    );
}

/// MS-DOC 2.9.203's `PLRSID.reserved1` constant is a writer rule on an inert
/// lookup table: a Word document carrying another value keeps its RSIDs and
/// records the repair.
#[test]
fn a_wrong_plrsid_reserved_marker_is_ignored_and_recorded() {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/document/47950_normal.doc"),
    )
    .expect("fixture");
    let rsid_defect = Defect::Rsid(litchi_doc::RsidDefect::ReservedMarker);
    let original = open_degraded(bytes.clone());
    assert_eq!(original.rsids().expect("PLRSID").ids().len(), 3);
    assert!(
        original
            .tolerance_report()
            .defects()
            .iter()
            .all(|tolerated| tolerated.defect != rsid_defect)
    );

    let bytes = edit_table_stream(bytes, |word, table| {
        let plrsid = table_pointer(word, 113);
        table[plrsid + 12..plrsid + 16].copy_from_slice(&0xBEEFu32.to_le_bytes());
    });
    let document = open_degraded(bytes);
    assert_eq!(
        document.rsids().expect("PLRSID still read").ids(),
        original.rsids().unwrap().ids()
    );
    assert!(
        document
            .tolerance_report()
            .defects()
            .iter()
            .any(|tolerated| tolerated.defect == rsid_defect && tolerated.context == 0xBEEF),
        "{:?}",
        document.tolerance_report().defects()
    );
}

/// MS-DOC 2.9.27's Boolean8 domain is enforced by writers, but readers treat
/// the byte as a CHP toggle (POI `getCHPFlag`): a CHPX carrying
/// `sprmCFRMark` 0x02 keeps its run's other properties and records the
/// repair instead of silently dropping every property of the run.
#[test]
fn a_non_boolean_chp_revision_flag_is_repaired_and_recorded() {
    let clean = writer_document(|_| {});
    let original = open_degraded(clean.clone());
    assert!(original.tolerance_report().is_clean());

    // PlcfBteChpx (FIB pointer 12) names one page; re-point its first run
    // at a CHPX holding sprmCFRMark (0x0801) = 2 and sprmCFBold = 1.
    let (page, _) = first_bin_table_page(&clean, 12);
    let chpx = 400usize;
    let bytes = edit_word_stream(&clean, |word| {
        let page = &mut word[page..page + 512];
        let crun = usize::from(page[511]);
        assert!(crun >= 1, "the fixture page holds a run");
        page[(crun + 1) * 4] = u8::try_from(chpx / 2).unwrap();
        page[chpx..chpx + 7].copy_from_slice(&[6, 0x01, 0x08, 0x02, 0x35, 0x08, 0x01]);
    });
    let document = open_degraded(bytes);
    assert!(document.text().unwrap().contains("Body text"));
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::Chp(litchi_doc::ChpDefect::Boolean),
            context: 0x0801,
        }]
    );
}

/// A one-section `SummaryInformation` stream whose title is a `VT_LPWSTR`
/// holding an unpaired high surrogate, which the property-set reader decodes
/// lossily (U+FFFD) and records against the title's property identifier.
fn summary_information_with_unpaired_surrogate() -> Vec<u8> {
    const SUMMARY_INFORMATION_FMTID: [u8; 16] = [
        0xE0, 0x85, 0x9F, 0xF2, 0xF9, 0x4F, 0x68, 0x10, 0xAB, 0x91, 0x08, 0x00, 0x2B, 0x27, 0xB3,
        0xD9,
    ];
    let code_page = [2u16.to_le_bytes(), [0, 0], 1200u16.to_le_bytes(), [0, 0]].concat();
    let title = [
        0x1Fu16.to_le_bytes(),
        [0, 0],
        3u16.to_le_bytes(),
        [0, 0],
        0xD800u16.to_le_bytes(),
        0x41u16.to_le_bytes(),
        [0, 0],
        [0, 0],
    ]
    .concat();
    let table_end = 8 + 2 * 8;
    let mut section = Vec::new();
    section.extend_from_slice(&((table_end + code_page.len() + title.len()) as u32).to_le_bytes());
    section.extend_from_slice(&2u32.to_le_bytes());
    section.extend_from_slice(&1u32.to_le_bytes());
    section.extend_from_slice(&(table_end as u32).to_le_bytes());
    section.extend_from_slice(&2u32.to_le_bytes());
    section.extend_from_slice(&((table_end + code_page.len()) as u32).to_le_bytes());
    section.extend_from_slice(&code_page);
    section.extend_from_slice(&title);
    let mut stream = Vec::new();
    stream.extend_from_slice(&0xFFFEu16.to_le_bytes());
    stream.extend_from_slice(&0u16.to_le_bytes());
    stream.extend_from_slice(&0u32.to_le_bytes());
    stream.extend_from_slice(&[0; 16]);
    stream.extend_from_slice(&1u32.to_le_bytes());
    stream.extend_from_slice(&SUMMARY_INFORMATION_FMTID);
    stream.extend_from_slice(&48u32.to_le_bytes());
    stream.extend_from_slice(&section);
    stream
}

/// Property-set repairs surface through the document report: a title with an
/// unpaired surrogate in `\x05SummaryInformation` decodes lossily the way
/// ole32 and POI decode it, and the repair is listed as a
/// `Defect::PropertySet` entry located by the title's property identifier.
#[test]
fn property_set_repairs_join_the_document_report() {
    let clean = writer_document(|_| {});
    assert!(open_degraded(clean.clone()).tolerance_report().is_clean());

    let mut package =
        PackageEditor::open(clean, Targets::default(), Limits::default()).expect("fixture package");
    let path = vec!["\u{5}SummaryInformation".to_string()];
    let stream = summary_information_with_unpaired_surrogate();
    if package.stream(&path).is_some() {
        package.put_stream(&path, stream).expect("summary stream");
    } else {
        package.add_stream(path, stream).expect("summary stream");
    }
    let bytes = package.finish().expect("package with lossy metadata");

    let document = open_degraded(bytes.clone());
    assert!(document.text().unwrap().contains("Body text"));
    assert_eq!(
        document.tolerance_report().defects(),
        &[litchi_doc::ToleratedDefect {
            defect: Defect::PropertySet(litchi_ole_common::leniency::Defect::String(
                litchi_ole_common::leniency::StringDefect::LossyDecode,
            )),
            context: 2,
        }]
    );
    // The metadata accessor still hands the lossy title back with its own
    // report.
    let summary = Package::from_reader(Cursor::new(bytes))
        .unwrap()
        .summary_information()
        .unwrap()
        .expect("summary stream");
    assert!(!summary.tolerance_report().is_clean());
}
