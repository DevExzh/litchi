//! Native Word RTF formatting and positioned paragraphs; no layout or shaping.
use litchi_core::Position;
use litchi_rtf::edit::{self, Error, ParagraphLayoutPatch, RunFormatPatch, RunFormatUpdate};
use litchi_rtf::{
    Formatting, HeaderFooterType as Story, Package, ParagraphFrame,
    ParagraphFrameHorizontalPosition as HPos, ParagraphFrameHorizontalReference as HRef,
    ParagraphFrameVerticalPosition as VPos, ParagraphFrameVerticalReference as VRef,
    ParagraphFrameWrap as Wrap, TextDirection, UnderlineStyle,
};
use std::num::NonZeroU16;

const MINIMAL: &str = "word-rtf-minimal.rtf";
const BIDI: &str = "word-bidi-eastasia.rtf";
const MATH: &str = "word-math-navigation.rtf";
const TABLES: &str = "word-tables-layout.rtf";
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
// Paragraph marks are structural; compare every visible character's complete
// retained formatting regardless of how lexical escapes split parsed runs.
fn characters(p: &Package) -> Vec<(char, Formatting)> {
    p.raw()
        .blocks()
        .iter()
        .flat_map(|b| {
            b.text
                .chars()
                .filter(|c| *c != '\n')
                .map(move |c| (c, b.formatting))
        })
        .collect()
}
fn paragraphs(p: &Package) -> Vec<(String, litchi_rtf::Paragraph)> {
    p.raw()
        .paragraphs_with_content()
        .iter()
        .map(|p| (p.text(), p.properties))
        .collect()
}
fn assert_metadata(source: &Package, output: &Package) {
    assert_eq!(
        output.raw().font_table().fonts(),
        source.raw().font_table().fonts()
    );
    assert_eq!(
        output.raw().color_table().colors(),
        source.raw().color_table().colors()
    );
    assert_eq!(output.raw().stylesheet(), source.raw().stylesheet());
    assert_eq!(output.raw().info(), source.raw().info());
    assert_eq!(output.raw().data_store(), source.raw().data_store());
    assert_eq!(output.raw().theme(), source.raw().theme());
}
fn save_reopen(source: &Package, commit: &edit::Commit) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    commit.package().save(file.path()).unwrap();
    let read = Package::open(file.path()).unwrap();
    assert_metadata(source, &read);
    assert_eq!(
        source
            .apply(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        read.to_bytes().unwrap()
    );
    assert_eq!(
        read.apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        source.to_bytes().unwrap()
    );
    assert!(matches!(
        read.apply(commit.patch()),
        Err(Error::PatchConflict)
    ));
    read
}
fn frame() -> ParagraphFrame {
    ParagraphFrame {
        width_twips: Some(1440),
        height_twips: Some(-720),
        horizontal_reference: HRef::Margin,
        horizontal_position: HPos::NegativeOffset(-120),
        vertical_reference: VRef::Page,
        vertical_position: VPos::Offset(240),
        anchor_locked: Some(true),
        horizontal_text_distance_twips: Some(90),
        horizontal_text_offset_twips: Some(10),
        vertical_text_offset_twips: Some(20),
        wrap: Wrap::Around,
        no_overlap: Some(true),
        ..ParagraphFrame::default()
    }
}
fn footer_frames(p: &Package) -> Vec<(Story, usize, ParagraphFrame)> {
    p.sections()[0]
        .headers_footers
        .iter()
        .flat_map(|story| {
            story
                .paragraphs
                .iter()
                .enumerate()
                .filter_map(move |(index, p)| {
                    p.paragraph
                        .frame
                        .map(|frame| (story.header_type, index, frame))
                })
        })
        .collect()
}

#[test]
fn native_headings_expose_exact_fonts_colors_spacing_and_numbered_indents() {
    let bytes = native(MATH);
    let p = Package::from_bytes(&bytes).unwrap();
    let body = p.body().paragraphs().collect::<Vec<_>>();
    assert_eq!(body.len(), 19);
    for (index, text, size, bold, before, after, keep) in [
        (
            0,
            "Native Word equations and navigation",
            56,
            false,
            0,
            80,
            false,
        ),
        (3, "Table of Contents", 28, true, 480, 0, true),
        (7, "TOC_SOURCE_H1_005 Equations", 40, false, 360, 80, true),
        (
            12,
            "TOC_SOURCE_H2_005 Numbered cases",
            32,
            false,
            160,
            80,
            true,
        ),
    ] {
        let para = body[index];
        assert_eq!(para.to_text(), text);
        let runs = para.runs().collect::<Vec<_>>();
        assert_eq!(runs.len(), 1);
        let f = runs[0].format();
        assert_eq!(f.font().unwrap().name(), "Aptos Display");
        assert_eq!(f.size().get(), size);
        assert_eq!(f.bold(), bold);
        assert!(!f.italic());
        assert_eq!(f.underline(), UnderlineStyle::None);
        assert_eq!(
            f.foreground(),
            Some(if index == 0 {
                litchi_rtf::color::Value::Automatic
            } else {
                litchi_rtf::color::Value::Rgb(litchi_rtf::color::Color::new(15, 71, 97))
            })
        );
        let layout = para.format().layout();
        assert_eq!(
            (layout.space_before(), layout.space_after()),
            (before, after)
        );
        assert_eq!(
            (layout.keep_together(), layout.keep_with_next()),
            (keep, keep)
        );
        assert!(!layout.page_break_before());
    }
    for (index, left, first) in [(13, 360, -360), (14, 792, -432), (15, 360, -360)] {
        let l = body[index].format().layout();
        assert_eq!(
            (l.left_indent(), l.right_indent(), l.first_line_indent()),
            (left, 0, first)
        );
    }
    assert_eq!(
        p.edit().commit().unwrap().package().to_bytes().unwrap(),
        bytes
    );
}

#[test]
fn native_script_fonts_and_associated_languages_survive_model_writing() {
    let bytes = native(BIDI);
    let p = Package::from_bytes(&bytes).unwrap();
    let body = p.body().paragraphs().collect::<Vec<_>>();
    assert_eq!(body.len(), 7);
    for (paragraph, run, text, font, associated_font, associated_language) in [
        (2, 1, "مرحبا", "Arial (Arabic)", 1555, 1025),
        (3, 1, "שלום", "Arial (Hebrew)", 1554, 1037),
        (4, 1, "日", "Hiragino Mincho ProN W3", 180, 1025),
        (5, 1, "中", "Songti SC Light", 345, 1025),
    ] {
        let r = body[paragraph].runs().nth(run).unwrap();
        assert_eq!(r.text(), text);
        assert_eq!(r.format().font().unwrap().name(), font);
        assert_eq!(r.format().size().get(), 24);
        // Explicit local LTR state is separate from script/language metadata.
        assert_eq!(r.format().direction(), Some(TextDirection::LeftToRight));
        let raw = p
            .raw()
            .blocks()
            .iter()
            .find(|b| b.text == text)
            .unwrap()
            .formatting;
        assert_eq!(
            raw.language,
            Some(litchi_rtf::LanguageId::new(1033).unwrap())
        );
        assert_eq!(raw.associated.font_ref, Some(associated_font));
        assert_eq!(raw.associated.font_size.unwrap().get(), 24);
        assert_eq!(
            raw.associated.language,
            Some(litchi_rtf::LanguageId::new(associated_language).unwrap())
        );
    }
    let mut output = Vec::new();
    litchi_rtf::write::Writer::new(&mut output)
        .write_document(p.raw())
        .unwrap();
    let written = Package::from_bytes(&output).unwrap();
    assert_eq!(written.text(), p.text());
    assert_eq!(characters(&written), characters(&p));
    assert_eq!(paragraphs(&written), paragraphs(&p));
    assert_eq!(written.opaque(), p.opaque());
    assert_metadata(&p, &written);
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_run_format_batches_preserve_other_facets_and_prior_staged_changes() {
    for name in [MINIMAL, BIDI] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let mut expected = characters(&p);
        let first = p.body().paragraphs().next().unwrap().to_text();
        let size = NonZeroU16::new(28).unwrap();
        let patch = RunFormatPatch::new()
            .with_bold(true)
            .with_italic(true)
            .with_underline(UnderlineStyle::Double)
            .with_size(size);
        let mut updates = vec![RunFormatUpdate::new(
            Position::new(0),
            Position::new(0),
            patch,
        )];
        let mut text_targets = vec![first.as_str()];
        if name == BIDI {
            updates.push(RunFormatUpdate::new(
                Position::new(2),
                Position::new(1),
                patch,
            ));
            text_targets.push("مرحبا");
        }
        for target in text_targets {
            let offset = p.text().find(target).unwrap();
            let start = p.text()[..offset].chars().filter(|c| *c != '\n').count();
            for (_, f) in &mut expected[start..start + target.chars().count()] {
                f.bold = true;
                f.italic = true;
                f.underline = UnderlineStyle::Double;
                f.font_size = size;
            }
        }
        let mut edit = p.edit();
        assert!(matches!(
            edit.patch_body_run_formats(&[
                updates[0],
                RunFormatUpdate::new(Position::new(0), Position::new(99), patch)
            ]),
            Err(Error::RunOutOfRange {
                paragraph: 0,
                run: 99,
                ..
            })
        ));
        assert_eq!(edit.operation_count(), 0);
        edit.patch_body_run_formats(&updates).unwrap();
        assert!(matches!(
            edit.patch_run_format(Position::new(99), Position::new(0), patch),
            Err(Error::ParagraphOutOfRange { position: 99, .. })
        ));
        assert_eq!(edit.operation_count(), updates.len());
        let commit = edit.commit().unwrap();
        let output = save_reopen(&p, &commit);
        assert_eq!(output.text(), p.text());
        assert_eq!(characters(&output), expected);
        assert_eq!(paragraphs(&output), paragraphs(&p));
        assert_eq!(output.opaque(), p.opaque());
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_frame_crud_preserves_opaque_tokens_and_unicode_character_coverage() {
    for name in [MINIMAL, BIDI] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let original_paragraphs = paragraphs(&p);
        let original_characters = characters(&p);
        let mut create = p.edit();
        create
            .patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new()
                    .with_frame(frame())
                    .with_space_before(120),
            )
            .unwrap();
        let invalid = ParagraphFrame {
            width_twips: Some(u32::MAX),
            ..frame()
        };
        assert!(matches!(
            create.patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new().with_frame(invalid)
            ),
            Err(Error::Rtf(_))
        ));
        assert_eq!(create.operation_count(), 1);
        let created = create.commit().unwrap();
        let created = save_reopen(&p, &created);
        let mut expected = original_paragraphs.clone();
        expected[0].1.frame = Some(frame());
        expected[0].1.spacing.before = 120;
        assert_eq!(paragraphs(&created), expected);
        assert_eq!(characters(&created), original_characters);
        assert_eq!(created.text(), p.text());
        assert_eq!(created.opaque(), p.opaque());
        let next = ParagraphFrame {
            width_twips: Some(960),
            horizontal_position: HPos::Right,
            wrap: Wrap::Through,
            ..frame()
        };
        let mut update = created.edit();
        update
            .patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new().with_frame(next),
            )
            .unwrap();
        let updated = update.commit().unwrap();
        let updated = save_reopen(&created, &updated);
        expected[0].1.frame = Some(next);
        assert_eq!(paragraphs(&updated), expected);
        assert_eq!(characters(&updated), original_characters);
        let mut remove = updated.edit();
        remove
            .patch_paragraph_layout(
                Position::new(0),
                ParagraphLayoutPatch::new()
                    .clear_frame()
                    .clear_space_before(),
            )
            .unwrap();
        let removed = remove.commit().unwrap();
        let removed = save_reopen(&updated, &removed);
        assert_eq!(paragraphs(&removed), original_paragraphs);
        assert_eq!(characters(&removed), original_characters);
        assert_eq!(removed.opaque(), p.opaque());
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_footer_frames_and_page_fields_survive_an_independent_header_edit() {
    let bytes = native(TABLES);
    let p = Package::from_bytes(&bytes).unwrap();
    let expected_frame = ParagraphFrame {
        horizontal_reference: HRef::Margin,
        horizontal_position: HPos::Right,
        vertical_reference: VRef::Paragraph,
        vertical_position: VPos::Offset(0),
        overlay: true,
        ..ParagraphFrame::default()
    };
    let expected = vec![
        (Story::FooterLeft, 0, expected_frame),
        (Story::FooterRight, 0, expected_frame),
    ];
    assert_eq!(footer_frames(&p), expected);
    assert_eq!(p.raw().fields().len(), 2);
    assert!(
        p.raw()
            .fields()
            .iter()
            .all(|f| f.field_type == litchi_rtf::FieldType::Page
                && f.owner == litchi_rtf::FieldOwner::Footer)
    );
    assert!(
        p.body()
            .paragraphs()
            .all(|para| para.format().frame().is_none())
    );
    let mut edit = p.edit();
    edit.set_header_footer_text(
        edit::HeaderFooterParagraph::new(0, Story::HeaderRight, 0),
        "HEADER_REVIEWED_002",
    )
    .unwrap();
    assert!(
        matches!(edit.set_header_footer_text(edit::HeaderFooterParagraph::new(0,Story::FooterRight,1),"refused"),Err(Error::Refused(r)) if r.reason()=="header/footer text has dependent positioned content")
    );
    assert_eq!(edit.operation_count(), 1);
    let commit = edit.commit().unwrap();
    let output = save_reopen(&p, &commit);
    let source = std::str::from_utf8(&bytes).unwrap();
    assert_eq!(source.matches("HEADER_DEFAULT_002").count(), 1);
    assert_eq!(
        output.to_bytes().unwrap(),
        source
            .replace("HEADER_DEFAULT_002", "HEADER_REVIEWED_002")
            .as_bytes()
    );
    assert_eq!(footer_frames(&output), expected);
    assert_eq!(output.text(), p.text());
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_dependent_formatting_boundaries_keep_original_artifacts() {
    for name in [MATH, TABLES] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let mut layout = p.edit();
        assert!(
            matches!(layout.patch_paragraph_layout(Position::new(0),ParagraphLayoutPatch::new().with_frame(frame())),Err(Error::Refused(r)) if r.reason()=="paragraph-layout editing refuses opaque dependent destinations")
        );
        assert_eq!(layout.operation_count(), 0);
        assert_eq!(
            layout.commit().unwrap().package().to_bytes().unwrap(),
            bytes
        );
        let mut edit = p.edit();
        edit.patch_run_format(
            Position::new(0),
            Position::new(0),
            RunFormatPatch::new().with_bold(true),
        )
        .unwrap();
        assert!(
            matches!(edit.commit(),Err(Error::Refused(r)) if r.reason()=="run-format editing refuses dependent body destinations or review content")
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}
