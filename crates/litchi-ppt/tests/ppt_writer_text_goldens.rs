#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::too_many_lines,
    reason = "golden fixtures favor explicit inputs and panic-driven assertions"
)]

//! Byte-exact goldens for the fresh PPT writer's text and drawing paths.
//!
//! Every fixture below was written by the writer at `6d989cad63`, before
//! change 0753 wrote `ClientTextbox`, shape, drawing and slide records in
//! place, and its SHA-256 recorded here. The same inputs must keep producing
//! the same bytes: ASCII, Latin-1, CJK and supplementary-plane text, empty and
//! very long text, text interactions, rich text, tables, pictures, comments,
//! transitions, timings, per-slide header/footer records, notes (through
//! `save`), and the three benchmark writer shapes.

use litchi_core::validation::EvidenceDigest;
use litchi_ppt::writer::text_format::{Paragraph, TextRun};
use litchi_ppt::writer::{Hyperlink, SlideComment, SlideTiming, Table, Writer};
use litchi_ppt::{
    AdvanceMode, HeaderFooter, HeaderFooterOptions, HeaderFooterScope, Interaction,
    InteractionAction, InteractionJump, InteractionLinkTarget, InteractionTrigger, TextInteraction,
    TextRange, TransitionInfo, TransitionSpeed, TransitionType,
};
use std::io::Cursor;

fn repeat_to(seed: &str, bytes: usize) -> String {
    let mut text = String::with_capacity(bytes + seed.len());
    while text.len() < bytes {
        text.push_str(seed);
    }
    text
}

fn written(mut writer: Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ascii_text_boxes() -> Vec<u8> {
    let mut writer = Writer::new();
    let first = writer.add_slide().unwrap();
    writer.add_textbox(first, 10, 10, 200, 40, "").unwrap();
    writer.add_textbox(first, 10, 60, 200, 40, "a").unwrap();
    writer
        .add_textbox(first, 10, 110, 200, 40, "Plain ASCII text box")
        .unwrap();
    writer
        .add_textbox(
            first,
            10,
            160,
            200,
            40,
            "two\rparagraphs\x0bwith a line break",
        )
        .unwrap();
    let second = writer.add_slide().unwrap();
    writer
        .add_textbox(
            second,
            10,
            10,
            400,
            300,
            &repeat_to("ascii payload ", 40_000),
        )
        .unwrap();
    writer
        .add_textbox(
            second,
            10,
            320,
            400,
            100,
            &repeat_to("beyond 64 KiB ", 70_001),
        )
        .unwrap();
    written(writer)
}

fn unicode_text_boxes() -> Vec<u8> {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 200, 40, "Latin-1: àéîõüçñ ß ÿ")
        .unwrap();
    writer
        .add_textbox(slide, 10, 60, 200, 40, "CJK: 漢字かなカナ한국어")
        .unwrap();
    writer.add_textbox(slide, 10, 110, 200, 40, "😀").unwrap();
    writer
        .add_textbox(
            slide,
            10,
            160,
            200,
            40,
            "\u{80}\u{7ff}\u{800}\u{ffff}\u{10000}\u{10ffff}",
        )
        .unwrap();
    let long = writer.add_slide().unwrap();
    writer
        .add_textbox(
            long,
            10,
            10,
            400,
            300,
            &repeat_to("混合 mixed ✓ 😀 text ", 90_000),
        )
        .unwrap();
    written(writer)
}

fn text_interactions() -> Vec<u8> {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer.add_textbox(slide, 10, 10, 240, 40, "A😀BC").unwrap();
    let hyperlink_id = writer.add_hyperlink(Hyperlink::url("https://example.invalid/text"));
    writer
        .set_last_shape_text_hyperlink(slide, TextRange::new(1, 3).unwrap(), hyperlink_id)
        .unwrap();
    let hover = TextInteraction::new(
        TextRange::new(3, 5).unwrap(),
        Interaction::new(
            InteractionTrigger::MouseOver,
            InteractionAction::RunProgram,
            InteractionLinkTarget::OtherFile,
        )
        .with_macro_name("viewer.exe")
        .unwrap(),
    )
    .unwrap();
    writer
        .set_last_shape_text_interaction(slide, hover)
        .unwrap();
    let mut shape_action = Interaction::new(
        InteractionTrigger::Click,
        InteractionAction::Jump,
        InteractionLinkTarget::NextSlide,
    );
    shape_action.jump = InteractionJump::NextSlide;
    writer
        .set_last_shape_interaction(slide, shape_action)
        .unwrap();
    writer
        .add_textbox(slide, 10, 60, 240, 40, "ascii with a link")
        .unwrap();
    let ascii_link = writer.add_hyperlink(Hyperlink::url("https://example.invalid/ascii"));
    writer
        .set_last_shape_text_hyperlink(slide, TextRange::new(0, 5).unwrap(), ascii_link)
        .unwrap();
    writer
        .add_rich_textbox(
            slide,
            10,
            110,
            240,
            60,
            vec![
                Paragraph::with_runs(vec![TextRun::new("Hi😀").bold(), TextRun::new(" plain")]),
                Paragraph::new("there"),
            ],
        )
        .unwrap();
    written(writer)
}

fn fixture_png() -> Vec<u8> {
    std::fs::read(
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/images/png/lena.png"),
    )
    .unwrap()
}

fn footer(text: &str) -> HeaderFooter {
    HeaderFooter {
        scope: HeaderFooterScope::PresentationSlides,
        options: HeaderFooterOptions {
            show_footer: true,
            ..HeaderFooterOptions::default()
        },
        user_date: None,
        header: None,
        footer: Some(text.to_string()),
        placeholder_display: None,
    }
}

fn mixed_slide_children() -> Vec<u8> {
    let mut writer = Writer::new();
    let first = writer.add_slide().unwrap();
    writer.add_rectangle(first, 20, 20, 100, 50).unwrap();
    writer
        .add_textbox(first, 20, 80, 300, 40, "Next to a table 表")
        .unwrap();
    let mut table = Table::new(2, 2).unwrap();
    table.set_cell_text(0, 0, "A1").unwrap();
    table.set_cell_text(0, 1, "漢字").unwrap();
    table.set_cell_text(1, 0, "😀").unwrap();
    table.set_cell_text(1, 1, "").unwrap();
    writer.add_table(first, 40, 200, table).unwrap();
    writer
        .add_picture(first, 400, 20, 120, 120, fixture_png())
        .unwrap();
    writer
        .add_comment(first, SlideComment::new("Alice 😀", "Great slide", 10, 20))
        .unwrap();
    writer
        .set_slide_transition(
            first,
            TransitionInfo::with_type(TransitionType::Dissolve)
                .with_speed(TransitionSpeed::Slow)
                .with_advance_mode(AdvanceMode::Automatic)
                .with_advance_time(3000),
        )
        .unwrap();
    writer
        .set_slide_header_footer(first, footer("Slide footer 页脚"))
        .unwrap();
    let second = writer.add_slide().unwrap();
    writer
        .add_textbox(second, 20, 20, 300, 40, "Timed slide")
        .unwrap();
    writer
        .set_slide_timing(second, SlideTiming::default())
        .unwrap();
    writer.add_slide().unwrap();
    written(writer)
}

static NEXT_FILE: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn saved_with_notes() -> Vec<u8> {
    let mut writer = Writer::new();
    let first = writer.add_slide().unwrap();
    writer
        .add_textbox(first, 10, 10, 300, 40, "Slide with notes 😀")
        .unwrap();
    writer.set_slide_notes(first, "Speaker notes 漢字").unwrap();
    let second = writer.add_slide().unwrap();
    writer
        .add_textbox(second, 10, 10, 300, 40, &repeat_to("saved ", 5_000))
        .unwrap();
    let path = std::env::temp_dir().join(format!(
        "litchi-ppt-text-golden-{}-{}.ppt",
        std::process::id(),
        NEXT_FILE.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
    ));
    writer.save(&path).unwrap();
    let bytes = std::fs::read(&path).unwrap();
    std::fs::remove_file(&path).unwrap();
    bytes
}

fn harness_shape(slides: usize, boxes: usize, payload: Option<usize>) -> Vec<u8> {
    let mut writer = Writer::new();
    for slide_number in 0..slides {
        let slide = writer.add_slide().unwrap();
        for box_number in 0..boxes {
            let mut text = format!(
                "litchi-perf-baseline-ppt-v1-{slide_number:03}-{box_number:05}-000 deterministic payload"
            );
            if let Some(length) = payload {
                while text.len() < length {
                    text.push_str("litchi-perf-baseline-payload-heavy-v1 ");
                }
                text.truncate(length);
            }
            let x = 36 + i32::try_from(box_number % 3).unwrap() * 180;
            let y = 36 + i32::try_from(box_number / 3).unwrap() * 90;
            writer.add_textbox(slide, x, y, 144, 54, &text).unwrap();
        }
    }
    written(writer)
}

fn harness_tiny() -> Vec<u8> {
    harness_shape(1, 2, None)
}

fn harness_large() -> Vec<u8> {
    harness_shape(12, 12, None)
}

fn harness_payload_heavy() -> Vec<u8> {
    harness_shape(16, 8, Some(40_000))
}

#[test]
fn fresh_ppt_writer_text_paths_match_the_pre_0753_goldens() {
    let fixtures: [(&str, fn() -> Vec<u8>, &str); 9] = [
        (
            "ascii_text_boxes",
            ascii_text_boxes,
            "46cc4291a83541d2823b07cbb454ee39b80f33353af504bbc0a41cf2f68bcc0e",
        ),
        (
            "unicode_text_boxes",
            unicode_text_boxes,
            "6f5cb6298f6db4271b6ac2dc46eaac58a899df00aa0159701ea4ebdaa7838b94",
        ),
        (
            "text_interactions",
            text_interactions,
            "50926d62d243b1a5b703899cd0005f03e941168931fa8d4e685dd9de7a04a776",
        ),
        (
            "mixed_slide_children",
            mixed_slide_children,
            "83ab9e77c32a95dbd3b913319afefd000ba3ee7ebc1a94aa716f40939a515f71",
        ),
        (
            "saved_with_notes",
            saved_with_notes,
            "474cedf8347135a9b6b948ecbb24dab5385fdf68bcb3c4ba7a6bc8a676e2e034",
        ),
        (
            "harness_tiny",
            harness_tiny,
            "e233c6b63928578c2429178c3ac8589b32d73d44df9953ac18c9d27f6968d8b4",
        ),
        (
            "harness_large",
            harness_large,
            "229052cd918c0e5b7ef44070bafe20833531eee119b5943b18499503e225ff52",
        ),
        (
            "harness_payload_heavy",
            harness_payload_heavy,
            "51686d86b2ca22c444c382565d065be7d92aecd9b6f7fdfb8597a2593d155ffb",
        ),
        (
            "empty_presentation",
            || written(Writer::new()),
            "42248efbfd7169bb976fc1a31ef163bde3239fe4348913f8dfcd77e8f30317b0",
        ),
    ];
    let mut mismatches = Vec::new();
    for (name, build, expected) in fixtures {
        let first = build();
        assert_eq!(first, build(), "{name} is not deterministic");
        let actual = EvidenceDigest::of(&first).to_string();
        if actual != expected {
            mismatches.push(format!("{name}: {actual} ({} bytes)", first.len()));
        }
    }
    assert!(
        mismatches.is_empty(),
        "digest mismatches:\n{}",
        mismatches.join("\n")
    );
}

/// Writing the same writer twice gives the same bytes, and a refused write
/// leaves the destination untouched even though slide records are now
/// assembled in place.
#[test]
fn a_writer_written_twice_writes_the_same_bytes_and_a_refusal_writes_nothing() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 300, 40, &repeat_to("twice ✓ ", 50_000))
        .unwrap();
    let mut first = Cursor::new(Vec::new());
    writer.write_to(&mut first).unwrap();
    let mut second = Cursor::new(Vec::new());
    writer.write_to(&mut second).unwrap();
    assert_eq!(first.get_ref(), second.get_ref());

    // A later slide is refused while it is being assembled: its drawing is
    // already in the stream when the over-long comment author is rejected.
    let refused = writer.add_slide().unwrap();
    writer
        .add_textbox(refused, 10, 10, 100, 40, &repeat_to("in place ", 20_000))
        .unwrap();
    writer
        .add_comment(refused, SlideComment::new(&"a".repeat(60), "text", 1, 1))
        .unwrap();
    let mut output = Cursor::new(vec![0xA5; 16]);
    let error = writer.write_to(&mut output).unwrap_err();
    assert!(error.to_string().contains("comment"), "{error}");
    assert_eq!(output.into_inner(), vec![0xA5; 16]);
}
