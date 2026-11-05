#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_core::Position;
use litchi_rtf::write::Writer;
use litchi_rtf::{Document, OpenOptions, Package, SourceLimits};
use litchi_rtf::{text, text::Inline};
use std::io::{self, Write};

struct FailAfter {
    accepted: usize,
    limit: usize,
}

impl Write for FailAfter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted >= self.limit {
            return Err(io::Error::other("injected RTF sink failure"));
        }
        let accepted = bytes.len().min(self.limit - self.accepted);
        self.accepted += accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn assert_snapshot_traits<T: Clone + Send + Sync>() {}
fn assert_borrowed_view_traits<T: Copy + Send + Sync>() {}

#[test]
fn retained_paragraph_count_matches_hidden_and_structural_content() {
    for source in [
        r"{\rtf1\par}",
        r"{\rtf1\line}",
        r"{\rtf1{\*\revtbl{A;}}visible{\deleted\revauthdel0 hidden\par} tail}",
        r"{\rtf1{\header hidden\par header tail}visible\par tail}",
        r"{\rtf1{\*\unknown hidden\par}visible\par tail}",
        r"{\rtf1\trowd\cellx1000\intbl cell\cell\row visible\par tail}",
    ] {
        let document = Package::parse(source).unwrap();
        assert_eq!(
            document.paragraph_count(),
            document.body().paragraphs().count(),
            "retained paragraph count differed for {source:?}"
        );
    }
}

#[test]
fn document_is_a_small_shared_snapshot() {
    assert_snapshot_traits::<Package>();
    assert_snapshot_traits::<Document>();
    assert_borrowed_view_traits::<litchi_rtf::font::Catalog<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::font::Font<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::color::Palette<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::color::Value>();
    assert_eq!(size_of::<Package>(), size_of::<usize>());
    assert_eq!(size_of::<Document>(), size_of::<usize>());

    let document = Package::parse(r"{\rtf1\ansi first\par second}").unwrap();
    let shared = document.clone();

    assert!(document.same_snapshot(&shared));
    assert_eq!(document.paragraph_count(), 2);
    assert_eq!(document.text(), shared.text());
    assert!(std::ptr::eq(document.text(), shared.text()));
}

#[test]
fn package_projection_survives_owner_drop_and_retains_open_policy() {
    let limits = SourceLimits::builder()
        .max_input_bytes(512)
        .expect("source-byte limit is representable")
        .build()
        .expect("source profile is consistent");
    let package = Package::parse_with_options(
        r"{\rtf1\ansi first\par second}",
        OpenOptions::new().source_limits(limits),
    )
    .unwrap();
    assert_eq!(package.source_limits(), limits);
    assert_eq!(package.open_options().source_limits_value(), limits);

    let document = package.document();
    let clone = document.clone();
    assert!(std::ptr::eq(package.text(), document.text()));
    assert!(std::ptr::eq(document.text(), clone.text()));
    assert_eq!(document.text(), "first\nsecond");
    assert_eq!(document.paragraph_count(), 2);

    drop(package);
    assert_eq!(document.text(), "first\nsecond");
    assert_eq!(document.body().paragraphs().count(), 2);
}

#[test]
fn contextual_reader_uses_checked_limits() {
    let limits = SourceLimits::builder()
        .max_input_bytes(8)
        .expect("source-byte limit is representable")
        .build()
        .expect("source profile is consistent");
    assert!(
        Package::parse_with_options(
            r"{\rtf1 too large}",
            OpenOptions::new().source_limits(limits)
        )
        .is_err()
    );
}

#[test]
fn snapshot_round_trips_through_concise_writer() {
    let document =
        Package::parse(r"{\rtf1\ansi{\fonttbl{\f0\fswiss Helvetica;}}\f0\pard Hello!\par}")
            .unwrap();

    let mut bytes = Vec::new();
    Writer::new(&mut bytes).write(&document).unwrap();
    let reparsed = Package::from_bytes(&bytes).unwrap();

    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.fonts().len(), 1);
    assert!(!reparsed.is_empty());
}

#[test]
fn font_catalog_hides_sparse_rtf_slots_and_resolves_runs() {
    let document = Package::parse(
        r"{\rtf1\ansi{\fonttbl{\f7\fswiss Sparse Sans;}{\f42\froman Story Serif;}}\f42 text}",
    )
    .unwrap();

    let fonts = document.fonts();
    assert_eq!(fonts.len(), 2);
    assert_eq!(fonts.iter().count(), 2);
    assert_eq!(
        fonts.at(0).map(litchi_rtf::font::Font::name),
        Some("Sparse Sans")
    );
    assert_eq!(
        fonts.at(1).map(litchi_rtf::font::Font::name),
        Some("Story Serif")
    );
    assert!(fonts.at(2).is_none());
    assert_eq!(
        fonts
            .find("Story Serif")
            .unwrap()
            .map(litchi_rtf::font::Font::name),
        Some("Story Serif")
    );
    assert!(fonts.find("missing").unwrap().is_none());

    let run = document.body().runs().next().unwrap();
    assert_eq!(
        run.format().font().map(litchi_rtf::font::Font::name),
        Some("Story Serif")
    );
}

#[test]
fn font_lookup_reports_ambiguous_names() {
    let document =
        Package::parse(r"{\rtf1\ansi{\fonttbl{\f1\fswiss Shared;}{\f9\froman Shared;}}\f1 text}")
            .unwrap();

    assert!(matches!(
        document.fonts().find("Shared"),
        Err(litchi_rtf::font::LookupError::AmbiguousName)
    ));
}

#[test]
fn run_format_resolves_checked_color_values() {
    let document = Package::parse(
        r"{\rtf1\ansi{\fonttbl{\f0\fnil Arial;}}{\colortbl;\red12\green34\blue56;}\f0\cf1\cb1\highlight1\ul\ulc1 colored}",
    )
    .unwrap();

    let expected = litchi_rtf::color::Color::new(12, 34, 56);
    let colors = document.colors();
    assert_eq!(colors.len(), 2);
    assert_eq!(colors.at(0), Some(litchi_rtf::color::Value::Automatic));
    assert_eq!(colors.at(1), Some(litchi_rtf::color::Value::Rgb(expected)));
    assert_eq!(
        colors.iter().next_back(),
        Some(litchi_rtf::color::Value::Rgb(expected))
    );

    let format = document.body().runs().next().unwrap().format();
    assert_eq!(
        format.foreground(),
        Some(litchi_rtf::color::Value::Rgb(expected))
    );
    assert_eq!(format.foreground_color(), Some(expected));
    assert_eq!(format.background_color(), Some(expected));
    assert_eq!(format.highlight_color(), Some(expected));
    assert_eq!(format.underline_color(), Some(expected));

    let reparsed = Package::from_bytes(&document.to_bytes().unwrap()).unwrap();
    assert_eq!(
        reparsed.colors().at(0),
        Some(litchi_rtf::color::Value::Automatic)
    );
    assert_eq!(
        reparsed.colors().at(1),
        Some(litchi_rtf::color::Value::Rgb(expected))
    );
}

#[test]
fn inline_character_controls_split_native_runs_and_restore_groups() {
    let document = Package::parse(
        r"{\rtf1\ansi{\fonttbl{\f0\fnil Base;}{\f1\fswiss Alt;}}{\colortbl;\red255\green0\blue0;}\f0 base\f1 alt\f0 base_again\i italic\i0 italic_off\fs28 big\fs24 size_off\uldb under_double\ulnone under_off\strike struck\strike0 strike_off\striked dstruck\striked0 dstrike_off\v hidden\v0 hidden_off\caps caps\caps0 caps_off\scaps smallcaps\scaps0 smallcaps_off\outl outline\outl0 outline_off\cf1 red\cf0 color_off{\b grouped_bold{\i nested_italic}grouped_restored}after_group}",
    )
    .unwrap();

    let runs: Vec<_> = document.body().runs().collect();
    let texts: Vec<_> = runs.iter().map(|run| run.text()).collect();
    assert_eq!(
        texts,
        vec![
            "base",
            "alt",
            "base_again",
            "italic",
            "italic_off",
            "big",
            "size_off",
            "under_double",
            "under_off",
            "struck",
            "strike_off",
            "dstruck",
            "dstrike_off",
            "hidden",
            "hidden_off",
            "caps",
            "caps_off",
            "smallcaps",
            "smallcaps_off",
            "outline",
            "outline_off",
            "red",
            "color_off",
            "grouped_bold",
            "nested_italic",
            "grouped_restored",
            "after_group",
        ]
    );

    let run = |text: &str| runs.iter().find(|run| run.text() == text).unwrap();
    assert_eq!(
        run("base")
            .format()
            .font()
            .map(litchi_rtf::font::Font::name),
        Some("Base")
    );
    assert_eq!(
        run("alt").format().font().map(litchi_rtf::font::Font::name),
        Some("Alt")
    );
    assert_eq!(
        run("base_again")
            .format()
            .font()
            .map(litchi_rtf::font::Font::name),
        Some("Base")
    );
    assert!(run("italic").format().italic());
    assert!(!run("italic_off").format().italic());
    assert_eq!(run("big").format().size().get(), 28);
    assert_eq!(run("size_off").format().size().get(), 24);
    assert_eq!(
        run("under_double").format().underline(),
        text::Underline::Double
    );
    assert_eq!(run("under_off").format().underline(), text::Underline::None);
    assert!(run("struck").format().strike());
    assert!(!run("strike_off").format().strike());
    assert!(run("dstruck").format().double_strike());
    assert!(!run("dstrike_off").format().double_strike());
    assert!(run("hidden").format().hidden());
    assert!(!run("hidden_off").format().hidden());
    assert!(run("caps").format().all_caps());
    assert!(!run("caps_off").format().all_caps());
    assert!(run("smallcaps").format().small_caps());
    assert!(!run("smallcaps_off").format().small_caps());
    assert!(run("outline").format().outline());
    assert!(!run("outline_off").format().outline());
    assert_eq!(
        run("red").format().foreground(),
        Some(litchi_rtf::color::Value::Rgb(
            litchi_rtf::color::Color::new(255, 0, 0)
        ))
    );
    assert_eq!(
        run("color_off").format().foreground(),
        Some(litchi_rtf::color::Value::Automatic)
    );

    assert!(run("grouped_bold").format().bold());
    assert!(!run("grouped_bold").format().italic());
    assert!(run("nested_italic").format().bold());
    assert!(run("nested_italic").format().italic());
    assert!(run("grouped_restored").format().bold());
    assert!(!run("grouped_restored").format().italic());
    assert!(!run("after_group").format().bold());
    assert!(!run("after_group").format().italic());
}

#[test]
fn every_native_underline_control_starts_a_distinct_run() {
    const STYLES: &[(&str, text::Underline)] = &[
        (r"\ul", text::Underline::Single),
        (r"\ulnone", text::Underline::None),
        (r"\uldb", text::Underline::Double),
        (r"\uld", text::Underline::Dotted),
        (r"\uldash", text::Underline::Dashed),
        (r"\uldashd", text::Underline::DashDot),
        (r"\uldashdd", text::Underline::DashDotDot),
        (r"\ulw", text::Underline::Words),
        (r"\ulth", text::Underline::Thick),
        (r"\ulwave", text::Underline::Wave),
        (r"\ulhair", text::Underline::Hairline),
        (r"\ulthd", text::Underline::ThickDotted),
        (r"\ulthdash", text::Underline::ThickDashed),
        (r"\ulthdashd", text::Underline::ThickDashDot),
        (r"\ulthdashdd", text::Underline::ThickDashDotDot),
        (r"\ulthldash", text::Underline::ThickLongDash),
        (r"\ulldash", text::Underline::LongDash),
        (r"\ulhwave", text::Underline::HeavyWave),
        (r"\ululdbwave", text::Underline::DoubleWave),
    ];
    let mut source = String::from(r"{\rtf1\ansi");
    for (index, (control, _)) in STYLES.iter().enumerate() {
        source.push_str(control);
        source.push(' ');
        source.push('u');
        source.push_str(&index.to_string());
        source.push_str(r"\ulnone off");
        source.push_str(&index.to_string());
    }
    source.push('}');

    let document = Package::parse(&source).unwrap();
    let runs: Vec<_> = document.body().runs().collect();
    assert_eq!(runs.len(), STYLES.len() * 2);
    for (index, (control, style)) in STYLES.iter().enumerate() {
        let styled_text = format!("u{index}");
        let reset_text = format!("off{index}");
        let styled = runs
            .iter()
            .find(|run| run.text() == styled_text.as_str())
            .unwrap();
        let reset = runs
            .iter()
            .find(|run| run.text() == reset_text.as_str())
            .unwrap();
        assert_eq!(styled.format().underline(), *style, "{control}");
        assert_eq!(
            reset.format().underline(),
            text::Underline::None,
            "reset after {control}"
        );
    }
}

#[test]
fn snapshot_to_bytes_matches_streaming_facade() {
    let document = Package::parse(r"{\rtf1\ansi concise}").unwrap();
    let direct = document.to_bytes().unwrap();
    let reparsed = Package::from_bytes(&direct).unwrap();
    assert_eq!(reparsed.text(), "concise");
}

#[test]
fn streaming_facade_propagates_partial_sink_failure_without_mutating_snapshot() {
    let source = br"{\rtf1\ansi immutable sequential output}";
    let document = Package::from_bytes(source).unwrap();
    let before = document.text().to_owned();
    let mut sink = FailAfter {
        accepted: 0,
        limit: source.len() / 2,
    };

    let error = document.write_to(&mut sink).unwrap_err();

    assert_eq!(error.kind(), io::ErrorKind::Other);
    assert_eq!(sink.accepted, source.len() / 2);
    assert_eq!(document.text(), before);
    assert_eq!(document.to_bytes().unwrap(), source);
}

#[test]
fn story_lazily_distinguishes_lines_paragraphs_and_empty_paragraphs() {
    let document = Package::parse(r"{\rtf1\ansi\qc\b one\line two\par\par three}").unwrap();
    assert_eq!(document.body().len(), document.text().len());
    assert_eq!(document.body().is_empty(), document.text().is_empty());
    let paragraphs: Vec<_> = document.body().paragraphs().collect();

    assert_eq!(document.paragraph_count(), 3);
    assert_eq!(paragraphs.len(), 3);
    assert_eq!(
        paragraphs.first().and_then(|paragraph| paragraph.as_str()),
        None
    );
    assert_eq!(
        paragraphs.first().map(|paragraph| paragraph.to_text()),
        Some("one\ntwo".to_string())
    );
    assert!(
        paragraphs
            .get(1)
            .is_some_and(|paragraph| paragraph.is_empty())
    );
    assert_eq!(
        paragraphs.get(2).and_then(|paragraph| paragraph.as_str()),
        Some("three")
    );
    assert_eq!(
        paragraphs
            .first()
            .map(|paragraph| paragraph.format().alignment()),
        Some(text::Alignment::Center)
    );

    let mut inlines = paragraphs.first().unwrap().inlines();
    let first = inlines.next();
    assert!(matches!(first, Some(Inline::Text(run)) if run.text() == "one" && run.format().bold()));
    assert!(matches!(
        inlines.next(),
        Some(Inline::Break(text::Break::Line))
    ));
    assert!(matches!(inlines.next(), Some(Inline::Text(run)) if run.text() == "two"));
    assert!(inlines.next().is_none());
}

#[test]
fn story_retains_exact_text_length_across_fragmentation_and_edits() {
    for source in [
        r"{\rtf1\ansi}",
        r"{\rtf1\ansi A{\b B}{\i C}\par D}",
        r"{\rtf1\ansi first\par\par trailing}",
        r"{\rtf1\ansi A\u10?B}",
    ] {
        let document = Package::parse(source).unwrap();
        assert_eq!(document.body().len(), document.text().len());
        assert_eq!(document.body().is_empty(), document.text().is_empty());
    }

    let source = Package::parse(r"{\rtf1\ansi old\par retained}").unwrap();
    let source_len = source.body().len();
    let projection = source.document();
    let mut edit = source.edit();
    edit.replace_paragraph_text(Position::new(0), "a longer replacement")
        .unwrap();
    let changed = edit.commit().unwrap().package().clone();

    assert_eq!(source.body().len(), source_len);
    assert_eq!(source.body().len(), source.text().len());
    assert_eq!(changed.body().len(), changed.text().len());
    assert_ne!(changed.body().len(), source.body().len());
    assert_eq!(projection.text(), "old\nretained");
}

#[test]
fn structural_break_kinds_survive_snapshot_round_trip() {
    let document = Package::parse(r"{\rtf1 first\line second\par third}").unwrap();
    let encoded = document.to_bytes().unwrap();
    let source = std::str::from_utf8(&encoded).unwrap();

    assert!(source.contains("\\line "));
    assert!(source.contains("\\par "));

    let reparsed = Package::from_bytes(&encoded).unwrap();
    let paragraphs: Vec<_> = reparsed
        .body()
        .paragraphs()
        .map(text::Paragraph::to_text)
        .collect();
    assert_eq!(paragraphs, ["first\nsecond", "third"]);
}

#[test]
fn literal_line_feed_is_text_not_a_structural_break() {
    let document = Package::parse(r"{\rtf1 A\u10?B}").unwrap();
    assert_eq!(document.paragraph_count(), 1);
    assert_eq!(
        document
            .body()
            .paragraphs()
            .next()
            .map(text::Paragraph::to_text),
        Some("A\nB".to_string())
    );
    assert!(
        !document
            .body()
            .inlines()
            .any(|inline| matches!(inline, Inline::Break(_)))
    );

    let encoded = document.to_bytes().unwrap();
    assert_eq!(encoded, br"{\rtf1 A\u10?B}");

    let mut canonical = Vec::new();
    Writer::with_options(
        &mut canonical,
        litchi_rtf::write::Options {
            indent: true,
            ..Default::default()
        },
    )
    .write(&document)
    .unwrap();
    let source = std::str::from_utf8(&canonical).unwrap();
    assert!(source.contains("\\'0a"));
    assert!(!source.contains("\\line "));
    assert!(!source.contains("\\par "));

    let reparsed = Package::from_bytes(&canonical).unwrap();
    assert_eq!(reparsed.text(), "A\nB");
    assert_eq!(reparsed.paragraph_count(), 1);
}
