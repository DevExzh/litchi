use litchi_rtf::write::Writer;
use litchi_rtf::{Document, read};
use litchi_rtf::{text, text::Inline};

fn assert_snapshot_traits<T: Clone + Send + Sync>() {}
fn assert_borrowed_view_traits<T: Copy + Send + Sync>() {}

#[test]
fn document_is_a_small_shared_snapshot() {
    assert_snapshot_traits::<Document>();
    assert_borrowed_view_traits::<litchi_rtf::font::Catalog<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::font::Font<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::color::Palette<'static>>();
    assert_borrowed_view_traits::<litchi_rtf::color::Value>();
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
fn font_catalog_hides_sparse_rtf_slots_and_resolves_runs() {
    let document = Document::parse(
        r"{\rtf1\ansi{\fonttbl{\f7\fswiss Sparse Sans;}{\f42\froman Story Serif;}}\f42 text}",
    )
    .unwrap();

    let fonts = document.fonts();
    assert_eq!(fonts.len(), 2);
    assert_eq!(fonts.iter().count(), 2);
    assert_eq!(fonts.at(0).map(|font| font.name()), Some("Sparse Sans"));
    assert_eq!(fonts.at(1).map(|font| font.name()), Some("Story Serif"));
    assert!(fonts.at(2).is_none());
    assert_eq!(
        fonts.find("Story Serif").unwrap().map(|font| font.name()),
        Some("Story Serif")
    );
    assert!(fonts.find("missing").unwrap().is_none());

    let run = document.body().runs().next().unwrap();
    assert_eq!(
        run.format().font().map(|font| font.name()),
        Some("Story Serif")
    );
}

#[test]
fn font_lookup_reports_ambiguous_names() {
    let document =
        Document::parse(r"{\rtf1\ansi{\fonttbl{\f1\fswiss Shared;}{\f9\froman Shared;}}\f1 text}")
            .unwrap();

    assert!(matches!(
        document.fonts().find("Shared"),
        Err(litchi_rtf::font::LookupError::AmbiguousName)
    ));
}

#[test]
fn run_format_resolves_checked_color_values() {
    let document = Document::parse(
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

    let reparsed = Document::from_bytes(&document.to_bytes().unwrap()).unwrap();
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
fn snapshot_to_bytes_matches_streaming_facade() {
    let document = Document::parse(r"{\rtf1\ansi concise}").unwrap();
    let direct = document.to_bytes().unwrap();
    let reparsed = Document::from_bytes(&direct).unwrap();
    assert_eq!(reparsed.text(), "concise");
}

#[test]
fn story_lazily_distinguishes_lines_paragraphs_and_empty_paragraphs() {
    let document = Document::parse(r"{\rtf1\ansi\qc\b one\line two\par\par three}").unwrap();
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
fn structural_break_kinds_survive_snapshot_round_trip() {
    let document = Document::parse(r"{\rtf1 first\line second\par third}").unwrap();
    let encoded = document.to_bytes().unwrap();
    let source = std::str::from_utf8(&encoded).unwrap();

    assert!(source.contains("\\line "));
    assert!(source.contains("\\par "));

    let reparsed = Document::from_bytes(&encoded).unwrap();
    let paragraphs: Vec<_> = reparsed
        .body()
        .paragraphs()
        .map(|paragraph| paragraph.to_text())
        .collect();
    assert_eq!(paragraphs, ["first\nsecond", "third"]);
}

#[test]
fn literal_line_feed_is_text_not_a_structural_break() {
    let document = Document::parse(r"{\rtf1 A\u10?B}").unwrap();
    assert_eq!(document.paragraph_count(), 1);
    assert_eq!(
        document
            .body()
            .paragraphs()
            .next()
            .map(|value| value.to_text()),
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

    let reparsed = Document::from_bytes(&canonical).unwrap();
    assert_eq!(reparsed.text(), "A\nB");
    assert_eq!(reparsed.paragraph_count(), 1);
}
