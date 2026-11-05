#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "run edit tests use panic-driven assertions"
)]

//! Character runs and text hyperlinks through the canonical primary edit
//! (record 0323).

use litchi_ppt::Package;
use litchi_ppt::hyperlink::{Click, Target};
use litchi_ppt::package::edit::{Position, Target as ShapeTarget};
use litchi_ppt::text_edit::{RunFormat, TextRuns};
use std::path::PathBuf;

fn fixture(path: &str) -> Vec<u8> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path);
    std::fs::read(path).expect("PPT fixture")
}

fn runs(package: &Package, slide: usize, shape: usize) -> TextRuns {
    let presentation = package.presentation().unwrap();
    let selected = presentation.slide_at(slide).unwrap().unwrap();
    let id = selected.shapes().unwrap()[shape].id();
    selected.shape_text_runs(id).unwrap().expect("text body")
}

fn built() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Builder::new();
    let slide = writer.add_slide().unwrap();
    writer
        .add_textbox(slide, 10, 10, 400, 60, "Hello brave new world")
        .unwrap();
    writer.add_slide().unwrap();
    let mut output = std::io::Cursor::new(Vec::new());
    litchi_ppt::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

#[test]
fn powerpoint_runs_read_with_their_formatting() {
    let deck = Package::from_bytes(&fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let presentation = deck.presentation().unwrap();
    let slide = presentation.slide_at(0).unwrap().unwrap();
    let mut seen_italic = false;
    for shape in slide.shapes().unwrap().iter() {
        let Some(body) = slide.shape_text_runs(shape.id()).unwrap() else {
            continue;
        };
        let units = body.text().encode_utf16().count();
        assert_eq!(
            body.runs().iter().map(|run| run.length).sum::<usize>(),
            units + 1
        );
        seen_italic |= body
            .runs()
            .iter()
            .any(|run| run.format.font_style().is_some_and(|style| style.italic()));
    }
    assert!(seen_italic);
}

#[test]
fn run_format_splits_runs_and_round_trips() {
    let source_bytes = built();
    let source = Package::from_bytes(&source_bytes).unwrap();
    let before = runs(&source, 0, 0);
    let target = ShapeTarget::new(Position::new(0), Position::new(0));
    let format = RunFormat {
        bold: Some(true),
        size: Some(28),
        color: Some([0x1F, 0x4E, 0x79]),
        ..RunFormat::default()
    };
    let mut edit = source.edit().unwrap();
    edit.set_run_format(target, 6, 5, &format).unwrap();
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let after = runs(&published, 0, 0);
    assert_eq!(after.text(), before.text());
    let bold = after
        .runs()
        .iter()
        .find(|run| run.start == 6)
        .expect("split at the range start");
    assert_eq!(bold.length, 5);
    assert!(bold.format.font_style().unwrap().bold());
    assert_eq!(bold.format.font_size(), Some(28));
    let color = bold.format.color().unwrap();
    assert_eq!((color.red, color.green, color.blue), (0x1F, 0x4E, 0x79));
    assert!(
        after
            .runs()
            .iter()
            .filter(|run| run.start != 6)
            .all(|run| run.format.font_style().is_none_or(|style| !style.bold()))
    );
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        source_bytes
    );

    // Applying the same format again is an exact no-op.
    let mut edit = published.edit().unwrap();
    edit.set_run_format(target, 6, 5, &format).unwrap();
    assert!(edit.commit().unwrap().patch().is_noop());
    // A range past the text is refused.
    let mut edit = source.edit().unwrap();
    assert!(edit.set_run_format(target, 20, 9, &format).is_err());
}

#[test]
fn outline_placeholder_runs_are_formatted_in_the_slide_list() {
    let source = Package::from_bytes(&fixture("ole/ppt/SampleShow.ppt")).unwrap();
    let title = runs(&source, 0, 0);
    let mut edit = source.edit().unwrap();
    edit.set_run_format(
        ShapeTarget::new(Position::new(0), Position::new(0)),
        0,
        3,
        &RunFormat {
            underline: Some(true),
            ..RunFormat::default()
        },
    )
    .unwrap();
    let published =
        Package::from_bytes(&edit.commit().unwrap().package().to_bytes().unwrap()).unwrap();
    let after = runs(&published, 0, 0);
    assert_eq!(after.text(), title.text());
    let first = &after.runs()[0];
    assert_eq!((first.start, first.length), (0, 3));
    assert!(first.format.font_style().unwrap().underline());
}

#[test]
fn text_hyperlinks_add_and_remove() {
    let source_bytes = built();
    let source = Package::from_bytes(&source_bytes).unwrap();
    let target = ShapeTarget::new(Position::new(0), Position::new(0));
    let click = Click {
        target: Target::Url("https://example.com/brave".into()),
        tooltip: None,
    };
    let mut edit = source.edit().unwrap();
    edit.set_text_hyperlink(target, 6, 5, Some(&click)).unwrap();
    let jump = Click {
        target: Target::Slide(Position::new(1)),
        tooltip: Some("Next".into()),
    };
    edit.set_text_hyperlink(target, 12, 3, Some(&jump)).unwrap();
    let commit = edit.commit().unwrap();
    let published = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let body = runs(&published, 0, 0);
    let presentation = published.presentation().unwrap();
    let links = body
        .links()
        .iter()
        .map(|link| {
            (
                link.start,
                link.end,
                presentation.resolve_click(&link.interaction).unwrap(),
            )
        })
        .collect::<Vec<_>>();
    assert_eq!(links, [(6, 11, Some(click.clone())), (12, 15, Some(jump))]);
    // Overlapping another link is refused; removal drops the pair.
    let mut edit = published.edit().unwrap();
    assert!(edit.set_text_hyperlink(target, 8, 6, Some(&click)).is_err());
    let mut edit = published.edit().unwrap();
    edit.set_text_hyperlink(target, 6, 5, None).unwrap();
    let removed = edit.commit().unwrap();
    assert_eq!(runs(removed.package(), 0, 0).links().len(), 1);
}

#[test]
fn outline_text_hyperlinks_are_read() {
    let deck = Package::from_bytes(&fixture("poi/test-data/slideshow/WithLinks.ppt")).unwrap();
    let presentation = deck.presentation().unwrap();
    let mut urls = Vec::new();
    for slide in presentation.slides().unwrap() {
        for shape in slide.shapes().unwrap().iter() {
            if let Some(body) = slide.shape_text_runs(shape.id()).unwrap() {
                for link in body.links() {
                    if let Some(Click {
                        target: Target::Url(url),
                        ..
                    }) = presentation.resolve_click(&link.interaction).unwrap()
                    {
                        urls.push(url);
                    }
                }
            }
        }
    }
    assert!(
        urls.contains(&"http://jakarta.apache.org/poi/".to_owned()),
        "{urls:?}"
    );
}
