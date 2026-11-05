//! Speaker-notes run styles (record 0329): character properties on a span
//! of a slide's notes, read back through the notes run reader, the text
//! kept, the inverse exact, over every corpus deck with notes.

use std::path::{Path, PathBuf};

use litchi_keynote::slide::drawable::{Error, Style, TextSpan, Unsupported};
use litchi_keynote::{Package, SlideSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn corpus() -> Vec<PathBuf> {
    let mut files = Vec::new();
    let mut stack =
        vec![Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote")];
    while let Some(dir) = stack.pop() {
        for entry in std::fs::read_dir(dir).expect("corpus directory") {
            let path = entry.expect("entry").path();
            if path.extension().is_some_and(|extension| extension == "key") && path.is_file() {
                files.push(path);
            } else if path.is_dir() {
                stack.push(path);
            }
        }
    }
    files.sort();
    files
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn notes_spans_are_styled_read_back_and_inverted() -> TestResult {
    let (mut styled, mut refused) = (0_usize, 0_usize);
    for path in corpus() {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Ok(slides) = package.slides().map(<[_]>::len) else {
            continue;
        };
        for slide in 0..slides {
            let Ok(Some(notes)) = package.slide_notes(SlideSelector::index(slide)) else {
                continue;
            };
            let length = notes.chars().count();
            if length == 0 {
                continue;
            }
            let span = TextSpan::new(0, length.min(4));
            let style = Style::new().with_italic(true).with_font_size(19.0);
            let mut edit = package.edit_slide_drawables();
            edit.set_notes_text_style(SlideSelector::index(slide), span, style.clone())?;
            let commit = match edit.commit() {
                Ok(commit) => commit,
                Err(Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                    reason: Unsupported::ForeignComponent | Unsupported::ExternalReferences,
                })) => {
                    refused += 1;
                    continue;
                },
                Err(error) => panic!("{} slide {slide}: {error:?}", path.display()),
            };
            let reopened = Package::from_bytes(&bytes(commit.package())?)?;
            assert_eq!(
                reopened
                    .slide_notes(SlideSelector::index(slide))?
                    .as_deref(),
                Some(notes.as_str()),
                "{}",
                path.display()
            );
            let runs = reopened.slide_notes_text_runs(SlideSelector::index(slide))?;
            let first = runs.first().expect("a run");
            assert_eq!(first.span().start(), 0, "{}", path.display());
            assert_eq!(first.style().italic, Some(true), "{}", path.display());
            assert_eq!(first.style().font_size, Some(19.0), "{}", path.display());
            let restored = commit
                .package()
                .apply_slide_drawables(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, bytes(&package)?);
            styled += 1;
        }
    }
    eprintln!("styled {styled}, refused {refused}");
    assert!(styled >= 2, "styled {styled}");
    Ok(())
}

#[test]
fn written_notes_take_styles() -> TestResult {
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/multi-slide.key");
    let package = Package::open(path)?;
    for slide in 0..package.slides()?.len() {
        let mut notes = package.edit_slide_notes(SlideSelector::index(slide))?;
        notes.set("Remember the\nsecond point")?;
        let written = notes.commit()?.into_package();
        let mut edit = written.edit_slide_drawables();
        edit.set_notes_text_style(
            SlideSelector::index(slide),
            TextSpan::new(13, 19),
            Style::new().with_bold(true).with_color([0xc0, 0x10, 0x20]),
        )?;
        let commit = edit.commit()?;
        let runs = commit
            .package()
            .slide_notes_text_runs(SlideSelector::index(slide))?;
        let bold: Vec<_> = runs
            .iter()
            .filter(|run| run.style().bold == Some(true))
            .map(|run| run.text().to_owned())
            .collect();
        assert_eq!(bold, ["second"], "slide {slide}");
        assert_eq!(
            commit
                .package()
                .slide_notes(SlideSelector::index(slide))?
                .as_deref(),
            Some("Remember the\nsecond point")
        );
        let restored = commit
            .package()
            .apply_slide_drawables(&commit.patch().inverse())?;
        assert_eq!(bytes(restored.package())?, bytes(&written)?);
    }
    Ok(())
}

#[test]
fn notes_styles_refuse_typed() -> TestResult {
    let path =
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote/basic.key");
    let package = Package::open(path)?;
    let mut edit = package.edit_slide_drawables();
    // A style that sets nothing, a span past the notes.
    let notes = package.slide_notes(SlideSelector::index(0))?;
    if let Some(notes) = notes.filter(|notes| !notes.is_empty()) {
        assert!(matches!(
            edit.set_notes_text_style(SlideSelector::index(0), TextSpan::new(0, 1), Style::new()),
            Err(Error::Refused(
                litchi_keynote::slide::drawable::Refusal::Unsupported {
                    reason: Unsupported::TextStyling
                }
            ))
        ));
        let past = notes.chars().count() + 1;
        assert!(matches!(
            edit.set_notes_text_style(
                SlideSelector::index(0),
                TextSpan::new(0, past),
                Style::new().with_bold(true)
            ),
            Err(Error::Refused(
                litchi_keynote::slide::drawable::Refusal::InvalidSpan
            ))
        ));
    }
    Ok(())
}
