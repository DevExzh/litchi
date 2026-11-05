//! Character styling of body text spans (record 0309).

use std::path::PathBuf;

use litchi_pages::section::text_style::{CharacterProperties, Error};
use litchi_pages::{Package, TextSpan};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages")
        .join(name)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn span(start: usize, end: usize) -> TestResult<TextSpan> {
    Ok(TextSpan::from_utf16_indexes(start, end)?)
}

#[test]
fn body_spans_take_character_styles_and_invert_exactly() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let source = bytes(&package)?;
    let text = package.sections()[0]
        .body_text()
        .expect("body text")
        .to_owned();
    let length = text.encode_utf16().count();
    assert!(length >= 12, "fixture text too short");
    let before = package.section_text_runs(0)?;

    let mut bold = CharacterProperties::new();
    bold.bold = Some(true);
    bold.color = Some([10, 120, 200]);
    let mut italic = CharacterProperties::new();
    italic.italic = Some(true);
    italic.size = Some(20.0);
    italic.font_name = Some("Helvetica-Bold".to_owned());
    let mut edit = package.edit_section_text_style(0)?;
    edit.set(span(2, 8)?, bold.clone())?;
    edit.set(span(5, 11)?, italic.clone())?;
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let styled = commit.package();
    assert_eq!(styled.sections()[0].body_text(), Some(text.as_str()));
    let reopened = Package::from_bytes(&bytes(styled)?)?;
    for candidate in [styled, &reopened] {
        let runs = candidate.section_text_runs(0)?;
        let at = |index: u32| {
            runs.iter()
                .find(|run| {
                    run.span().start().utf16_index() <= index
                        && run.span().end().utf16_index() > index
                })
                .expect("run")
        };
        assert_eq!(at(3).style().bold, Some(true));
        assert_eq!(at(3).style().color, Some([10, 120, 200]));
        // The overlap is a variation of the bold run: bold and italic.
        assert_eq!(at(6).style().bold, Some(true));
        assert_eq!(at(6).style().italic, Some(true));
        assert_eq!(at(6).style().size, Some(20.0));
        assert_eq!(at(9).style().font_name.as_deref(), Some("Helvetica-Bold"));
        assert_eq!(*at(0).style(), CharacterProperties::new());
        assert_eq!(*at(11).style(), *before[0].style());
    }
    let restored = styled.apply_section_text_style(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    assert!(matches!(
        styled.apply_section_text_style(commit.patch()),
        Err(Error::PatchConflict)
    ));

    // A later text edit keeps the styled characters styled.
    let mut edit = styled.edit_section_text(0)?;
    edit.insert(litchi_pages::TextPosition::from_utf16_index(0)?, "XY")?;
    let retexted = edit.commit()?.into_package();
    let runs = retexted.section_text_runs(0)?;
    let styled_run = runs
        .iter()
        .find(|run| run.style().bold == Some(true))
        .expect("bold run");
    assert_eq!(styled_run.span().start().utf16_index(), 4);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let mut bold = CharacterProperties::new();
    bold.bold = Some(true);
    let mut edit = package.edit_section_text_style(0)?;
    assert!(matches!(
        edit.set(span(0, 0)?, bold.clone()),
        Err(Error::Refused(
            litchi_pages::section::text_style::Refusal::InvalidSpan
        ))
    ));
    assert!(matches!(
        edit.set(span(0, 1_000_000)?, bold.clone()),
        Err(Error::Refused(
            litchi_pages::section::text_style::Refusal::InvalidSpan
        ))
    ));
    assert!(matches!(
        edit.set(span(0, 1)?, CharacterProperties::new()),
        Err(Error::Refused(
            litchi_pages::section::text_style::Refusal::EmptyStyle
        ))
    ));
    assert!(matches!(
        package.edit_section_text_style(40),
        Err(Error::Refused(
            litchi_pages::section::text_style::Refusal::SectionNotFound
        ))
    ));
    assert!(matches!(
        package.edit_section_text_style(0)?.commit(),
        Err(Error::Refused(
            litchi_pages::section::text_style::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}

/// Every corpus document with body text takes a bold span, or refuses
/// typed; the inverse restores each source byte for byte.
#[test]
fn corpus_bodies_take_bold_or_refuse_typed() -> TestResult {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/pages");
    let mut paths: Vec<PathBuf> = std::fs::read_dir(root)?
        .map(|entry| entry.map(|entry| entry.path()))
        .collect::<Result<_, _>>()?;
    paths.retain(|path| {
        path.extension()
            .is_some_and(|extension| extension == "pages")
    });
    paths.sort();
    let mut styled = 0_usize;
    let mut refused = Vec::new();
    for path in paths {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Some(length) = package
            .sections()
            .first()
            .and_then(|section| section.body_text())
            .map(|text| text.encode_utf16().count())
            .filter(|length| *length >= 2)
        else {
            continue;
        };
        let source = bytes(&package)?;
        let mut bold = CharacterProperties::new();
        bold.bold = Some(true);
        let mut edit = package.edit_section_text_style(0)?;
        edit.set(span(0, length.min(4))?, bold)?;
        match edit.commit() {
            Ok(commit) => {
                let runs = commit.package().section_text_runs(0)?;
                assert_eq!(runs[0].style().bold, Some(true), "{}", path.display());
                let restored = commit
                    .package()
                    .apply_section_text_style(&commit.patch().inverse())?;
                assert_eq!(bytes(restored.package())?, source, "{}", path.display());
                styled += 1;
            },
            Err(
                error @ (Error::Refused(
                    litchi_pages::section::text_style::Refusal::UnsupportedDependency,
                )
                | Error::Refused(
                    litchi_pages::section::text_style::Refusal::UnsupportedSource,
                )),
            ) => {
                refused.push((path, error));
            },
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("styled {styled}; refused {refused:?}");
    assert!(styled >= 10, "only {styled} styled");
    Ok(())
}
