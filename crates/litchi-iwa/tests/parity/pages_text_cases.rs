//! Pages body text parity cases (wave W1-B, record 0371): the host's
//! whole-body text verbs against the leaf section-text family
//! (`litchi_pages::section::text`).

use litchi_iwa::pages;
use litchi_pages::{Package, SectionSelector, TextPosition, TextSpan};

use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const BASIC: &str = "iwork/pages/basic.pages";
const SECTIONS: &str = "iwork/pages/body-sections-unicode.pages";

fn host_body_text(bytes: &[u8]) -> TestResult<String> {
    Ok(pages::Draft::from_bytes(bytes)?.body_text()?)
}

fn leaf_sections(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    (0..package.sections().len())
        .map(|index| {
            Ok(package
                .section_text(SectionSelector::index(index))?
                .to_owned())
        })
        .collect()
}

fn leaf_commit(commit: &litchi_pages::section::text::Commit) -> TestResult<LeafRun> {
    let restored = commit
        .package()
        .apply_section_text(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[test]
fn pages_set_body_text_parity() -> TestResult {
    let input = fixture(BASIC)?;
    Case::new(
        "pages_set_body_text_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.set_body_text("Replaced body — ünïcode")?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_text()?;
        edit.set("Replaced body — ünïcode")?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

#[test]
fn pages_clear_body_parity() -> TestResult {
    let input = fixture(BASIC)?;
    Case::new(
        "pages_clear_body_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.clear_body()?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_body_text()?;
        edit.clear()?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

/// `replace_body_text` takes a body-absolute UTF-16 range that may not
/// consume a section break, so it always falls inside one section; the leaf
/// edits that section with the section-relative span.
#[test]
fn pages_replace_body_text_parity() -> TestResult {
    let input = fixture(SECTIONS)?;
    let body = host_body_text(&input.bytes)?;
    let units = body.encode_utf16().collect::<Vec<_>>();
    // The second section starts after the first section-break marker.
    let marker = units
        .iter()
        .position(|unit| *unit == 0x0004)
        .ok_or("fixture has one section")?;
    let section_start = marker + 1;
    let (start, end) = (section_start + 1, section_start + 3);
    Case::new(
        "pages_replace_body_text_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = pages::Draft::from_bytes(source)?;
        editor.replace_body_text(start..end, "Ω inserted")?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        let package = Package::from_bytes(source)?;
        let mut edit = package.edit_section_text(SectionSelector::index(1))?;
        edit.replace(
            TextSpan::new(
                TextPosition::from_utf16_index(start - section_start)?,
                TextPosition::from_utf16_index(end - section_start)?,
            )?,
            "Ω inserted",
        )?;
        leaf_commit(&edit.commit()?)
    })
    .read("host body_text", host_body_text)
    .read("leaf section_text", leaf_sections)
    .run()?;
    Ok(())
}

#[test]
#[ignore = "callgrind cost probe (record 0371)"]
fn text_set_body_text_cost_abba() -> TestResult {
    crate::pages_cost_probes::abba(
        &fixture(BASIC)?,
        |source| {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_text("Replaced body — ünïcode")?;
            Ok(editor.to_bytes()?)
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let mut edit = package.edit_body_text()?;
            edit.set("Replaced body — ünïcode")?;
            Ok(leaf_bytes!(edit.commit()?.package()))
        },
    )
}
