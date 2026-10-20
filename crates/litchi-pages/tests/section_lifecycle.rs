//! Adding and removing sections (request 11, record 0312).

use std::path::PathBuf;

use litchi_pages::Package;
use litchi_pages::section::lifecycle::Error;

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

fn texts(package: &Package) -> Vec<String> {
    package
        .sections()
        .iter()
        .map(|section| section.body_text().unwrap_or_default().to_owned())
        .collect()
}

/// The header and footer texts of section `position`, by template, kind
/// and slot.
fn regions(package: &Package, position: usize) -> TestResult<Vec<(String, String)>> {
    let mut regions: Vec<(String, String)> = package
        .header_footers()?
        .iter()
        .filter(|region| region.section_position().get() == position)
        .map(|region| {
            (
                format!(
                    "{:?} {:?} {}",
                    region.template(),
                    region.kind(),
                    region.slot_index()
                ),
                region.text().to_owned(),
            )
        })
        .collect();
    regions.sort();
    Ok(regions)
}

#[test]
fn a_section_is_added_after_another_with_its_page_masters() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let source = bytes(&package)?;
    let before = texts(&package);
    let commit = package
        .edit_section_lifecycle()
        .add(0, Some("Appendix"))?
        .commit()?;
    assert!(commit.diagnostics().changed());
    let added = commit.package();
    assert_eq!(added.sections().len(), 2);
    assert_eq!(added.sections()[1].name(), Some("Appendix"));
    assert_eq!(texts(added), [before[0].clone(), String::new()]);
    assert_eq!(regions(added, 1)?, regions(&package, 0)?);
    // The candidate reopens from its bytes.
    let reopened = Package::from_bytes(&bytes(added)?)?;
    assert_eq!(texts(&reopened), texts(added));
    // The inverse restores the source exactly.
    let restored = added.apply_section_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // The new section is removed again.
    let removed = added.edit_section_lifecycle().remove(1)?.commit()?;
    assert_eq!(texts(removed.package()), before);
    Ok(())
}

#[test]
fn a_section_is_removed_and_its_text_joins_the_previous_one() -> TestResult {
    let package = Package::open(fixture("body-sections-unicode.pages"))?;
    let source = bytes(&package)?;
    let before = texts(&package);
    assert_eq!(before.len(), 2);
    let commit = package.edit_section_lifecycle().remove(1)?.commit()?;
    assert_eq!(
        texts(commit.package()),
        [format!("{}{}", before[0], before[1])]
    );
    assert!(commit.diagnostics().objects() > 1);
    let restored = commit
        .package()
        .apply_section_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // Removing then re-adding keeps every other section intact.
    let added = package.edit_section_lifecycle().add(1, None)?.commit()?;
    assert_eq!(
        texts(added.package()),
        [before[0].clone(), before[1].clone(), String::new()]
    );
    assert_eq!(added.package().sections()[2].name(), None);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    assert_eq!(
        package.edit_section_lifecycle().remove(0).err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::InitialSection
        ))
    );
    assert_eq!(
        package.edit_section_lifecycle().remove(5).err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::SectionNotFound
        ))
    );
    assert_eq!(
        package.edit_section_lifecycle().add(0, Some("a\0b")).err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::InvalidName
        ))
    );
    assert_eq!(
        package.edit_section_lifecycle().commit().err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::NoStagedOperation
        ))
    );
    let added = package.edit_section_lifecycle().add(0, None)?.commit()?;
    assert_eq!(
        added
            .package()
            .edit_section_lifecycle()
            .add(1, None)?
            .commit()
            .err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::EmptySection
        ))
    );
    let staged = package.edit_section_lifecycle().add(0, None)?;
    assert_eq!(
        staged.remove(0).err(),
        Some(Error::Refused(
            litchi_pages::section::lifecycle::Refusal::OperationAlreadyStaged
        ))
    );
    Ok(())
}

#[test]
fn every_corpus_document_adds_and_removes_a_section_or_refuses_typed() -> TestResult {
    let mut added = 0;
    let mut refused = Vec::new();
    for entry in std::fs::read_dir(fixture(""))? {
        let path = entry?.path();
        if path.extension().and_then(|value| value.to_str()) != Some("pages") || !path.is_file() {
            continue;
        }
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let last = package.sections().len() - 1;
        let before = texts(&package);
        match package
            .edit_section_lifecycle()
            .add(last, Some("Added"))?
            .commit()
        {
            Ok(commit) => {
                let mut expected = before.clone();
                expected.push(String::new());
                assert_eq!(texts(commit.package()), expected, "{}", path.display());
                let removed = commit
                    .package()
                    .edit_section_lifecycle()
                    .remove(last + 1)?
                    .commit()?;
                assert_eq!(texts(removed.package()), before, "{}", path.display());
                added += 1;
            },
            Err(
                error @ (Error::Refused(litchi_pages::section::lifecycle::Refusal::EmptySection)
                | Error::Refused(
                    litchi_pages::section::lifecycle::Refusal::UnsupportedContent,
                )
                | Error::Refused(
                    litchi_pages::section::lifecycle::Refusal::UnsupportedSource,
                )
                | Error::Refused(
                    litchi_pages::section::lifecycle::Refusal::UnsupportedDependency,
                )),
            ) => refused.push((path.file_name().map(|name| name.to_owned()), error)),
            Err(error) => panic!("{}: {error:?}", path.display()),
        }
    }
    eprintln!("added {added}, refused {refused:?}");
    assert!(added >= 20, "{added} documents took a section");
    Ok(())
}
