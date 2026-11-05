//! Paragraph-level section transactions (record 0299): set, insert, remove,
//! relocate and swap through `Package::edit_section_paragraphs`; text sets
//! coalesced into one section-text transaction where that publishes what
//! they publish one at a time (record 0314), and exact inverses on every
//! corpus file.

use std::path::PathBuf;

use litchi_pages::section::paragraphs::Error;
use litchi_pages::{Package, Position};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages")
        .join(name)
}

fn corpus() -> TestResult<Vec<(String, Vec<u8>)>> {
    let mut files = Vec::new();
    for entry in std::fs::read_dir(fixture(""))? {
        let path = entry?.path();
        if path
            .extension()
            .is_some_and(|extension| extension == "pages")
        {
            let name = path
                .file_name()
                .map(|name| name.to_string_lossy().into_owned())
                .unwrap_or_default();
            files.push((name, std::fs::read(&path)?));
        }
    }
    files.sort();
    Ok(files)
}

fn at(index: usize) -> Position {
    Position::new(index)
}

#[test]
fn paragraph_operations_round_trip_on_a_plain_section() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let original = package.to_bytes()?;
    assert_eq!(
        package.section_paragraphs(0_usize)?,
        [
            "Litchi native Pages fixture",
            "Buffa lazy-view migration verification",
            "2026-08-07"
        ]
    );

    let mut edit = package.edit_section_paragraphs(0_usize)?;
    edit.set(at(0), "Retitled")?
        .insert(at(1), "Inserted")?
        .remove(at(3))?
        .relocate(at(2), at(0))?
        .swap(at(1), at(2))?;
    let commit = edit.commit()?;
    // Record 0314: structural operations publish one transaction each, as
    // committing them one at a time does.
    assert_eq!(
        commit.diagnostics().transactions(),
        5,
        "one splice per operation"
    );
    let edited = commit.package();
    assert_eq!(
        edited.section_paragraphs(0_usize)?,
        [
            "Buffa lazy-view migration verification",
            "Inserted",
            "Retitled"
        ]
    );
    // The commit reopens: the published bytes decode to the same paragraphs.
    let reopened = Package::from_bytes(&edited.to_bytes()?)?;
    assert_eq!(
        reopened.section_paragraphs(0_usize)?,
        edited.section_paragraphs(0_usize)?
    );

    let restored = edited.apply_section_paragraphs(&commit.patch().inverse())?;
    assert_eq!(restored.package().to_bytes()?, original);
    Ok(())
}

#[test]
fn an_exact_no_op_publishes_the_source() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let mut edit = package.edit_section_paragraphs(0_usize)?;
    edit.set(at(0), "Litchi native Pages fixture")?;
    let commit = edit.commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes()?, package.to_bytes()?);
    Ok(())
}

#[test]
fn refusals_are_typed_and_leave_the_source_untouched() -> TestResult {
    let package = Package::open(fixture("basic.pages"))?;
    let mut edit = package.edit_section_paragraphs(0_usize)?;
    assert!(matches!(
        edit.set(at(9), "x"),
        Err(Error::Refused(
            litchi_pages::section::paragraphs::Refusal::PositionNotFound { count: 3, .. }
        ))
    ));
    assert!(matches!(
        edit.set(at(0), "a\nb"),
        Err(Error::Refused(
            litchi_pages::section::paragraphs::Refusal::ParagraphBreak
        ))
    ));
    assert!(edit.insert(at(0), "\u{fffc}").is_err());

    let single = Package::open(fixture("chart-grid-authored-native.pages"))?;
    let mut edit = single.edit_section_paragraphs(0_usize)?;
    assert!(matches!(
        edit.remove(at(0)),
        Err(Error::Refused(
            litchi_pages::section::paragraphs::Refusal::LastParagraph
        ))
    ));

    // The table's anchor sits in the last paragraph: removing or moving it
    // would delete or move the table.
    let table = Package::open(fixture("body-table-read-native.pages"))?;
    let count = table.section_paragraphs(0_usize)?.len();
    let mut edit = table.edit_section_paragraphs(0_usize)?;
    edit.relocate(at(count - 1), at(0))?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(
            litchi_pages::section::paragraphs::Refusal::AnchoredParagraph
        ))
    ));
    let mut edit = table.edit_section_paragraphs(0_usize)?;
    edit.remove(at(count - 1))?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(
            litchi_pages::section::paragraphs::Refusal::AnchoredParagraph
        ))
    ));
    Ok(())
}

#[test]
fn setting_an_anchored_paragraph_keeps_its_object() -> TestResult {
    let package = Package::open(fixture("chart-grid-authored-native.pages"))?;
    let before = package.section_paragraphs(0_usize)?[0].to_owned();
    assert!(before.ends_with('\u{fffc}'));
    let mut edit = package.edit_section_paragraphs(0_usize)?;
    edit.set(at(0), "Renamed grid")?;
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().section_paragraphs(0_usize)?,
        ["Renamed grid\u{fffc}"]
    );
    Ok(())
}

#[test]
fn sets_on_both_sides_of_an_anchor_fall_back_to_one_splice_each() -> TestResult {
    let package = Package::open(fixture("body-table-read-native.pages"))?;
    let paragraphs = package.section_paragraphs(0_usize)?;
    let anchored = paragraphs
        .iter()
        .position(|paragraph| paragraph.contains('\u{fffc}'))
        .ok_or("the fixture anchors a table")?;
    assert!(anchored > 0);
    let mut edit = package.edit_section_paragraphs(0_usize)?;
    edit.set(at(0), "First")?
        .set(at(anchored), "Table caption")?;
    let commit = edit.commit()?;
    let after = commit.package().section_paragraphs(0_usize)?;
    assert_eq!(after[0], "First");
    assert!(after[anchored].contains("Table caption"));
    assert!(after[anchored].contains('\u{fffc}'));
    Ok(())
}

#[test]
fn every_corpus_section_accepts_paragraph_set_and_insert_and_reverts_exactly() -> TestResult {
    for (name, bytes) in corpus()? {
        let package = Package::from_bytes(&bytes)?;
        for section in 0..package.sections().len() {
            let count = package.section_paragraphs(section)?.len();
            let mut edit = package.edit_section_paragraphs(section)?;
            edit.set(at(0), "Unified paragraph")?
                .insert(at(count), "Appended paragraph")?;
            let commit = edit
                .commit()
                .map_err(|error| format!("{name} section {section}: {error}"))?;
            let paragraphs = commit.package().section_paragraphs(section)?;
            assert_eq!(paragraphs.len(), count + 1, "{name}");
            assert!(paragraphs[0].starts_with("Unified paragraph"), "{name}");
            let restored = commit
                .package()
                .apply_section_paragraphs(&commit.patch().inverse())?;
            assert_eq!(restored.package().to_bytes()?, bytes, "{name}");
        }
    }
    Ok(())
}

/// Record 0314: a commit of several staged operations publishes exactly
/// what committing each on its own, in order, publishes — on every corpus
/// section, for sets of every paragraph, repeated sets, a set that cancels
/// an earlier one, and a mix with an insertion. (One splice across the
/// paragraphs used to drop the paragraph-style entries between them.)
#[test]
fn several_operations_publish_what_they_publish_one_at_a_time() -> TestResult {
    let mut compared = 0;
    for (name, bytes) in corpus()? {
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let Ok(paragraphs) = package.section_paragraphs(0_usize) else {
            continue;
        };
        let count = paragraphs.len();
        let last = count - 1;
        let original = paragraphs[0].to_string();
        let sequences: Vec<Vec<(usize, Option<String>)>> = vec![
            (0..count)
                .map(|line| (line, Some(format!("Set {line}"))))
                .collect(),
            vec![
                (last, Some("B".into())),
                (0, Some("A".into())),
                (last, Some("C".into())),
            ],
            vec![(0, Some(String::new())), (0, Some("Again".into()))],
            vec![(0, Some("Changed".into())), (0, Some(original.clone()))],
            vec![
                (0, Some("First".into())),
                (count, None),
                (last, Some("Last".into())),
            ],
        ];
        for sequence in sequences {
            let stage = |edit: &mut litchi_pages::section::paragraphs::Edit<'_>,
                         (line, text): &(usize, Option<String>)|
             -> Result<(), Error> {
                match text {
                    Some(text) => edit.set(at(*line), text).map(|_| ()),
                    None => edit.insert(at(*line), "Inserted").map(|_| ()),
                }
            };
            let mut edit = package.edit_section_paragraphs(0_usize)?;
            let staged = sequence.iter().try_for_each(|item| stage(&mut edit, item));
            let together = staged
                .and_then(|()| edit.commit())
                .map(|commit| commit.into_package());
            let mut one_by_one = Ok(package.clone());
            for item in &sequence {
                one_by_one = one_by_one.and_then(|current| {
                    let mut edit = current.edit_section_paragraphs(0_usize)?;
                    stage(&mut edit, item)?;
                    Ok(edit.commit()?.into_package())
                });
            }
            match (together, one_by_one) {
                (Ok(together), Ok(one_by_one)) => {
                    assert!(
                        together.to_bytes()? == one_by_one.to_bytes()?,
                        "{name}: {sequence:?}"
                    );
                    compared += 1;
                },
                (Err(left), Err(right)) => assert_eq!(left, right, "{name}: {sequence:?}"),
                (left, right) => panic!("{name}: {sequence:?}: {:?} {:?}", left.err(), right.err()),
            }
        }
    }
    assert!(compared >= 60, "{compared}");
    Ok(())
}
