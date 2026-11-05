//! Slide-table creation (record 0329): a new, empty table cloned from the
//! structure of the deck's first slide table, read back through the slide
//! table reader, reopened, inverted exactly; decks without a table refuse,
//! typed.

use std::path::{Path, PathBuf};

use litchi_keynote::slide::drawable::{
    Error, Geometry, Kind, NewTable, TableCellValue, Unsupported,
};
use litchi_keynote::slide::table::cells::{At, Value};
use litchi_keynote::{Package, SlideSelector, TableSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn keynote_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/keynote")
}

fn corpus() -> Vec<PathBuf> {
    let mut files = Vec::new();
    let mut stack = vec![keynote_dir()];
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

const FRAME: Geometry = Geometry::new(120.0, 160.0, 600.0, 240.0);

/// The z-order ordinal of the slide's last table.
fn last_table(package: &Package, slide: usize) -> TestResult<usize> {
    let tables = package
        .slide_drawable_list(SlideSelector::index(slide))?
        .iter()
        .filter(|drawable| drawable.kind() == Kind::Table)
        .count();
    Ok(tables - 1)
}

#[test]
fn tables_are_created_read_back_reopened_and_inverted() -> TestResult {
    let package = Package::open(keynote_dir().join("slide-table-read-native.key"))?;
    let source = bytes(&package)?;
    let before = package.slide_drawable_list(SlideSelector::index(0))?.len();
    let mut edit = package.edit_slide_drawables();
    edit.add_table(
        SlideSelector::index(0),
        NewTable::new(4, 3, FRAME)
            .with_name("Forecast")
            .with_cell(0, 0, TableCellValue::Text("Region".into()))
            .with_cell(1, 2, TableCellValue::Number(12.5))
            .with_cell(3, 1, TableCellValue::Boolean(true)),
    )?;
    let commit = edit.commit()?;
    assert_eq!(commit.created().len(), 1);
    let created = commit.created()[0];
    let edited = commit.package();
    let list = edited.slide_drawable_list(SlideSelector::index(0))?;
    assert_eq!(list.len(), before + 1);
    let new = &list[created.drawable().as_index()];
    assert_eq!(new.kind(), Kind::Table);
    let frame = new.geometry().expect("frame");
    assert_eq!(
        (frame.x, frame.y, frame.width, frame.height),
        (120.0, 160.0, 600.0, 240.0)
    );
    let reopened = Package::from_bytes(&bytes(edited)?)?;
    let read = reopened.slide_table_cells(
        SlideSelector::index(0),
        TableSelector::index(last_table(&reopened, 0)?),
    )?;
    let table = read.table();
    assert_eq!(table.name(), "Forecast");
    assert_eq!(
        (table.dimensions().rows(), table.dimensions().columns()),
        (4, 3)
    );
    assert_eq!(
        table.get(At::new(0, 0)),
        Some(&Value::Text("Region".into()))
    );
    assert!(matches!(
        table.get(At::new(1, 2)),
        Some(Value::Number(number)) if number.get() == 12.5
    ));
    assert_eq!(table.get(At::new(3, 1)), Some(&Value::Boolean(true)));
    // The template table reads as before.
    let original = package.slide_table_cells(SlideSelector::index(0), TableSelector::index(0))?;
    let again = reopened.slide_table_cells(SlideSelector::index(0), TableSelector::index(0))?;
    assert_eq!(original.table(), again.table());
    // The inverse restores the source exactly.
    let restored = edited.apply_slide_drawables(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn every_corpus_deck_creates_a_table_or_refuses_typed() -> TestResult {
    let (mut created, mut refused) = (0_usize, 0_usize);
    for path in corpus() {
        let Ok(package) = Package::open(&path) else {
            continue;
        };
        let Ok(slides) = package.slides().map(<[_]>::len) else {
            continue;
        };
        for slide in [0, slides.saturating_sub(1)] {
            if slide >= slides {
                continue;
            }
            let mut edit = package.edit_slide_drawables();
            let staged = edit.add_table(
                SlideSelector::index(slide),
                NewTable::new(3, 2, FRAME).with_cell(1, 1, TableCellValue::Text("x".into())),
            );
            match staged.map(|_| ()) {
                Ok(()) => {},
                Err(Error::Refused(litchi_keynote::slide::drawable::Refusal::Unsupported {
                    reason: Unsupported::TableTemplate,
                })) => {
                    refused += 1;
                    continue;
                },
                Err(error) => panic!("{}: staging failed: {error:?}", path.display()),
            }
            let commit = edit
                .commit()
                .unwrap_or_else(|error| panic!("{}: {error:?}", path.display()));
            let reopened = Package::from_bytes(&bytes(commit.package())?)?;
            let read = reopened.slide_table_cells(
                SlideSelector::index(slide),
                TableSelector::index(last_table(&reopened, slide)?),
            )?;
            assert_eq!(
                read.table().get(At::new(1, 1)),
                Some(&Value::Text("x".into())),
                "{}",
                path.display()
            );
            let restored = commit
                .package()
                .apply_slide_drawables(&commit.patch().inverse())?;
            assert_eq!(bytes(restored.package())?, bytes(&package)?);
            created += 1;
        }
    }
    eprintln!("created {created}, refused {refused}");
    assert!(created >= 3);
    Ok(())
}
