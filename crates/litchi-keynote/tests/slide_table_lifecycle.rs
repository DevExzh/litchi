//! Slide-table duplication and removal over the shared table structure set
//! (record 0372).

use std::path::PathBuf;

use litchi_core::edit::RefusalKind;
use litchi_keynote::slide::drawable::{Geometry, Kind};
use litchi_keynote::slide::table::lifecycle::{Error, Refusal};
use litchi_keynote::{OpenOptions, Package, SemanticLimits, SlideSelector, TableSelector};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn package() -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/table-discovery.key");
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut out = Vec::new();
    package.write_to(&mut out)?;
    Ok(out)
}

fn names(package: &Package) -> TestResult<Vec<String>> {
    let count = tables(package)?.len();
    (0..count)
        .map(|table| {
            Ok(package
                .slide_table_name(SlideSelector::index(0), TableSelector::index(table))?
                .as_str()
                .to_owned())
        })
        .collect()
}

/// The first slide's table frames, back to front.
fn tables(package: &Package) -> TestResult<Vec<Geometry>> {
    let mut drawables = package
        .slide_drawable_list(SlideSelector::index(0))?
        .into_iter()
        .filter(|drawable| drawable.kind() == Kind::Table)
        .collect::<Vec<_>>();
    drawables.sort_by_key(|drawable| drawable.z_position().map(|z| z.get()));
    drawables
        .iter()
        .map(|drawable| {
            drawable
                .geometry()
                .ok_or_else(|| "table without frame".into())
        })
        .collect()
}

fn cells(package: &Package, table: usize) -> TestResult<Vec<String>> {
    let read = package.slide_table_cells(SlideSelector::index(0), TableSelector::index(table))?;
    Ok(read
        .iter_cells()
        .map(|cell| format!("{:?} {:?}", cell.position(), cell.value()))
        .collect())
}

#[test]
fn duplicate_names_offsets_and_inverts_exactly() -> TestResult {
    let source = package()?;
    let original = bytes(&source)?;
    let commit = source
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0))?
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    let copied = commit.package();
    assert_eq!(names(copied)?, ["Table 1", "Table 1 copy"]);
    let frames = tables(copied)?;
    assert_eq!(
        (frames[1].x, frames[1].y, frames[1].width, frames[1].height),
        (
            frames[0].x + 10.0,
            frames[0].y + 10.0,
            frames[0].width,
            frames[0].height
        )
    );
    assert_eq!(cells(copied, 1)?, cells(copied, 0)?);
    let reopened = Package::from_bytes(&bytes(copied)?)?;
    assert_eq!(names(&reopened)?, ["Table 1", "Table 1 copy"]);

    let again = copied
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0))?
        .commit()?;
    assert_eq!(
        names(again.package())?,
        ["Table 1", "Table 1 copy", "Table 1 copy 2"]
    );

    let restored = copied.apply_slide_table_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, original);
    let replayed = source.apply_slide_table_lifecycle(commit.patch())?;
    assert_eq!(bytes(replayed.package())?, bytes(copied)?);
    assert_eq!(
        copied.apply_slide_table_lifecycle(commit.patch()).err(),
        Some(Error::PatchConflict)
    );
    Ok(())
}

#[test]
fn remove_detaches_the_table_and_inverts_exactly() -> TestResult {
    let source = package()?;
    let original = bytes(&source)?;
    let commit = source
        .edit_slide_table_lifecycle()
        .remove(SlideSelector::index(0), TableSelector::index(0))?
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().touched_components() > 1);
    assert!(tables(commit.package())?.is_empty());
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    assert!(tables(&reopened)?.is_empty());
    assert!(
        reopened
            .slide_table_cells(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    let restored = commit
        .package()
        .apply_slide_table_lifecycle(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, original);

    // A copy removed again leaves the source table as it was.
    let copied = source
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0))?
        .commit()?;
    let removed = copied
        .package()
        .edit_slide_table_lifecycle()
        .remove(SlideSelector::index(0), TableSelector::index(1))?
        .commit()?;
    assert_eq!(names(removed.package())?, ["Table 1"]);
    assert_eq!(cells(removed.package(), 0)?, cells(&source, 0)?);
    Ok(())
}

#[test]
fn refusals_are_typed() -> TestResult {
    let source = package()?;
    let empty = source.edit_slide_table_lifecycle().commit().err();
    assert_eq!(empty, Some(Error::Refused(Refusal::NoStagedOperation)));
    let twice = source
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0))?
        .remove(SlideSelector::index(0), TableSelector::index(0))
        .err();
    assert_eq!(twice, Some(Error::Refused(Refusal::OperationAlreadyStaged)));
    for (slide, table) in [(0, 1), (99, 0)] {
        let missing = source
            .edit_slide_table_lifecycle()
            .remove(SlideSelector::index(slide), TableSelector::index(table))
            .err();
        assert_eq!(missing, Some(Error::Refused(Refusal::TableNotFound)));
    }
    assert_eq!(Refusal::TableNotFound.kind(), RefusalKind::NotFound);
    assert_eq!(
        Refusal::UnsupportedDependency.kind(),
        RefusalKind::UnsupportedContent
    );
    assert_eq!(
        Refusal::NoStagedOperation.kind(),
        RefusalKind::UnsupportedOperation
    );
    Ok(())
}

#[test]
fn low_selection_budget_is_typed_and_atomic() -> TestResult {
    let source = package()?;
    let source_bytes = bytes(&source)?;
    let semantic = SemanticLimits::builder().max_references(1)?.build()?;
    let restricted = Package::from_bytes_with_options(
        &source_bytes,
        OpenOptions::new().semantic_limits(semantic),
    )?;
    let before = bytes(&restricted)?;
    let result = restricted
        .edit_slide_table_lifecycle()
        .duplicate(SlideSelector::index(0), TableSelector::index(0));
    assert!(matches!(
        result,
        Err(Error::Limit {
            kind: litchi_keynote::slide::table::lifecycle::LimitKind::Selection,
            ..
        })
    ));
    assert_eq!(bytes(&restricted)?, before);
    Ok(())
}
