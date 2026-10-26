//! Sheet-drawable families on native fixtures (record 0378): the drawable
//! list, frame/property/playback edits and z-order moves, their exact
//! inverses and their refusals.

use std::path::PathBuf;
use std::time::Duration;

use litchi_core::Position;
use litchi_core::edit::RefusalKind;
use litchi_numbers::Package;
use litchi_numbers::sheet::drawable::{
    self, DrawableSelector, FlipAxis, Kind, MediaLoopMode, MediaPlaybackSettings, Point,
    Properties, properties,
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn open(name: &str) -> TestResult<Package> {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers")
        .join(name);
    Ok(Package::open(path)?)
}

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

#[test]
fn the_list_reports_kinds_in_back_to_front_order() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let list = package.sheet_drawable_list(0)?;
    assert!(list.iter().any(|item| item.kind() == Kind::Table));
    let image = list
        .iter()
        .find(|item| item.kind() == Kind::Image)
        .ok_or("no image")?;
    assert_eq!(image.kind_selector(), DrawableSelector::image(0));
    assert!(image.geometry().is_some());
    for (position, item) in list.iter().enumerate() {
        assert_eq!(item.z_position(), Position::new(position));
        assert_eq!(item.selector(), DrawableSelector::index(position));
    }
    let audio = open("audio-playback-native.numbers")?.sheet_drawable_list(0)?;
    assert!(audio.iter().any(|item| item.kind() == Kind::Audio));
    Ok(())
}

#[test]
fn a_frame_edit_round_trips_and_its_inverse_restores_the_source() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let before = package.sheet_drawable_properties(0, DrawableSelector::image(0))?;
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::image(0))?;
    edit.set_position(Point { x: 12.0, y: 34.0 })?
        .flip(FlipAxis::Horizontal)?
        .set_properties(Properties {
            locked: Some(true),
            ..before.properties().clone()
        });
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    let after = commit
        .package()
        .sheet_drawable_properties(0, DrawableSelector::image(0))?;
    let frame = after.geometry().ok_or("no frame")?;
    assert_eq!(frame.position, Some(Point { x: 12.0, y: 34.0 }));
    assert!(frame.flags.ok_or("no flags")?.is_reflected());
    assert_eq!(after.properties().locked, Some(true));
    let restored = commit
        .package()
        .apply_sheet_drawable_properties(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);
    assert!(matches!(
        package.apply_sheet_drawable_properties(&commit.patch().inverse()),
        Err(properties::Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn an_unchanged_edit_is_an_exact_no_op() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let commit = package
        .edit_sheet_drawable_properties(0, DrawableSelector::image(0))?
        .commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package())?, bytes(&package)?);
    Ok(())
}

#[test]
fn audio_playback_is_edited_through_the_same_family() -> TestResult {
    let package = open("audio-playback-native.numbers")?;
    let settings = MediaPlaybackSettings {
        start_time: Some(Duration::from_millis(250)),
        end_time: Duration::from_millis(1_000),
        poster_time: None,
        loop_mode: Some(MediaLoopMode::Repeat),
        volume: None,
    };
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    edit.set_playback(settings)?;
    let commit = edit.commit()?;
    let after = commit
        .package()
        .sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    assert_eq!(after.playback(), Some(settings.canonicalize()?));
    Ok(())
}

#[test]
fn verbs_that_do_not_apply_are_typed_refusals() -> TestResult {
    let package = open("audio-playback-native.numbers")?;
    let mut edit = package.edit_sheet_drawable_properties(0, DrawableSelector::audio(0))?;
    let error = edit.flip(FlipAxis::Vertical).map(drop).unwrap_err();
    let properties::Error::Refused(refusal) = error else {
        return Err(format!("unexpected {error:?}").into());
    };
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedOperation);
    assert!(matches!(
        package.sheet_drawable_properties(0, DrawableSelector::image(5)),
        Err(properties::Error::Refused(
            properties::Refusal::DrawableNotFound { .. }
        ))
    ));
    assert!(matches!(
        package.sheet_drawable_properties("No such sheet", 0),
        Err(properties::Error::Refused(
            properties::Refusal::SheetNotFound
        ))
    ));
    Ok(())
}

#[test]
fn z_order_moves_reorder_the_list_and_tables_are_not_removed() -> TestResult {
    let package = open("image-adjustments-native.numbers")?;
    let list = package.sheet_drawable_list(0)?;
    let last = list.len() - 1;
    let mut edit = package.edit_sheet_drawables(0)?;
    edit.move_to(DrawableSelector::index(last), Position::new(0))?;
    let commit = edit.commit()?;
    let moved = commit.package().sheet_drawable_list(0)?;
    assert_eq!(moved[0].kind(), list[last].kind());
    let restored = commit
        .package()
        .apply_sheet_drawables(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, bytes(&package)?);

    let mut edit = package.edit_sheet_drawables(0)?;
    let table = list
        .iter()
        .find(|item| item.kind() == Kind::Table)
        .ok_or("no table")?;
    assert!(matches!(
        edit.remove(table.selector()).map(drop),
        Err(drawable::Error::Refused(
            drawable::Refusal::UnsupportedOperation { .. }
        ))
    ));
    assert!(matches!(
        edit.move_to(DrawableSelector::index(0), Position::new(list.len()))
            .map(drop),
        Err(drawable::Error::Refused(
            drawable::Refusal::ZPositionOutOfRange { .. }
        ))
    ));
    assert!(matches!(
        package.edit_sheet_drawables(0)?.commit(),
        Err(drawable::Error::Refused(
            drawable::Refusal::NoStagedOperation
        ))
    ));
    Ok(())
}
