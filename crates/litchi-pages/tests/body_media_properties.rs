//! Body-media frame, property and playback transactions
//! (`litchi_pages::media::properties`, record 0371).

use std::time::Duration;

use litchi_pages::drawable::{FlipAxis, Geometry, Point, Properties, Size};
use litchi_pages::media::properties::{Error, Refusal};
use litchi_pages::media::{MediaLoopMode, MediaPlaybackSettings, MediaVolume};
use litchi_pages::{MediaSelector, Package};

const IMAGES: &[u8] =
    include_bytes!("../../../test-data/iwork/pages/image-adjustments-native.pages");
const AUDIO: &[u8] = include_bytes!("../../../test-data/iwork/pages/audio-playback-native.pages");

fn bytes(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn image_frame_edit_round_trips_and_inverts_exactly() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let before = package
        .body_media_properties(MediaSelector::image(0))
        .unwrap();
    let frame = before.geometry().unwrap();
    assert!(before.playback().is_none());

    let moved = Geometry {
        position: Some(Point { x: 10.0, y: 20.0 }),
        size: Some(Size {
            width: 30.0,
            height: 40.0,
        }),
        ..frame
    };
    let mut edit = package
        .edit_body_media_properties(MediaSelector::image(0))
        .unwrap();
    edit.set_geometry(moved).unwrap();
    edit.flip(FlipAxis::Horizontal).unwrap();
    edit.set_properties(Properties {
        accessibility_description: Some("A chart".to_owned()),
        ..before.properties().clone()
    });
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(!commit.patch().is_noop());

    let after = commit
        .package()
        .body_media_properties(MediaSelector::image(0))
        .unwrap();
    let geometry = after.geometry().unwrap();
    assert_eq!(geometry.position, moved.position);
    assert_eq!(geometry.size, moved.size);
    assert_ne!(geometry.flags, frame.flags);
    assert_eq!(
        after.properties().accessibility_description.as_deref(),
        Some("A chart")
    );
    assert_eq!(after.original_size(), before.original_size());

    let restored = commit
        .package()
        .apply_body_media_properties(&commit.patch().inverse())
        .unwrap();
    assert_eq!(bytes(restored.package()), IMAGES);

    // Replaying the forward patch on the source reproduces the commit.
    let replayed = package.apply_body_media_properties(commit.patch()).unwrap();
    assert_eq!(bytes(replayed.package()), bytes(commit.package()));
    // A patch for another artifact conflicts.
    assert_eq!(
        commit
            .package()
            .apply_body_media_properties(commit.patch())
            .unwrap_err(),
        Error::PatchConflict
    );
}

#[test]
fn unchanged_edit_is_an_exact_no_op() {
    let package = Package::from_bytes(IMAGES).unwrap();
    let before = package
        .body_media_properties(MediaSelector::image(0))
        .unwrap();
    let mut edit = package
        .edit_body_media_properties(MediaSelector::image(0))
        .unwrap();
    edit.set_geometry(before.geometry().unwrap()).unwrap();
    edit.set_properties(before.properties().clone());
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(bytes(commit.package()), IMAGES);
}

#[test]
fn audio_playback_edit_changes_only_playback() {
    let package = Package::from_bytes(AUDIO).unwrap();
    let before = package
        .body_media_properties(MediaSelector::audio(0))
        .unwrap();
    let settings = MediaPlaybackSettings {
        start_time: None,
        end_time: before.playback().unwrap().end_time,
        poster_time: None,
        loop_mode: Some(MediaLoopMode::Repeat),
        volume: Some(MediaVolume::new(0.5).unwrap()),
    };
    let mut edit = package
        .edit_body_media_properties(MediaSelector::audio(0))
        .unwrap();
    edit.set_playback(settings).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit
        .package()
        .body_media_properties(MediaSelector::audio(0))
        .unwrap();
    assert_eq!(after.playback(), Some(settings.canonicalize().unwrap()));
    assert_eq!(after.geometry(), before.geometry());
    assert_eq!(after.properties(), before.properties());
}

#[test]
fn verbs_that_do_not_apply_to_the_kind_are_refused() {
    let audio = Package::from_bytes(AUDIO).unwrap();
    let mut edit = audio
        .edit_body_media_properties(MediaSelector::audio(0))
        .unwrap();
    for error in [
        edit.set_geometry(Geometry::default()).unwrap_err(),
        edit.flip(FlipAxis::Vertical).unwrap_err(),
        edit.restore_original_size().unwrap_err(),
    ] {
        assert!(matches!(
            error,
            Error::Refused(Refusal::UnsupportedOperation { kind: "audio", .. })
        ));
    }
    let images = Package::from_bytes(IMAGES).unwrap();
    let mut edit = images
        .edit_body_media_properties(MediaSelector::image(0))
        .unwrap();
    let error = edit
        .set_playback(MediaPlaybackSettings::new(Duration::from_secs(1)))
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Refused(Refusal::UnsupportedOperation { kind: "image", .. })
    ));
    let error = edit
        .set_position(Point {
            x: f32::NAN,
            y: 0.0,
        })
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Refused(Refusal::InvalidValue { .. })
    ));
    if let Error::Refused(refusal) = error {
        assert_eq!(refusal.kind(), litchi_core::edit::RefusalKind::InvalidValue);
    }
}

#[test]
fn missing_drawables_are_not_found() {
    let images = Package::from_bytes(IMAGES).unwrap();
    for selector in [
        MediaSelector::image(99),
        MediaSelector::movie(0),
        MediaSelector::audio(0),
    ] {
        let error = images.body_media_properties(selector).unwrap_err();
        assert!(
            matches!(error, Error::Refused(Refusal::NotFound { .. })),
            "{selector:?}: {error:?}"
        );
    }
}
