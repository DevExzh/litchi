//! Keynote build parity cases (wave W1-C, record 0372): build creation for
//! every host constructor family, update and removal through
//! `slide::build`, and the build read.

use litchi_iwa::keynote::{self, BuildSettings};
use litchi_iwa_protos::kn;
use litchi_keynote::build::{
    Acceleration, Action, Emphasis, FlipDirection, HorizontalDirection, Keyboard,
    KeyboardDirection, Motion, ObjectEffect, Opacity, Path, Point, Rotation, RotationDirection,
    Scale, Start, SwooshDirection, TimingCurve,
};
use litchi_keynote::slide::build::{Settings, Stage};
use litchi_keynote::{DrawableSelector, Package, Position, Seconds, SlideSelector};
use prost::Message as _;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::keynote_drawable_cases::{TEXT_BOX, host_id};
use crate::known_fields::{Decoded, decoded};
use crate::leaf_bytes;

const SLIDE: usize = 0;

/// Documented divergence: the leaf derives a build's random seed and its
/// chunk UUID deterministically (the host draws them at random); both are
/// cleared before comparing.
fn build_without_seed(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut build = kn::BuildArchive::decode(data)?;
    if let Some(animation) = build.attributes.animation_attributes.as_mut() {
        animation.random_number_seed = None;
    }
    Ok(decoded(build, data))
}

fn chunk_without_uuid(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    let mut chunk = kn::BuildChunkArchive::decode(data)?;
    chunk.build_id = None;
    if let Some(identifier) = chunk.build_chunk_identifier.as_mut() {
        identifier.build_id = None;
    }
    Ok(decoded(chunk, data))
}

/// Host read: every build of slide 0 (effect, stage, timing, start).
fn host_builds(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok(editor
        .slide_builds(SLIDE)?
        .into_iter()
        .map(|build| {
            let settings = &build.settings;
            format!(
                "{} {} {:?} {:?} {:?} chunks {}",
                settings
                    .semantic()
                    .map(|s| format!("{:?}", s.effect()))
                    .unwrap_or_default(),
                settings
                    .start()
                    .map(|s| format!("{s:?}"))
                    .unwrap_or_default(),
                settings.duration().map(Seconds::as_f64).ok(),
                settings.delay().map(Seconds::as_f64).ok(),
                build.drawable_object_id != 0,
                build.chunks.len()
            )
        })
        .collect())
}

/// Leaf read: every build of slide 0.
fn leaf_builds(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    Ok(package
        .slide_build_list(SlideSelector::index(SLIDE))?
        .iter()
        .map(|build| {
            let settings = build.settings();
            format!(
                "{:?} {:?} {:?} {} {} chunks {}",
                settings.stage(),
                settings.effect(),
                settings.start(),
                settings.duration().as_f64(),
                settings.delay().as_f64(),
                build.chunks()
            )
        })
        .collect())
}

fn leaf_run(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_keynote::slide::build::Edit<'_>) -> TestResult,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_builds();
    stage(&mut edit)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_builds(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    input: &Fixture,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft, &[u8]) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    Case::new(name, Format::Keynote, tier, evidence, input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor, source)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .known_fields(|known| {
            known.with(8, "KN.BuildArchive", build_without_seed).with(
                153,
                "KN.BuildChunkArchive",
                chunk_without_uuid,
            )
        })
        .read("host builds", host_builds)
        .read("leaf builds", leaf_builds)
        .run()?;
    Ok(())
}

fn seconds(value: f64) -> TestResult<Seconds> {
    Ok(Seconds::new(value)?)
}

/// Host and leaf settings for every host constructor family.
fn families() -> TestResult<Vec<(&'static str, BuildSettings, Settings)>> {
    use keynote::{
        FlipDirection as HostFlip, HorizontalBuildDirection as HostHorizontal,
        JiggleIntensity as HostJiggle, KeyboardDirection as HostKeyboard,
        RotationDirection as HostRotation, SwooshDirection as HostSwoosh,
    };
    let ease = Acceleration::EaseInOut;
    Ok(vec![
        (
            "appear_in",
            BuildSettings::appear_in(),
            Settings::appear(Stage::In),
        ),
        (
            "appear_out",
            BuildSettings::appear_out(),
            Settings::appear(Stage::Out),
        ),
        (
            "dissolve_in",
            BuildSettings::dissolve_in(),
            Settings::object(Stage::In, ObjectEffect::Dissolve),
        ),
        (
            "dissolve_out",
            BuildSettings::dissolve_out(),
            Settings::object(Stage::Out, ObjectEffect::Dissolve),
        ),
        (
            "shimmer_in",
            BuildSettings::shimmer_in(),
            Settings::object(Stage::In, ObjectEffect::Shimmer),
        ),
        (
            "shimmer_out",
            BuildSettings::shimmer_out(),
            Settings::object(Stage::Out, ObjectEffect::Shimmer),
        ),
        (
            "skid_in",
            BuildSettings::skid_in(HostHorizontal::RightToLeft),
            Settings::object(
                Stage::In,
                ObjectEffect::Skid(HorizontalDirection::RightToLeft),
            ),
        ),
        (
            "skid_out",
            BuildSettings::skid_out(HostHorizontal::LeftToRight),
            Settings::object(
                Stage::Out,
                ObjectEffect::Skid(HorizontalDirection::LeftToRight),
            ),
        ),
        (
            "swoosh_in",
            BuildSettings::swoosh_in(HostSwoosh::FromLeft),
            Settings::object(Stage::In, ObjectEffect::Swoosh(SwooshDirection::FromLeft)),
        ),
        (
            "swoosh_out",
            BuildSettings::swoosh_out(HostSwoosh::Center),
            Settings::object(Stage::Out, ObjectEffect::Swoosh(SwooshDirection::Center)),
        ),
        (
            "trace_in",
            BuildSettings::trace_in(HostHorizontal::LeftToRight),
            Settings::object(
                Stage::In,
                ObjectEffect::Trace(HorizontalDirection::LeftToRight),
            ),
        ),
        (
            "trace_out",
            BuildSettings::trace_out(HostHorizontal::RightToLeft),
            Settings::object(
                Stage::Out,
                ObjectEffect::Trace(HorizontalDirection::RightToLeft),
            ),
        ),
        (
            "rotate_action",
            BuildSettings::rotate_action(270.0, HostRotation::Counterclockwise),
            Settings::action(Action::Rotate(Rotation::new(
                270.0,
                RotationDirection::Counterclockwise,
                ease,
            )?)),
        ),
        (
            "scale_action",
            BuildSettings::scale_action(1.5),
            Settings::action(Action::Scale(Scale::new(1.5, ease)?)),
        ),
        (
            "opacity_action",
            BuildSettings::opacity_action(40.0),
            Settings::action(Action::Opacity(Opacity::new(40.0, ease)?)),
        ),
        (
            "move_action",
            BuildSettings::move_action(120.0, -40.0),
            Settings::action(Action::Move(Motion::new(
                Path::straight(120.0, -40.0)?,
                false,
                ease,
            ))),
        ),
        (
            "blink_action",
            BuildSettings::blink_action(3, true),
            Settings::emphasis(Emphasis::blink(3, true)?),
        ),
        (
            "bounce_action",
            BuildSettings::bounce_action(2, false),
            Settings::emphasis(Emphasis::bounce(2, false)?),
        ),
        (
            "flip_action",
            BuildSettings::flip_action(1, HostFlip::RightToLeft),
            Settings::emphasis(Emphasis::flip(1, FlipDirection::RightToLeft)?),
        ),
        (
            "jiggle_action",
            BuildSettings::jiggle_action(HostJiggle::Large),
            Settings::emphasis(Emphasis::jiggle(
                litchi_keynote::build::JiggleIntensity::Large,
            )),
        ),
        (
            "pop_action",
            BuildSettings::pop_action(150.0),
            Settings::emphasis(Emphasis::pop(150.0)?),
        ),
        (
            "pulse_action",
            BuildSettings::pulse_action(2, 120.0),
            Settings::emphasis(Emphasis::pulse(2, 120.0)?),
        ),
        (
            "keyboard_in",
            BuildSettings::keyboard_in(HostKeyboard::Forward, true),
            Settings::keyboard(Stage::In, Keyboard::new(KeyboardDirection::Forward, true)),
        ),
        (
            "keyboard_out",
            BuildSettings::keyboard_out(HostKeyboard::Backward, false),
            Settings::keyboard(
                Stage::Out,
                Keyboard::new(KeyboardDirection::Backward, false),
            ),
        ),
        (
            "with_timing",
            BuildSettings::appear_in()
                .with_start(Start::AfterTransition)?
                .with_delay(seconds(0.25)?)?
                .with_duration(seconds(2.0)?)?,
            Settings::appear(Stage::In)
                .with_start(Start::AfterTransition)?
                .with_delay(seconds(0.25)?)?
                .with_duration(seconds(2.0)?)?,
        ),
        (
            "with_custom_timing_curve",
            BuildSettings::scale_action(2.0).with_custom_timing_curve(
                keynote::BuildTimingCurve::cubic(
                    keynote::MotionPathPoint::new(0.25, 0.0),
                    keynote::MotionPathPoint::new(0.75, 1.0),
                ),
            )?,
            Settings::action(Action::Scale(Scale::new(2.0, ease)?)).with_timing_curve(
                TimingCurve::cubic(Point::new(0.25, 0.0)?, Point::new(0.75, 1.0)?)?,
            )?,
        ),
    ])
}

#[test]
fn keynote_slide_build_add_parity() -> TestResult {
    let input = fixture(TEXT_BOX)?;
    for (name, host_settings, leaf_settings) in families()? {
        run(
            "keynote_slide_build_add_parity",
            &input,
            Tier::B,
            Evidence::Queued,
            |editor, source| {
                let id = host_id(editor, source, 0)?;
                editor.add_slide_build(SLIDE, id, host_settings)?;
                Ok(())
            },
            |source| {
                leaf_run(source, |edit| {
                    edit.add(
                        SlideSelector::index(SLIDE),
                        DrawableSelector::index(0),
                        leaf_settings,
                    )?;
                    Ok(())
                })
            },
        )
        .map_err(|error| format!("{name}: {error}"))?;
    }
    Ok(())
}

/// `basic.key` with two leaf-authored builds on its text box: Appear In,
/// then Dissolve Out After Previous.
fn builds_fixture() -> TestResult<Fixture> {
    let base = fixture(TEXT_BOX)?;
    let package = Package::from_bytes(&base.bytes)?;
    let mut edit = package.edit_slide_builds();
    edit.add(
        SlideSelector::index(SLIDE),
        DrawableSelector::index(0),
        Settings::appear(Stage::In),
    )?
    .add(
        SlideSelector::index(SLIDE),
        DrawableSelector::index(0),
        Settings::object(Stage::Out, ObjectEffect::Dissolve).with_start(Start::AfterPrevious)?,
    )?;
    let commit = edit.commit()?;
    Ok(Fixture {
        path: "iwork/keynote/basic.key + leaf builds (appear in; dissolve out after previous)",
        bytes: leaf_bytes!(commit.package()),
    })
}

fn host_build(editor: &keynote::Draft, position: usize) -> TestResult<u64> {
    Ok(editor
        .slide_builds(SLIDE)?
        .get(position)
        .ok_or("no such build")?
        .object_id)
}

#[test]
fn keynote_slide_build_set_parity() -> TestResult {
    let input = builds_fixture()?;
    run(
        "keynote_slide_build_set_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, _| {
            let build = host_build(editor, 1)?;
            let mut settings = editor.slide_builds(SLIDE)?[1].settings.clone();
            settings.set_duration(seconds(2.5)?)?;
            settings.set_delay(seconds(0.5)?)?;
            editor.set_slide_build(SLIDE, build, settings)?;
            Ok(())
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let current = package.slide_build_list(SlideSelector::index(SLIDE))?[1]
                .settings()
                .clone();
            leaf_run(source, |edit| {
                edit.set(
                    SlideSelector::index(SLIDE),
                    Position::new(1),
                    current
                        .with_duration(seconds(2.5)?)?
                        .with_delay(seconds(0.5)?)?,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_build_remove_parity() -> TestResult {
    let input = builds_fixture()?;
    run(
        "keynote_slide_build_remove_parity",
        &input,
        Tier::B,
        Evidence::Queued,
        |editor, _| {
            let build = host_build(editor, 1)?;
            editor.remove_slide_build(SLIDE, build)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.remove(SlideSelector::index(SLIDE), Position::new(1))?;
                Ok(())
            })
        },
    )
}

/// Host and leaf build reads agree on native fixtures with builds.
fn builds_agree(bytes: &[u8]) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let host = editor.slide_builds(SLIDE)?;
    let leaf = Package::from_bytes(bytes)?.slide_build_list(SlideSelector::index(SLIDE))?;
    if host.len() != leaf.len() {
        return Err(format!("build count: host {}, leaf {}", host.len(), leaf.len()).into());
    }
    for (host, leaf) in host.iter().zip(leaf.iter()) {
        let settings = leaf.settings();
        if host.settings.duration()?.as_f64() != settings.duration().as_f64()
            || host.settings.delay()?.as_f64() != settings.delay().as_f64()
            || host.settings.start()? != settings.start()
            || host.chunks.len() != leaf.chunks()
            || host.settings.effect()? != *settings.effect()
        {
            return Err(format!("host {:?} vs leaf {:?}", host.settings, settings).into());
        }
    }
    Ok(leaf.len())
}

#[test]
fn keynote_slide_build_read_parity() -> TestResult {
    let derived = builds_fixture()?;
    for input in [
        fixture("iwork/keynote/slide-movie-creation-focused-native.key")?,
        fixture("iwork/keynote/media-lifecycle-duplicate-native.key")?,
        derived,
    ] {
        Case::new(
            "keynote_slide_build_read_parity",
            Format::Keynote,
            Tier::A,
            Evidence::Queued,
            &input,
        )
        .read_only()
        .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
        .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
        .read("host and leaf builds agree", builds_agree)
        .run()?;
    }
    Ok(())
}
