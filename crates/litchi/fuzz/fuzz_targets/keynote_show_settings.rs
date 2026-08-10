#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::keynote::{
    Limits, Package, ReadError, ReadOptions, Seconds, SemanticLimits,
    show::{Mode, Settings, Size},
    soundtrack::{
        Commit as SoundtrackCommit, Error as SoundtrackError, Mode as SoundtrackMode,
        Settings as SoundtrackSettings,
    },
};

const MAX_INPUT_BYTES: u64 = 1024 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 1024 * 1024 + 1;
const MAX_ENTRIES: usize = 256;
const MAX_ENTRY_BYTES: u64 = 2 * 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 8 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 2 * 1024 * 1024;
const MAX_OBJECTS: usize = 16 * 1024;
const MAX_SLIDES: usize = 512;
const MAX_REFERENCES: usize = 32 * 1024;
const MAX_TEXT_STORAGES: usize = 8 * 1024;
const MAX_TEXT_FRAGMENTS: usize = 32 * 1024;
const MAX_TEXT_BYTES: usize = 2 * 1024 * 1024;
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_keynote_show_input_e642__";
const NATIVE_KEYNOTE: &[u8] = include_bytes!("../../../../test-data/iwork/keynote/basic.key");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_options(data, fuzz_options()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // CRC-protected arbitrary packages rarely reach the strict show codec.
    // Reuse a fixed prefix as bounded settings commands against a genuine
    // Keynote package so every input reaches the public transaction surface.
    exercise_package(native_package(), data);
    exercise_semantic_validation(data);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
});

fn fuzz_options() -> ReadOptions {
    static OPTIONS: OnceLock<ReadOptions> = OnceLock::new();
    *OPTIONS.get_or_init(|| {
        let archive = Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz archive limits: {error}"));
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            MAX_REFERENCES,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Keynote fuzz semantic limits: {error}"));
        ReadOptions::new(archive, semantic)
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_options(NATIVE_KEYNOTE, fuzz_options())
            .unwrap_or_else(|error| panic!("native Keynote fuzz seed must open: {error}"));
        package.show_settings().unwrap_or_else(|error| {
            panic!("native Keynote fuzz seed must expose show settings: {error}")
        });
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    exercise_soundtrack_settings(package, data);

    let before = match package.show_settings() {
        Ok(settings) => settings,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_settings(before);

    let command = control(data, 0) & 3;
    let after = match command {
        0 => before,
        1 => change_playback(before, data),
        2 => change_rendering(before, data),
        _ => change_rendering(change_playback(before, data), data),
    };
    let Ok(edit) = package.edit_show_settings() else {
        return;
    };
    black_box(edit.settings());
    let edit = edit.set(after);
    assert_eq!(edit.settings(), after);
    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    if command == 1 {
        assert_ne!(
            before, after,
            "the playback-only command must change settings"
        );
        assert_eq!(
            diagnostics.deleted_previews(),
            0,
            "playback-only settings must preserve root previews"
        );
    }
    assert_eq!(
        commit
            .package()
            .show_settings()
            .unwrap_or_else(|error| panic!("committed show settings must be readable: {error}")),
        after,
    );
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_show_settings(&patch)
        .unwrap_or_else(|error| panic!("fresh show-settings patch must apply: {error}"));
    assert_eq!(
        applied
            .package()
            .show_settings()
            .unwrap_or_else(|error| panic!("applied show settings must be readable: {error}")),
        after,
    );
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(commit.package())
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if before != after {
        match applied.package().apply_show_settings(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed show-settings patch must conflict with its target"),
        }
        match package.apply_show_settings(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed show-settings inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_show_settings(&inverse)
        .unwrap_or_else(|error| panic!("fresh show-settings inverse must apply: {error}"));
    assert_eq!(
        restored
            .package()
            .show_settings()
            .unwrap_or_else(|error| panic!("restored show settings must be readable: {error}")),
        before,
    );
    assert_eq!(package_bytes(restored.package()), package_bytes(package));
}

fn exercise_soundtrack_settings(package: &Package, data: &[u8]) {
    let before = match package.soundtrack_settings() {
        Ok(Some(settings)) => settings,
        Ok(None) => {
            assert!(matches!(
                package.edit_soundtrack_settings(),
                Err(SoundtrackError::SoundtrackNotFound)
            ));
            return;
        },
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    observe_soundtrack_settings(before);

    let after = change_soundtrack_settings(before, data);
    let edit = match package.edit_soundtrack_settings() {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    assert_eq!(edit.settings(), before);
    let commit = match edit.set(after).commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    publish_soundtrack_and_reverse(package, before, after, commit);
}

fn change_soundtrack_settings(mut settings: SoundtrackSettings, data: &[u8]) -> SoundtrackSettings {
    match control(data, 20) % 6 {
        0 => {},
        1 => settings
            .set_volume(Some(distinct_soundtrack_volume(settings.volume(), data)))
            .unwrap_or_else(|error| unreachable!("bounded soundtrack volume is valid: {error}")),
        2 => match settings.volume() {
            Some(_) => settings.set_volume(None).unwrap_or_else(|error| {
                unreachable!("clearing soundtrack volume is valid: {error}")
            }),
            None => settings
                .set_volume(Some(distinct_soundtrack_volume(None, data)))
                .unwrap_or_else(|error| {
                    unreachable!("bounded soundtrack volume is valid: {error}")
                }),
        },
        3 => settings
            .set_mode(Some(distinct_soundtrack_mode(settings.mode(), data)))
            .unwrap_or_else(|error| unreachable!("named soundtrack mode is valid: {error}")),
        4 => match settings.mode() {
            Some(_) => settings
                .set_mode(None)
                .unwrap_or_else(|error| unreachable!("clearing soundtrack mode is valid: {error}")),
            None => settings
                .set_mode(Some(distinct_soundtrack_mode(None, data)))
                .unwrap_or_else(|error| unreachable!("named soundtrack mode is valid: {error}")),
        },
        _ => {
            settings
                .set_volume(Some(distinct_soundtrack_volume(settings.volume(), data)))
                .unwrap_or_else(|error| {
                    unreachable!("bounded soundtrack volume is valid: {error}")
                });
            settings
                .set_mode(Some(distinct_future_soundtrack_mode(settings.mode(), data)))
                .unwrap_or_else(|error| {
                    unreachable!("a future soundtrack mode is canonical: {error}")
                });
        },
    }
    settings
}

fn distinct_soundtrack_volume(current: Option<f64>, data: &[u8]) -> f64 {
    let candidate = f64::from(read_u16(data, 22)) / f64::from(u16::MAX);
    if current == Some(candidate) {
        if candidate == 1.0 { 0.0 } else { 1.0 }
    } else {
        candidate
    }
}

fn distinct_soundtrack_mode(current: Option<SoundtrackMode>, data: &[u8]) -> SoundtrackMode {
    let candidates = [
        SoundtrackMode::PlayOnce,
        SoundtrackMode::Loop,
        SoundtrackMode::DoNotPlay,
    ];
    let mut index = usize::from(control(data, 24)) % candidates.len();
    if current == Some(candidates[index]) {
        index = (index + 1) % candidates.len();
    }
    candidates[index]
}

fn distinct_future_soundtrack_mode(current: Option<SoundtrackMode>, data: &[u8]) -> SoundtrackMode {
    let mut raw = i32::from_le_bytes(read_u32(data, 26).to_le_bytes());
    if (0..=2).contains(&raw) {
        raw = 3;
    }
    if current == Some(SoundtrackMode::Unknown(raw)) {
        raw = if raw == i32::MAX { -1 } else { raw + 1 };
        if (0..=2).contains(&raw) {
            raw = 3;
        }
    }
    SoundtrackMode::Unknown(raw)
}

fn publish_soundtrack_and_reverse(
    package: &Package,
    before: SoundtrackSettings,
    after: SoundtrackSettings,
    commit: SoundtrackCommit,
) {
    let patch = commit.patch().clone();
    let diagnostics = commit.diagnostics();
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_eq!(patch.is_noop(), before == after);
    assert_eq!(diagnostics.changed(), before != after);
    assert_eq!(diagnostics.full_reparse_performed(), before != after);
    if before == after {
        assert_eq!(diagnostics.touched_components(), 0);
    } else {
        assert!(diagnostics.touched_components() > 0);
    }
    assert_eq!(
        commit
            .package()
            .soundtrack_settings()
            .unwrap_or_else(|error| panic!(
                "committed soundtrack settings must be readable: {error}"
            )),
        Some(after),
    );
    let source_bytes = package_bytes(package);
    let committed_bytes = package_bytes(commit.package());
    assert_eq!(patch.is_noop(), source_bytes == committed_bytes);
    black_box((
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
    ));

    let applied = package
        .apply_soundtrack_settings(&patch)
        .unwrap_or_else(|error| panic!("fresh soundtrack patch must apply: {error}"));
    assert_eq!(package_bytes(applied.package()), committed_bytes);
    assert_eq!(
        applied
            .package()
            .soundtrack_settings()
            .unwrap_or_else(|error| panic!(
                "applied soundtrack settings must be readable: {error}"
            )),
        Some(after),
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if !patch.is_noop() {
        match applied.package().apply_soundtrack_settings(&patch) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed soundtrack patch must conflict with its target"),
        }
        match package.apply_soundtrack_settings(&inverse) {
            Err(error) => observe_error(error),
            Ok(_) => panic!("a changed soundtrack inverse must conflict with its source"),
        }
    }

    let restored = applied
        .package()
        .apply_soundtrack_settings(&inverse)
        .unwrap_or_else(|error| panic!("fresh soundtrack inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_bytes);
    assert_eq!(
        restored
            .package()
            .soundtrack_settings()
            .unwrap_or_else(|error| panic!(
                "restored soundtrack settings must be readable: {error}"
            )),
        Some(before),
    );
}

fn change_playback(mut settings: Settings, data: &[u8]) -> Settings {
    settings.set_loop_presentation(Some(!settings.loop_presentation().unwrap_or(false)));
    settings.set_idle_timer_active(Some(!settings.idle_timer_active().unwrap_or(false)));
    settings.set_automatically_plays_upon_open(Some(
        !settings.automatically_plays_upon_open().unwrap_or(false),
    ));
    let mode = if settings.mode() == Some(Mode::SelfPlaying) {
        Mode::LinksOnly
    } else {
        Mode::SelfPlaying
    };
    settings
        .set_mode(Some(mode))
        .unwrap_or_else(|error| unreachable!("named Keynote mode is canonical: {error}"));

    let transition = duration(data, 2);
    let build = duration(data, 4);
    let idle = duration(data, 6);
    settings.set_autoplay_transition_delay(Some(transition));
    settings.set_autoplay_build_delay(Some(build));
    settings.set_idle_timer_delay(Some(idle));
    settings
}

fn change_rendering(mut settings: Settings, data: &[u8]) -> Settings {
    settings.set_slide_numbers_visible(Some(!settings.slide_numbers_visible().unwrap_or(false)));
    let size = settings.size();
    let width = distinct_dimension(size.width(), 1.0 + f32::from(read_u16(data, 2)));
    let height = distinct_dimension(size.height(), 1.0 + f32::from(read_u16(data, 4)));
    settings.set_size(
        Size::new(width, height)
            .unwrap_or_else(|error| unreachable!("bounded Keynote size is valid: {error}")),
    );
    settings
}

fn distinct_dimension(current: f32, candidate: f32) -> f32 {
    if current != candidate {
        candidate
    } else if current > 1.0 {
        current / 2.0
    } else {
        current + 1.0
    }
}

fn duration(data: &[u8], offset: usize) -> Seconds {
    Seconds::new(f64::from(read_u16(data, offset)) / 10.0)
        .unwrap_or_else(|error| unreachable!("bounded Keynote duration is valid: {error}"))
}

fn observe_settings(settings: Settings) {
    let size = settings.size();
    black_box((
        size.width(),
        size.height(),
        settings.slide_numbers_visible(),
        settings.loop_presentation(),
        settings.mode(),
        settings.autoplay_transition_delay(),
        settings.autoplay_build_delay(),
        settings.idle_timer_active(),
        settings.idle_timer_delay(),
        settings.automatically_plays_upon_open(),
    ));
}

fn observe_soundtrack_settings(settings: SoundtrackSettings) {
    black_box((settings.volume(), settings.mode()));
}

fn exercise_semantic_validation(data: &[u8]) {
    let raw = i32::from_le_bytes(read_u32(data, 4).to_le_bytes());
    observe_result(Mode::unknown(raw));
    observe_result(Size::new(
        f32::from_bits(read_u32(data, 8)),
        f32::from_bits(read_u32(data, 12)),
    ));
    observe_result(Seconds::new(f64::from_bits(read_u64(data, 16))));
    observe_result(Settings::default().validate());
    exercise_soundtrack_semantic_validation(data);
    observe_result(SemanticLimits::new(
        0,
        MAX_SLIDES,
        MAX_REFERENCES,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    ));
}

fn exercise_soundtrack_semantic_validation(data: &[u8]) {
    let raw = i32::from_le_bytes(read_u32(data, 28).to_le_bytes());
    let mode = SoundtrackMode::from_raw(raw);
    assert_eq!(mode.as_raw(), raw);
    black_box(mode.is_canonical());

    let absent = SoundtrackSettings::new(None, None)
        .unwrap_or_else(|error| unreachable!("absent soundtrack settings are valid: {error}"));
    assert_eq!((absent.volume(), absent.mode()), (None, None));
    let future = distinct_future_soundtrack_mode(None, data);
    let future_settings = SoundtrackSettings::new(Some(0.5), Some(future))
        .unwrap_or_else(|error| unreachable!("future soundtrack mode is valid: {error}"));
    assert_eq!(future_settings.mode(), Some(future));
    observe_result(SoundtrackSettings::new(
        Some(f64::from_bits(read_u64(data, 32))),
        Some(SoundtrackMode::Unknown(raw)),
    ));
    for known in [0, 1, 2] {
        assert!(SoundtrackSettings::new(None, Some(SoundtrackMode::Unknown(known))).is_err());
    }
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_options(PRIVATE_MALFORMED_INPUT, fuzz_options()) {
        Err(error) => observe_redacted(error, PRIVATE_MALFORMED_INPUT),
        Ok(_) => panic!("a private malformed sentinel must not parse as Keynote"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_options(bytes, fuzz_options()) {
        Err(ReadError::Archive(error)) => observe_error(error),
        Err(error) => panic!("an oversized Keynote input must return an archive limit: {error}"),
        Ok(_) => panic!("an oversized Keynote input must be rejected"),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([control(data, offset), control(data, offset + 1)])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control(data, offset),
        control(data, offset + 1),
        control(data, offset + 2),
        control(data, offset + 3),
    ])
}

fn read_u64(data: &[u8], offset: usize) -> u64 {
    u64::from_le_bytes([
        control(data, offset),
        control(data, offset + 1),
        control(data, offset + 2),
        control(data, offset + 3),
        control(data, offset + 4),
        control(data, offset + 5),
        control(data, offset + 6),
        control(data, offset + 7),
    ])
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a package to memory must succeed: {error}"));
    bytes
}

fn observe_result<T, E>(result: Result<T, E>)
where
    T: Debug,
    E: Debug + Display,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted(error: impl Debug + Display, private: &[u8]) {
    let private = std::str::from_utf8(private)
        .unwrap_or_else(|error| unreachable!("private sentinel is valid UTF-8: {error}"));
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
