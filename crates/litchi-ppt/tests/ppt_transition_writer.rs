#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Round-trip tests for slide transitions authored with the create-side
//! writer: a transition set via `Writer::set_slide_transition` must be emitted
//! as an `SSSlideInfoAtom` (MS-PPT 2.6.6) and read back by the crate's own
//! reader.

use std::io::Cursor;

use litchi_ppt::writer::{SlideTiming, WriteError, Writer};
use litchi_ppt::{
    AdvanceMode, Package, Record, SoundAction, TransitionDirection, TransitionInfo,
    TransitionSpeed, TransitionType,
};

fn write(writer: &mut Writer) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Payloads of every `SSSlideInfoAtom` (record type 1017) found in the Slide
/// containers (record type 1006) of the written presentation, in slide order.
fn slide_info_atom_payloads(bytes: &[u8]) -> Vec<Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["PowerPoint Document"]).unwrap();
    let mut payloads = Vec::new();
    let mut offset = 0;
    while offset < stream.len() {
        let (record, consumed) = Record::parse(&stream, offset).unwrap();
        if record.record_type_raw == 1006
            && let Some(atom) = record
                .children
                .iter()
                .find(|child| child.record_type_raw == 1017)
        {
            payloads.push(atom.data.clone());
        }
        offset += consumed;
    }
    payloads
}

#[test]
fn authored_transition_round_trips_through_reader() {
    let mut writer = Writer::new();
    let wipe_slide = writer.add_slide().unwrap();
    writer
        .set_slide_transition(
            wipe_slide,
            TransitionInfo::with_type(TransitionType::Wipe)
                .with_speed(TransitionSpeed::Fast)
                .with_direction(TransitionDirection::FromLeft)
                .with_advance_mode(AdvanceMode::OnClick),
        )
        .unwrap();
    let dissolve_slide = writer.add_slide().unwrap();
    writer
        .set_slide_transition(
            dissolve_slide,
            TransitionInfo::with_type(TransitionType::Dissolve)
                .with_speed(TransitionSpeed::Slow)
                .with_advance_mode(AdvanceMode::Automatic)
                .with_advance_time(3000),
        )
        .unwrap();
    let plain_slide = writer.add_slide().unwrap();

    let bytes = write(&mut writer);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    assert_eq!(slides.len(), 3);

    let wipe = slides[0].transition().unwrap().unwrap();
    assert_eq!(wipe.transition_type, TransitionType::Wipe);
    assert_eq!(wipe.speed, TransitionSpeed::Fast);
    assert_eq!(wipe.direction, TransitionDirection::FromLeft);
    assert_eq!(wipe.advance_mode, AdvanceMode::OnClick);
    assert_eq!(wipe.advance_time_ms, None);
    assert!(wipe.sound.is_none());
    assert!(wipe.has_effect());

    let dissolve = slides[1].transition().unwrap().unwrap();
    assert_eq!(dissolve.transition_type, TransitionType::Dissolve);
    assert_eq!(dissolve.speed, TransitionSpeed::Slow);
    assert_eq!(dissolve.advance_mode, AdvanceMode::Automatic);
    assert_eq!(dissolve.advance_time_ms, Some(3000));

    assert!(slides[plain_slide].transition().unwrap().is_none());
}

#[test]
fn authored_transition_record_layout_matches_spec() {
    let mut writer = Writer::new();
    let wipe_slide = writer.add_slide().unwrap();
    writer
        .set_slide_transition(
            wipe_slide,
            TransitionInfo::with_type(TransitionType::Wipe)
                .with_speed(TransitionSpeed::Fast)
                .with_direction(TransitionDirection::FromLeft)
                .with_advance_mode(AdvanceMode::OnClick),
        )
        .unwrap();
    let dissolve_slide = writer.add_slide().unwrap();
    writer
        .set_slide_transition(
            dissolve_slide,
            TransitionInfo::with_type(TransitionType::Dissolve)
                .with_speed(TransitionSpeed::Slow)
                .with_advance_mode(AdvanceMode::Automatic)
                .with_advance_time(3000),
        )
        .unwrap();
    writer.add_slide().unwrap();

    let payloads = slide_info_atom_payloads(&write(&mut writer));
    assert_eq!(payloads.len(), 2, "only transitioned slides carry the atom");

    // MS-PPT 2.6.6: slideTime(4) soundIdRef(4) effectDirection(1)
    // effectType(1) effectTransitionFlags(2) speed(1) unused(3)
    let wipe = &payloads[0];
    assert_eq!(wipe.len(), 16);
    assert_eq!(u32::from_le_bytes(wipe[0..4].try_into().unwrap()), 0);
    assert_eq!(u32::from_le_bytes(wipe[4..8].try_into().unwrap()), 0);
    assert_eq!(wipe[8], 0, "Wipe FromLeft direction (MS-PPT 2.6.6: 0=Left)");
    assert_eq!(wipe[9], 10, "MS-PPT 2.6.6 effect value for Wipe");
    assert_eq!(
        u16::from_le_bytes(wipe[10..12].try_into().unwrap()),
        0x0001,
        "fManualAdvance only"
    );
    assert_eq!(wipe[12], 2, "Fast speed");
    assert_eq!(&wipe[13..16], &[0, 0, 0]);

    let dissolve = &payloads[1];
    assert_eq!(dissolve.len(), 16);
    assert_eq!(u32::from_le_bytes(dissolve[0..4].try_into().unwrap()), 3000);
    assert_eq!(dissolve[9], 5, "MS-PPT 2.6.6 effect value for Dissolve");
    assert_eq!(
        u16::from_le_bytes(dissolve[10..12].try_into().unwrap()),
        0x0400,
        "fAutoAdvance only"
    );
    assert_eq!(dissolve[12], 0, "Slow speed");
}

#[test]
fn directional_transitions_round_trip() {
    let cases = [
        (TransitionType::Blinds, TransitionDirection::Vertical),
        (TransitionType::Blinds, TransitionDirection::Horizontal),
        (TransitionType::Checkerboard, TransitionDirection::Vertical),
        (TransitionType::Cover, TransitionDirection::FromRight),
        (TransitionType::Uncover, TransitionDirection::FromBottom),
        (TransitionType::Push, TransitionDirection::FromTop),
        (TransitionType::Strips, TransitionDirection::LeftDown),
        (TransitionType::Strips, TransitionDirection::RightUp),
        (TransitionType::Box, TransitionDirection::Out),
        (TransitionType::Box, TransitionDirection::In),
        (TransitionType::Split, TransitionDirection::HorizontalOut),
        (TransitionType::Split, TransitionDirection::HorizontalIn),
        (TransitionType::Split, TransitionDirection::VerticalOut),
        (TransitionType::Split, TransitionDirection::VerticalIn),
        (TransitionType::Wheel, TransitionDirection::Spokes3),
    ];

    let mut writer = Writer::new();
    for (transition_type, direction) in cases {
        let slide = writer.add_slide().unwrap();
        writer
            .set_slide_transition(
                slide,
                TransitionInfo::with_type(transition_type).with_direction(direction),
            )
            .unwrap();
    }

    let bytes = write(&mut writer);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();

    for (index, (transition_type, direction)) in cases.iter().enumerate() {
        let transition = slides[index].transition().unwrap().unwrap();
        assert_eq!(
            &transition.transition_type, transition_type,
            "slide {index} transition type"
        );
        assert_eq!(
            &transition.direction, direction,
            "slide {index} transition direction"
        );
    }
}

#[test]
fn hidden_timing_survives_when_transition_set() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .set_slide_transition(
            slide,
            TransitionInfo::with_type(TransitionType::Fade).with_advance_mode(AdvanceMode::OnClick),
        )
        .unwrap();
    writer
        .set_slide_timing(slide, SlideTiming::hidden())
        .unwrap();

    let bytes = write(&mut writer);

    let payloads = slide_info_atom_payloads(&bytes);
    assert_eq!(payloads.len(), 1, "transition and timing share one record");
    let flags = u16::from_le_bytes(payloads[0][10..12].try_into().unwrap());
    assert_ne!(flags & (1 << 0), 0, "fManualAdvance from the transition");
    assert_ne!(flags & (1 << 2), 0, "fHidden preserved from the timing");

    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();
    let timing = slides[0].timing().unwrap();
    assert!(timing.hidden);
    assert!(timing.advance_on_click);
    let transition = slides[0].transition().unwrap().unwrap();
    assert_eq!(transition.transition_type, TransitionType::Fade);
}

#[test]
fn timing_still_emitted_for_slides_without_transition() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    writer
        .set_slide_timing(slide, SlideTiming::auto_advance(5000))
        .unwrap();

    let bytes = write(&mut writer);
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let presentation = package.presentation().unwrap();
    let slides = presentation.slides().unwrap();

    let timing = slides[0].timing().unwrap();
    assert_eq!(timing.advance_time_ms, 5000);
    assert!(timing.auto_advance);
    assert!(timing.advance_on_click);

    // A timing-only atom has no transition effect but still round-trips.
    let transition = slides[0].transition().unwrap().unwrap();
    assert_eq!(transition.transition_type, TransitionType::None);
    assert!(!transition.has_effect());
}

#[test]
fn transition_with_sound_is_refused() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let transition = TransitionInfo {
        sound: Some(SoundAction::builtin("Applause")),
        ..TransitionInfo::default()
    };
    let result = writer.set_slide_transition(slide, transition);
    assert!(matches!(result, Err(WriteError::InvalidData(_))));
}

#[test]
fn transition_without_binary_ppt_representation_is_refused() {
    let mut writer = Writer::new();
    let slide = writer.add_slide().unwrap();
    let result =
        writer.set_slide_transition(slide, TransitionInfo::with_type(TransitionType::Morph));
    assert!(matches!(result, Err(WriteError::InvalidData(_))));
}

#[test]
fn transition_on_missing_slide_is_refused() {
    let mut writer = Writer::new();
    let result = writer.set_slide_transition(0, TransitionInfo::with_type(TransitionType::Fade));
    assert!(matches!(result, Err(WriteError::InvalidData(_))));
}
