use super::codec::*;
use super::model::*;
use crate::consts::RecordType;
use crate::records::Record;

fn cd_audio() -> CdAudio {
    CdAudio {
        media: Media {
            id: 17,
            loop_playback: true,
            rewind_after_playing: true,
            narration: true,
            unused: [0xaa, 0x55],
        },
        start: CdTime::new(1, 2, 3, 4).unwrap(),
        end: CdTime::new(2, 0, 0, 0).unwrap(),
    }
}

#[test]
fn protocol_shaped_cd_audio_roundtrips_losslessly() {
    let expected = cd_audio();
    let parsed = CdAudio::parse(&expected.to_record().unwrap()).unwrap();
    assert_eq!(parsed, expected);
    assert_eq!(
        parsed.to_record_bytes().unwrap(),
        expected.to_record_bytes().unwrap()
    );
}

#[test]
fn rejects_hostile_media_flags_times_order_and_headers() {
    for components in [
        (0, 0, 0, 0),
        (101, 0, 0, 0),
        (1, 61, 0, 0),
        (1, 0, 60, 0),
        (1, 0, 0, 74),
    ] {
        assert!(CdTime::new(components.0, components.1, components.2, components.3).is_err());
    }
    let mut value = cd_audio();
    value.end = CdTime::new(1, 0, 0, 0).unwrap();
    assert!(value.to_record_bytes().is_err());
    let mut bytes = cd_audio().to_record_bytes().unwrap();
    bytes[20] = 0x08;
    let record = Record::parse(&bytes, 0).unwrap().0;
    assert!(CdAudio::parse(&record).is_err());
    bytes = cd_audio().to_record_bytes().unwrap();
    bytes[0] = 0;
    let record = Record::parse(&bytes, 0).unwrap().0;
    assert!(CdAudio::parse(&record).is_err());
}

#[test]
fn external_video_roundtrips_optional_inert_paths() {
    for path in [None, Some(r"\\server\share\movie.avi".to_string())] {
        let expected = Video {
            media: Media {
                id: 23,
                loop_playback: false,
                rewind_after_playing: true,
                narration: false,
                unused: [3, 4],
            },
            path,
        };
        let parsed = Video::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }
}

#[test]
fn external_video_rejects_narration_and_hostile_paths() {
    let mut video = Video {
        media: cd_audio().media,
        path: Some("movie.avi".into()),
    };
    assert!(video.to_record_bytes().is_err());
    video.media.narration = false;
    video.path = Some("bad\0path".into());
    assert!(video.to_record_bytes().is_err());
    video.path = Some("x".repeat(MAX_PATH_UNITS + 1));
    assert!(video.to_record_bytes().is_err());
}

#[test]
fn avi_and_mci_movie_containers_roundtrip_canonically() {
    for kind in [MovieKind::Avi, MovieKind::Mci] {
        let expected = Movie {
            kind,
            video: Video {
                media: Media {
                    id: 41,
                    loop_playback: true,
                    rewind_after_playing: false,
                    narration: false,
                    unused: [0x12, 0x34],
                },
                path: Some(r"\\server\share\movie.avi".into()),
            },
        };
        let parsed = Movie::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }
}

#[test]
fn movie_containers_reject_wrong_headers_and_extra_children() {
    let movie = Movie {
        kind: MovieKind::Avi,
        video: Video {
            media: Media {
                id: 42,
                loop_playback: false,
                rewind_after_playing: false,
                narration: false,
                unused: [0, 0],
            },
            path: None,
        },
    };
    let mut bytes = movie.to_record_bytes().unwrap();
    bytes[0] = 0;
    assert!(Movie::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());

    let child = movie.video.to_record_bytes().unwrap();
    let doubled = [child.as_slice(), child.as_slice()].concat();
    let bytes = record_bytes(0x0f, 0, RecordType::ExternalAviMovie.as_u16(), &doubled).unwrap();
    assert!(Movie::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

#[test]
fn midi_and_linked_wav_roundtrip_optional_inert_paths() {
    for kind in [LinkedAudioKind::Midi, LinkedAudioKind::Wav] {
        let expected = LinkedAudio {
            kind,
            media: cd_audio().media,
            path: Some(r"C:\media\theme.mid".into()),
        };
        assert_eq!(
            LinkedAudio::parse(&expected.to_record().unwrap()).unwrap(),
            expected
        );
    }
}

#[test]
fn linked_audio_rejects_misplaced_children_and_hostile_paths() {
    let mut audio = LinkedAudio {
        kind: LinkedAudioKind::Wav,
        media: cd_audio().media,
        path: Some("bad\0path".into()),
    };
    assert!(audio.to_record_bytes().is_err());
    audio.path = None;
    let child = audio.media.to_record_bytes().unwrap();
    let payload = [child.as_slice(), child.as_slice()].concat();
    let bytes = record_bytes(0x0f, 0, RecordType::ExternalWavAudioLink.as_u16(), &payload).unwrap();
    assert!(LinkedAudio::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

#[test]
fn embedded_wav_roundtrips_nullable_sound_reference() {
    for sound_id in [None, Some(7)] {
        let expected = EmbeddedWav {
            media: cd_audio().media,
            sound_id,
            duration_ms: 90_000,
        };
        assert_eq!(
            EmbeddedWav::parse(&expected.to_record().unwrap()).unwrap(),
            expected
        );
    }
}

#[test]
fn embedded_wav_rejects_negative_or_overflowing_duration() {
    let value = EmbeddedWav {
        media: cd_audio().media,
        sound_id: Some(1),
        duration_ms: i32::MAX as u32 + 1,
    };
    assert!(value.to_record_bytes().is_err());

    let mut bytes = EmbeddedWav {
        duration_ms: 1,
        ..value
    }
    .to_record_bytes()
    .unwrap();
    let duration_offset = 8 + 16 + 8 + 4;
    bytes[duration_offset..duration_offset + 4].copy_from_slice(&(-1i32).to_le_bytes());
    assert!(EmbeddedWav::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

fn external_object_list(seed: i32, objects: &[Vec<u8>]) -> Record {
    let mut payload = record_bytes(
        0,
        0,
        RecordType::ExObjListAtom.as_u16(),
        &seed.to_le_bytes(),
    )
    .unwrap();
    for object in objects {
        payload.extend_from_slice(object);
    }
    let bytes = record_bytes(0x0f, 0, RecordType::ExObjList.as_u16(), &payload).unwrap();
    Record::parse(&bytes, 0).unwrap().0
}

fn hyperlink(id: u32) -> Vec<u8> {
    let unicode = |value: &str| {
        value
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>()
    };
    let mut payload = record_bytes(
        0,
        0,
        RecordType::ExternalHyperlinkAtom.as_u16(),
        &id.to_le_bytes(),
    )
    .unwrap();
    payload
        .extend(record_bytes(0, 0, RecordType::CString.as_u16(), &unicode("Media link")).unwrap());
    payload.extend(
        record_bytes(
            0,
            1,
            RecordType::CString.as_u16(),
            &unicode("https://example.test"),
        )
        .unwrap(),
    );
    payload.extend(record_bytes(0, 3, RecordType::CString.as_u16(), &unicode("slide")).unwrap());
    record_bytes(0x0f, 0, RecordType::ExternalHyperlink.as_u16(), &payload).unwrap()
}

#[test]
fn media_collection_discovers_typed_objects_and_resolves_ids() {
    let movie = Movie {
        kind: MovieKind::Mci,
        video: Video {
            media: Media {
                id: 3,
                loop_playback: false,
                rewind_after_playing: false,
                narration: false,
                unused: [0, 0],
            },
            path: None,
        },
    };
    let audio = LinkedAudio {
        kind: LinkedAudioKind::Midi,
        media: Media {
            id: 7,
            ..cd_audio().media
        },
        path: Some("theme.mid".into()),
    };
    let root = external_object_list(
        7,
        &[
            movie.to_record_bytes().unwrap(),
            audio.to_record_bytes().unwrap(),
        ],
    );
    let parsed = Collection::parse(&root).unwrap().unwrap();
    assert_eq!(parsed.id_seed, 7);
    assert_eq!(parsed.objects.len(), 2);
    assert!(parsed.get(3).is_some());
    assert!(parsed.get(7).is_some());
}

#[test]
fn media_collection_rejects_duplicate_ids_and_low_seed() {
    let first = LinkedAudio {
        kind: LinkedAudioKind::Midi,
        media: Media {
            id: 9,
            ..cd_audio().media
        },
        path: None,
    }
    .to_record_bytes()
    .unwrap();
    let second = LinkedAudio {
        kind: LinkedAudioKind::Wav,
        media: Media {
            id: 9,
            ..cd_audio().media
        },
        path: None,
    }
    .to_record_bytes()
    .unwrap();
    assert!(Collection::parse(&external_object_list(9, &[first.clone(), second])).is_err());
    assert!(Collection::parse(&external_object_list(8, &[first])).is_err());
}

#[test]
fn media_collection_rejects_cross_family_hyperlink_id_collisions() {
    let media = LinkedAudio {
        kind: LinkedAudioKind::Midi,
        media: Media {
            id: 11,
            ..cd_audio().media
        },
        path: None,
    }
    .to_record_bytes()
    .unwrap();
    let root = external_object_list(11, &[media, hyperlink(11)]);
    assert!(Collection::parse(&root).is_err());
}
