use super::codec::*;
use super::model::*;
use super::transaction::Snapshot;
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
    Record::parse(&external_object_list_bytes(seed, objects), 0)
        .unwrap()
        .0
}

fn external_object_list_bytes(seed: i32, objects: &[Vec<u8>]) -> Vec<u8> {
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
    record_bytes(0x0f, 0, RecordType::ExObjList.as_u16(), &payload).unwrap()
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

fn movie_object(id: u32, path: Option<&str>) -> Object {
    Object::Movie(Movie {
        kind: MovieKind::Avi,
        video: Video {
            media: Media {
                id,
                loop_playback: false,
                rewind_after_playing: false,
                narration: false,
                unused: [0xA5, 0x5A],
            },
            path: path.map(str::to_owned),
        },
    })
}

fn document_bytes(payload: &[u8]) -> Vec<u8> {
    record_bytes(0x0F, 0, RecordType::Document.as_u16(), payload).unwrap()
}

fn officeart_media_owner(id: u32) -> Vec<u8> {
    let reference = record_bytes(
        0,
        0,
        RecordType::ExternalObjectRefAtom.as_u16(),
        &id.to_le_bytes(),
    )
    .unwrap();
    let client_data = record_bytes(0, 0, 0xF011, &reference).unwrap();
    record_bytes(0x0F, 0, 0xF004, &client_data).unwrap()
}

#[test]
fn snapshot_transaction_rewrites_only_media_and_preserves_opaque_records() {
    let object = movie_object(3, None);
    let opaque = record_bytes(0, 7, 0x7ABC, &[0xDE, 0xAD, 0xBE, 0xEF]).unwrap();
    let list = external_object_list_bytes(3, &[opaque.clone(), object.to_record_bytes().unwrap()]);
    let trailer = record_bytes(0, 0, 0x7ABD, &[1, 2, 3, 4, 5]).unwrap();
    let source_bytes = document_bytes(&[list.clone(), trailer.clone()].concat());
    let source = Snapshot::parse(&source_bytes).unwrap();

    let mut editor = source.edit();
    editor
        .set_path(3, Some(r"\\server\share\clip.avi".into()))
        .unwrap();
    let commit = editor.commit().unwrap();
    let target = commit.snapshot();

    assert_ne!(target.bytes(), source.bytes());
    assert!(
        target
            .bytes()
            .windows(opaque.len())
            .any(|window| window == opaque)
    );
    assert!(
        target
            .bytes()
            .windows(trailer.len())
            .any(|window| window == trailer)
    );
    assert_eq!(
        u32::from_le_bytes(target.bytes()[4..8].try_into().unwrap()) as usize,
        target.bytes().len() - 8
    );
    assert_eq!(
        target
            .collection()
            .unwrap()
            .get(3)
            .and_then(|object| match object {
                Object::Movie(value) => value.video.path.as_deref(),
                _ => None,
            }),
        Some(r"\\server\share\clip.avi")
    );
    assert_eq!(commit.patch().changes().len(), 1);
    assert_eq!(commit.patch().undo(target).unwrap().bytes(), source.bytes());
    assert_eq!(
        commit.patch().redo(&source).unwrap().bytes(),
        target.bytes()
    );
}

#[test]
fn snapshot_transaction_supports_exact_noop_and_source_stale_rejection() {
    let object = movie_object(3, Some("clip.avi"));
    let source_bytes = document_bytes(&external_object_list_bytes(
        3,
        &[object.to_record_bytes().unwrap()],
    ));
    let source = Snapshot::parse(&source_bytes).unwrap();

    let noop = source.edit().commit().unwrap();
    assert!(noop.patch().is_empty());
    assert_eq!(noop.snapshot().bytes(), source.bytes());

    let mut editor = source.edit();
    editor.set_path(3, Some("clip.avi".into())).unwrap();
    let noop = editor.commit().unwrap();
    assert!(noop.patch().is_empty());
    assert!(noop.patch().changes().is_empty());

    let changed = {
        let mut editor = source.edit();
        editor.set_path(3, Some("new.avi".into())).unwrap();
        editor.commit().unwrap()
    };
    let mut stale_bytes = source.bytes().to_vec();
    *stale_bytes.last_mut().unwrap() ^= 1;
    let stale = Snapshot::parse(&stale_bytes).unwrap();
    assert!(changed.patch().apply(&stale).is_err());
}

#[test]
fn owner_validation_blocks_media_removal_and_keeps_failed_edits_atomic() {
    let object = movie_object(3, None);
    let list = external_object_list_bytes(3, &[object.to_record_bytes().unwrap()]);
    let drawing = record_bytes(
        0,
        0,
        RecordType::PPDrawing.as_u16(),
        &officeart_media_owner(3),
    )
    .unwrap();
    let source = Snapshot::parse(&document_bytes(&[list, drawing].concat())).unwrap();
    assert_eq!(source.owner_ids(), &[3]);
    assert_eq!(source.owner_count(3), 1);

    let mut editor = source.edit();
    assert!(editor.remove(3).is_err());
    assert_eq!(editor.objects().len(), 1);
    assert!(
        editor
            .set_playback(3, Playback::new(false, false, true))
            .is_err()
    );
    assert_eq!(
        editor.objects()[0],
        *source.collection().unwrap().get(3).unwrap()
    );
}

#[test]
fn insert_attaches_a_missing_list_to_document_and_remove_preserves_unknown_order() {
    let trailer = record_bytes(0, 4, 0x7AC0, &[8, 9]).unwrap();
    let source = Snapshot::parse(&document_bytes(&trailer)).unwrap();
    assert!(source.collection().is_none());

    let mut editor = source.edit();
    editor.insert(movie_object(12, Some("new.avi"))).unwrap();
    let commit = editor.commit().unwrap();
    assert_eq!(commit.snapshot().collection().unwrap().id_seed, 12);
    assert!(
        commit
            .snapshot()
            .bytes()
            .windows(trailer.len())
            .any(|window| window == trailer)
    );

    let mut editor = commit.snapshot().edit();
    let removed = editor.remove(12).unwrap();
    assert_eq!(removed.id(), 12);
    let empty = editor.commit().unwrap();
    assert!(empty.snapshot().collection().unwrap().objects.is_empty());
}

#[test]
fn bounded_snapshot_rejects_truncated_roots_and_overlarge_sources() {
    let object = movie_object(3, None);
    let source_bytes = document_bytes(&external_object_list_bytes(
        3,
        &[object.to_record_bytes().unwrap()],
    ));
    assert!(Snapshot::parse(&source_bytes[..source_bytes.len() - 1]).is_err());
    let limits = Limits {
        max_root_bytes: source_bytes.len() - 1,
        ..Limits::default()
    };
    assert!(Snapshot::parse_with_limits(&source_bytes, limits).is_err());
}
