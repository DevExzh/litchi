use super::VT_STREAM;
use super::{Limits, Snapshot};
use crate::property_set::model::{
    VT_STORAGE, VT_STORED_OBJECT, VT_STREAMED_OBJECT, VT_VERSIONED_STREAM,
};
use crate::property_set::{Guid, IndirectPropertyName, Section, Stream, Value, VersionedStream};
use litchi_cfb::OleWriter;
use std::io::Cursor;
use std::sync::Arc;

fn source_with_property(name: &str, value: Value, payload: &[u8], empty_storage: bool) -> Vec<u8> {
    let format_id = Guid::from_bytes([0x11; 16]);
    let mut section = Section::new(format_id);
    section.add(2, value).expect("synthetic indirect property");
    let contents = Stream {
        version: Stream::VERSION_0,
        system_identifier: 0,
        class_identifier: Guid::from_bytes([0; 16]),
        sections: vec![section],
    }
    .to_bytes()
    .expect("synthetic contents");

    let mut writer = OleWriter::new();
    writer
        .create_stream(&["CONTENTS"], &contents)
        .expect("CONTENTS stream");
    writer
        .create_stream(&[name], payload)
        .expect("indirect payload");
    if empty_storage {
        writer.create_storage(&["Empty"]).expect("empty storage");
        writer
            .set_storage_clsid(&["Empty"], [0xA5; 16])
            .expect("empty storage CLSID");
        writer
            .set_storage_metadata(&["Empty"], 7, 11, 13)
            .expect("empty storage metadata");
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("synthetic CFB");
    output.into_inner()
}

fn snapshot_with_property(name: &str, value: Value) -> Snapshot {
    Snapshot::open(
        source_with_property(name, value, b"payload", true),
        Limits::default(),
    )
    .expect("synthetic non-simple snapshot")
}

fn contents_without_indirect_reference(source: &Snapshot) -> Stream {
    let mut contents = source.contents().expect("CONTENTS");
    contents.sections[0].replace(2, Value::I4(7));
    contents
}

#[test]
fn ordinary_open_is_lazy_and_noop_shares_exact_source() {
    let name = IndirectPropertyName::new(2).expect("prop2");
    let source = Arc::<[u8]>::from(source_with_property(
        "prop2",
        Value::Stream(name),
        b"payload",
        true,
    ));
    let snapshot = Snapshot::open_shared(Arc::clone(&source), Limits::default()).expect("snapshot");
    assert_eq!(snapshot.source_shared().as_ref(), source.as_ref());
    assert_eq!(snapshot.elements().expect("elements").len(), 3);

    let commit = snapshot.edit().commit().expect("exact no-op");
    assert!(!commit.changed());
    assert!(Arc::ptr_eq(&commit.snapshot().source_shared(), &source));
}

#[test]
fn physical_paths_use_cfb_unicode_name_matching_and_reject_reserved_characters() {
    let lazy_source = Snapshot::open(
        source_with_property(
            "é",
            Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
            b"payload",
            false,
        ),
        Limits::default(),
    )
    .expect("Unicode-name source");
    assert_eq!(
        lazy_source
            .stream(&["É"])
            .expect("Unicode-name lookup")
            .expect("Unicode stream")
            .as_ref(),
        b"payload"
    );

    let source = snapshot_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
    );
    for invalid_name in ["bad\0name", "bad/name", "bad\\name", "bad:name", "bad!name"] {
        let mut editor = source.edit();
        assert!(editor.set_stream(&[invalid_name], Vec::new()).is_err());
        assert!(!editor.is_changed());
    }
}

#[test]
fn owned_input_limit_is_checked_before_arc_conversion() {
    let result = Snapshot::open(
        vec![0; 32],
        Limits {
            max_input_bytes: 1,
            ..Limits::default()
        },
    );
    assert!(matches!(
        result,
        Err(litchi_cfb::OleError::LimitExceeded {
            resource: "non-simple Property Set input bytes",
            observed: 32,
            maximum: 1,
        })
    ));
}

#[test]
fn copied_path_budget_rejects_a_wide_directory_before_catalog_retention() {
    let source = source_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
        b"payload",
        true,
    );
    let result = Snapshot::open(
        source,
        Limits {
            max_total_path_bytes: 1,
            ..Limits::default()
        },
    );
    assert!(matches!(
        result,
        Err(litchi_cfb::OleError::LimitExceeded {
            resource: "non-simple Property Set copied path bytes",
            ..
        })
    ));
}

#[test]
fn output_limit_accepts_exact_rendered_size_and_rejects_one_byte_less() {
    let source = snapshot_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
    );
    let mut baseline_editor = source.edit();
    baseline_editor
        .set_stream(&["prop2"], b"changed".to_vec())
        .expect("baseline edit");
    let baseline = baseline_editor.commit().expect("baseline commit");
    let exact_output = baseline.snapshot().source_shared().len() as u64;

    let exact_source = Snapshot::open(
        source.source_shared().as_ref().to_vec(),
        Limits {
            max_output_bytes: exact_output,
            ..Limits::default()
        },
    )
    .expect("exact-limit source");
    let mut exact_editor = exact_source.edit();
    exact_editor
        .set_stream(&["prop2"], b"changed".to_vec())
        .expect("exact-limit edit");
    let exact_commit = exact_editor.commit().expect("exact-limit commit");
    assert_eq!(
        exact_commit.snapshot().source_shared().len() as u64,
        exact_output
    );

    let under_source = Snapshot::open(
        source.source_shared().as_ref().to_vec(),
        Limits {
            max_output_bytes: exact_output - 1,
            ..Limits::default()
        },
    )
    .expect("under-limit source");
    let mut under_editor = under_source.edit();
    under_editor
        .set_stream(&["prop2"], b"changed".to_vec())
        .expect("under-limit edit");
    assert!(matches!(
        under_editor.commit(),
        Err(litchi_cfb::OleError::LimitExceeded {
            resource: "non-simple Property Set output bytes",
            ..
        })
    ));
}

#[test]
fn output_layout_preflight_covers_empty_mini_and_regular_stream_boundaries() {
    for length in [0usize, 4095, 4096] {
        let payload: Vec<u8> = (0..length)
            .map(|index| (index as u8).wrapping_add(3))
            .collect();
        let source_bytes = source_with_property(
            "prop2",
            Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
            &payload,
            true,
        );
        let source = Snapshot::open(source_bytes.clone(), Limits::default()).expect("source");

        // Replacing the typed CONTENTS value with an inert scalar forces a
        // physical publication even for the zero-length stream case, while
        // retaining the boundary-sized child stream in the output topology.
        let mut baseline_editor = source.edit();
        baseline_editor
            .replace_contents(contents_without_indirect_reference(&source))
            .expect("changed CONTENTS");
        let baseline = baseline_editor.commit().expect("baseline publication");
        let emitted = baseline.snapshot().source_shared();
        let exact_output = emitted.len() as u64;
        assert!(exact_output > 1);

        let reopened = Snapshot::open_shared(Arc::clone(&emitted), Limits::default())
            .expect("reopen emitted CFB");
        let element = reopened
            .element(&["prop2"])
            .expect("prop2 element")
            .expect("prop2 present");
        assert_eq!(element.size(), length as u64);
        assert_eq!(
            reopened
                .stream(&["prop2"])
                .expect("prop2 stream")
                .expect("prop2 bytes")
                .as_ref(),
            payload.as_slice()
        );

        let exact_source = Snapshot::open(
            source_bytes.clone(),
            Limits {
                max_output_bytes: exact_output,
                ..Limits::default()
            },
        )
        .expect("exact-cap source");
        let mut exact_editor = exact_source.edit();
        exact_editor
            .replace_contents(contents_without_indirect_reference(&exact_source))
            .expect("exact-cap CONTENTS");
        let exact = exact_editor.commit().expect("exact-cap publication");
        assert_eq!(exact.snapshot().source_shared().len() as u64, exact_output);
        let exact_reopened =
            Snapshot::open_shared(exact.snapshot().source_shared(), Limits::default())
                .expect("reopen exact-cap CFB");
        assert_eq!(
            exact_reopened
                .element(&["prop2"])
                .expect("exact prop2 element")
                .expect("exact prop2 present")
                .size(),
            length as u64
        );

        let under_source = Snapshot::open(
            source_bytes.as_slice().to_vec(),
            Limits {
                max_output_bytes: exact_output - 1,
                ..Limits::default()
            },
        )
        .expect("under-cap source");
        let mut under_editor = under_source.edit();
        under_editor
            .replace_contents(contents_without_indirect_reference(&under_source))
            .expect("under-cap CONTENTS");
        let staged_contents = under_editor
            .contents()
            .expect("staged under-cap contents")
            .to_bytes()
            .expect("staged under-cap bytes");
        let result = under_editor.snapshot();
        assert!(matches!(
            result,
            Err(litchi_cfb::OleError::LimitExceeded {
                resource: "non-simple Property Set output bytes",
                ..
            })
        ));
        assert!(under_editor.is_changed());
        assert_eq!(
            under_editor
                .contents()
                .expect("contents after refusal")
                .to_bytes()
                .expect("contents after refusal bytes"),
            staged_contents
        );
        assert_eq!(
            under_editor.source().source_shared().as_ref(),
            source_bytes.as_slice()
        );
    }
}

#[test]
fn rejected_physical_edits_do_not_materialize_or_disturb_a_staged_edit() {
    let source = snapshot_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
    );

    let mut fresh = source.edit();
    assert!(fresh.add_storage(&["prop2"], None).is_err());
    assert!(!fresh.is_changed());
    assert!(fresh.remove_element(&["missing"]).is_err());
    assert!(!fresh.is_changed());
    assert!(fresh.set_stream(&["Empty"], b"bad".to_vec()).is_err());
    assert!(!fresh.is_changed());

    let contents_size = source
        .element(&["CONTENTS"])
        .expect("CONTENTS element")
        .expect("CONTENTS present")
        .size();
    let limited = Snapshot::open(
        source.source_shared().as_ref().to_vec(),
        Limits {
            max_stream_bytes: contents_size,
            ..Limits::default()
        },
    )
    .expect("limited source");
    let mut limited_editor = limited.edit();
    let too_large = vec![0; usize::try_from(contents_size).unwrap() + 1];
    assert!(limited_editor.set_stream(&["prop2"], too_large).is_err());
    assert!(!limited_editor.is_changed());

    let mut staged = source.edit();
    staged
        .set_stream(&["prop2"], b"changed".to_vec())
        .expect("valid staged edit");
    let staged_contents = staged
        .contents()
        .expect("staged contents")
        .to_bytes()
        .unwrap();
    assert!(staged.add_storage(&["prop2"], None).is_err());
    assert!(staged.remove_element(&["missing"]).is_err());
    assert!(staged.set_stream(&["Empty"], b"bad".to_vec()).is_err());
    assert!(staged.is_changed());
    assert_eq!(
        staged
            .contents()
            .expect("contents after refusals")
            .to_bytes()
            .unwrap(),
        staged_contents
    );
}

#[test]
fn typed_indirect_values_bind_case_insensitive_and_preserve_selector_spelling() {
    let name = IndirectPropertyName::from_wire("PROP02".into(), Some(2))
        .expect("case-insensitive decimal prop selector");
    let snapshot = snapshot_with_property("prop02", Value::Stream(name));
    let contents = snapshot.contents().expect("typed contents");
    let value = contents.sections[0].property(2).expect("property");
    let Value::Stream(name) = value else {
        panic!("expected VT_STREAM");
    };
    assert_eq!(name.as_str(), "PROP02");
    assert_eq!(
        snapshot
            .stream(&["PROP02"])
            .expect("stream")
            .expect("present")
            .as_ref(),
        b"payload"
    );
}

#[test]
fn generic_property_parser_keeps_indirect_type_opaque() {
    let name = IndirectPropertyName::new(2).expect("prop2");
    let format_id = Guid::from_bytes([0x22; 16]);
    let mut section = Section::new(format_id);
    section.add(2, Value::Stream(name)).expect("property");
    let bytes = Stream::new(section).to_bytes().expect("wire");
    let parsed = Stream::parse(&bytes).expect("generic parser");
    assert!(matches!(
        parsed.sections[0].property(2),
        Some(Value::Unknown {
            variant_type: VT_STREAM,
            ..
        })
    ));
}

#[test]
fn storage_variant_keeps_empty_storage_and_reversible_source() {
    let name = IndirectPropertyName::new(2).expect("prop2");
    let format_id = Guid::from_bytes([0x33; 16]);
    let mut section = Section::new(format_id);
    section
        .add(2, Value::Storage(name))
        .expect("storage property");
    let contents = Stream {
        version: Stream::VERSION_0,
        system_identifier: 0,
        class_identifier: Guid::from_bytes([0; 16]),
        sections: vec![section],
    }
    .to_bytes()
    .expect("contents");
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["CONTENTS"], &contents)
        .expect("contents");
    writer.create_storage(&["prop2"]).expect("prop2 storage");
    writer.create_storage(&["Empty"]).expect("empty storage");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB");
    let snapshot = Snapshot::open(output.into_inner(), Limits::default()).expect("snapshot");

    let mut editor = snapshot.edit();
    editor
        .set_property(
            format_id,
            2,
            Value::Storage(IndirectPropertyName::new(2).unwrap()),
        )
        .expect("same typed property");
    let commit = editor.commit().expect("semantic no-op");
    assert!(!commit.changed());
}

#[test]
fn every_indirect_variant_checks_its_cfb_element_kind_and_versioned_tail() {
    let format_id = Guid::from_bytes([0x34; 16]);
    let version_guid = Guid::from_bytes([0x35; 16]);
    let mut section = Section::new(format_id);
    section
        .add(
            2,
            Value::VersionedStream(
                VersionedStream::new(version_guid, 2).expect("versioned stream"),
            ),
        )
        .expect("versioned stream property");
    section
        .add(
            3,
            Value::StreamedObject(
                IndirectPropertyName::from_wire("PROP03".into(), Some(3)).unwrap(),
            ),
        )
        .expect("streamed object property");
    section
        .add(
            4,
            Value::StoredObject(IndirectPropertyName::new(4).unwrap()),
        )
        .expect("stored object property");
    let contents = Stream {
        version: Stream::VERSION_0,
        system_identifier: 0,
        class_identifier: Guid::from_bytes([0; 16]),
        sections: vec![section],
    }
    .to_bytes()
    .expect("contents");
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["CONTENTS"], &contents)
        .expect("contents stream");
    writer
        .create_stream(&["prop2"], b"versioned")
        .expect("versioned stream");
    writer
        .create_stream(&["prop03"], b"streamed object")
        .expect("streamed object");
    writer.create_storage(&["prop4"]).expect("stored object");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB");
    let snapshot = Snapshot::open(output.into_inner(), Limits::default()).expect("snapshot");
    let parsed = snapshot.contents().expect("typed contents");
    assert!(matches!(
        parsed.sections[0].property(2),
        Some(Value::VersionedStream(_))
    ));
    assert!(matches!(
        parsed.sections[0].property(3),
        Some(Value::StreamedObject(_))
    ));
    assert!(matches!(
        parsed.sections[0].property(4),
        Some(Value::StoredObject(_))
    ));

    let mut wrong_kind = Section::new(format_id);
    wrong_kind
        .add(
            2,
            Value::StoredObject(IndirectPropertyName::new(2).unwrap()),
        )
        .expect("stored object property");
    let wrong_contents = Stream {
        version: Stream::VERSION_0,
        system_identifier: 0,
        class_identifier: Guid::from_bytes([0; 16]),
        sections: vec![wrong_kind],
    }
    .to_bytes()
    .expect("wrong contents");
    let mut wrong_writer = OleWriter::new();
    wrong_writer
        .create_stream(&["CONTENTS"], &wrong_contents)
        .expect("contents stream");
    wrong_writer
        .create_stream(&["prop2"], b"stream")
        .expect("stream");
    let mut wrong_output = Cursor::new(Vec::new());
    wrong_writer.write_to(&mut wrong_output).expect("CFB");
    let wrong_snapshot =
        Snapshot::open(wrong_output.into_inner(), Limits::default()).expect("snapshot");
    assert!(wrong_snapshot.contents().is_err());
}

#[test]
fn changed_stream_preserves_empty_storage_and_inverse() {
    let name = IndirectPropertyName::new(2).expect("prop2");
    let snapshot = snapshot_with_property("prop2", Value::Stream(name));
    let source = snapshot.source_shared();
    let mut editor = snapshot.edit();
    editor
        .set_stream(&["prop2"], b"changed".to_vec())
        .expect("stream edit");
    let commit = editor.commit().expect("commit");
    assert!(commit.changed());
    assert_eq!(
        commit
            .snapshot()
            .stream(&["prop2"])
            .expect("stream")
            .expect("present")
            .as_ref(),
        b"changed"
    );
    assert!(commit.snapshot().contents().is_ok());
    let empty = commit
        .snapshot()
        .element(&["Empty"])
        .expect("element")
        .expect("present");
    let metadata = empty.metadata().expect("storage metadata");
    assert_eq!(metadata.class_id(), Some(Guid::from_bytes([0xA5; 16])));
    assert_eq!(metadata.state_bits(), 7);
    assert_eq!(metadata.creation_time(), 11);
    assert_eq!(metadata.modified_time(), 13);
    let reverted = commit.patch().revert(commit.snapshot()).expect("inverse");
    assert_eq!(reverted.source_shared().as_ref(), source.as_ref());
    assert!(Arc::ptr_eq(&reverted.source_shared(), &source));
}

#[test]
fn invalid_indirect_property_identifier_is_rejected_before_serialization() {
    let format_id = Guid::from_bytes([0x44; 16]);
    let mut section = Section::new(format_id);
    let invalid_name = IndirectPropertyName::from_wire("prop3".into(), None)
        .expect("wire selector itself is valid");
    section.replace(2, Value::Stream(invalid_name));
    assert!(Stream::new(section).to_bytes().is_err());
}

#[test]
fn every_new_indirect_variant_rejects_a_replaced_property_id_mismatch() {
    let format_id = Guid::from_bytes([0x45; 16]);
    let values = [
        Value::Stream(IndirectPropertyName::from_wire("prop3".into(), None).unwrap()),
        Value::Storage(IndirectPropertyName::from_wire("prop3".into(), None).unwrap()),
        Value::StreamedObject(IndirectPropertyName::from_wire("prop3".into(), None).unwrap()),
        Value::StoredObject(IndirectPropertyName::from_wire("prop3".into(), None).unwrap()),
    ];
    for value in values {
        let mut section = Section::new(format_id);
        section.replace(2, value);
        assert!(Stream::new(section).to_bytes().is_err());
    }
}

#[test]
fn root_clsid_mismatch_is_checked_when_contents_is_requested() {
    let format_id = Guid::from_bytes([0x55; 16]);
    let mut section = Section::new(format_id);
    section
        .add(2, Value::Stream(IndirectPropertyName::new(2).unwrap()))
        .expect("property");
    let contents = Stream {
        version: Stream::VERSION_0,
        system_identifier: 0,
        class_identifier: Guid::from_bytes([0xAA; 16]),
        sections: vec![section],
    }
    .to_bytes()
    .expect("contents");
    let mut writer = OleWriter::new();
    writer.set_root_clsid([0xBB; 16]);
    writer
        .create_stream(&["CONTENTS"], &contents)
        .expect("contents");
    writer
        .create_stream(&["prop2"], b"payload")
        .expect("payload");
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("CFB");
    let snapshot = Snapshot::open(output.into_inner(), Limits::default()).expect("capture");
    assert!(snapshot.contents().is_err());
}

#[test]
fn limits_reject_oversized_indirect_stream_before_render() {
    let name = IndirectPropertyName::new(2).unwrap();
    let bytes = source_with_property("prop2", Value::Stream(name), b"payload", false);
    let limits = Limits {
        max_stream_bytes: 3,
        ..Limits::default()
    };
    assert!(Snapshot::open(bytes, limits).is_err());
}

#[test]
fn physical_and_typed_crud_commit_atomically() {
    let source = snapshot_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).unwrap()),
    );
    let format_id = Guid::from_bytes([0x11; 16]);
    let mut editor = source.edit();
    editor.add_storage(&["prop3"], None).expect("new storage");
    editor
        .set_storage_property(format_id, 3)
        .expect("new storage property");
    let commit = editor.commit().expect("storage commit");
    assert_eq!(
        commit
            .snapshot()
            .element(&["prop3"])
            .expect("element")
            .expect("present")
            .kind(),
        Some(super::ElementKind::Storage)
    );
    let mut editor = commit.snapshot().edit();
    let removed = editor
        .remove_property(format_id, 3)
        .expect("remove property")
        .expect("removed");
    assert!(matches!(removed, Value::Storage(_)));
    editor.remove_element(&["prop3"]).expect("remove storage");
    let reverted = editor.commit().expect("remove commit");
    assert!(
        reverted
            .snapshot()
            .element(&["prop3"])
            .expect("element")
            .is_none()
    );
}

fn generic_unknown_data_for(value: Value) -> Vec<u8> {
    let format_id = Guid::from_bytes([0x66; 16]);
    let mut section = Section::new(format_id);
    section.add(3, value).expect("indirect property");
    let bytes = Stream::new(section).to_bytes().expect("typed wire");
    let parsed = Stream::parse(&bytes).expect("generic wire parse");
    let Value::Unknown { data, .. } = parsed.sections[0]
        .property(3)
        .expect("generic unknown property")
    else {
        panic!("generic parser unexpectedly typed the indirect value");
    };
    data.clone()
}

#[test]
fn owner_closure_validates_opaque_indirect_variant_tags() {
    let source = snapshot_with_property(
        "prop2",
        Value::Stream(IndirectPropertyName::new(2).expect("prop2")),
    );
    let version_guid = Guid::from_bytes([0x68; 16]);
    let variants = [
        (
            VT_STREAM,
            generic_unknown_data_for(Value::Stream(
                IndirectPropertyName::from_wire("prop3".into(), None).unwrap(),
            )),
        ),
        (
            VT_STORAGE,
            generic_unknown_data_for(Value::Storage(
                IndirectPropertyName::from_wire("prop3".into(), None).unwrap(),
            )),
        ),
        (
            VT_STREAMED_OBJECT,
            generic_unknown_data_for(Value::StreamedObject(
                IndirectPropertyName::from_wire("prop3".into(), None).unwrap(),
            )),
        ),
        (
            VT_STORED_OBJECT,
            generic_unknown_data_for(Value::StoredObject(
                IndirectPropertyName::from_wire("prop3".into(), None).unwrap(),
            )),
        ),
        (
            VT_VERSIONED_STREAM,
            generic_unknown_data_for(Value::VersionedStream(
                VersionedStream::new(version_guid, 3).expect("prop3 versioned stream"),
            )),
        ),
    ];

    for (variant_type, data) in variants {
        let mut candidate = source.contents().expect("contents");
        candidate.sections[0].replace(2, Value::Unknown { variant_type, data });
        assert!(source.edit().replace_contents(candidate).is_err());
    }

    for variant_type in [
        VT_STREAM,
        VT_STORAGE,
        VT_STREAMED_OBJECT,
        VT_STORED_OBJECT,
        VT_VERSIONED_STREAM,
    ] {
        let mut candidate = source.contents().expect("contents");
        candidate.sections[0].replace(
            2,
            Value::Unknown {
                variant_type,
                data: Vec::new(),
            },
        );
        assert!(source.edit().replace_contents(candidate).is_err());
    }
}
