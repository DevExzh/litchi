#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{
    EditOptions, Error, Flags, Kind, LimitKind, Limits, OpenOptions, Package, RefusalKind, View,
};
use crate::prop::{Id, Prop, Props};
use crate::{Record, RecordKind};

fn opt_record(properties: &[(u16, i32, Option<&[u8]>)]) -> Vec<u8> {
    let headers = properties.len() * 6;
    let complex = properties
        .iter()
        .filter_map(|(_, _, data)| *data)
        .map(<[u8]>::len)
        .sum::<usize>();
    let mut body = Vec::with_capacity(headers + complex);
    for (opid, value, _) in properties {
        body.extend_from_slice(&opid.to_le_bytes());
        body.extend_from_slice(&value.to_le_bytes());
    }
    for (_, _, data) in properties {
        if let Some(complex_data) = data {
            body.extend_from_slice(complex_data);
        }
    }
    let mut record = Vec::with_capacity(body.len() + 8);
    let property_count = u16::try_from(properties.len()).expect("property count");
    let version_instance = (property_count << 4) | 3;
    record.extend_from_slice(&version_instance.to_le_bytes());
    record.extend_from_slice(&RecordKind::Opt.raw().to_le_bytes());
    let body_len = u32::try_from(body.len()).expect("body length");
    record.extend_from_slice(&body_len.to_le_bytes());
    record.extend_from_slice(&body);
    record
}

fn name_bytes(text: &str) -> Vec<u8> {
    let mut bytes = text
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    bytes.extend_from_slice(&0_u16.to_le_bytes());
    bytes
}

#[test]
fn decodes_picture_name_and_retains_reserved_flags() {
    let name = name_bytes("old");
    let bytes = opt_record(&[
        (
            0x8000 | Id::PictureFileName.raw(),
            i32::try_from(name.len()).expect("name length"),
            Some(&name),
        ),
        (Id::BlipFlags.raw(), 0x8000_0009_u32.cast_signed(), None),
    ]);
    let (record, consumed) =
        Record::parse(&bytes, 0, crate::Limits::default()).expect("picture Opt");
    assert_eq!(consumed, bytes.len());

    let props = Props::parse(&record, crate::Limits::default()).expect("properties");
    let metadata = props
        .picture(Limits::default())
        .expect("picture metadata")
        .expect("metadata present");
    let decoded_name = metadata.name().expect("name");
    assert_eq!(decoded_name.text().expect("UTF-16"), "old");
    assert_eq!(decoded_name.raw_bytes(), name_bytes("old").as_slice());
    assert_eq!(metadata.flags().kind(), Kind::File);
    assert!(metadata.flags().link_to_file());
    assert_eq!(metadata.flags().reserved(), 0x8000_0000);

    let source = Package::open(&bytes).expect("package");
    assert_eq!(source.to_bytes(), bytes);
    assert_eq!(
        View::parse(&record, Limits::default())
            .expect("raw view")
            .record()
            .data(),
        &bytes[8..]
    );
}

#[test]
fn edits_only_modeled_picture_values_and_preserves_opaque_neighbors() {
    let old_name = name_bytes("old");
    let opaque = [0xDE, 0xAD, 0xBE, 0xEF];
    let bytes = opt_record(&[
        (0x0600, -7, None),
        (
            0x8000 | Id::PictureFileName.raw(),
            i32::try_from(old_name.len()).expect("name length"),
            Some(&old_name),
        ),
        (Id::BlipFlags.raw(), 0x8000_0009_u32.cast_signed(), None),
        (
            0x8000 | 0x0601,
            i32::try_from(opaque.len()).expect("opaque length"),
            Some(&opaque),
        ),
    ]);
    let source = Package::open(&bytes).expect("package");

    let mut edit = source.edit();
    edit.set_name("new").expect("new name");
    edit.set_flags(Flags::new(Kind::Url, true, false).expect("URL flags"));
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_noop());
    assert_eq!(
        commit.patch().name().expect("name patch").before().copied(),
        Some(old_name.as_ref())
    );
    assert_eq!(
        commit.patch().name().expect("name patch").after().copied(),
        Some(name_bytes("new").as_ref())
    );
    assert_eq!(
        commit.patch().flags().expect("flags patch").before(),
        Some(&(0x8000_0009_u32))
    );
    assert_eq!(
        commit.patch().flags().expect("flags patch").after(),
        Some(&0x0A)
    );

    let edited = commit.package().view().expect("edited view");
    assert!(commit.diagnostics().changed());
    let metadata = edited
        .metadata()
        .expect("metadata")
        .expect("metadata present");
    assert_eq!(metadata.name().expect("name").text().expect("text"), "new");
    assert_eq!(metadata.flags().raw(), 0x0A);
    assert_eq!(
        edited
            .properties()
            .get_int(Id::unknown(0x0600).expect("unknown")),
        Some(-7)
    );
    assert_eq!(
        edited
            .properties()
            .get_binary(Id::unknown(0x0601).expect("unknown")),
        Some(&opaque[..])
    );
    assert_eq!(
        edited
            .properties()
            .iter()
            .map(Prop::raw_id)
            .collect::<Vec<_>>(),
        vec![
            0x0600,
            Id::PictureFileName.raw(),
            Id::BlipFlags.raw(),
            0x0601
        ]
    );

    let source_metadata = source
        .metadata()
        .expect("source metadata")
        .expect("source present");
    assert_eq!(
        source_metadata
            .name()
            .expect("source name")
            .text()
            .expect("text"),
        "old"
    );
    assert_eq!(source.to_bytes(), bytes);
}

#[test]
fn validates_utf16_bounds_and_flag_dependencies() {
    assert!(matches!(
        Flags::from_raw(3),
        Err(Error::Format(crate::Error::MalformedProperties {
            reason: "picture name kind uses a reserved flag value"
        }))
    ));
    assert!(matches!(
        Flags::from_raw(4),
        Err(Error::Format(crate::Error::MalformedProperties {
            reason: "picture do-not-save requires link-to-file"
        }))
    ));
    assert!(Flags::new(Kind::Url, true, true).is_ok());
    assert!(Flags::new(Kind::Comment, true, false).is_err());

    let odd = [0x41, 0x00, 0x00];
    let bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        i32::try_from(odd.len()).expect("name length"),
        Some(&odd),
    )]);
    let (record, _) = Record::parse(&bytes, 0, crate::Limits::default()).expect("picture Opt");
    assert!(matches!(
        View::parse(&record, Limits::default()),
        Err(Error::Format(crate::Error::MalformedProperties {
            reason: "picture name must be a bounded even-length UTF-16 string"
        }))
    ));

    let interior_nul = [0x41, 0x00, 0x00, 0x00, 0x00, 0x00];
    let interior_bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        i32::try_from(interior_nul.len()).expect("name length"),
        Some(&interior_nul),
    )]);
    let (interior_record, _) =
        Record::parse(&interior_bytes, 0, crate::Limits::default()).expect("picture Opt");
    assert!(matches!(
        View::parse(&interior_record, Limits::default()),
        Err(Error::Format(crate::Error::MalformedProperties {
            reason: "picture name contains an interior NUL"
        }))
    ));

    let too_long = [0x41, 0x00].repeat(Limits::default().name_units_value() + 2);
    let too_long_bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        i32::try_from(too_long.len()).expect("name length"),
        Some(&too_long),
    )]);
    let (too_long_record, _) =
        Record::parse(&too_long_bytes, 0, crate::Limits::default()).expect("picture Opt");
    assert!(matches!(
        View::parse(&too_long_record, Limits::default()),
        Err(Error::Limit {
            kind: LimitKind::NameUnits,
            observed: 4097,
            maximum: 4096
        })
    ));
}

#[test]
fn appends_and_clears_names_without_touching_other_properties() {
    let opaque = [1, 2, 3];
    let bytes = opt_record(&[(
        0x8000 | 0x0601,
        i32::try_from(opaque.len()).expect("opaque length"),
        Some(&opaque),
    )]);
    let source = Package::open(&bytes).expect("package");

    let mut add = source.edit();
    add.set_name("added").expect("name");
    let added = add.commit().expect("add name");
    let added_view = added.package().view().expect("added view");
    assert_eq!(
        added_view
            .metadata()
            .expect("metadata")
            .expect("present")
            .name()
            .expect("name")
            .text()
            .expect("text"),
        "added"
    );
    assert_eq!(
        added_view
            .properties()
            .get_binary(Id::unknown(0x0601).expect("unknown")),
        Some(&opaque[..])
    );

    let mut clear = added.package().edit();
    clear.clear_name();
    let cleared = clear.commit().expect("clear name");
    assert!(cleared.package().metadata().expect("metadata").is_none());
}

#[test]
fn replay_matches_exact_content_and_inverse_restores_absent_descriptors() {
    let bytes = opt_record(&[(0x0600, -7, None)]);
    let source = Package::open(&bytes).expect("source");
    let mut edit = source.edit();
    edit.set_name("added").expect("name");
    let committed = edit.commit().expect("commit");
    assert!(committed.patch().flags().is_none());
    let copy = bytes.clone();
    let independent = Package::open(&copy).expect("independent base");
    let replay = independent.apply(committed.patch()).expect("matching copy");
    assert_eq!(replay.package().to_bytes(), committed.package().to_bytes());
    let restored = replay
        .package()
        .apply(&replay.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.package().to_bytes(), bytes);
    assert_eq!(
        restored
            .package()
            .view()
            .expect("restored view")
            .properties()
            .len(),
        1
    );
    assert!(matches!(
        replay.package().apply(committed.patch()),
        Err(Error::PatchConflict)
    ));

    let different = opt_record(&[(0x0600, -8, None)]);
    assert!(matches!(
        Package::open(&different)
            .expect("other")
            .apply(committed.patch()),
        Err(Error::PatchConflict)
    ));
    let mut different_kind = bytes.clone();
    different_kind[2..4].copy_from_slice(&RecordKind::SecondaryOpt.raw().to_le_bytes());
    assert!(matches!(
        Package::open(&different_kind)
            .expect("other kind")
            .apply(committed.patch()),
        Err(Error::PatchConflict)
    ));

    let mut flags = source.edit();
    flags.set_flags(Flags::from_raw(0).expect("explicit zero"));
    let committed = flags.commit().expect("flags");
    assert!(committed.diagnostics().changed());
    assert_eq!(
        committed.patch().flags().expect("presence change").before(),
        None
    );
    assert_eq!(
        committed.patch().flags().expect("presence change").after(),
        Some(&0)
    );
    let restored = committed
        .package()
        .apply(&committed.patch().inverse())
        .expect("inverse flags");
    assert_eq!(restored.package().to_bytes(), bytes);
}

#[test]
fn no_op_drop_and_chained_edits_keep_the_original_bytes() {
    let bytes = opt_record(&[]);
    let source = Package::open(&bytes).expect("source");
    let committed = source.edit().commit().expect("no-op");
    assert!(committed.patch().is_noop());
    assert!(!committed.diagnostics().changed());
    let copy = bytes.clone();
    assert!(
        !Package::open(&copy)
            .expect("copy")
            .apply(committed.patch())
            .expect("replay")
            .diagnostics()
            .changed()
    );
    {
        let mut discarded = source.edit();
        discarded.set_name("discarded").expect("name");
    }
    assert_eq!(source.to_bytes(), bytes);
    let committed = {
        let temporary_owner = source.clone();
        let mut edit = temporary_owner.edit();
        edit.set_name("one").expect("name");
        edit.commit().expect("commit outlives edit and owner")
    };
    let mut same = committed.package().edit();
    same.set_name("one").expect("same");
    assert!(!same.commit().expect("same name").diagnostics().changed());
    let mut next = committed.package().edit();
    next.set_name("two").expect("next");
    let next = next.commit().expect("next commit");
    let independent_bytes = committed.package().to_bytes();
    let independent = Package::open(&independent_bytes).expect("independent edited base");
    assert_eq!(
        independent
            .apply(next.patch())
            .expect("chained replay")
            .package()
            .to_bytes(),
        next.package().to_bytes()
    );
    assert_eq!(source.to_bytes(), bytes);
}

#[test]
fn exact_limits_cover_ingress_edits_and_recipient_policy() {
    let bytes = opt_record(&[]);
    let limits = Limits::new()
        .name_units(2)
        .expect("name cap")
        .properties(1)
        .expect("count cap")
        .property_bytes(12)
        .expect("body cap");
    let source =
        Package::open_with_options(&bytes, OpenOptions::new().limits(limits)).expect("source");
    let mut exact = source.edit();
    exact.set_name("😀").expect("two UTF-16 units");
    let exact = exact
        .commit()
        .expect("six-byte descriptor and six-byte name");
    let encoded = exact.package().to_bytes();
    assert!(Package::open_with_options(&encoded, OpenOptions::new().limits(limits)).is_ok());
    let mut too_long = source.edit_with_options(EditOptions::new().limits(Limits::new()));
    assert!(matches!(
        too_long.set_name("abc"),
        Err(Error::Limit {
            kind: LimitKind::NameUnits,
            observed: 3,
            maximum: 2
        })
    ));
    let mut too_many = exact.package().edit();
    too_many.set_flags(Flags::from_raw(0).expect("flags"));
    assert!(matches!(
        too_many.commit(),
        Err(Error::Limit {
            kind: LimitKind::Properties,
            observed: 2,
            maximum: 1
        })
    ));
    for (limited, kind, observed, maximum) in [
        (
            limits.name_units(1).expect("name cap"),
            LimitKind::NameUnits,
            2,
            1,
        ),
        (
            limits.properties(0).expect("count cap"),
            LimitKind::Properties,
            1,
            0,
        ),
        (
            limits.property_bytes(11).expect("byte cap"),
            LimitKind::PropertyBytes,
            12,
            11,
        ),
    ] {
        assert!(
            matches!(Package::open_with_options(&encoded, OpenOptions::new().limits(limited)), Err(Error::Limit { kind: actual, observed: got, maximum: cap }) if (actual, got, cap) == (kind, observed, maximum))
        );
        let recipient = Package::open_with_options(&bytes, OpenOptions::new().limits(limited))
            .expect("empty recipient");
        assert!(
            matches!(recipient.apply(exact.patch()), Err(Error::Limit { kind: actual, observed: got, maximum: cap }) if (actual, got, cap) == (kind, observed, maximum))
        );
    }
    let stricter = EditOptions::new().limits(limits.name_units(1).expect("stricter"));
    assert!(matches!(
        exact.package().edit_with_options(stricter).commit(),
        Err(Error::Limit {
            kind: LimitKind::NameUnits,
            observed: 2,
            maximum: 1
        })
    ));
    assert!(matches!(
        Limits::new().properties(4096),
        Err(Error::InvalidLimit {
            kind: LimitKind::Properties,
            value: 4096,
            maximum: 4095
        })
    ));
}

#[test]
fn no_op_commit_retains_tighter_edit_policy() {
    let bytes = opt_record(&[]);
    let source = Package::open(&bytes).expect("source");
    let committed = source
        .edit_with_options(EditOptions::new().limits(Limits::new().name_units(1).expect("cap")))
        .commit()
        .expect("no-op");
    assert!(committed.patch().is_noop());
    assert_eq!(committed.package().to_bytes(), bytes);
    assert!(matches!(
        committed.package().edit().set_name("ab"),
        Err(Error::Limit {
            kind: LimitKind::NameUnits,
            observed: 2,
            maximum: 1
        })
    ));
    source
        .edit()
        .set_name("ab")
        .expect("source policy remains intact");
}

#[test]
fn default_edit_inherits_a_larger_admitted_name_budget() {
    let bytes = opt_record(&[]);
    let source = Package::open_with_options(
        &bytes,
        OpenOptions::new().limits(Limits::new().name_units(5000).expect("larger cap")),
    )
    .expect("source");
    let mut edit = source.edit();
    edit.set_name(&"x".repeat(5000)).expect("inherited cap");
    let committed = edit.commit().expect("commit");
    assert_eq!(
        committed
            .package()
            .metadata()
            .expect("metadata")
            .expect("present")
            .name()
            .expect("name")
            .unit_len(),
        5000
    );
    assert!(
        committed
            .package()
            .edit()
            .commit()
            .expect("no-op")
            .patch()
            .is_noop()
    );
}

#[test]
fn validation_rejects_bad_unicode_without_disclosing_authored_text() {
    for raw in [vec![0x00, 0xd8, 0, 0], vec![0x41, 0x00], vec![]] {
        let bytes = opt_record(&[(
            0x8000 | Id::PictureFileName.raw(),
            i32::try_from(raw.len()).expect("length"),
            Some(&raw),
        )]);
        assert!(matches!(
            Package::open(&bytes),
            Err(Error::Format(crate::Error::MalformedProperties { .. }))
        ));
    }
    let bytes = opt_record(&[]);
    let source = Package::open(&bytes).expect("source");
    let mut edit = source.edit();
    let error = match edit.set_name("private text\0tail") {
        Ok(_) => panic!("interior NUL accepted"),
        Err(error) => error,
    };
    assert!(
        matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::InvalidValue)
    );
    assert!(!error.to_string().contains("private text"));
    assert!(
        !edit
            .commit()
            .expect("failed operation leaves edit unchanged")
            .diagnostics()
            .changed()
    );
}

#[test]
fn output_streams_exact_bytes_and_propagates_io_failures() {
    struct Failing;
    impl std::io::Write for Failing {
        fn write(&mut self, _: &[u8]) -> std::io::Result<usize> {
            Err(std::io::Error::from(std::io::ErrorKind::BrokenPipe))
        }
        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    fn send_sync<T: Send + Sync>() {}
    send_sync::<Package<'static>>();
    send_sync::<super::Patch<'static>>();
    let bytes = opt_record(&[]);
    let source = Package::open(&bytes).expect("source");
    let mut edit = source.edit();
    edit.set_name("output").expect("name");
    let committed = edit.commit().expect("commit");
    let mut output = Vec::new();
    committed.package().write_to(&mut output).expect("write");
    assert_eq!(output, committed.package().to_bytes());
    assert!(
        matches!(committed.package().write_to(Failing), Err(Error::Io(error)) if error.kind() == std::io::ErrorKind::BrokenPipe)
    );
}

#[test]
fn streaming_crosses_descriptor_and_payload_buffer_boundaries_exactly() {
    for length in [1023, 1024, 1025] {
        let opaque = vec![0xab; length];
        let mut properties: Vec<_> = (0..300_u16)
            .map(|index| (0x3000 + index, i32::from(index), None))
            .collect();
        properties.push((
            0xbffe,
            i32::try_from(length).expect("length"),
            Some(opaque.as_slice()),
        ));
        let bytes = opt_record(&properties);
        let source = Package::open(&bytes).expect("source");
        let mut edit = source.edit();
        edit.set_name("new").expect("name");
        let committed = edit.commit().expect("commit");
        let name = name_bytes("new");
        properties.push((
            0x8000 | Id::PictureFileName.raw(),
            i32::try_from(name.len()).expect("name length"),
            Some(&name),
        ));
        let expected = opt_record(&properties);
        let mut actual = Vec::new();
        committed.package().write_to(&mut actual).expect("stream");
        assert_eq!(actual, expected);
        assert_eq!(
            committed.package().view().expect("view").record().data(),
            &expected[8..]
        );
    }
}

#[test]
fn public_debug_and_io_messages_omit_authored_content() {
    let name = name_bytes("private name");
    let bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        i32::try_from(name.len()).expect("length"),
        Some(&name),
    )]);
    let source = Package::open(&bytes).expect("source");
    let mut edit = source.edit();
    edit.set_name("replacement secret").expect("name");
    let committed = edit.commit().expect("commit");
    let view = source.view().expect("view");
    let metadata = source.metadata().expect("metadata").expect("present");
    let io = Error::Io(std::io::Error::other("private writer message"));
    let messages = [
        format!("{source:?}"),
        format!("{view:?}"),
        format!("{:?}", view.properties()),
        format!("{metadata:?}"),
        format!("{:?}", metadata.name().expect("name")),
        format!("{committed:?}"),
        format!("{:?}", committed.patch().name().expect("name change")),
        format!("{io:?}"),
        io.to_string(),
    ];
    for message in messages {
        for secret in ["private", "replacement", "[112, 0, 114", "[114, 0, 101"] {
            assert!(!message.contains(secret), "{message}");
        }
    }
}

#[test]
fn detached_owner_and_patch_outlive_the_original_input() {
    fn fixture() -> Vec<u8> {
        let array = [1, 0, 1, 0, 4, 0, 1, 2, 3, 4];
        let opaque = [9, 8, 7];
        opt_record(&[
            (0x0600, -7, None),
            (0xc000 | Id::Vertices.raw(), 10, Some(&array)),
            (0xc601, 3, Some(&opaque)),
        ])
    }
    let committed: super::Commit<'static> = {
        let bytes = fixture();
        let source = Package::open(&bytes).expect("source").into_owned();
        let mut edit = source.edit();
        edit.set_name("detached").expect("name");
        edit.commit().expect("owned commit")
    };
    let expected = fixture();
    let view = committed.package().view().expect("detached view");
    assert_eq!(
        view.properties()
            .get_array(Id::Vertices)
            .expect("array")
            .raw_data(),
        [1, 0, 1, 0, 4, 0, 1, 2, 3, 4]
    );
    assert_eq!(
        committed
            .package()
            .apply(&committed.patch().inverse())
            .expect("owned inverse")
            .package()
            .to_bytes(),
        expected
    );

    let (patch, no_op): (super::Patch<'static>, super::Patch<'static>) = {
        let source = Package::open(&expected).expect("source");
        let no_op = source
            .edit()
            .commit()
            .expect("no-op")
            .patch()
            .clone()
            .into_owned();
        let mut edit = source.edit();
        edit.set_name("detached").expect("name");
        (
            edit.commit().expect("commit").patch().clone().into_owned(),
            no_op,
        )
    };
    assert!(no_op.is_noop());
    let independent = Package::open(&expected).expect("independent recipient");
    assert_eq!(
        independent
            .apply(&patch)
            .expect("detached replay")
            .package()
            .to_bytes(),
        committed.package().to_bytes()
    );
}

#[test]
fn chained_appends_and_clear_preserve_slot_order() {
    let bytes = opt_record(&[]);
    let source = Package::open(&bytes).expect("source");
    let mut flags = source.edit();
    flags.set_flags(Flags::from_raw(9).expect("flags"));
    let flags = flags.commit().expect("flags commit");
    let mut name = flags.package().edit();
    name.set_name("new").expect("name");
    let name = name.commit().expect("name commit");
    let expected = opt_record(&[
        (Id::BlipFlags.raw(), 9, None),
        (
            0x8000 | Id::PictureFileName.raw(),
            8,
            Some(&name_bytes("new")),
        ),
    ]);
    assert_eq!(name.package().to_bytes(), expected);
    let mut clear = name.package().edit();
    clear.clear_name();
    let clear = clear.commit().expect("clear commit");
    assert_eq!(
        clear.package().to_bytes(),
        opt_record(&[
            (Id::BlipFlags.raw(), 9, None),
            (Id::PictureFileName.raw(), 0, None)
        ])
    );
    assert_eq!(
        clear
            .package()
            .apply(&clear.patch().inverse())
            .expect("inverse")
            .package()
            .to_bytes(),
        expected
    );
    let mut restore = clear.package().edit();
    restore.set_name("new").expect("restore");
    let restore = restore.commit().expect("restored");
    assert_eq!(restore.package().to_bytes(), expected);
}
