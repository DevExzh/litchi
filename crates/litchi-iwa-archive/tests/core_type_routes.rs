use litchi_iwa_archive::iwa;
use litchi_iwa_archive::package::{PackageEntry, PackageState};
use litchi_iwa_core as iwa_core;

fn sample_archive() -> iwa::Archive {
    let path = iwa::FieldPath::new(vec![1, 2]);
    let mut field_info = iwa::FieldInfo::new(path);
    field_info.r#type = Some(iwa::FieldType::ObjectReference);
    field_info.unknown_field_rule = Some(iwa::UnknownFieldRule::IgnoreAndPreserve);
    field_info.object_references = vec![12];
    field_info.data_references = vec![22];

    let mut message_info = iwa::MessageInfo::new(99, 2);
    message_info.field_infos.push(field_info);
    message_info.object_references = vec![11];
    message_info.data_references = vec![21];
    let archive_info = iwa::ArchiveInfo::new(7, vec![message_info]);

    let mut object = iwa::ArchiveObject::new(
        7,
        vec![iwa::RawMessage {
            type_: 99,
            data: vec![0x08, 0x01],
        }],
    )
    .expect("sample archive object should be valid");
    object.archive_info = archive_info;
    iwa::Archive {
        objects: vec![object],
    }
}

#[test]
fn archive_owned_routes_preserve_core_type_identity() {
    fn accepts_archive(_: &iwa::Archive) {}
    fn accepts_core_archive(_: &iwa_core::Archive) {}
    fn accepts_core_archive_limits(_: iwa_core::ArchiveLimits) {}
    fn accepts_core_error(_: iwa_core::Error) {}
    fn accepts_core_limit_kind(_: iwa_core::LimitKind) {}
    fn accepts_core_result(_: iwa_core::Result<iwa_core::Archive>) {}
    fn accepts_core_snappy_limits(_: iwa_core::SnappyLimits) {}

    let state = PackageState::from_entries(
        vec![PackageEntry::new(
            "Index/Document.iwa".to_owned(),
            Vec::new(),
        )],
        iwa::ArchiveLimits::default(),
    )
    .expect("package state should accept one entry");
    let source = sample_archive();
    let parsed = state
        .get_or_parse_archive("Index/Document.iwa", |_| Ok((source.clone(), 1)))
        .expect("archive cache should accept the routed archive type");

    // `PackageState` is implemented against litchi-iwa-core::Archive. This
    // assignment is therefore a compile-time identity check for the hidden
    // archive-owned route, rather than a conversion check.
    accepts_archive(parsed.as_ref());
    accepts_core_archive(parsed.as_ref());

    let routed_archive_limits = iwa::ArchiveLimits::default();
    accepts_core_archive_limits(routed_archive_limits);
    let routed_result: iwa::Result<iwa::Archive> = Ok(source);
    accepts_core_result(routed_result);

    let routed_error = iwa::Error::InvalidArchive {
        offset: 0,
        reason: "route identity",
    };
    accepts_core_error(routed_error);
    accepts_core_limit_kind(iwa::LimitKind::ArchiveBytes);
    accepts_core_snappy_limits(iwa::SnappyLimits::default());
}

#[test]
fn routed_archive_and_snappy_values_round_trip_with_references() -> iwa::Result<()> {
    fn accepts_core_field_transition(_: iwa_core::archive::FieldObjectReferenceTransition<'_>) {}
    fn accepts_core_object_transition(_: iwa_core::archive::ObjectReferenceTransition<'_>) {}
    fn accepts_core_occurrence(_: iwa_core::ArchiveReferenceOccurrence) {}
    fn accepts_core_policy(_: iwa_core::ArchiveReferencePolicy) {}
    fn accepts_core_visitor<T: iwa_core::ArchiveReferenceVisitor>(_: &mut T) {}

    let archive = sample_archive();
    let decompressed = archive.to_bytes()?;
    let compressed = iwa::SnappyStream::compress(&decompressed)?;
    let stream =
        iwa::SnappyStream::decompress_with_limits(&compressed, iwa::SnappyLimits::default())?;
    assert_eq!(stream.as_bytes(), decompressed.as_slice());

    let reparsed =
        iwa::Archive::parse_with_limits(stream.as_bytes(), iwa::ArchiveLimits::default())?;
    assert_eq!(
        reparsed.objects[0].archive_info,
        archive.objects[0].archive_info
    );
    assert_eq!(reparsed.objects[0].messages, archive.objects[0].messages);

    #[derive(Default)]
    struct References(Vec<iwa::ArchiveReferenceOccurrence>);

    impl iwa::ArchiveReferenceVisitor for References {
        fn visit_reference(
            &mut self,
            occurrence: iwa::ArchiveReferenceOccurrence,
        ) -> iwa::Result<()> {
            self.0.push(occurrence);
            Ok(())
        }
    }

    let mut references = References::default();
    accepts_core_visitor(&mut references);
    let policy = iwa::ArchiveReferencePolicy::KnownReferences;
    accepts_core_policy(policy);
    let count = reparsed.objects[0].inspect_references_with_policy_and_limits(
        &mut references,
        policy,
        iwa::ArchiveLimits::default(),
    )?;
    assert_eq!(count, 4);
    assert_eq!(references.0.len(), count);
    accepts_core_occurrence(references.0[0]);
    let expected_path = [1, 2];
    let before = [12];
    let after = [12, 13];
    let field_transition = iwa::FieldObjectReferenceTransition {
        field_info_index: 0,
        expected_path: &expected_path,
        before: &before,
        after: &after,
    };
    accepts_core_field_transition(field_transition);
    let fields = [field_transition];
    let transition = iwa::ObjectReferenceTransition {
        aggregate_before: &[11],
        aggregate_after: &[11, 13],
        fields: &fields,
    };
    accepts_core_object_transition(transition);
    assert_eq!(transition.fields.len(), 1);

    assert_eq!(iwa::FieldType::from_raw(1), iwa::FieldType::ObjectReference);
    assert_eq!(
        iwa::UnknownFieldRule::from_raw(-1),
        iwa::UnknownFieldRule::NotSupported
    );
    let _ = iwa::LimitKind::ArchiveBytes;
    assert!(matches!(
        iwa::Archive::parse(&[0]),
        Err(iwa::Error::InvalidArchive { .. })
    ));

    Ok(())
}
