use super::{Flags, Kind, MAX_NAME_UNITS, Snapshot};
use crate::prop::{Id, Props};
use crate::{Error, Record, RecordKind};

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
        if let Some(data) = data {
            body.extend_from_slice(data);
        }
    }
    let mut record = Vec::with_capacity(body.len() + 8);
    let version_instance = ((properties.len() as u16) << 4) | 3;
    record.extend_from_slice(&version_instance.to_le_bytes());
    record.extend_from_slice(&RecordKind::Opt.raw().to_le_bytes());
    record.extend_from_slice(&(body.len() as u32).to_le_bytes());
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
            name.len() as i32,
            Some(&name),
        ),
        (Id::BlipFlags.raw(), 0x8000_0009_u32 as i32, None),
    ]);
    let (record, consumed) = Record::parse(&bytes, 0).expect("picture Opt");
    assert_eq!(consumed, bytes.len());

    let props = Props::parse(&record).expect("properties");
    let metadata = props
        .picture()
        .expect("picture metadata")
        .expect("metadata present");
    let name = metadata.name().expect("name");
    assert_eq!(name.text().expect("UTF-16"), "old");
    assert_eq!(name.raw_bytes(), name_bytes("old").as_slice());
    assert_eq!(metadata.flags().kind(), Kind::File);
    assert!(metadata.flags().link_to_file());
    assert_eq!(metadata.flags().reserved(), 0x8000_0000);

    let snapshot = Snapshot::parse(&record).expect("snapshot");
    assert_eq!(snapshot.encode().expect("lossless encode"), bytes);
}

#[test]
fn edits_only_modeled_picture_values_and_preserves_opaque_neighbors() {
    let old_name = name_bytes("old");
    let opaque = [0xDE, 0xAD, 0xBE, 0xEF];
    let bytes = opt_record(&[
        (0x0600, -7, None),
        (
            0x8000 | Id::PictureFileName.raw(),
            old_name.len() as i32,
            Some(&old_name),
        ),
        (Id::BlipFlags.raw(), 0x8000_0009_u32 as i32, None),
        (0x8000 | 0x0601, opaque.len() as i32, Some(&opaque)),
    ]);
    let (record, _) = Record::parse(&bytes, 0).expect("picture Opt");
    let source = Snapshot::parse(&record).expect("snapshot");

    let mut edit = source.edit();
    edit.set_name("new").expect("new name");
    edit.set_flags(Flags::new(Kind::Url, true, false).expect("URL flags"));
    let commit = edit.commit().expect("commit");
    assert!(!commit.patch().is_empty());
    assert_eq!(
        commit.patch().name().expect("name patch").before(),
        Some(old_name.as_ref())
    );
    assert_eq!(
        commit.patch().name().expect("name patch").after(),
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

    let owned = commit.snapshot();
    let edited = owned.snapshot().expect("edited snapshot");
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
            .map(|property| property.raw_id())
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
    assert_eq!(source.encode().expect("source unchanged"), bytes);
}

#[test]
fn validates_utf16_bounds_and_flag_dependencies() {
    assert!(matches!(
        Flags::from_raw(3),
        Err(Error::MalformedProperties {
            reason: "picture name kind uses a reserved flag value"
        })
    ));
    assert!(matches!(
        Flags::from_raw(4),
        Err(Error::MalformedProperties {
            reason: "picture do-not-save requires link-to-file"
        })
    ));
    assert!(Flags::new(Kind::Url, true, true).is_ok());
    assert!(Flags::new(Kind::Comment, true, false).is_err());

    let odd = [0x41, 0x00, 0x00];
    let bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        odd.len() as i32,
        Some(&odd),
    )]);
    let (record, _) = Record::parse(&bytes, 0).expect("picture Opt");
    assert!(matches!(
        Snapshot::parse(&record),
        Err(Error::MalformedProperties {
            reason: "picture name must be a bounded even-length UTF-16 string"
        })
    ));

    let interior_nul = [0x41, 0x00, 0x00, 0x00, 0x00, 0x00];
    let bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        interior_nul.len() as i32,
        Some(&interior_nul),
    )]);
    let (record, _) = Record::parse(&bytes, 0).expect("picture Opt");
    assert!(matches!(
        Snapshot::parse(&record),
        Err(Error::MalformedProperties {
            reason: "picture name contains an interior NUL"
        })
    ));

    let too_long = [0x41, 0x00].repeat(MAX_NAME_UNITS + 2);
    let bytes = opt_record(&[(
        0x8000 | Id::PictureFileName.raw(),
        too_long.len() as i32,
        Some(&too_long),
    )]);
    let (record, _) = Record::parse(&bytes, 0).expect("picture Opt");
    assert!(matches!(
        Snapshot::parse(&record),
        Err(Error::MalformedProperties {
            reason: "picture name must be a bounded even-length UTF-16 string"
        })
    ));
}

#[test]
fn appends_and_clears_names_without_touching_other_properties() {
    let opaque = [1, 2, 3];
    let bytes = opt_record(&[(0x8000 | 0x0601, opaque.len() as i32, Some(&opaque))]);
    let (record, _) = Record::parse(&bytes, 0).expect("picture Opt");
    let source = Snapshot::parse(&record).expect("snapshot");

    let mut add = source.edit();
    add.set_name("added").expect("name");
    let added = add.commit().expect("add name");
    let added_snapshot = added.snapshot().snapshot().expect("added snapshot");
    assert_eq!(
        added_snapshot
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
        added_snapshot
            .properties()
            .get_binary(Id::unknown(0x0601).expect("unknown")),
        Some(&opaque[..])
    );

    let mut clear = added_snapshot.edit();
    clear.clear_name();
    let cleared = clear.commit().expect("clear name");
    assert!(
        cleared
            .snapshot()
            .snapshot()
            .expect("cleared snapshot")
            .metadata()
            .expect("metadata")
            .is_none()
    );
}
