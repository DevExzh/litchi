#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::*;
use super::model::*;
use crate::consts::RecordType;
use crate::records::Record;

fn metadata() -> Metadata {
    Metadata {
        draw_aspect: DrawAspect::Icon,
        object_type: ObjectType::Embedded,
        id: 17,
        subtype: ObjectSubtype::ExcelChart,
        persist_id: 9,
        unused: [1, 2, 3, 4],
    }
}

#[test]
fn ole_object_metadata_roundtrips_exactly() {
    let expected = metadata();
    let parsed = Metadata::parse(&expected.to_record().unwrap()).unwrap();
    assert_eq!(parsed, expected);
    assert_eq!(
        parsed.to_record_bytes().unwrap(),
        expected.to_record_bytes().unwrap()
    );
}

#[test]
fn ole_object_metadata_rejects_invalid_domains_and_ids() {
    let mut bytes = metadata().to_record_bytes().unwrap();
    bytes[8..12].copy_from_slice(&99u32.to_le_bytes());
    assert!(Metadata::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
    let mut value = metadata();
    value.persist_id = 0;
    assert!(value.to_record_bytes().is_err());
}

#[test]
fn embed_preferences_preserve_recommendation_and_unused_bytes() {
    let expected = EmbedPreferences {
        color_follow: ColorFollow::TextAndBackground,
        cannot_lock_server: true,
        dimension_policy: DimensionPolicy::ProducerDefined(7),
        is_word_table: false,
        unused: 0xa5,
    };
    let parsed = EmbedPreferences::parse(&expected.to_record().unwrap()).unwrap();
    assert_eq!(parsed, expected);
}

#[test]
fn link_info_roundtrips_nullable_slide_and_rejects_update_domain() {
    let expected = LinkInfo {
        slide_id: None,
        update_mode: UpdateMode::OnCall,
        unused: [0xde, 0xad, 0xbe, 0xef],
    };
    assert_eq!(
        LinkInfo::parse(&expected.to_record().unwrap()).unwrap(),
        expected
    );
    let mut bytes = expected.to_record_bytes().unwrap();
    bytes[12..16].copy_from_slice(&2u32.to_le_bytes());
    assert!(LinkInfo::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

fn definition(kind: ContainerKind) -> Definition {
    let object_type = match kind {
        ContainerKind::Embedded(_) => ObjectType::Embedded,
        ContainerKind::Linked(_) => ObjectType::Linked,
    };
    Definition {
        kind,
        object: Metadata {
            object_type,
            ..metadata()
        },
        menu_name: Some("Worksheet".into()),
        program_id: Some("Excel.Sheet.12".into()),
        clipboard_name: Some("Microsoft Excel Worksheet".into()),
        metafile: Some(vec![0xd7, 0xcd, 0xc6, 0x9a]),
    }
}

#[test]
fn embedded_and_linked_containers_roundtrip_canonically() {
    let embedded = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::EntireScheme,
        cannot_lock_server: true,
        dimension_policy: DimensionPolicy::Omit,
        is_word_table: false,
        unused: 7,
    }));
    let linked = definition(ContainerKind::Linked(LinkInfo {
        slide_id: Some(256),
        update_mode: UpdateMode::OnCall,
        unused: [1, 2, 3, 4],
    }));
    for expected in [embedded, linked] {
        let parsed = Definition::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
    }
}

#[test]
fn containers_reject_type_mismatch_and_hostile_strings() {
    let mut value = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    }));
    value.object.object_type = ObjectType::Linked;
    assert!(value.to_record_bytes().is_err());
    value.object.object_type = ObjectType::Embedded;
    value.program_id = Some("bad\nprogram".into());
    assert!(value.to_record_bytes().is_err());
    value.program_id = Some("x".repeat(MAX_OLE_NAME_UNITS + 1));
    assert!(value.to_record_bytes().is_err());
}

#[test]
fn containers_reject_duplicate_or_out_of_order_optional_atoms() {
    let value = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    }));
    let mut children = value.kind_embedded_bytes_for_test();
    children.extend_from_slice(&value.object.to_record_bytes().unwrap());
    children.extend_from_slice(
        &record_bytes(
            0,
            2,
            RecordType::CString,
            &encode_ole_string("Prog", true).unwrap(),
        )
        .unwrap(),
    );
    children.extend_from_slice(
        &record_bytes(
            0,
            1,
            RecordType::CString,
            &encode_ole_string("Menu", false).unwrap(),
        )
        .unwrap(),
    );
    let bytes = record_bytes(0x0f, 0, RecordType::ExternalOleEmbed, &children).unwrap();
    assert!(Definition::parse(&Record::parse(&bytes, 0).unwrap().0).is_err());
}

impl Definition {
    fn kind_embedded_bytes_for_test(&self) -> Vec<u8> {
        match self.kind {
            ContainerKind::Embedded(value) => value.to_record_bytes().unwrap(),
            ContainerKind::Linked(_) => unreachable!(),
        }
    }
}

fn external_object_list(seed: i32, objects: &[Vec<u8>]) -> Record {
    external_object_list_with_children(seed, objects).0
}

fn external_object_list_with_children(seed: i32, children: &[Vec<u8>]) -> (Record, Vec<u8>) {
    let mut child_bytes =
        record_bytes(0, 0, RecordType::ExObjListAtom, &seed.to_le_bytes()).unwrap();
    for child in children {
        child_bytes.extend_from_slice(child);
    }
    let bytes = record_bytes(0x0f, 0, RecordType::ExObjList, &child_bytes).unwrap();
    (Record::parse(&bytes, 0).unwrap().0, bytes)
}

#[test]
fn ole_collection_preserves_unknown_children_and_source_slots() {
    let value = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    }));
    let first_unknown = record_bytes_raw(0, 7, 0x7777, b"before").unwrap();
    let second_unknown = record_bytes_raw(0, 9, 0x8888, b"after").unwrap();
    let (root, original) = external_object_list_with_children(
        i32::try_from(value.object.id).unwrap(),
        &[
            first_unknown.clone(),
            value.to_record_bytes().unwrap(),
            second_unknown.clone(),
        ],
    );
    let collection = Collection::parse(&root).unwrap().unwrap();
    assert_eq!(collection.unknown_records().len(), 2);
    assert_eq!(collection.unknown_records()[0].record_type(), 0x7777);
    assert_eq!(collection.unknown_records()[0].data(), b"before");
    assert_eq!(collection.unknown_records()[1].record_type(), 0x8888);
    assert_eq!(collection.unknown_records()[1].data(), b"after");
    assert_eq!(collection.to_record_bytes().unwrap(), original);
    assert_eq!(
        collection.unknown_records()[0].to_record_bytes().unwrap(),
        first_unknown
    );
    assert_eq!(
        collection.unknown_records()[1].to_record_bytes().unwrap(),
        second_unknown
    );
}

#[test]
fn ole_collection_reorders_typed_objects_without_losing_unknown_slots() {
    let kind = ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    });
    let mut first = definition(kind);
    first.object.id = 1;
    let mut second = first.clone();
    second.object.id = 2;
    second.object.persist_id = 10;
    let first_unknown = record_bytes_raw(0, 1, 0x7777, b"slot-0").unwrap();
    let middle_unknown = record_bytes_raw(0, 2, 0x8888, b"slot-1").unwrap();
    let last_unknown = record_bytes_raw(0, 3, 0x9999, b"slot-2").unwrap();
    let (root, _) = external_object_list_with_children(
        2,
        &[
            first_unknown.clone(),
            first.to_record_bytes().unwrap(),
            middle_unknown.clone(),
            second.to_record_bytes().unwrap(),
            last_unknown.clone(),
        ],
    );
    let mut collection = Collection::parse(&root).unwrap().unwrap();
    collection.reorder(&[2, 1]).unwrap();
    let (_, reordered) = external_object_list_with_children(
        2,
        &[
            first_unknown,
            second.to_record_bytes().unwrap(),
            middle_unknown,
            first.to_record_bytes().unwrap(),
            last_unknown,
        ],
    );
    assert_eq!(collection.to_record_bytes().unwrap(), reordered);
}

#[test]
fn ole_collection_discovers_objects_and_enforces_seed() {
    let mut first = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    }));
    first.object.id = 21;
    let root = external_object_list(21, &[first.to_record_bytes().unwrap()]);
    let parsed = Collection::parse(&root).unwrap().unwrap();
    assert_eq!(parsed.id_seed, 21);
    assert!(parsed.get(21).is_some());
    let wrong_seed_root = external_object_list(20, &[first.to_record_bytes().unwrap()]);
    assert!(Collection::parse(&wrong_seed_root).is_err());
}

#[test]
fn ole_collection_rejects_duplicate_ids() {
    let first = definition(ContainerKind::Embedded(EmbedPreferences {
        color_follow: ColorFollow::None,
        cannot_lock_server: false,
        dimension_policy: DimensionPolicy::Send,
        is_word_table: false,
        unused: 0,
    }));
    let mut second = first.clone();
    second.object.persist_id += 1;
    let root = external_object_list(
        i32::try_from(first.object.id).unwrap(),
        &[
            first.to_record_bytes().unwrap(),
            second.to_record_bytes().unwrap(),
        ],
    );
    assert!(Collection::parse(&root).is_err());
}

#[test]
fn activex_control_roundtrips_as_inert_metadata() {
    let expected = Control {
        slide_id: Some(512),
        object: Metadata {
            object_type: ObjectType::ActiveXControl,
            ..metadata()
        },
        menu_name: Some("Calendar".into()),
        program_id: Some("MSCAL.Calendar.7".into()),
        clipboard_name: None,
        metafile: Some(vec![1, 2, 3]),
    };
    let parsed = Control::parse(&expected.to_record().unwrap()).unwrap();
    assert_eq!(parsed, expected);
}
