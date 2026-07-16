use litchi_ole::doc::{FieldStory, FieldType, Package};
use std::path::PathBuf;

fn fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/document")
        .join(name)
}

fn reference_plcf(cps: &[u32], descriptors: &[[u8; 2]]) -> Vec<u8> {
    assert_eq!(cps.len(), descriptors.len() + 1);
    let mut bytes = Vec::with_capacity(cps.len() * 4 + descriptors.len() * 2);
    for cp in cps {
        bytes.extend_from_slice(&cp.to_le_bytes());
    }
    for descriptor in descriptors {
        bytes.extend_from_slice(descriptor);
    }
    bytes
}

#[test]
fn apache_poi_reference_plcfs_cover_all_seven_story_tables() {
    // Ported from Apache POI `TestFieldsTables.EXPECTED`, whose source fixture
    // has a malformed CFB root name that strict litchi-cfb intentionally rejects.
    let cases: &[(FieldStory, &[u32], &[[u8; 2]], &[FieldType])] = &[
        (
            FieldStory::Comment,
            &[19, 43, 54, 59],
            &[[0x13, 0x1F], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::Date],
        ),
        (
            FieldStory::Endnote,
            &[31, 59, 61, 66],
            &[[0x13, 0x45], [0x14, 0xFF], [0x15, 0x80]],
            &[FieldType::FileSize],
        ),
        (
            FieldStory::Footnote,
            &[23, 49, 64, 69],
            &[[0x13, 0x11], [0x14, 0xFF], [0x15, 0x80]],
            &[FieldType::Author],
        ),
        (
            FieldStory::Header,
            &[18, 42, 44, 47, 75, 85, 91],
            &[
                [0x13, 0x21],
                [0x14, 0xFF],
                [0x15, 0x81],
                [0x13, 0x1D],
                [0x14, 0xFF],
                [0x15, 0x81],
            ],
            &[FieldType::Page, FieldType::FileName],
        ),
        (
            FieldStory::HeaderTextbox,
            &[30, 54, 62, 68],
            &[[0x13, 0x20], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::Time],
        ),
        (
            FieldStory::Main,
            &[1, 31, 51, 541],
            &[[0x13, 0x15], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::CreateDate],
        ),
        (
            FieldStory::Textbox,
            &[19, 47, 49, 55],
            &[[0x13, 0x19], [0x14, 0xFF], [0x15, 0x81]],
            &[FieldType::EditTime],
        ),
    ];

    for &(story, cps, descriptors, expected_types) in cases {
        let reference = reference_plcf(cps, descriptors);
        let table = litchi_ole::doc::FieldStoryTable::parse_plcf(
            story,
            *cps.last().unwrap(),
            &reference,
        )
        .unwrap();
        assert_eq!(table.terminal_cp(), *cps.last().unwrap());
        assert_eq!(
            table.fields().iter().map(|field| field.field_type).collect::<Vec<_>>(),
            expected_types
        );
        assert!(table.fields().iter().all(|field| field.end_flags.has_separator));
        assert_eq!(table.to_plcf_bytes().unwrap(), reference);
    }
}

#[test]
fn poi_hyperlink_fixture_retains_inert_hyperlink_field() {
    let mut package = Package::open(fixture("hyperlink.doc")).unwrap();
    let document = package.document().unwrap();
    let fields = document.fields_table().expect("fields table");
    assert!(
        fields
            .main_document_fields()
            .iter()
            .any(|field| field.field_type == FieldType::Hyperlink)
    );
}
