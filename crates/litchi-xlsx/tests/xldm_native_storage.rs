//! Native XLDM outer storage and explicitly framed XPRESS validation.
#![allow(clippy::unwrap_used, reason = "test assertions panic on failure")]

use litchi_opc::{OpcPackage, PackURI};
const SOURCE: &[u8] = include_bytes!("data/data_model/tdf167689_x15_namespace.xlsx");

#[test]
fn native_tabular_xpress_members_decode_with_explicit_framing_and_limits() {
    use litchi_xlsx::package::xldm::compression::{
        CodecLimits, XpressFraming, decompress_xpress, decompress_xpress_framed,
    };
    // Stored offsets/sizes come from the native VirtualDirectory (CRC excluded);
    // decoded lengths come independently from the native backup log's FileList.
    let members = [
        (4584, 1077, 3877),
        (5665, 148, 144),
        (5817, 1936, 7692),
        (7757, 135, 169),
        (7896, 2203, 6497),
        (10103, 1195, 3114),
        (11302, 224, 405),
        (11530, 1047, 2535),
        (12581, 8992, 33156),
        (21577, 792, 4701),
        (22373, 33, 152),
        (22410, 22, 144),
        (22436, 83, 1213),
        (22523, 140, 1213),
        (22667, 45, 168),
        (22716, 1656, 5965),
        (24376, 33, 152),
        (24413, 43, 168),
        (24460, 35, 64),
        (24499, 24, 64),
        (24527, 38, 168),
        (24569, 13228, 48873),
        (37801, 37, 160),
        (37842, 101, 1213),
        (37947, 45, 72),
        (37996, 1506, 5020),
        (39506, 2497, 10589),
        (42007, 38, 168),
        (42049, 83, 1213),
        (42136, 1518, 5020),
        (43658, 1657, 5969),
        (45319, 1519, 5021),
        (46842, 95, 1213),
        (46941, 1520, 5020),
        (48465, 86, 155),
        (48555, 52, 48),
        (48611, 86, 155),
        (48701, 52, 48),
        (48757, 32, 56),
        (48793, 2520, 10610),
        (51317, 34, 64),
        (51355, 32, 152),
        (51391, 22, 64),
        (51417, 95, 1213),
        (51516, 35, 64),
        (51555, 34, 64),
    ];
    let raw = OpcPackage::from_bytes(SOURCE).unwrap();
    let part = raw
        .get_part(&PackURI::new("/xl/model/item.data").unwrap())
        .unwrap();
    for (index, (offset, size, decoded)) in members.into_iter().enumerate() {
        let bytes = &part.blob()[offset..offset + size];
        let limits = CodecLimits {
            max_input_bytes: size,
            max_output_bytes: decoded,
            ..CodecLimits::default()
        };
        let output = decompress_xpress_framed(bytes, XpressFraming::Tabular16, limits)
            .unwrap_or_else(|error| panic!("native member {index}: {error}"));
        assert_eq!(output.len(), decoded);
        assert!(
            decompress_xpress_framed(
                bytes,
                XpressFraming::Tabular16,
                CodecLimits {
                    max_output_bytes: decoded - 1,
                    ..limits
                }
            )
            .is_err()
        );
        if index == 0 {
            assert!(output.starts_with(b"<Load "));
            assert!(output.ends_with(b"</Load>"));
            assert!(decompress_xpress(bytes, limits).is_err());
        }
    }
}

#[test]
fn native_outer_storage_is_borrowed_exact_and_checks_member_crcs() {
    use litchi_xlsx::package::xldm::{StorageProfile, inspect, write};
    let package = OpcPackage::from_bytes(SOURCE).unwrap();
    let part = package
        .get_part(&PackURI::new("/xl/model/item.data").unwrap())
        .unwrap();
    let bytes = part.blob();
    let storage = inspect(bytes).unwrap();
    assert_eq!(storage.profile(), StorageProfile::Tabular150);
    assert_eq!(storage.files.len(), 48);
    assert!(std::ptr::eq(storage.bytes().as_ptr(), bytes.as_ptr()));
    assert_eq!(write(&storage).unwrap(), bytes);
    for entry in &storage.files {
        let mut corrupt = bytes.to_vec();
        let crc_byte = usize::try_from(entry.offset.0 + entry.stored_size.0 - 1).unwrap();
        corrupt[crc_byte] ^= 1;
        assert!(
            inspect(&corrupt).is_err(),
            "accepted corrupt CRC for {}",
            entry.path
        );
    }
}

#[test]
fn public_xldm_path_classifier_preserves_xlsx_error_type() {
    let result = litchi_xlsx::package::xldm::classify_generated_path("../escape");
    assert!(matches!(result, Err(litchi_xlsx::Error::Invalid(_))));
}
