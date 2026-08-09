#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn sample(build_type: BuildType) -> Container {
    Container::new(
        Build::new(17, 42, true, false).with_reserved([0xA5, 0x5A]),
        Atom::new(build_type),
    )
    .unwrap()
}

#[test]
fn fixed_container_round_trips_all_known_diagram_build_values() {
    for raw in 0..=0x10 {
        let value = sample(BuildType::from_raw(raw));
        let bytes = value.to_bytes();
        assert_eq!(bytes.len(), Container::RECORD_LEN);
        let parsed = Container::parse_bytes(&bytes).unwrap();
        assert_eq!(parsed, value);
        assert_eq!(parse_bytes(&bytes).unwrap(), value);
        assert_eq!(Container::parse_record(&value.to_record()).unwrap(), value);
        assert_eq!(
            crate::animation::parse_diagram_build_record(&value.to_record()).unwrap(),
            value
        );
    }
}

#[test]
fn unknown_fixed_enums_and_reserved_bytes_are_lossless() {
    let value = sample(BuildType::Unknown(0xDEAD_BEEF));
    let parsed = Container::parse_bytes(&value.to_bytes()).unwrap();
    assert_eq!(parsed.atom().build_type, BuildType::Unknown(0xDEAD_BEEF));
    assert_eq!(parsed.build().kind(), Kind::Diagram);
    assert_eq!(parsed.build().reserved(), [0xA5, 0x5A]);

    let build = Build::parse_bytes(&{
        let mut bytes = value.build().to_bytes();
        bytes[8..12].copy_from_slice(&0x7FFF_FFFEu32.to_le_bytes());
        bytes
    })
    .unwrap();
    assert_eq!(build.kind(), Kind::Unknown(0x7FFF_FFFE));
}

#[test]
fn rejects_wrong_sizes_headers_child_order_and_bool_values() {
    let valid = sample(BuildType::AllAtOnce).to_bytes();
    for length in 0..valid.len() {
        assert!(Container::parse_bytes(&valid[..length]).is_err());
    }
    let mut trailing = valid.clone();
    trailing.push(0);
    assert!(Container::parse_bytes(&trailing).is_err());

    for (offset, value) in [(0, 1u8), (1, 1), (2, 1), (3, 1)] {
        let mut malformed = valid.clone();
        malformed[offset] = value;
        assert!(Container::parse_bytes(&malformed).is_err());
    }

    let mut wrong_bool = valid.clone();
    wrong_bool[8 + 8 + 12] = 2;
    assert!(Container::parse_bytes(&wrong_bool).is_err());

    let mut wrong_order = valid.clone();
    let build = wrong_order[8..32].to_vec();
    let atom = wrong_order[32..44].to_vec();
    wrong_order[8..20].copy_from_slice(&atom);
    wrong_order[20..44].copy_from_slice(&build);
    assert!(Container::parse_bytes(&wrong_order).is_err());
}

#[test]
fn rejects_known_non_diagram_build_kinds_but_keeps_unknown_kinds_bounded() {
    for raw in [1u32, 2] {
        let mut bytes = sample(BuildType::Custom).to_bytes();
        bytes[16..20].copy_from_slice(&raw.to_le_bytes());
        assert!(Container::parse_bytes(&bytes).is_err());
    }

    let mut bytes = sample(BuildType::Custom).to_bytes();
    bytes[16..20].copy_from_slice(&0x8000_0000u32.to_le_bytes());
    let parsed = Container::parse_bytes(&bytes).unwrap();
    assert_eq!(parsed.build().kind(), Kind::Unknown(0x8000_0000));
}

#[test]
fn generic_records_validate_raw_types_lengths_and_children() {
    let value = sample(BuildType::AsOneObject);
    let mut wrong_type = value.to_record();
    wrong_type.record_type_raw ^= 1;
    assert!(Container::parse_record(&wrong_type).is_err());

    let mut wrong_length = value.to_record();
    wrong_length.data_length = 35;
    assert!(Container::parse_record(&wrong_length).is_err());

    let mut extra_child = value.to_record();
    extra_child.children[1]
        .children
        .push(value.atom().to_record());
    assert!(Container::parse_record(&extra_child).is_err());

    let mut wrong_data = value.to_record();
    wrong_data.data[0] ^= 1;
    assert!(Container::parse_record(&wrong_data).is_err());
}

#[test]
fn atom_and_build_facades_reject_unsafe_boundaries() {
    let value = sample(BuildType::Down);
    assert_eq!(
        Atom::parse_bytes(&value.atom().to_bytes()).unwrap(),
        value.atom()
    );
    assert_eq!(
        Build::parse_bytes(&value.build().to_bytes()).unwrap(),
        value.build()
    );

    let mut atom = value.atom().to_record();
    atom.record_type = RecordType::BuildAtom;
    assert!(Atom::parse_record(&atom).is_err());

    let mut build = value.build().to_record();
    build.children.push(value.atom().to_record());
    assert!(Build::parse_record(&build).is_err());

    let mut parsed = Record::parse_strict(&value.to_bytes(), 0).unwrap().0;
    parsed.children.pop();
    assert!(Container::parse_record(&parsed).is_err());
}
