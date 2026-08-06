//! Animation-info round trips and default-state invariants.
use super::super::*;
use super::support::*;

#[test]
fn round_trips_exact_animation_info_atoms_and_containers() {
    let atom = sample_legacy_atom();
    let bytes = write_animation_info_atom(&atom).unwrap();
    assert_eq!(bytes.len(), 36);
    let (record, consumed) = Record::parse(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    assert_eq!(parse_animation_info_atom(&record).unwrap(), atom);

    let mut info = AnimationInfo::new();
    info.legacy_atom = Some(atom.clone());
    let (container, sound_ref) = write_animation_info(&info).unwrap();
    assert_eq!(sound_ref, 42);
    let (record, consumed) = Record::parse(&container, 0).unwrap();
    assert_eq!(consumed, container.len());
    let parsed = parse_animation_info(&record).unwrap();
    assert_eq!(parsed.legacy_atom, Some(atom));
    assert_eq!(parsed.animation_count(), 1);
    assert_eq!(parsed.after_effect_color, Some(0x0011_2233));
    assert_eq!(parsed.iteration, IterationType::ByLetter);
}

#[test]
fn rejects_malformed_animation_info_atoms() {
    let valid = write_animation_info_atom(&sample_legacy_atom()).unwrap();
    let mutations: &[(usize, u8)] = &[
        (12, 0x02), // invalid bool2 value
        (28, 0xFF), // invalid build type
        (29, 0x0F), // undefined effect
        (30, 0xFF), // invalid direction for Fly
        (31, 0x04), // invalid after effect
        (32, 0x03), // invalid text subdivision
    ];
    for &(offset, value) in mutations {
        let mut bytes = valid.clone();
        bytes[offset] = value;
        let (record, _) = Record::parse(&bytes, 0).unwrap();
        assert!(
            parse_animation_info_atom(&record).is_err(),
            "accepted mutation at byte {offset}"
        );
    }

    let mut short = valid;
    short[4..8].copy_from_slice(&27u32.to_le_bytes());
    let (record, _) = Record::parse(&short, 0).unwrap();
    assert!(parse_animation_info_atom(&record).is_err());
}

#[test]
fn test_animation_info_default() {
    let info = AnimationInfo::default();
    assert!(!info.has_animations());
    assert_eq!(info.animation_count(), 0);
}
