//! Focused writer validation tests.

use super::build::validate_paragraph_levels;
use super::*;
use crate::animation::{
    AnimationInfo, AnimationSound, BuildAtom, BuildList, BuildListEntry, BuiltinSound, ChartBuild,
    ChartBuildAtom, ChartBuildType, ExtendedTimeNode, LegacyAnimationAtom, LegacyAnimationEffect,
    ParagraphBuildLevel, TimeNodeAtom,
};
#[test]
fn animation_level_sound_is_serialized_without_a_build_list() {
    let mut info = AnimationInfo::new();
    info.sound = Some(AnimationSound::builtin(BuiltinSound::Whoosh));

    let (_, sound_ref) = write_animation_info(&info).unwrap();
    assert_eq!(sound_ref, BuiltinSound::Whoosh.id());
}

#[test]
fn test_write_build_list_empty() {
    let build_info = BuildList::new();
    let data = write_build_list(&build_info).unwrap();

    assert_eq!(data.len(), 8);
}

#[test]
fn rejects_invalid_paragraph_builds() {
    let time_node = ExtendedTimeNode {
        atom: TimeNodeAtom::default(),
        ..ExtendedTimeNode::default()
    };
    let level = ParagraphBuildLevel {
        level: 10,
        time_node,
    };
    assert!(
        validate_paragraph_levels(
            &crate::animation::types::ParagraphBuildType::AllAtOnce,
            &[level]
        )
        .is_err()
    );
    assert!(
        validate_paragraph_levels(&crate::animation::types::ParagraphBuildType::AsAWhole, &[])
            .is_err()
    );
}

#[test]
fn rejects_duplicate_build_id_shape_pairs() {
    let entry = || {
        BuildListEntry::Chart(ChartBuild {
            atom: BuildAtom {
                build_id: 5,
                shape_id_ref: 9,
                expanded: false,
                ui_expanded: false,
            },
            chart: ChartBuildAtom {
                build_type: ChartBuildType::AsOneObject,
                animate_background: false,
            },
        })
    };
    let list = BuildList {
        builds: vec![entry(), entry()],
    };
    assert!(write_build_list(&list).is_err());
}

#[test]
fn rejects_invalid_legacy_animation_atom_combinations() {
    let mut atom = LegacyAnimationAtom {
        effect: LegacyAnimationEffect::Wheel,
        effect_direction: 7,
        ..LegacyAnimationAtom::default()
    };
    assert!(write_animation_info_atom(&atom).is_err());

    atom.effect = LegacyAnimationEffect::Cut;
    atom.effect_direction = 0;
    atom.automatic = true;
    atom.delay_time_ms = -1;
    assert!(write_animation_info_atom(&atom).is_err());
}
