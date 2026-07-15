//! Animation record parser.
//!
//! Parses PowerPoint binary animation records into structured types.

use super::triggers::IterationType;
use super::types::{
    AfterEffect, AnimationInfo, BuildAtom, BuildKind, BuildList, BuildListEntry, ChartBuild,
    ChartBuildAtom, ChartBuildType, DiagramBuild, DiagramBuildAtom, DiagramBuildType,
    ExtendedTimeNode, LegacyAnimationAtom, LegacyAnimationBuild, LegacyAnimationEffect,
    LegacyTextBuildSubEffect, ParagraphBuild, ParagraphBuildAtom, ParagraphBuildLevel,
    ParagraphBuildType, TimeNodeAtom, TimeNodeFill, TimeNodeKind, TimeNodeRestart,
};
use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;

/// Parse animation info from AnimationInfo container record.
pub fn parse_animation_info(record: &PptRecord) -> Result<AnimationInfo> {
    if record.record_type != PptRecordType::AnimationInfo {
        return Err(PptError::InvalidFormat(format!(
            "Expected AnimationInfo record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 0x0F || record.instance != 0 {
        return Err(PptError::Corrupted(format!(
            "AnimationInfo requires version 15 and instance 0; got version {} and instance {}",
            record.version, record.instance
        )));
    }

    let mut info = AnimationInfo::new();
    let atom_record = record.children.first().ok_or_else(|| {
        PptError::Corrupted("AnimationInfo is missing its AnimationInfoAtom".to_string())
    })?;
    if atom_record.record_type != PptRecordType::AnimationInfoAtom {
        return Err(PptError::InvalidFormat(
            "AnimationInfoAtom must be the first AnimationInfo child".to_string(),
        ));
    }
    let atom = parse_animation_info_atom(atom_record)?;
    info.after_effect_color = Some(atom.dim_color);
    info.iteration = match atom.text_build_sub_effect {
        LegacyTextBuildSubEffect::AllAtOnce => IterationType::All,
        LegacyTextBuildSubEffect::ByWord => IterationType::ByWord,
        LegacyTextBuildSubEffect::ByCharacter => IterationType::ByLetter,
    };
    info.legacy_atom = Some(atom);
    for child in record.children.iter().skip(1) {
        if child.record_type == PptRecordType::AnimationInfoAtom {
            return Err(PptError::InvalidFormat(
                "AnimationInfo contains multiple AnimationInfoAtom records".to_string(),
            ));
        }
        info.raw_records.push(child.clone());
    }

    Ok(info)
}

/// Parse the exact PowerPoint 97 `AnimationInfoAtom` payload.
pub fn parse_animation_info_atom(record: &PptRecord) -> Result<LegacyAnimationAtom> {
    if record.record_type != PptRecordType::AnimationInfoAtom {
        return Err(PptError::InvalidFormat(format!(
            "Expected AnimationInfoAtom record, got {:?}",
            record.record_type
        )));
    }
    if record.version != 1 || record.instance != 0 || record.data.len() != 28 {
        return Err(PptError::Corrupted(format!(
            "AnimationInfoAtom requires version 1, instance 0, and 28 data bytes; got version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }

    let dim_color = u32::from_le_bytes(record.data[0..4].try_into().expect("length checked"));
    let flags = u16::from_le_bytes(record.data[4..6].try_into().expect("length checked"));
    let mut decoded_flags = [false; 8];
    for (index, decoded) in decoded_flags.iter_mut().enumerate() {
        let value = (flags >> (index * 2)) & 0x03;
        if value > 1 {
            return Err(PptError::InvalidFormat(format!(
                "AnimationInfoAtom flag field {index} has invalid bool2 value {value}"
            )));
        }
        *decoded = value == 1;
    }
    let sound_id_ref = u32::from_le_bytes(record.data[8..12].try_into().expect("length checked"));
    let delay_time_ms = i32::from_le_bytes(record.data[12..16].try_into().expect("length checked"));
    if decoded_flags[1] && delay_time_ms < 0 {
        return Err(PptError::InvalidFormat(
            "automatic AnimationInfoAtom has a negative delay".to_string(),
        ));
    }
    let order_id = i16::from_le_bytes(record.data[16..18].try_into().expect("length checked"));
    if order_id < -2 {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom orderID {order_id} is less than -2"
        )));
    }
    let slide_count = u16::from_le_bytes(record.data[18..20].try_into().expect("length checked"));
    let build_type = LegacyAnimationBuild::parse(record.data[20]).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid AnimationInfoAtom animBuildType {:#04X}",
            record.data[20]
        ))
    })?;
    let effect = LegacyAnimationEffect::parse(record.data[21]).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid AnimationInfoAtom animEffect {:#04X}",
            record.data[21]
        ))
    })?;
    let effect_direction = record.data[22];
    if !effect.accepts_direction(effect_direction) {
        return Err(PptError::InvalidFormat(format!(
            "AnimationInfoAtom direction {effect_direction:#04X} is invalid for {effect:?}"
        )));
    }
    let after_effect = match record.data[23] {
        0 => AfterEffect::None,
        1 => AfterEffect::DimToColor,
        2 => AfterEffect::HideOnNextClick,
        3 => AfterEffect::Hide,
        value => {
            return Err(PptError::InvalidFormat(format!(
                "invalid AnimationInfoAtom animAfterEffect {value:#04X}"
            )));
        },
    };
    let text_build_sub_effect =
        LegacyTextBuildSubEffect::parse(record.data[24]).ok_or_else(|| {
            PptError::InvalidFormat(format!(
                "invalid AnimationInfoAtom textBuildSubEffect {:#04X}",
                record.data[24]
            ))
        })?;

    Ok(LegacyAnimationAtom {
        dim_color,
        reverse: decoded_flags[0],
        automatic: decoded_flags[1],
        has_sound: decoded_flags[2],
        stop_sound: decoded_flags[3],
        play: decoded_flags[4],
        synchronous: decoded_flags[5],
        hide_while_not_playing: decoded_flags[6],
        animate_background: decoded_flags[7],
        sound_id_ref,
        delay_time_ms,
        order_id,
        slide_count,
        build_type,
        effect,
        effect_direction,
        after_effect,
        text_build_sub_effect,
        ole_verb: record.data[25],
    })
}

/// Parse the exact envelope and atom of an extended PowerPoint 2002 time node.
pub fn parse_extended_time_node(record: &PptRecord) -> Result<ExtendedTimeNode> {
    require_container(record, PptRecordType::ExtTimeNode, 1, "ExtTimeNode")?;
    let atom_record = record.children.first().ok_or_else(|| {
        PptError::Corrupted("ExtTimeNode is missing its TimeNodeAtom".to_string())
    })?;
    let atom = parse_time_node_atom(atom_record)?;
    if record.children[1..]
        .iter()
        .any(|child| child.record_type == PptRecordType::TimeNode)
    {
        return Err(PptError::InvalidFormat(
            "ExtTimeNode contains multiple TimeNodeAtom records".to_string(),
        ));
    }
    Ok(ExtendedTimeNode {
        atom,
        children: record.children[1..].to_vec(),
    })
}

/// Parse the exact 32-byte payload of a `TimeNodeAtom`.
pub fn parse_time_node_atom(record: &PptRecord) -> Result<TimeNodeAtom> {
    require_atom(record, PptRecordType::TimeNode, 0, 32, "TimeNodeAtom")?;
    let flags = read_u32(&record.data, 28);
    let fill = parse_optional_time_value(
        flags & 0x01 != 0,
        read_u32(&record.data, 12),
        TimeNodeFill::parse,
        "TimeNodeAtom fill",
    )?;
    let restart = parse_optional_time_value(
        flags & 0x02 != 0,
        read_u32(&record.data, 4),
        TimeNodeRestart::parse,
        "TimeNodeAtom restart",
    )?;
    let node_type = parse_optional_time_value(
        flags & 0x08 != 0,
        read_u32(&record.data, 8),
        TimeNodeKind::parse,
        "TimeNodeAtom type",
    )?;
    let duration = i32::from_le_bytes(record.data[24..28].try_into().expect("length checked"));
    let duration_ms = if flags & 0x10 != 0 {
        Some(duration)
    } else if duration == 0 {
        None
    } else {
        return Err(PptError::InvalidFormat(
            "TimeNodeAtom duration must be zero when not explicitly set".to_string(),
        ));
    };
    Ok(TimeNodeAtom {
        fill,
        restart,
        node_type,
        duration_ms,
    })
}

/// Parse build list from BuildList container record.
pub fn parse_build_list(record: &PptRecord) -> Result<BuildList> {
    if record.record_type != PptRecordType::BuildList {
        return Err(PptError::InvalidFormat(format!(
            "Expected BuildList record, got {:?}",
            record.record_type
        )));
    }

    require_container(record, PptRecordType::BuildList, 0, "BuildList")?;
    let mut build_info = BuildList::new();
    let mut identities = std::collections::HashSet::with_capacity(record.children.len());
    for child in &record.children {
        let build = match child.record_type {
            PptRecordType::ParaBuild => BuildListEntry::Paragraph(parse_paragraph_build(child)?),
            PptRecordType::ChartBuild => BuildListEntry::Chart(parse_chart_build(child)?),
            PptRecordType::DiagramBuild => BuildListEntry::Diagram(parse_diagram_build(child)?),
            other => {
                return Err(PptError::InvalidFormat(format!(
                    "BuildList contains invalid child {other:?}"
                )));
            },
        };
        let atom = match &build {
            BuildListEntry::Paragraph(build) => &build.atom,
            BuildListEntry::Chart(build) => &build.atom,
            BuildListEntry::Diagram(build) => &build.atom,
        };
        if !identities.insert((atom.build_id, atom.shape_id_ref)) {
            return Err(PptError::InvalidFormat(format!(
                "duplicate build identity ({}, {})",
                atom.build_id, atom.shape_id_ref
            )));
        }
        build_info.add_build(build);
    }
    Ok(build_info)
}

fn parse_build_atom(record: &PptRecord, expected: BuildKind) -> Result<BuildAtom> {
    if record.record_type != PptRecordType::BuildAtom {
        return Err(PptError::InvalidFormat(format!(
            "Expected BuildAtom, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0, 0, Some(16), "BuildAtom")?;
    let kind = BuildKind::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid BuildAtom build type {}",
            read_u32(&record.data, 0)
        ))
    })?;
    if kind != expected {
        return Err(PptError::InvalidFormat(format!(
            "BuildAtom type {kind:?} does not match {expected:?} container"
        )));
    }
    Ok(BuildAtom {
        build_id: read_u32(&record.data, 4),
        shape_id_ref: read_u32(&record.data, 8),
        expanded: parse_bool1(record.data[12], "BuildAtom.fExpanded")?,
        ui_expanded: parse_bool1(record.data[13], "BuildAtom.fUIExpanded")?,
    })
}

fn parse_paragraph_build(record: &PptRecord) -> Result<ParagraphBuild> {
    require_container(record, PptRecordType::ParaBuild, 0, "ParaBuild")?;
    if record.children.len() < 4 || (record.children.len() - 2) % 2 != 0 {
        return Err(PptError::Corrupted(
            "ParaBuild requires two atoms followed by level/time-node pairs".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Paragraph)?;
    let paragraph = parse_paragraph_build_atom(&record.children[1])?;
    let mut levels = Vec::with_capacity((record.children.len() - 2) / 2);
    for pair in record.children[2..].chunks_exact(2) {
        let level = parse_level_info_atom(&pair[0])?;
        let time_node = parse_extended_time_node(&pair[1])?;
        if levels
            .last()
            .is_some_and(|previous: &ParagraphBuildLevel| previous.level >= level)
        {
            return Err(PptError::InvalidFormat(
                "ParaBuild levels must be strictly increasing".to_string(),
            ));
        }
        levels.push(ParagraphBuildLevel { level, time_node });
    }
    if paragraph.build_type == ParagraphBuildType::AsAWhole && levels.len() != 1 {
        return Err(PptError::InvalidFormat(
            "AsAWhole ParaBuild requires exactly one level".to_string(),
        ));
    }
    Ok(ParagraphBuild {
        atom,
        paragraph,
        levels,
    })
}

fn parse_paragraph_build_atom(record: &PptRecord) -> Result<ParagraphBuildAtom> {
    require_atom(record, PptRecordType::ParaBuildAtom, 1, 16, "ParaBuildAtom")?;
    let build_type = ParagraphBuildType::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid ParaBuildAtom type {}",
            read_u32(&record.data, 0)
        ))
    })?;
    Ok(ParagraphBuildAtom {
        build_type,
        build_level: read_u32(&record.data, 4),
        animate_background: parse_bool1(record.data[8], "ParaBuildAtom.fAnimBackground")?,
        reverse: parse_bool1(record.data[9], "ParaBuildAtom.fReverse")?,
        user_set_animate_background: parse_bool1(
            record.data[10],
            "ParaBuildAtom.fUserSetAnimBackground",
        )?,
        automatic: parse_bool1(record.data[11], "ParaBuildAtom.fAutomatic")?,
        delay_time_ms: read_u32(&record.data, 12),
    })
}

fn parse_level_info_atom(record: &PptRecord) -> Result<u32> {
    require_atom(record, PptRecordType::LevelInfoAtom, 0, 4, "LevelInfoAtom")?;
    let level = read_u32(&record.data, 0);
    if level > 9 {
        return Err(PptError::InvalidFormat(format!(
            "LevelInfoAtom level {level} exceeds 9"
        )));
    }
    Ok(level)
}

fn parse_chart_build(record: &PptRecord) -> Result<ChartBuild> {
    require_container(record, PptRecordType::ChartBuild, 0, "ChartBuild")?;
    if record.children.len() != 2 {
        return Err(PptError::Corrupted(
            "ChartBuild requires exactly BuildAtom and ChartBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Chart)?;
    let chart_record = &record.children[1];
    require_atom(
        chart_record,
        PptRecordType::ChartBuildAtom,
        0,
        8,
        "ChartBuildAtom",
    )?;
    let build_type = ChartBuildType::parse(read_u32(&chart_record.data, 0)).ok_or_else(|| {
        PptError::InvalidFormat(format!(
            "invalid ChartBuildAtom type {}",
            read_u32(&chart_record.data, 0)
        ))
    })?;
    Ok(ChartBuild {
        atom,
        chart: ChartBuildAtom {
            build_type,
            animate_background: parse_bool1(
                chart_record.data[4],
                "ChartBuildAtom.fAnimBackground",
            )?,
        },
    })
}

fn parse_diagram_build(record: &PptRecord) -> Result<DiagramBuild> {
    require_container(record, PptRecordType::DiagramBuild, 0, "DiagramBuild")?;
    if record.children.len() != 2 {
        return Err(PptError::Corrupted(
            "DiagramBuild requires exactly BuildAtom and DiagramBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Diagram)?;
    let diagram_record = &record.children[1];
    require_atom(
        diagram_record,
        PptRecordType::DiagramBuildAtom,
        0,
        4,
        "DiagramBuildAtom",
    )?;
    let build_type =
        DiagramBuildType::parse(read_u32(&diagram_record.data, 0)).ok_or_else(|| {
            PptError::InvalidFormat(format!(
                "invalid DiagramBuildAtom type {}",
                read_u32(&diagram_record.data, 0)
            ))
        })?;
    Ok(DiagramBuild {
        atom,
        diagram: DiagramBuildAtom { build_type },
    })
}

fn require_container(
    record: &PptRecord,
    record_type: PptRecordType,
    instance: u16,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(PptError::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0x0F, instance, None, name)?;
    let encoded_children_length = record.children.iter().try_fold(0usize, |length, child| {
        length.checked_add(8 + child.data.len())
    });
    if encoded_children_length != Some(record.data.len()) {
        return Err(PptError::Corrupted(format!(
            "{name} child records do not cover its complete payload"
        )));
    }
    Ok(())
}

fn require_atom(
    record: &PptRecord,
    record_type: PptRecordType,
    version: u16,
    length: usize,
    name: &str,
) -> Result<()> {
    if record.record_type != record_type {
        return Err(PptError::InvalidFormat(format!(
            "Expected {name}, got {:?}",
            record.record_type
        )));
    }
    require_header(record, version, 0, Some(length), name)
}

fn require_header(
    record: &PptRecord,
    version: u16,
    instance: u16,
    length: Option<usize>,
    name: &str,
) -> Result<()> {
    if record.version != version
        || record.instance != instance
        || record.data_length as usize != record.data.len()
        || length.is_some_and(|length| record.data.len() != length)
    {
        return Err(PptError::Corrupted(format!(
            "invalid {name} header: version {}, instance {}, length {}",
            record.version,
            record.instance,
            record.data.len()
        )));
    }
    Ok(())
}

fn parse_bool1(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(PptError::InvalidFormat(format!(
            "{field} has invalid bool1 value {value}"
        ))),
    }
}

fn parse_optional_time_value<T>(
    is_set: bool,
    value: u32,
    parse: impl FnOnce(u32) -> Option<T>,
    field: &str,
) -> Result<Option<T>> {
    if is_set {
        parse(value)
            .map(Some)
            .ok_or_else(|| PptError::InvalidFormat(format!("{field} has invalid value {value}")))
    } else if value == 0 {
        Ok(None)
    } else {
        Err(PptError::InvalidFormat(format!(
            "{field} must be zero when not explicitly set"
        )))
    }
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(data[offset..offset + 4].try_into().expect("length checked"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ppt::animation::{
        BuildInfo, ChartBuildType, DiagramBuildType, LegacyAnimationAtom, LegacyAnimationBuild,
        LegacyAnimationEffect, LegacyTextBuildSubEffect, ParagraphBuildType, write_animation_info,
        write_animation_info_atom, write_build_list, write_extended_time_node,
        write_time_node_atom,
    };

    fn sample_legacy_atom() -> LegacyAnimationAtom {
        LegacyAnimationAtom {
            dim_color: 0x0011_2233,
            reverse: true,
            automatic: true,
            has_sound: true,
            stop_sound: true,
            play: true,
            synchronous: true,
            hide_while_not_playing: true,
            animate_background: true,
            sound_id_ref: 42,
            delay_time_ms: 750,
            order_id: -2,
            slide_count: 3,
            build_type: LegacyAnimationBuild::Level3,
            effect: LegacyAnimationEffect::Fly,
            effect_direction: 0x1C,
            after_effect: AfterEffect::HideOnNextClick,
            text_build_sub_effect: LegacyTextBuildSubEffect::ByCharacter,
            ole_verb: 2,
        }
    }

    fn empty_time_node() -> ExtendedTimeNode {
        ExtendedTimeNode {
            atom: TimeNodeAtom::default(),
            children: Vec::new(),
        }
    }

    fn sample_build_list() -> BuildList {
        BuildList {
            builds: vec![
                BuildListEntry::Paragraph(ParagraphBuild {
                    atom: BuildAtom {
                        build_id: 10,
                        shape_id_ref: 100,
                        expanded: true,
                        ui_expanded: false,
                    },
                    paragraph: ParagraphBuildAtom {
                        build_type: ParagraphBuildType::AsAWhole,
                        build_level: 4,
                        animate_background: true,
                        reverse: true,
                        user_set_animate_background: true,
                        automatic: true,
                        delay_time_ms: 750,
                    },
                    levels: vec![ParagraphBuildLevel {
                        level: 0,
                        time_node: empty_time_node(),
                    }],
                }),
                BuildListEntry::Chart(ChartBuild {
                    atom: BuildAtom {
                        build_id: 11,
                        shape_id_ref: 101,
                        expanded: false,
                        ui_expanded: true,
                    },
                    chart: ChartBuildAtom {
                        build_type: ChartBuildType::ByElementInCategory,
                        animate_background: true,
                    },
                }),
                BuildListEntry::Diagram(DiagramBuild {
                    atom: BuildAtom {
                        build_id: 12,
                        shape_id_ref: 102,
                        expanded: true,
                        ui_expanded: true,
                    },
                    diagram: DiagramBuildAtom {
                        build_type: DiagramBuildType::CounterClockwiseOut,
                    },
                }),
            ],
        }
    }

    #[test]
    fn round_trips_exact_animation_info_atoms_and_containers() {
        let atom = sample_legacy_atom();
        let bytes = write_animation_info_atom(&atom).unwrap();
        assert_eq!(bytes.len(), 36);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_animation_info_atom(&record).unwrap(), atom);

        let mut info = AnimationInfo::new();
        info.legacy_atom = Some(atom.clone());
        let (container, sound_ref) = write_animation_info(&info).unwrap();
        assert_eq!(sound_ref, 42);
        let (record, consumed) = PptRecord::parse(&container, 0).unwrap();
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
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert!(
                parse_animation_info_atom(&record).is_err(),
                "accepted mutation at byte {offset}"
            );
        }

        let mut short = valid;
        short[4..8].copy_from_slice(&27u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&short, 0).unwrap();
        assert!(parse_animation_info_atom(&record).is_err());
    }

    #[test]
    fn round_trips_exact_time_node_atoms_and_envelopes() {
        let atom = TimeNodeAtom {
            fill: Some(TimeNodeFill::ResetWhenInactiveLegacy),
            restart: Some(TimeNodeRestart::NeverLegacy),
            node_type: Some(TimeNodeKind::Behavior),
            duration_ms: Some(-1),
        };
        let bytes = write_time_node_atom(&atom);
        assert_eq!(bytes.len(), 40);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_time_node_atom(&record).unwrap(), atom);

        let raw_child = PptRecord {
            record_type: PptRecordType::Unknown,
            record_type_raw: 0xF999,
            version: 0,
            instance: 7,
            data_length: 3,
            data: vec![1, 2, 3],
            children: Vec::new(),
        };
        let node = ExtendedTimeNode {
            atom,
            children: vec![raw_child],
        };
        let bytes = write_extended_time_node(&node).unwrap();
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(parse_extended_time_node(&record).unwrap(), node);
    }

    #[test]
    fn rejects_malformed_time_node_atoms() {
        let default = write_time_node_atom(&TimeNodeAtom::default());
        let mutations: &[(usize, u32)] = &[
            (12, 1), // restart value without fRestartProperty
            (16, 1), // type value without fGroupingTypeProperty
            (20, 1), // fill value without fFillProperty
            (32, 1), // duration without fDurationProperty
        ];
        for &(offset, value) in mutations {
            let mut bytes = default.clone();
            bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
            assert!(parse_time_node_atom(&record).is_err());
        }

        let mut invalid_enum = default;
        invalid_enum[20..24].copy_from_slice(&5u32.to_le_bytes());
        invalid_enum[36..40].copy_from_slice(&1u32.to_le_bytes());
        let (record, _) = PptRecord::parse(&invalid_enum, 0).unwrap();
        assert!(parse_time_node_atom(&record).is_err());
    }

    #[test]
    fn supports_every_time_node_atom_enum_value_and_ignores_reserved_fields() {
        let fills = [
            TimeNodeFill::HoldUntilParentEnds,
            TimeNodeFill::ResetWhenInactive,
            TimeNodeFill::HoldUntilNext,
            TimeNodeFill::HoldUntilParentEndsLegacy,
            TimeNodeFill::ResetWhenInactiveLegacy,
        ];
        let restarts = [
            TimeNodeRestart::Never,
            TimeNodeRestart::Always,
            TimeNodeRestart::WhenNotActive,
            TimeNodeRestart::NeverLegacy,
        ];
        let kinds = [
            TimeNodeKind::Parallel,
            TimeNodeKind::Sequential,
            TimeNodeKind::Behavior,
            TimeNodeKind::Media,
        ];
        for fill in fills {
            for restart in restarts {
                for node_type in kinds {
                    let expected = TimeNodeAtom {
                        fill: Some(fill),
                        restart: Some(restart),
                        node_type: Some(node_type),
                        duration_ms: Some(i32::MIN),
                    };
                    let mut bytes = write_time_node_atom(&expected);
                    bytes[8..12].copy_from_slice(&u32::MAX.to_le_bytes());
                    bytes[24..28].copy_from_slice(&u32::MAX.to_le_bytes());
                    bytes[28] = 0xFF;
                    let flags = u32::from_le_bytes(bytes[36..40].try_into().unwrap()) | 0xFFFF_FFE0;
                    bytes[36..40].copy_from_slice(&flags.to_le_bytes());
                    let (record, _) = PptRecord::parse(&bytes, 0).unwrap();
                    assert_eq!(parse_time_node_atom(&record).unwrap(), expected);
                }
            }
        }
    }

    #[test]
    fn round_trips_exact_powerpoint_2002_build_lists() {
        assert_eq!(PptRecordType::BuildList.as_u16(), 0x2B02);
        assert_eq!(PptRecordType::LevelInfoAtom.as_u16(), 0x2B0A);
        let bytes = write_build_list(&sample_build_list()).unwrap();
        assert_eq!(PptRecordType::TimeNode.as_u16(), 0xF127);
        assert_eq!(bytes.len(), 216);
        let (record, consumed) = PptRecord::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        assert_eq!(record.record_type, PptRecordType::BuildList);
        assert_eq!(record.children.len(), 3);

        let parsed = parse_build_list(&record).unwrap();
        assert_eq!(parsed.builds.len(), 3);
        let BuildListEntry::Paragraph(paragraph) = &parsed.builds[0] else {
            panic!("expected paragraph build");
        };
        assert_eq!(paragraph.atom.build_id, 10);
        assert_eq!(paragraph.atom.shape_id_ref, 100);
        assert_eq!(paragraph.paragraph.build_type, ParagraphBuildType::AsAWhole);
        assert_eq!(paragraph.paragraph.delay_time_ms, 750);
        assert_eq!(paragraph.levels.len(), 1);
        assert_eq!(paragraph.levels[0].level, 0);
        assert_eq!(paragraph.levels[0].time_node.atom, TimeNodeAtom::default());

        let BuildListEntry::Chart(chart) = &parsed.builds[1] else {
            panic!("expected chart build");
        };
        assert_eq!(chart.chart.build_type, ChartBuildType::ByElementInCategory);
        assert!(chart.chart.animate_background);

        let BuildListEntry::Diagram(diagram) = &parsed.builds[2] else {
            panic!("expected diagram build");
        };
        assert_eq!(
            diagram.diagram.build_type,
            DiagramBuildType::CounterClockwiseOut
        );
    }

    #[test]
    fn rejects_malformed_powerpoint_2002_build_lists() {
        let bytes = write_build_list(&sample_build_list()).unwrap();
        let (valid, _) = PptRecord::parse(&bytes, 0).unwrap();

        let mut truncated = bytes.clone();
        let claimed_length = u32::from_le_bytes(truncated[4..8].try_into().unwrap()) + 1;
        truncated[4..8].copy_from_slice(&claimed_length.to_le_bytes());
        let (truncated, _) = PptRecord::parse(&truncated, 0).unwrap();
        assert_eq!(truncated.data_length, claimed_length);
        assert!(parse_build_list(&truncated).is_err());

        let mut malformed = Vec::new();
        let mut wrong_header = valid.clone();
        wrong_header.version = 0;
        malformed.push(wrong_header);

        let mut wrong_bool = valid.clone();
        wrong_bool.children[1].children[1].data[4] = 2;
        malformed.push(wrong_bool);

        let mut wrong_kind = valid.clone();
        wrong_kind.children[2].children[0].data[0..4].copy_from_slice(&1u32.to_le_bytes());
        malformed.push(wrong_kind);

        let mut wrong_level = valid.clone();
        wrong_level.children[0].children[2].data[0..4].copy_from_slice(&10u32.to_le_bytes());
        malformed.push(wrong_level);

        let mut duplicate = valid.clone();
        duplicate.children[2].children[0].data[4..8].copy_from_slice(&11u32.to_le_bytes());
        duplicate.children[2].children[0].data[8..12].copy_from_slice(&101u32.to_le_bytes());
        malformed.push(duplicate);

        let mut wrong_order = valid.clone();
        wrong_order.children[1].children.swap(0, 1);
        malformed.push(wrong_order);

        for record in malformed {
            assert!(parse_build_list(&record).is_err());
        }
    }

    #[test]
    fn test_animation_info_default() {
        let info = AnimationInfo::default();
        assert!(!info.has_animations());
        assert_eq!(info.animation_count(), 0);
    }

    #[test]
    fn test_build_info_default() {
        let build_info = BuildInfo::default();
        assert!(build_info.builds.is_empty());
    }
}
