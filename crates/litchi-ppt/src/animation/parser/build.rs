//! `PowerPoint` 2002 build-list records.

use super::support::{parse_bool1, read_u32, require_atom, require_container, require_header};
use super::time_node::parse_extended_time_node;
use crate::animation::types::{
    BuildAtom, BuildKind, BuildList, BuildListEntry, ChartBuild, ChartBuildAtom, ChartBuildType,
    DiagramBuild, DiagramBuildAtom, DiagramBuildType, ParagraphBuild, ParagraphBuildAtom,
    ParagraphBuildLevel, ParagraphBuildType,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Parse build list from `BuildList` container record.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
#[allow(
    clippy::wildcard_enum_match_arm,
    reason = "`RecordType` spans the full MS-PPT record ID space; every child other than the three build containers is rejected uniformly"
)]
pub fn parse_build_list(record: &Record) -> Result<BuildList> {
    if record.record_type != RecordType::BuildList {
        return Err(Error::InvalidFormat(format!(
            "Expected BuildList record, got {:?}",
            record.record_type
        )));
    }

    require_container(record, RecordType::BuildList, 0, "BuildList")?;
    let mut build_info = BuildList::new();
    let mut identities = std::collections::HashSet::with_capacity(record.children.len());
    for child in &record.children {
        let build = match child.record_type {
            RecordType::ParaBuild => BuildListEntry::Paragraph(parse_paragraph_build(child)?),
            RecordType::ChartBuild => BuildListEntry::Chart(parse_chart_build(child)?),
            RecordType::DiagramBuild => BuildListEntry::Diagram(parse_diagram_build(child)?),
            other => {
                return Err(Error::InvalidFormat(format!(
                    "BuildList contains invalid child {other:?}"
                )));
            },
        };
        let atom = match &build {
            BuildListEntry::Paragraph(paragraph) => &paragraph.atom,
            BuildListEntry::Chart(chart) => &chart.atom,
            BuildListEntry::Diagram(diagram) => &diagram.atom,
        };
        if !identities.insert((atom.build_id, atom.shape_id_ref)) {
            return Err(Error::InvalidFormat(format!(
                "duplicate build identity ({}, {})",
                atom.build_id, atom.shape_id_ref
            )));
        }
        build_info.add_build(build);
    }
    Ok(build_info)
}

fn parse_build_atom(record: &Record, expected: BuildKind) -> Result<BuildAtom> {
    if record.record_type != RecordType::BuildAtom {
        return Err(Error::InvalidFormat(format!(
            "Expected BuildAtom, got {:?}",
            record.record_type
        )));
    }
    require_header(record, 0, 0, Some(16), "BuildAtom")?;
    let kind = BuildKind::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "invalid BuildAtom build type {}",
            read_u32(&record.data, 0)
        ))
    })?;
    if kind != expected {
        return Err(Error::InvalidFormat(format!(
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

fn parse_paragraph_build(record: &Record) -> Result<ParagraphBuild> {
    require_container(record, RecordType::ParaBuild, 0, "ParaBuild")?;
    if record.children.len() < 4 || !(record.children.len() - 2).is_multiple_of(2) {
        return Err(Error::Corrupted(
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
            return Err(Error::InvalidFormat(
                "ParaBuild levels must be strictly increasing".to_string(),
            ));
        }
        levels.push(ParagraphBuildLevel { level, time_node });
    }
    if paragraph.build_type == ParagraphBuildType::AsAWhole && levels.len() != 1 {
        return Err(Error::InvalidFormat(
            "AsAWhole ParaBuild requires exactly one level".to_string(),
        ));
    }
    Ok(ParagraphBuild {
        atom,
        paragraph,
        levels,
    })
}

fn parse_paragraph_build_atom(record: &Record) -> Result<ParagraphBuildAtom> {
    require_atom(record, RecordType::ParaBuildAtom, 1, 16, "ParaBuildAtom")?;
    let build_type = ParagraphBuildType::parse(read_u32(&record.data, 0)).ok_or_else(|| {
        Error::InvalidFormat(format!(
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

fn parse_level_info_atom(record: &Record) -> Result<u32> {
    require_atom(record, RecordType::LevelInfoAtom, 0, 4, "LevelInfoAtom")?;
    let level = read_u32(&record.data, 0);
    if level > 9 {
        return Err(Error::InvalidFormat(format!(
            "LevelInfoAtom level {level} exceeds 9"
        )));
    }
    Ok(level)
}

fn parse_chart_build(record: &Record) -> Result<ChartBuild> {
    require_container(record, RecordType::ChartBuild, 0, "ChartBuild")?;
    if record.children.len() != 2 {
        return Err(Error::Corrupted(
            "ChartBuild requires exactly BuildAtom and ChartBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Chart)?;
    let chart_record = &record.children[1];
    require_atom(
        chart_record,
        RecordType::ChartBuildAtom,
        0,
        8,
        "ChartBuildAtom",
    )?;
    let build_type = ChartBuildType::parse(read_u32(&chart_record.data, 0)).ok_or_else(|| {
        Error::InvalidFormat(format!(
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

fn parse_diagram_build(record: &Record) -> Result<DiagramBuild> {
    require_container(record, RecordType::DiagramBuild, 0, "DiagramBuild")?;
    if record.children.len() != 2 {
        return Err(Error::Corrupted(
            "DiagramBuild requires exactly BuildAtom and DiagramBuildAtom".to_string(),
        ));
    }
    let atom = parse_build_atom(&record.children[0], BuildKind::Diagram)?;
    let diagram_record = &record.children[1];
    require_atom(
        diagram_record,
        RecordType::DiagramBuildAtom,
        0,
        4,
        "DiagramBuildAtom",
    )?;
    let build_type =
        DiagramBuildType::parse(read_u32(&diagram_record.data, 0)).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "invalid DiagramBuildAtom type {}",
                read_u32(&diagram_record.data, 0)
            ))
        })?;
    Ok(DiagramBuild {
        atom,
        diagram: DiagramBuildAtom { build_type },
    })
}
