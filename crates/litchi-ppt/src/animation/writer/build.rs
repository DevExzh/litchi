//! `PowerPoint` 2002 build-list records and legacy effect mapping.

use super::support::{create_record_header, wrap_record};
use super::time_node::write_extended_time_node;
use crate::animation::types::{
    AnimationEffect, BuildAtom, BuildKind, BuildList, BuildListEntry, ChartBuild, DiagramBuild,
    EffectDirection, ParagraphBuild, ParagraphBuildLevel,
};
use crate::consts::RecordType;
use crate::package::{Error, Result};

/// Map a high-level animation effect to PPT97 fly method and direction codes.
/// Based on `LibreOffice` ppt97animations.cxx mapping.
#[allow(
    clippy::match_same_arms,
    clippy::wildcard_enum_match_arm,
    reason = "the arms mirror the effect categories of the ppt97animations.cxx mapping; identical codes are deliberate per-category defaults and each wildcard arm encodes the documented fallback direction for that effect"
)]
pub(super) fn map_effect_to_ppt97(effect: AnimationEffect, direction: EffectDirection) -> (u8, u8) {
    use AnimationEffect::{
        Appear, Ascend, Blinds, BlindsOut, BoldFlash, Bounce, Box, BoxOut, ChangeFillColor,
        ChangeFontColor, ChangeFontSize, ChangeLineColor, Checkerboard, CheckerboardOut, Collapse,
        ColorPulse, ComplementaryColor, ComplementaryColor2, Compress, ContrastingColor, CrawlIn,
        CrawlOut, Custom, Darken, Descend, DescendOut, Diamond, DiamondOut, Disappear, Dissolve,
        Expand, FadeIn, FadeOut, Flicker, FloatIn, FlyIn, FlyOut, GrowAndTurn, GrowShrink, Lighten,
        MotionPath, MotionPathArcDown, MotionPathArcUp, MotionPathCircle, MotionPathCurvedX,
        MotionPathCurves, MotionPathDiagonalDownRight, MotionPathDiagonalUpRight,
        MotionPathDiamond, MotionPathDown, MotionPathHeart, MotionPathHexagon, MotionPathLeft,
        MotionPathLines, MotionPathLoopDeLoop, MotionPathOctagon, MotionPathPentagon,
        MotionPathRight, MotionPathSCurve1, MotionPathSCurve2, MotionPathShapes,
        MotionPathSineWave, MotionPathSpiralLeft, MotionPathSpiralRight, MotionPathSpring,
        MotionPathSquare, MotionPathStar4, MotionPathStar5, MotionPathStar6, MotionPathStar8,
        MotionPathTriangle, MotionPathUp, MotionPathZigzag, ObjectColor, PeekIn, PeekOut, Plus,
        PlusOut, Pulse, Random, RandomBars, RandomBarsOut, RiseUp, SinkDown, Spin, SpiralIn,
        SpiralOut, Split, SplitOut, Stretch, Strips, StripsOut, Swivel, Teeter, Transparency,
        Underline, VerticalHighlight, Wave, Wedge, Wheel, Wipe, WipeOut, Zoom,
    };
    use EffectDirection::{
        FromBottom, FromBottomLeft, FromBottomRight, FromLeft, FromRight, FromTop, FromTopLeft,
        FromTopRight, Horizontal, In, Out, Vertical,
    };

    match effect {
        // Entrance effects
        Appear => (0x00, 0),
        FadeIn => (0x06, 0),
        FlyIn => match direction {
            FromLeft => (0x0c, 0x00),
            FromTop => (0x0c, 0x01),
            FromRight => (0x0c, 0x02),
            FromBottom => (0x0c, 0x03),
            FromTopLeft => (0x0c, 0x04),
            FromTopRight => (0x0c, 0x05),
            FromBottomLeft => (0x0c, 0x06),
            FromBottomRight => (0x0c, 0x07),
            _ => (0x0c, 0x00),
        },
        Wipe => match direction {
            FromRight => (0x0a, 0x00),
            FromBottom => (0x0a, 0x01),
            FromLeft => (0x0a, 0x02),
            FromTop => (0x0a, 0x03),
            _ => (0x0a, 0x00),
        },
        Split => (0x0d, 0),
        Dissolve => (0x05, 0),
        Box => match direction {
            Out => (0x0b, 0x00),
            In => (0x0b, 0x01),
            _ => (0x0b, 0x00),
        },
        Checkerboard => match direction {
            Horizontal => (0x03, 0x00),
            Vertical => (0x03, 0x01),
            _ => (0x03, 0x00),
        },
        Blinds => match direction {
            Horizontal => (0x02, 0x00),
            Vertical => (0x02, 0x01),
            _ => (0x02, 0x00),
        },
        RandomBars => match direction {
            Horizontal => (0x08, 0x00),
            Vertical => (0x08, 0x01),
            _ => (0x08, 0x00),
        },
        GrowAndTurn => (0x00, 0),
        // Zoom sub-effects per ppt97animations.cxx:
        // 0x10=zoom-in, 0x11=zoom-in-slightly, 0x12=zoom-out,
        // 0x13=zoom-out-slightly, 0x14=from-screen-center, 0x15=out-from-screen-center
        Zoom => match direction {
            In => (0x0c, 0x10),
            Out => (0x0c, 0x12),
            _ => (0x0c, 0x10),
        },
        Expand => (0x0c, 0x10),   // zoom-in
        Compress => (0x0c, 0x12), // zoom-out
        // Stretch sub-effects: 0x16=across, 0x17=from-left, 0x18=from-top,
        // 0x19=from-right, 0x1a=from-bottom
        Stretch => match direction {
            FromLeft => (0x0c, 0x17),
            FromTop => (0x0c, 0x18),
            FromRight => (0x0c, 0x19),
            FromBottom => (0x0c, 0x1a),
            _ => (0x0c, 0x16),
        },
        // Swivel: 0x1b=vertical
        Swivel => (0x0c, 0x1b),
        // SpiralIn: 0x1c
        SpiralIn => (0x0c, 0x1c),
        Bounce => (0x00, 0),
        // PeekIn sub-effects: 0x08=from-left, 0x09=from-bottom, 0x0a=from-right, 0x0b=from-top
        PeekIn => match direction {
            FromLeft => (0x0c, 0x08),
            FromBottom => (0x0c, 0x09),
            FromRight => (0x0c, 0x0a),
            FromTop => (0x0c, 0x0b),
            _ => (0x0c, 0x08),
        },
        // CrawlIn = slow fly: 0x0c=from-left, 0x0d=from-top, 0x0e=from-right, 0x0f=from-bottom
        CrawlIn => match direction {
            FromLeft => (0x0c, 0x0c),
            FromTop => (0x0c, 0x0d),
            FromRight => (0x0c, 0x0e),
            FromBottom => (0x0c, 0x0f),
            _ => (0x0c, 0x0c),
        },
        FloatIn | Ascend => (0x0c, 0x03), // fly from bottom
        Descend => (0x0c, 0x01),          // fly from top
        RiseUp => (0x0c, 0x03),           // fly from bottom
        Random => (0x01, 0),              // random
        Wheel => (0x1a, 1),
        Plus => (0x12, 0),
        Diamond => (0x11, 0),
        Wedge => (0x13, 0),
        Strips => (0x09, 4),

        // Emphasis effects (map to appear as PPT97 doesn't have these)
        Pulse | Spin | Teeter | Wave | Lighten | Darken => (0x00, 0),
        ChangeFillColor | ChangeLineColor | ChangeFontColor | ChangeFontSize => (0x00, 0),
        GrowShrink | BoldFlash | Underline | ColorPulse => (0x00, 0),
        ComplementaryColor | ComplementaryColor2 | ContrastingColor => (0x00, 0),
        Transparency | ObjectColor | VerticalHighlight | Flicker => (0x00, 0),

        // Exit effects (reverse of entrance)
        FadeOut | Disappear => (0x00, 0),
        FlyOut | WipeOut | BoxOut | CheckerboardOut => (0x00, 0),
        BlindsOut | RandomBarsOut | StripsOut | SplitOut => (0x00, 0),
        PeekOut | PlusOut | DiamondOut | CrawlOut => (0x00, 0),
        DescendOut | Collapse | SinkDown | SpiralOut => (0x00, 0),

        // Motion paths (not supported in PPT97)
        MotionPath | MotionPathLines | MotionPathCurves | MotionPathShapes => (0x00, 0),
        MotionPathLeft | MotionPathRight | MotionPathUp | MotionPathDown => (0x00, 0),
        MotionPathDiagonalUpRight | MotionPathDiagonalDownRight => (0x00, 0),
        MotionPathArcDown | MotionPathArcUp | MotionPathCircle => (0x00, 0),
        MotionPathDiamond | MotionPathHeart | MotionPathHexagon => (0x00, 0),
        MotionPathOctagon | MotionPathPentagon | MotionPathSquare => (0x00, 0),
        MotionPathStar4 | MotionPathStar5 | MotionPathStar6 | MotionPathStar8 => (0x00, 0),
        MotionPathTriangle | MotionPathLoopDeLoop | MotionPathCurvedX => (0x00, 0),
        MotionPathSCurve1 | MotionPathSCurve2 | MotionPathSineWave => (0x00, 0),
        MotionPathSpiralLeft | MotionPathSpiralRight | MotionPathSpring => (0x00, 0),
        MotionPathZigzag => (0x00, 0),

        Custom => (0x00, 0),
    }
}

/// Write `BuildList` container record.
///
/// # Errors
///
/// Returns an error if two builds share a build identity, a build entry fails
/// validation or serialization, or the record exceeds the 4 GiB limit.
pub fn write_build_list(build_info: &BuildList) -> Result<Vec<u8>> {
    let mut identities = std::collections::HashSet::with_capacity(build_info.builds.len());
    let mut children = Vec::new();
    for build in &build_info.builds {
        let atom = match build {
            BuildListEntry::Paragraph(paragraph_build) => &paragraph_build.atom,
            BuildListEntry::Chart(chart_build) => &chart_build.atom,
            BuildListEntry::Diagram(diagram_build) => &diagram_build.atom,
        };
        if !identities.insert((atom.build_id, atom.shape_id_ref)) {
            return Err(Error::InvalidFormat(format!(
                "duplicate build identity ({}, {})",
                atom.build_id, atom.shape_id_ref
            )));
        }
        children.extend(match build {
            BuildListEntry::Paragraph(paragraph_build) => write_paragraph_build(paragraph_build)?,
            BuildListEntry::Chart(chart_build) => write_chart_build(chart_build)?,
            BuildListEntry::Diagram(diagram_build) => write_diagram_build(diagram_build)?,
        });
    }
    wrap_record(RecordType::BuildList, 0x0F, 0, children)
}

fn write_build_atom(atom: &BuildAtom, kind: BuildKind) -> Vec<u8> {
    let mut data = Vec::with_capacity(16);
    data.extend(kind.as_u32().to_le_bytes());
    data.extend(atom.build_id.to_le_bytes());
    data.extend(atom.shape_id_ref.to_le_bytes());
    data.push(u8::from(atom.expanded));
    data.push(u8::from(atom.ui_expanded));
    data.extend([0, 0]);
    let mut result = create_record_header(RecordType::BuildAtom, 0, 0, 16);
    result.extend(data);
    result
}

fn write_paragraph_build(build: &ParagraphBuild) -> Result<Vec<u8>> {
    validate_paragraph_levels(&build.paragraph.build_type, &build.levels)?;
    let mut children = write_build_atom(&build.atom, BuildKind::Paragraph);
    let mut atom = Vec::with_capacity(16);
    atom.extend(build.paragraph.build_type.as_u32().to_le_bytes());
    atom.extend(build.paragraph.build_level.to_le_bytes());
    atom.push(u8::from(build.paragraph.animate_background));
    atom.push(u8::from(build.paragraph.reverse));
    atom.push(u8::from(build.paragraph.user_set_animate_background));
    atom.push(u8::from(build.paragraph.automatic));
    atom.extend(build.paragraph.delay_time_ms.to_le_bytes());
    children.extend(create_record_header(RecordType::ParaBuildAtom, 1, 0, 16));
    children.extend(atom);
    for level in &build.levels {
        children.extend(create_record_header(RecordType::LevelInfoAtom, 0, 0, 4));
        children.extend(level.level.to_le_bytes());
        children.extend(write_extended_time_node(&level.time_node)?);
    }
    wrap_record(RecordType::ParaBuild, 0x0F, 0, children)
}

fn write_chart_build(build: &ChartBuild) -> Result<Vec<u8>> {
    let mut children = write_build_atom(&build.atom, BuildKind::Chart);
    let mut atom = Vec::with_capacity(8);
    atom.extend(build.chart.build_type.as_u32().to_le_bytes());
    atom.push(u8::from(build.chart.animate_background));
    atom.extend([0, 0, 0]);
    children.extend(create_record_header(RecordType::ChartBuildAtom, 0, 0, 8));
    children.extend(atom);
    wrap_record(RecordType::ChartBuild, 0x0F, 0, children)
}

fn write_diagram_build(build: &DiagramBuild) -> Result<Vec<u8>> {
    let mut children = write_build_atom(&build.atom, BuildKind::Diagram);
    children.extend(create_record_header(RecordType::DiagramBuildAtom, 0, 0, 4));
    children.extend(build.diagram.build_type.as_u32().to_le_bytes());
    wrap_record(RecordType::DiagramBuild, 0x0F, 0, children)
}

#[allow(
    clippy::trivially_copy_pass_by_ref,
    reason = "callers validate a borrowed field of the build struct; taking the build type by reference keeps the validator consistent with the other reference-taking validators in this module"
)]
pub(super) fn validate_paragraph_levels(
    build_type: &crate::animation::types::ParagraphBuildType,
    levels: &[ParagraphBuildLevel],
) -> Result<()> {
    if levels.is_empty() {
        return Err(Error::InvalidFormat(
            "ParaBuild requires at least one level".to_string(),
        ));
    }
    if *build_type == crate::animation::types::ParagraphBuildType::AsAWhole && levels.len() != 1 {
        return Err(Error::InvalidFormat(
            "AsAWhole ParaBuild requires exactly one level".to_string(),
        ));
    }
    for (index, level) in levels.iter().enumerate() {
        if level.level > 9 {
            return Err(Error::InvalidFormat(format!(
                "paragraph build level {} exceeds 9",
                level.level
            )));
        }
        if index > 0 && levels[index - 1].level >= level.level {
            return Err(Error::InvalidFormat(
                "ParaBuild levels must be strictly increasing".to_string(),
            ));
        }
    }
    Ok(())
}
