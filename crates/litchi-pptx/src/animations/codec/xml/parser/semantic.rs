use super::super::super::super::invalid;
use super::super::super::super::model::{
    DiagramBuild, DiagramBuildType, Duration, Effect, EffectInstance, EventFilter, Fill,
    GraphicBuild, GraphicBuildMode, GraphicChartBuildType, GraphicDiagramBuildType, GroupId,
    OleChartBuild, OleChartBuildType, ParagraphBuild, ParagraphBuildType, ParagraphTemplate,
    Repeat, Restart, Sequence, SequenceContext, Speed, SyncBehavior, TemplateTimeNode, TimeFilter,
};
use super::super::semantic::{
    PendingAnimation, PendingGraphicBuild, PendingParagraphTemplate, TimeNodeFrame,
    parse_sequence_context, trigger,
};
use super::validation::{
    DIAGRAM_NS, MAX_ANIMATION_BUILDS, MAX_ANIMATIONS, MAX_PARAGRAPH_TEMPLATES,
    MAX_PRESERVED_TIMING_BYTES, MAX_TEMPLATE_TIME_NODE_BYTES, TimingValue, attribute,
    check_attribute_count, direction_from_subtype, is_chartml_name, is_drawingml_name,
    is_known_ole_chart_program_id, is_namespace_name, is_presentationml_name,
    parse_build_auto_advance, parse_group_id, parse_progress, parse_shape_id, parse_timing_value,
    parse_xml_bool,
};
use crate::{Error, Result};
use quick_xml::encoding::Decoder;
use quick_xml::events::BytesStart;
use quick_xml::name::{QName, ResolveResult};
use std::collections::HashSet;
use std::ops::Range;

pub(super) struct TimingParser {
    sequence: Sequence,
    shape_ids: HashSet<u32>,
    time_nodes: Vec<TimeNodeFrame>,
    pending: Vec<PendingAnimation>,
    timing_depth: Option<usize>,
    start_conditions_depth: Vec<usize>,
    condition_depth: Vec<usize>,
    condition_target_depth: Option<usize>,
    build_list_depth: Option<usize>,
    saw_build_list: bool,
    timing_group_ids: HashSet<GroupId>,
    build_group_ids: HashSet<GroupId>,
    build_pairs: HashSet<(u8, u32, GroupId)>,
    paragraph_build_depth: Option<usize>,
    paragraph_build_index: Option<usize>,
    template_list_depth: Option<usize>,
    template_levels: HashSet<u8>,
    pending_template: Option<PendingParagraphTemplate>,
    template_ranges: Vec<(usize, u8, Range<usize>)>,
    diagram_build_depth: Option<usize>,
    ole_chart_build_depth: Option<usize>,
    pending_graphic_build: Option<PendingGraphicBuild>,
    graphic_frame_depth: Option<usize>,
    graphic_depth: Option<usize>,
    graphic_data_depth: Option<usize>,
    graphic_frame_shape_id: Option<u32>,
    graphic_frame_has_ole_object: bool,
    graphic_frame_has_ole_chart: bool,
    graphic_frame_has_chart: bool,
    graphic_frame_has_diagram: bool,
    ole_diagram_shape_ids: HashSet<u32>,
    ole_chart_shape_ids: HashSet<u32>,
    chart_shape_ids: HashSet<u32>,
    graphical_diagram_shape_ids: HashSet<u32>,
    saw_timing: bool,
    require_valid_targets: bool,
    timing_start: Option<usize>,
    timing_range: Option<Range<usize>>,
}

impl TimingParser {
    pub(super) fn new(require_valid_targets: bool) -> Self {
        Self {
            sequence: Sequence::new(),
            shape_ids: HashSet::new(),
            time_nodes: Vec::new(),
            pending: Vec::new(),
            timing_depth: None,
            start_conditions_depth: Vec::new(),
            condition_depth: Vec::new(),
            condition_target_depth: None,
            build_list_depth: None,
            saw_build_list: false,
            timing_group_ids: HashSet::new(),
            build_group_ids: HashSet::new(),
            build_pairs: HashSet::new(),
            paragraph_build_depth: None,
            paragraph_build_index: None,
            template_list_depth: None,
            template_levels: HashSet::new(),
            pending_template: None,
            template_ranges: Vec::new(),
            diagram_build_depth: None,
            ole_chart_build_depth: None,
            pending_graphic_build: None,
            graphic_frame_depth: None,
            graphic_depth: None,
            graphic_data_depth: None,
            graphic_frame_shape_id: None,
            graphic_frame_has_ole_object: false,
            graphic_frame_has_ole_chart: false,
            graphic_frame_has_chart: false,
            graphic_frame_has_diagram: false,
            ole_diagram_shape_ids: HashSet::new(),
            ole_chart_shape_ids: HashSet::new(),
            chart_shape_ids: HashSet::new(),
            graphical_diagram_shape_ids: HashSet::new(),
            saw_timing: false,
            require_valid_targets,
            timing_start: None,
            timing_range: None,
        }
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "parser frame carries one slot per XML element field"
    )]
    pub(super) fn start(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: Decoder,
        depth: usize,
        empty: bool,
        event_start: usize,
        event_end: usize,
    ) -> Result<()> {
        check_attribute_count(element)?;

        if self.require_valid_targets
            && is_presentationml_name(namespace, element.name(), b"graphicFrame")
            && !empty
        {
            if self.graphic_frame_depth.is_some() {
                return Err(invalid("nested graphic frames are not supported"));
            }
            self.graphic_frame_depth = Some(depth);
            self.graphic_depth = None;
            self.graphic_data_depth = None;
            self.graphic_frame_shape_id = None;
            self.graphic_frame_has_ole_object = false;
            self.graphic_frame_has_ole_chart = false;
            self.graphic_frame_has_chart = false;
            self.graphic_frame_has_diagram = false;
        }

        if self.require_valid_targets
            && is_presentationml_name(namespace, element.name(), b"cNvPr")
            && let Some(value) = attribute(element, b"id", decoder)?
        {
            let id = parse_shape_id(&value)?;
            if self.graphic_frame_depth.is_some() {
                if self.graphic_frame_shape_id.is_none() {
                    self.graphic_frame_shape_id = Some(id);
                    if !self.shape_ids.insert(id) {
                        return Err(invalid("duplicate shape ID in slide"));
                    }
                }
            } else if !self.shape_ids.insert(id) {
                return Err(invalid("duplicate shape ID in slide"));
            }
        }

        if self.require_valid_targets
            && self
                .graphic_data_depth
                .is_some_and(|data_depth| depth == data_depth + 1)
            && is_presentationml_name(namespace, element.name(), b"oleObj")
        {
            if self.graphic_frame_has_ole_object {
                return Err(invalid("graphic frame has multiple direct OLE objects"));
            }
            self.graphic_frame_has_ole_object = true;
            self.graphic_frame_has_ole_chart = attribute(element, b"progId", decoder)?
                .as_deref()
                .is_none_or(is_known_ole_chart_program_id);
        }

        if self.require_valid_targets
            && self
                .graphic_frame_depth
                .is_some_and(|frame_depth| depth == frame_depth + 1)
            && is_drawingml_name(namespace, element.name(), b"graphic")
        {
            if self.graphic_depth.is_some() {
                return Err(invalid("graphic frame has multiple direct graphic hosts"));
            }
            if !empty {
                self.graphic_depth = Some(depth);
            }
        }

        if self.require_valid_targets
            && self
                .graphic_depth
                .is_some_and(|graphic_depth| depth == graphic_depth + 1)
            && is_drawingml_name(namespace, element.name(), b"graphicData")
        {
            if self.graphic_data_depth.is_some() {
                return Err(invalid(
                    "graphic host has multiple direct graphic-data elements",
                ));
            }
            if !empty {
                self.graphic_data_depth = Some(depth);
            }
        }

        if self.require_valid_targets
            && self
                .graphic_data_depth
                .is_some_and(|data_depth| depth == data_depth + 1)
        {
            if is_chartml_name(namespace, element.name(), b"chart") {
                if self.graphic_frame_has_chart || self.graphic_frame_has_diagram {
                    return Err(invalid(
                        "graphic frame has duplicate or ambiguous subtype markers",
                    ));
                }
                self.graphic_frame_has_chart = true;
            }
            if is_namespace_name(namespace, element.name(), DIAGRAM_NS, b"relIds") {
                if self.graphic_frame_has_chart || self.graphic_frame_has_diagram {
                    return Err(invalid(
                        "graphic frame has duplicate or ambiguous subtype markers",
                    ));
                }
                self.graphic_frame_has_diagram = true;
            }
        }

        if is_presentationml_name(namespace, element.name(), b"timing") {
            if self.saw_timing {
                return Err(invalid("slide contains multiple timing trees"));
            }
            self.saw_timing = true;
            if empty {
                self.timing_range = Some(event_start..event_end);
            } else {
                self.timing_depth = Some(depth);
                self.timing_start = Some(event_start);
            }
            return Ok(());
        }

        let Some(timing_depth) = self.timing_depth else {
            return Ok(());
        };
        if depth <= timing_depth {
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"bldLst") {
            if self.saw_build_list {
                return Err(invalid("timing tree contains multiple build lists"));
            }
            self.saw_build_list = true;
            if !empty {
                self.build_list_depth = Some(depth);
            }
            return Ok(());
        }

        if self.diagram_build_depth.is_some() {
            return Err(invalid(
                "diagram build elements cannot contain child elements",
            ));
        }
        if self.ole_chart_build_depth.is_some() {
            return Err(invalid(
                "OLE chart build elements cannot contain child elements",
            ));
        }

        if let Some(pending) = self.pending_graphic_build.as_mut() {
            if depth == pending.depth + 1 {
                if pending.mode.is_some() || pending.sub_build_depth.is_some() {
                    return Err(invalid(
                        "graphical-object build has multiple content choices",
                    ));
                }
                if is_presentationml_name(namespace, element.name(), b"bldAsOne") {
                    if !empty {
                        return Err(invalid("graphical-object build-as-one must be empty"));
                    }
                    pending.mode = Some(GraphicBuildMode::AsOne);
                    return Ok(());
                }
                if is_presentationml_name(namespace, element.name(), b"bldSub") {
                    if empty {
                        return Err(invalid(
                            "graphical-object sub-build is missing its build type",
                        ));
                    }
                    pending.sub_build_depth = Some(depth);
                    return Ok(());
                }
                return Err(invalid("graphical-object build has invalid child content"));
            }
            if pending
                .sub_build_depth
                .is_some_and(|sub_depth| depth == sub_depth + 1)
            {
                if pending.mode.is_some() {
                    return Err(invalid(
                        "graphical-object sub-build has multiple build types",
                    ));
                }
                if !empty {
                    return Err(invalid("graphical-object DrawingML build must be empty"));
                }
                if is_drawingml_name(namespace, element.name(), b"bldDgm") {
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| GraphicDiagramBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let reverse = attribute(element, b"rev", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    pending.mode = Some(GraphicBuildMode::Diagram {
                        build_type,
                        reverse,
                    });
                    return Ok(());
                }
                if is_drawingml_name(namespace, element.name(), b"bldChart") {
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| GraphicChartBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(true);
                    pending.mode = Some(GraphicBuildMode::Chart {
                        build_type,
                        animate_background,
                    });
                    return Ok(());
                }
                return Err(invalid(
                    "graphical-object sub-build has invalid DrawingML content",
                ));
            }
            return Err(invalid("graphical-object build has invalid nested content"));
        }

        if let Some(pending) = self.pending_template.as_mut() {
            if let Some(root_depth) = pending.root_depth
                && depth > root_depth
            {
                return Ok(());
            }
            if let Some(time_list_depth) = pending.time_list_depth {
                if depth == time_list_depth + 1 {
                    if !is_presentationml_name(namespace, element.name(), b"par") {
                        return Err(invalid(
                            "paragraph template time list must contain a par node",
                        ));
                    }
                    if pending.root_start.is_some() || pending.root_range.is_some() {
                        return Err(invalid(
                            "paragraph template time list has multiple root nodes",
                        ));
                    }
                    if empty {
                        return Err(invalid("paragraph template par node cannot be empty"));
                    }
                    pending.root_depth = Some(depth);
                    pending.root_start = Some(event_start);
                    return Ok(());
                }
                return Err(invalid(
                    "paragraph template time list has invalid content order",
                ));
            }
            if depth == pending.depth + 1
                && is_presentationml_name(namespace, element.name(), b"tnLst")
            {
                if pending.saw_time_list {
                    return Err(invalid("paragraph template has multiple time lists"));
                }
                if empty {
                    return Err(invalid("paragraph template time list cannot be empty"));
                }
                pending.saw_time_list = true;
                pending.time_list_depth = Some(depth);
                return Ok(());
            }
            return Err(invalid("paragraph template has invalid child content"));
        }

        if let Some(template_list_depth) = self.template_list_depth {
            if depth != template_list_depth + 1
                || !is_presentationml_name(namespace, element.name(), b"tmpl")
            {
                return Err(invalid("paragraph template list has invalid child content"));
            }
            if self.template_levels.len() >= MAX_PARAGRAPH_TEMPLATES {
                return Err(invalid("paragraph template count exceeds PowerPoint limit"));
            }
            let level = attribute(element, b"lvl", decoder)?
                .map(|value| {
                    value
                        .parse::<u8>()
                        .map_err(|_err| invalid("invalid paragraph template level"))
                })
                .transpose()?
                .unwrap_or(0);
            if level > 9 {
                return Err(invalid("paragraph template level exceeds PowerPoint limit"));
            }
            if !self.template_levels.insert(level) {
                return Err(invalid("duplicate paragraph template level"));
            }
            if empty {
                return Err(invalid("paragraph template is missing its time list"));
            }
            let build_index = self
                .paragraph_build_index
                .ok_or_else(|| invalid("paragraph template has no containing build"))?;
            self.pending_template = Some(PendingParagraphTemplate {
                depth,
                build_index,
                level,
                time_list_depth: None,
                saw_time_list: false,
                root_depth: None,
                root_start: None,
                root_range: None,
            });
            return Ok(());
        }

        if let Some(paragraph_build_depth) = self.paragraph_build_depth {
            if depth != paragraph_build_depth + 1
                || !is_presentationml_name(namespace, element.name(), b"tmplLst")
            {
                return Err(invalid("paragraph build has invalid child content"));
            }
            if self.template_list_depth.is_some() {
                return Err(invalid("paragraph build has multiple template lists"));
            }
            if !empty {
                self.template_list_depth = Some(depth);
                self.template_levels.clear();
            }
            return Ok(());
        }

        if self
            .build_list_depth
            .is_some_and(|list_depth| depth == list_depth + 1)
        {
            let kind = if is_presentationml_name(namespace, element.name(), b"bldP") {
                Some(1)
            } else if is_presentationml_name(namespace, element.name(), b"bldDgm") {
                Some(2)
            } else if is_presentationml_name(namespace, element.name(), b"bldGraphic") {
                Some(3)
            } else if is_presentationml_name(namespace, element.name(), b"bldOleChart") {
                Some(4)
            } else {
                None
            };
            if let Some(kind) = kind {
                if self.build_pairs.len() >= MAX_ANIMATION_BUILDS {
                    return Err(invalid("animation build count exceeds safety limit"));
                }
                let shape_id = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation build is missing spid"))
                    .and_then(|value| parse_shape_id(&value))?;
                let group_id = attribute(element, b"grpId", decoder)?
                    .ok_or_else(|| invalid("animation build is missing grpId"))
                    .and_then(|value| parse_group_id(&value))?;
                if !self.build_pairs.insert((kind, shape_id, group_id)) {
                    return Err(invalid("duplicate animation build shape/group pair"));
                }
                self.build_group_ids.insert(group_id);
                if kind == 1 {
                    let build_type = attribute(element, b"build", decoder)?
                        .map(|value| ParagraphBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_level_attribute = attribute(element, b"bldLvl", decoder)?;
                    let build_level = build_level_attribute
                        .as_deref()
                        .map(|value| {
                            value
                                .parse::<u32>()
                                .map_err(|_err| invalid("invalid paragraph build level"))
                        })
                        .transpose()?
                        .unwrap_or(1);
                    if build_level_attribute.is_some()
                        && build_type != ParagraphBuildType::Paragraph
                    {
                        return Err(invalid(
                            "bldLvl is only supported when paragraph build type is p",
                        ));
                    }
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let auto_update_animate_background =
                        attribute(element, b"autoUpdateAnimBg", decoder)?
                            .map(|value| parse_xml_bool(&value))
                            .transpose()?
                            .unwrap_or(true);
                    let reverse_attribute = attribute(element, b"rev", decoder)?;
                    let reverse = reverse_attribute
                        .as_deref()
                        .map(parse_xml_bool)
                        .transpose()?
                        .unwrap_or(false);
                    if reverse_attribute.is_some() && build_type != ParagraphBuildType::Paragraph {
                        return Err(invalid(
                            "rev is only supported when paragraph build type is p",
                        ));
                    }
                    let auto_advance = attribute(element, b"advAuto", decoder)?
                        .map(|value| parse_build_auto_advance(&value))
                        .transpose()?
                        .unwrap_or(Duration::Indefinite);
                    self.sequence.paragraph_builds.push(ParagraphBuild {
                        shape_id,
                        group_id,
                        ui_expand,
                        build_type,
                        build_level,
                        animate_background,
                        auto_update_animate_background,
                        reverse,
                        auto_advance,
                        templates: Vec::new(),
                    });
                    if !empty {
                        self.paragraph_build_depth = Some(depth);
                        self.paragraph_build_index = Some(self.sequence.paragraph_builds.len() - 1);
                    }
                } else if kind == 2 {
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| DiagramBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    self.sequence.diagram_builds.push(DiagramBuild {
                        shape_id,
                        group_id,
                        ui_expand,
                        build_type,
                    });
                    if !empty {
                        self.diagram_build_depth = Some(depth);
                    }
                } else if kind == 3 {
                    if empty {
                        return Err(invalid(
                            "graphical-object build is missing its content choice",
                        ));
                    }
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    self.pending_graphic_build = Some(PendingGraphicBuild {
                        depth,
                        shape_id,
                        group_id,
                        ui_expand,
                        sub_build_depth: None,
                        mode: None,
                    });
                } else if kind == 4 {
                    let ui_expand = attribute(element, b"uiExpand", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(false);
                    let build_type = attribute(element, b"bld", decoder)?
                        .map(|value| OleChartBuildType::parse(&value))
                        .transpose()?
                        .unwrap_or_default();
                    let animate_background = attribute(element, b"animBg", decoder)?
                        .map(|value| parse_xml_bool(&value))
                        .transpose()?
                        .unwrap_or(true);
                    self.sequence.ole_chart_builds.push(OleChartBuild {
                        shape_id,
                        group_id,
                        ui_expand,
                        build_type,
                        animate_background,
                    });
                    if !empty {
                        self.ole_chart_build_depth = Some(depth);
                    }
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"stCondLst") {
            if !empty {
                self.start_conditions_depth.push(depth);
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"cond")
            && !self.start_conditions_depth.is_empty()
        {
            let value = attribute(element, b"delay", decoder)?
                .map(|value| parse_timing_value(&value))
                .transpose()?
                .unwrap_or(TimingValue::Milliseconds(0));
            let frame = self
                .time_nodes
                .last_mut()
                .ok_or_else(|| invalid("animation condition has no containing time node"))?;
            if frame.start_delay.is_none() {
                frame.start_delay = Some(value);
                frame.start_on_click =
                    attribute(element, b"evt", decoder)?.as_deref() == Some("onClick");
            }
            if !empty {
                self.condition_depth.push(depth);
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"tgtEl") {
            if !self.condition_depth.is_empty() {
                if self.condition_target_depth.replace(depth).is_some() {
                    return Err(invalid("animation condition has multiple target elements"));
                }
                return Ok(());
            }
            if let Some(pending) = self.pending.last_mut() {
                if pending.target_element_depth.is_some() {
                    return Err(invalid("animation has multiple target elements"));
                }
                if !empty {
                    pending.target_element_depth = Some(depth);
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"spTgt") {
            if self.condition_target_depth.is_some() {
                let value = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation condition shape target is missing spid"))?;
                let id = parse_shape_id(&value)?;
                let frame = self
                    .time_nodes
                    .last_mut()
                    .ok_or_else(|| invalid("animation condition has no containing time node"))?;
                if frame.start_target.replace(id).is_some() {
                    return Err(invalid("animation condition has multiple shape targets"));
                }
                return Ok(());
            }
            if let Some(pending) = self.pending.last_mut()
                && pending.target_element_depth.is_some()
            {
                let value = attribute(element, b"spid", decoder)?
                    .ok_or_else(|| invalid("animation shape target is missing spid"))?;
                let id = parse_shape_id(&value)?;
                if pending.target.replace(id).is_some() {
                    return Err(invalid("animation has multiple shape targets"));
                }
            }
            return Ok(());
        }

        if is_presentationml_name(namespace, element.name(), b"cTn") {
            let node_type = attribute(element, b"nodeType", decoder)?;
            let event_filter = attribute(element, b"evtFilter", decoder)?;
            let group_id = attribute(element, b"grpId", decoder)?
                .map(|value| parse_group_id(&value))
                .transpose()?;
            if let Some(group_id) = group_id {
                self.timing_group_ids.insert(group_id);
            }
            let is_interactive = node_type.as_deref() == Some("interactiveSeq");
            if event_filter.is_some() && !is_interactive {
                return Err(invalid(
                    "animation event filter is only valid on an interactive sequence",
                ));
            }
            let interactive_event_filter = if is_interactive {
                Some(
                    event_filter
                        .map(|value| EventFilter::parse(&value))
                        .transpose()?,
                )
            } else {
                None
            };
            let preset_id = attribute(element, b"presetID", decoder)?;
            if let Some(preset_id) = preset_id {
                if is_interactive {
                    return Err(invalid(
                        "interactive sequence cannot also be a preset effect",
                    ));
                }
                if self.sequence.len() >= MAX_ANIMATIONS {
                    return Err(invalid("slide animation count exceeds safety limit"));
                }
                let preset_id = preset_id
                    .parse::<u32>()
                    .map_err(|_err| invalid("invalid animation preset ID"))?;
                let preset_class = attribute(element, b"presetClass", decoder)?
                    .unwrap_or_else(|| "entr".to_string());
                if !matches!(
                    preset_class.as_str(),
                    "entr" | "exit" | "emph" | "path" | "verb" | "mediacall"
                ) {
                    return Err(invalid("invalid animation preset class"));
                }
                let preset_subtype = attribute(element, b"presetSubtype", decoder)?
                    .map(|value| {
                        value
                            .parse::<u32>()
                            .map_err(|_err| invalid("invalid animation preset subtype"))
                    })
                    .transpose()?
                    .unwrap_or(0);
                let duration = match attribute(element, b"dur", decoder)? {
                    Some(value) => match parse_timing_value(&value)? {
                        TimingValue::Milliseconds(value) => Duration::Finite(value),
                        TimingValue::Indefinite => Duration::Indefinite,
                    },
                    None => Duration::Finite(0),
                };
                let fill = attribute(element, b"fill", decoder)?
                    .map(|value| Fill::parse(&value))
                    .transpose()?;
                let restart = attribute(element, b"restart", decoder)?
                    .map(|value| Restart::parse(&value))
                    .transpose()?;
                let auto_reverse = attribute(element, b"autoRev", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?
                    .unwrap_or(false);
                let repeat = attribute(element, b"repeatCount", decoder)?
                    .map(|value| {
                        Ok::<Repeat, Error>(match parse_timing_value(&value)? {
                            TimingValue::Milliseconds(value) => Repeat::Finite(value),
                            TimingValue::Indefinite => Repeat::Indefinite,
                        })
                    })
                    .transpose()?;
                let speed = attribute(element, b"spd", decoder)?
                    .map(|value| {
                        let value = value
                            .parse::<i32>()
                            .map_err(|_err| invalid("invalid animation speed percentage"))?;
                        Speed::new(value)
                    })
                    .transpose()?;
                let acceleration = attribute(element, b"accel", decoder)?
                    .map(|value| parse_progress(&value, "acceleration"))
                    .transpose()?;
                let deceleration = attribute(element, b"decel", decoder)?
                    .map(|value| parse_progress(&value, "deceleration"))
                    .transpose()?;
                let display = attribute(element, b"display", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?;
                let repeat_duration = attribute(element, b"repeatDur", decoder)?
                    .map(|value| {
                        Ok::<Duration, Error>(match parse_timing_value(&value)? {
                            TimingValue::Milliseconds(value) => Duration::Finite(value),
                            TimingValue::Indefinite => Duration::Indefinite,
                        })
                    })
                    .transpose()?;
                let sync_behavior = attribute(element, b"syncBehavior", decoder)?
                    .map(|value| SyncBehavior::parse(&value))
                    .transpose()?;
                let after_effect = attribute(element, b"afterEffect", decoder)?
                    .map(|value| parse_xml_bool(&value))
                    .transpose()?;
                let time_filter = attribute(element, b"tmFilter", decoder)?
                    .map(|value| TimeFilter::parse(&value))
                    .transpose()?;
                let sequence_context = parse_sequence_context(&self.time_nodes)?;
                let trigger = trigger(node_type.as_deref(), &self.time_nodes);
                let delay = self
                    .time_nodes
                    .iter()
                    .rev()
                    .find_map(|node| match node.start_delay {
                        Some(TimingValue::Milliseconds(value)) => Some(value),
                        _ => None,
                    })
                    .unwrap_or(0);
                let order = u32::try_from(self.sequence.len() + 1)
                    .map_err(|_err| invalid("animation order exceeds u32"))?;
                let effect = Effect::from_preset_parts(&preset_class, preset_id);
                self.pending.push(PendingAnimation {
                    depth,
                    animation: EffectInstance {
                        shape_id: 0,
                        direction: direction_from_subtype(&effect, preset_subtype),
                        effect,
                        trigger,
                        duration,
                        delay,
                        fill,
                        restart,
                        auto_reverse,
                        repeat,
                        speed,
                        acceleration,
                        deceleration,
                        display,
                        repeat_duration,
                        sync_behavior,
                        after_effect,
                        time_filter,
                        sequence_context,
                        group_id,
                        order,
                    },
                    target: None,
                    target_element_depth: None,
                });
                if empty {
                    return Err(invalid("preset animation has no shape target"));
                }
            }
            if !empty {
                self.time_nodes.push(TimeNodeFrame {
                    depth,
                    start_delay: None,
                    start_on_click: false,
                    start_target: None,
                    interactive_event_filter,
                });
            }
        }

        Ok(())
    }

    pub(super) fn end(
        &mut self,
        namespace: &ResolveResult<'_>,
        name: QName<'_>,
        depth: usize,
        event_end: usize,
    ) -> Result<()> {
        if self.require_valid_targets
            && self.graphic_data_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphicData")
        {
            self.graphic_data_depth = None;
        }
        if self.require_valid_targets
            && self.graphic_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphic")
        {
            if self.graphic_data_depth.is_some() {
                return Err(invalid("graphic frame has an incomplete graphic-data host"));
            }
            self.graphic_depth = None;
        }
        if self.require_valid_targets
            && self.graphic_frame_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"graphicFrame")
        {
            if self.graphic_frame_has_ole_object
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.ole_diagram_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_ole_chart
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.ole_chart_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_chart
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.chart_shape_ids.insert(shape_id);
            }
            if self.graphic_frame_has_diagram
                && let Some(shape_id) = self.graphic_frame_shape_id
            {
                self.graphical_diagram_shape_ids.insert(shape_id);
            }
            self.graphic_frame_depth = None;
            self.graphic_depth = None;
            self.graphic_data_depth = None;
            self.graphic_frame_shape_id = None;
            self.graphic_frame_has_ole_object = false;
            self.graphic_frame_has_ole_chart = false;
            self.graphic_frame_has_chart = false;
            self.graphic_frame_has_diagram = false;
        }

        if self.diagram_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldDgm")
        {
            self.diagram_build_depth = None;
            return Ok(());
        }
        if self.ole_chart_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldOleChart")
        {
            self.ole_chart_build_depth = None;
            return Ok(());
        }
        if let Some(pending) = self.pending_graphic_build.as_mut() {
            if pending.sub_build_depth == Some(depth)
                && is_presentationml_name(namespace, name, b"bldSub")
            {
                if pending.mode.is_none() {
                    return Err(invalid(
                        "graphical-object sub-build is missing its build type",
                    ));
                }
                pending.sub_build_depth = None;
                return Ok(());
            }
            if pending.depth == depth && is_presentationml_name(namespace, name, b"bldGraphic") {
                let pending = self
                    .pending_graphic_build
                    .take()
                    .expect("pending graphical-object build checked above");
                if pending.sub_build_depth.is_some() {
                    return Err(invalid(
                        "graphical-object build has an incomplete sub-build",
                    ));
                }
                let mode = pending.mode.ok_or_else(|| {
                    invalid("graphical-object build is missing its content choice")
                })?;
                self.sequence.graphic_builds.push(GraphicBuild {
                    shape_id: pending.shape_id,
                    group_id: pending.group_id,
                    ui_expand: pending.ui_expand,
                    mode,
                });
                return Ok(());
            }
        }
        if let Some(pending) = self.pending_template.as_mut() {
            if let Some(root_depth) = pending.root_depth {
                if depth > root_depth {
                    return Ok(());
                }
                if depth == root_depth && is_presentationml_name(namespace, name, b"par") {
                    let start = pending
                        .root_start
                        .take()
                        .ok_or_else(|| invalid("paragraph template root offset is missing"))?;
                    pending.root_range = Some(start..event_end);
                    pending.root_depth = None;
                    return Ok(());
                }
            }
            if pending.time_list_depth == Some(depth)
                && is_presentationml_name(namespace, name, b"tnLst")
            {
                if pending.root_range.is_none() {
                    return Err(invalid("paragraph template time list has no root par node"));
                }
                pending.time_list_depth = None;
                return Ok(());
            }
            if pending.depth == depth && is_presentationml_name(namespace, name, b"tmpl") {
                let pending = self
                    .pending_template
                    .take()
                    .expect("pending template checked above");
                if pending.time_list_depth.is_some() || !pending.saw_time_list {
                    return Err(invalid("paragraph template has an incomplete time list"));
                }
                let range = pending
                    .root_range
                    .ok_or_else(|| invalid("paragraph template has no root time node"))?;
                self.template_ranges
                    .push((pending.build_index, pending.level, range));
                return Ok(());
            }
            return Ok(());
        }

        if self.template_list_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"tmplLst")
        {
            self.template_list_depth = None;
            self.template_levels.clear();
            return Ok(());
        }

        if self.paragraph_build_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"bldP")
        {
            self.paragraph_build_depth = None;
            self.paragraph_build_index = None;
            return Ok(());
        }

        if is_presentationml_name(namespace, name, b"tgtEl") {
            if self.condition_target_depth == Some(depth) {
                self.condition_target_depth = None;
            }
            if let Some(pending) = self.pending.last_mut()
                && pending.target_element_depth == Some(depth)
            {
                pending.target_element_depth = None;
            }
        }

        if is_presentationml_name(namespace, name, b"cond")
            && self.condition_depth.last() == Some(&depth)
        {
            self.condition_depth.pop();
        }

        if is_presentationml_name(namespace, name, b"cTn") {
            if let Some(mut pending) = self.pending.pop_if(|pending| pending.depth == depth) {
                pending.animation.shape_id = pending
                    .target
                    .ok_or_else(|| invalid("preset animation has no shape target"))?;
                self.sequence.add(pending.animation);
            }
            let frame = self
                .time_nodes
                .pop()
                .ok_or_else(|| invalid("unbalanced animation time node"))?;
            if frame.depth != depth {
                return Err(invalid("unbalanced animation time-node depth"));
            }
        }

        if is_presentationml_name(namespace, name, b"stCondLst")
            && self.start_conditions_depth.last() == Some(&depth)
        {
            self.start_conditions_depth.pop();
        }
        if is_presentationml_name(namespace, name, b"bldLst")
            && self.build_list_depth == Some(depth)
        {
            self.build_list_depth = None;
        }
        if is_presentationml_name(namespace, name, b"timing") && self.timing_depth == Some(depth) {
            self.timing_depth = None;
            let start = self
                .timing_start
                .take()
                .ok_or_else(|| invalid("timing subtree start offset is missing"))?;
            self.timing_range = Some(start..event_end);
        }
        Ok(())
    }

    pub(super) fn finish(mut self, xml: &[u8]) -> Result<Sequence> {
        if !self.pending.is_empty()
            || !self.time_nodes.is_empty()
            || self.timing_depth.is_some()
            || self.build_list_depth.is_some()
            || self.paragraph_build_depth.is_some()
            || self.template_list_depth.is_some()
            || self.pending_template.is_some()
            || self.diagram_build_depth.is_some()
            || self.ole_chart_build_depth.is_some()
            || self.pending_graphic_build.is_some()
            || self.graphic_frame_depth.is_some()
            || self.graphic_depth.is_some()
            || self.graphic_data_depth.is_some()
        {
            return Err(invalid("incomplete animation timing tree"));
        }
        if self.timing_group_ids != self.build_group_ids {
            return Err(invalid(
                "animation cTn group IDs and build-list group IDs do not match",
            ));
        }
        for (build_index, level, range) in self.template_ranges {
            let raw = xml
                .get(range)
                .ok_or_else(|| invalid("paragraph template range is outside slide XML"))?;
            if raw.len() > MAX_TEMPLATE_TIME_NODE_BYTES {
                return Err(invalid("paragraph template time node exceeds safety limit"));
            }
            let raw = std::str::from_utf8(raw)
                .map_err(|_err| invalid("paragraph template time node is not UTF-8"))?;
            let build = self
                .sequence
                .paragraph_builds
                .get_mut(build_index)
                .ok_or_else(|| invalid("paragraph template build index is invalid"))?;
            build.templates.push(ParagraphTemplate {
                level,
                time_node: TemplateTimeNode::parse(raw)?,
            });
        }
        for build in &self.sequence.paragraph_builds {
            if build.build_type == ParagraphBuildType::Whole && build.templates.len() > 1 {
                return Err(invalid(
                    "whole paragraph builds support exactly one template effect",
                ));
            }
        }
        if self.require_valid_targets {
            for animation in &self.sequence.animations {
                if !self.shape_ids.contains(&animation.shape_id) {
                    return Err(invalid(format!(
                        "animation target {} is not a shape on the current slide",
                        animation.shape_id
                    )));
                }
                if let SequenceContext::Interactive {
                    trigger_shape_id, ..
                } = &animation.sequence_context
                    && !self.shape_ids.contains(trigger_shape_id)
                {
                    return Err(invalid(format!(
                        "interactive animation trigger {trigger_shape_id} is not a shape on the current slide"
                    )));
                }
            }
            for (_, shape_id, _) in &self.build_pairs {
                if !self.shape_ids.contains(shape_id) {
                    return Err(invalid(format!(
                        "animation build target {shape_id} is not a shape on the current slide"
                    )));
                }
            }
            for (kind, shape_id, _) in &self.build_pairs {
                if *kind == 2 && !self.ole_diagram_shape_ids.contains(shape_id) {
                    return Err(invalid(format!(
                        "diagram build target {shape_id} is not an OLE graphic-frame shape"
                    )));
                }
            }
            for build in &self.sequence.graphic_builds {
                let valid = match build.mode {
                    GraphicBuildMode::AsOne => {
                        self.chart_shape_ids.contains(&build.shape_id)
                            || self.graphical_diagram_shape_ids.contains(&build.shape_id)
                    },
                    GraphicBuildMode::Diagram { .. } => {
                        self.graphical_diagram_shape_ids.contains(&build.shape_id)
                    },
                    GraphicBuildMode::Chart { .. } => {
                        self.chart_shape_ids.contains(&build.shape_id)
                    },
                };
                if !valid {
                    return Err(invalid(format!(
                        "graphical-object build target {} does not match its chart/diagram build type",
                        build.shape_id
                    )));
                }
            }
            for build in &self.sequence.ole_chart_builds {
                if !self.ole_chart_shape_ids.contains(&build.shape_id) {
                    return Err(invalid(format!(
                        "OLE chart build target {} is not an embedded chart graphic-frame shape",
                        build.shape_id
                    )));
                }
            }
        }
        if let Some(range) = self.timing_range {
            let raw = xml
                .get(range)
                .ok_or_else(|| invalid("timing subtree range is outside slide XML"))?;
            if raw.len() > MAX_PRESERVED_TIMING_BYTES {
                return Err(invalid("preserved timing subtree exceeds safety limit"));
            }
            let raw =
                std::str::from_utf8(raw).map_err(|_err| invalid("timing subtree is not UTF-8"))?;
            self.sequence.source_animations =
                Some(self.sequence.animations.clone().into_boxed_slice());
            self.sequence.source_paragraph_builds =
                Some(self.sequence.paragraph_builds.clone().into_boxed_slice());
            self.sequence.source_diagram_builds =
                Some(self.sequence.diagram_builds.clone().into_boxed_slice());
            self.sequence.source_graphic_builds =
                Some(self.sequence.graphic_builds.clone().into_boxed_slice());
            self.sequence.source_ole_chart_builds =
                Some(self.sequence.ole_chart_builds.clone().into_boxed_slice());
            let slide_xml = std::str::from_utf8(xml)
                .map_err(|_err| invalid("slide timing XML is not UTF-8"))?;
            let timing_tree = super::wire::parse_recursive_timing_tree(slide_xml)?;
            self.sequence.source_timing_tree = Some(Box::new(timing_tree.clone()));
            self.sequence.timing_tree = Some(timing_tree);
            self.sequence.source_timing_xml = Some(raw.to_string().into_boxed_str());
        }
        Ok(self.sequence)
    }
}
