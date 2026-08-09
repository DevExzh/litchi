//! Typed animation facade and model-preserving XML operations.

use super::super::model::{
    DiagramBuildType, Duration, GraphicBuildMode, GraphicChartBuildType, GraphicDiagramBuildType,
    OleChartBuildType, ParagraphBuildType, Sequence, SequenceContext, TemplateTimeNode, TimingTree,
};
use super::validation::{check_xml_size, validate_template_time_node};
use super::xml::{
    parse_processed_timing, parse_recursive_timing_tree, write_animation_xml, write_timing_child,
};
use crate::Result;

impl TemplateTimeNode {
    /// Validate and store one bounded `p:par` template time node.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &str) -> Result<Self> {
        validate_template_time_node(xml)?;
        Ok(Self {
            xml: xml.to_string().into_boxed_str(),
        })
    }

    /// Exact validated XML for the root `p:par` node.
    #[must_use]
    pub fn as_xml(&self) -> &str {
        &self.xml
    }
}

impl TimingTree {
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &str) -> Result<Self> {
        check_xml_size(xml.len())?;
        let processed = litchi_ooxml_common::mce::process_str(xml)?;
        parse_recursive_timing_tree(&processed)
    }
    #[must_use]
    pub fn to_xml(&self) -> String {
        if let (Some(xml), Some(roots), Some(opaque)) = (
            &self.source_xml,
            &self.source_roots,
            &self.source_opaque_children,
        ) && self.roots.as_slice() == roots.as_ref()
            && self.opaque_children.as_slice() == opaque.as_ref()
        {
            return xml.to_string();
        }
        let mut xml = String::from("<p:timing><p:tnLst>");
        for child in &self.roots {
            write_timing_child(&mut xml, child);
        }
        xml.push_str("</p:tnLst>");
        for child in &self.opaque_children {
            xml.push_str(child);
        }
        xml.push_str("</p:timing>");
        xml
    }
}

impl Sequence {
    /// Parse timing XML from a slide.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_timing_xml(xml: &str) -> Result<Self> {
        check_xml_size(xml.len())?;
        let xml = litchi_ooxml_common::mce::process_str(xml)?;
        check_xml_size(xml.len())?;
        parse_processed_timing(xml.as_bytes(), false)
    }

    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_slide_xml(xml: &[u8]) -> Result<Self> {
        check_xml_size(xml.len())?;
        parse_processed_timing(xml, true)
    }
    /// Generate timing XML for a slide.
    #[must_use]
    pub fn to_xml(&self) -> String {
        if let (
            Some(xml),
            Some(source),
            Some(source_builds),
            Some(source_diagram_builds),
            Some(source_graphic_builds),
            Some(source_ole_chart_builds),
            source_timing_tree,
        ) = (
            &self.source_timing_xml,
            &self.source_animations,
            &self.source_paragraph_builds,
            &self.source_diagram_builds,
            &self.source_graphic_builds,
            &self.source_ole_chart_builds,
            &self.source_timing_tree,
        ) && self.animations.as_slice() == source.as_ref()
            && self.paragraph_builds.as_slice() == source_builds.as_ref()
            && self.diagram_builds.as_slice() == source_diagram_builds.as_ref()
            && self.graphic_builds.as_slice() == source_graphic_builds.as_ref()
            && self.ole_chart_builds.as_slice() == source_ole_chart_builds.as_ref()
            && self.timing_tree.as_ref() == source_timing_tree.as_deref()
        {
            return xml.to_string();
        }
        if self.animations.as_slice() == self.source_animations.as_deref().unwrap_or_default()
            && self.paragraph_builds.as_slice()
                == self.source_paragraph_builds.as_deref().unwrap_or_default()
            && self.diagram_builds.as_slice()
                == self.source_diagram_builds.as_deref().unwrap_or_default()
            && self.graphic_builds.as_slice()
                == self.source_graphic_builds.as_deref().unwrap_or_default()
            && self.ole_chart_builds.as_slice()
                == self.source_ole_chart_builds.as_deref().unwrap_or_default()
            && let Some(timing_tree) = &self.timing_tree
        {
            return timing_tree.to_xml();
        }
        if self.is_empty() {
            return String::new();
        }

        let mut xml = String::with_capacity(2048);
        xml.push_str("<p:timing>");
        xml.push_str("<p:tnLst>");
        xml.push_str(r#"<p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">"#);
        xml.push_str(r#"<p:childTnLst><p:seq concurrent="1" nextAc="seek">"#);
        xml.push_str(r#"<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>"#);

        let mut tn_id = 3u32;
        for anim in self
            .animations
            .iter()
            .filter(|anim| anim.sequence_context == SequenceContext::Main)
        {
            write_animation_xml(&mut xml, anim, &mut tn_id, None);
        }

        xml.push_str("</p:childTnLst></p:cTn>");
        xml.push_str(r#"<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>"#);
        xml.push_str(r#"<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>"#);
        xml.push_str("</p:seq>");

        let mut contexts = Vec::<&SequenceContext>::new();
        for animation in &self.animations {
            if animation.sequence_context != SequenceContext::Main
                && !contexts.contains(&&animation.sequence_context)
            {
                contexts.push(&animation.sequence_context);
            }
        }
        for context in contexts {
            let SequenceContext::Interactive {
                trigger_shape_id,
                event_filter,
            } = context
            else {
                continue;
            };
            xml.push_str(r#"<p:seq concurrent="1" nextAc="seek"><p:cTn"#);
            xml.push_str(&format!(
                r#" id="{tn_id}" dur="indefinite" restart="whenNotActive" nodeType="interactiveSeq""#
            ));
            tn_id += 1;
            if let Some(event_filter) = event_filter {
                xml.push_str(&format!(r#" evtFilter="{}""#, event_filter.as_str()));
            }
            xml.push_str("><p:childTnLst>");
            for animation in self
                .animations
                .iter()
                .filter(|animation| &animation.sequence_context == context)
            {
                write_animation_xml(&mut xml, animation, &mut tn_id, Some(*trigger_shape_id));
            }
            xml.push_str("</p:childTnLst></p:cTn></p:seq>");
        }

        xml.push_str("</p:childTnLst></p:cTn></p:par>");
        xml.push_str("</p:tnLst>");
        if !self.paragraph_builds.is_empty()
            || !self.diagram_builds.is_empty()
            || !self.graphic_builds.is_empty()
            || !self.ole_chart_builds.is_empty()
        {
            xml.push_str("<p:bldLst>");
            for build in &self.paragraph_builds {
                xml.push_str(&format!(
                    r#"<p:bldP spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != ParagraphBuildType::Whole {
                    xml.push_str(&format!(r#" build="{}""#, build.build_type.as_str()));
                }
                if build.build_level != 1 {
                    xml.push_str(&format!(r#" bldLvl="{}""#, build.build_level));
                }
                if build.animate_background {
                    xml.push_str(r#" animBg="1""#);
                }
                if !build.auto_update_animate_background {
                    xml.push_str(r#" autoUpdateAnimBg="0""#);
                }
                if build.reverse {
                    xml.push_str(r#" rev="1""#);
                }
                if build.auto_advance != Duration::Indefinite {
                    xml.push_str(&format!(
                        r#" advAuto="{}""#,
                        build.auto_advance.write_value()
                    ));
                }
                if build.templates.is_empty() {
                    xml.push_str("/>");
                } else {
                    xml.push_str("><p:tmplLst>");
                    for template in &build.templates {
                        xml.push_str("<p:tmpl");
                        if template.level != 0 {
                            xml.push_str(&format!(r#" lvl="{}""#, template.level));
                        }
                        xml.push_str("><p:tnLst>");
                        xml.push_str(template.time_node.as_xml());
                        xml.push_str("</p:tnLst></p:tmpl>");
                    }
                    xml.push_str("</p:tmplLst></p:bldP>");
                }
            }
            for build in &self.diagram_builds {
                xml.push_str(&format!(
                    r#"<p:bldDgm spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != DiagramBuildType::Whole {
                    xml.push_str(&format!(r#" bld="{}""#, build.build_type.as_str()));
                }
                xml.push_str("/>");
            }
            for build in &self.graphic_builds {
                xml.push_str(&format!(
                    r#"<p:bldGraphic spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                xml.push('>');
                match build.mode {
                    GraphicBuildMode::AsOne => xml.push_str("<p:bldAsOne/>"),
                    GraphicBuildMode::Diagram {
                        build_type,
                        reverse,
                    } => {
                        xml.push_str("<p:bldSub><a:bldDgm");
                        if build_type != GraphicDiagramBuildType::AllAtOnce {
                            xml.push_str(&format!(r#" bld="{}""#, build_type.as_str()));
                        }
                        if reverse {
                            xml.push_str(r#" rev="1""#);
                        }
                        xml.push_str("/></p:bldSub>");
                    },
                    GraphicBuildMode::Chart {
                        build_type,
                        animate_background,
                    } => {
                        xml.push_str("<p:bldSub><a:bldChart");
                        if build_type != GraphicChartBuildType::AllAtOnce {
                            xml.push_str(&format!(r#" bld="{}""#, build_type.as_str()));
                        }
                        if !animate_background {
                            xml.push_str(r#" animBg="0""#);
                        }
                        xml.push_str("/></p:bldSub>");
                    },
                }
                xml.push_str("</p:bldGraphic>");
            }
            for build in &self.ole_chart_builds {
                xml.push_str(&format!(
                    r#"<p:bldOleChart spid="{}" grpId="{}""#,
                    build.shape_id,
                    build.group_id.value()
                ));
                if build.ui_expand {
                    xml.push_str(r#" uiExpand="1""#);
                }
                if build.build_type != OleChartBuildType::AllAtOnce {
                    xml.push_str(&format!(r#" bld="{}""#, build.build_type.as_str()));
                }
                if !build.animate_background {
                    xml.push_str(r#" animBg="0""#);
                }
                xml.push_str("/>");
            }
            xml.push_str("</p:bldLst>");
        }
        xml.push_str("</p:timing>");

        xml
    }
}
