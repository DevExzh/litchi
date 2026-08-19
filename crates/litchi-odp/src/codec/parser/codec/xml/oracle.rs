//! Equivalence oracle for the fused `content.xml` pass.
//!
//! `parse_pages_with_styles_sequential` below is the historical two-pass
//! implementation kept verbatim under `cfg(test)`: it resolves transition
//! styles with a full standalone pre-scan of `content.xml`
//! (`resolved_transition_styles`) and then re-tokenizes the same bytes for
//! the slide scan. The production `parse_pages_with_styles` fuses both
//! logical passes into one tokenization; the tests here prove the two are
//! observationally identical — same slides, same errors, same error
//! precedence — on every presentation fixture and on adversarial synthetic
//! inputs that pin each error class.

use super::super::{
    AnimationKind, AnimationNode, DRAW_NAMESPACE, DrawingHyperlink, DrawingShapeKind, Element,
    EnhancedGeometry, Error, Event, Node, NsClass, NsReader, ParagraphText, Parser, Result, Shape,
    ShapeBuilder, ShapeContainerScope, ShapeType, Slide, Speed, Transition, TransitionStyles,
    XLINK_NAMESPACE, XmlVersion, validate_legacy_animation_root,
};
use super::codec::TransitionStyleCollector;

impl Parser {
    /// Sequential oracle for `parse_slides_with_styles`.
    fn parse_slides_with_styles_sequential(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles_sequential::<false>(
            xml_content,
            styles_xml,
            0,
            false,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Sequential oracle for `parse_slide_with_styles_at`.
    fn parse_slide_with_styles_at_sequential(
        xml_content: &str,
        styles_xml: Option<&str>,
        index: usize,
    ) -> Result<Option<Slide>> {
        let mut slides = Self::parse_pages_with_styles_sequential::<true>(
            xml_content,
            styles_xml,
            index,
            false,
            ShapeContainerScope::DrawPages,
        )?;
        Ok(slides.pop())
    }

    /// Sequential oracle for `parse_drawing_pages`.
    fn parse_drawing_pages_sequential(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles_sequential::<false>(
            xml_content,
            styles_xml,
            0,
            true,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Sequential oracle for `parse_sheet_shape_tables`.
    fn parse_sheet_shape_tables_sequential(xml_content: &str) -> Result<Vec<Vec<Shape>>> {
        let tables = Self::parse_pages_with_styles_sequential::<false>(
            xml_content,
            None,
            0,
            true,
            ShapeContainerScope::SpreadsheetTables,
        )?;
        Ok(tables.into_iter().map(|table| table.shapes).collect())
    }
    pub(super) fn parse_pages_with_styles_sequential<const SELECT_ONE: bool>(
        xml_content: &str,
        styles_xml: Option<&str>,
        selected_index: usize,
        retain_text_shapes: bool,
        container_scope: ShapeContainerScope,
    ) -> Result<Vec<Slide>> {
        let sheet_scope = container_scope == ShapeContainerScope::SpreadsheetTables;
        let (transition_styles, default_transition) =
            Self::resolved_transition_styles(xml_content, styles_xml)?;
        let mut reader = NsReader::from_str(xml_content);
        let mut buf = Vec::new();
        // Nested subtree parsers feed their events to a throwaway collector:
        // historically the standalone pre-scan had already collected every
        // transition definition before this slide scan ran.
        let mut discarded_collector = TransitionStyleCollector::default();
        let mut slides = Vec::new();

        // State tracking
        let mut current_slide_text = String::new();
        let mut current_slide_title: Option<String> = None;
        let mut current_shapes: Vec<Shape> = Vec::new();
        let mut in_slide = false;
        let mut slide_index = 0;
        let mut current_notes_text = String::new();
        let mut current_notes_has_paragraph = false;
        let mut in_notes = false;
        let mut current_slide_has_segment = false;
        let mut current_transition: Option<Transition> = None;
        let mut current_animations = Vec::new();
        let mut animation_node_count = 0;
        let mut current_legacy_animation = None;
        let mut legacy_animation_node_count = 0;
        let mut shape_node_count = 0usize;

        // Shape parsing state
        let mut shape_stack: Vec<ShapeBuilder> = Vec::new();
        let mut current_paragraph: Option<ParagraphText> = None;
        let mut in_media_plugin = false;
        let mut in_media_parameter = false;
        let mut current_hyperlink: Option<DrawingHyperlink> = None;
        let mut hyperlink_parent_depth = None;
        let mut hyperlink_shape_seen = false;

        // Spreadsheet `table:shapes` container state
        let mut element_depth = 0usize;
        let mut spreadsheet_depth: Option<usize> = None;
        let mut sheet_table_depth: Option<usize> = None;
        let mut sheet_shapes_depth: Option<usize> = None;
        let mut sheet_table_has_shapes = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) => {
                    element_depth = element_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XML element depth overflow".to_string())
                    })?;
                    let element_type = Self::classify(
                        NsClass::from_resolve(&namespace),
                        element.local_name().as_ref(),
                    );
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                if !SELECT_ONE || slide_index == selected_index {
                                    slides.push(Slide {
                                        title: current_slide_title.take(),
                                        text: std::mem::take(&mut current_slide_text),
                                        index: slide_index,
                                        notes: (!current_notes_text.is_empty())
                                            .then(|| std::mem::take(&mut current_notes_text)),
                                        transition: current_transition.take(),
                                        animations: std::mem::take(&mut current_animations),
                                        legacy_animation: current_legacy_animation.take(),
                                        shapes: std::mem::take(&mut current_shapes),
                                    });
                                } else {
                                    current_slide_text.clear();
                                    current_notes_text.clear();
                                    current_animations.clear();
                                    current_legacy_animation = None;
                                    current_shapes.clear();
                                }
                                slide_index += 1;
                            }
                            current_slide_title = None;
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            let style_name =
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            current_transition = (!transition.is_empty()).then_some(transition);
                            in_slide = true;
                        },
                        Element::Notes if in_slide => in_notes = true,
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last() {
                                if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                    return Err(Error::InvalidFormat(
                                        "draw:enhanced-geometry requires draw:custom-shape"
                                            .to_string(),
                                    ));
                                }
                                if builder.enhanced_geometry.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "draw:custom-shape contains multiple enhanced geometries"
                                            .to_string(),
                                    ));
                                }
                            }
                            let geometry = Self::parse_enhanced_geometry(
                                &mut reader,
                                &mut discarded_collector,
                                element,
                            )?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.enhanced_geometry = Some(geometry);
                            }
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::LegacyAnimation(kind)
                            if in_slide
                                && !in_notes
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if kind != AnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            let root = Self::parse_legacy_animation_node(
                                &mut reader,
                                &mut discarded_collector,
                                element,
                                kind,
                                1,
                                &mut legacy_animation_node_count,
                            )?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::Plugin if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if !builder.is_frame {
                                    return Err(Error::InvalidFormat(
                                        "draw:plugin must be contained directly by draw:frame"
                                            .to_string(),
                                    ));
                                }
                                if builder.media.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "ODP frame contains multiple draw:plugin elements"
                                            .to_string(),
                                    ));
                                }
                                builder.shape_type = ShapeType::GraphicFrame;
                                builder.media = Some(Self::media_reference(&reader, element)?);
                                in_media_plugin = true;
                            }
                        },
                        Element::Plugin if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin must be contained by a drawing shape".to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            if let Some(media) = shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                            {
                                media.add_parameter(Self::media_parameter(&reader, element)?)?;
                                in_media_parameter = true;
                            }
                        },
                        Element::DrawingHyperlink
                            if in_slide && !in_notes && current_hyperlink.is_none() =>
                        {
                            current_hyperlink = Some(Self::drawing_hyperlink(&reader, element)?);
                            hyperlink_parent_depth = Some(shape_stack.len());
                            hyperlink_shape_seen = false;
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "nested or misplaced draw:a presentation hyperlink".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder
                                    .drawing_kind
                                    .is_some_and(DrawingShapeKind::is_three_dimensional)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D shapes cannot contain presentation event listeners"
                                            .to_string(),
                                    ));
                                }
                                if builder.event_listeners_seen {
                                    return Err(Error::InvalidFormat(
                                        "ODP shape contains multiple office:event-listeners elements"
                                            .to_string(),
                                    ));
                                }
                                builder.event_listeners = Self::parse_event_listeners(
                                    &mut reader,
                                    &mut discarded_collector,
                                )?;
                                element_depth = Self::rewind_consumed_subtree(element_depth);
                                builder.event_listeners_seen = true;
                            }
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            if current_paragraph.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested ODP text paragraphs are not supported".to_string(),
                                ));
                            }
                            current_paragraph = Some(ParagraphText::default());
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            if let Some(paragraph) = current_paragraph.as_mut() {
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::push_text_control(
                                        &reader,
                                        element,
                                        element_type,
                                        paragraph,
                                    )?;
                                } else {
                                    let mut ignored = ParagraphText::default();
                                    Self::push_text_control(
                                        &reader,
                                        element,
                                        element_type,
                                        &mut ignored,
                                    )?;
                                }
                            }
                        },
                        _ if in_notes => {},
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            current_animations.push(Self::parse_animation_node(
                                &mut reader,
                                &mut discarded_collector,
                                element,
                                kind,
                                1,
                                &mut animation_node_count,
                            )?);
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                        },
                        Element::SpreadsheetRoot if sheet_scope => {
                            spreadsheet_depth = Some(element_depth);
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            sheet_table_depth = Some(element_depth);
                            sheet_table_has_shapes = false;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                            sheet_shapes_depth = Some(element_depth);
                            in_slide = true;
                        },
                        Element::Shape(shape_element) => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            if shape_stack.len() >= 64 {
                                return Err(Error::InvalidFormat(
                                    "ODP shape groups exceed 64 levels".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            if in_slide && shape_stack.is_empty() {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                if current_hyperlink.is_some() && !hyperlink_applies {
                                    return Err(Error::InvalidFormat(
                                        "misplaced draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            } else if let Some(parent) = shape_stack.last() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            }
                        },
                        Element::Image if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        Element::Table if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        Element::Object if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Image
                        | Element::Table
                        | Element::Object
                        | Element::Plugin
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::Text(ref text) if current_paragraph.is_some() => {
                    let decoded = Self::decode_text(text)?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&decoded);
                    }
                },
                Event::Text(ref text) if in_media_plugin => {
                    let decoded = Self::decode_text(text)?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::Text(ref text)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    let decoded = Self::decode_text(text)?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "3D drawing elements cannot contain text".to_string(),
                        ));
                    }
                },
                Event::CData(ref text) if current_paragraph.is_some() => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid presentation CDATA: {error}"))
                    })?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&decoded);
                    }
                },
                Event::CData(ref text) if in_media_plugin => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid media plugin CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::GeneralRef(ref reference) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    if (!SELECT_ONE || slide_index == selected_index)
                        && let Some(paragraph) = current_paragraph.as_mut()
                    {
                        paragraph.push_text(&text);
                    }
                },
                Event::GeneralRef(_) if in_media_plugin => {
                    return Err(Error::InvalidFormat(
                        "draw:plugin cannot contain character references".to_string(),
                    ));
                },
                Event::GeneralRef(_)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain character references".to_string(),
                    ));
                },
                Event::CData(ref data)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) && !data.iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain CDATA text".to_string(),
                    ));
                },
                Event::Empty(ref element) => {
                    let element_type = Self::classify(
                        NsClass::from_resolve(&namespace),
                        element.local_name().as_ref(),
                    );
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope && !in_slide => {
                            let style_name =
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            if !SELECT_ONE || slide_index == selected_index {
                                slides.push(Slide {
                                    title: None,
                                    text: String::new(),
                                    index: slide_index,
                                    notes: None,
                                    transition: (!transition.is_empty()).then_some(transition),
                                    animations: Vec::new(),
                                    legacy_animation: None,
                                    shapes: Vec::new(),
                                });
                            }
                            slide_index += 1;
                        },
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                    return Err(Error::InvalidFormat(
                                        "draw:enhanced-geometry requires draw:custom-shape"
                                            .to_string(),
                                    ));
                                }
                                if builder.enhanced_geometry.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "draw:custom-shape contains multiple enhanced geometries"
                                            .to_string(),
                                    ));
                                }
                                builder.enhanced_geometry = Some(EnhancedGeometry {
                                    attributes: Self::exact_geometry_attributes(&reader, element)?,
                                    children: Vec::new(),
                                });
                            }
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::Plugin => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if !builder.is_frame {
                                    return Err(Error::InvalidFormat(
                                        "draw:plugin must be contained directly by draw:frame"
                                            .to_string(),
                                    ));
                                }
                                if builder.media.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "ODP frame contains multiple draw:plugin elements"
                                            .to_string(),
                                    ));
                                }
                                builder.shape_type = ShapeType::GraphicFrame;
                                builder.media = Some(Self::media_reference(&reader, element)?);
                            } else if in_slide {
                                return Err(Error::InvalidFormat(
                                    "draw:plugin must be contained by a drawing shape".to_string(),
                                ));
                            }
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:a must wrap exactly one non-empty drawing shape".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if builder
                                    .drawing_kind
                                    .is_some_and(DrawingShapeKind::is_three_dimensional)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D shapes cannot contain presentation event listeners"
                                            .to_string(),
                                    ));
                                }
                                if builder.event_listeners_seen {
                                    return Err(Error::InvalidFormat(
                                        "ODP shape contains multiple office:event-listeners elements"
                                            .to_string(),
                                    ));
                                }
                                builder.event_listeners_seen = true;
                            }
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            if let Some(media) = shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                            {
                                media.add_parameter(Self::media_parameter(&reader, element)?)?;
                            }
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            if !SELECT_ONE || slide_index == selected_index {
                                Self::push_parsed_paragraph(
                                    "",
                                    in_notes,
                                    &mut current_notes_text,
                                    &mut current_notes_has_paragraph,
                                    shape_stack.last_mut(),
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                );
                            }
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            if let Some(paragraph) = current_paragraph.as_mut() {
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::push_text_control(
                                        &reader,
                                        element,
                                        element_type,
                                        paragraph,
                                    )?;
                                } else {
                                    let mut ignored = ParagraphText::default();
                                    Self::push_text_control(
                                        &reader,
                                        element,
                                        element_type,
                                        &mut ignored,
                                    )?;
                                }
                            }
                        },
                        _ if in_notes => {},
                        Element::LegacyAnimation(kind) if in_slide => {
                            if kind != AnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            legacy_animation_node_count =
                                legacy_animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "legacy ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            let root = AnimationNode::from_parsed(
                                kind,
                                Self::animation_attributes(&reader, element)?,
                                Vec::new(),
                            );
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            animation_node_count =
                                animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            if animation_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP animation tree exceeds 65536 nodes".to_string(),
                                ));
                            }
                            current_animations.push(Node::from_parsed(
                                kind,
                                Self::animation_attributes(&reader, element)?,
                                Vec::new(),
                            ));
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: String::new(),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                        },
                        Element::Image => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        Element::Table => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        Element::Object => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        Element::Shape(shape_element) if in_slide => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            let mut builder = Self::shape_builder(&reader, element, shape_element)?;
                            if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                builder.hyperlink = Some(hyperlink.clone());
                                hyperlink_shape_seen = true;
                            }
                            if let Some(parent) = shape_stack.last_mut() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    parent.children.push(builder.build());
                                }
                            } else {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::finish_shape(
                                        builder,
                                        &mut current_slide_title,
                                        &mut current_slide_text,
                                        &mut current_slide_has_segment,
                                        &mut current_shapes,
                                        retain_text_shapes,
                                    );
                                }
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Shape(_)
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::End(ref element) => {
                    element_depth = element_depth.saturating_sub(1);
                    let element_type = Self::classify(
                        NsClass::from_resolve(&namespace),
                        element.local_name().as_ref(),
                    );
                    if matches!(element_type, Element::TextParagraph)
                        && let Some(parsed_paragraph) = current_paragraph.take()
                    {
                        if !SELECT_ONE || slide_index == selected_index {
                            let paragraph = parsed_paragraph.finish();
                            Self::push_parsed_paragraph(
                                &paragraph,
                                in_notes,
                                &mut current_notes_text,
                                &mut current_notes_has_paragraph,
                                shape_stack.last_mut(),
                                &mut current_slide_text,
                                &mut current_slide_has_segment,
                            );
                        }
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Notes) {
                        in_notes = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Plugin) {
                        in_media_plugin = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::PluginParameter) && in_media_parameter {
                        in_media_parameter = false;
                        buf.clear();
                        continue;
                    }
                    if in_notes {
                        buf.clear();
                        continue;
                    }
                    match element_type {
                        Element::DrawingHyperlink if current_hyperlink.is_some() => {
                            if hyperlink_parent_depth != Some(shape_stack.len())
                                || !hyperlink_shape_seen
                            {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one complete drawing shape"
                                        .to_string(),
                                ));
                            }
                            current_hyperlink = None;
                            hyperlink_parent_depth = None;
                            hyperlink_shape_seen = false;
                        },
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                if current_hyperlink.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "unterminated draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    slides.push(Slide {
                                        title: current_slide_title.take(),
                                        text: std::mem::take(&mut current_slide_text),
                                        index: slide_index,
                                        notes: (!current_notes_text.is_empty())
                                            .then(|| std::mem::take(&mut current_notes_text)),
                                        transition: current_transition.take(),
                                        animations: std::mem::take(&mut current_animations),
                                        legacy_animation: current_legacy_animation.take(),
                                        shapes: std::mem::take(&mut current_shapes),
                                    });
                                } else {
                                    current_slide_title = None;
                                    current_slide_text.clear();
                                    current_notes_text.clear();
                                    current_transition = None;
                                    current_animations.clear();
                                    current_legacy_animation = None;
                                    current_shapes.clear();
                                }
                                slide_index += 1;
                            }
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = false;
                        },
                        Element::SpreadsheetRoot
                            if sheet_scope
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            spreadsheet_depth = None;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_shapes_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            if current_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "unterminated draw:a drawing hyperlink".to_string(),
                                ));
                            }
                            sheet_shapes_depth = None;
                            in_slide = false;
                        },
                        Element::Table
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: std::mem::take(&mut current_slide_text),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: std::mem::take(&mut current_shapes),
                            });
                            slide_index += 1;
                            sheet_table_depth = None;
                            current_slide_has_segment = false;
                        },
                        Element::Shape(_) => {
                            if let Some(builder) = shape_stack.pop() {
                                if let Some(parent) = shape_stack.last_mut() {
                                    if !SELECT_ONE || slide_index == selected_index {
                                        parent.children.push(builder.build());
                                    }
                                    buf.clear();
                                    continue;
                                }
                                if !SELECT_ONE || slide_index == selected_index {
                                    Self::finish_shape(
                                        builder,
                                        &mut current_slide_title,
                                        &mut current_slide_text,
                                        &mut current_slide_has_segment,
                                        &mut current_shapes,
                                        retain_text_shapes,
                                    );
                                }
                            }
                        },
                        Element::Page
                        | Element::Notes
                        | Element::SheetShapes
                        | Element::SpreadsheetRoot
                        | Element::Image
                        | Element::Table
                        | Element::Object
                        | Element::Plugin
                        | Element::PluginParameter
                        | Element::DrawingHyperlink
                        | Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                        | Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                        | Element::TextParagraph
                        | Element::TextSpace
                        | Element::TextTab
                        | Element::TextLineBreak
                        | Element::Animation(_)
                        | Element::UnknownAnimation
                        | Element::LegacyAnimation(_)
                        | Element::Other => {},
                    }
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
            buf.clear();
        }

        Ok(slides)
    }
}

/// Drives a [`TransitionStyleCollector`] over one XML source exactly like the
/// fused pass does, returning the first collection error in scan order.
fn collect_via_collector(xml: &str) -> Result<TransitionStyles> {
    let mut reader = NsReader::from_str(xml);
    let mut collector = TransitionStyleCollector::default();
    let mut buf = Vec::new();
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buf)
            .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
        let ns_class = NsClass::from_resolve(&namespace);
        collector.feed(&reader, ns_class, &event);
        if let Some(error) = collector.take_error() {
            return Err(error);
        }
        if matches!(event, Event::Eof) {
            break;
        }
        buf.clear();
    }
    Ok(collector.finish())
}

/// Asserts the fused and sequential outcomes are identical: equal values on
/// success, equal formatted errors on failure.
fn assert_outcome<T: PartialEq + std::fmt::Debug>(
    fused: &Result<T>,
    sequential: &Result<T>,
    context: &str,
) {
    match (fused, sequential) {
        (Ok(fused_value), Ok(sequential_value)) => {
            assert_eq!(fused_value, sequential_value, "value mismatch: {context}");
        },
        (Err(fused_error), Err(sequential_error)) => {
            assert_eq!(
                format!("{fused_error}"),
                format!("{sequential_error}"),
                "error mismatch: {context}"
            );
        },
        _ => panic!("outcome divergence: {context}: fused={fused:?} sequential={sequential:?}"),
    }
}

/// Cross-checks every public entry point of the page parser on one input.
fn assert_equivalent(content: &str, styles: Option<&str>, context: &str) {
    let fused = Parser::parse_slides_with_styles(content, styles);
    let sequential = Parser::parse_slides_with_styles_sequential(content, styles);
    assert_outcome(
        &fused,
        &sequential,
        &format!("{context} [parse_slides_with_styles]"),
    );

    let count = fused.as_ref().map_or(0, Vec::len);
    for index in 0..=count + 1 {
        let fused_one = Parser::parse_slide_with_styles_at(content, styles, index);
        let sequential_one = Parser::parse_slide_with_styles_at_sequential(content, styles, index);
        assert_outcome(
            &fused_one,
            &sequential_one,
            &format!("{context} [parse_slide_with_styles_at {index}]"),
        );
    }

    let fused_drawing = Parser::parse_drawing_pages(content, styles);
    let sequential_drawing = Parser::parse_drawing_pages_sequential(content, styles);
    assert_outcome(
        &fused_drawing,
        &sequential_drawing,
        &format!("{context} [parse_drawing_pages]"),
    );
}

/// Loads `(content.xml, styles.xml, name)` for every `.odp` and `.fodp`
/// fixture under `test-data/`, skipping archives that do not decode.
fn presentation_fixtures() -> Vec<(String, Option<String>, String)> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut fixtures = Vec::new();
    let mut stack = vec![root];
    while let Some(directory) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&directory) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            match path.extension().and_then(|extension| extension.to_str()) {
                Some("odp") => {
                    let Ok(bytes) = std::fs::read(&path) else {
                        continue;
                    };
                    let Ok(archive) = soapberry_zip::office::ArchiveReader::new(&bytes) else {
                        continue;
                    };
                    let Ok(content) = archive.read_string("content.xml") else {
                        continue;
                    };
                    let styles = archive
                        .contains("styles.xml")
                        .then(|| archive.read_string("styles.xml"))
                        .transpose()
                        .ok()
                        .flatten();
                    fixtures.push((content, styles, path.display().to_string()));
                },
                Some("fodp") => {
                    let Ok(xml) = std::fs::read_to_string(&path) else {
                        continue;
                    };
                    fixtures.push((xml.clone(), Some(xml), path.display().to_string()));
                },
                _ => {},
            }
        }
    }
    fixtures
}

const NS: &str = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
    xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" \
    xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" \
    xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" \
    xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" \
    xmlns:anim=\"urn:oasis:names:tc:opendocument:xmlns:animation:1.0\"";

fn doc(body: &str) -> String {
    format!("<office:document-content {NS}>{body}</office:document-content>")
}

#[test]
fn fused_pass_matches_sequential_oracle_on_fixtures() {
    let fixtures = presentation_fixtures();
    assert!(
        fixtures.len() >= 15,
        "expected at least 15 presentation fixtures, found {}",
        fixtures.len()
    );
    for (content, styles, name) in &fixtures {
        assert_equivalent(content, styles.as_deref(), name);
    }
}

#[test]
fn collector_matches_standalone_transition_scan_on_fixtures() {
    let fixtures = presentation_fixtures();
    assert!(
        fixtures.len() >= 15,
        "expected at least 15 presentation fixtures, found {}",
        fixtures.len()
    );
    for (content, _, name) in &fixtures {
        let standalone = Parser::parse_transition_style_definitions(content);
        let collected = collect_via_collector(content);
        assert_outcome(&standalone, &collected, name);
    }
}

#[test]
fn transition_definition_after_referencing_page() {
    // The fused pass resolves transitions only after collecting every
    // definition, so a definition behind its referencing page still applies —
    // exactly what the historical up-front pre-scan produced.
    let content = doc("<draw:page draw:style-name=\"later\"/>\
         <style:style style:name=\"later\" style:family=\"drawing-page\">\
         <style:drawing-page-properties presentation:transition-type=\"manual\"/>\
         </style:style>");
    assert_equivalent(&content, None, "definition after referencing page");
    let slides = Parser::parse_slides_with_styles(&content, None).unwrap();
    assert_eq!(slides.len(), 1);
    assert!(slides[0].transition.is_some());
}

#[test]
fn collection_error_beats_slide_scan_error() {
    // The style:style nested in anim:seq is a slide-scan error ("contains a
    // non-animation element"); its invalid transition property is a
    // transition-collection error. Historically the pre-scan surfaced the
    // collection error first, so the fused pass must as well.
    let content = doc("<draw:page><anim:seq>\
         <style:style style:name=\"x\" style:family=\"drawing-page\">\
         <style:drawing-page-properties presentation:duration=\"bogus\"/>\
         </style:style></anim:seq></draw:page>");
    assert_equivalent(&content, None, "collection error precedence");
    let error = Parser::parse_slides_with_styles(&content, None).unwrap_err();
    let message = format!("{error}");
    assert!(message.contains("presentation:duration"), "{message}");
    assert!(!message.contains("non-animation"), "{message}");
}

#[test]
fn tokenization_error_after_slide_scan_error_wins() {
    // The unknown animation element is an early slide-scan error, but the
    // truncated tag later in the stream is a tokenization error that the
    // historical pre-scan always surfaced first.
    let content = format!(
        "{}<draw:page><anim:bogus/></draw:page><broken",
        "<office:document-content ".to_string() + NS + ">"
    );
    assert_equivalent(&content, None, "read error after semantic error");
    let error = Parser::parse_slides_with_styles(&content, None).unwrap_err();
    assert!(
        format!("{error}").contains("XML parsing error: "),
        "{error}"
    );
}

#[test]
fn tokenization_error_inside_transition_region() {
    let content = format!(
        "<office:document-content {NS}>\
         <style:style style:name=\"x\" style:family=\"drawing-page\">\
         <style:drawing-page-properties presentation:transition-type=\"manual\""
    );
    assert_equivalent(&content, None, "truncated transition region");
    let error = Parser::parse_slides_with_styles(&content, None).unwrap_err();
    assert!(
        format!("{error}").contains("XML parsing error: "),
        "{error}"
    );
}

#[test]
fn tokenization_error_inside_enhanced_geometry_uses_read_mapping() {
    // A tokenization failure observed by the nested enhanced-geometry parser
    // must surface with the same "XML parsing error" mapping as every other
    // read site: historically the pre-scan tokenized the whole part first, so
    // no other message could ever reach the caller.
    let content = format!(
        "<office:document-content {NS}>\
         <draw:page><draw:custom-shape><draw:enhanced-geometry><draw:equation"
    );
    assert_equivalent(&content, None, "truncated enhanced geometry");
    let error = Parser::parse_slides_with_styles(&content, None).unwrap_err();
    assert!(
        format!("{error}").contains("XML parsing error: "),
        "{error}"
    );
}

#[test]
fn duplicate_attributes_in_style_region_and_body() {
    let style_region = doc(
        "<style:style style:name=\"x\" style:name=\"y\" style:family=\"drawing-page\"/>\
         <draw:page draw:style-name=\"x\"/>",
    );
    assert_equivalent(&style_region, None, "duplicate attribute in style region");

    let body = doc("<draw:page draw:style-name=\"a\" draw:style-name=\"b\"/>");
    assert_equivalent(&body, None, "duplicate attribute in page element");
}

#[test]
fn custom_namespace_prefixes() {
    let content = "<office:document-content \
        xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
        xmlns:d=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" \
        xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" \
        xmlns:p=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\">\
        <s:style s:name=\"cp\" s:family=\"drawing-page\">\
        <s:drawing-page-properties p:transition-type=\"manual\"/>\
        </s:style><d:page d:style-name=\"cp\"/></office:document-content>";
    assert_equivalent(content, None, "custom prefixes");
    let slides = Parser::parse_slides_with_styles(content, None).unwrap();
    assert_eq!(slides.len(), 1);
    assert!(slides[0].transition.is_some());
}

#[test]
fn cyclic_inheritance_error_beats_slide_scan_error() {
    // Style resolution runs after collection but before the deferred
    // slide-scan error, matching the historical up-front resolution.
    let styles = "<office:document-styles \
        xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
        xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\">\
        <style:style style:name=\"a\" style:family=\"drawing-page\" style:parent-style-name=\"b\"/>\
        <style:style style:name=\"b\" style:family=\"drawing-page\" style:parent-style-name=\"a\"/>\
        </office:document-styles>";
    let content = doc("<draw:page draw:style-name=\"a\"><anim:bogus/></draw:page>");
    // Both passes must fail with the cyclic-inheritance error rather than the
    // slide-scan's unknown-animation error. The failing style name itself is
    // not compared: resolution iterates a HashMap, so which side of the cycle
    // is named first was already nondeterministic in the historical code.
    const CYCLIC: &str = "cyclic or excessively deep drawing-page style inheritance at '";
    let fused = Parser::parse_slides_with_styles(&content, Some(styles));
    let sequential = Parser::parse_slides_with_styles_sequential(&content, Some(styles));
    for (label, outcome) in [("fused", fused), ("sequential", sequential)] {
        let message = format!("{}", outcome.unwrap_err());
        assert!(message.contains(CYCLIC), "{label}: {message}");
        assert!(
            !message.contains("unknown ODF animation"),
            "{label}: {message}"
        );
    }
}

#[test]
fn content_default_style_overrides_styles_default() {
    let styles = "<office:document-styles \
        xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
        xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" \
        xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\">\
        <style:default-style style:family=\"drawing-page\">\
        <style:drawing-page-properties presentation:transition-type=\"manual\"/>\
        </style:default-style></office:document-styles>";
    let content = doc("<style:default-style style:family=\"drawing-page\">\
         <style:drawing-page-properties presentation:transition-speed=\"fast\"/>\
         </style:default-style><draw:page/>");
    assert_equivalent(
        &content,
        Some(styles),
        "content default overrides styles default",
    );
    let slides = Parser::parse_slides_with_styles(&content, Some(styles)).unwrap();
    assert_eq!(slides.len(), 1);
    let transition = slides[0].transition.as_ref().unwrap();
    assert_eq!(transition.speed, Some(Speed::Fast));
    assert_eq!(transition.transition_type, None);
}

#[test]
fn select_one_with_errors_around_selected_index() {
    let styles = "<office:document-styles \
        xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
        xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" \
        xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\">\
        <style:style style:name=\"p0\" style:family=\"drawing-page\">\
        <style:drawing-page-properties presentation:transition-speed=\"slow\"/>\
        </style:style>\
        <style:style style:name=\"p1\" style:family=\"drawing-page\">\
        <style:drawing-page-properties presentation:transition-speed=\"medium\"/>\
        </style:style></office:document-styles>";
    let content = doc("<draw:page draw:style-name=\"p0\"/>\
         <draw:page draw:style-name=\"p1\"/>\
         <draw:page draw:style-name=\"missing\"/>");
    assert_equivalent(&content, Some(styles), "select-one across indices");
    let slide = Parser::parse_slide_with_styles_at(&content, Some(styles), 1)
        .unwrap()
        .unwrap();
    assert_eq!(slide.index, 1);
    assert_eq!(
        slide
            .transition
            .as_ref()
            .and_then(|transition| transition.speed),
        Some(Speed::Medium)
    );

    // Slide-scan errors surface even when the failing page is not selected.
    let broken = doc("<draw:page/><draw:page><anim:bogus/></draw:page><draw:page/>");
    assert_equivalent(&broken, None, "select-one with error on deselected page");
    assert!(Parser::parse_slide_with_styles_at(&broken, None, 2).is_err());
}

#[test]
fn sheet_shape_tables_match_sequential_oracle() {
    let content = "<office:spreadsheet \
        xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" \
        xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" \
        xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" \
        xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" \
        xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" \
        xmlns:presentation=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\">\
        <style:style style:name=\"s\" style:family=\"drawing-page\">\
        <style:drawing-page-properties presentation:transition-type=\"manual\"/>\
        </style:style>\
        <table:table><table:shapes><draw:frame><text:p>hi</text:p></draw:frame>\
        </table:shapes></table:table></office:spreadsheet>";
    let fused = Parser::parse_sheet_shape_tables(content);
    let sequential = Parser::parse_sheet_shape_tables_sequential(content);
    assert_outcome(&fused, &sequential, "sheet shape tables");
    let tables = fused.unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].len(), 1);
    assert_eq!(tables[0][0].text, "hi");
}

#[test]
fn empty_and_missing_style_references() {
    assert_equivalent("", None, "empty content");
    let content = doc("<draw:page draw:style-name=\"undefined\"/>");
    assert_equivalent(&content, None, "undefined style reference");
    let slides = Parser::parse_slides_with_styles(&content, None).unwrap();
    assert_eq!(slides.len(), 1);
    assert!(slides[0].transition.is_none());
}
